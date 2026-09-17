"""AI Excel 公式生成 —— 自然语言交互入口。

一句话说需求，程序自动完成：读表结构 → DeepSeek 生成公式 → 本地校验（失败最多重试 2 次）
→ 预览确认 → openpyxl 写入 → 打印结果与 Token 用量。

用法：
    python main.py                      # 进入交互模式
    python main.py "给每科算总分写到G2"   # 单条指令模式
"""
from __future__ import annotations

import re
import sys
import time
from pathlib import Path

from excel_formula.config import ConfigError, SecurityError, Settings, normalize_thinking
from excel_formula.evaluator import UnsupportedFormula
from excel_formula.formula_parser import FormulaSyntaxError
from excel_formula.intent import (
    INTENT_DESCRIBE,
    INTENT_EXPLAIN,
    INTENT_GENERATE,
    INTENT_VALIDATE,
    classify,
    file_candidates,
)
from excel_formula.llm_client import LLMError
from excel_formula.logger import get_logger
from excel_formula.pipeline import FormulaService, Proposal

BANNER = """
============================================================
  AI Excel 公式生成助手（DeepSeek + openpyxl）
------------------------------------------------------------
  支持的说法示例：
    · 帮我在 G2 算每个科目的总分，填充到 G6
    · 统计 B 列大于 85 分的人数，放在 B8
    · 这个表有哪些列？
    · =SUMIF(B2:B6,">85") 这个公式对不对？
    · 解释一下 G2 里的公式
  指令：:file <路径>  :sheet <表名>  :think <1|0|auto>  :info  :help  :quit
============================================================
"""

HELP = """
可用指令：
  :file <路径>   切换当前 Excel 文件（仅允许工作目录内的 .xlsx/.xlsm）
  :sheet <表名>  指定工作表（默认取第一张有数据的表）
  :think <档位>  思考模式：1 深度思考常开 / 0 关闭（最快）/ auto 自动；不带参数看当前
  :info          查看当前文件结构（不消耗 Token）
  :help          显示帮助
  :quit / :exit  退出
其他输入按自然语言处理；写文件前一定会先预览并请你确认。
"""

# 剔除文件名后剩下的标点/空白，不算需求内容
_FILLER_RE = re.compile(r"[\s\"'“”‘’，,。；;：:、（）()]+")


def _is_bare_file_reference(text: str, file: str | None) -> bool:
    """只报了个文件名、没说要算什么。

    这类输入以前会把文件名本身当成需求发给模型，模型只能反复猜用户想干什么，
    白花 Token 还容易把额度耗在思考上导致返回空内容。
    """
    if not file:
        return False
    return not _FILLER_RE.sub("", text.replace(file, " "))


def _display_width(text: str) -> int:
    """估算终端显示宽度：CJK 字符按 2 列计，用于进度行重绘时补空格对齐。"""
    return sum(2 if ord(ch) > 0x2E7F else 1 for ch in text)


class _StreamProgress:
    """流式生成期间的一行原地刷新进度；非交互终端自动静默。

    思考型模型（DeepSeek V4.x 默认开启思考）可能几十秒里只输出思维链，
    这个进度行让用户明确知道模型还在工作。
    """

    def __init__(self, interval: float = 0.2) -> None:
        self.enabled = bool(getattr(sys.stdout, "isatty", lambda: False)())
        self._interval = interval
        self._thinking = 0
        self._body = 0
        self._started = time.perf_counter()
        self._last_paint = 0.0
        self._painted_width = 0

    def __call__(self, piece: str, kind: str) -> None:
        if kind == "reasoning":
            self._thinking += len(piece)
        else:
            self._body += len(piece)
        if not self.enabled:
            return
        now = time.perf_counter()
        if now - self._last_paint < self._interval:
            return  # 高频分片按固定节奏合并重绘，避免刷屏
        self._last_paint = now
        self._paint(
            f"… 模型思考 {self._thinking} 字 / 正文 {self._body} 字，已用 {now - self._started:.0f}s"
        )

    def _paint(self, message: str) -> None:
        width = _display_width(message)
        padding = " " * max(0, self._painted_width - width)
        print(f"\r{message}{padding}", end="", flush=True)
        self._painted_width = width

    def finish(self) -> None:
        """清掉进度行，把终端交还给正常输出。"""
        if self._painted_width:
            print("\r" + " " * self._painted_width + "\r", end="", flush=True)
            self._painted_width = 0


_THINKING_LABELS = {
    "enabled": "深度思考常开（最准也最慢）",
    "disabled": "关闭思考（最快）",
    "auto": "自动（按需求复杂度切换）",
    "": "服务端默认（思考开启、强度 high）",
}


def apply_thinking_command(settings: Settings, arg: str) -> str:
    """处理 :think 交互命令：更新 settings.thinking 并返回反馈文本。

    1/on=深度思考常开，0/off=关闭思考（最快），auto=按复杂度自动；
    不带参数时只报告当前档位。改动立即生效，client 与 pipeline 共享同一份 settings。
    """
    current = _THINKING_LABELS.get(settings.thinking, settings.thinking)
    if not arg.strip():
        return (
            f"当前思考模式：{current}；"
            "用法：:think 1 深度思考 / :think 0 关闭（最快）/ :think auto 自动"
        )
    try:
        value = normalize_thinking(arg)
    except ConfigError as exc:
        return f"× {exc}（当前仍为：{current}）"
    settings.thinking = value
    return f"√ 思考模式已切换：{_THINKING_LABELS.get(value, value)}"


class Console:
    """交互会话：维护当前文件/工作表，负责追问与确认。"""

    def __init__(self) -> None:
        self.settings = Settings.from_env(workspace=Path.cwd())
        self.logger = get_logger(self.settings.log_dir)
        self.service = FormulaService(self.settings, self.logger)
        self.current_file: Path | None = self._guess_file()
        self.current_sheet: str | None = None

    # -------------------------------------------------------------- 启动辅助
    def _guess_file(self) -> Path | None:
        """工作目录里只有一个 Excel 时自动选中，省去用户输入路径。"""
        candidates = [
            p for p in sorted(Path.cwd().glob("*.xls[xm]"))
            if not p.name.startswith("~$") and "backups" not in p.parts
        ]
        return candidates[0] if len(candidates) >= 1 else None

    def _need_file(self) -> Path | None:
        if self.current_file and self.current_file.is_file():
            return self.current_file
        answer = input("要操作哪个 Excel 文件？（输入路径）> ").strip().strip('"')
        if not answer:
            return None
        return self._set_file(answer)

    def _set_file(self, raw: str) -> Path | None:
        if not raw.strip():
            # 空参数会解析到工作目录本身，报出“收到：ExcelCR”这种误导信息，直接给用法
            print("? 用法：:file <文件名>，例如 :file 销售数据.xlsx")
            return None
        path, first_error = None, None
        for candidate in file_candidates(raw):
            try:
                path = self.settings.resolve_path(candidate)
                break
            except (SecurityError, FileNotFoundError) as exc:
                first_error = first_error or exc
        if path is None:
            print(f"× {first_error}")
            return None
        self.current_file = path
        self.current_sheet = None
        print(f"√ 当前文件：{path.name}")
        return path

    # -------------------------------------------------------------- 主循环
    def run(self, once: str | None = None) -> int:
        if not once:
            print(BANNER)
            current = _THINKING_LABELS.get(self.settings.thinking, self.settings.thinking)
            print(f"思考模式：{current}（:think 1/0/auto 随时切换）")
            if self.current_file:
                print(f"当前文件：{self.current_file.name}（用 :file 可切换）\n")
            if not self.settings.api_key:
                print("提示：尚未配置 DEEPSEEK_API_KEY，生成/解释类功能不可用；")
                print("      复制 .env.example 为 .env 并填入密钥即可。\n")

        if once:
            return 0 if self.handle(once) else 1

        while True:
            try:
                line = input("你 > ").strip()
            except (EOFError, KeyboardInterrupt):
                print("\n再见。")
                return 0
            if not line:
                continue
            if line in {":quit", ":exit", "quit", "exit"}:
                print("再见。")
                return 0
            if line == ":help":
                print(HELP)
                continue
            if line.startswith(":file"):
                self._set_file(line[5:].strip().strip('"'))
                continue
            if line.startswith(":sheet"):
                self.current_sheet = line[6:].strip() or None
                print(f"√ 当前工作表：{self.current_sheet or '自动'}")
                continue
            if line == ":think" or line.startswith(":think "):
                print(apply_thinking_command(self.settings, line[6:].strip()))
                continue
            if line == ":info":
                self.show_info()
                continue
            self.handle(line)

    # -------------------------------------------------------------- 意图分发
    def handle(self, text: str) -> bool:
        intent = classify(text)
        if intent.file and not self._set_file(intent.file):
            # 指定的文件没能打开时直接停下，不能默默拿上一个（或启动时自动猜的）文件继续
            return False
        if intent.kind == INTENT_GENERATE and _is_bare_file_reference(text, intent.file):
            print("? 还不知道要算什么，请把需求说出来，例如「在G2算每个人的总分」。")
            return True
        sheet = intent.sheet or self.current_sheet
        path = self._need_file()
        if not path:
            print("× 没有可操作的文件。")
            return False

        self.logger.info("用户指令 kind=%s text=%s", intent.kind, text[:100])
        try:
            if intent.kind == INTENT_DESCRIBE:
                self.show_info(sheet)
            elif intent.kind == INTENT_VALIDATE:
                self.do_validate(path, intent.formula, sheet, intent.target)
            elif intent.kind == INTENT_EXPLAIN:
                self.do_explain(path, intent.formula, sheet, intent.cell)
            else:
                self.do_generate(path, intent.request, sheet, intent.target)
        except ConfigError as exc:
            print(f"× 配置错误：{exc}")
            return False
        except (SecurityError, FileNotFoundError, KeyError, ValueError, FormulaSyntaxError) as exc:
            print(f"× {exc}")
            return False
        except (LLMError, UnsupportedFormula) as exc:
            print(f"× 模型调用失败：{exc}")
            return False
        except PermissionError as exc:
            print(f"× {exc}")
            return False
        return True

    # -------------------------------------------------------------- 各意图实现
    def show_info(self, sheet: str | None = None) -> None:
        path = self._need_file()
        if not path:
            return
        result = self.service.describe(path, sheet or self.current_sheet)
        print(f"\n文件：{result['file']}")
        for item in result["sheets"]:
            first_row = ", ".join(h for h in item["first_row"] if h) or "（首行为空）"
            print(f"  · {item['sheet']}：{item['rows']} 行 × {item['columns']} 列 | 首行：{first_row}")
        print("\n交给模型的表格内容：")
        print(result["digest_text"], "\n")

    def do_generate(self, path: Path, request: str, sheet: str | None, target: str | None) -> None:
        print("… 正在读取表结构并生成公式")
        proposal = self._propose_with_progress(path, request, sheet, target)

        if proposal.clarification:
            print(f"? {proposal.clarification}")
            extra = input("你 > ").strip()
            if not extra:
                print("× 已取消。")
                return
            merged = f"{request}（补充：{extra}）"
            new_target = target or classify(extra).target
            proposal = self._propose_with_progress(path, merged, sheet, new_target)
            if proposal.clarification:
                print(f"? 仍缺少信息：{proposal.clarification}")
                return

        print("\n----- 公式预览 -----")
        print(proposal.render())
        print("--------------------")
        self._print_usage(proposal.usage.to_dict())
        if not proposal.ok:
            print("× 未生成可用公式，未对文件做任何修改。")
            return

        cells = ", ".join(w.cell for w in proposal.writes[:5])
        suffix = " ..." if len(proposal.writes) > 5 else ""
        answer = input(f"确认写入 {proposal.sheet}!{cells}{suffix} ？[y/N] ").strip().lower()
        if answer not in {"y", "yes", "是"}:
            print("已取消，文件未改动。")
            self.logger.info("用户取消写入 target=%s", proposal.target)
            return

        applied = self.service.apply(proposal)
        print(f"√ 已写入 {applied['count']} 个单元格 → {Path(applied['file']).name}")
        if applied.get("backup"):
            print(f"  原文件已备份：{Path(applied['backup']).name}")
        print("  日志：logs/excelcr.log")

    def _propose_with_progress(
        self, path: Path, request: str, sheet: str | None, target: str | None
    ) -> Proposal:
        """调用 propose 并展示流式进度；非交互终端自动退回静默等待。"""
        progress = _StreamProgress()
        try:
            return self.service.propose(
                path,
                request,
                sheet=sheet,
                target=target,
                on_delta=progress if progress.enabled else None,
            )
        finally:
            progress.finish()

    def do_validate(self, path: Path, formula: str | None, sheet: str | None, target: str | None) -> None:
        if not formula:
            formula = input("要校验哪个公式？（以 = 开头）> ").strip()
        if not formula:
            print("× 没有收到公式。")
            return
        result = self.service.validate(path, formula, sheet=sheet, target=target)
        validation = result["validation"]
        print(f"\n公式：{result['formula']}")
        print("校验：" + ("通过 √" if validation["ok"] else "未通过 ×"))
        for item in validation["errors"]:
            print(f"  错误：{item}")
        for item in validation["warnings"]:
            print(f"  提示：{item}")
        if validation["functions"]:
            print("  使用函数：" + ", ".join(validation["functions"]))
        if result["predicted_value"] is not None:
            print(f"  本地独立计算结果：{result['predicted_value']}")
        elif result["predicted_note"]:
            print(f"  本地计算：未验证（{result['predicted_note']}）")
        print()

    def do_explain(self, path: Path, formula: str | None, sheet: str | None, cell: str | None) -> None:
        if not formula and not cell:
            cell = input("要解释哪个单元格的公式？（如 G2）> ").strip().upper() or None
        progress = _StreamProgress()
        try:
            result = self.service.explain(
                formula,
                file=path,
                sheet=sheet,
                cell=cell,
                on_delta=progress if progress.enabled else None,
            )
        finally:
            progress.finish()
        print(f"\n公式：{result['formula']}\n{result['explanation']}")
        if not result["validation"]["ok"]:
            print("注意：" + "；".join(result["validation"]["errors"]))
        self._print_usage(result["usage"])
        print()

    @staticmethod
    def _print_usage(usage: dict) -> None:
        if usage.get("calls"):
            print(
                f"[本次消耗] 模型调用 {usage['calls']} 次，"
                f"输入 {usage['prompt_tokens']} tokens，输出 {usage['completion_tokens']} tokens，"
                f"耗时 {usage['elapsed_seconds']}s"
            )


def main() -> int:
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8")  # type: ignore[union-attr]
        except (AttributeError, ValueError):
            pass
    once = " ".join(sys.argv[1:]).strip() or None
    return Console().run(once)


if __name__ == "__main__":
    raise SystemExit(main())
