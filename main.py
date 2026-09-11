"""AI Excel 公式生成 —— 自然语言交互入口。

一句话说需求，程序自动完成：读表结构 → DeepSeek 生成公式 → 本地校验（失败最多重试 2 次）
→ 预览确认 → openpyxl 写入 → 打印结果与 Token 用量。

用法：
    python main.py                      # 进入交互模式
    python main.py "给每科算总分写到G2"   # 单条指令模式
"""
from __future__ import annotations

import sys
from pathlib import Path

from excel_formula.config import ConfigError, SecurityError, Settings
from excel_formula.evaluator import UnsupportedFormula
from excel_formula.formula_parser import FormulaSyntaxError
from excel_formula.intent import (
    INTENT_DESCRIBE,
    INTENT_EXPLAIN,
    INTENT_VALIDATE,
    classify,
)
from excel_formula.llm_client import LLMError
from excel_formula.logger import get_logger
from excel_formula.pipeline import FormulaService

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
  指令：:file <路径>  :sheet <表名>  :info  :help  :quit
============================================================
"""

HELP = """
可用指令：
  :file <路径>   切换当前 Excel 文件（仅允许工作目录内的 .xlsx/.xlsm）
  :sheet <表名>  指定工作表（默认取第一张有数据的表）
  :info          查看当前文件结构（不消耗 Token）
  :help          显示帮助
  :quit / :exit  退出
其他输入按自然语言处理；写文件前一定会先预览并请你确认。
"""


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
        try:
            path = self.settings.resolve_path(raw)
        except (SecurityError, FileNotFoundError) as exc:
            print(f"× {exc}")
            return None
        self.current_file = path
        self.current_sheet = None
        print(f"√ 当前文件：{path.name}")
        return path

    # -------------------------------------------------------------- 主循环
    def run(self, once: str | None = None) -> int:
        if not once:
            print(BANNER)
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
            if line == ":info":
                self.show_info()
                continue
            self.handle(line)

    # -------------------------------------------------------------- 意图分发
    def handle(self, text: str) -> bool:
        intent = classify(text)
        if intent.file:
            self._set_file(intent.file)
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
        proposal = self.service.propose(path, request, sheet=sheet, target=target)

        if proposal.clarification:
            print(f"? {proposal.clarification}")
            extra = input("你 > ").strip()
            if not extra:
                print("× 已取消。")
                return
            merged = f"{request}（补充：{extra}）"
            new_target = target or classify(extra).target
            proposal = self.service.propose(path, merged, sheet=sheet, target=new_target)
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
        result = self.service.explain(formula, file=path, sheet=sheet, cell=cell)
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
