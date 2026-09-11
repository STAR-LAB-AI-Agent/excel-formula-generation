"""命令行接口：可独立运行、可被 Skill/智能体调用的确定性入口。

退出码约定：0 成功 / 2 校验未通过或需要用户补充信息 / 3 参数与安全错误 / 4 模型或网络错误。
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from .config import ConfigError, SecurityError, Settings
from .evaluator import UnsupportedFormula
from .formula_parser import FormulaSyntaxError
from .intent import (
    INTENT_DESCRIBE,
    INTENT_EXPLAIN,
    INTENT_VALIDATE,
    classify,
)
from .llm_client import LLMError
from .logger import get_logger
from .pipeline import FormulaService, Proposal

EXIT_OK = 0
EXIT_REJECTED = 2
EXIT_BAD_INPUT = 3
EXIT_LLM = 4


def _ensure_utf8() -> None:
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8")  # type: ignore[union-attr]
        except (AttributeError, ValueError):
            pass


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="excel-formula",
        description="AI Excel 公式生成：自然语言 → DeepSeek → 校验 → 确认 → 写入",
    )
    parser.add_argument("--json", action="store_true", help="以 JSON 输出，便于智能体解析")
    parser.add_argument("--verbose", action="store_true", help="控制台输出调试日志")
    parser.add_argument("--workspace", help="额外允许访问的目录（默认当前工作目录）")
    sub = parser.add_subparsers(dest="command", required=True)

    p_desc = sub.add_parser("describe", help="查看工作簿结构与紧凑表格描述")
    p_desc.add_argument("--file", required=True)
    p_desc.add_argument("--sheet")

    p_gen = sub.add_parser("generate", help="根据自然语言需求生成公式（默认只预览）")
    p_gen.add_argument("--file", required=True)
    p_gen.add_argument("--request", required=True, help="自然语言需求")
    p_gen.add_argument("--sheet")
    p_gen.add_argument("--target", help="目标单元格，如 G2")
    p_gen.add_argument("--apply", action="store_true", help="校验通过后写入文件")
    p_gen.add_argument("--yes", action="store_true", help="跳过交互确认（脚本模式）")
    p_gen.add_argument("--output", help="另存为新文件而不是原地修改")

    p_val = sub.add_parser("validate", help="校验公式并给出本地计算的预期结果")
    p_val.add_argument("--file", required=True)
    p_val.add_argument("--formula", required=True)
    p_val.add_argument("--sheet")
    p_val.add_argument("--target")

    p_exp = sub.add_parser("explain", help="解释一个公式")
    p_exp.add_argument("--formula")
    p_exp.add_argument("--file")
    p_exp.add_argument("--sheet")
    p_exp.add_argument("--cell", help="解释该单元格中已有的公式")

    p_write = sub.add_parser("write", help="直接写入指定公式（仍会先校验）")
    p_write.add_argument("--file", required=True)
    p_write.add_argument("--target", required=True)
    p_write.add_argument("--formula", required=True)
    p_write.add_argument("--sheet")
    p_write.add_argument("--fill-to", dest="fill_to")
    p_write.add_argument("--yes", action="store_true")
    p_write.add_argument("--output")

    p_nl = sub.add_parser("nl", help="一句自然语言，自动判断意图")
    p_nl.add_argument("text", help="例如：在测试数据.xlsx里给每科算总分，写到G2")
    p_nl.add_argument("--file", help="指令里没写文件名时使用")
    p_nl.add_argument("--sheet")
    p_nl.add_argument("--apply", action="store_true")
    p_nl.add_argument("--yes", action="store_true")
    return parser


# ------------------------------------------------------------------ 输出辅助
def _emit(payload: dict, text: str, as_json: bool) -> None:
    if as_json:
        print(json.dumps(payload, ensure_ascii=False, indent=2))
    else:
        print(text)


def _confirm(prompt: str, assume_yes: bool) -> bool:
    """高风险操作（写文件）前的用户确认。"""
    if assume_yes:
        return True
    if not sys.stdin or not sys.stdin.isatty():
        print("需要确认才能写入文件：非交互环境请显式加 --yes", file=sys.stderr)
        return False
    answer = input(f"{prompt} [y/N] ").strip().lower()
    return answer in {"y", "yes", "是"}


def _proposal_output(proposal: Proposal, applied: dict | None) -> tuple[dict, str]:
    payload = proposal.to_dict()
    payload["applied"] = applied
    text = proposal.render()
    if applied:
        text += f"\n已写入 {applied['count']} 个单元格 → {applied['file']}"
        if applied.get("backup"):
            text += f"\n备份: {applied['backup']}"
    elif proposal.ok:
        text += "\n（预览模式，未写入。确认无误后加 --apply 执行）"
    text += "\nToken 用量: " + json.dumps(payload["usage"], ensure_ascii=False)
    return payload, text


# ------------------------------------------------------------------ 各子命令
def _cmd_describe(service: FormulaService, args) -> int:
    result = service.describe(args.file, args.sheet)
    lines = [f"文件: {result['file']}"]
    for sheet in result["sheets"]:
        first_row = ", ".join(h for h in sheet["first_row"] if h) or "（首行为空）"
        lines.append(
            f"  - {sheet['sheet']}: {sheet['rows']} 行 × {sheet['columns']} 列 | 首行: {first_row}"
        )
    lines.append("")
    lines.append(result["digest_text"])
    _emit(result, "\n".join(lines), args.json)
    return EXIT_OK


def _cmd_generate(service: FormulaService, args) -> int:
    proposal = service.propose(args.file, args.request, sheet=args.sheet, target=args.target)
    applied = None
    if proposal.ok and args.apply:
        cells = ", ".join(w.cell for w in proposal.writes[:5])
        if _confirm(f"确认把公式写入 {proposal.sheet}!{cells} ?", args.yes):
            applied = service.apply(proposal, output=args.output)
        else:
            payload, text = _proposal_output(proposal, None)
            _emit(payload, text + "\n已取消写入。", args.json)
            return EXIT_OK
    payload, text = _proposal_output(proposal, applied)
    _emit(payload, text, args.json)
    if proposal.clarification:
        return EXIT_REJECTED
    return EXIT_OK if proposal.ok else EXIT_REJECTED


def _cmd_validate(service: FormulaService, args) -> int:
    result = service.validate(args.file, args.formula, sheet=args.sheet, target=args.target)
    validation = result["validation"]
    lines = [f"公式: {result['formula']}", "校验: " + ("通过" if validation["ok"] else "未通过")]
    lines.extend(f"  错误: {e}" for e in validation["errors"])
    lines.extend(f"  提示: {w}" for w in validation["warnings"])
    if validation["functions"]:
        lines.append("  使用函数: " + ", ".join(validation["functions"]))
    if validation["refs"]:
        lines.append("  引用: " + ", ".join(validation["refs"]))
    if result["predicted_value"] is not None:
        lines.append(f"  本地计算结果: {result['predicted_value']}")
    elif result["predicted_note"]:
        lines.append(f"  本地计算: 未验证（{result['predicted_note']}）")
    _emit(result, "\n".join(lines), args.json)
    return EXIT_OK if validation["ok"] else EXIT_REJECTED


def _cmd_explain(service: FormulaService, args) -> int:
    result = service.explain(args.formula, file=args.file, sheet=args.sheet, cell=args.cell)
    text = f"公式: {result['formula']}\n{result['explanation']}"
    if not result["validation"]["ok"]:
        text += "\n注意: " + "；".join(result["validation"]["errors"])
    _emit(result, text, args.json)
    return EXIT_OK


def _cmd_write(service: FormulaService, args) -> int:
    if not _confirm(f"确认把 {args.formula} 写入 {args.target} ?", args.yes):
        _emit({"applied": None, "cancelled": True}, "已取消写入。", args.json)
        return EXIT_OK
    result = service.write_formula(
        args.file,
        args.target,
        args.formula,
        sheet=args.sheet,
        fill_to=args.fill_to,
        output=args.output,
    )
    _emit(result, f"已写入 {result['count']} 个单元格 → {result['file']}", args.json)
    return EXIT_OK


def _cmd_nl(service: FormulaService, args) -> int:
    intent = classify(args.text)
    file = intent.file or args.file
    sheet = intent.sheet or args.sheet
    if not file:
        payload = {"ok": False, "intent": intent.to_dict(), "question": "请告诉我要操作哪个 Excel 文件？"}
        _emit(payload, payload["question"], args.json)
        return EXIT_REJECTED

    if intent.kind == INTENT_DESCRIBE:
        return _cmd_describe(service, argparse.Namespace(file=file, sheet=sheet, json=args.json))
    if intent.kind == INTENT_VALIDATE:
        return _cmd_validate(
            service,
            argparse.Namespace(
                file=file, formula=intent.formula, sheet=sheet, target=intent.target, json=args.json
            ),
        )
    if intent.kind == INTENT_EXPLAIN:
        return _cmd_explain(
            service,
            argparse.Namespace(
                formula=intent.formula, file=file, sheet=sheet, cell=intent.cell, json=args.json
            ),
        )
    return _cmd_generate(
        service,
        argparse.Namespace(
            file=file,
            request=intent.request,
            sheet=sheet,
            target=intent.target,
            apply=args.apply,
            yes=args.yes,
            output=None,
            json=args.json,
        ),
    )


_COMMANDS = {
    "describe": _cmd_describe,
    "generate": _cmd_generate,
    "validate": _cmd_validate,
    "explain": _cmd_explain,
    "write": _cmd_write,
    "nl": _cmd_nl,
}


def main(argv: list[str] | None = None) -> int:
    _ensure_utf8()
    args = build_parser().parse_args(argv)
    settings = Settings.from_env(workspace=args.workspace or Path.cwd())
    logger = get_logger(settings.log_dir, verbose=args.verbose)
    service = FormulaService(settings, logger)

    try:
        return _COMMANDS[args.command](service, args)
    except (SecurityError, FileNotFoundError, KeyError, ValueError, FormulaSyntaxError) as exc:
        logger.warning("请求被拒绝：%s", exc)
        _emit({"ok": False, "error": str(exc)}, f"错误: {exc}", args.json)
        return EXIT_BAD_INPUT
    except ConfigError as exc:
        logger.error("配置错误：%s", exc)
        _emit({"ok": False, "error": str(exc)}, f"配置错误: {exc}", args.json)
        return EXIT_BAD_INPUT
    except (LLMError, UnsupportedFormula) as exc:
        logger.error("模型调用失败：%s", exc)
        _emit({"ok": False, "error": str(exc)}, f"模型错误: {exc}", args.json)
        return EXIT_LLM
    except PermissionError as exc:
        logger.error("写入失败：%s", exc)
        _emit({"ok": False, "error": str(exc)}, f"写入失败: {exc}", args.json)
        return EXIT_BAD_INPUT
    except KeyboardInterrupt:
        print("\n已中断。", file=sys.stderr)
        return EXIT_OK


if __name__ == "__main__":
    raise SystemExit(main())
