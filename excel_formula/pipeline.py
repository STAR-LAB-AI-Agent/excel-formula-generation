"""流水线编排：需求 -> 摘要 -> 模型 -> 校验(带重试) -> 预览 -> 写入。"""
from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable

from .config import Settings
from .evaluator import UnsupportedFormula, evaluate_formula, format_value
from .excel_reader import SheetDigest, WorkbookView
from .formula_parser import FormulaSyntaxError, parse_target
from .llm_client import (
    DeepSeekClient,
    LLMError,
    Usage,
    build_explain_messages,
    build_generate_messages,
    build_repair_message,
    extract_candidates,
    parse_json_payload,
)
from .validator import ValidationResult, validate_formula
from .writer import CellWrite, apply_writes, expand_fill


@dataclass
class Attempt:
    """一次生成尝试的记录，用于日志与"重试了几次"的展示。"""

    round: int
    formula: str
    ok: bool
    errors: list[str] = field(default_factory=list)

    def to_dict(self) -> dict:
        return {"round": self.round, "formula": self.formula, "ok": self.ok, "errors": self.errors}


@dataclass
class Proposal:
    """待用户确认的公式方案。"""

    ok: bool
    file: Path
    sheet: str
    writes: list[CellWrite] = field(default_factory=list)
    explanation: str = ""
    validation: ValidationResult | None = None
    predicted: str | None = None
    predicted_note: str | None = None
    assumptions: list[str] = field(default_factory=list)
    clarification: str | None = None
    attempts: list[Attempt] = field(default_factory=list)
    usage: Usage = field(default_factory=Usage)
    error: str | None = None

    @property
    def formula(self) -> str:
        return self.writes[0].formula if self.writes else ""

    @property
    def target(self) -> str:
        return self.writes[0].cell if self.writes else ""

    def to_dict(self) -> dict:
        return {
            "ok": self.ok,
            "file": str(self.file),
            "sheet": self.sheet,
            "target": self.target,
            "formula": self.formula,
            "explanation": self.explanation,
            "cells": [w.to_dict() for w in self.writes],
            "validation": self.validation.to_dict() if self.validation else None,
            "predicted_value": self.predicted,
            "predicted_note": self.predicted_note,
            "assumptions": self.assumptions,
            "clarification": self.clarification,
            "attempts": [a.to_dict() for a in self.attempts],
            "usage": self.usage.to_dict(),
            "error": self.error,
        }

    def render(self) -> str:
        """给用户看的预览文本。"""
        if self.clarification:
            return f"需要补充信息：{self.clarification}"
        if not self.ok:
            return f"未能生成可用公式：{self.error or '校验未通过'}"

        lines = [f"文件: {self.file.name}    工作表: {self.sheet}"]
        # 值写入恒列出内容；公式只列每个填充块的首格（其余落进下方的“填充”范围行）
        anchors = [
            w for w in self.writes
            if not w.formula or w.explanation or w.predicted or w.predicted_note
        ]
        if len(anchors) > 1:
            for item in anchors:
                if item.formula:
                    line = f"  {item.cell} = {item.formula}"
                    if item.predicted is not None:
                        line += f"    预期结果: {item.predicted}"
                    elif item.predicted_note:
                        line += f"    预期结果: 未验证（{item.predicted_note}）"
                else:
                    line = f"  {item.cell} = {item.value}（值）"
                lines.append(line)
                if item.explanation:
                    lines.append(f"      说明: {item.explanation}")
            lines.append("（预期结果为本地独立计算，未写入文件）")
        else:
            first = self.writes[0] if self.writes else None
            is_value = first is not None and not first.formula
            if is_value:
                lines.append(f"目标: {self.target}    值: {first.value}")
            else:
                lines.append(f"目标: {self.target}    公式: {self.formula}")
            if self.explanation:
                lines.append(f"说明: {self.explanation}")
            if not is_value:
                if self.predicted is not None:
                    lines.append(f"预期结果: {self.predicted}（本地独立计算，未写入文件）")
                elif self.predicted_note:
                    lines.append(f"预期结果: 未验证（{self.predicted_note}）")
        # 自动填充的单元格按连续段报告范围（同时填充多列时不会合并成误导的大范围）
        for span in self._fill_spans(anchors):
            lines.append(f"填充: {span[0]} → {span[-1]}（共 {len(span)} 个单元格）")
        overwrite = [w for w in self.writes if w.old_formula or w.old_value is not None]
        if overwrite:
            preview = ", ".join(
                f"{w.cell}={w.old_formula or w.old_value}" for w in overwrite[:3]
            )
            lines.append(f"⚠ 将覆盖已有内容: {preview}{' ...' if len(overwrite) > 3 else ''}")
        if self.validation and self.validation.warnings:
            lines.extend(f"提示: {w}" for w in self.validation.warnings)
        if self.assumptions:
            lines.extend(f"假设: {a}" for a in self.assumptions)
        if len(self.attempts) > 1:
            lines.append(f"（经过 {len(self.attempts)} 次生成，前 {len(self.attempts) - 1} 次校验未通过）")
        return "\n".join(lines)

    def _fill_spans(self, anchors: list[CellWrite]) -> list[list[str]]:
        """自动填充格子按连续段分组；返回每段单元格列表（单格段与空段剔除）。"""
        anchor_ids = {id(w) for w in anchors}
        spans: list[list[str]] = []
        current: list[str] = []
        for item in self.writes:
            if id(item) in anchor_ids or not item.formula:
                # 锚点或值写入会打断填充段
                if current:
                    spans.append(current)
                current = []
            else:
                current.append(item.cell)
        if current:
            spans.append(current)
        return [span for span in spans if len(span) > 1]


# EXCELCR_THINKING=auto 时的"复杂需求"关键词：命中即保留思考模式
_COMPLEX_SIGNALS = (
    "跨表", "另一张", "多条件", "同时满足", "并且", "嵌套", "查找", "匹配",
    "排名", "占比", "百分比", "累计", "环比", "同比", "汇总", "分档", "评级",
)
# 值写入（分类名、标题等常量）的长度上限：防止模型把整段说明塞进单元格
_MAX_VALUE_LENGTH = 200


class FormulaService:
    """三类意图的统一入口：生成公式、校验公式、解释公式（外加表结构概览）。"""

    def __init__(self, settings: Settings, logger=None, client: object | None = None):
        self.settings = settings
        self.logger = logger
        self.client = client or DeepSeekClient(settings, logger)

    # ============================================================ 意图 0：概览
    def describe(self, file: str | Path, sheet: str | None = None) -> dict:
        path = self.settings.resolve_path(file)
        with WorkbookView(path) as view:
            digest = view.digest(sheet)
            result = {
                "file": str(path),
                "sheets": view.overview(),
                "digest": digest.to_dict(),
                "digest_text": digest.to_prompt(),
            }
        if self.logger:
            self.logger.info("读取结构 file=%s sheet=%s", path.name, result["digest"]["sheet"])
        return result

    # ============================================================ 意图 1：生成
    def propose(
        self,
        file: str | Path,
        request: str,
        *,
        sheet: str | None = None,
        target: str | None = None,
        on_delta: Callable[[str, str], None] | None = None,
    ) -> Proposal:
        path = self.settings.resolve_path(file)
        if not request or not request.strip():
            raise ValueError("请描述你想要的计算，例如“在G列算每科总分”")

        with WorkbookView(path) as view:
            sheet_name = view.resolve_sheet(sheet)
            digest = view.digest(sheet_name)
            if self.logger:
                self.logger.info(
                    "生成公式 file=%s sheet=%s target=%s 需求=%s",
                    path.name, sheet_name, target or "自动", request.strip()[:80],
                )
            proposal = Proposal(ok=False, file=path, sheet=sheet_name)
            messages = build_generate_messages(
                digest.to_prompt(), request, target=target, sheet_names=view.sheet_names
            )

            for round_index in range(self.settings.max_repair_rounds + 1):
                # 传了回调才走流式；不传时保持一次性返回，兼容程序化调用与测试替身
                chat_kwargs: dict = {"on_delta": on_delta} if on_delta is not None else {}
                if self.settings.thinking == "auto":
                    # auto 档：简单需求跳过思考提速；首轮失败后自动升级为思考再修
                    chat_kwargs["thinking"] = self._auto_thinking(request, digest, round_index)
                try:
                    result = self.client.chat(messages, **chat_kwargs)
                except LLMError as exc:
                    if exc.usage:  # 输出被截断等失败仍占用了 Token，不能漏算
                        proposal.usage.add(exc.usage)
                    proposal.error = str(exc)
                    if self.logger:
                        self.logger.error("模型调用失败：%s", exc)
                    return proposal
                proposal.usage.add(result.usage)

                try:
                    payload = parse_json_payload(result.content)
                except LLMError as exc:
                    proposal.error = str(exc)
                    return proposal

                candidates, assumptions, clarification = extract_candidates(payload)
                proposal.assumptions = assumptions
                if clarification and not candidates:
                    proposal.clarification = clarification
                    if self.logger:
                        self.logger.info("模型请求补充信息：%s", clarification)
                    return proposal
                if not candidates:
                    proposal.error = "模型没有给出公式"
                    return proposal

                # 模型可能一次给出多个目标单元格（例如“E2 F2 G2 分别求最小值”），
                # 逐个校验，全部通过才算成功；任何一个失败就把所有错误一起回传
                accepted: list[CellWrite] = []
                errors: list[str] = []
                validation: ValidationResult | None = None
                seen: set[str] = set()
                labelled = len(candidates) > 1

                for index, candidate in enumerate(candidates):
                    formula = candidate["formula"]
                    cell = candidate["target"] or (target if index == 0 else None)
                    if not cell:
                        proposal.clarification = "请告诉我把内容写到哪个单元格，例如 G2。"
                        return proposal
                    cell = cell.upper()
                    if cell in seen:
                        errors.append(f"{cell}: 同一单元格被指定了多个写入")
                        continue
                    seen.add(cell)

                    if not formula:
                        # 值写入：分类名、标题等已知常量，无需公式校验
                        try:
                            accepted.append(
                                self._build_value_write(
                                    view, sheet_name, cell,
                                    candidate.get("value"), candidate["explanation"],
                                )
                            )
                        except ValueError as exc:
                            prefix = f"{cell}: " if labelled else ""
                            errors.append(prefix + str(exc))
                        continue

                    one, cell_errors = self._check_candidate(
                        view, sheet_name, digest, cell, formula, candidate.get("fill_to")
                    )
                    validation = _merge_validation(validation, one)
                    if cell_errors:
                        prefix = f"{cell}: " if labelled else ""
                        errors.extend(prefix + message for message in cell_errors)
                        continue

                    writes = self._build_writes(
                        view, sheet_name, cell, formula, candidate.get("fill_to")
                    )
                    writes[0].explanation = candidate["explanation"]
                    writes[0].predicted, writes[0].predicted_note = self._predict(
                        view, sheet_name, formula
                    )
                    accepted.extend(writes)

                formulas = [c["formula"] for c in candidates]
                proposal.attempts.append(
                    Attempt(
                        round=round_index + 1,
                        formula="; ".join(formulas),
                        ok=not errors,
                        errors=list(errors),
                    )
                )
                if errors:
                    if self.logger:
                        self.logger.warning(
                            "第 %d 次校验失败 formulas=%s errors=%s",
                            round_index + 1, formulas, errors,
                        )
                    # 命中“重试也修不好”的错误类别：直接收手，不再白烧一轮 Token
                    unrepairable = validation is not None and validation.unrepairable
                    if unrepairable or round_index >= self.settings.max_repair_rounds:
                        proposal.validation = validation
                        proposal.error = "；".join(errors)
                        if unrepairable:
                            proposal.error += "。该写法本地不支持，重试无法解决，已停止"
                        return proposal
                    messages = messages + [
                        {"role": "assistant", "content": result.content},
                        build_repair_message(formulas, errors),
                    ]
                    continue

                proposal.ok = True
                proposal.writes = accepted
                proposal.explanation = accepted[0].explanation
                proposal.validation = validation
                proposal.predicted = accepted[0].predicted
                proposal.predicted_note = accepted[0].predicted_note
                if self.logger:
                    self.logger.info(
                        "公式通过校验 共 %d 个单元格 targets=%s 预期值=%s",
                        len(accepted),
                        ", ".join(w.cell for w in accepted[:5]),
                        accepted[0].predicted or accepted[0].predicted_note,
                    )
                return proposal
        return proposal

    # ============================================================ 意图 2：校验
    def validate(
        self,
        file: str | Path,
        formula: str,
        *,
        sheet: str | None = None,
        target: str | None = None,
    ) -> dict:
        path = self.settings.resolve_path(file)
        with WorkbookView(path) as view:
            sheet_name = view.resolve_sheet(sheet)
            digest = view.digest(sheet_name)
            validation, _ = validate_formula(
                formula,
                target=target,
                sheet=sheet_name,
                digest=digest,
                sheet_names=view.sheet_names,
            )
            predicted, note = (
                self._predict(view, sheet_name, formula)
                if validation.ok
                else (None, "公式未通过校验，已跳过本地计算")
            )
        if self.logger:
            self.logger.info(
                "校验公式 formula=%s ok=%s errors=%s", formula, validation.ok, validation.errors
            )
        return {
            "file": str(path),
            "sheet": sheet_name,
            "formula": formula,
            "target": target,
            "validation": validation.to_dict(),
            "predicted_value": predicted,
            "predicted_note": note,
        }

    # ============================================================ 意图 3：解释
    def explain(
        self,
        formula: str | None = None,
        *,
        file: str | Path | None = None,
        sheet: str | None = None,
        cell: str | None = None,
        on_delta: Callable[[str, str], None] | None = None,
    ) -> dict:
        digest_text = None
        resolved_formula = (formula or "").strip()
        path = None

        if file:
            path = self.settings.resolve_path(file)
            with WorkbookView(path) as view:
                sheet_name = view.resolve_sheet(sheet)
                digest_text = view.digest(sheet_name).to_prompt()
                if cell and not resolved_formula:
                    found = view.cell_formula(sheet_name, cell.upper())
                    if not found:
                        raise ValueError(f"{sheet_name}!{cell.upper()} 中没有公式")
                    resolved_formula = found
        if not resolved_formula:
            raise ValueError("请提供要解释的公式，或指定包含公式的单元格")

        # 先本地校验一次，静态问题不必花 Token 就能告知用户
        validation, _ = validate_formula(resolved_formula)
        chat_kwargs = {"on_delta": on_delta} if on_delta is not None else {}
        if self.settings.thinking == "auto":
            # 短公式直译即可；嵌套较深的公式才值得开思考
            chat_kwargs["thinking"] = "enabled" if len(resolved_formula) >= 60 else "disabled"
        result = self.client.chat(
            build_explain_messages(resolved_formula, digest_text),
            json_mode=False,
            max_tokens=1200,
            **chat_kwargs,
        )
        if self.logger:
            self.logger.info("解释公式 formula=%s", resolved_formula)
        return {
            "file": str(path) if path else None,
            "formula": resolved_formula,
            "explanation": result.content.strip(),
            "validation": validation.to_dict(),
            "usage": result.usage.to_dict(),
        }

    # ============================================================ 写入
    def apply(self, proposal: Proposal, *, output: str | Path | None = None) -> dict:
        if not proposal.ok or not proposal.writes:
            raise ValueError("方案未通过校验，拒绝写入")
        out_path = self.settings.resolve_path(output, must_exist=False) if output else None
        return apply_writes(
            proposal.file, proposal.writes, self.settings, logger=self.logger, output=out_path
        )

    def write_formula(
        self,
        file: str | Path,
        target: str,
        formula: str,
        *,
        sheet: str | None = None,
        fill_to: str | None = None,
        output: str | Path | None = None,
    ) -> dict:
        """直接写入指定公式（仍然先做本地校验），供"我自己写好了公式"的场景使用。"""
        path = self.settings.resolve_path(file)
        with WorkbookView(path) as view:
            sheet_name = view.resolve_sheet(sheet)
            digest = view.digest(sheet_name)
            validation, errors = self._check_candidate(
                view, sheet_name, digest, target, formula, fill_to
            )
            if errors:
                raise ValueError("公式未通过校验：" + "；".join(errors))
            writes = self._build_writes(view, sheet_name, target, formula, fill_to)
        out_path = self.settings.resolve_path(output, must_exist=False) if output else None
        return apply_writes(path, writes, self.settings, logger=self.logger, output=out_path)

    # ============================================================ 内部方法
    def _auto_thinking(self, request: str, digest: SheetDigest, round_index: int) -> str:
        """EXCELCR_THINKING=auto 的复杂度启发式：简单直算跳过思考省时间，复杂需求保留思考。

        硬信号优先：修复轮次（上一轮没过本地校验，说明问题不简单）、表格规模
        （大表定位数据区更容易出错）；需求侧看长度与复杂关键词。
        """
        if round_index > 0:
            return "enabled"
        if len(request) >= 24:
            return "enabled"
        if digest.max_row > 40 or digest.max_column > 12:
            return "enabled"
        if any(word in request for word in _COMPLEX_SIGNALS):
            return "enabled"
        return "disabled"

    def _check_candidate(
        self,
        view: WorkbookView,
        sheet_name: str,
        digest: SheetDigest,
        target: str,
        formula: str,
        fill_to: str | None,
    ) -> tuple[ValidationResult, list[str]]:
        """校验首个单元格；若需填充，额外校验最后一个单元格平移后的公式。"""
        validation, _ = validate_formula(
            formula,
            target=target,
            sheet=sheet_name,
            digest=digest,
            sheet_names=view.sheet_names,
        )
        errors = list(validation.errors)
        if not errors and fill_to:
            try:
                writes = expand_fill(sheet_name, target, formula, fill_to)
            except (ValueError, FormulaSyntaxError) as exc:
                errors.append(str(exc))
                return validation, errors
            tail = writes[-1]
            tail_validation, _ = validate_formula(
                tail.formula,
                target=tail.cell,
                sheet=sheet_name,
                digest=digest,
                sheet_names=view.sheet_names,
            )
            for message in tail_validation.errors:
                errors.append(f"填充到 {tail.cell} 时: {message}")
            validation.warnings.extend(
                f"填充到 {tail.cell} 时: {w}" for w in tail_validation.warnings
            )
        return validation, errors

    def _mark_old_content(self, view: WorkbookView, sheet_name: str, item: CellWrite) -> None:
        """记录目标单元格当前内容，供预览里的覆盖提醒使用。"""
        _, col, row = parse_target(item.cell)
        raw = view.formula_sheet(sheet_name).cell(row=row, column=col).value
        if isinstance(raw, str) and raw.startswith("="):
            item.old_formula = raw
        else:
            item.old_value = view.value_sheet(sheet_name).cell(row=row, column=col).value

    def _build_writes(
        self,
        view: WorkbookView,
        sheet_name: str,
        target: str,
        formula: str,
        fill_to: str | None,
    ) -> list[CellWrite]:
        writes = expand_fill(sheet_name, target, formula, fill_to)
        for item in writes:
            self._mark_old_content(view, sheet_name, item)
        return writes

    def _build_value_write(
        self,
        view: WorkbookView,
        sheet_name: str,
        cell: str,
        value: object,
        explanation: str,
    ) -> CellWrite:
        """值写入：把已知常量（区域名、标题等）直接写入单元格，不经过公式校验。"""
        if isinstance(value, bool) or not isinstance(value, (str, int, float)):
            raise ValueError("value 只能是文本或数字")
        if isinstance(value, str):
            if not value.strip():
                raise ValueError("常量值不能为空")
            if value.startswith("="):
                raise ValueError("常量值不能以 = 开头，公式请使用 formula 字段")
            if len(value) > _MAX_VALUE_LENGTH:
                raise ValueError(f"常量值超过 {_MAX_VALUE_LENGTH} 字，疑似把说明文字写进了单元格")
        write = CellWrite(sheet=sheet_name, cell=cell, formula="", value=value)
        write.explanation = explanation
        write.predicted = str(value)
        self._mark_old_content(view, sheet_name, write)
        return write

    def _predict(
        self, view: WorkbookView, sheet_name: str, formula: str
    ) -> tuple[str | None, str | None]:
        """可选功能：在 Python 中独立算一遍公式，作为写入前的预期结果。"""
        validation, ast = validate_formula(formula, sheet=sheet_name)
        if ast is None:
            return None, "公式无法解析"
        sheets = {name: view.wb_values[name] for name in view.sheet_names}
        try:
            value = evaluate_formula(ast, sheets, sheet_name)
        except UnsupportedFormula as exc:
            return None, str(exc)
        except (ZeroDivisionError, ValueError, TypeError, IndexError, KeyError) as exc:
            return None, f"本地计算出错：{exc}"
        return format_value(value), None


def _merge_validation(
    base: ValidationResult | None, extra: ValidationResult
) -> ValidationResult:
    """多个候选公式共用一份校验结果：警告累加，ok 取与。"""
    if base is None:
        return extra
    base.ok = base.ok and extra.ok
    base.errors.extend(extra.errors)
    base.warnings.extend(extra.warnings)
    base.unrepairable = base.unrepairable or extra.unrepairable
    base.functions.extend(f for f in extra.functions if f not in base.functions)
    base.refs.extend(r for r in extra.refs if r not in base.refs)
    return base
