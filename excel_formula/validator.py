"""公式静态校验：语法、函数白名单、参数个数、引用范围、循环引用。

校验完全在本地完成，不消耗 Token；失败信息会被回传给模型用于自动修复。
"""
from __future__ import annotations

from dataclasses import dataclass, field

from openpyxl.utils import get_column_letter

from .config import (
    ALLOWED_FUNCTIONS,
    FORBIDDEN_FUNCTIONS,
    MAX_EXCEL_COLUMN,
    MAX_EXCEL_ROW,
)
from .excel_reader import SheetDigest
from .formula_parser import (
    DefinedName,
    ErrorLiteral,
    FormulaSyntaxError,
    FuncCall,
    Ref,
    parse_formula,
    parse_target,
    walk,
)

# 常用函数的参数个数约束：函数名 -> (最少, 最多; None 表示不限)
ARITY: dict[str, tuple[int, int | None]] = {
    "SUM": (1, None), "AVERAGE": (1, None), "MAX": (1, None), "MIN": (1, None),
    "COUNT": (1, None), "COUNTA": (1, None), "COUNTBLANK": (1, 1), "PRODUCT": (1, None),
    "MEDIAN": (1, None), "STDEV": (1, None), "STDEV.S": (1, None), "STDEV.P": (1, None),
    "COUNTIF": (2, 2), "SUMIF": (2, 3), "AVERAGEIF": (2, 3),
    "COUNTIFS": (2, None), "SUMIFS": (3, None), "AVERAGEIFS": (3, None),
    "IF": (2, 3), "IFERROR": (2, 2), "IFNA": (2, 2), "AND": (1, None), "OR": (1, None),
    "NOT": (1, 1), "XOR": (1, None), "SWITCH": (3, None), "IFS": (2, None),
    "ROUND": (1, 2), "ROUNDUP": (2, 2), "ROUNDDOWN": (2, 2), "INT": (1, 1), "TRUNC": (1, 2),
    "ABS": (1, 1), "SQRT": (1, 1), "POWER": (2, 2), "MOD": (2, 2), "SIGN": (1, 1),
    "EXP": (1, 1), "LN": (1, 1), "LOG": (1, 2), "LOG10": (1, 1),
    "LARGE": (2, 2), "SMALL": (2, 2), "RANK": (2, 3), "RANK.EQ": (2, 3),
    "VLOOKUP": (3, 4), "HLOOKUP": (3, 4), "XLOOKUP": (3, 6), "INDEX": (2, 3),
    "MATCH": (2, 3), "CHOOSE": (2, None), "OFFSET": (3, 5),
    "LEN": (1, 1), "LEFT": (1, 2), "RIGHT": (1, 2), "MID": (3, 3),
    "FIND": (2, 3), "SEARCH": (2, 3), "SUBSTITUTE": (3, 4), "REPLACE": (4, 4),
    "TRIM": (1, 1), "UPPER": (1, 1), "LOWER": (1, 1), "PROPER": (1, 1),
    "TEXT": (2, 2), "VALUE": (1, 1), "TEXTJOIN": (3, None), "CONCAT": (1, None),
    "CONCATENATE": (1, None), "REPT": (2, 2),
    "TODAY": (0, 0), "NOW": (0, 0), "DATE": (3, 3), "YEAR": (1, 1), "MONTH": (1, 1),
    "DAY": (1, 1), "DATEDIF": (3, 3), "EOMONTH": (2, 2), "EDATE": (2, 2),
    "ISBLANK": (1, 1), "ISNUMBER": (1, 1), "ISTEXT": (1, 1), "ISERROR": (1, 1),
    "ROW": (0, 1), "COLUMN": (0, 1), "ROWS": (1, 1), "COLUMNS": (1, 1),
    "SUBTOTAL": (2, None), "SUMPRODUCT": (1, None), "RANDBETWEEN": (2, 2), "RAND": (0, 0),
}

# 每次重算结果都会变化的函数：能用但要提醒用户
VOLATILE_FUNCTIONS = {"RAND", "RANDBETWEEN", "NOW", "TODAY", "OFFSET"}


@dataclass
class ValidationResult:
    ok: bool
    errors: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    functions: list[str] = field(default_factory=list)
    refs: list[str] = field(default_factory=list)

    def error_text(self) -> str:
        return "；".join(self.errors)

    def to_dict(self) -> dict:
        return {
            "ok": self.ok,
            "errors": self.errors,
            "warnings": self.warnings,
            "functions": self.functions,
            "refs": self.refs,
        }


def validate_formula(
    formula: str,
    *,
    target: str | None = None,
    sheet: str | None = None,
    digest: SheetDigest | None = None,
    sheet_names: list[str] | None = None,
) -> tuple[ValidationResult, object | None]:
    """校验公式，返回 (校验结果, AST)。语法错误时 AST 为 None。"""
    errors: list[str] = []
    warnings: list[str] = []

    try:
        ast = parse_formula(formula)
    except FormulaSyntaxError as exc:
        return ValidationResult(ok=False, errors=[f"语法错误: {exc}"]), None

    nodes = list(walk(ast))
    functions = sorted({n.name for n in nodes if isinstance(n, FuncCall)})
    refs = [n for n in nodes if isinstance(n, Ref)]

    _check_functions(nodes, errors, warnings)
    _check_identifiers(nodes, errors)
    _check_sheets(refs, sheet_names, errors)
    _check_ranges(refs, digest, sheet, warnings, errors)
    _check_target(target, sheet, refs, digest, errors, warnings)

    return (
        ValidationResult(
            ok=not errors,
            errors=errors,
            warnings=warnings,
            functions=functions,
            refs=sorted({r.normalized() for r in refs}),
        ),
        ast,
    )


# ------------------------------------------------------------------ 分项检查
def _check_functions(nodes: list, errors: list[str], warnings: list[str]) -> None:
    for node in nodes:
        if isinstance(node, ErrorLiteral):
            errors.append(f"公式中包含错误值字面量 {node.value}")
        if not isinstance(node, FuncCall):
            continue
        name = node.name
        if name in FORBIDDEN_FUNCTIONS:
            errors.append(f"禁止使用高风险函数 {name}（可能引用外部资源或动态求值）")
            continue
        if name not in ALLOWED_FUNCTIONS:
            errors.append(
                f"函数 {name} 不在允许列表中，请改用常见 Excel 函数（如 SUM/AVERAGE/IF/COUNTIF 等）"
            )
            continue
        low, high = ARITY.get(name, (0, None))
        count = len(node.args)
        if count < low or (high is not None and count > high):
            expect = f"{low}" if high == low else f"{low}~{'不限' if high is None else high}"
            errors.append(f"函数 {name} 参数个数为 {count}，期望 {expect} 个")
        if any(arg is None for arg in node.args):
            errors.append(f"函数 {name} 存在空参数")
        if name in VOLATILE_FUNCTIONS:
            warnings.append(f"{name} 属于易变函数，每次重算结果可能变化")


def _check_identifiers(nodes: list, errors: list[str]) -> None:
    for node in nodes:
        if isinstance(node, DefinedName):
            errors.append(
                f"无法识别的标识符 {node.name!r}：请使用单元格地址（如 B2、B2:F2）而不是名称或中文列名"
            )


def _check_sheets(refs: list[Ref], sheet_names: list[str] | None, errors: list[str]) -> None:
    if not sheet_names:
        return
    for ref in refs:
        if ref.sheet and ref.sheet not in sheet_names:
            errors.append(
                f"引用的工作表 {ref.sheet!r} 不存在，可用工作表：{', '.join(sheet_names)}"
            )


def _check_ranges(
    refs: list[Ref],
    digest: SheetDigest | None,
    sheet: str | None,
    warnings: list[str],
    errors: list[str],
) -> None:
    for ref in refs:
        if ref.col2 > MAX_EXCEL_COLUMN or ref.row2 > MAX_EXCEL_ROW:
            errors.append(f"引用 {ref.text} 超出 Excel 地址范围")
        if ref.whole_column:
            warnings.append(f"{ref.text} 为整列引用，建议改为具体区域以提升性能")

    if digest is None or not digest.max_row:
        return
    for ref in refs:
        if ref.sheet and ref.sheet != digest.name:
            continue  # 跨表引用无摘要信息，交给 Excel 自身校验
        if ref.whole_column:
            continue
        if ref.col1 > digest.max_column or ref.row1 > digest.max_row:
            warnings.append(
                f"引用 {ref.normalized()} 完全落在数据区 {digest.data_range} 之外，可能是空白区域"
            )
        elif ref.col2 > digest.max_column or ref.row2 > digest.max_row:
            warnings.append(
                f"引用 {ref.normalized()} 部分超出数据区 {digest.data_range}，超出部分为空"
            )


def _check_target(
    target: str | None,
    sheet: str | None,
    refs: list[Ref],
    digest: SheetDigest | None,
    errors: list[str],
    warnings: list[str],
) -> None:
    if not target:
        return
    try:
        target_sheet, col, row = parse_target(target)
    except FormulaSyntaxError as exc:
        errors.append(f"目标单元格非法: {exc}")
        return

    current_sheet = target_sheet or sheet or (digest.name if digest else None)
    for ref in refs:
        ref_sheet = ref.sheet or current_sheet
        if ref_sheet == current_sheet and ref.contains(col, row):
            errors.append(
                f"循环引用：公式写入 {get_column_letter(col)}{row}，但它自身被 {ref.normalized()} 引用"
            )
            break

    if digest and digest.max_row and not refs:
        warnings.append("公式没有引用任何单元格，请确认这是常量公式")
