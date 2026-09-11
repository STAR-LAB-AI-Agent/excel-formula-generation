"""公式独立求值器：在 Python 中重算一遍公式，用于写入前的"公式自动验证"。

只支持常用函数子集；遇到不支持的函数会抛出 UnsupportedFormula，
上层据此把结果标记为"未验证"，而不是阻断流程。
"""
from __future__ import annotations

import datetime as _dt
import math
import re
import statistics
from dataclasses import dataclass

from openpyxl.worksheet.worksheet import Worksheet

from .formula_parser import (
    Binary,
    Bool,
    DefinedName,
    ErrorLiteral,
    FuncCall,
    Number,
    Ref,
    Text,
    Unary,
)


class UnsupportedFormula(RuntimeError):
    """本地求值器不支持该公式（函数未实现或引用无缓存值）。"""


@dataclass(frozen=True)
class ExcelError:
    code: str

    def __str__(self) -> str:
        return self.code


DIV0 = ExcelError("#DIV/0!")
VALUE_ERR = ExcelError("#VALUE!")
NA_ERR = ExcelError("#N/A")
NUM_ERR = ExcelError("#NUM!")


class Evaluator:
    """按 AST 求值。所有取值来自 data_only 工作簿的缓存值。"""

    def __init__(self, sheets: dict[str, Worksheet], default_sheet: str):
        self.sheets = sheets
        self.default_sheet = default_sheet

    # -------------------------------------------------------------- 对外入口
    def evaluate(self, node: object):
        value = self._eval(node)
        if isinstance(value, list):  # 顶层拿到区域时取首个值，与 Excel 隐式交叉行为近似
            flat = _flatten(value)
            return flat[0] if flat else None
        return value

    # -------------------------------------------------------------- 递归求值
    def _eval(self, node: object):
        if isinstance(node, Number):
            return node.value
        if isinstance(node, Text):
            return node.value
        if isinstance(node, Bool):
            return node.value
        if isinstance(node, ErrorLiteral):
            return ExcelError(node.value)
        if isinstance(node, Ref):
            return self._eval_ref(node)
        if isinstance(node, Unary):
            return self._eval_unary(node)
        if isinstance(node, Binary):
            return self._eval_binary(node)
        if isinstance(node, FuncCall):
            return self._eval_func(node)
        if isinstance(node, DefinedName):
            raise UnsupportedFormula(f"不支持定义名称 {node.name}")
        raise UnsupportedFormula(f"不支持的表达式节点 {type(node).__name__}")

    def _eval_ref(self, ref: Ref):
        sheet_name = ref.sheet or self.default_sheet
        ws = self.sheets.get(sheet_name)
        if ws is None:
            raise UnsupportedFormula(f"找不到工作表 {sheet_name}")
        if ref.whole_column:
            row2 = max(ws.max_row or 1, 1)
        else:
            row2 = ref.row2
        if not ref.is_range:
            return _coerce(ws.cell(row=ref.row1, column=ref.col1).value)
        values = []
        for row in range(ref.row1, row2 + 1):
            values.append(
                [_coerce(ws.cell(row=row, column=col).value) for col in range(ref.col1, ref.col2 + 1)]
            )
        return values

    def _eval_unary(self, node: Unary):
        value = self._eval(node.operand)
        if isinstance(value, ExcelError):
            return value
        if node.op == "%":
            return _to_number(value) / 100
        number = _to_number(value)
        if isinstance(number, ExcelError):
            return number
        return -number if node.op == "-" else number

    def _eval_binary(self, node: Binary):
        left, right = self._eval(node.left), self._eval(node.right)
        for value in (left, right):
            if isinstance(value, ExcelError):
                return value
        left, right = _single(left), _single(right)
        op = node.op

        if op == "&":
            return _to_text(left) + _to_text(right)
        if op in {"=", "<>", "<", ">", "<=", ">="}:
            return _compare(op, left, right)

        a, b = _to_number(left), _to_number(right)
        if isinstance(a, ExcelError):
            return a
        if isinstance(b, ExcelError):
            return b
        if op == "+":
            return a + b
        if op == "-":
            return a - b
        if op == "*":
            return a * b
        if op == "/":
            return DIV0 if b == 0 else a / b
        if op == "^":
            try:
                return math.pow(a, b)
            except (ValueError, OverflowError):
                return NUM_ERR
        raise UnsupportedFormula(f"不支持的运算符 {op}")

    # -------------------------------------------------------------- 函数实现
    def _eval_func(self, node: FuncCall):
        name = node.name
        handler = _FUNCTIONS.get(name)
        if handler is None:
            raise UnsupportedFormula(f"本地验证暂不支持函数 {name}")
        # IF / IFERROR 需要延迟求值分支，单独处理
        if name == "IF":
            condition = _single(self._eval(node.args[0]))
            if isinstance(condition, ExcelError):
                return condition
            if _to_bool(condition):
                return _single(self._eval(node.args[1]))
            if len(node.args) > 2:
                return _single(self._eval(node.args[2]))
            return False
        if name in {"IFERROR", "IFNA"}:
            try:
                value = _single(self._eval(node.args[0]))
            except UnsupportedFormula:
                raise
            except Exception:  # noqa: BLE001 - 求值异常等价于 Excel 报错
                return _single(self._eval(node.args[1]))
            if isinstance(value, ExcelError):
                return _single(self._eval(node.args[1]))
            return value
        args = [self._eval(arg) for arg in node.args]
        for arg in args:
            if isinstance(arg, ExcelError) and name not in {"ISERROR", "ISERR", "ISNA"}:
                return arg
        return handler(args)


# ------------------------------------------------------------------ 值处理工具
def _coerce(value: object):
    if isinstance(value, str) and value.startswith("="):
        # data_only 视图里仍是公式，说明该单元格没有缓存值
        raise UnsupportedFormula("引用的单元格是尚未计算的公式，无法本地验证")
    if isinstance(value, (_dt.datetime, _dt.date, _dt.time)):
        raise UnsupportedFormula("引用包含日期类型，本地验证暂不支持日期运算")
    return value


def _flatten(value) -> list:
    if isinstance(value, list):
        out: list = []
        for item in value:
            out.extend(_flatten(item))
        return out
    return [value]


def _single(value):
    if isinstance(value, list):
        flat = _flatten(value)
        return flat[0] if flat else None
    return value


def _numbers(values) -> list[float]:
    out: list[float] = []
    for item in _flatten(values):
        if isinstance(item, bool) or item is None or isinstance(item, ExcelError):
            continue
        if isinstance(item, (int, float)):
            out.append(float(item))
    return out


def _to_number(value):
    if value is None:
        return 0.0
    if isinstance(value, bool):
        return 1.0 if value else 0.0
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, ExcelError):
        return value
    try:
        return float(str(value).strip())
    except ValueError:
        return VALUE_ERR


def _to_text(value) -> str:
    if value is None:
        return ""
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value)


def _to_bool(value) -> bool:
    if value is None:
        return False
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)):
        return value != 0
    return str(value).strip().upper() == "TRUE"


def _compare(op: str, left, right):
    if isinstance(left, str) or isinstance(right, str):
        a, b = _to_text(left).upper(), _to_text(right).upper()
    else:
        a, b = _to_number(left), _to_number(right)
        if isinstance(a, ExcelError) or isinstance(b, ExcelError):
            return VALUE_ERR
    if op == "=":
        return a == b
    if op == "<>":
        return a != b
    if op == "<":
        return a < b
    if op == ">":
        return a > b
    if op == "<=":
        return a <= b
    return a >= b


_CRITERIA_RE = re.compile(r"^\s*(<=|>=|<>|=|<|>)?\s*(.*)$", re.DOTALL)


def _match_criteria(value, criteria) -> bool:
    """实现 COUNTIF/SUMIF 的条件匹配，如 ">80"、"Math"、"<>"。"""
    if isinstance(criteria, (int, float)) and not isinstance(criteria, bool):
        return _compare("=", value, criteria) is True
    match = _CRITERIA_RE.match(_to_text(criteria))
    op = match.group(1) or "="
    operand_text = match.group(2)
    try:
        operand: object = float(operand_text)
    except ValueError:
        operand = operand_text
    if isinstance(operand, str) and ("*" in operand or "?" in operand):
        raise UnsupportedFormula("本地验证不支持通配符条件")
    if value is None and op in {"=", "<>"}:
        return (operand_text == "") if op == "=" else (operand_text != "")
    result = _compare(op, value, operand)
    return result is True


def _sum_if(args, *, mode: str):
    """SUMIF / AVERAGEIF / COUNTIF 的统一实现。"""
    if mode == "COUNTIF":
        rng, criteria = args[0], _single(args[1])
        picked = [v for v in _flatten(rng) if _match_criteria(v, criteria)]
        return float(len(picked))
    rng, criteria = args[0], _single(args[1])
    target = args[2] if len(args) > 2 else rng
    source = _flatten(rng)
    values = _flatten(target)
    if len(values) < len(source):
        raise UnsupportedFormula("条件区域与求和区域大小不一致")
    picked = [
        values[i]
        for i, item in enumerate(source)
        if _match_criteria(item, criteria) and isinstance(values[i], (int, float))
        and not isinstance(values[i], bool)
    ]
    numbers = [float(v) for v in picked]
    if mode == "SUMIF":
        return math.fsum(numbers)
    return math.fsum(numbers) / len(numbers) if numbers else DIV0


def _multi_criteria(args, *, mode: str):
    """COUNTIFS / SUMIFS / AVERAGEIFS。"""
    if mode == "COUNTIFS":
        pairs = list(zip(args[0::2], args[1::2]))
        sum_range = None
    else:
        sum_range = _flatten(args[0])
        pairs = list(zip(args[1::2], args[2::2]))
    if not pairs:
        raise UnsupportedFormula("条件参数不成对")
    length = len(_flatten(pairs[0][0]))
    hits = []
    for index in range(length):
        ok = True
        for rng, criteria in pairs:
            flat = _flatten(rng)
            if len(flat) != length:
                raise UnsupportedFormula("多条件区域大小不一致")
            if not _match_criteria(flat[index], _single(criteria)):
                ok = False
                break
        if ok:
            hits.append(index)
    if mode == "COUNTIFS":
        return float(len(hits))
    numbers = [
        float(sum_range[i])
        for i in hits
        if i < len(sum_range) and isinstance(sum_range[i], (int, float)) and not isinstance(sum_range[i], bool)
    ]
    if mode == "SUMIFS":
        return math.fsum(numbers)
    return math.fsum(numbers) / len(numbers) if numbers else DIV0


def _safe_div(numbers: list[float], func):
    return func(numbers) if numbers else DIV0


def _round(args, mode: str):
    value = _to_number(_single(args[0]))
    if isinstance(value, ExcelError):
        return value
    digits = int(_to_number(_single(args[1]))) if len(args) > 1 else 0
    factor = 10 ** digits
    if mode == "ROUND":
        return math.floor(abs(value) * factor + 0.5) / factor * (1 if value >= 0 else -1)
    if mode == "UP":
        return math.ceil(abs(value) * factor) / factor * (1 if value >= 0 else -1)
    return math.floor(abs(value) * factor) / factor * (1 if value >= 0 else -1)


def _rank(args):
    value = _to_number(_single(args[0]))
    numbers = _numbers(args[1])
    order = int(_to_number(_single(args[2]))) if len(args) > 2 else 0
    if not numbers or isinstance(value, ExcelError):
        return NA_ERR
    ordered = sorted(numbers) if order else sorted(numbers, reverse=True)
    return float(ordered.index(value) + 1) if value in ordered else NA_ERR


def _vlookup(args):
    key = _single(args[0])
    table = args[1]
    if not isinstance(table, list):
        raise UnsupportedFormula("VLOOKUP 第二参数必须是区域")
    col_index = int(_to_number(_single(args[2])))
    exact = len(args) > 3 and not _to_bool(_single(args[3]))
    if not exact:
        raise UnsupportedFormula("本地验证仅支持精确匹配的 VLOOKUP（第 4 参数为 FALSE）")
    for row in table:
        cells = row if isinstance(row, list) else [row]
        if cells and _compare("=", cells[0], key) is True:
            if col_index > len(cells):
                return ExcelError("#REF!")
            return cells[col_index - 1]
    return NA_ERR


def _text_func(args, func):
    return func(_to_text(_single(args[0])))


_FUNCTIONS: dict[str, object] = {
    "SUM": lambda a: math.fsum(_numbers(a)),
    "PRODUCT": lambda a: math.prod(_numbers(a)) if _numbers(a) else 0.0,
    "AVERAGE": lambda a: _safe_div(_numbers(a), lambda n: math.fsum(n) / len(n)),
    "MEDIAN": lambda a: _safe_div(_numbers(a), statistics.median),
    "MAX": lambda a: max(_numbers(a)) if _numbers(a) else 0.0,
    "MIN": lambda a: min(_numbers(a)) if _numbers(a) else 0.0,
    "COUNT": lambda a: float(len(_numbers(a))),
    "COUNTA": lambda a: float(sum(1 for v in _flatten(a) if v is not None and v != "")),
    "COUNTBLANK": lambda a: float(sum(1 for v in _flatten(a) if v is None or v == "")),
    "STDEV": lambda a: _safe_div(_numbers(a), lambda n: statistics.stdev(n) if len(n) > 1 else DIV0),
    "STDEV.S": lambda a: _safe_div(_numbers(a), lambda n: statistics.stdev(n) if len(n) > 1 else DIV0),
    "STDEV.P": lambda a: _safe_div(_numbers(a), statistics.pstdev),
    "COUNTIF": lambda a: _sum_if(a, mode="COUNTIF"),
    "SUMIF": lambda a: _sum_if(a, mode="SUMIF"),
    "AVERAGEIF": lambda a: _sum_if(a, mode="AVERAGEIF"),
    "COUNTIFS": lambda a: _multi_criteria(a, mode="COUNTIFS"),
    "SUMIFS": lambda a: _multi_criteria(a, mode="SUMIFS"),
    "AVERAGEIFS": lambda a: _multi_criteria(a, mode="AVERAGEIFS"),
    "IF": lambda a: None,  # 由 Evaluator 特殊处理
    "IFERROR": lambda a: None,
    "IFNA": lambda a: None,
    "AND": lambda a: all(_to_bool(v) for v in _flatten(a) if v is not None),
    "OR": lambda a: any(_to_bool(v) for v in _flatten(a) if v is not None),
    "NOT": lambda a: not _to_bool(_single(a[0])),
    "ABS": lambda a: abs(_to_number(_single(a[0]))),
    "SQRT": lambda a: (
        math.sqrt(_to_number(_single(a[0]))) if _to_number(_single(a[0])) >= 0 else NUM_ERR
    ),
    "POWER": lambda a: math.pow(_to_number(_single(a[0])), _to_number(_single(a[1]))),
    "MOD": lambda a: (
        DIV0 if _to_number(_single(a[1])) == 0
        else math.fmod(_to_number(_single(a[0])), _to_number(_single(a[1])))
    ),
    "INT": lambda a: float(math.floor(_to_number(_single(a[0])))),
    "TRUNC": lambda a: float(math.trunc(_to_number(_single(a[0])))),
    "SIGN": lambda a: float((_to_number(_single(a[0])) > 0) - (_to_number(_single(a[0])) < 0)),
    "ROUND": lambda a: _round(a, "ROUND"),
    "ROUNDUP": lambda a: _round(a, "UP"),
    "ROUNDDOWN": lambda a: _round(a, "DOWN"),
    "LARGE": lambda a: (
        sorted(_numbers(a[0]), reverse=True)[int(_to_number(_single(a[1]))) - 1]
        if 0 < int(_to_number(_single(a[1]))) <= len(_numbers(a[0])) else NUM_ERR
    ),
    "SMALL": lambda a: (
        sorted(_numbers(a[0]))[int(_to_number(_single(a[1]))) - 1]
        if 0 < int(_to_number(_single(a[1]))) <= len(_numbers(a[0])) else NUM_ERR
    ),
    "RANK": _rank,
    "RANK.EQ": _rank,
    "ROWS": lambda a: float(len(a[0]) if isinstance(a[0], list) else 1),
    "COLUMNS": lambda a: float(
        len(a[0][0]) if isinstance(a[0], list) and a[0] and isinstance(a[0][0], list) else 1
    ),
    "LEN": lambda a: float(len(_to_text(_single(a[0])))),
    "TRIM": lambda a: _text_func(a, lambda s: s.strip()),
    "UPPER": lambda a: _text_func(a, str.upper),
    "LOWER": lambda a: _text_func(a, str.lower),
    "PROPER": lambda a: _text_func(a, str.title),
    "LEFT": lambda a: _to_text(_single(a[0]))[: int(_to_number(_single(a[1]))) if len(a) > 1 else 1],
    "RIGHT": lambda a: _to_text(_single(a[0]))[
        -(int(_to_number(_single(a[1]))) if len(a) > 1 else 1):
    ],
    "MID": lambda a: _to_text(_single(a[0]))[
        int(_to_number(_single(a[1]))) - 1: int(_to_number(_single(a[1]))) - 1 + int(_to_number(_single(a[2])))
    ],
    "CONCAT": lambda a: "".join(_to_text(v) for v in _flatten(a)),
    "CONCATENATE": lambda a: "".join(_to_text(v) for v in _flatten(a)),
    "TEXTJOIN": lambda a: _to_text(_single(a[0])).join(
        _to_text(v) for v in _flatten(a[2:]) if not (_to_bool(_single(a[1])) and (v is None or v == ""))
    ),
    "ISBLANK": lambda a: _single(a[0]) is None,
    "ISNUMBER": lambda a: isinstance(_single(a[0]), (int, float)) and not isinstance(_single(a[0]), bool),
    "ISTEXT": lambda a: isinstance(_single(a[0]), str),
    "ISERROR": lambda a: isinstance(_single(a[0]), ExcelError),
    "ISERR": lambda a: isinstance(_single(a[0]), ExcelError),
    "ISNA": lambda a: _single(a[0]) == NA_ERR,
    "VLOOKUP": _vlookup,
    "SUMPRODUCT": lambda a: math.fsum(
        math.prod(vals) for vals in zip(*[[float(x) for x in _numbers(arg)] for arg in a])
    ),
}


def evaluate_formula(ast: object, sheets: dict[str, Worksheet], default_sheet: str):
    """对外接口：求值成功返回结果值，不支持时抛出 UnsupportedFormula。"""
    return Evaluator(sheets, default_sheet).evaluate(ast)


def format_value(value) -> str:
    if value is None:
        return "(空)"
    if isinstance(value, ExcelError):
        return value.code
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, float):
        if value.is_integer():
            return str(int(value))
        return f"{value:.6g}"
    return str(value)
