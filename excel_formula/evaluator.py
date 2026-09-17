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
    ArrayLiteral,
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


@dataclass(frozen=True)
class DateValue:
    """只读视图中的日期占位值。

    日期在 Excel 里本质是序列号，但本地求值器不实现日期运算。为了不让一整列日期
    把 VLOOKUP/SUMIFS 这类「按行取值、不碰日期列」的公式拖成"未验证"，取值阶段
    只包装不报错；一旦日期真的参与算术、文本或条件判断，就在对应算子里抛
    UnsupportedFormula，宁可标"未验证"也不给出静默算错的结果。
    """

    raw: object

    def __str__(self) -> str:
        return format_value(self)


DATE_HINT = "区域中包含日期值，本地验证暂不支持日期运算"

# 区域直接参与 +-*/ 或比较属于数组公式语义，本地只做标量求值
_ARRAY_IN_BINARY_HINT = (
    "区域与值直接运算属于数组公式语义（本地不支持），"
    "请改用 SUMIF/COUNTIFS/SUMPRODUCT 等聚合函数"
)

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
        if isinstance(node, ArrayLiteral):
            return [[self._eval(item) for item in row] for row in node.rows]
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
        if isinstance(value, list):
            raise UnsupportedFormula(_ARRAY_IN_BINARY_HINT)
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
        # Excel 对区域参与运算走数组公式语义（逐元素展开），本地只做标量求值。
        # 旧实现把区域塌缩成首元素再比较/计算，会给出静默算错的结果（例如
        # (D2:D9="华北") 只比较了 D2），这里改为明确标记"未验证"。
        if isinstance(left, list) or isinstance(right, list):
            raise UnsupportedFormula(_ARRAY_IN_BINARY_HINT)
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
            args = [self._eval(arg) for arg in node.args]
            if any(isinstance(arg, list) for arg in args):
                return _if_array(args)
            condition = args[0]
            if isinstance(condition, ExcelError):
                return condition
            if _to_bool(condition):
                return _single(args[1])
            if len(args) > 2:
                return _single(args[2])
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
        if name == "IFS":
            # 条件-结果成对，命中第一个为真的条件即返回，未命中的分支不求值
            for index in range(0, len(node.args) - 1, 2):
                condition = self._eval(node.args[index])
                if isinstance(condition, ExcelError):
                    return condition
                if _to_bool(_single(condition)):
                    return _single(self._eval(node.args[index + 1]))
            return NA_ERR
        if name == "CHOOSE":
            # 只对选中的分支求值：CHOOSE(2, 坏公式, 好结果) 在 Excel 中不会报错
            index_value = self._eval(node.args[0])
            if isinstance(index_value, ExcelError):
                return index_value
            picks = _single(index_value)
            number = _to_number(picks)
            if isinstance(number, ExcelError):
                return number
            choice = int(number)
            if choice < 1 or choice > len(node.args) - 1:
                return VALUE_ERR
            return _single(self._eval(node.args[choice]))
        if name == "SWITCH":
            # 命中分支才求值：SWITCH(x, 1, "a", 2, 1/0) 不该因未命中的 1/0 而报错
            target = _single(self._eval(node.args[0]))
            if isinstance(target, ExcelError):
                return target
            index = 1
            while index + 1 < len(node.args):
                candidate = self._eval(node.args[index])
                if isinstance(candidate, ExcelError):
                    return candidate
                if _compare("=", _single(candidate), target) is True:
                    return _single(self._eval(node.args[index + 1]))
                index += 2
            if index < len(node.args):  # 最后的落单参数是默认值
                return _single(self._eval(node.args[index]))
            return NA_ERR
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
        # 不在这里报错：VLOOKUP/SUMIFS 等按行取值的函数可能根本不碰这一列日期，
        # 真正参与运算时由 _to_number/_to_text 等算子抛出"未验证"。
        return DateValue(value)
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


def _if_array(args: list):
    """IF 的数组形式，用于 IF({1,0},E2:E9,D2:D9) 这类“反向查找”构造。

    按 Excel 的数组广播规则：先取所有参与值形状的最大值当作结果形状，
    长度不足的维度自动伸展（1x2 的条件配 8x1 的区域得到 8x2）。
    """
    yes = args[1]
    no = args[2] if len(args) > 2 else False
    rows = max(_layer_count(value) for value in args) or 1
    cols = max(_layer_width(value) for value in args) or 1
    return [
        [_if_at(args[0], yes, no, row, col) for col in range(cols)]
        for row in range(rows)
    ]


def _layer_count(value) -> int:
    if not isinstance(value, list):
        return 1
    return len(value) if value and isinstance(value[0], list) else 1


def _layer_width(value) -> int:
    if not isinstance(value, list):
        return 1
    if not value:
        return 0
    return len(value[0]) if isinstance(value[0], list) else len(value)


def _at(value, row: int, col: int = 0):
    """取广播后 (row, col) 处的值：标量不变，一维退化为列向量，越界时钳到首行/首列。"""
    if not isinstance(value, list):
        return value
    if not value:
        return None
    line = value[row] if row < len(value) else value[0]
    if not isinstance(line, list):
        return line
    if not line:
        return None
    return line[col] if col < len(line) else line[0]


def _if_at(condition, yes, no, row: int, col: int):
    chosen = _at(condition, row, col)
    if isinstance(chosen, ExcelError):
        return chosen
    return _at(yes, row, col) if _to_bool(chosen) else _at(no, row, col)


def _numbers(values) -> list[float]:
    out: list[float] = []
    for item in _flatten(values):
        if isinstance(item, DateValue):
            raise UnsupportedFormula(DATE_HINT)
        if isinstance(item, bool) or item is None or isinstance(item, ExcelError):
            continue
        if isinstance(item, (int, float)):
            out.append(float(item))
    return out


def _to_number(value):
    if isinstance(value, DateValue):
        raise UnsupportedFormula(DATE_HINT)
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
    if isinstance(value, DateValue):
        raise UnsupportedFormula(DATE_HINT)
    if value is None:
        return ""
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value)


def _to_bool(value) -> bool:
    if isinstance(value, DateValue):
        raise UnsupportedFormula(DATE_HINT)
    if value is None:
        return False
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)):
        return value != 0
    return str(value).strip().upper() == "TRUE"


def _typed(value):
    """类型判断类函数的取值：日期在 Excel 里属于数字（序列号），
    本地无序列号语义，归到哪一类都给不出准确答案，统一标"未验证"。"""
    if isinstance(value, DateValue):
        raise UnsupportedFormula(DATE_HINT)
    return value


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


def _mod(args):
    """MOD：结果符号跟随除数（Excel 语义）。math.fmod 跟随被除数，负被除数时会给出相反符号。"""
    divisor = _to_number(_single(args[1]))
    if isinstance(divisor, ExcelError):
        return divisor
    dividend = _to_number(_single(args[0]))
    if isinstance(dividend, ExcelError):
        return dividend
    if divisor == 0:
        return DIV0
    return dividend % divisor


def _trunc(args):
    value = _to_number(_single(args[0]))
    if isinstance(value, ExcelError):
        return value
    digits = _require_int(args[1]) if len(args) > 1 else 0
    if isinstance(digits, ExcelError):
        return digits
    factor = 10 ** digits
    return math.trunc(value * factor) / factor


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
        raise UnsupportedFormula("VLOOKUP 第二参数必须是区域或数组")
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


def _sumproduct(args):
    """SUMPRODUCT：各参数逐位相乘再求和。

    参数长度不一致时 Excel 报 #VALUE!，这里必须显式检查——沿用 zip 会把长数组
    静默截断，算出一个偏小的"看起来正常"的数字。
    """
    columns = [[float(value) for value in _numbers(arg)] for arg in args]
    if len({len(column) for column in columns}) > 1:
        return VALUE_ERR
    return math.fsum(math.prod(pair) for pair in zip(*columns))


def _text_func(args, func):
    return func(_to_text(_single(args[0])))


# ------------------------------------------------------------------ 查找族
def _index(args):
    """INDEX(区域, 行号, [列号])：标量取值；单行区域只给一个下标时按列号解释。"""
    table = args[0]
    if not isinstance(table, list):
        raise UnsupportedFormula("INDEX 第一参数必须是区域或数组")
    grid = table if (table and isinstance(table[0], list)) else [table]
    height = len(grid)
    width = len(grid[0]) if grid else 0
    row_num = _require_int(args[1])
    if isinstance(row_num, ExcelError):
        return row_num
    if height == 1 and len(args) <= 2 and width > 1:
        # 单行区域 + 单下标：按列号解释，行号固定为 1
        col_num = row_num
        row_num = 1
    else:
        if len(args) > 2:
            col_num = _require_int(args[2])
            if isinstance(col_num, ExcelError):
                return col_num
        else:
            col_num = 1
    if row_num < 1 or col_num < 1:
        raise UnsupportedFormula("本地验证不支持 INDEX 的 0 下标（整行/整列返回）")
    if row_num > height or col_num > width:
        return ExcelError("#REF!")
    return grid[row_num - 1][col_num - 1]


def _match(args):
    """MATCH(查找值, 单行或单列区域, [匹配类型])：0 精确，1 升序近似（默认），-1 降序近似。"""
    key = _single(args[0])
    flat = _flatten(args[1])
    match_type = _require_int(args[2]) if len(args) > 2 else 1
    if isinstance(match_type, ExcelError):
        return match_type
    if match_type == 0:
        for index, value in enumerate(flat):
            if _compare("=", value, key) is True:
                return float(index + 1)
        return NA_ERR
    best = None
    for index, value in enumerate(flat):
        hit = _compare("<=", value, key) if match_type > 0 else _compare(">=", value, key)
        if hit is True:
            best = index + 1
    return float(best) if best else NA_ERR


def _hlookup(args):
    key = _single(args[0])
    table = args[1]
    if not isinstance(table, list):
        raise UnsupportedFormula("HLOOKUP 第二参数必须是区域或数组")
    row_index = _require_int(args[2])
    if isinstance(row_index, ExcelError):
        return row_index
    exact = len(args) > 3 and not _to_bool(_single(args[3]))
    if not exact:
        raise UnsupportedFormula("本地验证仅支持精确匹配的 HLOOKUP（第 4 参数为 FALSE）")
    grid = table if (table and isinstance(table[0], list)) else [table]
    header = grid[0] if grid else []
    for col, cell in enumerate(header):
        if _compare("=", cell, key) is True:
            if row_index > len(grid):
                return ExcelError("#REF!")
            row = grid[row_index - 1]
            return row[col] if col < len(row) else ExcelError("#REF!")
    return NA_ERR


def _xlookup(args):
    """XLOOKUP：仅支持默认的精确匹配、从头搜索。"""
    key = _single(args[0])
    lookup_values = _flatten(args[1])
    return_values = _flatten(args[2])
    if len(lookup_values) != len(return_values):
        return VALUE_ERR
    if len(args) > 4:
        mode = _to_number(_single(args[4]))
        if mode != 0:
            raise UnsupportedFormula("本地验证仅支持 XLOOKUP 精确匹配（第 5 参数为 0/省略）")
    for index, candidate in enumerate(lookup_values):
        if _compare("=", candidate, key) is True:
            return return_values[index]
    return _single(args[3]) if len(args) > 3 else NA_ERR


def _lookup(args):
    """LOOKUP 向量形式：在升序的查找列中做近似匹配，返回结果列同位置的值。"""
    key = _single(args[0])
    lookup_values = _flatten(args[1])
    result_values = _flatten(args[2]) if len(args) > 2 else lookup_values
    if len(result_values) < len(lookup_values):
        return VALUE_ERR
    best = None
    for index, value in enumerate(lookup_values):
        if _compare("<=", value, key) is True:
            best = index
    return result_values[best] if best is not None else NA_ERR


def _require_int(node_value):
    """把可能是错误值/文本的标量参数转成 int，错误值原样返回。"""
    number = _to_number(_single(node_value))
    if isinstance(number, ExcelError):
        return number
    return int(number)


# ------------------------------------------------------------------ 统计扩展
def _mode(args):
    counts: dict = {}
    ordered: list[float] = []
    for number in _numbers(args):
        if number not in counts:
            ordered.append(number)
        counts[number] = counts.get(number, 0) + 1
    if not counts:
        return NA_ERR
    best = max(counts.values())
    if best == 1:
        return NA_ERR  # 每个数都只出现一次，Excel 返回 #N/A
    return next(number for number in ordered if counts[number] == best)


def _conditional_extreme(args, mode: str):
    """MAXIFS / MINIFS：首参数为取值区域，其后是成对的条件区域与条件。"""
    target = _flatten(args[0])
    pairs = list(zip(args[1::2], args[2::2]))
    if not pairs:
        raise UnsupportedFormula("条件参数不成对")
    length = len(target)
    picked: list[float] = []
    for index in range(length):
        ok = True
        for rng, criteria in pairs:
            flat = _flatten(rng)
            if len(flat) != length:
                raise UnsupportedFormula("条件区域与取值区域大小不一致")
            if not _match_criteria(flat[index], _single(criteria)):
                ok = False
                break
        value = target[index]
        if ok and isinstance(value, (int, float)) and not isinstance(value, bool):
            picked.append(float(value))
    if not picked:
        return 0.0  # 没有任何满足条件的数值时 Excel 返回 0
    return max(picked) if mode == "MAX" else min(picked)


# ------------------------------------------------------------------ 数学扩展
def _exp(args):
    try:
        return math.exp(_to_number(_single(args[0])))
    except OverflowError:
        return NUM_ERR


def _ln(args):
    value = _to_number(_single(args[0]))
    if isinstance(value, ExcelError):
        return value
    return math.log(value) if value > 0 else NUM_ERR


def _log(args):
    value = _to_number(_single(args[0]))
    if isinstance(value, ExcelError):
        return value
    base = _to_number(_single(args[1])) if len(args) > 1 else 10.0
    if isinstance(base, ExcelError):
        return base
    if value <= 0 or base <= 0 or base == 1:
        return NUM_ERR
    return math.log(value, base)


def _log10(args):
    value = _to_number(_single(args[0]))
    if isinstance(value, ExcelError):
        return value
    return math.log10(value) if value > 0 else NUM_ERR


def _ceiling_floor(args, mode: str):
    """CEILING / FLOOR（正显著性）：结果朝 +∞ / -∞ 方向凑到显著性的整数倍。"""
    value = _to_number(_single(args[0]))
    significance = _to_number(_single(args[1]))
    if isinstance(value, ExcelError):
        return value
    if isinstance(significance, ExcelError):
        return significance
    if significance == 0:
        return DIV0
    if significance < 0:
        raise UnsupportedFormula("本地验证仅支持正显著性参数的 CEILING/FLOOR")
    ratio = value / significance
    if abs(ratio - round(ratio)) < 1e-9:  # 消除 0.1+0.2 类浮点噪声
        return round(ratio) * significance
    return (math.ceil(ratio) if mode == "CEILING" else math.floor(ratio)) * significance


# ------------------------------------------------------------------ 文本扩展
def _find(args, *, wildcard: bool):
    """FIND（区分大小写）/ SEARCH（不区分大小写，支持 * ? 通配符）：返回 1 起始位置。"""
    needle = _to_text(_single(args[0]))
    haystack = _to_text(_single(args[1]))
    start = _require_int(args[2]) if len(args) > 2 else 1
    if isinstance(start, ExcelError):
        return start
    if start < 1 or start > len(haystack) + 1:
        return VALUE_ERR
    hay = haystack[start - 1 :]
    if wildcard:
        pattern = "".join(
            ".*" if ch == "*" else "." if ch == "?" else re.escape(ch) for ch in needle
        )
        found = re.search(pattern, hay, re.IGNORECASE)
        return float(start + found.start()) if found else VALUE_ERR
    position = hay.find(needle)
    return float(start + position) if position >= 0 else VALUE_ERR


def _substitute(args):
    text = _to_text(_single(args[0]))
    old = _to_text(_single(args[1]))
    new = _to_text(_single(args[2]))
    if not old:
        return text
    if len(args) > 3:
        instance = _require_int(args[3])
        if isinstance(instance, ExcelError):
            return instance
        if instance < 1:
            return VALUE_ERR
        count = 0
        for index in range(len(text)):
            if text.startswith(old, index):
                count += 1
                if count == instance:
                    return text[:index] + new + text[index + len(old) :]
        return text  # 出现次数不足时原样返回
    return text.replace(old, new)


def _replace(args):
    text = _to_text(_single(args[0]))
    start = _require_int(args[1])
    length = _require_int(args[2])
    new = _to_text(_single(args[3]))
    if isinstance(start, ExcelError) or isinstance(length, ExcelError):
        return VALUE_ERR
    if start < 1 or length < 0:
        return VALUE_ERR
    return text[: start - 1] + new + text[start - 1 + length :]


def _rept(args):
    text = _to_text(_single(args[0]))
    count = _require_int(args[1])
    if isinstance(count, ExcelError):
        return count
    if count < 0 or len(text) * count > 32767:  # Excel 单元格上限
        return VALUE_ERR
    return text * count


def _char(args):
    code = _require_int(args[0])
    if isinstance(code, ExcelError):
        return code
    return chr(code) if 1 <= code <= 255 else VALUE_ERR


def _code(args):
    text = _to_text(_single(args[0]))
    return float(ord(text[0])) if text else VALUE_ERR


def _value(args):
    """VALUE：把文本转成数值，容忍千分位逗号、货币符号与百分号。"""
    raw = _single(args[0])
    if isinstance(raw, (int, float)) and not isinstance(raw, bool):
        return float(raw)
    text = _to_text(raw).strip()
    percent = text.endswith("%")
    if percent:
        text = text[:-1]
    cleaned = text.replace(",", "").replace("￥", "").replace("$", "").strip()
    try:
        number = float(cleaned)
    except ValueError:
        return VALUE_ERR
    return number / 100 if percent else number


# ------------------------------------------------------------------ 逻辑扩展
def _parity(args, *, want_even: bool):
    number = _to_number(_single(args[0]))
    if isinstance(number, ExcelError):
        return number
    remainder = abs(int(number)) % 2
    return remainder == 0 if want_even else remainder == 1


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
    "MOD": _mod,
    "INT": lambda a: float(math.floor(_to_number(_single(a[0])))),
    "TRUNC": _trunc,
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
    "ISNUMBER": lambda a: isinstance(_typed(_single(a[0])), (int, float))
    and not isinstance(_single(a[0]), bool),
    "ISTEXT": lambda a: isinstance(_typed(_single(a[0])), str),
    "ISERROR": lambda a: isinstance(_single(a[0]), ExcelError),
    "ISERR": lambda a: isinstance(_single(a[0]), ExcelError) and _single(a[0]) != NA_ERR,
    "ISNA": lambda a: _single(a[0]) == NA_ERR,
    "ISEVEN": lambda a: _parity(a, want_even=True),
    "ISODD": lambda a: _parity(a, want_even=False),
    "NA": lambda a: NA_ERR,
    "TRUE": lambda a: True,
    "FALSE": lambda a: False,
    "VLOOKUP": _vlookup,
    "HLOOKUP": _hlookup,
    "XLOOKUP": _xlookup,
    "LOOKUP": _lookup,
    "INDEX": _index,
    "MATCH": _match,
    "IFS": lambda a: None,  # 由 Evaluator 特殊处理（惰性求值）
    "CHOOSE": lambda a: None,  # 由 Evaluator 特殊处理（惰性求值）
    "SWITCH": lambda a: None,  # 由 Evaluator 特殊处理（惰性求值）
    "MODE": _mode,
    "MODE.SNGL": _mode,
    "MAXIFS": lambda a: _conditional_extreme(a, "MAX"),
    "MINIFS": lambda a: _conditional_extreme(a, "MIN"),
    "EXP": _exp,
    "LN": _ln,
    "LOG": _log,
    "LOG10": _log10,
    "CEILING": lambda a: _ceiling_floor(a, "CEILING"),
    "FLOOR": lambda a: _ceiling_floor(a, "FLOOR"),
    "FIND": lambda a: _find(a, wildcard=False),
    "SEARCH": lambda a: _find(a, wildcard=True),
    "SUBSTITUTE": _substitute,
    "REPLACE": _replace,
    "REPT": _rept,
    "CHAR": _char,
    "CODE": _code,
    "VALUE": _value,
    "XOR": lambda a: sum(1 for v in _flatten(a) if v is not None and _to_bool(v)) % 2 == 1,
    "VAR": lambda a: _safe_div(_numbers(a), lambda n: statistics.variance(n) if len(n) > 1 else DIV0),
    "VAR.S": lambda a: _safe_div(_numbers(a), lambda n: statistics.variance(n) if len(n) > 1 else DIV0),
    "VAR.P": lambda a: _safe_div(_numbers(a), statistics.pvariance),
    "STDEVP": lambda a: _safe_div(_numbers(a), statistics.pstdev),
    "SUMPRODUCT": _sumproduct,
}


def evaluate_formula(ast: object, sheets: dict[str, Worksheet], default_sheet: str):
    """对外接口：求值成功返回结果值，不支持时抛出 UnsupportedFormula。"""
    return Evaluator(sheets, default_sheet).evaluate(ast)


def format_value(value) -> str:
    if value is None:
        return "(空)"
    if isinstance(value, ExcelError):
        return value.code
    if isinstance(value, DateValue):
        raw = value.raw
        return raw.isoformat()[:10] if isinstance(raw, (_dt.datetime, _dt.date)) else raw.isoformat()
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, float):
        if value.is_integer():
            return str(int(value))
        return f"{value:.6g}"
    return str(value)
