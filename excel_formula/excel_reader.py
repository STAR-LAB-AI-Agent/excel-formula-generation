"""基于 openpyxl 读取工作表结构，生成"紧凑表格描述"（低 Token 的关键环节）。

同一个文件会被打开两次：
* formulas 模式：拿到单元格里真正的公式文本；
* data_only 模式：拿到 Excel 上次计算后的缓存值，用于类型推断与公式自动验证。
"""
from __future__ import annotations

import datetime as _dt
from dataclasses import dataclass, field
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from .config import (
    DIGEST_HEAD_ROWS,
    DIGEST_MAX_CELL_TEXT,
    DIGEST_MAX_COLUMNS,
    DIGEST_MAX_ROWS,
    DIGEST_TAIL_ROWS,
)


@dataclass
class SheetDigest:
    """一张工作表的无损文本化结果。

    刻意不推断表头行与列类型：那属于语义判断，交给模型；这里只负责忠实搬运。
    """

    name: str
    max_row: int
    max_column: int
    # (行号, 各列文本)，行号连续；仅在超出行数闸门时出现缺口
    rows: list[tuple[int, list[str]]] = field(default_factory=list)
    truncated_rows: int = 0
    truncated_columns: int = 0

    @property
    def data_range(self) -> str:
        if not self.max_row or not self.max_column:
            return "空表"
        return f"A1:{get_column_letter(self.max_column)}{self.max_row}"

    def to_prompt(self) -> str:
        """整表 TSV：首列行号、首行列字母，单元格原始值不做任何解释。"""
        if not self.rows:
            return f"工作表 {self.name}：空表（无数据）"

        limit = min(self.max_column, DIGEST_MAX_COLUMNS)
        lines = [
            f"工作表: {self.name} | 已用区域: {self.data_range} | 总行数: {self.max_row}",
            "下表为单元格原始值（TSV，首列是行号，首行是列字母，空白即空单元格）。",
            "\t" + "\t".join(get_column_letter(c) for c in range(1, limit + 1)),
        ]
        previous = 0
        for row_index, values in self.rows:
            if previous and row_index > previous + 1:
                lines.append(f"…（第 {previous + 1}-{row_index - 1} 行已省略，共 {row_index - previous - 1} 行）")
            lines.append(f"{row_index}\t" + "\t".join(values))
            previous = row_index
        if self.truncated_columns:
            lines.append(f"（另有 {self.truncated_columns} 列未列出）")
        return "\n".join(lines)

    def to_dict(self) -> dict:
        return {
            "sheet": self.name,
            "data_range": self.data_range,
            "max_row": self.max_row,
            "max_column": self.max_column,
            "rows": [{"row": r, "values": v} for r, v in self.rows],
            "truncated_rows": self.truncated_rows,
            "truncated_columns": self.truncated_columns,
        }


class WorkbookView:
    """成对持有 formulas / values 两个视图的工作簿封装。"""

    def __init__(self, path: str | Path):
        self.path = Path(path)
        self.wb_formulas = load_workbook(self.path, data_only=False)
        self.wb_values = load_workbook(self.path, data_only=True)

    # -------------------------------------------------------------- 基本信息
    @property
    def sheet_names(self) -> list[str]:
        return list(self.wb_formulas.sheetnames)

    def resolve_sheet(self, sheet: str | None) -> str:
        """未指定工作表时取第一张"有数据"的表，避免默认落在空表上。"""
        if sheet:
            if sheet not in self.wb_formulas.sheetnames:
                raise KeyError(
                    f"工作表 {sheet!r} 不存在，可用工作表：{', '.join(self.sheet_names)}"
                )
            return sheet
        for name in self.wb_formulas.sheetnames:
            ws = self.wb_formulas[name]
            if ws.max_row > 1 or ws.cell(row=1, column=1).value is not None:
                return name
        return self.wb_formulas.sheetnames[0]

    def formula_sheet(self, sheet: str | None = None) -> Worksheet:
        return self.wb_formulas[self.resolve_sheet(sheet)]

    def value_sheet(self, sheet: str | None = None) -> Worksheet:
        return self.wb_values[self.resolve_sheet(sheet)]

    def cell_formula(self, sheet: str | None, ref: str) -> str | None:
        value = self.formula_sheet(sheet)[ref].value
        return value if isinstance(value, str) and value.startswith("=") else None

    def close(self) -> None:
        self.wb_formulas.close()
        self.wb_values.close()

    def __enter__(self) -> "WorkbookView":
        return self

    def __exit__(self, *exc_info: object) -> None:
        self.close()

    # -------------------------------------------------------------- 结构摘要
    def digest(
        self,
        sheet: str | None = None,
        *,
        max_rows: int = DIGEST_MAX_ROWS,
        max_columns: int = DIGEST_MAX_COLUMNS,
    ) -> SheetDigest:
        name = self.resolve_sheet(sheet)
        ws_f = self.wb_formulas[name]
        ws_v = self.wb_values[name]

        max_row = _effective_max_row(ws_f)
        max_col = _effective_max_col(ws_f)
        if max_row == 0 or max_col == 0:
            return SheetDigest(name=name, max_row=0, max_column=0)

        limit = min(max_col, max_columns)
        kept = _rows_to_keep(max_row, max_rows)
        rows = [
            (
                row,
                [_cell_text(_display_value(ws_f, ws_v, row, col)) for col in range(1, limit + 1)],
            )
            for row in kept
        ]

        return SheetDigest(
            name=name,
            max_row=max_row,
            max_column=max_col,
            rows=rows,
            truncated_rows=max_row - len(kept),
            truncated_columns=max(0, max_col - limit),
        )

    def overview(self) -> list[dict]:
        """整个工作簿的一行式概览，用于“这个文件里有什么”这类提问。"""
        result = []
        for name in self.sheet_names:
            ws = self.wb_formulas[name]
            rows, cols = _effective_max_row(ws), _effective_max_col(ws)
            first_row: list[str] = []
            if rows and cols:
                first_row = [
                    _cell_text(ws.cell(row=1, column=c).value)
                    for c in range(1, min(cols, 12) + 1)
                ]
            result.append({"sheet": name, "rows": rows, "columns": cols, "first_row": first_row})
        return result


# ------------------------------------------------------------------ 内部工具
def _effective_max_row(ws: Worksheet) -> int:
    """openpyxl 的 max_row 可能包含只带格式的空行，这里向上收缩到真实数据。"""
    row = ws.max_row or 0
    while row > 0:
        if any(ws.cell(row=row, column=c).value is not None for c in range(1, (ws.max_column or 1) + 1)):
            return row
        row -= 1
    return 0


def _effective_max_col(ws: Worksheet) -> int:
    col = ws.max_column or 0
    while col > 0:
        if any(ws.cell(row=r, column=col).value is not None for r in range(1, (ws.max_row or 1) + 1)):
            return col
        col -= 1
    return 0


def _rows_to_keep(max_row: int, max_rows: int) -> list[int]:
    """行数超限时只保留头尾两段，省略位置由 to_prompt 显式标注。"""
    if max_row <= max_rows:
        return list(range(1, max_row + 1))
    head = range(1, DIGEST_HEAD_ROWS + 1)
    tail = range(max(DIGEST_HEAD_ROWS + 1, max_row - DIGEST_TAIL_ROWS + 1), max_row + 1)
    return list(head) + list(tail)


def _display_value(ws_f: Worksheet, ws_v: Worksheet, row: int, col: int):
    """优先展示公式文本，其次展示缓存值。"""
    raw = ws_f.cell(row=row, column=col).value
    if isinstance(raw, str) and raw.startswith("="):
        return raw
    cached = ws_v.cell(row=row, column=col).value
    return raw if cached is None else cached


def _cell_text(value: object, limit: int = DIGEST_MAX_CELL_TEXT) -> str:
    if value is None:
        return ""
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, float) and value.is_integer():
        text = str(int(value))
    elif isinstance(value, (_dt.datetime, _dt.date)):
        text = value.isoformat()[:10]
    else:
        text = str(value)
    text = text.replace("\n", " ").replace("\t", " ").strip()
    return text if len(text) <= limit else text[: limit - 1] + "…"
