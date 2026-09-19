"""真实大作业文件冒烟测试：直接打开 Excel大作业.xlsm 做"求值器重算 vs Excel 缓存值"对账。

文件不在仓库里（属个人作业文件）时整组自动跳过。对账成功说明求值器在这份
真实数据上与 Excel 的独立计算结果一致；本地不支持的公式（VBA 自定义函数、
定义名称、模拟运算表等）计入"未验证"清单，不影响通过。
"""
from __future__ import annotations

import datetime as _dt
from pathlib import Path

import pytest

from excel_formula.evaluator import (
    DateValue,
    ExcelError,
    UnsupportedFormula,
    evaluate_formula,
    format_value,
)
from excel_formula.excel_reader import WorkbookView
from excel_formula.external import ExternalBookLoader
from excel_formula.formula_parser import FormulaSyntaxError, parse_formula

HOMEWORK = Path(__file__).resolve().parents[1] / "Excel大作业.xlsm"

pytestmark = pytest.mark.skipif(
    not HOMEWORK.is_file(), reason="本地没有 Excel大作业.xlsm，跳过真实文件冒烟测试"
)


def _iter_formulas(view: WorkbookView):
    """遍历所有工作表的公式单元格；DataTableFormula 等非字符串对象自动跳过。"""
    for name in view.sheet_names:
        ws = view.wb_formulas[name]
        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell.value, str) and cell.value.startswith("="):
                    yield name, cell.coordinate, cell.value


def _as_number(value):
    """把一端结果折算成数值用于对账：日期类（本地 DateValue / Excel 缓存 datetime）→ 序列号。"""
    if isinstance(value, DateValue):
        value = value.raw
    if isinstance(value, _dt.datetime):
        return (value - _dt.datetime(1899, 12, 30)).total_seconds() / 86400
    if isinstance(value, _dt.date):
        return float((value - _dt.date(1899, 12, 30)).days)
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return float(value)
    return None


def _same(computed, cached) -> bool:
    """求值结果与 Excel 缓存值是否一致（数值给 1e-6 容差，文本精确比对）。

    日期在 Excel 里就是序列号：本地 DateValue 与缓存 datetime 先折算成序列号
    再按数值容差比对，避免同一时刻因表示形式不同被误判为不一致。
    """
    if isinstance(computed, bool) and isinstance(cached, bool):
        return computed is cached
    local, reference = _as_number(computed), _as_number(cached)
    if local is not None and reference is not None:
        return abs(local - reference) < 1e-6
    return format_value(computed) == str(cached)


def test_homework_workbook_reads():
    """真实文件能被无损打开：关键工作表在位，概览接口可用。"""
    with WorkbookView(HOMEWORK) as view:
        assert "原始数据表" in view.sheet_names
        overview = view.overview()
        assert overview
        assert any(item["sheet"] == "原始数据表" and item["rows"] > 1 for item in overview)


def test_homework_formulas_match_excel_cached_values():
    """把大作业里的每个公式重算一遍，与 Excel 上次计算缓存的值逐一对账。"""
    with WorkbookView(HOMEWORK) as view:
        sheets = {name: view.wb_values[name] for name in view.sheet_names}
        entries = list(_iter_formulas(view))
        cached_values = {
            (name, coord): view.wb_values[name][coord].value for name, coord, _ in entries
        }

    assert entries, "大作业文件里应当存在公式，一个都没找到说明测试用错文件"

    checked = 0
    unsupported: list[str] = []
    mismatches: list[str] = []
    loader = ExternalBookLoader()
    try:
        for sheet, coord, formula in entries:
            try:
                computed = evaluate_formula(
                    parse_formula(formula),
                    sheets,
                    sheet,
                    load_external=loader,
                    book_path=HOMEWORK,
                )
            except (UnsupportedFormula, FormulaSyntaxError):
                unsupported.append(f"{sheet}!{coord} {formula}")
                continue
            if isinstance(computed, ExcelError):
                continue  # 错误值与缓存错误文本的比对不在冒烟测试范围内
            checked += 1
            cached = cached_values[(sheet, coord)]
            if not _same(computed, cached):
                mismatches.append(
                    f"{sheet}!{coord} {formula} -> 本地 {format_value(computed)!r} / Excel {cached!r}"
                )
    finally:
        loader.close()

    assert checked >= 5, (
        f"参与对账的公式只有 {checked} 个，冒烟测试失去意义。"
        f"未验证清单：\n" + "\n".join(unsupported)
    )
    assert not mismatches, "以下公式本地重算与 Excel 缓存值不一致：\n" + "\n".join(mismatches)
