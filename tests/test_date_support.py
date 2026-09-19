"""日期按 Excel 序列号语义参与运算：算术、比较、条件匹配、日期函数与文本化。"""
from __future__ import annotations

import datetime as _dt

import pytest

from excel_formula.evaluator import UnsupportedFormula, evaluate_formula, format_value
from excel_formula.excel_reader import WorkbookView
from excel_formula.formula_parser import parse_formula

SHEET = "原始数据表"


def _eval(path, formula):
    with WorkbookView(path) as view:
        sheets = {name: view.wb_values[name] for name in view.sheet_names}
        return evaluate_formula(parse_formula(formula), sheets, SHEET)


def _fmt(path, formula):
    return format_value(_eval(path, formula))


# ------------------------------------------------------------------ 算术与比较
@pytest.mark.parametrize(
    "formula,expected",
    [
        ("=B2+30", "2025-02-04"),  # 日期 + 天数仍是日期
        ("=B2-5", "2024-12-31"),
        ("=30+B2", "2025-02-04"),
        ("=B9-B2", "27"),  # 日期 - 日期 = 天数
        ("=B2=DATE(2025,1,5)", "TRUE"),
        ("=B2<B3", "TRUE"),
        ('=B2>"2025-01-01"', "TRUE"),  # 文本日期按 ISO 解析后比较
        ("=SUM(B2:B9)", "365375"),  # 日期列按序列号汇总
        ("=MAX(B2:B9)", "45689"),  # 2025-02-01 的序列号
    ],
)
def test_date_arithmetic_and_comparison(sales_xlsx, formula, expected):
    assert _fmt(sales_xlsx, formula) == expected


def test_date_with_unparseable_text_still_unverified(sales_xlsx):
    """日期与无法解析的文本比较时仍标"未验证"，不给静默算错的答案。"""
    with pytest.raises(UnsupportedFormula, match="日期"):
        _eval(sales_xlsx, '=B2>"昨天"')


# ------------------------------------------------------------------ DATE 族构造
@pytest.mark.parametrize(
    "formula,expected",
    [
        ("=DATE(2025,1,5)", "2025-01-05"),
        ("=DATE(2025,13,1)", "2026-01-01"),  # 月份溢出进位
        ("=DATE(2025,1,0)", "2024-12-31"),  # 0 日 = 上月末
        ("=DATE(2025,2,30)", "2025-03-02"),  # 日溢出进位
        ("=DATE(100,1,1)", "2000-01-01"),  # 0~1899 的年自动加 1900
        ("=DATE(1899,1,1)", "3799-01-01"),  # 同上：1899 + 1900
    ],
)
def test_date_constructor(sales_xlsx, formula, expected):
    assert _fmt(sales_xlsx, formula) == expected


# ------------------------------------------------------------------ 日期函数
@pytest.mark.parametrize(
    "formula,expected",
    [
        ("=YEAR(B2)", "2025"),
        ("=MONTH(B2)", "1"),
        ("=DAY(B2)", "5"),
        ("=WEEKDAY(B2)", "1"),  # 2025-01-05 是周日
        ("=WEEKDAY(B2,2)", "7"),
        ("=WEEKDAY(B2,3)", "6"),
        ("=WEEKNUM(B2)", "2"),  # 周日起算：1/5 进入第 2 周
        ("=WEEKNUM(B2,2)", "1"),  # 周一起算：仍在第 1 周
        ("=EOMONTH(B2,0)", "2025-01-31"),
        ("=EOMONTH(B2,1)", "2025-02-28"),
        ("=EDATE(B2,1)", "2025-02-05"),
        ("=EDATE(DATE(2025,1,31),1)", "2025-02-28"),  # 目标月不足时取月末
        ("=DAYS(B9,B2)", "27"),
        ('=DATEVALUE("2025/1/5")', "2025-01-05"),
        ('=DATEVALUE("不是日期")', "#VALUE!"),
        ("=HOUR(0.5)", "12"),
        ("=MINUTE(0.5)", "0"),
        ("=SECOND(B2)", "0"),
        ("=NETWORKDAYS(DATE(2025,1,5),DATE(2025,1,10))", "5"),
        ("=WORKDAY(DATE(2025,1,5),1)", "2025-01-06"),
    ],
)
def test_date_functions(sales_xlsx, formula, expected):
    assert _fmt(sales_xlsx, formula) == expected


@pytest.mark.parametrize(
    "unit,expected",
    [("Y", "1"), ("M", "12"), ("D", "366"), ("YM", "0"), ("MD", "1"), ("YD", "1")],
)
def test_datedif_units(sales_xlsx, unit, expected):
    formula = f'=DATEDIF(B2,DATE(2026,1,6),"{unit}")'
    assert _fmt(sales_xlsx, formula) == expected


def test_datedif_end_before_start_is_num_error(sales_xlsx):
    assert _fmt(sales_xlsx, '=DATEDIF(DATE(2026,1,6),B2,"Y")') == "#NUM!"


# ------------------------------------------------------------------ 文本与条件
def test_date_to_text_uses_iso(sales_xlsx):
    """日期与文本拼接走 ISO 文本；">="&DATE(...) 这类条件文本也能还原成日期。"""
    assert _fmt(sales_xlsx, '="下单日期："&B2') == "下单日期：2025-01-05"


@pytest.mark.parametrize(
    "formula,expected",
    [
        ('=COUNTIF(B2:B9,">="&DATE(2025,1,15))', "4"),
        ('=COUNTIF(B2:B9,">2025-01-15")', "3"),
        ('=COUNTIF(B2:B9,"<="&DATE(2025,1,10))', "3"),
        ('=SUMIFS(J2:J9,B2:B9,">="&DATE(2025,1,15))', "40000"),
        ('=SUMIF(B2:B9,"<="&DATE(2025,1,10),J2:J9)', "43500"),
    ],
)
def test_date_criteria_matching(sales_xlsx, formula, expected):
    assert _fmt(sales_xlsx, formula) == expected


def test_date_condition_in_if(sales_xlsx):
    assert _fmt(sales_xlsx, '=IF(B2<DATE(2025,1,10),"早批次","晚批次")') == "早批次"


def test_isnumber_treats_date_as_number(sales_xlsx):
    """日期在 Excel 里就是数字：类型判断按序列号归属。"""
    assert _fmt(sales_xlsx, "=ISNUMBER(B2)") == "TRUE"
    assert _fmt(sales_xlsx, "=ISTEXT(B2)") == "FALSE"


def test_today_flows_through_functions(sales_xlsx):
    assert _fmt(sales_xlsx, "=YEAR(TODAY())") == str(_dt.date.today().year)
