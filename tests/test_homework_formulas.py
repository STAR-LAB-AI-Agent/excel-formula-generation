"""Excel 大作业实景测试：把课程作业里的公式清单原样搬进测试。

数据形态复刻 `Excel大作业.xlsm`（销售订单表 + 跨表查询区，含日期列与中文表名），
覆盖嵌套 IF 评级、统计区六件套、跨表 SUMIFS、IFERROR+VLOOKUP、单元格引用、
日期列防护与区域数组语义防护。期望值都按夹具里那 8 条订单手工核算过。
"""
from __future__ import annotations

import pytest

from excel_formula.evaluator import UnsupportedFormula, evaluate_formula, format_value
from excel_formula.excel_reader import WorkbookView
from excel_formula.formula_parser import parse_formula
from excel_formula.validator import validate_formula


def _eval(path, sheet, formula):
    with WorkbookView(path) as view:
        sheets = {name: view.wb_values[name] for name in view.sheet_names}
        return evaluate_formula(parse_formula(formula), sheets, sheet)


def _fmt(path, sheet, formula):
    return format_value(_eval(path, sheet, formula))


# ------------------------------------------------------------------ 读取器
def test_digest_keeps_chinese_sheet_and_dates(sales_xlsx):
    with WorkbookView(sales_xlsx) as view:
        digest = view.digest("原始数据表")
    assert digest.name == "原始数据表"
    assert digest.data_range == "A1:L9"
    header = digest.rows[0][1]
    assert header[0] == "订单编号" and header[9] == "总销售额(元)"
    first_order = digest.rows[1][1]
    assert first_order[0] == "ORD-001"
    assert first_order[1] == "2025-01-05"  # 日期以 ISO 文本搬运，不做类型推断


# ------------------------------------------------------------------ 嵌套 IF 评级
@pytest.mark.parametrize(
    "row,expected",
    [(2, "优秀"), (3, "良好"), (7, "优秀"), (8, "中等"), (4, "待提升")],
)
def test_three_level_nested_if_rating(sales_xlsx, row, expected):
    """=IF(J>=15000,"优秀",IF(J>=10000,"良好",IF(J>=6000,"中等","待提升"))) 的四档全覆盖。"""
    formula = (
        f'=IF(J{row}>=15000,"优秀",IF(J{row}>=10000,"良好",'
        f'IF(J{row}>=6000,"中等","待提升")))'
    )
    assert _fmt(sales_xlsx, "原始数据表", formula) == expected


# ------------------------------------------------------------------ 统计区
@pytest.mark.parametrize(
    "formula,expected",
    [
        ("=SUM(J2:J9)", "88000"),
        ("=AVERAGE(J2:J9)", "11000"),
        ("=MAX(J2:J9)", "27500"),
        ("=MIN(J2:J9)", "4000"),
        ("=COUNTA(A2:A9)", "8"),
        ('=COUNTIF(D2:D9,"华北")', "2"),
        ('=COUNTIF(E2:E9,"电子产品")', "2"),
        ("=H2*I2", "27500"),
        ("=J2>=15000", "TRUE"),
    ],
)
def test_summary_block_formulas(sales_xlsx, formula, expected):
    assert _fmt(sales_xlsx, "原始数据表", formula) == expected


# ------------------------------------------------------------------ 跨表查询
@pytest.mark.parametrize("cell,expected", [("D2", "51500"), ("D4", "8500")])
def test_cross_sheet_sumifs_by_owner(sales_xlsx, cell, expected):
    """按负责人汇总销售额：张伟 27500+24000，李娜 4000+4500。"""
    formula = f"=SUMIFS(原始数据表!$J$2:$J$9, 原始数据表!$G$2:$G$9, {cell})"
    assert _fmt(sales_xlsx, "查询表", formula) == expected


def test_cross_sheet_sumifs_with_quoted_sheet_name(sales_xlsx):
    """带单引号的中文表名同样可用（Excel 对含特殊字符的表名会加引号）。"""
    formula = "=SUMIFS('原始数据表'!$J$2:$J$9, '原始数据表'!$G$2:$G$9, D2)"
    assert _fmt(sales_xlsx, "查询表", formula) == "51500"


def test_cross_sheet_vlookup_with_iferror(sales_xlsx):
    formula = '=IFERROR(VLOOKUP($B$3, 原始数据表!$A$2:$L$9, 3, FALSE), "未找到订单")'
    assert _fmt(sales_xlsx, "查询表", formula) == "赵丽"


def test_vlookup_miss_falls_back_to_message(sales_xlsx):
    formula = '=IFERROR(VLOOKUP("ORD-999", 原始数据表!$A$2:$L$9, 3, FALSE), "未找到订单")'
    assert _fmt(sales_xlsx, "查询表", formula) == "未找到订单"


def test_vlookup_can_return_date_column(sales_xlsx):
    """查找区域里含日期列不再阻断整个查询；命中日期列时以 ISO 日期展示。"""
    formula = '=VLOOKUP("ORD-003", 原始数据表!$A$2:$L$9, 2, FALSE)'
    assert _fmt(sales_xlsx, "查询表", formula) == "2025-01-10"


def test_vlookup_returns_sales_amount(sales_xlsx):
    formula = '=VLOOKUP("ORD-006", 原始数据表!$A$2:$L$9, 10, FALSE)'
    assert _fmt(sales_xlsx, "查询表", formula) == "24000"


def test_cell_reference_chain(sales_xlsx):
    """=E7 这类纯引用在这份作业里大量出现，取值必须原样传递。"""
    assert _fmt(sales_xlsx, "查询表", "=B3") == "ORD-003"


# ------------------------------------------------------------------ 校验链路
def test_homework_formulas_pass_validation(sales_xlsx):
    statuses = [
        ('=IF(J2>=15000,"优秀",IF(J2>=10000,"良好",IF(J2>=6000,"中等","待提升")))', "原始数据表", "M2"),
        ('=COUNTIF(D2:D9,"华北")', "原始数据表", "M12"),
        ("=SUMIFS(原始数据表!$J$2:$J$9, 原始数据表!$G$2:$G$9, D2)", "查询表", "E2"),
        ('=IFERROR(VLOOKUP($B$3, 原始数据表!$A$2:$L$9, 3, FALSE), "未找到订单")', "查询表", "B5"),
    ]
    with WorkbookView(sales_xlsx) as view:
        sheet_names = view.sheet_names
        digests = {name: view.digest(name) for name in ("原始数据表", "查询表")}
    for formula, sheet, target in statuses:
        result, ast = validate_formula(
            formula, target=target, sheet=sheet, digest=digests[sheet], sheet_names=sheet_names
        )
        assert result.ok, f"{formula} -> {result.errors}"
        assert ast is not None


def test_defined_name_is_rejected_by_validator(sales_xlsx):
    """名称管理器里的 saledata 不是单元格地址，校验阶段直接拒绝并给出指引。"""
    result, _ = validate_formula("=SUM(saledata)")
    assert not result.ok
    assert "无法识别的标识符" in result.error_text()


def test_new_functions_pass_validation():
    """本轮新实现的函数都应能通过白名单与参数个数检查。"""
    for formula in [
        "=INDEX(B2:F6,2,3)",
        "=MATCH(85,B2:F2,0)",
        "=HLOOKUP(85,B2:F6,2,FALSE)",
        "=XLOOKUP(85,B2:F2,B3:F3)",
        "=LOOKUP(85,B2:B6,B3:B7)",
        "=IFS(B2>90,\"优\",B2>80,\"良\",TRUE,\"差\")",
        "=MAXIFS(B2:F6,B2:B6,\">80\")",
        "=MODE({1,2,2,3})",
        "=CEILING(87.2,1)",
        "=VAR.P(B2:F2)",
    ]:
        result, _ = validate_formula(formula)
        assert result.ok, f"{formula} -> {result.errors}"


# ------------------------------------------------------------------ 防护行为
def test_defined_name_marks_unverified_in_evaluator(sales_xlsx):
    with pytest.raises(UnsupportedFormula, match="定义名称"):
        _eval(sales_xlsx, "原始数据表", "=SUM(saledata)")


def test_aggregating_over_date_column_is_unverified(sales_xlsx):
    """日期列聚合：Excel 会把日期当序列号求和，本地给不出该语义，标"未验证"。"""
    with pytest.raises(UnsupportedFormula, match="日期"):
        _eval(sales_xlsx, "原始数据表", "=SUM(B2:B9)")


def test_array_style_condition_is_unverified(sales_xlsx):
    """(D2:D9="华北")*J2:J9 属于数组公式语义，宁可"未验证"也不能只比首元素。"""
    with pytest.raises(UnsupportedFormula, match="数组"):
        _eval(sales_xlsx, "原始数据表", '=SUM((D2:D9="华北")*J2:J9)')
