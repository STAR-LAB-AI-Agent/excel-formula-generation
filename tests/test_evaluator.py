"""求值器与读取器测试：整表无损文本化、本地公式计算（可选功能"公式自动验证"的基础）。"""
from __future__ import annotations

import pytest

from excel_formula.evaluator import UnsupportedFormula, evaluate_formula, format_value
from excel_formula.excel_reader import WorkbookView
from excel_formula.formula_parser import parse_formula


def _eval(path, formula):
    with WorkbookView(path) as view:
        sheets = {name: view.wb_values[name] for name in view.sheet_names}
        return evaluate_formula(parse_formula(formula), sheets, "Sheet1")


# ------------------------------------------------------------------ 读取器
def test_digest_dumps_every_cell(sample_xlsx):
    """读取器只搬运不判断：每一行都要在里面，行号连续。"""
    with WorkbookView(sample_xlsx) as view:
        digest = view.digest()
        assert digest.name == "Sheet1"  # Sheet2 为空表，应自动跳过
        assert digest.data_range == "A1:F6"
        assert [row for row, _ in digest.rows] == [1, 2, 3, 4, 5, 6]
        assert digest.rows[0][1] == [
            "Subject", "Student1", "Student2", "Student3", "Student4", "Student5"
        ]
        assert digest.rows[1][1] == ["Math", "85", "78", "92", "88", "95"]
        assert digest.truncated_rows == 0


def test_digest_prompt_is_lossless_tsv(sample_xlsx):
    """提示文本里必须带坐标系，且末行数据不能丢。"""
    with WorkbookView(sample_xlsx) as view:
        text = view.digest().to_prompt()
    assert "\tA\tB\tC\tD\tE\tF" in text  # 列字母表头
    assert "1\tSubject\tStudent1" in text  # 行号 + 原始值
    assert "6\tBiology\t78\t80\t86\t83\t81" in text  # 最后一行也在
    assert "已省略" not in text  # 小表不应触发截断


def test_digest_truncates_large_table_at_both_ends(tmp_path):
    """超过行数闸门时保留头尾两段，并显式告知已截断。"""
    from openpyxl import Workbook

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Sheet1"
    sheet.append(["序号", "重量"])
    for index in range(1, 500):
        sheet.append([index, index * 2])
    path = tmp_path / "big.xlsx"
    workbook.save(path)
    workbook.close()

    with WorkbookView(path) as view:
        digest = view.digest()
        text = digest.to_prompt()
    assert digest.max_row == 500
    assert digest.truncated_rows == 500 - 60  # 头 50 行 + 尾 10 行
    assert [row for row, _ in digest.rows][:3] == [1, 2, 3]
    assert [row for row, _ in digest.rows][-1] == 500
    assert "已省略" in text


# ------------------------------------------------------------------ 求值器
@pytest.mark.parametrize(
    "formula,expected",
    [
        ("=SUM(B2:F2)", "438"),
        ("=AVERAGE(B2:F2)", "87.6"),
        ("=MAX(B2:F6)", "95"),
        ("=MIN(B2:F6)", "75"),
        ("=COUNT(B2:F2)", "5"),
        ('=COUNTIF(B2:F2,">85")', "3"),
        ('=SUMIF(A2:A6,"Math",B2:B6)', "85"),
        ('=IF(AVERAGE(B2:F2)>=85,"优秀","一般")', "优秀"),
        ("=ROUND(AVERAGE(B2:F2),1)", "87.6"),
        ("=B2/SUM(B2:B6)", "0.2019"),  # 85/421
        ("=B2&\"分\"", "85分"),
        ('=VLOOKUP("Physics",A2:F6,2,FALSE)', "80"),
    ],
)
def test_evaluate_common_formulas(sample_xlsx, formula, expected):
    assert format_value(_eval(sample_xlsx, formula)) == expected


def test_division_by_zero_returns_excel_error(sample_xlsx):
    assert format_value(_eval(sample_xlsx, "=B2/(B2-B2)")) == "#DIV/0!"


def test_unsupported_function_raises(sample_xlsx):
    with pytest.raises(UnsupportedFormula):
        _eval(sample_xlsx, "=NETWORKDAYS(A2,A3)")
