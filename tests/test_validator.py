"""校验器测试：正常公式、语法错误、函数白名单、参数个数、循环引用、越界引用。"""
from __future__ import annotations

from excel_formula.excel_reader import WorkbookView
from excel_formula.validator import validate_formula


def _digest(path):
    with WorkbookView(path) as view:
        return view.digest("Sheet1"), view.sheet_names


def test_valid_formula_passes(sample_xlsx):
    digest, names = _digest(sample_xlsx)
    result, ast = validate_formula(
        "=SUM(B2:F2)", target="G2", sheet="Sheet1", digest=digest, sheet_names=names
    )
    assert result.ok
    assert result.errors == []
    assert result.functions == ["SUM"]
    assert "B2:F2" in result.refs
    assert ast is not None


def test_missing_equal_sign_is_syntax_error():
    result, ast = validate_formula("SUM(B2:F2)")
    assert not result.ok
    assert "必须以 = 开头" in result.error_text()
    assert ast is None


def test_unbalanced_parenthesis():
    result, _ = validate_formula("=SUM(B2:F2")
    assert not result.ok
    assert "右括号" in result.error_text()


def test_forbidden_function_rejected():
    result, _ = validate_formula('=INDIRECT("A1")')
    assert not result.ok
    assert "高风险函数" in result.error_text()


def test_unknown_function_rejected():
    result, _ = validate_formula("=MYSUM(B2:F2)")
    assert not result.ok
    assert "不在允许列表" in result.error_text()


def test_wrong_argument_count():
    result, _ = validate_formula("=ROUND()")
    assert not result.ok
    assert "参数个数" in result.error_text()


def test_circular_reference_detected(sample_xlsx):
    digest, names = _digest(sample_xlsx)
    result, _ = validate_formula(
        "=SUM(B2:G2)", target="G2", sheet="Sheet1", digest=digest, sheet_names=names
    )
    assert not result.ok
    assert "循环引用" in result.error_text()


def test_out_of_data_range_warns(sample_xlsx):
    digest, names = _digest(sample_xlsx)
    result, _ = validate_formula(
        "=SUM(B20:F20)", target="G2", sheet="Sheet1", digest=digest, sheet_names=names
    )
    assert result.ok  # 只是提醒，不阻断
    assert any("数据区" in w for w in result.warnings)


def test_unknown_sheet_rejected(sample_xlsx):
    digest, names = _digest(sample_xlsx)
    result, _ = validate_formula(
        "=SUM(汇总!A1:A3)", target="G2", sheet="Sheet1", digest=digest, sheet_names=names
    )
    assert not result.ok
    assert "工作表" in result.error_text()


def test_chinese_column_name_rejected():
    result, _ = validate_formula("=SUM(成绩)")
    assert not result.ok
    assert "无法识别的标识符" in result.error_text()


def test_nested_formula_and_cross_sheet_ok(sample_xlsx):
    digest, names = _digest(sample_xlsx)
    result, _ = validate_formula(
        '=IF(AVERAGE(B2:F2)>=85,"优秀",IF(AVERAGE(B2:F2)>=75,"良好","一般"))',
        target="H2",
        sheet="Sheet1",
        digest=digest,
        sheet_names=names,
    )
    assert result.ok
    assert set(result.functions) == {"IF", "AVERAGE"}


def test_volatile_function_warns():
    result, _ = validate_formula("=TODAY()")
    assert result.ok
    assert any("易变函数" in w for w in result.warnings)


def test_absolute_and_percent_syntax():
    result, _ = validate_formula("=$B$2/SUM($B$2:$F$2)*100%")
    assert result.ok, result.errors
