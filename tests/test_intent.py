"""意图识别与日志脱敏测试（这两部分完全本地、零 Token）。"""
from __future__ import annotations

from excel_formula.intent import (
    INTENT_DESCRIBE,
    INTENT_EXPLAIN,
    INTENT_GENERATE,
    INTENT_VALIDATE,
    classify,
    file_candidates,
)
from excel_formula.logger import redact


def test_generate_intent_with_target_and_fill():
    intent = classify("帮我在G2算每个科目的总分，填充到G6")
    assert intent.kind == INTENT_GENERATE
    assert intent.target == "G2"
    assert intent.fill_to == "G6"


def test_generate_intent_extracts_file_and_sheet():
    intent = classify("在 测试数据.xlsx 的工作表Sheet1里统计平均分，写入H2")
    assert intent.kind == INTENT_GENERATE
    assert intent.file == "测试数据.xlsx"
    assert intent.sheet == "Sheet1"
    assert intent.target == "H2"


def test_file_candidates_strips_leading_chinese_prefix():
    # 中文里“在VLOOKUP.xlsx”没有空格，抓到的名字会粘上前面的介词
    assert file_candidates("在VLOOKUP.xlsx") == ["在VLOOKUP.xlsx", "VLOOKUP.xlsx"]
    # 纯汉字文件名也合法，逐字剥离但不能提前丢掉真名
    assert "测试数据.xlsx" in file_candidates("测试数据.xlsx")
    # 无中文前缀时只有原名
    assert file_candidates("VLOOKUP.xlsx") == ["VLOOKUP.xlsx"]


def test_describe_intent():
    intent = classify("这个表有哪些列？")
    assert intent.kind == INTENT_DESCRIBE


def test_validate_intent_with_formula():
    intent = classify('检查一下 =SUMIF(A2:A6,"Math",B2:B6) 对不对')
    assert intent.kind == INTENT_VALIDATE
    assert intent.formula == '=SUMIF(A2:A6,"Math",B2:B6)'


def test_pasted_formula_defaults_to_validate():
    intent = classify("=AVERAGE(B2:F2)")
    assert intent.kind == INTENT_VALIDATE


def test_explain_intent_with_cell():
    intent = classify("解释一下G2里的公式")
    assert intent.kind == INTENT_EXPLAIN
    assert intent.cell == "G2"


def test_source_range_not_taken_as_target():
    intent = classify("统计B2:F2里大于85的个数")
    assert intent.fill_to is None
    assert intent.kind == INTENT_GENERATE


def test_biaoge_not_mistaken_for_sheet_name():
    # “表格”里的“表”不是工作表前缀：截断出的“格”不应被当成表名
    intent = classify("在金额列求出每个订单的总金额，并在表格下方合适位置列一新表统计不同区域的总金额")
    assert intent.sheet is None


def test_single_char_candidate_is_skipped():
    # 停用字截断后只剩单个汉字时不当作表名（如“表单里”切出的“单”）
    assert classify("在表单里填一列序号").sheet is None


def test_action_verb_after_new_sheet_is_not_a_sheet_name():
    # “列一个新表汇总……”中的“表”是“新表”的尾字，后面跟动作词，不是表名
    assert classify("在表格下方列一个新表汇总各区域的总金额").sheet is None


def test_empty_input_marks_missing():
    assert "request" in classify("   ").missing


def test_redact_hides_api_key():
    text = "Authorization: Bearer sk-abcdef1234567890 用于调用"
    masked = redact(text)
    assert "sk-abcdef1234567890" not in masked
    assert "***" in masked


def test_redact_hides_key_assignment():
    assert "1234" not in redact("DEEPSEEK_API_KEY=sk-1234567890abcdef")
