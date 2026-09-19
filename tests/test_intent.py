"""意图识别与日志脱敏测试（这两部分完全本地、零 Token）。"""
from __future__ import annotations

from excel_formula.intent import (
    INTENT_DESCRIBE,
    INTENT_DROPDOWN,
    INTENT_EXPLAIN,
    INTENT_FORMAT,
    INTENT_GENERATE,
    INTENT_VALIDATE,
    align_sheet_name,
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


def test_new_table_intent_detection():
    """“新建一张表”的写法由本地规则识别：写入后自动套用表格格式。"""
    assert classify("生成一个新的表格，统计各班平均分").new_table is True
    assert classify("在Sheet1里新建一张汇总表").new_table is True
    assert classify("帮我做一张统计表").new_table is True
    assert classify("建一张分组表").new_table is True
    assert classify("生成一个新的表格").to_dict()["new_table"] is True


def test_new_table_intent_ignores_column_addition_and_sheet_names():
    """给已有表加列、或只是引用已有表名（生成成绩表的汇总）不应误判为新建表格。"""
    assert classify("在成绩表里新增一列总分").new_table is False
    assert classify("生成成绩表的汇总列").new_table is False
    assert classify("新建工作表").new_table is False
    assert classify("统计各科平均分写入H2").new_table is False


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


# ------------------------------------------------------------------ 下拉列表
def test_dropdown_shortform_intent():
    """“加下拉/做个下拉菜单”等简写也判下拉；“下拉到B9”是填充语义，不能误判。"""
    assert classify("给订单查询B3加下拉，选项来自销售订单!A2:A25").kind == INTENT_DROPDOWN
    assert classify("在B3做个下拉菜单，选项为：甲、乙").kind == INTENT_DROPDOWN
    assert classify("把B3设置下拉，选项是C2:C3").kind == INTENT_DROPDOWN
    assert classify("在B3生成公式并下拉到B9").kind == INTENT_GENERATE
    assert classify("把B3的公式下拉填充到B9").kind == INTENT_GENERATE


# ------------------------------------------------------------------ 表格加框
def test_frame_range_intent_detection():
    """“把范围框起来/加边框”是纯格式需求：本地处理（0 Token），范围归一化为 A14:B18。"""
    intent = classify("将a14到B18的表格范围框起来")
    assert intent.kind == INTENT_FORMAT
    assert intent.table_range == "A14:B18"
    assert classify("把 A14:B18 框起来").table_range == "A14:B18"
    assert classify("a14到B18框住").kind == INTENT_FORMAT
    assert classify("帮我把 A1：B2 画框").table_range == "A1:B2"


def test_frame_intent_without_range_keeps_format():
    """没说范围时也拦在本地：由流水线追问，而不是送去生成公式白跑一轮。"""
    intent = classify("给这个表格加边框")
    assert intent.kind == INTENT_FORMAT
    assert intent.table_range is None


def test_frame_intent_yields_to_generate_actions():
    """带公式生成动作词（统计/求/算…）的需求仍走模型：边框由新建表格流程顺带套用。"""
    assert classify("新建一张表格并加边框，统计各班平均分").kind == INTENT_GENERATE
    assert classify("把A14:B18框起来并算总分").kind == INTENT_GENERATE


def test_frame_intent_not_hijacked_by_negative_phrases():
    """“去掉/取消/不要边框”是删格式，不能被当成加框需求反向操作。"""
    for text in (
        "去掉A14:B18的边框",
        "把A14:B18的边框去掉",
        "取消表格边框",
        "不要加边框",
        "移除边框",
    ):
        assert classify(text).kind != INTENT_FORMAT, text


def test_frame_negative_word_far_away_does_not_block():
    """负向词与加框分属两句时不算删除：“去掉G2那格的旧格式，然后把A1:B2加边框”仍是加框。"""
    intent = classify("去掉G2那格旧的格式，然后把A1:B2加边框")
    assert intent.kind == INTENT_FORMAT
    assert intent.table_range == "A1:B2"


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


def test_ordinal_table_word_is_not_a_sheet_prefix():
    # “表一/表2”是序数引用而不是表名前缀：“来自表一成绩单”里没有表名
    assert classify("成绩来自表一成绩单，请在H2求和").sheet is None


def test_sentence_fragment_after_sheet_word_is_rejected():
    # “班级信息表教室列后面增加一列…”里的“表”是“班级信息”的尾字，不能把后半句当成表名
    intent = classify("在班级信息表教室列后面增加一列班级学生平均成绩，成绩来自表一成绩单")
    assert intent.sheet is None


# ------------------------------------------------------------------ 表名对齐
def test_align_sheet_recovers_target_from_source_mention():
    # 提取器抓出句子片段；原文里“班级信息”是操作目标，“成绩单”带来源标记
    text = "在班级信息表教室列后面增加一列班级学生平均成绩，成绩来自表一成绩单"
    wrong = "教室列后面增加一列班级学"
    assert align_sheet_name(text, wrong, ["成绩单", "班级信息"]) == "班级信息"


def test_align_sheet_prefers_target_hint_over_bare_mention():
    text = "把成绩单的均分写入班级信息的F2"
    assert align_sheet_name(text, "单的均分写入班级信息的F2", ["成绩单", "班级信息"]) == "班级信息"


def test_align_sheet_keeps_valid_and_unknown_names():
    sheets = ["成绩单", "班级信息"]
    assert align_sheet_name("统计成绩单的平均分", "成绩单", sheets) == "成绩单"
    assert align_sheet_name("随便说点什么", "不存在的表", sheets) == "不存在的表"
    assert align_sheet_name("随便说点什么", None, sheets) is None


def test_align_sheet_fills_none_from_text():
    # 提取器没给出表名时，原文提到的唯一真实表可作为默认表（优于“第一张有数据的表”）
    assert align_sheet_name("统计成绩单的平均分", None, ["成绩单", "班级信息"]) == "成绩单"


def test_align_sheet_restores_official_case():
    assert align_sheet_name("统计sheet2的均分", "sheet2", ["Sheet1", "Sheet2"]) == "Sheet2"


def test_align_sheet_ignores_shorter_name_inside_longer_one():
    # “成绩”落在“成绩单”内部，只按长表名计分，不干扰消歧
    text = "统计成绩单里的总分并写到班级信息"
    assert align_sheet_name(text, "抓错的表名", ["成绩", "成绩单", "班级信息"]) == "班级信息"


def test_align_sheet_gives_up_when_ambiguous():
    # 两个表名都出现且没有目标/来源标记：宁可不改，也不能猜错表
    text = "把成绩单的均分整理进班级信息"
    assert align_sheet_name(text, "抓错了", ["成绩单", "班级信息"]) == "抓错了"


def test_empty_input_marks_missing():
    assert "request" in classify("   ").missing


def test_redact_hides_api_key():
    text = "Authorization: Bearer sk-abcdef1234567890 用于调用"
    masked = redact(text)
    assert "sk-abcdef1234567890" not in masked
    assert "***" in masked


def test_redact_hides_key_assignment():
    assert "1234" not in redact("DEEPSEEK_API_KEY=sk-1234567890abcdef")
