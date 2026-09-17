"""函数库系统性测试：把求值器已实现的函数逐族覆盖一遍。

数据基于 sample_xlsx（5 行成绩表 A1:F6）。期望值按 Excel 语义手工核算，
发现实现与 Excel 不一致时以 Excel 为准修实现（如 MOD 负数取模、ISERR 的 #N/A 例外）。
最后一个 section 是护栏测试：新增函数忘了登记白名单或参数个数约束时会直接报红。
"""
from __future__ import annotations

import pytest

from excel_formula.evaluator import UnsupportedFormula, evaluate_formula, format_value
from excel_formula.excel_reader import WorkbookView
from excel_formula.formula_parser import parse_formula


def _eval(path, formula):
    with WorkbookView(path) as view:
        sheets = {name: view.wb_values[name] for name in view.sheet_names}
        return evaluate_formula(parse_formula(formula), sheets, "Sheet1")


def _fmt(path, formula):
    return format_value(_eval(path, formula))


# ------------------------------------------------------------------ 数学族
@pytest.mark.parametrize(
    "formula,expected",
    [
        ("=ABS(-3.5)", "3.5"),
        ("=SQRT(144)", "12"),
        ("=SQRT(-1)", "#NUM!"),
        ("=POWER(2,10)", "1024"),
        ("=POWER(4,0.5)", "2"),
        ("=MOD(10,3)", "1"),
        ("=MOD(-10,3)", "2"),  # Excel 取模结果符号跟随除数
        ("=MOD(10,-3)", "-2"),
        ("=MOD(10,0)", "#DIV/0!"),
        ('=MOD("abc",3)', "#VALUE!"),  # 文本不能直接把求值器炸掉
        ("=INT(87.9)", "87"),
        ("=INT(-87.1)", "-88"),
        ("=TRUNC(87.678)", "87"),
        ("=TRUNC(87.678,2)", "87.67"),
        ("=TRUNC(-87.678,2)", "-87.67"),
        ("=TRUNC(87.678,-1)", "80"),
        ("=SIGN(-5)", "-1"),
        ("=SIGN(0)", "0"),
        ("=ROUND(2.5,0)", "3"),
        ("=ROUND(-2.5,0)", "-3"),
        ("=ROUND(87.674,2)", "87.67"),
        ("=ROUNDUP(87.1,0)", "88"),
        ("=ROUNDDOWN(87.9,0)", "87"),
        ("=EXP(0)", "1"),
        ("=EXP(1000)", "#NUM!"),
        ("=LN(1)", "0"),
        ("=LN(0)", "#NUM!"),
        ("=LOG(100)", "2"),
        ("=LOG(8,2)", "3"),
        ("=LOG(100,1)", "#NUM!"),
        ("=LOG10(1000)", "3"),
        ("=CEILING(87.2,1)", "88"),
        ("=CEILING(2.5,1)", "3"),
        ("=CEILING(-2.5,2)", "-2"),  # 朝 +∞ 方向凑
        ("=CEILING(7.3,0.5)", "7.5"),
        ("=CEILING(5,0)", "#DIV/0!"),
        ("=FLOOR(87.8,1)", "87"),
        ("=FLOOR(-2.5,2)", "-4"),  # 朝 -∞ 方向凑
        ("=FLOOR(7.7,0.5)", "7.5"),
        ("=SUMPRODUCT({1,2,3},{4,5,6})", "32"),
        ("=SUMPRODUCT(B2:F2,B3:F3)", "38248"),
        ("=SUMPRODUCT(B2:F2,B3:F5)", "#VALUE!"),  # 长度不一致必须报错，不能静默截断
    ],
)
def test_math_functions(sample_xlsx, formula, expected):
    assert _fmt(sample_xlsx, formula) == expected


# ------------------------------------------------------------------ 统计族
@pytest.mark.parametrize(
    "formula,expected",
    [
        ("=MEDIAN(B2:F2)", "88"),
        ("=MEDIAN({1,2,3,4})", "2.5"),
        ("=MODE({1,2,2,3})", "2"),
        ("=MODE({1,2,3})", "#N/A"),  # 无重复值 Excel 返回 #N/A
        ("=MODE.SNGL({5,5,7})", "5"),
        ("=STDEV.S(B2:F2)", "6.58027"),
        ("=STDEV.P(B2:F2)", "5.88558"),
        ("=STDEVP(B2:F2)", "5.88558"),
        ("=VAR(B2:F2)", "43.3"),
        ("=VAR.S(B2:F2)", "43.3"),
        ("=VAR.P(B2:F2)", "34.64"),
        ("=RANK(B2,B2:F2)", "4"),  # 降序：95,92,88,85,78
        ("=RANK(B2,B2:F2,1)", "2"),  # 升序：78,85,...
        ("=RANK.EQ(B2,B2:F2)", "4"),
        ("=LARGE(B2:F2,2)", "92"),
        ("=SMALL(B2:F2,2)", "85"),
        ('=MAXIFS(B2:F2,B2:F2,">85")', "95"),
        ('=MINIFS(B2:F2,B2:F2,">85")', "88"),
        ('=MAXIFS(B2:B6,C2:C6,">85")', "90"),
        ('=MINIFS(B2:B6,C2:C6,">=85")', "88"),
        ('=COUNTIFS(B2:B6,">80",C2:C6,"<90")', "3"),
        ('=SUMIFS(B2:B6,C2:C6,">=87")', "90"),
        ('=AVERAGEIF(B2:B6,">80")', "87.6667"),
        ('=AVERAGEIFS(B2:B6,C2:C6,">80")', "86"),
    ],
)
def test_statistics_functions(sample_xlsx, formula, expected):
    assert _fmt(sample_xlsx, formula) == expected


# ------------------------------------------------------------------ 查找族
@pytest.mark.parametrize(
    "formula,expected",
    [
        ("=INDEX(B2:F2,3)", "92"),  # 单行区域单下标按列号解释
        ("=INDEX(B2:F2,1,3)", "92"),
        ("=INDEX(B2:F6,2,3)", "91"),  # 区域第 2 行(English)第 3 列(D)
        ("=INDEX(B2:F2,9)", "#REF!"),
        ('=MATCH("Physics",A2:A6,0)', "3"),
        ("=MATCH(88,B2:F2,0)", "4"),
        ("=MATCH(87,B2:F2,0)", "#N/A"),
        ("=MATCH(88,{78,80,85,88,92},1)", "4"),
        ("=MATCH(88,{92,88,85,80,78},-1)", "2"),
        ('=HLOOKUP("Student2",A1:F2,2,FALSE)', "78"),
        ('=HLOOKUP("Student3",A1:F6,4,FALSE)', "89"),
        ('=HLOOKUP("Student9",A1:F2,2,FALSE)', "#N/A"),
        ('=XLOOKUP("Physics",A2:A6,B2:B6)', "80"),
        ('=XLOOKUP("Chemistry",A2:A6,D2:D6)', "84"),
        ('=XLOOKUP("查无",A2:A6,B2:B6,"无记录")', "无记录"),
        ('=XLOOKUP("查无",A2:A6,B2:B6)', "#N/A"),
        ('=XLOOKUP("Physics",A2:A6,B2:B7)', "#VALUE!"),  # 查找列与返回列长度必须一致
        ('=VLOOKUP("Biology",A2:F6,5,FALSE)', "83"),
        ('=VLOOKUP("Math",A2:F6,6,FALSE)', "95"),
        ('=LOOKUP(2.5,{1,2,3},{"低","中","高"})', "中"),
        ('=LOOKUP(0.5,{1,2,3},{"低","中","高"})', "#N/A"),
    ],
)
def test_lookup_functions(sample_xlsx, formula, expected):
    assert _fmt(sample_xlsx, formula) == expected


# ------------------------------------------------------------------ 逻辑族
@pytest.mark.parametrize(
    "formula,expected",
    [
        ('=IFS(B2>90,"优",B2>80,"良",TRUE,"差")', "良"),
        ('=IFS(1>2,"a",2>3,"b")', "#N/A"),  # 全部条件不命中
        ('=IFS(TRUE,"ok",1/0,"坏")', "ok"),  # 未命中分支不求值
        ('=CHOOSE(2,"金","银","铜")', "银"),
        ('=CHOOSE(9,"金","银")', "#VALUE!"),
        ('=CHOOSE(1,"ok",1/0)', "ok"),  # 未选中分支不求值
        ('=SWITCH(3,1,"一",2,"二",3,"三")', "三"),
        ('=SWITCH(9,1,"一",2,"二","其他")', "其他"),  # 落单参数是默认值
        ('=SWITCH(9,1,"一",2,"二")', "#N/A"),
        ('=SWITCH(2,1,"一",2,"二",1/0,"三")', "二"),  # 命中即返回，后面不求值
        ("=XOR(TRUE,TRUE)", "FALSE"),
        ("=XOR(TRUE,FALSE,TRUE)", "FALSE"),
        ("=XOR(1,0)", "TRUE"),
        ("=IFERROR(1/0,\"兜底\")", "兜底"),
        ('=IFNA(NA(),"缺失")', "缺失"),
        ("=TRUE()", "TRUE"),
        ("=FALSE()", "FALSE"),
        ("=NA()", "#N/A"),
    ],
)
def test_logic_functions(sample_xlsx, formula, expected):
    assert _fmt(sample_xlsx, formula) == expected


# ------------------------------------------------------------------ 文本族
@pytest.mark.parametrize(
    "formula,expected",
    [
        ("=UPPER(\"abc\")", "ABC"),
        ("=LOWER(\"ABC\")", "abc"),
        ('=PROPER("hello world")', "Hello World"),
        ("=LEFT(A4,4)", "Phys"),
        ("=RIGHT(A4,4)", "sics"),
        ("=MID(A4,2,3)", "hys"),
        ("=LEN(A4)", "7"),
        ('=TRIM("  a  ")', "a"),
        ('=CONCAT("订单",A2)', "订单Math"),
        ('=TEXTJOIN("-",TRUE,"a","","b")', "a-b"),
        ('=TEXTJOIN("-",FALSE,"a","","b")', "a--b"),  # FALSE 时不跳过空值
        ('=FIND("i","Physics")', "5"),
        ('=FIND("I","Physics")', "#VALUE!"),  # FIND 区分大小写
        ('=FIND("s","Physics",4)', "4"),
        ('=FIND("x","Physics")', "#VALUE!"),
        ('=SEARCH("I","Physics")', "5"),  # SEARCH 不区分大小写
        ('=SEARCH("?h*","Physics")', "1"),  # 支持 ? 和 * 通配符
        ('=SUBSTITUTE("banana","a","o")', "bonono"),
        ('=SUBSTITUTE("banana","a","o",2)', "banona"),  # 只换第 2 个 a
        ('=SUBSTITUTE("banana","x","o")', "banana"),
        ('=REPLACE("abcdef",2,3,"XY")', "aXYef"),
        ('=REPLACE("abcdef",1,0,"X")', "Xabcdef"),
        ('=REPT("ab",3)', "ababab"),
        ('=REPT("x",-1)', "#VALUE!"),
        ("=CHAR(65)", "A"),
        ("=CHAR(0)", "#VALUE!"),
        ('=CODE("Apple")', "65"),
        ("=CODE(A2)", "77"),  # "Math" 的 M
        ('=VALUE("1,234.5")', "1234.5"),
        ('=VALUE("12%")', "0.12"),
        ('=VALUE("￥88")', "88"),
        ('=VALUE("abc")', "#VALUE!"),
    ],
)
def test_text_functions(sample_xlsx, formula, expected):
    assert _fmt(sample_xlsx, formula) == expected


# ------------------------------------------------------------------ 信息族
@pytest.mark.parametrize(
    "formula,expected",
    [
        ("=ISBLANK(Z99)", "TRUE"),  # 空单元格
        ("=ISBLANK(A2)", "FALSE"),
        ("=ISNUMBER(B2)", "TRUE"),
        ("=ISNUMBER(A2)", "FALSE"),
        ("=ISTEXT(A2)", "TRUE"),
        ("=ISERROR(NA())", "TRUE"),
        ("=ISERROR(1/0)", "TRUE"),
        ("=ISERR(NA())", "FALSE"),  # ISERR 对 #N/A 返回 FALSE
        ("=ISERR(1/0)", "TRUE"),
        ("=ISNA(NA())", "TRUE"),
        ("=ISNA(1/0)", "FALSE"),
        ("=ISEVEN(4)", "TRUE"),
        ("=ISEVEN(-3)", "FALSE"),
        ("=ISODD(7)", "TRUE"),
        ("=ISODD(8)", "FALSE"),
        ("=ROWS(B2:F6)", "5"),
        ("=COLUMNS(B2:F6)", "5"),
    ],
)
def test_information_functions(sample_xlsx, formula, expected):
    assert _fmt(sample_xlsx, formula) == expected


# ------------------------------------------------------------------ 不支持的路径
@pytest.mark.parametrize(
    "formula,match",
    [
        ("=VLOOKUP(85,B2:F6,2)", "精确匹配"),  # 近似匹配本地不实现
        ("=HLOOKUP(85,A1:F6,2)", "精确匹配"),
        ('=XLOOKUP("Physics",A2:A6,B2:B6,"",1)', "精确匹配"),
        ("=INDEX(B2:F6,0,1)", "0 下标"),
        ('=COUNTIF(A2:A6,"*i*")', "通配符"),
        ("=CEILING(5,-1)", "正显著性"),
        ('=SUMIFS(B2:B6,B2:B6,">0",C2:C7,">0")', "多条件区域大小不一致"),
    ],
)
def test_unsupported_paths_raise_unverified(sample_xlsx, formula, match):
    """不支持的语义必须抛 UnsupportedFormula（上层标"未验证"），而不是给个错答案。"""
    with pytest.raises(UnsupportedFormula, match=match):
        _eval(sample_xlsx, formula)


# ------------------------------------------------------------------ 护栏：注册表一致性
def test_evaluator_functions_stay_inside_whitelist():
    """求值器实现的所有函数都必须在白名单内，防止绕过安全边界。"""
    from excel_formula.config import ALLOWED_FUNCTIONS
    from excel_formula.evaluator import _FUNCTIONS

    missing = set(_FUNCTIONS) - ALLOWED_FUNCTIONS
    assert not missing, f"实现了但不在白名单的函数：{sorted(missing)}"


def test_arity_table_covers_implemented_functions():
    """已实现函数都必须有参数个数约束，防止校验器漏检参数个数错误。"""
    from excel_formula.evaluator import _FUNCTIONS
    from excel_formula.validator import ARITY

    missing = set(_FUNCTIONS) - set(ARITY)
    assert not missing, f"已实现但缺参数个数约束的函数：{sorted(missing)}"
