"""本地意图识别与参数抽取（纯规则，不消耗 Token）。

只负责把一句中文指令拆成"做什么 + 用哪些参数"，真正的公式生成仍交给 DeepSeek。
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field

INTENT_GENERATE = "generate"
INTENT_EXPLAIN = "explain"
INTENT_VALIDATE = "validate"
INTENT_DESCRIBE = "describe"
INTENT_FORMAT = "format"
INTENT_DROPDOWN = "dropdown"

_EXPLAIN_WORDS = ("解释", "什么意思", "啥意思", "讲讲", "说明一下", "含义", "作用是", "看不懂")
_VALIDATE_WORDS = ("校验", "验证", "检查", "对不对", "有没有错", "是否正确", "合法吗", "有问题吗")
_DESCRIBE_WORDS = ("结构", "有哪些列", "有什么列", "表头", "概览", "多少行", "几张表", "哪些工作表", "看看表")
_GENERATE_WORDS = ("公式", "计算", "统计", "求和", "求", "算", "填", "写入", "汇总", "平均", "排名", "占比")

_FILE_RE = re.compile(r"[^\s\"'，,；;（）()]+\.xls[xm]", re.IGNORECASE)
_QUOTED_RE = re.compile(r"[\"'“”‘’《》]([^\"'“”‘’《》]+\.xls[xm])[\"'“”‘’《》]", re.IGNORECASE)
# 中文与字母之间没有 \b 边界（中文也算 \w），因此用显式的 ASCII 前后置断言
_CELL_TOKEN = r"(?<![A-Za-z0-9$])(\$?[A-Za-z]{1,3}\$?[1-9][0-9]{0,6})(?![A-Za-z0-9])"
_SHEET_RE = re.compile(
    # 裸“表”要排除两类误伤：“表格下方”里的“表”不是工作表前缀；
    # “来自表一成绩单”“在表2里”的“表一/表2”是序数引用，也不是表名前缀
    r"(?:工作表|sheet|表(?![格一二三四五六七八九十0-9]))\s*[:：]?\s*"
    r"(?:[\"'“”‘’《]([^\"'“”‘’》]{1,31})[\"'“”‘’》]|([A-Za-z0-9_]{1,31})|([\u4e00-\u9fff]{1,12}))",
    re.IGNORECASE,
)
_CELL_RE = re.compile(_CELL_TOKEN)
_RANGE_RE = re.compile(_CELL_TOKEN + r"\s*[:：]\s*" + _CELL_TOKEN)
# 工作表名里不会出现的动作词，命中即说明抓到的是后半句而不是表名
_SHEET_STOP = ("里", "中", "的", "上", "内", "下", "这", "那", "有")
_SHEET_VERBS = (
    "统计", "计算", "汇总", "求", "算", "填", "写", "生成",
    "添加", "增加", "新增", "插入", "解释", "校验", "检查",
)
_TARGET_HINT_RE = re.compile(
    r"(?:写到|写入|放到|放在|填到|填入|存到|在)\s*([A-Za-z]{1,3}[1-9][0-9]{0,6})\s*(?:单元格|格子)?"
)
_FILL_RE = re.compile(
    r"(?:填充到|一直到|下拉到|到)\s*([A-Za-z]{1,3}[1-9][0-9]{0,6})"
)
# 新建表格的识别：创建动词 +（可选的量词“一张/一个/一份”与“新/新的”）+ 表格名词。
# 名词必须紧跟修饰语，避免“生成成绩表的汇总”这类引用已有表名的需求被误判；
# 裸“表”只认“一张/一个/一份表”（允许“分组/对照”等简短定语）或“新建/创建/建表”这两种明确写法。
_NEW_TABLE_RE = re.compile(
    r"(?:新建|创建|新增|增加|生成|制作|整理成|汇总成|做成|建成|建|做)"
    r"\s*(?:"
    r"(?:一[张个份]|[张个份])?\s*(?:新的|新)?\s*(?:表格|汇总表|统计表|明细表|报表|清单|数据表)"
    r"|一[张个份]\s*(?:新的|新)?\s*[^\s，。；;！？!]{0,8}表"
    r"|(?:新建|创建|建)\s*表(?![格头脑条单])"
    r")"
)
# “把范围框起来/加边框”这类纯格式需求：本地直接套细边框（0 Token），不送模型——
# 模型只会给公式，给不了格式。负向写法（去掉/取消边框）不在这里承接。
_FRAME_WORDS = ("框起来", "框住", "加框", "加边框", "加个框", "边框", "画框")
_FRAME_NEG_WORDS = ("去掉", "取消", "删除", "清除", "移除", "撤销", "不要")
# 负向词与边框词相距超过这个字符数才视为互不相干（“去掉旧格式，给A1:B2加边框”）
_FRAME_NEG_GAP = 12
_FRAME_RANGE_RE = re.compile(
    _CELL_TOKEN + r"\s*(?:到|至|[:：]|－|-|—|–|~|～)\s*" + _CELL_TOKEN
)


# 下拉列表关键词：完整名词命中即判；“下拉”简写需配合设置类动词，
# 且不能是“下拉到B9/下拉填充”这类填充语义。
_DROPDOWN_WORDS = ("下拉列表", "下拉框", "下拉选项", "下拉菜单", "数据验证", "数据有效性")
_DROPDOWN_VERBS = ("设置", "设为", "设成", "添加", "创建", "新增", "建立", "做成", "改为", "改成", "加")
_FILL_TAIL_RE = re.compile(r"下拉\s*(?:到|至|填充|复制|拖动)")


def is_dropdown_request(text: str) -> bool:
    """是否是设置下拉列表的请求（纯规则，0 Token）。"""
    if _hit(text, _DROPDOWN_WORDS):
        return True
    if "下拉" not in text or _FILL_TAIL_RE.search(text):
        return False
    return _hit(text, _DROPDOWN_VERBS)


@dataclass
class Intent:
    kind: str
    request: str
    file: str | None = None
    sheet: str | None = None
    target: str | None = None
    fill_to: str | None = None
    formula: str | None = None
    cell: str | None = None
    missing: list[str] = field(default_factory=list)
    # 需求是在“新建一张表”：写入后自动套用表头浅蓝底与整表边框
    new_table: bool = False
    # “框起来”类需求要加边框的目标范围（规范化为 A14:B18 写法），None 表示原文里没给
    table_range: str | None = None

    def to_dict(self) -> dict:
        return {
            "kind": self.kind,
            "file": self.file,
            "sheet": self.sheet,
            "target": self.target,
            "fill_to": self.fill_to,
            "formula": self.formula,
            "cell": self.cell,
            "missing": self.missing,
            "new_table": self.new_table,
            "table_range": self.table_range,
        }


def wants_new_table(text: str) -> bool:
    """需求是否在“新建一张表格”（纯规则，0 Token）。"""
    return bool(_NEW_TABLE_RE.search(text or ""))


def classify(text: str) -> Intent:
    """把自然语言指令解析为 Intent。"""
    raw = (text or "").strip()
    intent = Intent(kind=INTENT_GENERATE, request=raw)
    if not raw:
        intent.missing.append("request")
        return intent

    intent.file = _extract_file(raw)
    intent.formula = _extract_formula(raw)
    keyword_text = raw
    if intent.formula:
        # 公式字符串里的“下拉列表”等只是数据，不改变原来的公式意图。
        keyword_text = raw.replace(intent.formula, re.sub(r'"(?:[^"]|"")*"', '""', intent.formula), 1)
    if is_dropdown_request(keyword_text):
        # 表名和地址需结合工作簿解析；不能沿用公式抽取器把来源表识别为目标。
        intent.kind = INTENT_DROPDOWN
        return intent
    intent.sheet = _extract_sheet(raw)
    intent.new_table = wants_new_table(raw)

    range_match = _RANGE_RE.search(raw)
    target_match = _TARGET_HINT_RE.search(raw)
    if range_match and _looks_like_target_range(raw, range_match):
        intent.target = range_match.group(1).replace("$", "").upper()
        intent.fill_to = range_match.group(2).replace("$", "").upper()
    elif target_match:
        intent.target = target_match.group(1).upper()
        fill_match = _FILL_RE.search(raw[target_match.end():])
        if fill_match:
            intent.fill_to = fill_match.group(1).upper()

    has_formula = bool(intent.formula)
    wants_write = _hit(raw, ("写入", "写到", "填入", "填到", "生成公式", "帮我算"))
    if _hit(raw, _EXPLAIN_WORDS) and not wants_write and (has_formula or _CELL_RE.search(raw)):
        intent.kind = INTENT_EXPLAIN
        intent.cell = intent.target or _first_cell(raw)
    elif _hit(raw, _VALIDATE_WORDS) and has_formula:
        intent.kind = INTENT_VALIDATE
    elif _hit(raw, _DESCRIBE_WORDS) and not _hit(raw, _GENERATE_WORDS):
        intent.kind = INTENT_DESCRIBE
    elif has_formula and not _hit(raw, _GENERATE_WORDS):
        # 用户直接贴了公式又没说要算什么，默认按校验处理，避免误写文件
        intent.kind = INTENT_VALIDATE
    elif _looks_like_frame(raw) and not intent.new_table:
        intent.kind = INTENT_FORMAT
        frame = _FRAME_RANGE_RE.search(raw)
        if frame:
            intent.table_range = (
                f"{frame.group(1)}:{frame.group(2)}".replace("$", "").upper()
            )
    else:
        intent.kind = INTENT_GENERATE

    if intent.kind == INTENT_EXPLAIN and not (intent.formula or intent.cell):
        intent.missing.append("formula")
    if intent.kind == INTENT_VALIDATE and not intent.formula:
        intent.missing.append("formula")
    return intent


# ------------------------------------------------------------------ 抽取工具
def _hit(text: str, words: tuple[str, ...]) -> bool:
    return any(word in text for word in words)


def _looks_like_frame(text: str) -> bool:
    """是否“框起来/加边框”这类纯格式需求（本地规则，0 Token）。

    只承接没有任何公式与生成动作词的需求（如“将 A14 到 B18 的表格范围框起来”）；
    “新建一张表格并加边框，统计……”仍走模型生成（边框由新建表格流程自动套用），
    “去掉/取消边框”这类负向写法也不在这里拦下。
    """
    if not _hit(text, _FRAME_WORDS):
        return False
    if _extract_formula(text) or _hit(text, _GENERATE_WORDS):
        return False
    # 负向词与边框词相距很近（前后都算）时属于同一短语：“去掉A14:B18的边框”
    # 是删除格式，不能被当成加框需求反向操作；相距很远才可能是两句不相干的话。
    for neg in _spans_of(text, _FRAME_NEG_WORDS):
        for frame in _spans_of(text, _FRAME_WORDS):
            if max(neg[0], frame[0]) - min(neg[1], frame[1]) <= _FRAME_NEG_GAP:
                return False
    return True


def _spans_of(text: str, words: tuple[str, ...]) -> list[tuple[int, int]]:
    """给定词表在文本中出现的区间列表（含重叠出现）。"""
    spans: list[tuple[int, int]] = []
    for word in words:
        start = text.find(word)
        while start >= 0:
            spans.append((start, start + len(word)))
            start = text.find(word, start + 1)
    return spans


def _extract_file(text: str) -> str | None:
    quoted = _QUOTED_RE.search(text)
    if quoted:
        return quoted.group(1)
    plain = _FILE_RE.search(text)
    return plain.group(0) if plain else None


def file_candidates(name: str) -> list[str]:
    """中文里“在xxx.xlsx”没有空格分隔，抓到的文件名可能粘上了前面的介词。

    汉字本身也是合法文件名（如 测试数据.xlsx），所以不能简单剔除汉字。
    这里只逐字剥掉开头的汉字生成由长到短的候选，由调用方按文件是否真的存在来挑。
    """
    candidates = [name]
    stem = name
    while re.match(r"[\u4e00-\u9fff]", stem):
        stem = stem[1:]
        if stem.startswith("."):  # 再剥就只剩后缀了
            break
        candidates.append(stem)
    return candidates


def _extract_sheet(text: str) -> str | None:
    for match in _SHEET_RE.finditer(text):
        quoted, ascii_name, chinese = match.groups()
        name = (quoted or ascii_name or chinese or "").strip()
        if chinese and not quoted:
            # 中文表名容易把后半句一起吃进来，遇到停用字或动作词就截断
            for stop in _SHEET_STOP:
                name = name.split(stop)[0]
            # 截断后只剩单个汉字（如“表格下方”切出的“格”）几乎不可能是表名
            if len(name) < 2 or any(verb in name for verb in _SHEET_VERBS):
                continue
        if not name or name.lower().endswith((".xlsx", ".xlsm")):
            continue
        return name
    return None


# ------------------------------------------------------------------ 表名对齐
# 表名出现位置前的窗口里若有这些词，说明它是数据来源而不是要操作的表
_SHEET_SOURCE_HINTS = ("来自", "取自", "源于", "根据")
# 反过来，这些词说明它是操作目标
_SHEET_TARGET_HINTS = (
    "写入", "写到", "填入", "填到", "放到", "放在", "录入", "补充到", "汇总到", "输出到", "加到",
)


def align_sheet_name(text: str, extracted: str | None, sheet_names: list[str]) -> str | None:
    """提取出的表名与真实工作表对不上时，回到原文里找回用户实际提到的表。

    典型场景：“在班级信息表教室列后面增加一列……”——提取器把“班级信息表”的
    “表”当成了前缀，抓出一段句子片段（或干脆没给出表名）。这类情况不值得
    直接报错：原文里往往就写着真实表名。只有结果唯一确定时才替换，否则原样
    返回交给调用方报错/走默认表，宁可报错也不能猜错表——猜错会把公式写到
    错误的表里。
    """
    if extracted:
        for name in sheet_names:
            if extracted == name:
                return extracted
        folded = extracted.casefold()
        for name in sheet_names:
            if folded == name.casefold():
                return name  # 只有大小写对不上：回正成真实表名
    if not text:
        return extracted

    lowered = text.casefold()
    hits: list[tuple[str, int]] = []
    for name in sheet_names:
        key = name.casefold()
        start = lowered.find(key)
        while start >= 0:
            hits.append((name, start))
            start = lowered.find(key, start + 1)
    if not hits:
        return extracted

    # 短表名落在长表名内部（如“成绩”与“成绩单”）时不重复计分
    def covered(name: str, index: int) -> bool:
        end = index + len(name)
        return any(
            other_name != name and other_index <= index and end <= other_index + len(other_name)
            for other_name, other_index in hits
        )

    hits = [hit for hit in hits if not covered(*hit)]
    if not hits:
        return extracted

    scores: dict[str, int] = {}
    for name, index in hits:
        scores[name] = max(scores.get(name, -2), _mention_score(text, index))

    non_negative = [name for name, score in scores.items() if score >= 0]
    if len(non_negative) == 1:
        return non_negative[0]
    positives = [name for name in non_negative if scores[name] > 0]
    if len(positives) == 1:
        return positives[0]
    return extracted


def _mention_score(text: str, index: int) -> int:
    """给表名在原文中的一次出现打分：来源标记 -1，目标标记 +1，其余 0。"""
    window = text[max(0, index - 8): index]
    if any(hint in window for hint in _SHEET_SOURCE_HINTS):
        return -1
    if any(hint in window for hint in _SHEET_TARGET_HINTS):
        return 1
    return 0


def _extract_formula(text: str) -> str | None:
    index = text.find("=")
    if index < 0:
        return None
    candidate = text[index:].strip()
    # 截掉句末的中文标点与说明性文字
    candidate = re.split(r"[，。；、？！\s]{1,}(?=[\u4e00-\u9fff])|[，。；？！]$", candidate)[0].strip()
    if len(candidate) < 3:
        return None
    if not re.match(r"^=\s*[A-Za-z$'\"(\-+0-9]", candidate):
        return None
    return candidate


def _first_cell(text: str) -> str | None:
    match = _CELL_RE.search(text)
    return match.group(1).replace("$", "").upper() if match else None


def _looks_like_target_range(text: str, match: re.Match) -> bool:
    """区分"写到 G2:G6"（目标）与"统计 B2:F2"（数据源）。"""
    prefix = text[max(0, match.start() - 8): match.start()]
    return bool(re.search(r"写|填|放|存|输出|生成", prefix))
