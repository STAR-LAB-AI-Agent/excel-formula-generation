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

_EXPLAIN_WORDS = ("解释", "什么意思", "啥意思", "讲讲", "说明一下", "含义", "作用是", "看不懂")
_VALIDATE_WORDS = ("校验", "验证", "检查", "对不对", "有没有错", "是否正确", "合法吗", "有问题吗")
_DESCRIBE_WORDS = ("结构", "有哪些列", "有什么列", "表头", "概览", "多少行", "几张表", "哪些工作表", "看看表")
_GENERATE_WORDS = ("公式", "计算", "统计", "求和", "求", "算", "填", "写入", "汇总", "平均", "排名", "占比")

_FILE_RE = re.compile(r"[^\s\"'，,；;（）()]+\.xls[xm]", re.IGNORECASE)
_QUOTED_RE = re.compile(r"[\"'“”‘’《》]([^\"'“”‘’《》]+\.xls[xm])[\"'“”‘’《》]", re.IGNORECASE)
# 中文与字母之间没有 \b 边界（中文也算 \w），因此用显式的 ASCII 前后置断言
_CELL_TOKEN = r"(?<![A-Za-z0-9$])(\$?[A-Za-z]{1,3}\$?[1-9][0-9]{0,6})(?![A-Za-z0-9])"
_SHEET_RE = re.compile(
    r"(?:工作表|sheet|表)\s*[:：]?\s*"
    r"(?:[\"'“”‘’《]([^\"'“”‘’》]{1,31})[\"'“”‘’》]|([A-Za-z0-9_]{1,31})|([\u4e00-\u9fff]{1,12}))",
    re.IGNORECASE,
)
_CELL_RE = re.compile(_CELL_TOKEN)
_RANGE_RE = re.compile(_CELL_TOKEN + r"\s*[:：]\s*" + _CELL_TOKEN)
# 工作表名里不会出现的动作词，命中即说明抓到的是后半句而不是表名
_SHEET_STOP = ("里", "中", "的", "上", "内", "下", "这", "那", "有")
_SHEET_VERBS = ("统计", "计算", "求", "算", "填", "写", "生成", "添加", "解释", "校验", "检查")
_TARGET_HINT_RE = re.compile(
    r"(?:写到|写入|放到|放在|填到|填入|存到|在)\s*([A-Za-z]{1,3}[1-9][0-9]{0,6})\s*(?:单元格|格子)?"
)
_FILL_RE = re.compile(
    r"(?:填充到|一直到|下拉到|到)\s*([A-Za-z]{1,3}[1-9][0-9]{0,6})"
)


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
        }


def classify(text: str) -> Intent:
    """把自然语言指令解析为 Intent。"""
    raw = (text or "").strip()
    intent = Intent(kind=INTENT_GENERATE, request=raw)
    if not raw:
        intent.missing.append("request")
        return intent

    intent.file = _extract_file(raw)
    intent.sheet = _extract_sheet(raw)
    intent.formula = _extract_formula(raw)

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


def _extract_file(text: str) -> str | None:
    quoted = _QUOTED_RE.search(text)
    if quoted:
        return quoted.group(1)
    plain = _FILE_RE.search(text)
    return plain.group(0) if plain else None


def _extract_sheet(text: str) -> str | None:
    for match in _SHEET_RE.finditer(text):
        quoted, ascii_name, chinese = match.groups()
        name = (quoted or ascii_name or chinese or "").strip()
        if chinese and not quoted:
            # 中文表名容易把后半句一起吃进来，遇到停用字或动作词就截断
            for stop in _SHEET_STOP:
                name = name.split(stop)[0]
            if not name or any(verb in name for verb in _SHEET_VERBS):
                continue
        if not name or name.lower().endswith((".xlsx", ".xlsm")):
            continue
        return name
    return None


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
