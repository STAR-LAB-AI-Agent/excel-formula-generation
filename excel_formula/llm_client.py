"""DeepSeek API 客户端与提示词构造。

只通过 requests 直连 OpenAI 兼容的 /chat/completions 接口，不引入额外 SDK。
密钥仅存在于请求头，绝不写入日志。
"""
from __future__ import annotations

import json
import re
import time
import urllib.request
from dataclasses import dataclass, field
from typing import Callable
from urllib.parse import urlsplit, urlunsplit

import requests

from .config import Settings

SYSTEM_PROMPT = """你是 Excel 公式专家，负责把中文自然语言需求转换为可直接写入单元格的 Excel 公式。

规则：
1. 只输出 JSON，不要任何解释性文字或 Markdown 代码块。
2. 公式必须以 = 开头，使用标准 A1 引用；跨表引用写成 Sheet1!A1。
3. 只能使用常见 Excel 内置函数；严禁使用 INDIRECT、HYPERLINK、WEBSERVICE 等高风险函数。
4. 公式不得引用目标单元格自身（避免循环引用）。
5. 表格已以 TSV 完整给出（首列是行号，首行是列字母）。请自己判断哪一行是表头、数据从第几行开始；
   不要凭空发明列。
6. 区域的起始行必须是第一个真正的数据行，表头行、标题行、单位说明行都要排除在外。
   即使 SUM、MIN 这类函数会自动忽略文本，COUNTA、ROWS、INDEX、MATCH、SUMPRODUCT 仍会因为多算一行而出错。
   在 assumptions 里写明认定的数据区，例如「数据区 A3:C10，第 2 行是表头」。
7. 表中可能同时存在多个数据块（例如主数据区旁边另有说明或参考区），请按列位置分辨，不要混用。
8. 反向查找（要返回的列在查找列左侧）不能用普通 VLOOKUP，但可以用数组常量翻转列顺序：
   =VLOOKUP(查找值,IF({1,0},查找列,返回列),2,0)。
   写之前先在 TSV 里核对列的字母位置（首行列字母就是列标），不要想当然地猜列号。
9. 目标单元格未给出时，选择数据区右侧或下方第一个空位，并在 assumptions 里说明。
10. 需要用户补充关键信息（例如无法判断统计哪一列）时，把问题写进 clarification 并让 formulas 为空数组。
11. explanation 用一句中文说明公式含义，不超过 60 字。
12. 需求包含分类名、标题等已知常量时（例如新建汇总表要写入区域名、班级名、列标题），
    用 value 字段把这些常量直接写进单元格，不要留给用户手动填；也不要写成 ="华东" 这类“公式化的常量”。
13. 目标表是写入位置，其他工作表内容仅供跨表引用。每个引用都须核对对应表的 TSV 列字母和数据区；
    不得把目标表的列布局套到来源表，也不得臆造 A2:B100 等范围。空白的平均分列不是已有成绩。
14. 表格可能因规模限制而省略部分行、列或工作表；缺少必要来源结构或统计口径不明确时，
    返回 clarification，不得猜测列号。均分须明确科目范围、权重及空成绩处理，在 assumptions 中说明。
15. 本地试算报告 #DIV/0!、#VALUE!、#N/A 等错误时，先核对真实数据与引用；
    不得仅用 IFERROR(...,0) 掩盖引用错误。确实无数据且需求未说明处理方式时请求澄清。
16. 需求是新建一张表（写表头行并填充数据）时，在 JSON 顶层加 table 字段给出几何：
    {"header": "A2:D2", "range": "A2:D10"}。header 是表头行、必须是 range 的第一行，
    标题、说明行不要写进 range。写入工具会据此给表头铺浅蓝底、给整表范围加边框。
    仅给已有表增加一列/一行时不要给 table（输出 null 或省略）。

输出 JSON 结构（table 仅新建表格时给出）：
{"formulas": [{"target": "G2", "formula": "=SUM(B2:F2)", "explanation": "求B2到F2的总分",
"fill_to": "G6"}, {"target": "A24", "value": "华东", "explanation": "区域名称"}],
"table": {"header": "A24:D24", "range": "A24:D28"}, "assumptions": ["..."], "clarification": null}

每个元素 formula 与 value 二选一：formula 为以 = 开头的公式；value 为直接写入的文本或数字。
fill_to 可选：需要整列向下填充时给出结束单元格，公式会按相对引用自动调整（value 不支持 fill_to，多个常量逐个列出）。"""

EXPLAIN_PROMPT = """你是 Excel 公式讲解者。用中文解释给定公式：先一句话说明作用，
再分点说明关键部分（函数、引用区域、判断条件），最后指出常见易错点。总长度不超过 200 字，
不要输出 Markdown 标题，不要复述表格全部数据。"""

DROPDOWN_PROMPT = """你是 Excel 下拉列表（数据验证）助手。本地规则没能从用户原句解析出设置指令，
请读懂原句并给出结构化参数。

规则：
1. 只输出 JSON，不要任何解释文字或 Markdown 代码块。
2. 先判断 intent：创建/设置下拉列表为 "create"；删除、解释、询问或与下拉无关为 "remove"、
   "query" 或 "other"，并在 clarification 里说明本系统只支持创建下拉列表。
3. intent 为 "create" 时给出三部分：
   - target：目标单元格或区域（如 "B3"、"C2:C13"），必须是需求里明确提到的位置；
   - sheet：目标工作表名；只有需求把表名明确连到目标单元格（如“订单查询的B3”“在订单查询里给B3设置…”）
     才填写，照抄表名；表名只是数据来源（如“选项用销售订单的订单号”）时必须留空，使用默认工作表；
   - 来源三选一：{"source": {"sheet": "表名", "range": "A2:A25"}} 引用某表的区域；
     {"source": {"sheet": "表名", "header": "列标题"}} 引用某个列标题对应列的数据；
     {"options": ["甲", "乙"]} 使用需求里直接列举的选项。
4. 表名与列标题必须真实存在于给出的表格内容中，不得凭空发明；用户的口语说法
   （如“订单编号”）要映射到表格里实际的列标题（如“订单号”）。
5. 无法确定 target 或来源时，把问题写进 clarification 并省略不确定的字段，不要猜测。
6. 只支持静态区域、列标题或固定选项；不要输出公式、名称或 INDIRECT 等动态来源。

输出 JSON 结构（不需要的字段用 null 或省略）：
{"intent": "create", "sheet": "订单查询", "target": "B3",
 "source": {"sheet": "销售订单", "header": "订单号", "range": null},
 "options": null, "clarification": null}"""


class LLMError(RuntimeError):
    """调用模型失败（网络、鉴权、返回格式非法）。

    ``usage`` 可选：当失败发生在一次已计费的调用上（如输出被截断），
    携上用量以便上层如实统计，不漏算这次 Token。
    """

    def __init__(self, *args, usage: "Usage | None" = None):
        super().__init__(*args)
        self.usage = usage


# 本机代理软件的监听地址，几乎都是明文 HTTP 而非 HTTPS
_LOOPBACK_HOSTS = {"127.0.0.1", "localhost", "::1", "0.0.0.0"}

_PROXY_HINT = (
    "。代理不可用：请确认代理软件正在运行，或在 .env 里设置 "
    "EXCELCR_PROXY=http://127.0.0.1:端口 指定代理，或设置 EXCELCR_TRUST_ENV=0 直连"
)


def normalize_proxies(raw: dict) -> dict:
    """修正系统代理里 https:// 指向本机代理的写法。

    Windows 的系统代理在注册表里只存 host:port，Python 会给 https 通道补成
    ``https://127.0.0.1:端口``，requests 于是对本地明文代理发起 TLS 握手，
    报出 ProxyError('Unable to connect to proxy')。这里统一降级为 http://。
    """
    fixed: dict = {}
    for scheme, url in (raw or {}).items():
        if not url or "://" not in url:
            fixed[scheme] = url
            continue
        parts = urlsplit(url)
        if parts.scheme == "https" and parts.hostname in _LOOPBACK_HOSTS:
            parts = parts._replace(scheme="http")
        fixed[scheme] = urlunsplit(parts)
    return fixed


@dataclass
class Usage:
    prompt_tokens: int = 0
    completion_tokens: int = 0
    calls: int = 0
    elapsed: float = 0.0

    @property
    def total_tokens(self) -> int:
        return self.prompt_tokens + self.completion_tokens

    def add(self, other: "Usage") -> None:
        self.prompt_tokens += other.prompt_tokens
        self.completion_tokens += other.completion_tokens
        self.calls += other.calls
        self.elapsed += other.elapsed

    def to_dict(self) -> dict:
        return {
            "calls": self.calls,
            "prompt_tokens": self.prompt_tokens,
            "completion_tokens": self.completion_tokens,
            "total_tokens": self.total_tokens,
            "elapsed_seconds": round(self.elapsed, 2),
        }


@dataclass
class ChatResult:
    content: str
    usage: Usage = field(default_factory=Usage)


class DeepSeekClient:
    """OpenAI 兼容的 Chat Completions 客户端。"""

    def __init__(self, settings: Settings, logger=None):
        self.settings = settings
        self.logger = logger
        self.total_usage = Usage()
        self._session = requests.Session()
        self._proxy_bypassed = False
        # 流式请求会先带 stream_options 索取 usage；服务端不接受时自动降级
        self._stream_usage_supported = True
        self._configure_proxies()

    def _configure_proxies(self) -> None:
        """显式接管代理，不让 requests 直接照抄系统里写错 scheme 的代理。"""
        if self.settings.proxy:
            self._session.trust_env = False
            self._session.proxies = {"http": self.settings.proxy, "https": self.settings.proxy}
        elif not self.settings.trust_env:
            self._bypass_proxy()
            return
        else:
            self._session.proxies = normalize_proxies(urllib.request.getproxies())
        if self.logger and self._session.proxies:
            self.logger.info("使用代理 %s", self._session.proxies.get("https") or "系统代理")

    def _bypass_proxy(self) -> None:
        """放弃代理改直连（DeepSeek 接口在国内可直连）。"""
        self._proxy_bypassed = True
        self._session.trust_env = False
        self._session.proxies = {}

    def chat(
        self,
        messages: list[dict],
        *,
        json_mode: bool = True,
        max_tokens: int = 8000,
        stream: bool | None = None,
        on_delta: Callable[[str, str], None] | None = None,
        thinking: str | None = None,
    ) -> ChatResult:
        """max_tokens 同时限制思考与正文，思考型模型的实际长度波动很大。

        给足额度：只有真正用完才计费，额度不够反而会在推理中途被截断、白花一次调用。

        stream 为 None 时按调用方意图推断：传了 on_delta（逐段回调，用于流式进度展示）
        就走 SSE 流式，否则保持一次性返回，既有调用方式零改动。

        thinking 为 None 时用配置值；传入 enabled/disabled 可逐次覆盖（auto 档由
        流水线按需求复杂度决定，见 FormulaService）。
        """
        api_key = self.settings.require_api_key()
        use_stream = (on_delta is not None) if stream is None else stream
        payload: dict = {
            "model": self.settings.model,
            "messages": messages,
            "temperature": self.settings.temperature,
            "max_tokens": max_tokens,
            "stream": use_stream,
        }
        if json_mode:
            payload["response_format"] = {"type": "json_object"}
        effective_thinking = thinking if thinking is not None else self.settings.thinking
        if effective_thinking == "auto":
            effective_thinking = ""  # 调用方未给出具体档位时按服务端默认
        if effective_thinking:
            payload["thinking"] = {"type": effective_thinking}
        if self.settings.reasoning_effort:
            payload["reasoning_effort"] = self.settings.reasoning_effort
        if use_stream and self._stream_usage_supported:
            # 不显式索取时，流式的最后一个分片不会带 usage，Token 统计会少算
            payload["stream_options"] = {"include_usage": True}

        headers = {"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"}
        if self.logger:
            self.logger.info(
                "调用模型 model=%s json_mode=%s 流式=%s 消息数=%d 输入字符数=%d",
                self.settings.model,
                json_mode,
                use_stream,
                len(messages),
                sum(len(m.get("content", "")) for m in messages),
            )

        last_error: Exception | None = None
        attempt, max_attempts = 0, 2  # 网络层重试，与"公式修复重试"是两件事
        while attempt < max_attempts:
            attempt += 1
            started = time.perf_counter()
            try:
                response = self._session.post(
                    self.settings.base_url,
                    headers=headers,
                    data=json.dumps(payload, ensure_ascii=False).encode("utf-8"),
                    timeout=self.settings.timeout,
                    stream=use_stream,
                )
            except requests.RequestException as exc:
                last_error = exc
                if self.logger:
                    self.logger.warning("模型请求异常（第 %d 次）：%s", attempt, exc)
                # 代理连不上时立刻改直连再试一次，不占用原有的重试次数
                if isinstance(exc, requests.exceptions.ProxyError) and not self._proxy_bypassed:
                    if self.logger:
                        self.logger.warning("代理不可用，改为直连重试")
                    self._bypass_proxy()
                    max_attempts += 1
                    continue
                time.sleep(1.5 * attempt)
                continue

            action, status_error = self._classify_status(response)
            if action != "ok":
                response.close()
            if action == "drop_stream_options" and "stream_options" in payload:
                # 服务端不认 stream_options：去掉它再试，宁愿少算用量也不阻塞用户
                payload.pop("stream_options", None)
                self._stream_usage_supported = False
                if self.logger:
                    self.logger.warning("服务端不接受 stream_options，改为不带用量参数重试")
                max_attempts += 1
                continue
            if action == "retry":
                last_error = status_error
                if self.logger:
                    self.logger.warning("模型返回 %s，准备重试", response.status_code)
                time.sleep(1.5 * attempt)
                continue
            if action != "ok":
                raise status_error  # 鉴权失败、余额不足等直接抛给上层

            if use_stream:
                try:
                    content, raw_usage, finish_reason, ttft = self._read_stream(response, on_delta)
                except requests.RequestException as exc:
                    last_error = exc
                    if self.logger:
                        self.logger.warning("流式响应读取中断（第 %d 次）：%s", attempt, exc)
                    time.sleep(1.5 * attempt)
                    continue
                finally:
                    response.close()
            else:
                try:
                    data = response.json()
                    choice = data["choices"][0]
                    content = choice["message"]["content"]
                    finish_reason = choice.get("finish_reason")
                except (ValueError, KeyError, IndexError) as exc:
                    raise LLMError(f"模型返回格式异常：{exc}") from exc
                raw_usage = data.get("usage") or {}
                ttft = None

            elapsed = time.perf_counter() - started
            usage = Usage(
                prompt_tokens=int(raw_usage.get("prompt_tokens", 0)),
                completion_tokens=int(raw_usage.get("completion_tokens", 0)),
                calls=1,
                elapsed=elapsed,
            )
            self.total_usage.add(usage)
            if self.logger:
                self.logger.info(
                    "模型返回 tokens(in/out)=%d/%d 耗时=%.2fs",
                    usage.prompt_tokens,
                    usage.completion_tokens,
                    elapsed,
                )
                if ttft is not None:
                    self.logger.info("流式首包 %.2fs", ttft)
                elif use_stream and not raw_usage:
                    self.logger.warning("流式响应未携带 usage，本次 Token 用量按 0 计")

            # 思考型模型的推理过程也计入 completion_tokens，额度用尽时 content 会是空串，
            # 直接交给 parse_json_payload 只会报"未返回 JSON"，掩盖真实原因
            if finish_reason == "length":
                raise LLMError(
                    f"模型输出被 max_tokens({max_tokens}) 截断，已消耗 "
                    f"{usage.completion_tokens} 输出 tokens 但没拿到完整结果。"
                    "请把需求说得更具体，或缩小表格范围后重试",
                    usage=usage,
                )
            if not (content or "").strip():
                raise LLMError(
                    f"模型返回了空内容（finish_reason={finish_reason}）", usage=usage
                )
            return ChatResult(content=content, usage=usage)

        hint = _PROXY_HINT if isinstance(last_error, requests.exceptions.ProxyError) else ""
        raise LLMError(f"模型调用失败：{last_error}{hint}")

    # -------------------------------------------------------------- 流式与状态处理
    @staticmethod
    def _classify_status(response) -> tuple[str, LLMError | None]:
        """把响应状态码归类：ok / retry / drop_stream_options / 直接抛出。"""
        code = response.status_code
        if code == 200:
            return "ok", None
        if code == 401:
            return "raise", LLMError("DeepSeek 鉴权失败（401），请检查 DEEPSEEK_API_KEY 是否有效")
        if code == 402:
            return "raise", LLMError("DeepSeek 账户余额不足（402）")
        if code in {429, 500, 502, 503, 504}:
            return "retry", LLMError(f"服务暂时不可用（HTTP {code}）")
        body = response.text[:200]
        if code == 400 and "stream_options" in body:
            return "drop_stream_options", LLMError(f"模型返回 HTTP 400: {body}")
        return "raise", LLMError(f"模型返回 HTTP {code}: {body}")

    def _read_stream(self, response, on_delta) -> tuple[str, dict, str | None, float | None]:
        """读取 SSE 流：拼回正文、收集 usage 与 finish_reason，并返回首包耗时。

        on_delta(piece, kind) 会分别收到思考片段（reasoning）与正文片段（content），
        由调用方决定怎么展示；中途断网由 requests 抛 RequestException，交给上层重试。
        """
        parts: list[str] = []
        raw_usage: dict = {}
        finish_reason: str | None = None
        started = time.perf_counter()
        ttft: float | None = None
        for line in response.iter_lines(decode_unicode=True):
            if not line:
                continue
            chunk_text = line[5:].strip() if line.startswith("data:") else line.strip()
            if not chunk_text or chunk_text == "[DONE]":
                continue
            try:
                chunk = json.loads(chunk_text)
            except json.JSONDecodeError:
                continue  # 心跳或注释行
            if chunk.get("usage"):
                raw_usage = chunk["usage"]
            choices = chunk.get("choices") or []
            if not choices:
                continue
            choice = choices[0]
            delta = choice.get("delta") or {}
            thinking = delta.get("reasoning_content") or ""
            piece = delta.get("content") or ""
            if (thinking or piece) and ttft is None:
                ttft = time.perf_counter() - started
            if thinking and on_delta is not None:
                on_delta(thinking, "reasoning")
            if piece:
                parts.append(piece)
                if on_delta is not None:
                    on_delta(piece, "content")
            if choice.get("finish_reason"):
                finish_reason = choice["finish_reason"]
        return "".join(parts), raw_usage, finish_reason, ttft


# ------------------------------------------------------------------ 提示词构造
def build_generate_messages(
    digest_text: str,
    request: str,
    *,
    target: str | None = None,
    sheet_names: list[str] | None = None,
    source_digests: dict[str, str] | None = None,
    omitted_sheets: list[str] | None = None,
    new_table: bool = False,
) -> list[dict]:
    parts = [f"目标工作表（所有 target 均写入此表）\n表格内容:\n{digest_text}"]
    if sheet_names and len(sheet_names) > 1:
        parts.append("工作簿内的工作表: " + ", ".join(sheet_names))
    for name, text in (source_digests or {}).items():
        parts.append(f"来源工作表（仅供引用）: {name}\n{text}")
    if omitted_sheets:
        parts.append(
            "因上下文预算未提供内容的工作表: " + ", ".join(omitted_sheets)
            + "。若需求依赖这些表且现有信息不足，请返回 clarification，不要猜测列号。"
        )
    parts.append(f"目标单元格: {target}" if target else "目标单元格: 未指定，请自行选择并说明理由")
    parts.append(f"用户需求: {request}")
    if new_table:
        # 本地规则已识别出“新建表格”，明确提醒模型按规则给出 table 几何
        parts.append(
            "注意：这条需求是在新建一张表格，请按规则在 JSON 顶层给出 table 字段"
            "（表头 header 与整表 range，header 是 range 的第一行）。"
        )
    return [
        {"role": "system", "content": SYSTEM_PROMPT},
        {"role": "user", "content": "\n".join(parts)},
    ]


def build_repair_message(formulas: list[str] | str, errors: list[str]) -> dict:
    """校验失败后的修复请求：只回传公式与错误，不重复发送表格内容。"""
    items = [formulas] if isinstance(formulas, str) else list(formulas)
    label = "上一个公式" if len(items) == 1 else "上一批公式"
    return {
        "role": "user",
        "content": (
            f"{label} {'; '.join(items)} 未通过本地校验，错误如下：\n"
            + "\n".join(f"- {e}" for e in errors)
            + "\n请修正后重新输出同样结构的 JSON（包含全部目标单元格），只改必要的部分。"
        ),
    }


def build_explain_messages(formula: str, digest_text: str | None = None) -> list[dict]:
    user = f"公式: {formula}"
    if digest_text:
        user += f"\n所在表格内容（仅供参考）:\n{digest_text}"
    return [
        {"role": "system", "content": EXPLAIN_PROMPT},
        {"role": "user", "content": user},
    ]


def build_dropdown_messages(
    context_text: str,
    request: str,
    *,
    default: str | None = None,
    sheet_names: list[str] | None = None,
) -> list[dict]:
    """下拉兜底理解的上下文：目标表与来源表的 TSV，附上用户原句。"""
    parts = [f"工作簿内容:\n{context_text}"]
    if sheet_names:
        parts.append("工作表列表: " + ", ".join(sheet_names))
    if default:
        parts.append(f"默认工作表（需求未指定目标表时使用）: {default}")
    parts.append(f"用户需求: {request}")
    return [
        {"role": "system", "content": DROPDOWN_PROMPT},
        {"role": "user", "content": "\n".join(parts)},
    ]


# ------------------------------------------------------------------ 返回值解析
_JSON_BLOCK = re.compile(r"\{.*\}", re.DOTALL)


def parse_json_payload(content: str) -> dict:
    """从模型返回中提取 JSON；兼容偶发的 Markdown 包裹。"""
    text = (content or "").strip()
    if text.startswith("```"):
        text = re.sub(r"^```[a-zA-Z]*\s*|\s*```$", "", text).strip()
    try:
        data = json.loads(text)
    except json.JSONDecodeError:
        match = _JSON_BLOCK.search(text)
        if not match:
            raise LLMError(f"模型未返回 JSON：{text[:120]}")
        try:
            data = json.loads(match.group(0))
        except json.JSONDecodeError as exc:
            raise LLMError(f"模型返回的 JSON 无法解析：{exc}") from exc
    if not isinstance(data, dict):
        raise LLMError("模型返回的 JSON 顶层不是对象")
    return data


def extract_candidates(data: dict) -> tuple[list[dict], list[str], str | None]:
    """标准化模型输出，返回 (公式列表, 假设列表, 追问问题)。

    列表元素 formula 与 value 二选一：formula 为公式（以 = 开头），
    value 为直接写入的常量（分类名、标题等）。
    """
    raw = data.get("formulas")
    if raw is None and (data.get("formula") or data.get("value") is not None):
        # 兼容模型只给单个目标的情况
        raw = [{
            "target": data.get("target"),
            "formula": data.get("formula"),
            "value": data.get("value"),
            "explanation": data.get("explanation", ""),
            "fill_to": data.get("fill_to"),
        }]
    items: list[dict] = []
    for entry in raw or []:
        if not isinstance(entry, dict):
            continue
        formula = str(entry.get("formula") or "").strip()
        value = entry.get("value")
        if not formula and value is None:
            continue
        items.append(
            {
                "target": str(entry.get("target") or "").strip(),
                "formula": formula,
                "value": value,
                "explanation": str(entry.get("explanation") or "").strip(),
                "fill_to": str(entry.get("fill_to") or "").strip() or None,
            }
        )
    assumptions = [str(a) for a in (data.get("assumptions") or []) if str(a).strip()]
    clarification = data.get("clarification")
    clarification = str(clarification).strip() if clarification else None
    return items, assumptions, clarification


def _clean_range_text(value: object) -> str | None:
    """清洗模型给出的坐标文本：去空格与 $、转大写；空值返回 None。"""
    text = re.sub(r"[\s$]", "", str(value or "")).upper()
    return text or None


def extract_table_spec(data: dict) -> dict | None:
    """提取“新建表格”标记与几何，返回 {"header": str|None, "range": str|None}。

    兼容三种写法：table 为对象（推荐，给出表头行与整表范围）、table 为 true（仅标记）、
    顶层 new_table 字段。坐标只做清洗，合法性（越界、与写入不匹配等）由 pipeline 校验。
    未标记新建表格时返回 None。
    """
    raw = data.get("table")
    if raw is None:
        raw = data.get("new_table")
    if raw is None or raw is False:
        return None
    if raw is True:
        return {"header": None, "range": None}
    if isinstance(raw, dict):
        header = _clean_range_text(raw.get("header"))
        table_range = _clean_range_text(raw.get("range"))
        if header is None and table_range is None:
            return None
        return {"header": header, "range": table_range}
    return None


def extract_dropdown_spec(data: dict) -> dict | None:
    """从模型输出中提取下拉创建参数；只接受 intent 明确为 "create" 的结果。

    返回 {"sheet": str|None, "target": str, "source": str|None, "options": list|None,
    "source_sheet": str|None}：source 会拼成“表名的列标题”或“表名!区域”表达式，
    表、列、区域的真实性仍由 data_validation 的严格校验把关，这里只做形态清洗与二选一取舍；
    source_sheet 供兜底流程识别“模型把来源表填进 sheet”的情况。
    """
    if str(data.get("intent") or "").strip().lower() != "create":
        return None
    target = _clean_range_text(data.get("target")) or ""
    target = re.sub(r"(?:到|至)", ":", target)
    if not target:
        return None
    sheet = str(data.get("sheet") or "").strip() or None
    options: list[str] | None = None
    raw_options = data.get("options")
    if isinstance(raw_options, list):
        cleaned = [str(item).strip() for item in raw_options
                   if item is not None and str(item).strip()]
        options = cleaned or None
    source_text: str | None = None
    source_sheet: str | None = None
    raw_source = data.get("source")
    if isinstance(raw_source, dict):
        src_sheet = str(raw_source.get("sheet") or "").strip()
        source_sheet = src_sheet or None
        header = re.sub(r"[\s　]+", "", str(raw_source.get("header") or ""))
        area = _clean_range_text(raw_source.get("range"))
        if src_sheet and header:
            source_text = f"{src_sheet}的{header}"
        elif src_sheet and area:
            source_text = f"{src_sheet}!{area}"
        elif area:
            source_text = area
        elif src_sheet:
            source_text = src_sheet
    elif isinstance(raw_source, str) and raw_source.strip():
        source_text = raw_source.strip()
    if source_text is None and options is None:
        return None
    if source_text is not None:
        options = None  # 同时给出时优先区域/列标题来源：它能被本地严格校验
    return {"sheet": sheet, "target": target, "source": source_text,
            "options": options, "source_sheet": source_sheet}
