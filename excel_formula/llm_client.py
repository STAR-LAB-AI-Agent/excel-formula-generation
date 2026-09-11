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
8. 目标单元格未给出时，选择数据区右侧或下方第一个空位，并在 assumptions 里说明。
9. 需要用户补充关键信息（例如无法判断统计哪一列）时，把问题写进 clarification 并让 formulas 为空数组。
10. explanation 用一句中文说明公式含义，不超过 60 字。

输出 JSON 结构：
{"formulas": [{"target": "G2", "formula": "=SUM(B2:F2)", "explanation": "求B2到F2的总分",
"fill_to": "G6"}], "assumptions": ["..."], "clarification": null}

fill_to 可选：需要整列向下填充时给出结束单元格，公式会按相对引用自动调整。"""

EXPLAIN_PROMPT = """你是 Excel 公式讲解者。用中文解释给定公式：先一句话说明作用，
再分点说明关键部分（函数、引用区域、判断条件），最后指出常见易错点。总长度不超过 200 字，
不要输出 Markdown 标题，不要复述表格全部数据。"""


class LLMError(RuntimeError):
    """调用模型失败（网络、鉴权、返回格式非法）。"""


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

    def chat(self, messages: list[dict], *, json_mode: bool = True, max_tokens: int = 800) -> ChatResult:
        api_key = self.settings.require_api_key()
        payload: dict = {
            "model": self.settings.model,
            "messages": messages,
            "temperature": self.settings.temperature,
            "max_tokens": max_tokens,
            "stream": False,
        }
        if json_mode:
            payload["response_format"] = {"type": "json_object"}

        headers = {"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"}
        if self.logger:
            self.logger.info(
                "调用模型 model=%s json_mode=%s 消息数=%d 输入字符数=%d",
                self.settings.model,
                json_mode,
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

            elapsed = time.perf_counter() - started
            if response.status_code == 401:
                raise LLMError("DeepSeek 鉴权失败（401），请检查 DEEPSEEK_API_KEY 是否有效")
            if response.status_code == 402:
                raise LLMError("DeepSeek 账户余额不足（402）")
            if response.status_code in {429, 500, 502, 503, 504}:
                last_error = LLMError(f"服务暂时不可用（HTTP {response.status_code}）")
                if self.logger:
                    self.logger.warning("模型返回 %s，准备重试", response.status_code)
                time.sleep(1.5 * attempt)
                continue
            if response.status_code != 200:
                raise LLMError(f"模型返回 HTTP {response.status_code}: {response.text[:200]}")

            try:
                data = response.json()
                content = data["choices"][0]["message"]["content"]
            except (ValueError, KeyError, IndexError) as exc:
                raise LLMError(f"模型返回格式异常：{exc}") from exc

            raw_usage = data.get("usage") or {}
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
            return ChatResult(content=content, usage=usage)

        hint = _PROXY_HINT if isinstance(last_error, requests.exceptions.ProxyError) else ""
        raise LLMError(f"模型调用失败：{last_error}{hint}")


# ------------------------------------------------------------------ 提示词构造
def build_generate_messages(
    digest_text: str,
    request: str,
    *,
    target: str | None = None,
    sheet_names: list[str] | None = None,
) -> list[dict]:
    parts = [f"表格内容:\n{digest_text}"]
    if sheet_names and len(sheet_names) > 1:
        parts.append("工作簿内的工作表: " + ", ".join(sheet_names))
    parts.append(f"目标单元格: {target}" if target else "目标单元格: 未指定，请自行选择并说明理由")
    parts.append(f"用户需求: {request}")
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
    """标准化模型输出，返回 (公式列表, 假设列表, 追问问题)。"""
    raw = data.get("formulas")
    if raw is None and data.get("formula"):  # 兼容模型只给单个公式的情况
        raw = [{
            "target": data.get("target"),
            "formula": data.get("formula"),
            "explanation": data.get("explanation", ""),
            "fill_to": data.get("fill_to"),
        }]
    items: list[dict] = []
    for entry in raw or []:
        if not isinstance(entry, dict):
            continue
        formula = str(entry.get("formula", "")).strip()
        if not formula:
            continue
        items.append(
            {
                "target": str(entry.get("target") or "").strip(),
                "formula": formula,
                "explanation": str(entry.get("explanation") or "").strip(),
                "fill_to": str(entry.get("fill_to") or "").strip() or None,
            }
        )
    assumptions = [str(a) for a in (data.get("assumptions") or []) if str(a).strip()]
    clarification = data.get("clarification")
    clarification = str(clarification).strip() if clarification else None
    return items, assumptions, clarification
