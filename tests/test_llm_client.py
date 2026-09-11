"""DeepSeek 客户端测试：用假 Session 替代网络，覆盖鉴权失败、重试、JSON 解析。"""
from __future__ import annotations

import json

import pytest
import requests

from excel_formula.config import ConfigError, Settings
from excel_formula.llm_client import (
    DeepSeekClient,
    LLMError,
    build_generate_messages,
    build_repair_message,
    extract_candidates,
    normalize_proxies,
    parse_json_payload,
)


class FakeResponse:
    def __init__(self, status_code: int, payload: dict | None = None, text: str = ""):
        self.status_code = status_code
        self._payload = payload
        self.text = text or json.dumps(payload or {})

    def json(self):
        if self._payload is None:
            raise ValueError("no json")
        return self._payload


class FakeSession:
    """按顺序返回预置响应，并记录请求头/请求体。"""

    def __init__(self, responses: list[FakeResponse], errors: list[Exception] | None = None):
        self.responses = list(responses)
        self.errors = list(errors or [])
        self.requests: list[dict] = []
        self.proxies: dict = {"https": "http://127.0.0.1:10808"}
        self.trust_env = True

    def post(self, url, headers=None, data=None, timeout=None):
        if self.errors:
            raise self.errors.pop(0)
        self.requests.append({"url": url, "headers": headers or {}, "data": data})
        return self.responses.pop(0)


def _client(responses, *, errors=None, **kwargs) -> DeepSeekClient:
    settings = Settings(api_key="sk-test", **kwargs)
    client = DeepSeekClient(settings)
    client._session = FakeSession(responses, errors)
    return client


def _ok_response(content: str) -> FakeResponse:
    return FakeResponse(
        200,
        {
            "choices": [{"message": {"content": content}}],
            "usage": {"prompt_tokens": 300, "completion_tokens": 50},
        },
    )


def test_chat_parses_content_and_usage():
    client = _client([_ok_response('{"formulas": []}')])
    result = client.chat([{"role": "user", "content": "hi"}])

    assert result.content == '{"formulas": []}'
    assert result.usage.total_tokens == 350
    assert client.total_usage.calls == 1
    request = client._session.requests[0]
    assert request["headers"]["Authorization"] == "Bearer sk-test"
    body = json.loads(request["data"].decode("utf-8"))
    assert body["response_format"] == {"type": "json_object"}
    assert body["stream"] is False


def test_missing_api_key_raises_config_error():
    client = DeepSeekClient(Settings(api_key=""))
    with pytest.raises(ConfigError):
        client.chat([{"role": "user", "content": "hi"}])


def test_auth_failure_message():
    client = _client([FakeResponse(401, {"error": "bad key"})])
    with pytest.raises(LLMError, match="鉴权失败"):
        client.chat([])


def test_retries_on_server_error_then_succeeds():
    client = _client([FakeResponse(500, {}), _ok_response('{"formulas": []}')])
    result = client.chat([])
    assert result.content == '{"formulas": []}'
    assert len(client._session.requests) == 2


def test_broken_json_response_raises():
    client = _client([FakeResponse(200, {"unexpected": True})])
    with pytest.raises(LLMError, match="格式异常"):
        client.chat([])


# ------------------------------------------------------------------ 代理处理
def test_normalize_proxies_downgrades_local_https_proxy():
    fixed = normalize_proxies(
        {
            "http": "http://127.0.0.1:10808",
            "https": "https://127.0.0.1:10808",  # Windows 系统代理会写成这样
            "ftp": "ftp://127.0.0.1:10808",
        }
    )
    assert fixed["https"] == "http://127.0.0.1:10808"
    assert fixed["http"] == "http://127.0.0.1:10808"


def test_normalize_proxies_keeps_remote_https_proxy():
    fixed = normalize_proxies({"https": "https://proxy.example.com:8443"})
    assert fixed["https"] == "https://proxy.example.com:8443"


def test_trust_env_off_means_direct_connection():
    client = DeepSeekClient(Settings(api_key="sk-test", trust_env=False))
    assert client._session.proxies == {}
    assert client._session.trust_env is False


def test_explicit_proxy_setting_overrides_system_proxy():
    client = DeepSeekClient(Settings(api_key="sk-test", proxy="http://127.0.0.1:7890"))
    assert client._session.proxies["https"] == "http://127.0.0.1:7890"
    assert client._session.trust_env is False


def test_proxy_error_falls_back_to_direct_connection():
    proxy_error = requests.exceptions.ProxyError("Unable to connect to proxy")
    client = _client([_ok_response('{"formulas": []}')], errors=[proxy_error])

    result = client.chat([])

    assert result.content == '{"formulas": []}'
    assert client._session.proxies == {}  # 已放弃代理
    assert client._session.trust_env is False


def test_persistent_proxy_error_message_tells_how_to_fix():
    errors = [requests.exceptions.ProxyError("Unable to connect to proxy") for _ in range(3)]
    client = _client([], errors=errors)
    with pytest.raises(LLMError, match="EXCELCR_TRUST_ENV"):
        client.chat([])


# ------------------------------------------------------------------ 返回值解析
def test_parse_json_payload_handles_markdown_wrapper():
    data = parse_json_payload('```json\n{"formula": "=SUM(A1:A2)"}\n```')
    assert data["formula"] == "=SUM(A1:A2)"


def test_parse_json_payload_rejects_plain_text():
    with pytest.raises(LLMError):
        parse_json_payload("抱歉，我无法完成")


def test_extract_candidates_normalizes_single_formula():
    items, assumptions, clarification = extract_candidates(
        {"formula": "=SUM(B2:F2)", "target": "g2", "explanation": "求和"}
    )
    assert items == [
        {"target": "g2", "formula": "=SUM(B2:F2)", "explanation": "求和", "fill_to": None}
    ]
    assert assumptions == []
    assert clarification is None


def test_repair_message_only_carries_formula_and_errors():
    digest_text = "工作表: Sheet1\n\tA\tB\n1\tSubject\tStudent1\n2\tMath\t85"
    messages = build_generate_messages(digest_text, "算总分", target="G2")
    repair = build_repair_message("=SUM(B2:G2)", ["循环引用：..."])
    assert len(messages) == 2
    assert "Student1" in messages[1]["content"]
    assert "Student1" not in repair["content"]  # 修复时不重发表格内容
    assert "循环引用" in repair["content"]
