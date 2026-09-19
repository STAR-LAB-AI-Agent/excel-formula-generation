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
    extract_table_spec,
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

    def close(self):
        pass


class FakeSession:
    """按顺序返回预置响应，并记录请求头/请求体/是否流式。"""

    def __init__(self, responses: list[FakeResponse], errors: list[Exception] | None = None):
        self.responses = list(responses)
        self.errors = list(errors or [])
        self.requests: list[dict] = []
        self.proxies: dict = {"https": "http://127.0.0.1:10808"}
        self.trust_env = True

    def post(self, url, headers=None, data=None, timeout=None, stream=False):
        if self.errors:
            raise self.errors.pop(0)
        self.requests.append({"url": url, "headers": headers or {}, "data": data, "stream": stream})
        return self.responses.pop(0)

    def request_body(self, index: int = 0) -> dict:
        return json.loads(self.requests[index]["data"].decode("utf-8"))


def _client(responses, *, errors=None, **kwargs) -> DeepSeekClient:
    settings = Settings(api_key="sk-test", **kwargs)
    client = DeepSeekClient(settings)
    client._session = FakeSession(responses, errors)
    return client


def _ok_response(content: str, finish_reason: str = "stop") -> FakeResponse:
    return FakeResponse(
        200,
        {
            "choices": [{"message": {"content": content}, "finish_reason": finish_reason}],
            "usage": {"prompt_tokens": 300, "completion_tokens": 50},
        },
    )


class FakeStreamResponse:
    """模拟 requests 的流式响应：iter_lines 逐行产出 SSE 数据。"""

    def __init__(self, lines: list[str], status_code: int = 200, text: str = ""):
        self.status_code = status_code
        self._lines = lines
        self.text = text or ""
        self.closed = False

    def iter_lines(self, decode_unicode=False):
        yield from self._lines

    def close(self):
        self.closed = True


def _sse(payload: dict) -> str:
    return "data: " + json.dumps(payload, ensure_ascii=False)


def _stream_chunk(content=None, reasoning=None, finish_reason=None, usage=None) -> dict:
    delta: dict = {}
    if content is not None:
        delta["content"] = content
    if reasoning is not None:
        delta["reasoning_content"] = reasoning
    chunk: dict = {"choices": [{"delta": delta, "finish_reason": finish_reason}]}
    if usage is not None:
        chunk["usage"] = usage
    return chunk


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
    assert request["stream"] is False


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


def test_truncated_output_reports_max_tokens_not_json_error():
    """思考型模型把额度耗在推理上时 content 为空，错误信息必须指向 max_tokens。"""
    client = _client([_ok_response("", finish_reason="length")])
    with pytest.raises(LLMError, match="max_tokens"):
        client.chat([], max_tokens=800)
    assert client.total_usage.calls == 1  # 被截断的这次调用仍要计入用量


def test_empty_content_raises_instead_of_returning_blank():
    client = _client([_ok_response("   ")])
    with pytest.raises(LLMError, match="空内容"):
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
        {"target": "g2", "formula": "=SUM(B2:F2)", "value": None, "explanation": "求和", "fill_to": None}
    ]
    assert assumptions == []
    assert clarification is None


def test_extract_candidates_keeps_value_writes():
    """分类名等常量用 value 提交，与公式候选并存；两者都没有的条目丢弃。"""
    items, _, _ = extract_candidates(
        {"formulas": [
            {"target": "A24", "value": "华东", "explanation": "区域名"},
            {"target": "B24", "formula": "=SUMIF($E$2:$E$21,A24,$H$2:$H$21)", "explanation": "汇总"},
            {"target": "C24", "explanation": "只有说明，没有可写入的内容"},
        ]}
    )
    assert [i["target"] for i in items] == ["A24", "B24"]
    assert items[0]["value"] == "华东"
    assert items[0]["formula"] == ""


def test_extract_candidates_single_value_payload():
    items, _, _ = extract_candidates({"target": "A1", "value": "销售汇总", "explanation": "标题"})
    assert len(items) == 1
    assert items[0]["value"] == "销售汇总"
    assert items[0]["formula"] == ""


def test_extract_table_spec_normalizes_coordinates():
    """新建表格几何：去空格与 $、转大写；boolean 写法只带标记不带坐标。"""
    assert extract_table_spec({"table": {"header": "a2:d2", "range": "$A$2:$D$10"}}) == {
        "header": "A2:D2", "range": "A2:D10"
    }
    assert extract_table_spec({"table": True}) == {"header": None, "range": None}
    assert extract_table_spec({"new_table": {"range": "a1 : c5"}}) == {
        "header": None, "range": "A1:C5"
    }


def test_extract_table_spec_absent_or_empty():
    assert extract_table_spec({"formulas": []}) is None
    assert extract_table_spec({"table": None}) is None
    assert extract_table_spec({"table": {"header": "", "range": ""}}) is None
    assert extract_table_spec({"new_table": False}) is None


def test_generation_prompt_hints_new_table_only_when_flagged():
    messages = build_generate_messages("工作表: Sheet1", "新建一个汇总表", new_table=True)
    assert "新建一张表格" in messages[1]["content"]
    plain = build_generate_messages("工作表: Sheet1", "统计各科平均分")
    assert "新建一张表格" not in plain[1]["content"]


def test_repair_message_only_carries_formula_and_errors():
    digest_text = "工作表: Sheet1\n\tA\tB\n1\tSubject\tStudent1\n2\tMath\t85"
    messages = build_generate_messages(digest_text, "算总分", target="G2")
    repair = build_repair_message("=SUM(B2:G2)", ["循环引用：..."])
    assert len(messages) == 2
    assert "Student1" in messages[1]["content"]
    assert "Student1" not in repair["content"]  # 修复时不重发表格内容
    assert "循环引用" in repair["content"]


def test_generation_prompt_distinguishes_target_and_sources():
    messages = build_generate_messages(
        "工作表: 班级信息", "计算三科平均分", target="D2",
        sheet_names=["成绩单", "班级信息", "历史"],
        source_digests={"成绩单": "工作表: 成绩单\n1\t学号\t姓名\t班级\t语文\t数学\t英语"},
        omitted_sheets=["历史"],
    )
    context = messages[1]["content"]
    assert "目标工作表（所有 target 均写入此表）" in context
    assert "来源工作表（仅供引用）: 成绩单" in context
    assert "因上下文预算未提供内容的工作表: 历史" in context
    assert "目标单元格: D2" in context
    assert "不得臆造" in messages[0]["content"]
    assert "IFERROR(...,0)" in messages[0]["content"]


# ------------------------------------------------------------------ 流式（SSE）
def test_streaming_via_on_delta_assembles_content_and_usage():
    """传了回调就自动走 SSE：思考与正文分轨回调，usage 从最后分片带回。"""
    lines = [
        _sse(_stream_chunk(reasoning="先看看表结构")),
        _sse(_stream_chunk(reasoning="，再决定区域")),
        _sse(_stream_chunk(content='{"formulas":')),
        _sse(_stream_chunk(content=" []}")),
        _sse(_stream_chunk(finish_reason="stop")),
        _sse({"choices": [], "usage": {"prompt_tokens": 500, "completion_tokens": 90}}),
        "data: [DONE]",
    ]
    client = _client([FakeStreamResponse(lines)])
    seen: list[tuple[str, str]] = []
    result = client.chat(
        [{"role": "user", "content": "hi"}],
        on_delta=lambda piece, kind: seen.append((kind, piece)),
    )

    assert result.content == '{"formulas": []}'
    assert result.usage.prompt_tokens == 500
    assert result.usage.completion_tokens == 90
    assert ("reasoning", "先看看表结构") in seen
    assert ("content", " []}") in seen
    # 请求体：走 SSE 并显式索取 usage；HTTP 层也用流式读取
    body = client._session.request_body(0)
    assert body["stream"] is True
    assert body["stream_options"] == {"include_usage": True}
    assert client._session.requests[0]["stream"] is True


def test_streaming_truncation_counts_usage_toward_total():
    lines = [
        _sse(_stream_chunk(reasoning="想得太久了")),
        _sse(_stream_chunk(finish_reason="length")),
        _sse({"choices": [], "usage": {"prompt_tokens": 300, "completion_tokens": 8000}}),
    ]
    client = _client([FakeStreamResponse(lines)])
    with pytest.raises(LLMError, match="max_tokens"):
        client.chat([], max_tokens=8000, on_delta=lambda piece, kind: None)
    assert client.total_usage.completion_tokens == 8000


def test_streaming_without_usage_does_not_crash():
    """服务端不回报 usage 时照常拿内容，用量按 0 计。"""
    lines = [
        _sse(_stream_chunk(content='{"formulas": []}')),
        _sse(_stream_chunk(finish_reason="stop")),
        "data: [DONE]",
    ]
    client = _client([FakeStreamResponse(lines)])
    result = client.chat([], on_delta=lambda piece, kind: None)
    assert result.content == '{"formulas": []}'
    assert result.usage.total_tokens == 0


def test_streaming_falls_back_when_stream_options_rejected():
    """服务端不认识 stream_options 时自动去掉重试，后续请求不再携带。"""
    rejected = FakeStreamResponse(
        [], status_code=400, text='{"error": {"message": "unknown field stream_options"}}'
    )
    ok_lines = [
        _sse(_stream_chunk(content='{"formulas": []}')),
        _sse(_stream_chunk(finish_reason="stop")),
        "data: [DONE]",
    ]
    client = _client([rejected, FakeStreamResponse(ok_lines)])
    result = client.chat([], on_delta=lambda piece, kind: None)

    assert result.content == '{"formulas": []}'
    assert client._stream_usage_supported is False
    assert "stream_options" in client._session.request_body(0)
    assert "stream_options" not in client._session.request_body(1)


def test_explicit_stream_true_without_callback_uses_sse():
    lines = [
        _sse(_stream_chunk(content="你好")),
        _sse(_stream_chunk(finish_reason="stop")),
        "data: [DONE]",
    ]
    client = _client([FakeStreamResponse(lines)])
    result = client.chat([{"role": "user", "content": "hi"}], stream=True)
    assert result.content == "你好"
    assert client._session.requests[0]["stream"] is True


def test_thinking_settings_are_sent_to_api():
    client = _client(
        [_ok_response('{"formulas": []}')], thinking="disabled", reasoning_effort="low"
    )
    client.chat([])
    body = client._session.request_body(0)
    assert body["thinking"] == {"type": "disabled"}
    assert body["reasoning_effort"] == "low"


def test_thinking_defaults_leave_payload_untouched():
    """未配置时不干预服务端默认行为，请求体里不应出现思考字段。"""
    client = _client([_ok_response('{"formulas": []}')])
    client.chat([])
    body = client._session.request_body(0)
    assert "thinking" not in body
    assert "reasoning_effort" not in body


def test_thinking_override_beats_settings():
    """单次调用的 thinking 参数优先级高于配置值（auto 档逐次决策的通道）。"""
    client = _client([_ok_response('{"formulas": []}')], thinking="disabled")
    client.chat([], thinking="enabled")
    assert client._session.request_body(0)["thinking"] == {"type": "enabled"}


def test_thinking_auto_without_override_sends_nothing():
    """auto 是流水线内部档位，绝不能原样发给 API；缺省时按服务端默认。"""
    client = _client([_ok_response('{"formulas": []}')], thinking="auto")
    client.chat([])
    assert "thinking" not in client._session.request_body(0)
