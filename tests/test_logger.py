"""日志测试：脱敏必须生效，但不能破坏 %d / %.2f 这类数字占位符。

回归背景：RedactFilter 曾把所有参数一律 str() 化，导致 llm_client 里
"消息数=%d 输入字符数=%d"、"tokens(in/out)=%d/%d 耗时=%.2fs" 两条日志
在格式化时抛 TypeError —— logging 把错误打到 stderr 并丢弃记录，
Token 用量因此从未写进日志文件。
"""
from __future__ import annotations

import json
import logging

import pytest

from excel_formula.config import Settings
from excel_formula.logger import RedactFilter, redact
from excel_formula.llm_client import DeepSeekClient
from tests.test_llm_client import FakeResponse, FakeSession


@pytest.fixture
def logfile(tmp_path, request):
    """独立 logger + 文件 handler，避免与全局 excelcr logger 相互干扰。"""
    path = tmp_path / "test.log"
    logger = logging.getLogger(f"excelcr-test-{request.node.name}")
    logger.handlers.clear()
    logger.setLevel(logging.DEBUG)
    logger.propagate = False
    handler = logging.FileHandler(path, encoding="utf-8")
    handler.setFormatter(logging.Formatter("%(levelname)s|%(message)s"))
    handler.addFilter(RedactFilter())
    logger.addHandler(handler)

    yield logger, path

    handler.close()
    logger.handlers.clear()


def _read(path) -> str:
    return path.read_text(encoding="utf-8")


# ------------------------------------------------------- 数字占位符（本次回归点）
def test_numeric_placeholders_are_not_broken(logfile, capsys):
    logger, path = logfile
    logger.info("消息数=%d 输入字符数=%d", 2, 1051)
    logger.info("tokens(in/out)=%d/%d 耗时=%.2fs", 547, 69, 0.9584626)

    content = _read(path)
    assert "消息数=2 输入字符数=1051" in content
    assert "tokens(in/out)=547/69 耗时=0.96s" in content
    # logging 的格式化错误只打到 stderr，不会让测试失败，必须显式断言它没出现
    assert "Logging error" not in capsys.readouterr().err


def test_bool_and_float_args_survive(logfile):
    logger, path = logfile
    logger.info("json_mode=%s 耗时=%.1fs", True, 1.25)
    assert "json_mode=True 耗时=1.2s" in _read(path)


def test_dict_style_args_keep_numbers(logfile):
    logger, path = logfile
    logger.info("%(name)s 用了 %(tokens)d tokens", {"name": "deepseek-chat", "tokens": 616})
    assert "deepseek-chat 用了 616 tokens" in _read(path)


# ------------------------------------------------------------------- 脱敏能力
def test_api_key_in_string_arg_is_redacted(logfile):
    logger, path = logfile
    logger.info("请求头 %s，第 %d 次", "Authorization: Bearer sk-abcdef1234567890", 1)

    content = _read(path)
    assert "sk-abcdef1234567890" not in content
    assert "***" in content
    assert "第 1 次" in content  # 数字仍然正常


def test_api_key_in_message_itself_is_redacted(logfile):
    logger, path = logfile
    logger.info("DEEPSEEK_API_KEY=sk-0123456789abcdef 已加载")
    assert "sk-0123456789abcdef" not in _read(path)


def test_exception_arg_is_still_redacted(logfile):
    logger, path = logfile
    logger.warning("请求异常（第 %d 次）：%s", 2, RuntimeError("api_key=sk-secretvalue123"))

    content = _read(path)
    assert "sk-secretvalue123" not in content
    assert "第 2 次" in content


@pytest.mark.parametrize(
    "text",
    [
        "sk-abcdefgh12345678",
        "authorization: Bearer sk-xyz98765432",
        "password = hunter2000",
        "token=abcdefghijklmn",
    ],
)
def test_redact_covers_common_secret_shapes(text):
    cleaned = redact(text)
    assert "***" in cleaned
    for secret in ("sk-abcdefgh12345678", "hunter2000", "abcdefghijklmn"):
        assert secret not in cleaned


# --------------------------------------------- 真实调用链：Token 用量必须落盘
def test_client_writes_token_usage_to_log(logfile, capsys):
    """还原 bug 现场：走一遍 DeepSeekClient.chat，两条日志都要出现在文件里。"""
    logger, path = logfile
    client = DeepSeekClient(Settings(api_key="sk-test-key-1234567890"), logger=logger)
    client._session = FakeSession(
        [
            FakeResponse(
                200,
                {
                    "choices": [{"message": {"content": json.dumps({"formulas": []})}}],
                    "usage": {"prompt_tokens": 547, "completion_tokens": 69},
                },
            )
        ]
    )

    client.chat([{"role": "system", "content": "x" * 600}, {"role": "user", "content": "y" * 40}])

    content = _read(path)
    assert "调用模型" in content and "消息数=2 输入字符数=640" in content
    assert "模型返回 tokens(in/out)=547/69" in content
    assert "sk-test-key-1234567890" not in content
    assert "Logging error" not in capsys.readouterr().err
