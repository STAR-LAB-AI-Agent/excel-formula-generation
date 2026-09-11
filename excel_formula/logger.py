"""运行日志：文件 + 控制台双通道，并对密钥类信息做脱敏。"""
from __future__ import annotations

import logging
import re
from pathlib import Path

LOGGER_NAME = "excelcr"

# 需要脱敏的模式：sk-xxx 密钥、Authorization 头、显式的 key=xxx
_REDACT_PATTERNS = [
    re.compile(r"sk-[A-Za-z0-9\-_]{8,}"),
    re.compile(r"(?i)(authorization\s*[:=]\s*)(bearer\s+)?\S+"),
    re.compile(r"(?i)((?:api[_-]?key|token|password|secret)\s*[:=]\s*)\S+"),
]


class RedactFilter(logging.Filter):
    """兜底防护：即便调用方误把密钥写进日志，也不会落盘。"""

    def filter(self, record: logging.LogRecord) -> bool:
        record.msg = redact(str(record.msg))
        if record.args:
            if isinstance(record.args, dict):
                record.args = {k: _redact_arg(v) for k, v in record.args.items()}
            else:
                record.args = tuple(_redact_arg(a) for a in record.args)
        return True


def _redact_arg(value):
    """参数脱敏：数字必须原样放过。

    若一律转成字符串，使用 %d / %.2f 占位符的日志会在格式化时抛
    TypeError，logging 会把错误打到 stderr 并丢弃这条记录（日志就残了）。
    数字里不可能藏密钥，跳过它们不削弱脱敏能力。
    """
    if isinstance(value, str):
        return redact(value)
    if isinstance(value, (int, float, complex)):  # bool 是 int 子类，一并放过
        return value
    return redact(str(value))  # 异常对象等仍然先脱敏再输出


def redact(text: str) -> str:
    for pattern in _REDACT_PATTERNS:
        if pattern.groups >= 1:
            text = pattern.sub(lambda m: (m.group(1) or "") + "***", text)
        else:
            text = pattern.sub("***", text)
    return text


def get_logger(log_dir: str | Path = "logs", *, verbose: bool = False) -> logging.Logger:
    """获取全局 logger；重复调用不会重复添加 handler。"""
    logger = logging.getLogger(LOGGER_NAME)
    if logger.handlers:
        return logger

    logger.setLevel(logging.DEBUG)
    logger.propagate = False
    fmt = logging.Formatter(
        "%(asctime)s | %(levelname)-7s | %(name)s | %(message)s", datefmt="%Y-%m-%d %H:%M:%S"
    )

    directory = Path(log_dir)
    try:
        directory.mkdir(parents=True, exist_ok=True)
        file_handler = logging.FileHandler(directory / "excelcr.log", encoding="utf-8")
        file_handler.setLevel(logging.DEBUG)
        file_handler.setFormatter(fmt)
        file_handler.addFilter(RedactFilter())
        logger.addHandler(file_handler)
    except OSError:  # 日志目录不可写时不影响主流程
        pass

    # 控制台默认保持安静：用户可见的提示由 CLI / main.py 统一输出，避免同一条信息重复出现
    console = logging.StreamHandler()
    console.setLevel(logging.DEBUG if verbose else logging.CRITICAL + 1)
    console.setFormatter(logging.Formatter("[%(levelname)s] %(message)s"))
    console.addFilter(RedactFilter())
    logger.addHandler(console)
    return logger
