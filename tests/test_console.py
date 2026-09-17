"""交互入口测试：裸文件名不触发模型调用；:think 开关即时切换。"""
from __future__ import annotations

from main import _is_bare_file_reference, apply_thinking_command


def test_bare_filename_is_not_a_request():
    assert _is_bare_file_reference("VLOOKUP.xlsx", "VLOOKUP.xlsx")
    assert _is_bare_file_reference(" “测试数据.xlsx” ", "测试数据.xlsx")


def test_filename_with_real_request_still_goes_to_model():
    assert not _is_bare_file_reference("VLOOKUP.xlsx 里算每个人的总分", "VLOOKUP.xlsx")


def test_no_filename_means_normal_request():
    assert not _is_bare_file_reference("在G2算每科总分", None)


# ------------------------------------------------------------ :think 思考开关

def test_think_command_enables_deep_thinking(settings):
    message = apply_thinking_command(settings, "1")
    assert settings.thinking == "enabled"
    assert "深度思考" in message


def test_think_command_disables_thinking(settings):
    message = apply_thinking_command(settings, "0")
    assert settings.thinking == "disabled"
    assert "最快" in message


def test_think_command_switches_to_auto(settings):
    apply_thinking_command(settings, "auto")
    assert settings.thinking == "auto"


def test_think_command_without_argument_reports_current(settings):
    settings.thinking = "disabled"
    message = apply_thinking_command(settings, "")
    assert "当前思考模式" in message
    assert "关闭" in message


def test_think_command_rejects_unknown_value(settings):
    settings.thinking = "auto"
    message = apply_thinking_command(settings, "banana")
    assert settings.thinking == "auto"  # 非法输入不改变当前档位
    assert message.startswith("×")
