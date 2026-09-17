"""配置层测试：思考开关的数字写法（1/0）归一化与非法值拦截。"""
from __future__ import annotations

import pytest

from excel_formula.config import ConfigError, Settings, normalize_thinking


def test_numeric_switch_maps_to_enabled_and_disabled():
    assert normalize_thinking("1") == "enabled"
    assert normalize_thinking("0") == "disabled"


def test_alias_and_keyword_values_pass_through():
    assert normalize_thinking("ON") == "enabled"
    assert normalize_thinking("off") == "disabled"
    assert normalize_thinking("auto") == "auto"
    assert normalize_thinking("") == ""


def test_unknown_value_raises_config_error():
    with pytest.raises(ConfigError):
        normalize_thinking("maybe")


def test_from_env_accepts_numeric_switch(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)  # 避免读到仓库根目录的 .env
    monkeypatch.setenv("EXCELCR_THINKING", "1")
    assert Settings.from_env(workspace=tmp_path).thinking == "enabled"
    monkeypatch.setenv("EXCELCR_THINKING", "0")
    assert Settings.from_env(workspace=tmp_path).thinking == "disabled"


def test_from_env_rejects_invalid_thinking(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    monkeypatch.setenv("EXCELCR_THINKING", "sometimes")
    with pytest.raises(ConfigError):
        Settings.from_env(workspace=tmp_path)
