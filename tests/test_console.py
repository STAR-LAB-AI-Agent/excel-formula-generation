"""交互入口测试：裸文件名不触发模型调用；:think 开关即时切换；:web 演示台幂等。"""
from __future__ import annotations

from unittest.mock import Mock

import pytest
from openpyxl import Workbook

from main import Console, _is_bare_file_reference, _parse_cli_args, apply_thinking_command


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


# ------------------------------------------------------------ 启动参数与演示台

def test_interactive_mode_opens_web_by_default():
    assert _parse_cli_args([]) == (None, True)


def test_no_web_flag_keeps_cli_only():
    assert _parse_cli_args(["--no-web"]) == (None, False)


def test_once_mode_stays_quiet_without_explicit_flag():
    assert _parse_cli_args(["给G2算总分"]) == ("给G2算总分", False)
    assert _parse_cli_args(["--web", "给G2算总分"]) == ("给G2算总分", True)


def test_open_web_starts_once_and_reports_address(console, monkeypatch, capsys):
    started = []

    def fake_start(service, **kwargs):
        started.append(service)
        return object(), "http://127.0.0.1:8765/"

    monkeypatch.setattr("webapp.start_in_background", fake_start)
    assert console.open_web() == "http://127.0.0.1:8765/"
    assert console.open_web() == "http://127.0.0.1:8765/"
    assert len(started) == 1  # 第二次只提示地址，不重复起服务
    assert "已在运行" in capsys.readouterr().out


def test_open_web_reports_port_exhaustion(console, monkeypatch, capsys):
    monkeypatch.setattr("webapp.start_in_background", lambda service, **kw: (None, None))
    assert console.open_web() is None
    assert "被占用" in capsys.readouterr().out


# ------------------------------------------------------------ 编号选文件
@pytest.fixture()
def console(settings, monkeypatch):
    monkeypatch.setattr("main.Settings.from_env", lambda **kwargs: settings)
    monkeypatch.setattr("main.get_logger", lambda *args: Mock())
    monkeypatch.setattr("main.FormulaService", lambda *args: Mock())
    return Console()


def _excel(path):
    book = Workbook()
    book.active["A1"] = "测试数据"
    book.save(path)
    book.close()
    return path


def _inputs(monkeypatch, lines):
    values = iter(lines)
    monkeypatch.setattr("builtins.input", lambda *args: next(values))


def test_file_list_filters_and_sorts(console, tmp_path, capsys):
    _excel(tmp_path / "b.XLSX")
    _excel(tmp_path / "a.xlsx")
    _excel(tmp_path / "宏.xlsm")
    (tmp_path / "~$a.xlsx").touch()
    (tmp_path / "old.xls").touch()
    (tmp_path / "notes.txt").touch()
    (tmp_path / "folder.xlsx").mkdir()
    (tmp_path / "backups").mkdir()
    _excel(tmp_path / "backups" / "backup.xlsx")
    console.show_files()
    assert [p.name for p in console._file_choices] == ["a.xlsx", "b.XLSX", "宏.xlsm"]
    output = capsys.readouterr().out
    assert "1. a.xlsx" in output and "2. b.XLSX" in output
    assert "backup.xlsx" not in output and "~$" not in output
    assert console._guess_file() is None


def test_single_file_is_auto_selected(console, tmp_path):
    path = _excel(tmp_path / "唯一.xlsx")
    assert console._guess_file() == path


def test_startup_number_selection_does_not_call_model(console, tmp_path, monkeypatch, capsys):
    _excel(tmp_path / "a.xlsx")
    selected = _excel(tmp_path / "b.xlsx")
    before = selected.read_bytes()
    console.current_sheet = "上一个文件的表"
    _inputs(monkeypatch, ["2", ":quit"])
    assert console.run() == 0
    assert console.current_file == selected
    assert console.current_sheet is None
    console.service.propose.assert_not_called()
    console.service.apply.assert_not_called()
    assert selected.read_bytes() == before
    assert "2. b.xlsx" in capsys.readouterr().out


@pytest.mark.parametrize("number", ["0", "-1", "999", "9" * 100])
def test_bad_number_preserves_current_selection(console, tmp_path, number, capsys):
    path = _excel(tmp_path / "a.xlsx")
    console._set_file(str(path))
    console.current_sheet = "Sheet1"
    console.show_files()
    assert not console.handle(number)
    assert console.current_file == path and console.current_sheet == "Sheet1"
    assert "编号无效" in capsys.readouterr().out
    console.service.propose.assert_not_called()


def test_empty_directory_and_eof_are_safe(console, monkeypatch, capsys):
    assert console._guess_file() is None
    console.show_files()
    assert not console.handle("1")
    monkeypatch.setattr("builtins.input", Mock(side_effect=EOFError))
    assert console.run() == 0
    assert "未找到 .xlsx/.xlsm" in capsys.readouterr().out


def test_menu_numbers_remain_stable_until_refresh(console, tmp_path):
    first = _excel(tmp_path / "b.xlsx")
    _excel(tmp_path / "c.xlsx")
    console.show_files()
    new = _excel(tmp_path / "a.xlsx")
    assert console._set_file("1") == first
    console.show_files()
    assert console._set_file("1") == new


def test_vanished_numbered_file_is_not_replaced(console, tmp_path, capsys):
    first = _excel(tmp_path / "a.xlsx")
    second = _excel(tmp_path / "b.xlsx")
    console.show_files()
    first.rename(tmp_path / "moved.xlsx")
    console._set_file(str(second))
    assert console._set_file("1") is None
    assert console.current_file == second
    assert "文件不可用" in capsys.readouterr().out


def test_file_command_number_and_refresh(console, tmp_path, monkeypatch):
    _excel(tmp_path / "a.xlsx")
    selected = _excel(tmp_path / "b.xlsx")
    _inputs(monkeypatch, [":files", ":file 2", ":think 0", ":quit"])
    assert console.run() == 0
    assert console.current_file == selected
    assert console.settings.thinking == "disabled"
    console.service.propose.assert_not_called()


def test_path_selection_and_blank_file_command(console, tmp_path):
    selected = _excel(tmp_path / "含 空格.xlsx")
    assert console._set_file(f'"{selected}"') == selected
    assert console._set_file("") is None
    assert console.current_file == selected
    assert console._file_choices == [selected]


def test_need_file_retries_invalid_number_and_accepts_choice(console, tmp_path, monkeypatch):
    selected = _excel(tmp_path / "a.xlsx")
    _inputs(monkeypatch, ["9", "1"])
    assert console._need_file() == selected


def test_number_during_clarification_is_not_file_selection(console, tmp_path, monkeypatch):
    selected = _excel(tmp_path / "a.xlsx")
    console.show_files()
    console._set_file("1")
    console._propose_with_progress = Mock(side_effect=[
        Mock(clarification="需要统计多少项？"),
        Mock(clarification=None, ok=False, render=lambda: "未生成公式",
             usage=Mock(to_dict=lambda: {})),
    ])
    _inputs(monkeypatch, ["2"])
    console.do_generate(selected, "计算数据", None, None)
    assert console.current_file == selected
    assert "补充：2" in console._propose_with_progress.call_args.args[1]


def test_failed_prediction_never_prompts_for_confirmation(console, tmp_path, monkeypatch):
    selected = _excel(tmp_path / "a.xlsx")
    console._propose_with_progress = Mock(return_value=Mock(
        clarification=None, ok=False, render=lambda: "本地试算得到 #DIV/0!",
        usage=Mock(to_dict=lambda: {}),
    ))
    ask = Mock(side_effect=AssertionError("错误方案不应询问写入"))
    monkeypatch.setattr("builtins.input", ask)
    console.do_generate(selected, "计算均分", None, None)
    ask.assert_not_called()
    console.service.apply.assert_not_called()


def _confirmed_write(console, monkeypatch, path, applied):
    console._propose_with_progress = Mock(return_value=Mock(
        clarification=None, ok=True, usage=Mock(to_dict=lambda: {}),
        sheet="Sheet1", target="D1", writes=[Mock(cell="D1")],
        render=lambda: "（预览）",
    ))
    console.service.apply = Mock(return_value=applied)
    _inputs(monkeypatch, ["y"])
    console.do_generate(path, "增加一列", None, None)


def test_write_success_reports_style_inheritance(console, tmp_path, monkeypatch, capsys):
    """写入成功后在控制台说明哪些新单元格套用了相邻格式。"""
    selected = _excel(tmp_path / "班级成绩.xlsx")
    _confirmed_write(console, monkeypatch, selected, {
        "file": str(selected), "count": 4, "backup": None, "fidelity": {},
        "styled": ["D1 ← C1"],
    })
    assert "已为新单元格套用相邻格式：D1 ← C1" in capsys.readouterr().out


def test_write_without_style_inheritance_has_no_extra_line(console, tmp_path, monkeypatch, capsys):
    """没有可沿用的格式时不显示提示行。"""
    selected = _excel(tmp_path / "普通.xlsx")
    _confirmed_write(console, monkeypatch, selected, {
        "file": str(selected), "count": 1, "backup": None, "fidelity": {},
        "styled": [],
    })
    assert "套用相邻格式" not in capsys.readouterr().out


def test_write_success_reports_table_format(console, tmp_path, monkeypatch, capsys):
    """新建表格写入成功后，控制台列出已套用的表格格式。"""
    selected = _excel(tmp_path / "汇总.xlsx")
    label = "Sheet1!A1:B5（表头 A1:B1 浅蓝底、范围加细边框）"
    _confirmed_write(console, monkeypatch, selected, {
        "file": str(selected), "count": 6, "backup": None, "fidelity": {},
        "styled": [], "tables": [label],
    })
    assert f"已套用表格格式：{label}" in capsys.readouterr().out
