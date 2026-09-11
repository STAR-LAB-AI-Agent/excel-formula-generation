"""端到端流水线测试：用假客户端替代 DeepSeek，覆盖成功、重试修复、追问、安全边界。"""
from __future__ import annotations

import json

import pytest
from openpyxl import load_workbook

from excel_formula.config import SecurityError
from excel_formula.llm_client import ChatResult, DeepSeekClient, Usage
from excel_formula.pipeline import FormulaService


class FakeClient:
    """按顺序返回预置回复，并记录收到的消息，便于断言"重试只回传错误"。"""

    def __init__(self, replies: list[dict | str]):
        self.replies = list(replies)
        self.calls: list[list[dict]] = []

    def chat(self, messages, *, json_mode=True, max_tokens=800):
        self.calls.append(messages)
        reply = self.replies.pop(0)
        content = reply if isinstance(reply, str) else json.dumps(reply, ensure_ascii=False)
        return ChatResult(content=content, usage=Usage(prompt_tokens=120, completion_tokens=40, calls=1))


def _service(settings, replies):
    return FormulaService(settings, logger=None, client=FakeClient(replies))


# ------------------------------------------------------------------ 多目标
def test_keeps_every_target_when_model_returns_several(sample_xlsx, settings):
    """“在 H2 H3 H4 分别求…”这类需求：模型给几个就要全部保留，不能只取第一个。"""
    service = _service(
        settings,
        [{
            "formulas": [
                {"target": "H2", "formula": "=MIN(B2:F2)", "explanation": "Math 最小值"},
                {"target": "H3", "formula": "=MIN(B3:F3)", "explanation": "English 最小值"},
                {"target": "H4", "formula": "=MIN(B4:F4)", "explanation": "Physics 最小值"},
            ],
            "assumptions": [],
            "clarification": None,
        }],
    )
    proposal = service.propose(sample_xlsx, "在H2 H3 H4分别求前三科的最低分")

    assert proposal.ok
    assert [w.cell for w in proposal.writes] == ["H2", "H3", "H4"]
    assert [w.formula for w in proposal.writes] == [
        "=MIN(B2:F2)", "=MIN(B3:F3)", "=MIN(B4:F4)"
    ]
    # 每个单元格都应有自己的本地试算结果
    assert [w.predicted for w in proposal.writes] == ["78", "82", "75"]
    rendered = proposal.render()
    assert "H3 = =MIN(B3:F3)" in rendered
    assert "H4 = =MIN(B4:F4)" in rendered

    applied = service.apply(proposal)
    assert applied["count"] == 3


def test_rejects_batch_when_one_formula_is_invalid(sample_xlsx, settings):
    """一批里有一个不合法就整批不写，并把出错单元格标出来。"""
    bad_batch = {
        "formulas": [
            {"target": "H2", "formula": "=MIN(B2:F2)", "explanation": "正常"},
            {"target": "H3", "formula": '=INDIRECT("B3")', "explanation": "高风险"},
        ],
        "assumptions": [],
        "clarification": None,
    }
    service = _service(settings, [bad_batch, bad_batch, bad_batch])
    proposal = service.propose(sample_xlsx, "在H2 H3分别求最低分")

    assert not proposal.ok
    assert not proposal.writes
    assert "H3" in proposal.error  # 错误带上了出错的单元格
    assert "INDIRECT" in proposal.error


# ------------------------------------------------------------------ 正常链路
def test_generate_and_apply(sample_xlsx, settings):
    service = _service(
        settings,
        [{"formulas": [{"target": "G2", "formula": "=SUM(B2:F2)", "explanation": "B2到F2求和",
                        "fill_to": "G6"}], "assumptions": [], "clarification": None}],
    )
    proposal = service.propose(sample_xlsx, "帮我算每个科目的总分")

    assert proposal.ok
    assert proposal.formula == "=SUM(B2:F2)"
    assert [w.cell for w in proposal.writes] == ["G2", "G3", "G4", "G5", "G6"]
    assert proposal.writes[-1].formula == "=SUM(B6:F6)"  # 相对引用已平移
    assert proposal.predicted == "438"  # 写入前的本地预期结果
    assert proposal.usage.total_tokens == 160

    applied = service.apply(proposal)
    assert applied["count"] == 5
    assert applied["backup"] is not None

    workbook = load_workbook(sample_xlsx)
    assert workbook["Sheet1"]["G2"].value == "=SUM(B2:F2)"
    assert workbook["Sheet1"]["G6"].value == "=SUM(B6:F6)"
    workbook.close()


def test_preview_does_not_touch_file(sample_xlsx, settings):
    before = sample_xlsx.read_bytes()
    service = _service(
        settings,
        [{"formulas": [{"target": "G2", "formula": "=AVERAGE(B2:F2)", "explanation": "平均分"}]}],
    )
    proposal = service.propose(sample_xlsx, "算平均分")
    assert proposal.ok
    assert sample_xlsx.read_bytes() == before  # 预览阶段绝不写盘


# ------------------------------------------------------------------ 校验重试
def test_repair_loop_fixes_invalid_formula(sample_xlsx, settings):
    service = _service(
        settings,
        [
            {"formulas": [{"target": "G2", "formula": "=SUM(B2:G2)", "explanation": "错误：含自身"}]},
            {"formulas": [{"target": "G2", "formula": "=SUM(B2:F2)", "explanation": "修正版"}]},
        ],
    )
    proposal = service.propose(sample_xlsx, "算总分写到G2")

    assert proposal.ok
    assert len(proposal.attempts) == 2
    assert not proposal.attempts[0].ok
    assert "循环引用" in proposal.attempts[0].errors[0]
    # 修复请求只带公式与错误，不重复发送表格内容
    repair_message = service.client.calls[1][-1]["content"]
    assert "循环引用" in repair_message
    assert "Student1" not in repair_message


def test_gives_up_after_max_rounds(sample_xlsx, settings):
    bad = {"formulas": [{"target": "G2", "formula": '=INDIRECT("B2")', "explanation": "高风险"}]}
    service = _service(settings, [bad, bad, bad])
    proposal = service.propose(sample_xlsx, "随便算点东西")

    assert not proposal.ok
    assert len(proposal.attempts) == settings.max_repair_rounds + 1
    assert "高风险函数" in proposal.error
    with pytest.raises(ValueError):
        service.apply(proposal)  # 未通过校验的方案禁止写入


def test_clarification_is_returned(sample_xlsx, settings):
    service = _service(
        settings,
        [{"formulas": [], "clarification": "你想统计哪一列的分数？", "assumptions": []}],
    )
    proposal = service.propose(sample_xlsx, "帮我统计一下")
    assert not proposal.ok
    assert proposal.clarification == "你想统计哪一列的分数？"


def test_missing_target_triggers_question(sample_xlsx, settings):
    service = _service(
        settings, [{"formulas": [{"target": "", "formula": "=SUM(B2:F2)", "explanation": "总分"}]}]
    )
    proposal = service.propose(sample_xlsx, "算总分")
    assert proposal.clarification is not None


# ------------------------------------------------------------------ 其他意图
def test_validate_reports_predicted_value(sample_xlsx, settings):
    service = _service(settings, [])
    result = service.validate(sample_xlsx, '=COUNTIF(B2:F2,">85")', target="H2")
    assert result["validation"]["ok"]
    assert result["predicted_value"] == "3"


def test_write_formula_rejects_invalid(sample_xlsx, settings):
    service = _service(settings, [])
    with pytest.raises(ValueError, match="未通过校验"):
        service.write_formula(sample_xlsx, "G2", "=SUM(B2:G2")


def test_overwrite_is_reported_in_preview(sample_xlsx, settings):
    service = _service(
        settings, [{"formulas": [{"target": "A2", "formula": "=UPPER(A3)", "explanation": "覆盖测试"}]}]
    )
    proposal = service.propose(sample_xlsx, "把A2改成大写的A3")
    assert proposal.ok
    assert proposal.writes[0].old_value == "Math"
    assert "将覆盖已有内容" in proposal.render()


# ------------------------------------------------------------------ 安全边界
def test_path_outside_workspace_rejected(settings, tmp_path):
    service = _service(settings, [])
    outside = tmp_path.parent / "outside.xlsx"
    with pytest.raises(SecurityError):
        service.describe(outside)


def test_non_excel_suffix_rejected(settings, tmp_path):
    service = _service(settings, [])
    text_file = tmp_path / "data.csv"
    text_file.write_text("a,b", encoding="utf-8")
    with pytest.raises(SecurityError):
        service.describe(text_file)


def test_fill_range_cell_limit(sample_xlsx, settings):
    service = _service(
        settings,
        [{"formulas": [{"target": "G2", "formula": "=SUM(B2:F2)", "explanation": "过大范围",
                        "fill_to": "G5000"}]}] * 3,
    )
    proposal = service.propose(sample_xlsx, "整列都填上总分")
    assert not proposal.ok
    assert "最多写入" in proposal.error


# ------------------------------------------------------------------ 真实客户端接缝
def test_pipeline_with_real_client_and_fake_socket(sample_xlsx, settings):
    """用真的 DeepSeekClient（仅替换 HTTP 层）跑一遍全链路，验证两者的接口契约。"""

    class FakeResponse:
        status_code = 200
        text = ""

        @staticmethod
        def json():
            return {
                "choices": [{"message": {"content": json.dumps(
                    {"formulas": [{"target": "H2", "formula": "=AVERAGE(B2:F2)",
                                   "explanation": "五人平均分"}]}, ensure_ascii=False)}}],
                "usage": {"prompt_tokens": 380, "completion_tokens": 52},
            }

    class FakeSession:
        def post(self, url, headers=None, data=None, timeout=None):
            assert headers["Authorization"].startswith("Bearer ")
            return FakeResponse()

    client = DeepSeekClient(settings)
    client._session = FakeSession()
    service = FormulaService(settings, logger=None, client=client)

    proposal = service.propose(sample_xlsx, "算每科平均分", target="H2")
    assert proposal.ok
    assert proposal.formula == "=AVERAGE(B2:F2)"
    assert proposal.predicted == "87.6"
    assert proposal.usage.prompt_tokens == 380
