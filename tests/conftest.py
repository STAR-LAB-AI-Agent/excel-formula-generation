"""pytest 公共夹具：构造一份与 测试数据.xlsx 结构一致的临时工作簿。"""
from __future__ import annotations

import sys
from pathlib import Path

import pytest
from openpyxl import Workbook

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from excel_formula.config import Settings  # noqa: E402

ROWS = [
    ["Subject", "Student1", "Student2", "Student3", "Student4", "Student5"],
    ["Math", 85, 78, 92, 88, 95],
    ["English", 88, 85, 91, 82, 90],
    ["Physics", 80, 82, 89, 75, 88],
    ["Chemistry", 90, 87, 84, 91, 86],
    ["Biology", 78, 80, 86, 83, 81],
]


@pytest.fixture()
def sample_xlsx(tmp_path: Path) -> Path:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Sheet1"
    for row in ROWS:
        sheet.append(row)
    workbook.create_sheet("Sheet2")
    path = tmp_path / "scores.xlsx"
    workbook.save(path)
    workbook.close()
    return path


@pytest.fixture()
def settings(tmp_path: Path) -> Settings:
    return Settings(
        api_key="sk-test-key-not-real",
        allowed_roots=[tmp_path.resolve()],
        log_dir=tmp_path / "logs",
        max_repair_rounds=2,
    )
