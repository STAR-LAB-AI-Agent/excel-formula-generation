"""公式写入：备份 -> 写入 -> 保存，并把相对引用填充展开为多个单元格。"""
from __future__ import annotations

import shutil
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.formula.translate import Translator
from openpyxl.utils import get_column_letter

from .config import MAX_WRITE_CELLS, Settings
from .formula_parser import parse_target


@dataclass
class CellWrite:
    """一次待写入操作。old_formula/old_value/predicted 均仅用于预览与回溯。"""

    sheet: str
    cell: str
    formula: str
    explanation: str = ""
    old_formula: str | None = None
    old_value: object = None
    predicted: str | None = None
    predicted_note: str | None = None

    def to_dict(self) -> dict:
        return {
            "sheet": self.sheet,
            "cell": self.cell,
            "formula": self.formula,
            "explanation": self.explanation,
            "overwrites": self.old_formula or (
                None if self.old_value is None else str(self.old_value)
            ),
            "predicted_value": self.predicted,
            "predicted_note": self.predicted_note,
        }


def expand_fill(sheet: str, target: str, formula: str, fill_to: str | None) -> list[CellWrite]:
    """把 target + fill_to 展开成一列（或一行）单元格，相对引用自动平移。"""
    _, col, row = parse_target(target)
    base = f"{get_column_letter(col)}{row}"
    writes = [CellWrite(sheet=sheet, cell=base, formula=formula)]
    if not fill_to:
        return writes

    _, end_col, end_row = parse_target(fill_to)
    if end_col != col and end_row != row:
        raise ValueError(f"填充范围 {target}:{fill_to} 必须在同一行或同一列")

    cells: list[str] = []
    if end_row != row:
        step = 1 if end_row > row else -1
        cells = [f"{get_column_letter(col)}{r}" for r in range(row + step, end_row + step, step)]
    elif end_col != col:
        step = 1 if end_col > col else -1
        cells = [f"{get_column_letter(c)}{row}" for c in range(col + step, end_col + step, step)]

    if len(cells) + 1 > MAX_WRITE_CELLS:
        raise ValueError(
            f"一次最多写入 {MAX_WRITE_CELLS} 个单元格，当前请求 {len(cells) + 1} 个，请缩小范围"
        )
    translator = Translator(formula, origin=base)
    for cell in cells:
        writes.append(CellWrite(sheet=sheet, cell=cell, formula=translator.translate_formula(cell)))
    return writes


def backup_file(path: Path, settings: Settings) -> Path | None:
    """写入前备份到 backups/ 目录，便于人工回退。"""
    if not settings.backup:
        return None
    backup_dir = path.parent / "backups"
    backup_dir.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    destination = backup_dir / f"{path.stem}_{stamp}{path.suffix}"
    shutil.copy2(path, destination)
    return destination


def apply_writes(
    path: Path,
    writes: list[CellWrite],
    settings: Settings,
    *,
    logger=None,
    output: Path | None = None,
) -> dict:
    """执行写入。output 为空表示原地保存（会先备份）。"""
    if not writes:
        raise ValueError("没有需要写入的公式")
    if len(writes) > MAX_WRITE_CELLS:
        raise ValueError(f"一次最多写入 {MAX_WRITE_CELLS} 个单元格")

    target_path = output or path
    backup = backup_file(path, settings) if target_path == path else None

    workbook = load_workbook(path, keep_vba=path.suffix.lower() == ".xlsm")
    try:
        for item in writes:
            if item.sheet not in workbook.sheetnames:
                raise KeyError(f"工作表 {item.sheet!r} 不存在")
            workbook[item.sheet][item.cell] = item.formula
        try:
            workbook.save(target_path)
        except PermissionError as exc:
            raise PermissionError(
                f"无法写入 {target_path.name}，文件可能正被 Excel 打开，请关闭后重试"
            ) from exc
    finally:
        workbook.close()

    if logger:
        logger.info(
            "写入完成 file=%s cells=%s backup=%s",
            target_path.name,
            ",".join(f"{w.sheet}!{w.cell}" for w in writes),
            backup.name if backup else "无",
        )
    return {
        "file": str(target_path),
        "backup": str(backup) if backup else None,
        "written": [w.to_dict() for w in writes],
        "count": len(writes),
    }
