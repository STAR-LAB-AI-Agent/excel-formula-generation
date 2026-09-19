"""跨工作簿引用支持：解析 [book]Sheet 形式的引用并按需打开外部工作簿。

公式引用其他工作簿（如 =SUM([库存.xlsx]Sheet1!B2:B10)）时，本地求值器
需要能真正读到那份工作簿的数据。这里提供：

* `parse_external_sheet`：把 `[路径/][文件名]工作表` 拆成三段；
* `ExternalBookLoader`：按公式给出的位置找到文件、只读打开（同一文件只开一次），
  失败时给出中文原因，由求值器转成"未验证"提示。

加载不写任何文件；单个外部工作簿超过 `MAX_EXTERNAL_FILE_BYTES` 时跳过试算。
"""
from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path

from openpyxl.worksheet.worksheet import Worksheet

from .excel_reader import WorkbookView

# 单个外部工作簿的读取上限：避免误引用超大文件拖垮预览
MAX_EXTERNAL_FILE_BYTES = 50 * 1024 * 1024
# 公式里省略扩展名时依次尝试的候选
_CANDIDATE_SUFFIXES = (".xlsx", ".xlsm", ".xls")

_EXTERNAL_PARTS = re.compile(r"^(?P<dir>.*?)\[(?P<book>[^\[\]]+)\](?P<sheet>[^\[\]]*)$")


@dataclass(frozen=True)
class ExternalRef:
    """跨工作簿引用的拆解结果：可选的目录部分 + 文件名 + 内部工作表名。"""

    dir_part: str
    book: str
    sheet: str


def parse_external_sheet(sheet_name: str) -> ExternalRef | None:
    """拆解 [目录/][文件名]工作表 形式的引用；不是跨簿引用时返回 None。"""
    match = _EXTERNAL_PARTS.match(sheet_name)
    if match is None:
        return None
    return ExternalRef(
        dir_part=match.group("dir"),
        book=match.group("book"),
        sheet=match.group("sheet"),
    )


@dataclass
class ExternalBook:
    """一份已打开的外部工作簿：缓存值视图 + 公式视图（均按工作表名索引）。"""

    path: Path
    values: dict[str, Worksheet]
    formulas: dict[str, Worksheet]
    view: WorkbookView

    def find_values(self, sheet: str) -> Worksheet | None:
        return _find(self.values, sheet)

    def find_formulas(self, sheet: str) -> Worksheet | None:
        return _find(self.formulas, sheet)


def _find(sheets: dict[str, Worksheet], sheet: str) -> Worksheet | None:
    """Excel 的工作表名不区分大小写：先精确匹配，再大小写不敏感兜底。"""
    if sheet in sheets:
        return sheets[sheet]
    lowered = sheet.casefold()
    for name, ws in sheets.items():
        if name.casefold() == lowered:
            return ws
    return None


class ExternalBookLoader:
    """按需打开外部工作簿的加载器；用完调用 close() 释放文件视图。

    实例可直接作为 evaluate_formula 的 load_external 参数使用。
    """

    def __init__(self) -> None:
        self._books: dict[str, ExternalBook] = {}
        self._failures: dict[str, str] = {}
        self.loaded: list[Path] = []  # 成功加载的文件，供日志与调试

    # -------------------------------------------------------------- 对外入口
    def __call__(
        self, sheet_name: str, parent_dir: Path | None = None
    ) -> tuple[ExternalBook | None, str]:
        """返回 (外部工作簿, 内部工作表名)；失败时返回 (None, 原因)。"""
        ref = parse_external_sheet(sheet_name)
        if ref is None:
            return None, "不是跨工作簿引用"
        path, reason = self._resolve(ref, parent_dir)
        if path is None:
            return None, reason
        key = str(path)
        if key in self._failures:
            return None, self._failures[key]
        book = self._books.get(key)
        if book is None:
            book, reason = self._open(path)
            if book is None:
                self._failures[key] = reason
                return None, reason
            self._books[key] = book
            self.loaded.append(path)
        return book, ref.sheet

    def close(self) -> None:
        for book in self._books.values():
            book.view.close()
        self._books.clear()

    # -------------------------------------------------------------- 内部实现
    def _resolve(self, ref: ExternalRef, parent_dir: Path | None) -> tuple[Path | None, str]:
        """把引用解析成磁盘路径：带目录的按目录，省略目录的按当前工作簿同目录。"""
        if ref.dir_part:
            base = Path(ref.dir_part)
            if not base.is_absolute():
                if parent_dir is None:
                    return None, f"无法解析外部工作簿的相对路径 {ref.dir_part}"
                base = parent_dir / base
        else:
            if parent_dir is None:
                return None, "未提供当前工作簿目录，无法解析同目录的外部引用"
            base = parent_dir
        if Path(ref.book).suffix:
            candidates = [base / ref.book]
        else:
            candidates = [base / f"{ref.book}{suffix}" for suffix in _CANDIDATE_SUFFIXES]
        for candidate in candidates:
            if candidate.is_file():
                return candidate, ""
        tried = "、".join(candidate.name for candidate in candidates)
        return None, f"本地找不到外部工作簿（已查找：{tried}）"

    def _open(self, path: Path) -> tuple[ExternalBook | None, str]:
        if path.suffix.lower() == ".xls":
            return None, f"外部工作簿 {path.name} 是旧版 .xls 格式，openpyxl 无法读取"
        try:
            size = path.stat().st_size
        except OSError as exc:
            return None, f"外部工作簿 {path.name} 无法读取（{exc}）"
        if size > MAX_EXTERNAL_FILE_BYTES:
            limit_mb = MAX_EXTERNAL_FILE_BYTES // (1024 * 1024)
            return None, f"外部工作簿 {path.name} 超过 {limit_mb} MB，已跳过本地试算"
        try:
            view = WorkbookView(path)
        except Exception as exc:  # noqa: BLE001 - 损坏/加密等情况一律降级为"未验证"
            return None, f"外部工作簿 {path.name} 无法打开（{type(exc).__name__}: {exc}）"
        values = {name: view.wb_values[name] for name in view.sheet_names}
        formulas = {name: view.wb_formulas[name] for name in view.sheet_names}
        return ExternalBook(path=path, values=values, formulas=formulas, view=view), ""
