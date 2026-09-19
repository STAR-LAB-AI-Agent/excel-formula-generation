"""零依赖 Web 演示台：标准库 http.server 承载「三栏答辩大屏」。

复用 excel_formula.FormulaService 的全部能力（意图识别 / 生成 / 校验 / 解释 /
写入），模型输出经 SSE 以打字机形式实时推给浏览器；不引入任何第三方依赖，
也不改动 excel_formula 包内代码。

启动：
    python webapp.py                     # 127.0.0.1:8765，自动打开浏览器
    python webapp.py --port 9000 --no-browser

接口约定（前端 web/index.html 只依赖这些）：
    GET  /                 三栏页面
    GET  /api/bootstrap    文件列表、思考档位、模型名、示例需求
    GET  /api/sheet        指定工作表的结构 + 网格数据（本地，0 Token）
    POST /api/chat         一句话 → SSE 事件流（stage / delta / result / error）
    POST /api/apply        用 proposal_id 把待确认方案写入文件
    POST /api/open         用系统默认程序打开当前表格（本机查看结果）
    POST /api/upload       上传本地 .xlsx/.xlsm 到工作目录（重名自动改名）
    POST /api/think        切换思考档位（1 / 0 / auto）
"""
from __future__ import annotations

import argparse
import json
import os
import socket
import subprocess
import sys
import threading
import uuid
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, urlparse

from excel_formula import data_validation as dv
from excel_formula.config import (
    ALLOWED_SUFFIXES,
    DIGEST_MAX_CELL_TEXT,
    ConfigError,
    SecurityError,
    Settings,
    normalize_thinking,
)
from excel_formula.evaluator import (
    ExcelError,
    UnsupportedFormula,
    evaluate_formula,
    format_value,
)
from excel_formula.excel_reader import SheetDigest, WorkbookView, cell_formula_text
from excel_formula.formula_parser import FormulaSyntaxError
from excel_formula.intent import (
    INTENT_DESCRIBE,
    INTENT_DROPDOWN,
    INTENT_EXPLAIN,
    INTENT_FORMAT,
    INTENT_GENERATE,
    INTENT_VALIDATE,
    classify,
    file_candidates,
)
from excel_formula.llm_client import LLMError
from excel_formula.logger import get_logger
from excel_formula.pipeline import FormulaService, Proposal
from excel_formula.validator import validate_formula

WEB_DIR = Path(__file__).resolve().parent / "web"

# 示例需求数据：随 /api/bootstrap 返回（前端目前以输入区简介文案呈现，数据保留备用）
SAMPLE_REQUESTS = [
    {"label": "生成并写入", "text": "帮我在 G2 算每个科目的总分，填充到 G6"},
    {"label": "本地校验 · 0 Token", "text": '=SUMIF(B2:B6,">85") 这个公式对不对？'},
    {"label": "解释公式", "text": "解释一下 G2 里的公式"},
    {"label": "表结构概览 · 0 Token", "text": "这个表有哪些列？"},
    {"label": "表格加框 · 0 Token", "text": "将 A14 到 B18 的表格范围框起来"},
    {"label": "下拉列表 · 0 Token", "text": "把 B3 设置为下拉列表，选项为：华东、华南、华北"},
]

THINKING_LABELS = {
    "enabled": "深度思考常开（最准也最慢）",
    "disabled": "关闭思考（最快）",
    "auto": "自动（按需求复杂度切换）",
    "": "服务端默认（思考开启、强度 high）",
}

_INTENT_LABELS = {
    INTENT_GENERATE: "生成公式",
    INTENT_VALIDATE: "校验公式",
    INTENT_EXPLAIN: "解释公式",
    INTENT_DESCRIBE: "表结构概览",
    INTENT_FORMAT: "表格加框",
    INTENT_DROPDOWN: "设置下拉列表",
}

# 会被翻译成 400 的输入类错误（安全、路径、参数）
_BAD_INPUT = (
    SecurityError,
    FileNotFoundError,
    KeyError,
    ValueError,
    FormulaSyntaxError,
    ConfigError,
)

# 单表本地试算的公式格上限：防止极端文件把页面拖成"转圈"
_GRID_EVAL_LIMIT = 500

# 上传体积上限：答辩用表远小于此，防止超大请求体撑爆内存
_MAX_UPLOAD_BYTES = 30 * 1024 * 1024


def _grid_text(value: object) -> str:
    """单格展示文本：与 digest 相同的清理规则（换行/制表归一、超长截断）。"""
    if value is None:
        return ""
    text = format_value(value)  # ExcelError -> #code；bool/float/DateValue 归一
    text = text.replace("\n", " ").replace("\t", " ").strip()
    return text if len(text) <= DIGEST_MAX_CELL_TEXT else text[: DIGEST_MAX_CELL_TEXT - 1] + "…"


def _evaluate_cell(
    formula: str, sheet_name: str, sheets: dict, formula_sheets: dict,
    current_cell: tuple[int, int],
) -> str:
    """本地试算单格公式：成功返回结果文本，失败退回公式原文（悬停仍可查看）。

    expand_uncached：引用的其他公式格没有缓存时也递归求值——文件被 openpyxl
    保存过（写入后）会丢掉全部公式缓存，没有它链式公式只能显示原文。
    current_cell 是公式所在位置（行、列），供 ROW()/COLUMN() 无参形式求值。
    """
    try:
        _, ast = validate_formula(formula, sheet=sheet_name)
        if ast is None:
            return formula
        value = evaluate_formula(
            ast, sheets, sheet_name, formula_sheets=formula_sheets,
            expand_uncached=True, current_cell=current_cell,
        )
    except (UnsupportedFormula, ZeroDivisionError, ValueError, TypeError, IndexError, KeyError):
        return formula
    return _grid_text(value)


def build_display_grid(view: WorkbookView, digest: SheetDigest) -> dict:
    """把 digest 转成前端网格：公式格优先展示计算结果——Excel 缓存值 → 本地试算 →
    退回公式原文；每个单元格附带平行公式表（formulas），供鼠标悬停查看原公式。

    缓存值优先是因为它来自 Excel/WPS 最近一次真实计算；openpyxl 写出的公式没有
    缓存，用与校验环节相同的本地求值引擎试算（0 Token），做不到才退回公式文本。
    """
    ws_formulas = view.wb_formulas[digest.name]
    ws_values = view.wb_values[digest.name]
    sheets = None
    formula_sheets = None
    evaluated = 0
    rows = []
    for row_index, texts in digest.rows:
        values: list[str] = []
        formulas: list[str] = []
        for offset, text in enumerate(texts):
            col = offset + 1
            cell = ws_formulas.cell(row=row_index, column=col)
            raw = cell_formula_text(cell) if cell.data_type == "f" else None
            if raw is None:
                # 非公式格；或对象型公式（模拟运算表 DataTableFormula 等，openpyxl
                # 不给出公式文本）——展示 digest 文本（已退回缓存结果），悬停留空，
                # 也确保不会有对象进入 JSON 序列化
                values.append(text)
                formulas.append("")
                continue
            formulas.append(raw)
            cached = ws_values.cell(row=row_index, column=col).value
            if cached is not None:
                values.append(_grid_text(cached))
                continue
            if sheets is None:  # 真遇到无缓存公式才组装求值上下文
                sheets = {name: view.wb_values[name] for name in view.sheet_names}
                formula_sheets = {name: view.wb_formulas[name] for name in view.sheet_names}
            if evaluated < _GRID_EVAL_LIMIT:
                evaluated += 1
                values.append(
                    _evaluate_cell(raw, digest.name, sheets, formula_sheets, (row_index, col))
                )
            else:
                values.append(raw)  # 公式过多：放弃试算，保留原文
        rows.append({"row": row_index, "values": values, "formulas": formulas})
    payload = digest.to_dict()
    payload["rows"] = rows
    warning = dv.extension_warning(view.path)
    payload["dropdowns"] = dv.dropdown_metadata(view, digest.name, warning)
    payload["dropdown_warning"] = warning
    return payload


def _open_externally(path: Path) -> None:
    """用系统默认程序打开文件（Windows 下即 Excel/WPS）；立即返回，不阻塞服务线程。"""
    if hasattr(os, "startfile"):  # Windows
        os.startfile(str(path))  # type: ignore[attr-defined]
        return
    opener = "open" if sys.platform == "darwin" else "xdg-open"
    subprocess.Popen([opener, str(path)], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)


def _unique_upload_path(directory: Path, name: str) -> Path:
    """上传落点：重名自动加 _1/_2…，绝不覆盖用户已有文件。"""
    target = directory / name
    if not target.exists():
        return target
    stem, suffix = target.stem, target.suffix
    for index in range(1, 1000):
        candidate = directory / f"{stem}_{index}{suffix}"
        if not candidate.exists():
            return candidate
    raise ValueError(f"同名文件过多，请先整理目录：{name}")


class ProposalStore:
    """单用户演示会话的待确认方案暂存。

    生成接口把 Proposal 存在服务端内存里并返回短 token；浏览器点「确认写入」
    时凭 token 取回——这样公式预览与写入之间隔着的用户思考时间不影响正确性。
    只保留最近若干条，避免长时间演示时无限增长。
    """

    def __init__(self, limit: int = 20):
        self._items: dict[str, Proposal] = {}
        self._order: list[str] = []
        self._limit = limit
        self._lock = threading.Lock()

    def put(self, proposal: Proposal) -> str:
        token = uuid.uuid4().hex[:12]
        with self._lock:
            self._items[token] = proposal
            self._order.append(token)
            while len(self._order) > self._limit:
                self._items.pop(self._order.pop(0), None)
        return token

    def get(self, token: str) -> Proposal:
        with self._lock:
            proposal = self._items.get(token)
        if proposal is None:
            raise ValueError("这个方案已失效（可能已写入或被新方案顶掉），请重新生成")
        return proposal

    def drop(self, token: str) -> None:
        with self._lock:
            self._items.pop(token, None)


class Handler(BaseHTTPRequestHandler):
    """路由 + SSE 推送。线程模型沿用 ThreadingHTTPServer（每个请求一个线程）。"""

    protocol_version = "HTTP/1.1"  # 启用长连接与 chunked 流式
    server_version = "ExcelCR-Web/1.0"

    # ---------------------------------------------------------------- 基础设施
    @property
    def service(self) -> FormulaService:
        return self.server.service  # type: ignore[attr-defined]

    @property
    def proposals(self) -> ProposalStore:
        return self.server.proposals  # type: ignore[attr-defined]

    def log_message(self, fmt: str, *args) -> None:  # noqa: A003
        logger = getattr(self.service, "logger", None)
        if logger:
            logger.info("web %s", fmt % args)

    def _send_json(self, payload: dict, status: int = 200) -> None:
        body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(body)

    def _send_bytes(self, body: bytes, content_type: str) -> None:
        self.send_response(200)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(body)

    def _read_json(self) -> dict:
        length = int(self.headers.get("Content-Length") or 0)
        raw = self.rfile.read(length) if length > 0 else b""
        if not raw:
            return {}
        try:
            payload = json.loads(raw.decode("utf-8"))
        except (json.JSONDecodeError, UnicodeDecodeError) as exc:
            raise ValueError(f"请求体不是合法 JSON：{exc}") from exc
        if not isinstance(payload, dict):
            raise ValueError("请求体必须是 JSON 对象")
        return payload

    # ---------------------------------------------------------------- SSE 流
    def _sse_open(self) -> None:
        self.send_response(200)
        self.send_header("Content-Type", "text/event-stream; charset=utf-8")
        self.send_header("Cache-Control", "no-cache")
        self.send_header("X-Accel-Buffering", "no")  # 中间有反代时也别缓冲
        self.send_header("Transfer-Encoding", "chunked")
        self.end_headers()

    def _sse_send(self, payload: dict) -> None:
        """一个事件 = 一行 data: + JSON + 空行，按 chunked 编码落盘到 socket。"""
        chunk = f"data: {json.dumps(payload, ensure_ascii=False)}\n\n".encode("utf-8")
        self.wfile.write(f"{len(chunk):X}\r\n".encode("ascii"))
        self.wfile.write(chunk)
        self.wfile.write(b"\r\n")
        self.wfile.flush()

    def _sse_close(self) -> None:
        self.wfile.write(b"0\r\n\r\n")
        self.wfile.flush()

    # ---------------------------------------------------------------- GET
    def do_GET(self) -> None:  # noqa: N802
        parsed = urlparse(self.path)
        try:
            if parsed.path in {"/", "/index.html"}:
                return self._serve_page()
            if parsed.path == "/api/bootstrap":
                return self._send_json(self._bootstrap())
            if parsed.path == "/api/sheet":
                return self._send_json(self._sheet_info(parse_qs(parsed.query)))
            self._send_json({"ok": False, "error": f"未知接口：{parsed.path}"}, 404)
        except _BAD_INPUT as exc:
            self._send_json({"ok": False, "error": str(exc)}, 400)

    def _serve_page(self) -> None:
        page = WEB_DIR / "index.html"
        if not page.is_file():
            return self._send_json({"ok": False, "error": "缺少 web/index.html"}, 500)
        self._send_bytes(page.read_bytes(), "text/html; charset=utf-8")

    def _bootstrap(self) -> dict:
        settings = self.service.settings
        return {
            "ok": True,
            "workspace": str(settings.allowed_roots[0]),
            "files": [p.name for p in self._list_files()],
            "model": settings.model,
            "thinking": settings.thinking,
            "thinking_label": THINKING_LABELS.get(settings.thinking, settings.thinking),
            "has_key": bool(settings.api_key),
            "samples": SAMPLE_REQUESTS,
        }

    def _list_files(self) -> list[Path]:
        """当前目录直属的 .xlsx/.xlsm，规则与 CLI 启动编号列表一致。"""
        directory = self.service.settings.allowed_roots[0]
        found: list[Path] = []
        try:
            paths = sorted(directory.iterdir(), key=lambda p: (p.name.casefold(), p.name))
        except OSError:
            return found
        for path in paths:
            if path.name.startswith("~$") or path.suffix.lower() not in ALLOWED_SUFFIXES:
                continue
            try:
                self.service.settings.resolve_path(path)
            except (SecurityError, OSError):
                continue
            found.append(path)
        return found

    def _sheet_info(self, query: dict) -> dict:
        name = (query.get("file") or [""])[0]
        sheet = (query.get("sheet") or [""])[0] or None
        path = self.service.settings.resolve_path(name)
        with dv.WRITE_LOCK:
            version = dv.file_version(path)
            with WorkbookView(path) as view:
                sheet_name = view.resolve_sheet(sheet)
                digest = view.digest(sheet_name)
                result = {
                    "ok": True,
                    "file": path.name,
                    "file_version": version,
                    "sheets": view.overview(),
                    "current": sheet_name,
                    "grid": build_display_grid(view, digest),
                }
            dv.require_version(path, version)
            return result

    # ---------------------------------------------------------------- POST
    def do_POST(self) -> None:  # noqa: N802
        parsed = urlparse(self.path)
        try:
            if parsed.path == "/api/upload":  # 请求体是原始文件字节，不走 JSON 解析
                return self._upload(parse_qs(parsed.query))
            payload = self._read_json()
            if parsed.path == "/api/chat":
                return self._chat(payload)
            if parsed.path == "/api/apply":
                return self._apply(payload)
            if parsed.path == "/api/dropdown/select":
                return self._select_dropdown(payload)
            if parsed.path == "/api/open":
                return self._open_file(payload)
            if parsed.path == "/api/think":
                return self._switch_thinking(payload)
            self._send_json({"ok": False, "error": f"未知接口：{parsed.path}"}, 404)
        except _BAD_INPUT as exc:
            self._send_json({"ok": False, "error": str(exc)}, 400)
        except PermissionError as exc:
            self._send_json({"ok": False, "error": str(exc)}, 400)
        except (LLMError, UnsupportedFormula) as exc:
            self._send_json({"ok": False, "error": f"模型调用失败：{exc}"}, 502)
        except (BrokenPipeError, ConnectionResetError, ConnectionAbortedError):
            self.close_connection = True

    # ---------------------------------------------------------------- /api/chat
    def _chat(self, payload: dict) -> None:
        """SSE 入口：一旦开流，所有错误都转为 error 事件，不再走 JSON 响应。"""
        text = str(payload.get("request") or "").strip()
        if not text:
            raise ValueError("请先说一句需求，例如「帮我在 G2 算总分，填充到 G6」")
        file = str(payload.get("file") or "").strip() or None
        sheet = str(payload.get("sheet") or "").strip() or None

        self._sse_open()
        try:
            self._run_chat(text, file, sheet)
        except _BAD_INPUT as exc:
            self._sse_send({"type": "error", "message": str(exc)})
        except (LLMError, UnsupportedFormula) as exc:
            self._sse_send({"type": "error", "message": f"模型调用失败：{exc}"})
        except PermissionError as exc:
            self._sse_send({"type": "error", "message": str(exc)})
        except (BrokenPipeError, ConnectionResetError, ConnectionAbortedError):
            self.close_connection = True  # 浏览器关掉页面：静默收场
            return
        finally:
            try:
                self._sse_close()
            except OSError:
                pass

    def _run_chat(self, text: str, file: str | None, sheet: str | None) -> None:
        """按意图分发。规则识别与读表都在本地完成，0 Token 的环节如实标注。"""
        intent = classify(text)
        label = _INTENT_LABELS.get(intent.kind, intent.kind)
        note = "，识别为新建表格（自动套表头浅蓝底与边框）" if intent.new_table else ""
        self._sse_send({
            "type": "stage",
            "text": f"意图识别：{label}（本地规则，0 Token）{note}",
        })
        path = self._request_file(intent.file, file)
        sheet_arg = intent.sheet or sheet

        if intent.kind == INTENT_DESCRIBE:
            result = self.service.describe(path, sheet_arg)
            self._sse_send({
                "type": "result",
                "kind": "describe",
                "payload": {
                    "file": Path(result["file"]).name,
                    "sheets": result["sheets"],
                    "digest_text": result["digest_text"],
                },
            })
            return

        if intent.kind == INTENT_VALIDATE:
            result = self.service.validate(
                path, intent.formula, sheet=sheet_arg, target=intent.target
            )
            self._sse_send({"type": "result", "kind": "validate", "payload": result})
            return

        if intent.kind == INTENT_EXPLAIN:
            self._sse_send({"type": "stage", "text": "调用 DeepSeek 解释公式（SSE 流式）…"})
            result = self.service.explain(
                intent.formula,
                file=path,
                sheet=sheet_arg,
                cell=intent.cell,
                on_delta=self._send_delta,
            )
            self._sse_send({"type": "result", "kind": "explain", "payload": result})
            return

        if intent.kind == INTENT_DROPDOWN:
            self._sse_send({"type": "stage", "text": "解析列表数据验证需求…"})
            proposal = self.service.propose_dropdown(
                path, text, sheet=sheet_arg,
                # 本地未命中而走模型兜底时，追加真实的处理阶段（保证 0 Token 文案诚实）
                on_stage=lambda note: self._sse_send({"type": "stage", "text": note}),
            )
            payload = proposal.to_dict()
            payload["proposal_id"] = self.proposals.put(proposal)
            payload["file_name"] = proposal.file.name
            self._sse_send({"type": "result", "kind": "generate", "payload": payload})
            return

        if intent.kind == INTENT_FORMAT:
            # “把范围框起来”：本地套细边框，不调用模型（模型只会给公式，给不了格式）
            self._sse_send({"type": "stage", "text": "本地套用边框：不调用模型（0 Token）"})
            proposal = self.service.frame_table(
                path, sheet=sheet_arg, cell_range=intent.table_range
            )
            token = self.proposals.put(proposal)
            payload = proposal.to_dict()
            payload["proposal_id"] = token
            payload["file_name"] = Path(payload["file"]).name
            self._sse_send({"type": "result", "kind": "generate", "payload": payload})
            return

        # INTENT_GENERATE：全链路
        self._sse_send({"type": "stage", "text": "读取表结构与同簿来源表（本地，0 Token）"})
        self._sse_send({
            "type": "stage",
            "text": "调用 DeepSeek 生成公式：校验失败自动回传修复，最多重试 2 轮",
        })
        proposal = self.service.propose(
            path,
            intent.request or text,
            sheet=sheet_arg,
            target=intent.target,
            new_table=intent.new_table,
            on_delta=self._send_delta,
        )
        token = self.proposals.put(proposal)
        payload = proposal.to_dict()
        payload["proposal_id"] = token
        payload["file_name"] = Path(payload["file"]).name
        self._sse_send({"type": "result", "kind": "generate", "payload": payload})

    def _send_delta(self, piece: str, kind: str) -> None:
        """模型增量回调：reasoning=思考链 / content=正文，浏览器端按轨打字机展示。"""
        self._sse_send({
            "type": "delta",
            "kind": "reasoning" if kind == "reasoning" else "content",
            "text": piece,
        })

    def _request_file(self, from_text: str | None, from_client: str | None) -> Path:
        """需求原文里提到的文件优先（可能粘着虚词，逐个候选试），否则用界面选中的。"""
        settings = self.service.settings
        for raw in (from_text, from_client):
            if not raw:
                continue
            for candidate in file_candidates(raw):
                try:
                    return settings.resolve_path(candidate)
                except (SecurityError, FileNotFoundError):
                    continue
        raise ValueError("请先在中间栏选择一个 Excel 文件，或在需求里写明文件名")

    # ---------------------------------------------------------------- /api/apply
    def _apply(self, payload: dict) -> None:
        token = str(payload.get("proposal_id") or "").strip()
        if not token:
            raise ValueError("缺少 proposal_id，请先生成公式")
        with dv.WRITE_LOCK:
            proposal = self.proposals.get(token)
            applied = self.service.apply(proposal)
            self.proposals.drop(token)
        result = dict(applied)
        result["file_name"] = Path(result["file"]).name
        if result.get("backup"):
            result["backup_name"] = Path(result["backup"]).name
        self._send_json({"ok": True, "applied": result})

    # ---------------------------------------------------------------- /api/dropdown/select
    def _select_dropdown(self, payload: dict) -> None:
        """网页确认选择后提交，服务端按规则和索引取值并安全保存。"""
        if not isinstance(payload, dict):
            raise ValueError("请求必须为 JSON 对象")
        for key in ("file", "sheet", "cell", "rule_id", "file_version"):
            if not isinstance(payload.get(key), str) or not payload[key].strip():
                raise ValueError(f"缺少或无效参数：{key}")
        origin = self.headers.get("Origin")
        if origin and urlparse(origin).netloc != self.headers.get("Host"):
            raise ValueError("拒绝跨站写入请求")
        if self.headers.get_content_type() != "application/json":
            raise ValueError("下拉选择请求必须使用 application/json")
        applied = self.service.select_dropdown(
            payload["file"], sheet=payload["sheet"], cell=payload["cell"],
            rule_id=payload["rule_id"], option_index=payload.get("option_index"),
            file_version=payload["file_version"],
        )
        applied["file_name"] = Path(applied["file"]).name
        if applied.get("backup"):
            applied["backup_name"] = Path(applied["backup"]).name
        self._send_json({"ok": True, "applied": applied})

    # ---------------------------------------------------------------- /api/open
    def _open_file(self, payload: dict) -> None:
        """用系统默认程序打开工作目录内的表格：演示时对照「网页预览 vs Excel 实际」。"""
        name = str(payload.get("file") or "").strip()
        if not name:
            raise ValueError("请先选择要打开的 Excel 文件")
        path = self.service.settings.resolve_path(name)
        if not path.is_file():
            raise FileNotFoundError(f"文件不存在：{path.name}")
        _open_externally(path)
        self._send_json({"ok": True, "file": path.name})

    # ---------------------------------------------------------------- /api/upload
    def _upload(self, query: dict) -> None:
        """接收浏览器选中的本地表格，校验后存入工作目录（重名自动改名）。"""
        length = int(self.headers.get("Content-Length") or 0)
        if length <= 0:
            raise ValueError("上传内容为空")
        if length > _MAX_UPLOAD_BYTES:
            self.close_connection = True  # 超限不读请求体，直接断开连接
            raise ValueError(f"文件超过 {_MAX_UPLOAD_BYTES // 1024 // 1024} MB 上限")

        data = self.rfile.read(length)  # 先读完，后续报错也不会把字节留在连接里
        raw_name = (query.get("name") or [""])[0]
        name = Path(raw_name).name
        if not name or name != raw_name:
            raise ValueError("文件名无效（不能包含目录）")
        if name.startswith("~$"):
            raise ValueError("这是 Excel 临时锁文件，请选择正式表格")
        if Path(name).suffix.lower() not in ALLOWED_SUFFIXES:
            raise ValueError(f"只支持 {' / '.join(sorted(ALLOWED_SUFFIXES))} 文件")

        directory = self.service.settings.allowed_roots[0]
        target = _unique_upload_path(directory, name)
        try:
            target.write_bytes(data)
        except OSError as exc:
            raise PermissionError(f"写入失败：{exc}") from exc
        try:
            with WorkbookView(target):  # 读一遍确认不是坏文件，别等演示时才炸
                pass
        except Exception as exc:
            target.unlink(missing_ok=True)
            raise ValueError(f"{name} 不是有效的 Excel 文件（{exc.__class__.__name__}）") from exc

        self._send_json({
            "ok": True,
            "file": target.name,
            "renamed": target.name != name,
            "files": [p.name for p in self._list_files()],
        })

    # ---------------------------------------------------------------- /api/think
    def _switch_thinking(self, payload: dict) -> None:
        value = normalize_thinking(str(payload.get("value") or ""))
        self.service.settings.thinking = value  # client 与 pipeline 共享同一份 settings
        if self.service.logger:
            self.service.logger.info("web 思考模式切换：%s", value)
        self._send_json({
            "ok": True,
            "thinking": value,
            "label": THINKING_LABELS.get(value, value),
        })


class _ExclusiveServer(ThreadingHTTPServer):
    """Windows 下默认 SO_REUSEADDR 允许第二个进程重复绑定同一端口，
    会让「端口顺延」失效（残留实例与新实例抢同一端口，请求去向不定）。
    Windows 改用独占绑定；其他平台保持原有行为。"""

    allow_reuse_address = os.name != "nt"

    def server_bind(self) -> None:
        if os.name == "nt":
            self.socket.setsockopt(socket.SOL_SOCKET, socket.SO_EXCLUSIVEADDRUSE, 1)
        super().server_bind()


def make_server(
    service: FormulaService, host: str = "127.0.0.1", port: int = 8765
) -> ThreadingHTTPServer:
    """组装可注入 service 的服务器；测试用它挂 FakeClient 驱动。"""
    server = _ExclusiveServer((host, port), Handler)
    server.service = service  # type: ignore[attr-defined]
    server.proposals = ProposalStore()  # type: ignore[attr-defined]
    return server


def _bind_server(
    service: FormulaService, host: str, port: int, attempts: int = 10
) -> ThreadingHTTPServer | None:
    """按端口顺延绑定服务器；连续被占用时返回 None。"""
    for candidate in range(port, port + attempts):
        try:
            return make_server(service, host=host, port=candidate)
        except OSError:
            continue
    return None


def start_in_background(
    service: FormulaService,
    *,
    host: str = "127.0.0.1",
    port: int = 8765,
    open_browser: bool = True,
) -> tuple[ThreadingHTTPServer | None, str | None]:
    """在守护线程里启动演示台（供 main.py 复用），返回 (server, url)。

    与调用方共享同一个 FormulaService：网页里切换的思考档位立即作用于 CLI，
    反之亦然。主程序退出时守护线程随进程结束，无需显式关闭。
    """
    server = _bind_server(service, host, port)
    if server is None:
        return None, None
    threading.Thread(target=server.serve_forever, daemon=True).start()
    url = f"http://{host}:{server.server_address[1]}/"
    if open_browser:
        webbrowser.open(url)
    return server, url


def main(argv: list[str] | None = None) -> int:
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8")  # type: ignore[union-attr]
        except (AttributeError, ValueError):
            pass

    parser = argparse.ArgumentParser(
        prog="webapp", description="AI Excel 公式生成平台（零依赖，三栏大屏）"
    )
    parser.add_argument("--host", default="127.0.0.1", help="监听地址（默认仅本机）")
    parser.add_argument("--port", type=int, default=8765, help="监听端口，被占用时自动顺延")
    parser.add_argument("--no-browser", action="store_true", help="不自动打开浏览器")
    parser.add_argument("--workspace", help="允许访问的目录（默认当前工作目录）")
    args = parser.parse_args(argv)

    settings = Settings.from_env(workspace=args.workspace or Path.cwd())
    logger = get_logger(settings.log_dir)
    service = FormulaService(settings, logger)

    server = _bind_server(service, args.host, args.port)
    if server is None:
        print(f"× 端口 {args.port}-{args.port + 9} 都被占用，无法启动。", file=sys.stderr)
        return 1

    url = f"http://{args.host}:{server.server_address[1]}/"
    thinking_label = THINKING_LABELS.get(settings.thinking, settings.thinking)
    print("=" * 58)
    print("  AI Excel 公式生成平台")
    print(f"  地址：{url}")
    print(f"  模型：{settings.model}    思考档位：{thinking_label}")
    if not settings.api_key:
        print("  ⚠ 未配置 DEEPSEEK_API_KEY：生成/解释不可用，概览与校验仍可演示")
    print("  Ctrl+C 停止")
    print("=" * 58)

    if not args.no_browser:
        threading.Timer(0.8, lambda: webbrowser.open(url)).start()
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\n已停止。")
    finally:
        server.server_close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
