"""运行配置与安全边界（目录白名单、函数白名单、模型参数）。"""
from __future__ import annotations

import os
from dataclasses import dataclass, field
from pathlib import Path

# ---------------------------------------------------------------- 模型相关默认值
DEFAULT_BASE_URL = "https://api.deepseek.com/chat/completions"
DEFAULT_MODEL = "deepseek-chat"
DEFAULT_TIMEOUT = 60
DEFAULT_TEMPERATURE = 0.0
# 校验失败后允许把错误信息回传模型重新生成的次数
DEFAULT_MAX_REPAIR_ROUNDS = 2

# ---------------------------------------------------------------- 表格文本化参数
# 表格以无损 TSV 全量交给模型，由模型自行判断表头行与数据起始行。
# 行数闸门：超过则只保留头尾两段，中间在提示里显式标注为已省略
DIGEST_MAX_ROWS = 200
DIGEST_HEAD_ROWS = 50
DIGEST_TAIL_ROWS = 10
# 单个单元格文本在描述里的最大长度
DIGEST_MAX_CELL_TEXT = 24
# 描述里最多列出的列数
DIGEST_MAX_COLUMNS = 40

# ---------------------------------------------------------------- 公式安全控制
# 允许出现在生成公式里的函数（白名单之外一律拒绝）
ALLOWED_FUNCTIONS: frozenset[str] = frozenset(
    """
    SUM SUMIF SUMIFS SUMPRODUCT SUBTOTAL PRODUCT
    AVERAGE AVERAGEIF AVERAGEIFS MEDIAN MODE
    COUNT COUNTA COUNTBLANK COUNTIF COUNTIFS
    MAX MAXIFS MIN MINIFS LARGE SMALL RANK RANK.EQ PERCENTILE QUARTILE
    STDEV STDEV.S STDEV.P STDEVP VAR VAR.S VAR.P
    ABS ROUND ROUNDUP ROUNDDOWN INT TRUNC MOD POWER SQRT SIGN EXP LN LOG LOG10
    CEILING FLOOR RAND RANDBETWEEN
    IF IFS IFERROR IFNA AND OR NOT XOR TRUE FALSE SWITCH
    ISBLANK ISNUMBER ISTEXT ISERROR ISERR ISNA ISEVEN ISODD NA
    VLOOKUP HLOOKUP XLOOKUP LOOKUP INDEX MATCH CHOOSE OFFSET
    LEN LEFT RIGHT MID FIND SEARCH SUBSTITUTE REPLACE REPT TRIM
    CONCAT CONCATENATE TEXTJOIN TEXT VALUE NUMBERVALUE UPPER LOWER PROPER CHAR CODE
    DATE DATEVALUE TODAY NOW YEAR MONTH DAY HOUR MINUTE SECOND WEEKDAY WEEKNUM
    EOMONTH EDATE DATEDIF DAYS NETWORKDAYS WORKDAY
    ROW ROWS COLUMN COLUMNS TRANSPOSE UNIQUE SORT FILTER
    """.split()
)

# 明确禁止的函数：可外发数据、执行外部内容或产生动态引用，属于高风险
FORBIDDEN_FUNCTIONS: frozenset[str] = frozenset(
    """
    INDIRECT HYPERLINK WEBSERVICE FILTERXML RTD CALL REGISTER REGISTER.ID
    EXEC DDE EVALUATE IMPORTDATA IMPORTXML IMPORTRANGE PY
    """.split()
)

# 允许写入的文件后缀（openpyxl 可写格式）
ALLOWED_SUFFIXES: frozenset[str] = frozenset({".xlsx", ".xlsm"})

MAX_EXCEL_ROW = 1_048_576
MAX_EXCEL_COLUMN = 16_384
# 单次请求最多写入的单元格数量，避免误操作大面积改表
MAX_WRITE_CELLS = 200


class ConfigError(RuntimeError):
    """配置缺失或非法。"""


class SecurityError(RuntimeError):
    """越权访问：超出目录白名单或文件类型不被允许。"""


def load_dotenv(path: str | os.PathLike[str] = ".env") -> None:
    """极简 .env 加载器：只处理 KEY=VALUE，不覆盖已存在的环境变量。"""
    env_path = Path(path)
    if not env_path.is_file():
        return
    for raw in env_path.read_text(encoding="utf-8").splitlines():
        line = raw.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, _, value = line.partition("=")
        key = key.strip()
        value = value.strip().strip('"').strip("'")
        if key and key not in os.environ:
            os.environ[key] = value


@dataclass
class Settings:
    """一次运行所需的全部配置。"""

    api_key: str = ""
    model: str = DEFAULT_MODEL
    base_url: str = DEFAULT_BASE_URL
    timeout: int = DEFAULT_TIMEOUT
    temperature: float = DEFAULT_TEMPERATURE
    max_repair_rounds: int = DEFAULT_MAX_REPAIR_ROUNDS
    # 显式代理地址（空表示不指定）；trust_env 为 False 时忽略系统/环境代理直连
    proxy: str = ""
    trust_env: bool = True
    # 目录白名单：所有读写的 Excel 必须落在这些目录之内
    allowed_roots: list[Path] = field(default_factory=list)
    log_dir: Path = Path("logs")
    backup: bool = True

    @classmethod
    def from_env(cls, workspace: str | os.PathLike[str] | None = None) -> "Settings":
        load_dotenv()
        roots_raw = os.environ.get("EXCELCR_ALLOWED_DIRS", "").strip()
        roots = [Path(p).expanduser().resolve() for p in roots_raw.split(os.pathsep) if p.strip()]
        if workspace:
            roots.insert(0, Path(workspace).expanduser().resolve())
        if not roots:
            roots = [Path.cwd().resolve()]
        return cls(
            api_key=os.environ.get("DEEPSEEK_API_KEY", "").strip(),
            model=os.environ.get("DEEPSEEK_MODEL", DEFAULT_MODEL).strip() or DEFAULT_MODEL,
            base_url=os.environ.get("DEEPSEEK_BASE_URL", DEFAULT_BASE_URL).strip() or DEFAULT_BASE_URL,
            timeout=int(os.environ.get("DEEPSEEK_TIMEOUT", DEFAULT_TIMEOUT)),
            temperature=float(os.environ.get("DEEPSEEK_TEMPERATURE", DEFAULT_TEMPERATURE)),
            max_repair_rounds=int(os.environ.get("EXCELCR_MAX_REPAIR", DEFAULT_MAX_REPAIR_ROUNDS)),
            proxy=os.environ.get("EXCELCR_PROXY", "").strip(),
            trust_env=os.environ.get("EXCELCR_TRUST_ENV", "1") != "0",
            allowed_roots=roots,
            log_dir=Path(os.environ.get("EXCELCR_LOG_DIR", "logs")),
            backup=os.environ.get("EXCELCR_BACKUP", "1") != "0",
        )

    def require_api_key(self) -> str:
        if not self.api_key:
            raise ConfigError(
                "未配置 DEEPSEEK_API_KEY。请复制 .env.example 为 .env 并填入密钥，"
                "或设置环境变量 DEEPSEEK_API_KEY。"
            )
        return self.api_key

    def resolve_path(self, path: str | os.PathLike[str], *, must_exist: bool = True) -> Path:
        """把用户给的路径限制在目录白名单内，返回绝对路径。"""
        target = Path(path).expanduser()
        if not target.is_absolute():
            target = (self.allowed_roots[0] / target)
        target = target.resolve()

        if target.suffix.lower() not in ALLOWED_SUFFIXES:
            raise SecurityError(
                f"只允许操作 {'/'.join(sorted(ALLOWED_SUFFIXES))} 文件，收到：{target.name}"
            )
        if not any(_is_within(target, root) for root in self.allowed_roots):
            allowed = "、".join(str(r) for r in self.allowed_roots)
            raise SecurityError(f"路径超出允许目录（{allowed}）：{target}")
        if must_exist and not target.is_file():
            raise FileNotFoundError(f"文件不存在：{target}")
        return target


def _is_within(target: Path, root: Path) -> bool:
    try:
        target.relative_to(root)
    except ValueError:
        return False
    return True
