# AI Excel 公式生成（ExcelCR）

《智能体开发实战》课程实验 · 选题 23「AI Excel 公式生成」

用一句中文说出你要算什么，程序自动读表结构、调用 DeepSeek 生成 Excel 公式、**在本地严格校验**、
预览并经你确认后用 openpyxl 写回文件，同时给出公式解释、预期结果和可追溯日志。

## 1. 技术路线

```
用户自然语言 + Excel 文件
        ↓
openpyxl 逐格读出工作表（只搬运，不推断表头与列类型）
        ↓
无损 TSV 文本化（带行号/列字母坐标系，超 200 行才头尾截断）
        ↓
DeepSeek API：自行判断表头行与数据区 → 输出公式 JSON
        ↓
本地校验：语法 AST + 函数白名单 + 引用范围 + 循环引用 + 参数个数
        ↓
校验失败？→ 只回传公式与错误信息给 DeepSeek，最多重试 2 次
        ↓
预览（公式 / 影响单元格 / 预期结果 / 是否覆盖旧内容）+ 用户确认
        ↓
openpyxl 写入并保存（自动备份原文件）
        ↓
返回结果 + Token 用量 + logs/excelcr.log
```

## 2. 支持的三类自然语言意图（外加一个辅助意图）

| 意图 | 用户说法示例 | 走到哪一步 |
|---|---|---|
| ① 生成公式并写入 | "帮我在 G2 算每个科目的总分，填充到 G6" | 全链路，含确认与写入 |
| ② 校验公式 | "=SUMIF(B2:B6,\">85\") 这个公式对不对？" | 纯本地校验 + 本地试算，**零 Token** |
| ③ 解释公式 | "解释一下 G2 里的公式" | 读单元格公式 → DeepSeek 讲解 |
| ④ 表结构概览（辅助） | "这个表有哪些列？" | 纯本地读取，**零 Token** |

意图识别由 `excel_formula/intent.py` 用规则完成，不花 Token；只有真正需要"理解语义/生成公式"时才调用模型。

## 3. 安装与运行

```powershell
# 1) 创建并激活虚拟环境（Python 3.8 及以上；实测 3.8.4 与 3.12.7 均通过全部测试）
python -m venv venv
.\venv\Scripts\Activate.ps1

# 2) 安装依赖（注意：必须装进上面这个 venv，PyCharm 的运行配置用的就是它）
pip install -r requirements.txt
python -c "import requests, openpyxl"   # 自检，无报错即依赖就位

# 3) 配置密钥（不要提交 .env）
copy .env.example .env
notepad .env      # 填入 DEEPSEEK_API_KEY

# 4) 交互模式
python main.py

# 或单条指令模式
python main.py "帮我在G2算每个科目的总分，填充到G6"
```

交互模式内置指令：`:file <路径>` 切换文件、`:sheet <表名>` 指定工作表、`:info` 查看表结构（零 Token）、
`:help`、`:quit`。

### 一次真实交互长什么样

```
你 > 帮我在G2算每个科目的总分，填充到G6
… 正在读取表结构并生成公式

----- 公式预览 -----
文件: 测试数据.xlsx    工作表: Sheet1
目标: G2    公式: =SUM(B2:F2)
说明: 对B2到F2的五名学生分数求和
填充: G2 → G6（共 5 个单元格）
预期结果: 438（本地独立计算，未写入文件）
--------------------
[本次消耗] 模型调用 1 次，输入 421 tokens，输出 58 tokens，耗时 2.31s
确认写入 Sheet1!G2, G3, G4, G5, G6 ？[y/N] y
√ 已写入 5 个单元格 → 测试数据.xlsx
  原文件已备份：测试数据_20260910_142530.xlsx
  日志：logs/excelcr.log
```

## 4. 底层 Script/CLI（可独立运行、可被 Skill 调用）

```powershell
# 表结构概览（含交给模型的整表 TSV）
python -m excel_formula.cli describe --file 测试数据.xlsx

# 生成公式（默认只预览，不改文件）
python -m excel_formula.cli generate --file 测试数据.xlsx --request "算每科总分" --target G2

# 确认无误后写入
python -m excel_formula.cli generate --file 测试数据.xlsx --request "算每科总分" --target G2 --apply --yes

# 校验公式并本地试算（零 Token）
python -m excel_formula.cli validate --file 测试数据.xlsx --formula "=SUM(B2:F2)" --target G2

# 解释公式
python -m excel_formula.cli explain --file 测试数据.xlsx --cell G2

# 直接写入自己写好的公式（仍会先校验）
python -m excel_formula.cli write --file 测试数据.xlsx --target G2 --formula "=SUM(B2:F2)" --fill-to G6 --yes

# 一句话自动判断意图
python -m excel_formula.cli nl "这个表有哪些列？" --file 测试数据.xlsx
```

加 `--json` 输出结构化结果（供智能体解析）；退出码：`0` 成功 / `2` 校验未通过或需追问 /
`3` 参数与安全错误 / `4` 模型或网络错误。

## 5. Skill 集成

Skill 说明位于 [`skills/excel-formula/SKILL.md`](skills/excel-formula/SKILL.md)，包含使用场景、
参数表、结果格式、3 个完整示例和失败处理策略。智能体（nanobot 或其他框架）按 SKILL.md 调用上面的 CLI 即可：

```
用户自然语言 → 智能体 → SKILL.md → python -m excel_formula.cli → openpyxl / DeepSeek → 结果
```

Skill 中强制约定：**先预览、展示给用户、拿到同意后才允许带 `--apply` 再执行一次**。

## 6. 代码结构

| 文件 | 职责 |
|---|---|
| `main.py` | 自然语言交互入口（REPL + 单条指令），负责追问与确认 |
| `excel_formula/cli.py` | 命令行接口，6 个子命令，可独立测试 |
| `excel_formula/intent.py` | 规则意图识别与参数抽取（文件/工作表/目标单元格/填充范围/公式） |
| `excel_formula/excel_reader.py` | openpyxl 读取、**无损 TSV 文本化**（不推断表头与类型，只做行数闸门） |
| `excel_formula/llm_client.py` | DeepSeek 调用、提示词、JSON 解析、Token 统计 |
| `excel_formula/formula_parser.py` | 公式分词器 + 递归下降语法分析器（产出 AST） |
| `excel_formula/validator.py` | 静态校验：语法、函数白名单、参数个数、引用范围、循环引用 |
| `excel_formula/evaluator.py` | 公式独立求值器（Python 重算一遍，用于"公式自动验证"） |
| `excel_formula/writer.py` | 备份、相对引用填充展开、写入保存 |
| `excel_formula/pipeline.py` | 流水线编排：生成 → 校验 → 重试 → 预览 → 写入（支持一次多个目标单元格） |
| `excel_formula/config.py` | 配置、目录白名单、函数白/黑名单、写入上限 |
| `excel_formula/logger.py` | 日志与密钥脱敏 |
| `scripts/measure_digest.py` | 提示词规模检查脚本（文本化后有多大、是否触发截断） |
| `skills/excel-formula/SKILL.md` | 供智能体调用的 Skill 说明 |
| `tests/` | 75 个自动化测试用例（不需要联网）+ 手工验收用例清单 |

## 7. 开源项目集成

| 项目 | 版本 | 许可证 | 本项目中的实际用途 |
|---|---|---|---|
| [openpyxl](https://foss.heptapod.net/openpyxl/openpyxl) | 3.1.5 | MIT | 读取工作表结构与缓存值（`load_workbook` 双视图）、写入公式并保存、`formula.translate.Translator` 做相对引用平移 |
| [requests](https://github.com/psf/requests) | 2.32.3 | Apache-2.0 | 调用 DeepSeek 的 OpenAI 兼容 `/chat/completions` 接口 |
| [pytest](https://github.com/pytest-dev/pytest) | ≥7.4 | MIT | 运行自动化测试 |

核心 Excel 能力（结构解析、公式写入、引用平移）完全由 openpyxl 支撑，本项目不重新实现 xlsx 读写。
DeepSeek 为模型服务（需自备密钥），不属于开源依赖。

## 8. 安全与异常处理

- **目录白名单**：所有路径经 `Settings.resolve_path` 收敛，越界抛 `SecurityError`；默认仅允许当前工作目录。
- **文件类型限制**：只接受 `.xlsx/.xlsm`，`.csv/.exe` 等一律拒绝。
- **函数黑白名单**：`INDIRECT/HYPERLINK/WEBSERVICE/FILTERXML/RTD/CALL/DDE...` 直接拒绝（可外发数据或动态求值）；
  白名单外的函数也拒绝，防止模型编造函数名。
- **循环引用拦截**：公式引用区域包含目标单元格时拒绝写入。
- **写入前确认**：预览 → 用户输入 `y` 才写；非交互环境必须显式 `--yes`，否则拒绝执行。
- **写入前备份**：原文件复制到 `backups/<名称>_<时间戳>.xlsx`，可人工回退。
- **写入规模上限**：单次最多 200 个单元格（`MAX_WRITE_CELLS`），避免误伤整表。
- **覆盖提示**：目标单元格已有值或公式时，预览里用 `⚠ 将覆盖已有内容` 明确列出。
- **密钥保护**：密钥只从环境变量/`.env` 读取，只出现在请求头；日志经 `RedactFilter` 脱敏（`sk-*`、
  `Authorization`、`api_key=` 一律打码），`.env` 已入 `.gitignore`。
- **异常分类**：文件被 Excel 占用（`PermissionError`）、工作表不存在、模型鉴权失败/余额不足/超时、
  模型返回非 JSON，都有独立的中文提示，不会抛裸栈给用户。

## 9. 上下文规模与准确性的取舍（已实测）

早期版本把表格压成“表头 + 推断类型 + 2 行样例”，对 500 行表能压 120 倍，但付了正确性的代价：
`_detect_header_row` 的打分包含“表头必须铺满 max_column”的假设，一旦表里除主数据区外另有标注块，
真表头的得分上限就被列数比例压到阈值以下→ 返回 `None` → 数据起始行退化为 1、列名退化为
`(列A)` 占位符、样例行从标题行开始取，模型根本看不到真实数据。结果就是把表头文本当数据算。

现在的做法：**Python 只搬运，语义判断全部交给模型**。运行 `python scripts/measure_digest.py` 可现场复现：

```
文件              规模                    提示字符      发送行数      省略行数
测试数据.xlsx       6行 × 6列                263         6         0
large.xlsx      500行 × 8列             2867        60       440
```

| 优化点 | 做法 | 实测效果 |
|---|---|---|
| **无损 TSV 文本化** | 逐格原样输出，带行号/列字母坐标系，不推断表头与类型；单元格文本截断到 24 字符，列数上限 40 | 6×6 小表反而比旧摘要更省（263 vs 旧版 273，旧版把字符花在 `A=(列A)(number)` 这类无效元数据上）；代价是中等表的提示会变长 |
| **区域起点约束写进提示词** | 要求区域首行必须是第一个真正的数据行，并把认定的数据区写进 `assumptions` 供人工核对 | 模型不再把表头行包进区域（`=MIN(A2:A10)` → `=MIN(A3:A10)`）；COUNTA、ROWS、INDEX、MATCH、SUMPRODUCT 这类对文本敏感的函数不再多算一行 |
| **行数闸门而非抽样** | 超过 `DIGEST_MAX_ROWS`（200）才截断，且保留头 50 行 + 尾 10 行，中间在提示里显式标注“第 X-Y 行已省略” | 500 行表发 60 行、省 440 行，提示稳定在 ~2900 字符；模型知道自己没看全，不会把截断当完整表 |
| **重试不重发上下文** | 校验失败时只追加“上一个公式 + 错误列表”，不重发表格内容与系统提示词 | 一次修复的新增输入仅 **89 字符**（`test_repair_loop_fixes_invalid_formula` 断言修复消息不含表格单元格值）；表越大，这一项的收益越大 |
| **本地拦截无效调用** | 意图识别、公式校验、表结构概览、公式试算全部本地完成 | 校验 / 概览类请求 **0 次模型调用**（`describe`、`validate` 子命令无需密钥即可运行） |
| **可观测** | 每次调用记录 `prompt_tokens / completion_tokens / 耗时`，交互界面直接打印 | 真实 Token 数可当场核对 |

## 10. 可选功能：公式自动验证

`excel_formula/evaluator.py` 用 Python 独立实现了常用函数子集（SUM/AVERAGE/MAX/MIN/COUNT/COUNTIF/
SUMIF/SUMIFS/IF/IFERROR/ROUND/RANK/VLOOKUP/文本函数/四则运算与比较等），在**写入之前**把公式算一遍：

- 算得出结果 → 预览显示 `预期结果: 438`，用户可当场判断对不对；
- 算出 `#DIV/0!` 等错误值 → 提前暴露问题；
- 遇到不支持的函数或未计算的公式引用 → 显示 `未验证（原因）`，不阻断流程。

这是独立于 DeepSeek 的第二条计算路径，属于"生成结果的交叉校验"，而不是让模型自证。

## 11. 测试

```powershell
python -m pytest tests -q          # 72 个用例，全部离线（用 FakeClient / FakeSession 替代网络）
python -m pytest tests -v          # 查看每个用例
```

覆盖范围：正常公式、缺 `=`、括号不配对、未知函数、禁用函数、参数个数错误、循环引用、越界引用警告、
不存在的工作表、中文列名、易变函数、绝对引用与百分号、整表无损文本化与头尾截断、12 个公式求值对照、
除零、不支持函数、生成并写入、预览不落盘、重试修复、连续失败放弃、追问、覆盖提示、
目录越界、非法后缀、写入规模上限、日志脱敏、DeepSeek 鉴权失败/服务端错误重试/返回值解析。

`tests/test_logger.py` 单独钉住了一个真实踩过的坑：脱敏过滤器曾把所有日志参数一律 `str()` 化，
导致 `%d` / `%.2f` 占位符在格式化时抛 `TypeError`——logging 会把错误打到 stderr 并**丢弃这条记录**，
Token 用量因此从未写进日志文件。现在数字参数原样放过，并断言 stderr 不出现 `Logging error`。

手工验收用例（含现场演示脚本）见 [`tests/testcases.md`](tests/testcases.md)。

## 12. 已知问题与边界

- openpyxl 保存会丢失原文件中的图表、图片、数据透视表等对象；含这些内容的文件建议先 `--output` 另存验证。
- 公式写入后 openpyxl 不会计算结果，单元格缓存值为空，需在 Excel/WPS 中打开一次才显示数值；
  这也是"本地预期结果"存在的原因。
- 本地求值器只覆盖常用函数子集，日期运算、数组公式、通配符条件（`"张*"`）会返回"未验证"。
- 跨工作簿引用（`[book2]Sheet1!A1`）未支持，会被校验拦下。
- 意图识别是规则实现，极端口语化的表述可能落到"生成"分支；此时程序会追问而不是乱写。
- `describe` 对超宽表只列出前 40 列，其余以"另有 N 列未列出"提示。

## 13. 数据与合规

`测试数据.xlsx` 为课程虚构数据（5 个科目 × 5 名学生的分数），不含真实个人信息。
仓库内不存放任何真实密钥，只提供 `.env.example`。
