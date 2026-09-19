# AI Excel 公式生成（ExcelCR）

《智能体开发实战》课程实验 · 选题 23「AI Excel 公式生成」

用一句中文说出你要算什么，程序自动读表结构、调用 DeepSeek 生成 Excel 公式、**在本地严格校验**、
预览并经你确认后用 openpyxl 写回文件，同时给出公式解释、预期结果和可追溯日志。

## 1. 技术路线

```
用户自然语言 + Excel 文件
        ↓
openpyxl 逐格读出目标表及同簿来源表（只搬运，不推断表头与列类型）
        ↓
无损 TSV 文本化（带行号/列字母坐标系，超 200 行才头尾截断）
        ↓
DeepSeek API：自行判断表头行与数据区 → 输出公式/常量值 JSON
        ↓
本地校验：语法 AST + 函数白名单 + 引用范围 + 循环引用 + 参数个数 + 逐格试算
        ↓
校验失败？→ 只回传公式与错误信息给 DeepSeek，最多重试 2 次
        ↓
预览（公式或常量值 / 影响单元格 / 预期结果 / 是否覆盖旧内容）+ 用户确认
        ↓
openpyxl 写入并保存（自动备份原文件；新单元格沿用同行相邻单元格的格式）
        ↓
返回结果 + Token 用量 + logs/excelcr.log
```

## 2. 支持的四类自然语言意图（外加一个辅助意图）

| 意图 | 用户说法示例 | 走到哪一步 |
|---|---|---|
| ① 生成公式并写入 | "帮我在 G2 算每个科目的总分，填充到 G6" | 全链路，含确认与写入 |
| ② 校验公式 | "=SUMIF(B2:B6,\">85\") 这个公式对不对？" | 纯本地校验 + 本地试算，**零 Token** |
| ③ 解释公式 | "解释一下 G2 里的公式" | 读单元格公式 → DeepSeek 讲解 |
| ④ 表格加框 | "将 A14 到 B18 的表格范围框起来" | 纯本地套细边框，**零 Token**（模型只会给公式、给不了格式） |
| ⑤ 表结构概览（辅助） | "这个表有哪些列？" | 纯本地读取，**零 Token** |

意图识别由 `excel_formula/intent.py` 用规则完成，不花 Token；表名与文件里的真实工作表对不上时，会回到原话找回真实表名（只有唯一确定才替换）；只有真正需要“理解语义/生成公式”时才调用模型。加框类需求同样在本地闭环：边框词 + 范围抽取（“A14 到 B18”/“A14:B18”）命中即直接出加框预览，逐格套细边框且不改单元格内容；“新建表格并加边框，统计……”这类带生成动作的说法仍走模型，边框由新建表格流程顺带套用，“去掉/取消边框”不会被误当成加框执行。

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

# 4) 交互模式（同时自动打开 Web 演示台；--no-web 只保留 CLI）
python main.py

# 或单条指令模式（默认安静；加 --web 可同时打开演示台）
python main.py "帮我在G2算每个科目的总分，填充到G6"
```

启动后自动显示当前工作目录中的 `.xlsx/.xlsm` 编号列表，直接输入数字即可选择文件（不消耗 Token，不写文件）：

```text
  1. 库存表.xlsx
  2. 成绩表.xlsx
  3. 销售数据.xlsx
你 > 2
√ 当前文件：成绩表.xlsx
你 > 在班级信息表教室列后面增加一列班级学生平均成绩，成绩来自表一成绩单
```

编号以实际显示为准，按文件名排序；只有一个文件时自动选中，多文件时不默认选第一项。
列表仅包含当前目录直属文件，忽略 `~$` 临时锁文件及子目录；文件增删后用 `:files` 刷新，刷新前编号保持不变。

交互模式内置指令：`:files` 刷新列表、`:file <编号或路径>` 切换文件（不带参数则显示列表）、
`:sheet <表名>` 指定工作表、`:think <1|0|auto>` 思考开关（立即生效）、
`:info` 查看表结构（零 Token）、`:web` 启动（或查看）Web 演示台地址、`:help`、`:quit`。选择文件仅切换程序的操作对象，不启动 Excel/WPS。

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
  已为新单元格套用相邻格式：G2 ← F2、G3 ← F3、G4 ← F4、G5 ← F5 等共 5 个
  原文件已备份：测试数据_20260910_142530.xlsx
  日志：logs/excelcr.log
```

### 响应速度与流式输出

- **流式输出**：生成公式与解释公式时走 SSE 流式接口，终端上以一行原地刷新的进度显示
  `… 模型思考 1234 字 / 正文 56 字，已用 12s`——模型是在思考还是卡住了，一目了然；
  流式请求会自动携带 `stream_options.include_usage`，Token 用量照常从最后一个分片回收。
  非交互终端（重定向/管道）自动退回一次性返回，程序化调用零变化。
- **思考开关（1/0）**：DeepSeek V4.x 默认开启思考且强度为 `high`，思考内容同样计入 `max_tokens`
  与计费（这是“明明只是生成一条公式却等了 40 秒”的主因）。`.env` 里用一个数字开关控制：

  ```
  EXCELCR_THINKING=0             # 最快：始终关闭思考，直接出公式
  EXCELCR_THINKING=1             # 深度思考常开：每次都思考，最准也最慢
  EXCELCR_THINKING=auto          # 自动：按需求复杂度切换（规则见下）
  EXCELCR_REASONING_EFFORT=low   # 可叠加：保留思考时降低推理强度
  ```

  `on/off` 与 `enabled/disabled` 是等价别名。交互模式里可随时切换、立即生效，不用改配置重启：

  ```
  :think 1      # 打开深度思考
  :think 0      # 关闭，追求最快
  :think auto   # 恢复按复杂度自动
  :think        # 查看当前档位
  ```

  流式进度行在两种档位下都会显示，区别只在看到的是「模型思考 N 字」增长（开思考时）
  还是「正文 N 字」直接跳动（关闭时）。

  `auto` 档的判定规则：需求超过 24 字、表格超过 40 行/12 列、或命中跨表/多条件/查找匹配等
  关键词时保留思考；**首轮校验失败后的修复轮次一律升级为思考兜底**；其余情况关闭思考
  （`explain` 按公式长度 60 字符分界）。留空则完全保持服务端默认行为。

### Web 演示台（答辩演示用，零依赖）

```powershell
python webapp.py                             # 默认 127.0.0.1:8765，自动打开浏览器；端口被占用自动顺延
python webapp.py --port 9000 --no-browser    # 可选参数
```

也可以由交互入口自动拉起：`python main.py` 启动时会在后台运行演示台并打开浏览器
（`--no-web` 关闭；`--web` 也可用于单条指令模式）；两者共享同一个 `FormulaService`，
在网页里切换思考档位，CLI 立即生效——`:web` 随时可再次打开或查看地址。

三栏答辩大屏（会话面板 / Excel 表格 / 对话，窄屏自动收起会话面板），复用与 CLI 完全相同的 `FormulaService`
流水线，不引入任何第三方依赖、不改动包内代码：

- **思考链打字机**：`propose/explain` 的 `on_delta` 增量经 SSE 实时推送，思考与正文分轨展示，
  状态行实时刷新「思考 N 字 / 正文 M 字 / 已用 Ts」——模型在工作还是卡住，一目了然；
- **一句话全流程**：意图识别（0 Token）→ 生成（流式）→ 本地校验含试算（0 Token）→
  公式预览卡片（预期结果 / 覆盖警告 / 修复轮次）→ 点击「确认写入」才落盘；
- **表格联动**：网格在数据区之外补出空白行列并自动铺满中间栏（至少 20 行 / 12 列，随后按可视区
  实测尺寸继续补足，无数据处全是空白格子；窗口尺寸变化时自动重铺），像打开一张真实工作表，
  便于一眼看清可写范围；公式格直接显示计算结果（Excel 缓存值优先，缺失时本地试算、0 Token，
  链式依赖沿引用链逐层递归求值），鼠标悬停浮出原始公式；预览时目标单元格虚线高亮并自动滚入视野（数据区外同样有效）、写入成功后变绿；
  结果卡片同步显示写入数、备份名与格式沿用；
- **网格缩放**：工具栏 `− 100% +`（也可 Ctrl+滚轮）在 50%–200% 之间缩放表格（步进 10%），
  列宽 / 行高 / 字号等比缩放并自动重新铺满；档位记入浏览器本地存储，刷新后保持；
- **打开原文件 / 提交表格**：一键用本机 Excel/WPS 打开当前文件对照查看结果（`/api/open`，仅限工作目录内文件）；
  「提交表格」可直接上传自己电脑上的 .xlsx/.xlsm（`/api/upload`：校验文件有效性、重名自动改名、
  后缀与体积受限、文件名不可含目录），上传后自动加入列表并选中加载；
- **请求细节**：`/api/chat` 以 SSE（chunked）返回 `stage / delta / result / error` 事件；
  待确认方案由服务端暂存并发放一次性令牌，`/api/apply` 凭令牌写入、二次使用被拒绝；
  文件路径同样经目录白名单收敛；顶栏可随时切换思考档位（1 / 0 / auto，立即生效）。

`tests/test_webapp.py`（23 个用例）真实启动本机服务后验证：五类意图分发（校验/概览/加框断言零模型调用）、
SSE 事件序列、生成→确认→写入闭环（含备份与令牌一次性）、越界路径拒绝、思考档位即时切换、后台自动拉起与端口顺延、
公式格结果展示（含链式依赖递归求值）与悬停公式元数据、打开原文件与上传表格（重名自动改名、无效内容拒绝）、
加框预览→确认后逐格套细边框且零写入。

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
| `excel_formula/intent.py` | 规则意图识别与参数抽取（文件/工作表/目标单元格/填充范围/公式；错抓的表名会回原文校验修正） |
| `excel_formula/excel_reader.py` | openpyxl 读取、**无损 TSV 文本化**（不推断表头与类型，只做行数闸门） |
| `excel_formula/llm_client.py` | DeepSeek 调用（含 SSE 流式与 Token 统计）、提示词、JSON 解析 |
| `excel_formula/formula_parser.py` | 公式分词器 + 递归下降语法分析器（产出 AST） |
| `excel_formula/validator.py` | 静态校验：语法、函数白名单、参数个数、引用范围、循环引用 |
| `excel_formula/evaluator.py` | 公式独立求值器（Python 重算一遍，用于"公式自动验证"） |
| `excel_formula/writer.py` | 备份、相对引用填充展开、写入保存（新单元格沿用同行相邻格式） |
| `excel_formula/pipeline.py` | 流水线编排：生成 → 校验 → 重试 → 预览 → 写入（支持一次多个目标单元格） |
| `excel_formula/config.py` | 配置、目录白名单、函数白/黑名单、写入上限 |
| `excel_formula/logger.py` | 日志与密钥脱敏 |
| `webapp.py` | 零依赖 Web 演示台（标准库 http.server + SSE 流式），答辩用三栏大屏 |
| `web/index.html` | 前端单文件（原生 JS/CSS）：对话流、表格高亮、思考链打字机 |
| `scripts/measure_digest.py` | 提示词规模检查脚本（文本化后有多大、是否触发截断） |
| `skills/excel-formula/SKILL.md` | 供智能体调用的 Skill 说明 |
| `tests/` | 498 个自动化测试用例（不需要联网，含大作业实景复刻与真实文件对账）+ 手工验收用例清单 |

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
- **生成结果试算拦截**：逐格检查生成的整批公式及填充区域，明确的 Excel 错误值进入有限重试；修复失败则整批拒绝，不进入写入确认。重叠写入也会拒绝。
- **写入前确认**：预览 → 用户输入 `y` 才写；非交互环境必须显式 `--yes`，否则拒绝执行。
- **写入前备份**：原文件复制到 `backups/<名称>_<时间戳>.xlsx`，可人工回退。
- **写入规模上限**：单次最多 200 个单元格（`MAX_WRITE_CELLS`），避免误伤整表。
- **覆盖提示**：目标单元格已有值或公式时，预览里用 `⚠ 将覆盖已有内容` 明确列出。
- **格式沿用**：写入的新单元格若本身没有格式，会沿用同行相邻单元格的格式
  （新表头继承表头颜色/加粗，新数据格继承边框等；左邻优先，无则往右找）；
  已有自定义格式的格子保持原样，表内其他单元格的格式不受影响。
- **密钥保护**：密钥只从环境变量/`.env` 读取，只出现在请求头；日志经 `RedactFilter` 脱敏（`sk-*`、
  `Authorization`、`api_key=` 一律打码），`.env` 已入 `.gitignore`。
- **异常分类**：文件被 Excel 占用（`PermissionError`）、工作表不存在、模型鉴权失败/余额不足/超时、
  模型返回非 JSON，都有独立的中文提示，不会抛裸栈给用户。
- **代理容错**：Windows 的系统代理在注册表里只存 `host:port`，Python 会把 https 通道拼成
  `https://127.0.0.1:端口`，对本机明文代理做 TLS 握手就会报
  `ProxyError('Unable to connect to proxy')`。客户端启动时把这类回环代理统一降级为 `http://`，
  真遇到代理连不通还会自动改直连重试一次；也可用 `.env` 里的 `EXCELCR_PROXY` 指定代理、
  `EXCELCR_TRUST_ENV=0` 强制直连。

## 9. 上下文规模与准确性的取舍（已实测）

早期版本把表格压成“表头 + 推断类型 + 2 行样例”，对 500 行表能压 120 倍，但付了正确性的代价：
`_detect_header_row` 的打分包含“表头必须铺满 max_column”的假设，一旦表里除主数据区外另有标注块，
真表头的得分上限就被列数比例压到阈值以下→ 返回 `None` → 数据起始行退化为 1、列名退化为
`(列A)` 占位符、样例行从标题行开始取，模型根本看不到真实数据。结果就是把表头文本当数据算。

现在的做法：**Python 只搬运，语义判断全部交给模型**。生成时同时提供目标表与同一工作簿其他表的 TSV，明确标注写入目标及引用来源，避免模型只看到表名就猜列。
来源表沿用每表的行列截断规则，另共享 40,000 字符预算，优先保留需求明确提到的表；超限的表名会明确列出，并要求信息不足时追问，禁止猜测列号。
以下规模数据针对单张表的 TSV。运行 `python scripts/measure_digest.py` 可现场复现：

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

`excel_formula/evaluator.py` 用 Python 独立实现了常用函数子集，在**写入之前**把公式算一遍：

- 数学：`SUM/PRODUCT/ABS/ROUND/ROUNDUP/ROUNDDOWN/INT/TRUNC/MOD/POWER/SQRT/SIGN/EXP/LN/LOG/LOG10/CEILING/FLOOR/SUMPRODUCT`
- 统计：`AVERAGE/AVERAGEIF/AVERAGEIFS/MEDIAN/MODE(.SNGL)/COUNT/COUNTA/COUNTBLANK/COUNTIF/COUNTIFS/MAX/MAXIFS/MIN/MINIFS/LARGE/SMALL/RANK(.EQ)/STDEV(.S/.P)/STDEVP/VAR(.S/.P)`
- 查找：`VLOOKUP/HLOOKUP/XLOOKUP/LOOKUP/INDEX/MATCH`
- 逻辑：`IF/IFS/IFERROR/IFNA/AND/OR/NOT/XOR/CHOOSE/SWITCH/TRUE/FALSE`
- 文本：`LEN/LEFT/RIGHT/MID/FIND/SEARCH/SUBSTITUTE/REPLACE/REPT/TRIM/CONCAT/CONCATENATE/TEXTJOIN/UPPER/LOWER/PROPER/CHAR/CODE/VALUE`
- 信息：`ISBLANK/ISNUMBER/ISTEXT/ISERROR/ISERR/ISNA/ISEVEN/ISODD/NA/ROWS/COLUMNS/ROW/COLUMN`
- 日期：`TODAY/NOW/DATE/YEAR/MONTH/DAY/HOUR/MINUTE/SECOND/WEEKDAY/WEEKNUM/EOMONTH/EDATE/DATEDIF/DAYS/DATEVALUE/NETWORKDAYS/WORKDAY`

行为约定：

- 算得出结果 → 预览显示 `预期结果: 438`，用户可当场判断对不对；
- 生成流程算出 `#DIV/0!`、`#VALUE!`、`#N/A` 等确定错误值 → 回传模型修复（含出错表名及单元格），持续失败则拒绝整批写入；普通文本 `"#DIV/0!"` 不会被误判；
- 遇到不支持的函数或未计算的公式引用 → 显示 `未验证（原因）`，不阻断流程
  （这是校验路径的保守语义；网页表格展示会沿引用链递归试算缺失缓存的公式，见第 5 节）；
- 生成流程试算前，在内存中暂存本批常量与公式，避免用尚未写入的旧分类名/分母计算；本批公式之间的依赖及主簿无缓存公式仍保守标未验证，试算结束后恢复内存，不写盘；
- `validate` 子命令仍分别报告静态校验与试算结果；直接 `write` 指令保持静态校验规则，不走模型修复；
- 日期按 Excel 1900 日期系统序列号语义参与运算：日期 ± 天数得日期、日期 − 日期得天数，
  比较与汇总按序列号折算，`&` 拼接输出 ISO 文本（`">="&DATE(2025,1,15)` 条件可被还原）；
  日期与无法解析的文本比较、通配符条件（`"张*"`）仍标"未验证"，而不是给出静默算错的结果；
- 区域直接参与一元/二元运算（`=SUM(B2:B6*C2:C6)`、`(D2:D9="华北")*J2:J9` 条件数组写法）
  按 Excel 数组语义逐元素求值：同形区域对应元素运算，行/列单维区域自动伸展（外积），
  维度不兼容返回 `#N/A`。这类公式写入时自动以数组公式（CSE）形态保存——裸公式会被
  Excel/WPS 按传统"隐式交叉"只取交叉单值；以 CSE 写入后，Excel/WPS 打开、网页显示
  与本地试算三方一致（实测 `=SUM(B2:B6*C2:C6)` 得 34740，隐式交叉只会得到 6630）；
- 跨工作簿引用（`[book2.xlsx]Sheet1!A1`，引号式与带路径写法均可）会在预览时按公式里的
  位置自动打开外部工作簿试算：优先读 Excel 缓存值，外部公式没有缓存时递归本地求值
  （带循环检测）；省略目录的引用按当前工作簿同目录查找；文件找不到或打不开时标
  "未验证（原因）"，不影响写入；
- 所有"未验证"路径都抛 `UnsupportedFormula`（异常消息带原因），
  对应上面第 8 节点的"宁可未验证也不静默算错"。

这是独立于 DeepSeek 的第二条计算路径，属于"生成结果的交叉校验"，而不是让模型自证。

## 11. 测试

```powershell
python -m pytest tests -q          # 638 个用例，全部离线（用 FakeClient / FakeSession 替代网络）
python -m pytest tests -v          # 查看每个用例
python -m pytest tests/test_homework_smoke.py -v   # 真实大作业文件对账（文件不在时自动跳过）
```

覆盖范围：正常公式、缺 `=`、括号不配对、未知函数、禁用函数、参数个数错误、循环引用、越界引用警告、
不存在的工作表、跨工作簿引用试算、中文列名、易变函数、日期序列号语义、绝对引用与百分号、
整表无损文本化与头尾截断、12 个公式求值对照、除零、不支持函数、区域数组运算与 CSE 写入、生成并写入、预览不落盘、
跨表来源上下文、试算错误自动修复与阻止写入、填充中间格错误、无缓存公式降级、
链式无缓存公式递归求值（循环引用检测与依赖结果复用）、
启动编号选文件、列表过滤与刷新、无效编号、失效文件、追问数字不误切文件、
重试修复、连续失败放弃、追问、覆盖提示、
目录越界、非法后缀、写入规模上限、日志脱敏、DeepSeek 鉴权失败/服务端错误重试/返回值解析、
SSE 流式分片重组与 usage 回收（含 `stream_options` 被拒后的自动降级与思考开关下发）、
auto 思考档的复杂度启发式（简单直算关闭、修复轮次升级）。

2026-09 扩展（以大作业 `Excel大作业.xlsm` 为素材）：

- `tests/test_formula_library.py`（163 个）：逐族覆盖全部已实现函数的返回值与错误值，
  并钉住与 Excel 的语义一致性（`MOD` 负数取模符号跟随除数、`ISERR` 对 `#N/A` 返回 FALSE、
  `TRUNC` 支持位数、`CEILING/FLOOR` 负数方向等）；
- `tests/test_homework_formulas.py`（31 个）：把课程作业里的公式清单原样搬进测试——
  三层嵌套 IF 评级四档全覆盖、统计区六件套、跨表 SUMIFS、`IFERROR+VLOOKUP`、
  日期列序列号汇总与区域数组语义求值；
- `tests/test_homework_smoke.py`：直接打开真实 `Excel大作业.xlsm`，把 85 个公式逐个
  本地重算并与 Excel 缓存值对账（83 个一致、2 个定义名称公式标"未验证"）；
- `tests/test_pipeline.py`：跨表与跨工作簿端到端用例，以及班级均分错列→除零→修复→临时文件写入回归；
- `tests/test_console.py`：启动编号菜单、数字选文件不触发模型、无效编号与列表快照、错误方案不询问写入；
- `tests/test_date_support.py`（52 个）：日期序列号语义逐项钉住——日期 ± 天数、日期 − 日期、
  比较与条件匹配（`">="&DATE(...)`、`">2025-01-15"`）、DATE/EDATE/EOMONTH/DATEDIF/WEEKDAY/
  WEEKNUM/NETWORKDAYS/WORKDAY 函数族与 Excel 构造怪癖（月份/日溢出进位、1899 年自动加 1900），
  日期与无法解析文本比较仍标"未验证"；
- `tests/test_external_refs.py`（19 个）：跨工作簿引用（未引号/引号/路径式）通过校验并给出
  去重警告；预览试算会真正打开外部工作簿读取缓存值，外部无缓存公式递归求值，循环引用、
  缺失文件/工作表与损坏文件一律降级为"未验证（原因）"；

`tests/test_logger.py` 单独钉住了一个真实踩过的坑：脱敏过滤器曾把所有日志参数一律 `str()` 化，
导致 `%d` / `%.2f` 占位符在格式化时抛 `TypeError`——logging 会把错误打到 stderr 并**丢弃这条记录**，
Token 用量因此从未写进日志文件。现在数字参数原样放过，并断言 stderr 不出现 `Logging error`。

手工验收用例（含现场演示脚本）见 [`tests/testcases.md`](tests/testcases.md)。

## 12. 已知问题与边界

- openpyxl 保存会丢失原文件中的图表、图片、数据透视表等对象；含这些内容的文件建议先 `--output` 另存验证。
- 公式写入后 openpyxl 不会计算结果，单元格缓存值为空（任何经本工具保存的文件，原有公式缓存也会被清空）；
  网页表格用本地试算兜底显示、链式依赖逐层递归求值，用 Excel/WPS 打开文件时会自动重算并写回缓存。
  这也是"本地预期结果"存在的原因。
- 本地求值器只覆盖常用函数子集：通配符条件（`"张*"`）仍返回"未验证"；区域数组运算
  已支持并按数组语义求值、写入时 CSE 化（见第 10 节）；日期已按序列号语义参与
  算术、比较、汇总与条件匹配。
- 跨工作簿引用（`[book2.xlsx]Sheet1!A1`）的本地试算：按公式里的目录（省略时取当前文件
  同目录）打开外部工作簿读 Excel 缓存值；外部公式没有缓存时递归本地求值；文件找不到、
  损坏、旧版 .xls 或超过 50 MB 时预览标"未验证（原因）"，写入不受影响。
- 意图识别是规则实现，极端口语化的表述可能落到“生成”分支；此时程序会追问而不是乱写。说“班级信息表教室列后面…”这类“表名+表字后缀”的话时，提取器可能抓出句子片段；程序会回到原话找回真实表名，找不回时才报“工作表不存在”。
- `describe` 对超宽表只列出前 40 列，其余以"另有 N 列未列出"提示。

## 13. 数据与合规

`测试数据.xlsx` 为课程虚构数据（5 个科目 × 5 名学生的分数），不含真实个人信息。
`Excel大作业.xlsm` 为个人课程作业文件（虚构销售订单），仅用于本地实景测试，
`tests/test_homework_smoke.py` 在文件不存在时自动跳过，仓库中不含该文件。
仓库内不存放任何真实密钥，只提供 `.env.example`。
