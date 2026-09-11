---
name: excel-formula
description: 把中文自然语言需求转成 Excel 公式并写入指定单元格。当用户提到"生成/写公式""算总分、平均分、排名、占比、条件计数""校验这个公式对不对""解释某个公式的含义""看看这个表有哪些列"，且涉及 .xlsx/.xlsm 文件时使用本 Skill。
---

# AI Excel 公式生成 Skill

## 能力说明

底层是本仓库的 Python CLI（`python -m excel_formula.cli`），链路为：

```
openpyxl 逐格读表 → 无损 TSV 文本化 → DeepSeek 判断表头并生成公式 JSON → 本地校验(语法/函数白名单/引用/循环引用)
→ 失败自动回传错误重试(≤2次) → 预览 → 用户确认 → openpyxl 写入 → 日志
```

**重要约束**

1. 写文件属于高风险操作：必须先执行不带 `--apply` 的预览，把预览内容原样展示给用户，得到用户明确同意后，才允许再执行带 `--apply --yes` 的命令。
2. 不要自己臆造公式再让用户确认，公式一律由 CLI 产出（CLI 内部已做校验）。
3. 只操作用户明确给出的 `.xlsx/.xlsm` 文件，且必须位于项目工作目录内。
4. 不要自己把整张表读进上下文；需要了解表结构时用 `describe`，它返回带行号/列字母坐标系的整表 TSV
   （超 200 行会自动头尾截断并标注省略位置）。

## 何时使用哪个子命令

| 用户意图 | 子命令 |
|---|---|
| "这个表有哪些列/几行/几张工作表" | `describe` |
| "算总分/平均分/排名/及格人数并写到某单元格" | `generate`（先预览，后 `--apply`） |
| "这个公式对不对/帮我检查" | `validate` |
| "解释一下这个公式/G2 里的公式什么意思" | `explain` |
| 用户已给出确定公式，只要求写入 | `write` |
| 一句话含糊、想让程序自己判断 | `nl` |

## 参数

所有子命令共用：`--json`（结构化输出，推荐智能体使用）、`--verbose`、`--workspace <目录>`。

| 子命令 | 必填 | 可选 |
|---|---|---|
| `describe` | `--file` | `--sheet` |
| `generate` | `--file` `--request` | `--sheet` `--target` `--apply` `--yes` `--output` |
| `validate` | `--file` `--formula` | `--sheet` `--target` |
| `explain` | `--formula` 或（`--file` + `--cell`） | `--sheet` |
| `write` | `--file` `--target` `--formula` | `--sheet` `--fill-to` `--yes` `--output` |
| `nl` | 位置参数：整句指令 | `--file` `--sheet` `--apply` `--yes` |

退出码：`0` 成功；`2` 校验未通过或需要向用户追问；`3` 参数/路径/安全错误；`4` 模型或网络错误。

## 结果格式（`--json`）

`generate` 的关键字段：

```json
{
  "ok": true,
  "sheet": "Sheet1",
  "target": "G2",
  "formula": "=SUM(B2:F2)",
  "explanation": "对B2到F2求和得到总分",
  "cells": [{"sheet": "Sheet1", "cell": "G2", "formula": "=SUM(B2:F2)", "overwrites": null}],
  "validation": {"ok": true, "errors": [], "warnings": [], "functions": ["SUM"], "refs": ["B2:F2"]},
  "predicted_value": "438",
  "clarification": null,
  "attempts": [{"round": 1, "formula": "=SUM(B2:F2)", "ok": true, "errors": []}],
  "usage": {"calls": 1, "prompt_tokens": 412, "completion_tokens": 63, "total_tokens": 475},
  "applied": null
}
```

- `clarification` 非空：不要写入，把这个问题转述给用户并等待补充。
- `ok=false`：把 `error` 或 `validation.errors` 用自然语言解释给用户，不要重试写入。
- `predicted_value`：程序在本地独立算出的预期结果，可在确认时展示；为 `null` 时看 `predicted_note`。
- `applied`：只有真正写入后才非空，包含 `file`、`backup`、`count`。

## 示例

### 示例 1：生成并写入（正常两步流程）

用户："帮我在 测试数据.xlsx 里算每个科目的总分，写到 G2 并填充到 G6"

第一步，预览（不改文件）：

```bash
python -m excel_formula.cli generate --file 测试数据.xlsx --request "算每个科目的总分，G2到G6" --target G2 --json
```

把 `formula`、`cells`、`predicted_value`、`validation.warnings` 展示给用户并询问是否写入。

第二步，用户回答"可以/确认"后：

```bash
python -m excel_formula.cli generate --file 测试数据.xlsx --request "算每个科目的总分，G2到G6" --target G2 --apply --yes --json
```

回复用户：已写入 G2:G6 共 5 个单元格，原文件已备份到 backups/。

### 示例 2：校验用户自己写的公式

用户："=SUMIF(B2:B6,">85") 这个公式有问题吗？"

```bash
python -m excel_formula.cli validate --file 测试数据.xlsx --formula "=SUMIF(B2:B6,\">85\")" --json
```

`validation.ok=false` 时按 `errors` 说明原因（例如 SUMIF 缺少求和区域会被参数个数检查拦下）；
`ok=true` 时告诉用户校验通过，并给出 `predicted_value` 作为参考结果。

### 示例 3：解释单元格里已有的公式

用户："G2 里的公式是什么意思？"

```bash
python -m excel_formula.cli explain --file 测试数据.xlsx --cell G2 --json
```

直接把 `explanation` 转述给用户。

## 失败处理

| 情况 | 处理方式 |
|---|---|
| 退出码 3 且提示"路径超出允许目录" | 告知用户只能操作项目目录内的表格，请他移动文件或改用相对路径 |
| 退出码 3 且提示"文件可能正被 Excel 打开" | 请用户关闭 Excel 后重试 |
| 退出码 4 且提示鉴权失败 | 提示用户检查 `.env` 里的 `DEEPSEEK_API_KEY`，不要在回复里回显密钥 |
| `clarification` 有值 | 向用户追问缺失信息（通常是目标单元格或统计哪一列），补充后重跑 |
| 连续 3 次校验失败（`attempts` 长度为 3） | 把最后一次的 `errors` 讲清楚，建议用户换一种说法或指明数据区域 |
