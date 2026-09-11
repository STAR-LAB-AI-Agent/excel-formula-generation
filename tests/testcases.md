# 手工验收测试用例

数据文件：`测试数据.xlsx`（Sheet1，A1:F6，表头为 Subject / Student1..Student5，5 个科目 5 名学生分数）

| 编号 | 类型 | 输入（自然语言 / 命令） | 预期结果 |
|---|---|---|---|
| T01 | 正常 | `帮我在G2算每个科目的总分，填充到G6` | 生成 `=SUM(B2:F2)`，预期结果 438，确认后写入 G2:G6 且 G6 为 `=SUM(B6:F6)`，生成备份 |
| T02 | 正常 | `在H2算每个科目的平均分` | 生成 `=AVERAGE(B2:F2)`，预期结果 87.6 |
| T03 | 正常 | `统计Math这一行里超过85分的人数，放到I2` | 生成 `=COUNTIF(B2:F2,">85")`，预期结果 3 |
| T04 | 正常 | `这个表有哪些列？` | 列出 3 张工作表与 Sheet1 的 6 列表头，**不调用模型**（Token 用量为 0） |
| T05 | 正常 | `解释一下G2里的公式` | 读出 G2 的公式并给出中文讲解（需先完成 T01） |
| T06 | 正常 | `=SUMIF(A2:A6,"Math",B2:B6) 这个公式对不对？` | 校验通过，本地计算结果 85，**不调用模型** |
| T07 | 异常 | `=SUM(B2:F2 对不对？` | 校验未通过：`语法错误: 此处应为 右括号 )...`，退出码 2 |
| T08 | 异常 | `用INDIRECT("B2")取值写到G2` | 拒绝：禁止使用高风险函数 INDIRECT；重试 2 次后放弃，文件未改动 |
| T09 | 边界 | `在G2算B2到G2的和` | 拒绝：循环引用（公式写入 G2 却引用 G2） |
| T10 | 边界 | `=SUM(B20:F20)` 校验 | 校验通过但给出提示：引用完全落在数据区 A1:F6 之外 |
| T11 | 边界 | `帮我统计一下`（信息不足） | 触发追问（`clarification`），不写文件 |
| T12 | 安全 | `python -m excel_formula.cli describe --file C:\Windows\a.xlsx` | 拒绝：路径超出允许目录，退出码 3 |
| T13 | 安全 | `python -m excel_formula.cli describe --file requirements.txt` | 拒绝：只允许 .xlsx/.xlsm，退出码 3 |
| T14 | 安全 | 预览阶段回答 `n` | 输出"已取消，文件未改动"，文件字节数不变 |
| T15 | 异常 | 用 Excel 打开 测试数据.xlsx 后执行写入 | 提示"文件可能正被 Excel 打开，请关闭后重试" |
| T16 | 异常 | 清空 `DEEPSEEK_API_KEY` 后执行生成 | 提示配置缺失并指向 .env.example；校验类功能仍可用 |

## 现场演示脚本（约 1 分钟）

```powershell
python main.py
你 > :info                                  # T04：零 Token 看结构
你 > 帮我在G2算每个科目的总分，填充到G6        # T01：预览 → y → 写入
你 > =SUM(B2:G2) 这个公式对不对？             # T09：循环引用被拦下
你 > 解释一下G2里的公式                       # T05：解释
你 > :quit
```

## 自动化测试与手工用例的对应关系

| 手工用例 | 自动化测试 |
|---|---|
| T01 | `tests/test_pipeline.py::test_generate_and_apply` |
| T03 / T06 | `tests/test_evaluator.py::test_evaluate_common_formulas` |
| T04 | `tests/test_evaluator.py::test_digest_detects_headers_and_types` |
| T07 | `tests/test_validator.py::test_unbalanced_parenthesis` |
| T08 | `tests/test_validator.py::test_forbidden_function_rejected`、`test_gives_up_after_max_rounds` |
| T09 | `tests/test_validator.py::test_circular_reference_detected` |
| T10 | `tests/test_validator.py::test_out_of_data_range_warns` |
| T11 | `tests/test_pipeline.py::test_clarification_is_returned`、`test_missing_target_triggers_question` |
| T12 / T13 | `tests/test_pipeline.py::test_path_outside_workspace_rejected`、`test_non_excel_suffix_rejected` |
| T14 | `tests/test_pipeline.py::test_preview_does_not_touch_file` |
