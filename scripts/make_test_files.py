"""生成 ExcelCR 自测用的示例 Excel 文件（与 main.py 同目录）。

运行：python scripts/make_test_files.py
生成 4 个文件（同名的已存在文件会被覆盖）：
  成绩表.xlsx    成绩单 + 班级信息：SUM/AVERAGE/COUNTIF/AVERAGEIF/嵌套 IF/跨表 VLOOKUP
  销售数据.xlsx  原始数据 + 统计报表：乘法列、跨表 SUMIFS/COUNTIF、含日期列区域的 VLOOKUP
  库存表.xlsx    库存台账：IF 判断、COUNTIF/SUMIF、INDEX+MATCH
  边界情况.xlsx  标题占位 + 空行 + 文本混排：考验表头识别与区域起点约束

约定：所有"要写公式"的目标单元格一律留空，数据均为虚构。
"""
from __future__ import annotations

import datetime as dt
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.worksheet import Worksheet

ROOT = Path(__file__).resolve().parents[1]
HEADER_FILL = PatternFill("solid", fgColor="D9E1F2")
BOLD = Font(bold=True)
TITLE = Font(bold=True, size=13)
CENTER = Alignment(horizontal="center")


# ------------------------------------------------------------------ 通用工具
def _header(ws: Worksheet, row: int, ncols: int) -> None:
    for col in range(1, ncols + 1):
        cell = ws.cell(row=row, column=col)
        cell.font = BOLD
        cell.fill = HEADER_FILL
        cell.alignment = CENTER


def _title(ws: Worksheet, ref: str, text: str, *, merge: str | None = None) -> None:
    if merge:
        ws.merge_cells(merge)
    ws[ref] = text
    ws[ref].font = TITLE


def _widths(ws: Worksheet, widths: dict[str, int]) -> None:
    for col, width in widths.items():
        ws.column_dimensions[col].width = width


def _write_rows(ws: Worksheet, start_row: int, rows) -> None:
    for offset, row in enumerate(rows):
        for col, value in enumerate(row, start=1):
            if value is not None:
                ws.cell(row=start_row + offset, column=col, value=value)


# ------------------------------------------------------------------ 1. 成绩表
def make_scores(path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "成绩单"
    ws.append(["学号", "姓名", "班级", "语文", "数学", "英语", "总分", "平均分", "评级"])
    _write_rows(ws, 2, [
        ("S01", "张伟", "高一1班", 88, 92, 85),
        ("S02", "王芳", "高一1班", 76, 68, 90),
        ("S03", "李娜", "高一1班", 95, 88, 92),
        ("S04", "刘强", "高一1班", 52, 61, 47),
        ("S05", "陈晨", "高一2班", 79, 83, 74),
        ("S06", "杨帆", "高一2班", 91, 78, 84),
        ("S07", "赵敏", "高一2班", 68, 72, 77),
        ("S08", "周涛", "高一2班", 85, 93, 81),
        ("S09", "吴悠", "高一3班", 73, 66, 69),
        ("S10", "郑浩", "高一3班", 90, 87, 89),
        ("S11", "孙悦", "高一3班", 64, 59, 72),
        ("S12", "何雪", "高一3班", 82, 90, 76),
    ])  # 数据区 A1:I13，G/H/I 列为留空的写入目标
    _header(ws, 1, 9)
    ws["A15"] = "统计区"
    ws["A15"].font = BOLD
    ws["A16"] = "数学85分及以上人数"
    ws["A17"] = "高一1班语文平均分"
    ws["A18"] = "英语最高分"
    _widths(ws, {"A": 8, "B": 10, "C": 11, "D": 18, "E": 18, "F": 18, "G": 8, "H": 11, "I": 10})
    ws.freeze_panes = "A2"

    ws2 = wb.create_sheet("班级信息")
    ws2.append(["班级", "班主任", "教室"])
    _write_rows(ws2, 2, [
        ("高一1班", "李红", 301),
        ("高一2班", "王强", 302),
        ("高一3班", "孙丽", 303),
    ])
    _header(ws2, 1, 3)
    _widths(ws2, {"A": 11, "B": 10, "C": 8})
    ws2.freeze_panes = "A2"
    wb.save(path)


# ------------------------------------------------------------------ 2. 销售数据
def make_sales(path: Path) -> None:
    orders = [
        ("SO-2601", dt.date(2026, 7, 1), "杭州晨光商贸", "智能音箱", "华东", 399, 12),
        ("SO-2602", dt.date(2026, 7, 2), "广州南粤电子", "无线键盘", "华南", 129, 30),
        ("SO-2603", dt.date(2026, 7, 3), "北京京华数码", "显示器27寸", "华北", 1099, 8),
        ("SO-2604", dt.date(2026, 7, 5), "成都蜀通科技", "移动电源", "西南", 89, 50),
        ("SO-2605", dt.date(2026, 7, 7), "西安古都智联", "智能音箱", "西北", 399, 15),
        ("SO-2606", dt.date(2026, 7, 9), "上海浦江办公", "无线鼠标", "华东", 59, 80),
        ("SO-2607", dt.date(2026, 7, 11), "深圳鹏城数码", "显示器27寸", "华南", 1099, 10),
        ("SO-2608", dt.date(2026, 7, 12), "天津海河商贸", "移动电源", "华北", 89, 60),
        ("SO-2609", dt.date(2026, 7, 15), "重庆山城电子", "无线键盘", "西南", 129, 25),
        ("SO-2610", dt.date(2026, 7, 18), "兰州丝路科技", "智能音箱", "西北", 399, 18),
        ("SO-2611", dt.date(2026, 7, 20), "南京金陵办公", "显示器27寸", "华东", 1099, 6),
        ("SO-2612", dt.date(2026, 7, 23), "厦门鹭岛商贸", "无线鼠标", "华南", 59, 100),
        ("SO-2613", dt.date(2026, 7, 25), "石家庄燕赵数码", "无线键盘", "华北", 129, 40),
        ("SO-2614", dt.date(2026, 7, 28), "昆明春城智联", "移动电源", "西南", 89, 45),
        ("SO-2615", dt.date(2026, 8, 2), "乌鲁木齐天山科技", "显示器27寸", "西北", 1099, 5),
        ("SO-2616", dt.date(2026, 8, 5), "苏州姑苏电子", "智能音箱", "华东", 399, 20),
        ("SO-2617", dt.date(2026, 8, 8), "珠海香洲数码", "无线键盘", "华南", 129, 35),
        ("SO-2618", dt.date(2026, 8, 10), "太原并州商贸", "移动电源", "华北", 89, 70),
        ("SO-2619", dt.date(2026, 8, 12), "贵阳黔灵科技", "无线鼠标", "西南", 59, 90),
        ("SO-2620", dt.date(2026, 8, 15), "银川塞上电子", "智能音箱", "西北", 399, 22),
    ]
    wb = Workbook()
    ws = wb.active
    ws.title = "原始数据"
    ws.append(["订单编号", "下单日期", "客户名称", "产品", "区域", "单价", "数量", "金额"])
    _write_rows(ws, 2, orders)  # 数据区 A1:H21，H 列（金额）留空待写公式
    _header(ws, 1, 8)
    for row in range(2, 22):
        ws.cell(row=row, column=2).number_format = "yyyy-mm-dd"
        ws.cell(row=row, column=6).number_format = "#,##0"
    _widths(ws, {"A": 10, "B": 12, "C": 18, "D": 12, "E": 8, "F": 8, "G": 8, "H": 10})
    ws.freeze_panes = "A2"

    rpt = wb.create_sheet("统计报表")
    _title(rpt, "A1", "销售统计报表", merge="A1:C1")
    rpt["A3"], rpt["B3"], rpt["C3"] = "区域", "销售额", "订单数"
    _header(rpt, 3, 3)
    for offset, region in enumerate(["华东", "华南", "华北", "西南", "西北"]):
        rpt.cell(row=4 + offset, column=1, value=region)  # B4:C8 留空待写公式
    rpt["A10"] = "总销售额"
    rpt["A11"] = "平均订单金额"
    rpt["A12"] = "最大单笔金额"
    rpt["A14"] = "订单查询"
    rpt["A14"].font = BOLD
    rpt["A15"], rpt["B15"] = "订单编号", "SO-2607"  # B15 为查询输入单元格
    rpt["A16"], rpt["A17"], rpt["A18"] = "客户名称", "产品", "金额"  # B16:B18 留空待写公式
    _widths(rpt, {"A": 14, "B": 16, "C": 10})
    wb.save(path)


# ------------------------------------------------------------------ 3. 库存表
def make_stock(path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "库存台账"
    ws.append(["商品编码", "商品名称", "类别", "库存数量", "安全库存", "单价", "库存金额", "状态"])
    _write_rows(ws, 2, [
        ("P001", "A4复印纸", "办公用品", 120, 50, 25.5),
        ("P002", "中性笔(黑)", "办公用品", 300, 100, 2.5),
        ("P003", "订书机", "办公用品", 40, 30, 18),
        ("P004", "文件柜", "办公用品", 12, 15, 480),
        ("P005", "笔记本支架", "电子设备", 35, 20, 89),
        ("P006", "无线鼠标", "电子设备", 18, 25, 59),
        ("P007", "机械键盘", "电子设备", 22, 15, 329),
        ("P008", "显示器27寸", "电子设备", 8, 10, 1099),
        ("P009", "USB扩展坞", "电子设备", 45, 20, 129),
        ("P010", "墨盒(黑)", "耗材", 60, 40, 159),
        ("P011", "硒鼓", "耗材", 25, 30, 349),
        ("P012", "打印纸A3", "耗材", 80, 40, 42),
        ("P013", "色带", "耗材", 15, 25, 68),
        ("P014", "标签纸", "耗材", 200, 80, 12.5),
        ("P015", "封箱胶带", "耗材", 90, 50, 6.8),
    ])  # 数据区 A1:H16，G/H 列留空待写公式
    _header(ws, 1, 8)
    for row in range(2, 17):
        ws.cell(row=row, column=6).number_format = "0.00"
    ws["A18"] = "统计区"
    ws["A18"].font = BOLD
    ws["A19"] = "需补货商品数"
    ws["A20"] = "电子设备库存总金额"
    ws["A21"] = "单价最高的商品名称"
    _widths(ws, {"A": 12, "B": 14, "C": 11, "D": 10, "E": 10, "F": 9, "G": 11, "H": 9})
    ws.freeze_panes = "A2"
    wb.save(path)


# ------------------------------------------------------------------ 4. 边界情况
def make_edge_cases(path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "季度数据"
    _title(ws, "A1", "红星事业部 2026 年第三季度经营数据", merge="A1:E1")
    ws.merge_cells("A2:E2")
    ws["A2"] = "说明：金额单位为万元；「暂无」表示尚未发生"
    _write_rows(ws, 4, [("项目", "7月", "8月", "9月", "备注")])  # 第 3 行留空
    _header(ws, 4, 5)
    _write_rows(ws, 5, [
        ("销售收入", 128.5, 142.0, 156.8, "含税"),
        ("采购成本", 76.2, 80.5, 88.1, None),
        ("人工成本", 32.0, 32.0, 33.5, None),
        ("差旅费", 5.6, "暂无", 4.8, None),
        ("办公费", 3.2, 3.5, 3.1, None),
        ("市场推广", 12.8, None, 15.2, None),
        ("设备折旧", 6.0, 6.0, 6.0, None),
        ("其他支出", 2.4, 1.8, 2.2, None),
    ])  # 数据区 A4:E12，其中 B8 为文本、C10 为空
    ws["A14"] = "合计"  # B14:D14 留空待写公式
    ws["F4"], ws["G4"] = "指标", "数值"
    for ref in ("F4", "G4"):
        ws[ref].font = BOLD
        ws[ref].fill = HEADER_FILL
        ws[ref].alignment = CENTER
    ws["F5"], ws["F6"], ws["F7"] = "收入合计", "成本合计", "毛利率"  # G5:G7 留空待写公式
    _widths(ws, {"A": 11, "B": 8, "C": 8, "D": 8, "E": 8, "F": 11, "G": 9})
    ws.freeze_panes = "A5"
    wb.save(path)


# ------------------------------------------------------------------ 入口
FILES = {
    "成绩表.xlsx": make_scores,
    "销售数据.xlsx": make_sales,
    "库存表.xlsx": make_stock,
    "边界情况.xlsx": make_edge_cases,
}


def main() -> int:
    for name, builder in FILES.items():
        target = ROOT / name
        if target.exists():
            print(f"覆盖已存在文件：{name}")
        builder(target)
        wb = load_workbook(target)
        dims = " | ".join(
            f"{ws.title}: {ws.max_row}行×{ws.max_column}列" for ws in wb.worksheets
        )
        wb.close()
        print(f"√ {name}  {dims}")
    print("完成，目标单元格（要写公式的位置）均为空。")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
