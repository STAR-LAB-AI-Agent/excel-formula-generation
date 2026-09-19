"""生成 5 分钟答辩用的简约风 PPT（与 main.py 同目录）。

运行：python scripts/make_defense_ppt.py
生成：答辩PPT.pptx —— 10 页，白底简约风，覆盖：背景 → 架构 → 六类意图 →
      三大亮点 → 现场演示 → 测试质量 → 总结展望。

风格约定：白底 + 深灰正文 + 单一强调色（深海蓝）；不加图片与装饰图形，
标题下一条细线，右下角页码；封面深蓝底白字。
"""
from __future__ import annotations

from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "答辩PPT.pptx"

FONT = "微软雅黑"
INK = RGBColor(0x22, 0x2A, 0x38)      # 主文字：深灰
SUB = RGBColor(0x5B, 0x66, 0x77)      # 次级文字
ACCENT = RGBColor(0x1D, 0x4E, 0x89)   # 强调色：深海蓝
LIGHT = RGBColor(0xEE, 0xF3, 0xF9)    # 浅蓝底（表格头）
LINE = RGBColor(0xD5, 0xDC, 0xE4)     # 细线
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
COVER_BG = RGBColor(0x16, 0x23, 0x3B)  # 封面：深蓝黑

SLIDE_W, SLIDE_H = 13.333, 7.5


def _ea(run) -> None:
    rPr = run._r.get_or_add_rPr()
    ea = rPr.find(qn("a:ea"))
    if ea is None:
        ea = etree.SubElement(rPr, qn("a:ea"))
    ea.set("typeface", FONT)


def _run(run, text, size, *, bold=False, color=INK, italic=False):
    run.text = text
    f = run.font
    f.name = FONT
    f.size = Pt(size)
    f.bold = bold
    f.italic = italic
    f.color.rgb = color
    _ea(run)
    return run


def _textbox(slide, x, y, w, h):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    return tf


def _para(tf, first: bool, *, space_before=0, space_after=6, line=1.25):
    p = tf.paragraphs[0] if first else tf.add_paragraph()
    p.space_before = Pt(space_before)
    p.space_after = Pt(space_after)
    p.line_spacing = line
    return p


def _rect(slide, x, y, w, h, color):
    from pptx.enum.shapes import MSO_SHAPE
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    shape.line.fill.background()
    shape.shadow.inherit = False
    return shape


def _blank(prs):
    return prs.slides.add_slide(prs.slide_layouts[6])


def _content_slide(prs, title: str, page: int):
    """内容页骨架：左上强调色小条 + 标题 + 细线 + 右下页码。"""
    slide = _blank(prs)
    _rect(slide, 0.9, 0.62, 0.09, 0.42, ACCENT)
    tf = _textbox(slide, 1.15, 0.52, 11.3, 0.7)
    p = _para(tf, True, space_after=0)
    _run(p.add_run(), title, 26, bold=True)
    _rect(slide, 0.9, 1.32, 11.53, 0.016, LINE)
    foot = _textbox(slide, 11.4, 6.95, 1.0, 0.35)
    p = _para(foot, True, space_after=0)
    p.alignment = PP_ALIGN.RIGHT
    _run(p.add_run(), str(page), 11, color=SUB)
    return slide


def _bullets(slide, items, *, x=1.15, y=1.75, w=11.2, h=5.0, size=17, gap=12):
    """items: (文本, 粗体附加说明或 None) 或 str；'-' 开头为二级缩进。"""
    tf = _textbox(slide, x, y, w, h)
    first = True
    for item in items:
        if isinstance(item, str):
            text, note = item, None
        else:
            text, note = item
        level2 = text.startswith("-")
        text = text.lstrip("-").strip()
        p = _para(tf, first, space_after=gap)
        first = False
        if level2:
            p.level = 1
        _run(p.add_run(), ("• " if not level2 else "– ") + text, size if not level2 else size - 2,
             bold=False, color=INK if not level2 else SUB)
        if note:
            _run(p.add_run(), "   " + note, size - 3, color=ACCENT)
    return tf


# ------------------------------------------------------------------ 各页
def slide_cover(prs):
    slide = _blank(prs)
    _rect(slide, 0, 0, SLIDE_W, SLIDE_H, COVER_BG)
    _rect(slide, 1.1, 2.28, 0.62, 0.07, RGBColor(0x5B, 0x8F, 0xD6))
    tf = _textbox(slide, 1.1, 2.62, 11.0, 1.4)
    p = _para(tf, True, space_after=0)
    _run(p.add_run(), "ExcelCR", 54, bold=True, color=WHITE)
    tf2 = _textbox(slide, 1.13, 3.95, 11.0, 1.0)
    p = _para(tf2, True, space_after=0)
    _run(p.add_run(), "自然语言驱动的 Excel 公式生成智能体", 24, color=RGBColor(0xC9, 0xD6, 0xE8))
    tf3 = _textbox(slide, 1.13, 5.35, 11.0, 1.0)
    p = _para(tf3, True, space_after=4)
    _run(p.add_run(), "《智能体开发实战》课程实验 · 答辩", 16, color=RGBColor(0x9F, 0xB3, 0xCC))
    p = _para(tf3, False, space_after=0)
    _run(p.add_run(), "汇报人：＿＿＿＿　　2026 年 9 月", 14, color=RGBColor(0x8A, 0x9E, 0xB8))


def slide_background(prs):
    slide = _content_slide(prs, "一、背景与目标", 2)
    _bullets(slide, [
        ("Excel 公式门槛高：", "函数多、语法细、跨表引用容易写错"),
        ("目标：", "一句话需求 → 生成公式 → 本地校验 → 预览确认 → 安全写入，人始终在场"),
        ("交付形态：", "命令行 CLI ＋ 标准库 Web 演示台 ＋ Skill 技能包，共享核心服务"),
        ("能力扩展：", "公式与常量写入、跨表统计、格式处理、下拉列表设置与网页选择"),
        "核心原则：明确错误拦截，未验证如实提示，预览确认后才执行写入",
    ], size=18, gap=16)


def slide_arch(prs):
    slide = _content_slide(prs, "二、系统架构：本地分流，统一写入", 3)
    _bullets(slide, [
        ("① 本地意图分流　", "规则识别六类需求；概览、校验、加框直接在本地完成"),
        ("② 读取表格上下文　", "目标表＋同簿来源表转 TSV，保留坐标，超限显式提示"),
        ("③ 按需调用模型　", "生成公式 / 常量 JSON；下拉规则未命中时由 DeepSeek 兜底"),
        ("④ 校验与预览确认　", "静态检查＋整批试算，展示影响范围、结果及覆盖警告"),
        ("⑤ 可靠写入　", "openpyxl 保存，原地写入默认备份，临时文件原子替换"),
    ], size=18, gap=14)
    slide.notes_slide.notes_text_frame.text = (
        "规则命中的下拉需求为 0 Token；本地无法理解时最多调用一次模型，输出仍经本地验证。\n"
        "TSV 不是无限量无损传输：默认超过 200 行保留头 50 行和尾 10 行，最多 40 列，"
        "单元格文本上限 24 字；同簿来源表另共享 40,000 字符预算，优先保留需求点名的表。\n"
        "表头及数据起点由模型判断，不再由 Python 猜列类型；信息不足时应追问。"
    )
    tf = _textbox(slide, 1.15, 6.35, 11.2, 0.5)
    p = _para(tf, True, space_after=0)
    _run(p.add_run(), "三层交付共享同一套核心：excel_formula 包 → CLI / Web / Skill", 14, color=SUB, italic=True)


def slide_intents(prs):
    slide = _content_slide(prs, "三、六类自然语言意图", 4)
    rows = [
        ("意图", "示例", "执行方式"),
        ("生成公式 / 常量", "“在 H2 算三科均分，填充到 H13”", "模型生成 → 校验 → 确认 → 写入"),
        ("解释公式", "“解释一下 G2 里的公式”", "DeepSeek 讲解，支持流式输出"),
        ("校验公式", "“=SUM(D2:F2) 这个公式对不对？”", "纯本地校验＋试算（0 Token）"),
        ("表结构概览", "“这个表有哪些列？”", "纯本地读取（0 Token）"),
        ("表格加框", "“把 A3 到 B8 框起来”", "纯本地套边框（0 Token）"),
        ("设置下拉列表", "“B3 设置下拉，选项为：华东、华南”", "本地规则优先，模型理解兜底"),
    ]
    shape = slide.shapes.add_table(len(rows), 3, Inches(1.15), Inches(1.65), Inches(11.2), Inches(4.3))
    table = shape.table
    table.columns[0].width = Inches(2.6)
    table.columns[1].width = Inches(4.4)
    table.columns[2].width = Inches(4.2)
    for r, row in enumerate(rows):
        for c, cell_text in enumerate(row):
            cell = table.cell(r, c)
            cell.margin_left = cell.margin_right = Inches(0.12)
            cell.margin_top = cell.margin_bottom = Inches(0.04)
            if r == 0:
                cell.fill.solid()
                cell.fill.fore_color.rgb = ACCENT
            elif r % 2 == 0:
                cell.fill.solid()
                cell.fill.fore_color.rgb = LIGHT
            tf = cell.text_frame
            p = tf.paragraphs[0]
            p.space_after = Pt(0)
            _run(p.add_run(), cell_text, 14.5 if r else 15,
                 bold=(r == 0), color=WHITE if r == 0 else INK)
    tf = _textbox(slide, 1.15, 6.15, 11.2, 0.65)
    p = _para(tf, True, space_after=0)
    _run(p.add_run(), "下拉来源支持固定选项或单元格区域；本地命中为 0 Token，兜底调用单独计量；确认后可在网页选值。",
         14, color=SUB)
    slide.notes_slide.notes_text_frame.text = (
        "意图识别本身不消耗 Token，不代表所有意图都零调用。下拉列表本地解析失败后，"
        "DeepSeek 只理解目标与来源，结果必须通过本地校验。\n"
        "没有明确指定目标工作表时，下拉落在当前表，来源表名不会把目标迁移走。"
        "网页选择只接受服务器规则中的选项索引，并校验规则和文件版本。"
    )


def slide_verify(prs):
    slide = _content_slide(prs, "四、亮点一：本地严格校验闭环", 5)
    _bullets(slide, [
        ("六道检查：", "语法 AST · 函数白名单 · 引用范围 · 循环引用 · 参数个数 · 逐格试算"),
        ("整批错误拦截：", "含填充中间格；#DIV/0!、#N/A 等确定错误进入修复，失败则拒绝写入"),
        ("有限自动修复：", "默认最多 2 轮，追加错误反馈；预期结果由 Python 独立试算"),
        ("求值能力扩展：", "日期序列号、跨工作簿引用、区域数组运算、VLOOKUP 左向查找"),
    ], h=3.65, size=18, gap=16)
    tf = _textbox(slide, 1.15, 5.55, 11.2, 1.1)
    p = _para(tf, True, space_after=6)
    _run(p.add_run(), "明确错误阻断，能力边界标注；预览结果仍需人工确认。", 16, bold=True, color=ACCENT)
    p = _para(tf, False, space_after=0)
    _run(p.add_run(), "缺失外部文件、无缓存依赖或尚未支持的求值场景标为“未验证”，不等于验证通过。", 14, color=SUB)
    slide.notes_slide.notes_text_frame.text = (
        "区域数组运算已支持：区域直接参与一元/二元运算按数组语义逐元素求值，"
        "这类公式写入时自动以数组公式（CSE）形态保存，Excel/WPS 打开与网页显示一致。\n"
        "主簿无缓存依赖在写入前校验中保守标未验证；网页显示可以递归试算，两条路径语义不同。\n"
        "跨工作簿试算优先读缓存，缺缓存可递归；外部文件不可用时标未验证。"
        "静态检查合法但无法试算的公式可以带警告进入确认，不宣传为百分之百正确。"
    )


def slide_safety(prs):
    slide = _content_slide(prs, "五、亮点二：安全与可靠写入", 6)
    _bullets(slide, [
        ("写入前确认：", "预览影响范围与覆盖内容；Web 方案使用一次性令牌，默认自动备份"),
        ("一致性保护：", "保存共用锁＋文件版本核对＋原子替换；下拉方案拒绝过期版本"),
        ("输入边界：", "主文件目录白名单、危险函数拦截、常量禁止以 = 开头、单批写入上限"),
        ("富对象保真：", "写前检测风险、写后对比部件数量；复杂对象不承诺完整保留"),
        ("格式智能：", "新单元格沿用相邻样式；新建表格自动设置浅蓝表头与细边框"),
    ], size=18, gap=15)
    slide.notes_slide.notes_text_frame.text = (
        "默认单批公式或常量写入不超过 200 格。备份可人工恢复，不是自动回滚功能。\n"
        "下拉预览保存文件版本，确认和选值时复核；所有保存共用锁，写盘前后再次核对版本。\n"
        "缺少 Pillow 而工作簿含普通图片时阻止保存；WMF 等对象可能丢失，"
        "保存后比较图片、图表、透视表部件数量并报告。不能据此宣称复杂 Excel 对象完全保真。"
    )


def slide_engineering(prs):
    slide = _content_slide(prs, "六、亮点三：工程与质量", 7)
    _bullets(slide, [
        ("轻量 Web：", "标准库 http.server ＋ 原生 JS/CSS；SSE 流式输出，无额外 Web 框架"),
        ("交互可观测：", "思考 1 / 0 / auto、用量统计；自适应网格与 50%–200% 缩放"),
        ("网格联动：", "公式结果与悬停原式、写入高亮、按工作簿 / 工作表隔离标记、下拉选值"),
        ("当前回归结果：", "610 通过 / 8 跳过；模型调用由 FakeClient / FakeSession 替代"),
        ("验证范围：", "函数求值、跨表 / 外部引用、下拉、保真、Web 闭环与大作业缓存对账"),
    ], h=4.3, size=18, gap=15)
    tf = _textbox(slide, 1.15, 6.3, 11.2, 0.4)
    p = _para(tf, True, space_after=0)
    _run(p.add_run(), "2026-09-18 实测；8 项因当前 PATH 未找到 Node.js 跳过，不计为通过。", 13, color=SUB)
    slide.notes_slide.notes_text_frame.text = (
        "验证命令：venv\\Scripts\\python.exe -m pytest -q。结果：610 passed, 8 skipped, 2 warnings。\n"
        "跳过项为 1 个前端脚本语法检查和 7 个前端事件测试，原因均为当前 PATH 未找到 Node.js。\n"
        "2 条警告来自大作业读取：openpyxl 不支持扩展数据验证。大作业对账只验证支持的公式，"
        "未知语法与特定错误值不在缓存一致性断言范围内。\n"
        "网页公式结果优先使用 Excel 缓存，无缓存时本地递归试算；模拟运算表对象退回缓存展示，"
        "不等于 Python 已实现 Excel 模拟运算表引擎。"
    )


def slide_demo(prs):
    slide = _content_slide(prs, "七、现场演示（约 90 秒）", 8)
    _bullets(slide, [
        ("① 读取成绩单　", "选择工作簿与工作表，询问“这个表有哪些列？”（0 Token）"),
        ("② 生成三科平均分　", "在 H2 计算平均分，保留 1 位小数，填充到 H13"),
        ("③ 预览并确认写入　", "核对预期结果与范围，观察写入高亮、格式沿用和备份"),
        ("④ 设置订单下拉　", "订单查询 B3 引用销售订单 A2:A25；确认后在网页选择订单号"),
        ("补充展示：", "跨表统计、解释 / 校验公式、表格加框与新建表格自动套格式"),
    ], size=18, gap=14)
    tf = _textbox(slide, 1.15, 6.3, 11.2, 0.5)
    p = _para(tf, True, space_after=0)
    _run(p.add_run(), "演示文件：答辩演示.xlsx · 5 张工作表；先展示公式主链路，再展示下拉交互。", 14, color=SUB, italic=True)
    slide.notes_slide.notes_text_frame.text = (
        "建议分配：背景与架构 45 秒、意图 30 秒、三项亮点 100 秒、演示 90 秒、总结 35 秒。\n"
        "演示前：打开 Web，选答辩演示.xlsx，可将思考设为 0；只对演示副本操作，不重建现有数据。\n"
        "成绩单指令：在 H2 计算每位学生语文、数学、英语的平均分，保留 1 位小数，填充到 H13。"
        "首行 D2:F2 为 88、92、85，预期 H2 为 88.3。先确认预览，再写入。\n"
        "切到订单查询，输入：把 B3 设置为下拉列表，来源为销售订单!A2:A25。"
        "本地规则命中应为 0 Token；先确认规则，再选择订单号。\n"
        "当前演示文件的订单查询 B4 没有预置公式，不演示解释空格；可改为解释成绩单 G2。"
        "也不要承诺选订单号后客户自动更新，除非先生成对应查找公式。\n"
        "可选指令：=SUM(D2:F2) 这个公式对不对？；将 A3 到 B8 框起来。"
        "网络较慢时保留核心生成与下拉流程，其他能力口头说明。"
    )


def slide_summary(prs):
    slide = _content_slide(prs, "八、总结与展望", 9)
    _bullets(slide, [
        ("达成的目标：", "自然语言驱动公式、格式与下拉交互；校验、确认、备份形成闭环"),
        ("关键数字：", "6 类意图 · 638 项测试通过（另 8 项跳过）· 3 个入口共享核心"),
        ("当前边界：", "本地求值不是完整 Excel 引擎；复杂对象、未验证结果仍需人工复核"),
        ("下一步：", "扩展财务与统计函数求值、公式审计、多用户会话隔离与大表处理"),
    ], size=18, gap=18)
    tf = _textbox(slide, 1.15, 5.5, 11.2, 0.8)
    p = _para(tf, True, space_after=0)
    _run(p.add_run(), "让 Excel 公式“说人话”——谢谢各位老师！", 20, bold=True, color=ACCENT)


def slide_thanks(prs):
    slide = _blank(prs)
    _rect(slide, 0, 0, SLIDE_W, SLIDE_H, COVER_BG)
    tf = _textbox(slide, 0, 2.9, SLIDE_W, 1.2)
    p = _para(tf, True, space_after=0)
    p.alignment = PP_ALIGN.CENTER
    _run(p.add_run(), "谢谢聆听", 44, bold=True, color=WHITE)
    tf2 = _textbox(slide, 0, 4.25, SLIDE_W, 0.8)
    p = _para(tf2, True, space_after=0)
    p.alignment = PP_ALIGN.CENTER
    _run(p.add_run(), "恳请各位老师批评指正", 18, color=RGBColor(0x9F, 0xB3, 0xCC))


def main() -> None:
    prs = Presentation()
    prs.slide_width = Inches(SLIDE_W)
    prs.slide_height = Inches(SLIDE_H)
    slide_cover(prs)
    slide_background(prs)
    slide_arch(prs)
    slide_intents(prs)
    slide_verify(prs)
    slide_safety(prs)
    slide_engineering(prs)
    slide_demo(prs)
    slide_summary(prs)
    slide_thanks(prs)
    prs.save(OUT)
    print(f"已生成 {OUT.name}（{len(prs.slides._sldIdLst)} 页）")


if __name__ == "__main__":
    main()
