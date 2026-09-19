# -*- coding: utf-8 -*-
"""把 report_content.CONTENT 渲染为《智能体开发实战》课程实验报告 docx。

环境无 python-docx 依赖，直接按 OOXML 最小可用结构手工构建：
  [Content_Types].xml / _rels/.rels / word/document.xml / word/styles.xml
  / word/footer1.xml / word/_rels/document.xml.rels
生成后自动做三重校验：ZIP 完整性、XML 良构性、段落标签配平。
"""
import os
import re
import sys
import zipfile
from xml.etree import ElementTree as ET
from xml.sax.saxutils import escape

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from report_content import CONTENT, TITLE, SUBTITLE, META

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_DECL = ('xmlns:w="{}" xmlns:r="{}"'.format(W_NS, R_NS))

PAGE_WIDTH = 11906          # A4 宽（twips）
MARGIN = 1800               # 左右页边距（twips，约 3.17cm）
CONTENT_WIDTH = PAGE_WIDTH - 2 * MARGIN   # 表格可用宽度

# --------------------------------------------------------------------- 样式
STYLES_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles {ns}>
 <w:docDefaults>
  <w:rPrDefault><w:rPr>
   <w:rFonts w:ascii="Times New Roman" w:eastAsia="\u5b8b\u4f53" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>
   <w:sz w:val="24"/><w:szCs w:val="24"/>
  </w:rPr></w:rPrDefault>
  <w:pPrDefault><w:pPr><w:spacing w:line="360" w:lineRule="auto"/></w:pPr></w:pPrDefault>
 </w:docDefaults>
 <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
  <w:name w:val="Normal"/><w:qFormat/>
  <w:pPr>
   <w:spacing w:before="0" w:after="60" w:line="360" w:lineRule="auto"/>
   <w:ind w:firstLineChars="200" w:firstLine="480"/>
   <w:jc w:val="both"/>
  </w:pPr>
  <w:rPr><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr>
 </w:style>
 <w:style w:type="paragraph" w:styleId="Title">
  <w:name w:val="Title"/><w:basedOn w:val="Normal"/><w:qFormat/>
  <w:pPr><w:keepNext/><w:spacing w:before="480" w:after="120" w:line="360" w:lineRule="auto"/><w:ind w:firstLine="0" w:firstLineChars="0"/><w:jc w:val="center"/></w:pPr>
  <w:rPr><w:rFonts w:ascii="Times New Roman" w:eastAsia="\u9ed1\u4f53" w:hAnsi="Times New Roman"/><w:sz w:val="44"/><w:szCs w:val="44"/></w:rPr>
 </w:style>
 <w:style w:type="paragraph" w:styleId="SubTitle">
  <w:name w:val="Subtitle"/><w:basedOn w:val="Normal"/><w:qFormat/>
  <w:pPr><w:keepNext/><w:spacing w:before="60" w:after="240" w:line="360" w:lineRule="auto"/><w:ind w:firstLine="0" w:firstLineChars="0"/><w:jc w:val="center"/></w:pPr>
  <w:rPr><w:rFonts w:ascii="Times New Roman" w:eastAsia="\u9ed1\u4f53" w:hAnsi="Times New Roman"/><w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr>
 </w:style>
 <w:style w:type="paragraph" w:styleId="MetaLine">
  <w:name w:val="MetaLine"/><w:basedOn w:val="Normal"/>
  <w:pPr><w:spacing w:before="0" w:after="120" w:line="360" w:lineRule="auto"/><w:ind w:firstLine="0" w:firstLineChars="0"/><w:jc w:val="center"/></w:pPr>
  <w:rPr><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr>
 </w:style>
 <w:style w:type="paragraph" w:styleId="Heading1">
  <w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:qFormat/>
  <w:pPr>
   <w:keepNext/><w:keepLines/>
   <w:spacing w:before="240" w:after="120" w:line="360" w:lineRule="auto"/>
   <w:ind w:firstLine="0" w:firstLineChars="0"/>
   <w:outlineLvl w:val="0"/>
  </w:pPr>
  <w:rPr><w:rFonts w:ascii="Times New Roman" w:eastAsia="\u9ed1\u4f53" w:hAnsi="Times New Roman"/><w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr>
 </w:style>
 <w:style w:type="paragraph" w:styleId="Heading2">
  <w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:qFormat/>
  <w:pPr>
   <w:keepNext/><w:keepLines/>
   <w:spacing w:before="180" w:after="100" w:line="360" w:lineRule="auto"/>
   <w:ind w:firstLine="0" w:firstLineChars="0"/>
   <w:outlineLvl w:val="1"/>
  </w:pPr>
  <w:rPr><w:rFonts w:ascii="Times New Roman" w:eastAsia="\u9ed1\u4f53" w:hAnsi="Times New Roman"/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr>
 </w:style>
 <w:style w:type="paragraph" w:styleId="Mono">
  <w:name w:val="Mono"/><w:basedOn w:val="Normal"/>
  <w:pPr>
   <w:keepLines/>
   <w:spacing w:before="0" w:after="0" w:line="300" w:lineRule="exact"/>
   <w:ind w:firstLine="0" w:firstLineChars="0"/>
  </w:pPr>
  <w:rPr><w:rFonts w:ascii="Consolas" w:eastAsia="\u5b8b\u4f53" w:hAnsi="Consolas"/><w:sz w:val="21"/><w:szCs w:val="21"/></w:rPr>
 </w:style>
 <w:style w:type="paragraph" w:styleId="Caption">
  <w:name w:val="Caption"/><w:basedOn w:val="Normal"/>
  <w:pPr>
   <w:keepNext/>
   <w:spacing w:before="160" w:after="60" w:line="300" w:lineRule="auto"/>
   <w:ind w:firstLine="0" w:firstLineChars="0"/>
   <w:jc w:val="center"/>
  </w:pPr>
  <w:rPr><w:b/><w:bCs/><w:sz w:val="21"/><w:szCs w:val="21"/></w:rPr>
 </w:style>
 <w:style w:type="paragraph" w:styleId="TableText">
  <w:name w:val="TableText"/><w:basedOn w:val="Normal"/>
  <w:pPr>
   <w:spacing w:before="20" w:after="20" w:line="280" w:lineRule="auto"/>
   <w:ind w:firstLine="0" w:firstLineChars="0"/>
   <w:jc w:val="left"/>
  </w:pPr>
  <w:rPr><w:sz w:val="21"/><w:szCs w:val="21"/></w:rPr>
 </w:style>
</w:styles>
""".replace("{ns}", NS_DECL)

FOOTER_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr {ns}>
 <w:p>
  <w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/><w:ind w:firstLine="0" w:firstLineChars="0"/><w:jc w:val="center"/></w:pPr>
  <w:r><w:rPr><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr><w:t xml:space="preserve">\u7b2c </w:t></w:r>
  <w:r><w:rPr><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr><w:fldChar w:fldCharType="begin"/></w:r>
  <w:r><w:rPr><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>
  <w:r><w:rPr><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr><w:fldChar w:fldCharType="separate"/></w:r>
  <w:r><w:rPr><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr><w:t>1</w:t></w:r>
  <w:r><w:rPr><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr><w:fldChar w:fldCharType="end"/></w:r>
  <w:r><w:rPr><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr><w:t xml:space="preserve"> \u9875</w:t></w:r>
 </w:p>
</w:ftr>
""".replace("{ns}", NS_DECL)

CONTENT_TYPES_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
 <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
 <Default Extension="xml" ContentType="application/xml"/>
 <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
 <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
 <Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>
</Types>
"""

ROOT_RELS_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
 <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>
"""

DOC_RELS_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
 <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
 <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>
</Relationships>
"""

SECT_PR = (
    '<w:sectPr><w:footerReference w:type="default" r:id="rId2"/>'
    '<w:pgSz w:w="11906" w:h="16838"/>'
    '<w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800" '
    'w:header="851" w:footer="992" w:gutter="0"/>'
    '<w:cols w:space="425"/><w:docGrid w:type="lines" w:linePitch="312"/></w:sectPr>'
)

TABLE_BORDERS = (
    '<w:tblBorders>'
    '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
    '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
    '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
    '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
    '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
    '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
    '</w:tblBorders>'
)


def run(text, bold=False):
    rpr = '<w:rPr><w:b/><w:bCs/></w:rPr>' if bold else ''
    return '<w:r>{}<w:t xml:space="preserve">{}</w:t></w:r>'.format(rpr, escape(text))


def para(style, runs_xml):
    return '<w:p><w:pPr><w:pStyle w:val="{}"/></w:pPr>{}</w:p>'.format(style, runs_xml)


def simple_para(style, text, bold=False):
    return para(style, run(text, bold))


def table_xml(spec):
    headers, rows, widths = spec["headers"], spec["rows"], spec["widths"]
    cols = len(headers)
    ws = [int(round(CONTENT_WIDTH * f)) for f in widths]
    ws[-1] = CONTENT_WIDTH - sum(ws[:-1])

    def cell(text, w, bold=False, shaded=False):
        # CT_TcPr 子元素顺序：tcW → shd → vAlign
        shd = '<w:shd w:val="clear" w:color="auto" w:fill="EFEFEF"/>' if shaded else ''
        return (
            '<w:tc><w:tcPr><w:tcW w:w="{}" w:type="dxa"/>{}</w:tcPr>'
            '<w:p><w:pPr><w:pStyle w:val="TableText"/></w:pPr>{}</w:p></w:tc>'
        ).format(w, shd, run(text, bold))

    grid = ''.join('<w:gridCol w:w="{}"/>'.format(w) for w in ws)
    head_row = '<w:tr><w:trPr><w:tblHeader/></w:trPr>{}</w:tr>'.format(
        ''.join(cell(h, w, bold=True, shaded=True) for h, w in zip(headers, ws))
    )
    body_rows = ''.join(
        '<w:tr>{}</w:tr>'.format(''.join(cell(c, w) for c, w in zip(row, ws)))
        for row in rows
    )
    return (
        '<w:tbl><w:tblPr><w:tblW w:w="{}" w:type="dxa"/>{}</w:tblPr>'
        '<w:tblGrid>{}</w:tblGrid>{}{}</w:tbl>'
    ).format(CONTENT_WIDTH, TABLE_BORDERS, grid, head_row, body_rows)


def build_body():
    parts = [simple_para("Title", TITLE), simple_para("SubTitle", SUBTITLE)]
    for line in META.split("\n"):
        parts.append(simple_para("MetaLine", line))
    parts.append('<w:p><w:pPr><w:pStyle w:val="MetaLine"/></w:pPr>'
                 '<w:r><w:br w:type="page"/></w:r></w:p>')  # 封面信息后分页
    for item in CONTENT:
        kind = item[0]
        if kind in ("h1", "h2", "h3"):
            parts.append(simple_para({"h1": "Heading1", "h2": "Heading2", "h3": "Heading3"}[kind], item[1]))
        elif kind == "p":
            parts.append(simple_para("Normal", item[1]))
        elif kind == "caption":
            parts.append(simple_para("Caption", item[1]))
        elif kind == "diagram":
            parts.extend(simple_para("Mono", line) for line in item[1])
        elif kind == "table":
            parts.append(table_xml(item[1]))
        else:
            raise ValueError("未知内容类型：{}".format(kind))
    return ''.join(parts)


def build_document():
    body = build_body() + SECT_PR
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<w:document {}><w:body>{}</w:body></w:document>'.format(NS_DECL, body)
    )


def write_docx(path):
    parts = [
        ("[Content_Types].xml", CONTENT_TYPES_XML),
        ("_rels/.rels", ROOT_RELS_XML),
        ("word/document.xml", build_document()),
        ("word/styles.xml", STYLES_XML),
        ("word/footer1.xml", FOOTER_XML),
        ("word/_rels/document.xml.rels", DOC_RELS_XML),
    ]
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in parts:
            z.writestr(name, data)


def verify(path):
    ok = True
    with zipfile.ZipFile(path) as z:
        bad = z.testzip()
        print("ZIP 完整性:", "OK" if bad is None else "损坏: {}".format(bad))
        ok = ok and bad is None
        for name in ("word/document.xml", "word/styles.xml", "word/footer1.xml",
                     "[Content_Types].xml", "_rels/.rels", "word/_rels/document.xml.rels"):
            try:
                ET.fromstring(z.read(name))
                print("XML 良构  :", name, "OK")
            except ET.ParseError as exc:
                ok = False
                print("XML 良构  :", name, "失败 ->", exc)
        doc = z.read("word/document.xml").decode("utf-8")
        open_ps = len(re.findall(r"<w:p[ >]", doc))
        close_ps = len(re.findall(r"</w:p>", doc))
        print("段落配平  : <w:p>={} </w:p>={} -> {}".format(
            open_ps, close_ps, "OK" if open_ps == close_ps else "不配平"))
        ok = ok and open_ps == close_ps
        tables = len(re.findall(r"<w:tbl>", doc))
        print("表格数量  :", tables)
        # 反向抽取文本，确认正文关键词完整
        text = "".join(re.findall(r"<w:t[^>]*>(.*?)</w:t>", doc, re.S))
        for kw in ("一、实验概述", "二、需求分析", "三、总体架构", "四、关键实现",
                   "五、测试结果", "六、安全措施与异常处理", "七、低 Token 与性能优化",
                   "八、可选功能：公式自动验证", "九、不足与改进方向", "十、总结"):
            if kw not in text:
                ok = False
                print("关键词缺失:", kw)
        cjk = len(re.findall(r"[\u4e00-\u9fff]", doc))
        print("中文字数  :", cjk, "（仅统计汉字，未含数字/英文/标点）")
        print("总字符数  :", len(text))
    print("\n总体校验  :", "全部通过" if ok else "存在问题！")
    return ok


if __name__ == "__main__":
    target = sys.argv[1] if len(sys.argv) > 1 else "实验报告.docx"
    write_docx(target)
    print("已生成:", os.path.abspath(target), "\n")
    sys.exit(0 if verify(target) else 1)
