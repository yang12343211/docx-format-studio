# -*- coding: utf-8 -*-
"""生成测试用的 DOCX：一个格式模板 + 一个格式混乱的待处理文档"""
import zipfile, os, io

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
DIR = os.path.dirname(os.path.abspath(__file__))

XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'

CONTENT_TYPES = XML_DECL + '''<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>'''

ROOT_RELS = XML_DECL + '''<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>'''

DOC_RELS = XML_DECL + '''<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>'''


def doc(body, sectpr):
    return (XML_DECL +
            '<w:document xmlns:w="%s"><w:body>%s%s</w:body></w:document>' % (W, body, sectpr))


def p(text, style=None, rpr=None, ppr=None, ):
    """构造一个段落"""
    ppr_xml = ''
    if style or ppr:
        inner = ''
        if style:
            inner += '<w:pStyle w:val="%s"/>' % style
        if ppr:
            inner += ppr
        ppr_xml = '<w:pPr>%s</w:pPr>' % inner
    rpr_xml = '<w:rPr>%s</w:rPr>' % rpr if rpr else ''
    return ('<w:p>%s<w:r>%s<w:t xml:space="preserve">%s</w:t></w:r></w:p>'
            % (ppr_xml, rpr_xml, text))


def table(rows):
    trs = ''
    for row in rows:
        tcs = ''
        for cell in row:
            tcs += '<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>%s</w:tc>' % p(cell)
        trs += '<w:tr>%s</w:tr>' % tcs
    return ('<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>'
            '<w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>%s</w:tbl>' % trs)


# ---------------------------------------------------------------- 模板
TEMPLATE_STYLES = XML_DECL + '''<w:styles xmlns:w="%s">
<w:docDefaults>
  <w:rPrDefault><w:rPr>
    <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="宋体" w:cs="Times New Roman"/>
    <w:sz w:val="24"/><w:szCs w:val="24"/>
  </w:rPr></w:rPrDefault>
  <w:pPrDefault><w:pPr><w:spacing w:line="360" w:lineRule="auto"/></w:pPr></w:pPrDefault>
</w:docDefaults>

<w:style w:type="paragraph" w:default="1" w:styleId="Normal">
  <w:name w:val="Normal"/><w:qFormat/>
  <w:pPr><w:jc w:val="both"/><w:spacing w:line="360" w:lineRule="auto"/>
    <w:ind w:firstLineChars="200" w:firstLine="480"/></w:pPr>
  <w:rPr><w:rFonts w:ascii="Times New Roman" w:eastAsia="宋体" w:hAnsi="Times New Roman"/>
    <w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr>
</w:style>

<w:style w:type="paragraph" w:styleId="Heading1">
  <w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>
  <w:pPr><w:keepNext/><w:spacing w:before="240" w:after="240" w:line="360" w:lineRule="auto"/>
    <w:jc w:val="center"/><w:outlineLvl w:val="0"/></w:pPr>
  <w:rPr><w:rFonts w:ascii="Times New Roman" w:eastAsia="黑体" w:hAnsi="Times New Roman"/>
    <w:b w:val="0"/><w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr>
</w:style>

<w:style w:type="paragraph" w:styleId="Heading2">
  <w:name w:val="标题 2"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>
  <w:pPr><w:keepNext/><w:spacing w:before="120" w:after="120" w:line="360" w:lineRule="auto"/>
    <w:jc w:val="left"/><w:outlineLvl w:val="1"/></w:pPr>
  <w:rPr><w:rFonts w:ascii="Times New Roman" w:eastAsia="黑体" w:hAnsi="Times New Roman"/>
    <w:b w:val="0"/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr>
</w:style>

<w:style w:type="paragraph" w:styleId="3">
  <w:name w:val="标题 3"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>
  <w:pPr><w:spacing w:before="60" w:after="60" w:line="360" w:lineRule="auto"/>
    <w:jc w:val="left"/><w:outlineLvl w:val="2"/></w:pPr>
  <w:rPr><w:rFonts w:ascii="Times New Roman" w:eastAsia="黑体" w:hAnsi="Times New Roman"/>
    <w:b/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr>
</w:style>

<w:style w:type="paragraph" w:styleId="Caption">
  <w:name w:val="题注"/><w:basedOn w:val="Normal"/>
  <w:pPr><w:spacing w:before="60" w:after="60" w:line="280" w:lineRule="auto"/>
    <w:jc w:val="center"/><w:ind w:firstLine="0" w:firstLineChars="0"/></w:pPr>
  <w:rPr><w:rFonts w:ascii="Times New Roman" w:eastAsia="宋体" w:hAnsi="Times New Roman"/>
    <w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr>
</w:style>
</w:styles>''' % W

TEMPLATE_SECTPR = ('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
                   '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1800" '
                   'w:header="851" w:footer="992" w:gutter="0"/>'
                   '<w:docGrid w:type="lines" w:linePitch="312"/></w:sectPr>')

template_body = ''.join([
    p('高校实验设备管理信息系统', style='Heading1'),
    p('本文档用于展示学校要求的格式规范。', style='Normal'),
    p('1.1 研究背景', style='Heading2'),
    p('随着高校实验室规模不断扩大，传统的人工管理方式已难以满足需求。', style='Normal'),
    p('1.1.1 现状分析', style='3'),
])

# ---------------------------------------------------------------- 待处理文档
TARGET_STYLES = XML_DECL + '''<w:styles xmlns:w="%s">
<w:docDefaults>
  <w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:eastAsia="等线" w:hAnsi="Calibri"/>
    <w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:rPrDefault>
  <w:pPrDefault><w:pPr/></w:pPrDefault>
</w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal">
  <w:name w:val="Normal"/><w:qFormat/>
  <w:rPr><w:rFonts w:ascii="Calibri" w:eastAsia="等线" w:hAnsi="Calibri"/>
    <w:sz w:val="22"/></w:rPr>
</w:style>
<w:style w:type="paragraph" w:styleId="Heading1">
  <w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:qFormat/>
  <w:rPr><w:b/><w:sz w:val="36"/></w:rPr>
</w:style>
<w:style w:type="paragraph" w:styleId="Heading2">
  <w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:qFormat/>
  <w:rPr><w:b/><w:sz w:val="30"/></w:rPr>
</w:style>
<w:style w:type="paragraph" w:styleId="Caption">
  <w:name w:val="caption"/><w:basedOn w:val="Normal"/>
</w:style>
<w:style w:type="paragraph" w:styleId="TOC1">
  <w:name w:val="toc 1"/><w:basedOn w:val="Normal"/>
</w:style>
</w:styles>''' % W

TARGET_SECTPR = ('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
                 '<w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720" '
                 'w:header="851" w:footer="992" w:gutter="0"/></w:sectPr>')

# 手动加粗变红的 H1（模拟学生乱调格式）
BAD_H1_RPR = '<w:rFonts w:ascii="Arial" w:eastAsia="微软雅黑" w:hAnsi="Arial"/><w:color w:val="FF0000"/><w:sz w:val="40"/><w:b/>'
# 正文里手工设了 18 磅、单倍行距
BAD_BODY_PPR = '<w:spacing w:line="240" w:lineRule="auto"/>'

target_body = ''.join([
    p('第一章  绪论', style='Heading1', rpr=BAD_H1_RPR),
    p('高校实验室是培养人才的重要场所，设备管理信息化势在必行。', style='Normal', ppr=BAD_BODY_PPR),
    p('1.1 研究背景', style='Heading2'),
    p('本课题来源于实际教学需求，旨在解决设备台账混乱的问题。', style='Normal'),
    table([['设备编号', '设备名称'], ['20240101001', '示波器'], ['20240101002', '万用表']]),
    p('图 1-1 系统总体架构图', style='Caption'),
    p('2.1 系统需求分析'),          # 无样式，靠启发式识别
    p('系统需要实现用户管理、设备管理、查询统计等功能。', style='Normal'),
    p('目  录', style='TOC1'),
])

# ---------------------------------------------------------------- 打包


def build(path, styles_xml, body, sectpr):
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', CONTENT_TYPES)
        z.writestr('_rels/.rels', ROOT_RELS)
        z.writestr('word/_rels/document.xml.rels', DOC_RELS)
        z.writestr('word/styles.xml', styles_xml)
        z.writestr('word/document.xml', doc(body, sectpr))
    print('written:', path, os.path.getsize(path), 'bytes')


build(os.path.join(DIR, 'template.docx'), TEMPLATE_STYLES, template_body, TEMPLATE_SECTPR)
build(os.path.join(DIR, 'target.docx'), TARGET_STYLES, target_body, TARGET_SECTPR)
