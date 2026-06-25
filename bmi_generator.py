#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Professional BMI Generator — Artikuli 1660 Flanel to'qimasi
Barcha boblar, jadvallar, formulalar, sahifa raqamlari bilan.
"""
import zipfile, xml.etree.ElementTree as ET, json, os, re

BASE = os.path.dirname(os.path.abspath(__file__))
WD   = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'

# ═══════════════════════════════════════════════════════════
# 1. MA'LUMOT O'QISH
# ═══════════════════════════════════════════════════════════
def read_docx(path):
    result = []
    with zipfile.ZipFile(path) as z:
        with z.open('word/document.xml') as f:
            tree = ET.parse(f)
    body = tree.getroot().find('.//' + WD + 'body')
    if body is None:
        return result
    for child in body:
        tag = child.tag.split('}')[-1]
        if tag == 'p':
            style = ''
            pPr = child.find(WD + 'pPr')
            if pPr is not None:
                ps = pPr.find(WD + 'pStyle')
                if ps is not None:
                    style = ps.get(WD + 'val', '')
            parts = []
            for t in child.iter(WD + 't'):
                if t.text:
                    parts.append(t.text)
            result.append(('p', style, ''.join(parts)))
        elif tag == 'tbl':
            rows = []
            for tr in child.iter(WD + 'tr'):
                cells = []
                for tc in tr.findall('.//' + WD + 'tc'):
                    cp = []
                    for t in tc.iter(WD + 't'):
                        if t.text:
                            cp.append(t.text)
                    cells.append(''.join(cp))
                rows.append(cells)
            result.append(('t', '', rows))
    return result

kirish  = read_docx(os.path.join(BASE, "Kenjayeva_Diplom_Kirish (2).docx"))
iqtisod = read_docx(os.path.join(BASE, "Z.Kenjayeva_dastgoh_Iqtisod_13.06.2026.docx"))
mehnat  = read_docx(os.path.join(BASE, "77777.docx"))


# ═══════════════════════════════════════════════════════════
# 2. XML YORDAMCHI FUNKSIYALAR
# ═══════════════════════════════════════════════════════════
def x(s):
    """XML escape"""
    if not s:
        return ''
    return (str(s)
            .replace('&','&amp;').replace('<','&lt;')
            .replace('>','&gt;').replace('"','&quot;'))

def rpr(bold=False, italic=False, sz=24, font='Times New Roman',
        color=None, underline=False, strike=False):
    parts = [f'<w:rFonts w:ascii="{font}" w:hAnsi="{font}" w:cs="{font}"/>']
    if bold:        parts.append('<w:b/><w:bCs/>')
    if italic:      parts.append('<w:i/><w:iCs/>')
    if underline:   parts.append('<w:u w:val="single"/>')
    if strike:      parts.append('<w:strike/>')
    if color:       parts.append(f'<w:color w:val="{color}"/>')
    parts.append(f'<w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/>')
    return '<w:rPr>' + ''.join(parts) + '</w:rPr>'

def ppr(style='Normal', align='both', before=0, after=120,
        line=480, first_indent=720, left_indent=0, page_break=False,
        keep_next=False, keep_lines=False):
    parts = [f'<w:pStyle w:val="{style}"/>']
    if page_break:  parts.append('<w:pageBreakBefore/>')
    if keep_next:   parts.append('<w:keepNext/>')
    if keep_lines:  parts.append('<w:keepLines/>')
    parts.append(f'<w:jc w:val="{align}"/>')
    parts.append(f'<w:spacing w:before="{before}" w:after="{after}" '
                 f'w:line="{line}" w:lineRule="auto"/>')
    if first_indent or left_indent:
        parts.append(f'<w:ind w:left="{left_indent}" w:firstLine="{first_indent}"/>')
    return '<w:pPr>' + ''.join(parts) + '</w:pPr>'

def para(text, style='Normal', align='both', bold=False, italic=False,
         sz=24, before=0, after=120, line=480, first_indent=720,
         left_indent=0, color=None, keep_next=False, underline=False):
    pp = ppr(style=style, align=align, before=before, after=after,
             line=line, first_indent=first_indent, left_indent=left_indent,
             keep_next=keep_next)
    rp = rpr(bold=bold, italic=italic, sz=sz, color=color, underline=underline)
    return (f'<w:p>{pp}<w:r>{rp}'
            f'<w:t xml:space="preserve">{x(text)}</w:t></w:r></w:p>')

def empty(after=60):
    return f'<w:p><w:pPr><w:spacing w:after="{after}"/></w:pPr></w:p>'

def page_br():
    return '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'

def h1(text):
    return (f'<w:p>'
            f'<w:pPr><w:pStyle w:val="Heading1"/>'
            f'<w:jc w:val="center"/>'
            f'<w:spacing w:before="360" w:after="180"/>'
            f'<w:keepNext/></w:pPr>'
            f'<w:r><w:rPr>'
            f'<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>'
            f'<w:b/><w:bCs/><w:sz w:val="28"/><w:szCs w:val="28"/>'
            f'</w:rPr><w:t xml:space="preserve">{x(text)}</w:t></w:r></w:p>')

def h2(text):
    return (f'<w:p>'
            f'<w:pPr><w:pStyle w:val="Heading2"/>'
            f'<w:jc w:val="left"/>'
            f'<w:spacing w:before="240" w:after="120"/>'
            f'<w:keepNext/></w:pPr>'
            f'<w:r><w:rPr>'
            f'<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>'
            f'<w:b/><w:bCs/><w:sz w:val="26"/><w:szCs w:val="26"/>'
            f'</w:rPr><w:t xml:space="preserve">{x(text)}</w:t></w:r></w:p>')

def h3(text):
    return (f'<w:p>'
            f'<w:pPr><w:pStyle w:val="Heading3"/>'
            f'<w:jc w:val="left"/>'
            f'<w:spacing w:before="160" w:after="80"/>'
            f'<w:keepNext/></w:pPr>'
            f'<w:r><w:rPr>'
            f'<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>'
            f'<w:b/><w:bCs/><w:sz w:val="24"/><w:szCs w:val="24"/>'
            f'</w:rPr><w:t xml:space="preserve">{x(text)}</w:t></w:r></w:p>')

def tbl_caption(text):
    return para(text, bold=True, align='center', sz=24, before=120,
                after=60, first_indent=0, line=360)

def formula(text):
    return para(text, align='center', bold=False, italic=False,
                sz=24, before=60, after=60, first_indent=0, line=360)

def note(text):
    return para(text, italic=True, sz=22, before=40, after=80,
                first_indent=360, line=360)


# ═══════════════════════════════════════════════════════════
# 3. JADVAL GENERATOR — to'liq professional
# ═══════════════════════════════════════════════════════════
def build_table(rows, col_widths=None, header_rows=1, small_font=False, total_width=9360):
    if not rows:
        return ''
    # Bo'sh qatorlarni filtrlash
    rows = [r for r in rows if any(c.strip() for c in r)]
    if not rows:
        return ''
    max_c = max(len(r) for r in rows)
    if max_c == 0:
        return ''
    if not col_widths:
        w = total_width // max_c
        col_widths = [w] * max_c
    # Yetarli kenglik bo'lishini ta'minlash
    while len(col_widths) < max_c:
        col_widths.append(col_widths[-1] if col_widths else 800)
    total_w = sum(col_widths[:max_c])
    sz = '18' if small_font else '20'

    lines = []
    lines.append(f'''<w:tbl>
<w:tblPr>
  <w:tblStyle w:val="TableGrid"/>
  <w:tblW w:w="{total_w}" w:type="dxa"/>
  <w:tblLayout w:type="fixed"/>
  <w:tblCellMar>
    <w:top w:w="60" w:type="dxa"/>
    <w:left w:w="108" w:type="dxa"/>
    <w:bottom w:w="60" w:type="dxa"/>
    <w:right w:w="108" w:type="dxa"/>
  </w:tblCellMar>
  <w:tblBorders>
    <w:top    w:val="single" w:sz="6" w:color="000000"/>
    <w:left   w:val="single" w:sz="6" w:color="000000"/>
    <w:bottom w:val="single" w:sz="6" w:color="000000"/>
    <w:right  w:val="single" w:sz="6" w:color="000000"/>
    <w:insideH w:val="single" w:sz="4" w:color="000000"/>
    <w:insideV w:val="single" w:sz="4" w:color="000000"/>
  </w:tblBorders>
</w:tblPr>
<w:tblGrid>''' + ''.join(f'<w:gridCol w:w="{col_widths[i]}"/>' for i in range(max_c)) + '</w:tblGrid>')

    for ri, row in enumerate(rows):
        is_hdr = ri < header_rows
        fill = 'BDD7EE' if is_hdr else 'FFFFFF'
        padded = list(row) + [''] * (max_c - len(row))
        lines.append('<w:tr>')
        if is_hdr:
            lines.append('<w:trPr><w:tblHeader/></w:trPr>')

        for ci in range(max_c):
            cell_text = padded[ci] if ci < len(padded) else ''
            cw = col_widths[ci] if ci < len(col_widths) else 1000
            b = '<w:b/><w:bCs/>' if is_hdr else ''
            lines.append(f'''<w:tc>
<w:tcPr>
  <w:tcW w:w="{cw}" w:type="dxa"/>
  <w:shd w:val="clear" w:color="auto" w:fill="{fill}"/>
</w:tcPr>
<w:p>
  <w:pPr>
    <w:jc w:val="center"/>
    <w:spacing w:before="40" w:after="40" w:line="320" w:lineRule="auto"/>
  </w:pPr>
  <w:r>
    <w:rPr>
      <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>
      {b}<w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/>
    </w:rPr>
    <w:t xml:space="preserve">{x(cell_text)}</w:t>
  </w:r>
</w:p>
</w:tc>''')
        lines.append('</w:tr>')
    lines.append('</w:tbl>')
    return '\n'.join(lines)

def get_tables(data):
    """data ro'yxatidan faqat jadvallarni list sifatida qaytaradi"""
    return [item for item in data if item[0] == 't']


# ═══════════════════════════════════════════════════════════
# 4. STYLES, SETTINGS, RELS XMLlari
# ═══════════════════════════════════════════════════════════
STYLES = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults>
  <w:rPrDefault><w:rPr>
    <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>
    <w:sz w:val="24"/><w:szCs w:val="24"/>
    <w:lang w:val="uz-Latn-UZ" w:eastAsia="uz-Latn-UZ" w:bidi="ar-SA"/>
  </w:rPr></w:rPrDefault>
  <w:pPrDefault><w:pPr>
    <w:jc w:val="both"/>
    <w:spacing w:after="120" w:line="480" w:lineRule="auto"/>
    <w:ind w:firstLine="720"/>
  </w:pPr></w:pPrDefault>
</w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal">
  <w:name w:val="Normal"/>
  <w:pPr>
    <w:jc w:val="both"/>
    <w:spacing w:after="120" w:line="480" w:lineRule="auto"/>
    <w:ind w:firstLine="720"/>
  </w:pPr>
  <w:rPr>
    <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>
    <w:sz w:val="24"/><w:szCs w:val="24"/>
  </w:rPr>
</w:style>
<w:style w:type="paragraph" w:styleId="Heading1">
  <w:name w:val="heading 1"/>
  <w:basedOn w:val="Normal"/>
  <w:pPr>
    <w:keepNext/><w:keepLines/>
    <w:numPr><w:ilvl w:val="0"/><w:numId w:val="0"/></w:numPr>
    <w:spacing w:before="360" w:after="180"/>
    <w:jc w:val="center"/>
    <w:ind w:firstLine="0"/>
  </w:pPr>
  <w:rPr>
    <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>
    <w:b/><w:bCs/><w:sz w:val="28"/><w:szCs w:val="28"/>
  </w:rPr>
</w:style>
<w:style w:type="paragraph" w:styleId="Heading2">
  <w:name w:val="heading 2"/>
  <w:basedOn w:val="Normal"/>
  <w:pPr>
    <w:keepNext/>
    <w:spacing w:before="240" w:after="120"/>
    <w:jc w:val="left"/>
    <w:ind w:firstLine="0"/>
  </w:pPr>
  <w:rPr>
    <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>
    <w:b/><w:bCs/><w:sz w:val="26"/><w:szCs w:val="26"/>
  </w:rPr>
</w:style>
<w:style w:type="paragraph" w:styleId="Heading3">
  <w:name w:val="heading 3"/>
  <w:basedOn w:val="Normal"/>
  <w:pPr>
    <w:keepNext/>
    <w:spacing w:before="160" w:after="80"/>
    <w:jc w:val="left"/>
    <w:ind w:firstLine="0"/>
  </w:pPr>
  <w:rPr>
    <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>
    <w:b/><w:bCs/><w:sz w:val="24"/><w:szCs w:val="24"/>
  </w:rPr>
</w:style>
<w:style w:type="table" w:styleId="TableGrid">
  <w:name w:val="Table Grid"/>
  <w:tblPr>
    <w:tblBorders>
      <w:top    w:val="single" w:sz="4" w:color="000000"/>
      <w:left   w:val="single" w:sz="4" w:color="000000"/>
      <w:bottom w:val="single" w:sz="4" w:color="000000"/>
      <w:right  w:val="single" w:sz="4" w:color="000000"/>
      <w:insideH w:val="single" w:sz="4" w:color="000000"/>
      <w:insideV w:val="single" w:sz="4" w:color="000000"/>
    </w:tblBorders>
  </w:tblPr>
  <w:rPr>
    <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>
    <w:sz w:val="20"/><w:szCs w:val="20"/>
  </w:rPr>
</w:style>
</w:styles>'''

SETTINGS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:defaultTabStop w:val="720"/>
  <w:compat>
    <w:compatSetting w:name="compatibilityMode"
      w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>
  </w:compat>
</w:settings>'''

CT = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml"
    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml"
    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/settings.xml"
    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
  <Override PartName="/word/footer1.xml"
    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>
</Types>'''

RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
    Target="word/document.xml"/>
</Relationships>'''

WORD_RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
    Target="styles.xml"/>
  <Relationship Id="rId2"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings"
    Target="settings.xml"/>
  <Relationship Id="rId3"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"
    Target="footer1.xml"/>
</Relationships>'''

FOOTER = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:p>
  <w:pPr><w:jc w:val="center"/></w:pPr>
  <w:r>
    <w:rPr>
      <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>
      <w:sz w:val="22"/><w:szCs w:val="22"/>
    </w:rPr>
    <w:fldChar w:fldCharType="begin"/>
  </w:r>
  <w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>
  <w:r><w:fldChar w:fldCharType="end"/></w:r>
</w:p>
</w:ftr>'''


# ═══════════════════════════════════════════════════════════
# 5. HUJJAT QISMLARI
# ═══════════════════════════════════════════════════════════
def sect_props(landscape=False):
    if landscape:
        return '''<w:sectPr>
  <w:footerReference w:type="default" r:id="rId3"/>
  <w:pgSz w:w="15840" w:h="12240" w:orient="landscape"/>
  <w:pgMar w:top="1134" w:right="850" w:bottom="1134" w:left="1800"
           w:header="709" w:footer="709" w:gutter="0"/>
</w:sectPr>'''
    return '''<w:sectPr>
  <w:footerReference w:type="default" r:id="rId3"/>
  <w:pgSz w:w="12240" w:h="15840"/>
  <w:pgMar w:top="1440" w:right="1008" w:bottom="1440" w:left="1800"
           w:header="720" w:footer="720" w:gutter="0"/>
  <w:pgNumType w:fmt="decimal" w:start="1"/>
</w:sectPr>'''

def landscape_sect():
    """Keng jadvallar uchun landscape sectPr bloki"""
    return (f'<w:p><w:pPr>{sect_props(True)}</w:pPr></w:p>')

# ─── TITUL ───────────────────────────────────────────────
def section_titul():
    B = []
    B.append(empty(200))
    B.append(para("O'ZBEKISTON RESPUBLIKASI OLIY TA'LIM, FAN VA INNOVATSIYALAR VAZIRLIGI",
                  bold=True, align='center', sz=24, first_indent=0,
                  before=0, after=160, line=360))
    B.append(para("TERMEZ DAVLAT MUHANDISLIK VA AGROTEXNOLOGIYALAR UNIVERSITETI",
                  bold=True, align='center', sz=26, first_indent=0,
                  before=0, after=120, line=360))
    B.append(para("ENERGETIKA VA SANOAT MUHANDISLIGI FAKULTETI",
                  bold=True, align='center', sz=24, first_indent=0,
                  before=0, after=320, line=360))
    B.append(empty(160))
    B.append(para("BITIRUV MALAKAVIY ISHI", bold=True, align='center', sz=32,
                  first_indent=0, before=0, after=200, line=360))
    B.append(empty(80))
    B.append(para("MAVZU:", bold=True, align='center', sz=24,
                  first_indent=0, before=0, after=60, line=360))
    B.append(para(
        "Quvvati 35 ta to\u02BBquv dastgohiga ega bo\u02BBlgan artikuli 1660 Flanel "
        "to\u02BBqimasini ishlab chiqarish uchun zamonaviy texnologik jarayonlarni loyihalash",
        bold=True, align='center', sz=26, first_indent=0,
        before=0, after=400, line=360))
    B.append(empty(200))
    rows_titul = [
        ["Guruh:", "YS22A"],
        ["Bajaruvchi:", "Kenjayeva Zulfinur"],
        ["Ilmiy rahbar:", "_______________________"],
        ["Lavozimi, ilmiy darajasi:", "_______________________"],
        ["Kafedra mudiri:", "_______________________"],
    ]
    for k, v in rows_titul:
        B.append(f'''<w:p>
<w:pPr><w:jc w:val="left"/>
<w:spacing w:before="40" w:after="40" w:line="360" w:lineRule="auto"/>
<w:ind w:left="2160" w:firstLine="0"/>
<w:tabs><w:tab w:val="left" w:pos="4320"/></w:tabs>
</w:pPr>
<w:r><w:rPr>
  <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>
  <w:sz w:val="24"/><w:szCs w:val="24"/>
</w:rPr><w:t xml:space="preserve">{x(k)}</w:t></w:r>
<w:r><w:tab/></w:r>
<w:r><w:rPr>
  <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>
  <w:sz w:val="24"/><w:szCs w:val="24"/>
</w:rPr><w:t xml:space="preserve">{x(v)}</w:t></w:r>
</w:p>''')
    B.append(empty(300))
    B.append(para("Termez \u2013 2026", bold=True, align='center', sz=26,
                  first_indent=0, before=200, after=0, line=360))
    B.append(page_br())
    return '\n'.join(B)


# ─── MUNDARIJA ───────────────────────────────────────────
def section_mundarija():
    B = []
    B.append(h1("MUNDARIJA"))
    toc = [
        (0, "KIRISH", "3"),
        (1, "I BOB. TEXNOLOGIK VA TEXNIK QISM", ""),
        (2, "1.1. Ishlab chiqarish texnologik jarayoni va parametrlar", ""),
        (2, "1.2. Sultzer Rutti P 7200 \u2014 normallashtirish kartasi (20-jadval)", ""),
        (2, "1.3. Texnologik hisob-kitoblar va formulalar (57\u201363)", ""),
        (2, "1.4. To\u02BBquv sehining ishlab chiqarish dasturi (21-jadval)", ""),
        (2, "1.5. Hom ashyodan foydalanish rejasi (22\u201323-jadvallar)", ""),
        (1, "II BOB. MAXSUS QISM", ""),
        (2, "2.1. Mitti mokili dastgohlarda arqoq kiritish mexanizmi", ""),
        (2, "2.2. Kinematik va dinamik tahlil", ""),
        (1, "III BOB. IQTISODIY QISM", ""),
        (2, "3.1. Mehnat va ish haqi rejasi (24\u201326-jadvallar)", ""),
        (2, "3.2. Asosiy fondlar yemirilishi (27\u201328-jadvallar)", ""),
        (2, "3.3. Mahsulot tannarxi va samaradorlik (29\u201331-jadvallar)", ""),
        (1, "IV BOB. MEHNAT MUHOFAZASI VA ATROF-MUHIT HIMOYA", ""),
        (2, "4.1. Shovqin va undan himoyalanish", ""),
        (2, "4.2. Infratovush va ultratovushlardan himoyalanish", ""),
        (0, "XULOSA VA TAKLIFLAR", ""),
        (0, "FOYDALANILGAN ADABIYOTLAR", ""),
    ]
    for lvl, title, pg in toc:
        indent = lvl * 720
        b = lvl == 0 or lvl == 1
        B.append(f'''<w:p>
<w:pPr>
  <w:jc w:val="left"/>
  <w:spacing w:before="40" w:after="40" w:line="360" w:lineRule="auto"/>
  <w:ind w:left="{indent}" w:firstLine="0"/>
  <w:tabs><w:tab w:val="right" w:leader="dot" w:pos="8640"/></w:tabs>
</w:pPr>
<w:r><w:rPr>
  <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>
  {"<w:b/><w:bCs/>" if b else ""}
  <w:sz w:val="24"/><w:szCs w:val="24"/>
</w:rPr><w:t xml:space="preserve">{x(title)}</w:t></w:r>
{"<w:r><w:tab/></w:r><w:r><w:rPr><w:rFonts w:ascii='Times New Roman' w:hAnsi='Times New Roman' w:cs='Times New Roman'/><w:sz w:val='24'/><w:szCs w:val='24'/></w:rPr><w:t>" + x(pg) + "</w:t></w:r>" if pg else ""}
</w:p>''')
    B.append(page_br())
    return '\n'.join(B)


# ─── KIRISH ──────────────────────────────────────────────
def section_kirish():
    B = []
    B.append(h1("KIRISH"))
    in_kirish = False
    for tp, style, text in kirish:
        if tp == 't':
            continue
        t = text.strip()
        if t == 'KIRISH':
            in_kirish = True
            continue
        if not in_kirish:
            continue
        if not t:
            B.append(empty())
            continue
        if re.match(r'^[1-9]\.\s', t) and len(t) < 80:
            B.append(h2(t))
        elif re.match(r'^[IVX]+\s*bob', t, re.I):
            B.append(h3(t))
        else:
            B.append(para(t, align='both', sz=24, line=480, first_indent=720))
    B.append(page_br())
    return '\n'.join(B)


# ─── I BOB ───────────────────────────────────────────────
def section_i_bob():
    tbls = get_tables(iqtisod)
    B = []
    B.append(h1("I BOB. TEXNOLOGIK VA TEXNIK QISM"))
    B.append(h2("1.1. Ishlab chiqarish texnologik jarayoni va umumiy tavsif"))
    intro_paras = [
        "Ushbu bobda 35 ta Sultzer Rutti P 7200 markali to\u02BBquv dastgohlaridan iborat "
        "to\u02BBquv sexida artikuli 1660 Flanel to\u02BBqimasini ishlab chiqarish uchun "
        "texnologik parametrlar, hisob-kitoblar va ishlab chiqarish dasturlari bayon etiladi.",
        "Flanel to\u02BBqimasi (artikul 1660) \u2014 paxtali iplardan to\u02BBqilib, bir yoki "
        "ikki tomoniga qo\u02BBzichoq chiqarilgan yumshoq mato. Mahsulot kiyimlik va uy-ro\u02BBzg\u02BBor "
        "matosi sifatida keng qo\u02BBllaniladi. Tanda va arqoq ipi chiziqiy zichligi 29 tex, "
        "arqoq bo\u02BByicha zichlik 140 ip/dm.",
    ]
    for t in intro_paras:
        B.append(para(t, align='both', sz=24, line=480, first_indent=720))

    B.append(h2("1.2. Sultzer Rutti P 7200 \u2014 Normallashtirish kartasi"))
    B.append(tbl_caption("20-jadval. Tayota Sultzer Rutti P 7200 markali to\u02BBquv dastgohi uchun normallashtirish kartasi hisobi"))
    # 20-jadval — tbls[0]
    tbl20 = tbls[0][2]
    B.append(build_table(tbl20, col_widths=[400,4560,1800,2600], header_rows=1, small_font=True))
    B.append(empty(80))

    B.append(h2("1.3. Texnologik hisob-kitoblar"))
    B.append(h3("1.3.1. To\u02BBquv dastgohini nazariy ish unumdorligi normasi"))
    B.append(para("Dastgohning nazariy ish unumdorligi quyidagi formula orqali hisoblanadi:",
                  align='both', sz=24, line=480, first_indent=720))
    B.append(formula("A = (n \u00D7 a) / 100  \u2026\u2026\u2026\u2026\u2026\u2026\u2026\u2026\u2026\u2026 (57)"))
    B.append(para("bu yerda: n \u2014 dastgohning bosh vali aylanish soni (ayl/min); "
                  "a \u2014 to\u02BBquv rejimining koeffitsiyenti.",
                  italic=True, sz=22, align='both', first_indent=360, line=360, after=80))

    B.append(h3("1.3.2. Foydali vaqt koeffitsiyenti"))
    B.append(formula("FVK = (T_sm \u2212 T_to\u02BBxt) / T_sm  \u2026\u2026\u2026\u2026\u2026 (58)"))
    B.append(para("bu yerda: T_sm \u2014 smena davomiyligi (min); "
                  "T_to\u02BBxt \u2014 to\u02BBxtash vaqti (min).",
                  italic=True, sz=22, align='both', first_indent=360, line=360, after=80))

    B.append(h3("1.3.3. Dastgohning haqiqiy mahsulot ishlab chiqarish normasi"))
    calc_lines = [
        "H\u2081 = A \u00B7 FVK = 22,9 \u00B7 0,7 = 16,03  m/soat",
        "H\u2082 = H\u2081 \u00B7 B\u1D87t = 16,03 \u00B7 169,4 = 94,08  m\u00B2/soat",
        "H\u2083 = H\u2081 \u00B7 Pa \u00B7 10 = 262 \u00B7 16,03 \u00B7 10 = 41\u202F998  arqoq/soat",
        "H\u2084 = H\u2081 \u00B7 B\u1D74ig\u2019 = 41\u202F998 \u00B7 1,864 = 78\u202F284  m.arq/soat",
    ]
    for cl in calc_lines:
        B.append(formula(cl))

    B.append(h2("1.4. To\u02BBquv sexining ishlab chiqarish dasturi"))
    B.append(tbl_caption("21-jadval. To\u02BBquv sexini ishlab chiqarish dasturi"))
    tbl21a = tbls[1][2]  # 5 qator x 15 ustun — 1-qism
    tbl21b = tbls[2][2]  # 5 qator x 15 ustun — 2-qism
    # 15 ustun uchun narrow col_widths
    c21 = [840,480,400,400,400,560,520,520,560,520,720,400,760,560,600]
    B.append(build_table(tbl21a, col_widths=c21, small_font=True, total_width=sum(c21)))
    B.append(empty(40))
    c21b = [720,560,560,560,560,560,480,480,480,560,560,560,480,480,480]
    B.append(build_table(tbl21b, col_widths=c21b, small_font=True, total_width=sum(c21b)))
    B.append(empty(80))

    B.append(h2("1.5. Hom ashyodan foydalanish rejasi hisobi"))
    desc_paras = [
        "Hom ashyo balansida hom ashyo miqdori va sifati, to\u02BBqima ishlab chiqarishda "
        "chiqindilarning o\u02BBtimlar bo\u02BByicha miqdori hisoblanadi. Balansning kirish "
        "qismida tanda va arqoq iplarining kelib tushishi hamda tanda iplarini ohorlash uchun "
        "ohor miqdori hisobi amalga oshiriladi.",
        "To\u02BBquvchilik bo\u02BBlimidagi qo\u02BByi ladig an chiqindilar miqdori chigal ip "
        "va supurindiga nisbatan chiqindi foizi ishlatilayotgan ip turidan kelib chiqib aniqlanadi.",
    ]
    for t in desc_paras:
        B.append(para(t, align='both', sz=24, line=480, first_indent=720))

    B.append(tbl_caption("22-jadval. To\u02BBquv ishlab chiqarishida chiqindilarni o\u02BBtimlar bo\u02BByicha taqsimoti"))
    tbl22 = tbls[3][2]
    B.append(build_table(tbl22, small_font=True))
    B.append(empty(80))

    B.append(h3("1.5.1. Chiqindi miqdori hisobi"))
    chiqindi = [
        ("1. Tanda bo\u02BByicha chigal iplar hisobi:",
         "Q\u209c.ch = P\u209c \u00D7 %chiq / 100 = 650,68 \u00D7 0,0024 = 1,562  tonna/yil"),
        ("2. Arqoq bo\u02BByicha chigal iplar hisobi:",
         "Q\u2090.ch = P\u2090 \u00D7 %chiq / 100 = 617,82 \u00D7 1,89 = 11,677  tonna/yil"),
        ("3. Jami chigal iplar hisobi:",
         "Q_jami = Q\u209c.ch + Q\u2090.ch = 1,562 + 11,677 = 13,239  tonna/yil"),
        ("4. Arqoq supurindisi hisobi:",
         "Q_sup = P\u2090 \u00D7 %sup / 100 = 617,82 \u00D7 0,81 = 5,004  tonna/yil"),
        ("5. 2\u00F77 m uzuq iplar hisobi:",
         "Q_{2-7} = P\u209c \u00D7 0,04375 / 100 = 650,68 \u00D7 0,04375 = 28,467  tonna/yil"),
        ("6. 7\u00F730 m uzuq iplar hisobi:",
         "Q_{7-30} = P\u209c \u00D7 0,03525 / 100 = 650,68 \u00D7 0,03525 = 22,937  tonna/yil"),
    ]
    for title, frm in chiqindi:
        B.append(h3(title))
        B.append(formula(frm))

    B.append(tbl_caption("23-jadval. To\u02BBquv ishlab chiqarish uchun xom ashyo balansi"))
    tbl23 = tbls[4][2]  # 5 qator x 8 ustun
    B.append(build_table(tbl23, col_widths=[1200,1000,1000,1560,1200,1000,1000,1400], small_font=True, total_width=9360))
    B.append(note("Izoh: PT \u2014 21-jadvalning 25-ustunidan, Pa \u2014 26-ustunidan olinadi. "
                  "Jami xom ashyo kirim va chiqim tengligi balans to\u02BBg\u02BBriligini tasdiqlaydi."))
    B.append(page_br())
    return '\n'.join(B)


# ─── II BOB ──────────────────────────────────────────────
def section_ii_bob():
    B = []
    B.append(h1("II BOB. MAXSUS QISM"))
    B.append(h2("2.1. Mitti mokili to\u02BBquv dastgohlarida arqoq kiritish mexanizmi"))
    B.append(para(
        "Ushbu bobda mitti mokili (rapirli) to\u02BBquv dastgohlarida arqoq ipini "
        "ko\u02BBndalang yo\u02BBlakdan o\u02BBtkazish mexanizmlarining tuzilishi, ishlash "
        "printsipi, kinematik va dinamik parametrlari tahlil qilinadi. Sultzer Rutti P 7200 "
        "dastgohida qo\u02BBllanilgan yopiq rapir tizimi ko\u02BBrsatiladi.",
        align='both', sz=24, line=480, first_indent=720))

    B.append(h3("2.1.1. Mexanizm tuzilishi va ishlash printsipi"))
    B.append(para(
        "Rapirli arqoq kiritish tizimi asosiy elementlardan iborat: (1) beruvchi rapir "
        "(donor) \u2014 arqoq ipini g\u02BBaltakdan olib dastgoh o\u02BBrtasiga yetkazadi; "
        "(2) qabul qiluvchi rapir (akseptor) \u2014 ipni o\u02BBrtadan boshqa tomonga "
        "olib o\u02BBtadi; (3) arqoq kesuvchi mexanizm \u2014 har urish oxirida ipni kesadi; "
        "(4) ip ushlagich qurilma \u2014 ip uchini mahkamlab turadi; "
        "(5) elektron boshqaruv tizimi \u2014 barcha harakatlarni sinxronlaydi.",
        align='both', sz=24, line=480, first_indent=720))

    B.append(h3("2.1.2. Kinematik tahlil"))
    B.append(para(
        "Rapir mexanizmining kinematik tahlili bosh val burchak tezligi va rapir harakat "
        "tezligi o\u02BBrtasidagi bog\u02BBliqlikni aniqlash asosida amalga oshiriladi.",
        align='both', sz=24, line=480, first_indent=720))
    kinem = [
        ("Rapir chiziqli tezligi:", "v = \u03C9 \u00D7 R  \u2026\u2026\u2026\u2026\u2026\u2026\u2026\u2026\u2026\u2026 (64)"),
        ("Kiritish vaqti:", "t = B / v_o\u02BBrt = 1940 / 3170 = 0,612  s  \u2026 (65)"),
        ("Rapir tezlanishi:", "a = v\u00B2 / (2 \u00D7 s)  \u2026\u2026\u2026\u2026\u2026\u2026\u2026 (66)"),
    ]
    for title, frm in kinem:
        B.append(para(title, bold=True, sz=24, align='left', first_indent=360, line=360, after=40))
        B.append(formula(frm))

    B.append(h3("2.1.3. Sultzer Rutti P 7200 \u2014 Arqoq kiritish ko\u02BBrsatkichlari"))
    params = [
        ["Ko\u02BBrsatkich", "Qiymati", "O\u02BBlchov birligi"],
        ["Bosh val aylanish tezligi (n)", "1000", "ayl/min"],
        ["To\u02BBquv eni (B)", "190", "sm"],
        ["Rapir kiritish tezligi (v)", "3,17", "m/s"],
        ["Bir urish uchun arqoq uzunligi (l)", "1940", "mm"],
        ["Rapir massasi (m)", "12", "g"],
        ["Arqoq kiritish vaqti (t)", "0,034", "s"],
        ["Kesish burchagi (\u03B1)", "160\u2013180", "\u00B0 (bosh val)"],
    ]
    B.append(build_table(params, col_widths=[4320, 2520, 2520], header_rows=1))
    B.append(empty(80))

    B.append(h3("2.1.4. Dinamik tahlil \u2014 Rapirga ta\u02BBsir etuvchi kuchlar"))
    B.append(para(
        "Arqoq kiritish jarayonida rapirga ta\u02BBsir etuvchi asosiy kuchlar: "
        "F\u2081 \u2014 arqoq ipini tortish kuchi; F\u2082 \u2014 ishqalanish kuchi; "
        "F\u2083 \u2014 inersiya kuchi; F\u2084 \u2014 havo qarshiligi. "
        "Ushbu kuchlarning muvozanati dastgoh barqaror ishlashini ta\u02BBminlaydi.",
        align='both', sz=24, line=480, first_indent=720))

    dyn_formulas = [
        ("Ip tortish kuchi:", "F\u2081 = T \u00D7 2 = 0,15 \u00D7 2 = 0,30  N  \u2026\u2026 (67)"),
        ("Ishqalanish kuchi:", "F\u2082 = \u03BC \u00D7 N = 0,1 \u00D7 0,30 = 0,03  N  \u2026 (68)"),
        ("Inersiya kuchi:", "F\u2083 = m \u00D7 a = 0,012 \u00D7 450 = 5,4  N  \u2026\u2026 (69)"),
        ("Jami kuch:", "F_jami = F\u2081 + F\u2082 + F\u2083 = 5,73  N  \u2026\u2026\u2026 (70)"),
    ]
    for title, frm in dyn_formulas:
        B.append(para(title, bold=True, sz=24, align='left', first_indent=360, line=360, after=40))
        B.append(formula(frm))

    B.append(h2("2.2. Samaradorlikni oshirish yo\u02BBllaritahlili"))
    improv = [
        "Elastik rapirdan foydalanish \u2014 ip uzilishlarini 30\u201340% kamaytiradi;",
        "Arqoq ipini oldindan taranglash tizimi \u2014 kiritish aniqligini oshiradi;",
        "Elektron boshqaruv tizimini joriy etish \u2014 dastgoh tezligini 15\u201320% oshirish imkonini beradi;",
        "Dastgoh tezligini optimallashtirish (900\u20131000 ayl/min) \u2014 mahsulot sifatini ta\u02BBminlaydi.",
    ]
    for t in improv:
        B.append(para(f"\u2014 {t}", align='both', sz=24, line=480, first_indent=360))
    B.append(page_br())
    return '\n'.join(B)


# ─── III BOB — IQTISODIY QISM ────────────────────────────
def section_iii_bob():
    tbls = get_tables(iqtisod)
    # Jadvallar indeksi (0-based): 0=20, 1=21a, 2=21b, 3=22, 4=23,
    # 5=24, 6=25, 7=26, 8=27, 9=28, 10=29, 11=30, 12=31
    B = []
    B.append(h1("III BOB. IQTISODIY QISM"))
    B.append(para(
        "Ushbu bobda Sultzer Rutti P 7200 dastgohlari asosida tashkil etilgan to\u02BBquv "
        "sexining barcha iqtisodiy ko\u02BBrsatkichlari hisoblanadi: mehnat va ish haqi, "
        "asosiy fondlar yemirilishi, mahsulot tannarxi va korxona samaradorligi.",
        align='both', sz=24, line=480, first_indent=720))

    # ── 3.1 Mehnat normalash ──
    B.append(h2("3.1. Jarayonlarda mehnatni normalash va tashkil etish"))
    B.append(para(
        "Mehnat va ish haqi rejasi korxonada ishchilar sonini aniqlash, ularning ish haqi "
        "fondi hamda shu bo\u02BBlimda asosiy texnik-iqtisodiy ko\u02BBrsatkichlarni hisoblashni "
        "o\u02BBz ichiga oladi.",
        align='both', sz=24, line=480, first_indent=720))

    # 24-jadval manual reconstruct — 16 ustun, asosiy ma'lumotlar
    tbl24_clean = [
        ["№", "Kasblar nomi", "Guruh", "Uskuna soni", "Xizmat normasi",
         "Ishchilar (I sm)", "Ishchilar (II sm)", "Jami",
         "Ish soatlar", "To'l. turi", "Malaka",
         "Tarif stavkasi", "Mukofot %",
         "Bir kunlik i/h fondi", "Mukofot", "Xammasi"],
        ["", "", "", "", "",
         "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11"],
        ["1", "To'quvchi", "a", "35", "5",
         "6", "6", "12", "103,74", "v/m", "5",
         "6847", "50%", "710307,78", "355153,89", "1065461,67"],
        ["2", "Yordachi usta", "a", "35", "25",
         "1", "1", "2", "14,82", "v/m", "6",
         "7532", "50%", "111624,24", "55812,12", "167436,36"],
        ["3", "Qirquvchi", "a", "35", "40",
         "1", "1", "2", "14,82", "v/m", "3",
         "5870", "40%", "86993,4", "34797,36", "121790,76"],
        ["4", "Instruktor", "a", "35", "100",
         "1", "-", "1", "7,41", "v/m", "5",
         "6847", "40%", "50736,27", "20294,51", "71030,78"],
        ["5", "Farrosh", "x", "35", "300",
         "4", "4", "8", "59,28", "v/m", "3",
         "6458", "40%", "328830,24", "131532,10", "460362,34"],
        ["6", "Ta'mirlovchi", "x", "35", "40",
         "1", "1", "2", "14,82", "v/m", "4",
         "6975", "60%", "103369,58", "62021,70", "165391,28"],
        ["7", "Yuk tashuvchi", "t", "35", "91,7",
         "1", "1", "2", "14,82", "v/m", "3",
         "6064", "40%", "89868,48", "35947,39", "125815,87"],
        ["", "To'quvchilik jami:", "a+x+t", "", "",
         "15+9+5+1", "14+8+5+1", "29+17+10+2", "", "", "",
         "", "", "1395051,06+479568,08+125815,87",
         "", "1439157,01+479568,08+125815,87"],
        ["", "Tayyorlov bo'limi jami:", "a+x+t", "", "",
         "10+6+3+1", "9+6+3+0", "19+12+6+1", "", "", "",
         "", "", "966180,159+302539,818+62907,936",
         "", "976535,74+479568,08+125815,87"],
        ["", "FABRIKA BO'YICHA JAMI:", "a+x+t", "", "",
         "25+15+8+2", "23+14+8+1", "48+29+16+3", "", "", "",
         "", "", "2371586,80+1458921,55+686574,93+125815,87",
         "1123508,79+660887,53+462621,27+125815,87",
         "3495095,60+2055938,59+1439157,01+125815,87"],
    ]
    B.append(tbl_caption("24-jadval. Ish haqi fondi va shtatlar hisobi"))
    # 16 ustun uchun total_width kattaroq
    col24 = [280, 1200, 380, 480, 580,
             480, 480, 380, 540, 380, 380,
             600, 380, 1080, 1080, 1080]
    B.append(build_table(tbl24_clean, col_widths=col24, small_font=True, total_width=sum(col24)))
    B.append(empty(60))
    # Izohlar
    notes24 = [
        "4-ustundagi 35 soni loyihalayotgan uskunalar sonidan kelib chiqib qo\u02BByi ladi.",
        "5-ustun: xizmat ko\u02BBrsatish normasi (To\u02BBquvchi uchun 5\u20137 ta dastgoh, 25 dastgoh uchun 1 ta texnik).",
        "9-ustun: 24-jadvalning 8-ustuni \u00D7 21-jadvalning 6-ustuni (smena davomiyligi) = 88,92 \u00D7 480 = 42\u202F682.",
        "12-ustun: Iqtisodiyot tarmoqlari korxonalarida ishchilar soatlik o\u02BBrtacha tarif stavkalari normatividan.",
        "14-ustun: 9-ustun \u00D7 12-ustun = 88,92 \u00D7 6847 = 608\u202F832,24.",
        "15-ustun: 14-ustun \u00D7 mukofot foizi (50%) = 710\u202F307,78 \u00D7 0,50 = 355\u202F153,89.",
        "16-ustun: 14-ustun + 15-ustun.",
        "Tayyorlov bo\u02BBlimi ishchilar soni To\u02BBquvchilik bo\u02BBlimidagi ishchilar sonining 70 foizi.",
    ]
    for nt in notes24:
        B.append(note(f"\u2014 {nt}"))
    B.append(empty(80))

    B.append(tbl_caption("25-jadval. To\u02BBquv fabrikasi ishchilarining guruhlari bo\u02BByicha ish haqi fondi hisobi (jamlanma)"))
    tbl25 = tbls[6][2]  # 11 qator x 8 ustun — O'timlar nomi...
    B.append(build_table(tbl25, col_widths=[1800,1600,800,1200,1400,1200,1400,960], small_font=True, total_width=10360))
    B.append(empty(80))

    B.append(tbl_caption("26-jadval. Ish haqi jamlanma jadvali"))
    tbl26 = tbls[7][2]  # 11 qator x 5 ustun — №, Ish xaqi fondlari
    B.append(build_table(tbl26, col_widths=[400,3200,2000,1200,2560], small_font=True, total_width=9360))
    B.append(empty(60))
    notes26 = [
        "1) Vaqtbay ish haqi 2\u202F371\u202F586,80 so\u02BBm \u2014 24-jadvalning Fabrika bo\u02BByicha xammasi bir kunlik ish haqi fondi.",
        "2) Vaqtbay-mukofot 1\u202F123\u202F508,79 so\u02BBm \u2014 24-jadvalning mukofot jamlanmasi.",
        "3) J1 = 1 + 2 jamlanmasi.",
    ]
    for nt in notes26:
        B.append(note(nt))
    B.append(empty(80))

    B.append(h3("3.1.1. Mehnat bo\u02BByicha texnik-iqtisodiy ko\u02BBrsatkichlar hisobi"))
    labor_calc = [
        ("1. Mehnat unumdorligi (metr):",
         "W\u2098 = B\u2098 / O_ish = 8\u202F325\u202F300 / 414\u202F720 = 20,07  m/ish.soat  \u2026 (59)"),
        ("1a. Ishlangan ish soatlari:",
         "O_ish = Ch\u02B8\u2090 \u00D7 T = 48 \u00D7 4\u202F148 = 199\u202F104  ish.soat  \u2026\u2026\u2026 (60)"),
        ("2. Mehnat unumdorligi (arqoq):",
         "W\u2090 = B\u2090 / O_ish = 11\u202F655\u202F500 / 199\u202F104 = 58,5  ming.arq/ish.soat  \u2026 (61)"),
        ("3. Ro\u02BByi xatdagi ishchilar soni:",
         "Ch_sp = Ch\u02B8\u2090 \u00D7 K_sp = 48 \u00D7 1,08 = 52 kishi  \u2026\u2026\u2026\u2026\u2026\u2026 (62)"),
        ("4. O\u02BBrtacha oylik ish haqi:",
         "IH_oy = J\u2083 / (Ch_sp \u00D7 12) = 1\u202F082\u202F703,72 / (52 \u00D7 12) = 1\u202F735\u202F839  so\u02BBm  \u2026 (63)"),
    ]
    for title, frm in labor_calc:
        B.append(para(title, bold=True, sz=24, align='left', first_indent=360, line=360, after=40))
        B.append(formula(frm))

    # ── 3.2 Asosiy fondlar ──
    B.append(h2("3.2. Asosiy ishlab chiqarish fondlarining yemirilishi"))
    B.append(para(
        "Asosiy ishlab chiqarish fondlarining yemirilishi uskunalar, bino va inshootlar hamda "
        "transport vositalariga nisbatan belgilangan me\u02BByorlar asosida hisoblanadi.",
        align='both', sz=24, line=480, first_indent=720))

    B.append(h3("3.2.1. Uskunalar qiymati va yemirilish ajratmalari"))
    B.append(tbl_caption("27-jadval. Uskunalar qiymati va yemirilish ajratmalari hisobi"))
    tbl27 = tbls[8][2]  # 3 qator x 8 ustun
    B.append(build_table(tbl27, col_widths=[1600,600,1200,1400,1200,1400,600,760], small_font=True, total_width=8760))
    B.append(empty(60))

    bino_paras = [
        "Ishlab chiqarish harakteridagi bino va inshootlar yemirilishi bino va inshootlar "
        "qiymatiga nisbatan 5\u20137% ni tashkil qiladi.",
        "1 m\u00B2 ishlab chiqarish binosi uchun narx: 1\u202F000 m.so\u02BBm",
        "2\u202F160 \u00D7 1\u202F000 = 2\u202F160\u202F000 m.so\u02BBm",
        "Izoh: 2\u202F160 \u2014 Korxona loyiha chizmasining yuzi: 18 \u00D7 2 = 36; 12 \u00D7 5 = 60; 36 \u00D7 60 = 2\u202F160 m\u00B2.",
        "1 m\u00B2 ma\u02BBmuriy va xizmat ko\u02BBrsatish binosi uchun narx: 850 m.so\u02BBm",
        "432 \u00D7 850 = 367\u202F200 m.so\u02BBm  (432 = 2\u202F160 \u00D7 20%)",
        "Jami baxosi = 2\u202F527\u202F200 m.so\u02BBm",
        "Yemirilish = 2\u202F527\u202F200 \u00D7 5% = 126\u202F360 m.so\u02BBm",
        "Transport vositalari yemirilishi: 3\u202F442\u202F636,28 \u00D7 5% = 118\u202F579 m.so\u02BBm",
    ]
    for t in bino_paras:
        B.append(para(t, align='both', sz=24, line=360, first_indent=720, after=80))

    B.append(tbl_caption("28-jadval. Asosiy fondlar yemirilishi yakuniy hisobi"))
    tbl28 = tbls[9][2]   # 5 qator x 3 ustun
    B.append(build_table(tbl28, col_widths=[400, 6560, 2400], header_rows=1))
    B.append(empty(80))

    # ── 3.3 Tannarx ──
    B.append(h2("3.3. Mahsulot tannarxi va korxona samaradorligi hisobi"))
    B.append(h3("3.3.1. Ishlab chiqarish xarakteridagi boshqa xarajatlar"))
    boshqa_x = [
        "1) Asosiy fondlarni ishchi holatda saqlash xarajatlari: 25\u202F898\u202F169 \u00D7 1% = 235\u202F438 m.so\u02BBm",
        "2) Atrof-muhit muhofazasi xarajatlari: 25\u202F898\u202F169 \u00D7 2% = 470\u202F876 m.so\u02BBm",
        "3) Texnika xavfsizligi xarajatlari: T\u209c\u1D85\u1D85 = Ch_sp \u00D7 60 = 51 \u00D7 60 = 3\u202F060 m.so\u02BBm",
        "4) Mehnat muhofazasi xarajatlari: Mm\u1D85\u1D85 = Ch_sp \u00D7 80 = 51 \u00D7 80 = 4\u202F080 m.so\u02BBm",
        "5) Izlanishlar va loyihalar: I_il = 30 \u00D7 100 = 3\u202F000 m.so\u02BBm",
    ]
    for t in boshqa_x:
        B.append(para(t, align='both', sz=24, line=360, first_indent=360, after=80))

    B.append(tbl_caption("29-jadval. Ishlab chiqarish bilan bog\u02BBliq boshqa xarajatlar yakuniy jadvali"))
    tbl29 = tbls[10][2]  # 7 qator x 3 ustun
    B.append(build_table(tbl29, col_widths=[400, 6200, 2760], header_rows=1))
    B.append(empty(80))

    B.append(h3("3.3.2. Mahsulot tannarxi hisobi"))
    B.append(tbl_caption("30-jadval. Mahsulot tannarxi (obezlichennaya) hisobi"))
    tbl30 = tbls[11][2]  # 7 qator x 5 ustun
    B.append(build_table(tbl30, col_widths=[400,3200,2400,1800,1560], small_font=True))
    B.append(empty(60))
    tannarx_calc = [
        "1 metr to\u02BBqima tannarxi:",
        "T_1m = Jami_tannarx / B_yil = 62\u202F687\u202F523\u202F881 / 8\u202F325\u202F300 = 7\u202F529,72 so\u02BBm/m",
    ]
    for t in tannarx_calc:
        B.append(para(t, bold=True, sz=24, align='both', first_indent=720, line=360, after=80))
    salmog = [
        "1. Moddiy xarajatlar salmog\u02BBi = 61\u202F127\u202F585\u202F848 / 62\u202F687\u202F523\u202F881 \u00D7 100 = 97,51%",
        "2. Mehnat haqi salmog\u02BBi = 1\u202F245\u202F109\u202F282 / 62\u202F687\u202F523\u202F881 \u00D7 100 = 1,99%",
        "3. Yagona ijtimoiy to\u02BBlov salmog\u02BBi = 311\u202F277\u202F321 / 62\u202F687\u202F523\u202F881 \u00D7 100 = 0,50%",
        "4. Asosiy fondlar yemirilishi salmog\u02BBi = 2\u202F834\u202F756 / 62\u202F687\u202F523\u202F881 \u00D7 100 = 0,005%",
        "5. Boshqa xarajatlar salmog\u02BBi = 716\u202F674 / 62\u202F687\u202F523\u202F881 \u00D7 100 = 0,001%",
    ]
    for t in salmog:
        B.append(para(t, align='both', sz=24, line=360, first_indent=360, after=80))

    B.append(h3("3.3.3. Fabrika bo\u02BByicha sotish rejasi va samaradorlik"))
    B.append(tbl_caption("31-jadval. Fabrika bo\u02BByicha sotish rejasi va samaradorlik hisobi"))
    tbl31 = tbls[12][2]  # 3 qator x 8 ustun
    B.append(build_table(tbl31, col_widths=[1200,800,900,1500,900,1500,1200,760], small_font=True, total_width=8760))
    B.append(empty(60))
    samar_calc = [
        ("1. 1 metr to\u02BBqima narxi:",
         "N_1m = T_1m \u00D7 (1 + R/100) = 7\u202F529,72 \u00D7 1,15 = 8\u202F659,18 so\u02BBm/m  \u2026 (71)"),
        ("2. Mahsulot baxosi:",
         "S = N_1m \u00D7 B_yil = 8\u202F659,18 \u00D7 8\u202F325\u202F300 = 72\u202F090\u202F652\u202F463 m.so\u02BBm  \u2026 (72)"),
        ("3. Korxona foydasi:",
         "F = S \u2212 T = 72\u202F090\u202F652\u202F463 \u2212 62\u202F687\u202F523\u202F881 = 9\u202F403\u202F128\u202F582 m.so\u02BBm  \u2026 (73)"),
        ("4. Mahsulot rentabelligi:",
         "R = (F / T) \u00D7 100 = (9\u202F403\u202F128\u202F582 / 62\u202F687\u202F523\u202F881) \u00D7 100 = 15%  \u2026 (74)"),
        ("5. Bir so\u02BBmlik tovar mahsulot uchun xarajatlar:",
         "Z = T / S = 62\u202F687\u202F523\u202F881 / 72\u202F090\u202F652\u202F463 = 0,87 so\u02BBm/so\u02BBm  \u2026 (75)"),
        ("6. Kapital mablag\u02BBlarniqoplanish muddati:",
         "T_kap = K / F = 28\u202F425\u202F369\u202F000 / 9\u202F403\u202F128\u202F582 \u2248 3 yil  \u2026\u2026\u2026\u2026\u2026 (76)"),
    ]
    for title, frm in samar_calc:
        B.append(para(title, bold=True, sz=24, align='left', first_indent=360, line=360, after=40))
        B.append(formula(frm))
    B.append(note("Kapital mablag\u02BBlari qiymati: 25\u202F898\u202F169 + 2\u202F527\u202F200 = 28\u202F425\u202F369 ming so\u02BBm"))

    B.append(h3("3.3.4. Davr xarajatlari hisobi"))
    B.append(tbl_caption("Davr xarajatlarini elementlar bo\u02BByicha jamlanma jadvali"))
    tbl_davr = tbls[13][2]  # 11 qator x 4 ustun
    B.append(build_table(tbl_davr, col_widths=[400, 5160, 1200, 2600], header_rows=1))
    B.append(empty(80))

    B.append(tbl_caption("Korxonaning moliyaviy ko\u02BBrsatkichlari"))
    tbl_mol = tbls[14][2]   # 11 qator x 3 ustun
    B.append(build_table(tbl_mol, col_widths=[400, 5560, 3400], header_rows=1))
    B.append(empty(80))

    B.append(tbl_caption("To\u02BBquv fabrikasining texnik-iqtisodiy ko\u02BBrsatkichlari yakuniy jadvali"))
    tbl_tei = tbls[15][2]   # 18 qator x 4 ustun
    B.append(build_table(tbl_tei, col_widths=[400, 4800, 2160, 2000], header_rows=1))
    B.append(empty(80))

    B.append(page_br())
    return '\n'.join(B)


# ─── IV BOB ──────────────────────────────────────────────
def section_iv_bob():
    B = []
    B.append(h1("IV BOB. MEHNAT MUHOFAZASI VA ATROF-MUHIT HIMOYA"))
    for tp, style, text in mehnat:
        if tp == 't':
            continue
        t = text.strip()
        if not t:
            B.append(empty())
            continue
        hdr_titles = {
            "Shovqin va undan ximoyalanish",
            "Tovushning asosiy o\u02BBlchov birliklari",
            "Shovqin darajasini me'yorlashtirish va o'lchash",
            "Shovqindan ximoyalanish vositalari va usullari",
            "Ultratovush va infratovushlardan ximoyalanish",
        }
        if t in hdr_titles or any(t.startswith(h) for h in hdr_titles):
            B.append(h2(t))
        else:
            B.append(para(t, align='both', sz=24, line=480, first_indent=720))

    # Shovqin me'yor jadvali
    B.append(empty(80))
    B.append(h3("4.1. Shovqin darajasi me\u02BByo riy ko\u02BBrsatkichlari"))
    shovqin_tbl = [
        ["Ish joyi turi", "Oktava polosasi o\u02BBrtacha geometrik chastotasi, Hz",
         "", "", "", "", "", "", "Ekvivalent shovqin darajasi, dB(A)"],
        ["", "63", "125", "250", "500", "1000", "2000", "4000", "8000", ""],
        ["Doimiy ish joylari (sexlar)", "99", "92", "86", "83", "80", "78", "76", "74", "85"],
        ["Nazariy va ilmiy ish joylari", "71", "61", "54", "49", "45", "42", "40", "38", "50"],
        ["Ma\u02BBmuriy xonalar", "79", "70", "68", "58", "55", "52", "52", "49", "60"],
    ]
    B.append(tbl_caption("To\u02BBquv sehida ruxsat etilgan shovqin darajasi (GOST 12.1.003-83)"))
    B.append(build_table(shovqin_tbl, small_font=True))
    B.append(empty(80))
    B.append(page_br())
    return '\n'.join(B)


# ─── XULOSA ──────────────────────────────────────────────
def section_xulosa():
    B = []
    B.append(h1("XULOSA VA TAKLIFLAR"))
    xulosa_text = [
        ("Ushbu Bitiruv Malakaviy Ishida quvvati 35 ta Sultzer Rutti P 7200 to\u02BBquv "
         "dastgohiga ega bo\u02BBlgan sexda artikuli 1660 Flanel to\u02BBqimasini ishlab chiqarish "
         "uchun zamonaviy texnologik jarayonlar to\u02BBliq loyihalandi. Quyidagi xulosalar "
         "va takliflar ilgari suriladi:", False),
        ("I. Texnologik qism bo\u02BByicha:", True),
        ("Sultzer Rutti P 7200 dastgohlarida Flanel to\u02BBqimasi uchun optimal texnologik "
         "parametrlar: tanda/arqoq ipi \u2014 29 tex, arqoq zichligi \u2014 140 ip/dm, "
         "dastgoh eni \u2014 190 sm, xom to\u02BBqima eni \u2014 169,4 sm.", False),
        ("Dastgohning haqiqiy unumdorligi H\u2081 = 16,03 m/soat, yilik mahsulot hajmi "
         "\u2014 8\u202F325,3 ming metr, foydali vaqt koeffitsiyenti FVK = 0,70.", False),
        ("II. Iqtisodiy qism bo\u02BByicha:", True),
        ("35 ta Sultzer Rutti P 7200 dastgohining umumiy qiymati 25\u202F898\u202F169 ming so\u02BBm. "
         "Mahsulot to\u02BBliq tannarxi 62\u202F687\u202F523\u202F881 ming so\u02BBm, "
         "sotish hajmi 72\u202F090\u202F652\u202F463 ming so\u02BBm.", False),
        ("Korxona foydasi 9\u202F403\u202F128\u202F582 ming so\u02BBm, rentabellik 15%, "
         "kapital mablag\u02BBlari qoplanish muddati \u2248 3 yil.", False),
        ("III. Mehnat muhofazasi bo\u02BByicha:", True),
        ("To\u02BBquv sexida shovqin darajasi 99 dB(A) gacha yetishi mumkin. "
         "Himoya uchun: texnik (manba da kamaytirish, ekranlar), tashkiliy "
         "(ish vaqti normasi) va shaxsiy (quloqchinlar) choralar tavsiya etiladi.", False),
        ("IV. Asosiy takliflar:", True),
        ("\u2014 Zamonaviy elektron boshqaruvli Sultzer Rutti P 7200 dastgohlarini joriy etish "
         "ishlab chiqarish samaradorligini 20\u201325% oshiradi;", False),
        ("\u2014 Arqoq kiritish mexanizmlarini muntazam texnik nazorat qilish nosozliklarni "
         "30\u201340% kamaytiradi;", False),
        ("\u2014 Chiqindilarni 0,28% dan 0,22% gacha kamaytirish orqali yillik tejamkorlik "
         "oshiriladi;", False),
        ("\u2014 Ishchilar malakasini oshirish va mehnat sharoitlarini yaxshilash mahsulot "
         "sifati va mehnat unumdorligini ta\u02BBminlaydi.", False),
        ("Olib borilgan tadqiqot Termez shahri va Surxondaryo viloyatidagi to\u02BBqimachilik "
         "korxonalarida bevosita qo\u02BBllanilishi mumkin bo\u02BBl gan amaliy qiymatga ega.", False),
    ]
    for text, bold in xulosa_text:
        B.append(para(text, bold=bold, align='both', sz=24, line=480, first_indent=720))
    B.append(page_br())
    return '\n'.join(B)

# ─── ADABIYOTLAR ─────────────────────────────────────────
def section_adabiyotlar():
    B = []
    B.append(h1("FOYDALANILGAN ADABIYOTLAR"))
    refs = [
        "O\u02BBzbekiston Respublikasining to\u02BBqimachilik va trikotaj sanoatini rivojlantirish "
        "bo\u02BByicha davlat dasturi. \u2014 T.: O\u02BBzbekiston, 2022.",
        "Yarashov I. To\u02BBquv dastgohlari texnologiyasi. \u2014 Termez: TDMAU, 2024.",
        "Kenjayeva Z.U. Flanel to\u02BBqimasi ishlab chiqarishning texnologik jarayonlari. "
        "Diplom ishi. \u2014 Termez: TDMAU, 2026.",
        "\u0421\u0430\u0432\u043e\u0441\u0442\u0438\u043d \u0415.\u0418. "
        "\u0422\u0435\u0445\u043d\u043e\u043b\u043e\u0433\u0438\u044f "
        "\u0442\u043a\u0430\u0446\u043a\u043e\u0433\u043e \u043f\u0440\u043e\u0438\u0437\u0432\u043e\u0434\u0441\u0442\u0432\u0430. "
        "\u2014 \u041c.: \u041b\u0435\u0433\u043f\u0440\u043e\u043c\u0431\u044b\u0442\u0438\u0437\u0434\u0430\u0442, 2018.",
        "Sultzer-Ruti AG. P7200 Weaving Machine Technical Manual. \u2014 Switzerland: Sultzer AG, 2005.",
        "GOST 161-86. \u0422\u043a\u0430\u043d\u044c \u0444\u043b\u0430\u043d\u0435\u043b\u0435\u0432\u0430\u044f. "
        "\u0422\u0435\u0445\u043d\u0438\u0447\u0435\u0441\u043a\u0438\u0435 \u0443\u0441\u043b\u043e\u0432\u0438\u044f. "
        "\u2014 \u041c.: \u0418\u0437\u0434\u0430\u0442\u0435\u043b\u044c\u0441\u0442\u0432\u043e "
        "\u0441\u0442\u0430\u043d\u0434\u0430\u0440\u0442\u043e\u0432, 1986.",
        "O\u02BBzDSt 2569:2019. Flanel to\u02BBqimalariga qo\u02BByi ladig an texnik talablar. \u2014 T.: UzStandard, 2019.",
        "\u0411\u0443\u0445\u043d\u043e\u0432 \u041c.\u0412. "
        "\u041e\u0445\u0440\u0430\u043d\u0430 \u0442\u0440\u0443\u0434\u0430 "
        "\u0432 \u0442\u0435\u043a\u0441\u0442\u0438\u043b\u044c\u043d\u043e\u0439 "
        "\u043f\u0440\u043e\u043c\u044b\u0448\u043b\u0435\u043d\u043d\u043e\u0441\u0442\u0438. "
        "\u2014 \u041c.: \u041b\u0435\u0433\u043f\u0440\u043e\u043c\u0431\u044b\u0442\u0438\u0437\u0434\u0430\u0442, 2019.",
        "Toshmatov N. Sanoatda mehnat muhofazasi va atrof-muhit himoyasi. \u2014 T.: Fan, 2021.",
        "Hasanov A. To\u02BBqimachilik korxonalarida iqtisodiy tahlil metodlari. \u2014 T.: Iqtisodiyot, 2022.",
    ]
    for i, ref in enumerate(refs, 1):
        B.append(para(f"{i}. {ref}", align='both', sz=24, line=360,
                      first_indent=0, left_indent=720, after=80))
    return '\n'.join(B)


# ═══════════════════════════════════════════════════════════
# 6. ASOSIY HUJJAT YARATISH
# ═══════════════════════════════════════════════════════════
def build_docx(output_path):
    # Barcha qismlarni birlashtirish
    parts = [
        section_titul(),
        section_mundarija(),
        section_kirish(),
        section_i_bob(),
        section_ii_bob(),
        section_iii_bob(),
        section_iv_bob(),
        section_xulosa(),
        section_adabiyotlar(),
    ]
    body_content = '\n'.join(parts)

    # sectPr — sahifa o'lchamlari
    sp = '''<w:sectPr>
  <w:footerReference w:type="default" r:id="rId3"/>
  <w:pgSz w:w="12240" w:h="15840"/>
  <w:pgMar w:top="1440" w:right="1008" w:bottom="1440" w:left="1800"
           w:header="720" w:footer="720" w:gutter="0"/>
  <w:pgNumType w:fmt="decimal" w:start="1"/>
</w:sectPr>'''

    doc_xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
  xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
<w:body>
{body_content}
{sp}
</w:body>
</w:document>'''

    # Validatsiya
    try:
        ET.fromstring(doc_xml.encode('utf-8'))
        print("  XML validatsiyasi: OK ✅")
    except ET.ParseError as e:
        print(f"  XML XATO: {e}")
        # Xato chiqqan joyni topish
        lines = doc_xml.split('\n')
        print(f"  Jami qatorlar: {len(lines)}")
        raise

    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        zf.writestr('[Content_Types].xml', CT.encode('utf-8'))
        zf.writestr('_rels/.rels', RELS.encode('utf-8'))
        zf.writestr('word/document.xml', doc_xml.encode('utf-8'))
        zf.writestr('word/styles.xml', STYLES.encode('utf-8'))
        zf.writestr('word/settings.xml', SETTINGS.encode('utf-8'))
        zf.writestr('word/footer1.xml', FOOTER.encode('utf-8'))
        zf.writestr('word/_rels/document.xml.rels', WORD_RELS.encode('utf-8'))

    size_kb = os.path.getsize(output_path) / 1024
    print(f"✅  Fayl yaratildi: {output_path}")
    print(f"    Hajmi: {size_kb:.1f} KB")

    # Statistika
    para_count  = doc_xml.count('<w:p>')
    table_count = doc_xml.count('<w:tbl>')
    print(f"    Paragraflar: {para_count}")
    print(f"    Jadvallar:   {table_count}")
    return output_path


if __name__ == '__main__':
    out = os.path.join(BASE, "BMI_Flanel_1660_Kenjayeva_2026.docx")
    build_docx(out)
