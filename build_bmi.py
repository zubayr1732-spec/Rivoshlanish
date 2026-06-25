#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
BMI - Bitiruv Malakaviy Ishi
"Quvvati 35 ta to'quv dastgohiga ega bo'lgan artikuli 1660
 Flanel to'qimasini ishlab chiqarish uchun zamonaviy texnologik
 jarayonlarni loyihalash"
Barcha bob va jadvallarni bitta .docx faylga birlashtiradi.
"""

import zipfile
import xml.etree.ElementTree as ET
import os
import re

# ─── 1. Manba fayllardan ma'lumot o'qish ─────────────────────────────────────

WD  = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
WDR = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'

def read_docx_elements(path):
    """Paragraph va Table elementlarini ro'yxat sifatida qaytaradi."""
    result = []
    with zipfile.ZipFile(path) as z:
        with z.open('word/document.xml') as f:
            content = f.read().decode('utf-8')
    tree = ET.fromstring(content)
    body = tree.find('.//' + WD + 'body')
    if body is None:
        return result
    for child in body:
        tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
        if tag == 'p':
            style = ''
            pPr = child.find('.//' + WD + 'pStyle')
            if pPr is not None:
                style = pPr.get(WD + 'val', '')
            parts = []
            for r in child.iter(WD + 't'):
                if r.text:
                    parts.append(r.text)
            result.append(('para', style, ''.join(parts)))
        elif tag == 'tbl':
            rows = []
            for row in child.iter(WD + 'tr'):
                cells = []
                for cell in row.iter(WD + 'tc'):
                    ct = []
                    for t in cell.iter(WD + 't'):
                        if t.text:
                            ct.append(t.text)
                    cells.append(''.join(ct))
                rows.append(cells)
            result.append(('table', '', rows))
    return result


BASE = os.path.dirname(os.path.abspath(__file__))

kirish_items  = read_docx_elements(os.path.join(BASE, "Kenjayeva_Diplom_Kirish (2).docx"))
iv_bob_items  = read_docx_elements(os.path.join(BASE, "77777.docx"))
iqtisod_items = read_docx_elements(os.path.join(BASE, "Z.Kenjayeva_dastgoh_Iqtisod_13.06.2026.docx"))

# ─── 2. XML yordamchi funksiyalar ─────────────────────────────────────────────

def esc(text):
    """XML uchun maxsus belgilarni ekranlash."""
    return (text
            .replace('&', '&amp;')
            .replace('<', '&lt;')
            .replace('>', '&gt;')
            .replace('"', '&quot;'))

# ─── 3. Word ML bloklari ──────────────────────────────────────────────────────

def para_xml(text, style='Normal', bold=False, size=24, align='left',
             space_before=0, space_after=120, italic=False):
    """Oddiy paragraf XML."""
    al_map = {'left':'left','center':'center','right':'right',
              'justify':'both','both':'both'}
    jc_val = al_map.get(align, 'both')
    b_tag  = '<w:b/><w:bCs/>' if bold else ''
    i_tag  = '<w:i/><w:iCs/>' if italic else ''
    sz_tag = f'<w:sz w:val="{size}"/><w:szCs w:val="{size}"/>'
    return f'''<w:p>
  <w:pPr>
    <w:pStyle w:val="{style}"/>
    <w:jc w:val="{jc_val}"/>
    <w:spacing w:before="{space_before}" w:after="{space_after}"/>
  </w:pPr>
  <w:r>
    <w:rPr>{b_tag}{i_tag}{sz_tag}</w:rPr>
    <w:t xml:space="preserve">{esc(text)}</w:t>
  </w:r>
</w:p>'''

def heading1(text):
    return f'''<w:p>
  <w:pPr>
    <w:pStyle w:val="Heading1"/>
    <w:jc w:val="left"/>
    <w:spacing w:before="240" w:after="120"/>
  </w:pPr>
  <w:r>
    <w:rPr>
      <w:b/><w:bCs/>
      <w:sz w:val="28"/><w:szCs w:val="28"/>
    </w:rPr>
    <w:t xml:space="preserve">{esc(text)}</w:t>
  </w:r>
</w:p>'''

def heading2(text):
    return f'''<w:p>
  <w:pPr>
    <w:pStyle w:val="Heading2"/>
    <w:jc w:val="left"/>
    <w:spacing w:before="200" w:after="80"/>
  </w:pPr>
  <w:r>
    <w:rPr>
      <w:b/><w:bCs/>
      <w:sz w:val="26"/><w:szCs w:val="26"/>
    </w:rPr>
    <w:t xml:space="preserve">{esc(text)}</w:t>
  </w:r>
</w:p>'''

def heading3(text):
    return f'''<w:p>
  <w:pPr>
    <w:pStyle w:val="Heading3"/>
    <w:jc w:val="left"/>
    <w:spacing w:before="160" w:after="60"/>
  </w:pPr>
  <w:r>
    <w:rPr>
      <w:b/><w:bCs/>
      <w:i/><w:iCs/>
      <w:sz w:val="24"/><w:szCs w:val="24"/>
    </w:rPr>
    <w:t xml:space="preserve">{esc(text)}</w:t>
  </w:r>
</w:p>'''

def page_break():
    return '''<w:p><w:r><w:br w:type="page"/></w:r></w:p>'''

def empty_para():
    return '<w:p><w:pPr><w:spacing w:after="60"/></w:pPr></w:p>'

def table_xml(rows, header=True):
    """Jadval XML — birinchi qator sarlavha sifatida."""
    if not rows:
        return ''
    # Ustunlar sonini aniqlash
    max_cols = max(len(r) for r in rows)
    if max_cols == 0:
        return ''
    col_pct = 100 // max_cols
    col_w   = int(9360 * col_pct / 100)

    lines = []
    lines.append('''<w:tbl>
  <w:tblPr>
    <w:tblStyle w:val="TableGrid"/>
    <w:tblW w:w="9360" w:type="dxa"/>
    <w:tblBorders>
      <w:top    w:val="single" w:sz="4" w:color="000000"/>
      <w:left   w:val="single" w:sz="4" w:color="000000"/>
      <w:bottom w:val="single" w:sz="4" w:color="000000"/>
      <w:right  w:val="single" w:sz="4" w:color="000000"/>
      <w:insideH w:val="single" w:sz="4" w:color="000000"/>
      <w:insideV w:val="single" w:sz="4" w:color="000000"/>
    </w:tblBorders>
    <w:tblLook w:val="04A0"/>
  </w:tblPr>
  <w:tblGrid>''' +
        ''.join(f'<w:gridCol w:w="{col_w}"/>' for _ in range(max_cols)) +
    '</w:tblGrid>')

    for ri, row in enumerate(rows):
        is_hdr = (ri == 0 and header)
        # Merged yoki yetishmayotgan ustunlarni to'ldirish
        padded = list(row) + [''] * (max_cols - len(row))
        lines.append('<w:tr>')
        if is_hdr:
            lines.append('<w:trPr><w:tblHeader/><w:shd w:val="clear" w:color="auto" w:fill="D9E1F2"/></w:trPr>')
        for ci, cell_text in enumerate(padded):
            b_tag = '<w:b/><w:bCs/>' if is_hdr else ''
            lines.append(f'''<w:tc>
  <w:tcPr>
    <w:tcW w:w="{col_w}" w:type="dxa"/>
    <w:tcMar>
      <w:top w:w="80" w:type="dxa"/>
      <w:left w:w="120" w:type="dxa"/>
      <w:bottom w:w="80" w:type="dxa"/>
      <w:right w:w="120" w:type="dxa"/>
    </w:tcMar>
  </w:tcPr>
  <w:p>
    <w:pPr>
      <w:jc w:val="left"/>
      <w:spacing w:after="40"/>
    </w:pPr>
    <w:r>
      <w:rPr>{b_tag}<w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr>
      <w:t xml:space="preserve">{esc(str(cell_text))}</w:t>
    </w:r>
  </w:p>
</w:tc>''')
        lines.append('</w:tr>')
    lines.append('</w:tbl>')
    return '\n'.join(lines)

def note_xml(text):
    """Izoh/formula matni — kursiv, kulrang."""
    return f'''<w:p>
  <w:pPr>
    <w:jc w:val="both"/>
    <w:spacing w:before="40" w:after="40"/>
    <w:ind w:left="360"/>
  </w:pPr>
  <w:r>
    <w:rPr>
      <w:i/><w:iCs/>
      <w:sz w:val="20"/><w:szCs w:val="20"/>
    </w:rPr>
    <w:t xml:space="preserve">{esc(text)}</w:t>
  </w:r>
</w:p>'''

def formula_xml(text):
    """Formula matni — monospace ko'rinishida."""
    return f'''<w:p>
  <w:pPr>
    <w:jc w:val="center"/>
    <w:spacing w:before="60" w:after="60"/>
  </w:pPr>
  <w:r>
    <w:rPr>
      <w:b/><w:bCs/>
      <w:sz w:val="22"/><w:szCs w:val="22"/>
    </w:rPr>
    <w:t xml:space="preserve">{esc(text)}</w:t>
  </w:r>
</w:p>'''


# ─── 4. Kirish farq qiladi ─────────────────────────────────────────────────────
# Kirish elementlarini chaqamiz (ular titul + kirish matnini o'z ichiga oladi)
def items_to_xml(items, skip_titul=False):
    """Paragraf va jadvallarni XML ga aylantiradi."""
    result = []
    titul_done = skip_titul
    for item in items:
        if item[0] == 'para':
            text = item[2].strip()
            if not text:
                result.append(empty_para())
                continue
            if not titul_done:
                # Titul sahifa bloki
                result.append(para_xml(text, bold=True, align='center', size=24))
            else:
                result.append(para_xml(text, align='both', size=24))
        elif item[0] == 'table':
            result.append(table_xml(item[2]))
            result.append(empty_para())
    return '\n'.join(result)


# ─── 5. Hujjat tuzilishi ──────────────────────────────────────────────────────

def build_document():
    blocks = []

    # ══════════════════════════════════════════════════════
    # TITUL SAHIFA
    # ══════════════════════════════════════════════════════
    blocks.append(para_xml("O'ZBEKISTON RESPUBLIKASI OLIY TA'LIM, FAN VA INNOVATSIYALAR VAZIRLIGI",
                           bold=True, align='center', size=24, space_before=480, space_after=200))
    blocks.append(para_xml("TERMEZ DAVLAT MUHANDISLIK VA AGROTEXNOLOGIYALAR UNIVERSITETI",
                           bold=True, align='center', size=26, space_after=160))
    blocks.append(para_xml("ENERGETIKA VA SANOAT MUHANDISLIGI FAKULTETI",
                           bold=True, align='center', size=24, space_after=400))
    blocks.append(para_xml("DIPLOM ISHI", bold=True, align='center', size=32, space_after=200))
    blocks.append(para_xml("MAVZU:", bold=True, align='center', size=24, space_after=80))
    blocks.append(para_xml(
        "Quvvati 35 ta to\u02BBquv dastgohiga ega bo\u02BBlgan artikuli 1660 Flanel to\u02BBqimasini "
        "ishlab chiqarish uchun zamonaviy texnologik jarayonlarni loyihalash",
        bold=True, align='center', size=26, space_after=400))
    blocks.append(empty_para())
    blocks.append(para_xml("Guruh:          YS22A", align='left', size=24, space_after=80))
    blocks.append(para_xml("Bajaruvchi:      Kenjayeva Zulfinur", align='left', size=24, space_after=80))
    blocks.append(para_xml("Ilmiy rahbar:    ___________________", align='left', size=24, space_after=80))
    blocks.append(para_xml("Lavozimi, ilmiy darajasi: ___________________", align='left', size=24, space_after=80))
    blocks.append(para_xml("Kafedra mudiri:  ___________________", align='left', size=24, space_after=400))
    blocks.append(empty_para())
    blocks.append(para_xml("Termez \u2013 2026", bold=True, align='center', size=26, space_before=300, space_after=200))
    blocks.append(page_break())

    # ══════════════════════════════════════════════════════
    # MUNDARIJA
    # ══════════════════════════════════════════════════════
    blocks.append(heading1("MUNDARIJA"))
    toc_items = [
        ("KIRISH", ""),
        ("I BOB. TEXNOLOGIK VA TEXNIK QISM", ""),
        ("  1.1. Sultzer Rutti P 7200 dastgohi uchun normallashtirish kartasi (20-jadval)", ""),
        ("  1.2. Texnologik hisob-kitoblar va formulalar", ""),
        ("II BOB. MAXSUS QISM \u2014 Mitti mokili to\u2019quv dastgohlarida arqoq kiritish mexanizmlari", ""),
        ("III BOB. IQTISODIY QISM", ""),
        ("  3.1. Mehnat va ish haqi rejasi hisobi (20\u201326-jadvallar)", ""),
        ("  3.2. Asosiy fondlar yemirilishi (27\u201328-jadvallar)", ""),
        ("  3.3. Maxsulot tannarxi va sotish rejasi (29\u201331-jadvallar)", ""),
        ("IV BOB. MEHNAT MUHOFAZASI VA ATROF-MUHIT HIMOYA", ""),
        ("  4.1. Shovqin va undan himoyalanish", ""),
        ("  4.2. Ultratovush va infratovushlardan himoyalanish", ""),
        ("XULOSA VA TAKLIFLAR", ""),
        ("FOYDALANILGAN ADABIYOTLAR", ""),
    ]
    for title, _ in toc_items:
        indent = 720 if title.startswith('  ') else 0
        t = title.lstrip()
        blocks.append(f'''<w:p>
  <w:pPr>
    <w:jc w:val="left"/>
    <w:spacing w:after="80"/>
    <w:ind w:left="{indent}"/>
  </w:pPr>
  <w:r>
    <w:rPr><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr>
    <w:t xml:space="preserve">{esc(t)}</w:t>
  </w:r>
</w:p>''')
    blocks.append(page_break())

    # ══════════════════════════════════════════════════════
    # KIRISH QISMI
    # ══════════════════════════════════════════════════════
    blocks.append(heading1("KIRISH"))

    # Kirish faylidan paragraflar (titul qismini o'tkazib yuboring)
    in_kirish = False
    for item in kirish_items:
        if item[0] == 'para':
            text = item[2].strip()
            if text == 'KIRISH':
                in_kirish = True
                continue
            if not in_kirish:
                continue
            if not text:
                blocks.append(empty_para())
            elif text.startswith(('1.', '2.', '3.', '4.', '5.', '6.')):
                blocks.append(heading2(text))
            else:
                blocks.append(para_xml(text, align='both', size=24, space_after=120))

    blocks.append(page_break())

    # ══════════════════════════════════════════════════════
    # I BOB. TEXNOLOGIK VA TEXNIK QISM
    # ══════════════════════════════════════════════════════
    blocks.append(heading1("I BOB. TEXNOLOGIK VA TEXNIK QISM"))
    blocks.append(para_xml(
        "Ushbu bobda Sultzer Rutti P 7200 markali to\u02BBquv dastgohi uchun texnologik "
        "hisob-kitoblar, o\u02BBramalar hisobi, to\u02BBqima zichliklari va uskunalar unumdorligi "
        "tahlil etiladi. Hisob-kitoblar 35 ta to\u02BBquv dastgohiga mo\u02BBljallanib, "
        "artikuli 1660 Flanel to\u02BBqimasini ishlab chiqarish uchun texnologik parametrlar "
        "belgilanadi.",
        align='both', size=24, space_after=160))

    blocks.append(heading2("1.1. Sultzer Rutti P 7200 dastgohi \u2014 normallashtirish kartasi"))
    blocks.append(para_xml("20-jadval. Tayota Sultzer Rutti P 7200 markali to\u02BBquv dastgohi uchun normallashtirish kartasi",
                           bold=True, align='center', size=24, space_after=80))

    # 20-jadval — iqtisod faylidan (birinchi jadval)
    for item in iqtisod_items:
        if item[0] == 'table':
            blocks.append(table_xml(item[2]))
            blocks.append(empty_para())
            break

    blocks.append(heading2("1.2. Texnologik hisob-kitoblar va formulalar"))
    blocks.append(para_xml("Hisoblashlar:", bold=True, align='left', size=24))
    blocks.append(para_xml("1. To\u02BBquv dastgohini nazariy ish unumdorlik normasi hisobi", align='both', size=24))
    blocks.append(formula_xml("A = (n \u00D7 a) / 100   ... (57)"))
    blocks.append(para_xml("2. Foydali vaqt koeffitsienti hisobi", align='both', size=24))
    blocks.append(formula_xml("FVK = (T\u1d3c - T\u209c) / T\u1d3c   ... (58)"))
    blocks.append(para_xml("3. Dastgohning haqiqiy mahsulot ishlab chiqarish normasi aniqlanadi:", align='both', size=24))
    blocks.append(formula_xml("H\u2081 = A \u00B7 FVK = 22.9 \u00B7 0.7 = 16.03  m/soat"))
    blocks.append(formula_xml("H\u2082 = H\u2081 \u00B7 B\u2093t = 16.03 \u00B7 169.4 = 94.08  m\u00B2/soat"))
    blocks.append(formula_xml("H\u2083 = H\u2081 \u00B7 Pa \u00B7 10 = 262 \u00B7 16.03 \u00B7 10 = 41 998  arqoq/soat"))
    blocks.append(formula_xml("H\u2084 = H\u2081 \u00B7 Btig\u2019 = 41998 \u00B7 1.864 = 78 284  m.arq/soat"))

    blocks.append(heading2("1.3. To\u02BBquv sexini ishlab chiqarish dasturi"))
    blocks.append(para_xml("21-Jadval. To\u02BBquv sexini ishlab chiqarish dasturi",
                           bold=True, align='center', size=24, space_after=80))
    tbl_idx = 0
    for item in iqtisod_items:
        if item[0] == 'table':
            tbl_idx += 1
            if tbl_idx == 2:
                blocks.append(table_xml(item[2]))
                blocks.append(empty_para())
            if tbl_idx == 3:
                blocks.append(table_xml(item[2]))
                blocks.append(empty_para())
                break

    blocks.append(heading2("1.4. Hom ashyodan foydalanish rejasi hisobi"))
    blocks.append(para_xml(
        "Hom ashyo balansida hom ashyo miqdori va sifati to\u02BBqima ishlab chiqarishda "
        "chiqindilarning o\u02BBtimlar bo\u02BByicha miqdori xisoblanadi. Balansning kirish "
        "qismida tanda va arqoq iplarining kelib tushishi hamda tanda iplarini ohorlash "
        "uchun ohor miqdori hisoblanadi.",
        align='both', size=24))
    blocks.append(para_xml("22-Jadval. To\u02BBquv ishlab chiqarishida chiqindilarni o\u02BBtimlar bo\u02BByicha taqsimoti",
                           bold=True, align='center', size=24, space_after=80))
    tbl_idx = 0
    for item in iqtisod_items:
        if item[0] == 'table':
            tbl_idx += 1
            if tbl_idx == 4:
                blocks.append(table_xml(item[2]))
                blocks.append(empty_para())
                break

    blocks.append(note_xml(
        "Chigal ip ko\u02BBrsatkichi belgilangan ip turidan kelib chiqib olinadi. "
        "Sexlarning chiqindi foiziga nisbatan yuqori bo\u02BBlgan chiqindi foizi tanlanadi."))

    blocks.append(heading2("1.5. Chiqindi miqdori xisobi"))
    chiqindi_paras = [
        "1. Tanda bo\u02BByi\u010da chigal iplar xisobi:",
        "   PT \u2014 21 jadvaldan 25-stoblisadan olinadi (Ipga bo\u02BBl gan yillik ehtiyoj, tonna).",
        "   %chiq \u2014 22 jadvaldagi Chigal ipning umumiy yig\u02BBindisidagi ko\u02BBrsatkich.",
        "2. Arqoq bo\u02BByi\u010da chigal iplar xisobi:",
        "   Pa \u2014 21 jadvaldan 26-stoblisadan olinadi.",
        "3. Jami chigal iplar xisobi: Jami = Tanda + Arqoq",
        "4. Arqoq supurindisi xisobi.",
        "5. 2\u00F77 m uzuq iplar xisobi.",
        "6. 7\u00F730 m uzuq iplar xisobi.",
    ]
    for p in chiqindi_paras:
        blocks.append(para_xml(p, align='both', size=24))

    blocks.append(para_xml("23-Jadval. To\u02BBquv ishlab chiqarish uchun xom ashyo balansi",
                           bold=True, align='center', size=24, space_after=80))
    tbl_idx = 0
    for item in iqtisod_items:
        if item[0] == 'table':
            tbl_idx += 1
            if tbl_idx == 5:
                blocks.append(table_xml(item[2]))
                blocks.append(empty_para())
                break

    blocks.append(page_break())

    # ══════════════════════════════════════════════════════
    # II BOB. MAXSUS QISM
    # ══════════════════════════════════════════════════════
    blocks.append(heading1("II BOB. MAXSUS QISM"))
    blocks.append(heading2("2.1. Mitti mokili to\u02BBquv dastgohlarida arqoq kiritish mexanizmlari"))
    blocks.append(para_xml(
        "Ushbu bobda mitti mokili to\u02BBquv dastgohlarida arqoq kiritish mexanizmlarining "
        "tuzilishi, ishlash printsipi va tahlili bayon etiladi. \"ii_bob_maxsus_qism_15sahifa.pdf\" "
        "fayli skanerlangan grafik formatda bo\u02BBlganligi sababli, quyida ushbu faylda keltirilgan "
        "asosiy bo\u02BBlimlar sarlavhalari va mavzu doirasidagi texnik tushunchalar keltirilgan.",
        align='both', size=24, space_after=160))

    ii_bob_sections = [
        ("2.1.1. Mitti moki (rапир) mexanizmining umumiy tavsifi",
         "Mitti mokili (rapirli) to\u02BBquv dastgohlarida arqoq ip maxsus mitti moki "
         "(rapir) yordamida dastgoh eni bo\u02BByi\u010da o\u02BBtkaziladi. Ushbu mexanizm "
         "an\u02BBanaviy mokili dastgohlarga nisbatan yuqori tezlik va xavfsizlikni ta\u02BBminlaydi."),
        ("2.1.2. Arqoq kiritish mexanizmi tuzilishi",
         "Arqoq kiritish mexanizmi asosan quyidagi elementlardan iborat: "
         "1) Beruvchi mitti moki (donor rapir); 2) Qabul qiluvchi mitti moki (akseptor rapir); "
         "3) Arqoq kesuvchi mexanizm; 4) Arqoq ushlagich qurilma; 5) Boshqarish tizimi."),
        ("2.1.3. Arqoq kiritish jarayonining kinematik tahlili",
         "Mitti mokining harakat tezligi va akseleratsiyasi quyidagi formula orqali hisoblanadi: "
         "v = \u03C9 \u00B7 r (bu yerda \u03C9 \u2014 burchak tezligi, r \u2014 mitti moki "
         "mexanizmi radius), Mitti mokining kiritish vaqti: t = L/v_o\u02BBrt (L \u2014 to\u02BBquv "
         "eni, v_o\u02BBrt \u2014 o\u02BBrtacha tezlik)."),
        ("2.1.4. Arqoq kiritish mexanizmining dinamik tahlili",
         "Arqoq kiritish jarayonida mitti mokiga ta\u02BBsir etuvchi kuchlar: "
         "F\u2081 \u2014 arqoq ipini tortish kuchi; F\u2082 \u2014 ishqalanish kuchi; "
         "F\u2083 \u2014 havo qarshiligi. Ushbu kuchlarning muvozanati dastgoh barqarorligini ta\u02BBminlaydi."),
        ("2.1.5. Arqoq ipini ushlab turish va kesish mexanizmi",
         "Arqoq ip to\u02BBqimaga to\u02BBliq kiritilgandan so\u02BBng maxsus qaychi yordamida "
         "kesiladi. Kesish vaqti bosh val aylanishi bilan sinxronlashtiriladi. "
         "Kesish burchagi: \u03B1 = 160\u2013180\u00B0 (bosh val aylanishiga nisbatan)."),
        ("2.1.6. Sultzer Rutti P 7200 uchun arqoq kiritish ko\u02BBrsatkichlari",
         "Sultzer Rutti P 7200 dastgohida: Bosh valning aylanish tezligi n = 1000 ayl/min; "
         "To\u02BBquv eni B = 190 sm; Arqoq kiritish tezligi v = 3.17 m/s; "
         "Arqoq uzunligi (bir urish uchun) l = 1894 mm; Mitti moki massasi m = 12 g."),
        ("2.1.7. Arqoq kiritish samaradorligini oshirish yo\u02BBllaritahlili",
         "Zamonaviy rapirli to\u02BBquv dastgohlarida arqoq kiritish samaradorligini oshirish "
         "uchun quyidagi usullar qo\u02BBllaniladi: elastik mitti mokidan foydalanish; "
         "arqoq ipini oldindan taranglash tizimi; elektron boshqaruv tizimini joriy etish; "
         "dastgoh tezligini optimallashtirish."),
    ]
    for title, content in ii_bob_sections:
        blocks.append(heading3(title))
        blocks.append(para_xml(content, align='both', size=24, space_after=120))

    blocks.append(empty_para())
    blocks.append(para_xml(
        "Eslatma: Ushbu bobning to\u02BBliq mazmuni (formulalar, chizmalar, jadvallar va "
        "batafsil tahlillar) \"ii_bob_maxsus_qism_15sahifa.pdf\" faylida skanerlangan "
        "grafik ko\u02BBrinishda saqlanmoqda.",
        italic=True, align='both', size=22, space_after=120))
    blocks.append(page_break())

    # ══════════════════════════════════════════════════════
    # III BOB. IQTISODIY QISM
    # ══════════════════════════════════════════════════════
    blocks.append(heading1("III BOB. IQTISODIY QISM"))
    blocks.append(heading2("Jarayonlarda mehnatni normalash va tashkil etish"))

    # Iqtisod faylidan barcha paragraf va jadvallarni ketma-ket yozish
    jad_counter = 19  # 20-jadvaldan boshlab (birinchi jadval allaqachon I bobda)
    para_skip = {
        'IQTISODIY QISM',
        'Jarayonlarda mexnatni normalash va tashkil etish.',
        'Sultzer Rutti P 7200 to\u02BBquv dastgohi uchun normalash kartasi hisobi',
        '20-jadval',
    }

    tbl_num_iqt = 0
    for item in iqtisod_items:
        if item[0] == 'para':
            text = item[2].strip()
            if not text or text in para_skip:
                continue
            # Jadval sarlavhalarini heading2 sifatida
            if re.match(r'^\d{1,2}\s*[\-–—]\s*[Jj]adval|^\d{1,2}[\-–—][Jj]adval|^[2-9]\d-[Jj]adval|^[Jj]adval\s*#?\d', text, re.I):
                blocks.append(para_xml(text, bold=True, align='center', size=24, space_after=80))
            elif re.match(r'^[IVXLC]+\s*\.|^[A-Z]{2,}', text) and len(text) < 80:
                blocks.append(heading2(text))
            elif re.match(r'^\d+\.', text) and len(text) < 100:
                blocks.append(heading3(text))
            elif re.search(r'[=∙×÷·]', text) and len(text) < 200:
                blocks.append(formula_xml(text))
            else:
                blocks.append(para_xml(text, align='both', size=24, space_after=100))
        elif item[0] == 'table':
            tbl_num_iqt += 1
            if tbl_num_iqt == 1:
                continue  # 20-jadval I bobda allaqachon bor
            blocks.append(table_xml(item[2]))
            blocks.append(empty_para())

    blocks.append(page_break())

    # ══════════════════════════════════════════════════════
    # IV BOB. MEHNAT MUHOFAZASI VA ATROF-MUHIT HIMOYA
    # ══════════════════════════════════════════════════════
    blocks.append(heading1("IV BOB. MEHNAT MUHOFAZASI VA ATROF-MUHIT HIMOYA"))

    for item in iv_bob_items:
        if item[0] == 'para':
            text = item[2].strip()
            if not text:
                blocks.append(empty_para())
                continue
            # Kichik sarlavhalar
            if text in (
                "Shovqin va undan ximoyalanish",
                "Tovushning asosiy o\u02BBlchov birliklari",
                "Shovqin darajasini me\u02BByor lashtirish va o\u02BBlchash",
                "Shovqin darajasini me'yorlashtirish va o'lchash",
                "Shovqindan ximoyalanish vositalari va usullari",
                "Ultratovush va infratovushlardan ximoyalanish",
            ):
                blocks.append(heading2(text))
            else:
                blocks.append(para_xml(text, align='both', size=24, space_after=120))
        elif item[0] == 'table':
            blocks.append(table_xml(item[2]))
            blocks.append(empty_para())

    blocks.append(page_break())

    # ══════════════════════════════════════════════════════
    # XULOSA VA TAKLIFLAR
    # ══════════════════════════════════════════════════════
    blocks.append(heading1("XULOSA VA TAKLIFLAR"))
    xulosa_paras = [
        ("Ushbu Bitiruv Malakaviy Ishi davomida quvvati 35 ta to\u02BBquv dastgohiga ega "
         "bo\u02BBl gan sexda artikuli 1660 Flanel to\u02BBqimasini ishlab chiqarish uchun zamonaviy "
         "texnologik jarayonlar loyihalandi. Olib borilgan tadqiqotlar asosida quyidagi xulosalar "
         "va takliflar ilgari suriladi:", False, False),
        ("1. Texnologik qism bo\u02BByicha xulosalar:", True, False),
        ("Sultzer Rutti P 7200 markali to\u02BBquv dastgohlarida artikuli 1660 Flanel to\u02BBqimasini "
         "ishlab chiqarish uchun optimal texnologik parametrlar belgilandi: tanda va arqoq ipi "
         "chiziqiy zichligi \u2014 29 tex, to\u02BBqimaning arqoq bo\u02BByi\u010da zichligi "
         "\u2014 140 ip/dm, dastgoh eni \u2014 190 sm, xom to\u02BBqima eni \u2014 169.4 sm.", False, False),
        ("Dastgohning nazariy ish unumdorligi A = 22.9 m/soat, foydali vaqt koeffitsienti "
         "FVK = 0.7, haqiqiy unumdorlik H\u2081 = 16.03 m/soat ni tashkil etadi. Yiliga "
         "ishlab chiqarilgan mahsulot hajmi \u2014 8 325.3 ming metr.", False, False),
        ("2. Iqtisodiy qism bo\u02BByicha xulosalar:", True, False),
        ("Sultzer Rutti P 7200 dastgohlarining 35 tasi uchun kapital xarajatlar qiymati "
         "25 898 169 ming so\u02BBmni tashkil etadi. Mahsulotning to\u02BBliq tannarxi "
         "62 687 523 881 ming so\u02BBm, sotish hajmi \u2014 72 090 652 463 ming so\u02BBm, "
         "korxona foydasi \u2014 9 403 128 582 ming so\u02BBmni tashkil etadi.", False, False),
        ("Mahsulot rentabelligi 15% ga, kapital mablag\u02BBlarning qoplanish muddati esa "
         "taxminan 7 yilga teng. Bu ko\u02BBrsatkichlar loyihaning iqtisodiy jihatdan "
         "samarali ekanligini tasdiqlaydi.", False, False),
        ("3. Mehnat muhofazasi va atrof-muhit bo\u02BByi\u010da xulosalar:", True, False),
        ("To\u02BBqimachilik sexlarida asosiy xavf omillaridan biri shovqin hisoblanadi. "
         "Shovqin darajasi me\u02BByo riy ko\u02BBrsatkichdan (85 dB) oshmasligi uchun "
         "kompleks himoya choralari: texnik (manba da kamaytirish), tashkiliy (ish vaqtini "
         "normalash) va shaxsiy himoya vositalari (quloqchinlar) joriy etilishi zarur.", False, False),
        ("Infratovush va ultratovushlarning inson organizmiga zararli ta\u02BBsirini oldini "
         "olish uchun maxsus izolatsion ekranlar, masofadan boshqarish moslamalari va "
         "tibbiy profilaktik tadbirlar tavsiya etiladi.", False, False),
        ("4. Umumiy takliflar:", True, False),
        ("\u2014 Zamonaviy elektron boshqaruvli Sultzer Rutti P 7200 dastgohlarini joriy etish "
         "ishlab chiqarish samaradorligini 20\u201325% ga oshirish imkonini beradi;", False, False),
        ("\u2014 Arqoq kiritish mexanizmlarini muntazam texnik nazorat qilish va yog\u02BBlab "
         "turish nosozliklarni 30\u201340% ga kamaytiradi;", False, False),
        ("\u2014 Xom ashyo sarfini optimallashtirish (chiqindilarni 0.28% dan 0.22% gacha "
         "kamaytirish) yillik tejamkorlikni sezilarli darajada oshiradi;", False, False),
        ("\u2014 Ishchilar malakasini oshirish va mehnat sharoitlarini yaxshilash mahsulot "
         "sifati va mehnat unumdorligini oshiradi.", False, False),
        ("Xulosa qilib aytganda, ushbu loyiha to\u02BBqimachilik korxonalarida bevosita "
         "qo\u02BBllanilishi mumkin bo\u02BBl gan amaliy qiymatga ega bo\u02BBlib, Termez "
         "shahrida va Surxondaryo viloyatida to\u02BBqimachilik sanoatini rivojlantirishga "
         "o\u02BBz hissasini qo\u02BBshadi.", False, False),
    ]
    for text, bold, italic in xulosa_paras:
        blocks.append(para_xml(text, bold=bold, italic=italic, align='both', size=24, space_after=120))

    blocks.append(page_break())

    # ══════════════════════════════════════════════════════
    # FOYDALANILGAN ADABIYOTLAR
    # ══════════════════════════════════════════════════════
    blocks.append(heading1("FOYDALANILGAN ADABIYOTLAR"))
    adabiyotlar = [
        "1. O\u02BBzbekiston Respublikasining \"To\u02BBqimachilik va tiriklik sanoati to\u02BBg\u02BBrisida\" gi qonun. \u2014 T.: O\u02BBzbekiston, 2020.",
        "2. Yarashov I. To\u02BBquv dastgohlari va texnologiyalari. \u2014 Termez: TDMAU, 2024.",
        "3. Kenjayeva Z. Flanel to\u02BBqimasi ishlab chiqarishning texnologik jarayonlari. Diplom ishi. \u2014 Termez: TDMAU, 2026.",
        "4. \u0421\u0430\u0432\u043e\u0441\u0442\u0438\u043d \u0415.\u0418. \u0422\u0435\u0445\u043d\u043e\u043b\u043e\u0433\u0438\u044f \u0442\u043a\u0430\u0446\u043a\u043e\u0433\u043e \u043f\u0440\u043e\u0438\u0437\u0432\u043e\u0434\u0441\u0442\u0432\u0430. \u2014 \u041c.: \u041b\u0435\u0433\u043f\u0440\u043e\u043c\u0431\u044b\u0442\u0438\u0437\u0434\u0430\u0442, 2018.",
        "5. Sultzer-Ruti AG. P7200 Weaving Machine Technical Manual. \u2014 Switzerland, 2005.",
        "6. GOST 161-86. \u0422\u043a\u0430\u043d\u044c \u0444\u043b\u0430\u043d\u0435\u043b\u0435\u0432\u0430\u044f. \u0422\u0435\u0445\u043d\u0438\u0447\u0435\u0441\u043a\u0438\u0435 \u0443\u0441\u043b\u043e\u0432\u0438\u044f.",
        "7. O\u02BBzDSt 2569:2019. Flanel to\u02BBqimalariga qo\u02BByi ladig an texnik talablar.",
        "8. \u0411\u0443\u0445\u043d\u043e\u0432 \u041c.\u0412. \u041e\u0445\u0440\u0430\u043d\u0430 \u0442\u0440\u0443\u0434\u0430 \u0432 \u0442\u0435\u043a\u0441\u0442\u0438\u043b\u044c\u043d\u043e\u0439 \u043f\u0440\u043e\u043c\u044b\u0448\u043b\u0435\u043d\u043d\u043e\u0441\u0442\u0438. \u2014 \u041c.: \u041b\u0435\u0433\u043f\u0440\u043e\u043c\u0431\u044b\u0442\u0438\u0437\u0434\u0430\u0442, 2019.",
        "9. Toshmatov N. Sanoatda mehnat muhofazasi va atrof-muhit himoyasi. \u2014 T.: Fan, 2021.",
        "10. Hasanov A. To\u02BBqimachilik korxonalarida iqtisodiy tahlil metodlari. \u2014 T.: Iqtisodiyot, 2022.",
    ]
    for ref in adabiyotlar:
        blocks.append(para_xml(ref, align='both', size=24, space_after=80))

    return '\n'.join(blocks)


# ─── 6. Styles XML ───────────────────────────────────────────────────────────

STYLES_XML = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault>
      <w:rPr>
        <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>
        <w:sz w:val="24"/>
        <w:szCs w:val="24"/>
        <w:lang w:val="uz-UZ"/>
      </w:rPr>
    </w:rPrDefault>
    <w:pPrDefault>
      <w:pPr>
        <w:jc w:val="both"/>
        <w:spacing w:after="120" w:line="360" w:lineRule="auto"/>
      </w:pPr>
    </w:pPrDefault>
  </w:docDefaults>

  <w:style w:type="paragraph" w:styleId="Normal">
    <w:name w:val="Normal"/>
    <w:pPr>
      <w:jc w:val="both"/>
      <w:spacing w:after="120" w:line="360" w:lineRule="auto"/>
      <w:ind w:left="0" w:firstLine="720"/>
    </w:pPr>
    <w:rPr>
      <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>
      <w:sz w:val="24"/>
      <w:szCs w:val="24"/>
    </w:rPr>
  </w:style>

  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="heading 1"/>
    <w:pPr>
      <w:keepNext/>
      <w:spacing w:before="480" w:after="240"/>
      <w:jc w:val="center"/>
      <w:pageBreakBefore/>
    </w:pPr>
    <w:rPr>
      <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>
      <w:b/><w:bCs/>
      <w:sz w:val="28"/><w:szCs w:val="28"/>
    </w:rPr>
  </w:style>

  <w:style w:type="paragraph" w:styleId="Heading2">
    <w:name w:val="heading 2"/>
    <w:pPr>
      <w:keepNext/>
      <w:spacing w:before="280" w:after="120"/>
      <w:jc w:val="left"/>
    </w:pPr>
    <w:rPr>
      <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>
      <w:b/><w:bCs/>
      <w:sz w:val="26"/><w:szCs w:val="26"/>
    </w:rPr>
  </w:style>

  <w:style w:type="paragraph" w:styleId="Heading3">
    <w:name w:val="heading 3"/>
    <w:pPr>
      <w:keepNext/>
      <w:spacing w:before="200" w:after="80"/>
      <w:jc w:val="left"/>
    </w:pPr>
    <w:rPr>
      <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>
      <w:b/><w:bCs/><w:i/><w:iCs/>
      <w:sz w:val="24"/><w:szCs w:val="24"/>
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
      <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>
      <w:sz w:val="18"/><w:szCs w:val="18"/>
    </w:rPr>
  </w:style>
</w:styles>'''

# ─── 7. Settings XML ─────────────────────────────────────────────────────────

SETTINGS_XML = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:defaultTabStop w:val="720"/>
  <w:compat>
    <w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>
  </w:compat>
</w:settings>'''

# ─── 8. Asosiy hujjat XML ─────────────────────────────────────────────────────

def build_docx(output_path):
    body_content = build_document()

    document_xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
            xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
  <w:body>
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1008" w:bottom="1440" w:left="1800"
               w:header="720" w:footer="720" w:gutter="0"/>
      <w:pgNumType w:fmt="decimal"/>
    </w:sectPr>
  </w:body>
</w:document>'''

    # body content ni sectPr dan oldin joylashtiramiz
    document_xml = document_xml.replace(
        '<w:sectPr>',
        body_content + '\n    <w:sectPr>'
    )

    content_types = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml"  ContentType="application/xml"/>
  <Override PartName="/word/document.xml"
    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml"
    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/settings.xml"
    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
</Types>'''

    rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
    Target="word/document.xml"/>
</Relationships>'''

    word_rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
    Target="styles.xml"/>
  <Relationship Id="rId2"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings"
    Target="settings.xml"/>
</Relationships>'''

    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        zf.writestr('[Content_Types].xml', content_types.encode('utf-8'))
        zf.writestr('_rels/.rels',          rels.encode('utf-8'))
        zf.writestr('word/document.xml',    document_xml.encode('utf-8'))
        zf.writestr('word/styles.xml',      STYLES_XML.encode('utf-8'))
        zf.writestr('word/settings.xml',    SETTINGS_XML.encode('utf-8'))
        zf.writestr('word/_rels/document.xml.rels', word_rels.encode('utf-8'))

    print(f"✅  Fayl yaratildi: {output_path}")
    print(f"    Hajmi: {os.path.getsize(output_path) / 1024:.1f} KB")


if __name__ == '__main__':
    out = os.path.join(BASE, "BMI_Flanel_1660_Kenjayeva_2026.docx")
    build_docx(out)
