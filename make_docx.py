"""
Builds Iqtisodiy_qism_Yarashov_Flanel_1660.docx
Pure stdlib: zipfile + xml only. No python-docx needed.
"""

import zipfile, os
from xml.sax.saxutils import escape

OUT = "Iqtisodiy_qism_Yarashov_Flanel_1660.docx"

WNS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
RNS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
MC  = "http://schemas.openxmlformats.org/markup-compatibility/2006"

def e(s): return escape(str(s))

# ── low-level XML helpers ────────────────────────────────────────────
def _rpr(bold=False, sz=22, font="Times New Roman"):
    b = "<w:b/><w:bCs/>" if bold else ""
    return (
        f"<w:rPr>{b}"
        f"<w:rFonts w:ascii=\"{font}\" w:hAnsi=\"{font}\" w:cs=\"{font}\"/>"
        f"<w:sz w:val=\"{sz}\"/><w:szCs w:val=\"{sz}\"/>"
        f"</w:rPr>"
    )

def _ppr(align="left", indent=0, keeplines=False):
    jc  = f'<w:jc w:val="{align}"/>'
    ind = f'<w:ind w:left="{indent}"/>' if indent else ""
    kl  = "<w:keepLines/>" if keeplines else ""
    return f"<w:pPr>{kl}{jc}{ind}{_rpr()}</w:pPr>"

def _run(text, bold=False, sz=22):
    return f"<w:r>{_rpr(bold,sz)}<w:t xml:space=\"preserve\">{e(text)}</w:t></w:r>"

def para(text, bold=False, sz=22, align="left", indent=0):
    return (
        f"<w:p>"
        f"{_ppr(align, indent)}"
        f"{_run(text, bold, sz)}"
        f"</w:p>"
    )

def heading(text, level=1):
    sizes = {1: 32, 2: 26, 3: 24}
    return para(text, bold=True, sz=sizes.get(level, 24), align="center")

def empty_para():
    return f"<w:p>{_ppr()}</w:p>"

def formula(text):
    return para(text, bold=False, sz=22, align="center")

# ── table builder (clean, no string multiplication) ─────────────────
def tbl(rows, widths=None, hdr=1):
    """
    rows   : list[list[str]]
    widths : list[int] in twips (1 cm ~ 567 twips)
    hdr    : number of header rows (blue bg + bold)
    """
    ncols = max(len(r) for r in rows)
    if widths is None:
        w = 9200 // ncols
        widths = [w] * ncols
    # pad widths if short
    while len(widths) < ncols:
        widths.append(widths[-1])

    tw = sum(widths)
    border = (
        "<w:tblBorders>"
        '<w:top    w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:left   w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:right  w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        "</w:tblBorders>"
    )
    xml = []
    xml.append(
        f"<w:tbl>"
        f"<w:tblPr>"
        f'<w:tblW w:w="{tw}" w:type="dxa"/>'
        f"{border}"
        f'<w:tblLook w:val="04A0"/>'
        f"</w:tblPr>"
        f"<w:tblGrid>"
    )
    for cw in widths:
        xml.append(f'<w:gridCol w:w="{cw}"/>')
    xml.append("</w:tblGrid>")

    for ri, row in enumerate(rows):
        is_hdr = ri < hdr
        xml.append("<w:tr>")
        if is_hdr:
            xml.append("<w:trPr><w:tblHeader/></w:trPr>")
        for ci in range(ncols):
            cell_text = str(row[ci]) if ci < len(row) else ""
            cw = widths[ci]
            shd = '<w:shd w:val="clear" w:color="auto" w:fill="BDD7EE"/>' if is_hdr else ""
            cell_rpr = _rpr(bold=is_hdr, sz=18)
            xml.append(
                f"<w:tc>"
                f"<w:tcPr>"
                f'<w:tcW w:w="{cw}" w:type="dxa"/>'
                f"{shd}"
                f"</w:tcPr>"
                f"<w:p>"
                f'<w:pPr><w:jc w:val="center"/>'
                f"{cell_rpr}"
                f"</w:pPr>"
                f"<w:r>{cell_rpr}"
                f'<w:t xml:space="preserve">{e(cell_text)}</w:t>'
                f"</w:r>"
                f"</w:p>"
                f"</w:tc>"
            )
        xml.append("</w:tr>")
    xml.append("</w:tbl>")
    return "".join(xml)

# ════════════════════════════════════════════════════════════════════
# BUILD DOCUMENT CONTENT
# ════════════════════════════════════════════════════════════════════
P = []
add = P.append

# ── Sarlavha ────────────────────────────────────────────────────────
add(heading("IQTISODIY QISM", 1))
add(heading("I.Yarashov | 35 ta Sultzer Ruti P7200 to'quv dastgohi", 2))
add(heading("Flanel  Artikul 1660", 2))
add(empty_para())

# ── Boshlangich malumotlar ───────────────────────────────────────────
add(heading("BOSHLANGICH MALUMOTLAR", 2))
add(tbl([
    ["No","Korsatkich","Olchov birligi","Qiymat"],
    ["1","Dastgoh nomi","—","Sultzer Ruti P7200"],
    ["2","Dastgohlar soni","dona","35"],
    ["3","Toqima nomi va artikuli","—","Flanel, Art. 1660"],
    ["4","Tayyor toqima eni (Btay)","sm","160"],
    ["5","Xom toqima eni (Bxt)","sm","169.4"],
    ["6","Tig boyicha eni (Btig)","sm","183.93"],
    ["7","Tanda ipi zichligi (Tt)","teks","29"],
    ["8","Arqoq ipi zichligi (Ta)","teks","29"],
    ["9","Tanda boyicha zichlik (Pt)","ip/dm","152"],
    ["10","Arqoq boyicha zichlik (Pa)","ip/dm","140"],
    ["11","Umumiy tanda iplari soni","ip","2 579"],
    ["12","Bosh val aylanish soni","ayl/min","400"],
    ["13","Tanda chiqindi foizi","%","0.74"],
    ["14","Arqoq chiqindi foizi","%","1.40"],
    ["15","Smena davomiyligi","soat","7.41"],
    ["16","Smenalar soni","smena","2"],
    ["17","Yildagi ish kunlari","kun","280"],
], widths=[500, 3500, 1500, 1500]))
add(empty_para())

# ══ 20-JADVAL ══════════════════════════════════════════════════════
add(heading("I. JARAYONLARDA MEHNATNI NORMALASH VA TASHKIL ETISH", 2))
add(para("20-jadval. Sultzer Ruti P7200 markali toquv dastgohi uchun normallashtirish kartasi",
         bold=True, sz=22, align="center"))
add(tbl([
    ["No","Korsatkich","Olchov birligi","Qiymat"],
    ["","I. Dastgoh tavsifi","",""],
    ["1","Dastgoh nomi va markasi","—","Sultzer Ruti P7200"],
    ["2","Dastgohning ishchi eni","sm","183.93"],
    ["3","Arqoq kiritish mexanizmi","—","Pnevmatik (mitti-mokili)"],
    ["4","Arqoq bobina turi","—","Konussimon"],
    ["","II. Dastgohni ishga tayyorlash va mahsulot tavsifi","",""],
    ["5","Xom toqima eni","sm","169.4"],
    ["6","Tig boyicha toqima eni","sm","183.93"],
    ["7","Tayyor toqima eni","sm","160.0"],
    ["8","Bosh valning aylanish soni","ayl/min","400"],
    ["9","Tanda ipining chiziqiy zichligi","teks","29"],
    ["10","Arqoq ipining chiziqiy zichligi","teks","29"],
    ["11","Toqimaning arqoq boyicha zichligi","ip/dm","140"],
    ["12","Iplarning umumiy soni","ip","2 579"],
    ["13","\"a\" guruhi toxtashlar koeffitsienti","—","0.93"],
    ["14","\"b\" guruhi toxtashlar koeffitsienti","—","0.85"],
    ["","III. Tashkiliy shartlar tavsifi","",""],
    ["15","Smena davomiyligi","soat","7.41"],
    ["16","Smenalar soni","smena","2"],
    ["17","Yildagi ish kunlari","kun","280"],
], widths=[500, 4200, 1500, 1700]))
add(empty_para())
add(para("Hisoblashlar:", bold=True, sz=22))
add(para("1. Nazariy unumdorlik normasi hisobi (formula 57):", sz=22))
add(formula("A = (n_bv x 60) / (Pa x 10) = (400 x 60) / (140 x 10) = 24000 / 1400 = 17,143 m/soat"))
add(para("bu yerda: n_bv = 400 ayl/min;  Pa = 140 ip/dm", sz=20))
add(empty_para())
add(para("2. Foydali vaqt koeffitsienti hisobi (formula 58):", sz=22))
add(formula("FVK = Ka x Kb = 0,93 x 0,85 = 0,7905"))
add(para("bu yerda: Ka = 0,93 — \"a\" guruh;  Kb = 0,85 — \"b\" guruh toxtashlari koeffitsienti", sz=20))
add(empty_para())
add(para("3. Dastgohning haqiqiy mahsulot ishlab chiqarish normasi:", sz=22))
add(formula("H1 = A x FVK = 17,143 x 0,7905 = 13,551 m/soat"))
add(formula("H2 = H1 x Bxt / 100 = 13,551 x 169,4 / 100 = 22,956 m2/soat"))
add(formula("H3 = H1 x Pa x 10 = 13,551 x 140 x 10 = 18 971 arqoq/soat"))
add(formula("H4 = H3 x Btig / 100 = 18 971 x 183,93 / 100 = 34 893 m.arq/soat"))
add(empty_para())

# ══ 21-JADVAL ══════════════════════════════════════════════════════
add(heading("II. TOQUV SEXINI ISHLAB CHIQARISH DASTURI", 2))
add(para("21-jadval. Toquv sexini ishlab chiqarish dasturi", bold=True, sz=22, align="center"))
add(tbl([
    ["No","Korsatkich","Olchov birligi","Qiymat","Formula / Izoh"],
    ["1","Toqima nomi","—","Flanel","—"],
    ["2","Artikul","—","1660","—"],
    ["3","Toqima eni (xom)","sm","169.4","—"],
    ["4","Tanda ipi zichligi","teks","29","—"],
    ["5","Arqoq ipi zichligi","teks","29","—"],
    ["6","Arqoq zichligi (10 smda)","ip/dm","140","—"],
    ["7","Smena davomiyligi","soat","7.41","—"],
    ["8","Smenalar soni","—","2","—"],
    ["9","Yildagi ish kunlari","kun","280","—"],
    ["10","Yildagi ish soatlari","soat","4 149.6","7,41 x 2 x 280"],
    ["11","O'rnatilgan dastgohlar soni","dona","35","—"],
    ["12","O'rnatilgan dastgoh soatlari","soat","145 236.0","35 x 4 149,6"],
    ["13","IUK","koef","0.956","qabul qilingan"],
    ["14","Ishlayotgan dastgoh soatlari","soat","138 845.6","145 236 x 0,956"],
    ["15","H1 — 1 soatlik unumdorlik","m/soat","13.551","formuladan"],
    ["16","H2 — 1 soatlik unumdorlik","m2/soat","22.956","formuladan"],
    ["17","H3 — 1 soatlik unumdorlik","arqoq/soat","18 971","formuladan"],
    ["18","H4 — 1 soatlik unumdorlik","m.arq/soat","34 893","formuladan"],
    ["19","Yalpi ishlab chiqarish (m)","ming metr","1 881.56","H1 x S_ish / 1000"],
    ["20","Yalpi ishlab chiqarish (m2)","ming m2","3 187.20","H2 x S_ish / 1000"],
    ["21","Yalpi (mln arqoq)","mln arq","2 634.10","H3 x S_ish / 10^6"],
    ["22","Yalpi (mln m.arq)","mln m.arq","4 844.89","H4 x S_ish / 10^6"],
    ["23","100 m uchun tanda ip sarfi","kg","7.911","n*Tt*100*k_t*(1+chiq_t)/10^6"],
    ["24","100 m uchun arqoq ip sarfi","kg","7.532","Pa*10*100*Bxt/100*Ta*k_a*(1+chiq_a)/10^6"],
    ["25","Yillik tanda ip ehtiyoji","tonna","148.85","kol19 x kol23 / 100"],
    ["26","Yillik arqoq ip ehtiyoji","tonna","141.71","kol19 x kol24 / 100"],
    ["27","Yillik ip ehtiyoji (jami)","tonna","290.56","kol25 + kol26"],
    ["28","Soatlik tanda ip ehtiyoji","kg/soat","35.87","kol25 x 1000 / yillik soat"],
    ["29","Soatlik arqoq ip ehtiyoji","kg/soat","34.15","kol26 x 1000 / yillik soat"],
    ["30","Soatlik jami ip ehtiyoji","kg/soat","70.02","kol28 + kol29"],
], widths=[400, 2800, 1200, 1200, 2900]))
add(empty_para())
add(para("Hisoblash izohlari:", bold=True, sz=22))
add(para("- O'rnatilgan dastgoh soatlari = 35 x (7,41 x 2 x 280) = 145 236,0 soat", sz=21))
add(para("- Ishlayotgan dastgoh soatlari = 145 236,0 x 0,956 = 138 845,6 soat", sz=21))
add(para("- Yalpi mahsulot (m) = 13,551 x 138 845,6 / 1 000 = 1 881,56 ming metr", sz=21))
add(para("- 100 m tanda sarfi: 2579 x 29 x 100 x 1,05 x 1,0074 / 10^6 = 7,911 kg", sz=21))
add(para("- 100 m arqoq sarfi: 1400 x 1,694 x 29 x 1,08 x 1,014 / 10^6 = 7,532 kg", sz=21))
add(para("- Yillik tanda: 1 881 560 m x 7,911 / 100 / 1 000 = 148,85 t", sz=21))
add(para("- Yillik arqoq: 1 881 560 m x 7,532 / 100 / 1 000 = 141,71 t", sz=21))
add(empty_para())

# ══ 22-JADVAL ══════════════════════════════════════════════════════
add(heading("III. HOM ASHYODAN FOYDALANISH REJASI HISOBI", 2))
add(para("Hom ashyo balansida hom ashyo miqdori va sifati, toqima ishlab chiqarishda chiqindilarning "
         "o'timlar boyicha miqdori hamda hom toqima ishlab chiqarish uchun olinadigan ipning qiymatlari aniqlanadi.",
         sz=22))
add(empty_para())
add(para("22-jadval. Toquv ishlab chiqarishida chiqindilarni o'timlar boyicha taqsimoti",
         bold=True, sz=22, align="center"))
add(tbl([
    ["Sexlar","Norm. chiqindi % (jami)","Chigal ip","Yumshoq ip 1-7m","Yumshoq ip 7-30m","Supurindi","Normalanmaydigan"],
    ["TANDA BOYICHA","","","","","",""],
    ["Qayta o'rash","0.003","0.0024","—","—","0.0006","0.003"],
    ["Tandalash","0.015","0.0135","0.01125","0.00225","—","0.015"],
    ["Ip boglash","0.042","0.042","0.0315","—","—","0.042"],
    ["Toquvchilik","0.680","0.476","—","0.204","—","0.680"],
    ["JAMI tanda:","0.740","0.534","0.043","0.206","0.0006","0.740"],
    ["ARQOQ BOYICHA","","","","","",""],
    ["Toquvchilik","1.40","0.98","—","—","0.42","1.40"],
    ["JAMI arqoq:","1.40","0.98","—","—","0.42","1.40"],
], widths=[1900, 1300, 1000, 1100, 1100, 900, 1100]))
add(empty_para())
add(para("Chiqindi miqdori hisobi:", bold=True, sz=22))
add(formula("Tanda chigal ip:    Q = 148,85 x 0,534 / 100 = 0,7949 t"))
add(formula("Arqoq chigal ip:   Q = 141,71 x 0,98  / 100 = 1,3888 t"))
add(formula("Jami chigal ip:     Q = 0,7949 + 1,3888 = 2,1837 t"))
add(formula("Arqoq supurindisi (0,42%):   Q = 141,71 x 0,42/100 = 0,5952 t"))
add(formula("Tanda 1-7m uzuq iplar (0,043%):   Q = 148,85 x 0,043/100 x 1000 = 64,0 kg"))
add(formula("Tanda 7-30m uzuq iplar (0,206%):  Q = 148,85 x 0,206/100 x 1000 = 306,6 kg"))
add(empty_para())
add(para("23-jadval. Toquv ishlab chiqarish uchun xom ashyo balansi", bold=True, sz=22, align="center"))
add(tbl([
    ["Nomlanishi","Miqdori (kg)","1 kg bahosi (so'm)","Jami bahosi (ming so'm)"],
    ["KIRIM:","","",""],
    ["Tanda ipi","148 850","52 000","7 740 200"],
    ["Arqoq ipi","141 710","52 000","7 368 920"],
    ["Jami kirim:","290 560","","15 109 120"],
    ["CHIQIM:","","",""],
    ["Hom toqima (ip)","275 214","53 460","14 711 594"],
    ["Chigal ip","2 183.7","15 600","34 066"],
    ["1-7 m uzuq iplar","64.0","10 400","666"],
    ["7-30 m uzuq iplar","306.6","8 320","2 551"],
    ["Supurindi","595.2","5 200","3 095"],
    ["Momiq va tokilganlar","196.5","2 600","511"],
    ["Jami chiqindi:","3 346.0","","40 889"],
    ["Jami chiqim:","278 560","","14 752 483"],
], widths=[2500, 1800, 1800, 2400]))
add(empty_para())

# ══ 24-JADVAL ══════════════════════════════════════════════════════
add(heading("IV. MEHNAT VA ISH HAQI REJASI HISOBI", 2))
add(para("24-jadval. Ish haqi fondi va shtatlar hisobi", bold=True, sz=22, align="center"))
add(tbl([
    ["No","Kasb nomi","Guruh","Xizmat norma","I smena","II smena","Jami","Ish soat/ kun","Tarif stavkasi (so'm/s)","Mukofot %","Bir kunlik IH (so'm)","Mukofot (so'm)","Jami (so'm)"],
    ["1","Toquvchi","a","5 dastgoh","7","6","13","96.33","6 847","50%","659 627","329 814","989 441"],
    ["2","Yordamchi usta","x","25 dastgoh","2","2","4","29.64","7 532","50%","223 148","111 574","334 722"],
    ["3","Qirquvchi","a","40 dastgoh","1","1","2","14.82","5 870","40%","86 993","34 797","121 790"],
    ["4","Instruktor","a","100 dastgoh","1","—","1","7.41","6 847","40%","50 736","20 294","71 030"],
    ["5","Farrosh","a","300 m2","2","2","4","29.64","6 458","40%","191 415","76 566","267 981"],
    ["6","Ta'mirlovchi","x","40 dastgoh","1","1","2","14.82","6 975","60%","103 370","62 022","165 392"],
    ["7","Yuk tashuvchi","t","< 200 kg/s","1","1","2","14.82","6 064","40%","89 869","35 948","125 817"],
    ["","Toquvchilik jami:","","","15","13","28","207.48","","","1 405 158","671 015","2 076 173"],
    ["","Tayyorlov (70%):","","","11","9","20","","","","983 611","469 711","1 453 322"],
    ["","FABRIKA JAMI:","","","26","22","48","","","","2 388 769","1 140 726","3 529 495"],
], widths=[300, 1400, 500, 900, 550, 550, 500, 750, 1000, 650, 1200, 1000, 1100]))
add(empty_para())
add(para("Izohlar:", bold=True, sz=22))
add(para("- Bir kunlik IH = ishchilar soni x smena davomiyligi (7,41 soat) x tarif stavkasi", sz=21))
add(para("- Mukofot = Bir kunlik IH x mukofot foizi", sz=21))
add(para("- Tayyorlov bolimi = toquvchilik bolimi ishchilar sonining 70 foizi", sz=21))
add(para("- Yuk tashuvchi: soatlik ip ehtiyoji 70,02 kg/soat < 200 kg => 1 ta/smena x 2 = 2 ta", sz=21))
add(empty_para())

# ══ 25-JADVAL ══════════════════════════════════════════════════════
add(para("25-jadval. Guruhlar boyicha ish haqi fondi — jamlanma jadval", bold=True, sz=22, align="center"))
add(tbl([
    ["O'tim nomi","Bir kunlik IH (so'm)","Ish kunlari","Yillik soatbay IH (ming so'm)","Asosiy IH (ming so'm)","Qo'shimcha 20%","Oylik IH (ming so'm)","Ijtimoiy to'lov (ming so'm)"],
    ["ASOSIY ISHCHILAR","","","","","","",""],
    ["Tayyorlov bolimi","983 611","280","275 411","275 411","55 082","330 493","66 099"],
    ["Toquvchilik","1 405 158","280","393 444","393 444","78 689","472 133","94 427"],
    ["Jami:","2 388 769","280","668 855","668 855","133 771","802 626","160 526"],
    ["USKUNA XIZMAT ISHCHILARI","","","","","","",""],
    ["Tayyorlov bolimi","351 289","280","98 361","98 361","19 672","118 033","23 607"],
    ["Toquvchilik","501 892","280","140 530","140 530","28 106","168 636","33 727"],
    ["Jami:","853 181","280","238 891","238 891","47 778","286 669","57 334"],
    ["TRANSPORT ISHCHILARI","","","","","","",""],
    ["Tayyorlov bolimi","130 400","280","36 512","36 512","7 302","43 814","8 763"],
    ["Toquvchilik","186 286","280","52 160","52 160","10 432","62 592","12 518"],
    ["Jami:","316 686","280","88 672","88 672","17 734","106 406","21 281"],
    ["FABRIKA BOYICHA JAMI:","3 558 636","280","996 418","996 418","199 283","1 195 701","239 141"],
], widths=[1700, 1200, 700, 1400, 1400, 1000, 1300, 1300]))
add(empty_para())

# ══ 26-JADVAL ══════════════════════════════════════════════════════
add(para("26-jadval. Ish haqi jamlanma jadvali", bold=True, sz=22, align="center"))
add(tbl([
    ["No","Ish haqi fondlarining nomlanishi","Bir kunlik soatbay IH (so'm)","Yillik ish kunlari","Bir yildagi soatbay IH (ming so'm)"],
    ["1","Vaqtbay ish haqi","2 388 769","280","668 855.3"],
    ["2","Vaqtbay-mukofot","1 140 726","280","319 403.3"],
    ["J1","Jami soatbay ish haqi fondi","3 529 495","280","988 258.6"],
    ["3","Smena toxtashlari qo'shimchasi 1,5%","52 942","280","14 823.8"],
    ["J2","Jami kunlik ish haqi fondi","3 582 437","280","1 003 082.4"],
    ["4","Davlat buyurtmalari uchun qo'shimcha 0,5%","17 912","280","5 015.4"],
    ["5","Mehnat ta'tiliga qo'shimcha 8,5%","304 507","280","85 262.0"],
    ["J3","Jami oylik ish haqi fondi","3 904 856","280","1 093 359.8"],
    ["6","Yagona ijtimoiy to'lov 25%","976 214","280","273 340.0"],
], widths=[600, 3200, 1500, 800, 1400]))
add(empty_para())
add(para("Izohlar:", bold=True, sz=22))
add(para("- J1 = vaqtbay + mukofot = 668 855,3 + 319 403,3 = 988 258,6 ming so'm", sz=21))
add(para("- J2 = J1 + toxtash 1,5% = 988 258,6 + 14 823,8 = 1 003 082,4 ming so'm", sz=21))
add(para("- J3 = J2 + davlat 0,5% + ta'til 8,5% = 1 003 082,4 + 5 015,4 + 85 262,0 = 1 093 359,8 ming so'm", sz=21))
add(para("- Ijtimoiy to'lov = J3 x 25% = 1 093 359,8 x 25% = 273 340,0 ming so'm", sz=21))
add(empty_para())

# ══ MEHNAT KO'RSATKICHLARI ══════════════════════════════════════════
add(heading("V. MEHNAT BOYICHA TEXNIK-IQTISODIY KORSATKICHLAR", 2))
add(para("Mehnat unumdorligi (m/ishchi-soat) — formula 59:", sz=22))
add(formula("Pi_m = Bm / O_ch = 1 881 560 / (48 x 4 149,6) = 1 881 560 / 199 181 = 9,45 m/ishchi-soat"))
add(empty_para())
add(para("Mehnat unumdorligi (arqoq/ishchi-soat) — formula 61:", sz=22))
add(formula("Pi_arq = 2 634 100 000 / 199 181 = 13 224 arqoq/ishchi-soat"))
add(empty_para())
add(para("100 ta dastgoh uchun ishchi sarfi — formula 62:", sz=22))
add(formula("Pi_100 = (48 / 35) x 100 = 137,14 kishi/100 dastgoh"))
add(empty_para())
add(para("O'rtacha oylik ish haqi — formula 63:", sz=22))
add(formula("Z_oy = J3 x 1000 / (Ch_sp x 12) = 1 093 359 800 / (48 x 12) = 1 898 194 so'm/oy"))
add(empty_para())

# ══ 27-JADVAL ══════════════════════════════════════════════════════
add(heading("VI. TOQUVCHILIKDA MAHSULOT TANNARXI HISOBI", 2))
add(para("27-jadval. Uskunalar qiymati va yemirilish ajratmalari hisobi", bold=True, sz=22, align="center"))
add(tbl([
    ["Uskuna nomi va markasi","Soni","1 ta bahosi (ming so'm)","Umumiy qiymati (ming so'm)","Montaj 10% (ming so'm)","O'rnatish bilan (ming so'm)","Yemirish %","Yemirish (ming so'm)"],
    ["Sultzer Ruti P7200","35","850 000","29 750 000","2 975 000","32 725 000","10","3 272 500"],
    ["Qayta o'rash mashinasi","2","145 220","290 440","29 044","319 484","10","31 948"],
    ["Tandalash mashinasi","1","800 000","800 000","80 000","880 000","10","88 000"],
    ["Ip boglash mashinasi","1","101 600","101 600","10 160","111 760","10","11 176"],
    ["Ip o'tkazish mashinasi","1","31 750","31 750","3 175","34 925","10","3 493"],
    ["Saralash-hisoblash","1","45 000","45 000","4 500","49 500","10","4 950"],
    ["O'lchash mashinasi","1","50 000","50 000","5 000","55 000","10","5 500"],
    ["JAMI:","42","","31 068 790","3 106 879","34 175 669","10","3 417 567"],
], widths=[2200, 500, 1300, 1500, 1200, 1500, 700, 1200]))
add(empty_para())
add(para("Bino va inshootlar yemirilishi:", bold=True, sz=22))
add(para("- Sex binosi:  2 520 m2 x 1 000 ming so'm/m2 = 2 520 000 ming so'm", sz=22))
add(para("  (35/30 x 2 160 = 2 520 m2 — loyiha chizmasidan proporsional)", sz=20))
add(para("- Ma'muriy bino:  504 m2 x 850 ming so'm = 428 400 ming so'm  (504 = 2 520 x 20%)", sz=22))
add(para("- Bino jami qiymati = 2 520 000 + 428 400 = 2 948 400 ming so'm", sz=22))
add(para("- Bino yemirilishi = 2 948 400 x 5% = 147 420 ming so'm", sz=22))
add(para("- Transport yemirilishi = 3 417 567 x 5% = 170 878 ming so'm", sz=22))
add(empty_para())

add(para("28-jadval. Asosiy fondlar yemirilishi yakuniy jadvali", bold=True, sz=22, align="center"))
add(tbl([
    ["No","Asosiy fondlar nomlanishi","Umumiy qiymati (ming so'm)","Yemirilish (ming so'm)"],
    ["1","Uskunalar yemirilishi","34 175 669","3 417 567"],
    ["2","Bino va inshootlar yemirilishi","2 948 400","147 420"],
    ["3","Transport vositalari yemirilishi","—","170 878"],
    ["","Yemirilish ajratmalarining umumiy qiymati:","","3 735 865"],
], widths=[500, 3800, 2000, 1700]))
add(empty_para())
add(para("Boshqa harajatlar hisobi:", bold=True, sz=22))
add(para("1. Joriy ta'mirlash va saqlash (1%):  31 068 790 x 1% = 310 688 ming so'm", sz=22))
add(para("2. O'rta va kapital ta'mirlash (2%):  31 068 790 x 2% = 621 376 ming so'm", sz=22))
add(para("3. Atrof muhit muhofazasi (0,2% binodan):  2 948 400 x 0,2% = 5 897 ming so'm", sz=22))
add(para("4. Texnika xavfsizligi (48 x 60 000) = 2 880 ming so'm", sz=22))
add(para("5. Mehnat muhofazasi (48 x 80 000) = 3 840 ming so'm", sz=22))
add(para("6. Izlanishlar va loyihalar (35 x 100 000) = 3 500 ming so'm", sz=22))
add(empty_para())

add(para("29-jadval. Boshqa ishlab chiqarish harajatlarining yakuniy jadvali", bold=True, sz=22, align="center"))
add(tbl([
    ["No","Harajatlarning nomlanishi","Harajatlar qiymati (ming so'm)"],
    ["1","Joriy ta'mirlash va uskunalarni saqlash","310 688"],
    ["2","Uskunalarni o'rta va kapital ta'mirlash","621 376"],
    ["3","Atrof muhit muhofazasi","5 897"],
    ["4","Texnika xavfsizligi","2 880"],
    ["5","Mehnat muhofazasi","3 840"],
    ["6","Izlanishlar va loyihalar, ratsionallashtirish","3 500"],
    ["","JAMI:","948 181"],
], widths=[500, 4800, 2700]))
add(empty_para())

add(para("30-jadval. Mahsulot tannarxi hisobi", bold=True, sz=22, align="center"))
add(tbl([
    ["No","Harajat turlari","Umumiy tannarx (ming so'm)","1 m xom to'qima tannarxi (so'm)","Jamiga nisbatan %"],
    ["1","Moddiy ishlab chiqarish harajatlari","16 015 722","8 512.8","71.96"],
    ["2","Ishlovchilar mehnatiga haq to'lash","1 257 364","668.3","5.65"],
    ["3","Yagona ijtimoiy to'lov","314 341","167.1","1.41"],
    ["4","Asosiy fondlar yemirilishi","3 735 865","1 985.6","16.80"],
    ["5","Ishlab chiqarish boshqa harajatlari","948 181","504.0","4.26"],
    ["","JAMI tannarx:","22 271 473","11 837.8","100"],
], widths=[500, 2900, 1800, 1800, 1000]))
add(empty_para())
add(para("Hisoblash izohlari:", bold=True, sz=22))
add(para("I.  Moddiy harajatlar: xom ashyo 15 109 120 + 6% qo'shimcha 906 547 = 16 015 667 ming so'm", sz=21))
add(para("II. Ish haqi: J3 (1 093 360) + sex xodimlari 15% (164 004) = 1 257 364 ming so'm", sz=21))
add(para("III.Ijtimoiy to'lov: 1 257 364 x 25% = 314 341 ming so'm", sz=21))
add(para("IV. Yemirilish: 28-jadvaldan = 3 735 865 ming so'm", sz=21))
add(para("V.  Boshqa: 29-jadvaldan = 948 181 ming so'm", sz=21))
add(formula("Umumiy tannarx = 16 015 722 + 1 257 364 + 314 341 + 3 735 865 + 948 181 = 22 271 473 ming so'm"))
add(formula("1 m tannarxi = 22 271 473 000 / 1 881 560 = 11 837,8 so'm/m"))
add(empty_para())

# ══ 31-JADVAL ══════════════════════════════════════════════════════
add(heading("VII. FABRIKA BOYICHA SOTISH REJASI VA SAMARADORLIK", 2))
add(para("31-jadval. Fabrika boyicha sotish rejasi va samaradorlik hisobi", bold=True, sz=22, align="center"))
add(tbl([
    ["To'qima nomi","Mahsulot hajmi (m)","Tannarx 1m (so'm)","Tannarx jami (ming so'm)","Baho 1m (so'm)","Baho jami (ming so'm)","Foyda (ming so'm)","Rentabellik %"],
    ["Flanel Art.1660","1 881 560","11 837.8","22 271 473","13 613.5","25 612 145","3 340 672","15"],
], widths=[1300, 1000, 1000, 1600, 1000, 1600, 1300, 700]))
add(empty_para())
add(para("Hisoblash izohlari:", bold=True, sz=22))
add(formula("1. Bir metr bahosi:  B_1m = T_1m x (1 + R%) = 11 837,8 x 1,15 = 13 613,5 so'm/m"))
add(formula("2. Baho jami:  13 613,5 x 1 881 560 = 25 612 145 ming so'm"))
add(formula("3. Korxona foydasi:  F = S - T = 25 612 145 - 22 271 473 = 3 340 672 ming so'm"))
add(formula("4. Rentabellik:  R = F/T x 100 = 3 340 672 / 22 271 473 x 100 = 15,00%"))
add(formula("5. 1 so'mlik tovar uchun harajat:  22 271 473 / 25 612 145 = 0,87 so'm/so'm"))
add(formula("6. Kapital:  K = 34 175 669 + 2 948 400 = 37 124 069 ming so'm"))
add(formula("7. Qoplanish muddati:  T = 37 124 069 / 3 340 672 = 11,1 yil"))
add(empty_para())

# ══ DAVR HARAJATLARI ═══════════════════════════════════════════════
add(heading("VIII. DAVR HARAJATLARI HISOBI", 2))
add(para("Davr harajatlarini elementlar boyicha jamlanma jadvali", bold=True, sz=22, align="center"))
add(tbl([
    ["No","Harajatlar nomi","Foizi","Miqdori (ming so'm)"],
    ["1","Boshqaruv rahbarlari ish haqi","25%","314 341"],
    ["2","Devonxona va bogliq harajatlar","13%","163 457"],
    ["3","Xizmat safari harajatlari","16%","201 178"],
    ["4","Ma'muriy binoni ta'mirlash","17%","213 752"],
    ["5","Umum korxona laboratoriya harajatlari","12%","150 840"],
    ["6","Ixtiro va loyihalash harajatlari","10%","125 700"],
    ["7","Boshqa umumxo'jalik harajatlari","7%","87 990"],
    ["","Jami davr harajatlari:","100%","1 257 258"],
    ["8a","Mulk solig'i 2%:  37 124 069 x 2%","","742 481"],
    ["8b","Suvga to'lov (50 so'm/m):  1 881 560 x 50/1000","","94 078"],
    ["8v","Yer solig'i (6 500 ming so'm/ga):  0,252 ga x 6 500","","1 638"],
    ["8g","Yo'l fondi 1%:  25 612 145 x 1%","","256 121"],
    ["","Davr harajatlari hammasi:","","2 351 576"],
], widths=[600, 3700, 1000, 1700]))
add(empty_para())

# ══ MOLIYAVIY KO'RSATKICHLAR ════════════════════════════════════════
add(heading("IX. KORXONANING MOLIYAVIY KORSATKICHLARI", 2))
add(tbl([
    ["No","Korsatkichlar nomi","Miqdori (ming so'm)","Formula"],
    ["1","Fabrika boyicha mahsulot tannarxi","22 271 473","30-jadvaldan"],
    ["2","Mahsulotni sotish hajmi (bahosi)","25 612 145","31-jadvaldan"],
    ["3","Mahsulot sotishdan yalpi foyda","3 340 672","(2-1)"],
    ["4","Davr harajatlari — hammasi","2 351 576","jadvaldan"],
    ["5","Asosiy faoliyatdan foyda","989 096","(3-4)"],
    ["6","Foydadan tolanadigan soliq — 15%","148 364","5 x 15%"],
    ["7","Soliq tolangandan song foyda","840 732","5 - 6"],
    ["8","Boshqa tolovlar va majburiyatlar — 4%","33 629","7 x 4%"],
    ["9","Korxona ixtiyoridagi foyda","807 103","7 - 8"],
    ["10","Korxona zaxira fondiga ajratma — 5%","40 355","9 x 5%"],
], widths=[500, 3500, 1600, 2400]))
add(empty_para())

# ══ YAKUNIY KO'RSATKICHLAR ══════════════════════════════════════════
add(heading("X. TOQUV FABRIKASINING TEXNIK-IQTISODIY KORSATKICHLARI", 2))
add(tbl([
    ["No","Korsatkichlar nomi","Olchov birligi","Korsatkichlar"],
    ["1","Toqima nomi","—","Flanel, Art. 1660"],
    ["2","O'rnatilgan dastgohlar soni","dona","35"],
    ["3","Dastgoh markasi","—","Sultzer Ruti P7200"],
    ["4","Rejali toxtashlar foizi","%","4.4"],
    ["5","Uskunalar unumdorligi","m/soat","13.551"],
    ["6","Uskunalar unumdorligi","arqoq/soat","18 971"],
    ["7","Bir yilda ishlab chiqarilgan mahsulot","ming metr","1 881.56"],
    ["8","Mehnat unumdorligi","m/ishchi-soat","9.45"],
    ["9","Mehnat unumdorligi","arqoq/ishchi-soat","13 224"],
    ["10","O'rtacha oylik ish haqi","so'm","1 898 194"],
    ["11","Royxatdagi ishchilar soni","kishi","48"],
    ["12","Fabrika boyicha mahsulot tannarxi","ming so'm","22 271 473"],
    ["13","1 m toqima tannarxi","so'm","11 837.8"],
    ["14","Sotish hajmi (jami baho)","ming so'm","25 612 145"],
    ["15","1 m toqima bahosi","so'm","13 613.5"],
    ["16","Fabrika boyicha foyda","ming so'm","3 340 672"],
    ["17","Mahsulot samaradorligi (rentabellik)","%","15"],
    ["18","Asosiy ishlab chiqarish fondlari","ming so'm","37 124 069"],
    ["19","Kapital mablaglarni qoplanish muddati","yil","11.1"],
], widths=[500, 3500, 1700, 2300]))
add(empty_para())
add(empty_para())

add(para("XULOSA", bold=True, sz=28, align="center"))
add(para(
    "Mazkur hisob-kitoblar asosida \"Flanel\" (Artikul 1660) toqimasini ishlab chiqaruvchi "
    "35 ta Sultzer Ruti P7200 markali toquv dastgohi uchun quyidagi texnik-iqtisodiy "
    "korsatkichlar aniqlandi:",
    sz=22))
add(para("- Yillik ishlab chiqarish hajmi:    1 881,56 ming metr", sz=22))
add(para("- Bir dastgohning soatlik unumdorligi:  13,551 m/soat", sz=22))
add(para("- Jami ishchilar soni:  48 kishi", sz=22))
add(para("- 1 m xom toqima tannarxi:  11 837,8 so'm", sz=22))
add(para("- 1 m toqima sotish bahosi:  13 613,5 so'm", sz=22))
add(para("- Yillik foyda:  3 340 672 ming so'm", sz=22))
add(para("- Rentabellik:  15%", sz=22))
add(para("- Kapital mablaglar qoplanish muddati:  11,1 yil", sz=22))
add(empty_para())
add(para(
    "Hisob-kitoblar \"Iqtisod qismini ishlanishi namuna\" metodologiyasi asosida, "
    "I. Yarashov kurs loyihasining texnik malumotlari boyicha qayta ishlandi.",
    sz=20, align="both"))

# ════════════════════════════════════════════════════════════════════
# ASSEMBLE ZIP
# ════════════════════════════════════════════════════════════════════
body_content = "\n".join(P)

document_xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document
  xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"
  xmlns:mc="{MC}"
  xmlns:r="{RNS}"
  xmlns:w="{WNS}"
  mc:Ignorable="w14">
<w:body>
{body_content}
<w:sectPr>
  <w:pgSz w:w="12240" w:h="15840"/>
  <w:pgMar w:top="1134" w:right="850" w:bottom="1134" w:left="1701"
           w:header="709" w:footer="709" w:gutter="0"/>
</w:sectPr>
</w:body>
</w:document>"""

styles_xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="{WNS}"
          xmlns:r="{RNS}">
<w:docDefaults>
  <w:rPrDefault>
    <w:rPr>
      <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>
      <w:sz w:val="24"/><w:szCs w:val="24"/>
    </w:rPr>
  </w:rPrDefault>
</w:docDefaults>
</w:styles>"""

doc_rels = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="{RNS}/styles"
    Target="styles.xml"/>
</Relationships>"""

pkg_rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
    Target="word/document.xml"/>
</Relationships>"""

content_types = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels"
    ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml"
    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml"
    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>"""

with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED) as z:
    z.writestr("[Content_Types].xml", content_types)
    z.writestr("_rels/.rels",          pkg_rels)
    z.writestr("word/document.xml",    document_xml)
    z.writestr("word/styles.xml",      styles_xml)
    z.writestr("word/_rels/document.xml.rels", doc_rels)

size = os.path.getsize(OUT)
print(f"OK — {OUT}  ({size:,} bytes)")
