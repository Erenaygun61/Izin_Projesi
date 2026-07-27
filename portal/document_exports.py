from io import BytesIO
from pathlib import Path
from xml.sax.saxutils import escape
from decimal import Decimal, InvalidOperation
import re

from openpyxl import Workbook
from openpyxl.drawing.image import Image as ExcelImage
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import Image as PdfImage
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle
from .models import _ad_soyad_bicimi


STATIC_PORTAL = Path(__file__).resolve().parent.parent / "static" / "portal"
BELEDIYE_LOGO = STATIC_PORTAL / "trabzon-buyuksehir-logo.jpg"
TISKI_LOGO = STATIC_PORTAL / "logo1.png"
ARIAL = Path(r"C:\Windows\Fonts\arial.ttf")
ARIAL_BOLD = Path(r"C:\Windows\Fonts\arialbd.ttf")


def _fontlar():
    """Türkçe karakterleri PDF'de eksiksiz gösterecek yazı tiplerini kaydeder."""
    regular, bold = "Helvetica", "Helvetica-Bold"
    try:
        if "PortalArial" not in pdfmetrics.getRegisteredFontNames():
            pdfmetrics.registerFont(TTFont("PortalArial", str(ARIAL)))
        if "PortalArialBold" not in pdfmetrics.getRegisteredFontNames():
            pdfmetrics.registerFont(TTFont("PortalArialBold", str(ARIAL_BOLD)))
        regular, bold = "PortalArial", "PortalArialBold"
    except Exception:
        # Yerel Windows yazı tipi bulunamazsa belge yine oluşturulsun.
        pass
    return regular, bold


def _metin(value, fallback="-"):
    if value is None:
        return fallback
    value = str(value).strip()
    return value or fallback


def _tarih(value):
    return value.strftime("%d/%m/%Y") if value else "-"


def _izin_gunu(value):
    """Tam günlerde ,00 göstermeden, kesirli günleri koruyarak yazar."""
    if value is None or value == "":
        return "-"
    try:
        return format(Decimal(str(value)).normalize(), "f").replace(".", ",")
    except (InvalidOperation, ValueError, TypeError):
        return _metin(value)


def _telefon_bicimi(value):
    """Cep numarasını PDF'de okunur standart biçimde gösterir."""
    raw = _metin(value, "")
    rakamlar = re.sub(r"\D", "", raw)
    if len(rakamlar) == 11 and rakamlar.startswith("0"):
        rakamlar = rakamlar[1:]
    if len(rakamlar) == 10:
        return f"({rakamlar[:3]}) {rakamlar[3:6]}-{rakamlar[6:8]}-{rakamlar[8:]}"
    return raw or "-"


def _kullanici_adi(user):
    if not user:
        return "Tanımlanmamış"
    return _ad_soyad_bicimi(user.get_full_name().strip() or user.username)


def _kullanici_parafladi_mi(talep, user):
    if not talep or not user:
        return False
    return talep.islem_gecmisi.filter(
        kullanici=user,
        islem__in=["IMZA", "KONTROL_PARAFI", "ONAY"],
    ).exists()


def _izin_aciklamasi(talep):
    izin_turu = talep.izin_turu
    if izin_turu and izin_turu.yillik_izin_mi:
        return (
            f"{_metin(talep.ait_oldugu_yil)} yılı yıllık iznimden "
            f"{_izin_gunu(talep.kullanilan_gun)} gün izin kullanmak istiyorum."
        )
    return f"{_metin(getattr(izin_turu, 'ad', None), 'İzin')} kullanmak istiyorum."


def _birim_satirlari(personel):
    ust_birim = _metin(getattr(personel, "gorev_birimi", ""), "")
    alt_birim = _metin(getattr(personel, "alt_birimi", ""), "")
    return ust_birim, alt_birim


def _belge_gorevi(personel):
    """Belgede işçi/şirket için statü, memur için özlük unvanı gösterilir."""
    if not personel:
        return "-"
    tip = str(getattr(personel, "personel_tipi", "") or "").upper()
    sinif = personel.__class__.__name__.lower()
    if tip == "SIRKET" or "sirket" in sinif:
        return "Şirket Personeli"
    if tip == "ISCI" or "isci" in sinif:
        return "Kadrolu İşçi"
    return _metin(getattr(personel, "unvani", "") or getattr(personel, "unvan", "") or getattr(personel, "gorevi", ""))


def _personel_tipi_etiketi(talep):
    """Belgedeki statüyü okunur başlık biçiminde yazar."""
    return {
        "MEMUR": "Memur",
        "ISCI": "İşçi",
        "SIRKET": "Şirket",
    }.get(str(getattr(talep, "personel_tipi", "") or "").upper(), _metin(talep.get_personel_tipi_display()))


def _p(value, style):
    return Paragraph(escape(_metin(value)), style)


def izin_pdf_olustur(talep, personel, imzacilar, goreve_baslama):
    cep_telefonu = getattr(talep, "cep_telefonu", "") or getattr(personel, "cep_tel", "")
    """İzin isteği onayını tek sayfalık, indirilebilir PDF olarak üretir."""
    regular, bold = _fontlar()
    kurum_lacivert = colors.HexColor("#123B59")
    kurum_turkuaz = colors.HexColor("#0A8C91")
    etiket_zemini = colors.HexColor("#EAF3F4")
    yumusak_zemin = colors.HexColor("#F7FAFB")
    buffer = BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        leftMargin=15 * mm,
        rightMargin=15 * mm,
        topMargin=11 * mm,
        bottomMargin=12 * mm,
        title=f"İzin İsteği - {talep.evrak_no}",
        author="TİSKİ İzin Portalı",
    )
    styles = getSampleStyleSheet()
    title = ParagraphStyle("portal_title", parent=styles["Normal"], fontName=bold, fontSize=13, leading=15, alignment=TA_CENTER, textColor=kurum_lacivert)
    subtitle = ParagraphStyle("portal_subtitle", parent=styles["Normal"], fontName=bold, fontSize=13, leading=15, alignment=TA_CENTER, textColor=kurum_lacivert)
    cell = ParagraphStyle("portal_cell", parent=styles["Normal"], fontName=regular, fontSize=10.5, leading=12.5)
    cell_center = ParagraphStyle("portal_cell_center", parent=cell, alignment=TA_CENTER)
    cell_value = ParagraphStyle("portal_cell_value", parent=cell, alignment=TA_LEFT, leftIndent=8)
    cell_bold = ParagraphStyle("portal_cell_bold", parent=cell, fontName=bold)
    birim_bold = ParagraphStyle("portal_birim_bold", parent=cell_center, fontName=bold, leading=14)
    birim_normal = ParagraphStyle("portal_birim_normal", parent=cell_center, leading=14)
    statu_style = ParagraphStyle("portal_statu", parent=cell, fontSize=9.2, leading=11, alignment=TA_LEFT)
    tiny_left = ParagraphStyle("portal_tiny_left", parent=cell, fontSize=9.8, leading=11.5, alignment=TA_LEFT)
    tiny_center = ParagraphStyle("portal_tiny_center", parent=cell_center, fontSize=10.8, leading=12.5)
    sign_center = ParagraphStyle("portal_sign_center", parent=tiny_center, fontName=bold)
    foot = ParagraphStyle("portal_foot", parent=cell, fontSize=9.4, leading=11)
    story = []

    left_logo = PdfImage(str(BELEDIYE_LOGO), width=19 * mm, height=19 * mm)
    # TİSKİ kaynak görselinin çevresinde boşluk bulunduğu için, iki logo
    # belgede gözle aynı büyüklükte görünsün diye sağ antet biraz büyütülür.
    right_logo = PdfImage(str(TISKI_LOGO), width=23 * mm, height=23 * mm)
    antet = Table(
        [[
            left_logo,
            Paragraph(
                "T.C.<br/><b>TRABZON BÜYÜKŞEHİR BELEDİYESİ</b><br/>"
                "Trabzon İçmesuyu ve Kanalizasyon İdaresi Genel Müdürlüğü",
                title,
            ),
            right_logo,
        ]],
        colWidths=[27 * mm, 126 * mm, 27 * mm],
    )
    antet.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ALIGN", (0, 0), (0, 0), "LEFT"),
        ("ALIGN", (2, 0), (2, 0), "RIGHT"),
        ("LEFTPADDING", (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (-1, -1), 0),
        ("TOPPADDING", (0, 0), (-1, -1), 0),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
    ]))
    story += [antet, Spacer(1, 3 * mm)]

    # Tarih/sayı soldadır; birim adı antette tekrarlanmaz.
    story.append(Paragraph(
        f"<b>Evrak Tarih ve Sayısı:</b> {_tarih(talep.kurum_kayit_tarihi)} - E.{escape(_metin(talep.evrak_no))}",
        ParagraphStyle("doc_number", parent=cell_bold, alignment=TA_LEFT, fontSize=9.4),
    ))
    story += [Spacer(1, 2 * mm), Paragraph("İZİN İSTEĞİ ONAYI", subtitle), Spacer(1, 3 * mm)]

    ust_birim, alt_birim = _birim_satirlari(personel)
    # Birim adları hücre hizasından bağımsız olarak paragraf seviyesinde ortalanır.
    birim_satirlari = [[Paragraph(escape(ust_birim), birim_bold)]]
    if alt_birim:
        birim_satirlari.append([Paragraph(escape(alt_birim), birim_normal)])
    birim_degeri = Table(birim_satirlari, colWidths=[151 * mm])
    birim_degeri.setStyle(TableStyle([
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 2.8),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2.8),
        # Daire ve alt birimi ayıran ince çizgi korunur.
        ("LINEBELOW", (0, 0), (0, 0), 0.65, colors.black),
    ]))
    izin_adi = _metin(getattr(talep.izin_turu, "ad", None), "İzin")
    rows = [
        [_p("Birimi", cell_bold), birim_degeri, "", ""],
        [_p("Adı Soyadı", cell_bold), _p(_ad_soyad_bicimi(personel.ad_soyad), cell_value), _p("Ait Olduğu Yıl", cell_bold), _p(talep.ait_oldugu_yil, cell_value)],
        [_p("Görevi", cell_bold), _p(_belge_gorevi(personel), cell_value), _p("İzin Başlangıç Tarihi", cell_bold), _p(_tarih(talep.baslangic_tarihi), cell_value)],
        [_p("T.C. Kimlik No", cell_bold), _p(personel.tc_no, cell_value), _p("İzin Bitiş Tarihi", cell_bold), _p(_tarih(talep.bitis_tarihi), cell_value)],
        [_p("Kurum Sicil No", cell_bold), _p(personel.sicil_no, cell_value), _p("Göreve Başlama Tarihi", cell_bold), _p(_tarih(goreve_baslama), cell_value)],
        [_p("İzin Türü", cell_bold), _p(izin_adi, cell_value), _p("Toplam İzin", cell_bold), _p(_izin_gunu(talep.izin_hakki), cell_value)],
        [_p("Talep Ettiği İzin Süresi", cell_bold), _p(_izin_gunu(talep.kullanilan_gun), cell_value), _p("Kalan İzin Süresi", cell_bold), _p(_izin_gunu(talep.kalan_izin), cell_value)],
    ]
    info_table = Table(rows, colWidths=[29 * mm, 63 * mm, 36 * mm, 52 * mm])
    info_table.setStyle(TableStyle([
        ("BOX", (0, 0), (-1, -1), 1.4, colors.black),
        ("INNERGRID", (0, 0), (-1, -1), 0.95, colors.black),
        ("SPAN", (1, 0), (3, 0)),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ALIGN", (1, 0), (3, 0), "LEFT"),
        ("LEFTPADDING", (0, 0), (-1, -1), 5),
        ("RIGHTPADDING", (0, 0), (-1, -1), 5),
        ("LEFTPADDING", (1, 0), (3, 0), 0),
        ("RIGHTPADDING", (1, 0), (3, 0), 0),
        ("TOPPADDING", (0, 0), (-1, -1), 4.5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4.5),
    ]))
    story.append(info_table)

    note_html = (
        f'<font name="{bold}">Açıklama:</font><br/>'
        + escape(_izin_aciklamasi(talep))
        + "<br/><br/>Gereğini arz ederim."
    )
    note = Table([
        [Paragraph(note_html, cell)],
        [Paragraph(f'<font name="{bold}">Statü:</font> {escape(_personel_tipi_etiketi(talep))}', statu_style)],
    ], colWidths=[180 * mm])
    note.setStyle(TableStyle([
        ("BOX", (0, 0), (-1, -1), 1.25, colors.black),
        ("LINEABOVE", (0, 1), (-1, 1), 0.9, colors.black),
        ("LEFTPADDING", (0, 0), (-1, -1), 7),
        ("RIGHTPADDING", (0, 0), (-1, -1), 7),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
    ]))
    story.append(note)

    requester = Table([
        [Paragraph(f'<font name="{bold}">Cep Telefonu:</font> {escape(_telefon_bicimi(cep_telefonu))}', cell)],
    ], colWidths=[180 * mm])
    requester.setStyle(TableStyle([
        ("BOX", (0, 0), (-1, -1), 1.25, colors.black),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 7),
        ("RIGHTPADDING", (0, 0), (-1, -1), 7),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
    ]))
    story.append(requester)

    paraf_kutulari = [
        Paragraph(
            f"<b>İzin Talebinde Bulunan</b><br/><br/>"
            f"e-Paraflıdır<br/>{escape(_ad_soyad_bicimi(personel.ad_soyad))}<br/>{escape(_belge_gorevi(personel))}",
            sign_center,
        )
    ]
    if len(imzacilar) > 1:
        rol, user = imzacilar[0]
        paraf_kutulari.append(Paragraph(
            f"<b>Teklif Eden</b><br/><br/>"
            f"{'e-Paraflıdır' if talep.durum in {'IMZA_SURECINDE', 'ONAYLANDI'} or talep.imza_sirasi > 0 else 'Paraf Bekliyor'}<br/>"
            f"{escape(_kullanici_adi(user))}<br/>{escape(rol)}",
            sign_center,
        ))
    paraf_table = Table([paraf_kutulari], colWidths=[180 * mm / len(paraf_kutulari)] * len(paraf_kutulari), rowHeights=[52 * mm])
    paraf_table.setStyle(TableStyle([
        ("BOX", (0, 0), (-1, -1), 1.4, colors.black),
        ("INNERGRID", (0, 0), (-1, -1), 0.95, colors.black),
        ("BACKGROUND", (0, 0), (-1, -1), colors.white),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ALIGN", (0, 0), (-1, -1), "LEFT"),
    ]))
    story.append(paraf_table)

    if imzacilar:
        rol, user = imzacilar[-1]
        approval = Paragraph(
            f"<b>Onaylayan</b><br/><br/>{'e-İmzalıdır' if talep.durum == 'ONAYLANDI' or talep.imza_tamamlanma_tarihi else 'İmza Bekliyor'}<br/>"
            f"{escape(_kullanici_adi(user))}<br/>{escape(rol)}",
            sign_center,
        )
    else:
        approval = Paragraph("<b>Onaylayan</b><br/>Birimi için imzacı tanımlanmamış.", sign_center)
    approval_table = Table([[approval]], colWidths=[180 * mm], rowHeights=[48 * mm])
    approval_table.setStyle(TableStyle([
        ("BOX", (0, 0), (-1, -1), 1.4, colors.black),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ALIGN", (0, 0), (-1, -1), "LEFT"),
    ]))
    story.append(approval_table)
    story += [Spacer(1, 2 * mm), Paragraph(
        f"<b>Kontrol Eden</b><br/><b>Personel İzin Kontrolü:</b> {escape(_kullanici_adi(talep.kontrol_eden))}", foot
    )]
    doc.build(story)
    return buffer.getvalue()


def izin_excel_olustur(talep, personel, imzacilar, goreve_baslama):
    cep_telefonu = getattr(talep, "cep_telefonu", "") or getattr(personel, "cep_tel", "")
    """Aynı evrak verilerini düzenlenebilir Excel dosyası olarak üretir."""
    wb = Workbook()
    ws = wb.active
    ws.title = "İzin İsteği"
    ws.sheet_view.showGridLines = False
    for col, width in {"A": 23, "B": 29, "C": 24, "D": 28}.items():
        ws.column_dimensions[col].width = width
    for row in range(1, 30):
        ws.row_dimensions[row].height = 21
    thin = Side(style="thin", color="111111")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    heading = Font(name="Arial", size=12, bold=True)
    body = Font(name="Arial", size=10)
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left = Alignment(horizontal="left", vertical="center", wrap_text=True)

    try:
        image = ExcelImage(str(BELEDIYE_LOGO))
        image.width, image.height = 65, 65
        ws.add_image(image, "A1")
    except Exception:
        pass
    ws.merge_cells("B1:D1")
    ws["B1"] = "T.C.\nTRABZON BÜYÜKŞEHİR BELEDİYESİ\nTrabzon İçmesuyu ve Kanalizasyon İdaresi Genel Müdürlüğü"
    ws["B1"].font = heading
    ws["B1"].alignment = center
    ws.row_dimensions[1].height = 63
    ws.merge_cells("A3:D3")
    ws["A3"] = f"Evrak Tarih ve Sayısı: {_tarih(talep.kurum_kayit_tarihi)} - E.{talep.evrak_no}"
    ws["A3"].font = Font(name="Arial", size=10, bold=True)
    ws["A3"].alignment = left
    ws.merge_cells("A4:D4")
    ws["A4"] = "İZİN İSTEĞİ ONAYI"
    ws["A4"].font = heading
    ws["A4"].alignment = center

    ust_birim, alt_birim = _birim_satirlari(personel)
    values = [
        ("Birimi", f"{ust_birim}\n{alt_birim}".strip()),
        ("Adı Soyadı", _ad_soyad_bicimi(personel.ad_soyad), "Ait Olduğu Yıl", _metin(talep.ait_oldugu_yil)),
        ("Görevi", _belge_gorevi(personel), "İzin Başlangıç Tarihi", _tarih(talep.baslangic_tarihi)),
        ("T.C. Kimlik No", personel.tc_no, "İzin Bitiş Tarihi", _tarih(talep.bitis_tarihi)),
        ("Kurum Sicil No", personel.sicil_no, "Göreve Başlama Tarihi", _tarih(goreve_baslama)),
        ("İzin Türü", _metin(getattr(talep.izin_turu, "ad", None)), "Toplam İzin", _izin_gunu(talep.izin_hakki)),
        ("Talep Ettiği İzin Süresi", _izin_gunu(talep.kullanilan_gun), "Kalan İzin Süresi", _izin_gunu(talep.kalan_izin)),
    ]
    row = 6
    for value in values:
        if len(value) == 2:
            ws.cell(row, 1, value[0]); ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=4); ws.cell(row, 2, value[1])
            cells = [ws.cell(row, 1), ws.cell(row, 2)]
        else:
            for col, text in enumerate(value, 1):
                ws.cell(row, col, text)
            cells = [ws.cell(row, col) for col in range(1, 5)]
        for item in cells:
            item.font = body; item.border = border; item.alignment = left
        ws.cell(row, 1).font = Font(name="Arial", size=10, bold=True)
        ws.cell(row, 2).alignment = left
        if len(value) == 4:
            ws.cell(row, 3).font = Font(name="Arial", size=10, bold=True)
            ws.cell(row, 4).alignment = left
        row += 1
    ws.merge_cells(start_row=row, start_column=1, end_row=row + 3, end_column=4)
    ws.cell(row, 1, "Açıklama:\n" + _izin_aciklamasi(talep) + "\n\nGereğini arz ederim.")
    ws.cell(row, 1).font = body; ws.cell(row, 1).alignment = left; ws.cell(row, 1).border = border
    row += 5
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=4)
    ws.cell(row, 1, f"e-Paraflıdır\n{_ad_soyad_bicimi(personel.ad_soyad)}\n{_belge_gorevi(personel)}")
    ws.cell(row, 1).font = Font(name="Arial", size=10, bold=True); ws.cell(row, 1).alignment = center; ws.cell(row, 1).border = border
    row += 2
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=4)
    onaylayan = _kullanici_adi(imzacilar[-1][1]) if imzacilar else "Tanımlanmamış"
    rol = imzacilar[-1][0] if imzacilar else ""
    ws.cell(row, 1, f"Onaylayan\n{onaylayan}\n{rol}")
    ws.cell(row, 1).font = Font(name="Arial", size=10, bold=True); ws.cell(row, 1).alignment = center; ws.cell(row, 1).border = border
    row += 2
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=4)
    ws.cell(row, 1, f"Kontrol Eden: {_kullanici_adi(talep.kontrol_eden)}")
    ws.cell(row, 1).font = Font(name="Arial", size=9); ws.cell(row, 1).alignment = left

    buffer = BytesIO()
    wb.save(buffer)
    return buffer.getvalue()
