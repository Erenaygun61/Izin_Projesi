from io import BytesIO
from pathlib import Path
from xml.sax.saxutils import escape
from decimal import Decimal, InvalidOperation

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


def _kullanici_adi(user):
    if not user:
        return "Tanımlanmamış"
    return _ad_soyad_bicimi(user.get_full_name().strip() or user.username)


def _izin_aciklamasi(talep):
    izin_turu = talep.izin_turu
    if izin_turu and izin_turu.yillik_izin_mi:
        return (
            f"Özel işlerimden dolayı {_metin(talep.ait_oldugu_yil)} yılı yıllık iznimden "
            f"{_izin_gunu(talep.kullanilan_gun)} gün izin kullanmak istiyorum."
        )
    return f"{_metin(getattr(izin_turu, 'ad', None), 'İzin')} kullanmak istiyorum."


def _birim_satirlari(personel):
    ust_birim = _metin(getattr(personel, "gorev_birimi", ""), "")
    alt_birim = _metin(getattr(personel, "alt_birimi", ""), "")
    return ust_birim, alt_birim


def _p(value, style):
    return Paragraph(escape(_metin(value)), style)


def izin_pdf_olustur(talep, personel, imzacilar, goreve_baslama):
    """İzin isteği onayını tek sayfalık, indirilebilir PDF olarak üretir."""
    regular, bold = _fontlar()
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
    title = ParagraphStyle("portal_title", parent=styles["Normal"], fontName=bold, fontSize=12.5, leading=14.5, alignment=TA_CENTER)
    subtitle = ParagraphStyle("portal_subtitle", parent=styles["Normal"], fontName=bold, fontSize=12, leading=14, alignment=TA_CENTER)
    cell = ParagraphStyle("portal_cell", parent=styles["Normal"], fontName=regular, fontSize=10.5, leading=12.5)
    cell_center = ParagraphStyle("portal_cell_center", parent=cell, alignment=TA_CENTER)
    cell_bold = ParagraphStyle("portal_cell_bold", parent=cell, fontName=bold)
    tiny_left = ParagraphStyle("portal_tiny_left", parent=cell, fontSize=9.8, leading=11.5, alignment=TA_LEFT)
    tiny_center = ParagraphStyle("portal_tiny_center", parent=cell_center, fontSize=9.8, leading=11.5)
    foot = ParagraphStyle("portal_foot", parent=cell, fontSize=9.4, leading=11)
    story = []

    left_logo = PdfImage(str(BELEDIYE_LOGO), width=19 * mm, height=19 * mm)
    right_logo = PdfImage(str(TISKI_LOGO), width=19 * mm, height=19 * mm)
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
    story += [antet, Spacer(1, 2 * mm)]

    # Tarih/sayı soldadır; birim adı antette tekrarlanmaz.
    story.append(Paragraph(
        f"<b>Evrak Tarih ve Sayısı:</b> {_tarih(talep.kurum_kayit_tarihi)} - E.{escape(_metin(talep.evrak_no))}",
        ParagraphStyle("doc_number", parent=cell_bold, alignment=TA_LEFT, fontSize=9.4),
    ))
    story += [Spacer(1, 2 * mm), Paragraph("İZİN İSTEĞİ ONAYI", subtitle), Spacer(1, 3 * mm)]

    ust_birim, alt_birim = _birim_satirlari(personel)
    birim_satirlari = [[Paragraph("<b>" + escape(ust_birim) + "</b>", cell_center)]]
    if alt_birim:
        birim_satirlari.append([Paragraph(escape(alt_birim), cell_center)])
    birim_degeri = Table(birim_satirlari, colWidths=[151 * mm])
    birim_degeri.setStyle(TableStyle([
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 1),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 1),
        ("LINEABOVE", (0, 1), (-1, 1), 0.85, colors.black),
    ]))
    izin_adi = _metin(getattr(talep.izin_turu, "ad", None), "İzin")
    rows = [
        [_p("Birimi", cell_bold), birim_degeri, "", ""],
        [_p("Adı Soyadı", cell_bold), _p(_ad_soyad_bicimi(personel.ad_soyad), cell), _p("Ait Olduğu Yıl", cell_bold), _p(talep.ait_oldugu_yil, cell_center)],
        [_p("Görevi", cell_bold), _p(personel.unvani, cell), _p("İzin Başlangıç Tarihi", cell_bold), _p(_tarih(talep.baslangic_tarihi), cell_center)],
        [_p("T.C. Kimlik No", cell_bold), _p(personel.tc_no, cell_center), _p("İzin Bitiş Tarihi", cell_bold), _p(_tarih(talep.bitis_tarihi), cell_center)],
        [_p("Kurum Sicil No", cell_bold), _p(personel.sicil_no, cell_center), _p("Göreve Başlama Tarihi", cell_bold), _p(_tarih(goreve_baslama), cell_center)],
        [_p("İzin Türü", cell_bold), _p(izin_adi, cell_center), _p("Toplam İzin", cell_bold), _p(_izin_gunu(talep.izin_hakki), cell_center)],
        [_p("Talep Ettiği İzin Süresi", cell_bold), _p(_izin_gunu(talep.kullanilan_gun), cell_center), _p("Kalan İzin Süresi", cell_bold), _p(_izin_gunu(talep.kalan_izin), cell_center)],
    ]
    info_table = Table(rows, colWidths=[29 * mm, 63 * mm, 36 * mm, 52 * mm])
    info_table.setStyle(TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.85, colors.black),
        ("SPAN", (1, 0), (3, 0)),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ALIGN", (1, 0), (3, 0), "CENTER"),
        ("LEFTPADDING", (0, 0), (-1, -1), 5),
        ("RIGHTPADDING", (0, 0), (-1, -1), 5),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]))
    story.append(info_table)

    note_html = (
        "<b>Açıklama:</b><br/>"
        + escape(_izin_aciklamasi(talep))
        + "<br/><br/>Gereğini arz ederim."
        + f"<br/><b>Statü:</b> {escape(_metin(talep.get_personel_tipi_display()).upper())}"
    )
    note = Table([[Paragraph(note_html, cell)]], colWidths=[180 * mm])
    note.setStyle(TableStyle([
        ("BOX", (0, 0), (-1, -1), 0.85, colors.black),
        ("LEFTPADDING", (0, 0), (-1, -1), 7),
        ("RIGHTPADDING", (0, 0), (-1, -1), 7),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
    ]))
    story.append(note)

    requester = Table([
        [Paragraph("<b>İzin Talebinde Bulunan</b>", cell_bold)],
        [Paragraph(f"Cep Telefonu: {escape(_metin(personel.cep_tel))}", cell)],
    ], colWidths=[180 * mm])
    requester.setStyle(TableStyle([
        ("BOX", (0, 0), (-1, -1), 0.85, colors.black),
        ("LINEBELOW", (0, 0), (-1, 0), 0.85, colors.black),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 7),
        ("RIGHTPADDING", (0, 0), (-1, -1), 7),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
    ]))
    story.append(requester)

    paraf_kutulari = [
        Paragraph(f"e-Paraflıdır<br/>{escape(_ad_soyad_bicimi(personel.ad_soyad))}<br/>{escape(personel.unvani)}", tiny_center)
    ]
    if len(imzacilar) > 1:
        rol, user = imzacilar[0]
        paraf_kutulari.append(Paragraph(
            f"<b>Teklif Eden</b><br/>"
            f"{'e-Paraflıdır' if talep.durum in {'IMZA_SURECINDE', 'ONAYLANDI'} else 'Paraf Bekliyor'}<br/>"
            f"{escape(_kullanici_adi(user))}<br/>{escape(rol)}",
            tiny_center,
        ))
    paraf_table = Table([paraf_kutulari], colWidths=[180 * mm / len(paraf_kutulari)] * len(paraf_kutulari), rowHeights=[28 * mm])
    paraf_table.setStyle(TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.85, colors.black),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
    ]))
    story.append(paraf_table)

    if imzacilar:
        rol, user = imzacilar[-1]
        approval = Paragraph(
            f"<b>Onaylayan</b><br/>{'e-İmzalıdır' if talep.durum == 'ONAYLANDI' else 'İmza Bekliyor'}<br/>"
            f"{escape(_kullanici_adi(user))}<br/>{escape(rol)}",
            tiny_center,
        )
    else:
        approval = Paragraph("<b>Onaylayan</b><br/>Birimi için imzacı tanımlanmamış.", tiny_center)
    approval_table = Table([[approval]], colWidths=[180 * mm], rowHeights=[26 * mm])
    approval_table.setStyle(TableStyle([
        ("BOX", (0, 0), (-1, -1), 0.85, colors.black),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
    ]))
    story.append(approval_table)
    story += [Spacer(1, 2 * mm), Paragraph(
        f"<b>Kontrol Eden:</b> Personel İzin Kontrolü: {escape(_kullanici_adi(talep.kontrol_eden))}", foot
    )]
    doc.build(story)
    return buffer.getvalue()


def izin_excel_olustur(talep, personel, imzacilar, goreve_baslama):
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
        ("Görevi", personel.unvani, "İzin Başlangıç Tarihi", _tarih(talep.baslangic_tarihi)),
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
        ws.cell(row, 2).alignment = center
        if len(value) == 4:
            ws.cell(row, 3).font = Font(name="Arial", size=10, bold=True)
            ws.cell(row, 4).alignment = center
        row += 1
    ws.merge_cells(start_row=row, start_column=1, end_row=row + 3, end_column=4)
    ws.cell(row, 1, "Açıklama:\n" + _izin_aciklamasi(talep) + "\n\nGereğini arz ederim.")
    ws.cell(row, 1).font = body; ws.cell(row, 1).alignment = left; ws.cell(row, 1).border = border
    row += 5
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=4)
    ws.cell(row, 1, f"e-Paraflıdır\n{_ad_soyad_bicimi(personel.ad_soyad)}\n{personel.unvani}")
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
