from decimal import Decimal, InvalidOperation

from django import template
from portal.models import _ad_soyad_bicimi


register = template.Library()


def _turkce_baslik(metin):
    metin = str(metin or "").strip()
    if not metin:
        return "-"

    def duzelt(kelime):
        ilk = {"i": "İ", "ı": "I"}.get(kelime[:1], kelime[:1].upper())
        devam = kelime[1:].replace("İ", "i").replace("I", "ı").lower()
        return ilk + devam

    return " ".join(duzelt(kelime) for kelime in metin.split())


@register.filter
def ad_soyad_yazimi(value):
    """Tamamı büyük tutulan kullanıcı adlarını belgede okunur gösterir."""
    return _ad_soyad_bicimi(value)


@register.filter
def izin_gunu(value):
    """35,00 yerine 35; kesir varsa 0,5 / 1,25 biçiminde gösterir."""
    if value in (None, ""):
        return "-"
    try:
        number = Decimal(str(value))
    except (InvalidOperation, ValueError, TypeError):
        return value
    return format(number.normalize(), "f").replace(".", ",")


@register.filter
def belge_gorevi(personel):
    """İzin evrakında işçi/şirket için sade statü metni, memur için özlük unvanı gösterir."""
    if not personel:
        return "-"
    tip = str(getattr(personel, "personel_tipi", "") or "").upper()
    sinif = personel.__class__.__name__.lower()
    if tip == "SIRKET" or "sirket" in sinif:
        return "Şirket Personeli"
    if tip == "ISCI" or "isci" in sinif:
        return "Kadrolu İşçi"
    return getattr(personel, "unvani", "") or getattr(personel, "unvan", "") or getattr(personel, "gorevi", "") or "-"


@register.filter
def paraf_verdi_mi(talep, user):
    from portal.imza_durumu import paraf_verildi_mi
    return paraf_verildi_mi(talep, user)


@register.simple_tag
def makam_parafli_mi(talep, user, rol):
    from portal.imza_durumu import paraf_verildi_mi
    return paraf_verildi_mi(talep, user, rol)
