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
