import logging

from django.db import OperationalError
from django.db.models import F, OuterRef, Subquery, Value
from django.db.models.functions import Coalesce

from .models import IsciPersonel, Personel, PortalPersonel, SirketPersonel

logger = logging.getLogger(__name__)
PERSONEL_MODELLERI = {"MEMUR": Personel, "ISCI": IsciPersonel, "SIRKET": SirketPersonel}


def personel_durumlarini_esitle():
    toplam = 0
    tipler = set(PortalPersonel.objects.values_list("personel_tipi", flat=True).distinct())
    for tip in tipler:
        model = PERSONEL_MODELLERI.get(tip)
        if model is None:
            continue
        kaynak_durum = Coalesce(Subquery(
            model.objects.filter(pk=OuterRef("source_personel_id")).values("durum")[:1]
        ), Value("Pasif"))
        kartlar = PortalPersonel.objects.filter(personel_tipi=tip).annotate(kaynak_durum=kaynak_durum)
        # Only status follows the source automatically; local names/units stay untouched.
        degisenler = kartlar.exclude(durum=F("kaynak_durum"))
        if degisenler.exists():
            toplam += degisenler.update(durum=kaynak_durum)
    return toplam


def personel_aktif_mi(personel):
    if isinstance(personel, PortalPersonel):
        model = PERSONEL_MODELLERI.get(personel.personel_tipi)
        return bool(model and model.objects.filter(pk=personel.source_personel_id, durum__iexact="Aktif").exists())
    return str(getattr(personel, "durum", "")).strip().casefold() == "aktif"


class PersonelDurumMiddleware:
    def __init__(self, get_response):
        self.get_response = get_response

    def __call__(self, request):
        if request.user.is_authenticated and not request.path.startswith("/static/"):
            try:
                personel_durumlarini_esitle()
            except OperationalError:
                logger.warning("Personel durumu esitlemesi bu istekte tamamlanamadi.", exc_info=True)
        return self.get_response(request)
