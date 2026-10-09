import re

from django.db import transaction
from django.db.models import Q
from django.utils import timezone

from .models import EvrakNoSayaci, IzinTalebi


def sonraki_izin_evrak_no(yil=None):
    yil = yil or timezone.localdate().year
    with transaction.atomic():
        sayac, _ = EvrakNoSayaci.objects.select_for_update().get_or_create(yil=yil)
        if sayac.son_no == 0:
            prefix = f"{yil}-"
            son = 0
            for no in IzinTalebi.objects.filter(evrak_no__startswith=prefix).values_list("evrak_no", flat=True):
                match = re.match(rf"^{yil}-(\d+)$", no or "")
                if match:
                    son = max(son, int(match.group(1)))
            sayac.son_no = son
        sayac.son_no += 1
        sayac.save(update_fields=["son_no", "guncelleme_zamani"])
        return f"{yil}-{sayac.son_no}"


def talep_numarasi_filtresi(deger):
    filtre = Q(evrak_no__icontains=deger)
    eslesme = re.fullmatch(r"TALEP-(\d{4})-([1-9]\d{0,18})", deger, flags=re.IGNORECASE)
    if eslesme:
        yil, pk = map(int, eslesme.groups())
        if 1 <= yil <= 9999 and pk <= 9223372036854775807:
            filtre |= Q(pk=pk, olusturma_zamani__year=yil)
    return filtre


def resmi_evrak_kaydini_hazirla(talep):
    """Called inside the locked final-signature transaction, before HR transfer."""
    if talep.durum != "ONAYLANDI":
        raise ValueError("Resmi sayı yalnızca son imza aşamasında verilebilir.")
    tarih = timezone.localdate()
    if not talep.evrak_no:
        talep.evrak_no = sonraki_izin_evrak_no(tarih.year)
        talep.kurum_kayit_tarihi = tarih
    elif not talep.kurum_kayit_tarihi:
        talep.kurum_kayit_tarihi = tarih
    talep.save(update_fields=["evrak_no", "kurum_kayit_tarihi"])
