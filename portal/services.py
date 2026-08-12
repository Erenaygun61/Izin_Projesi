import re

from django.db import transaction
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
