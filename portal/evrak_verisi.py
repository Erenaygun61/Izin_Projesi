from copy import copy
from types import SimpleNamespace

from django.contrib.auth.models import User
from django.utils import timezone

from .models import PortalIzinHakki
from .templatetags.portal_extras import belge_gorevi


DUZENLENEBILIR_DURUMLAR = {"TASLAK", "REDDEDILDI", "IADE_EDILDI"}


def sabit_veri(talep):
    return talep.evrak_verisi if talep.durum not in DUZENLENEBILIR_DURUMLAR else {}


def kullanici_verisi(user):
    return {"id": user.pk, "ad": user.get_full_name() or user.username, "username": user.username} if user else None


def sabit_kullanici(veri):
    if not veri:
        return None
    user = User.objects.filter(pk=veri["id"]).first()
    # An unsaved historical user still renders its recorded name, not a
    # replacement approver when an account has been removed.
    user = copy(user) if user else User(id=veri["id"])
    user.first_name, user.last_name, user.username = veri["ad"], "", veri["username"]
    return user


def evraki_sabitle(talep, personel, imzacilar):
    if talep.evrak_verisi:
        return
    from .views import _yillara_gore_haklar

    haklar = _yillara_gore_haklar(personel, talep.ait_oldugu_yil, talep.izin_turu) if talep.izin_turu_id else []
    talep.evrak_verisi = {
        "surum": 1,
        "tarih": timezone.now().isoformat(),
        "personel": {
            **{alan: str(getattr(personel, alan, "") or "") for alan in ("ad_soyad", "tc_no", "sicil_no", "gorev_birimi", "alt_birimi", "cep_tel")},
            "unvani": belge_gorevi(personel), "personel_tipi": talep.personel_tipi,
        },
        "imzacilar": [{"rol": rol, "kullanici": kullanici_verisi(user), "asil_kullanici_id": getattr(user, "_asil_kullanici_id", user.pk)} for rol, user in imzacilar],
        "kontrol_eden": kullanici_verisi(talep.kontrol_eden),
        "izin_turu": {"ad": talep.izin_turu.ad, "yillik_izin_mi": talep.izin_turu.yillik_izin_mi} if talep.izin_turu_id else None,
        "bakiye": {"portal": bool(haklar and isinstance(haklar[0], PortalIzinHakki)), "hak_ids": [hak.pk for hak in haklar]},
    }
    talep.save(update_fields=["evrak_verisi"])


def belge_verileri(talep, personel, imzacilar):
    veri = sabit_veri(talep)
    if not veri:
        return talep, personel, imzacilar
    belge = copy(talep)
    belge._state = copy(talep._state)
    belge._state.fields_cache = talep._state.fields_cache.copy()
    if talep.izin_turu_id and veri.get("izin_turu"):
        tur = copy(talep.izin_turu)
        tur.ad = veri["izin_turu"]["ad"]
        tur.yillik_izin_mi = veri["izin_turu"]["yillik_izin_mi"]
        belge.izin_turu = tur
    belge.kontrol_eden = sabit_kullanici(veri.get("kontrol_eden"))
    imzacilar = [(item["rol"], sabit_kullanici(item["kullanici"])) for item in veri["imzacilar"]]
    return belge, SimpleNamespace(**veri["personel"]), imzacilar
