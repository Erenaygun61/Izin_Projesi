from django.db.models import Q

from .evrak_verisi import DUZENLENEBILIR_DURUMLAR


def paraf_verildi_mi(talep, user, rol=None):
    """Only signatures valid in this submission and before the current step."""
    if not talep or not user or talep.durum in DUZENLENEBILIR_DURUMLAR:
        return False
    gecmis = talep.islem_gecmisi.filter(kullanici_id=user.pk, islem__in=["IMZA", "KONTROL_PARAFI", "ONAY"])
    baslangic = talep.islem_gecmisi.filter(Q(islem="ONAYA_SUNULDU") | Q(islem="GERI_CEK", rol="İzin Talebinde Bulunan")).order_by("-pk").first()
    if baslangic:
        gecmis = gecmis.filter(pk__gt=baslangic.pk)
    if rol == "Teklif Eden" or (rol is None and user.pk == talep.kullanici_id):
        return gecmis.filter(islem="IMZA", rol="Teklif Eden").exists()
    if talep.durum == "PARAF_BEKLIYOR":
        return False
    imzacilar = (talep.evrak_verisi or {}).get("imzacilar", [])
    for sira, item in enumerate(imzacilar):
        if item["kullanici"]["id"] == user.pk and (rol is None or item["rol"] == rol):
            return (talep.durum == "ONAYLANDI" or sira < talep.imza_sirasi) and gecmis.filter(rol=item["rol"]).exists()
    return gecmis.exists() if not imzacilar else False
