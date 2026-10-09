import logging
import uuid

from django.db import DatabaseError, OperationalError, transaction
from django.http import JsonResponse
from django.shortcuts import get_object_or_404
from django.views.decorators.http import require_GET, require_POST

from .models import PersonelAktarim
from .personel_senkron import PERSONEL_MODELLERI

logger = logging.getLogger(__name__)
PARCA_BOYUTU = 5


def _yetki_hatasi(request):
    from .views import _personel_veri_yonetimi_yetkisi
    if not request.user.is_authenticated:
        return JsonResponse({"error": "Oturum suresi doldu. Yeniden giris yapin."}, status=401)
    if not _personel_veri_yonetimi_yetkisi(request.user):
        return JsonResponse({"error": "Personel aktarim yetkiniz yok."}, status=403)


def _ozet(aktarim):
    toplam = len(aktarim.kayitlar)
    return {
        "id": str(aktarim.pk), "total": toplam, "processed": aktarim.islenen,
        "success": aktarim.basarili, "failed": len(aktarim.hatalar),
        "errors": aktarim.hatalar, "done": aktarim.tamamlandi,
        "percent": 100 if aktarim.tamamlandi else min(99, int(aktarim.islenen * 100 / toplam)) if toplam else 0,
    }


def aktarim_planini_olustur(islem, tip, secilenler):
    if tip not in PERSONEL_MODELLERI:
        raise ValueError("Gecersiz personel tipi.")
    if islem not in {"tumunu_aktar", "tum_tipleri_aktar", "secilenleri_guncelle"}:
        raise ValueError("Gecersiz aktarim islemi.")
    if islem == "secilenleri_guncelle":
        kayitlar = set()
        for value in secilenler:
            parcalar = value.split(":", 1)
            kaynak_tip, source_id = parcalar if len(parcalar) == 2 else (tip, value)
            kaynak_tip = kaynak_tip.upper()
            if kaynak_tip not in PERSONEL_MODELLERI or not source_id.isdigit() or int(source_id) <= 0:
                raise ValueError("Secili personel bilgisi gecersiz.")
            kayitlar.add((kaynak_tip, int(source_id)))
        if not kayitlar:
            raise ValueError("Aktarmak icin en az bir personel secin.")
        return sorted(kayitlar)
    tipler = list(PERSONEL_MODELLERI) if islem == "tum_tipleri_aktar" else [tip]
    return [(kod, pk) for kod in tipler for pk in PERSONEL_MODELLERI[kod].objects.order_by("pk").values_list("pk", flat=True)]


@require_POST
def personel_aktarim_baslat(request):
    hata = _yetki_hatasi(request)
    if hata:
        return hata
    try:
        request_id = uuid.UUID(request.POST["request_id"]) if request.POST.get("request_id") else uuid.uuid4()
        mevcut = PersonelAktarim.objects.filter(pk=request_id).first()
        if mevcut:
            if mevcut.kullanici_id != request.user.pk:
                return JsonResponse({"error": "Aktarim bulunamadi."}, status=404)
            return JsonResponse(_ozet(mevcut))
        if request.POST.get("retry"):
            onceki = get_object_or_404(PersonelAktarim, pk=uuid.UUID(request.POST["retry"]), kullanici=request.user)
            if not onceki.tamamlandi:
                raise ValueError("Once mevcut aktarimi tamamlayin.")
            kayitlar = [(h["tip"], h["source_id"]) for h in onceki.hatalar]
        else:
            kayitlar = aktarim_planini_olustur(
                request.POST.get("islem", ""), request.POST.get("tip", "MEMUR").upper(), request.POST.getlist("source_ids")
            )
        if not kayitlar:
            raise ValueError("Aktarilacak personel bulunamadi.")
        aktarim, _ = PersonelAktarim.objects.get_or_create(pk=request_id, defaults={"kullanici": request.user, "kayitlar": kayitlar})
        if aktarim.kullanici_id != request.user.pk:
            return JsonResponse({"error": "Aktarim bulunamadi."}, status=404)
        return JsonResponse(_ozet(aktarim), status=201)
    except ValueError as exc:
        return JsonResponse({"error": str(exc)}, status=400)
    except DatabaseError:
        logger.exception("Personel aktarim plani olusturulamadi.")
        return JsonResponse({"error": "Veritabanina ulasilamadi. Biraz sonra yeniden deneyin."}, status=503)


@require_GET
def personel_aktarim_durum(request, aktarim_id):
    hata = _yetki_hatasi(request)
    if hata:
        return hata
    return JsonResponse(_ozet(get_object_or_404(PersonelAktarim, pk=aktarim_id, kullanici=request.user)))


@require_POST
def personel_aktarim_adim(request, aktarim_id):
    hata = _yetki_hatasi(request)
    if hata:
        return hata
    try:
        beklenen = int(request.POST.get("processed", "-1"))
        with transaction.atomic():
            aktarim = get_object_or_404(PersonelAktarim.objects.select_for_update(), pk=aktarim_id, kullanici=request.user)
            if aktarim.tamamlandi or beklenen != aktarim.islenen:
                return JsonResponse(_ozet(aktarim))
            # Claim the write lock before importing; repeated/lost responses cannot duplicate a batch.
            PersonelAktarim.objects.filter(pk=aktarim.pk, islenen=beklenen).update(guncelleme_zamani=aktarim.guncelleme_zamani)
            from .views import ozlukten_portala_aktar
            for tip, source_id in aktarim.kayitlar[beklenen:beklenen + PARCA_BOYUTU]:
                try:
                    with transaction.atomic():
                        kaynak = PERSONEL_MODELLERI[tip].objects.select_related("unvan").get(pk=source_id)
                        ozlukten_portala_aktar(kaynak, tip)
                    aktarim.basarili += 1
                except OperationalError:
                    raise
                except Exception:
                    logger.exception("Personel aktarimi basarisiz: tip=%s source_id=%s", tip, source_id)
                    aktarim.hatalar.append({"tip": tip, "source_id": source_id, "message": "Kart aktarilamadi; yeniden deneyin."})
                aktarim.islenen += 1
            aktarim.tamamlandi = aktarim.islenen == len(aktarim.kayitlar)
            aktarim.save(update_fields=["islenen", "basarili", "hatalar", "tamamlandi", "guncelleme_zamani"])
            return JsonResponse(_ozet(aktarim))
    except ValueError:
        return JsonResponse({"error": "Gecersiz aktarim sirasi."}, status=400)
    except OperationalError:
        logger.warning("Personel aktariminda gecici veritabani hatasi.", exc_info=True)
        response = JsonResponse({"error": "Veritabani mesgul. Aktarim kaldigi yerden tekrar denenebilir."}, status=503)
        response["Retry-After"] = "1"
        return response
