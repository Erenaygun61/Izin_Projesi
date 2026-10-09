"""Permission catalogue and atomic, versioned bulk edits for the leave portal."""
import json

from django.contrib import messages
from django.contrib.auth.models import User
from django.core.exceptions import ValidationError
from django.core.paginator import Paginator
from django.db import models, transaction
from django.db.models import Q
from django.shortcuts import redirect, render

from .models import IzinDenetimKaydi, YonetimYetkisi


KONTROL_ALANLARI = ("memur_kontrol_eden", "isci_kontrol_eden", "sirket_kontrol_eden")
ETIKETLER = {
    "tam_yonetim_yetkisi": "Tam yönetim",
    "evrak_yonetimi": "Evrak yönetimi ekranı",
    "tum_izinleri_goruntuleme": "Tüm evrakları gör ve ara",
    "personel_veri_yonetimi": "Personel aktarımı ve düzenleme",
    "gorev_birimi_yonetimi": "Birim ve imzacı yönetimi",
    "vekalet_yonetimi": "Vekalet yönetimi",
    "izin_goruntuleme": "Evrakları görüntüle",
    "izin_arama": "Evrak ve personel ara",
    "kontrol_eden": "Kontrolcü olarak ata",
    "evrak_geri_cekme": "Önceki imza adımına geri çek",
    "sahibine_iade": "Düzenleme için sahibine iade et",
    "izin_silme": "Tamamlanmamış evrakı sil",
    "tamamlanan_izin_bildirimi": "Tamamlanma bildirimlerini al",
}


def yetki_katalogu():
    gruplar = {"Genel": [], "Memur": [], "İşçi": [], "Şirket": []}
    tipler = {"memur": "Memur", "isci": "İşçi", "sirket": "Şirket"}
    for alan in YonetimYetkisi._meta.fields:
        if not isinstance(alan, models.BooleanField) or alan.name in {"aktif_mi", "birim_ve_vekalet_yonetimi"}:
            continue
        prefix, _, eylem = alan.name.partition("_")
        grup = tipler.get(prefix, "Genel")
        etiket = ETIKETLER.get(eylem if prefix in tipler else alan.name, str(alan.verbose_name).replace("_", " "))
        gruplar[grup].append({"alan": alan.name, "etiket": etiket, "kontrol": alan.name in KONTROL_ALANLARI})
    return [{"ad": ad, "alanlar": alanlar} for ad, alanlar in gruplar.items()]


def yetki_surumu(yetki):
    return yetki.guncelleme_zamani.isoformat() if yetki else "yok"


def _denetle(aktor, hedef, onceki, sonraki):
    degisenler = [(alan, deger, sonraki[alan]) for alan, deger in onceki.items() if deger != sonraki[alan]]
    for alan, eski, yeni in degisenler:
        IzinDenetimKaydi.objects.create(kullanici=aktor, hedef_kullanici=hedef, islem="GUNCELLE",
                                       aciklama=f"Portal yetkisi: {alan}; {int(eski)} -> {int(yeni)}")


def yetkileri_uygula(request):
    from .views import _bekleyen_vekalet_evraklarini_guncelle, _yetki_form_degerleri

    katalog = yetki_katalogu()
    alanlar = {item["alan"] for grup in katalog for item in grup["alanlar"]}
    secili = set(request.POST.getlist("yetkiler"))
    ids = request.POST.getlist("kullanicilar")
    islem = request.POST.get("islem", "ekle")
    if islem not in {"ekle", "kaldir", "degistir", "tumunu_kaldir"}:
        raise ValidationError("Yetki işlemi geçersiz.")
    if not ids or len(ids) > 100 or any(not item.isdigit() for item in ids):
        raise ValidationError("İşlem için 1 ile 100 arasında kullanıcı seçin.")
    if secili - alanlar:
        raise ValidationError("Tanımsız yetki gönderildi. Sayfayı yenileyin.")
    if not secili and islem in {"ekle", "kaldir"}:
        raise ValidationError("Değiştirilecek yetkileri seçin.")
    if secili.intersection(KONTROL_ALANLARI) and islem in {"ekle", "degistir"} and len(set(ids)) != 1:
        raise ValidationError("Her personel tipi için yalnızca bir kontrolcü atanabilir. Tek kullanıcı seçin.")
    try:
        surumler = json.loads(request.POST.get("surumler", "{}"))
    except (ValueError, TypeError):
        raise ValidationError("Yetki sürümü geçersiz. Sayfayı yenileyin.")
    if not isinstance(surumler, dict):
        raise ValidationError("Yetki sürümü geçersiz.")
    with transaction.atomic():
        # All permission writes serialize on the same existing user row.
        User.objects.select_for_update().order_by("pk").first()
        kullanicilar = list(User.objects.filter(pk__in=ids, is_active=True).order_by("pk"))
        if len(kullanicilar) != len(set(ids)):
            raise ValidationError("Seçilen kullanıcılardan biri pasif veya bulunamıyor. İşlem uygulanmadı.")
        kayitlar = {y.kullanici_id: y for y in YonetimYetkisi.objects.select_for_update().filter(kullanici_id__in=ids)}
        for user in kullanicilar:
            if surumler.get(str(user.pk)) != yetki_surumu(kayitlar.get(user.pk)):
                raise ValidationError("Yetkiler başka bir işlemle değişti. Sayfayı yenileyin; toplu işlem uygulanmadı.")
        for user in kullanicilar:
            yetki = kayitlar.get(user.pk) or YonetimYetkisi(kullanici=user)
            onceki = _yetki_form_degerleri(yetki)
            for alan in alanlar:
                yeni = alan in secili if islem == "degistir" else bool(onceki.get(alan))
                if islem == "tumunu_kaldir" or (islem == "kaldir" and alan in secili):
                    yeni = False
                elif islem == "ekle" and alan in secili:
                    yeni = True
                setattr(yetki, alan, yeni)
            for alan in KONTROL_ALANLARI:
                if not getattr(yetki, alan):
                    continue
                oncekiler = list(YonetimYetkisi.objects.filter(aktif_mi=True, **{alan: True}).exclude(kullanici=user).select_related("kullanici"))
                if oncekiler and request.POST.get("kontrol_devri_onay") != "on":
                    raise ValidationError("Mevcut kontrolcünün görevini devralmak için kontrolcü devrini onaylayın.")
                for eski in oncekiler:
                    _denetle(request.user, eski.kullanici, {alan: True}, {alan: False})
            yetki.birim_ve_vekalet_yonetimi = False
            yetki.aktif_mi = any(getattr(yetki, alan) for alan in alanlar)
            yetki.aciklama = (request.POST.get("aciklama") or "").strip()[:200]
            yetki.save()
            _denetle(request.user, user, onceki, _yetki_form_degerleri(yetki))
        _bekleyen_vekalet_evraklarini_guncelle(kaydeden=request.user)
    return len(kullanicilar)


def yetki_yonetimi(request):
    if request.method == "POST":
        try:
            sayi = yetkileri_uygula(request)
        except ValidationError as hata:
            messages.error(request, " ".join(hata.messages))
        else:
            messages.success(request, f"{sayi} kullanıcının yetkileri güncellendi.")
        return redirect("kullanici_yetkileri")
    q = (request.GET.get("q") or "").strip()
    qs = User.objects.filter(is_active=True).select_related("portal_yetki").order_by("first_name", "last_name", "username")
    for parca in q.split():
        qs = qs.filter(Q(username__icontains=parca) | Q(first_name__icontains=parca) | Q(last_name__icontains=parca))
    page = Paginator(qs, 25).get_page(request.GET.get("page"))
    rows = []
    for user in page:
        try:
            yetki = user.portal_yetki
        except YonetimYetkisi.DoesNotExist:
            yetki = None
        rows.append({"kullanici": user, "surum": yetki_surumu(yetki), "sayi": sum(bool(getattr(yetki, item["alan"], False)) for grup in yetki_katalogu() for item in grup["alanlar"]), "yetki": yetki})
    kontrolculer = []
    for alan in KONTROL_ALANLARI:
        y = YonetimYetkisi.objects.filter(aktif_mi=True, kullanici__is_active=True, **{alan: True}).select_related("kullanici").first()
        kontrolculer.append({"tip": {"memur": "Memur", "isci": "İşçi", "sirket": "Şirket"}[alan.split("_")[0]], "kullanici": y.kullanici if y else None})
    params = request.GET.copy()
    params.pop("page", None)
    return render(request, "portal/kullanici_yetkileri.html", {"rows": rows, "page_obj": page, "query": params.urlencode(), "q": q, "katalog": yetki_katalogu(), "kontrolculer": kontrolculer})
