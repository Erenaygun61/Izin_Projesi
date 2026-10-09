from django.utils import timezone
from django.urls import reverse

from .models import ImzaciVekaleti, YonetimYetkisi


def _kullanici_yetkisi(user):
    if not user or not user.is_authenticated or not user.is_active:
        return None
    try:
        yetki = user.portal_yetki
    except YonetimYetkisi.DoesNotExist:
        return None
    return yetki if yetki.aktif_mi else None


def _yonetim_nav_items(request):
    user = getattr(request, "user", None)
    if not user or not user.is_authenticated:
        return []

    url_name = getattr(getattr(request, "resolver_match", None), "url_name", "")
    sekme = request.GET.get("sekme", "")

    yetki = _kullanici_yetkisi(user)

    if user.is_active and (user.is_superuser or (yetki and yetki.tam_yonetim_yetkisi)):
        return [{
            "label": "Yönetim",
            "href": reverse("yonetim"),
            "active": url_name in {
                "yonetim",
                "personel_veri_yonetimi",
                "kullanici_yetkileri",
                "kontrol_eden_yonetimi",
                "gorev_birimi_yonetimi",
                "izin_talepleri_yonetim",
            },
        }]

    if not yetki:
        return []

    eski_birlesik_yetki = bool(getattr(yetki, "birim_ve_vekalet_yonetimi", False))
    personel_veri_yetkili = bool(getattr(yetki, "personel_veri_yonetimi", False))
    gorev_birimi_yetkili = bool(getattr(yetki, "gorev_birimi_yonetimi", False) or eski_birlesik_yetki)
    vekalet_yetkili = bool(getattr(yetki, "vekalet_yonetimi", False) or eski_birlesik_yetki)
    items = []
    if personel_veri_yetkili:
        items.append({
            "label": "Personel Veri Yönetimi",
            "href": reverse("personel_veri_yonetimi"),
            "active": url_name == "personel_veri_yonetimi",
        })
    if gorev_birimi_yetkili:
        items.append({
            "label": "Görev Birimi Yönetimi",
            "href": reverse("gorev_birimi_yonetimi"),
            "active": url_name == "gorev_birimi_yonetimi" and sekme != "vekalet",
        })
    if vekalet_yetkili:
        items.append({
            "label": "Vekalet Düzenleme",
            "href": f"{reverse('gorev_birimi_yonetimi')}?sekme=vekalet#vekalet",
            "active": url_name == "gorev_birimi_yonetimi" and sekme == "vekalet",
        })
    return items


def aktif_vekaletler(request):
    """ıst barda vekilin bugön temsil ettiği asıl kullanıcıları gösterir.

    Sıresi biten kayötlar silinmez; sadece bitiı tarihini izleyen gön pasife
    alınır. Bıylece yönetim ekranında tım vekılet geımiıi korunur.
    """
    if not getattr(request, "user", None) or not request.user.is_authenticated:
        return {"aktif_vekaletler": [], "yonetim_nav_items": []}

    bugun = timezone.localdate()
    ImzaciVekaleti.objects.filter(
        aktif_mi=True,
        bitis_tarihi__lt=bugun,
    ).update(aktif_mi=False)

    kayitlar = ImzaciVekaleti.objects.filter(
        vekil_kullanici=request.user,
        aktif_mi=True,
        baslangic_tarihi__lte=bugun,
        bitis_tarihi__gte=bugun,
    ).select_related("asil_kullanici").order_by("asil_kullanici__first_name", "asil_kullanici__last_name")
    return {"aktif_vekaletler": list(kayitlar), "yonetim_nav_items": _yonetim_nav_items(request)}
