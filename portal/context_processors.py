from django.utils import timezone

from .models import ImzaciVekaleti


def aktif_vekaletler(request):
    """Üst barda vekilin bugün temsil ettiği asıl kullanıcıları gösterir.

    Süresi biten kayıtlar silinmez; sadece bitiş tarihini izleyen gün pasife
    alınır. Böylece yönetim ekranında tüm vekâlet geçmişi korunur.
    """
    if not getattr(request, "user", None) or not request.user.is_authenticated:
        return {"aktif_vekaletler": []}

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
    return {"aktif_vekaletler": list(kayitlar)}
