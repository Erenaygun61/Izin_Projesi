from datetime import date, timedelta
from decimal import Decimal
import re
import unicodedata

from django.contrib import messages
from django.contrib.auth.decorators import login_required, user_passes_test
from django.core.paginator import Paginator
from django.db import IntegrityError, transaction
from django.db.models import Q, Sum
from django.http import HttpResponse, JsonResponse
from django.shortcuts import get_object_or_404, redirect, render

from .forms import IzinEvrakForm, TcSorguForm
from django.contrib.auth.models import User
from django.utils import timezone

from .models import GorevBirimiImzaci, HakEdilenIzin, ImzaciVekaleti, IzinTalebi, IzinBildirimKapatma, IzinIslemGecmisi, KullanilanIzin, Personel, PortalAyar, PortalPersonel, YonetimYetkisi
from .document_exports import izin_excel_olustur, izin_pdf_olustur

EVRAK_TIPLERI = {
    "memur": {"baslik": "\u0130zin Evrak\u0131 (Memur)", "calisan_tipi": "Memur", "tip_kodu": "MEMUR", "aktif": True},
    "isci": {"baslik": "\u0130zin Evrak\u0131 (\u0130\u015f\u00e7i)", "calisan_tipi": "\u0130\u015f\u00e7i", "tip_kodu": "ISCI", "aktif": False},
    "sirket": {"baslik": "\u0130zin Evrak\u0131 (\u015eirket)", "calisan_tipi": "\u015eirket", "tip_kodu": "SIRKET", "aktif": False},
}


def _normalize(text):
    text = (text or "").replace("\u0130", "i").replace("I", "i").lower().strip()
    text = unicodedata.normalize("NFKD", text)
    return "".join(ch for ch in text if not unicodedata.combining(ch))


def _personel_veri_modeli():
    """Aktarım yapıldıktan sonra portal yalnızca kendi personel kartlarını kullanır."""
    return PortalPersonel if PortalPersonel.objects.exists() else Personel


def _personel_kaynak_id(personel):
    return getattr(personel, "source_personel_id", personel.id)


def _talep_personel_idleri(qs):
    if qs.model is PortalPersonel:
        return qs.values("source_personel_id")
    return qs.values("id")


def _personel_kaynak_kaydi(personel):
    return Personel.objects.filter(id=_personel_kaynak_id(personel)).first()


def ozlukten_portala_aktar(kaynak_personel):
    """Bir özlük kaydını izin portalındaki düzenlenebilir karta kopyalar."""
    defaults = {
        "sicil_no": kaynak_personel.sicil_no,
        "ad": kaynak_personel.ad,
        "soyad": kaynak_personel.soyad,
        "tc_no": kaynak_personel.tc_no,
        "unvan": kaynak_personel.unvani,
        "cep_tel": kaynak_personel.cep_tel,
        "gorev_birimi": kaynak_personel.gorev_birimi,
        "alt_birimi": kaynak_personel.alt_birimi,
        "gorevi": kaynak_personel.gorevi,
        "durum": kaynak_personel.durum,
        "ise_baslama_tarihi": kaynak_personel.ise_baslama_tarihi,
        "ozluk_senkron_zamani": timezone.now(),
    }
    return PortalPersonel.objects.update_or_create(source_personel_id=kaynak_personel.id, defaults=defaults)


def kullanicinin_personeli(user):
    username = (user.username or "").strip()
    for model in (PortalPersonel, Personel):
        if username:
            personel = model.objects.filter(sicil_no=username).first()
            if personel:
                return personel
        if user.first_name or user.last_name:
            personel = model.objects.filter(ad__iexact=user.first_name.strip(), soyad__iexact=user.last_name.strip()).first()
            if personel:
                return personel
    compact_username = _normalize(username).replace(".", "").replace("_", "").replace(" ", "")
    if compact_username:
        for model in (PortalPersonel, Personel):
            for personel in model.objects.all():
                compact_name = _normalize(personel.ad + personel.soyad).replace(" ", "")
                if compact_username == compact_name:
                    return personel
    return None





def gorev_birimlerini_senkronla():
    mevcut = set()
    for row in _personel_veri_modeli().objects.values("gorev_birimi", "alt_birimi").distinct():
        birim = (row.get("gorev_birimi") or "").strip()
        alt = (row.get("alt_birimi") or "").strip()
        if not birim:
            continue
        mevcut.add((birim, alt))
        GorevBirimiImzaci.objects.get_or_create(birim_adi=birim, alt_birim_adi=alt)
    # ?zl?kte art?k olmayanlar? silmiyoruz; yanl??l?kla imzac? kayb? olmas?n diye pasife de alm?yoruz.
    return mevcut

def sonraki_izin_evrak_no(yil=None):
    yil = yil or timezone.localdate().year
    prefix = f"{yil}-"
    son = 0
    for no in IzinTalebi.objects.filter(evrak_no__startswith=prefix).values_list("evrak_no", flat=True):
        m = re.match(rf"^{yil}-(\d+)$", no or "")
        if m:
            son = max(son, int(m.group(1)))
    return f"{yil}-{son + 1}"

def _user_label(user):
    if not user:
        return ""
    full = user.get_full_name().strip()
    return f"{full} ({user.username})" if full else user.username


def _suresi_dolan_vekaletleri_pasife_al(tarih=None):
    """Bitiş tarihini geçen vekâletleri silmeden otomatik pasifleştirir."""
    tarih = tarih or timezone.localdate()
    return ImzaciVekaleti.objects.filter(
        aktif_mi=True,
        bitis_tarihi__lt=tarih,
    ).update(aktif_mi=False)


def _vekil_kullanici(user, tarih=None):
    if not user:
        return None
    tarih = tarih or timezone.localdate()
    vekalet = _aktif_vekalet_kaydi(user, tarih)
    return vekalet.vekil_kullanici if vekalet else user


def _aktif_vekalet_kaydi(user, tarih=None):
    if not user:
        return None
    tarih = tarih or timezone.localdate()
    _suresi_dolan_vekaletleri_pasife_al(tarih)
    return ImzaciVekaleti.objects.filter(
        asil_kullanici=user,
        aktif_mi=True,
        baslangic_tarihi__lte=tarih,
        bitis_tarihi__gte=tarih,
    ).select_related("vekil_kullanici", "asil_kullanici").order_by("-baslangic_tarihi", "-id").first()


def _imzaci_rol_ve_kullanici(rol, user, tarih):
    vekalet = _aktif_vekalet_kaydi(user, tarih)
    if vekalet:
        return (rol if rol.endswith("V.") else f"{rol} V.", vekalet.vekil_kullanici)
    return rol, user


def imzacilari_bul(personel, tarih=None):
    tarih = tarih or timezone.localdate()
    birim = (personel.gorev_birimi or "").strip()
    alt = (getattr(personel, "alt_birimi", "") or "").strip()
    parent = GorevBirimiImzaci.objects.filter(birim_adi=birim, alt_birim_adi="", aktif_mi=True).select_related("daire_baskani", "sube_muduru").first()
    child = None
    if alt:
        child = GorevBirimiImzaci.objects.filter(birim_adi=birim, alt_birim_adi=alt, aktif_mi=True).select_related("daire_baskani", "sube_muduru").first()
    imzacilar = []
    if child and child.sube_muduru_id:
        imzacilar.append(_imzaci_rol_ve_kullanici("Şube Müdürü", child.sube_muduru, tarih))
    daire_user = (child.daire_baskani if child and child.daire_baskani_id else None) or (parent.daire_baskani if parent and parent.daire_baskani_id else None)
    if daire_user:
        imzacilar.append(_imzaci_rol_ve_kullanici("Daire Başkanı", daire_user, tarih))
    # Ayn? ki?i iki role denk gelirse tekrar g?stermeyelim.
    # Genel Müdürlük, ana birim olmasına rağmen daire başkanlığı değildir.
    # Bu nedenle aynı kayıt yapısı kullanılsa da evraktaki imza rolünü
    # doğru unvanla gösteriyoruz.
    if daire_user and "genel mudur" in _normalize(birim):
        mevcut_rol, mevcut_user = imzacilar[-1]
        imzacilar[-1] = ("Genel Müdür V." if mevcut_rol.endswith("V.") else "Genel Müdür", mevcut_user)

    temiz = []
    seen = set()
    for rol, user in imzacilar:
        if user and user.id not in seen:
            temiz.append((rol, user))
            seen.add(user.id)
    return temiz


def imzaci_metni(personel, tarih=None):
    imzacilar = imzacilari_bul(personel, tarih)
    if not imzacilar:
        return "?mzac? tan?mlanmam??"
    return "\n".join(f"{rol}: {_user_label(user)}" for rol, user in imzacilar)


def talep_personeli(talep):
    return (
        PortalPersonel.objects.filter(source_personel_id=talep.personel_id).first()
        or Personel.objects.filter(id=talep.personel_id).first()
    )


def talep_imzacilari(talep):
    personel = talep_personeli(talep)
    return imzacilari_bul(personel, talep.kurum_kayit_tarihi or timezone.localdate()) if personel else []


def goreve_baslama_tarihi(talep):
    if talep.goreve_baslama_tarihi:
        return talep.goreve_baslama_tarihi
    if not talep.bitis_tarihi:
        return None
    tarih = talep.bitis_tarihi + timedelta(days=1)
    while tarih.weekday() >= 5:
        tarih += timedelta(days=1)
    return tarih


def kontrol_edeni_bul(personel_tipi):
    ayar = PortalAyar.aktif()
    alan = {
        "MEMUR": "kontrol_eden_memur",
        "ISCI": "kontrol_eden_isci",
        "SIRKET": "kontrol_eden_sirket",
    }.get(personel_tipi, "kontrol_eden_memur")
    kontrol_eden = getattr(ayar, alan, None)
    # Eski ayar, geçiş sürecinde yalnızca memur için güvenli geri uyumluluk sağlar.
    return kontrol_eden or (ayar.kontrol_eden if personel_tipi == "MEMUR" else None)


def _islem_gecmisi_ekle(talep, kullanici, islem, rol="", aciklama=""):
    return IzinIslemGecmisi.objects.create(
        talep=talep,
        kullanici=kullanici,
        islem=islem,
        rol=rol,
        aciklama=aciklama[:300],
    )


YONETIM_YETKI_ALANLARI = (
    "evrak_yonetimi",
    "tum_izinleri_goruntuleme",
    "memur_kontrol_eden",
    "isci_kontrol_eden",
    "sirket_kontrol_eden",
    "evrak_geri_cekme",
    "sahibine_iade",
    "izin_silme",
    "tamamlanan_izin_bildirimi",
    "birim_ve_vekalet_yonetimi",
)


def _kullanici_yetkisi(user):
    if not user.is_authenticated:
        return None
    try:
        yetki = user.portal_yetki
    except YonetimYetkisi.DoesNotExist:
        return None
    return yetki if yetki.aktif_mi else None


def _izin_yetkisi_var(user, alan):
    if user.is_authenticated and (user.is_superuser or user.is_staff):
        return True
    yetki = _kullanici_yetkisi(user)
    return bool(yetki and getattr(yetki, alan, False))


def _kontrol_yetkisi_var(user, personel_tipi):
    alan = {
        "MEMUR": "memur_kontrol_eden",
        "ISCI": "isci_kontrol_eden",
        "SIRKET": "sirket_kontrol_eden",
    }.get(personel_tipi, "memur_kontrol_eden")
    return _izin_yetkisi_var(user, alan)


def _aktif_vekalet_asilleri(user, tarih=None):
    """Kullanıcının kendi hesabı ile o gün vekil olduğu asıl hesapların kimlikleri."""
    if not user or not user.is_authenticated:
        return set()
    tarih = tarih or timezone.localdate()
    _suresi_dolan_vekaletleri_pasife_al(tarih)
    asil_ids = {user.id}
    asil_ids.update(
        ImzaciVekaleti.objects.filter(
            vekil_kullanici=user,
            aktif_mi=True,
            baslangic_tarihi__lte=tarih,
            bitis_tarihi__gte=tarih,
        ).values_list("asil_kullanici_id", flat=True)
    )
    return asil_ids


def _yonettigi_personeller(user, tarih=None):
    """Amirin görebileceği özlük personelini döndürür.

    Daire başkanı kendi ana birimindeki bütün alt birimleri; şube müdürü ise
    yalnızca tanımlı alt birimdeki personeli görür. Güncel vekaletler de bu
    kapsama dâhildir.
    """
    asil_ids = _aktif_vekalet_asilleri(user, tarih)
    if not asil_ids:
        return _personel_veri_modeli().objects.none()

    kapsam = Q(pk__in=[])
    # Ana birimdeki daire başkanı, o daireye bağlı tüm şubeleri kapsar.
    for birim in GorevBirimiImzaci.objects.filter(
        alt_birim_adi="", daire_baskani_id__in=asil_ids, aktif_mi=True
    ).values_list("birim_adi", flat=True):
        kapsam |= Q(gorev_birimi=birim)
    # Şube müdürü yalnızca kendi alt birimini kapsar.
    for birim, alt_birim in GorevBirimiImzaci.objects.filter(
        sube_muduru_id__in=asil_ids, aktif_mi=True
    ).exclude(alt_birim_adi="").values_list("birim_adi", "alt_birim_adi"):
        kapsam |= Q(gorev_birimi=birim, alt_birimi=alt_birim)
    return _personel_veri_modeli().objects.filter(kapsam)


def _kontrol_ettigi_tipler(user):
    """Merkezî kontrol edenin aramada görebileceği personel tipleri."""
    if not user or not user.is_authenticated:
        return []
    ayar = PortalAyar.aktif()
    alanlar = {
        "MEMUR": "kontrol_eden_memur_id",
        "ISCI": "kontrol_eden_isci_id",
        "SIRKET": "kontrol_eden_sirket_id",
    }
    return [
        tip for tip, alan in alanlar.items()
        if getattr(ayar, alan, None) == user.id and _kontrol_yetkisi_var(user, tip)
    ]


def _talep_arama_filtresi(user):
    """Kullanıcının izin aramasında erişebileceği evrakların filtresi."""
    if _izin_yetkisi_var(user, "tum_izinleri_goruntuleme"):
        return Q()

    # Kendi evrakları, aktif görevi, kontrol ettiği evraklar ve geçmişte
    # paraf/imza attığı tüm evraklar herkes için görünürdür.
    filtre = (
        Q(kullanici=user)
        | Q(kontrol_eden=user)
        | Q(kimin_uzerinde=user)
        | Q(islem_gecmisi__kullanici=user)
    )
    kontrol_tipleri = _kontrol_ettigi_tipler(user)
    if kontrol_tipleri:
        filtre |= Q(personel_tipi__in=kontrol_tipleri)
    yonetilenler = _yonettigi_personeller(user)
    if yonetilenler.exists():
        filtre |= Q(personel_id__in=_talep_personel_idleri(yonetilenler))
    return filtre


def _talep_gorme_yetkisi(user, talep):
    if not user or not user.is_authenticated:
        return False
    if _izin_yetkisi_var(user, "tum_izinleri_goruntuleme"):
        return True
    if (
        talep.kullanici_id == user.id
        or talep.kontrol_eden_id == user.id
        or talep.kimin_uzerinde_id == user.id
        or IzinIslemGecmisi.objects.filter(talep=talep, kullanici=user).exists()
    ):
        return True
    if talep.personel_tipi in _kontrol_ettigi_tipler(user):
        return True
    yonetilenler = _yonettigi_personeller(user)
    if yonetilenler.model is PortalPersonel:
        return yonetilenler.filter(source_personel_id=talep.personel_id).exists()
    return yonetilenler.filter(id=talep.personel_id).exists()


def _talep_aktif_gorev_yetkisi(user, talep):
    if user.is_superuser or user.is_staff:
        return True
    if talep.durum == "PARAF_BEKLIYOR":
        return (
            (talep.kontrol_eden_id == user.id or talep.kimin_uzerinde_id == user.id)
            and _kontrol_yetkisi_var(user, talep.personel_tipi)
        )
    if talep.durum == "IMZA_SURECINDE":
        return talep.kimin_uzerinde_id == user.id
    return False


def _yonetim_erisim_yetkisi(user):
    if not user.is_authenticated:
        return False
    if user.is_superuser or user.is_staff:
        return True
    return YonetimYetkisi.objects.filter(kullanici=user, aktif_mi=True).filter(
        Q(evrak_yonetimi=True)
        | Q(evrak_geri_cekme=True)
        | Q(sahibine_iade=True)
        | Q(izin_silme=True)
        | Q(birim_ve_vekalet_yonetimi=True)
    ).exists()


def _tam_yonetim_yetkisi(user):
    return user.is_authenticated and (user.is_superuser or user.is_staff)


def _evrak_yonetim_yetkisi(user):
    return (
        _izin_yetkisi_var(user, "evrak_yonetimi")
        or _izin_yetkisi_var(user, "evrak_geri_cekme")
        or _izin_yetkisi_var(user, "sahibine_iade")
        or _izin_yetkisi_var(user, "izin_silme")
    )


def _birim_ve_vekalet_yetkisi(user):
    return _izin_yetkisi_var(user, "birim_ve_vekalet_yonetimi")

def izin_ozeti(personel):
    kaynak_id = _personel_kaynak_id(personel)
    haklar = list(HakEdilenIzin.objects.filter(personel_id=kaynak_id).order_by("-izin_yili"))
    kullanimlar = KullanilanIzin.objects.filter(personel_id=kaynak_id, hak_edilen_izin__isnull=False).values("hak_edilen_izin_id").annotate(toplam=Sum("kullanilan_gun"))
    kullanim_map = {row["hak_edilen_izin_id"]: row["toplam"] or Decimal("0.00") for row in kullanimlar}
    toplam_kalan = Decimal("0.00")
    for hak in haklar:
        hak.kullanilan = kullanim_map.get(hak.id, Decimal("0.00"))
        hak.kalan = (hak.hak_edilen_gun - hak.kullanilan).quantize(Decimal("0.01"))
        toplam_kalan += hak.kalan
    return haklar, toplam_kalan.quantize(Decimal("0.01"))


def _yillara_gore_haklar(personel, ait_oldugu_yil):
    """Yıllık izin evrakının seçtiği yıl(lar)daki hak edişleri eskiden yeniye döndürür."""
    kaynak_id = _personel_kaynak_id(personel)
    qs = HakEdilenIzin.objects.filter(personel_id=kaynak_id)
    yillar = [int(yil) for yil in re.findall(r"\d{4}", ait_oldugu_yil or "")]
    if yillar:
        qs = qs.filter(izin_yili__in=yillar)
    return list(qs.order_by("izin_yili", "id"))


def izni_ozluge_aktar(talep, personel, islemi_yapan=None):
    """Son imzası tamamlanan evrağı özlükteki kullanılan izin kayıtlarına işler."""
    if talep.ozluk_aktarildi_mi:
        return
    if not talep.izin_turu_id or not talep.baslangic_tarihi or not talep.bitis_tarihi or not talep.kullanilan_gun:
        raise ValueError("Özlük aktarımı için izin türü, tarih aralığı ve izin süresi gereklidir.")

    gun = Decimal(talep.kullanilan_gun)
    kaynak_id = _personel_kaynak_id(personel)
    kaynak = Personel.objects.filter(id=kaynak_id).first()
    if not kaynak:
        raise ValueError("Özlükte kaynak personel kaydı bulunamadı.")

    ortak_kayit = {
        "ise_baslama_tarihi": kaynak.ise_baslama_tarihi,
        "vekalet_eden": "",
        "izin_gecirecegi_adres": "",
        "ilgili_birim_onay_sayisi": "",
        "ik_onay_sayisi": "",
        "sisteme_giris_yapabilir_mi": True,
        "olusturma_zamani": timezone.now(),
        "guncelleme_zamani": timezone.now(),
    }
    if islemi_yapan:
        ortak_kayit["olusturan_id"] = islemi_yapan.id
        ortak_kayit["guncelleyen_id"] = islemi_yapan.id

    with transaction.atomic():
        if talep.izin_turu.yillik_izin_mi:
            haklar = _yillara_gore_haklar(personel, talep.ait_oldugu_yil)
            kullanilanlar = KullanilanIzin.objects.filter(
                personel_id=kaynak_id,
                hak_edilen_izin_id__in=[hak.id for hak in haklar],
            ).values("hak_edilen_izin_id").annotate(toplam=Sum("kullanilan_gun"))
            kullanilan_map = {row["hak_edilen_izin_id"]: row["toplam"] or Decimal("0.00") for row in kullanilanlar}
            kalan = gun
            aktarim_plani = []
            for hak in haklar:
                uygun = max(Decimal("0.00"), Decimal(hak.hak_edilen_gun) - Decimal(kullanilan_map.get(hak.id, 0)))
                if uygun <= 0:
                    continue
                parca = min(kalan, uygun)
                aktarim_plani.append((hak.id, parca))
                kalan -= parca
                if kalan <= 0:
                    break
            if kalan > 0:
                raise ValueError("Seçilen yıl için özlükte yeterli izin bakiyesi bulunamadı.")
            for hak_id, parca in aktarim_plani:
                KullanilanIzin.objects.create(
                    personel_id=kaynak_id,
                    izin_turu_id=talep.izin_turu_id,
                    hak_edilen_izin_id=hak_id,
                    izin_baslangic_tarihi=talep.baslangic_tarihi,
                    izin_bitis_tarihi=talep.bitis_tarihi,
                    kullanilan_gun=parca,
                    aciklama=f"İzin Portalı evrakı: {talep.evrak_no}",
                    **ortak_kayit,
                )
        else:
            KullanilanIzin.objects.create(
                personel_id=kaynak_id,
                izin_turu_id=talep.izin_turu_id,
                izin_baslangic_tarihi=talep.baslangic_tarihi,
                izin_bitis_tarihi=talep.bitis_tarihi,
                kullanilan_gun=gun,
                aciklama=f"İzin Portalı evrakı: {talep.evrak_no}",
                **ortak_kayit,
            )
        talep.ozluk_aktarildi_mi = True
        talep.ozluk_aktarim_tarihi = timezone.now()


def _evrak_formu_kwargs(personel, haklar, toplam_kalan, tip="memur", **extra):
    bilgi = _evrak_tipi_or_404(tip)
    return {
        "personel": personel,
        "haklar": haklar,
        "calisan_tipi": bilgi["calisan_tipi"],
        "imzacilar": imzaci_metni(personel),
        "toplam_izin": toplam_kalan,
        **extra,
    }


def _talebi_kaydet(talep, form, *, personel, kullanici, evrak_no, onaya_sun, personel_tipi="MEMUR"):
    """Formdan gelen veriyi hem yeni hem iade edilen taslağa güvenle uygular."""
    talep = form.save(commit=False)
    talep.kullanici = kullanici
    talep.personel_id = _personel_kaynak_id(personel)
    talep.personel_tipi = personel_tipi
    talep.evrak_no = evrak_no
    talep.kurum_kayit_tarihi = form.cleaned_data.get("kurum_kayit_tarihi") or timezone.localdate()
    talep.izin_sebebi = ""
    talep.adres = ""
    talep.kontrol_tarihi = None
    talep.imza_sirasi = 0
    talep.imza_tamamlanma_tarihi = None

    if onaya_sun:
        kontrol_eden = kontrol_edeni_bul(personel_tipi)
        if not kontrol_eden:
            return None, "Kontrol eden personel henüz tanımlanmamış. Lütfen İK birimi ile görüşünüz."
        if not _kontrol_yetkisi_var(kontrol_eden, personel_tipi):
            return None, "Tanımlı kontrol eden personelin bu personel tipi için yetkisi yok. Lütfen İK birimi ile görüşünüz."
        talep.durum = "PARAF_BEKLIYOR"
        talep.kontrol_eden = kontrol_eden
        talep.kimin_uzerinde = kontrol_eden
    else:
        talep.durum = "TASLAK"
        talep.kontrol_eden = None
        talep.kimin_uzerinde = kullanici
    talep.save()
    _islem_gecmisi_ekle(
        talep,
        kullanici,
        "ONAYA_SUNULDU" if onaya_sun else "TASLAK",
        "Teklif Eden",
    )
    return talep, None


def _evrak_tipi_or_404(tip):
    bilgi = EVRAK_TIPLERI.get(tip)
    if not bilgi:
        from django.http import Http404
        raise Http404("Evrak tipi bulunamad\u0131")
    return bilgi


@login_required
def dashboard(request):
    personel = kullanicinin_personeli(request.user)
    haklar, toplam_kalan, son_kullanimlar, talepler = [], Decimal("0.00"), [], []
    if personel:
        haklar, toplam_kalan = izin_ozeti(personel)
        personel_kaynak_id = _personel_kaynak_id(personel)
        # Bu alan paraf verilen evrakları değil, yalnızca giriş yapan kişinin
        # kendi oluşturup imza süreci tamamlanan izinlerini gösterir.
        son_kullanimlar = [
            {
                "tur": talep.izin_turu.ad if talep.izin_turu else "İzin türü seçilmedi",
                "baslangic": talep.baslangic_tarihi,
                "bitis": talep.bitis_tarihi,
                "gun": talep.kullanilan_gun,
            }
            for talep in IzinTalebi.objects.filter(
                kullanici=request.user,
                personel_id=personel_kaynak_id,
                durum="ONAYLANDI",
            ).select_related("izin_turu").order_by("-bitis_tarihi", "-guncelleme_zamani")[:6]
        ]
        talepler = IzinTalebi.objects.filter(kullanici=request.user, personel_id=personel_kaynak_id)[:6]
    gorev_filtresi = (
        Q(kimin_uzerinde=request.user, durum="IMZA_SURECINDE")
        | Q(kullanici=request.user, durum__in=["TASLAK", "IADE_EDILDI", "REDDEDILDI"])
    )
    for tip in ("MEMUR", "ISCI", "SIRKET"):
        if _kontrol_yetkisi_var(request.user, tip):
            gorev_filtresi |= Q(kontrol_eden=request.user, durum="PARAF_BEKLIYOR", personel_tipi=tip)
    gorevler = IzinTalebi.objects.select_related("izin_turu", "kullanici", "kimin_uzerinde").filter(
        gorev_filtresi
    ).order_by("-olusturma_zamani")[:25]

    imzalananlar_qs = IzinTalebi.objects.select_related("izin_turu", "kullanici").filter(durum="ONAYLANDI")
    if not _izin_yetkisi_var(request.user, "tamamlanan_izin_bildirimi"):
        imzalananlar_qs = imzalananlar_qs.filter(kullanici=request.user)
    imzalananlar = imzalananlar_qs.exclude(
        bildirim_kapatmalari__kullanici=request.user
    ).order_by("-guncelleme_zamani")[:10]
    return render(request, "portal/dashboard.html", {"personel": personel, "haklar": haklar, "toplam_kalan": toplam_kalan, "son_kullanimlar": son_kullanimlar, "talepler": talepler, "gorevler": gorevler, "imzalananlar": imzalananlar})


@login_required
def izin_menu(request):
    return render(request, "portal/izin_menu.html", {"evrak_tipleri": EVRAK_TIPLERI})


@login_required
def tc_sorgula(request, tip):
    bilgi = _evrak_tipi_or_404(tip)
    form = TcSorguForm(request.POST or None)
    personel = None
    if request.method == "POST" and form.is_valid():
        if not bilgi["aktif"]:
            messages.error(request, f"{bilgi['baslik']} i\u00e7in personel kayna\u011f\u0131 hen\u00fcz ba\u011flanmad\u0131.")
        else:
            personel = _personel_veri_modeli().objects.filter(tc_no=form.cleaned_data["tc_no"]).first()
            if personel:
                return redirect("izin_evrak", tip=tip, personel_id=personel.id)
            messages.error(request, "Bu T.C. Kimlik No ile kay\u0131tl\u0131 personel bulunamad\u0131.")
    return render(request, "portal/tc_sorgula.html", {"form": form, "tip": tip, "bilgi": bilgi, "personel": personel})


@login_required
def izin_evrak(request, tip, personel_id):
    bilgi = _evrak_tipi_or_404(tip)
    if not bilgi["aktif"]:
        messages.error(request, f"{bilgi['baslik']} i\u00e7in personel kayna\u011f\u0131 hen\u00fcz ba\u011flanmad\u0131.")
        return redirect("tc_sorgula", tip=tip)
    personel = get_object_or_404(_personel_veri_modeli(), id=personel_id)
    haklar, toplam_kalan = izin_ozeti(personel)
    imzacilar = imzacilari_bul(personel)
    if not haklar or toplam_kalan <= 0:
        messages.warning(request, "\u0130zin bakiyesi bulunamad\u0131. L\u00fctfen \u0130K birimi ile g\u00f6r\u00fc\u015f\u00fcn\u00fcz.")
    if not imzacilar:
        messages.warning(request, "Bu personelin birimi i\u00e7in imzac\u0131 tan\u0131mlanmam\u0131\u015f. L\u00fctfen \u0130K birimi ile g\u00f6r\u00fc\u015f\u00fcn\u00fcz.")
    if request.method == "GET":
        acik_taslak = IzinTalebi.objects.filter(
            kullanici=request.user,
            personel_id=_personel_kaynak_id(personel),
            durum="TASLAK",
        ).order_by("-guncelleme_zamani").first()
        if acik_taslak:
            messages.info(request, "Bu kişi için yarım bırakılmış bir izin evrakı var. Kaldığınız yerden devam ediyorsunuz.")
            return redirect("izin_talebi_duzenle", talep_id=acik_taslak.id)
        taslak = IzinTalebi.objects.create(
            kullanici=request.user,
            personel_id=_personel_kaynak_id(personel),
            personel_tipi=bilgi["tip_kodu"],
            kurum_kayit_tarihi=timezone.localdate(),
            evrak_no=sonraki_izin_evrak_no(),
            durum="TASLAK",
            kimin_uzerinde=request.user,
        )
        _islem_gecmisi_ekle(taslak, request.user, "TASLAK", "Teklif Eden")
        messages.info(request, "İzin evrakı taslak olarak oluşturuldu. Dilediğiniz zaman Görevlerim alanından kaldığınız yerden devam edebilirsiniz.")
        return redirect("izin_talebi_duzenle", talep_id=taslak.id)
    evrak_no = sonraki_izin_evrak_no()
    if request.method == "POST":
        islem = request.POST.get("islem", "onaya_sun")
        taslak = islem == "taslak_kaydet"
        form = IzinEvrakForm(
            request.POST, taslak=taslak, next_evrak_no=evrak_no,
            **_evrak_formu_kwargs(personel, haklar, toplam_kalan, tip),
        )
        if form.is_valid():
            talep, hata = _talebi_kaydet(
                IzinTalebi(), form, personel=personel, kullanici=request.user,
                evrak_no=evrak_no, onaya_sun=not taslak, personel_tipi=bilgi["tip_kodu"],
            )
            if hata:
                form.add_error(None, hata)
            elif taslak:
                messages.success(request, f"Evrak taslak olarak kaydedildi. Evrak No: {talep.evrak_no}")
                return redirect("dashboard")
            else:
                messages.success(request, f"\u0130zin evrak\u0131 paraf kontrol\u00fcne al\u0131nd\u0131. Evrak No: {talep.evrak_no}")
                return redirect("evrak_paraf", talep_id=talep.id)
    else:
        form = IzinEvrakForm(next_evrak_no=evrak_no, **_evrak_formu_kwargs(personel, haklar, toplam_kalan, tip))
    return render(request, "portal/izin_evrak.html", {"form": form, "personel": personel, "haklar": haklar, "toplam_kalan": toplam_kalan, "tip": tip, "bilgi": bilgi, "imzacilar": imzacilar, "duzenleme": False})


@login_required
def izin_talebi_duzenle(request, talep_id):
    talep = get_object_or_404(IzinTalebi.objects.select_related("izin_turu", "kullanici"), id=talep_id)
    if talep.kullanici_id != request.user.id:
        messages.error(request, "Bu evrağı düzenleme yetkiniz yok.")
        return redirect("dashboard")
    if talep.durum not in {"TASLAK", "IADE_EDILDI", "REDDEDILDI"}:
        messages.warning(request, "Onaya sunulmuş evrak düzenlenemez. Gerekirse yönetimden sahibine iade edilmesini isteyin.")
        return redirect("evrak_paraf", talep_id=talep.id)

    personel = talep_personeli(talep)
    if not personel:
        messages.error(request, "Evrağa bağlı personel kaydı bulunamadı.")
        return redirect("dashboard")
    haklar, toplam_kalan = izin_ozeti(personel)
    imzacilar = imzacilari_bul(personel)
    if request.method == "POST":
        islem = request.POST.get("islem", "onaya_sun")
        taslak = islem == "taslak_kaydet"
        form = IzinEvrakForm(
            request.POST, instance=talep, taslak=taslak, next_evrak_no=talep.evrak_no,
            **_evrak_formu_kwargs(personel, haklar, toplam_kalan),
        )
        if form.is_valid():
            talep, hata = _talebi_kaydet(
                talep, form, personel=personel, kullanici=request.user,
                evrak_no=talep.evrak_no, onaya_sun=not taslak, personel_tipi=talep.personel_tipi,
            )
            if hata:
                form.add_error(None, hata)
            elif taslak:
                messages.success(request, "Evrak bilgileri taslak olarak güncellendi.")
                return redirect("dashboard")
            else:
                messages.success(request, f"Evrak güncellendi ve paraf kontrolüne yeniden sunuldu. Evrak No: {talep.evrak_no}")
                return redirect("evrak_paraf", talep_id=talep.id)
    else:
        form = IzinEvrakForm(
            instance=talep, next_evrak_no=talep.evrak_no,
            **_evrak_formu_kwargs(personel, haklar, toplam_kalan),
        )
    tip = {"MEMUR": "memur", "ISCI": "isci", "SIRKET": "sirket"}.get(talep.personel_tipi, "memur")
    bilgi = _evrak_tipi_or_404(tip)
    return render(request, "portal/izin_evrak.html", {
        "form": form, "personel": personel, "haklar": haklar,
        "toplam_kalan": toplam_kalan, "tip": tip, "bilgi": bilgi,
        "imzacilar": imzacilar, "duzenleme": True, "talep": talep,
    })


@login_required
def yeni_talep(request):
    return redirect("izin_menu")


@login_required
def izin_bildirimini_kapat(request, talep_id):
    if request.method != "POST":
        return redirect("dashboard")
    qs = IzinTalebi.objects.filter(id=talep_id, durum="ONAYLANDI")
    if not _izin_yetkisi_var(request.user, "tamamlanan_izin_bildirimi"):
        qs = qs.filter(kullanici=request.user)
    talep = get_object_or_404(qs)
    IzinBildirimKapatma.objects.get_or_create(kullanici=request.user, talep=talep)
    messages.success(request, "Tamamlanan izin evrakı bildirimi kapatıldı.")
    return redirect("dashboard")


@login_required
def izin_bildirimlerini_kapat(request):
    if request.method != "POST":
        return redirect("dashboard")
    qs = IzinTalebi.objects.filter(durum="ONAYLANDI")
    if not _izin_yetkisi_var(request.user, "tamamlanan_izin_bildirimi"):
        qs = qs.filter(kullanici=request.user)
    talep_ids = qs.values_list("id", flat=True)
    IzinBildirimKapatma.objects.bulk_create(
        [IzinBildirimKapatma(kullanici=request.user, talep_id=talep_id) for talep_id in talep_ids],
        ignore_conflicts=True,
    )
    messages.success(request, "Tamamlanan izin evrakı bildirimleri kapatıldı.")
    return redirect("dashboard")


@user_passes_test(_yonetim_erisim_yetkisi, login_url="dashboard")
def yonetim(request):
    evrak_yetkili = _evrak_yonetim_yetkisi(request.user)
    talepler = []
    if evrak_yetkili:
        talepler = IzinTalebi.objects.select_related("izin_turu", "kullanici", "kimin_uzerinde").filter(
            durum__in=["PARAF_BEKLIYOR", "IMZA_SURECINDE", "IADE_EDILDI"]
        )[:8]
    return render(request, "portal/yonetim.html", {
        "talepler": talepler,
        "tam_yetkili": _tam_yonetim_yetkisi(request.user),
        "birim_yetkili": _birim_ve_vekalet_yetkisi(request.user),
        "evrak_yetkili": evrak_yetkili,
    })


@user_passes_test(_tam_yonetim_yetkisi, login_url="dashboard")
def personel_veri_yonetimi(request):
    """Özlükten seçilerek aktarılan portal personel kartlarını yönetir."""
    if request.method == "POST":
        islem = request.POST.get("islem")
        if islem in {"tumunu_aktar", "secilenleri_guncelle"}:
            if islem == "tumunu_aktar":
                kaynaklar = Personel.objects.select_related("unvan").all()
            else:
                source_ids = [int(value) for value in request.POST.getlist("source_ids") if value.isdigit()]
                kaynaklar = Personel.objects.select_related("unvan").filter(id__in=source_ids)
            aktarilan = 0
            for kaynak in kaynaklar:
                ozlukten_portala_aktar(kaynak)
                aktarilan += 1
            if aktarilan:
                messages.success(request, f"{aktarilan} personel özlükten izin portalına aktarıldı/güncellendi.")
            else:
                messages.warning(request, "Aktarmak için en az bir personel seçin.")
            return redirect("personel_veri_yonetimi")

        if islem == "yerel_kaydet":
            source_id = request.POST.get("source_personel_id")
            kaynak = Personel.objects.filter(id=source_id).first()
            if not kaynak:
                messages.error(request, "Özlükte kaynak personel kaydı bulunamadı.")
                return redirect("personel_veri_yonetimi")
            portal_personeli, _ = ozlukten_portala_aktar(kaynak)
            for alan in ("sicil_no", "ad", "soyad", "tc_no", "unvan", "cep_tel", "gorev_birimi", "alt_birimi", "gorevi", "durum"):
                setattr(portal_personeli, alan, (request.POST.get(alan) or "").strip())
            portal_personeli.save()
            messages.success(request, f"{portal_personeli.ad_soyad} için yerel izin portalı bilgileri kaydedildi.")
            return redirect("personel_veri_yonetimi")

    q = (request.GET.get("q") or "").strip()
    kaynak_qs = Personel.objects.select_related("unvan").all()
    if q:
        lookup = Q(sicil_no__icontains=q) | Q(tc_no__icontains=q) | Q(ad__icontains=q) | Q(soyad__icontains=q) | Q(gorev_birimi__icontains=q) | Q(alt_birimi__icontains=q) | Q(gorevi__icontains=q) | Q(unvan__ad__icontains=q)
        parcalar = q.split()
        if len(parcalar) >= 2:
            lookup |= Q(ad__icontains=parcalar[0], soyad__icontains=parcalar[-1])
        kaynak_qs = kaynak_qs.filter(lookup)
    kaynaklar = list(kaynak_qs.order_by("sicil_no"))
    yerel_map = {kayit.source_personel_id: kayit for kayit in PortalPersonel.objects.filter(source_personel_id__in=[k.id for k in kaynaklar])}
    satirlar = [
        {
            "kaynak": kaynak,
            "portal": yerel_map.get(kaynak.id),
            "sicil_no": (yerel_map.get(kaynak.id).sicil_no if yerel_map.get(kaynak.id) else kaynak.sicil_no),
            "ad_soyad": (yerel_map.get(kaynak.id).ad_soyad if yerel_map.get(kaynak.id) else kaynak.ad_soyad),
            "unvan": (yerel_map.get(kaynak.id).unvani if yerel_map.get(kaynak.id) else kaynak.unvani),
            "birim": " / ".join(filter(None, [
                (yerel_map.get(kaynak.id).gorev_birimi if yerel_map.get(kaynak.id) else kaynak.gorev_birimi),
                (yerel_map.get(kaynak.id).alt_birimi if yerel_map.get(kaynak.id) else kaynak.alt_birimi),
            ])),
        }
        for kaynak in kaynaklar
    ]

    duzenlenecek = None
    edit_source_id = request.GET.get("duzenle")
    if edit_source_id and edit_source_id.isdigit():
        kaynak = Personel.objects.select_related("unvan").filter(id=int(edit_source_id)).first()
        yerel = yerel_map.get(int(edit_source_id)) or PortalPersonel.objects.filter(source_personel_id=edit_source_id).first()
        if kaynak:
            duzenlenecek = {
                "source_personel_id": kaynak.id,
                "sicil_no": yerel.sicil_no if yerel else kaynak.sicil_no,
                "ad": yerel.ad if yerel else kaynak.ad,
                "soyad": yerel.soyad if yerel else kaynak.soyad,
                "tc_no": yerel.tc_no if yerel else kaynak.tc_no,
                "unvan": yerel.unvan if yerel else kaynak.unvani,
                "cep_tel": yerel.cep_tel if yerel else kaynak.cep_tel,
                "gorev_birimi": yerel.gorev_birimi if yerel else kaynak.gorev_birimi,
                "alt_birimi": yerel.alt_birimi if yerel else kaynak.alt_birimi,
                "gorevi": yerel.gorevi if yerel else kaynak.gorevi,
                "durum": yerel.durum if yerel else kaynak.durum,
                "ad_soyad": yerel.ad_soyad if yerel else kaynak.ad_soyad,
            }
    return render(request, "portal/personel_veri_yonetimi.html", {
        "satirlar": satirlar,
        "toplam_kaynak": Personel.objects.count(),
        "toplam_portal": PortalPersonel.objects.count(),
        "q": q,
        "duzenlenecek": duzenlenecek,
    })


@user_passes_test(_evrak_yonetim_yetkisi, login_url="dashboard")
def izin_talepleri_yonetim(request):
    if request.method == "POST":
        talep = get_object_or_404(IzinTalebi, id=request.POST.get("talep_id"))
        islem = request.POST.get("islem")
        not_text = (request.POST.get("not") or "").strip()
        if islem == "geri_cek":
            if not _izin_yetkisi_var(request.user, "evrak_geri_cekme"):
                messages.error(request, "Bu evrağı geri çekme yetkiniz yok.")
                return redirect("izin_talepleri_yonetim")
            if talep.durum in ["IMZA_SURECINDE", "PARAF_BEKLIYOR"]:
                talep.durum = "PARAF_BEKLIYOR"
                talep.kimin_uzerinde = talep.kontrol_eden
                talep.imza_sirasi = 0
                ek = f"\n[{timezone.localtime().strftime('%d.%m.%Y %H:%M')}] Evrak y\u00f6netim taraf\u0131ndan kontrol a\u015famas\u0131na geri \u00e7ekildi."
                if not_text:
                    ek += f" Not: {not_text}"
                talep.aciklama = (talep.aciklama or "") + ek
                talep.save(update_fields=["durum", "kimin_uzerinde", "imza_sirasi", "aciklama", "guncelleme_zamani"])
                _islem_gecmisi_ekle(talep, request.user, "GERI_CEK", "Evrak Yönetimi", not_text)
                messages.success(request, "Evrak kontrol a\u015famas\u0131na geri \u00e7ekildi.")
            else:
                messages.error(request, "Bu durumdaki evrak geri \u00e7ekilemez.")
        elif islem == "sahibine_iade":
            if not _izin_yetkisi_var(request.user, "sahibine_iade"):
                messages.error(request, "Bu evrağı sahibine iade etme yetkiniz yok.")
                return redirect("izin_talepleri_yonetim")
            talep.durum = "IADE_EDILDI"
            talep.kimin_uzerinde = talep.kullanici
            ek = f"\n[{timezone.localtime().strftime('%d.%m.%Y %H:%M')}] Evrak sahibine iade edildi."
            if not_text:
                ek += f" Not: {not_text}"
            talep.aciklama = (talep.aciklama or "") + ek
            talep.save(update_fields=["durum", "kimin_uzerinde", "aciklama", "guncelleme_zamani"])
            _islem_gecmisi_ekle(talep, request.user, "IADE", "Evrak Yönetimi", not_text)
            messages.warning(request, "Evrak sahibine iade edildi.")
        return redirect("izin_talepleri_yonetim")
    q = (request.GET.get("q") or "").strip()
    durum = (request.GET.get("durum") or "").strip()
    qs = IzinTalebi.objects.select_related("izin_turu", "kullanici", "kontrol_eden", "kimin_uzerinde").order_by("-olusturma_zamani")
    if q:
        qs = qs.filter(Q(evrak_no__icontains=q) | Q(kullanici__username__icontains=q) | Q(kullanici__first_name__icontains=q) | Q(kullanici__last_name__icontains=q))
    if durum:
        qs = qs.filter(durum=durum)
    rows = []
    for talep in qs[:200]:
        rows.append({"talep": talep, "personel": talep_personeli(talep)})
    return render(request, "portal/izin_talepleri_yonetim.html", {
        "rows": rows, "q": q, "durum": durum, "durumlar": IzinTalebi.DURUM_CHOICES,
        "geri_cekebilir": _izin_yetkisi_var(request.user, "evrak_geri_cekme"),
        "iade_edebilir": _izin_yetkisi_var(request.user, "sahibine_iade"),
    })


@user_passes_test(_tam_yonetim_yetkisi, login_url="dashboard")
def kullanici_yetkileri(request):
    if request.method == "POST":
        islem = request.POST.get("islem")
        if islem == "yetki_kaydet":
            user_id = request.POST.get("kullanici")
            if user_id:
                kullanici = User.objects.filter(id=user_id, is_active=True).first()
                if not kullanici:
                    messages.error(request, "Seçilen kullanıcı bulunamadı veya pasif durumda.")
                else:
                    yetki, _ = YonetimYetkisi.objects.get_or_create(kullanici=kullanici)
                    for alan in YONETIM_YETKI_ALANLARI:
                        setattr(yetki, alan, request.POST.get(alan) == "on")
                    yetki.aktif_mi = any(getattr(yetki, alan) for alan in YONETIM_YETKI_ALANLARI)
                    yetki.aciklama = request.POST.get("aciklama", "")
                    yetki.save()
                    if yetki.aktif_mi:
                        messages.success(request, "Kullanıcı yetkileri kaydedildi.")
                    else:
                        messages.warning(request, "Hiçbir yetki seçilmediği için kullanıcı pasif durumda bırakıldı.")
            else:
                messages.error(request, "Yetki verilecek kullanıcıyı seçin.")
        elif islem == "yetki_pasif":
            yetki = get_object_or_404(YonetimYetkisi, id=request.POST.get("yetki_id"))
            for alan in YONETIM_YETKI_ALANLARI:
                setattr(yetki, alan, False)
            yetki.aktif_mi = False
            yetki.save(update_fields=[*YONETIM_YETKI_ALANLARI, "aktif_mi", "guncelleme_zamani"])
            messages.warning(request, "Kullanıcının tüm portal yetkileri kaldırıldı.")
        return redirect("kullanici_yetkileri")
    yetkiler = YonetimYetkisi.objects.select_related("kullanici")[:100]
    return render(request, "portal/kullanici_yetkileri.html", {"yetkiler": yetkiler, "yetki_alanlari": YONETIM_YETKI_ALANLARI})


@user_passes_test(_tam_yonetim_yetkisi, login_url="dashboard")
def kullanici_yetkisi_detay(request):
    user_id = request.GET.get("kullanici")
    user = User.objects.filter(id=user_id, is_active=True).first()
    if not user:
        return JsonResponse({"error": "Kullanıcı bulunamadı."}, status=404)
    yetki = YonetimYetkisi.objects.filter(kullanici=user).first()
    return JsonResponse({
        "kullanici": {"id": user.id, "label": _user_label(user)},
        "aciklama": yetki.aciklama if yetki else "",
        "yetkiler": {alan: bool(yetki and getattr(yetki, alan)) for alan in YONETIM_YETKI_ALANLARI},
    })


@user_passes_test(_birim_ve_vekalet_yetkisi, login_url="dashboard")
def gorev_birimi_yonetimi(request):
    bugun = timezone.localdate()
    _suresi_dolan_vekaletleri_pasife_al(bugun)
    if request.method == "POST":
        islem = request.POST.get("islem")
        if islem == "birim_senkronize":
            gorev_birimlerini_senkronla()
            messages.success(request, "Görev birimleri portal personel kartlarından senkronize edildi.")
            return redirect("gorev_birimi_yonetimi")
        if islem == "imzaci_kaydet":
            for kayit in GorevBirimiImzaci.objects.all():
                if kayit.alt_birim_adi:
                    daire_id = None
                    sube_id = request.POST.get(f"sube_{kayit.id}") or None
                else:
                    daire_id = request.POST.get(f"daire_{kayit.id}") or None
                    sube_id = None
                kayit.daire_baskani_id = daire_id
                kayit.sube_muduru_id = sube_id
                kayit.aktif_mi = True
                kayit.save(update_fields=["daire_baskani", "sube_muduru", "aktif_mi", "guncelleme_zamani"])
            messages.success(request, "G\u00f6rev birimi imzac\u0131lar\u0131 kaydedildi.")
            return redirect("gorev_birimi_yonetimi")
        if islem == "vekalet_ekle":
            asil_id = request.POST.get("asil_kullanici")
            vekil_id = request.POST.get("vekil_kullanici")
            baslangic = request.POST.get("baslangic_tarihi")
            bitis = request.POST.get("bitis_tarihi")
            if asil_id and vekil_id and baslangic and bitis:
                try:
                    baslangic_tarihi = date.fromisoformat(baslangic)
                    bitis_tarihi = date.fromisoformat(bitis)
                except ValueError:
                    messages.error(request, "Vekalet tarihleri geçerli değil.")
                    return redirect("gorev_birimi_yonetimi")
                if asil_id == vekil_id:
                    messages.error(request, "Asıl kullanıcı ile vekil kullanıcı aynı olamaz.")
                    return redirect("gorev_birimi_yonetimi")
                if bitis_tarihi < baslangic_tarihi:
                    messages.error(request, "Vekalet bitiş tarihi başlangıçtan önce olamaz.")
                    return redirect("gorev_birimi_yonetimi")
                ImzaciVekaleti.objects.create(
                    asil_kullanici_id=asil_id,
                    vekil_kullanici_id=vekil_id,
                    baslangic_tarihi=baslangic_tarihi,
                    bitis_tarihi=bitis_tarihi,
                    aciklama=request.POST.get("aciklama", ""),
                    aktif_mi=True,
                )
                messages.success(request, "Vekalet kaydı eklendi.")
            else:
                messages.error(request, "Vekalet i\u00e7in asil, vekil ve tarih alanlar\u0131 zorunludur.")
            return redirect("gorev_birimi_yonetimi")
        if islem == "vekalet_pasif":
            vekalet = get_object_or_404(ImzaciVekaleti, id=request.POST.get("vekalet_id"))
            vekalet.aktif_mi = False
            vekalet.save(update_fields=["aktif_mi"])
            messages.success(request, "Vekalet pasife al\u0131nd\u0131.")
            return redirect("gorev_birimi_yonetimi")

    kullanicilar = User.objects.filter(is_active=True).order_by("first_name", "last_name", "username")
    kayitlar = list(GorevBirimiImzaci.objects.select_related("daire_baskani", "sube_muduru").order_by("birim_adi", "alt_birim_adi"))
    for kayit in kayitlar:
        kayit.ana_imzaci_etiketi = "Genel Müdür" if "genel mudur" in _normalize(kayit.birim_adi) else "Daire Başkanı"
    vekalet_q = (request.GET.get("vekalet_q") or "").strip()
    vekalet_qs = ImzaciVekaleti.objects.select_related("asil_kullanici", "vekil_kullanici")
    if vekalet_q:
        lookup = (
            Q(asil_kullanici__username__icontains=vekalet_q)
            | Q(asil_kullanici__first_name__icontains=vekalet_q)
            | Q(asil_kullanici__last_name__icontains=vekalet_q)
            | Q(vekil_kullanici__username__icontains=vekalet_q)
            | Q(vekil_kullanici__first_name__icontains=vekalet_q)
            | Q(vekil_kullanici__last_name__icontains=vekalet_q)
            | Q(aciklama__icontains=vekalet_q)
        )
        parcalar = vekalet_q.split()
        if len(parcalar) >= 2:
            lookup |= (
                Q(asil_kullanici__first_name__icontains=parcalar[0], asil_kullanici__last_name__icontains=parcalar[-1])
                | Q(vekil_kullanici__first_name__icontains=parcalar[0], vekil_kullanici__last_name__icontains=parcalar[-1])
            )
        vekalet_qs = vekalet_qs.filter(lookup)
    vekaletler = list(vekalet_qs.order_by("-baslangic_tarihi", "-id")[:300])
    for vekalet in vekaletler:
        vekalet.gecerli_mi = bool(vekalet.aktif_mi and vekalet.baslangic_tarihi <= bugun <= vekalet.bitis_tarihi)
        if vekalet.gecerli_mi:
            vekalet.durum_etiketi = "Aktif"
            vekalet.durum_sinifi = "active"
        elif vekalet.bitis_tarihi < bugun:
            vekalet.durum_etiketi = "Süresi doldu"
            vekalet.durum_sinifi = "expired"
        else:
            vekalet.durum_etiketi = "Pasif"
            vekalet.durum_sinifi = "passive"
    return render(request, "portal/gorev_birimi_yonetimi.html", {
        "kullanicilar": kullanicilar,
        "kayitlar": kayitlar,
        "vekaletler": vekaletler,
        "vekalet_q": vekalet_q,
        "ayar": PortalAyar.aktif(),
    })


@user_passes_test(_tam_yonetim_yetkisi, login_url="dashboard")
def kontrol_eden_yonetimi(request):
    ayar = PortalAyar.aktif()
    alanlar = ["kontrol_eden_memur", "kontrol_eden_isci", "kontrol_eden_sirket"]
    if request.method == "POST":
        tipler = {
            "kontrol_eden_memur": "MEMUR",
            "kontrol_eden_isci": "ISCI",
            "kontrol_eden_sirket": "SIRKET",
        }
        for alan in alanlar:
            user_id = request.POST.get(alan) or None
            user = User.objects.filter(id=user_id, is_active=True).first() if user_id else None
            if user and not _kontrol_yetkisi_var(user, tipler[alan]):
                messages.error(
                    request,
                    f"{_user_label(user)} için önce {tipler[alan].title()} kontrol eden yetkisini Kullanıcı Yetkileri ekranından verin.",
                )
                continue
            setattr(ayar, f"{alan}_id", user.id if user else None)
        # Eski alan yalnızca eski kayıtlarla uyumluluk içindir.
        ayar.kontrol_eden_id = ayar.kontrol_eden_memur_id
        ayar.save(update_fields=["kontrol_eden", *alanlar, "guncelleme_zamani"])
        messages.success(request, "Memur, işçi ve şirket izinleri için kontrol eden personeller kaydedildi.")
        return redirect("kontrol_eden_yonetimi")
    return render(request, "portal/kontrol_eden_yonetimi.html", {"ayar": ayar})


@login_required
def kullanici_ara(request):
    q = (request.GET.get("q") or "").strip()
    qs = User.objects.filter(is_active=True)
    if q:
        lookup = Q(username__icontains=q) | Q(first_name__icontains=q) | Q(last_name__icontains=q)
        parcalar = q.split()
        if len(parcalar) >= 2:
            lookup |= Q(first_name__icontains=parcalar[0], last_name__icontains=parcalar[-1])
        qs = qs.filter(lookup)
    data = []
    for user in qs.order_by("first_name", "last_name", "username")[:30]:
        full = user.get_full_name().strip()
        data.append({"id": user.id, "username": user.username, "label": f"{full} ({user.username})" if full else user.username})
    return JsonResponse({"results": data})


@login_required
def birim_ara(request):
    q = (request.GET.get("q") or "").strip()
    qs = _personel_veri_modeli().objects.exclude(gorev_birimi="")
    if q:
        qs = qs.filter(Q(gorev_birimi__icontains=q) | Q(alt_birimi__icontains=q))
    data = []
    seen = set()
    for row in qs.values("gorev_birimi", "alt_birimi").distinct().order_by("gorev_birimi", "alt_birimi")[:50]:
        birim = (row["gorev_birimi"] or "").strip()
        alt = (row["alt_birimi"] or "").strip()
        key = (birim, alt)
        if not birim or key in seen:
            continue
        seen.add(key)
        data.append({"birim": birim, "alt_birim": alt, "label": f"{birim}{' / ' + alt if alt else ''}"})
    return JsonResponse({"results": data})


def _personel_secim_kapsami(user):
    """Arama büyütecinde gösterilebilecek personelleri yetkiye göre sınırlar."""
    model = _personel_veri_modeli()
    if _izin_yetkisi_var(user, "tum_izinleri_goruntuleme") or _kontrol_ettigi_tipler(user):
        return model.objects.all()

    kapsam = Q(pk__in=[])
    kendi_personeli = kullanicinin_personeli(user)
    if kendi_personeli:
        if model is PortalPersonel:
            kapsam |= Q(source_personel_id=_personel_kaynak_id(kendi_personeli))
        else:
            kapsam |= Q(pk=kendi_personeli.id)
    yonetilenler = _yonettigi_personeller(user)
    if yonetilenler.exists():
        kapsam |= Q(pk__in=yonetilenler.values("id"))
    # Kullanıcının oluşturduğu, kontrol ettiği veya geçmişte parafladığı
    # evraklardaki personeller de doğrudan seçilebilir.
    gecmis_personeller = IzinTalebi.objects.filter(
        Q(kullanici=user)
        | Q(kontrol_eden=user)
        | Q(kimin_uzerinde=user)
        | Q(islem_gecmisi__kullanici=user)
    ).values("personel_id")
    gecmis_kapsami = Q(source_personel_id__in=gecmis_personeller) if model is PortalPersonel else Q(pk__in=gecmis_personeller)
    return model.objects.filter(kapsam | gecmis_kapsami).distinct()


@login_required
def personel_ara(request):
    q = (request.GET.get("q") or "").strip()
    qs = _personel_secim_kapsami(request.user)
    if q:
        lookup = (
            Q(sicil_no__icontains=q)
            | Q(ad__icontains=q)
            | Q(soyad__icontains=q)
            | Q(tc_no__icontains=q)
            | Q(gorev_birimi__icontains=q)
            | Q(alt_birimi__icontains=q)
            | Q(gorevi__icontains=q)
        )
        lookup |= Q(unvan__icontains=q) if qs.model is PortalPersonel else Q(unvan__ad__icontains=q)
        parcalar = q.split()
        if len(parcalar) >= 2:
            lookup |= Q(ad__icontains=parcalar[0], soyad__icontains=parcalar[-1])
        qs = qs.filter(lookup)
    data = [
        {
            "id": personel.id,
            "label": f"{personel.sicil_no} - {personel.ad_soyad}",
            "birim": " / ".join(filter(None, [personel.gorev_birimi, personel.alt_birimi])),
        }
        for personel in qs.order_by("ad", "soyad", "sicil_no")[:50]
    ]
    return JsonResponse({"results": data})


@login_required
def evrak_paraf(request, talep_id):
    talep = get_object_or_404(IzinTalebi.objects.select_related("izin_turu", "kullanici", "kontrol_eden", "kimin_uzerinde"), id=talep_id)
    personel = talep_personeli(talep)
    if not _talep_gorme_yetkisi(request.user, talep):
        messages.error(request, "Bu evrak\u0131 g\u00f6r\u00fcnt\u00fcleme yetkiniz yok.")
        return redirect("dashboard")
    if not personel:
        messages.error(request, "Evrağa bağlı personel kaydı bulunamadı.")
        return redirect("dashboard")
    if talep.durum == "TASLAK":
        if talep.kullanici_id == request.user.id:
            return redirect("izin_talebi_duzenle", talep_id=talep.id)
        messages.warning(request, "Taslak evrak önce sahibi tarafından tamamlanmalıdır.")
        return redirect("dashboard")
    imzacilar = talep_imzacilari(talep)
    if request.method == "POST":
        islem = request.POST.get("islem")
        if islem in {"parafla", "imzala"}:
            if not _talep_aktif_gorev_yetkisi(request.user, talep):
                messages.error(request, "Bu evrak \u015fu anda sizin \u00fczerinizde de\u011fil.")
                return redirect("dashboard")
            onceki_durum = talep.durum
            aktif_rol = "Kontrol Eden"
            if onceki_durum == "IMZA_SURECINDE" and talep.imza_sirasi < len(imzacilar):
                aktif_rol = imzacilar[talep.imza_sirasi][0]
            if talep.durum == "PARAF_BEKLIYOR":
                talep.kontrol_tarihi = timezone.now()
                talep.imza_sirasi = 0
                if imzacilar:
                    talep.durum = "IMZA_SURECINDE"
                    talep.kimin_uzerinde = imzacilar[0][1]
                    messages.success(request, f"Evrak parafland\u0131 ve {imzacilar[0][1].get_full_name() or imzacilar[0][1].username} kullan\u0131c\u0131s\u0131na g\u00f6nderildi.")
                else:
                    talep.durum = "ONAYLANDI"
                    talep.kimin_uzerinde = None
                    talep.imza_tamamlanma_tarihi = timezone.now()
                    try:
                        izni_ozluge_aktar(talep, personel, request.user)
                    except ValueError as hata:
                        messages.error(request, str(hata))
                        return redirect("dashboard")
                    except IntegrityError:
                        messages.error(request, "Özlükte kullanılan izin kaydı oluşturulamadı. Lütfen İK birimi ile görüşünüz.")
                        return redirect("dashboard")
                    messages.success(request, "Evrak parafland\u0131 ve imzac\u0131 tan\u0131m\u0131 olmad\u0131\u011f\u0131 i\u00e7in tamamland\u0131.")
            elif talep.durum == "IMZA_SURECINDE":
                sonraki_sira = talep.imza_sirasi + 1
                if sonraki_sira < len(imzacilar):
                    talep.imza_sirasi = sonraki_sira
                    talep.kimin_uzerinde = imzacilar[sonraki_sira][1]
                    messages.success(request, f"Evrak parafland\u0131 ve {imzacilar[sonraki_sira][1].get_full_name() or imzacilar[sonraki_sira][1].username} kullan\u0131c\u0131s\u0131na g\u00f6nderildi.")
                else:
                    talep.durum = "ONAYLANDI"
                    talep.kimin_uzerinde = None
                    talep.imza_tamamlanma_tarihi = timezone.now()
                    try:
                        izni_ozluge_aktar(talep, personel, request.user)
                    except ValueError as hata:
                        messages.error(request, str(hata))
                        return redirect("dashboard")
                    except IntegrityError:
                        messages.error(request, "Özlükte kullanılan izin kaydı oluşturulamadı. Lütfen İK birimi ile görüşünüz.")
                        return redirect("dashboard")
                    messages.success(request, "Evrak imzaland\u0131 ve izin s\u00fcreci tamamland\u0131. Olu\u015fturan personelin g\u00f6revlerine bilgi d\u00fc\u015ft\u00fc.")
            else:
                messages.error(request, "Bu evrak i\u00e7in aktif paraf/imza i\u015flemi bulunmuyor.")
                return redirect("dashboard")
            talep.save(update_fields=["durum", "kontrol_tarihi", "kimin_uzerinde", "imza_sirasi", "imza_tamamlanma_tarihi", "ozluk_aktarildi_mi", "ozluk_aktarim_tarihi", "guncelleme_zamani"])
            if onceki_durum == "PARAF_BEKLIYOR":
                _islem_gecmisi_ekle(talep, request.user, "KONTROL_PARAFI", aktif_rol)
            else:
                gecmis_islem = "ONAY" if talep.durum == "ONAYLANDI" else "IMZA"
                _islem_gecmisi_ekle(talep, request.user, gecmis_islem, aktif_rol)
            return redirect("dashboard")
        if islem == "reddet":
            if not _talep_aktif_gorev_yetkisi(request.user, talep):
                messages.error(request, "Bu evrak \u015fu anda sizin \u00fczerinizde de\u011fil.")
                return redirect("dashboard")
            talep.durum = "REDDEDILDI"
            talep.kimin_uzerinde = None
            talep.aciklama = request.POST.get("red_nedeni", "")
            talep.save(update_fields=["durum", "kimin_uzerinde", "aciklama", "guncelleme_zamani"])
            _islem_gecmisi_ekle(talep, request.user, "RED", "Kontrol Eden" if talep.kontrol_eden_id == request.user.id else "İmzacı", talep.aciklama or "")
            messages.error(request, "Evrak paraflamadan reddedildi.")
            return redirect("dashboard")
    return render(request, "portal/evrak_paraf.html", {
        "talep": talep,
        "personel": personel,
        "imzacilar": imzacilar,
        "indir": False,
        "aktif_gorev": _talep_aktif_gorev_yetkisi(request.user, talep),
        "son_imzaci_mi": (
            talep.durum == "IMZA_SURECINDE"
            and bool(imzacilar)
            and talep.imza_sirasi == len(imzacilar) - 1
            and talep.kimin_uzerinde_id == request.user.id
        ),
        "duzenlenebilir": talep.kullanici_id == request.user.id and talep.durum in {"IADE_EDILDI", "REDDEDILDI"},
        "goreve_baslama_tarihi": goreve_baslama_tarihi(talep),
    })


@login_required
def evrak_onizleme(request, talep_id):
    talep = get_object_or_404(IzinTalebi.objects.select_related("izin_turu", "kullanici", "kontrol_eden", "kimin_uzerinde"), id=talep_id)
    personel = talep_personeli(talep)
    if not _talep_gorme_yetkisi(request.user, talep):
        messages.error(request, "Bu evrakı görüntüleme yetkiniz yok.")
        return redirect("dashboard")
    if talep.durum == "TASLAK":
        if talep.kullanici_id == request.user.id:
            return redirect("izin_talebi_duzenle", talep_id=talep.id)
        messages.warning(request, "Taslak evrak önizlenemez.")
        return redirect("dashboard")
    return render(request, "portal/evrak_paraf.html", {"talep": talep, "personel": personel, "imzacilar": talep_imzacilari(talep), "indir": True, "goreve_baslama_tarihi": goreve_baslama_tarihi(talep)})


@login_required
def evrak_indir(request, talep_id):
    """Seçilen biçimde dosyayı doğrudan indirir; sayfa açmaz."""
    talep = get_object_or_404(IzinTalebi.objects.select_related("izin_turu", "kullanici", "kontrol_eden", "kimin_uzerinde"), id=talep_id)
    personel = talep_personeli(talep)
    if not _talep_gorme_yetkisi(request.user, talep):
        messages.error(request, "Bu evrakı indirme yetkiniz yok.")
        return redirect("dashboard")
    if talep.durum == "TASLAK":
        messages.warning(request, "Taslak evrak indirilemez.")
        return redirect("izin_talebi_duzenle", talep_id=talep.id)

    dosya_tipi = (request.GET.get("format") or "").lower()
    if dosya_tipi not in {"pdf", "excel"}:
        return redirect("evrak_onizleme", talep_id=talep.id)

    imzacilar = talep_imzacilari(talep)
    baslama = goreve_baslama_tarihi(talep)
    guvenli_no = re.sub(r"[^A-Za-z0-9_-]", "-", talep.evrak_no or str(talep.id))
    if dosya_tipi == "pdf":
        content = izin_pdf_olustur(talep, personel, imzacilar, baslama)
        response = HttpResponse(content, content_type="application/pdf")
        response["Content-Disposition"] = f'attachment; filename="izin-evraki-{guvenli_no}.pdf"'
        return response

    content = izin_excel_olustur(talep, personel, imzacilar, baslama)
    response = HttpResponse(
        content,
        content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    response["Content-Disposition"] = f'attachment; filename="izin-evraki-{guvenli_no}.xlsx"'
    return response


@login_required
def izin_ara(request):
    if request.method == "POST":
        if request.POST.get("islem") == "izin_sil":
            if not _izin_yetkisi_var(request.user, "izin_silme"):
                messages.error(request, "İzin evrakı silme yetkiniz yok.")
                return redirect("izin_ara")
            talep = get_object_or_404(IzinTalebi, id=request.POST.get("talep_id"))
            evrak_no = talep.evrak_no or str(talep.id)
            talep.delete()
            messages.success(request, f"{evrak_no} numaralı izin evrakı silindi.")
        next_url = request.POST.get("next") or ""
        return redirect(next_url if next_url.startswith("/izin-ara/") else "izin_ara")

    qs = IzinTalebi.objects.select_related("izin_turu", "kullanici", "kontrol_eden", "kimin_uzerinde").prefetch_related("islem_gecmisi__kullanici")
    qs = qs.filter(_talep_arama_filtresi(request.user)).distinct()
    q = (request.GET.get("q") or "").strip()
    evrak_no = (request.GET.get("evrak_no") or "").strip()
    durum = (request.GET.get("durum") or "").strip()
    tarih_baslangic = (request.GET.get("tarih_baslangic") or "").strip()
    tarih_bitis = (request.GET.get("tarih_bitis") or "").strip()
    birim = (request.GET.get("birim") or "").strip()
    alt_birim = (request.GET.get("alt_birim") or "").strip()
    birim_metin = (request.GET.get("birim_metin") or "").strip()
    personel_id = (request.GET.get("personel") or "").strip()
    personel_metin = (request.GET.get("personel_metin") or "").strip()
    imzalayan_id = (request.GET.get("imzalayan") or "").strip()
    imzalayan_metin = (request.GET.get("imzalayan_metin") or "").strip()
    if q:
        personel_lookup = (
            Q(sicil_no__icontains=q)
            | Q(tc_no__icontains=q)
            | Q(ad__icontains=q)
            | Q(soyad__icontains=q)
            | Q(gorev_birimi__icontains=q)
            | Q(alt_birimi__icontains=q)
            | Q(gorevi__icontains=q)
        )
        personel_lookup |= Q(unvan__icontains=q) if _personel_veri_modeli() is PortalPersonel else Q(unvan__ad__icontains=q)
        parcalar = q.split()
        if len(parcalar) >= 2:
            personel_lookup |= Q(ad__icontains=parcalar[0], soyad__icontains=parcalar[-1])
        personel_ids = _talep_personel_idleri(_personel_veri_modeli().objects.filter(personel_lookup))
        genel_eslesme = (
            Q(evrak_no__icontains=q)
            | Q(ait_oldugu_yil__icontains=q)
            | Q(izin_sebebi__icontains=q)
            | Q(aciklama__icontains=q)
            | Q(izin_turu__ad__icontains=q)
            | Q(kullanici__username__icontains=q)
            | Q(kullanici__first_name__icontains=q)
            | Q(kullanici__last_name__icontains=q)
            | Q(kontrol_eden__username__icontains=q)
            | Q(kontrol_eden__first_name__icontains=q)
            | Q(kontrol_eden__last_name__icontains=q)
            | Q(kimin_uzerinde__username__icontains=q)
            | Q(kimin_uzerinde__first_name__icontains=q)
            | Q(kimin_uzerinde__last_name__icontains=q)
            | Q(islem_gecmisi__kullanici__username__icontains=q)
            | Q(islem_gecmisi__kullanici__first_name__icontains=q)
            | Q(islem_gecmisi__kullanici__last_name__icontains=q)
            | Q(personel_id__in=personel_ids)
        )
        for durum_kodu, durum_adi in IzinTalebi.DURUM_CHOICES:
            if _normalize(q) in _normalize(durum_kodu) or _normalize(q) in _normalize(durum_adi):
                genel_eslesme |= Q(durum=durum_kodu)
        if q.isdigit() and len(q) == 4:
            genel_eslesme |= (
                Q(kurum_kayit_tarihi__year=int(q))
                | Q(baslangic_tarihi__year=int(q))
                | Q(bitis_tarihi__year=int(q))
            )
        qs = qs.filter(genel_eslesme).distinct()
    if evrak_no:
        qs = qs.filter(evrak_no__icontains=evrak_no)
    if durum:
        qs = qs.filter(durum=durum)
    for raw, lookup in ((tarih_baslangic, "kurum_kayit_tarihi__gte"), (tarih_bitis, "kurum_kayit_tarihi__lte")):
        if raw:
            try:
                qs = qs.filter(**{lookup: date.fromisoformat(raw)})
            except ValueError:
                messages.warning(request, "Geçersiz tarih filtresi dikkate alınmadı.")
    if birim or birim_metin:
        personeller = _personel_veri_modeli().objects.all()
        if birim:
            personeller = personeller.filter(gorev_birimi=birim)
            if alt_birim:
                personeller = personeller.filter(alt_birimi=alt_birim)
        else:
            personeller = personeller.filter(Q(gorev_birimi__icontains=birim_metin) | Q(alt_birimi__icontains=birim_metin))
        qs = qs.filter(personel_id__in=_talep_personel_idleri(personeller))
    secili_personel = None
    if personel_id.isdigit():
        secili_personel = _personel_secim_kapsami(request.user).filter(id=personel_id).first()
        if secili_personel:
            qs = qs.filter(personel_id=_personel_kaynak_id(secili_personel))
    elif personel_metin:
        personel_lookup = Q(sicil_no__icontains=personel_metin) | Q(ad__icontains=personel_metin) | Q(soyad__icontains=personel_metin)
        parcalar = personel_metin.split()
        if len(parcalar) >= 2:
            personel_lookup |= Q(ad__icontains=parcalar[0], soyad__icontains=parcalar[-1])
        qs = qs.filter(personel_id__in=_talep_personel_idleri(_personel_secim_kapsami(request.user).filter(personel_lookup)))
    secili_imzalayan = None
    if imzalayan_id.isdigit():
        secili_imzalayan = User.objects.filter(id=imzalayan_id, is_active=True).first()
        if secili_imzalayan:
            qs = qs.filter(islem_gecmisi__kullanici=secili_imzalayan).distinct()
    elif imzalayan_metin:
        lookup = Q(username__icontains=imzalayan_metin) | Q(first_name__icontains=imzalayan_metin) | Q(last_name__icontains=imzalayan_metin)
        parcalar = imzalayan_metin.split()
        if len(parcalar) >= 2:
            lookup |= Q(first_name__icontains=parcalar[0], last_name__icontains=parcalar[-1])
        kullanicilar = User.objects.filter(is_active=True).filter(lookup)
        qs = qs.filter(islem_gecmisi__kullanici__in=kullanicilar).distinct()

    sort = (request.GET.get("sort") or "olusturma").strip()
    direction = (request.GET.get("dir") or "desc").strip()
    sort_labels = {
        "evrak_tanimi": "Evrak Tanımı",
        "birim": "Evrakın Birimi",
        "konu": "Konu",
        "evrak_no": "Evrak No",
        "kurum_kayit": "Kurum Kayıt Tarihi",
        "olusturma": "Oluşturulma",
        "olusturan": "Oluşturan",
        "imzalayanlar": "İmzalayanlar",
        "durum": "Durum",
        "kimde": "Kimin Üzerinde",
    }
    if sort not in sort_labels:
        sort = "olusturma"
    if direction not in {"asc", "desc"}:
        direction = "desc"

    per_page_options = [10, 20, 50, 100]
    try:
        per_page = int(request.GET.get("per_page") or 10)
    except ValueError:
        per_page = 10
    if per_page not in per_page_options:
        per_page = 10

    rows = []
    for talep in qs.order_by("-olusturma_zamani", "-id"):
        personel = talep_personeli(talep)
        imzalayanlar = []
        seen = set()
        for gecmis in talep.islem_gecmisi.all():
            if gecmis.kullanici_id and gecmis.islem in {"KONTROL_PARAFI", "IMZA", "ONAY"} and gecmis.kullanici_id not in seen:
                imzalayanlar.append(gecmis.kullanici)
                seen.add(gecmis.kullanici_id)
        birim_text = ""
        if personel:
            birim_text = f"{personel.gorev_birimi or ''} {getattr(personel, 'alt_birimi', '') or ''}".strip()
        olusturan_text = talep.kullanici.get_full_name() or talep.kullanici.username
        imzalayan_text = " ".join(user.get_full_name() or user.username for user in imzalayanlar)
        kimde_text = talep.kimin_uzerinde.get_full_name() if talep.kimin_uzerinde else ""
        kimde_text = kimde_text or (talep.kimin_uzerinde.username if talep.kimin_uzerinde else "")
        rows.append({
            "talep": talep,
            "personel": personel,
            "kimde": talep.kimin_uzerinde,
            "imzalayanlar": imzalayanlar,
            "sort_values": {
                "evrak_tanimi": "İzin Talep Formu",
                "birim": birim_text,
                "konu": "İzin İsteği",
                "evrak_no": talep.evrak_no or "",
                "kurum_kayit": talep.kurum_kayit_tarihi or date.min,
                "olusturma": talep.olusturma_zamani,
                "olusturan": olusturan_text,
                "imzalayanlar": imzalayan_text,
                "durum": talep.get_durum_display(),
                "kimde": kimde_text,
            },
        })

    def row_sort_value(row):
        value = row["sort_values"].get(sort, "")
        if isinstance(value, str):
            return _normalize(value)
        return value or date.min

    rows.sort(key=row_sort_value, reverse=direction == "desc")
    paginator = Paginator(rows, per_page)
    page_obj = paginator.get_page(request.GET.get("page"))
    rows = list(page_obj.object_list)

    base_params = request.GET.copy()
    base_params["per_page"] = str(per_page)
    base_params.pop("page", None)
    pagination_query = base_params.urlencode()
    page_size_params = request.GET.copy()
    page_size_params.pop("page", None)
    page_size_params.pop("per_page", None)
    page_size_query = page_size_params.urlencode()

    sort_links = {}
    for key, label in sort_labels.items():
        params = request.GET.copy()
        params.pop("page", None)
        params["sort"] = key
        params["dir"] = "asc" if sort == key and direction == "desc" else "desc"
        sort_links[key] = {
            "key": key,
            "label": label,
            "url": params.urlencode(),
            "active": sort == key,
            "direction": direction if sort == key else "",
        }

    return render(request, "portal/izin_ara.html", {
        "rows": rows, "q": q, "evrak_no": evrak_no, "durum": durum,
        "durumlar": IzinTalebi.DURUM_CHOICES, "tarih_baslangic": tarih_baslangic,
        "tarih_bitis": tarih_bitis, "birim": birim, "alt_birim": alt_birim,
        "birim_metin": birim_metin, "secili_personel": secili_personel,
        "personel_metin": personel_metin, "secili_imzalayan": secili_imzalayan,
        "imzalayan_metin": imzalayan_metin,
        "page_obj": page_obj, "paginator": paginator, "per_page": per_page,
        "per_page_options": per_page_options, "pagination_query": pagination_query,
        "page_size_query": page_size_query, "sort_links": sort_links,
        "sort": sort, "direction": direction,
        "can_delete": _izin_yetkisi_var(request.user, "izin_silme"),
    })
