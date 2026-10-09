from datetime import date, timedelta
from copy import copy
from decimal import Decimal
import json
import mimetypes
from pathlib import Path
import re
import unicodedata

from django.contrib import messages
from django.contrib.auth.decorators import login_required, user_passes_test
from django.core.paginator import Paginator
from django.db import IntegrityError, transaction
from django.db.models import CharField, Exists, OuterRef, Q, Subquery, Sum, Value
from django.db.models.functions import Cast, Coalesce, Concat, ExtractYear, NullIf, TruncDate
from django.utils.http import url_has_allowed_host_and_scheme
from django.http import FileResponse, Http404, HttpResponse, JsonResponse
from django.shortcuts import get_object_or_404, redirect, render

from .forms import IzinEvrakForm, TcSorguForm
from .evrak_verisi import belge_verileri, evraki_sabitle, sabit_veri, sabit_kullanici
from .models import IzinTuru
from django.contrib.auth.models import User
from django.utils import timezone

from .models import GorevBirimiImzaci, HakEdilenIzin, ImzaciVekaleti, IsciPersonel, IzinDenetimKaydi, IzinEvrakEki, IzinTalebi, IzinBildirimKapatma, IzinIslemGecmisi, KaynakAltBirim, KaynakGorevBirimi, KullanilanIzin, Personel, PortalAyar, PortalIzinHakki, PortalKullanilanIzin, PortalPersonel, SirketPersonel, YonetimYetkisi
from .document_exports import izin_excel_olustur, izin_pdf_olustur
from .services import resmi_evrak_kaydini_hazirla, talep_numarasi_filtresi
from .islem_guvenligi import veritabani_yogunlugunu_yakala
from .personel_senkron import personel_aktif_mi

EVRAK_TIPLERI = {
    "memur": {"baslik": "\u0130zin Evrak\u0131 (Memur)", "calisan_tipi": "Memur", "tip_kodu": "MEMUR", "aktif": True},
    "isci": {"baslik": "\u0130zin Evrak\u0131 (\u0130\u015f\u00e7i)", "calisan_tipi": "\u0130\u015f\u00e7i", "tip_kodu": "ISCI", "aktif": True},
    "sirket": {"baslik": "\u0130zin Evrak\u0131 (\u015eirket)", "calisan_tipi": "\u015eirket", "tip_kodu": "SIRKET", "aktif": True},
}

PERSONEL_TIPI_MODELLERI = {
    "MEMUR": Personel,
    "ISCI": IsciPersonel,
    "SIRKET": SirketPersonel,
}
PERSONEL_TIPI_SLUG = {"MEMUR": "memur", "ISCI": "isci", "SIRKET": "sirket"}
PERSONEL_TIPI_ETIKET = {"MEMUR": "Memur", "ISCI": "İşçi", "SIRKET": "Şirket"}
ISLETME_DAIRESI_ADI = "İşletmeler Dairesi Başkanlığı"
ISLETME_YONETIM_ALT_KODLARI = ("ISL-01", "ISL-02", "ISL-03", "ISL-04")
ISLETME_ALT_BIRIM_GRUPLARI = {
    "ISL-01": {
        "ad": "İşletme Koordinasyon Şube Müdürlüğü",
        "on_ek": "isletme koordinasyon sube mudurlugu",
        "ilceler": (),
    },
    "ISL-02": {
        "ad": "Merkez İlçeleri İşletme Şube Müdürlüğü",
        "on_ek": "merkez ilceleri isletme sube mudurlugu",
        "ilceler": ("Ortahisar", "Maçka"),
    },
    "ISL-03": {
        "ad": "Doğu İlçeleri İşletme Şube Müdürlüğü",
        "on_ek": "dogu ilceleri isletme sube mudurlugu",
        "ilceler": ("Araklı", "Arsin", "Çaykara", "Dernekpazarı", "Hayrat", "Köprübaşı", "Of", "Sürmene", "Yomra"),
    },
    "ISL-04": {
        "ad": "Batı İlçeleri İşletme Şube Müdürlüğü",
        "on_ek": "bati ilceleri isletme sube mudurlugu",
        "ilceler": ("Akçaabat", "Beşikdüzü", "Çarşıbaşı", "Düzköy", "Şalpazarı", "Tonya", "Vakfıkebir"),
    },
}


def _normalize(text):
    text = (text or "").replace("\u0130", "i").replace("I", "i").replace("\u0131", "i").lower().strip()
    text = unicodedata.normalize("NFKD", text)
    return "".join(ch for ch in text if not unicodedata.combining(ch))


def _personel_veri_modeli(personel_tipi=None):
    """Aktarım yapıldöktan sonra portal yalnızca kendi personel kartlarını kullanır."""
    return PortalPersonel if PortalPersonel.objects.exists() else Personel


def _kaynak_personel_modeli(personel_tipi):
    return PERSONEL_TIPI_MODELLERI.get(personel_tipi, Personel)


def _personel_tip_kodu(personel):
    if isinstance(personel, PortalPersonel):
        return personel.personel_tipi or "MEMUR"
    if isinstance(personel, IsciPersonel):
        return "ISCI"
    if isinstance(personel, SirketPersonel):
        return "SIRKET"
    return "MEMUR"


def _personel_tip_slug(personel):
    return PERSONEL_TIPI_SLUG.get(_personel_tip_kodu(personel), "memur")


def _personel_kaynak_id(personel):
    return getattr(personel, "source_personel_id", personel.id)


def _ozluk_personel_filtresi(personel):
    tip = _personel_tip_kodu(personel)
    kaynak_id = _personel_kaynak_id(personel)
    if tip == "ISCI":
        return {"isci_personel_id": kaynak_id, "personel_tipi": "ISCI"}
    if tip == "SIRKET":
        return {"sirket_personel_id": kaynak_id, "personel_tipi": "SIRKET"}
    return {"personel_id": kaynak_id}


def _talep_personel_idleri(qs):
    if qs.model is PortalPersonel:
        return qs.values("source_personel_id")
    return qs.values("id")


def _turkce_upper(text):
    return (text or "").replace("i", "İ").replace("ı", "I").upper()


def _turkce_lower(text):
    return (text or "").replace("İ", "i").replace("I", "ı").lower()


def _arama_varyantlari(text):
    text = (text or "").strip()
    if not text:
        return []
    normalized = _normalize(text)
    variants = {
        text,
        text.lower(),
        text.upper(),
        text.title(),
        _turkce_lower(text),
        _turkce_upper(text),
        _turkce_lower(text).title(),
        normalized,
        normalized.upper(),
        normalized.title(),
    }
    return [variant for variant in variants if variant]


def _varyantli_icontains(alanlar, text):
    lookup = Q()
    for variant in _arama_varyantlari(text):
        for alan in alanlar:
            lookup |= Q(**{f"{alan}__icontains": variant})
    return lookup


def _personel_arama_lookup(text, personel_modeli=None):
    personel_modeli = personel_modeli or _personel_veri_modeli()
    alanlar = ["sicil_no", "tc_no", "ad", "soyad", "gorev_birimi", "alt_birimi", "gorevi"]
    lookup = _varyantli_icontains(alanlar, text)
    lookup |= _varyantli_icontains(["unvan"], text) if personel_modeli is PortalPersonel else _varyantli_icontains(["unvan__ad"], text)
    parcalar = [parca for parca in (text or "").split() if parca]
    if len(parcalar) >= 2:
        for ad in _arama_varyantlari(parcalar[0]):
            for soyad in _arama_varyantlari(parcalar[-1]):
                lookup |= Q(ad__icontains=ad, soyad__icontains=soyad)
    return lookup


def _personel_kaynak_kaydi(personel):
    return _kaynak_personel_modeli(_personel_tip_kodu(personel)).objects.filter(id=_personel_kaynak_id(personel)).first()


def _kaynak_personel_gorevi(kaynak_personel):
    for alan in ("gorevi", "fiili_gorevi", "birimdeki_gorevi", "meslek_adi"):
        deger = (getattr(kaynak_personel, alan, "") or "").strip()
        if deger:
            return deger
    return getattr(kaynak_personel, "unvani", "")


def _kaynak_personel_kullanici_adi_adayi(kaynak_personel):
    taban = _normalize(f"{getattr(kaynak_personel, 'ad', '')}{getattr(kaynak_personel, 'soyad', '')}").replace(" ", "").replace(".", "").replace("_", "")
    return re.sub(r"[^a-z0-9]", "", taban) or ""


def _kaynak_personel_arama_eslesir(kaynak_personel, q, personel_tipi=""):
    q_norm = _normalize(q)
    if not q_norm:
        return True
    q_compact = q_norm.replace(" ", "")
    tokens = [parca for parca in q_norm.split() if parca]
    metin = " ".join([
        getattr(kaynak_personel, "sicil_no", "") or "",
        getattr(kaynak_personel, "tc_no", "") or "",
        getattr(kaynak_personel, "ad", "") or "",
        getattr(kaynak_personel, "soyad", "") or "",
        getattr(kaynak_personel, "ad_soyad", "") or "",
        _kaynak_personel_kullanici_adi_adayi(kaynak_personel),
        getattr(kaynak_personel, "gorev_birimi", "") or "",
        getattr(kaynak_personel, "alt_birimi", "") or "",
        _kaynak_personel_gorevi(kaynak_personel),
        getattr(kaynak_personel, "unvani", "") or "",
        PERSONEL_TIPI_ETIKET.get(personel_tipi, personel_tipi),
    ])
    blob = _normalize(metin)
    blob_compact = blob.replace(" ", "")
    return q_norm in blob or q_compact in blob_compact or all(token in blob for token in tokens)


def _sicil_siralama_degeri(sicil_no):
    metin = str(sicil_no or "")
    if metin.isdigit():
        return (0, int(metin), metin)
    return (1, metin.casefold(), metin)


def _kullanici_adi_uret(ad, soyad, sicil_no="", zorunlu_sicil=False):
    taban = _normalize(f"{ad}{soyad}").replace(" ", "").replace(".", "").replace("_", "")
    taban = re.sub(r"[^a-z0-9]", "", taban) or "personel"
    if zorunlu_sicil and sicil_no:
        taban = f"{taban}{re.sub(r'[^0-9a-z]', '', str(sicil_no).lower())}"
    username = taban
    sayac = 2
    while User.objects.filter(username=username).exists():
        user = User.objects.filter(username=username).first()
        if user and _normalize(user.first_name) == _normalize(ad) and _normalize(user.last_name) == _normalize(soyad):
            return username
        username = f"{taban}{sayac}"
        sayac += 1
    return username


def _portal_kullanici_olustur(portal_personeli):
    if portal_personeli.kullanici_id:
        user = portal_personeli.kullanici
        if (user.first_name, user.last_name) != (portal_personeli.ad, portal_personeli.soyad):
            user.first_name, user.last_name = portal_personeli.ad, portal_personeli.soyad
            user.save(update_fields=["first_name", "last_name"])
        return user, False
    ayni_isim_sayisi = sum(
        1 for kisi in PortalPersonel.objects.only("ad", "soyad")
        if _normalize(kisi.ad) == _normalize(portal_personeli.ad)
        and _normalize(kisi.soyad) == _normalize(portal_personeli.soyad)
    )
    username = _kullanici_adi_uret(
        portal_personeli.ad,
        portal_personeli.soyad,
        portal_personeli.sicil_no,
        zorunlu_sicil=ayni_isim_sayisi > 1,
    )
    user, created = User.objects.get_or_create(
        username=username,
        defaults={
            "first_name": portal_personeli.ad,
            "last_name": portal_personeli.soyad,
            "is_active": True,
        },
    )
    # Resolve an existing link before modifying a potentially unrelated user.
    diger_kartlar = PortalPersonel.objects.filter(kullanici=user).exclude(pk=portal_personeli.pk)
    isim_cakismasi = diger_kartlar.exclude(tc_no=portal_personeli.tc_no).exists() if portal_personeli.tc_no else diger_kartlar.exists()
    if isim_cakismasi:
        taban = f"{username}{portal_personeli.personel_tipi.lower()}{portal_personeli.source_personel_id}"
        aday = taban
        sayac = 2
        while User.objects.filter(username=aday).exists():
            aday = f"{taban}{sayac}"
            sayac += 1
        user = User.objects.create_user(aday, password="12345678", first_name=portal_personeli.ad, last_name=portal_personeli.soyad)
        created = True
    degisti = False
    if user.first_name != portal_personeli.ad:
        user.first_name = portal_personeli.ad
        degisti = True
    if user.last_name != portal_personeli.soyad:
        user.last_name = portal_personeli.soyad
        degisti = True
    if created:
        user.set_password("12345678")
        degisti = True
    if degisti:
        user.save()
    portal_personeli.kullanici = user
    portal_personeli.save(update_fields=["kullanici"])
    return user, created


def _varsayilan_portal_izin_hakki(portal_personeli):
    if portal_personeli.personel_tipi == "MEMUR":
        return None
    return PortalIzinHakki.objects.get_or_create(
        personel=portal_personeli,
        izin_yili=2026,
        defaults={
            "hak_edilen_gun": Decimal("10.00"),
            "gecerlilik_baslangic_tarihi": date(2026, 1, 1),
            "gecerlilik_bitis_tarihi": date(2028, 1, 1),
            "ise_baslama_tarihi": portal_personeli.ise_baslama_tarihi,
        },
    )


@transaction.atomic
def ozlukten_portala_aktar(kaynak_personel, personel_tipi="MEMUR"):
    """Bir özlük kaydönı izin portalındaki dözenlenebilir karta kopyalar."""
    personel_tipi = personel_tipi or _personel_tip_kodu(kaynak_personel)
    defaults = {
        "personel_tipi": personel_tipi,
        "sicil_no": kaynak_personel.sicil_no,
        "ad": kaynak_personel.ad,
        "soyad": kaynak_personel.soyad,
        "tc_no": kaynak_personel.tc_no,
        "unvan": kaynak_personel.unvani,
        "cep_tel": kaynak_personel.cep_tel,
        "gorev_birimi": kaynak_personel.gorev_birimi,
        "alt_birimi": kaynak_personel.alt_birimi,
        "gorevi": _kaynak_personel_gorevi(kaynak_personel),
        "durum": kaynak_personel.durum,
        "ise_baslama_tarihi": kaynak_personel.ise_baslama_tarihi,
        "ozluk_senkron_zamani": timezone.now(),
    }
    portal_personeli, created = PortalPersonel.objects.update_or_create(
        personel_tipi=personel_tipi,
        source_personel_id=kaynak_personel.id,
        defaults=defaults,
    )
    _portal_kullanici_olustur(portal_personeli)
    _varsayilan_portal_izin_hakki(portal_personeli)
    return portal_personeli, created


def kullanicinin_personeli(user):
    bagli = PortalPersonel.objects.filter(kullanici=user).first()
    if bagli:
        return bagli
    username = (user.username or "").strip()
    for model in (PortalPersonel, Personel, IsciPersonel, SirketPersonel):
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
        for model in (PortalPersonel, Personel, IsciPersonel, SirketPersonel):
            for personel in model.objects.all():
                compact_name = _normalize(personel.ad + personel.soyad).replace(" ", "")
                compact_with_sicil = compact_name + re.sub(r"[^0-9a-z]", "", str(personel.sicil_no).lower())
                if compact_username in {compact_name, compact_with_sicil}:
                    return personel
    return None


def _kendi_personeli_mi(user, personel):
    """İzin oluıturma yalnızca oturum sahibinin kendi özlük kartı iıin yapılabilir."""
    kendi = kullanicinin_personeli(user)
    return bool(
        kendi and personel
        and _personel_tip_kodu(kendi) == _personel_tip_kodu(personel)
        and _personel_kaynak_id(kendi) == _personel_kaynak_id(personel)
    )





ORG_PARENT_NORMALIZED = {
    "1. hukuk musavirligi": "genel mudurluk",
    "rehberlik ve teftis kurulu baskanligi": "genel mudurluk",
    "ozel kalem mudurlugu": "genel mudurluk",
    "genel mudur yardimcisi idari mali": "genel mudurluk",
    "genel mudur yardimcisi idari ve mali": "genel mudurluk",
    "genel mudur yardimcisi teknik 1": "genel mudurluk",
    "genel mudur yardimcisi teknik 2": "genel mudurluk",
    "insan kaynaklari ve egitim dairesi baskanligi": "genel mudur yardimcisi idari mali",
    "abone isleri dairesi baskanligi": "genel mudur yardimcisi idari mali",
    "strateji gelistirme dairesi baskanligi": "genel mudur yardimcisi idari mali",
    "bilgi islem dairesi baskanligi": "genel mudur yardimcisi idari mali",
    "isletmeler dairesi baskanligi": "genel mudur yardimcisi teknik 1",
    "elektrik merkezleri dairesi baskanligi": "genel mudur yardimcisi teknik 1",
    "icme suyu ve kanalizasyon dairesi baskanligi": "genel mudur yardimcisi teknik 1",
    "planlama yatirim ve insaat dairesi baskanligi": "genel mudur yardimcisi teknik 2",
    "destek hizmetleri dairesi baskanligi": "genel mudur yardimcisi teknik 2",
    "tesisler dairesi baskanligi": "genel mudur yardimcisi teknik 2",
}


def _org_norm(text):
    text = _normalize(text)
    text = text.replace("(", " ").replace(")", " ").replace("-", " ")
    return " ".join(text.split())


def _isletme_birimi_mi(birim_adi):
    return _org_norm(birim_adi) == _org_norm(ISLETME_DAIRESI_ADI)


def _isletme_alt_birim_grubu(birim_adi, alt_birim_adi):
    """ılıe bazlı özlük seıimini dört gerıek ııletmeler ıubesinden birine başlar."""
    if not _isletme_birimi_mi(birim_adi):
        return None
    temiz_ad = re.sub(
        r"\bMüdürlüğü\s+Müdürlüğü\b",
        "Müdürlüğü",
        (alt_birim_adi or "").strip(),
        flags=re.IGNORECASE,
    )
    norm = _org_norm(temiz_ad)
    for kod, grup in ISLETME_ALT_BIRIM_GRUPLARI.items():
        if norm.startswith(grup["on_ek"]):
            return kod
    return None


def _isletme_kanonik_alt_birim_adi(birim_adi, alt_birim_adi):
    kod = _isletme_alt_birim_grubu(birim_adi, alt_birim_adi)
    return ISLETME_ALT_BIRIM_GRUPLARI[kod]["ad"] if kod else (alt_birim_adi or "").strip()


def _isletme_org_gorunen_ad(birim_adi, alt_birim_adi):
    kod = _isletme_alt_birim_grubu(birim_adi, alt_birim_adi)
    if not kod:
        return alt_birim_adi
    grup = ISLETME_ALT_BIRIM_GRUPLARI[kod]
    if not grup["ilceler"]:
        return grup["ad"]
    return f"{grup['ad']} ({', '.join(grup['ilceler'])})"


def _kaynak_parent_adi(kaynak_birim, kaynak_map):
    parent_key = ORG_PARENT_NORMALIZED.get(_org_norm(kaynak_birim.ad))
    if not parent_key:
        return ""
    parent = kaynak_map.get(parent_key)
    return parent.ad if parent else ""


def _user_for_personel_record(personel):
    if not personel:
        return None
    portal = PortalPersonel.objects.filter(
        personel_tipi=_personel_tip_kodu(personel),
        source_personel_id=_personel_kaynak_id(personel),
    ).first()
    adaylar = User.objects.filter(is_active=True)
    if portal:
        user = adaylar.filter(first_name__iexact=portal.ad, last_name__iexact=portal.soyad).first()
        if user:
            return user
    user = adaylar.filter(first_name__iexact=personel.ad, last_name__iexact=personel.soyad).first()
    if user:
        return user
    compact = _normalize(f"{personel.ad}{personel.soyad}").replace(" ", "")
    for user in adaylar:
        if _normalize(user.username).replace(".", "").replace("_", "").replace(" ", "") == compact:
            return user
    return None


def _kayit_amiri(kayit):
    if not kayit:
        return None
    if kayit.alt_birim_adi:
        return kayit.sube_muduru or kayit.amir_kullanici or kayit.daire_baskani
    return kayit.daire_baskani or kayit.amir_kullanici or kayit.sube_muduru


def _kayit_rolu(kayit, alt_mi=False):
    if alt_mi or (kayit and kayit.alt_birim_adi):
        return "\u015eube M\u00fcd\u00fcr\u00fc"
    tur = (kayit.birim_turu or "").upper() if kayit else ""
    ad = _org_norm(kayit.birim_adi if kayit else "")
    if tur == "GENEL_MUDURLUK":
        return "Genel M\u00fcd\u00fcr"
    if tur == "GENEL_MUDUR_YARDIMCILIGI":
        return "Genel M\u00fcd\u00fcr Yard\u0131mc\u0131s\u0131"
    if tur == "DAIRE_BASKANLIGI":
        return "Daire Ba\u015fkan\u0131"
    if "hukuk" in ad:
        return "1. Hukuk M\u00fc\u015faviri"
    if "teftis" in ad:
        return "Kurul Ba\u015fkan\u0131"
    if "ozel kalem" in ad:
        return "\u00d6zel Kalem M\u00fcd\u00fcr\u00fc"
    return "Birim Amiri"


def _personel_gorev_unvan_metni(personel):
    """Personelin görev/ınvan metinlerini tek normalize alanda toplar."""
    if not personel:
        return ""
    parcalar = []
    for alan in ("unvan", "gorevi", "fiili_gorevi", "kadro_unvani"):
        deger = getattr(personel, alan, "")
        if hasattr(deger, "ad"):
            deger = deger.ad
        if deger:
            parcalar.append(str(deger))
    return _org_norm(" ".join(parcalar))


def _personel_sube_muduru_mu(personel):
    metin = _personel_gorev_unvan_metni(personel)
    return "sube muduru" in metin or "sube mudur" in metin


def _personel_daire_baskani_mu(personel):
    metin = _personel_gorev_unvan_metni(personel)
    return "daire baskani" in metin or "daire baskan" in metin


def _personel_genel_mudur_yardimcisi_mu(personel):
    metin = _personel_gorev_unvan_metni(personel)
    return "genel mudur yardimcisi" in metin


def gorev_birimlerini_senkronla():
    kaynaklar = list(KaynakGorevBirimi.objects.select_related("amir_personel").prefetch_related("alt_birimler__amir_personel").all())
    kaynak_map = {_org_norm(k.ad): k for k in kaynaklar}
    mevcut = set()
    if kaynaklar:
        for kaynak in kaynaklar:
            birim = (kaynak.ad or "").strip()
            if not birim:
                continue
            ust_birim = _kaynak_parent_adi(kaynak, kaynak_map)
            amir_user = _user_for_personel_record(kaynak.amir_personel)
            kayit, _ = GorevBirimiImzaci.objects.get_or_create(birim_adi=birim, alt_birim_adi="")
            if not amir_user:
                amir_user = kayit.amir_kullanici or kayit.daire_baskani or kayit.sube_muduru
            kayit.birim_turu = kaynak.birim_turu or "DIGER"
            kayit.ust_birim_adi = ust_birim
            kayit.amir_kullanici = amir_user
            kayit.daire_baskani = amir_user
            kayit.sube_muduru = None
            kayit.aktif_mi = bool(kaynak.aktif_mi)
            kayit.save(update_fields=["birim_turu", "ust_birim_adi", "amir_kullanici", "daire_baskani", "sube_muduru", "aktif_mi", "guncelleme_zamani"])
            mevcut.add((birim, ""))
            alt_birimler = list(kaynak.alt_birimler.all())
            if _isletme_birimi_mi(birim):
                alt_birimler = [
                    alt
                    for alt in alt_birimler
                    if (alt.kod or "").strip().upper() in ISLETME_YONETIM_ALT_KODLARI
                ]
                alt_birimler.sort(
                    key=lambda alt: ISLETME_YONETIM_ALT_KODLARI.index((alt.kod or "").strip().upper())
                )
            for alt in alt_birimler:
                alt_ad = (alt.ad or "").strip()
                if not alt_ad:
                    continue
                if _isletme_birimi_mi(birim):
                    alt_ad = _isletme_kanonik_alt_birim_adi(birim, alt_ad)
                alt_user = _user_for_personel_record(alt.amir_personel)
                alt_kayit, _ = GorevBirimiImzaci.objects.get_or_create(birim_adi=birim, alt_birim_adi=alt_ad)
                if not alt_user:
                    alt_user = alt_kayit.amir_kullanici or alt_kayit.sube_muduru or alt_kayit.daire_baskani
                alt_kayit.birim_turu = "ALT_BIRIM"
                alt_kayit.ust_birim_adi = birim
                alt_kayit.amir_kullanici = alt_user
                alt_kayit.sube_muduru = alt_user
                alt_kayit.daire_baskani = None
                alt_kayit.aktif_mi = bool(alt.aktif_mi and kaynak.aktif_mi)
                alt_kayit.save(update_fields=["birim_turu", "ust_birim_adi", "amir_kullanici", "daire_baskani", "sube_muduru", "aktif_mi", "guncelleme_zamani"])
                mevcut.add((birim, alt_ad))
        for kayit in GorevBirimiImzaci.objects.all():
            anahtar = (kayit.birim_adi, kayit.alt_birim_adi)
            aktif_olmali = anahtar in mevcut
            if kayit.aktif_mi != aktif_olmali:
                kayit.aktif_mi = aktif_olmali
                kayit.save(update_fields=["aktif_mi", "guncelleme_zamani"])
        return mevcut

    for row in _personel_veri_modeli().objects.values("gorev_birimi", "alt_birimi").distinct():
        birim = (row.get("gorev_birimi") or "").strip()
        alt = (row.get("alt_birimi") or "").strip()
        if not birim:
            continue
        mevcut.add((birim, alt))
        GorevBirimiImzaci.objects.get_or_create(birim_adi=birim, alt_birim_adi=alt)
    return mevcut

def _user_label(user):
    if not user:
        return ""
    full = user.get_full_name().strip()
    return f"{full} ({user.username})" if full else user.username


def _suresi_dolan_vekaletleri_pasife_al(tarih=None):
    """Bitiı tarihini geıen vekıletleri silmeden otomatik pasifleıtirir."""
    tarih = tarih or timezone.localdate()
    qs = ImzaciVekaleti.objects.filter(
        aktif_mi=True,
        bitis_tarihi__lt=tarih,
    ).select_related("asil_kullanici", "vekil_kullanici")
    kullanici_idleri = set()
    for vekalet in qs:
        kullanici_idleri.add(vekalet.asil_kullanici_id)
        kullanici_idleri.add(vekalet.vekil_kullanici_id)
    sayi = qs.update(aktif_mi=False)
    if sayi:
        _bekleyen_vekalet_evraklarini_guncelle(kullanicilar=kullanici_idleri, kaydeden=None)
    return sayi


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
    return ImzaciVekaleti.objects.filter(
        asil_kullanici=user,
        aktif_mi=True,
        asil_kullanici__is_active=True,
        vekil_kullanici__is_active=True,
        baslangic_tarihi__lte=tarih,
        bitis_tarihi__gte=tarih,
    ).select_related("vekil_kullanici", "asil_kullanici").order_by("-baslangic_tarihi", "-id").first()


def _imzaci_rol_ve_kullanici(rol, user, tarih):
    asil_id = user.pk
    vekalet = _aktif_vekalet_kaydi(user, tarih)
    if vekalet:
        rol, user = (rol if rol.endswith("V.") else f"{rol} V."), vekalet.vekil_kullanici
    user = copy(user)
    user._asil_kullanici_id = asil_id
    return rol, user


def _birim_kaydi(birim_adi, alt_birim_adi=""):
    if alt_birim_adi:
        alt_birim_adi = _isletme_kanonik_alt_birim_adi(birim_adi, alt_birim_adi)
    return GorevBirimiImzaci.objects.filter(
        birim_adi=birim_adi or "",
        alt_birim_adi=alt_birim_adi or "",
        aktif_mi=True,
    ).select_related("amir_kullanici", "daire_baskani", "sube_muduru").first()


def _ust_birim_kaydi(kayit):
    if not kayit or not kayit.ust_birim_adi:
        return None
    return _birim_kaydi(kayit.ust_birim_adi, "")


def _ekle_imzaci(liste, kayit, tarih, rol=None):
    user = _kayit_amiri(kayit)
    if not user:
        liste.append((rol or _kayit_rolu(kayit), None))
        return
    rol = rol or _kayit_rolu(kayit)
    rol, user = _imzaci_rol_ve_kullanici(rol, user, tarih)
    if not any(mevcut_user and mevcut_user.id == user.id for _, mevcut_user in liste):
        liste.append((rol, user))


def _gorev_birimi_sira_degeri(ad):
    norm = _org_norm(ad)
    siralar = {
        "genel mudurluk": 0,
        "ozel kalem mudurlugu": 10,
        "rehberlik ve teftis kurulu baskanligi": 11,
        "1. hukuk musavirligi": 12,
        "genel mudur yardimcisi idari mali": 20,
        "genel mudur yardimcisi idari ve mali": 20,
        "insan kaynaklari ve egitim dairesi baskanligi": 21,
        "abone isleri dairesi baskanligi": 22,
        "strateji gelistirme dairesi baskanligi": 23,
        "bilgi islem dairesi baskanligi": 24,
        "genel mudur yardimcisi teknik 1": 40,
        "isletmeler dairesi baskanligi": 41,
        "isletme koordinasyon sube mudurlugu": 410,
        "merkez ilceleri isletme sube mudurlugu": 411,
        "dogu ilceleri isletme sube mudurlugu": 412,
        "bati ilceleri isletme sube mudurlugu": 413,
        "elektrik merkezleri dairesi baskanligi": 42,
        "icme suyu ve kanalizasyon dairesi baskanligi": 43,
        "genel mudur yardimcisi teknik 2": 60,
        "planlama yatirim ve insaat dairesi baskanligi": 61,
        "destek hizmetleri dairesi baskanligi": 62,
        "tesisler dairesi baskanligi": 63,
    }
    return siralar.get(norm, 500)


def _gorev_birimi_agaci(kayitlar):
    """Görev birimi imzacı kayötlarını özlükteki gibi katmanlı organizasyon açacına ıevirir."""
    node_map = {}
    for kayit in kayitlar:
        key = f"{kayit.birim_adi}||{kayit.alt_birim_adi or ''}"
        alt_mi = bool(kayit.alt_birim_adi)
        secili = kayit.sube_muduru if alt_mi else kayit.daire_baskani
        node_map[key] = {
            "key": key.replace(" ", "_").replace("/", "_").replace("|", "_"),
            "kayit": kayit,
            "title": (
                _isletme_org_gorunen_ad(kayit.birim_adi, kayit.alt_birim_adi)
                if alt_mi
                else kayit.birim_adi
            ),
            "subtitle": kayit.birim_adi if alt_mi else "",
            "type_label": _kayit_rolu(kayit, alt_mi=alt_mi) if alt_mi else _kayit_rolu(kayit),
            "field_name": f"sube_{kayit.id}" if alt_mi else f"daire_{kayit.id}",
            "other_field_name": f"daire_{kayit.id}" if alt_mi else f"sube_{kayit.id}",
            "selected_user": secili,
            "selected_user_id": secili.id if secili else "",
            "children": [],
            "depth": 0,
            "order": _gorev_birimi_sira_degeri(kayit.alt_birim_adi or kayit.birim_adi),
            "is_alt": alt_mi,
        }

    roots = []
    for node in node_map.values():
        kayit = node["kayit"]
        if kayit.alt_birim_adi:
            parent_key = f"{kayit.birim_adi}||"
        elif kayit.ust_birim_adi:
            parent_key = f"{kayit.ust_birim_adi}||"
        else:
            parent_key = ""

        parent = node_map.get(parent_key)
        if parent and parent is not node:
            parent["children"].append(node)
        else:
            roots.append(node)

    def sirala_ve_derinlik(nodes, depth=0):
        nodes.sort(key=lambda n: (n["order"], _normalize(n["title"])))
        for node in nodes:
            node["depth"] = depth
            sirala_ve_derinlik(node["children"], depth + 1)

    sirala_ve_derinlik(roots)
    return roots


def imzacilari_bul(personel, tarih=None):
    tarih = tarih or timezone.localdate()
    if not personel:
        return []
    birim = (personel.gorev_birimi or "").strip()
    alt = (getattr(personel, "alt_birimi", "") or "").strip()
    ana = _birim_kaydi(birim, "")
    alt_kayit = _birim_kaydi(birim, alt) if alt else None
    genel_mudur = GorevBirimiImzaci.objects.filter(birim_turu="GENEL_MUDURLUK", aktif_mi=True).select_related("amir_kullanici", "daire_baskani", "sube_muduru").first()
    personel_user = _user_for_personel_record(_personel_kaynak_kaydi(personel) or personel)
    daire_baskani_mu = _personel_daire_baskani_mu(personel)
    genel_mudur_yardimcisi_mu = _personel_genel_mudur_yardimcisi_mu(personel)
    imzacilar = []

    if alt_kayit:
        alt_amir = _kayit_amiri(alt_kayit)
        secili_sube_muduru_mu = bool(
            personel_user
            and GorevBirimiImzaci.objects.filter(
                birim_adi=birim,
                sube_muduru=personel_user,
                aktif_mi=True,
            ).exclude(alt_birim_adi="").exists()
        )
        if secili_sube_muduru_mu or (personel_user and alt_amir and personel_user.id == alt_amir.id):
            _ekle_imzaci(imzacilar, ana, tarih)
            _ekle_imzaci(imzacilar, _ust_birim_kaydi(ana), tarih)
        else:
            _ekle_imzaci(imzacilar, alt_kayit, tarih, "\u015eube M\u00fcd\u00fcr\u00fc")
            _ekle_imzaci(imzacilar, ana, tarih)
        return imzacilar

    ana_amir = _kayit_amiri(ana)
    if daire_baskani_mu or genel_mudur_yardimcisi_mu or (personel_user and ana_amir and personel_user.id == ana_amir.id):
        tur = (ana.birim_turu or "").upper() if ana else ""
        if daire_baskani_mu or tur == "DAIRE_BASKANLIGI":
            _ekle_imzaci(imzacilar, _ust_birim_kaydi(ana), tarih)
            _ekle_imzaci(imzacilar, genel_mudur, tarih)
        elif genel_mudur_yardimcisi_mu or tur in {"GENEL_MUDUR_YARDIMCILIGI", "DOGRU_GENEL_MUDURLUK", "DIGER"}:
            _ekle_imzaci(imzacilar, genel_mudur, tarih)
        return imzacilar

    _ekle_imzaci(imzacilar, ana, tarih)
    return imzacilar

def imzaci_metni(personel, tarih=None):
    imzacilar = imzacilari_bul(personel, tarih)
    if not imzacilar:
        return "\u0130mzac\u0131 tan\u0131mlanmam\u0131\u015f"
    return "\n".join(f"{rol}: {_user_label(user)}" for rol, user in imzacilar)


def talep_personeli(talep):
    return (
        PortalPersonel.objects.filter(personel_tipi=talep.personel_tipi, source_personel_id=talep.personel_id).first()
        or _kaynak_personel_modeli(talep.personel_tipi).objects.filter(id=talep.personel_id).first()
    )


def talep_imzacilari(talep):
    veri = sabit_veri(talep)
    if veri:
        return [(item["rol"], sabit_kullanici(item["kullanici"])) for item in veri["imzacilar"]]
    personel = talep_personeli(talep)
    return imzacilari_bul(personel, talep.kurum_kayit_tarihi or timezone.localdate()) if personel else []


def talep_aktif_imzacilari(talep):
    """Kontrol parafını veren kiıiyi tekrar imza görevine döıırmez.

    Talep sahibi, bir amire vekalet ettiği için imza makamında da yer
    alabilir. Bu durumda kişi aynı olsa bile rol farklıdır; belge ve görev
    akışı ayrı imza adımını korumalıdır.
    """
    veri = sabit_veri(talep)
    if veri:
        imzacilar = []
        for item in veri["imzacilar"]:
            asil_id = item.get("asil_kullanici_id", item["kullanici"]["id"])
            asil = User.objects.filter(pk=asil_id, is_active=True).first()
            user = _vekil_kullanici(asil)
            rol = item["rol"].removesuffix(" V.")
            if user and user.pk != asil_id:
                rol += " V."
            imzacilar.append((rol, user))
        return imzacilar
    atlanacaklar = {talep.kontrol_eden_id}
    return [(rol, user) for rol, user in talep_imzacilari(talep) if not user or user.id not in atlanacaklar]


def _talep_bekleyen_hedefini_bul(talep):
    if talep.durum == "PARAF_BEKLIYOR":
        teklif_parafi_verildi = _teklif_parafi_verildi_mi(talep)
        if teklif_parafi_verildi:
            return kontrol_edeni_bul(talep.personel_tipi)
        if not talep.kontrol_eden or not _kontrol_yetkisi_var(talep.kontrol_eden, talep.personel_tipi):
            talep.kontrol_eden = kontrol_edeni_bul(talep.personel_tipi)
            talep.save(update_fields=["kontrol_eden"])
        return talep.kullanici if talep.kullanici.is_active else None
    if talep.durum == "IMZA_SURECINDE":
        imzacilar = talep_aktif_imzacilari(talep)
        if talep.imza_sirasi < len(imzacilar):
            return imzacilar[talep.imza_sirasi][1]
        return None
    return talep.kimin_uzerinde


def _bekleyen_talep_hedefini_guncelle(talep, kaydeden=None):
    hedef = _talep_bekleyen_hedefini_bul(talep)
    hedef_id = hedef.id if hedef else None
    veri = sabit_veri(talep)
    beklenen_id = None
    if veri and talep.durum == "PARAF_BEKLIYOR" and _teklif_parafi_verildi_mi(talep):
        beklenen_id = (veri.get("kontrol_eden") or {}).get("id")
    elif veri and talep.durum == "IMZA_SURECINDE" and talep.imza_sirasi < len(veri["imzacilar"]):
        beklenen_id = veri["imzacilar"][talep.imza_sirasi]["kullanici"]["id"]
    if beklenen_id and beklenen_id != hedef_id:
        # Frozen contents remain untouched. An authority change starts a new
        # editable round at the owner, never rewrites the signed snapshot.
        onceki_durum = talep.durum
        talep.durum = "IADE_EDILDI"
        talep.kimin_uzerinde = talep.kullanici
        talep.imza_sirasi = 0
        talep.kontrol_tarihi = None
        talep.save(update_fields=["durum", "kimin_uzerinde", "imza_sirasi", "kontrol_tarihi", "guncelleme_zamani"])
        aciklama = "İmza/kontrol yetkisi değiştiği için evrak yeni paraf turu için ilk sahibine iade edildi. Sabit belge bilgileri değiştirilmedi."
        _islem_gecmisi_ekle(talep, kaydeden, "IADE", "Yetki Değişikliği", aciklama)
        _denetim_kaydi_ekle(talep, kaydeden, "IADE", onceki_durum, aciklama, talep.kullanici)
        return True
    kontrol_degisti = talep.durum == "PARAF_BEKLIYOR" and _teklif_parafi_verildi_mi(talep) and talep.kontrol_eden_id != hedef_id
    if talep.kimin_uzerinde_id == hedef_id and not kontrol_degisti:
        return False
    onceki = talep.kimin_uzerinde
    talep.kimin_uzerinde = hedef
    alanlar = ["kimin_uzerinde", "guncelleme_zamani"]
    if talep.durum == "PARAF_BEKLIYOR" and _teklif_parafi_verildi_mi(talep):
        talep.kontrol_eden = hedef
        alanlar.append("kontrol_eden")
    talep.save(update_fields=alanlar)
    _denetim_kaydi_ekle(
        talep,
        kaydeden,
        "GUNCELLE",
        talep.durum,
        "Vekalet durumu değiştiği için bekleyen evrak sahibine/asıl imzacıya aktarıldı.",
        hedef,
    )
    return bool(onceki or hedef)


def _bekleyen_vekalet_evraklarini_guncelle(kullanicilar=None, kaydeden=None):
    qs = IzinTalebi.objects.select_related("kullanici", "kontrol_eden", "kimin_uzerinde").filter(
        durum__in=["PARAF_BEKLIYOR", "IMZA_SURECINDE"]
    )
    if kullanicilar:
        user_ids = [user.id if hasattr(user, "id") else user for user in kullanicilar if user]
        qs = qs.filter(Q(kimin_uzerinde_id__in=user_ids) | Q(kullanici_id__in=user_ids) | Q(kontrol_eden_id__in=user_ids))
    sayi = 0
    for talep_id in qs.values_list("pk", flat=True):
        with transaction.atomic():
            talep = IzinTalebi.objects.select_for_update().select_related("kullanici", "kontrol_eden", "kimin_uzerinde").get(pk=talep_id)
            if talep.durum in {"PARAF_BEKLIYOR", "IMZA_SURECINDE"} and _bekleyen_talep_hedefini_guncelle(talep, kaydeden=kaydeden):
                sayi += 1
    return sayi


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
    """Personel tipi iıin tek geçerli kontrol edeni kullanıcı yetkilerinden bulur."""
    alan = {
        "MEMUR": "memur_kontrol_eden",
        "ISCI": "isci_kontrol_eden",
        "SIRKET": "sirket_kontrol_eden",
    }.get(personel_tipi, "memur_kontrol_eden")
    yetki = (
        YonetimYetkisi.objects
        .select_related("kullanici")
        .filter(aktif_mi=True, kullanici__is_active=True, **{alan: True})
        .order_by("guncelleme_zamani", "id")
        .last()
    )
    return yetki.kullanici if yetki else None


def _islem_gecmisi_ekle(talep, kullanici, islem, rol="", aciklama=""):
    return IzinIslemGecmisi.objects.create(
        talep=talep,
        kullanici=kullanici,
        islem=islem,
        rol=rol,
        aciklama=aciklama[:300],
    )


def _denetim_kaydi_ekle(talep, kullanici, islem, onceki_durum="", aciklama="", hedef_kullanici=None):
    return IzinDenetimKaydi.objects.create(
        talep=talep if getattr(talep, "pk", None) else None,
        evrak_no=talep.referans_no,
        islem=islem,
        kullanici=kullanici if getattr(kullanici, "pk", None) else None,
        hedef_kullanici=hedef_kullanici if getattr(hedef_kullanici, "pk", None) else None,
        onceki_durum=onceki_durum or "",
        yeni_durum=getattr(talep, "durum", "") or "",
        aciklama=(aciklama or "")[:500],
    )

def _teklif_parafi_verildi_mi(talep):
    """Talep sahibinin son onaya sunmadan sonraki ilk parafı verilmiı mi?"""
    if getattr(talep, "imza_sirasi", 0) > 0:
        return True
    qs = talep.islem_gecmisi.filter(islem="IMZA", rol="Teklif Eden")
    # Evrak daha ınce paraflayan kullanıcıya geri çekildiyse, teklif sahibinin
    # yeni paraf turunu eski turdaki paraftan başımsız başlatürüz.
    son_baslangic = talep.islem_gecmisi.filter(
        Q(islem="ONAYA_SUNULDU") | Q(islem="GERI_CEK", rol="İzin Talebinde Bulunan")
    ).order_by("-tarih", "-id").first()
    if son_baslangic:
        qs = qs.filter(tarih__gt=son_baslangic.tarih)
    return qs.exists()


def _geri_cekme_hedefleri(talep):
    """Mevcut turda gerıekten paraf/imza vermiı geri dönıı hedeflerini ıretir."""
    aktif_imzacilar = {
        user.id: (sira, rol)
        for sira, (rol, user) in enumerate(talep_aktif_imzacilari(talep))
        if user
    }
    son_onaya_sunma = talep.islem_gecmisi.filter(islem="ONAYA_SUNULDU").order_by("-tarih", "-id").first()
    gecmisler = talep.islem_gecmisi.filter(
        islem__in=["KONTROL_PARAFI", "IMZA"]
    ).select_related("kullanici").order_by("-tarih", "-id")
    if son_onaya_sunma:
        gecmisler = gecmisler.filter(tarih__gt=son_onaya_sunma.tarih)

    hedefler, gorulenler = [], set()
    for gecmis in gecmisler:
        user = gecmis.kullanici
        if not user or user.id in gorulenler:
            continue
        if gecmis.islem == "KONTROL_PARAFI":
            hedefler.append({
                "kullanici": user,
                "ad_soyad": user.get_full_name() or user.username,
                "rol": "Kontrol Eden",
                "durum": "PARAF_BEKLIYOR",
                "imza_sirasi": 1,
            })
        elif user.id == talep.kullanici_id or gecmis.rol == "Teklif Eden":
            hedefler.append({
                "kullanici": user,
                "ad_soyad": user.get_full_name() or user.username,
                "rol": "İzin Talebinde Bulunan",
                "durum": "IADE_EDILDI",
                "imza_sirasi": 0,
            })
        elif user.id in aktif_imzacilar:
            sira, rol = aktif_imzacilar[user.id]
            hedefler.append({
                "kullanici": user,
                "ad_soyad": user.get_full_name() or user.username,
                "rol": rol or "Parafçı",
                "durum": "IMZA_SURECINDE",
                "imza_sirasi": sira,
            })
        else:
            continue
        gorulenler.add(user.id)
    return hedefler

PERSONEL_TIPI_PREFIX = {"MEMUR": "memur", "ISCI": "isci", "SIRKET": "sirket"}
TIP_BAZLI_YETKI_EYLEMLERI = (
    "izin_goruntuleme",
    "izin_arama",
    "kontrol_eden",
    "evrak_geri_cekme",
    "sahibine_iade",
    "izin_silme",
    "tamamlanan_izin_bildirimi",
)
ESKI_GENEL_YETKI_KARSILIKLARI = {
    "izin_goruntuleme": ("tum_izinleri_goruntuleme",),
    "izin_arama": ("tum_izinleri_goruntuleme", "evrak_yonetimi"),
    "evrak_geri_cekme": ("evrak_geri_cekme",),
    "sahibine_iade": ("sahibine_iade",),
    "izin_silme": ("izin_silme",),
    "tamamlanan_izin_bildirimi": ("tamamlanan_izin_bildirimi",),
}
YONETIM_YETKI_ALANLARI = tuple(alan.name for alan in YonetimYetkisi._meta.fields if alan.get_internal_type() == 'BooleanField' and alan.name not in {'aktif_mi', 'birim_ve_vekalet_yonetimi'})

def _kullanici_yetkisi(user):
    if not user or not user.is_authenticated or not user.is_active:
        return None
    try:
        yetki = user.portal_yetki
    except YonetimYetkisi.DoesNotExist:
        return None
    return yetki if yetki.aktif_mi else None


def _yetki_form_degerleri(yetki):
    degerler = {alan: bool(yetki and yetki.aktif_mi and getattr(yetki, alan, False)) for alan in YONETIM_YETKI_ALANLARI}
    eski_birlesik = bool(yetki and yetki.aktif_mi and getattr(yetki, "birim_ve_vekalet_yonetimi", False))
    degerler["gorev_birimi_yonetimi"] = bool(degerler.get("gorev_birimi_yonetimi") or eski_birlesik)
    degerler["vekalet_yonetimi"] = bool(degerler.get("vekalet_yonetimi") or eski_birlesik)
    return degerler


def _tamamlanan_bildirim_filtresi(user):
    """
    Tamamlanan evrak bildirimleri yönetici/staff olma durumundan başımsızdör.

    Her kullanıcı kendi evrakını ve kontrol ettiği evrakı görır. Bunun döıındaki
    bildirimler yalnızca Kullanıcı Yetkileri ekranında açıkıa seçilen genel veya
    personel tipi bazlı "Tamamlanan bildirimleri al" yetkileriyle açılır.
    """
    yetki = _kullanici_yetkisi(user)
    if yetki and yetki.tamamlanan_izin_bildirimi:
        return None

    filtre = Q(kullanici=user) | Q(kontrol_eden=user)
    if yetki:
        for personel_tipi, prefix in PERSONEL_TIPI_PREFIX.items():
            if getattr(yetki, f"{prefix}_tamamlanan_izin_bildirimi", False):
                filtre |= Q(personel_tipi=personel_tipi)
    return filtre


def _reddedilen_bildirimler(user):
    return IzinTalebi.objects.filter(
        durum="REDDEDILDI", kontrol_eden=user,
    ).exclude(kullanici=user).exclude(reddeden=user).filter(
        Q(kontrol_tarihi__isnull=False)
        | Q(islem_gecmisi__islem="KONTROL_PARAFI", islem_gecmisi__kullanici=user)
    ).distinct()


def _kapatilabilir_izin_bildirimleri(user):
    tamamlananlar = IzinTalebi.objects.filter(durum="ONAYLANDI")
    filtre = _tamamlanan_bildirim_filtresi(user)
    if filtre is not None:
        tamamlananlar = tamamlananlar.filter(filtre)
    return IzinTalebi.objects.filter(
        Q(pk__in=tamamlananlar.values("pk"))
        | Q(pk__in=_reddedilen_bildirimler(user).values("pk"))
    )


def _izin_yetkisi_var(user, alan):
    if not user or not user.is_authenticated or not user.is_active:
        return False
    if user.is_superuser and not alan.endswith("kontrol_eden"):
        return True
    yetki = _kullanici_yetkisi(user)
    return bool(yetki and getattr(yetki, alan, False))


def _kontrol_yetkisi_var(user, personel_tipi):
    alan = {
        "MEMUR": "memur_kontrol_eden",
        "ISCI": "isci_kontrol_eden",
        "SIRKET": "sirket_kontrol_eden",
    }.get(personel_tipi)
    return bool(alan and _izin_yetkisi_var(user, alan))


def _tip_yetki_alani(personel_tipi, eylem):
    prefix = PERSONEL_TIPI_PREFIX.get(personel_tipi or "MEMUR", "memur")
    return f"{prefix}_{eylem}"


def _tip_yetkisi_var(user, personel_tipi, eylem):
    """Memur/iııi/Şirket bazlı evrak yetkisi."""
    if not user or not user.is_authenticated or not user.is_active or personel_tipi not in PERSONEL_TIPI_PREFIX:
        return False
    if user.is_superuser and eylem != "kontrol_eden":
        return True
    if eylem != "kontrol_eden" and _tam_yonetim_yetkisi(user):
        return True
    if eylem == "kontrol_eden":
        return _kontrol_yetkisi_var(user, personel_tipi)
    yetki = _kullanici_yetkisi(user)
    if not yetki:
        return False
    alan = _tip_yetki_alani(personel_tipi, eylem)
    if hasattr(yetki, alan) and getattr(yetki, alan, False):
        return True
    return any(getattr(yetki, eski, False) for eski in ESKI_GENEL_YETKI_KARSILIKLARI.get(eylem, ()))


def _tip_yetkili_tipler(user, eylem):
    if eylem != "kontrol_eden" and _tam_yonetim_yetkisi(user):
        return list(PERSONEL_TIPI_PREFIX.keys())
    return [tip for tip in PERSONEL_TIPI_PREFIX if _tip_yetkisi_var(user, tip, eylem)]


def _genis_izin_arama_yetkisi(user):
    return bool(
        user.is_authenticated
        and (
            user.is_superuser
            or _izin_yetkisi_var(user, "evrak_yonetimi")
            or _tip_yetkili_tipler(user, "izin_arama")
            or _tip_yetkili_tipler(user, "izin_goruntuleme")
            or _kontrol_ettigi_tipler(user)
        )
    )


def _aktif_vekalet_asilleri(user, tarih=None):
    """Kullanıcının kendi hesabı ile o gön vekil olduğu asıl hesapların kimlikleri."""
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
    """Amirin görebileceıi özlük personelini döndörır.

    Daire başkanı kendi ana birimindeki bıtın alt birimleri; ıube mıdörı ise
    yalnızca tanımlı alt birimdeki personeli görır. Gıncel vekaletler de bu
    kapsama döhildir.
    """
    asil_ids = _aktif_vekalet_asilleri(user, tarih)
    if not asil_ids:
        return _personel_veri_modeli().objects.none()

    personel_modeli = _personel_veri_modeli()
    kapsam = Q(pk__in=[])
    # Ana birimdeki daire başkanı, o daireye bağlı tım ıubeleri kapsar.
    for birim in GorevBirimiImzaci.objects.filter(
        alt_birim_adi="", daire_baskani_id__in=asil_ids, aktif_mi=True
    ).values_list("birim_adi", flat=True):
        kapsam |= Q(gorev_birimi=birim)
    # ıube mıdörı yalnızca kendi alt birimini kapsar.
    for birim, alt_birim in GorevBirimiImzaci.objects.filter(
        sube_muduru_id__in=asil_ids, aktif_mi=True
    ).exclude(alt_birim_adi="").values_list("birim_adi", "alt_birim_adi"):
        if _isletme_birimi_mi(birim):
            kanonik_ad = _isletme_kanonik_alt_birim_adi(birim, alt_birim)
            eslesen_idler = [
                personel.pk
                for personel in personel_modeli.objects.filter(gorev_birimi=birim).only("pk", "alt_birimi")
                if _isletme_kanonik_alt_birim_adi(birim, personel.alt_birimi) == kanonik_ad
            ]
            kapsam |= Q(pk__in=eslesen_idler)
        else:
            kapsam |= Q(gorev_birimi=birim, alt_birimi=alt_birim)
    return personel_modeli.objects.filter(kapsam)


def _kontrol_ettigi_tipler(user):
    """Kullanıcı yetkilerinden, kiıinin kontrolcı olduğu personel tiplerini döndörır."""
    if not user or not user.is_authenticated:
        return []
    yetki = _kullanici_yetkisi(user)
    if not yetki:
        return []
    alanlar = {
        "MEMUR": "memur_kontrol_eden",
        "ISCI": "isci_kontrol_eden",
        "SIRKET": "sirket_kontrol_eden",
    }
    return [tip for tip, alan in alanlar.items() if getattr(yetki, alan, False)]


def _talep_arama_filtresi(user):
    """Aramada kullanıcının akııta gerıekten yetkili olduğu evrakları gösterir."""
    # Onaya sunulmuı fakat teklif sahibi henız paraflamamıı evrak, bir taslak
    # niteliıindedir: sahibinin görevinde durur; genel arama listesine girmez.
    # İade/red evrakının paylaçılabilmesi iıin en az ilk teklif parafı verilmiı
    # olmalıdör. Bıylece sahibinin kendi ilk açamada reddettiği evrak gereksiz
    # biıimde yöneticilerin aramasında görınmez.
    if not user or not user.is_authenticated or not user.is_active:
        return Q(pk__in=[])
    # Correlated history prevents signatures from a previous submission from
    # exposing an uninitialled resubmission to its former participants.
    son_baslangic = IzinIslemGecmisi.objects.filter(talep_id=OuterRef("talep_id")).filter(
        Q(islem="ONAYA_SUNULDU") | Q(islem="GERI_CEK", rol="İzin Talebinde Bulunan")
    ).order_by("-id").values("id")[:1]
    bu_tur = IzinIslemGecmisi.objects.filter(talep_id=OuterRef("pk"), pk__gt=Coalesce(Subquery(son_baslangic), 0))
    teklif_var = Q(Exists(bu_tur.filter(islem="IMZA", rol="Teklif Eden")))
    imza_var = Q(Exists(bu_tur.filter(islem__in=["IMZA", "KONTROL_PARAFI", "ONAY"])))
    kullanici_imzasi = Q(Exists(bu_tur.filter(kullanici=user, islem__in=["IMZA", "KONTROL_PARAFI", "ONAY"])))
    akista_ilerleyen = (
        Q(durum__in=["IMZA_SURECINDE", "ONAYLANDI"])
        | (Q(durum="PARAF_BEKLIYOR") & teklif_var)
        | (Q(durum__in=["IADE_EDILDI", "REDDEDILDI"]) & imza_var)
    )
    sahibinin_evraklari = (
        Q(kullanici=user, durum__in=["TASLAK", "IADE_EDILDI", "REDDEDILDI"])
        | (Q(kullanici=user) & kullanici_imzasi)
    )
    # Aktif görev, ilk paraf bekleyen teklif sahibinin kendi evrağı değilse
    # aramadan da açılabilsin (kontrol eden veya sıradaki imzacı).
    aktif_gorev = Q(kimin_uzerinde=user) & ~Q(kullanici=user, durum="PARAF_BEKLIYOR")
    # Bir amirin, imza listesinde adö bulunuyor diye evrakı sırası gelmeden
    # görmemesi gerekir. Aıık evrakta görınırlık yalnızca evrak kendisindeyse
    # veya bu turda daha ınce gerıekten paraf/imza vermiıse doıar.
    akista_islem_yapmis = Q(kimin_uzerinde=user) | kullanici_imzasi
    kapsam = Q(kontrol_eden=user) | Q(islem_gecmisi__kullanici=user)
    tipler = set(_kontrol_ettigi_tipler(user))
    tipler.update(_tip_yetkili_tipler(user, "izin_goruntuleme"))
    tipler.update(_tip_yetkili_tipler(user, "izin_arama"))
    if tipler:
        kapsam |= Q(personel_tipi__in=list(tipler))
    yonetilenler = _yonettigi_personeller(user)
    if yonetilenler.exists():
        if yonetilenler.model is PortalPersonel:
            yonetilen_q = Q(pk__in=[])
            for tip, source_id in yonetilenler.values_list("personel_tipi", "source_personel_id"):
                yonetilen_q |= Q(personel_tipi=tip, personel_id=source_id)
            kapsam |= yonetilen_q
        else:
            kapsam |= Q(personel_tipi="MEMUR", personel_id__in=_talep_personel_idleri(yonetilenler))
    # Tım izinleri görme yetkisi kapsamı geniıletir; ilk parafı bekleyen
    # evrakları yine aramaya açmaz.
    if (
        _tam_yonetim_yetkisi(user)
        or _izin_yetkisi_var(user, "tum_izinleri_goruntuleme")
        or _izin_yetkisi_var(user, "evrak_yonetimi")
    ):
        return sahibinin_evraklari | aktif_gorev | akista_ilerleyen

    # Tip bazlı "evrakları gör/ara" yetkisi, tamamlanmıı evrakları ilgili
    # kapsamda açar. Ancak devam eden evrakta amir veya kontrol eden, kendi
    # adömına ulaçmadan hiıbir kayöt görmez.
    tamamlanmis_kapsam = Q(durum="ONAYLANDI") & kapsam
    return sahibinin_evraklari | aktif_gorev | (akista_ilerleyen & (akista_islem_yapmis | tamamlanmis_kapsam))

def _talep_gorme_yetkisi(user, talep):
    if not user or not user.is_authenticated or not user.is_active:
        return False
    # Genel arama ekranında ilk teklif parafını bekleyen evraklar görınmez.
    # Bu gizleme, evrak o anda kullanıcının görevindeyken açılmasını
    # engellememelidir: teklif sahibi kendi ilk parafını, kontrol eden veya
    # imzacı da yalnızca kendi üzerindeki aktif görevi her zaman görmelidir.
    if _talep_aktif_gorev_yetkisi(user, talep):
        return True
    if talep.durum == "REDDEDILDI" and _reddedilen_bildirimler(user).filter(pk=talep.pk).exists():
        return True
    return IzinTalebi.objects.filter(pk=talep.pk).filter(_talep_arama_filtresi(user)).exists()

def _talep_aktif_gorev_yetkisi(user, talep):
    # Yınetici veya staff olmak evrakı görüntüleme/yönetme yetkisi verebilir;
    # ancak başka bir kullanıcının yerine paraf ya da e-imza atma yetkisi
    # vermez. Akıı iılemini yalnızca evrak o anda kimin üzerindeyse yapar.
    if not user or not user.is_authenticated or not user.is_active or talep.kimin_uzerinde_id != user.id:
        return False
    if talep.durum == "PARAF_BEKLIYOR":
        if not _teklif_parafi_verildi_mi(talep):
            return talep.kullanici_id == user.id
        return talep.kontrol_eden_id == user.id and _kontrol_yetkisi_var(user, talep.personel_tipi)
    if talep.durum == "IMZA_SURECINDE":
        imzacilar = talep_aktif_imzacilari(talep)
        return 0 <= talep.imza_sirasi < len(imzacilar) and bool(imzacilar[talep.imza_sirasi][1]) and imzacilar[talep.imza_sirasi][1].pk == user.pk
    return False


ANA_YONETICI_KULLANICI_ADI = "erenaygun"
ONAYLI_SILME_KILITLI_DURUMLAR = {
    "ONAYLANDI",
    "IMZALANDI",
    "TAMAMLANDI",
    "PARAF_IMZA_TAMAMLANDI",
}


def _ana_yonetici_mi(user):
    kullanici_adi = user.get_username() if user and user.is_authenticated else ""
    return bool(kullanici_adi and user.is_active and user.is_superuser and kullanici_adi.casefold() == ANA_YONETICI_KULLANICI_ADI.casefold())


def _ana_yonetici_kullanicisi():
    return User.objects.filter(username__iexact=ANA_YONETICI_KULLANICI_ADI, is_active=True, is_superuser=True).first()


def _talep_onayli_silme_kilitli_mi(talep):
    if not talep:
        return False
    durum = (getattr(talep, "durum", "") or "").upper()
    if durum in ONAYLI_SILME_KILITLI_DURUMLAR:
        return True
    for alan in (
        "onaylanma_tarihi",
        "onay_tarihi",
        "imzalanma_tarihi",
        "imza_tamamlanma_tarihi",
        "tamamlanma_tarihi",
        "surec_tamamlanma_tarihi",
    ):
        if getattr(talep, alan, None):
            return True
    return False


def _posttan_silme_sifresi_al(request):
    for alan in ("yonetici_sifre", "admin_sifre", "admin_password", "password", "sifre", "silme_sifresi"):
        deger = request.POST.get(alan)
        if deger:
            return deger
    return ""


def _onayli_talep_silme_dogrula(request, talep):
    """Onaylanan evrakı yalnızca ana yönetici doıru şifreyle silebilir."""
    if not _talep_onayli_silme_kilitli_mi(talep):
        return True, ""
    if not _ana_yonetici_mi(request.user):
        return False, "Onaylanan izin evrakını yalnızca ana yönetici silebilir."
    girilen_sifre = _posttan_silme_sifresi_al(request).strip()
    if not girilen_sifre:
        return False, "Onaylanan evrakı silmek için ana yönetici şifresi zorunludur."
    # Onaylı evrak silme yetkisi devredilemez: yalnızca ana yönetici kendi
    # hesabının parolasıyla iılem yapabilir. Tarayöcıdan gelen herhangi bir
    # şifre deıeri tek başına asla yeterli değildir.
    ana_yonetici = _ana_yonetici_kullanicisi()
    if (
        not ana_yonetici
        or request.user.pk != ana_yonetici.pk
        or not request.user.check_password(girilen_sifre)
    ):
        return False, "Ana yönetici şifresi hatalı. Evrak silinmedi."
    return True, ""


def _talep_silinebilir_mi(user, talep):
    """Aıık evraklarda yetkili/evrak sahibi silebilir; onaylanan evrak yalnızca ana yönetici tarafından silinir."""
    if not user or not user.is_authenticated or not user.is_active:
        return False
    if _talep_onayli_silme_kilitli_mi(talep):
        return _ana_yonetici_mi(user)
    tip = getattr(talep, "personel_tipi", "MEMUR")
    if _tip_yetkisi_var(user, tip, "izin_silme"):
        return True
    return talep.kullanici_id == user.id and talep.durum in {"TASLAK", "IADE_EDILDI", "REDDEDILDI"}


def _evrak_surumu_eski(request, talep):
    surum = request.POST.get("evrak_surumu") if request.method == "POST" else None
    if surum is not None and surum != talep.guncelleme_zamani.isoformat():
        messages.warning(request, "Evrak başka bir işlemle güncellendi. Güncel ekranı kontrol ederek işlemi tekrar yapın.")
        return True
    return False


def _post_talebi(request):
    talep_id = request.POST.get("talep_id", "")
    if not talep_id.isdigit():
        raise Http404("Geçersiz evrak.")
    talep = get_object_or_404(IzinTalebi.objects.select_for_update(), pk=talep_id)
    if not _talep_gorme_yetkisi(request.user, talep):
        raise Http404("Evrak bulunamadı.")
    return talep


def _guvenli_donus(request):
    hedef = request.META.get("HTTP_REFERER") or ""
    return hedef if url_has_allowed_host_and_scheme(hedef, {request.get_host()}, require_https=request.is_secure()) else "dashboard"


EK_DOSYA_UZANTILARI = {".pdf", ".png", ".jpg", ".jpeg", ".doc", ".docx", ".xls", ".xlsx"}
EK_DOSYA_AZAMI_BOYUT = 10 * 1024 * 1024


def _talep_ek_duzenlenebilir_mi(kullanici, talep):
    """Ekler yalnızca evrak sahibince, akıı başlamadan ınce değiştirilebilir."""
    return bool(
        kullanici
        and kullanici.is_authenticated
        and kullanici.is_active
        and kullanici.id == talep.kullanici_id
        and talep.durum in {"TASLAK", "IADE_EDILDI", "REDDEDILDI"}
    )

def _herhangi_tip_yetkisi_var(user, eylemler=None):
    eylemler = eylemler or TIP_BAZLI_YETKI_EYLEMLERI
    return any(_tip_yetkili_tipler(user, eylem) for eylem in eylemler)


def _yonetim_erisim_yetkisi(user):
    return _tam_yonetim_yetkisi(user) or _personel_veri_yonetimi_yetkisi(user) or _birim_ve_vekalet_yetkisi(user) or _evrak_yonetim_yetkisi(user)


def _tam_yonetim_yetkisi(user):
    return bool(
        user
        and user.is_authenticated
        and user.is_active
        and (
            user.is_superuser
            or _izin_yetkisi_var(user, "tam_yonetim_yetkisi")
        )
    )


def _personel_veri_yonetimi_yetkisi(user):
    return _tam_yonetim_yetkisi(user) or _izin_yetkisi_var(user, "personel_veri_yonetimi")


def _evrak_yonetim_yetkisi(user):
    return bool(
        user and user.is_authenticated and user.is_active and (
            user.is_superuser
            or _tam_yonetim_yetkisi(user)
            or _izin_yetkisi_var(user, "evrak_yonetimi")
            or _izin_yetkisi_var(user, "evrak_geri_cekme")
            or _izin_yetkisi_var(user, "sahibine_iade")
            or _izin_yetkisi_var(user, "izin_silme")
            or _herhangi_tip_yetkisi_var(user, ("izin_goruntuleme", "izin_arama", "evrak_geri_cekme", "sahibine_iade", "izin_silme"))
            or _kontrol_ettigi_tipler(user)
        )
    )


def _birim_ve_vekalet_yetkisi(user):
    return _gorev_birimi_yonetimi_yetkisi(user) or _vekalet_yonetimi_yetkisi(user)


def _gorev_birimi_yonetimi_yetkisi(user):
    return _tam_yonetim_yetkisi(user) or _izin_yetkisi_var(user, "gorev_birimi_yonetimi") or _izin_yetkisi_var(user, "birim_ve_vekalet_yonetimi")


def _vekalet_yonetimi_yetkisi(user):
    return _tam_yonetim_yetkisi(user) or _izin_yetkisi_var(user, "vekalet_yonetimi") or _izin_yetkisi_var(user, "birim_ve_vekalet_yonetimi")


def izin_ozeti(personel, izin_turu=None):
    ozluk_filtre = _ozluk_personel_filtresi(personel)
    yillik = izin_turu is None or izin_turu.yillik_izin_mi
    tur_filtre = Q(izin_turu__isnull=True) | Q(izin_turu__yillik_izin_mi=True) if yillik else Q(izin_turu=izin_turu)
    haklar = list(HakEdilenIzin.objects.filter(tur_filtre, **ozluk_filtre).order_by("-izin_yili"))
    kullanimlar = KullanilanIzin.objects.filter(
        **ozluk_filtre,
        hak_edilen_izin__isnull=False,
    ).values("hak_edilen_izin_id").annotate(toplam=Sum("kullanilan_gun"))
    kullanim_map = {row["hak_edilen_izin_id"]: row["toplam"] or Decimal("0.00") for row in kullanimlar}
    toplam_kalan = Decimal("0.00")
    for hak in haklar:
        hak.kullanilan = kullanim_map.get(hak.id, Decimal("0.00"))
        hak.kalan = (hak.hak_edilen_gun - hak.kullanilan).quantize(Decimal("0.01"))
        toplam_kalan += hak.kalan
    if haklar or not yillik or _personel_tip_kodu(personel) == "MEMUR":
        return haklar, toplam_kalan.quantize(Decimal("0.01"))

    portal_personeli = personel if isinstance(personel, PortalPersonel) else PortalPersonel.objects.filter(
        personel_tipi=_personel_tip_kodu(personel),
        source_personel_id=_personel_kaynak_id(personel),
    ).first()
    if not portal_personeli:
        return [], Decimal("0.00")
    haklar = list(PortalIzinHakki.objects.filter(personel=portal_personeli).order_by("-izin_yili"))
    kullanimlar = PortalKullanilanIzin.objects.filter(
        personel=portal_personeli,
        hak_edilen_izin__isnull=False,
    ).values("hak_edilen_izin_id").annotate(toplam=Sum("kullanilan_gun"))
    kullanim_map = {row["hak_edilen_izin_id"]: row["toplam"] or Decimal("0.00") for row in kullanimlar}
    toplam_kalan = Decimal("0.00")
    for hak in haklar:
        hak.kullanilan = kullanim_map.get(hak.id, Decimal("0.00"))
        hak.kalan = (hak.hak_edilen_gun - hak.kullanilan).quantize(Decimal("0.01"))
        toplam_kalan += hak.kalan
    return haklar, toplam_kalan.quantize(Decimal("0.01"))


def _yillara_gore_haklar(personel, ait_oldugu_yil, izin_turu=None):
    yillik = izin_turu is None or izin_turu.yillik_izin_mi
    tur_filtre = Q(izin_turu__isnull=True) | Q(izin_turu__yillik_izin_mi=True) if yillik else Q(izin_turu=izin_turu)
    qs = HakEdilenIzin.objects.filter(tur_filtre, **_ozluk_personel_filtresi(personel))
    if yillik and _personel_tip_kodu(personel) != "MEMUR" and not qs.exists():
        portal_personeli = personel if isinstance(personel, PortalPersonel) else PortalPersonel.objects.filter(
            personel_tipi=_personel_tip_kodu(personel),
            source_personel_id=_personel_kaynak_id(personel),
        ).first()
        qs = PortalIzinHakki.objects.filter(personel=portal_personeli) if portal_personeli else PortalIzinHakki.objects.none()
    yillar = [int(yil) for yil in re.findall(r"\d{4}", ait_oldugu_yil or "")]
    if yillar:
        qs = qs.filter(izin_yili__in=yillar)
    return list(qs.order_by("izin_yili", "id"))

def izin_bakiyesi_onaya_sunma_kontrolu(talep, personel):
    """İzin bakiyesi yalnızca evrak ilk kez onaya sunulurken kontrol edilir."""
    mesaj = "Seçilen yıl için yeterli izin bakiyesi bulunamadı. Eğer yanlışlık olduğunu düşünüyorsanız İK ile iletişime geçiniz!"
    if not talep.izin_turu_id:
        return None
    gun = Decimal(talep.kullanilan_gun or 0)
    if gun <= 0:
        return None

    haklar = _yillara_gore_haklar(personel, talep.ait_oldugu_yil, talep.izin_turu)
    if not haklar:
        return mesaj if talep.izin_turu.yillik_izin_mi else None

    portal_hakki_mi = bool(haklar and isinstance(haklar[0], PortalIzinHakki))
    if portal_hakki_mi:
        portal_personeli = personel if isinstance(personel, PortalPersonel) else PortalPersonel.objects.filter(
            personel_tipi=_personel_tip_kodu(personel),
            source_personel_id=_personel_kaynak_id(personel),
        ).first()
        kullanilanlar = PortalKullanilanIzin.objects.filter(
            personel=portal_personeli,
            hak_edilen_izin_id__in=[hak.id for hak in haklar],
        ).values("hak_edilen_izin_id").annotate(toplam=Sum("kullanilan_gun"))
    else:
        kullanilanlar = KullanilanIzin.objects.filter(
            **_ozluk_personel_filtresi(personel),
            hak_edilen_izin_id__in=[hak.id for hak in haklar],
        ).values("hak_edilen_izin_id").annotate(toplam=Sum("kullanilan_gun"))

    kullanilan_map = {row["hak_edilen_izin_id"]: row["toplam"] or Decimal("0.00") for row in kullanilanlar}
    toplam_kalan = Decimal("0.00")
    for hak in haklar:
        toplam_kalan += max(Decimal("0.00"), Decimal(hak.hak_edilen_gun) - Decimal(kullanilan_map.get(hak.id, 0)))
    if toplam_kalan < gun:
        return mesaj
    return None


def izni_ozluge_aktar(talep, personel, islemi_yapan=None):
    if talep.ozluk_aktarildi_mi:
        return
    if not talep.izin_turu_id or not talep.baslangic_tarihi or not talep.bitis_tarihi or not talep.kullanilan_gun:
        raise ValueError("Özlük aktarımı için izin türü, tarih aralığı ve izin süresi gereklidir.")

    gun = Decimal(talep.kullanilan_gun)
    if gun <= 0 or talep.bitis_tarihi < talep.baslangic_tarihi:
        raise ValueError("İzin süresi pozitif ve tarih aralığı geçerli olmalıdır.")
    sabit_tur = getattr(talep, "evrak_verisi", {}).get("izin_turu") or {}
    yillik_izin_mi = sabit_tur.get("yillik_izin_mi", talep.izin_turu.yillik_izin_mi)
    tip = talep.personel_tipi or _personel_tip_kodu(personel)
    kaynak = _personel_kaynak_kaydi(personel)
    if not kaynak:
        raise ValueError("Özlükte kaynak personel kaydı bulunamadı.")

    personel_filtre = _ozluk_personel_filtresi(personel)
    ortak_kayit = {
        **personel_filtre,
        "personel_tipi": tip,
        "izin_turu_id": talep.izin_turu_id,
        "izin_baslangic_tarihi": talep.baslangic_tarihi,
        "izin_bitis_tarihi": talep.bitis_tarihi,
        "ise_baslama_tarihi": getattr(talep, "goreve_baslama_tarihi", None) or getattr(kaynak, "ise_baslama_tarihi", None),
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
        # Claim the transfer once; simultaneous final-signature requests must
        # not create duplicate usage, including on SQLite.
        if getattr(talep, "pk", None):
            if not IzinTalebi.objects.filter(pk=talep.pk, ozluk_aktarildi_mi=False).update(ozluk_aktarildi_mi=True):
                mevcut = IzinTalebi.objects.only("ozluk_aktarildi_mi", "ozluk_aktarim_tarihi").get(pk=talep.pk)
                talep.ozluk_aktarildi_mi = mevcut.ozluk_aktarildi_mi
                talep.ozluk_aktarim_tarihi = mevcut.ozluk_aktarim_tarihi
                return
        bakiye = getattr(talep, "evrak_verisi", {}).get("bakiye")
        if bakiye is not None:
            model = PortalIzinHakki if bakiye["portal"] else HakEdilenIzin
            kapsam = {"personel_id": personel.pk} if model is PortalIzinHakki else personel_filtre
            haklar = list(model.objects.select_for_update().filter(pk__in=bakiye["hak_ids"], **kapsam).order_by("izin_yili", "id"))
        else:
            haklar = _yillara_gore_haklar(personel, talep.ait_oldugu_yil, talep.izin_turu)
            if haklar:
                haklar = list(type(haklar[0]).objects.select_for_update().filter(pk__in=[hak.pk for hak in haklar]).order_by("izin_yili", "id"))
        if yillik_izin_mi or haklar or (bakiye and bakiye["hak_ids"]):
            portal_hakki_mi = bool(haklar and isinstance(haklar[0], PortalIzinHakki))
            if portal_hakki_mi:
                portal_personeli = personel if isinstance(personel, PortalPersonel) else PortalPersonel.objects.filter(
                    personel_tipi=tip,
                    source_personel_id=_personel_kaynak_id(personel),
                ).first()
                kullanilanlar = PortalKullanilanIzin.objects.filter(
                    personel=portal_personeli,
                    hak_edilen_izin_id__in=[hak.id for hak in haklar],
                ).values("hak_edilen_izin_id").annotate(toplam=Sum("kullanilan_gun"))
            else:
                kullanilanlar = KullanilanIzin.objects.filter(
                    **personel_filtre,
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
                aktarim_plani.append((hak, parca))
                kalan -= parca
                if kalan <= 0:
                    break
            if kalan > 0 or not aktarim_plani:
                raise ValueError("Tanımlı izin bakiyesi bu evrak için yeterli değil. Özlükteki izin hakkı kontrol edilmelidir.")
            else:
                for hak, parca in aktarim_plani:
                    KullanilanIzin.objects.create(
                        hak_edilen_izin_id=None if isinstance(hak, PortalIzinHakki) else hak.id,
                        kullanilan_gun=parca,
                        aciklama=f"İzin Portalı evrakı: {talep.evrak_no}",
                        **ortak_kayit,
                    )
                    if isinstance(hak, PortalIzinHakki):
                        PortalKullanilanIzin.objects.create(
                            personel_id=hak.personel_id, izin_turu_id=talep.izin_turu_id,
                            hak_edilen_izin=hak, izin_baslangic_tarihi=talep.baslangic_tarihi,
                            izin_bitis_tarihi=talep.bitis_tarihi, kullanilan_gun=parca,
                            ise_baslama_tarihi=ortak_kayit["ise_baslama_tarihi"],
                            aciklama=f"İzin Portalı evrakı: {talep.evrak_no}", olusturan=islemi_yapan,
                        )
        else:
            KullanilanIzin.objects.create(
                kullanilan_gun=gun,
                aciklama=f"İzin Portalı evrakı: {talep.evrak_no}",
                **ortak_kayit,
            )
        talep.ozluk_aktarildi_mi = True
        talep.ozluk_aktarim_tarihi = timezone.now()
        talep.save(update_fields=["ozluk_aktarildi_mi", "ozluk_aktarim_tarihi"])

def _evrak_formu_kwargs(personel, haklar, toplam_kalan, tip="memur", **extra):
    bilgi = _evrak_tipi_or_404(tip)
    return {
        "personel": personel,
        "haklar": haklar,
        "tur_haklari": {str(tur.pk): izin_ozeti(personel, tur)[0] for tur in IzinTuru.objects.filter(yillik_izin_mi=False)},
        "calisan_tipi": bilgi["calisan_tipi"],
        "imzacilar": imzaci_metni(personel),
        "toplam_izin": toplam_kalan,
        **extra,
    }


@transaction.atomic
def _talebi_kaydet(talep, form, *, personel, kullanici, evrak_no, onaya_sun, personel_tipi="MEMUR"):
    """Formdan gelen veriyi hem yeni hem iade edilen taslaça gövenle uygular."""
    if onaya_sun and not personel_aktif_mi(personel):
        return None, "Pasif personelin izin evraki onaya sunulamaz."
    talep = form.save(commit=False)
    talep.evrak_verisi = {}
    talep.kullanici = kullanici
    talep.personel_id = _personel_kaynak_id(personel)
    talep.personel_tipi = personel_tipi
    talep.evrak_no = evrak_no
    talep.izin_sebebi = ""
    talep.adres = ""
    talep.kontrol_tarihi = None
    talep.imza_sirasi = 0
    talep.imza_tamamlanma_tarihi = None

    if onaya_sun:
        bakiye_hatasi = izin_bakiyesi_onaya_sunma_kontrolu(talep, personel)
        if bakiye_hatasi:
            return None, bakiye_hatasi
        kontrol_eden = kontrol_edeni_bul(personel_tipi)
        if not kontrol_eden:
            return None, "Kontrol eden personel henüz tanımlanmamış. Lütfen İK birimi ile görüşünüz."
        if not _kontrol_yetkisi_var(kontrol_eden, personel_tipi):
            return None, "Tanımlı kontrol eden personelin bu personel tipi için yetkisi yok. Lütfen İK birimi ile görüşünüz."
        talep.kontrol_eden = kontrol_eden
        talep.durum = "PARAF_BEKLIYOR"
        talep.kimin_uzerinde = kullanici
        talep.red_nedeni = ""
        talep.reddeden = None
        talep.reddedilme_tarihi = None
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
    _denetim_kaydi_ekle(talep, kullanici, "ONAYA_SUN" if onaya_sun else "GUNCELLE")
    if onaya_sun:
        IzinBildirimKapatma.objects.filter(talep=talep).delete()
    return talep, None

def _evrak_tipi_or_404(tip):
    bilgi = EVRAK_TIPLERI.get(tip)
    if not bilgi:
        from django.http import Http404
        raise Http404("Evrak tipi bulunamad\u0131")
    return bilgi


@login_required
def dashboard(request):
    _bekleyen_vekalet_evraklarini_guncelle(kullanicilar=[request.user], kaydeden=None)
    personel = kullanicinin_personeli(request.user)
    haklar, toplam_kalan, son_kullanimlar, talepler = [], Decimal("0.00"), [], []
    if personel:
        haklar, toplam_kalan = izin_ozeti(personel)
        personel_kaynak_id = _personel_kaynak_id(personel)
        personel_tipi = _personel_tip_kodu(personel)
        # Bu alan paraf verilen evrakları değil, yalnızca giriı yapan kiıinin
        # kendi oluıturup imza süreci tamamlanan izinlerini gösterir.
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
                personel_tipi=personel_tipi,
                durum="ONAYLANDI",
            ).select_related("izin_turu").order_by("-bitis_tarihi", "-guncelleme_zamani")[:6]
        ]
        talepler = IzinTalebi.objects.filter(kullanici=request.user, personel_id=personel_kaynak_id, personel_tipi=personel_tipi)[:6]
    gorev_filtresi = (
        Q(kimin_uzerinde=request.user, durum="IMZA_SURECINDE")
        | Q(kimin_uzerinde=request.user, kullanici=request.user, durum="PARAF_BEKLIYOR")
        | Q(kullanici=request.user, durum__in=["TASLAK", "IADE_EDILDI", "REDDEDILDI"])
    )
    for tip in ("MEMUR", "ISCI", "SIRKET"):
        if _kontrol_yetkisi_var(request.user, tip):
            gorev_filtresi |= Q(kimin_uzerinde=request.user, durum="PARAF_BEKLIYOR", personel_tipi=tip)
    gorevler = IzinTalebi.objects.select_related("izin_turu", "kullanici", "kimin_uzerinde").filter(
        gorev_filtresi
    ).distinct().order_by("-olusturma_zamani")[:25]
    reddedilenler = _reddedilen_bildirimler(request.user).select_related("izin_turu", "kullanici").exclude(
        bildirim_kapatmalari__kullanici=request.user
    ).order_by("-reddedilme_tarihi", "-guncelleme_zamani")[:25]

    imzalananlar_qs = IzinTalebi.objects.select_related("izin_turu", "kullanici").filter(durum="ONAYLANDI")
    bildirim_filtresi = _tamamlanan_bildirim_filtresi(request.user)
    if bildirim_filtresi is not None:
        imzalananlar_qs = imzalananlar_qs.filter(bildirim_filtresi)
    imzalananlar = imzalananlar_qs.exclude(
        bildirim_kapatmalari__kullanici=request.user
    ).order_by("-guncelleme_zamani")[:10]
    return render(request, "portal/dashboard.html", {"personel": personel, "haklar": haklar, "toplam_kalan": toplam_kalan, "son_kullanimlar": son_kullanimlar, "talepler": talepler, "gorevler": gorevler, "reddedilenler": reddedilenler, "imzalananlar": imzalananlar})


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
            tip_kodu = bilgi["tip_kodu"]
            tc_no = form.cleaned_data["tc_no"]
            if PortalPersonel.objects.filter(personel_tipi=tip_kodu).exists():
                personel = PortalPersonel.objects.filter(personel_tipi=tip_kodu, tc_no=tc_no).first()
                farkli_tip = PortalPersonel.objects.filter(tc_no=tc_no).exclude(personel_tipi=tip_kodu).first()
            else:
                personel = _kaynak_personel_modeli(tip_kodu).objects.filter(tc_no=tc_no).first()
                farkli_tip = None
                for diger_tip, model in PERSONEL_TIPI_MODELLERI.items():
                    if diger_tip != tip_kodu:
                        farkli_tip = model.objects.filter(tc_no=tc_no).first()
                        if farkli_tip:
                            break
            if personel:
                if not personel_aktif_mi(personel):
                    messages.error(request, "Pasif personel icin yeni izin evraki olusturulamaz.")
                    return redirect("tc_sorgula", tip=tip)
                if not _kendi_personeli_mi(request.user, personel):
                    messages.error(request, "Kendi hesabınızla yalnızca kendi T.C. kimlik numaranız için izin evrakı oluşturabilirsiniz.")
                    return redirect("tc_sorgula", tip=tip)
                return redirect("izin_evrak", tip=tip, personel_id=personel.id)
            if farkli_tip:
                dogru_tip = _personel_tip_kodu(farkli_tip)
                dogru_slug = PERSONEL_TIPI_SLUG.get(dogru_tip, "memur")
                messages.error(request, f"Bu T.C. Kimlik No {PERSONEL_TIPI_ETIKET.get(dogru_tip, 'personel')} kaydında görünüyor. Lütfen {EVRAK_TIPLERI[dogru_slug]['baslik']} ekranından devam edin.")
                return redirect("tc_sorgula", tip=dogru_slug)
            messages.error(request, "Bu T.C. Kimlik No ile kay\u0131tl\u0131 personel bulunamad\u0131.")
    return render(request, "portal/tc_sorgula.html", {"form": form, "tip": tip, "bilgi": bilgi, "personel": personel})


@login_required
@veritabani_yogunlugunu_yakala
@transaction.atomic
def izin_evrak(request, tip, personel_id):
    bilgi = _evrak_tipi_or_404(tip)
    if not bilgi["aktif"]:
        messages.error(request, f"{bilgi['baslik']} için personel kaynağı henüz bağlanmadı.")
        return redirect("tc_sorgula", tip=tip)
    tip_kodu = bilgi["tip_kodu"]
    if PortalPersonel.objects.filter(personel_tipi=tip_kodu).exists():
        personel = get_object_or_404(PortalPersonel, id=personel_id, personel_tipi=tip_kodu)
    else:
        personel = get_object_or_404(_kaynak_personel_modeli(tip_kodu), id=personel_id)
    if not _kendi_personeli_mi(request.user, personel):
        messages.error(request, "Başka personel adına izin evrakı oluşturamazsınız.")
        return redirect("tc_sorgula", tip=tip)
    if not personel_aktif_mi(personel):
        messages.error(request, "Pasif personel icin yeni izin evraki olusturulamaz.")
        return redirect("dashboard")
    haklar, toplam_kalan = izin_ozeti(personel)
    imzacilar = imzacilari_bul(personel)
    if not haklar or toplam_kalan <= 0:
        messages.warning(request, "İzin bakiyesi bulunamadı. Lütfen İK birimi ile görüşünüz.")
    if not imzacilar:
        messages.warning(request, "Bu personelin birimi için imzacı tanımlanmamış. Lütfen İK birimi ile görüşünüz.")
    if request.method == "GET":
        acik_taslak = IzinTalebi.objects.filter(
            kullanici=request.user,
            personel_id=_personel_kaynak_id(personel),
            personel_tipi=bilgi["tip_kodu"],
            durum="TASLAK",
        ).order_by("-guncelleme_zamani").first()
        if acik_taslak:
            messages.info(request, "Bu kişi için yarım bırakılmış bir izin evrakı var. Kaldığınız yerden devam ediyorsunuz.")
            return redirect("izin_talebi_duzenle", talep_id=acik_taslak.id)
        taslak = IzinTalebi.objects.create(
            kullanici=request.user,
            personel_id=_personel_kaynak_id(personel),
            personel_tipi=bilgi["tip_kodu"],
            durum="TASLAK",
            kimin_uzerinde=request.user,
        )
        _islem_gecmisi_ekle(taslak, request.user, "TASLAK", "Teklif Eden")
        _denetim_kaydi_ekle(taslak, request.user, "OLUSTUR")
        messages.info(request, "İzin evrakı taslak olarak oluşturuldu. Dilediğiniz zaman Görevlerim alanından kaldığınız yerden devam edebilirsiniz.")
        return redirect("izin_talebi_duzenle", talep_id=taslak.id)
    evrak_no = ""
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
                messages.error(request, hata)
            elif taslak:
                messages.success(request, f"Evrak taslak olarak kaydedildi. Takip No: {talep.takip_no}")
                return redirect("dashboard")
            else:
                messages.success(request, f"İzin evrakı ilk parafınız için size yönlendirildi. Takip No: {talep.takip_no}")
                return redirect("evrak_paraf", talep_id=talep.id)
        else:
            messages.error(request, "İzin evrakında hatalı veya eksik alanlar var. Lütfen kırmızı uyarıları kontrol edin.")
    else:
        form = IzinEvrakForm(next_evrak_no=evrak_no, **_evrak_formu_kwargs(personel, haklar, toplam_kalan, tip))
    return render(request, "portal/izin_evrak.html", {"form": form, "personel": personel, "haklar": haklar, "toplam_kalan": toplam_kalan, "tip": tip, "bilgi": bilgi, "imzacilar": imzacilar, "duzenleme": False})

@login_required
@veritabani_yogunlugunu_yakala
@transaction.atomic
def izin_talebi_duzenle(request, talep_id):
    talep = get_object_or_404(IzinTalebi.objects.select_for_update().select_related("izin_turu", "kullanici"), id=talep_id)
    if _evrak_surumu_eski(request, talep):
        return redirect("izin_talebi_duzenle", talep_id=talep.pk)
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
    if not _kendi_personeli_mi(request.user, personel):
        messages.error(request, "Bu evrak başka personele ait olduğu için düzenlenemez.")
        return redirect("dashboard")
    haklar, toplam_kalan = izin_ozeti(personel)
    imzacilar = imzacilari_bul(personel)
    if not haklar or toplam_kalan <= 0:
        messages.warning(request, "İzin hakkınız bulunmadı. Yanlışlık olduğunu düşünüyorsanız lütfen İK birimi ile iletişime geçiniz.")
    if request.method == "POST" and request.POST.get("islem") == "sil":
        if not _talep_silinebilir_mi(request.user, talep):
            messages.error(request, 'Bu izin evrakını silme yetkiniz yok.')
            return redirect("dashboard")
        evrak_no = talep.referans_no
        ok, hata = _onayli_talep_silme_dogrula(request, talep)
        if not ok:
            messages.error(request, hata)
            return redirect(_guvenli_donus(request))
        _denetim_kaydi_ekle(talep, request.user, "SIL", talep.durum, "Evrak sahibi/yetkili tarafindan silindi.")
        talep.delete()
        messages.success(request, f'İzin evrakı silindi. Evrak No: {evrak_no}')
        return redirect("dashboard")
    if request.method == "POST":
        islem = request.POST.get("islem", "onaya_sun")
        taslak = islem == "taslak_kaydet"
        tip = PERSONEL_TIPI_SLUG.get(talep.personel_tipi, "memur")
        form = IzinEvrakForm(
            request.POST, instance=talep, taslak=taslak, next_evrak_no=talep.evrak_no,
            **_evrak_formu_kwargs(personel, haklar, toplam_kalan, tip),
        )
        if form.is_valid():
            kaydedilen_talep, hata = _talebi_kaydet(
                talep, form, personel=personel, kullanici=request.user,
                evrak_no=talep.evrak_no, onaya_sun=not taslak, personel_tipi=talep.personel_tipi,
            )
            if hata:
                form.add_error(None, hata)
                messages.error(request, hata)
            elif taslak:
                talep = kaydedilen_talep
                messages.success(request, "Evrak bilgileri taslak olarak güncellendi.")
                return redirect("dashboard")
            else:
                talep = kaydedilen_talep
                messages.success(request, f"Evrak güncellendi ve ilk parafınız için size yönlendirildi. Takip No: {talep.takip_no}")
                return redirect("evrak_paraf", talep_id=talep.id)
        else:
            messages.error(request, "İzin evrakında hatalı veya eksik alanlar var. Lütfen kırmızı uyarıları kontrol edin.")
    else:
        tip = PERSONEL_TIPI_SLUG.get(talep.personel_tipi, "memur")
        form = IzinEvrakForm(
            instance=talep, next_evrak_no=talep.evrak_no,
            **_evrak_formu_kwargs(personel, haklar, toplam_kalan, tip),
        )
    tip = {"MEMUR": "memur", "ISCI": "isci", "SIRKET": "sirket"}.get(talep.personel_tipi, "memur")
    bilgi = _evrak_tipi_or_404(tip)
    return render(request, "portal/izin_evrak.html", {
        "form": form, "personel": personel, "haklar": haklar,
        "toplam_kalan": toplam_kalan, "tip": tip, "bilgi": bilgi,
        "imzacilar": imzacilar, "duzenleme": True, "talep": talep,
        "silinebilir": _talep_silinebilir_mi(request.user, talep),
        "ekler": talep.ekler.select_related("ekleyen").all(),
        "ek_duzenlenebilir": _talep_ek_duzenlenebilir_mi(request.user, talep),
        "goreve_baslama_tarihi": goreve_baslama_tarihi(talep),
    })


@login_required
@veritabani_yogunlugunu_yakala
@transaction.atomic
def izin_talebi_islak_cikti(request, talep_id):
    talep = get_object_or_404(IzinTalebi.objects.select_for_update(of=("self",)).select_related("izin_turu", "kullanici"), id=talep_id)
    if _evrak_surumu_eski(request, talep):
        return redirect("izin_talebi_duzenle", talep_id=talep.pk)
    if talep.kullanici_id != request.user.id:
        messages.error(request, "Bu evrağın çıktısını alma yetkiniz yok.")
        return redirect("dashboard")
    if talep.durum not in {"TASLAK", "IADE_EDILDI", "REDDEDILDI"}:
        messages.warning(request, "Onaya sunulmuş evrak için ıslak imza çıktısı alınamaz.")
        return redirect("evrak_paraf", talep_id=talep.id)
    if request.method != "POST":
        return redirect("izin_talebi_duzenle", talep_id=talep.id)

    personel = talep_personeli(talep)
    if not personel:
        messages.error(request, "Evrağa bağlı personel kaydı bulunamadı.")
        return redirect("dashboard")
    if not _kendi_personeli_mi(request.user, personel):
        messages.error(request, "Bu evrak başka personele ait olduğu için çıktı alınamaz.")
        return redirect("dashboard")

    haklar, toplam_kalan = izin_ozeti(personel)
    imzacilar = imzacilari_bul(personel)
    tip = PERSONEL_TIPI_SLUG.get(talep.personel_tipi, "memur")
    bilgi = _evrak_tipi_or_404(tip)
    form = IzinEvrakForm(
        request.POST,
        instance=talep,
        taslak=False,
        next_evrak_no=talep.evrak_no,
        **_evrak_formu_kwargs(personel, haklar, toplam_kalan, tip),
    )
    if not form.is_valid():
        messages.error(request, "Çıktı almadan önce izin evrakındaki zorunlu alanları tamamlayın.")
        return render(request, "portal/izin_evrak.html", {
            "form": form, "personel": personel, "haklar": haklar,
            "toplam_kalan": toplam_kalan, "tip": tip, "bilgi": bilgi,
            "imzacilar": imzacilar, "duzenleme": True, "talep": talep,
            "silinebilir": _talep_silinebilir_mi(request.user, talep),
            "ekler": talep.ekler.select_related("ekleyen").all(),
            "ek_duzenlenebilir": _talep_ek_duzenlenebilir_mi(request.user, talep),
            "goreve_baslama_tarihi": goreve_baslama_tarihi(talep),
        })

    kaydedilen_talep, hata = _talebi_kaydet(
        talep,
        form,
        personel=personel,
        kullanici=request.user,
        evrak_no=talep.evrak_no,
        onaya_sun=False,
        personel_tipi=talep.personel_tipi,
    )
    if hata:
        messages.error(request, hata)
        return redirect("izin_talebi_duzenle", talep_id=talep.id)

    kaydedilen_talep.kontrol_eden = kaydedilen_talep.kontrol_eden or kontrol_edeni_bul(kaydedilen_talep.personel_tipi)
    content = izin_pdf_olustur(
        kaydedilen_talep,
        personel,
        imzacilar,
        goreve_baslama_tarihi(kaydedilen_talep),
        islak_imza=True,
    )
    guvenli_no = re.sub(r"[^A-Za-z0-9_-]", "-", kaydedilen_talep.referans_no)
    response = HttpResponse(content, content_type="application/pdf")
    response["Content-Disposition"] = f'inline; filename="islak-imza-izin-evraki-{guvenli_no}.pdf"'
    return response


@login_required
def yeni_talep(request):
    return redirect("izin_menu")


@login_required
def izin_bildirimini_kapat(request, talep_id):
    if request.method != "POST":
        return redirect("dashboard")
    qs = _kapatilabilir_izin_bildirimleri(request.user).filter(id=talep_id)
    talep = get_object_or_404(qs)
    IzinBildirimKapatma.objects.get_or_create(kullanici=request.user, talep=talep)
    messages.success(request, "İzin evrakı bildirimi kapatıldı.")
    return redirect("dashboard")


@login_required
def izin_bildirimlerini_kapat(request):
    if request.method != "POST":
        return redirect("dashboard")
    qs = _kapatilabilir_izin_bildirimleri(request.user)
    talep_ids = qs.values_list("id", flat=True)
    IzinBildirimKapatma.objects.bulk_create(
        [IzinBildirimKapatma(kullanici=request.user, talep_id=talep_id) for talep_id in talep_ids],
        ignore_conflicts=True,
    )
    messages.success(request, "İzin evrakı bildirimleri kapatıldı.")
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
        "personel_veri_yetkili": _personel_veri_yonetimi_yetkisi(request.user),
        "birim_yetkili": _birim_ve_vekalet_yetkisi(request.user),
        "gorev_birimi_yetkili": _gorev_birimi_yonetimi_yetkisi(request.user),
        "vekalet_yetkili": _vekalet_yonetimi_yetkisi(request.user),
        "evrak_yetkili": evrak_yetkili,
    })


@user_passes_test(_personel_veri_yonetimi_yetkisi, login_url="dashboard")
def personel_veri_yonetimi(request):
    """özlükten seıilerek aktarılan portal personel kartlarını yönetir."""
    tip = (request.GET.get("tip") or request.POST.get("tip") or "MEMUR").upper()
    if tip not in PERSONEL_TIPI_MODELLERI:
        tip = "MEMUR"
    kaynak_model = _kaynak_personel_modeli(tip)

    if request.method == "POST":
        islem = request.POST.get("islem")
        if islem in {"tumunu_aktar", "secilenleri_guncelle", "tum_tipleri_aktar"}:
            if islem == "tumunu_aktar":
                kaynaklar = kaynak_model.objects.select_related("unvan").all()
            else:
                kaynaklar = []
                if islem == "tum_tipleri_aktar":
                    for kaynak_tip, model in PERSONEL_TIPI_MODELLERI.items():
                        for kaynak in model.objects.select_related("unvan").all():
                            kaynaklar.append((kaynak_tip, kaynak))
                else:
                    secilenler = []
                    for value in request.POST.getlist("source_ids"):
                        if ":" in value:
                            kaynak_tip, source_id = value.split(":", 1)
                            kaynak_tip = kaynak_tip.upper()
                        else:
                            kaynak_tip, source_id = tip, value
                        if kaynak_tip in PERSONEL_TIPI_MODELLERI and source_id.isdigit():
                            secilenler.append((kaynak_tip, int(source_id)))
                    for kaynak_tip, source_id in secilenler:
                        kaynak = PERSONEL_TIPI_MODELLERI[kaynak_tip].objects.select_related("unvan").filter(id=source_id).first()
                        if kaynak:
                            kaynaklar.append((kaynak_tip, kaynak))
            aktarilan = 0
            for item in kaynaklar:
                kaynak_tip, kaynak = item if isinstance(item, tuple) else (tip, item)
                ozlukten_portala_aktar(kaynak, kaynak_tip)
                aktarilan += 1
            if aktarilan:
                messages.success(request, f"{aktarilan} personel özlükten izin portalına aktarıldı/güncellendi. Kullanıcı adları oluşturuldu, ilk şifreler 12345678.")
            else:
                messages.warning(request, "Aktarmak için en az bir personel seçin.")
            return redirect(f"{request.path}?tip={tip}")

        if islem == "yerel_kaydet":
            source_id = request.POST.get("source_personel_id", "")
            if not source_id.isdigit():
                messages.error(request, "Gecersiz kaynak personel numarasi.")
                return redirect(f"{request.path}?tip={tip}")
            kaynak = kaynak_model.objects.filter(id=source_id).first()
            if not kaynak:
                messages.error(request, "Özlükte kaynak personel kaydı bulunamadı.")
                return redirect(f"{request.path}?tip={tip}")
            portal_personeli, _ = ozlukten_portala_aktar(kaynak, tip)
            for alan in ("sicil_no", "ad", "soyad", "tc_no", "unvan", "cep_tel", "gorev_birimi", "alt_birimi", "gorevi"):
                setattr(portal_personeli, alan, (request.POST.get(alan) or "").strip())
            portal_personeli.save()
            _portal_kullanici_olustur(portal_personeli)
            messages.success(request, f"{portal_personeli.ad_soyad} için yerel izin portalı bilgileri kaydedildi.")
            return redirect(f"{request.path}?tip={tip}")

    q = (request.GET.get("q") or "").strip()
    secili_source = (request.GET.get("personel_source") or "").strip()
    kaynak_qs = kaynak_model.objects.select_related("unvan").all()
    kaynak_satirlari = []
    kaynak_tipleri = [tip]
    if ":" in secili_source:
        kaynak_tip, source_id = secili_source.split(":", 1)
        kaynak_tip = kaynak_tip.upper()
        if kaynak_tip in PERSONEL_TIPI_MODELLERI and source_id.isdigit():
            kaynak_tipleri = [kaynak_tip]
            kaynak = PERSONEL_TIPI_MODELLERI[kaynak_tip].objects.select_related("unvan").filter(id=int(source_id)).first()
            if kaynak:
                kaynak_satirlari.append((kaynak_tip, kaynak))
    elif q:
        kaynak_tipleri = list(PERSONEL_TIPI_MODELLERI.keys())
        for kaynak_tip, model in PERSONEL_TIPI_MODELLERI.items():
            for kaynak in model.objects.select_related("unvan").all():
                if _kaynak_personel_arama_eslesir(kaynak, q, kaynak_tip):
                    kaynak_satirlari.append((kaynak_tip, kaynak))
    else:
        kaynak_satirlari = [(tip, kaynak) for kaynak in kaynak_qs.order_by("sicil_no")]
    kaynak_satirlari.sort(key=lambda item: (_sicil_siralama_degeri(getattr(item[1], "sicil_no", "")), _normalize(getattr(item[1], "ad_soyad", "")), item[0]))
    kaynaklar = [kaynak for _, kaynak in kaynak_satirlari]
    yerel_map = {
        (kayit.personel_tipi, kayit.source_personel_id): kayit
        for kayit in PortalPersonel.objects.filter(
            personel_tipi__in=kaynak_tipleri,
            source_personel_id__in=[k.id for k in kaynaklar],
        )
    }
    satirlar = [
        {
            "kaynak": kaynak,
            "tip": kaynak_tip,
            "tip_etiket": PERSONEL_TIPI_ETIKET.get(kaynak_tip, kaynak_tip),
            "portal": yerel_map.get((kaynak_tip, kaynak.id)),
            "sicil_no": (yerel_map.get((kaynak_tip, kaynak.id)).sicil_no if yerel_map.get((kaynak_tip, kaynak.id)) else kaynak.sicil_no),
            "ad_soyad": (yerel_map.get((kaynak_tip, kaynak.id)).ad_soyad if yerel_map.get((kaynak_tip, kaynak.id)) else kaynak.ad_soyad),
            "unvan": (yerel_map.get((kaynak_tip, kaynak.id)).unvani if yerel_map.get((kaynak_tip, kaynak.id)) else kaynak.unvani),
            "birim": " / ".join(filter(None, [
                (yerel_map.get((kaynak_tip, kaynak.id)).gorev_birimi if yerel_map.get((kaynak_tip, kaynak.id)) else kaynak.gorev_birimi),
                (yerel_map.get((kaynak_tip, kaynak.id)).alt_birimi if yerel_map.get((kaynak_tip, kaynak.id)) else kaynak.alt_birimi),
            ])),
        }
        for kaynak_tip, kaynak in kaynak_satirlari
    ]

    duzenlenecek = None
    edit_source_id = request.GET.get("duzenle")
    if edit_source_id and edit_source_id.isdigit():
        kaynak = kaynak_model.objects.select_related("unvan").filter(id=int(edit_source_id)).first()
        yerel = yerel_map.get((tip, int(edit_source_id))) or PortalPersonel.objects.filter(personel_tipi=tip, source_personel_id=edit_source_id).first()
        if kaynak:
            duzenlenecek = {
                "tip": tip,
                "tip_etiket": PERSONEL_TIPI_ETIKET.get(tip, "Memur"),
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
        "tip": tip,
        "tip_slug": PERSONEL_TIPI_SLUG.get(tip, "memur"),
        "tip_etiket": PERSONEL_TIPI_ETIKET.get(tip, "Memur"),
        "tipler": [
            {"kod": kod, "etiket": PERSONEL_TIPI_ETIKET[kod], "kaynak": model.objects.count(), "portal": PortalPersonel.objects.filter(personel_tipi=kod).count()}
            for kod, model in PERSONEL_TIPI_MODELLERI.items()
        ],
        "toplam_kaynak": kaynak_model.objects.count(),
        "toplam_portal": PortalPersonel.objects.filter(personel_tipi=tip).count(),
        "genel_toplam_kaynak": sum(model.objects.count() for model in PERSONEL_TIPI_MODELLERI.values()),
        "genel_toplam_portal": PortalPersonel.objects.count(),
        "q": q,
        "secili_source": secili_source,
        "duzenlenecek": duzenlenecek,
        "son_aktarim": request.user.personel_aktarimlari.order_by("-olusturma_zamani").first(),
    })


@user_passes_test(_evrak_yonetim_yetkisi, login_url="dashboard")
@veritabani_yogunlugunu_yakala
@transaction.atomic
def izin_talepleri_yonetim(request):
    if request.method == "POST":
        talep = _post_talebi(request)
        islem = request.POST.get("islem")
        not_text = (request.POST.get("not") or "").strip()
        if islem == "geri_cek":
            if not _teklif_parafi_verildi_mi(talep):
                messages.error(request, "İlk paraf verilmeden evrak kontrolcüye gönderilemez.")
                return redirect("izin_talepleri_yonetim")
            if not _tip_yetkisi_var(request.user, talep.personel_tipi, "evrak_geri_cekme"):
                messages.error(request, "Bu personel tipi için evrağı geri çekme yetkiniz yok.")
                return redirect("izin_talepleri_yonetim")
            if talep.durum in ["IMZA_SURECINDE", "PARAF_BEKLIYOR"]:
                talep.durum = "PARAF_BEKLIYOR"
                talep.kimin_uzerinde = talep.kontrol_eden
                talep.imza_sirasi = 1
                ek = f"\n[{timezone.localtime().strftime('%d.%m.%Y %H:%M')}] Evrak yönetim tarafından kontrol aşamasına geri çekildi."
                if not_text:
                    ek += f" Not: {not_text}"
                talep.aciklama = (talep.aciklama or "") + ek
                talep.save(update_fields=["durum", "kimin_uzerinde", "imza_sirasi", "aciklama", "guncelleme_zamani"])
                _islem_gecmisi_ekle(talep, request.user, "GERI_CEK", "Evrak Yönetimi", not_text)
                messages.success(request, "Evrak kontrol aşamasına geri çekildi.")
            else:
                messages.error(request, "Bu durumdaki evrak geri çekilemez.")
        elif islem == "sahibine_iade":
            if talep.durum not in {"PARAF_BEKLIYOR", "IMZA_SURECINDE"}:
                messages.error(request, "Onaylanmış evrak sahibine iade edilemez.")
                return redirect("izin_talepleri_yonetim")
            if not _tip_yetkisi_var(request.user, talep.personel_tipi, "sahibine_iade"):
                messages.error(request, "Bu personel tipi için evrağı sahibine iade etme yetkiniz yok.")
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
        elif islem == "izin_sil":
            if not _talep_silinebilir_mi(request.user, talep):
                messages.error(request, "Bu personel tipi için izin evrakı silme yetkiniz yok.")
                return redirect("izin_talepleri_yonetim")
            evrak_no = talep.referans_no
            ok, hata = _onayli_talep_silme_dogrula(request, talep)
            if not ok:
                messages.error(request, hata)
                return redirect(_guvenli_donus(request))
            _denetim_kaydi_ekle(talep, request.user, "SIL", talep.durum, "Evrak yönetiminden silindi.")
            talep.delete()
            messages.success(request, f"{evrak_no} numaralı izin evrakı silindi.")
        return redirect("izin_talepleri_yonetim")
    q = (request.GET.get("q") or "").strip()
    durum = (request.GET.get("durum") or "").strip()
    qs = IzinTalebi.objects.select_related("izin_turu", "kullanici", "kontrol_eden", "kimin_uzerinde").filter(_talep_arama_filtresi(request.user)).distinct().order_by("-olusturma_zamani")
    if q:
        qs = qs.filter(talep_numarasi_filtresi(q) | Q(kullanici__username__icontains=q) | Q(kullanici__first_name__icontains=q) | Q(kullanici__last_name__icontains=q))
    if durum:
        qs = qs.filter(durum=durum)
    rows = []
    for talep in qs[:200]:
        rows.append({
            "talep": talep,
            "personel": talep_personeli(talep),
            "can_geri_cek": talep.durum != "ONAYLANDI" and _tip_yetkisi_var(request.user, talep.personel_tipi, "evrak_geri_cekme"),
            "can_iade": talep.durum != "ONAYLANDI" and _tip_yetkisi_var(request.user, talep.personel_tipi, "sahibine_iade"),
            "can_delete": _talep_silinebilir_mi(request.user, talep),
        })
    return render(request, "portal/izin_talepleri_yonetim.html", {
        "rows": rows, "q": q, "durum": durum, "durumlar": IzinTalebi.DURUM_CHOICES,
    })

@user_passes_test(_tam_yonetim_yetkisi, login_url="dashboard")
@veritabani_yogunlugunu_yakala
@transaction.atomic
def kullanici_yetkileri(request):
    from .yetki_yonetimi import yetki_yonetimi
    return yetki_yonetimi(request)


@user_passes_test(_tam_yonetim_yetkisi, login_url="dashboard")
def kullanici_yetkisi_detay(request):
    user_id = request.GET.get("kullanici")
    user = User.objects.filter(id=user_id, is_active=True).first()
    if not user:
        return JsonResponse({"error": "Kullanıcı bulunamadı."}, status=404)
    yetki = YonetimYetkisi.objects.filter(kullanici=user).first()
    return JsonResponse({
        "kullanici": {"id": user.id, "label": _user_label(user)},
        "surum": yetki.guncelleme_zamani.isoformat() if yetki else "yok",
        "aciklama": yetki.aciklama if yetki else "",
        "yetkiler": _yetki_form_degerleri(yetki),
    })


@user_passes_test(_birim_ve_vekalet_yetkisi, login_url="dashboard")
@veritabani_yogunlugunu_yakala
@transaction.atomic
def gorev_birimi_yonetimi(request):
    bugun = timezone.localdate()
    _suresi_dolan_vekaletleri_pasife_al(bugun)
    gorev_birimi_yetkili = _gorev_birimi_yonetimi_yetkisi(request.user)
    vekalet_yetkili = _vekalet_yonetimi_yetkisi(request.user)
    if request.method == "POST":
        islem = request.POST.get("islem")
        if islem == "birim_senkronize":
            if not gorev_birimi_yetkili:
                messages.error(request, "Görev birimi yönetimi için yetkiniz yok.")
                return redirect("gorev_birimi_yonetimi")
            gorev_birimlerini_senkronla()
            messages.success(request, "Görev birimleri portal personel kartlarından senkronize edildi.")
            return redirect("gorev_birimi_yonetimi")
        if islem == "imzaci_kaydet":
            if not gorev_birimi_yetkili:
                messages.error(request, "Görev birimi imzacılarını düzenleme yetkiniz yok.")
                return redirect("gorev_birimi_yonetimi")
            for kayit in GorevBirimiImzaci.objects.filter(aktif_mi=True):
                if kayit.alt_birim_adi:
                    alan_adi = f"sube_{kayit.id}"
                else:
                    alan_adi = f"daire_{kayit.id}"
                if alan_adi not in request.POST:
                    continue
                kullanici_id = request.POST.get(alan_adi) or None
                if kullanici_id and (not kullanici_id.isdigit() or not User.objects.filter(pk=kullanici_id, is_active=True).exists()):
                    messages.error(request, "İmzacı aktif bir kullanıcı olmalıdır. Değişiklikler kaydedilmedi.")
                    transaction.set_rollback(True)
                    return redirect("gorev_birimi_yonetimi")
                kayit.amir_kullanici_id = kullanici_id
                if kayit.alt_birim_adi:
                    kayit.daire_baskani_id = None
                    kayit.sube_muduru_id = kullanici_id
                else:
                    kayit.daire_baskani_id = kullanici_id
                    kayit.sube_muduru_id = None
                kayit.save(update_fields=["amir_kullanici", "daire_baskani", "sube_muduru", "guncelleme_zamani"])
            messages.success(request, "G\u00f6rev birimi imzac\u0131lar\u0131 kaydedildi.")
            return redirect("gorev_birimi_yonetimi")
        if islem == "vekalet_ekle":
            if not vekalet_yetkili:
                messages.error(request, "Vekalet düzenleme yetkiniz yok.")
                return redirect("gorev_birimi_yonetimi")
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
                if not asil_id.isdigit() or not vekil_id.isdigit() or User.objects.filter(pk__in=[asil_id, vekil_id], is_active=True).count() != 2:
                    messages.error(request, "Asıl ve vekil aktif kullanıcı olmalıdır.")
                    return redirect("gorev_birimi_yonetimi")
                # Lock the principal so overlapping delegation assignments
                # cannot choose different signers for the same date.
                User.objects.select_for_update().get(pk=asil_id)
                if ImzaciVekaleti.objects.filter(asil_kullanici_id=asil_id, aktif_mi=True, baslangic_tarihi__lte=bitis_tarihi, bitis_tarihi__gte=baslangic_tarihi).exists():
                    messages.error(request, "Asıl kullanıcı için bu tarih aralığında zaten aktif bir vekalet var.")
                    return redirect("gorev_birimi_yonetimi")
                ImzaciVekaleti.objects.create(
                    asil_kullanici_id=asil_id,
                    vekil_kullanici_id=vekil_id,
                    baslangic_tarihi=baslangic_tarihi,
                    bitis_tarihi=bitis_tarihi,
                    aciklama=request.POST.get("aciklama", ""),
                    aktif_mi=True,
                )
                _bekleyen_vekalet_evraklarini_guncelle(kaydeden=request.user)
                messages.success(request, "Vekalet kaydı eklendi.")
            else:
                messages.error(request, "Vekalet i\u00e7in asil, vekil ve tarih alanlar\u0131 zorunludur.")
            return redirect("gorev_birimi_yonetimi")
        if islem == "vekalet_pasif":
            if not vekalet_yetkili:
                messages.error(request, "Vekalet düzenleme yetkiniz yok.")
                return redirect("gorev_birimi_yonetimi")
            vekalet = get_object_or_404(ImzaciVekaleti, id=request.POST.get("vekalet_id"))
            vekalet.aktif_mi = False
            vekalet.save(update_fields=["aktif_mi"])
            tasinan = _bekleyen_vekalet_evraklarini_guncelle(
                kullanicilar=[vekalet.asil_kullanici_id, vekalet.vekil_kullanici_id],
                kaydeden=request.user,
            )
            mesaj = "Vekalet pasife al\u0131nd\u0131."
            if tasinan:
                mesaj += f" Bekleyen {tasinan} evrak aktif imzacıya geri aktarıldı."
            messages.success(request, mesaj)
            return redirect("gorev_birimi_yonetimi")

    kullanicilar = User.objects.filter(is_active=True).order_by("first_name", "last_name", "username")
    kayitlar = list(
        GorevBirimiImzaci.objects
        .filter(aktif_mi=True)
        .select_related("amir_kullanici", "daire_baskani", "sube_muduru")
        .order_by("birim_adi", "alt_birim_adi")
    )
    for kayit in kayitlar:
        kayit.ana_imzaci_etiketi = "Genel Müdür" if "genel mudur" in _normalize(kayit.birim_adi) else "Daire Başkanı"
    org_roots = _gorev_birimi_agaci(kayitlar)
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
        "org_roots": org_roots,
        "vekaletler": vekaletler,
        "vekalet_q": vekalet_q,
        "ayar": PortalAyar.aktif(),
        "gorev_birimi_yetkili": gorev_birimi_yetkili,
        "vekalet_yetkili": vekalet_yetkili,
    })


@user_passes_test(_tam_yonetim_yetkisi, login_url="dashboard")
def kontrol_eden_yonetimi(request):
    messages.info(request, "Kontrol eden seçimi artık Kullanıcı Yetkileri ekranından yapılır.")
    return redirect("kullanici_yetkileri")


@login_required
def kullanici_ara(request):
    if not (_tam_yonetim_yetkisi(request.user) or _birim_ve_vekalet_yetkisi(request.user) or _genis_izin_arama_yetkisi(request.user)):
        return JsonResponse({"error": "Kullanıcı arama yetkiniz yok."}, status=403)
    q = (request.GET.get("q") or "").strip()
    page = request.GET.get("page") or 1
    qs = User.objects.filter(is_active=True).order_by("first_name", "last_name", "username")
    if q:
        q_norm = _normalize(q)
        q_compact = q_norm.replace(" ", "")
        tokens = [parca for parca in q_norm.split() if parca]
        secilenler = []
        for user in qs[:2000]:
            full = user.get_full_name().strip()
            blob = _normalize(" ".join([
                user.username or "",
                user.first_name or "",
                user.last_name or "",
                full,
                full.replace(" ", ""),
            ]))
            blob_compact = blob.replace(" ", "")
            if q_norm in blob or q_compact in blob_compact or all(token in blob for token in tokens):
                secilenler.append(user.id)
        qs = User.objects.filter(id__in=secilenler).order_by("first_name", "last_name", "username")
    paginator = Paginator(qs.order_by("first_name", "last_name", "username"), 10)
    page_obj = paginator.get_page(page)
    data = []
    for user in page_obj.object_list:
        full = user.get_full_name().strip()
        data.append({
            "id": user.id,
            "username": user.username,
            "full_name": full,
            "label": f"{full} ({user.username})" if full else user.username,
        })
    return JsonResponse({
        "results": data,
        "pagination": {
            "page": page_obj.number,
            "pages": paginator.num_pages,
            "total": paginator.count,
            "has_previous": page_obj.has_previous(),
            "has_next": page_obj.has_next(),
        },
    })


@login_required
def birim_ara(request):
    q = (request.GET.get("q") or "").strip()
    mode = (request.GET.get("mode") or "ana").strip()
    birim = (request.GET.get("birim") or "").strip()
    page = request.GET.get("page") or 1
    data = []

    if mode == "alt":
        if not birim:
            return JsonResponse({"results": data, "pagination": {"page": 1, "pages": 1, "total": 0, "has_previous": False, "has_next": False}})
        qs = GorevBirimiImzaci.objects.filter(aktif_mi=True, birim_adi=birim).exclude(alt_birim_adi="")
        if q:
            qs = qs.filter(alt_birim_adi__icontains=q)
        paginator = Paginator(qs.order_by("alt_birim_adi"), 10)
        page_obj = paginator.get_page(page)
        for kayit in page_obj.object_list:
            alt = (kayit.alt_birim_adi or "").strip()
            if not alt:
                continue
            data.append({
                "birim": kayit.birim_adi,
                "alt_birim": alt,
                "label": alt,
            })
        return JsonResponse({
            "results": data,
            "pagination": {
                "page": page_obj.number,
                "pages": paginator.num_pages,
                "total": paginator.count,
                "has_previous": page_obj.has_previous(),
                "has_next": page_obj.has_next(),
            },
        })

    qs = GorevBirimiImzaci.objects.filter(aktif_mi=True, alt_birim_adi="")
    if q:
        qs = qs.filter(birim_adi__icontains=q)
    paginator = Paginator(qs.order_by("birim_adi"), 10)
    page_obj = paginator.get_page(page)
    for kayit in page_obj.object_list:
        ana_birim = (kayit.birim_adi or "").strip()
        if not ana_birim:
            continue
        has_children = GorevBirimiImzaci.objects.filter(
            aktif_mi=True,
            birim_adi=ana_birim,
        ).exclude(alt_birim_adi="").exists()
        data.append({
            "birim": ana_birim,
            "alt_birim": "",
            "label": ana_birim,
            "has_children": has_children,
        })
    return JsonResponse({
        "results": data,
        "pagination": {
            "page": page_obj.number,
            "pages": paginator.num_pages,
            "total": paginator.count,
            "has_previous": page_obj.has_previous(),
            "has_next": page_obj.has_next(),
        },
    })


def _personel_secim_kapsami(user):
    """Arama bıyötecinde gösterilebilecek personelleri yetkiye göre sınırlar."""
    model = _personel_veri_modeli()
    if (
        _izin_yetkisi_var(user, "tum_izinleri_goruntuleme")
        or _tam_yonetim_yetkisi(user)
    ):
        return model.objects.all()
    kapsam = Q(pk__in=[])
    for tip in PERSONEL_TIPI_MODELLERI:
        if (_tip_yetkisi_var(user, tip, "izin_arama")
                or _tip_yetkisi_var(user, tip, "izin_goruntuleme")
                or tip in _kontrol_ettigi_tipler(user)):
            if model is PortalPersonel:
                kapsam |= Q(personel_tipi=tip)
            elif tip == "MEMUR":
                return model.objects.all()
    kendi_personeli = kullanicinin_personeli(user)
    if kendi_personeli:
        kapsam |= Q(personel_tipi=_personel_tip_kodu(kendi_personeli), source_personel_id=_personel_kaynak_id(kendi_personeli)) if model is PortalPersonel else Q(pk=kendi_personeli.id)
    yonetilenler = _yonettigi_personeller(user)
    if yonetilenler.exists():
        kapsam |= Q(pk__in=yonetilenler.values("id"))
    gecmis_talepler = IzinTalebi.objects.filter(_talep_arama_filtresi(user) | Q(kullanici=user))
    gecmis_kapsami = Q(pk__in=[])
    if model is PortalPersonel:
        for tip in PERSONEL_TIPI_MODELLERI:
            gecmis_kapsami |= Q(personel_tipi=tip, source_personel_id__in=gecmis_talepler.filter(personel_tipi=tip).values("personel_id"))
    else:
        gecmis_kapsami = Q(pk__in=gecmis_talepler.filter(personel_tipi="MEMUR").values("personel_id"))
    return model.objects.filter(kapsam | gecmis_kapsami).distinct()

@user_passes_test(_personel_veri_yonetimi_yetkisi, login_url="dashboard")
def personel_veri_personel_ara(request):
    q = (request.GET.get("q") or "").strip()
    page = request.GET.get("page") or 1
    rows = []
    for kaynak_tip, model in PERSONEL_TIPI_MODELLERI.items():
        for kaynak in model.objects.select_related("unvan").all():
            if _kaynak_personel_arama_eslesir(kaynak, q, kaynak_tip):
                rows.append((kaynak_tip, kaynak))
    rows.sort(key=lambda item: (_sicil_siralama_degeri(getattr(item[1], "sicil_no", "")), _normalize(getattr(item[1], "ad_soyad", "")), item[0]))
    paginator = Paginator(rows, 10)
    page_obj = paginator.get_page(page)
    data = [
        {
            "id": f"{kaynak_tip}:{kaynak.id}",
            "source_id": kaynak.id,
            "tip": kaynak_tip,
            "tip_etiket": PERSONEL_TIPI_ETIKET.get(kaynak_tip, kaynak_tip),
            "label": f"{kaynak.sicil_no} - {kaynak.ad_soyad}",
            "username": _kaynak_personel_kullanici_adi_adayi(kaynak),
            "birim": " / ".join(filter(None, [kaynak.gorev_birimi, kaynak.alt_birimi])),
        }
        for kaynak_tip, kaynak in page_obj.object_list
    ]
    return JsonResponse({
        "results": data,
        "pagination": {
            "page": page_obj.number,
            "pages": paginator.num_pages,
            "total": paginator.count,
            "has_previous": page_obj.has_previous(),
            "has_next": page_obj.has_next(),
        },
    })


@login_required
def personel_ara(request):
    q = (request.GET.get("q") or "").strip()
    page = request.GET.get("page") or 1
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
    paginator = Paginator(qs.order_by("sicil_no", "ad", "soyad"), 10)
    page_obj = paginator.get_page(page)
    data = [
        {
            "id": personel.id,
            "label": f"{personel.sicil_no} - {personel.ad_soyad}",
            "birim": " / ".join(filter(None, [personel.gorev_birimi, personel.alt_birimi])),
        }
        for personel in page_obj.object_list
    ]
    return JsonResponse({
        "results": data,
        "pagination": {
            "page": page_obj.number,
            "pages": paginator.num_pages,
            "total": paginator.count,
            "has_previous": page_obj.has_previous(),
            "has_next": page_obj.has_next(),
        },
    })


@login_required
@veritabani_yogunlugunu_yakala
@transaction.atomic
def evrak_paraf(request, talep_id):
    talep = get_object_or_404(IzinTalebi.objects.select_for_update().select_related("izin_turu", "kullanici", "kontrol_eden", "kimin_uzerinde"), id=talep_id)
    _bekleyen_talep_hedefini_guncelle(talep, kaydeden=None)
    if talep.kimin_uzerinde_id:
        talep.kimin_uzerinde = User.objects.filter(pk=talep.kimin_uzerinde_id).first()
    personel = talep_personeli(talep)
    if not _talep_gorme_yetkisi(request.user, talep):
        messages.error(request, "Bu evrakı görüntüleme yetkiniz yok.")
        return redirect("dashboard")
    if _evrak_surumu_eski(request, talep):
        return redirect("evrak_paraf", talep_id=talep.pk)
    if not personel:
        messages.error(request, "Evrağa bağlı personel kaydı bulunamadı.")
        return redirect("dashboard")
    if talep.durum == "TASLAK":
        if talep.kullanici_id == request.user.id:
            return redirect("izin_talebi_duzenle", talep_id=talep.id)
        messages.warning(request, "Taslak evrak önce sahibi tarafından tamamlanmalıdır.")
        return redirect("dashboard")
    imzacilar = talep_aktif_imzacilari(talep)
    if request.method == "POST":
        islem = request.POST.get("islem")
        if islem in {"parafla", "imzala"}:
            if not personel_aktif_mi(personel):
                messages.error(request, "Pasif personelin evrakı ilerletilemez. İK birimi personel durumunu kontrol etmelidir.")
                return redirect("dashboard")
            if not _talep_aktif_gorev_yetkisi(request.user, talep):
                messages.error(request, "Bu evrak şu anda sizin üzerinizde değil.")
                return redirect("dashboard")
            kalan_imzacilar = imzacilar[talep.imza_sirasi:] if talep.durum == "IMZA_SURECINDE" else imzacilar
            if not imzacilar or any(not user or not user.is_active for _, user in kalan_imzacilar):
                messages.error(request, "Geçerli imzacı tanımı bulunamadı. Evrak tamamlanmadı; İK birimi imzacı tanımını kontrol etmelidir.")
                return redirect("evrak_paraf", talep_id=talep.pk)
            onceki_durum = talep.durum
            aktif_rol = "Kontrol Eden"
            if onceki_durum == "IMZA_SURECINDE" and talep.imza_sirasi < len(imzacilar):
                aktif_rol = imzacilar[talep.imza_sirasi][0]
            if talep.durum == "PARAF_BEKLIYOR":
                teklif_parafi_verildi = _teklif_parafi_verildi_mi(talep)
                if not teklif_parafi_verildi:
                    if not talep.kontrol_eden or not _kontrol_yetkisi_var(talep.kontrol_eden, talep.personel_tipi):
                        messages.error(request, "Geçerli kontrolcü tanımı bulunamadı. Evrak ilerletilmedi.")
                        return redirect("evrak_paraf", talep_id=talep.pk)
                    bakiye_hatasi = izin_bakiyesi_onaya_sunma_kontrolu(talep, personel)
                    if bakiye_hatasi:
                        messages.error(request, bakiye_hatasi)
                        return redirect("evrak_paraf", talep_id=talep.id)
                    evraki_sabitle(talep, personel, imzacilar)
                    talep.kimin_uzerinde = talep.kontrol_eden
                    talep.imza_sirasi = 1
                    hedef = talep.kontrol_eden.get_full_name() or talep.kontrol_eden.username
                    messages.success(request, f"Evrak teklif eden tarafından paraflandı ve kontrol için {hedef} kullanıcısına gönderildi.")
                else:
                    if not imzacilar[0][1] or not imzacilar[0][1].is_active:
                        messages.error(request, "Sıradaki imzacı pasif veya tanımsız. Evrak ilerletilmedi.")
                        return redirect("evrak_paraf", talep_id=talep.pk)
                    talep.kontrol_tarihi = timezone.now()
                    talep.imza_sirasi = 0
                    talep.durum = "IMZA_SURECINDE"
                    talep.kimin_uzerinde = imzacilar[0][1]
                    messages.success(request, f"Evrak kontrol parafından geçti ve {imzacilar[0][1].get_full_name() or imzacilar[0][1].username} kullanıcısına gönderildi.")
            elif talep.durum == "IMZA_SURECINDE":
                sonraki_sira = talep.imza_sirasi + 1
                if sonraki_sira < len(imzacilar):
                    talep.imza_sirasi = sonraki_sira
                    talep.kimin_uzerinde = imzacilar[sonraki_sira][1]
                    messages.success(request, f"Evrak paraflandı ve {imzacilar[sonraki_sira][1].get_full_name() or imzacilar[sonraki_sira][1].username} kullanıcısına gönderildi.")
                else:
                    talep.durum = "ONAYLANDI"
                    talep.kimin_uzerinde = None
                    talep.imza_tamamlanma_tarihi = timezone.now()
                    try:
                        with transaction.atomic():
                            resmi_evrak_kaydini_hazirla(talep)
                            izni_ozluge_aktar(talep, personel, request.user)
                    except ValueError as hata:
                        messages.error(request, str(hata))
                        return redirect("dashboard")
                    except IntegrityError:
                        messages.error(request, "Özlükte kullanılan izin kaydı oluşturulamadı. Lütfen İK birimi ile görüşünüz.")
                        return redirect("dashboard")
                    messages.success(request, f"Evrak imzalandı ve izin süreci tamamlandı. Resmi Evrak No: {talep.evrak_no}. Oluşturan personelin görevlerine bilgi düştü.")
            else:
                messages.error(request, "Bu evrak için aktif paraf/imza işlemi bulunmuyor.")
                return redirect("dashboard")
            talep.save(update_fields=["durum", "kontrol_tarihi", "kimin_uzerinde", "imza_sirasi", "imza_tamamlanma_tarihi", "ozluk_aktarildi_mi", "ozluk_aktarim_tarihi", "guncelleme_zamani"])
            if onceki_durum == "PARAF_BEKLIYOR":
                if teklif_parafi_verildi:
                    _islem_gecmisi_ekle(talep, request.user, "KONTROL_PARAFI", "Kontrol Eden")
                else:
                    _islem_gecmisi_ekle(talep, request.user, "IMZA", "Teklif Eden")
            else:
                gecmis_islem = "ONAY" if talep.durum == "ONAYLANDI" else "IMZA"
                _islem_gecmisi_ekle(talep, request.user, gecmis_islem, aktif_rol)
            _denetim_kaydi_ekle(talep, request.user, "PARAF", onceki_durum, aktif_rol, talep.kimin_uzerinde)
            return redirect("dashboard")
        if islem == "reddet":
            if not _talep_aktif_gorev_yetkisi(request.user, talep):
                messages.error(request, "Bu evrak şu anda sizin üzerinizde değil.")
                return redirect("dashboard")
            red_nedeni = (request.POST.get("red_nedeni") or "").strip()
            if not red_nedeni:
                messages.error(request, "Evrakı reddetmek için red nedenini yazmanız zorunludur.")
                return redirect("evrak_paraf", talep_id=talep.id)
            onceki_durum = talep.durum
            with transaction.atomic():
                talep.durum = "REDDEDILDI"
                talep.kimin_uzerinde = None
                talep.red_nedeni = red_nedeni
                talep.reddeden = request.user
                talep.reddedilme_tarihi = timezone.now()
                talep.save(update_fields=["durum", "kimin_uzerinde", "red_nedeni", "reddeden", "reddedilme_tarihi", "guncelleme_zamani"])
                _islem_gecmisi_ekle(talep, request.user, "RED", "Kontrol Eden" if talep.kontrol_eden_id == request.user.id else "İmzacı", red_nedeni)
                _denetim_kaydi_ekle(talep, request.user, "RED", onceki_durum, red_nedeni, talep.kullanici)
                # A fresh rejection must notify the controller even after a prior dismissal.
                if talep.kontrol_eden_id and talep.kontrol_eden_id not in (request.user.id, talep.kullanici_id):
                    IzinBildirimKapatma.objects.filter(talep=talep, kullanici_id=talep.kontrol_eden_id).delete()
            messages.success(request, "Evrak reddedildi ve ilk sahibine geri gönderildi. Ret nedeni evrak sahibine iletildi.")
            return redirect("dashboard")
    talep, personel, imzacilar = belge_verileri(talep, personel, imzacilar)
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
        "ekler": talep.ekler.select_related("ekleyen").all(),
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
    if _talep_aktif_gorev_yetkisi(request.user, talep):
        return redirect("evrak_paraf", talep_id=talep.id)
    talep, personel, imzacilar = belge_verileri(talep, personel, talep_aktif_imzacilari(talep))
    return render(request, "portal/evrak_paraf.html", {"talep": talep, "personel": personel, "imzacilar": imzacilar, "indir": True, "ekler": talep.ekler.select_related("ekleyen").all(), "goreve_baslama_tarihi": goreve_baslama_tarihi(talep)})


@login_required
@veritabani_yogunlugunu_yakala
@transaction.atomic
def izin_eki_ekle(request, talep_id):
    """Taslak evraka destekleyici belge ekler."""
    talep = get_object_or_404(IzinTalebi.objects.select_for_update(), id=talep_id)
    if request.method != "POST":
        return redirect("izin_talebi_duzenle", talep_id=talep.id)
    if not _talep_ek_duzenlenebilir_mi(request.user, talep):
        messages.error(request, "Ekler yalnızca kendi taslak evrağınıza eklenebilir. Onaya sunulduktan sonra ekler sabitlenir.")
        return redirect("izin_talebi_duzenle", talep_id=talep.id)

    dosya = request.FILES.get("dosya")
    if not dosya:
        messages.error(request, "Lütfen eklemek istediğiniz dosyayı seçin.")
        return redirect("izin_talebi_duzenle", talep_id=talep.id)
    dosya_adi = Path(dosya.name).name
    uzanti = Path(dosya_adi).suffix.lower()
    if uzanti not in EK_DOSYA_UZANTILARI:
        messages.error(request, "Bu dosya türü desteklenmiyor. PDF, görsel, Word veya Excel dosyası seçin.")
        return redirect("izin_talebi_duzenle", talep_id=talep.id)
    if dosya.size > EK_DOSYA_AZAMI_BOYUT:
        messages.error(request, "Ek dosya en fazla 10 MB olabilir.")
        return redirect("izin_talebi_duzenle", talep_id=talep.id)

    IzinEvrakEki.objects.create(talep=talep, dosya=dosya, dosya_adi=dosya_adi[:255], ekleyen=request.user)
    _denetim_kaydi_ekle(talep, request.user, "EK_EKLE", talep.durum, dosya_adi)
    messages.success(request, f"{dosya_adi} evrağa eklendi.")
    return redirect("izin_talebi_duzenle", talep_id=talep.id)


@login_required
@veritabani_yogunlugunu_yakala
@transaction.atomic
def izin_eki_sil(request, talep_id, ek_id):
    talep = get_object_or_404(IzinTalebi.objects.select_for_update(), id=talep_id)
    if request.method != "POST":
        return redirect("izin_talebi_duzenle", talep_id=talep.id)
    if not _talep_ek_duzenlenebilir_mi(request.user, talep):
        messages.error(request, "Onaya sunulmuş evrağın ekleri değiştirilemez.")
        return redirect("izin_talebi_duzenle", talep_id=talep.id)
    ek = get_object_or_404(IzinEvrakEki, id=ek_id, talep=talep)
    dosya_adi = ek.dosya_adi
    if ek.dosya:
        storage, name = ek.dosya.storage, ek.dosya.name
        transaction.on_commit(lambda: storage.delete(name))
    ek.delete()
    _denetim_kaydi_ekle(talep, request.user, "EK_SIL", talep.durum, dosya_adi)
    messages.success(request, f"{dosya_adi} evraktan kaldırıldı.")
    return redirect("izin_talebi_duzenle", talep_id=talep.id)


@login_required
def izin_eki_indir(request, talep_id, ek_id):
    talep = get_object_or_404(IzinTalebi, id=talep_id)
    if not _talep_gorme_yetkisi(request.user, talep):
        messages.error(request, "Bu evrağın eklerini görüntüleme yetkiniz yok.")
        return redirect("dashboard")
    ek = get_object_or_404(IzinEvrakEki, id=ek_id, talep=talep)
    if not ek.dosya:
        raise Http404("Ek dosya bulunamadı.")
    content_type = mimetypes.guess_type(ek.dosya_adi)[0] or "application/octet-stream"
    try:
        stream = ek.dosya.open("rb")
    except FileNotFoundError:
        raise Http404("Ek dosya bulunamadı.")
    response = FileResponse(stream, content_type=content_type, as_attachment=request.GET.get("inline") != "1", filename=Path(ek.dosya_adi).name)
    return response


@login_required
def evrak_indir(request, talep_id):
    """Seçilen biıimde dosyayö doğrudan indirir; sayfa açmaz."""
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

    imzacilar = talep_aktif_imzacilari(talep)
    baslama = goreve_baslama_tarihi(talep)
    guvenli_no = re.sub(r"[^A-Za-z0-9_-]", "-", talep.referans_no)
    if dosya_tipi == "pdf":
        content = izin_pdf_olustur(talep, personel, imzacilar, baslama)
        response = HttpResponse(content, content_type="application/pdf")
        # Yazdör döımesi aynı PDF belgesini tarayöcıda açar; HTML yazdörma
        # dözeniyle indirilen belge arasında görınım farkı oluımaz.
        disposition = "inline" if request.GET.get("inline") in {"1", "true", "True"} else "attachment"
        response["Content-Disposition"] = f'{disposition}; filename="izin-evraki-{guvenli_no}.pdf"'
        return response

    content = izin_excel_olustur(talep, personel, imzacilar, baslama)
    response = HttpResponse(
        content,
        content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    response["Content-Disposition"] = f'attachment; filename="izin-evraki-{guvenli_no}.xlsx"'
    return response


@login_required
@veritabani_yogunlugunu_yakala
@transaction.atomic
def izin_ara(request):
    if request.method == "POST":
        islem = request.POST.get("islem")
        if islem == "izin_sil":
            talep = _post_talebi(request)
            if not _talep_silinebilir_mi(request.user, talep):
                messages.error(request, "Bu personel tipi için izin evrakı silme yetkiniz yok.")
                return redirect("izin_ara")
            evrak_no = talep.referans_no
            ok, hata = _onayli_talep_silme_dogrula(request, talep)
            if not ok:
                messages.error(request, hata)
                return redirect(_guvenli_donus(request))
            _denetim_kaydi_ekle(talep, request.user, "SIL", talep.durum, "Yonetim ekranindan silindi.")
            talep.delete()
            messages.success(request, f"{evrak_no} numaralı izin evrakı silindi.")
        elif islem == "sahibine_iade":
            talep = _post_talebi(request)
            tip = getattr(talep, "personel_tipi", "MEMUR")
            if talep.durum not in {"PARAF_BEKLIYOR", "IMZA_SURECINDE"}:
                messages.error(request, "Onaylanan veya kapalı evrak sahibine iade edilemez.")
            elif not _tip_yetkisi_var(request.user, tip, "sahibine_iade"):
                messages.error(request, "Bu personel tipi için sahibine iade yetkiniz yok.")
            else:
                talep.durum = "IADE_EDILDI"
                talep.kimin_uzerinde = talep.kullanici
                talep.imza_sirasi = 0
                talep.save(update_fields=["durum", "kimin_uzerinde", "imza_sirasi", "guncelleme_zamani"])
                _islem_gecmisi_ekle(
                    talep,
                    request.user,
                    "IADE",
                    "Sahibine İade",
                    "Yetkili kullanıcı tarafından evrak sahibine iade edildi.",
                )
                messages.success(request, f"{talep.referans_no} numaralı evrak sahibine iade edildi.")
                _denetim_kaydi_ekle(talep, request.user, "IADE", hedef_kullanici=talep.kullanici)
        elif islem == "evraki_geri_cek":
            talep = _post_talebi(request)
            tip = getattr(talep, "personel_tipi", "MEMUR")
            if talep.durum not in {"PARAF_BEKLIYOR", "IMZA_SURECINDE"}:
                messages.error(request, "Onaylanan veya kapalı evrak geri çekilemez.")
            elif not _tip_yetkisi_var(request.user, tip, "evrak_geri_cekme"):
                messages.error(request, "Bu personel tipi için evrak geri çekme yetkiniz yok.")
            else:
                hedefler = {
                    str(hedef["kullanici"].id): hedef
                    for hedef in _geri_cekme_hedefleri(talep)
                }
                hedef = hedefler.get(str(request.POST.get("hedef_kullanici_id") or ""))
                if not hedef:
                    messages.error(request, "Geri çekme hedefi geçersiz veya bu evrakta daha önce paraf/imza vermemiş.")
                else:
                    talep.durum = hedef["durum"]
                    talep.kimin_uzerinde = hedef["kullanici"]
                    talep.imza_sirasi = hedef["imza_sirasi"]
                    talep.save(update_fields=["durum", "kimin_uzerinde", "imza_sirasi", "guncelleme_zamani"])
                    _islem_gecmisi_ekle(
                        talep,
                        request.user,
                        "GERI_CEK",
                        hedef["rol"],
                        f"Evrak {hedef['ad_soyad']} kullanıcısına geri çekildi.",
                    )
                    _denetim_kaydi_ekle(talep, request.user, "GERI_CEK", aciklama=hedef["rol"], hedef_kullanici=hedef["kullanici"])
                    messages.success(request, f"{talep.referans_no} numaralı evrak {hedef['ad_soyad']} kullanıcısına geri çekildi.")
        next_url = request.POST.get("next") or ""
        return redirect(next_url if next_url.startswith("/izin-ara/") else "izin_ara")

    qs = IzinTalebi.objects.select_related("izin_turu", "kullanici", "kontrol_eden", "kimin_uzerinde").prefetch_related("islem_gecmisi__kullanici")
    qs = qs.filter(_talep_arama_filtresi(request.user)).distinct()
    qs = qs.annotate(liste_tarihi=Coalesce("kurum_kayit_tarihi", TruncDate("olusturma_zamani")))
    q = (request.GET.get("q") or "").strip()
    evrak_no = (request.GET.get("evrak_no") or "").strip()
    durum = (request.GET.get("durum") or "").strip()
    tarih_baslangic = (request.GET.get("tarih_baslangic") or "").strip()
    tarih_bitis = (request.GET.get("tarih_bitis") or "").strip()
    tarih_araligi = (request.GET.get("tarih_araligi") or "7").strip()
    if tarih_araligi not in {"7", "30", "90", "180", "365", "all"}:
        tarih_araligi = "7"
    birim = (request.GET.get("birim") or "").strip()
    alt_birim = (request.GET.get("alt_birim") or "").strip()
    birim_metin = (request.GET.get("birim_metin") or "").strip()
    personel_id = (request.GET.get("personel") or "").strip()
    personel_metin = (request.GET.get("personel_metin") or "").strip()
    imzalayan_id = (request.GET.get("imzalayan") or "").strip()
    imzalayan_metin = (request.GET.get("imzalayan_metin") or "").strip()
    genel_arama_parcalari = []
    if q:
        q_parcalari = [parca for parca in q.split() if parca]
        genel_arama_parcalari = [_normalize(parca) for parca in q_parcalari if _normalize(parca)]
        arama_terimleri = []
        for terim in [q, *q_parcalari]:
            if terim and _normalize(terim) not in {_normalize(eski) for eski in arama_terimleri}:
                arama_terimleri.append(terim)

        personel_modeli = _personel_veri_modeli()
        personel_lookup = Q()
        for terim in arama_terimleri:
            personel_lookup |= _personel_arama_lookup(terim, personel_modeli)
        personel_ids = _talep_personel_idleri(personel_modeli.objects.filter(personel_lookup))

        evrak_alanlari = ["evrak_no", "ait_oldugu_yil", "izin_sebebi", "aciklama", "izin_turu__ad"]
        kullanici_alanlari = [
            "kullanici__username",
            "kullanici__first_name",
            "kullanici__last_name",
            "kontrol_eden__username",
            "kontrol_eden__first_name",
            "kontrol_eden__last_name",
            "kimin_uzerinde__username",
            "kimin_uzerinde__first_name",
            "kimin_uzerinde__last_name",
            "islem_gecmisi__kullanici__username",
            "islem_gecmisi__kullanici__first_name",
            "islem_gecmisi__kullanici__last_name",
        ]
        genel_eslesme = Q(personel_id__in=personel_ids)
        for terim in arama_terimleri:
            genel_eslesme |= talep_numarasi_filtresi(terim)
            genel_eslesme |= _varyantli_icontains(evrak_alanlari + kullanici_alanlari, terim)
        for durum_kodu, durum_adi in IzinTalebi.DURUM_CHOICES:
            if any(parca in _normalize(durum_kodu) or parca in _normalize(durum_adi) for parca in genel_arama_parcalari):
                genel_eslesme |= Q(durum=durum_kodu)
        if q.isdigit() and len(q) == 4:
            genel_eslesme |= (
                Q(kurum_kayit_tarihi__year=int(q))
                | Q(baslangic_tarihi__year=int(q))
                | Q(bitis_tarihi__year=int(q))
            )
        qs = qs.filter(genel_eslesme).distinct()
    if evrak_no:
        qs = qs.filter(talep_numarasi_filtresi(evrak_no))
    if durum:
        qs = qs.filter(durum=durum)
    for raw, lookup in ((tarih_baslangic, "liste_tarihi__gte"), (tarih_bitis, "liste_tarihi__lte")):
        if raw:
            try:
                qs = qs.filter(**{lookup: date.fromisoformat(raw)})
            except ValueError:
                messages.warning(request, "Geçersiz tarih filtresi dikkate alınmadı.")
    if not tarih_baslangic and not tarih_bitis and tarih_araligi != "all":
        qs = qs.filter(liste_tarihi__gte=timezone.localdate() - timedelta(days=int(tarih_araligi) - 1))
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
        "evrak_no": "Evrak / Takip No",
        "kurum_kayit": "Kayıt / Oluşturulma Tarihi",
        "olusturma": "Oluşturulma",
        "olusturan": "Oluşturan",
        "imzalayanlar": "İmzalayanlar",
        "durum": "Durum",
        "kimde": "Kimin üzerinde",
    }
    if sort not in sort_labels:
        sort = "olusturma"
    if direction not in {"asc", "desc"}:
        direction = "desc"

    per_page_options = [10, 20, 50]
    try:
        per_page = int(request.GET.get("per_page") or 10)
    except ValueError:
        per_page = 10
    if per_page not in per_page_options:
        per_page = 10

    if sort == "evrak_no":
        qs = qs.annotate(liste_numarasi=Coalesce(
            NullIf("evrak_no", Value("")),
            Concat(Value("TALEP-"), Cast(ExtractYear("olusturma_zamani"), CharField()), Value("-"), Cast("pk", CharField())),
        ))
    db_sorts = {"olusturma": "olusturma_zamani", "kurum_kayit": "liste_tarihi", "evrak_no": "liste_numarasi"}
    db_page = None
    if not q and sort in db_sorts:
        prefix = "-" if direction == "desc" else ""
        db_page = Paginator(qs.order_by(prefix + db_sorts[sort], prefix + "pk"), per_page).get_page(request.GET.get("page"))
    talepler = db_page.object_list if db_page else qs.order_by("-olusturma_zamani", "-id")
    rows = []
    for talep in talepler:
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
        if q and genel_arama_parcalari:
            personel_text = ""
            if personel:
                personel_text = " ".join(str(getattr(personel, alan, "") or "") for alan in (
                    "sicil_no", "tc_no", "ad", "soyad", "gorev_birimi", "alt_birimi", "gorevi", "unvani"
                ))
            kullanici_text = " ".join([
                talep.kullanici.username,
                talep.kullanici.first_name,
                talep.kullanici.last_name,
                olusturan_text,
                *(f"{user.username} {user.first_name} {user.last_name} {user.get_full_name()}" for user in imzalayanlar),
                *(
                    [f"{talep.kontrol_eden.username} {talep.kontrol_eden.first_name} {talep.kontrol_eden.last_name} {talep.kontrol_eden.get_full_name()}"]
                    if talep.kontrol_eden
                    else []
                ),
                *(
                    [f"{talep.kimin_uzerinde.username} {talep.kimin_uzerinde.first_name} {talep.kimin_uzerinde.last_name} {kimde_text}"]
                    if talep.kimin_uzerinde
                    else []
                ),
            ])
            arama_blob = _normalize(" ".join([
                talep.evrak_no or "",
                talep.takip_no,
                talep.ait_oldugu_yil or "",
                talep.izin_sebebi or "",
                talep.aciklama or "",
                talep.izin_turu.ad if talep.izin_turu else "",
                talep.get_durum_display(),
                personel_text,
                birim_text,
                kullanici_text,
            ]))
            q_norm = _normalize(q)
            if q_norm not in arama_blob and not all(parca in arama_blob for parca in genel_arama_parcalari):
                continue
        geri_cekme_hedefleri = _geri_cekme_hedefleri(talep) if talep.durum in {"PARAF_BEKLIYOR", "IMZA_SURECINDE"} and _tip_yetkisi_var(request.user, talep.personel_tipi, "evrak_geri_cekme") else []
        rows.append({
            "talep": talep,
            "personel": personel,
            "kimde": talep.kimin_uzerinde,
            "imzalayanlar": imzalayanlar,
            "can_sahibine_iade": (
                talep.durum in {"PARAF_BEKLIYOR", "IMZA_SURECINDE"}
                and _tip_yetkisi_var(request.user, getattr(talep, "personel_tipi", "MEMUR"), "sahibine_iade")
            ),
            "geri_cekme_hedefleri": geri_cekme_hedefleri,
            "can_geri_cek": (
                talep.durum in {"PARAF_BEKLIYOR", "IMZA_SURECINDE"}
                and _tip_yetkisi_var(request.user, getattr(talep, "personel_tipi", "MEMUR"), "evrak_geri_cekme")
                and bool(geri_cekme_hedefleri)
            ),
            "sort_values": {
                "evrak_tanimi": "İzin Talep Formu",
                "birim": birim_text,
                "konu": "İzin İsteği",
                "evrak_no": talep.referans_no,
                "kurum_kayit": talep.liste_tarihi or date.min,
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
    if db_page:
        paginator, page_obj = db_page.paginator, db_page
    else:
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
        "tarih_bitis": tarih_bitis, "tarih_araligi": tarih_araligi, "birim": birim, "alt_birim": alt_birim,
        "birim_metin": birim_metin, "secili_personel": secili_personel,
        "personel_metin": personel_metin, "secili_imzalayan": secili_imzalayan,
        "imzalayan_metin": imzalayan_metin,
        "page_obj": page_obj, "paginator": paginator, "per_page": per_page,
        "per_page_options": per_page_options, "pagination_query": pagination_query,
        "page_size_query": page_size_query, "sort_links": sort_links,
        "sort": sort, "direction": direction,
        "can_delete": any(row.get("can_delete") for row in rows),
        "genis_arama_yetkisi": _genis_izin_arama_yetkisi(request.user),
    })
