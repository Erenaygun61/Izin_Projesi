from datetime import date
from decimal import Decimal
from types import SimpleNamespace
from unittest.mock import MagicMock, patch

from django.contrib.auth.models import User
from django.db import connection
from django.test import TestCase
from django.urls import reverse

from .models import GorevBirimiImzaci, ImzaciVekaleti, IzinIslemGecmisi, IzinTalebi, IzinTuru, PortalPersonel, YonetimYetkisi
from . import views


class OzlukAktarimTests(TestCase):
    """İmza tamamlanınca kullanılan iznin özlüğe aktarımını korur."""

    @patch("portal.views.KullanilanIzin.objects.create")
    @patch("portal.views.KullanilanIzin.objects.filter")
    @patch("portal.views._yillara_gore_haklar", return_value=[])
    @patch("portal.views.Personel.objects.filter")
    def test_memur_izni_personel_tipiyle_ozluge_aktarilir(
        self,
        personel_filter,
        haklar,
        kullanilan_filter,
        kullanilan_create,
    ):
        kaynak = SimpleNamespace(ise_baslama_tarihi=date(2018, 1, 1))
        personel_filter.return_value.first.return_value = kaynak
        kullanilan_filter.return_value.values.return_value.annotate.return_value = []

        talep = SimpleNamespace(
            ozluk_aktarildi_mi=False,
            izin_turu_id=1,
            baslangic_tarihi=date(2026, 8, 3),
            bitis_tarihi=date(2026, 8, 7),
            kullanilan_gun=Decimal("5.00"),
            personel_tipi="MEMUR",
            izin_turu=SimpleNamespace(yillik_izin_mi=False),
            ait_oldugu_yil="2026",
            evrak_no="2026-TEST",
            save=MagicMock(),
        )
        portal_personeli = SimpleNamespace(id=184)

        views.izni_ozluge_aktar(talep, portal_personeli)

        self.assertTrue(talep.ozluk_aktarildi_mi)
        self.assertEqual(kullanilan_create.call_args.kwargs["personel_tipi"], "MEMUR")
        talep.save.assert_called_once()


class IzinAramaGorunurlukTests(TestCase):
    @classmethod
    def setUpTestData(cls):
        with connection.cursor() as cursor:
            cursor.execute(
                """
                CREATE TABLE IF NOT EXISTS ozluk_izinturu (
                    id integer PRIMARY KEY,
                    kod varchar(30) NOT NULL UNIQUE,
                    ad varchar(120) NOT NULL UNIQUE,
                    yillik_izin_mi bool NOT NULL,
                    aktif_mi bool NOT NULL
                )
                """
            )
            cursor.execute(
                """
                CREATE TABLE IF NOT EXISTS ozluk_personel (
                    id integer PRIMARY KEY,
                    sicil_no varchar(20) NOT NULL UNIQUE,
                    ad varchar(50) NOT NULL,
                    soyad varchar(50) NOT NULL,
                    tc_no varchar(11) NOT NULL UNIQUE,
                    unvan_id integer NULL,
                    cep_tel varchar(20) NOT NULL,
                    gorev_birimi varchar(100) NOT NULL,
                    alt_birimi varchar(100) NOT NULL,
                    gorevi varchar(150) NOT NULL,
                    durum varchar(20) NOT NULL,
                    ise_baslama_tarihi date NULL
                )
                """
            )

    def test_teklif_parafindan_sonra_yonetici_arama_listesinde_gorur(self):
        olusturan = User.objects.create_user(username="olusturan")
        kontrol_eden = User.objects.create_user(username="kontrol")
        yonetici = User.objects.create_superuser(username="yonetici")
        evrak_yetkilisi = User.objects.create_user(username="evrak_yetkilisi")
        YonetimYetkisi.objects.create(kullanici=evrak_yetkilisi, aktif_mi=True, evrak_yonetimi=True)
        talep = IzinTalebi.objects.create(
            kullanici=olusturan,
            personel_id=1,
            personel_tipi="MEMUR",
            evrak_no="TEST-1",
            durum="PARAF_BEKLIYOR",
            kontrol_eden=kontrol_eden,
            kimin_uzerinde=kontrol_eden,
            imza_sirasi=1,
        )
        IzinIslemGecmisi.objects.create(talep=talep, kullanici=olusturan, islem="ONAYA_SUNULDU", rol="Teklif Eden")
        IzinIslemGecmisi.objects.create(talep=talep, kullanici=olusturan, islem="IMZA", rol="Teklif Eden")

        self.assertTrue(IzinTalebi.objects.filter(views._talep_arama_filtresi(yonetici), pk=talep.pk).exists())
        self.assertTrue(IzinTalebi.objects.filter(views._talep_arama_filtresi(evrak_yetkilisi), pk=talep.pk).exists())

    def test_genel_arama_bosluklu_ad_soyadi_olusturan_ve_imzalayanda_bulur(self):
        olusturan = User.objects.create_user(
            username="erenaygun",
            first_name="EREN",
            last_name="AYGÜN",
        )
        imzalayan = User.objects.create_user(
            username="zekiyamic",
            first_name="ZEKİ",
            last_name="YAMİÇ",
        )
        arayan = User.objects.create_user(username="arama_yetkilisi", is_staff=True)
        YonetimYetkisi.objects.create(kullanici=arayan, evrak_yonetimi=True)
        PortalPersonel.objects.create(
            personel_tipi="MEMUR",
            source_personel_id=184,
            sicil_no="184",
            ad="EREN",
            soyad="AYGÜN",
            tc_no="12345678901",
            gorev_birimi="İnsan Kaynakları ve Eğitim Dairesi Başkanlığı",
            alt_birimi="İnsan Kaynakları Şube Müdürlüğü",
            gorevi="Mühendis",
        )
        talep = IzinTalebi.objects.create(
            kullanici=olusturan,
            personel_id=184,
            personel_tipi="MEMUR",
            evrak_no="2026-21",
            durum="IMZA_SURECINDE",
            kimin_uzerinde=imzalayan,
            kurum_kayit_tarihi=date(2026, 7, 30),
        )
        IzinIslemGecmisi.objects.create(talep=talep, kullanici=olusturan, islem="IMZA", rol="Teklif Eden")
        IzinIslemGecmisi.objects.create(talep=talep, kullanici=imzalayan, islem="KONTROL_PARAFI", rol="Kontrol Eden")

        self.client.force_login(arayan)
        response = self.client.get(reverse("izin_ara"), {"q": "eren aygün", "tarih_araligi": "all"})

        self.assertContains(response, "2026-21")

    def test_aktif_gorev_onizleme_paraf_ekranina_yonlenir(self):
        olusturan = User.objects.create_user(username="talep_sahibi")
        imzaci = User.objects.create_user(username="aktif_imzaci")
        PortalPersonel.objects.create(
            personel_tipi="MEMUR",
            source_personel_id=260,
            sicil_no="260",
            ad="ELİF BÜŞRA",
            soyad="EROĞLU",
            tc_no="28751428782",
            gorev_birimi="İnsan Kaynakları ve Eğitim Dairesi Başkanlığı",
            alt_birimi="İnsan Kaynakları Şube Müdürlüğü",
            gorevi="Eğitmen",
        )
        talep = IzinTalebi.objects.create(
            kullanici=olusturan,
            personel_id=260,
            personel_tipi="MEMUR",
            evrak_no="2026-26",
            durum="IMZA_SURECINDE",
            kimin_uzerinde=imzaci,
            kurum_kayit_tarihi=date(2026, 7, 31),
        )

        self.client.force_login(imzaci)
        with patch("portal.views.talep_aktif_imzacilari", return_value=[("Amir", imzaci)]):
            response = self.client.get(reverse("evrak_onizleme", args=[talep.id]))

        self.assertRedirects(response, reverse("evrak_paraf", args=[talep.id]), fetch_redirect_response=False)

    def test_vekalet_pasife_alininca_bekleyen_evrak_asli_imzaciya_doner(self):
        eren = User.objects.create_user(username="erenaygun", first_name="EREN", last_name="AYGÜN")
        zeki = User.objects.create_user(username="zekiyamic", first_name="ZEKİ", last_name="YAMİÇ")
        kontrol = User.objects.create_user(username="kontrol")
        PortalPersonel.objects.create(
            personel_tipi="MEMUR",
            source_personel_id=184,
            sicil_no="184",
            ad="EREN",
            soyad="AYGÜN",
            tc_no="42367983732",
            gorev_birimi="İnsan Kaynakları ve Eğitim Dairesi Başkanlığı",
            alt_birimi="İnsan Kaynakları Şube Müdürlüğü",
            gorevi="Mühendis",
        )
        GorevBirimiImzaci.objects.create(
            birim_adi="İnsan Kaynakları ve Eğitim Dairesi Başkanlığı",
            alt_birim_adi="İnsan Kaynakları Şube Müdürlüğü",
            birim_turu="ALT_BIRIM",
            sube_muduru=zeki,
            amir_kullanici=zeki,
            aktif_mi=True,
        )
        vekalet = ImzaciVekaleti.objects.create(
            asil_kullanici=zeki,
            vekil_kullanici=eren,
            baslangic_tarihi=date(2026, 8, 1),
            bitis_tarihi=date(2026, 8, 31),
            aktif_mi=True,
        )
        talep = IzinTalebi.objects.create(
            kullanici=eren,
            personel_id=184,
            personel_tipi="MEMUR",
            evrak_no="2026-32",
            durum="IMZA_SURECINDE",
            kontrol_eden=kontrol,
            kimin_uzerinde=eren,
            imza_sirasi=0,
            kurum_kayit_tarihi=date(2026, 8, 7),
        )

        imzacilar = views.talep_aktif_imzacilari(talep)
        self.assertEqual(imzacilar[0][1], eren)
        self.assertIn("V.", imzacilar[0][0])

        vekalet.aktif_mi = False
        vekalet.save(update_fields=["aktif_mi"])
        views._bekleyen_vekalet_evraklarini_guncelle(kullanicilar=[zeki.id, eren.id])

        talep.refresh_from_db()
        self.assertEqual(talep.kimin_uzerinde, zeki)

    @patch("portal.views.izin_pdf_olustur", return_value=b"%PDF-ISLAK%")
    @patch("portal.views.imzacilari_bul")
    @patch("portal.views.izin_ozeti")
    @patch("portal.views._kendi_personeli_mi", return_value=True)
    def test_islak_cikti_taslagi_onaya_sunmadan_pdf_uretir(self, kendi_personeli_mi, izin_ozeti, imzacilari_bul, pdf_olustur):
        izin_ozeti.return_value = ([SimpleNamespace(izin_yili=2026, pk=1, kalan=Decimal("25"))], Decimal("25"))
        olusturan = User.objects.create_user(username="erenaygun")
        imzaci = User.objects.create_user(username="imzaci", first_name="Zeki", last_name="Yamiç")
        imzacilari_bul.return_value = [("Şube Müdürü", imzaci)]
        izin_turu = IzinTuru.objects.create(kod="YILLIK", ad="Yıllık İzin", yillik_izin_mi=True, aktif_mi=True)
        personel = PortalPersonel.objects.create(
            personel_tipi="MEMUR",
            source_personel_id=184,
            sicil_no="184",
            ad="EREN",
            soyad="AYGÜN",
            tc_no="42367983732",
            gorev_birimi="İnsan Kaynakları ve Eğitim Dairesi Başkanlığı",
            alt_birimi="İnsan Kaynakları Şube Müdürlüğü",
            gorevi="Mühendis",
        )
        talep = IzinTalebi.objects.create(
            kullanici=olusturan,
            personel_id=personel.source_personel_id,
            personel_tipi="MEMUR",
            evrak_no="2026-27",
            durum="TASLAK",
            kimin_uzerinde=olusturan,
            kurum_kayit_tarihi=date(2026, 7, 31),
        )

        self.client.force_login(olusturan)
        response = self.client.post(reverse("izin_talebi_islak_cikti", args=[talep.id]), {
            "cep_telefonu": "",
            "izin_turu": str(izin_turu.id),
            "ait_oldugu_yil": "2026",
            "izin_hakki": "25",
            "kullanilan_gun": "5",
            "kalan_izin": "20",
            "baslangic_tarihi": "2026-08-10",
            "bitis_tarihi": "2026-08-14",
            "goreve_baslama_tarihi": "2026-08-17",
        })

        self.assertEqual(response.status_code, 200)
        self.assertEqual(response["Content-Type"], "application/pdf")
        self.assertIn("inline", response["Content-Disposition"])
        self.assertEqual(response.content, b"%PDF-ISLAK%")
        pdf_olustur.assert_called_once()
        self.assertTrue(pdf_olustur.call_args.kwargs["islak_imza"])
        talep.refresh_from_db()
        self.assertEqual(talep.durum, "TASLAK")
