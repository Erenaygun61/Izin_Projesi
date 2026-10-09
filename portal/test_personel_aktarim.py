import uuid
from unittest.mock import Mock, patch

from django.apps import apps
from django.contrib.auth.models import User
from django.db import OperationalError, connection
from django.test import Client, TestCase, override_settings
from django.urls import reverse

from . import views
from .models import IzinTalebi, PersonelAktarim, PortalPersonel, Unvan, YonetimYetkisi
from .personel_senkron import PERSONEL_MODELLERI, personel_durumlarini_esitle


@override_settings(PASSWORD_HASHERS=["django.contrib.auth.hashers.MD5PasswordHasher"])
class PersonelAktarimTests(TestCase):
    @classmethod
    def setUpClass(cls):
        # Source models are unmanaged: supply their schemas only in the isolated test DB.
        mevcut = set(connection.introspection.table_names())
        with connection.schema_editor() as editor:
            for model in apps.get_app_config("portal").get_models():
                if not model._meta.managed and model._meta.db_table not in mevcut:
                    editor.create_model(model)
                    mevcut.add(model._meta.db_table)
        super().setUpClass()

    @classmethod
    def setUpTestData(cls):
        cls.admin = User.objects.create_superuser("transfer_admin", password="test-only")
        unvan = Unvan.objects.create(ad="TEST UNVANI")
        cls.kaynaklar = []
        for index, (tip, model) in enumerate(PERSONEL_MODELLERI.items()):
            for sicil in (1, 2):
                cls.kaynaklar.append((tip, model.objects.create(
                    id=sicil, sicil_no=str(sicil), ad=f"Test{index}{sicil}", soyad="Personel",
                    tc_no=f"12345678{index}{sicil}0", unvan=unvan,
                    gorev_birimi="Test Birimi", durum="Aktif",
                )))

    def setUp(self):
        self.client.force_login(self.admin)

    def baslat(self, **params):
        values = {"tip": "MEMUR", "islem": "tum_tipleri_aktar"}
        values.update(params)
        response = self.client.post(reverse("personel_aktarim_baslat"), values)
        self.assertIn(response.status_code, (200, 201))
        return response.json()

    def adim(self, job, cursor=None):
        response = self.client.post(reverse("personel_aktarim_adim", args=[job["id"]]), {
            "processed": job["processed"] if cursor is None else cursor,
        })
        self.assertEqual(response.status_code, 200)
        return response.json()

    def tamamla(self, job):
        while not job["done"]:
            job = self.adim(job)
        return job

    def test_tum_tipler_gercek_sayacla_yuzde_yuze_ulasir(self):
        job = self.baslat()
        self.assertEqual((job["total"], job["processed"], job["percent"]), (6, 0, 0))
        job = self.adim(job)
        self.assertEqual((job["processed"], job["percent"]), (5, 83))
        job = self.tamamla(job)
        self.assertEqual((job["percent"], job["success"], job["failed"]), (100, 6, 0))
        self.assertEqual(PortalPersonel.objects.count(), 6)

    def test_tek_tip_ve_secili_kayit_aktarilir(self):
        job = self.tamamla(self.baslat(tip="ISCI", islem="tumunu_aktar"))
        self.assertEqual(job["success"], 2)
        self.assertEqual(set(PortalPersonel.objects.values_list("personel_tipi", flat=True)), {"ISCI"})
        job = self.tamamla(self.baslat(islem="secilenleri_guncelle", source_ids=["SIRKET:1", "SIRKET:1"]))
        self.assertEqual((job["total"], job["success"]), (1, 1))

    def test_baslat_ve_adim_tekrarlari_idempotenttir(self):
        request_id = str(uuid.uuid4())
        job = self.baslat(request_id=request_id)
        tekrar = self.baslat(request_id=request_id)
        self.assertEqual(job["id"], tekrar["id"])
        self.assertEqual(PersonelAktarim.objects.count(), 1)
        job = self.adim(job)
        with patch("portal.views.ozlukten_portala_aktar") as aktar:
            response = self.adim(job, cursor=0)
            aktar.assert_not_called()
        self.assertEqual(response["processed"], 5)
        self.assertEqual(PortalPersonel.objects.count(), 5)

    def test_hata_ayrilir_ve_yalniz_hatalilar_yeniden_denenir(self):
        original = views.ozlukten_portala_aktar
        def aktar(kaynak, tip):
            if tip == "MEMUR" and kaynak.pk == 2:
                raise ValueError("test import failure")
            return original(kaynak, tip)
        with patch("portal.views.ozlukten_portala_aktar", side_effect=aktar), patch("portal.personel_aktarim.logger"):
            job = self.tamamla(self.baslat())
        self.assertEqual((job["percent"], job["success"], job["failed"]), (100, 5, 1))
        self.assertFalse(PortalPersonel.objects.filter(personel_tipi="MEMUR", source_personel_id=2).exists())
        tekrar = self.tamamla(self.baslat(retry=job["id"]))
        self.assertEqual((tekrar["total"], tekrar["success"]), (1, 1))
        self.assertEqual(PortalPersonel.objects.count(), 6)

    def test_kilit_hatasinda_parca_geri_alinir_ve_devam_edilir(self):
        original = views.ozlukten_portala_aktar
        def aktar(kaynak, tip):
            if kaynak.pk == 2:
                raise OperationalError("database is locked")
            return original(kaynak, tip)
        job = self.baslat()
        with patch("portal.views.ozlukten_portala_aktar", side_effect=aktar), patch("portal.personel_aktarim.logger"):
            response = self.client.post(reverse("personel_aktarim_adim", args=[job["id"]]), {"processed": 0})
        self.assertEqual(response.status_code, 503)
        self.assertEqual(PortalPersonel.objects.count(), 0)
        record = PersonelAktarim.objects.get(pk=job["id"])
        self.assertEqual(record.islenen, 0)
        self.assertEqual(self.tamamla(job)["success"], 6)

    def test_sonradan_silinen_kayit_hata_olarak_raporlanir(self):
        job = self.baslat(tip="MEMUR", islem="tumunu_aktar")
        PERSONEL_MODELLERI["MEMUR"].objects.filter(pk=2).delete()
        with patch("portal.personel_aktarim.logger"):
            job = self.tamamla(job)
        self.assertEqual((job["success"], job["failed"]), (1, 1))

    def test_yetkisiz_ve_baska_kullanici_aktarimina_erisim_yok(self):
        job = self.baslat()
        ordinary = User.objects.create_user("ordinary")
        self.client.force_login(ordinary)
        self.assertEqual(self.client.post(reverse("personel_aktarim_baslat"), {"tip": "MEMUR"}).status_code, 403)
        self.assertEqual(self.client.get(reverse("personel_aktarim_durum", args=[job["id"]])).status_code, 403)
        staff = User.objects.create_user("other_admin", is_staff=True)
        self.client.force_login(staff)
        self.assertEqual(self.client.get(reverse("personel_aktarim_durum", args=[job["id"]])).status_code, 403)
        self.assertEqual(self.client.post(reverse("personel_aktarim_adim", args=[job["id"]]), {"processed": 0}).status_code, 403)
        YonetimYetkisi.objects.create(kullanici=staff, personel_veri_yonetimi=True)
        self.assertEqual(self.client.get(reverse("personel_aktarim_durum", args=[job["id"]])).status_code, 404)
        self.assertEqual(self.client.post(reverse("personel_aktarim_adim", args=[job["id"]]), {"processed": 0}).status_code, 404)
        self.client.logout()
        self.assertEqual(self.client.get(reverse("personel_aktarim_durum", args=[job["id"]])).status_code, 401)

    def test_csrf_ve_girdi_dogrulamasi(self):
        secure_client = Client(enforce_csrf_checks=True)
        secure_client.force_login(self.admin)
        self.assertEqual(secure_client.post(reverse("personel_aktarim_baslat"), {"tip": "MEMUR"}).status_code, 403)
        self.assertEqual(PersonelAktarim.objects.count(), 0)
        for params in ({"tip": "BOZUK"}, {"islem": "sil"}, {"islem": "secilenleri_guncelle"},
                       {"islem": "secilenleri_guncelle", "source_ids": ["MEMUR:yanlis"]}, {"request_id": "bad-uuid"}):
            with self.subTest(params=params):
                self.assertEqual(self.client.post(reverse("personel_aktarim_baslat"), params).status_code, 400)
        self.assertEqual(self.client.get(reverse("personel_aktarim_baslat")).status_code, 405)

    def test_pasif_ve_aktif_durumu_yerel_bilgileri_bozmadan_esitlenir(self):
        self.tamamla(self.baslat())
        for tip, model in PERSONEL_MODELLERI.items():
            kart = PortalPersonel.objects.get(personel_tipi=tip, source_personel_id=1)
            kart.gorev_birimi = "Yerel Birim"
            kart.unvan = "Yerel Unvan"
            kart.save()
            model.objects.filter(pk=1).update(durum="Pasif")
        self.client.get(reverse("personel_aktarim_durum", args=[PersonelAktarim.objects.first().pk]))
        for kart in PortalPersonel.objects.filter(source_personel_id=1):
            self.assertEqual((kart.durum, kart.gorev_birimi, kart.unvan), ("Pasif", "Yerel Birim", "Yerel Unvan"))
        for model in PERSONEL_MODELLERI.values():
            model.objects.filter(pk=1).update(durum="Aktif")
        self.assertEqual(personel_durumlarini_esitle(), 3)
        self.assertEqual(PortalPersonel.objects.filter(durum="Aktif").count(), 6)

    def test_kaynagi_silinen_portal_karti_pasiflesir_silinmez(self):
        kaynak = self.kaynaklar[0][1]
        kart, _ = views.ozlukten_portala_aktar(kaynak)
        PERSONEL_MODELLERI["MEMUR"].objects.filter(pk=kaynak.pk).delete()
        personel_durumlarini_esitle()
        kart.refresh_from_db()
        self.assertEqual(kart.durum, "Pasif")

    def test_devre_disi_kullanici_aktarimla_acilmaz_sifre_ve_bag_korunur(self):
        kart, _ = views.ozlukten_portala_aktar(self.kaynaklar[0][1])
        user = kart.kullanici
        user.username = "custom-login"
        user.is_active = False
        user.set_password("custom-password")
        user.save()
        tekrar, _ = views.ozlukten_portala_aktar(self.kaynaklar[0][1])
        user.refresh_from_db()
        self.assertEqual(tekrar.kullanici_id, user.pk)
        self.assertEqual(user.username, "custom-login")
        self.assertFalse(user.is_active)
        self.assertTrue(user.check_password("custom-password"))

    def test_ayni_isim_ve_sicil_farkli_kisilerin_hesabini_birlestirmez(self):
        ilk = self.kaynaklar[0][1]
        ikinci = self.kaynaklar[2][1]
        ikinci.ad, ikinci.soyad = ilk.ad, ilk.soyad
        ikinci.save()
        kart1, _ = views.ozlukten_portala_aktar(ilk, "MEMUR")
        kart2, _ = views.ozlukten_portala_aktar(ikinci, "ISCI")
        self.assertNotEqual(kart1.kullanici_id, kart2.kullanici_id)
        self.assertEqual(views.kullanicinin_personeli(kart2.kullanici).pk, kart2.pk)

    def test_pasif_personel_yeni_taslak_ve_onay_olusturamaz(self):
        kaynak = self.kaynaklar[0][1]
        kart, _ = views.ozlukten_portala_aktar(kaynak)
        PERSONEL_MODELLERI["MEMUR"].objects.filter(pk=kaynak.pk).update(durum="Pasif")
        self.client.force_login(kart.kullanici)
        response = self.client.get(reverse("izin_evrak", args=["memur", kart.pk]))
        self.assertEqual(response.status_code, 302)
        self.assertEqual(IzinTalebi.objects.count(), 0)
        form = Mock()
        _, hata = views._talebi_kaydet(IzinTalebi(), form, personel=kart, kullanici=kart.kullanici,
                                      evrak_no="TEST", onaya_sun=True)
        self.assertIn("Pasif", hata)
        form.save.assert_not_called()

    def test_personel_arama_ayni_idli_diger_tipi_acmaz(self):
        self.tamamla(self.baslat())
        kart = PortalPersonel.objects.get(personel_tipi="MEMUR", source_personel_id=1)
        self.client.force_login(kart.kullanici)
        response = self.client.get(reverse("personel_ara"))
        self.assertEqual([row["id"] for row in response.json()["results"]], [kart.pk])
        self.client.logout()
        self.assertEqual(self.client.get(reverse("personel_ara")).status_code, 302)

    def test_yonetim_ve_temel_ekranlar_acilir(self):
        names = ("dashboard", "izin_menu", "yonetim", "personel_veri_yonetimi", "izin_talepleri_yonetim",
                 "kullanici_yetkileri", "kontrol_eden_yonetimi", "gorev_birimi_yonetimi", "izin_ara")
        for name in names:
            with self.subTest(name=name):
                self.assertEqual(self.client.get(reverse(name), follow=True).status_code, 200)
        for tip in ("MEMUR", "ISCI", "SIRKET"):
            self.assertEqual(self.client.get(reverse("personel_veri_yonetimi"), {"tip": tip}).status_code, 200)

    def test_aktarim_hatasi_yarim_kart_birakmaz(self):
        onceki_user = User.objects.count()
        with patch("portal.views._varsayilan_portal_izin_hakki", side_effect=ValueError("test failure")):
            with self.assertRaises(ValueError):
                views.ozlukten_portala_aktar(self.kaynaklar[2][1], "ISCI")
        self.assertFalse(PortalPersonel.objects.exists())
        self.assertEqual(User.objects.count(), onceki_user)

    def test_eski_ayni_isimli_hesap_yeniden_olusturulmaz(self):
        kaynak1, kaynak2 = self.kaynaklar[0][1], self.kaynaklar[1][1]
        kaynak2.ad, kaynak2.soyad = kaynak1.ad, kaynak1.soyad
        kaynak2.save()
        kart1, _ = views.ozlukten_portala_aktar(kaynak1)
        kart2, _ = views.ozlukten_portala_aktar(kaynak2)
        user = kart2.kullanici
        user.set_unusable_password()
        user.save()
        kart2.kullanici = None
        kart2.save()
        count = User.objects.count()
        kart2, _ = views.ozlukten_portala_aktar(kaynak2)
        self.assertEqual(kart2.kullanici_id, user.pk)
        self.assertEqual(User.objects.count(), count)
        user.refresh_from_db()
        self.assertFalse(user.has_usable_password())

    def test_yerel_kayit_kaynak_pasifligini_degistiremez(self):
        kaynak = self.kaynaklar[0][1]
        PERSONEL_MODELLERI["MEMUR"].objects.filter(pk=kaynak.pk).update(durum="Pasif")
        response = self.client.post(reverse("personel_veri_yonetimi"), {
            "islem": "yerel_kaydet", "tip": "MEMUR", "source_personel_id": kaynak.pk,
            "ad": kaynak.ad, "soyad": kaynak.soyad, "sicil_no": kaynak.sicil_no,
            "tc_no": kaynak.tc_no, "durum": "Aktif", "unvan": "Yerel Test",
        })
        self.assertEqual(response.status_code, 302)
        kart = PortalPersonel.objects.get(personel_tipi="MEMUR", source_personel_id=kaynak.pk)
        self.assertEqual((kart.durum, kart.unvan), ("Pasif", "Yerel Test"))

    def test_gecersiz_yerel_kayit_sunucu_hatasi_vermez(self):
        response = self.client.post(reverse("personel_veri_yonetimi"), {
            "islem": "yerel_kaydet", "source_personel_id": "bad-id",
        })
        self.assertEqual(response.status_code, 302)
        self.assertFalse(PortalPersonel.objects.exists())

    def test_kismi_aktarim_diger_tipte_evrak_kaynagini_engellemez(self):
        self.tamamla(self.baslat(islem="tumunu_aktar", tip="MEMUR"))
        kaynak = self.kaynaklar[2][1]
        with patch("portal.views._kendi_personeli_mi", return_value=True), patch("portal.views.izin_ozeti", return_value=([], 0)):
            response = self.client.get(reverse("izin_evrak", args=["isci", kaynak.pk]))
        self.assertEqual(response.status_code, 302)
        self.assertTrue(IzinTalebi.objects.filter(personel_id=kaynak.pk, personel_tipi="ISCI").exists())

    def test_gecmis_talep_ayni_idli_diger_tipi_arama_kapsamina_almaz(self):
        self.tamamla(self.baslat())
        user = User.objects.create_user("history-only")
        IzinTalebi.objects.create(kullanici=user, personel_id=1, personel_tipi="ISCI", evrak_no="TEST-HISTORY")
        self.client.force_login(user)
        response = self.client.get(reverse("personel_ara"))
        kart = PortalPersonel.objects.get(personel_tipi="ISCI", source_personel_id=1)
        self.assertEqual([row["id"] for row in response.json()["results"]], [kart.pk])
