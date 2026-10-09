from decimal import Decimal
from unittest.mock import Mock, patch

from django.apps import apps
from django.contrib.auth.models import User
from django.contrib.messages import SUCCESS, get_messages
from django.db import connection
from django.test import TestCase
from django.urls import reverse
from django.utils import timezone

from . import views
from .models import IzinBildirimKapatma, IzinIslemGecmisi, IzinTalebi, PortalPersonel, YonetimYetkisi


class EvrakReddiTests(TestCase):
    @classmethod
    def setUpClass(cls):
        mevcut = set(connection.introspection.table_names())
        with connection.schema_editor() as editor:
            for model in apps.get_app_config("portal").get_models():
                if not model._meta.managed and model._meta.db_table not in mevcut:
                    editor.create_model(model)
                    mevcut.add(model._meta.db_table)
        super().setUpClass()

    @classmethod
    def setUpTestData(cls):
        cls.sahip = User.objects.create_user("test-sahip")
        cls.kontrol = User.objects.create_user("test-kontrol")
        cls.imzaci = User.objects.create_user("test-imzaci")
        YonetimYetkisi.objects.create(kullanici=cls.kontrol, aktif_mi=True, memur_kontrol_eden=True)
        PortalPersonel.objects.create(personel_tipi="MEMUR", source_personel_id=1,
                                      kullanici=cls.sahip, sicil_no="1", ad="Test", soyad="Sahip", tc_no="12345678901")

    def talep_olustur(self, durum="IMZA_SURECINDE"):
        return IzinTalebi.objects.create(kullanici=self.sahip, personel_id=1,
                                        evrak_no="TEST-RED", personel_tipi="MEMUR", durum=durum,
                                        kontrol_eden=self.kontrol, imza_sirasi=1 if durum == "PARAF_BEKLIYOR" else 0,
                                        kimin_uzerinde=self.kontrol if durum == "PARAF_BEKLIYOR" else self.imzaci)

    def reddet(self, talep, actor, neden="Test ret nedeni", follow=False):
        self.client.force_login(actor)
        with patch("portal.views.talep_aktif_imzacilari", return_value=[("Amir", self.imzaci)]), \
             patch("portal.views.izin_ozeti", return_value=([], Decimal("0"))):
            return self.client.post(reverse("evrak_paraf", args=[talep.pk]),
                                    {"islem": "reddet", "red_nedeni": neden}, follow=follow)

    def test_imzaci_reddi_basari_bildirimiyle_ana_sayfaya_doner(self):
        self.basariyi_dogrula("IMZA_SURECINDE", self.imzaci, "\u0130mzac\u0131")

    def test_kontrol_reddi_basari_bildirimiyle_ana_sayfaya_doner(self):
        self.basariyi_dogrula("PARAF_BEKLIYOR", self.kontrol, "Kontrol Eden")

    def basariyi_dogrula(self, durum, actor, rol):
        talep = self.talep_olustur(durum)
        response = self.reddet(talep, actor, follow=True)
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.redirect_chain, [(reverse("dashboard"), 302)])
        bildirimler = list(get_messages(response.wsgi_request))
        self.assertEqual(len(bildirimler), 1)
        self.assertEqual(bildirimler[0].level, SUCCESS)
        self.assertIn("ilk sahibine geri g\u00f6nderildi", str(bildirimler[0]))
        self.assertNotContains(response, "Bu evrak\u0131 g\u00f6r\u00fcnt\u00fcleme yetkiniz yok.")
        talep.refresh_from_db()
        self.assertEqual(talep.durum, "REDDEDILDI")
        self.assertEqual(talep.reddeden_id, actor.pk)
        self.assertEqual(talep.red_nedeni, "Test ret nedeni")
        self.assertIsNotNone(talep.reddedilme_tarihi)
        self.assertEqual(talep.islem_gecmisi.get(islem="RED").rol, rol)
        kayit = talep.denetim_kayitlari.get(islem="RED")
        self.assertEqual((kayit.onceki_durum, kayit.yeni_durum), (durum, "REDDEDILDI"))
        self.assertEqual(kayit.hedef_kullanici_id, self.sahip.pk)
        self.client.force_login(self.sahip)
        with patch("portal.views.izin_ozeti", return_value=([], Decimal("0"))):
            sahibi_ekrani = self.client.get(reverse("dashboard"))
        self.assertIn(talep, sahibi_ekrani.context["gorevler"])

    def test_ret_nedeni_zorunlu_kalmaya_devam_eder(self):
        talep = self.talep_olustur()
        response = self.reddet(talep, self.imzaci, neden="   ")
        self.assertRedirects(response, reverse("evrak_paraf", args=[talep.pk]), fetch_redirect_response=False)
        talep.refresh_from_db()
        self.assertEqual(talep.durum, "IMZA_SURECINDE")
        self.assertFalse(talep.islem_gecmisi.filter(islem="RED").exists())

    def test_baska_kullanici_reddedemez(self):
        talep = self.talep_olustur()
        self.reddet(talep, self.sahip)
        talep.refresh_from_db()
        self.assertEqual(talep.durum, "IMZA_SURECINDE")
        self.assertIsNone(talep.reddeden_id)

    def test_islem_gecmisi_hatasinda_yarim_reddetme_kaydedilmez(self):
        talep = self.talep_olustur()
        with patch("portal.views._denetim_kaydi_ekle", side_effect=RuntimeError("Test audit failure")):
            with self.assertRaises(RuntimeError):
                self.reddet(talep, self.imzaci)
        talep.refresh_from_db()
        self.assertEqual(talep.durum, "IMZA_SURECINDE")
        self.assertEqual(talep.kimin_uzerinde_id, self.imzaci.pk)
        self.assertFalse(talep.islem_gecmisi.filter(islem="RED").exists())

    def amir_reddi_olustur(self):
        talep = self.talep_olustur()
        talep.kontrol_tarihi = timezone.now()
        talep.save(update_fields=["kontrol_tarihi"])
        IzinIslemGecmisi.objects.create(talep=talep, kullanici=self.kontrol, islem="KONTROL_PARAFI", rol="Kontrol Eden")
        self.reddet(talep, self.imzaci)
        talep.refresh_from_db()
        return talep

    def dashboard(self, user):
        self.client.force_login(user)
        with patch("portal.views.izin_ozeti", return_value=([], Decimal("0"))):
            return self.client.get(reverse("dashboard"))

    def test_amir_reddi_kontrolcuye_islem_degil_bilgi_olarak_gelir(self):
        talep = self.amir_reddi_olustur()
        response = self.dashboard(self.kontrol)
        self.assertEqual(list(response.context["gorevler"]), [])
        self.assertEqual(list(response.context["reddedilenler"]), [talep])
        self.assertContains(response, "Bilgi:")
        self.assertNotContains(response, "Dosyas\u0131na i\u015fleyiniz")
        self.assertContains(response, reverse("izin_bildirimini_kapat", args=[talep.pk]))
        self.assertTrue(views._talep_gorme_yetkisi(self.kontrol, talep))

    def test_eski_kontrol_tarihli_ret_bilgisi_onizlenebilir(self):
        talep = self.talep_olustur()
        talep.durum = "REDDEDILDI"
        talep.kontrol_tarihi = timezone.now()
        talep.reddeden = self.imzaci
        talep.kimin_uzerinde = None
        talep.save()
        self.assertTrue(views._talep_gorme_yetkisi(self.kontrol, talep))
        self.assertFalse(views._talep_gorme_yetkisi(User.objects.create_user("test-outside"), talep))

    def test_kontrolcu_bilgilendirmeyi_kapatabilir_evrak_ve_sahip_gorevi_korunur(self):
        talep = self.amir_reddi_olustur()
        onceki_gecmis = talep.islem_gecmisi.count()
        self.client.force_login(self.kontrol)
        url = reverse("izin_bildirimini_kapat", args=[talep.pk])
        self.assertRedirects(self.client.post(url), reverse("dashboard"), fetch_redirect_response=False)
        self.client.post(url)
        self.assertEqual(IzinBildirimKapatma.objects.filter(kullanici=self.kontrol, talep=talep).count(), 1)
        self.assertEqual(list(self.dashboard(self.kontrol).context["reddedilenler"]), [])
        talep.refresh_from_db()
        self.assertEqual((talep.durum, talep.islem_gecmisi.count()), ("REDDEDILDI", onceki_gecmis))
        owner = self.dashboard(self.sahip)
        self.assertIn(talep, owner.context["gorevler"])
        self.assertEqual(list(owner.context["reddedilenler"]), [])

    def test_sahip_ve_ilgisiz_kisi_kontrolcunun_bildirimini_kapatamaz(self):
        talep = self.amir_reddi_olustur()
        for user in (self.sahip, self.imzaci, User.objects.create_user("test-unrelated")):
            self.client.force_login(user)
            response = self.client.post(reverse("izin_bildirimini_kapat", args=[talep.pk]))
            self.assertEqual(response.status_code, 404)
        self.assertFalse(IzinBildirimKapatma.objects.filter(talep=talep).exists())

    def test_toplu_kapatma_bilgileri_kapatir_aktif_ve_sahip_gorevlerine_dokunmaz(self):
        red = self.amir_reddi_olustur()
        onay = IzinTalebi.objects.create(kullanici=self.sahip, kontrol_eden=self.kontrol,
                                        personel_id=1, evrak_no="TEST-ONAY", durum="ONAYLANDI")
        aktif = self.talep_olustur("PARAF_BEKLIYOR")
        kendi_reddi = IzinTalebi.objects.create(kullanici=self.kontrol, personel_id=1,
                                              evrak_no="TEST-KENDI-RED", durum="REDDEDILDI")
        self.client.force_login(self.kontrol)
        response = self.client.post(reverse("izin_bildirimlerini_kapat"))
        self.assertRedirects(response, reverse("dashboard"), fetch_redirect_response=False)
        kapatilanlar = set(IzinBildirimKapatma.objects.filter(kullanici=self.kontrol).values_list("talep_id", flat=True))
        self.assertEqual(kapatilanlar, {red.pk, onay.pk})
        dashboard = self.dashboard(self.kontrol)
        self.assertEqual(list(dashboard.context["reddedilenler"]), [])
        self.assertEqual(list(dashboard.context["imzalananlar"]), [])
        self.assertEqual({talep.pk for talep in dashboard.context["gorevler"]}, {aktif.pk, kendi_reddi.pk})

    def test_ayni_evrak_yeniden_reddedilirse_bilgi_tekrar_gorunur(self):
        talep = self.amir_reddi_olustur()
        IzinBildirimKapatma.objects.create(kullanici=self.kontrol, talep=talep)
        talep.durum = "IMZA_SURECINDE"
        talep.kimin_uzerinde = self.imzaci
        talep.save(update_fields=["durum", "kimin_uzerinde"])
        self.reddet(talep, self.imzaci, neden="Test second rejection")
        self.assertFalse(IzinBildirimKapatma.objects.filter(kullanici=self.kontrol, talep=talep).exists())
        self.assertEqual([t.pk for t in self.dashboard(self.kontrol).context["reddedilenler"]], [talep.pk])

    def test_kontrolcunun_kendi_reddi_kendisine_bilgi_olarak_gelmez(self):
        talep = self.talep_olustur("PARAF_BEKLIYOR")
        talep.kontrol_tarihi = timezone.now()
        talep.save(update_fields=["kontrol_tarihi"])
        self.reddet(talep, self.kontrol)
        self.assertEqual(list(self.dashboard(self.kontrol).context["reddedilenler"]), [])

    def test_get_istegi_bildirimi_kapatmaz(self):
        talep = self.amir_reddi_olustur()
        self.client.force_login(self.kontrol)
        self.client.get(reverse("izin_bildirimini_kapat", args=[talep.pk]))
        self.client.get(reverse("izin_bildirimlerini_kapat"))
        self.assertFalse(IzinBildirimKapatma.objects.exists())

    def test_kapatilmis_ret_bilgisi_yeni_onay_bildirimini_gizlemez(self):
        talep = self.amir_reddi_olustur()
        IzinBildirimKapatma.objects.create(kullanici=self.kontrol, talep=talep)
        form = Mock(cleaned_data={})
        form.save.return_value = talep
        personel = PortalPersonel.objects.get(kullanici=self.sahip)
        with patch("portal.views.personel_aktif_mi", return_value=True), \
             patch("portal.views.izin_bakiyesi_onaya_sunma_kontrolu", return_value=None), \
             patch("portal.views.kontrol_edeni_bul", return_value=self.kontrol):
            _, hata = views._talebi_kaydet(talep, form, personel=personel, kullanici=self.sahip,
                                          evrak_no=talep.evrak_no, onaya_sun=True)
        self.assertIsNone(hata)
        self.assertFalse(IzinBildirimKapatma.objects.filter(talep=talep).exists())
        talep.durum = "ONAYLANDI"
        talep.kimin_uzerinde = None
        talep.save(update_fields=["durum", "kimin_uzerinde"])
        self.assertIn(talep, self.dashboard(self.kontrol).context["imzalananlar"])
