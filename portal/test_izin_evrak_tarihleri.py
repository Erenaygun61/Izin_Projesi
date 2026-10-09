from datetime import date
from decimal import Decimal
from html.parser import HTMLParser
from unittest.mock import patch

from django.apps import apps
from django.contrib.auth.models import User
from django.db import connection
from django.test import TestCase
from django.urls import reverse
from django.utils.translation import override

from .forms import DateInput
from .models import (HakEdilenIzin, IzinTalebi, IzinTuru, Personel, PortalAyar,
                     PortalPersonel, YonetimYetkisi)


class InputValues(HTMLParser):
    def __init__(self, html):
        super().__init__()
        self.values = {}
        self.feed(html)

    def handle_starttag(self, tag, attrs):
        if tag == "input":
            attrs = dict(attrs)
            self.values[attrs.get("name")] = attrs.get("value", "")


class IzinEvrakTarihleriTests(TestCase):
    tarihler = {
        "baslangic_tarihi": date(2026, 8, 10),
        "bitis_tarihi": date(2026, 8, 14),
        "goreve_baslama_tarihi": date(2026, 8, 17),
    }

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
        cls.sahip = User.objects.create_user("date_owner", first_name="Test", last_name="Sahip")
        cls.kontrol = User.objects.create_user("date_control")
        cls.imzaci = User.objects.create_user("date_approver")
        kaynak = Personel.objects.create(id=101, sicil_no="101", ad="Test", soyad="Sahip", tc_no="12345678901",
                                        durum="Aktif", gorev_birimi="Test Birimi", cep_tel="05551111111")
        cls.personel = PortalPersonel.objects.create(source_personel_id=kaynak.pk, personel_tipi="MEMUR",
                                                     kullanici=cls.sahip, sicil_no="101", ad="Test", soyad="Sahip",
                                                     tc_no=kaynak.tc_no, durum="Aktif", gorev_birimi="Test Birimi")
        cls.tur = IzinTuru.objects.create(kod="TEST-YILLIK", ad="Test Yillik Izin", yillik_izin_mi=True, aktif_mi=True)
        for yil, gun in ((2025, 2), (2026, 29)):
            HakEdilenIzin.objects.create(personel=kaynak, personel_tipi="MEMUR", izin_yili=yil, hak_edilen_gun=gun,
                                        gecerlilik_baslangic_tarihi=date(yil, 1, 1),
                                        gecerlilik_bitis_tarihi=date(yil + 2, 1, 1))
        YonetimYetkisi.objects.create(kullanici=cls.kontrol, aktif_mi=True, memur_kontrol_eden=True)
        PortalAyar.objects.create(id=1, kontrol_eden_memur=cls.kontrol)
        cls.talep = IzinTalebi.objects.create(kullanici=cls.sahip, personel_id=kaynak.pk, personel_tipi="MEMUR",
                                             izin_turu=cls.tur, evrak_no="TEST-TARIH", durum="REDDEDILDI",
                                             kontrol_eden=cls.kontrol, cep_telefonu="05550000000",
                                             ait_oldugu_yil="2025-2026", izin_hakki=31, kullanilan_gun=5, kalan_izin=26,
                                             red_nedeni="Test ret nedeni", reddeden=cls.imzaci, **cls.tarihler)

    def setUp(self):
        self.client.force_login(self.sahip)

    def url(self):
        return reverse("izin_talebi_duzenle", args=[self.talep.pk])

    def get_form(self):
        response = self.client.get(self.url())
        self.assertEqual(response.status_code, 200)
        return response.context["form"]

    def browser_values(self, form):
        data = {}
        for name, field in form.fields.items():
            if field.disabled:
                continue
            if field.widget.input_type == "date":
                data[name] = InputValues(form[name].as_widget()).values[name]
            else:
                value = form[name].value()
                data[name] = "" if value is None else str(value)
        return data

    def assert_saved_values(self):
        self.talep.refresh_from_db()
        for alan, tarih in self.tarihler.items():
            self.assertEqual(getattr(self.talep, alan), tarih)
        self.assertEqual(self.talep.cep_telefonu, "05550000000")
        self.assertEqual(self.talep.izin_turu_id, self.tur.pk)
        self.assertEqual(self.talep.ait_oldugu_yil, "2025-2026")
        self.assertEqual((self.talep.izin_hakki, self.talep.kullanilan_gun, self.talep.kalan_izin),
                         (Decimal("31"), Decimal("5"), Decimal("26")))
        self.assertEqual(self.talep.evrak_no, "TEST-TARIH")

    def test_tarih_widgeti_turkce_ortamda_tarayici_icin_iso_yazar(self):
        with override("tr"):
            html = DateInput().render("date", date(2026, 8, 10))
        self.assertEqual(InputValues(html).values["date"], "2026-08-10")

    def test_gercekten_bos_tarih_kendiliginden_doldurulmaz(self):
        self.assertEqual(InputValues(DateInput().render("date", None)).values["date"], "")

    def test_red_iade_ve_taslak_dolu_bilgilerle_acilir(self):
        for durum in ("REDDEDILDI", "IADE_EDILDI", "TASLAK"):
            with self.subTest(durum=durum):
                IzinTalebi.objects.filter(pk=self.talep.pk).update(durum=durum)
                with override("tr"):
                    form = self.get_form()
                values = self.browser_values(form)
                for alan, tarih in self.tarihler.items():
                    self.assertEqual(values[alan], tarih.isoformat())
                self.assertEqual(values["cep_telefonu"], "05550000000")
                self.assertEqual(values["ait_oldugu_yil"], "2025-2026")
                self.assert_saved_values()

    def test_hic_degistirmeden_taslak_kaydedilebilir(self):
        data = self.browser_values(self.get_form())
        data["islem"] = "taslak_kaydet"
        response = self.client.post(self.url(), data)
        self.assertRedirects(response, reverse("dashboard"), fetch_redirect_response=False)
        self.assert_saved_values()
        self.assertEqual(self.talep.durum, "TASLAK")

    def test_hic_degistirmeden_yeniden_onaya_sunulabilir(self):
        data = self.browser_values(self.get_form())
        data["islem"] = "onaya_sun"
        response = self.client.post(self.url(), data)
        self.assertRedirects(response, reverse("evrak_paraf", args=[self.talep.pk]), fetch_redirect_response=False)
        self.assert_saved_values()
        self.assertEqual(self.talep.durum, "PARAF_BEKLIYOR")
        self.assertEqual(self.talep.kimin_uzerinde_id, self.sahip.pk)
        self.assertEqual(IzinTalebi.objects.count(), 1)

    def test_kisi_isterse_bir_tarihi_degistirebilir(self):
        data = self.browser_values(self.get_form())
        data.update(islem="taslak_kaydet", goreve_baslama_tarihi="2026-08-18")
        response = self.client.post(self.url(), data)
        self.assertEqual(response.status_code, 302)
        self.talep.refresh_from_db()
        self.assertEqual(self.talep.goreve_baslama_tarihi, date(2026, 8, 18))
        self.assertEqual(self.talep.baslangic_tarihi, self.tarihler["baslangic_tarihi"])
        self.assertEqual(self.talep.bitis_tarihi, self.tarihler["bitis_tarihi"])

    def test_dogrulama_hatasinda_kayitli_tarihler_silinmez(self):
        data = self.browser_values(self.get_form())
        data.update(islem="onaya_sun", goreve_baslama_tarihi="")
        response = self.client.post(self.url(), data)
        self.assertEqual(response.status_code, 200)
        self.assertIn("goreve_baslama_tarihi", response.context["form"].errors)
        self.assert_saved_values()
        self.assertEqual(self.talep.durum, "REDDEDILDI")

    def test_reddetme_onceki_izin_bilgilerini_degistirmez(self):
        self.talep.durum = "IMZA_SURECINDE"
        self.talep.kimin_uzerinde = self.imzaci
        self.talep.save(update_fields=["durum", "kimin_uzerinde"])
        self.client.force_login(self.imzaci)
        with patch("portal.views.talep_aktif_imzacilari", return_value=[("Amir", self.imzaci)]):
            response = self.client.post(reverse("evrak_paraf", args=[self.talep.pk]),
                                        {"islem": "reddet", "red_nedeni": "Test second review"})
        self.assertEqual(response.status_code, 302)
        self.assert_saved_values()
        self.assertEqual(self.talep.durum, "REDDEDILDI")
