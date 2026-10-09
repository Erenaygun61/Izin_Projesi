from datetime import date
from decimal import Decimal
from io import BytesIO, StringIO
from unittest.mock import patch

from django.apps import apps
from django.contrib.auth.models import User
from django.db import connection
from django.core.management import call_command
from django.test import TestCase
from django.urls import reverse
from openpyxl import load_workbook
try:
    from pypdf import PdfReader
except ImportError:
    PdfReader = None

from . import test_izin_evrak_tarihleri as tarih_tests
from . import views
from .document_exports import izin_excel_olustur, izin_pdf_olustur
from .evrak_verisi import belge_verileri, evraki_sabitle
from .models import (HakEdilenIzin, IzinTalebi, IzinTuru, IsciPersonel, KullanilanIzin,
                     Personel, PortalIzinHakki, PortalKullanilanIzin, PortalPersonel, SirketPersonel, IzinIslemGecmisi)


class IzinTuruVeSabitEvrakTests(TestCase):
    tarihler = tarih_tests.IzinEvrakTarihleriTests.tarihler
    url = tarih_tests.IzinEvrakTarihleriTests.url
    get_form = tarih_tests.IzinEvrakTarihleriTests.get_form
    browser_values = tarih_tests.IzinEvrakTarihleriTests.browser_values

    @classmethod
    def setUpClass(cls):
        mevcut = set(connection.introspection.table_names())
        with connection.schema_editor() as editor:
            for model in apps.get_app_config('portal').get_models():
                if not model._meta.managed and model._meta.db_table not in mevcut:
                    editor.create_model(model)
                    mevcut.add(model._meta.db_table)
        super().setUpClass()

    @classmethod
    def setUpTestData(cls):
        tarih_tests.IzinEvrakTarihleriTests.setUpTestData.__func__(cls)
        cls.mesai = IzinTuru.objects.create(kod='TEST-MESAI', ad='Mesai İzni', aktif_mi=True)
        cls.diger = IzinTuru.objects.create(kod='TEST-DIGER', ad='Diğer İzin', aktif_mi=True)
        cls.personel.unvan = 'Eski Unvan'
        cls.personel.save()

    def setUp(self):
        self.client.force_login(self.sahip)

    def mesai_hakki(self, personel=None, tip='MEMUR', gun=5, yil=2026):
        alan = {'MEMUR': 'personel_id', 'ISCI': 'isci_personel_id', 'SIRKET': 'sirket_personel_id'}[tip]
        return HakEdilenIzin.objects.create(**{alan: personel.pk if personel else self.talep.personel_id},
                                          personel_tipi=tip, izin_turu=self.mesai, izin_yili=yil,
                                          hak_edilen_gun=gun, gecerlilik_baslangic_tarihi=date(yil, 1, 1),
                                          gecerlilik_bitis_tarihi=date(yil + 2, 1, 1))

    def mesai_talebi(self, gun=3):
        self.talep.izin_turu = self.mesai
        self.talep.ait_oldugu_yil = '2026'
        self.talep.kullanilan_gun = gun
        self.talep.izin_hakki = None
        self.talep.kalan_izin = None
        self.talep.save()
        return self.talep

    def formu_gonder(self, tur, gun, hak='', islem='onaya_sun'):
        data = self.browser_values(self.get_form())
        data.update(izin_turu=str(tur.pk), kullanilan_gun=str(gun), izin_hakki=str(hak), ait_oldugu_yil='2026', islem=islem)
        return self.client.post(self.url(), data)

    def test_tur_bakiyeleri_ayridir(self):
        self.mesai_hakki()
        self.assertEqual(views.izin_ozeti(self.personel)[1], Decimal('31'))
        self.assertEqual(views.izin_ozeti(self.personel, self.mesai)[1], Decimal('5'))
        form = self.get_form()
        self.assertEqual(form.tur_bakiyeleri[str(self.mesai.pk)][0]['kalan'], '5.00')
        self.assertEqual(form.tur_bakiyeleri[str(self.diger.pk)], [])

    def test_mesai_tanimliysa_sunucudaki_bakiye_kullanilir(self):
        self.mesai_hakki()
        self.assertEqual(self.formu_gonder(self.mesai, 3, 999).status_code, 302)
        self.talep.refresh_from_db()
        self.assertEqual((self.talep.izin_hakki, self.talep.kalan_izin), (Decimal('5'), Decimal('2')))
        self.assertEqual(KullanilanIzin.objects.count(), 0)

    def test_tanimli_bakiyeyi_asmak_elle_hak_yazarak_mumkun_degil(self):
        self.mesai_hakki()
        response = self.formu_gonder(self.mesai, 6, 999)
        self.assertEqual(response.status_code, 200)
        self.assertIn('kullanilan_gun', response.context['form'].errors)
        self.talep.refresh_from_db()
        self.assertEqual(self.talep.durum, 'REDDEDILDI')

    def test_sifir_bakiye_hak_yok_gibi_atlanamaz(self):
        hak = self.mesai_hakki()
        KullanilanIzin.objects.create(personel_id=self.talep.personel_id, izin_turu=self.mesai,
                                     hak_edilen_izin=hak, kullanilan_gun=5, izin_baslangic_tarihi=date(2026, 1, 1), izin_bitis_tarihi=date(2026, 1, 5))
        self.assertEqual(self.formu_gonder(self.mesai, 1, 999).status_code, 200)

    def test_tanimsiz_mesai_elle_gunle_gonderilir(self):
        self.assertEqual(self.formu_gonder(self.mesai, 2).status_code, 302)
        self.talep.refresh_from_db()
        self.assertIsNone(self.talep.izin_hakki)
        self.assertIsNone(self.talep.kalan_izin)
        self.assertEqual(self.talep.kullanilan_gun, Decimal('2'))
        self.assertEqual(KullanilanIzin.objects.count(), 0)

    def test_tam_imza_akisi_mesaiyi_sadece_son_imzada_duser(self):
        self.mesai_hakki()
        self.assertEqual(self.formu_gonder(self.mesai, 3).status_code, 302)
        url = reverse('evrak_paraf', args=[self.talep.pk])
        with patch('portal.views.imzacilari_bul', return_value=[('Amir', self.imzaci)]):
            self.assertEqual(self.client.post(url, {'islem': 'parafla'}).status_code, 302)
        self.assertEqual(KullanilanIzin.objects.count(), 0)
        self.client.force_login(self.kontrol)
        self.assertEqual(self.client.post(url, {'islem': 'parafla'}).status_code, 302)
        self.assertEqual(KullanilanIzin.objects.count(), 0)
        self.client.force_login(self.imzaci)
        self.assertEqual(self.client.post(url, {'islem': 'imzala'}).status_code, 302)
        self.talep.refresh_from_db()
        self.assertEqual(self.talep.durum, 'ONAYLANDI')
        self.assertTrue(self.talep.ozluk_aktarildi_mi)
        self.assertEqual(views.izin_ozeti(self.personel, self.mesai)[1], Decimal('2'))
        self.assertEqual(views.izin_ozeti(self.personel)[1], Decimal('31'))

    def test_diger_tur_yillik_izni_tuketmez(self):
        self.assertEqual(self.formu_gonder(self.diger, 2).status_code, 302)
        self.talep.refresh_from_db()
        views.izni_ozluge_aktar(self.talep, self.personel, self.imzaci)
        self.assertIsNone(KullanilanIzin.objects.get().hak_edilen_izin_id)
        self.assertEqual(views.izin_ozeti(self.personel)[1], Decimal('31'))

    def test_uc_personel_tipinde_bes_gunden_uc_gun_duser(self):
        for tip, model in [('MEMUR', Personel), ('ISCI', IsciPersonel), ('SIRKET', SirketPersonel)]:
            with self.subTest(tip=tip):
                kaynak = model.objects.create(sicil_no=f'NEW-{tip}', tc_no='22222222222', ad='Test', soyad=tip, durum='Aktif')
                hak = self.mesai_hakki(kaynak, tip)
                talep = IzinTalebi.objects.create(kullanici=self.sahip, personel_id=kaynak.pk, personel_tipi=tip,
                                                 izin_turu=self.mesai, kullanilan_gun=3, ait_oldugu_yil='2026',
                                                 durum='ONAYLANDI', **self.tarihler)
                views.izni_ozluge_aktar(talep, kaynak, self.imzaci)
                self.assertEqual(KullanilanIzin.objects.get(hak_edilen_izin=hak).kullanilan_gun, Decimal('3'))
                self.assertEqual(views.izin_ozeti(kaynak, self.mesai)[1], Decimal('2'))

    def test_tanimsiz_mesai_sonradan_hak_eklense_de_oradan_dusmez(self):
        talep = self.mesai_talebi(2)
        talep.durum = 'PARAF_BEKLIYOR'
        talep.save()
        evraki_sabitle(talep, self.personel, [('Amir', self.imzaci)])
        self.mesai_hakki()
        views.izni_ozluge_aktar(talep, self.personel, self.imzaci)
        self.assertIsNone(KullanilanIzin.objects.get().hak_edilen_izin_id)
        self.assertEqual(views.izin_ozeti(self.personel, self.mesai)[1], Decimal('5'))
        self.assertEqual(views.izin_ozeti(self.personel)[1], Decimal('31'))

    def test_ozluk_aktarimi_tekrarlanmaz(self):
        self.mesai_hakki()
        talep = self.mesai_talebi()
        stale = IzinTalebi.objects.get(pk=talep.pk)
        views.izni_ozluge_aktar(talep, self.personel, self.imzaci)
        views.izni_ozluge_aktar(stale, self.personel, self.imzaci)
        self.assertEqual(KullanilanIzin.objects.count(), 1)
        self.assertEqual(views.izin_ozeti(self.personel, self.mesai)[1], Decimal('2'))

    def test_bakiye_onay_arasinda_biterse_transfer_geri_alinir(self):
        hak = self.mesai_hakki()
        talep = self.mesai_talebi(6)
        with self.assertRaises(ValueError):
            views.izni_ozluge_aktar(talep, self.personel, self.imzaci)
        talep.refresh_from_db()
        self.assertFalse(talep.ozluk_aktarildi_mi)
        self.assertEqual(KullanilanIzin.objects.count(), 0)

    def test_sirket_portal_yillik_hakki_da_duser(self):
        kaynak = SirketPersonel.objects.create(sicil_no='COMPANY', tc_no='33333333333', ad='Test', soyad='Company', durum='Aktif')
        kart = PortalPersonel.objects.create(source_personel_id=kaynak.pk, personel_tipi='SIRKET', sicil_no='COMPANY', tc_no=kaynak.tc_no, ad='Test', soyad='Company')
        PortalIzinHakki.objects.create(personel=kart, izin_yili=2026, hak_edilen_gun=10,
                                      gecerlilik_baslangic_tarihi=date(2026, 1, 1), gecerlilik_bitis_tarihi=date(2028, 1, 1))
        talep = IzinTalebi.objects.create(kullanici=self.sahip, personel_id=kaynak.pk, personel_tipi='SIRKET', izin_turu=self.tur,
                                         kullanilan_gun=3, ait_oldugu_yil='2026', **self.tarihler)
        views.izni_ozluge_aktar(talep, kart, self.imzaci)
        self.assertEqual(views.izin_ozeti(kart)[1], Decimal('7'))
        self.assertEqual(PortalKullanilanIzin.objects.count(), 1)
        self.assertEqual(KullanilanIzin.objects.filter(sirket_personel=kaynak).count(), 1)

    def sabit_talep(self):
        self.talep.durum = 'PARAF_BEKLIYOR'
        self.talep.kimin_uzerinde = self.sahip
        self.talep.save()
        evraki_sabitle(self.talep, self.personel, [('Eski Amir Rolü', self.imzaci)])
        self.personel.unvan = 'Yeni Unvan'
        self.personel.gorev_birimi = 'Yeni Birim'
        self.personel.save()
        self.imzaci.first_name = 'Yeni Amir Adi'
        self.imzaci.save()
        self.tur.ad = 'Yeni İzin Adı'
        self.tur.save()
        self.talep.refresh_from_db()
        return self.talep

    def test_sabit_belge_unvan_birim_imzaci_ve_tur_degisikliginden_etkilenmez(self):
        talep = self.sabit_talep()
        for durum in ('PARAF_BEKLIYOR', 'IMZA_SURECINDE', 'ONAYLANDI'):
            talep.durum = durum
            belge, kisi, imzacilar = belge_verileri(talep, self.personel, [('Yeni Amir', self.sahip)])
            self.assertEqual(kisi.unvani, 'Eski Unvan')
            self.assertEqual(kisi.gorev_birimi, 'Test Birimi')
            self.assertEqual(imzacilar[0][0], 'Eski Amir Rolü')
            self.assertEqual(imzacilar[0][1].get_full_name(), 'date_approver')
            self.assertEqual(belge.izin_turu.ad, 'Test Yillik Izin')
            self.assertEqual(views.talep_aktif_imzacilari(talep)[0][1].pk, self.imzaci.pk)

    def test_ilk_paraf_bilgileri_sabitler_ve_ozluge_islemez(self):
        self.talep.durum = 'PARAF_BEKLIYOR'
        self.talep.kimin_uzerinde = self.sahip
        self.talep.save()
        with patch('portal.views.imzacilari_bul', return_value=[('Amir', self.imzaci)]):
            response = self.client.post(reverse('evrak_paraf', args=[self.talep.pk]), {'islem': 'parafla'})
        self.assertEqual(response.status_code, 302)
        self.talep.refresh_from_db()
        self.assertEqual(self.talep.evrak_verisi['personel']['unvani'], 'Eski Unvan')
        self.assertEqual(self.talep.kimin_uzerinde_id, self.kontrol.pk)
        self.assertEqual(KullanilanIzin.objects.count(), 0)

    def test_red_iade_edilen_belge_guncel_bilgilerle_duzenlenir(self):
        talep = self.sabit_talep()
        for durum in ('REDDEDILDI', 'IADE_EDILDI', 'TASLAK'):
            talep.durum = durum
            belge, kisi, imzacilar = belge_verileri(talep, self.personel, [('Yeni Amir', self.sahip)])
            self.assertEqual(kisi.unvan, 'Yeni Unvan')
            self.assertEqual(imzacilar[0][0], 'Yeni Amir')

    def test_yeniden_gonderilince_eski_sabitleme_temizlenir(self):
        talep = self.sabit_talep()
        talep.durum = 'REDDEDILDI'
        talep.save()
        response = self.formu_gonder(self.tur, 5)
        self.assertEqual(response.status_code, 302)
        talep.refresh_from_db()
        self.assertEqual(talep.evrak_verisi, {})
        evraki_sabitle(talep, self.personel, [('Yeni Amir', self.sahip)])
        self.assertEqual(talep.evrak_verisi['personel']['unvani'], 'Yeni Unvan')

    def test_onayli_belgenin_ekrani_ve_exceli_sabit_kalir(self):
        talep = self.sabit_talep()
        talep.durum = 'ONAYLANDI'
        talep.save()
        response = self.client.get(reverse('evrak_onizleme', args=[talep.pk]))
        self.assertContains(response, 'Eski Unvan')
        self.assertNotContains(response, 'Yeni Unvan')
        content = izin_excel_olustur(talep, self.personel, [], talep.goreve_baslama_tarihi)
        workbook = load_workbook(BytesIO(content))
        text = '\n'.join(str(cell.value or '') for row in workbook.active for cell in row)
        self.assertIn('Eski Unvan', text)
        self.assertNotIn('Yeni Unvan', text)
        pdf = izin_pdf_olustur(talep, self.personel, [], talep.goreve_baslama_tarihi)
        self.assertTrue(pdf.startswith(b'%PDF'))
        if PdfReader is not None:
            pdf_text = '\n'.join(page.extract_text() for page in PdfReader(BytesIO(pdf)).pages)
            self.assertIn('Eski Unvan', pdf_text)
            self.assertNotIn('Yeni Unvan', pdf_text)

    def test_imza_sonrasi_duzenleme_postu_engellenir(self):
        talep = self.sabit_talep()
        talep.durum = 'ONAYLANDI'
        talep.save()
        response = self.client.post(self.url(), {'islem': 'taslak_kaydet', 'kullanilan_gun': '999'})
        self.assertEqual(response.status_code, 302)
        talep.refresh_from_db()
        self.assertEqual(talep.kullanilan_gun, Decimal('5'))

    def test_eski_evrak_sabitleme_gercek_imzaciyi_korur_ve_tekrarlanabilir(self):
        self.talep.durum = 'ONAYLANDI'
        self.talep.save()
        IzinIslemGecmisi.objects.create(talep=self.talep, kullanici=self.sahip, islem='IMZA', rol='Teklif Eden')
        IzinIslemGecmisi.objects.create(talep=self.talep, kullanici=self.imzaci, islem='ONAY', rol='Gerçek Amir')
        with patch('portal.views.imzacilari_bul', return_value=[('Yeni Amir', self.sahip)]):
            call_command('evraklari_sabitle', stdout=StringIO())
        self.talep.refresh_from_db()
        saved = self.talep.evrak_verisi
        self.assertEqual(saved['imzacilar'][0]['kullanici']['id'], self.imzaci.pk)
        self.assertEqual(saved['imzacilar'][0]['rol'], 'Gerçek Amir')
        call_command('evraklari_sabitle', stdout=StringIO())
        self.talep.refresh_from_db()
        self.assertEqual(self.talep.evrak_verisi, saved)

    def test_henuz_ilk_paraf_verilmemis_eski_evrak_sabitlenmez(self):
        self.talep.durum = 'PARAF_BEKLIYOR'
        self.talep.kimin_uzerinde = self.sahip
        self.talep.save()
        call_command('evraklari_sabitle', stdout=StringIO())
        self.talep.refresh_from_db()
        self.assertEqual(self.talep.evrak_verisi, {})
