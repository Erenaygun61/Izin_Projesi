import json
from datetime import date, timedelta
from decimal import Decimal
from unittest.mock import patch

from django.apps import apps
from django.contrib.auth.models import User
from django.core.files.uploadedfile import SimpleUploadedFile
from django.db import connection, OperationalError
from django.test import TestCase, override_settings
from django.urls import reverse
from django.utils import timezone

from . import views
from .evrak_verisi import evraki_sabitle
from .models import (GorevBirimiImzaci, HakEdilenIzin, ImzaciVekaleti, IzinDenetimKaydi,
                     IzinIslemGecmisi, IzinTalebi, KullanilanIzin, Personel, YonetimYetkisi)
from .test_izin_turu_ve_sabit_evrak import IzinTuruVeSabitEvrakTests
from .yetki_yonetimi import yetki_katalogu, yetki_surumu


@override_settings(PASSWORD_HASHERS=['django.contrib.auth.hashers.MD5PasswordHasher'])
class IzinYetkiAkisTests(TestCase):
    @classmethod
    def setUpClass(cls):
        mevcut = set(connection.introspection.table_names())
        with connection.schema_editor() as editor:
            for model in apps.get_app_config('portal').get_models():
                if not model._meta.managed and model._meta.db_table not in mevcut:
                    editor.create_model(model)
                    mevcut.add(model._meta.db_table)
        super().setUpClass()
    setUpTestData = classmethod(IzinTuruVeSabitEvrakTests.setUpTestData.__func__)
    browser_values = IzinTuruVeSabitEvrakTests.browser_values
    mesai_hakki = IzinTuruVeSabitEvrakTests.mesai_hakki
    tarihler = IzinTuruVeSabitEvrakTests.tarihler

    def setUp(self):
        self.admin = User.objects.create_superuser('workflow_admin', password='test-only')
        self.yabanci = User.objects.create_user('workflow_outsider', is_staff=True)
        self.son_imzaci = User.objects.create_user('workflow_final')
        self.personel.alt_birimi = 'Test Sube'
        self.personel.save()
        GorevBirimiImzaci.objects.create(birim_adi='Test Birimi', alt_birim_adi='Test Sube',
                                        sube_muduru=self.imzaci, amir_kullanici=self.imzaci)
        GorevBirimiImzaci.objects.create(birim_adi='Test Birimi', daire_baskani=self.son_imzaci,
                                        amir_kullanici=self.son_imzaci)

    def evrak(self):
        self.talep.refresh_from_db()
        return self.talep

    def gonder(self):
        self.client.force_login(self.sahip)
        url = reverse('izin_talebi_duzenle', args=[self.talep.pk])
        response = self.client.get(url)
        data = self.browser_values(response.context['form'])
        data['islem'] = 'onaya_sun'
        response = self.client.post(url, data)
        self.assertEqual(response.status_code, 302)
        self.assertEqual(self.evrak().durum, 'PARAF_BEKLIYOR')

    def isle(self, user, islem='parafla', **params):
        self.client.force_login(user)
        params['islem'] = islem
        response = self.client.post(reverse('evrak_paraf', args=[self.talep.pk]), params)
        self.assertEqual(response.status_code, 302)
        return self.evrak()

    def kontrol_asamasina(self):
        self.gonder()
        self.isle(self.sahip)
        self.assertEqual(self.talep.kimin_uzerinde_id, self.kontrol.pk)

    def imza_asamasina(self):
        self.kontrol_asamasina()
        self.isle(self.kontrol)
        self.assertEqual(self.talep.kimin_uzerinde_id, self.imzaci.pk)

    def yetki_post(self, users, secili=(), islem='ekle', **extra):
        self.client.force_login(self.admin)
        versions = {str(user.pk): yetki_surumu(YonetimYetkisi.objects.filter(kullanici=user).first()) for user in users}
        data = {'kullanicilar': [str(user.pk) for user in users], 'yetkiler': list(secili),
                'islem': islem, 'surumler': json.dumps(versions)}
        data.update(extra)
        return self.client.post(reverse('kullanici_yetkileri'), data)

    def test_anonim_kullanici_yetki_veremez(self):
        response = self.client.post(reverse('kullanici_yetkileri'), {'islem':'yetki_kaydet', 'kullanici':self.yabanci.pk, 'tam_yonetim_yetkisi':'on'})
        self.assertEqual(response.status_code, 302)
        self.assertFalse(YonetimYetkisi.objects.filter(kullanici=self.yabanci).exists())

    def test_staff_yetki_ekranina_giremez_ve_kendine_yetki_veremez(self):
        self.client.force_login(self.yabanci)
        for method in ('get', 'post'):
            response = getattr(self.client, method)(reverse('kullanici_yetkileri'))
            self.assertEqual(response.status_code, 302)
        self.assertFalse(views._tam_yonetim_yetkisi(self.yabanci))
        self.assertFalse(views._kontrol_yetkisi_var(self.yabanci, 'MEMUR'))

    def test_staff_ayri_personel_tipi_yetkisini_asamaz(self):
        YonetimYetkisi.objects.create(kullanici=self.yabanci, isci_sahibine_iade=True)
        self.imza_asamasina()
        self.client.force_login(self.yabanci)
        self.client.post(reverse('izin_ara'), {'islem':'sahibine_iade', 'talep_id':self.talep.pk})
        self.assertEqual(self.evrak().durum, 'IMZA_SURECINDE')

    def test_pasif_superuser_yetkisi_yok(self):
        self.admin.is_active = False
        self.assertFalse(views._tam_yonetim_yetkisi(self.admin))
        self.assertFalse(views._tip_yetkisi_var(self.admin, 'MEMUR', 'izin_silme'))
        self.assertFalse(views._evrak_yonetim_yetkisi(self.admin))

    def test_tam_yonetim_kontrolcu_yetkisi_vermez(self):
        YonetimYetkisi.objects.create(kullanici=self.yabanci, tam_yonetim_yetkisi=True)
        self.assertEqual(views._tip_yetkili_tipler(self.yabanci, 'kontrol_eden'), [])
        self.assertFalse(views._kontrol_yetkisi_var(self.admin, 'MEMUR'))
        self.assertEqual(views.kontrol_edeni_bul('MEMUR'), self.kontrol)

    def test_dinamik_katalog_modeldeki_tum_yetkileri_icerir(self):
        actual = {item['alan'] for grup in yetki_katalogu() for item in grup['alanlar']}
        self.assertEqual(actual, set(views.YONETIM_YETKI_ALANLARI))

    def test_toplu_ekleme_diger_yetkileri_korur_ve_denetlenir(self):
        YonetimYetkisi.objects.create(kullanici=self.yabanci, isci_izin_arama=True)
        self.yetki_post([self.yabanci,self.imzaci], ['memur_izin_arama'])
        self.assertTrue(YonetimYetkisi.objects.get(kullanici=self.yabanci).isci_izin_arama)
        self.assertEqual(YonetimYetkisi.objects.filter(kullanici__in=[self.yabanci,self.imzaci], memur_izin_arama=True).count(), 2)
        self.assertEqual(IzinDenetimKaydi.objects.filter(kullanici=self.admin, aciklama__contains='memur_izin_arama').count(), 2)

    def test_toplu_kaldirma_diger_yetkileri_korur(self):
        YonetimYetkisi.objects.create(kullanici=self.yabanci, memur_izin_arama=True, isci_izin_arama=True)
        self.yetki_post([self.yabanci], ['memur_izin_arama'], 'kaldir')
        yetki = YonetimYetkisi.objects.get(kullanici=self.yabanci)
        self.assertFalse(yetki.memur_izin_arama)
        self.assertTrue(yetki.isci_izin_arama)

    def test_tum_yetkileri_kaldirma(self):
        YonetimYetkisi.objects.create(kullanici=self.yabanci, tam_yonetim_yetkisi=True)
        self.yetki_post([self.yabanci], islem='tumunu_kaldir')
        self.assertFalse(YonetimYetkisi.objects.get(kullanici=self.yabanci).aktif_mi)

    def test_eski_surum_toplu_islemin_tamamini_engeller(self):
        self.yetki_post([self.yabanci,self.imzaci], ['memur_izin_arama'], surumler=json.dumps({str(self.yabanci.pk):'yok', str(self.imzaci.pk):'eski'}))
        self.assertFalse(YonetimYetkisi.objects.filter(kullanici__in=[self.yabanci,self.imzaci]).exists())

    def test_birden_cok_kontrolcu_toplu_atanamaz(self):
        self.yetki_post([self.yabanci,self.imzaci], ['memur_kontrol_eden'])
        self.assertEqual(views.kontrol_edeni_bul('MEMUR'), self.kontrol)
        self.assertFalse(YonetimYetkisi.objects.filter(kullanici__in=[self.yabanci,self.imzaci]).exists())

    def test_kontrolcu_devri_onaysiz_uygulanmaz(self):
        self.yetki_post([self.yabanci], ['memur_kontrol_eden'])
        self.assertEqual(views.kontrol_edeni_bul('MEMUR'), self.kontrol)

    def test_kontrolcu_devrinde_bekleyen_evrak_aktarilir(self):
        self.kontrol_asamasina()
        self.yetki_post([self.yabanci], ['memur_kontrol_eden'], kontrol_devri_onay='on')
        self.assertEqual(views.kontrol_edeni_bul('MEMUR'), self.yabanci)
        self.assertEqual(self.evrak().durum, 'IADE_EDILDI')
        self.assertEqual(self.talep.kimin_uzerinde_id, self.sahip.pk)
        self.gonder()
        self.isle(self.sahip)
        self.assertEqual(self.talep.kimin_uzerinde_id, self.yabanci.pk)
        self.assertFalse(views._talep_aktif_gorev_yetkisi(self.kontrol, self.talep))

    def test_pasif_yetki_kaydi_aktif_kontrolcuyu_silemez(self):
        YonetimYetkisi.objects.create(kullanici=self.yabanci, aktif_mi=False, memur_kontrol_eden=True)
        self.assertEqual(views.kontrol_edeni_bul('MEMUR'), self.kontrol)

    def test_isci_kontrol_devri_memuru_degistirmez(self):
        self.yetki_post([self.yabanci], ['isci_kontrol_eden'])
        self.assertEqual(views.kontrol_edeni_bul('MEMUR'), self.kontrol)
        self.assertEqual(views.kontrol_edeni_bul('ISCI'), self.yabanci)

    def test_dar_yetkili_yonetim_merkezine_erisebilir(self):
        YonetimYetkisi.objects.create(kullanici=self.yabanci, vekalet_yonetimi=True)
        self.client.force_login(self.yabanci)
        self.assertEqual(self.client.get(reverse('yonetim')).status_code, 200)
        self.assertEqual(self.client.get(reverse('kullanici_yetkisi_detay'), {'kullanici':self.sahip.pk}).status_code, 302)

    def test_kullanici_ara_kisisel_hesaba_acik_degil(self):
        self.client.force_login(self.sahip)
        self.assertEqual(self.client.get(reverse('kullanici_ara')).status_code, 403)

    def test_tam_akis_son_imzaya_kadar_bakiye_dusurmez(self):
        self.imza_asamasina()
        self.assertEqual(KullanilanIzin.objects.count(), 0)
        self.isle(self.imzaci)
        self.assertEqual(self.talep.kimin_uzerinde_id, self.son_imzaci.pk)
        self.assertEqual(KullanilanIzin.objects.count(), 0)
        self.isle(self.son_imzaci, 'imzala')
        self.assertEqual(self.talep.durum, 'ONAYLANDI')
        self.assertTrue(self.talep.ozluk_aktarildi_mi)
        self.assertEqual(sum(KullanilanIzin.objects.values_list('kullanilan_gun',flat=True)), Decimal('5'))
        self.assertEqual(views.izin_ozeti(self.personel)[1], Decimal('26'))

    def test_yanlis_kisi_ve_yonetici_baskasi_yerine_imzalayamaz(self):
        self.imza_asamasina()
        for user in [self.admin,self.yabanci,self.son_imzaci,self.sahip]:
            with self.subTest(user=user.username):
                self.isle(user)
                self.assertEqual(self.talep.kimin_uzerinde_id, self.imzaci.pk)
                self.assertEqual(KullanilanIzin.objects.count(), 0)

    def test_eski_ekran_ile_paraf_tekrari_engellenir(self):
        self.imza_asamasina()
        version = self.talep.guncelleme_zamani.isoformat()
        self.isle(self.imzaci, evrak_surumu='eski')
        self.assertEqual(self.talep.kimin_uzerinde_id, self.imzaci.pk)
        self.isle(self.imzaci, evrak_surumu=version)
        self.assertEqual(self.talep.kimin_uzerinde_id, self.son_imzaci.pk)

    def test_son_imza_tekrari_cift_aktarim_yapmaz(self):
        self.imza_asamasina()
        self.isle(self.imzaci)
        self.isle(self.son_imzaci,'imzala')
        adet = KullanilanIzin.objects.count()
        self.isle(self.son_imzaci,'imzala')
        self.assertEqual(KullanilanIzin.objects.count(), adet)

    def test_imzaci_yoksa_otomatik_tamamlanmaz(self):
        self.gonder()
        GorevBirimiImzaci.objects.all().delete()
        self.isle(self.sahip)
        self.assertEqual(self.talep.kimin_uzerinde_id,self.sahip.pk)
        self.assertEqual(KullanilanIzin.objects.count(),0)

    def test_eski_tur_imzalari_yeniden_sunulan_evraki_acmaz(self):
        self.imza_asamasina()
        self.isle(self.imzaci,'reddet',red_nedeni='Tekrar kontrol')
        self.gonder()
        for user in [self.admin,self.imzaci,self.kontrol,self.son_imzaci]:
            with self.subTest(user=user.username):
                self.assertFalse(IzinTalebi.objects.filter(views._talep_arama_filtresi(user),pk=self.talep.pk).exists())
        self.isle(self.sahip)
        self.assertFalse(IzinTalebi.objects.filter(views._talep_arama_filtresi(self.imzaci),pk=self.talep.pk).exists())

    def test_reddedilen_evrak_dolu_tarihlerle_tekrar_ilerler(self):
        self.imza_asamasina()
        self.isle(self.imzaci,'reddet',red_nedeni='Tekrar gönderin')
        onceki = (self.talep.baslangic_tarihi,self.talep.bitis_tarihi,self.talep.goreve_baslama_tarihi)
        self.gonder()
        self.assertEqual(onceki,(self.talep.baslangic_tarihi,self.talep.bitis_tarihi,self.talep.goreve_baslama_tarihi))
        self.isle(self.sahip)
        self.isle(self.kontrol)
        self.assertEqual(self.talep.kimin_uzerinde_id,self.imzaci.pk)

    def test_sahibine_iade_duzenlemeyi_acar(self):
        self.imza_asamasina()
        self.client.force_login(self.admin)
        self.client.post(reverse('izin_ara'),{'islem':'sahibine_iade','talep_id':self.talep.pk})
        self.assertEqual(self.evrak().durum,'IADE_EDILDI')
        self.gonder()
        self.assertEqual(self.talep.kimin_uzerinde_id,self.sahip.pk)

    def test_geri_cekme_imza_vermemis_kisiye_gidemez(self):
        self.imza_asamasina()
        self.client.force_login(self.admin)
        self.client.post(reverse('izin_ara'),{'islem':'evraki_geri_cek','talep_id':self.talep.pk,'hedef_kullanici_id':self.son_imzaci.pk})
        self.assertEqual(self.evrak().kimin_uzerinde_id,self.imzaci.pk)

    def test_geri_cekme_sahibine_duzenlenebilir_doner(self):
        self.imza_asamasina()
        self.client.force_login(self.admin)
        self.client.post(reverse('izin_ara'),{'islem':'evraki_geri_cek','talep_id':self.talep.pk,'hedef_kullanici_id':self.sahip.pk})
        self.assertEqual(self.evrak().durum,'IADE_EDILDI')
        self.gonder()

    def test_onaylanan_evrak_iade_ve_geri_cekme_ile_degismez(self):
        self.imza_asamasina()
        self.isle(self.imzaci)
        self.isle(self.son_imzaci,'imzala')
        self.client.force_login(self.admin)
        for route,action in [('izin_ara','sahibine_iade'),('izin_ara','evraki_geri_cek'),('izin_talepleri_yonetim','sahibine_iade'),('izin_talepleri_yonetim','geri_cek')]:
            self.client.post(reverse(route),{'islem':action,'talep_id':self.talep.pk,'hedef_kullanici_id':self.sahip.pk})
            self.assertEqual(self.evrak().durum,'ONAYLANDI')

    def test_son_imzada_yillik_bakiye_azalmissa_aktarim_yok(self):
        self.imza_asamasina()
        self.isle(self.imzaci)
        HakEdilenIzin.objects.all().update(hak_edilen_gun=0)
        self.isle(self.son_imzaci,'imzala')
        self.assertEqual(self.talep.durum,'IMZA_SURECINDE')
        self.assertFalse(self.talep.ozluk_aktarildi_mi)
        self.assertFalse(KullanilanIzin.objects.exists())

    def test_islem_gecmisi_hatasi_final_aktarimi_da_geri_alir(self):
        self.imza_asamasina()
        self.isle(self.imzaci)
        self.client.force_login(self.son_imzaci)
        with patch('portal.views._islem_gecmisi_ekle',side_effect=RuntimeError('audit failure')):
            with self.assertRaises(RuntimeError):
                self.client.post(reverse('evrak_paraf',args=[self.talep.pk]),{'islem':'imzala'})
        self.assertEqual(self.evrak().durum,'IMZA_SURECINDE')
        self.assertFalse(self.talep.ozluk_aktarildi_mi)
        self.assertFalse(KullanilanIzin.objects.exists())

    def test_sabit_evrakta_vekalet_bitince_asila_doner(self):
        today = timezone.localdate()
        vekalet = ImzaciVekaleti.objects.create(asil_kullanici=self.imzaci,vekil_kullanici=self.yabanci,
                                               baslangic_tarihi=today-timedelta(days=1),bitis_tarihi=today+timedelta(days=1))
        self.kontrol_asamasina()
        self.isle(self.kontrol)
        self.assertEqual(self.talep.kimin_uzerinde_id,self.yabanci.pk)
        vekalet.aktif_mi=False
        vekalet.save()
        self.isle(self.yabanci)
        self.assertEqual(self.talep.durum,'IADE_EDILDI')
        self.assertEqual(self.talep.kimin_uzerinde_id,self.sahip.pk)
        self.assertEqual(self.talep.imza_sirasi,0)

    def test_cakisan_vekalet_eklenmez(self):
        today=timezone.localdate()
        ImzaciVekaleti.objects.create(asil_kullanici=self.imzaci,vekil_kullanici=self.yabanci,baslangic_tarihi=today,bitis_tarihi=today+timedelta(days=2))
        self.client.force_login(self.admin)
        self.client.post(reverse('gorev_birimi_yonetimi'),{'islem':'vekalet_ekle','asil_kullanici':self.imzaci.pk,'vekil_kullanici':self.son_imzaci.pk,'baslangic_tarihi':str(today),'bitis_tarihi':str(today+timedelta(days=1))})
        self.assertEqual(ImzaciVekaleti.objects.count(),1)

    def test_onayli_silme_username_ile_yetki_yukseltmez(self):
        user=User.objects.create_user('erenaygun')
        self.assertFalse(views._ana_yonetici_mi(user))

    def test_bozuk_imza_sirasi_evraki_tamamlayamaz(self):
        self.imza_asamasina()
        IzinTalebi.objects.filter(pk=self.talep.pk).update(imza_sirasi=99)
        self.isle(self.imzaci)
        self.assertEqual(self.talep.durum,'IMZA_SURECINDE')
        self.assertFalse(self.talep.ozluk_aktarildi_mi)

    def test_eksik_kullanici_ve_bilinmeyen_yetki_uygulanmaz(self):
        self.yetki_post([self.yabanci],['unknown_permission'])
        self.assertFalse(YonetimYetkisi.objects.filter(kullanici=self.yabanci).exists())

    def test_kontrol_asamasina_ilk_paraf_atlanarak_cekilemez(self):
        self.gonder()
        self.client.force_login(self.admin)
        self.client.post(reverse('izin_talepleri_yonetim'),{'islem':'geri_cek','talep_id':self.talep.pk})
        self.assertEqual(self.evrak().kimin_uzerinde_id,self.sahip.pk)

    def test_red_ve_yeniden_gonderimde_eski_paraf_belgede_gosterilmez(self):
        from .imza_durumu import paraf_verildi_mi
        self.imza_asamasina()
        self.isle(self.imzaci,'reddet',red_nedeni='Yeniden gönder')
        self.assertFalse(paraf_verildi_mi(self.talep,self.sahip))
        self.gonder()
        self.assertFalse(paraf_verildi_mi(self.talep,self.sahip))

    def test_kontrole_geri_cekmek_sonraki_paraflari_gecersiz_kilar(self):
        from .imza_durumu import paraf_verildi_mi
        self.imza_asamasina()
        self.isle(self.imzaci)
        self.assertTrue(paraf_verildi_mi(self.talep,self.imzaci))
        self.client.force_login(self.admin)
        self.client.post(reverse('izin_talepleri_yonetim'),{'islem':'geri_cek','talep_id':self.talep.pk})
        self.assertEqual(self.evrak().kimin_uzerinde_id,self.kontrol.pk)
        self.assertFalse(paraf_verildi_mi(self.talep,self.imzaci))

    def test_eksik_ikinci_imzaci_ilk_parafla_atlanamaz(self):
        self.gonder()
        GorevBirimiImzaci.objects.filter(alt_birim_adi='').update(amir_kullanici=None,daire_baskani=None)
        self.isle(self.sahip)
        self.assertEqual(self.talep.kimin_uzerinde_id,self.sahip.pk)

    def test_pasif_personel_onaya_sunduktan_sonra_paraflayamaz(self):
        self.gonder()
        Personel.objects.filter(pk=self.talep.personel_id).update(durum='Pasif')
        self.isle(self.sahip)
        self.assertEqual(self.talep.kimin_uzerinde_id,self.sahip.pk)

    def test_bozuk_talep_id_hata_500_uretmez(self):
        self.client.force_login(self.admin)
        self.assertEqual(self.client.post(reverse('izin_ara'),{'islem':'izin_sil','talep_id':'bozuk'}).status_code,404)

    def test_sqlite_yogunlugu_final_aktarimi_geri_alir(self):
        self.imza_asamasina()
        self.isle(self.imzaci)
        self.client.force_login(self.son_imzaci)
        with patch('portal.views._islem_gecmisi_ekle',side_effect=OperationalError('database is locked')):
            response=self.client.post(reverse('evrak_paraf',args=[self.talep.pk]),{'islem':'imzala'},HTTP_ACCEPT='application/json')
        self.assertEqual(response.status_code,409)
        self.assertEqual(self.evrak().durum,'IMZA_SURECINDE')
        self.assertFalse(self.talep.ozluk_aktarildi_mi)
        self.assertFalse(KullanilanIzin.objects.exists())

    def test_diger_veritabani_hatalari_yogunluk_diye_gizlenmez(self):
        self.imza_asamasina()
        self.isle(self.imzaci)
        self.client.force_login(self.son_imzaci)
        with patch('portal.views._islem_gecmisi_ekle',side_effect=OperationalError('no such table')):
            with self.assertRaises(OperationalError):
                self.client.post(reverse('evrak_paraf',args=[self.talep.pk]),{'islem':'imzala'})
        self.assertFalse(KullanilanIzin.objects.exists())

    def test_yetki_devrinde_sabit_belge_jsonu_degismez(self):
        self.kontrol_asamasina()
        snapshot=json.dumps(self.talep.evrak_verisi,sort_keys=True)
        self.yetki_post([self.yabanci],['memur_kontrol_eden'],kontrol_devri_onay='on')
        self.assertEqual(self.evrak().durum,'IADE_EDILDI')
        self.assertEqual(snapshot,json.dumps(self.talep.evrak_verisi,sort_keys=True))
