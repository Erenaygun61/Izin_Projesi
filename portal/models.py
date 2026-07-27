from django.conf import settings
from django.db import models
from django.utils import timezone


def _turkce_baslik(metin):
    """Unvanları belge üzerinde okunur başlık biçiminde gösterir."""
    metin = (metin or "").strip()
    if not metin:
        return "-"

    def kelimeyi_duzelt(kelime):
        if not kelime:
            return kelime
        ilk = {"i": "İ", "ı": "I"}.get(kelime[0], kelime[0].upper())
        devam = kelime[1:].replace("İ", "i").replace("I", "ı").lower()
        return ilk + devam

    return " ".join(kelimeyi_duzelt(kelime) for kelime in metin.split())


def _ad_soyad_bicimi(metin):
    """Belge üzerindeki adları baş harf, soyadını tamamı büyük gösterir."""
    parcalar = _turkce_baslik(metin).split()
    if not parcalar:
        return "-"
    soyad = parcalar[-1].replace("i", "İ").replace("ı", "I").upper()
    return " ".join([*parcalar[:-1], soyad])


class Unvan(models.Model):
    """Ana maaş/özlük uygulamasındaki unvan tanımları (yalnızca okunur)."""
    ad = models.CharField(max_length=150)

    class Meta:
        managed = False
        db_table = "bordro_unvan"
        ordering = ["ad"]

    def __str__(self):
        return self.ad


class Personel(models.Model):
    sicil_no = models.CharField(max_length=20, unique=True)
    ad = models.CharField(max_length=50)
    soyad = models.CharField(max_length=50)
    tc_no = models.CharField(max_length=11, unique=True)
    unvan = models.ForeignKey(
        Unvan,
        db_column="unvan_id",
        on_delete=models.DO_NOTHING,
        related_name="personeller",
        null=True,
        blank=True,
        db_constraint=False,
    )
    cep_tel = models.CharField(max_length=20, blank=True, default="")
    gorev_birimi = models.CharField(max_length=100)
    alt_birimi = models.CharField(max_length=100, blank=True, default="")
    gorevi = models.CharField(max_length=150, blank=True, default="")
    durum = models.CharField(max_length=20, blank=True, default="")
    ise_baslama_tarihi = models.DateField(blank=True, null=True)

    class Meta:
        managed = False
        db_table = "ozluk_personel"
        ordering = ["sicil_no"]

    def __str__(self):
        return f"{self.sicil_no} - {self.ad} {self.soyad}"

    @property
    def ad_soyad(self):
        return _turkce_baslik(f"{self.ad} {self.soyad}")

    @property
    def unvani(self):
        """Evraklarda görev alanı yerine özlükteki kadro unvanı kullanılır."""
        try:
            return _turkce_baslik((self.unvan.ad if self.unvan else "") or self.gorevi)
        except Unvan.DoesNotExist:
            return _turkce_baslik(self.gorevi)


class DigerKaynakPersonel(models.Model):
    """Ana özlük uygulamasındaki işçi/şirket personel kayıtları için ortak okuma modeli."""
    sicil_no = models.CharField(max_length=20)
    ad = models.CharField(max_length=50)
    soyad = models.CharField(max_length=50)
    tc_no = models.CharField(max_length=11)
    cep_tel = models.CharField(max_length=20, blank=True, default="")
    gorev_birimi = models.CharField(max_length=150)
    alt_birimi = models.CharField(max_length=150, blank=True, default="")
    kadro_unvani = models.CharField(max_length=150, blank=True, default="")
    unvan = models.ForeignKey(
        Unvan,
        db_column="unvan_id",
        on_delete=models.DO_NOTHING,
        related_name="+",
        null=True,
        blank=True,
        db_constraint=False,
    )
    gorevi = models.CharField(max_length=150, blank=True, default="")
    fiili_gorevi = models.CharField(max_length=150, blank=True, default="")
    meslek_adi = models.CharField(max_length=150, blank=True, default="")
    birimdeki_gorevi = models.CharField(max_length=150, blank=True, default="")
    durum = models.CharField(max_length=20, blank=True, default="")
    ise_baslama_tarihi = models.DateField(blank=True, null=True)

    class Meta:
        abstract = True
        managed = False
        ordering = ["sicil_no"]

    def __str__(self):
        return f"{self.sicil_no} - {self.ad_soyad}"

    @property
    def ad_soyad(self):
        return _turkce_baslik(f"{self.ad} {self.soyad}")

    @property
    def unvani(self):
        try:
            kaynak = (self.unvan.ad if self.unvan else "") or self.kadro_unvani or self.gorevi or self.fiili_gorevi
        except Unvan.DoesNotExist:
            kaynak = self.kadro_unvani or self.gorevi or self.fiili_gorevi
        return _turkce_baslik(kaynak)


class IsciPersonel(DigerKaynakPersonel):
    statusu = models.CharField(max_length=80, blank=True, default="")
    sigortali_durumu = models.CharField(max_length=120, blank=True, default="")

    class Meta(DigerKaynakPersonel.Meta):
        db_table = "ozluk_iscipersonel"
        verbose_name = "İşçi Personel"
        verbose_name_plural = "İşçi Personeller"


class SirketPersonel(DigerKaynakPersonel):
    sirket_bilgisi = models.CharField(max_length=180, blank=True, default="")
    alt_sirket_bilgisi = models.CharField(max_length=180, blank=True, default="")
    eski_sirket_bilgisi = models.CharField(max_length=180, blank=True, default="")

    class Meta(DigerKaynakPersonel.Meta):
        db_table = "ozluk_sirketpersonel"
        verbose_name = "Şirket Personeli"
        verbose_name_plural = "Şirket Personelleri"


class PortalPersonel(models.Model):
    """Özlükten seçilerek aktarılan, izin portalına ait düzenlenebilir personel kartı."""
    PERSONEL_TIPI_CHOICES = [
        ("MEMUR", "Memur"),
        ("ISCI", "İşçi"),
        ("SIRKET", "Şirket"),
    ]
    personel_tipi = models.CharField(max_length=12, choices=PERSONEL_TIPI_CHOICES, default="MEMUR", db_index=True)
    source_personel_id = models.PositiveIntegerField(db_index=True)
    sicil_no = models.CharField(max_length=20)
    ad = models.CharField(max_length=50)
    soyad = models.CharField(max_length=50)
    tc_no = models.CharField(max_length=11)
    unvan = models.CharField(max_length=150, blank=True, default="")
    cep_tel = models.CharField(max_length=20, blank=True, default="")
    gorev_birimi = models.CharField(max_length=180, blank=True, default="")
    alt_birimi = models.CharField(max_length=180, blank=True, default="")
    gorevi = models.CharField(max_length=150, blank=True, default="")
    durum = models.CharField(max_length=20, blank=True, default="")
    ise_baslama_tarihi = models.DateField(blank=True, null=True)
    ozluk_senkron_zamani = models.DateTimeField(default=timezone.now)
    yerel_guncelleme_zamani = models.DateTimeField(auto_now=True)

    class Meta:
        ordering = ["personel_tipi", "sicil_no"]
        unique_together = ("personel_tipi", "source_personel_id")
        verbose_name = "Portal Personeli"
        verbose_name_plural = "Portal Personelleri"

    def __str__(self):
        return f"{self.sicil_no} - {self.ad_soyad}"

    @property
    def ad_soyad(self):
        return _turkce_baslik(f"{self.ad} {self.soyad}")

    @property
    def unvani(self):
        return _turkce_baslik(self.unvan or self.gorevi)


class PortalIzinHakki(models.Model):
    """İşçi/şirket gibi özlükte ayrı izin tablosu henüz olmayan personel için portal içi geçici izin bakiyesi."""
    personel = models.ForeignKey(PortalPersonel, on_delete=models.CASCADE, related_name="yerel_izin_haklari")
    izin_yili = models.PositiveSmallIntegerField()
    hak_edilen_gun = models.DecimalField(max_digits=6, decimal_places=2)
    gecerlilik_baslangic_tarihi = models.DateField()
    gecerlilik_bitis_tarihi = models.DateField()
    ise_baslama_tarihi = models.DateField(blank=True, null=True)
    olusturma_zamani = models.DateTimeField(auto_now_add=True)
    guncelleme_zamani = models.DateTimeField(auto_now=True)

    class Meta:
        ordering = ["-izin_yili"]
        unique_together = ("personel", "izin_yili")
        verbose_name = "Portal İzin Hakkı"
        verbose_name_plural = "Portal İzin Hakları"


class PortalKullanilanIzin(models.Model):
    personel = models.ForeignKey(PortalPersonel, on_delete=models.CASCADE, related_name="yerel_izin_kullanimlari")
    izin_turu = models.ForeignKey("IzinTuru", on_delete=models.PROTECT)
    hak_edilen_izin = models.ForeignKey(PortalIzinHakki, on_delete=models.SET_NULL, null=True, blank=True, related_name="kullanimlar")
    izin_baslangic_tarihi = models.DateField()
    izin_bitis_tarihi = models.DateField()
    kullanilan_gun = models.DecimalField(max_digits=6, decimal_places=2)
    ise_baslama_tarihi = models.DateField(blank=True, null=True)
    aciklama = models.TextField(blank=True, default="")
    olusturan = models.ForeignKey(settings.AUTH_USER_MODEL, on_delete=models.SET_NULL, null=True, blank=True, related_name="+")
    olusturma_zamani = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ["-izin_baslangic_tarihi", "-id"]
        verbose_name = "Portal Kullanılan İzin"
        verbose_name_plural = "Portal Kullanılan İzinler"


class IzinTuru(models.Model):
    kod = models.CharField(max_length=30, unique=True)
    ad = models.CharField(max_length=120, unique=True)
    yillik_izin_mi = models.BooleanField(default=False)
    aktif_mi = models.BooleanField(default=True)

    class Meta:
        managed = False
        db_table = "ozluk_izinturu"
        ordering = ["ad"]

    def __str__(self):
        return self.ad


class HakEdilenIzin(models.Model):
    personel = models.ForeignKey(Personel, on_delete=models.DO_NOTHING, related_name="haklar")
    izin_yili = models.PositiveSmallIntegerField()
    gecerlilik_baslangic_tarihi = models.DateField()
    gecerlilik_bitis_tarihi = models.DateField()
    ise_baslama_tarihi = models.DateField(blank=True, null=True)
    hak_edilen_gun = models.DecimalField(max_digits=6, decimal_places=2)

    class Meta:
        managed = False
        db_table = "ozluk_hakedilenizin"
        ordering = ["-izin_yili"]


class KullanilanIzin(models.Model):
    personel = models.ForeignKey(Personel, on_delete=models.DO_NOTHING, related_name="kullanimlar")
    izin_turu = models.ForeignKey(IzinTuru, on_delete=models.DO_NOTHING)
    hak_edilen_izin = models.ForeignKey(HakEdilenIzin, on_delete=models.DO_NOTHING, blank=True, null=True)
    izin_baslangic_tarihi = models.DateField()
    izin_bitis_tarihi = models.DateField()
    kullanilan_gun = models.DecimalField(max_digits=6, decimal_places=2)
    ise_baslama_tarihi = models.DateField(blank=True, null=True)
    vekalet_eden = models.CharField(max_length=150, blank=True, default="")
    izin_gecirecegi_adres = models.TextField(blank=True, default="")
    ilgili_birim_onay_tarihi = models.DateField(blank=True, null=True)
    ilgili_birim_onay_sayisi = models.CharField(max_length=80, blank=True, default="")
    ik_onay_tarihi = models.DateField(blank=True, null=True)
    ik_onay_sayisi = models.CharField(max_length=80, blank=True, default="")
    aciklama = models.TextField(blank=True, default="")
    sisteme_giris_yapabilir_mi = models.BooleanField(default=True)
    olusturma_zamani = models.DateTimeField(blank=True, null=True)
    guncelleme_zamani = models.DateTimeField(blank=True, null=True)
    olusturan = models.ForeignKey(settings.AUTH_USER_MODEL, on_delete=models.DO_NOTHING, blank=True, null=True, related_name="+")
    guncelleyen = models.ForeignKey(settings.AUTH_USER_MODEL, on_delete=models.DO_NOTHING, blank=True, null=True, related_name="+")

    class Meta:
        managed = False
        db_table = "ozluk_kullanilanizin"
        ordering = ["-izin_baslangic_tarihi", "-id"]


class IzinTalebi(models.Model):
    PERSONEL_TIPI_CHOICES = [
        ("MEMUR", "Memur"),
        ("ISCI", "İşçi"),
        ("SIRKET", "Şirket"),
    ]
    DURUM_CHOICES = [
        ("TASLAK", "Taslak"),
        ("PARAF_BEKLIYOR", "Paraf Bekliyor"),
        ("IMZA_SURECINDE", "\u0130mza S\u00fcrecinde"),
        ("ONAYLANDI", "Onayland\u0131"),
        ("REDDEDILDI", "Reddedildi"),
        ("IADE_EDILDI", "Sahibine \u0130ade Edildi"),
    ]
    kullanici = models.ForeignKey(settings.AUTH_USER_MODEL, on_delete=models.CASCADE, related_name="izin_talepleri")
    personel_id = models.PositiveIntegerField()
    personel_tipi = models.CharField(max_length=12, choices=PERSONEL_TIPI_CHOICES, default="MEMUR")
    # Taslakta bu alanlar henüz doldurulmamış olabilir. Onaya sunulurken
    # form tarafından zorunlu olarak doğrulanırlar.
    izin_turu = models.ForeignKey(IzinTuru, on_delete=models.PROTECT, null=True, blank=True)
    kurum_kayit_tarihi = models.DateField(null=True, blank=True)
    evrak_no = models.CharField(max_length=30, blank=True, default="")
    cep_telefonu = models.CharField(max_length=20, blank=True, default="")
    izin_sebebi = models.CharField(max_length=200, blank=True, default="")
    ait_oldugu_yil = models.CharField(max_length=30, blank=True, default="")
    baslangic_tarihi = models.DateField(null=True, blank=True)
    bitis_tarihi = models.DateField(null=True, blank=True)
    goreve_baslama_tarihi = models.DateField(null=True, blank=True)
    ozluk_aktarildi_mi = models.BooleanField(default=False)
    ozluk_aktarim_tarihi = models.DateTimeField(null=True, blank=True)
    kullanilan_gun = models.DecimalField(max_digits=6, decimal_places=2, null=True, blank=True)
    izin_hakki = models.DecimalField(max_digits=6, decimal_places=2, null=True, blank=True)
    kalan_izin = models.DecimalField(max_digits=6, decimal_places=2, null=True, blank=True)
    kontrol_eden = models.ForeignKey(settings.AUTH_USER_MODEL, on_delete=models.SET_NULL, null=True, blank=True, related_name="kontrol_edecegi_izinler")
    kontrol_tarihi = models.DateTimeField(null=True, blank=True)
    kimin_uzerinde = models.ForeignKey(settings.AUTH_USER_MODEL, on_delete=models.SET_NULL, null=True, blank=True, related_name="uzerindeki_izinler")
    imza_sirasi = models.PositiveSmallIntegerField(default=0)
    imza_tamamlanma_tarihi = models.DateTimeField(null=True, blank=True)
    adres = models.TextField(blank=True, default="")
    aciklama = models.TextField(blank=True, default="")
    durum = models.CharField(max_length=20, choices=DURUM_CHOICES, default="PARAF_BEKLIYOR")
    olusturma_zamani = models.DateTimeField(auto_now_add=True)
    guncelleme_zamani = models.DateTimeField(auto_now=True)

    class Meta:
        ordering = ["-olusturma_zamani"]
        verbose_name = "İzin Talebi"
        verbose_name_plural = "İzin Talepleri"

    def __str__(self):
        return f"{self.personel_id} / {self.izin_turu} / {self.kullanilan_gun} gün"


class KaynakGorevBirimi(models.Model):
    ad = models.CharField(max_length=200)
    kod = models.CharField(max_length=20, blank=True, null=True)
    birim_turu = models.CharField(max_length=35, blank=True, default="DIGER")
    amir_personel = models.ForeignKey(Personel, db_column="amir_personel_id", on_delete=models.DO_NOTHING, null=True, blank=True, db_constraint=False, related_name="+")
    aktif_mi = models.BooleanField(default=True)

    class Meta:
        managed = False
        db_table = "bordro_gorevbirimi"
        ordering = ["ad"]


class KaynakAltBirim(models.Model):
    gorev_birimi = models.ForeignKey(KaynakGorevBirimi, db_column="gorev_birimi_id", on_delete=models.DO_NOTHING, db_constraint=False, related_name="alt_birimler")
    ad = models.CharField(max_length=200)
    kod = models.CharField(max_length=20, blank=True, null=True)
    amir_personel = models.ForeignKey(Personel, db_column="amir_personel_id", on_delete=models.DO_NOTHING, null=True, blank=True, db_constraint=False, related_name="+")
    aktif_mi = models.BooleanField(default=True)

    class Meta:
        managed = False
        db_table = "bordro_altbirim"
        ordering = ["gorev_birimi__ad", "ad"]


class GorevBirimiImzaci(models.Model):
    birim_adi = models.CharField(max_length=150)
    alt_birim_adi = models.CharField(max_length=150, blank=True, default="")
    birim_turu = models.CharField(max_length=35, blank=True, default="DIGER")
    ust_birim_adi = models.CharField(max_length=180, blank=True, default="")
    amir_kullanici = models.ForeignKey(settings.AUTH_USER_MODEL, on_delete=models.SET_NULL, null=True, blank=True, related_name="amir_oldugu_portal_birimleri")
    daire_baskani = models.ForeignKey(settings.AUTH_USER_MODEL, on_delete=models.SET_NULL, null=True, blank=True, related_name="daire_baskani_birimleri")
    sube_muduru = models.ForeignKey(settings.AUTH_USER_MODEL, on_delete=models.SET_NULL, null=True, blank=True, related_name="sube_muduru_birimleri")
    aktif_mi = models.BooleanField(default=True)
    guncelleme_zamani = models.DateTimeField(auto_now=True)

    class Meta:
        ordering = ["birim_adi", "alt_birim_adi"]
        unique_together = ("birim_adi", "alt_birim_adi")
        verbose_name = "Görev Birimi İmzacı"
        verbose_name_plural = "Görev Birimi İmzacıları"

    def __str__(self):
        return f"{self.birim_adi} / {self.alt_birim_adi or '-'}"


class ImzaciVekaleti(models.Model):
    asil_kullanici = models.ForeignKey(settings.AUTH_USER_MODEL, on_delete=models.CASCADE, related_name="asil_vekaletleri")
    vekil_kullanici = models.ForeignKey(settings.AUTH_USER_MODEL, on_delete=models.CASCADE, related_name="vekil_vekaletleri")
    baslangic_tarihi = models.DateField()
    bitis_tarihi = models.DateField()
    aciklama = models.CharField(max_length=200, blank=True, default="")
    aktif_mi = models.BooleanField(default=True)
    olusturma_zamani = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ["-baslangic_tarihi", "asil_kullanici__username"]
        verbose_name = "İmzacı Vekaleti"
        verbose_name_plural = "İmzacı Vekaletleri"

    def __str__(self):
        return f"{self.asil_kullanici} -> {self.vekil_kullanici}"


class PortalAyar(models.Model):
    # Eski alan korunur; mevcut Memur kontrol eden tanımı veri kaybetmeden
    # kontrol_eden_memur alanına taşınır.
    kontrol_eden = models.ForeignKey(settings.AUTH_USER_MODEL, on_delete=models.SET_NULL, null=True, blank=True, related_name="kontrol_eden_ayarlar")
    kontrol_eden_memur = models.ForeignKey(settings.AUTH_USER_MODEL, on_delete=models.SET_NULL, null=True, blank=True, related_name="memur_kontrol_ayarlari")
    kontrol_eden_isci = models.ForeignKey(settings.AUTH_USER_MODEL, on_delete=models.SET_NULL, null=True, blank=True, related_name="isci_kontrol_ayarlari")
    kontrol_eden_sirket = models.ForeignKey(settings.AUTH_USER_MODEL, on_delete=models.SET_NULL, null=True, blank=True, related_name="sirket_kontrol_ayarlari")
    guncelleme_zamani = models.DateTimeField(auto_now=True)

    class Meta:
        verbose_name = "Portal Ayarı"
        verbose_name_plural = "Portal Ayarları"

    def __str__(self):
        return "İzin Portalı Ayarları"

    @classmethod
    def aktif(cls):
        obj, _ = cls.objects.get_or_create(id=1)
        return obj


class YonetimYetkisi(models.Model):
    kullanici = models.OneToOneField(settings.AUTH_USER_MODEL, on_delete=models.CASCADE, related_name="portal_yetki")
    # Her yetki, tek bir genel "yönetici" bayrağı yerine ayrı ayrı verilir.
    # Eski kayıtların davranışını korumak için evrak_yonetimi alanı tutulur.
    evrak_yonetimi = models.BooleanField(default=False)
    tum_izinleri_goruntuleme = models.BooleanField(default=False)
    memur_kontrol_eden = models.BooleanField(default=False)
    isci_kontrol_eden = models.BooleanField(default=False)
    sirket_kontrol_eden = models.BooleanField(default=False)
    evrak_geri_cekme = models.BooleanField(default=False)
    sahibine_iade = models.BooleanField(default=False)
    izin_silme = models.BooleanField(default=False)
    tamamlanan_izin_bildirimi = models.BooleanField(default=False)
    birim_ve_vekalet_yonetimi = models.BooleanField(default=False)
    # Personel tipi bazlı ayrıntılı evrak yetkileri.
    memur_izin_goruntuleme = models.BooleanField(default=False)
    isci_izin_goruntuleme = models.BooleanField(default=False)
    sirket_izin_goruntuleme = models.BooleanField(default=False)
    memur_izin_arama = models.BooleanField(default=False)
    isci_izin_arama = models.BooleanField(default=False)
    sirket_izin_arama = models.BooleanField(default=False)
    memur_evrak_geri_cekme = models.BooleanField(default=False)
    isci_evrak_geri_cekme = models.BooleanField(default=False)
    sirket_evrak_geri_cekme = models.BooleanField(default=False)
    memur_sahibine_iade = models.BooleanField(default=False)
    isci_sahibine_iade = models.BooleanField(default=False)
    sirket_sahibine_iade = models.BooleanField(default=False)
    memur_izin_silme = models.BooleanField(default=False)
    isci_izin_silme = models.BooleanField(default=False)
    sirket_izin_silme = models.BooleanField(default=False)
    memur_tamamlanan_izin_bildirimi = models.BooleanField(default=False)
    isci_tamamlanan_izin_bildirimi = models.BooleanField(default=False)
    sirket_tamamlanan_izin_bildirimi = models.BooleanField(default=False)
    aktif_mi = models.BooleanField(default=True)
    aciklama = models.CharField(max_length=200, blank=True, default="")
    olusturma_zamani = models.DateTimeField(auto_now_add=True)
    guncelleme_zamani = models.DateTimeField(auto_now=True)

    class Meta:
        ordering = ["kullanici__first_name", "kullanici__last_name", "kullanici__username"]
        verbose_name = "Y\u00f6netim Yetkisi"
        verbose_name_plural = "Y\u00f6netim Yetkileri"

    def save(self, *args, **kwargs):
        super().save(*args, **kwargs)
        kontrol_alanlari = ("memur_kontrol_eden", "isci_kontrol_eden", "sirket_kontrol_eden")
        for alan in kontrol_alanlari:
            if getattr(self, alan, False):
                YonetimYetkisi.objects.exclude(pk=self.pk).filter(**{alan: True}).update(**{alan: False})

    def __str__(self):
        return f"{self.kullanici} / {'Aktif' if self.aktif_mi else 'Pasif'}"


class IzinIslemGecmisi(models.Model):
    ISLEM_CHOICES = [
        ("TASLAK", "Taslak oluşturuldu"),
        ("ONAYA_SUNULDU", "Onaya sunuldu"),
        ("KONTROL_PARAFI", "Kontrol parafı"),
        ("IMZA", "İmza / paraf"),
        ("ONAY", "Onaylandı"),
        ("RED", "Reddedildi"),
        ("IADE", "Sahibine iade edildi"),
        ("GERI_CEK", "Kontrole geri çekildi"),
    ]
    talep = models.ForeignKey(IzinTalebi, on_delete=models.CASCADE, related_name="islem_gecmisi")
    kullanici = models.ForeignKey(settings.AUTH_USER_MODEL, on_delete=models.SET_NULL, null=True, blank=True, related_name="izin_islem_gecmisi")
    islem = models.CharField(max_length=24, choices=ISLEM_CHOICES)
    rol = models.CharField(max_length=80, blank=True, default="")
    aciklama = models.CharField(max_length=300, blank=True, default="")
    tarih = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ["-tarih", "-id"]
        verbose_name = "İzin İşlem Geçmişi"
        verbose_name_plural = "İzin İşlem Geçmişleri"


class IzinBildirimKapatma(models.Model):
    kullanici = models.ForeignKey(settings.AUTH_USER_MODEL, on_delete=models.CASCADE, related_name="kapatilan_izin_bildirimleri")
    talep = models.ForeignKey(IzinTalebi, on_delete=models.CASCADE, related_name="bildirim_kapatmalari")
    kapatilma_zamani = models.DateTimeField(auto_now_add=True)

    class Meta:
        unique_together = ("kullanici", "talep")
        verbose_name = "Kapatılan İzin Bildirimi"
        verbose_name_plural = "Kapatılan İzin Bildirimleri"


