from django.core.management.base import BaseCommand
from django.db import transaction

from portal.evrak_verisi import evraki_sabitle
from portal.models import IzinTalebi
from portal.views import _teklif_parafi_verildi_mi, talep_aktif_imzacilari, talep_personeli


class Command(BaseCommand):
    help = 'Snapshot olmayan eski paraflanmış evrakları mevcut bilgilerle sabitler.'

    def handle(self, *args, **options):
        adet, atlanan = 0, 0
        adaylar = IzinTalebi.objects.filter(durum__in=['PARAF_BEKLIYOR', 'IMZA_SURECINDE', 'ONAYLANDI'], evrak_verisi={})
        for pk in adaylar.values_list('pk', flat=True):
            with transaction.atomic():
                talep = IzinTalebi.objects.select_for_update().get(pk=pk)
                if talep.evrak_verisi or (talep.durum == 'PARAF_BEKLIYOR' and not _teklif_parafi_verildi_mi(talep)):
                    continue
                personel = talep_personeli(talep)
                if not personel:
                    atlanan += 1
                    continue
                imzacilar = talep_aktif_imzacilari(talep)
                if talep.durum == 'ONAYLANDI':
                    imzalar, gorulen = [], set()
                    for kayit in talep.islem_gecmisi.select_related('kullanici').filter(islem__in=['IMZA', 'ONAY']).order_by('tarih', 'id'):
                        if not kayit.kullanici or kayit.kullanici_id == talep.kontrol_eden_id or (kayit.kullanici_id == talep.kullanici_id and kayit.rol == 'Teklif Eden'):
                            continue
                        key = (kayit.rol, kayit.kullanici_id)
                        if key not in gorulen:
                            imzalar.append((kayit.rol, kayit.kullanici))
                            gorulen.add(key)
                    if imzalar:
                        imzacilar = imzalar
                evraki_sabitle(talep, personel, imzacilar)
                adet += 1
        self.stdout.write(f'Sabitlenen: {adet}; kaynak personeli bulunamayan: {atlanan}')
        self.stdout.write('Eski unvan/ad bilgileri tarihsel olarak kaydedilmediğinden mevcut değerler korunmuştur.')
