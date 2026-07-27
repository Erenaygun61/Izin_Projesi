from django.core.management.base import BaseCommand

from portal.views import PERSONEL_TIPI_MODELLERI, ozlukten_portala_aktar


class Command(BaseCommand):
    help = "Özlükteki memur, işçi ve şirket personelini izin portalındaki yerel kartlara aktarır."

    def add_arguments(self, parser):
        parser.add_argument(
            "--tip",
            choices=["ALL", "MEMUR", "ISCI", "SIRKET"],
            default="ALL",
            help="Aktarılacak personel tipi. Varsayılan: ALL",
        )

    def handle(self, *args, **options):
        secilen_tip = options["tip"]
        tipler = (
            PERSONEL_TIPI_MODELLERI.items()
            if secilen_tip == "ALL"
            else [(secilen_tip, PERSONEL_TIPI_MODELLERI[secilen_tip])]
        )

        genel_toplam = 0
        for tip, model in tipler:
            sayac = 0
            for kaynak in model.objects.select_related("unvan").all():
                ozlukten_portala_aktar(kaynak, tip)
                sayac += 1
            genel_toplam += sayac
            self.stdout.write(self.style.SUCCESS(f"{tip}: {sayac} personel aktarıldı/güncellendi."))

        self.stdout.write(self.style.SUCCESS(f"Toplam {genel_toplam} personel aktarıldı/güncellendi."))
