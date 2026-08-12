from django.conf import settings
from django.db import migrations, models
import django.db.models.deletion


class Migration(migrations.Migration):

    dependencies = [
        ("portal", "0021_evraknosayaci"),
        migrations.swappable_dependency(settings.AUTH_USER_MODEL),
    ]

    operations = [
        migrations.CreateModel(
            name="IzinDenetimKaydi",
            fields=[
                ("id", models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name="ID")),
                ("evrak_no", models.CharField(blank=True, db_index=True, default="", max_length=30)),
                ("islem", models.CharField(choices=[("OLUSTUR", "Oluşturma"), ("GUNCELLE", "Güncelleme"), ("ONAYA_SUN", "Onaya sunma"), ("PARAF", "Paraf/imza"), ("RED", "Red"), ("IADE", "Sahibine iade"), ("GERI_CEK", "Geri çekme"), ("SIL", "Silme"), ("OZLUK_AKTAR", "Özlük aktarımı"), ("EK_EKLE", "Ek ekleme"), ("EK_SIL", "Ek silme")], db_index=True, max_length=24)),
                ("onceki_durum", models.CharField(blank=True, default="", max_length=20)),
                ("yeni_durum", models.CharField(blank=True, default="", max_length=20)),
                ("aciklama", models.CharField(blank=True, default="", max_length=500)),
                ("tarih", models.DateTimeField(auto_now_add=True, db_index=True)),
                ("hedef_kullanici", models.ForeignKey(blank=True, null=True, on_delete=django.db.models.deletion.SET_NULL, related_name="+", to=settings.AUTH_USER_MODEL)),
                ("kullanici", models.ForeignKey(blank=True, null=True, on_delete=django.db.models.deletion.SET_NULL, related_name="izin_denetim_kayitlari", to=settings.AUTH_USER_MODEL)),
                ("talep", models.ForeignKey(blank=True, null=True, on_delete=django.db.models.deletion.SET_NULL, related_name="denetim_kayitlari", to="portal.izintalebi")),
            ],
            options={
                "verbose_name": "İzin Denetim Kaydı",
                "verbose_name_plural": "İzin Denetim Kayıtları",
                "ordering": ["-tarih", "-id"],
            },
        ),
    ]
