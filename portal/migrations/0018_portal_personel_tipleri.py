from django.conf import settings
from django.db import migrations, models
import django.db.models.deletion


class Migration(migrations.Migration):

    dependencies = [
        migrations.swappable_dependency(settings.AUTH_USER_MODEL),
        ("portal", "0017_izintalebi_cep_telefonu"),
    ]

    operations = [
        migrations.AddField(
            model_name="portalpersonel",
            name="personel_tipi",
            field=models.CharField(
                choices=[("MEMUR", "Memur"), ("ISCI", "İşçi"), ("SIRKET", "Şirket")],
                db_index=True,
                default="MEMUR",
                max_length=12,
            ),
        ),
        migrations.AlterField(
            model_name="portalpersonel",
            name="source_personel_id",
            field=models.PositiveIntegerField(db_index=True),
        ),
        migrations.AlterField(
            model_name="portalpersonel",
            name="sicil_no",
            field=models.CharField(max_length=20),
        ),
        migrations.AlterField(
            model_name="portalpersonel",
            name="tc_no",
            field=models.CharField(max_length=11),
        ),
        migrations.AlterUniqueTogether(
            name="portalpersonel",
            unique_together={("personel_tipi", "source_personel_id")},
        ),
        migrations.CreateModel(
            name="PortalIzinHakki",
            fields=[
                ("id", models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name="ID")),
                ("izin_yili", models.PositiveSmallIntegerField()),
                ("hak_edilen_gun", models.DecimalField(decimal_places=2, max_digits=6)),
                ("gecerlilik_baslangic_tarihi", models.DateField()),
                ("gecerlilik_bitis_tarihi", models.DateField()),
                ("ise_baslama_tarihi", models.DateField(blank=True, null=True)),
                ("olusturma_zamani", models.DateTimeField(auto_now_add=True)),
                ("guncelleme_zamani", models.DateTimeField(auto_now=True)),
                ("personel", models.ForeignKey(on_delete=django.db.models.deletion.CASCADE, related_name="yerel_izin_haklari", to="portal.portalpersonel")),
            ],
            options={
                "verbose_name": "Portal İzin Hakkı",
                "verbose_name_plural": "Portal İzin Hakları",
                "ordering": ["-izin_yili"],
                "unique_together": {("personel", "izin_yili")},
            },
        ),
        migrations.CreateModel(
            name="PortalKullanilanIzin",
            fields=[
                ("id", models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name="ID")),
                ("izin_baslangic_tarihi", models.DateField()),
                ("izin_bitis_tarihi", models.DateField()),
                ("kullanilan_gun", models.DecimalField(decimal_places=2, max_digits=6)),
                ("ise_baslama_tarihi", models.DateField(blank=True, null=True)),
                ("aciklama", models.TextField(blank=True, default="")),
                ("olusturma_zamani", models.DateTimeField(auto_now_add=True)),
                ("hak_edilen_izin", models.ForeignKey(blank=True, null=True, on_delete=django.db.models.deletion.SET_NULL, related_name="kullanimlar", to="portal.portalizinhakki")),
                ("izin_turu", models.ForeignKey(on_delete=django.db.models.deletion.PROTECT, to="portal.izinturu")),
                ("olusturan", models.ForeignKey(blank=True, null=True, on_delete=django.db.models.deletion.SET_NULL, related_name="+", to=settings.AUTH_USER_MODEL)),
                ("personel", models.ForeignKey(on_delete=django.db.models.deletion.CASCADE, related_name="yerel_izin_kullanimlari", to="portal.portalpersonel")),
            ],
            options={
                "verbose_name": "Portal Kullanılan İzin",
                "verbose_name_plural": "Portal Kullanılan İzinler",
                "ordering": ["-izin_baslangic_tarihi", "-id"],
            },
        ),
    ]
