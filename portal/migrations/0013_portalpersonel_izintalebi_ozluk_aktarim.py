from django.db import migrations, models
import django.utils.timezone


class Migration(migrations.Migration):

    dependencies = [
        ("portal", "0012_izintalebi_goreve_baslama_tarihi"),
    ]

    operations = [
        migrations.CreateModel(
            name="PortalPersonel",
            fields=[
                ("id", models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name="ID")),
                ("source_personel_id", models.PositiveIntegerField(db_index=True, unique=True)),
                ("sicil_no", models.CharField(max_length=20, unique=True)),
                ("ad", models.CharField(max_length=50)),
                ("soyad", models.CharField(max_length=50)),
                ("tc_no", models.CharField(max_length=11, unique=True)),
                ("unvan", models.CharField(blank=True, default="", max_length=150)),
                ("cep_tel", models.CharField(blank=True, default="", max_length=20)),
                ("gorev_birimi", models.CharField(blank=True, default="", max_length=180)),
                ("alt_birimi", models.CharField(blank=True, default="", max_length=180)),
                ("gorevi", models.CharField(blank=True, default="", max_length=150)),
                ("durum", models.CharField(blank=True, default="", max_length=20)),
                ("ise_baslama_tarihi", models.DateField(blank=True, null=True)),
                ("ozluk_senkron_zamani", models.DateTimeField(default=django.utils.timezone.now)),
                ("yerel_guncelleme_zamani", models.DateTimeField(auto_now=True)),
            ],
            options={"verbose_name": "Portal Personeli", "verbose_name_plural": "Portal Personelleri", "ordering": ["sicil_no"]},
        ),
        migrations.AddField(
            model_name="izintalebi",
            name="ozluk_aktarildi_mi",
            field=models.BooleanField(default=False),
        ),
        migrations.AddField(
            model_name="izintalebi",
            name="ozluk_aktarim_tarihi",
            field=models.DateTimeField(blank=True, null=True),
        ),
    ]
