from django.db import migrations, models


def eski_birim_vekalet_yetkilerini_tasi(apps, schema_editor):
    YonetimYetkisi = apps.get_model("portal", "YonetimYetkisi")
    YonetimYetkisi.objects.filter(birim_ve_vekalet_yonetimi=True).update(
        gorev_birimi_yonetimi=True,
        vekalet_yonetimi=True,
    )


class Migration(migrations.Migration):

    dependencies = [
        ("portal", "0023_izintalebi_search_indexes"),
    ]

    operations = [
        migrations.AddField(
            model_name="yonetimyetkisi",
            name="gorev_birimi_yonetimi",
            field=models.BooleanField(default=False),
        ),
        migrations.AddField(
            model_name="yonetimyetkisi",
            name="vekalet_yonetimi",
            field=models.BooleanField(default=False),
        ),
        migrations.RunPython(eski_birim_vekalet_yetkilerini_tasi, migrations.RunPython.noop),
    ]
