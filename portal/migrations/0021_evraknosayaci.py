from django.db import migrations, models
import re


def sayaclari_doldur(apps, schema_editor):
    IzinTalebi = apps.get_model("portal", "IzinTalebi")
    EvrakNoSayaci = apps.get_model("portal", "EvrakNoSayaci")
    sonlar = {}
    for evrak_no in IzinTalebi.objects.exclude(evrak_no="").values_list("evrak_no", flat=True):
        match = re.match(r"^(\d{4})-(\d+)$", evrak_no or "")
        if not match:
            continue
        yil, sira = int(match.group(1)), int(match.group(2))
        sonlar[yil] = max(sonlar.get(yil, 0), sira)
    for yil, son_no in sonlar.items():
        EvrakNoSayaci.objects.update_or_create(yil=yil, defaults={"son_no": son_no})


class Migration(migrations.Migration):

    dependencies = [
        ("portal", "0020_iscipersonel_sirketpersonel_and_more"),
    ]

    operations = [
        migrations.CreateModel(
            name="EvrakNoSayaci",
            fields=[
                ("id", models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name="ID")),
                ("yil", models.PositiveSmallIntegerField(unique=True)),
                ("son_no", models.PositiveIntegerField(default=0)),
                ("guncelleme_zamani", models.DateTimeField(auto_now=True)),
            ],
            options={
                "verbose_name": "Evrak No Sayacı",
                "verbose_name_plural": "Evrak No Sayaçları",
                "ordering": ["-yil"],
            },
        ),
        migrations.RunPython(sayaclari_doldur, migrations.RunPython.noop),
    ]
