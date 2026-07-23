from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ("portal", "0013_portalpersonel_izintalebi_ozluk_aktarim"),
    ]

    operations = [
        migrations.AddField(
            model_name="yonetimyetkisi",
            name="izin_silme",
            field=models.BooleanField(default=False),
        ),
    ]
