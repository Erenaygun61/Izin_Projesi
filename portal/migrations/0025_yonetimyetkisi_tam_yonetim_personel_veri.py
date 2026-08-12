from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ("portal", "0024_yonetimyetkisi_detayli_yonetim_ekranlari"),
    ]

    operations = [
        migrations.AddField(
            model_name="yonetimyetkisi",
            name="tam_yonetim_yetkisi",
            field=models.BooleanField(default=False),
        ),
        migrations.AddField(
            model_name="yonetimyetkisi",
            name="personel_veri_yonetimi",
            field=models.BooleanField(default=False),
        ),
    ]
