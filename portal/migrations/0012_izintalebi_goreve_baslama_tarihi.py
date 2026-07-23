from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ("portal", "0011_unvan"),
    ]

    operations = [
        migrations.AddField(
            model_name="izintalebi",
            name="goreve_baslama_tarihi",
            field=models.DateField(blank=True, null=True),
        ),
    ]
