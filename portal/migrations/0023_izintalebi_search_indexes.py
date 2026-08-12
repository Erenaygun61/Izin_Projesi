from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ("portal", "0022_izindenetimkaydi"),
    ]

    operations = [
        migrations.AddIndex(
            model_name="izintalebi",
            index=models.Index(fields=["durum", "kimin_uzerinde"], name="portal_izin_durum__e39cf8_idx"),
        ),
        migrations.AddIndex(
            model_name="izintalebi",
            index=models.Index(fields=["kullanici", "durum"], name="portal_izin_kullani_3f8601_idx"),
        ),
        migrations.AddIndex(
            model_name="izintalebi",
            index=models.Index(fields=["personel_tipi", "personel_id"], name="portal_izin_person_7d849c_idx"),
        ),
        migrations.AddIndex(
            model_name="izintalebi",
            index=models.Index(fields=["evrak_no"], name="portal_izin_evrak_n_b6ee36_idx"),
        ),
        migrations.AddIndex(
            model_name="izintalebi",
            index=models.Index(fields=["kurum_kayit_tarihi"], name="portal_izin_kurum_k_8fcafe_idx"),
        ),
        migrations.AddIndex(
            model_name="izintalebi",
            index=models.Index(fields=["guncelleme_zamani"], name="portal_izin_guncell_7b808d_idx"),
        ),
    ]
