from django.conf import settings
from django.db import migrations, models
import django.db.models.deletion


class Migration(migrations.Migration):
    dependencies = [
        migrations.swappable_dependency(settings.AUTH_USER_MODEL),
        ("portal", "0018_portal_personel_tipleri"),
    ]

    operations = [
        migrations.AddField(model_name="izintalebi", name="red_nedeni", field=models.TextField(blank=True, default="")),
        migrations.AddField(
            model_name="izintalebi",
            name="reddeden",
            field=models.ForeignKey(blank=True, null=True, on_delete=django.db.models.deletion.SET_NULL, related_name="reddettigi_izin_evraklari", to=settings.AUTH_USER_MODEL),
        ),
        migrations.AddField(model_name="izintalebi", name="reddedilme_tarihi", field=models.DateTimeField(blank=True, null=True)),
    ]
