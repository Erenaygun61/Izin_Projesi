from django.db import migrations, models
import django.db.models.deletion


class Migration(migrations.Migration):
    dependencies = [('portal', '0027_evrak_verisi')]
    operations = [migrations.SeparateDatabaseAndState(state_operations=[
        migrations.AddField(
            model_name='hakedilenizin', name='izin_turu',
            field=models.ForeignKey(blank=True, null=True, db_constraint=False,
                                    on_delete=django.db.models.deletion.DO_NOTHING, to='portal.izinturu'),
        ),
    ])]
