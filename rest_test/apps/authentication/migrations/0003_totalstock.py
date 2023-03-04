# Generated by Django 3.2.16 on 2023-03-03 11:44

from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('authentication', '0002_userdetail'),
    ]

    operations = [
        migrations.CreateModel(
            name='TotalStock',
            fields=[
                ('ts_key', models.BigAutoField(primary_key=True, serialize=False)),
                ('article_number', models.CharField(max_length=50)),
                ('subinventory', models.CharField(blank=True, max_length=100, null=True)),
                ('quantity', models.BigIntegerField(blank=True, null=True)),
                ('country', models.CharField(blank=True, max_length=100, null=True)),
                ('prod_centre', models.CharField(blank=True, max_length=500, null=True)),
                ('prod_group', models.CharField(blank=True, max_length=500, null=True)),
                ('description', models.CharField(blank=True, max_length=500, null=True)),
                ('prod_status_type', models.CharField(blank=True, max_length=100, null=True)),
                ('bin_cur', models.CharField(blank=True, max_length=100, null=True)),
                ('std_day', models.CharField(blank=True, max_length=100, null=True)),
                ('state', models.CharField(blank=True, max_length=50, null=True)),
                ('state_time', models.CharField(blank=True, max_length=100, null=True)),
            ],
            options={
                'db_table': 'total_stock',
                'managed': False,
            },
        ),
    ]