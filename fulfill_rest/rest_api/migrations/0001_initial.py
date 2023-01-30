# Generated by Django 4.1.5 on 2023-01-29 20:13

from django.db import migrations, models


class Migration(migrations.Migration):

    initial = True

    dependencies = [
    ]

    operations = [
        migrations.CreateModel(
            name='AuthGroup',
            fields=[
                ('id', models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name='ID')),
                ('name', models.CharField(blank=True, max_length=150, null=True, unique=True)),
            ],
            options={
                'db_table': 'auth_group',
                'managed': False,
            },
        ),
        migrations.CreateModel(
            name='AuthGroupPermissions',
            fields=[
                ('id', models.BigAutoField(primary_key=True, serialize=False)),
            ],
            options={
                'db_table': 'auth_group_permissions',
                'managed': False,
            },
        ),
        migrations.CreateModel(
            name='AuthPermission',
            fields=[
                ('id', models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name='ID')),
                ('name', models.CharField(blank=True, max_length=255, null=True)),
                ('codename', models.CharField(blank=True, max_length=100, null=True)),
            ],
            options={
                'db_table': 'auth_permission',
                'managed': False,
            },
        ),
        migrations.CreateModel(
            name='AuthUser',
            fields=[
                ('id', models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name='ID')),
                ('password', models.CharField(blank=True, max_length=128, null=True)),
                ('last_login', models.DateTimeField(blank=True, null=True)),
                ('is_superuser', models.BooleanField()),
                ('username', models.CharField(blank=True, max_length=150, null=True, unique=True)),
                ('first_name', models.CharField(blank=True, max_length=150, null=True)),
                ('last_name', models.CharField(blank=True, max_length=150, null=True)),
                ('email', models.CharField(blank=True, max_length=254, null=True)),
                ('is_staff', models.BooleanField()),
                ('is_active', models.BooleanField()),
                ('date_joined', models.DateTimeField()),
            ],
            options={
                'db_table': 'auth_user',
                'managed': False,
            },
        ),
        migrations.CreateModel(
            name='AuthUserGroups',
            fields=[
                ('id', models.BigAutoField(primary_key=True, serialize=False)),
            ],
            options={
                'db_table': 'auth_user_groups',
                'managed': False,
            },
        ),
        migrations.CreateModel(
            name='AuthUserUserPermissions',
            fields=[
                ('id', models.BigAutoField(primary_key=True, serialize=False)),
            ],
            options={
                'db_table': 'auth_user_user_permissions',
                'managed': False,
            },
        ),
        migrations.CreateModel(
            name='DeliveryMethod',
            fields=[
                ('dm_key', models.CharField(max_length=50, primary_key=True, serialize=False)),
                ('del_med', models.CharField(max_length=50)),
            ],
            options={
                'db_table': 'delivery_method',
                'managed': False,
            },
        ),
        migrations.CreateModel(
            name='DjangoAdminLog',
            fields=[
                ('id', models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name='ID')),
                ('action_time', models.DateTimeField()),
                ('object_id', models.TextField(blank=True, null=True)),
                ('object_repr', models.CharField(blank=True, max_length=200, null=True)),
                ('action_flag', models.IntegerField()),
                ('change_message', models.TextField(blank=True, null=True)),
            ],
            options={
                'db_table': 'django_admin_log',
                'managed': False,
            },
        ),
        migrations.CreateModel(
            name='DjangoContentType',
            fields=[
                ('id', models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name='ID')),
                ('app_label', models.CharField(blank=True, max_length=100, null=True)),
                ('model', models.CharField(blank=True, max_length=100, null=True)),
            ],
            options={
                'db_table': 'django_content_type',
                'managed': False,
            },
        ),
        migrations.CreateModel(
            name='DjangoMigrations',
            fields=[
                ('id', models.BigAutoField(primary_key=True, serialize=False)),
                ('app', models.CharField(blank=True, max_length=255, null=True)),
                ('name', models.CharField(blank=True, max_length=255, null=True)),
                ('applied', models.DateTimeField()),
            ],
            options={
                'db_table': 'django_migrations',
                'managed': False,
            },
        ),
        migrations.CreateModel(
            name='DjangoSession',
            fields=[
                ('session_key', models.CharField(max_length=40, primary_key=True, serialize=False)),
                ('session_data', models.TextField(blank=True, null=True)),
                ('expire_date', models.DateTimeField()),
            ],
            options={
                'db_table': 'django_session',
                'managed': False,
            },
        ),
        migrations.CreateModel(
            name='MailDetail',
            fields=[
                ('ml_sub', models.CharField(max_length=100, primary_key=True, serialize=False)),
                ('ml_type_nm', models.BigIntegerField()),
                ('std_day', models.CharField(max_length=50)),
                ('ml_body', models.CharField(blank=True, max_length=2000, null=True)),
            ],
            options={
                'db_table': 'mail_detail',
                'managed': False,
            },
        ),
        migrations.CreateModel(
            name='MailStatus',
            fields=[
                ('ms_index', models.BigAutoField(primary_key=True, serialize=False)),
                ('ml_status', models.CharField(max_length=50)),
                ('up_time', models.CharField(max_length=100)),
                ('ml_bin', models.CharField(max_length=100)),
            ],
            options={
                'db_table': 'mail_status',
                'managed': False,
            },
        ),
        migrations.CreateModel(
            name='MailType',
            fields=[
                ('ml_type_nm', models.BigIntegerField(primary_key=True, serialize=False)),
                ('type_name', models.CharField(max_length=100)),
            ],
            options={
                'db_table': 'mail_type',
                'managed': False,
            },
        ),
        migrations.CreateModel(
            name='PodMethod',
            fields=[
                ('pod_key', models.CharField(max_length=50, primary_key=True, serialize=False)),
                ('pod_med', models.CharField(max_length=50)),
            ],
            options={
                'db_table': 'pod_method',
                'managed': False,
            },
        ),
        migrations.CreateModel(
            name='ShipmentInformation',
            fields=[
                ('si_index', models.BigIntegerField(primary_key=True, serialize=False)),
                ('awb_no', models.CharField(blank=True, max_length=100, null=True)),
                ('trip_no', models.CharField(blank=True, max_length=100, null=True)),
                ('shipment_nm', models.CharField(blank=True, max_length=100, null=True)),
                ('nm_of_package', models.CharField(blank=True, max_length=50, null=True)),
                ('invoice_date', models.CharField(blank=True, max_length=100, null=True)),
                ('order_nm', models.CharField(blank=True, max_length=50, null=True)),
                ('order_total', models.CharField(blank=True, max_length=100, null=True)),
                ('unit_price', models.CharField(blank=True, max_length=100, null=True)),
                ('ship_to', models.CharField(blank=True, max_length=100, null=True)),
                ('arrival_date', models.CharField(blank=True, max_length=100, null=True)),
                ('ship_date', models.CharField(blank=True, max_length=100, null=True)),
                ('pod_date', models.CharField(blank=True, max_length=100, null=True)),
                ('for_free', models.CharField(blank=True, max_length=50, null=True)),
                ('remark', models.CharField(blank=True, max_length=2000, null=True)),
                ('parcels_no', models.CharField(blank=True, max_length=100, null=True)),
                ('comment', models.CharField(blank=True, max_length=2000, null=True)),
                ('status', models.CharField(blank=True, max_length=50, null=True)),
            ],
            options={
                'db_table': 'shipment_information',
                'managed': False,
            },
        ),
        migrations.CreateModel(
            name='SoOut',
            fields=[
                ('so_index', models.BigAutoField(primary_key=True, serialize=False)),
                ('sht_row_idx', models.CharField(max_length=2000)),
                ('person_in_charge', models.CharField(max_length=100)),
                ('ship_date', models.CharField(max_length=50)),
                ('dm_key', models.CharField(max_length=50)),
                ('pod_key', models.CharField(max_length=50)),
                ('is_local', models.CharField(max_length=2000)),
                ('up_time', models.CharField(max_length=100)),
            ],
            options={
                'db_table': 'so_out',
                'managed': False,
            },
        ),
    ]