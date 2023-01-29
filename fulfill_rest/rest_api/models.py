# This is an auto-generated Django model module.
# You'll have to do the following manually to clean this up:
#   * Rearrange models' order
#   * Make sure each model has one field with primary_key=True
#   * Make sure each ForeignKey and OneToOneField has `on_delete` set to the desired behavior
#   * Remove `managed = False` lines if you wish to allow Django to create, modify, and delete the table
# Feel free to rename the models, but don't rename db_table values or field names.
from django.db import models


class AuthGroup(models.Model):
    name = models.CharField(unique=True, max_length=150, blank=True, null=True)

    class Meta:
        managed = False
        db_table = 'auth_group'


class AuthGroupPermissions(models.Model):
    id = models.BigAutoField(primary_key=True)
    group = models.ForeignKey(AuthGroup, models.DO_NOTHING)
    permission = models.ForeignKey('AuthPermission', models.DO_NOTHING)

    class Meta:
        managed = False
        db_table = 'auth_group_permissions'
        unique_together = (('group', 'permission'),)


class AuthPermission(models.Model):
    name = models.CharField(max_length=255, blank=True, null=True)
    content_type = models.ForeignKey('DjangoContentType', models.DO_NOTHING)
    codename = models.CharField(max_length=100, blank=True, null=True)

    class Meta:
        managed = False
        db_table = 'auth_permission'
        unique_together = (('content_type', 'codename'),)


class AuthUser(models.Model):
    password = models.CharField(max_length=128, blank=True, null=True)
    last_login = models.DateTimeField(blank=True, null=True)
    is_superuser = models.BooleanField()
    username = models.CharField(unique=True, max_length=150, blank=True, null=True)
    first_name = models.CharField(max_length=150, blank=True, null=True)
    last_name = models.CharField(max_length=150, blank=True, null=True)
    email = models.CharField(max_length=254, blank=True, null=True)
    is_staff = models.BooleanField()
    is_active = models.BooleanField()
    date_joined = models.DateTimeField()

    class Meta:
        managed = False
        db_table = 'auth_user'


class AuthUserGroups(models.Model):
    id = models.BigAutoField(primary_key=True)
    user = models.ForeignKey(AuthUser, models.DO_NOTHING)
    group = models.ForeignKey(AuthGroup, models.DO_NOTHING)

    class Meta:
        managed = False
        db_table = 'auth_user_groups'
        unique_together = (('user', 'group'),)


class AuthUserUserPermissions(models.Model):
    id = models.BigAutoField(primary_key=True)
    user = models.ForeignKey(AuthUser, models.DO_NOTHING)
    permission = models.ForeignKey(AuthPermission, models.DO_NOTHING)

    class Meta:
        managed = False
        db_table = 'auth_user_user_permissions'
        unique_together = (('user', 'permission'),)


class DeliveryMethod(models.Model):
    dm_key = models.CharField(primary_key=True, max_length=50)
    del_med = models.CharField(max_length=50)

    class Meta:
        managed = False
        db_table = 'delivery_method'


class DjangoAdminLog(models.Model):
    action_time = models.DateTimeField()
    object_id = models.TextField(blank=True, null=True)
    object_repr = models.CharField(max_length=200, blank=True, null=True)
    action_flag = models.IntegerField()
    change_message = models.TextField(blank=True, null=True)
    content_type = models.ForeignKey('DjangoContentType', models.DO_NOTHING, blank=True, null=True)
    user = models.ForeignKey(AuthUser, models.DO_NOTHING)

    class Meta:
        managed = False
        db_table = 'django_admin_log'


class DjangoContentType(models.Model):
    app_label = models.CharField(max_length=100, blank=True, null=True)
    model = models.CharField(max_length=100, blank=True, null=True)

    class Meta:
        managed = False
        db_table = 'django_content_type'
        unique_together = (('app_label', 'model'),)


class DjangoMigrations(models.Model):
    id = models.BigAutoField(primary_key=True)
    app = models.CharField(max_length=255, blank=True, null=True)
    name = models.CharField(max_length=255, blank=True, null=True)
    applied = models.DateTimeField()

    class Meta:
        managed = False
        db_table = 'django_migrations'


class DjangoSession(models.Model):
    session_key = models.CharField(primary_key=True, max_length=40)
    session_data = models.TextField(blank=True, null=True)
    expire_date = models.DateTimeField()

    class Meta:
        managed = False
        db_table = 'django_session'


class MailDetail(models.Model):
    ml_sub = models.CharField(primary_key=True, max_length=100)
    ml_type_nm = models.BigIntegerField()
    std_day = models.CharField(max_length=50)
    ml_body = models.CharField(max_length=2000, blank=True, null=True)

    class Meta:
        managed = False
        db_table = 'mail_detail'


class MailStatus(models.Model):
    ms_index = models.BigAutoField(primary_key=True)
    ml_sub = models.ForeignKey(MailDetail, models.DO_NOTHING, db_column='ml_sub')
    ml_status = models.CharField(max_length=50)
    up_time = models.CharField(max_length=100)
    ml_bin = models.CharField(max_length=100)

    class Meta:
        managed = False
        db_table = 'mail_status'
        unique_together = (('ms_index', 'ml_sub'),)


class MailType(models.Model):
    ml_type_nm = models.BigIntegerField(primary_key=True)
    type_name = models.CharField(max_length=100)

    class Meta:
        managed = False
        db_table = 'mail_type'


class PodMethod(models.Model):
    pod_key = models.CharField(primary_key=True, max_length=50)
    pod_med = models.CharField(max_length=50)

    class Meta:
        managed = False
        db_table = 'pod_method'


class ShipmentInformation(models.Model):
    si_index = models.BigIntegerField(primary_key=True)
    awb_no = models.CharField(max_length=100, blank=True, null=True)
    trip_no = models.CharField(max_length=100, blank=True, null=True)
    shipment_nm = models.CharField(max_length=100, blank=True, null=True)
    nm_of_package = models.CharField(max_length=50, blank=True, null=True)
    invoice_date = models.CharField(max_length=100, blank=True, null=True)
    order_nm = models.CharField(max_length=50, blank=True, null=True)
    order_total = models.CharField(max_length=100, blank=True, null=True)
    unit_price = models.CharField(max_length=100, blank=True, null=True)
    ship_to = models.CharField(max_length=100, blank=True, null=True)
    arrival_date = models.CharField(max_length=100, blank=True, null=True)
    ship_date = models.CharField(max_length=100, blank=True, null=True)
    pod_date = models.CharField(max_length=100, blank=True, null=True)
    for_free = models.CharField(max_length=50, blank=True, null=True)
    remark = models.CharField(max_length=2000, blank=True, null=True)
    parcels_no = models.CharField(max_length=100, blank=True, null=True)
    comment = models.CharField(max_length=2000, blank=True, null=True)
    status = models.CharField(max_length=50, blank=True, null=True)

    class Meta:
        managed = False
        db_table = 'shipment_information'


class SoOut(models.Model):
    so_index = models.BigAutoField(primary_key=True)
    sht_row_idx = models.CharField(max_length=2000)
    person_in_charge = models.CharField(max_length=100)
    ship_date = models.CharField(max_length=50)
    dm_key = models.CharField(max_length=50)
    pod_key = models.CharField(max_length=50)
    is_local = models.CharField(max_length=2000)
    up_time = models.CharField(max_length=100)

    class Meta:
        managed = False
        db_table = 'so_out'
