# This is an auto-generated Django model module.
# You'll have to do the following manually to clean this up:
#   * Rearrange models' order
#   * Make sure each model has one field with primary_key=True
#   * Make sure each ForeignKey and OneToOneField has `on_delete` set to the desired behavior
#   * Remove `managed = False` lines if you wish to allow Django to create, modify, and delete the table
# Feel free to rename the models, but don't rename db_table values or field names.
from django.db import models




class TotalStock(models.Model):
    ts_key = models.BigAutoField(primary_key=True)
    article_number = models.CharField(max_length=50)
    subinventory = models.CharField(max_length=100, blank=True, null=True)
    quantity = models.BigIntegerField(blank=True, null=True)
    country = models.CharField(max_length=100, blank=True, null=True)
    prod_centre = models.CharField(max_length=500, blank=True, null=True)
    prod_group = models.CharField(max_length=500, blank=True, null=True)
    description = models.CharField(max_length=500, blank=True, null=True)
    prod_status_type = models.CharField(max_length=100, blank=True, null=True)
    bin_cur = models.CharField(max_length=100, blank=True, null=True)
    std_day = models.CharField(max_length=100, blank=True, null=True)

    class Meta:
        managed = False
        db_table = 'total_stock'


class DeliveryMethod(models.Model):
    dm_key = models.CharField(primary_key=True, max_length=50)
    del_med = models.CharField(max_length=50)

    class Meta:
        managed = False
        db_table = 'delivery_method'


class LocalList(models.Model):
    lc_index = models.BigIntegerField(primary_key=True)
    arrival_date = models.CharField(max_length=50, blank=True, null=True)
    article_number = models.CharField(max_length=100, blank=True, null=True)
    description = models.CharField(max_length=500, blank=True, null=True)
    quantity = models.BigIntegerField(blank=True, null=True)
    so_no = models.CharField(max_length=100, blank=True, null=True)
    receipt_no = models.CharField(max_length=50, blank=True, null=True)
    field = models.CharField(max_length=50, blank=True, null=True)
    customer = models.CharField(max_length=100, blank=True, null=True)
    ship_date = models.CharField(max_length=1000, blank=True, null=True)
    pod_date = models.CharField(max_length=100, blank=True, null=True)
    remark = models.CharField(max_length=4000, blank=True, null=True)

    class Meta:
        managed = False
        db_table = 'local_list'


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


class ProdPose(models.Model):
    subinventory = models.CharField(primary_key=True, max_length=50)
    commnet = models.CharField(max_length=500, blank=True, null=True)

    class Meta:
        managed = False
        db_table = 'prod_pose'


class Products(models.Model):
    article_number = models.CharField(primary_key=True, max_length=50)
    country = models.CharField(max_length=50, blank=True, null=True)
    prod_centre = models.CharField(max_length=500, blank=True, null=True)
    prod_group = models.CharField(max_length=500, blank=True, null=True)
    description = models.CharField(max_length=500)
    prod_status_type = models.CharField(max_length=200, blank=True, null=True)
    bin_cur = models.CharField(max_length=50, blank=True, null=True)

    class Meta:
        managed = False
        db_table = 'products'


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


class SystemStock(models.Model):
    ss_key = models.BigAutoField(primary_key=True)
    article_number = models.ForeignKey(Products, models.DO_NOTHING, db_column='article_number')
    subinventory = models.ForeignKey(ProdPose, models.DO_NOTHING, db_column='subinventory')
    location = models.CharField(max_length=100, blank=True, null=True)
    quantity = models.BigIntegerField()
    in_date = models.CharField(max_length=50)
    expiry_date = models.CharField(max_length=50, blank=True, null=True)
    currency = models.CharField(max_length=50, blank=True, null=True)
    lot_cost = models.FloatField(blank=True, null=True)
    lot_cost_in_usd = models.FloatField(blank=True, null=True)
    std_day = models.CharField(max_length=50)

    class Meta:
        managed = False
        db_table = 'system_stock'
        unique_together = (('ss_key', 'article_number', 'subinventory'),)