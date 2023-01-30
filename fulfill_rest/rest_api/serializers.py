from rest_api.models import *
from rest_framework import serializers


class ShipmentInformationSerializer(serializers.HyperlinkedModelSerializer):
    class Meta:
        model = ShipmentInformation
        fields = [
            'si_index', 'awb_no', 'trip_no', 'shipment_nm', 'nm_of_package', 'invoice_date',
            'order_nm', 'order_total', 'unit_price', 'ship_to', 'arrival_date', 'ship_date',
            'pod_date', 'for_free', 'remark', 'parcels_no', 'comment', 'status'
        ]


