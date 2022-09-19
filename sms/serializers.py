from .models import *
from rest_framework.serializers import Serializer
from rest_framework.fields import CharField, IntegerField, DateTimeField

#Add Product Serializer
class ProductionSystemSerialzer(Serializer):
    id = IntegerField()
    system = CharField()
    organization = CharField()
    set_id = CharField()
    blt_status = CharField()
    pre_hil_status = CharField()
    vibaration_status = CharField()
    post_hil_status = CharField()
    fgt_status = CharField()
    final_integration_st = models.CharField()
    bhd_status = CharField()
    fqm_status = CharField()
    qm_certification_st  = CharField()
    attachment = CharField()
    remarks = CharField()
    Created_at = DateTimeField()


class FlightSystemSerialzer(Serializer):
    id = IntegerField()
    system = CharField()
    organization = CharField()
    set_id = CharField()
    blt_status = CharField()
    pre_hil_status = CharField()
    vibaration_status = CharField()
    post_hil_status = CharField()
    fgt_status = CharField()
    final_integration_st = models.CharField()
    bhd_status = CharField()
    fqm_status = CharField()
    qm_certification_st  = CharField()
    attachment = CharField()
    remarks = CharField()
    Created_at = DateTimeField()


class RelifingSystemSerialzer(Serializer):
    id = IntegerField()
    system = CharField()
    organization = CharField()
    set_id = CharField()
    blt_status = CharField()
    pre_hil_status = CharField()
    vibaration_status = CharField()
    post_hil_status = CharField()
    fgt_status = CharField()
    final_integration_st = models.CharField()
    bhd_status = CharField()
    fqm_status = CharField()
    qm_certification_st  = CharField()
    attachment = CharField()
    remarks = CharField()
    Created_at = DateTimeField()