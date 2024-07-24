from rest_framework.serializers import Serializer
from .models import *
from rest_framework.fields import CharField, IntegerField, DateTimeField


class PowerSerializer(Serializer):
    id = IntegerField()
    org_type = CharField()
    system_type = CharField()
    sys_name = CharField()
    module_name = CharField()
    pc_module_id = CharField()
    lot_no = CharField()
    design_version = CharField()
    power_compliance_status = CharField()
    pcs_ldc101 = CharField()
    pcs_ldc102 = CharField()
    pcs_ldc103 = CharField()
    pcs_ldc104 = CharField()
    pcs_ldc105 = CharField()
    pcs_ldc201 = CharField()
    pcs_ldc301 = CharField()
    pcs_ldc401 = CharField()
    pcs_ldc501 = CharField()
    pcs_ldc601 = CharField()
    pcs_ldc602 = CharField()
    compliance_status = CharField()
    report_status = CharField()
    compliance_date = DateTimeField()
    remarks = CharField()
