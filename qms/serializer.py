from rest_framework.serializers import Serializer
from rest_framework.fields import CharField, IntegerField, DateTimeField


class QmsAuditSerializer(Serializer):
    id = IntegerField()
    audit_id = IntegerField()
    Organization = CharField()
    site = CharField()
    certification_status= CharField()
    previous_standard = CharField()
    certification_validity = CharField()
    audit_type = CharField()
    planned_date = DateTimeField()
    audit_start_date =DateTimeField()
    audit_close_date = DateTimeField()
    audit_status = CharField()
    standard = CharField()
    remarks = CharField()
    Created_at = DateTimeField()

class CespAuditSerializer(Serializer):
    id = IntegerField()
    audit_id = IntegerField()
    Organization = CharField()
    site = CharField()
    certification_status= CharField()
    previous_standard = CharField()
    certification_validity = CharField()
    audit_type = CharField()
    planned_date = DateTimeField()
    audit_start_date =DateTimeField()
    audit_close_date = DateTimeField()
    audit_status = CharField()
    standard = CharField()
    remarks = CharField()
    Created_at = DateTimeField()