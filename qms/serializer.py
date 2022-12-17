from rest_framework.serializers import Serializer
from rest_framework.fields import CharField, IntegerField, DateTimeField
from .models import *


class QmsAuditSerializer(Serializer):
    id = IntegerField()
    audit_id = IntegerField()
    Organization = CharField()
    site = CharField()
    setup = CharField()
    certification_status = CharField()
    previous_standard = CharField()
    certification_validity_date = DateTimeField()
    certification_validity_rescheduling_date = DateTimeField()
    audit_type = CharField()
    planned_date = DateTimeField()
    audit_start_date = DateTimeField()
    audit_close_date = DateTimeField()
    audit_status = CharField()
    standard = CharField()
    remarks = CharField()
    certification_setup = CharField()
    Created_at = DateTimeField()

    #  class Meta:
    #     model = QmsAudit
    #     fields = '__a__'
    # ['id', 'audit_id', 'Organization', 'site', 'setup', 'certification_status', 'previous_standard', 'certification_validity_date',
    #       'certification_validity_rescheduling_date', 'audit_type', 'planned_date', 'audit_start_date', 'audit_close_date', 'audit_status',
    #       'standard', 'remarks','certification_setup','Created_at']


class CespAuditSerializer(Serializer):
    id = IntegerField()
    audit_id = IntegerField()
    Organization = CharField()
    site = CharField()
    certification_status = CharField()
    previous_standard = CharField()
    certification_validity = CharField()
    audit_type = CharField()
    planned_date = DateTimeField()
    audit_start_date = DateTimeField()
    audit_close_date = DateTimeField()
    audit_status = CharField()
    standard = CharField()
    remarks = CharField()
    Created_at = DateTimeField()


class QmsTrainingScheduleSerializer(Serializer):
    id = IntegerField()
    sr_no = CharField()
    training_type = CharField()
    training_start_date = DateTimeField()
    training_end_date = DateTimeField()
    trainer = CharField()
    standards = CharField()
    organizations = CharField()
    setups = CharField()
    Created_at = DateTimeField()


class QmsAuditScheduleSerializer(Serializer):
    id = IntegerField()
    sr_no = CharField()
    organization = CharField()
    standard = CharField()
    setup = CharField()
    audit_due_date = DateTimeField()
    audit_done = CharField()
    followup_done = CharField()
    certifification_validity_date = DateTimeField()
    certifification_validity_rescheduling_date = DateTimeField()
    category = CharField()
    status = CharField()
    next_due_date = DateTimeField()
    remarks = CharField()
    Created_at = DateTimeField()
