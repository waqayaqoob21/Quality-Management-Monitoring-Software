from django.db.models import TextField

from rest_framework.fields import CharField, IntegerField, DateTimeField
from rest_framework_simplejwt.serializers import TokenObtainPairSerializer
from rest_framework_simplejwt.tokens import RefreshToken
from rest_framework.serializers import Serializer

class MyTokenObtainPairSerializer(TokenObtainPairSerializer):
    @classmethod
    def get_token(cls, user):
        # token = super().get_token(user)
        token = RefreshToken.for_user(user)
        # Add custom claims
        token['name'] = user.username
        # Add more custom fields from your custom user model, If you have a
        # custom user model.
        # ...

        return token


class DocListSerializer(Serializer):
    id = IntegerField()
    doc_name = CharField()
    doc_type = CharField()
    sender = CharField()
    receive_date = DateTimeField()
    marked_to = CharField()
    marked_date = DateTimeField()
    due_date = DateTimeField()
    task_date = DateTimeField()
    status = CharField()
    sent_to = CharField()
    sent_date = DateTimeField()
    isActive = IntegerField()
    attachement = CharField()
    remarks = CharField()
    created_at = DateTimeField()
    product_sr_no = CharField()
    tracking_id = CharField()
    certificate_number = CharField()
    last_meeting_date = DateTimeField()

class CertificationSerializer(Serializer):
    id = IntegerField()
    organization = CharField()
    certificate_serial_number = CharField()
    certificate_copy = CharField()
    certificate_type = CharField()
    certificate_issue_date = DateTimeField()
    product_name = CharField()
    product = CharField()
    identification_no = CharField()
    qualification_date = DateTimeField()
    manufacturer = CharField()
    bhd_no = CharField()
    certificate_number = CharField()
    certificate_date = DateTimeField()
    audit_report_no = CharField()
    audit_report_date = DateTimeField()
    table_1_description = CharField()
    table_2_heading = CharField()
    table_2_description = CharField()
    remarks = CharField()

