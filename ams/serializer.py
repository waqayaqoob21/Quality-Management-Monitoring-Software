from rest_framework import serializers
from .models import *


# Add Product Serializer
class TaskSummarySerialzer(serializers.ModelSerializer):
    class Meta:
        model = TaskSummary
        fields = ['id', 'task_name', 'assigned_by', 'assigned_date',
                  'assigned_to', 'target_date', 'task_date', 'status', 'remarks', 'follow_up','priority']


class OcrDataSerializer(serializers.ModelSerializer):
    class Meta:
        model = OcrDataModel
        fields = ['id', 'organization', 'senior_directorate', 'site', 'test_report_no', 'job_card_no',
                  'test_report_date', 'product_name',
                  'id_no', 'lot_no_lot_size', 'test_name', 'test_type', 'qualification_standard', 'test_specifications',
                  'results', 'remarks', 'Created_at']
class PriorityNotificationSerialzer(serializers.ModelSerializer):
    class Meta:
        model = PriorityNotification
        fields = ['id', 'task_id', 'task_group', 'message','created_at']