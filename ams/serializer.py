from rest_framework import serializers
from .models import *

#Add Product Serializer
class TaskSummarySerialzer(serializers.ModelSerializer):
    class Meta:
        model = TaskSummary
        fields = ['id','task', 'assigned_by','assigned_date_by','target_date_by',
                'assigned_to','assigned_date_to','target_date_to','attachment','isActive']
        