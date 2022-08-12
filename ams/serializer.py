from rest_framework import serializers
from .models import *

#Add Product Serializer
class TaskSummarySerialzer(serializers.ModelSerializer):
    class Meta:
        model = TaskSummary
        fields = ['id','task_name', 'assigned_by','assigned_date',
                'assigned_to','target_date','status','remarks']
        