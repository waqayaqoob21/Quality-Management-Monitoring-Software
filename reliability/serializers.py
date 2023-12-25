from rest_framework import serializers
from .models import *

class ReliabilitySerializer(serializers.ModelSerializer):
    class Meta:
        model = Reliability
        fields = '__all__'