from rest_framework import serializers
from .models import *

class EmcSerializer(serializers.ModelSerializer):
    class Meta:
        model = Emc
        fields = '__all__'