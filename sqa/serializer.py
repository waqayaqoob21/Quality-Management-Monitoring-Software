from rest_framework import serializers
from .models import *

class SqaSerializer(serializers.ModelSerializer):
    class Meta:
        model = Sqa
        fields = '__all__'