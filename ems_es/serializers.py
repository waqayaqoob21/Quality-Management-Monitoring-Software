from rest_framework import serializers
from .models import *

class EmsEsSerializer(serializers.ModelSerializer):
    class Meta:
        model = EmsEs
        fields = '__all__'