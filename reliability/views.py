from rest_framework.permissions import AllowAny
from rest_framework.views import APIView

from .reliabilityController import *

qms_obj = reliabilityController()
# Create your views here.

class AddReliabilityAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def post(self, request):
        print(request.data)
        result = qms_obj.addReliability(request.data)
        return result

class GetReliabilityListAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def get(self, request):
        print(request.data)
        result = qms_obj.getReliabilityList(request)
        return result