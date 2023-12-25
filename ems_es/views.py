from rest_framework.permissions import AllowAny
from rest_framework.views import APIView

from .emsController import *

qms_obj = emsController()
# Create your views here.

class AddEmsEsAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def post(self, request):
        print(request.data)
        result = qms_obj.addEmsEs(request.data)
        return result

class GetEmsEsListAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def get(self, request):
        print(request.data)
        result = qms_obj.getEmsEs(request)
        return result