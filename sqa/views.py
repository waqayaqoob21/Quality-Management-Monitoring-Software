from rest_framework.permissions import AllowAny
from rest_framework.views import APIView

from .sqaController import *

qms_obj = sqaController()
# Create your views here.

class AddSqaAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def post(self, request):
        print(request.data)
        result = qms_obj.addSqa(request.data)
        return result

class GetSqaListAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def get(self, request):
        print(request.data)
        result = qms_obj.getSqaList(request)
        return result