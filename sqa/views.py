from rest_framework.permissions import AllowAny
from rest_framework.views import APIView

from .sqaController import *

sqa_obj = sqaController()
# Create your views here.

class GetModuleNameIDListAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def get(self, request):
        print(request.data)
        result = sqa_obj.getModuleNameIDList(request)
        return result
class AddSqaAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def post(self, request):
        print(request.data)
        result = sqa_obj.addSqa(request.data)
        return result
class GetSqaDashbaordCountAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def get(self, request):
        print(request.data)
        result = sqa_obj.getSqaDashboardCount(request)
        return result
class GetSqaListAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def get(self, request):
        print(request.data)
        result = sqa_obj.getSqaList(request)
        return result

class DeleteSqaAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def delete(self, request):
        result = sqa_obj.deleteSqa(request)
        return result

class GetSqaHistoryAPIVIEW(APIView):
    permission_classes = [AllowAny]
    def get(self, request):
        print(request.data)
        result = sqa_obj.getSqaHistory(request)
        return result

class GetSqaAPIVIEW(APIView):
    permission_classes = [AllowAny]
    def get(self, request):
        print(request.data)
        result = sqa_obj.getSqa(request)
        return result

class DeleteSqaHistoryAPIVIEW(APIView):
    permission_classes = [AllowAny]
    def get(self,request):
        print(request.data)
        result = sqa_obj.deleteSqaHistory(request)
        return result

class GetSqaSummaryAPIVIEW(APIView):
    permission_classes = [AllowAny]
    def get(self,request):
        print(request.data)
        result = sqa_obj.getSqaSummary(request)
        return result