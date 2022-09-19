from django.shortcuts import render

# Create your views here.
from rest_framework.permissions import AllowAny
from rest_framework.views import APIView

from sms.SmsController import *

task_obj = SmsController()
# Create your views here.

class AddProductionAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def post(self, request):
        result = task_obj.AddProductionStatus(request.data)
        return result

class EditProductionAPIVIEW(APIView):
    permission_classes = [AllowAny]
    def put(self, request):
        result = task_obj.AddProductionStatus(request.data)
        print(result)
        return result

class GetProductionListAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def get(self, request):
        result = task_obj.GetProductionList(request.data)
        return result
class GetProductionPDFAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def get(self, request):
        result = task_obj.GetProductionPDFList(request.data)
        return result
class GetProductionExcelAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def get(self, request):
        result = task_obj.GetProductionExcelList(request.data)
        return result

#Flight System Status

class AddFlightAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def post(self, request):
        result = task_obj.AddFlightStatus(request.data)
        return result

class EditFlightAPIVIEW(APIView):
    permission_classes = [AllowAny]
    def put(self, request):
        result = task_obj.AddFlightStatus(request.data)
        print(result)
        return result

class GetFlightListAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def get(self, request):
        result = task_obj.GetFlightList(request.data)
        return result
class GetFlightPDFAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def get(self, request):
        result = task_obj.GetFlightPDFList(request.data)
        return result
class GetFlightExcelAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def get(self, request):
        result = task_obj.GetFlightExcelList(request.data)
        return result




class AddRelifingAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def post(self, request):
        result = task_obj.AddRelifingStatus(request.data)
        return result

class EditRelifingAPIVIEW(APIView):
    permission_classes = [AllowAny]
    def put(self, request):
        result = task_obj.AddRelifingStatus(request.data)
        print(result)
        return result

class GetRelifingListAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def get(self, request):
        result = task_obj.GetRelifingList(request.data)
        return result
class GetRelifingPDFAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def get(self, request):
        result = task_obj.GetRelifingPDFList(request.data)
        return result
class GetRelifingExcelAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def get(self, request):
        result = task_obj.GetRelifingExcelList(request.data)
        return result