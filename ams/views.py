from django.shortcuts import render

# Create your views here.
from rest_framework.permissions import AllowAny
from rest_framework.views import APIView

from ams.AmsController import *

task_obj = AmsController()


# Create your views here.

class AddTaskAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def post(self, request):
        result = task_obj.AddTask(request.data)
        return result


class EidtTaskAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def put(self, request):
        result = task_obj.EditTask(request.data)
        print(result)
        return result


class GetTaskListAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def get(self, request):
        result = task_obj.GetTaskList(request)
        return result


class GetDocumentExcelAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def get(self, request):
        result = task_obj.GetDocumentExcelList(request)
        return result


class GetDocumentPDFAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def get(self, request):
        result = task_obj.GetDocumentPDFList(request)
        return result


class OcrPdfAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def post(self, request):
        result = task_obj.OcrPDF(request.data)
        return result


class GetOcrDataAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def get(self, request):
        result = task_obj.getOcrData(request.data)
        return result


class GetTaskListHistoryAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def get(self, request):
        result = task_obj.GetTaskListHistory(request)
        return result

class SendIntimationAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def post(self, request):
        result = task_obj.sendIntimation(request.data)
        return result

class GetTaskNotificationsAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def get(self, request):
        result = task_obj.getTaskNotifications(request)
        return result


class GetNotifiedTaskAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def get(self, request):
        result = task_obj.getNotifiedTask(request)
        return result