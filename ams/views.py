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
        result = task_obj.GetTaskList(request.data)
        return result
