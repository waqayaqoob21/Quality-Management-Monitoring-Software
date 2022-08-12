from django.http import JsonResponse
from ams.serializer import *
from ams.models import *


class AmsController:

    @staticmethod
    def AddTask(request):
        taskModel = TaskSummary()
        try:
            taskModel.task_name = request['task_name']
            taskModel.assigned_by = request['assigned_by']
            taskModel.assigned_date = request['assigned_date']
            taskModel.assigned_to = request['assigned_to']
            taskModel.target_date = request['target_date']
            taskModel.status = request['status']
            taskModel.remarks = request['remarks']
            taskModel.save()
            return JsonResponse({'Success': "Task Created Successfully"},status=200)
        except Exception as e:
            print(e)
            return JsonResponse({"Error": "Task Not Saved"}, status=500)

    @staticmethod
    def EditTask(request): 
        try:
            task = TaskSummary.objects.get(id=request['id'])
            task.task_name = request['task_name']
            task.assigned_by = request['assigned_by']
            task.assigned_date = request['assigned_date']
            task.assigned_to = request['assigned_to']
            task.target_date = request['target_date']
            task.status = request['status']
            task.remarks = request['remarks']
            task.save()
            return JsonResponse({'Success':'Task Updated Successfully!'}, status=201)
        except:
            return JsonResponse({'Error':'Task Could Not Update Successfully!'}, status=400)

    @staticmethod
    def GetTaskList(request):
        try:   
            data = TaskSummary.objects.all()
            serializer = TaskSummarySerialzer(data, many=True)
            return JsonResponse({'message':'Welcome to Home Page','Task List': serializer.data}, status=200)
        except:
            return JsonResponse({'message':'Sorry! No Task found.'}, status=200)

    