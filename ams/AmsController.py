from django.http import JsonResponse
from ams.serializer import *
from ams.models import *


class AmsController:

    @staticmethod
    def AddTask(request):
        taskModel = TaskSummary()
        try:
            taskModel.task = request['task']
            taskModel.assigned_by = request['assigned_by']
            taskModel.assigned_date_by = request['assigned_date_by']
            taskModel.target_date_by = request['target_date_by']
            taskModel.assigned_to = request['assigned_to']
            taskModel.assigned_date_to = request['assigned_date_to']
            taskModel.target_date_to = request['target_date_to']
            taskModel.status = request['status']
            taskModel.attachment = request['attachment']
            if request['isActive'] == '1':
                taskModel.isActive = 'True'
            else:
                taskModel.isActive = 'False'
            taskModel.save()
            return JsonResponse({'Success': "Task Created Successfully"},status=200)
        except Exception as e:
            print(e)
            return JsonResponse({"Error": "Task Not Saved"}, status=500)

    @staticmethod
    def EditTask(request): 
        try:
            pk = 1
            task = TaskSummary.objects.get(id=pk)
            task.task = request['task']
            task.assigned_by = request['assigned_by']
            task.assigned_date_by = request['assigned_date_by']
            task.target_date_by = request['target_date_by']
            task.assigned_to = request['assigned_to']
            task.assigned_date_to = request['assigned_date_to']
            task.target_date_to = request['target_date_to']
            task.status = request['status']
            task.attachment = request['attachment']
            if request['isActive'] == '1':
                task.isActive = 'True'
            else:
                task.isActive = 'False'
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

    