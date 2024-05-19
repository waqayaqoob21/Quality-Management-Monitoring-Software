from django.http import JsonResponse
from fpdf import FPDF
import xlwt
from datetime import datetime, timedelta
from django.http import FileResponse
from django.http import HttpResponse
from django.core.files.storage import FileSystemStorage
from django.db.models import F, Q
import itertools
from reliability.models import *
from reliability.serializers import *
from datetime import date
from django.db.models.functions import Extract
from datetime import datetime
from datetime import timedelta
from dateutil.relativedelta import relativedelta
class reliabilityController:
    @staticmethod
    def addReliability(request):
        try:
            id = request['id']
            get_obj = Reliability.objects.filter(id=id).first()
            if get_obj is None:
                rel_obj = Reliability()
                rel_obj.system_type = request['system_type']
                rel_obj.module_name = request['module_name']
                rel_obj.estimation_type = request['estimation_type']
                rel_obj.estimation_method = request['estimation_method']
                rel_obj.request_date = request['request_date']
                rel_obj.due_date = request['due_date']
                rel_obj.completion_date = request['completion_date']
                rel_obj.current_status = request['current_status']
                rel_obj.remarks = request['remarks']
                rel_obj.save()
                return JsonResponse({'Success': 'Reliability record inserted Successfully!'})
            else:
                if get_obj.system_type != request['system_type'] or get_obj.module_name != request['module_name'] or \
                get_obj.estimation_method != request['estimation_method'] or str(get_obj.request_date) != request['request_date'] or \
                str(get_obj.due_date) != request['due_date'] or str(get_obj.completion) != request['completion'] or \
                get_obj.current_status != request['current_status'] or get_obj.remarks != request['remarks']:
                    reliabilityHistoryModal = ReliabilityHistory()
                    reliabilityHistoryModal.reliability_id = get_obj.id
                    reliabilityHistoryModal.system_type = get_obj.system_type
                    reliabilityHistoryModal.module_name = get_obj.module_name
                    reliabilityHistoryModal.estimation_type = get_obj.estimation_type
                    reliabilityHistoryModal.estimation_method = get_obj.estimation_method
                    reliabilityHistoryModal.request_date = get_obj.request_date
                    reliabilityHistoryModal.due_date = get_obj.due_date
                    if request['completion_date'] != '':
                        reliabilityHistoryModal.completion_date = get_obj.completion_date
                    reliabilityHistoryModal.current_status = get_obj.current_status
                    reliabilityHistoryModal.remarks = get_obj.remarks
                    reliabilityHistoryModal.save()

                get_obj.system_type = request['system_type']
                get_obj.module_name = request['module_name']
                get_obj.estimation_type = request['estimation_type']
                get_obj.estimation_method = request['estimation_method']
                get_obj.request_date = request['request_date']
                get_obj.due_date = request['due_date']
                get_obj.completion_date = request['completion_date']
                get_obj.current_status = request['current_status']
                get_obj.remarks = request['remarks']
                get_obj.save()
                return JsonResponse({'Success': 'Reliability record updated Successfully!'})
        except Exception as e:
            print(e)
            return JsonResponse({'message': 'Record could not add.', 'data': [], 'success': False, 'staus': '500'},
                                status=500)

    @staticmethod
    def getReliabilityList(request):
        try:
            data = Reliability.objects.all()
            serializer = ReliabilitySerializer(data, many=True)
            return JsonResponse(
                {'Success': 'Reliability List Fetched Successfully!', 'data': serializer.data, 'success': True,
                 'status': '200'}, status=200)
        except Exception as e:
            print(e)
            return JsonResponse({'message': 'Record could not add.', 'data': [], 'success': False, 'staus': '500'},
                                status=500)