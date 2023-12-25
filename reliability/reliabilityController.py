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
            rel_obj = Reliability()
            id = request['id']
            if id == '0':
                rel_obj.sys_type = request['sys_type']
                rel_obj.sys_name = request['system_name']
                rel_obj.items = request['items']
                rel_obj.age_method = request['age_method']
                rel_obj.age_estimation = request['age_estimation']
                rel_obj.age_remarks = request['age_remarks']
                rel_obj.reliability_estimation = request['reliability_estimation']
                rel_obj.reliability_estimation_method = request['reliability_estimation_method']
                rel_obj.reliability_estimation_remarks = request['reliability_estimation_remarks']
                rel_obj.save()
                return JsonResponse({'Success': 'EMS&ES record inserted Successfully!'})
            else:
                get_obj = Reliability.objects.filter(id=id).first()
                get_obj.sys_type = request['sys_type']
                get_obj.sys_name = request['system_name']
                get_obj.items = request['items']
                get_obj.age_method = request['age_method']
                get_obj.age_estimation = request['age_estimation']
                get_obj.age_remarks = request['age_remarks']
                get_obj.reliability_estimation = request['reliability_estimation']
                get_obj.reliability_estimation_method = request['reliability_estimation_method']
                get_obj.reliability_estimation_remarks = request['reliability_estimation_remarks']
                get_obj.save()
                return JsonResponse({'Success': 'EMS&ES Record Updated Successfully!'})
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
                {'Success': 'EMS&ES List Fetched Successfully!', 'data': serializer.data, 'success': True,
                 'status': '200'}, status=200)
        except Exception as e:
            print(e)
            return JsonResponse({'message': 'Record could not add.', 'data': [], 'success': False, 'staus': '500'},
                                status=500)