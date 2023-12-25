from django.http import JsonResponse
from fpdf import FPDF
import xlwt
from datetime import datetime, timedelta
from django.http import FileResponse
from django.http import HttpResponse
from django.core.files.storage import FileSystemStorage
from django.db.models import F, Q
import itertools
from ems_es.models import *
from ems_es.serializers import *
from datetime import date
from django.db.models.functions import Extract
from datetime import datetime
from datetime import timedelta
from dateutil.relativedelta import relativedelta
class emsController:
    @staticmethod
    def addEmsEs(request):
        try:
            ems_obj = EmsEs()
            id = request['id']
            if id == '0':
                ems_obj.sys_type = request['sys_type']
                ems_obj.sys_name = request['system_name']
                ems_obj.product = request['product']
                ems_obj.test = request['test']
                ems_obj.platform = request['platform']
                ems_obj.test_requirements = request['test_requirements']
                ems_obj.test_conducted = request['test_conducted']
                ems_obj.test_compliance = request['test_compliance']
                ems_obj.selected_status = request['selected_status']
                ems_obj.report_status = request['report_status']
                ems_obj.module = request['module']
                ems_obj.compliance_status = request['compliance_status']
                ems_obj.save()
                return JsonResponse({'Success': 'EMS&ES record inserted Successfully!'})
            else:
                get_obj = EmsEs.objects.filter(id=id).first()
                get_obj.sys_type = request['sys_type']
                get_obj.sys_name = request['system_name']
                get_obj.product = request['product']
                get_obj.test = request['test']
                get_obj.platform = request['platform']
                get_obj.test_requirements = request['test_requirements']
                get_obj.test_conducted = request['test_conducted']
                get_obj.test_compliance = request['test_compliance']
                get_obj.selected_status = request['selected_status']
                get_obj.report_status = request['report_status']
                get_obj.module = request['module']
                get_obj.compliance_status = request['compliance_status']
                get_obj.save()
                return JsonResponse({'Success': 'EMS&ES Record Updated Successfully!'})
        except Exception as e:
            print(e)
            return JsonResponse({'message':'Record could not add.','data':[],'success':False,'staus':'500'},status=500)

    @staticmethod
    def getEmsEs(request):
        try:
            data = EmsEs.objects.all()
            serializer = EmsEsSerializer(data, many=True)
            return JsonResponse({'Success': 'EMS&ES List Fetched Successfully!','data':serializer.data,'success':True,'status':'200'},status=200)
        except Exception as e:
            print(e)
            return JsonResponse({'message':'Record could not add.','data':[],'success':False,'staus':'500'},status=500)