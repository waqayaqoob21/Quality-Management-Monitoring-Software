import json

from django.http import JsonResponse
from fpdf import FPDF
import xlwt
from datetime import datetime, timedelta
from django.http import FileResponse
from django.http import HttpResponse
from django.core.files.storage import FileSystemStorage
from django.db.models import F, Q
import itertools
from .models import *
from .serializers import *
from datetime import date
from django.db.models.functions import Extract
from datetime import datetime
from datetime import timedelta
from dateutil.relativedelta import relativedelta
class emcController:

    @staticmethod
    def getProduct(request):
        try:
            emc_id = request.query_params.get('emc_id')
            data = Emc.objects.filter(id = emc_id).first()
            serializer = EmcSerializer(data)
            return JsonResponse({'Success': 'EMS&ES List Fetched Successfully!','data':serializer.data,'success':True,'status':'200'},status=200)
        except Exception as e:
            print(e)
            return JsonResponse({'message':'Record could not add.','data':[],'success':False,'staus':'500'},status=500)
    @staticmethod
    def addEmc(request):
        try:
            emc_id = request['emc_id']
            is_emc = Emc.objects.filter(id = emc_id).first()
            if is_emc is None:
                product_tests = request['product_tests']
                emc_records = json.loads(product_tests)
                for item in emc_records:
                    emc_obj = Emc()
                    emc_obj.sys_type = item["sys_type"]
                    emc_obj.sys_name = item["sys_name"]
                    emc_obj.product = item["product"]
                    emc_obj.test = item["test"]
                    emc_obj.platform = item["plateform"]
                    emc_obj.test_requirements = item["test_requirements"]
                    emc_obj.test_conducted = item["test_conducted"]
                    emc_obj.test_compliance = item["test_compliance"]
                    emc_obj.selected_status = item["selected_status"]
                    emc_obj.report_status = item['report_status']
                    # emc_obj.module = item["module"]
                    emc_obj.compliance_status = item["compliance_status"]
                    emc_obj.due_date = item["due_date"]
                    emc_obj.audit_completion_date = item["emc_completion_date"]
                    emc_obj.save()
                    print("record saved successfully!")
                return JsonResponse({'Success': 'EMS&ES record inserted Successfully!'})
            else:
                is_emc.sys_type = request['sys_type']
                is_emc.sys_name = request['system_name']
                is_emc.product = request['product']
                is_emc.test = request['test']
                is_emc.test_requirements = request['test_requirements']
                is_emc.test_conducted = request['test_conducted']
                is_emc.test_compliance = request['test_compliance']
                is_emc.selected_status = request['selected_status']
                is_emc.report_status = request['report_status']
                # is_emc.module = request['module']
                is_emc.compliance_status = request['compliance_status']
                is_emc.due_date = request['due_date']
                is_emc.audit_completion_date = request['audit_completion_date']
                is_emc.save()
                return JsonResponse({'Success': 'EMS&ES Record Updated Successfully!'})
        except Exception as e:
            print(e)
            return JsonResponse({'message':'Record could not add.','data':[],'success':False,'staus':'500'},status=500)

    @staticmethod
    def getEmc(request):
        try:
            data = Emc.objects.all()
            serializer = EmcSerializer(data, many=True)
            return JsonResponse({'Success': 'EMS&ES List Fetched Successfully!','data':serializer.data,'success':True,'status':'200'},status=200)
        except Exception as e:
            print(e)
            return JsonResponse({'message':'Record could not add.','data':[],'success':False,'staus':'500'},status=500)

    @staticmethod
    def deleteEmc(request):
        try:
            emc_id = request.query_params['emc_id']
            emc = Emc.objects.filter(id=emc_id).first()
            emc.delete()
            return JsonResponse({'Success': 'EMC deleted successfully', 'data':[],'success': True,'status':'201'},status=201)
        except Exception as e:
            print(e)
            return JsonResponse({'message': 'EMC could not delete.', 'data': [], 'success': False, 'staus': '500'},
                                status=500)