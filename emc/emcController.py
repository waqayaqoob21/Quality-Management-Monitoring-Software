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
                    emc_obj.organization = item["organization"]
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
                is_emc.organization = request["organization"]
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
    def getEmcDashboardCount(request):
        try:
            selected_year = request.query_params.get('selected_year')
            selected_org = request.query_params.get('selected_org')
            selected_type = request.query_params.get('selected_type')
            selected_sys = request.query_params.get('selected_system')
            current_status = request.query_params.get('current_status')
            def get_filter(field_name, filter_condition, filter_value):
                if filter_condition.strip() == "contains":
                    kwargs = {
                        '{0}__icontains'.format(field_name): filter_value
                    }
                    return Q(**kwargs)

                if filter_condition.strip() == "not_equal":
                    kwargs = {
                        '{0}__iexact'.format(field_name): filter_value
                    }
                    return ~Q(**kwargs)

                if filter_condition.strip() == "starts_with":
                    kwargs = {
                        '{0}__istartswith'.format(field_name): filter_value
                    }
                    return Q(**kwargs)
                if filter_condition.strip() == "equal":
                    kwargs = {
                        '{0}__iexact'.format(field_name): filter_value
                    }
                    return Q(**kwargs)

                if filter_condition.strip() == "not_equal":
                    kwargs = {
                        '{0}__iexact'.format(field_name): filter_value
                    }

                    return ~Q(**kwargs)
            typeQuery = Q()
            filter_objects = Q()
            if selected_year != '':
                filter_objects &= get_filter('created_at__year', 'equal',selected_year)
            if selected_org != '':
                filter_objects &= get_filter('organization', 'equal',selected_org)
            if selected_type != '':
                filter_objects &= get_filter('sys_type', 'equal',selected_type)
            if selected_sys != '':
                filter_objects &= get_filter('sys_name', 'equal',selected_sys)
            current_year_modules = 0
            current_year_compliant_modules = 0
            current_year_non_compliant_modules = 0

            total_modules = 0
            total_compliant_modules = 0
            total_non_compliant_modules = 0
            dataList = Emc.objects.filter(filter_objects)

            if selected_org != '':
                filter_objects &= get_filter('organization', 'equal',selected_org)
            if selected_type != "":
                filter_objects &= get_filter('sys_type', 'equal', selected_type)
            if selected_sys != "":
                filter_objects &= get_filter('sys_name', 'equal', selected_sys)

            total_modules = Emc.objects.filter(filter_objects).count()
            total_compliant_modules = dataList.filter(filter_objects, selected_status = 'Compliant Modules').count()
            total_non_compliant_modules = dataList.filter(filter_objects, selected_status = 'Non Compliant Modules').count()

            if selected_year != '':
                filter_objects &= get_filter('created_at__year', 'equal',selected_year)
            current_year_modules = Emc.objects.filter(filter_objects).count()
            current_year_compliant_modules = Emc.objects.filter(selected_status = 'Compliant Modules').count()
            current_year_non_compliant_modules = Emc.objects.filter(selected_status = 'Non Compliant Modules').count()



            dict = {
                'current_year_modules': current_year_modules,
                'current_year_compliant_modules': current_year_compliant_modules,
                'current_year_non_compliant_modules': current_year_non_compliant_modules,
                'total_modules': total_modules,
                'total_compliant_modules': total_compliant_modules,
                'total_non_compliant_modules': total_non_compliant_modules
            }
            return JsonResponse({'Success': 'EMS&ES List Fetched Successfully!','data':dict,'success':True,'status':'200'},status=200)
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