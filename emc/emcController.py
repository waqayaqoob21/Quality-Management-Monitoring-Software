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
                    emc_obj.organization = item["organization"]
                    emc_obj.set_id = item['set_id']
                    emc_obj.module_name = item['module_name']
                    emc_obj.emi_emc_id = item['emi_emc_id']
                    testList = []
                    emc_obj.test_requirements = item["test_requirements"]
                    testList = item["test_requirements"].split(",")
                    print(testList)
                    for test in testList:
                        if test == "CE101":
                            emc_obj.ce_101 = item["test_conducted"]
                        if test == "CE102":
                            emc_obj.ce_102 = item["test_conducted"]
                        if test == "CE106":
                            emc_obj.ce_106 = item["test_conducted"]
                        if test == "RE101":
                            emc_obj.re_101 = item["test_conducted"]
                        if test == "RE102":
                            emc_obj.re_102 = item["test_conducted"]
                        if test == "RE103":
                            emc_obj.re_103 = item["test_conducted"]
                        if test == "CS101":
                            emc_obj.cs_101 = item["test_conducted"]
                        if test == "CS103":
                            emc_obj.cs_103 = item["test_conducted"]
                        if test == "CS104":
                            emc_obj.cs_104 = item["test_conducted"]
                        if test == "CS105":
                            emc_obj.cs_105 = item["test_conducted"]
                        if test == "CS109":
                            emc_obj.cs_109 = item["test_conducted"]
                        if test == "CS114":
                            emc_obj.cs_114 = item["test_conducted"]
                        if test == "CS115":
                            emc_obj.cs_115 = item["test_conducted"]
                        if test == "CS116":
                            emc_obj.cs_116 = item["test_conducted"]
                        if test == "CS117":
                            emc_obj.cs_117 = item["test_conducted"]
                        if test == "CS118":
                            emc_obj.cs_118 = item["test_conducted"]
                        if test == "RS101":
                            emc_obj.rs_101 = item["test_conducted"]
                        if test == "RS103":
                            emc_obj.rs_103 = item["test_conducted"]
                        if test == "RS105":
                            emc_obj.rs_105 = item["test_conducted"]
                    emc_obj.test_conducted = item["test_conducted"]
                    emc_obj.compliance_status = item["compliance_status"]
                    emc_obj.report_status = item['report_status']
                    emc_obj.remarks = item["remarks"]
                    emc_obj.save()
                    print("record saved successfully!")
                return JsonResponse({'Success': 'EMS&ES record inserted Successfully!'})
            else:
                is_emc.sys_type = request['sys_type']
                is_emc.sys_name = request['system_name']
                is_emc.organization = request["organization"]
                is_emc.set_id = request['set_id']
                is_emc.module_name = request['module_name']
                is_emc.emi_emc_id = request['emi_emc_id']
                testList = []
                ce_101 = ""
                ce_102 = ""
                ce_106 = ""
                re_101 = ""
                re_102 = ""
                re_103 = ""
                cs_101 = ""
                cs_103 = ""
                cs_104 = ""
                cs_105 = ""
                cs_109 = ""
                cs_114 = ""
                cs_115 = ""
                cs_116 = ""
                cs_117 = ""
                cs_118 = ""
                rs_101 = ""
                rs_103 = ""
                rs_105 = ""
                is_emc.test_requirements = request["test_requirements"]
                testList = request["test_requirements"].split(",")
                print(testList)
                for test in testList:
                    if test == "CE101":
                        ce_101 = request["test_conducted"]
                    if test == "CE102":
                        ce_102 = request["test_conducted"]
                    if test == "CE106":
                        ce_106 = request["test_conducted"]
                    if test == "RE101":
                        re_101 = request["test_conducted"]
                    if test == "RE102":
                        re_102 = request["test_conducted"]
                    if test == "RE103":
                        re_103 = request["test_conducted"]
                    if test == "CS101":
                        cs_101 = request["test_conducted"]
                    if test == "CS103":
                        cs_103 = request["test_conducted"]
                    if test == "CS104":
                        cs_104 = request["test_conducted"]
                    if test == "CS105":
                        cs_105 = request["test_conducted"]
                    if test == "CS109":
                        cs_109 = request["test_conducted"]
                    if test == "CS114":
                        cs_114 = request["test_conducted"]
                    if test == "CS115":
                        cs_115 = request["test_conducted"]
                    if test == "CS116":
                        cs_116 = request["test_conducted"]
                    if test == "CS117":
                        cs_117 = request["test_conducted"]
                    if test == "CS118":
                        cs_118 = request["test_conducted"]
                    if test == "RS101":
                        rs_101 = request["test_conducted"]
                    if test == "RS103":
                        rs_103 = request["test_conducted"]
                    if test == "RS105":
                        rs_105 = request["test_conducted"]
                is_emc.ce_101 = ce_101
                is_emc.ce_102 = ce_102
                is_emc.ce_106 = ce_106
                is_emc.re_101 = re_101
                is_emc.re_102 = re_102
                is_emc.re_103 = re_103
                is_emc.cs_101 = cs_101
                is_emc.cs_103 = cs_103
                is_emc.cs_104 = cs_104
                is_emc.cs_105 = cs_105
                is_emc.cs_109 = cs_109
                is_emc.cs_114 = cs_114
                is_emc.cs_115 = cs_115
                is_emc.cs_116 = cs_116
                is_emc.cs_117 = cs_117
                is_emc.cs_118 = cs_118
                is_emc.rs_101 = rs_101
                is_emc.rs_103 = rs_103
                is_emc.rs_105 = rs_105
                is_emc.test_conducted = request['test_conducted']
                is_emc.compliance_status = request['compliance_status']
                is_emc.report_status = request['report_status']
                is_emc.remarks = request['remarks']
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
            total_compliant_modules = dataList.filter(filter_objects, compliance_status = 'Compliant Modules').count()
            total_non_compliant_modules = dataList.filter(filter_objects, compliance_status = 'Non Compliant Modules').count()

            if selected_year != '':
                filter_objects &= get_filter('created_at__year', 'equal',selected_year)
            current_year_modules = Emc.objects.filter(filter_objects).count()
            current_year_compliant_modules = Emc.objects.filter(compliance_status = 'Compliant Modules').count()
            current_year_non_compliant_modules = Emc.objects.filter(compliance_status = 'Non Compliant Modules').count()



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
    def getEmcList(request):
        try:
            selected_year = request.query_params.get('selected_year')
            selected_org = request.query_params.get('selected_org')
            selected_type = request.query_params.get('selected_type')
            selected_sys = request.query_params.get('selected_system')
            current_status = request.query_params.get('selected_status')
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

            if selected_org != '':
                filter_objects &= get_filter('organization', 'equal',selected_org)
            if selected_type != "":
                filter_objects &= get_filter('sys_type', 'equal', selected_type)
            if selected_sys != "":
                filter_objects &= get_filter('sys_name', 'equal', selected_sys)

            dataList = Emc.objects.all().order_by('-id')
            if current_status == 'total_modules':
                dataList = dataList.filter(filter_objects)
            if current_status == 'total_compliant_modules':
                dataList = dataList.filter(filter_objects, compliance_status = 'Compliant Modules')
            if current_status == 'total_non_compliant_modules':
                dataList = dataList.filter(filter_objects, compliance_status = 'Non Compliant Modules')

            if current_status == 'current_year_modules':
                dataList = dataList.filter(filter_objects, created_at__year = selected_year)
            if current_status == 'current_year_compliant_modules':
                dataList = dataList.filter(filter_objects, created_at__year = selected_year,compliance_status = 'Compliant Modules')
            if current_status == 'current_year_non_compliant_modules':
                dataList = dataList.filter(filter_objects, created_at__year = selected_year,compliance_status = 'Non Compliant Modules')
            serializer = EmcSerializer(dataList, many=True)
            return JsonResponse({'Success': 'EMI/EMC List Fetched Successfully!','data':serializer.data,'success':True,'status':'200'},status=200)
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