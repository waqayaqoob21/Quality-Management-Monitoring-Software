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
        print("Update request",request)
        try:
            emc_id = request['emc_id']
            is_emc = Emc.objects.filter(id = emc_id).first()
            if is_emc is None:
                product_tests = request['product_tests']
                emc_records = json.loads(product_tests)
                for item in emc_records:
                    emc_obj = Emc()
                    emc_obj.org_type = item["organization"]
                    emc_obj.sys_type = item["sys_type"]
                    emc_obj.sys_name = item["sys_name"]
                    emc_obj.module_name = item['module_name']
                    emc_obj.emi_emc_id = item['emi_emc_id']
                    emc_obj.lot_no = item['lot_no']
                    emc_obj.design_version = item['design_version']
                    emc_obj.test_requirements = item["test_requirements"]
                    emc_obj.ce_101 = item["ce_101"]
                    emc_obj.ce_102 = item["ce_102"]
                    emc_obj.ce_106 = item["ce_106"]
                    emc_obj.re_101 = item["re_101"]
                    emc_obj.re_102 = item["re_102"]
                    emc_obj.re_103 = item["re_103"]
                    emc_obj.cs_101 = item["cs_101"]
                    emc_obj.cs_103 = item["cs_103"]
                    emc_obj.cs_104 = item["cs_104"]
                    emc_obj.cs_105 = item["cs_105"]
                    emc_obj.cs_109 = item["cs_109"]
                    emc_obj.cs_114 = item["cs_114"]
                    emc_obj.cs_115 = item["cs_115"]
                    emc_obj.cs_116 = item["cs_116"]
                    emc_obj.cs_117 = item["cs_117"]
                    emc_obj.cs_118 = item["cs_118"]
                    emc_obj.rs_101 = item["rs_101"]
                    emc_obj.rs_103 = item["rs_103"]
                    emc_obj.rs_105 = item["rs_105"]
                    emc_obj.test_conducted = item["test_conducted"]
                    emc_obj.compliance_status = item["compliance_status"]
                    emc_obj.report_status = item['report_status']
                    emc_obj.compliance_date = item["compliance_date"]
                    emc_obj.remarks = item["remarks"]
                    emc_obj.save()
                    print("record saved successfully!")
                return JsonResponse({'message': 'EMS&ES record inserted Successfully!', 'data': {}}, status=201)
            else:
                get_obj = Emc.objects.filter(id=emc_id).first()
                if (get_obj.org_type != request['organization'] or get_obj.sys_type != request['sys_type'] or get_obj.sys_name != request['system_name'] or \
                    get_obj.module_name != request['module_name'] or \
                    get_obj.lot_no != request['lot_no'] or get_obj.emi_emc_id != request['emi_emc_id'] or \
                    get_obj.design_version != request['design_version'] or get_obj.test_requirements != request['test_requirements'] or \
                    get_obj.compliance_status != request['compliance_status'] or get_obj.compliance_date != request['compliance_date'] or
                    get_obj.report_status != request['report_status'] or \
                    get_obj.remarks != request['remarks'] or get_obj.ce_101 != request['ce_101'] or \
                    get_obj.ce_102 != request['ce_102'] or get_obj.ce_106 != request['ce_106'] or \
                    get_obj.re_101 != request['re_101'] or get_obj.re_102 != request['re_102'] or \
                    get_obj.re_103 != request['re_103'] or get_obj.cs_101 != request['cs_101'] or \
                    get_obj.cs_103 != request['cs_103'] or get_obj.cs_104 != request['cs_104'] or \
                    get_obj.cs_105 != request['cs_105'] or get_obj.cs_109 != request['cs_109'] or get_obj.cs_114 != request['cs_114'] or \
                    get_obj.cs_115 != request['cs_115'] or get_obj.cs_116 != request['cs_116'] or \
                    get_obj.cs_117 != request['cs_117'] or get_obj.cs_118 != request['cs_118'] or \
                    get_obj.rs_101 != request['rs_101'] or get_obj.rs_103 != request['rs_103'] or get_obj.rs_105 != request['rs_105']):

                    EmcHistoryObj = EmcHistory()
                    EmcHistoryObj.emc_id = emc_id
                    EmcHistoryObj.org_type = get_obj.org_type
                    EmcHistoryObj.sys_type = get_obj.sys_type
                    EmcHistoryObj.sys_name = get_obj.sys_name
                    EmcHistoryObj.module_name = get_obj.module_name
                    EmcHistoryObj.emi_emc_id = get_obj.emi_emc_id
                    EmcHistoryObj.lot_no = get_obj.lot_no
                    EmcHistoryObj.design_version = get_obj.design_version
                    EmcHistoryObj.test_requirements = get_obj.test_requirements
                    EmcHistoryObj.ce_101 = get_obj.ce_101
                    EmcHistoryObj.ce_102 = get_obj.ce_102
                    EmcHistoryObj.ce_106 = get_obj.ce_106
                    EmcHistoryObj.re_101 = get_obj.re_101
                    EmcHistoryObj.re_102 = get_obj.re_102
                    EmcHistoryObj.re_103 = get_obj.re_103
                    EmcHistoryObj.cs_101 = get_obj.cs_101
                    EmcHistoryObj.cs_103 = get_obj.cs_103
                    EmcHistoryObj.cs_104 = get_obj.cs_104
                    EmcHistoryObj.cs_105 = get_obj.cs_105
                    EmcHistoryObj.cs_109 = get_obj.cs_109
                    EmcHistoryObj.cs_114 = get_obj.cs_114
                    EmcHistoryObj.cs_115 = get_obj.cs_115
                    EmcHistoryObj.cs_116 = get_obj.cs_116
                    EmcHistoryObj.cs_117 = get_obj.cs_117
                    EmcHistoryObj.cs_118 = get_obj.cs_118
                    EmcHistoryObj.cs_118 = get_obj.cs_118
                    EmcHistoryObj.cs_118 = get_obj.cs_118
                    EmcHistoryObj.rs_105 = get_obj.rs_105
                    EmcHistoryObj.test_conducted = get_obj.test_conducted
                    EmcHistoryObj.compliance_status = get_obj.compliance_status
                    EmcHistoryObj.report_status = get_obj.report_status
                    EmcHistoryObj.compliance_date = get_obj.compliance_date
                    EmcHistoryObj.remarks = get_obj.remarks
                    EmcHistoryObj.save()
                is_emc.org_type = request['organization']
                is_emc.sys_type = request['sys_type']
                is_emc.sys_name = request['system_name']
                is_emc.module_name = request['module_name']
                is_emc.emi_emc_id = request['emi_emc_id']
                is_emc.lot_no = request['lot_no']
                is_emc.design_version = request["design_version"]
                is_emc.test_requirements = request['test_requirements']
                is_emc.ce_101 = request["ce_101"]
                is_emc.ce_102 = request["ce_102"]
                is_emc.ce_106 = request["ce_106"]
                is_emc.re_101 = request["re_101"]
                is_emc.re_102 = request["re_102"]
                is_emc.re_103 = request["re_103"]
                is_emc.cs_101 = request["cs_101"]
                is_emc.cs_103 = request["cs_103"]
                is_emc.cs_104 = request["cs_104"]
                is_emc.cs_105 = request["cs_105"]
                is_emc.cs_109 = request["cs_109"]
                is_emc.cs_114 = request["cs_114"]
                is_emc.cs_115 = request["cs_115"]
                is_emc.cs_116 = request["cs_116"]
                is_emc.cs_117 = request["cs_117"]
                is_emc.cs_118 = request["cs_118"]
                is_emc.rs_101 = request["rs_101"]
                is_emc.rs_103 = request["rs_103"]
                is_emc.rs_105 = request["rs_105"]
                is_emc.test_conducted = request['test_conducted']
                is_emc.compliance_status = request['compliance_status']
                is_emc.report_status = request['report_status']
                is_emc.compliance_date = request['compliance_date']
                is_emc.remarks = request['remarks']
                is_emc.save()
                return JsonResponse({'Success': 'EMS&ES Record Updated Successfully!', 'data': {}}, status=201)
        except Exception as e:
            print(e)
            return JsonResponse({'message':'Record could not add.','data':[],'success':False,'staus':'500'},status=500)

    @staticmethod
    def getEmcDashboardCount(request):
        try:
            selected_year = request.query_params.get('selected_year')
            selected_org = request.query_params.get('selected_org')
            selected_type = request.query_params.get('selected_type')
            selected_system = request.query_params.get('selected_system')
            current_status = request.query_params.get('current_status')
            systemItems = ""
            noVal = ['']
            if selected_system != noVal:
                systemItems = selected_system.split(',')
                # for item in systemItems:
                #     systemItems = item.split(',')

            def get_filter(field_name, filter_condition, filter_value):
                if filter_value is None:
                    return Q()
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
                if filter_condition.strip() == "iregex":
                    kwargs = {
                        '{0}__iregex'.format(field_name): filter_value
                    }
                    return Q(**kwargs)
            typeQuery = Q()
            filter_objects = Q()
            filter_objects_2 = Q()
            # if selected_year != '':
            #     filter_objects &= get_filter('compliance_date__year', 'equal',selected_year)
            # if selected_system != '':
            #     filter_objects &= get_filter('sys_name', 'contains',selected_system)

            # current_year_modules = 0
            # current_year_compliant_modules = 0
            # current_year_non_compliant_modules = 0

            # total_modules = 0
            # total_compliant_modules = 0
            # total_non_compliant_modules = 0
            # dataList = Emc.objects.filter(filter_objects)

            # if selected_system != "":
            #     filter_objects &= get_filter('sys_name', 'contains', selected_system)

            # total_modules = Emc.objects.filter(filter_objects).count()
            # total_compliant_modules = dataList.filter(filter_objects, compliance_status = 'Compliant').count()
            # total_non_compliant_modules = dataList.filter(filter_objects, compliance_status = 'Non Compliant').count()
            # total_partial_compliant_modules = dataList.filter(filter_objects, compliance_status = 'Partial Compliant').count()


            if selected_org != '':
                filter_objects &= get_filter('org_type', 'equal',selected_org)

            current_year_modules = 0
            current_year_compliant_modules = 0
            current_year_non_compliant_modules = 0

            total_modules_2 = 0
            total_compliant_modules_2 = 0
            total_non_compliant_modules_2 = 0
            dataList = Emc.objects.filter(filter_objects)

            if selected_org != "":
                filter_objects &= get_filter('org_type', 'equal', selected_org)

            total_modules_2 = Emc.objects.filter(filter_objects).count()
            total_compliant_modules_2 = dataList.filter(filter_objects, compliance_status = 'Compliant').count()
            total_non_compliant_modules_2 = dataList.filter(filter_objects, compliance_status = 'Non Compliant').count()
            total_partial_compliant_modules_2 = dataList.filter(filter_objects, compliance_status = 'Partial Compliant').count()

            if selected_type != '':
                filter_objects &= get_filter('sys_type', 'equal',selected_type)

            current_year_modules = 0
            current_year_compliant_modules = 0
            current_year_non_compliant_modules = 0

            total_modules = 0
            total_compliant_modules = 0
            total_non_compliant_modules = 0
            dataList = Emc.objects.filter(filter_objects)

            if selected_type != "":
                filter_objects &= get_filter('sys_type', 'equal', selected_type)

            total_modules = Emc.objects.filter(filter_objects).count()
            total_compliant_modules = dataList.filter(filter_objects, compliance_status = 'Compliant').count()
            total_non_compliant_modules = dataList.filter(filter_objects, compliance_status = 'Non Compliant').count()
            total_partial_compliant_modules = dataList.filter(filter_objects, compliance_status = 'Partial Compliant').count()


            # if systemItems != '':
            #     for system in systemItems:
            #      filter_objects &= get_filter('sys_name', 'contains',system)

            if systemItems and any(system for system in systemItems if system and system.strip()):
                for system in systemItems:
                    if system == 'S2':
                        # Use a regex to match exact 'S2' and exclude 'S2B'
                        filter_objects &= get_filter('sys_name', 'iregex', r'^(?!S2B$)S2$')
                    if system == 'S1':
                        filter_objects &= get_filter('sys_name', 'iregex', r'^(?!S1A$)S1$')
                    else:
                        filter_objects &= get_filter('sys_name', 'contains', system)

            current_year_modules = 0
            current_year_compliant_modules = 0
            current_year_non_compliant_modules = 0

            total_modules = 0
            total_compliant_modules = 0
            total_non_compliant_modules = 0
            dataList = Emc.objects.filter(filter_objects)

            # if selected_system != "":
            #     filter_objects &= get_filter('sys_name', 'contains', selected_system)

            # if systemItems != "":
            #     for system in systemItems:
            #      filter_objects &= get_filter('sys_name', 'contains',system)


            total_modules = Emc.objects.filter(filter_objects).count()
            total_compliant_modules = dataList.filter(filter_objects, compliance_status = 'Compliant').count()
            total_non_compliant_modules = dataList.filter(filter_objects, compliance_status = 'Non Compliant').count()
            total_partial_compliant_modules = dataList.filter(filter_objects, compliance_status = 'Partial Compliant').count()

            if selected_year != '':
                filter_objects &= get_filter('compliance_date__year', 'equal',selected_year)
            current_year_modules = Emc.objects.filter(filter_objects).count()
            current_year_compliant_modules = Emc.objects.filter(filter_objects,compliance_status = 'Compliant').count()
            current_year_non_compliant_modules = Emc.objects.filter(filter_objects,compliance_status = 'Non Compliant').count()
            current_year_partial_compliant_modules = Emc.objects.filter(filter_objects,compliance_status = 'Partial Compliant').count()


            # if selected_year != '':
            #     filter_objects_2 &= get_filter('created_at__year', 'equal',selected_year)
            # current_year_modules = Emc.objects.filter(filter_objects_2).count()
            # current_year_compliant_modules = Emc.objects.filter(filter_objects_2,compliance_status = 'Compliant').count()
            # current_year_non_compliant_modules = Emc.objects.filter(filter_objects_2,compliance_status = 'Non Compliant').count()
            # current_year_partial_compliant_modules = Emc.objects.filter(filter_objects_2,compliance_status = 'Partial Compliant').count()



            dict = {
                'current_year_modules': current_year_modules,
                'current_year_compliant_modules': current_year_compliant_modules,
                'current_year_non_compliant_modules': current_year_non_compliant_modules,
                'current_year_partial_compliant_modules': current_year_partial_compliant_modules,
                # 'total_modules': total_modules,
                # 'total_compliant_modules': total_compliant_modules,
                # 'total_non_compliant_modules': total_non_compliant_modules,
                # 'total_partial_compliant_modules': total_partial_compliant_modules,
                'total_modules_2': total_modules_2,
                'total_compliant_modules_2': total_compliant_modules_2,
                'total_non_compliant_modules_2': total_non_compliant_modules_2,
                'total_partial_compliant_modules_2': total_partial_compliant_modules_2
            }
            print("res dict",dict)
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
            selected_system = request.query_params.get('selected_system')
            current_status = request.query_params.get('selected_status')
            systemItems = ""
            noVal = ['']
            if selected_system != noVal:
                systemItems = selected_system.split(',')
            def get_filter(field_name, filter_condition, filter_value):
                if filter_value is None:
                    return Q()
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
                if filter_condition.strip() == "iregex":
                    kwargs = {
                        '{0}__iregex'.format(field_name): filter_value
                    }
                    return Q(**kwargs)
            typeQuery = Q()
            filter_objects = Q()

            if selected_org != '':
                filter_objects &= get_filter('org_type', 'equal',selected_org)

            if selected_org != "":
                filter_objects &= get_filter('org_type', 'equal', selected_org)

            if selected_type != '':
                filter_objects &= get_filter('sys_type', 'equal',selected_type)

            if selected_type != "":
                filter_objects &= get_filter('sys_type', 'equal', selected_type)

            # if selected_system != '':
            #     filter_objects &= get_filter('sys_name', 'contains',selected_system)

            if systemItems and any(system for system in systemItems if system and system.strip()):
                for system in systemItems:
                    if system == 'S2':
                        # Use a regex to match exact 'S2' and exclude 'S2B'
                        filter_objects &= get_filter('sys_name', 'iregex', r'^(?!S2B$)S2$')
                    if system == 'S1':
                        filter_objects &= get_filter('sys_name', 'iregex', r'^(?!S1A$)S1$')
                    else:
                        filter_objects &= get_filter('sys_name', 'contains', system)

            # if selected_system != "":
            #     filter_objects &= get_filter('sys_name', 'contains', selected_system)

            dataList = Emc.objects.all().order_by('-id')
            if current_status == 'total_modules':
                dataList = dataList.filter(filter_objects)
            if current_status == 'total_compliant_modules':
                dataList = dataList.filter(filter_objects, compliance_status = 'Compliant')
            if current_status == 'total_non_compliant_modules':
                dataList = dataList.filter(filter_objects, compliance_status = 'Non Compliant')
            if current_status == 'total_partial_compliant_modules':
                dataList = dataList.filter(filter_objects, compliance_status = 'Partial Compliant')
            if current_status == 'total_systems':
                dataList = dataList.filter(filter_objects)
            if current_status == 'total_compliant_system':
                dataList = dataList.filter(filter_objects, compliance_status = 'Compliant')
            if current_status == 'total_non_compliant_system':
                dataList = dataList.filter(filter_objects, compliance_status = 'Non Compliant')
            if current_status == 'total_partial_compliant_system':
                dataList = dataList.filter(filter_objects, compliance_status = 'Partial Compliant')    

            if selected_year != '':
                filter_objects &= get_filter('compliance_date__year', 'equal',selected_year)
            if current_status == 'current_year_modules':
                dataList = dataList.filter(filter_objects)
            if current_status == 'current_year_compliant_modules':
                dataList = dataList.filter(filter_objects, compliance_status = 'Compliant')
            if current_status == 'current_year_non_compliant_modules':
                dataList = dataList.filter(filter_objects, compliance_status = 'Non Compliant')
            if current_status == 'current_year_partial_compliant_modules':
                dataList = dataList.filter(filter_objects, compliance_status = 'Partial Compliant')
            serializer = EmcSerializer(dataList, many=True)
            return JsonResponse({'Success': 'EMI/EMC List Fetched Successfully!','data':serializer.data,'success':True,'status':'200'},status=200)
        except Exception as e:
            print(e)
            return JsonResponse({'message':'Record could not add.','data':[],'success':False,'staus':'500'},status=500)

    @staticmethod
    def getEmcHistory(request):
        try:
            emc_id = request.query_params.get('id')
            data = EmcHistory.objects.filter(emc_id = emc_id)
            serializer = EmcHistorySerializer(data, many=True)
            print("history" , serializer.data)
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