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



class powerController:

    @staticmethod
    def getPowerToBeUpdated(request):
        try:
            power_id = request.query_params.get('power_id')
            data = Power.objects.filter(id = power_id).first()
            serializer = PowerSerializer(data)
            return JsonResponse({'Success': 'EMS&ES List Fetched Successfully!','data':serializer.data,'success':True,'status':'200'},status=200)
        except Exception as e:
            print(e)
            return JsonResponse({'message':'Record could not add.','data':[],'success':False,'staus':'500'},status=500)
    @staticmethod
    def addPower(request):
        try:
            power_id = request['power_id']
            is_power = Power.objects.filter(id = power_id).first()
            if is_power is None:
                product_tests = request['product_tests']
                power_records = json.loads(product_tests)
                for item in power_records:
                    power_obj = Power()
                    power_obj.org_type = item["organization"]
                    power_obj.system_type = item["sys_type"]
                    power_obj.sys_name = item["sys_name"]
                    power_obj.module_name = item['module_name']
                    power_obj.pc_module_id = item['pc_module_id']
                    power_obj.lot_no = item['lot_no']
                    power_obj.design_version = item['design_version']
                    power_obj.power_compliance_status = item['power_compliance_status']
                    power_obj.pcs_ldc101 = item["pcs_ldc101"]
                    power_obj.pcs_ldc102 = item["pcs_ldc102"]
                    power_obj.pcs_ldc103 = item["pcs_ldc103"]
                    power_obj.pcs_ldc104 = item["pcs_ldc104"]
                    power_obj.pcs_ldc105 = item["pcs_ldc105"]
                    power_obj.pcs_ldc201 = item["pcs_ldc201"]
                    power_obj.pcs_ldc301 = item["pcs_ldc301"]
                    power_obj.pcs_ldc401 = item["pcs_ldc401"]
                    power_obj.pcs_ldc501 = item["pcs_ldc501"]
                    power_obj.pcs_ldc601 = item["pcs_ldc601"]
                    power_obj.pcs_ldc602 = item["pcs_ldc602"]
                    power_obj.compliance_status = item["compliance_status"]
                    power_obj.report_status = item['report_status']
                    power_obj.compliance_date = item["compliance_date"]
                    power_obj.remarks = item["remarks"]
                    power_obj.save()
                    print("record saved successfully!")
                return JsonResponse({'Success': 'Power record inserted Successfully!'})
            else:
                get_obj = Power.objects.filter(id=power_id).first()
                if (get_obj.org_type != request['organization'] or get_obj.system_type != request['sys_type'] or get_obj.sys_name != request['system_name'] or 
                    get_obj.module_name != request['module_name'] or \
                    get_obj.lot_no != request['lot_no'] or get_obj.pc_module_id != request['pc_module_id'] or \
                    get_obj.design_version != request['design_version'] or \
                    get_obj.compliance_status != request['compliance_status'] or get_obj.compliance_date != request['compliance_date'] or
                    get_obj.report_status != request['report_status'] or \
                    get_obj.remarks != request['remarks'] or get_obj.pcs_ldc101 != request['pcs_ldc101'] or \
                    get_obj.pcs_ldc102 != request['pcs_ldc102'] or get_obj.pcs_ldc103 != request['pcs_ldc103'] or \
                    get_obj.pcs_ldc104 != request['pcs_ldc104'] or get_obj.pcs_ldc105 != request['pcs_ldc105'] or \
                    get_obj.pcs_ldc201 != request['pcs_ldc201'] or get_obj.pcs_ldc301 != request['pcs_ldc301'] or \
                    get_obj.pcs_ldc401 != request['pcs_ldc401'] or get_obj.pcs_ldc501 != request['pcs_ldc501'] or \
                    get_obj.pcs_ldc601 != request['pcs_ldc601'] or get_obj.pcs_ldc602 != request['pcs_ldc602']):

                    powerHistoryObj = PowerHistory()
                    powerHistoryObj.power_id = power_id
                    powerHistoryObj.org_type = get_obj.org_type
                    powerHistoryObj.system_type = get_obj.system_type
                    powerHistoryObj.sys_name = get_obj.sys_name
                    powerHistoryObj.module_name = get_obj.module_name
                    powerHistoryObj.pc_module_id = get_obj.pc_module_id
                    powerHistoryObj.lot_no = get_obj.lot_no
                    powerHistoryObj.design_version = get_obj.design_version
                    powerHistoryObj.power_compliance_status = get_obj.power_compliance_status
                    powerHistoryObj.pcs_ldc101 = get_obj.pcs_ldc101
                    powerHistoryObj.pcs_ldc102 = get_obj.pcs_ldc102
                    powerHistoryObj.pcs_ldc103 = get_obj.pcs_ldc103
                    powerHistoryObj.pcs_ldc104 = get_obj.pcs_ldc104
                    powerHistoryObj.pcs_ldc105 = get_obj.pcs_ldc105
                    powerHistoryObj.pcs_ldc201 = get_obj.pcs_ldc201
                    powerHistoryObj.pcs_ldc301 = get_obj.pcs_ldc301
                    powerHistoryObj.pcs_ldc401 = get_obj.pcs_ldc401
                    powerHistoryObj.pcs_ldc501 = get_obj.pcs_ldc501
                    powerHistoryObj.pcs_ldc601 = get_obj.pcs_ldc601
                    powerHistoryObj.pcs_ldc602 = get_obj.pcs_ldc602
                    powerHistoryObj.compliance_status = get_obj.compliance_status
                    powerHistoryObj.report_status = get_obj.report_status
                    powerHistoryObj.compliance_date = get_obj.compliance_date
                    powerHistoryObj.remarks = get_obj.remarks
                    powerHistoryObj.save()
                is_power.org_type = request["organization"]
                is_power.system_type = request["sys_type"]
                is_power.sys_name = request["system_name"]
                is_power.module_name = request['module_name']
                is_power.pc_module_id = request['pc_module_id']
                is_power.lot_no = request['lot_no']
                is_power.design_version = request['design_version']
                is_power.pcs_ldc101 = request["pcs_ldc101"]
                is_power.pcs_ldc102 = request["pcs_ldc102"]
                is_power.pcs_ldc103 = request["pcs_ldc103"]
                is_power.pcs_ldc104 = request["pcs_ldc104"]
                is_power.pcs_ldc105 = request["pcs_ldc105"]
                is_power.pcs_ldc201 = request["pcs_ldc201"]
                is_power.pcs_ldc301 = request["pcs_ldc301"]
                is_power.pcs_ldc401 = request["pcs_ldc401"]
                is_power.pcs_ldc501 = request["pcs_ldc501"]
                is_power.pcs_ldc601 = request["pcs_ldc601"]
                is_power.pcs_ldc602 = request["pcs_ldc602"]
                is_power.compliance_status = request["compliance_status"]
                is_power.report_status = request['report_status']
                is_power.compliance_date = request["compliance_date"]
                is_power.remarks = request["remarks"]
                is_power.save()
                return JsonResponse({'Success': 'Power Record Updated Successfully!'})
        except Exception as e:
            print(e)
            return JsonResponse({'message':'Record could not add.','data':[],'success':False,'staus':'500'},status=500)

    @staticmethod
    def getPowerDashboardCount(request):
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
                    return Q()  # Return an empty Q object instead of None
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

            current_year_modules = 0
            current_year_compliant_modules = 0
            current_year_non_compliant_modules = 0

            total_modules_2 = 0
            total_compliant_modules_2 = 0
            total_non_compliant_modules_2 = 0
            dataList = Power.objects.filter(filter_objects)

            if selected_org != "":
                filter_objects &= get_filter('org_type', 'equal', selected_org)

            total_modules_2 = Power.objects.filter(filter_objects).count()
            total_compliant_modules_2 = dataList.filter(filter_objects, compliance_status = 'Compliant').count()
            total_non_compliant_modules_2 = dataList.filter(filter_objects, compliance_status = 'Non Compliant').count()
            total_partial_compliant_modules_2 = dataList.filter(filter_objects, compliance_status = 'Partial Compliant').count()

            if selected_type != '':
                filter_objects &= get_filter('system_type', 'equal',selected_type)

            current_year_modules = 0
            current_year_compliant_modules = 0
            current_year_non_compliant_modules = 0

            total_modules = 0
            total_compliant_modules = 0
            total_non_compliant_modules = 0
            dataList = Power.objects.filter(filter_objects)

            if selected_type != "":
                filter_objects &= get_filter('system_type', 'equal', selected_type)

            total_modules = Power.objects.filter(filter_objects).count()
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
            dataList = Power.objects.filter(filter_objects)

            # if systemItems != "":
            #     for system in systemItems:
            #      filter_objects &= get_filter('sys_name', 'contains',system)

            total_modules = Power.objects.filter(filter_objects).count()
            total_compliant_modules = dataList.filter(filter_objects, compliance_status = 'Compliant').count()
            total_non_compliant_modules = dataList.filter(filter_objects, compliance_status = 'Non Compliant').count()
            total_partial_compliant_modules = dataList.filter(filter_objects, compliance_status = 'Partial Compliant').count()

            if selected_year != '':
                filter_objects &= get_filter('compliance_date__year', 'equal',selected_year)

            current_year_modules = Power.objects.filter(filter_objects).count()
            current_year_compliant_modules = Power.objects.filter(filter_objects,compliance_status = 'Compliant').count()
            current_year_non_compliant_modules = Power.objects.filter(filter_objects,compliance_status = 'Non Compliant').count()
            current_year_partial_compliant_modules = Power.objects.filter(filter_objects,compliance_status = 'Partial Compliant').count()



            dict = {
                'current_year_modules': current_year_modules,
                'current_year_compliant_modules': current_year_compliant_modules,
                'current_year_non_compliant_modules': current_year_non_compliant_modules,
                'current_year_partial_compliant_modules': current_year_partial_compliant_modules,
                'total_modules': total_modules,
                'total_compliant_modules': total_compliant_modules,
                'total_non_compliant_modules': total_non_compliant_modules,
                'total_partial_compliant_modules': total_partial_compliant_modules,
                'total_modules_2': total_modules_2,
                'total_compliant_modules_2': total_compliant_modules_2,
                'total_non_compliant_modules_2': total_non_compliant_modules_2,
                'total_partial_compliant_modules_2': total_partial_compliant_modules_2
            }
            return JsonResponse({'Success': 'EMS&ES List Fetched Successfully!','data':dict,'success':True,'status':'200'},status=200)
        except Exception as e:
            print(e)
            return JsonResponse({'message':'Record could not add.','data':[],'success':False,'staus':'500'},status=500)

    @staticmethod
    def getPowerList(request):
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

            if selected_type != '':
                filter_objects &= get_filter('system_type', 'equal',selected_type)

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


            dataList = Power.objects.all().order_by('-id')
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
                dataList = dataList.filter(filter_objects, compliance_date__year = selected_year)
            if current_status == 'current_year_compliant_modules':
                dataList = dataList.filter(filter_objects, compliance_date__year = selected_year,compliance_status = 'Compliant')
            if current_status == 'current_year_non_compliant_modules':
                dataList = dataList.filter(filter_objects, compliance_date__year = selected_year,compliance_status = 'Non Compliant')
            if current_status == 'current_year_partial_compliant_modules':
                dataList = dataList.filter(filter_objects, compliance_date__year = selected_year,compliance_status = 'Partial Compliant')
            serializer = PowerSerializer(dataList, many=True)
            return JsonResponse({'Success': 'EMI/Power List Fetched Successfully!','data':serializer.data,'success':True,'status':'200'},status=200)
        except Exception as e:
            print(e)
            return JsonResponse({'message':'Record could not add.','data':[],'success':False,'staus':'500'},status=500)

    @staticmethod
    def getPowerHistory(request):
        try:
            power_id = request.query_params.get('id')
            data = PowerHistory.objects.filter(power_id = power_id)
            serializer = PowerSerializer(data, many=True)
            return JsonResponse({'Success': 'Power List Fetched Successfully!','data':serializer.data,'success':True,'status':'200'},status=200)
        except Exception as e:
            print(e)
            return JsonResponse({'message':'Record could not add.','data':[],'success':False,'staus':'500'},status=500)
    @staticmethod
    def deletePower(request):
        try:
            power_id = request.query_params['power_id']
            power = Power.objects.filter(id=power_id).first()
            power.delete()
            return JsonResponse({'Success': 'Power deleted successfully', 'data':[],'success': True,'status':'201'},status=201)
        except Exception as e:
            print(e)
            return JsonResponse({'message': 'Power could not delete.', 'data': [], 'success': False, 'staus': '500'},
                                status=500)