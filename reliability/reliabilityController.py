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
    def getReliabilityDashboardCount(request):
        try:
            selected_year = request.query_params.get('selected_year')
            selected_type = request.query_params.get('selected_type')
            estimation_type = request.query_params.get('estimation_type')
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

            if selected_type != '':
                filter_objects &= get_filter('sys_type', 'equal',selected_type)

            if estimation_type != '':
                filter_objects &= get_filter('estimation_type', 'equal',estimation_type)
            current_year_modules = 0
            current_year_under_process = 0
            current_year_completed = 0

            total_modules = 0
            total_under_process = 0
            total_completed = 0
            dataList = Reliability.objects.filter(filter_objects)

            if selected_type != "":
                filter_objects &= get_filter('sys_type', 'equal', selected_type)

            total_modules = Reliability.objects.filter(filter_objects).count()
            total_under_process = dataList.filter(filter_objects, current_status = 'Under Process').count()
            total_completed = dataList.filter(filter_objects, current_status = 'Completed').count()

            if selected_year != '':
                filter_objects &= get_filter('created_at__year', 'equal',selected_year)
            current_year_modules = Reliability.objects.filter(filter_objects).count()
            current_year_under_process = Reliability.objects.filter(filter_objects,current_status = 'Under Process').count()


            dict = {
                'current_year_modules': current_year_modules,
                'current_year_under_process': current_year_under_process,
                'current_year_completed': current_year_completed,

                'total_modules': total_modules,
                'total_under_process': total_under_process,
                'total_completed': total_completed,
            }
            return JsonResponse({'Success': 'EMS&ES List Fetched Successfully!','data':dict,'success':True,'status':'200'},status=200)
        except Exception as e:
            print(e)
            return JsonResponse({'message':'Record could not add.','data':[],'success':False,'staus':'500'},status=500)

    @staticmethod
    def getReliabilityList(request):
        try:
            selected_year = request.query_params.get('selected_year')
            selected_type = request.query_params.get('selected_type')
            estimation_type = request.query_params.get('estimation_type')
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
            # if selected_org != '':
            #     filter_objects &= get_filter('organization', 'equal',selected_org)
            if selected_type != '':
                filter_objects &= get_filter('sys_type', 'equal',selected_type)

            if selected_type != "":
                filter_objects &= get_filter('sys_type', 'equal', selected_type)

            dataList = Reliability.objects.all().order_by('-id')
            if current_status == 'total_modules':
                dataList = dataList.filter(filter_objects)
            if current_status == 'total_compliant_modules':
                dataList = dataList.filter(filter_objects, compliance_status = 'Compliant')
            if current_status == 'total_non_compliant_modules':
                dataList = dataList.filter(filter_objects, compliance_status = 'Non Compliant')
            if current_status == 'total_partial_compliant_modules':
                dataList = dataList.filter(filter_objects, compliance_status = 'Partial Compliant')


            if current_status == 'current_year_modules':
                dataList = dataList.filter(filter_objects, created_at__year = selected_year)
            if current_status == 'current_year_compliant_modules':
                dataList = dataList.filter(filter_objects, created_at__year = selected_year,compliance_status = 'Compliant')
            if current_status == 'current_year_non_compliant_modules':
                dataList = dataList.filter(filter_objects, created_at__year = selected_year,compliance_status = 'Non Compliant')
            if current_status == 'current_year_partial_compliant_modules':
                dataList = dataList.filter(filter_objects, created_at__year = selected_year,compliance_status = 'Partial Compliant')
            serializer = ReliabilitySerializer(dataList, many=True)
            return JsonResponse({'Success': 'EMI/EMC List Fetched Successfully!','data':serializer.data,'success':True,'status':'200'},status=200)
        except Exception as e:
            print(e)
            return JsonResponse({'message':'Record could not add.','data':[],'success':False,'staus':'500'},status=500)

    @staticmethod
    def getReliabilityHistory(request):
        try:
            emc_id = request.query_params.get('id')
            data = Reliability.objects.filter(emc_id = emc_id)
            serializer = ReliabilitySerializer(data, many=True)
            return JsonResponse({'Success': 'EMS&ES List Fetched Successfully!','data':serializer.data,'success':True,'status':'200'},status=200)
        except Exception as e:
            print(e)
            return JsonResponse({'message':'Record could not add.','data':[],'success':False,'staus':'500'},status=500)
    @staticmethod
    def deleteReliability(request):
        try:
            emc_id = request.query_params['emc_id']
            emc = Reliability.objects.filter(id=emc_id).first()
            emc.delete()
            return JsonResponse({'Success': 'EMC deleted successfully', 'data':[],'success': True,'status':'201'},status=201)
        except Exception as e:
            print(e)
            return JsonResponse({'message': 'EMC could not delete.', 'data': [], 'success': False, 'staus': '500'},
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
