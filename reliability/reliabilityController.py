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
    def getReliabilityToBeUpdated(request):
        try:
            reliability_id = request.query_params.get('id')
            data = Reliability.objects.filter(id = reliability_id).first()
            serializer = ReliabilitySerializer(data)
            return JsonResponse({'Success': 'Age & Reliability List Fetched Successfully!','data':serializer.data,'success':True,'status':'200'},status=200)
        except Exception as e:
            print(e)
            return JsonResponse({'message':'Record could not add.','data':[],'success':False,'staus':'500'},status=500)
    @staticmethod
    def addReliability(request):
        try:
            id = request['id']
            get_obj = Reliability.objects.filter(id=id).first()
            if get_obj is None:
                rel_obj = Reliability()
                rel_obj.org_type = request["organization"]
                rel_obj.system_type = request['system_type']
                rel_obj.sys_name = request["sys_name"]
                rel_obj.module_name = request['module_name']
                rel_obj.estimation_type = request['estimation_type']
                rel_obj.estimation_method = request['estimation_method']
                rel_obj.request_date = request['request_date']
                rel_obj.due_date = request['due_date']
                rel_obj.completion_date = request['completion_date']
                rel_obj.current_status = request['current_status']
                rel_obj.remarks = request['remarks']
                rel_obj.save()
                return JsonResponse({'Success': 'Age & Reliability record inserted Successfully!'})
            else:
                if get_obj.org_type != request['organization'] or get_obj.system_type != request['system_type'] or get_obj.sys_name != request['sys_name'] or get_obj.module_name != request['module_name'] or \
                get_obj.estimation_method != request['estimation_method'] or str(get_obj.request_date) != request['request_date'] or \
                str(get_obj.due_date) != request['due_date'] or str(get_obj.completion) != request['completion'] or \
                get_obj.current_status != request['current_status'] or get_obj.remarks != request['remarks']:
                    reliabilityHistoryModal = ReliabilityHistory()
                    reliabilityHistoryModal.reliability_id = get_obj.id
                    reliabilityHistoryModal.org_type = get_obj.org_type
                    reliabilityHistoryModal.system_type = get_obj.system_type
                    reliabilityHistoryModal.sys_name = get_obj.sys_name
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

                get_obj.org_type = request["organization"]
                get_obj.system_type = request['system_type']
                get_obj.sys_name = request["sys_name"]
                get_obj.module_name = request['module_name']
                get_obj.estimation_type = request['estimation_type']
                get_obj.estimation_method = request['estimation_method']
                get_obj.request_date = request['request_date']
                get_obj.due_date = request['due_date']
                get_obj.completion_date = request['completion_date']
                get_obj.current_status = request['current_status']
                get_obj.remarks = request['remarks']
                get_obj.save()
                return JsonResponse({'Success': 'Age & Reliability record updated Successfully!'})
        except Exception as e:
            print(e)
            return JsonResponse({'message': 'Record could not add.', 'data': [], 'success': False, 'staus': '500'},
                                status=500)

    @staticmethod
    def getReliabilityDashboardCount(request):
        try:
            selected_year = request.query_params.get('selected_year')
            selected_org = request.query_params.get('selected_org')
            selected_type = request.query_params.get('selected_type')
            selected_system = request.query_params.get('selected_system')
            estimation_type = request.query_params.get('estimation_type')
            systemItems = ""
            noVal = ['']
            if selected_system != noVal:
                systemItems = selected_system.split(',')
                # for item in systemItems:
                #     systemItems = item.split(',')
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


            current_year_modules = 0
            current_year_under_process = 0
            current_year_completed = 0
            current_year_not_started_yet = 0

            total_modules = 0
            total_under_process = 0
            total_completed = 0
            total_not_started_yet = 0

            # total_modules_2 = 0
            # total_under_process_2 = 0
            # total_completed_2 = 0
            # total_not_started_yet_2 = 0

            if selected_org != '':
                filter_objects &= get_filter('org_type', 'equal',selected_org)
            if estimation_type != '':
                filter_objects &= get_filter('estimation_type', 'equal',estimation_type)
            dataList = Reliability.objects.filter(filter_objects)


            total_modules = dataList.count()
            total_under_process = dataList.filter(current_status = 'Under Process').count()
            total_completed = dataList.filter(current_status = 'Completed').count()
            total_not_started_yet = dataList.filter(current_status = 'Not Started Yet').count()

            if selected_type != '':
                filter_objects &= get_filter('system_type', 'equal',selected_type)

            if systemItems != '':
                for system in systemItems:
                 filter_objects &= get_filter('sys_name', 'contains',system)

            if selected_year != '':
                filter_objects &= get_filter('completion_date__year', 'equal',selected_year)
            current_year_modules = Reliability.objects.filter(filter_objects).count()
            current_year_under_process = Reliability.objects.filter(filter_objects,current_status = 'Under Process').count()
            current_year_completed = Reliability.objects.filter(filter_objects,current_status = 'Completed').count()
            current_year_not_started_yet = Reliability.objects.filter(filter_objects,current_status = 'Not Started Yet').count()


            dict = {
                'current_year_modules': current_year_modules,
                'current_year_under_process': current_year_under_process,
                'current_year_completed': current_year_completed,
                'current_year_not_started_yet': current_year_not_started_yet,
                
                'total_modules': total_modules,
                'total_under_process': total_under_process,
                'total_completed': total_completed,
                'total_not_started_yet': total_not_started_yet,

                # 'total_modules_2': total_modules_2,
                # 'total_under_process_2': total_under_process_2,
                # 'total_completed_2': total_completed_2,
                # 'total_not_started_yet_2': total_not_started_yet_2,

            }
            return JsonResponse({'Success': 'EMS&ES List Fetched Successfully!','data':dict,'success':True,'status':'200'},status=200)
        except Exception as e:
            print(e)
            return JsonResponse({'message':'Record could not add.','data':[],'success':False,'staus':'500'},status=500)

    @staticmethod
    def getReliabilityList(request):
        print(request)
        try:
            selected_year = request.query_params.get('selected_year')
            selected_org = request.query_params.get('selected_org')
            selected_type = request.query_params.get('selected_type')
            selected_system = request.query_params.get('selected_system')
            estimation_type = request.query_params.get('estimation_type')
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

            if selected_org != '':
                filter_objects &= get_filter('org_type', 'equal',selected_org)
            
            if selected_type != '':
                filter_objects &= get_filter('system_type', 'equal',selected_type)

            if selected_system != '':
                filter_objects &= get_filter('sys_name', 'contains',selected_system)

            if estimation_type != '':
                filter_objects &= get_filter('estimation_type', 'equal', estimation_type)

            # dataList = Reliability.objects.filter(filter_objects)
            dataList = Reliability.objects.all().order_by('-id')
            if current_status == 'total_modules':
                dataList = dataList.filter(filter_objects)
            if current_status == 'total_under_process':
                dataList = dataList.filter(filter_objects, current_status = 'Under Process')
            if current_status == 'total_completed':
                dataList = dataList.filter(filter_objects, current_status = 'Completed')
            if current_status == 'total_not_started_yet':
                dataList = dataList.filter(filter_objects, current_status = 'Not Started Yet')
            if current_status == 'total_systems':
                dataList = dataList.filter(filter_objects)

            if selected_year != '':
                filter_objects &= get_filter('completion_date__year', 'equal',selected_year)

            if current_status == 'current_year_modules':
                dataList = dataList.filter(filter_objects, completion_date__year = selected_year)
            if current_status == 'current_year_under_process':
                dataList = dataList.filter(filter_objects, completion_date__year = selected_year, current_status = 'Under Process')
            if current_status == 'current_year_completed':
                dataList = dataList.filter(filter_objects, completion_date__year = selected_year, current_status = 'Completed')
            if current_status == 'current_year_not_started_yet':
                dataList = dataList.filter(filter_objects, completion_date__year = selected_year, current_status = 'Not Started Yet')

            serializer = ReliabilitySerializer(dataList, many=True)
            # print("reliabilit list", serializer.data)
            return JsonResponse({'Success': 'Age & Reliability List Fetched Successfully!','data':serializer.data,'success':True,'status':'200'},status=200)
        except Exception as e:
            print(e)
            return JsonResponse({'message':'Record could not add.','data':[],'success':False,'staus':'500'},status=500)

    @staticmethod
    def getReliabilityHistory(request):
        try:
            reliability_id = request.query_params.get('id')
            data = ReliabilityHistory.objects.filter(reliability_id = reliability_id)
            serializer = ReliabilitySerializer(data, many=True)
            return JsonResponse({'Success': 'EMS&ES List Fetched Successfully!','data':serializer.data,'success':True,'status':'200'},status=200)
        except Exception as e:
            print(e)
            return JsonResponse({'message':'Record could not add.','data':[],'success':False,'staus':'500'},status=500)
    @staticmethod
    def deleteReliability(request):
        try:
            reliability_id = request.query_params['id']
            reliability = Reliability.objects.filter(id=reliability_id).first()
            reliability.delete()
            return JsonResponse({'Success': 'Reliability deleted successfully', 'data':[],'success': True,'status':'201'},status=201)
        except Exception as e:
            print(e)
            return JsonResponse({'message': 'Reliability could not delete.', 'data': [], 'success': False, 'staus': '500'},
                                status=500)