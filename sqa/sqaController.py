from django.http import JsonResponse
from fpdf import FPDF
import xlwt
from datetime import datetime, timedelta
from django.http import FileResponse
from django.http import HttpResponse
from django.core.files.storage import FileSystemStorage
from django.db.models import F, Q
import itertools
from sqa.models import *
from sqa.serializer import *
from datetime import date
from django.db.models.functions import Extract
from datetime import datetime
from datetime import timedelta
from dateutil.relativedelta import relativedelta
class sqaController:
    @staticmethod
    def addSqa(request):
        try:
            ems_obj = Sqa()
            id = request['id']
            if id == '0':
                ems_obj.sys_type = request['sys_type']
                ems_obj.system_name = request['system_name']
                ems_obj.module_name = request['module_name']
                ems_obj.module_id = request['module_id']
                ems_obj.organization = request['organization']
                ems_obj.software_type = request['software_type']
                ems_obj.software_version = request['software_version']
                ems_obj.sqa_certificate_no = request['sqa_certificate_no']
                ems_obj.request_date = request['request_date']
                ems_obj.urd = request['urd']
                ems_obj.srs = request['srs']
                ems_obj.sdd = request['sdd']
                ems_obj.rtm = request['rtm']
                ems_obj.stp = request['stp']
                ems_obj.unit_test = request['unit_test']
                ems_obj.static_analysis_report = request['static_analysis_report']
                ems_obj.assertion_density = request['assertion']
                ems_obj.eng_change_proposal = request['eng_change_proposal']
                ems_obj.bugs_observation = request['bugs_observation']
                ems_obj.cyclomatic_complexity = request['cyclomatic_complexity']
                ems_obj.code_coverage = request['code_coverage']
                ems_obj.functional_testing = request['functional_testing']
                ems_obj.formal_testing = request['formal_testing']
                ems_obj.status = request['status']
                ems_obj.due_date = request['due_date']
                ems_obj.audit_completion_date = request['audit_completion_date']
                ems_obj.remarks = request['remarks']
                ems_obj.save()
                return JsonResponse({'Success': 'SQA record inserted Successfully!'})
            else:
                get_obj = Sqa.objects.filter(id=id).first()
                get_obj.sys_type = request['sys_type']
                get_obj.system_name = request['system_name']
                get_obj.module_name = request['module_name']
                get_obj.module_id = request['module_id']
                get_obj.organization = request['organization']
                get_obj.software_type = request['software_type']
                get_obj.software_version = request['software_version']
                get_obj.sqa_certificate_no = request['sqa_certificate_no']
                get_obj.request_date = request['request_date']
                get_obj.urd = request['urd']
                get_obj.srs = request['srs']
                get_obj.sdd = request['sdd']
                get_obj.rtm = request['rtm']
                get_obj.stp = request['stp']
                get_obj.unit_test = request['unit_test']
                get_obj.static_analysis_report = request['static_analysis_report']
                get_obj.assertion_density = request['assertion']
                get_obj.eng_change_proposal = request['eng_change_proposal']
                get_obj.bugs_observation = request['bugs_observation']
                get_obj.cyclomatic_complexity = request['cyclomatic_complexity']
                get_obj.code_coverage = request['code_coverage']
                get_obj.functional_testing = request['functional_testing']
                get_obj.formal_testing = request['formal_testing']
                get_obj.status = request['status']
                get_obj.due_date = request['due_date']
                get_obj.audit_completion_date = request['audit_completion_date']
                get_obj.remarks = request['remarks']
                get_obj.save()
                return JsonResponse({'Success': 'SQA Record Updated Successfully!'})
        except Exception as e:
            print(e)
            return JsonResponse({'message': 'Record could not add.', 'data': [], 'success': False, 'staus': '500'},
                                status=500)

    @staticmethod
    def getSqaDashboardCount(request, self=None):
        try:

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

            selected_year = request.query_params.get('selected_year')
            selected_org = request.query_params.get('selected_org')

            typeQuery = Q()
            filter_objects = Q()
            total_filter_objects = Q()
            if selected_year != '':
                filter_objects &= get_filter('request_date__year', 'equal',selected_year)
            if selected_org != '':
                filter_objects &= get_filter('organization', 'equal',selected_org)
                total_filter_objects &= get_filter('organization', 'equal',selected_org)
                total_filter_objects &= get_filter(
                    'status', 'not_equal',
                    'Certified')

            total_audit_inprocess_overdue = 0
            total_sbhd = Sqa.objects.filter(total_filter_objects).count()
            total_sbhd_qm_qualified = Sqa.objects.filter(total_filter_objects,status = 'QM Qualified').count()
            total_sbhd_inprocess_qm = Sqa.objects.filter(total_filter_objects,status = 'Audit in-process').count()
            total_sbhd_inprocess_nco = Sqa.objects.filter(status = 'Audit in-process').count()
            # total_audit_inprocess_overdue = Sqa.objects.filter(total_filter_objects,status = 'total_audit_inprocess_overdue').count()
            docList = Sqa.objects.filter(total_filter_objects, status = 'Audit in-process')
            list = []
            if docList is not None:
                for item in docList:
                    if item.audit_completion_date is None:
                        item.audit_completion_date = datetime.today() + timedelta(hours=5)
                    if item.due_date.date() < item.audit_completion_date.date():
                        total_audit_inprocess_overdue += 1

            docList = Sqa.objects.filter(total_filter_objects)
            total_overdue_qual_obs = 0
            if docList is not None:
                for item in docList:
                    if item.audit_completion_date is None:
                        item.audit_completion_date = datetime.today() + timedelta(hours=5)
                    if item.due_date.date() < item.audit_completion_date.date():
                        total_overdue_qual_obs += 1

            current_year_count = Sqa.objects.filter(filter_objects).count()
            qm_qualified = Sqa.objects.filter(filter_objects, status = 'QM Qualified').count()
            qm_observation_forwarded = Sqa.objects.filter(filter_objects, status = 'QM Observations Forwarded').count()
            qm_observation_repeated = Sqa.objects.filter(filter_objects, status = 'QM Observations Repeated').count()
            audit_inProcess = Sqa.objects.filter(filter_objects, status = 'Audit in-process').count()


            docList = Sqa.objects.filter(filter_objects)
            over_due_sbhd = 0
            if docList is not None:
                for item in docList:
                    if item.audit_completion_date is None:
                        item.audit_completion_date = datetime.today() + timedelta(hours=5)
                    if item.due_date.date() < item.audit_completion_date.date():
                        over_due_sbhd += 1
            docList = Sqa.objects.filter(filter_objects)
            sbhd_qual_obs_overdue = 0
            if docList is not None:
                for item in docList:
                    if item.audit_completion_date is None:
                        item.audit_completion_date = datetime.today() + timedelta(hours=5)
                    if item.due_date.date() < item.audit_completion_date.date():
                        sbhd_qual_obs_overdue += 1
            dist = {
                'total_sbhd': total_sbhd,
                'total_sbhd_qm_qualified': total_sbhd_qm_qualified,
                'total_sbhd_inprocess_qm': total_sbhd_inprocess_qm,
                'total_sbhd_inprocess_nco': total_sbhd_inprocess_nco,
                'total_audit_inprocess_overdue': total_audit_inprocess_overdue,
                'total_sbhd_qual_obs': total_overdue_qual_obs,

                'current_year_count': current_year_count,
                'qm_qualified': qm_qualified,
                'qm_observation_forwarded': qm_observation_forwarded,
                'qm_observation_repeated': qm_observation_repeated,
                'audit_inProcess': audit_inProcess,
                'over_due_sbhd': over_due_sbhd,
                'sbhd_qual_obs_overdue': sbhd_qual_obs_overdue,

            }

            # DataCount.append(dist)
            return JsonResponse({'message': 'Data fetched successfully!', 'data': dist,'success':'True'},
                                status=200)

        except Exception as e:
            print(e)
            return JsonResponse({'status': 'False', "message": "Internal Server Error"}, status=500)
            pass

    @staticmethod
    def getSqaList(request):
        try:
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

            selected_year = request.query_params.get('selected_year')
            selected_org = request.query_params.get('selected_org')
            selected_status = request.query_params.get('selected_status')

            typeQuery = Q()
            filter_objects = Q()
            total_filter_objects = Q()
            if selected_year != '':
                filter_objects &= get_filter('request_date__year', 'equal',selected_year)
            if selected_org != '':
                filter_objects &= get_filter('organization', 'equal',selected_org)
                total_filter_objects &= get_filter('organization', 'equal',selected_org)
            if selected_status != '':
                total_filter_objects &= get_filter(
                    'status', 'not_equal',
                    'Audit in-process')


            sqaList = ""
            if selected_status == 'total_sbhds':
                sqaList = Sqa.objects.all()
            if selected_status == 'totalApproved':
                sqaList = Sqa.objects.filter(status = 'QM Qualified')
            if selected_status == 'totaln_sbhd_inprocess_qm':
                sqaList = Sqa.objects.filter(total_filter_objects, status='Audit in-process')
            if selected_status == 'total_sbhd_inprocess_nco':
                sqaList = Sqa.objects.filter(status='Audit in-process')
            if selected_status == 'total_audit_inprocess_overdue':
                docList = Sqa.objects.filter(organization = selected_org, status='Audit In-process')
                list = []
                if docList is not None:
                    for item in docList:
                        if item.audit_completion_date is None:
                            item.audit_completion_date = datetime.today() + timedelta(hours=5)
                        if item.due_date.date() < item.audit_completion_date.date():
                            list.append(item)
                    sqaList = list
            if selected_status == 'total_overdue_qual_obs':
                docList = Sqa.objects.filter(total_filter_objects)
                list = []
                if docList is not None:
                    for item in docList:
                        if item.audit_completion_date is None:
                            item.audit_completion_date = datetime.today() + timedelta(hours=5)
                        if item.due_date.date() < item.audit_completion_date.date():
                            list.append(item)
                    sqaList = list

            if selected_status == 'current_yar_count':
                sqaList = Sqa.objects.filter(filter_objects)
            if selected_status == 'current_qm_qualified':
                sqaList = Sqa.objects.filter(filter_objects,status='QM Qualified')
            if selected_status == 'current_QM_Observations_Forwarded':
                sqaList = Sqa.objects.filter(filter_objects,
                                                              status='QM Observations Forwarded')

            if selected_status == 'current_QM_Observations_Repeated':
                sqaList = Sqa.objects.filter(filter_objects, status='QM Observations Repeated')

            if selected_status == 'current_Audit_in_process':
                sqaList = Sqa.objects.filter(filter_objects, status='Audit in-process')

            if selected_status == 'inprocess_overdue':
                docList = Sqa.objects.filter(filter_objects)
                list = []
                if docList is not None:
                    for item in docList:
                        if item.audit_completion_date is None:
                            item.audit_completion_date = datetime.today() + timedelta(hours=5)
                        if item.due_date.date() < item.audit_completion_date.date():
                            list.append(item)
                    sqaList = list
            if selected_status == 'current_overdue_qual_obs':
                docList = Sqa.objects.filter(filter_objects)
                list = []
                if docList is not None:
                    for item in docList:
                        if item.audit_completion_date is None:
                            item.audit_completion_date = datetime.today() + timedelta(hours=5)
                        if item.due_date.date() < item.audit_completion_date.date():
                            list.append(item)

                    sqaList = list


            serializer = SqaSerializer(sqaList, many=True)

            return JsonResponse({'message': 'Data fetched successfully!', 'data': serializer.data,'success':'True'},
                                status=200)
        except Exception as e:
            print(e)
            return JsonResponse({'message': 'Record could not add.', 'data': [], 'success': False, 'staus': '500'},
                                status=500)

    @staticmethod
    def deleteSqa(request):
        try:
            id = request.query_params['id']
            sqa = Sqa.objects.filter(id=id).first()
            if sqa is None:
                return JsonResponse(
                    {'Success': 'No SQA found.', 'data': [], 'success': True, 'status': '401'},
                    status=401)
            else:
                sqa.delete()
                return JsonResponse({'Success': 'SQA deleted successfully', 'data':[],'success': True,'status':'201'},status=201)
        except Exception as e:
            print(e)
            return JsonResponse({'message': 'SQA could not delete.', 'data': [], 'success': False, 'staus': '500'},
                                status=500)
