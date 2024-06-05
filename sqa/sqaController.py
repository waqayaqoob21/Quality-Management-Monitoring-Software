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
    def getModuleNameIDList(request):
       try:
            data = Sqa.objects.all()
            temp_module_name = ""
            module_name_list = []
            temp_software_versions = ""
            software_version_list = []
            optional_system_types_list = []
            temp_optional_system_types = ""
            for item in data:
                if item.module_name != '' and item.module_name != 'other':
                    if item.module_name != temp_module_name and item.module_name not in module_name_list:
                        module_name_list.append(item.module_name)
                    temp_module_name = item.module_name
            for item in data:
                if item.software_version != '' and item.software_version != 'other':
                    if item.software_version != temp_software_versions and item.software_version not in software_version_list:
                        software_version_list.append(item.software_version)
                    temp_software_versions = item.software_version

            for item in data:
                if item.optional_system_types != '' and item.optional_system_types is not None and item.optional_system_types != 'Add More':
                    if item.optional_system_types != temp_optional_system_types and item.optional_system_types not in optional_system_types_list:
                        optional_system_types_list.append(item.optional_system_types)
                    temp_optional_system_types = item.optional_system_types
            dict = {
                'module_name' : module_name_list,
                'software_version' : software_version_list,
                'optional_system_types': optional_system_types_list
            }
            return JsonResponse({'message':'record fetched successfully!','success': True, 'data': dict, 'status': '200'}, status=200)
       except Exception as e:
           print(e)
           return JsonResponse({'message':'record could not fetch','data':[],'success':False, 'status': '500'},status=500)


    @staticmethod
    def addSqa(request):
        try:
            sqa_obj = Sqa()
            id_from_frontend = request['id']
            previous_record = '0'
            record_id = '0'
            endSign = '&'
            if endSign in id_from_frontend:
                previous_record = id_from_frontend.split("&")[1]
            elif id_from_frontend != '0':
                previous_record = id_from_frontend
            else:
                record_id = id_from_frontend
            if record_id == '0':
                sqa_obj.sys_type = request['sys_type']
                sqa_obj.system_name = request['system_name']
                sqa_obj.optional_system_types = request['optional_system_types']
                sqa_obj.module_name = request['module_name']
                sqa_obj.module_id = request['module_id']
                sqa_obj.organization = request['organization']
                sqa_obj.software_type = request['software_type']
                sqa_obj.svc_no = request['svc_no']
                sqa_obj.svc_date = request['svc_date']
                sqa_obj.received_date = request['received_date']
                sqa_obj.revision_no = request['revision_no']
                sqa_obj.revision_date = request['revision_date']
                sqa_obj.purpose = request['purpose']
                sqa_obj.set_no = request['set_no']
                sqa_obj.software_size = request['software_size']
                sqa_obj.installation_date = request['installation_date']
                sqa_obj.installation_date_obs = request['installation_date_obs']
                sqa_obj.attachment = request['attachment']
                sqa_obj.software_version = request['software_version']
                sqa_obj.sqa_certificate_no = request['sqa_certificate_no']
                sqa_obj.request_date = request['request_date']
                sqa_obj.urd = request['urd']
                sqa_obj.srs = request['srs']
                sqa_obj.sdd = request['sdd']
                sqa_obj.rtm = request['rtm']
                sqa_obj.stp = request['stp']
                sqa_obj.unit_test = request['unit_test']
                sqa_obj.static_analysis_report = request['static_analysis_report']
                sqa_obj.assertion_density = request['assertion']
                sqa_obj.eng_change_proposal = request['eng_change_proposal']
                sqa_obj.bugs_observation = request['bugs_observation']
                sqa_obj.cyclomatic_complexity = request['cyclomatic_complexity']
                sqa_obj.code_coverage = request['code_coverage']
                sqa_obj.functional_testing = request['functional_testing']
                sqa_obj.formal_testing = request['formal_testing']
                sqa_obj.design_coverage = request['design_coverage']
                sqa_obj.status = request['status']
                sqa_obj.due_date = request['due_date']
                sqa_obj.audit_completion_date = request['audit_completion_date']
                sqa_obj.remarks = request['remarks']
                sqa_obj.save()
            if previous_record != '0':
                get_obj = Sqa.objects.filter(id=previous_record).first()
                if get_obj.sys_type != request['sys_type'] or get_obj.system_name != request['system_name'] or \
                    get_obj.optional_system_types != request['optional_system_types'] or \
                    get_obj.organization != request['organization'] or get_obj.module_name != request['module_name'] or \
                    get_obj.module_id != request['module_id'] or get_obj.software_type != request['software_type'] or \
                    get_obj.svc_no != request['svc_no'] or get_obj.svc_date != request['svc_date'] or \
                    get_obj.received_date != request['received_date'] or get_obj.revision_no != request['revision_no'] or \
                    get_obj.revision_date != request['revision_date'] or get_obj.purpose != request['purpose'] or \
                    get_obj.set_no != request['set_no'] or get_obj.set_no != request['set_no'] or \
                    get_obj.software_size != request['software_size'] or get_obj.attachment != request['attachment'] or \
                    get_obj.software_version != request['software_version'] or get_obj.sqa_certificate_no != request['sqa_certificate_no'] or \
                    get_obj.request_date != request['request_date'] or get_obj.urd != request['urd'] or \
                    get_obj.srs != request['srs'] or get_obj.sdd != request['sdd'] or get_obj.installation_date != request['installation_date'] or \
                    get_obj.installation_date_obs != request['installation_date_obs'] or get_obj.rtm != request['rtm'] or get_obj.stp != request['stp'] or \
                    get_obj.unit_test != request['unit_test'] or get_obj.static_analysis_report != request['static_analysis_report'] or \
                    get_obj.assertion_density != request['assertion_density'] or get_obj.eng_change_proposal != request['eng_change_proposal'] or \
                    get_obj.bugs_observation != request['bugs_observation'] or get_obj.cyclomatic_complexity != request['cyclomatic_complexity'] or \
                    get_obj.code_coverage != request['code_coverage'] or get_obj.functional_testing != request['functional_testing'] or \
                    get_obj.formal_testing != request['formal_testing'] or get_obj.status != request['status'] or \
                    get_obj.due_date != request['due_date'] or get_obj.audit_completion_date != request['audit_completion_date'] or \
                    get_obj.remarks != request['remarks'] or get_obj.design_coverage != request['design_coverage']:


                    sqaHistoryObj = SqaHistory()
                    sqaHistoryObj.sys_type = get_obj.sys_type
                    sqaHistoryObj.system_name = get_obj.system_name
                    sqaHistoryObj.optional_system_types = get_obj.optional_system_types
                    sqaHistoryObj.module_name = get_obj.module_name
                    sqaHistoryObj.module_id = get_obj.module_id
                    sqaHistoryObj.organization = get_obj.organization
                    sqaHistoryObj.software_type = get_obj.software_type
                    sqaHistoryObj.svc_no = get_obj.svc_no
                    sqaHistoryObj.svc_date = get_obj.svc_date
                    sqaHistoryObj.received_date = get_obj.received_date
                    sqaHistoryObj.revision_no = get_obj.revision_no
                    sqaHistoryObj.revision_date = get_obj.revision_date
                    sqaHistoryObj.purpose = get_obj.purpose
                    sqaHistoryObj.set_no = get_obj.set_no
                    sqaHistoryObj.software_size = get_obj.software_size
                    sqaHistoryObj.installation_date = get_obj.installation_date
                    sqaHistoryObj.installation_date_obs = get_obj.installation_date_obs
                    sqaHistoryObj.attachment = get_obj.attachment
                    sqaHistoryObj.software_version = get_obj.software_version
                    sqaHistoryObj.sqa_certificate_no = get_obj.sqa_certificate_no
                    sqaHistoryObj.request_date = get_obj.request_date
                    sqaHistoryObj.urd = get_obj.urd
                    sqaHistoryObj.srs = get_obj.srs
                    sqaHistoryObj.sdd = get_obj.sdd
                    sqaHistoryObj.rtm = get_obj.rtm
                    sqaHistoryObj.stp = get_obj.stp
                    sqaHistoryObj.unit_test = get_obj.unit_test
                    sqaHistoryObj.static_analysis_report = get_obj.static_analysis_report
                    sqaHistoryObj.assertion_density = get_obj.assertion_density
                    sqaHistoryObj.eng_change_proposal = get_obj.eng_change_proposal
                    sqaHistoryObj.bugs_observation = get_obj.bugs_observation
                    sqaHistoryObj.cyclomatic_complexity = get_obj.cyclomatic_complexity
                    sqaHistoryObj.code_coverage = get_obj.code_coverage
                    sqaHistoryObj.design_coverage = get_obj.design_coverage
                    sqaHistoryObj.functional_testing = get_obj.functional_testing
                    sqaHistoryObj.formal_testing = get_obj.formal_testing
                    sqaHistoryObj.status = get_obj.status
                    sqaHistoryObj.due_date = get_obj.due_date
                    sqaHistoryObj.audit_completion_date = get_obj.audit_completion_date
                    sqaHistoryObj.remarks = get_obj.remarks
                    sqaHistoryObj.sqa = previous_record
                    sqaHistoryObj.save()
                    print("SQA History has been saved!")

                get_obj.sys_type = request['sys_type']
                get_obj.system_name = request['system_name']
                get_obj.optional_system_types = request['optional_system_types']
                get_obj.module_name = request['module_name']
                get_obj.module_id = request['module_id']
                get_obj.organization = request['organization']
                get_obj.software_type = request['software_type']
                get_obj.svc_no = request['svc_no']
                get_obj.svc_date = request['svc_date']
                get_obj.received_date = request['received_date']
                get_obj.revision_no = request['revision_no']
                get_obj.revision_date = request['revision_date']
                get_obj.purpose = request['purpose']
                get_obj.set_no = request['set_no']
                get_obj.software_size = request['software_size']
                get_obj.installation_date = request['installation_date']
                get_obj.installation_date_obs = request['installation_date_obs']
                get_obj.attachment = request['attachment']
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
                get_obj.design_coverage = request['design_coverage']
                get_obj.functional_testing = request['functional_testing']
                get_obj.formal_testing = request['formal_testing']
                get_obj.status = request['status']
                get_obj.due_date = request['due_date']
                get_obj.audit_completion_date = request['audit_completion_date']
                get_obj.remarks = request['remarks']
                get_obj.save()
            msg = ""
            if id_from_frontend != '0':
                msg  = "New record added and previous record updated"
            else:
                msg = "Record updated successfully!"
            return JsonResponse({'message': msg, 'data':{}},status=201)
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
            selected_type = request.query_params.get('selected_type')
            selected_sys = request.query_params.get('selected_system')

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
            if selected_type != "":
                filter_objects &= get_filter('sys_type', 'equal', selected_type)
                total_filter_objects &= get_filter('sys_type', 'equal', selected_type)
                total_filter_objects &= get_filter(
                    'status', 'not_equal',
                    'Certified')
            if selected_sys != "":
                filter_objects &= get_filter('system_name', 'equal', selected_sys)
                total_filter_objects &= get_filter('system_name', 'equal', selected_sys)
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
            selected_type = request.query_params.get('selected_type')
            selected_sys = request.query_params.get('selected_system')
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
            if selected_type != "":
                filter_objects &= get_filter('sys_type', 'equal', selected_type)
                total_filter_objects &= get_filter('sys_type', 'equal', selected_type)
            if selected_sys != "":
                filter_objects &= get_filter('system_name', 'equal', selected_sys)
                total_filter_objects &= get_filter('system_name', 'equal', selected_sys)


            sqaList = ""
            if selected_status == 'total_sbhds':
                sqaList = Sqa.objects.all().order_by('-id')
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

    @staticmethod
    def getSqaHistory(request):
        try:
            sqa_id = request.query_params.get('id')
            sqaObj = SqaHistory.objects.filter(sqa = sqa_id)
            serializer = SqaSerializer(sqaObj, many=True)
            return JsonResponse({'Success': 'SQA history has successfully fetched', 'data': serializer.data, 'success': True, 'status': '200'},
                            status=200)
        except Exception as e:
            print(e)
            return JsonResponse({'message': 'SQA history could not find.', 'data': [], 'success': False, 'staus': '500'},
                                status=500)

    @staticmethod
    def getSqa(request):
        try:
            software_version = request.query_params.get('software_version')
            module_name = request.query_params.get('module_name')
            sqaObj = Sqa.objects.filter(software_version = software_version, module_name = module_name).first()
            if sqaObj is not None:
                serializer = SqaSerializer(sqaObj)
                return JsonResponse({'Success': 'SQA has successfully fetched', 'data': serializer.data, 'success': True, 'status': '200'},
                                status=200)
            else:
                return JsonResponse({'Success': 'SQA could not fetch', 'data': '', 'success': True, 'status': '200'},
                                status=200)
        except Exception as e:
            print(e)
            return JsonResponse({'message': 'SQA could not fetch.', 'data': [], 'success': False, 'staus': '500'},
                                status=500)

    @staticmethod
    def deleteSqaHistory(request):
        try:
            id = request.query_params['id']
            sqa = SqaHistory.objects.filter(id=id).first()
            if sqa is None:
                return JsonResponse(
                    {'Success': 'No SQA history found.', 'data': [], 'success': True, 'status': '401'},
                    status=401)
            else:
                sqa.delete()
                return JsonResponse({'Success': 'SQA history deleted successfully', 'data':[],'success': True,'status':'201'},status=201)
        except Exception as e:
            print(e)
            return JsonResponse({'message': 'SQA history could not delete.', 'data': [], 'success': False, 'staus': '500'},
                                status=500)

    @staticmethod
    def getSqaSummary(request):
        try:
            system_type = request.query_params.get('system_type')
            system_name = request.query_params.get('system_name')
            software_version = request.query_params.get('software_version')
            module_name = request.query_params.get('module_name')
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
            sys_name_list = []
            sys_type_list = []
            module_name_list = []
            software_version_list = []

            typeQuery = Q()
            filter_objects = Q()
            dataList = []
            if system_type != '' and system_type.__contains__(","):
                system_type_list = system_type.split(",")
                for type in system_type_list:
                    filter_objects |= get_filter('sys_type', 'equal', system_type)
            elif system_type != '' and not system_type.__contains__(","):
                filter_objects &= get_filter('sys_type', 'equal', system_type)

            if system_name != '' and system_name.__contains__(","):
                system_name_list = system_name.split(",")
                for name in system_name_list:
                    filter_objects |= get_filter('system_name', 'equal', name)
            elif system_name != '' and not system_name.__contains__(","):
                filter_objects &= get_filter('system_name', 'equal', system_name)

            if module_name != '' and module_name.__contains__(","):
                module_name_list = module_name.split(",")
                for module in module_name_list:
                    filter_objects |= get_filter('module_name', 'equal',module)

            elif module_name != '' and not module_name.__contains__(","):
                filter_objects &= get_filter('module_name', 'equal', module_name)

            if software_version != '' and software_version.__contains__(","):
                software_version_list = software_version.split(",")
                for version in software_version_list:
                    filter_objects |= get_filter('software_version', 'equal', version)
            elif software_version != '' and not software_version.__contains__(","):
                filter_objects &= get_filter('software_version', 'equal', software_version)

            dataList = Sqa.objects.filter(filter_objects)
            serializer = SqaSerializer(dataList, many=True)
            return JsonResponse({'Success': 'SQA Summary could not fetch', 'data': serializer.data, 'success': True, 'status': '200'},
                                    status=200)
        except Exception as e:
            print(e)
            return JsonResponse({'message': 'SQA Summary could not fetch.', 'data': [], 'success': False, 'staus': '500'},
                                status=500)