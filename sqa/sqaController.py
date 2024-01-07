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
    def getSqaList(request):
        try:
            data = Sqa.objects.all()
            serializer = SqaSerializer(data, many=True)
            return JsonResponse({'Success': 'SQA list fetched', 'data':serializer.data,'success': True,'status':'200'},status=200)
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
