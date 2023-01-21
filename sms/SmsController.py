from importlib import import_module

from django.db.models import Q
from django.http import JsonResponse
from ams.serializer import *
from ams.models import *
from fpdf import FPDF
import xlwt
from datetime import datetime, timedelta
from django.http import FileResponse
from django.http import HttpResponse
from django.core.files.storage import FileSystemStorage
import re
import fitz  # pip install PyMuPDF Pillow
import io
import os
from PIL import Image
from pytesseract import pytesseract

from sms.models import FlightSystemStatus, ProductionSystemStatus, RelifingSystemStatus, ProductionSystemStatusHistory, \
    FlightSystemStatusHistory, RelifingSystemStatusHistory
from sms.serializers import FlightSystemSerialzer, ProductionSystemSerialzer, \
    RelifingSystemSerialzer  # install tesseract-ocr-w64-setup-v5.2.0.20220712.exe (64 bit) resp.


# from https://github.com/UB-Mannheim/tesseract/wiki
# pip install pytesseract

class SmsController:

    @staticmethod
    def AddProductionStatus(request):
        is_active = 0
        prodModel = ProductionSystemStatus()
        if request['is_active'] == 'true':
            is_active = 1
        try:
            id = request['id']
            if id == '0':
                prodModel.system = request['system']
                prodModel.organization = request['organization']
                prodModel.set_id = request['set_id']
                prodModel.blt_date = request['blt_date']
                prodModel.testing_date = request['testing_date']
                prodModel.sys_type = request['sys_type']
                prodModel.blt_status = request['blt_status']
                prodModel.blt_remarks = request['blt_remarks']
                prodModel.pre_hil_date = request['pre_hil_date']
                prodModel.pre_hil_status = request['pre_hil_status']
                prodModel.pre_hil_remarks = request['pre_hil_remarks']
                prodModel.vibaration_date = request['vibaration_date']
                prodModel.vibaration_status = request['vibaration_status']
                prodModel.vibaration_remarks = request['vibaration_remarks']
                prodModel.post_hil_date = request['post_hil_date']
                prodModel.post_hil_status = request['post_hil_status']
                prodModel.post_hil_remarks = request['post_hil_remarks']
                prodModel.fgt_date = request['fgt_date']
                prodModel.fgt_status = request['fgt_status']
                prodModel.fgt_remarks = request['fgt_remarks']
                prodModel.final_integration_date = request['final_integration_date']
                prodModel.final_integration_status = request['final_integration_status']
                prodModel.final_integration_remarks = request['final_integration_remarks']
                prodModel.bhd_date = request['bhd_date']
                prodModel.bhd_status = request['bhd_status']
                prodModel.bhd_remarks = request['bhd_remarks']
                prodModel.fqm_date = request['fqm_date']
                prodModel.fqm_status = request['fqm_status']
                prodModel.fqm_remarks = request['fqm_remarks']
                prodModel.qm_certification_date = request['qm_certification_date']
                prodModel.qm_certification_status = request['qm_certification_status']
                prodModel.qm_certification_remarks = request['qm_certification_remarks']
                prodModel.attachment = request['attachment']
                prodModel.remarks = request['remarks']
                prodModel.cgbalancing_date = request['cgbalancing_date']
                prodModel.cgbalancing_date_status = request['cgbalancing_date_status']
                prodModel.cgbalancing_date_remarks = request['cgbalancing_date_remarks']
                prodModel.enduser_date = request['enduser_date']
                prodModel.enduser_status = request['enduser_status']
                prodModel.enduser_remarks = request['enduser_remarks']
                prodModel.incapsulation_date = request['incapsulation_date']
                prodModel.incapsulation_status = request['incapsulation_status']
                prodModel.incapsulation_remarks = request['incapsulation_remarks']
                prodModel.sys_align_Date = request['sys_align_Date']
                prodModel.sys_align_status = request['sys_align_status']
                prodModel.sys_align_remarks = request['sys_align_remarks']
                prodModel.testing_type = request['testing_type']
                prodModel.emp_proofing = request['emp_proofing']
                prodModel.func_tst = request['func_tst']
                prodModel.func_tst_dummy_bird = request['func_tst_dummy_bird']
                prodModel.road_test = request['road_test']
                prodModel.post_road_test = request['post_road_test']
                prodModel.integrated_operation = request['integrated_operation']
                prodModel.rain_test = request['rain_test']
                prodModel.pre_user_inspection = request['pre_user_inspection']
                prodModel.final_integrated_testing = request['final_integrated_testing']
                prodModel.load_unload_on_mlv_hlf = request['load_unload_on_mlv_hlf']

                prodModel.isActive = is_active
                prodModel.save()
                return JsonResponse({'status': 'True', 'message': "Production Status Added Successfully!"},
                                    status=200)
            else:
                get_prod = ProductionSystemStatus.objects.filter(id=id).first()
                if get_prod is not None:
                    if request['blt_status'] != get_prod.blt_status or request[
                        'pre_hil_status'] != get_prod.pre_hil_status or request[
                        'vibaration_status'] != get_prod.vibaration_status or request[
                        'post_hil_status'] != get_prod.post_hil_status or request[
                        'fgt_status'] != get_prod.fgt_status or request[
                        'final_integration_status'] != get_prod.final_integration_status or request[
                        'bhd_status'] != get_prod.bhd_status or request['fqm_status'] != get_prod.fqm_status or request[
                        'qm_certification_status'] != get_prod.qm_certification_status or \
                            request['incapsulation_status'] != get_prod.incapsulation_status or \
                            request['sys_align_status'] != get_prod.sys_align_status:
                        print("add data in production history")
                        ProdHistoryModal = ProductionSystemStatusHistory()
                        ProdHistoryModal.prod_id = get_prod.id
                        ProdHistoryModal.system = get_prod.system
                        ProdHistoryModal.organization = get_prod.organization
                        ProdHistoryModal.set_id = get_prod.set_id
                        ProdHistoryModal.blt_date = get_prod.blt_date
                        ProdHistoryModal.blt_status = get_prod.blt_status
                        ProdHistoryModal.blt_remarks = get_prod.blt_remarks
                        ProdHistoryModal.testing_date = get_prod.testing_date
                        ProdHistoryModal.sys_type = get_prod.sys_type
                        ProdHistoryModal.pre_hil_date = get_prod.pre_hil_date
                        ProdHistoryModal.pre_hil_status = get_prod.pre_hil_status
                        ProdHistoryModal.pre_hil_remarks = get_prod.pre_hil_remarks
                        ProdHistoryModal.vibaration_date = get_prod.vibaration_date
                        ProdHistoryModal.vibaration_status = get_prod.vibaration_status
                        ProdHistoryModal.vibaration_remarks = get_prod.vibaration_remarks
                        ProdHistoryModal.post_hil_date = get_prod.post_hil_date
                        ProdHistoryModal.post_hil_status = get_prod.post_hil_status
                        ProdHistoryModal.post_hil_remarks = get_prod.post_hil_remarks
                        ProdHistoryModal.fgt_date = get_prod.fgt_date
                        ProdHistoryModal.fgt_status = get_prod.fgt_status
                        ProdHistoryModal.fgt_remarks = get_prod.fgt_remarks
                        ProdHistoryModal.final_integration_date = get_prod.final_integration_date
                        ProdHistoryModal.final_integration_status = get_prod.final_integration_status
                        ProdHistoryModal.final_integration_remarks = get_prod.final_integration_remarks
                        ProdHistoryModal.bhd_date = get_prod.bhd_date
                        ProdHistoryModal.bhd_status = get_prod.bhd_status
                        ProdHistoryModal.bhd_remarks = get_prod.bhd_remarks
                        ProdHistoryModal.fqm_date = get_prod.fqm_date
                        ProdHistoryModal.fqm_status = get_prod.fqm_status
                        ProdHistoryModal.fqm_remarks = get_prod.fqm_remarks
                        ProdHistoryModal.qm_certification_date = get_prod.qm_certification_date
                        ProdHistoryModal.qm_certification_status = get_prod.qm_certification_status
                        ProdHistoryModal.qm_certification_remarks = get_prod.qm_certification_remarks
                        ProdHistoryModal.attachment = get_prod.attachment
                        ProdHistoryModal.remarks = get_prod.remarks
                        ProdHistoryModal.isActive = get_prod.isActive
                        ProdHistoryModal.cgbalancing_date = get_prod.cgbalancing_date
                        ProdHistoryModal.cgbalancing_date_status = get_prod.cgbalancing_date_status
                        ProdHistoryModal.cgbalancing_date_remarks = get_prod.cgbalancing_date_remarks
                        ProdHistoryModal.enduser_date = get_prod.enduser_date
                        ProdHistoryModal.enduser_status = get_prod.enduser_status
                        ProdHistoryModal.enduser_remarks = get_prod.enduser_remarks
                        ProdHistoryModal.incapsulation_date = get_prod.incapsulation_date
                        ProdHistoryModal.incapsulation_status = get_prod.incapsulation_status
                        ProdHistoryModal.incapsulation_remarks = get_prod.incapsulation_remarks
                        ProdHistoryModal.sys_align_Date = get_prod.sys_align_Date
                        ProdHistoryModal.sys_align_status = get_prod.sys_align_status
                        ProdHistoryModal.sys_align_remarks = get_prod.sys_align_remarks
                        ProdHistoryModal.testing_type = get_prod.testing_type
                        ProdHistoryModal.emp_proofing = get_prod.emp_proofing
                        ProdHistoryModal.func_tst = get_prod.func_tst
                        ProdHistoryModal.func_tst_dummy_bird = get_prod.func_tst_dummy_bird
                        ProdHistoryModal.road_test = get_prod.road_test
                        ProdHistoryModal.post_road_test = get_prod.post_road_test
                        ProdHistoryModal.integrated_operation = get_prod.integrated_operation
                        ProdHistoryModal.rain_test = get_prod.rain_test
                        ProdHistoryModal.pre_user_inspection = get_prod.pre_user_inspection
                        ProdHistoryModal.final_integrated_testing = get_prod.final_integrated_testing
                        ProdHistoryModal.load_unload_on_mlv_hlf = get_prod.load_unload_on_mlv_hlf

                        ProdHistoryModal.save()

                        # history saved
                    get_prod.system = request['system']
                    get_prod.organization = request['organization']
                    get_prod.set_id = request['set_id']
                    get_prod.blt_date = request['blt_date']
                    get_prod.blt_status = request['blt_status']
                    get_prod.blt_remarks = request['blt_remarks']
                    get_prod.testing_date = request['testing_date']
                    get_prod.sys_type = request['sys_type']
                    get_prod.pre_hil_date = request['pre_hil_date']
                    get_prod.pre_hil_status = request['pre_hil_status']
                    get_prod.pre_hil_remarks = request['pre_hil_remarks']
                    get_prod.vibaration_date = request['vibaration_date']
                    get_prod.vibaration_status = request['vibaration_status']
                    get_prod.vibaration_remarks = request['vibaration_remarks']
                    get_prod.post_hil_date = request['post_hil_date']
                    get_prod.post_hil_status = request['post_hil_status']
                    get_prod.post_hil_remarks = request['post_hil_remarks']
                    get_prod.fgt_date = request['fgt_date']
                    get_prod.fgt_status = request['fgt_status']
                    get_prod.fgt_remarks = request['fgt_remarks']
                    get_prod.final_integration_date = request['final_integration_date']
                    get_prod.final_integration_status = request['final_integration_status']
                    get_prod.final_integration_remarks = request['final_integration_remarks']
                    get_prod.bhd_date = request['bhd_date']
                    get_prod.bhd_status = request['bhd_status']
                    get_prod.bhd_remarks = request['bhd_remarks']
                    get_prod.fqm_date = request['fqm_date']
                    get_prod.fqm_status = request['fqm_status']
                    get_prod.fqm_remarks = request['fqm_remarks']
                    get_prod.qm_certification_date = request['qm_certification_date']
                    get_prod.qm_certification_status = request['qm_certification_status']
                    get_prod.qm_certification_remarks = request['qm_certification_remarks']
                    if request['attachment'] != '':
                        get_prod.attachment = request['attachment']
                    get_prod.remarks = request['remarks']
                    get_prod.isActive = is_active
                    get_prod.cgbalancing_date = request['cgbalancing_date']
                    get_prod.cgbalancing_date_status = request['cgbalancing_date_status']
                    get_prod.cgbalancing_date_remarks = request['cgbalancing_date_remarks']
                    get_prod.enduser_date = request['enduser_date']
                    get_prod.enduser_status = request['enduser_status']
                    get_prod.enduser_remarks = request['enduser_remarks']
                    get_prod.incapsulation_date = request['incapsulation_date']
                    get_prod.incapsulation_status = request['incapsulation_status']
                    get_prod.incapsulation_remarks = request['incapsulation_remarks']
                    get_prod.sys_align_Date = request['sys_align_Date']
                    get_prod.sys_align_status = request['sys_align_status']
                    get_prod.sys_align_remarks = request['sys_align_remarks']
                    get_prod.testing_type = request['testing_type']
                    get_prod.emp_proofing = request['emp_proofing']
                    get_prod.func_tst = request['func_tst']
                    get_prod.func_tst_dummy_bird = request['func_tst_dummy_bird']
                    get_prod.road_test = request['road_test']
                    get_prod.post_road_test = request['post_road_test']
                    get_prod.integrated_operation = request['integrated_operation']
                    get_prod.rain_test = request['rain_test']
                    get_prod.pre_user_inspection = request['pre_user_inspection']
                    get_prod.final_integrated_testing = request['final_integrated_testing']
                    get_prod.load_unload_on_mlv_hlf = request['load_unload_on_mlv_hlf']

                    get_prod.save()
            return JsonResponse({'status': 'True', 'message': "Production Status Updated Successfully!"},
                                status=200)
        except Exception as e:
            return JsonResponse({'status': 'False', "message": "Status Not Saved"}, status=500)

    @staticmethod
    def GetProductionList(request):
        try:
            org = request.query_params['selected_org']
            year = request.query_params['selected_year']
            type = request.query_params['selected_type']
            system = request.query_params['selected_system']
            status = request.query_params['selected_system']
            filter_objects = Q()

            def get_filter(field_name, filter_condition, filter_value):
                # thanks to the below post
                # https://stackoverflow.com/questions/310732/in-django-how-does-one-filter-a-queryset-with-dynamic-field-lookups
                # the idea to this below logic is very similar to that in the above mentioned post
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

            # create dynamic filter
            if year != '':
                filter_objects &= get_filter(
                    'testing_date__year', 'equal',
                    year)
            if org != '':
                filter_objects &= get_filter(
                    'organization', 'equal',
                    org)
            if type != '':
                filter_objects &= get_filter(
                    'sys_type', 'equal',
                    type)
            if system != '':
                filter_objects &= get_filter(
                    'system', 'equal',
                    system)

            # if status == 'Total':
            #     data = ProductionSystemStatus.objects.all()
            #     serializer = ProductionSystemSerialzer(data, many=True)
            #     return JsonResponse({'status': 'true', 'data': serializer.data}, status=200)
            # else:
            data = ProductionSystemStatus.objects.filter(filter_objects)
            serializer = ProductionSystemSerialzer(data, many=True)
            print(serializer.data)
            return JsonResponse({'status': 'true', 'data': serializer.data}, status=200)
        except Exception as e:
            print(e)
            return JsonResponse({'status': 'false'}, status=200)

    @staticmethod
    def GetProductionListHistory(request):
        try:
            id = request.query_params['id']
            data = ProductionSystemStatusHistory.objects.filter(prod_id=id)
            serializer = ProductionSystemSerialzer(data, many=True)
            return JsonResponse({'status': 'true', 'data': serializer.data}, status=200)

        except Exception as e:
            print(e)
            return JsonResponse({'status': 'false'}, status=200)

    # API for PDF generator with table and Text_wraping

    @staticmethod
    def GetProductionPDFList(request):
        data = []
        status = request.query_params['status']
        if status == 'Total':
            data = ProductionSystemStatus.objects.all().order_by('-id').values_list('system', 'organization', 'set_id',
                                                                                    'blt_status', 'pre_hil_status',
                                                                                    'vibaration_status',
                                                                                    'post_hil_status',
                                                                                    'fgt_status',
                                                                                    'final_integration_status',
                                                                                    'bhd_status', 'fqm_status',
                                                                                    'qm_certification_status',
                                                                                    'remarks')
        else:
            data = ProductionSystemStatus.objects.filter(qm_certification_status=status).order_by('-id').values_list(
                'system', 'organization', 'set_id',
                'blt_status', 'pre_hil_status',
                'vibaration_status',
                'post_hil_status',
                'fgt_status',
                'final_integration_status',
                'bhd_status', 'fqm_status',
                'qm_certification_status',
                'remarks')
        TABLE_COL_NAMES = (
            "System", "Organization", "Set ID", "BLT Status", "Pre-Hil Status", "Vibration Status", "Post-Hil Status",
            "FGT Status", "Final Integration Status", "BHD Status", "FQM Status", "QM Certification Status", "Remarks")

        pdf = FPDF('L', 'mm', 'Legal')
        pdf.add_page()
        pdf.set_font('courier', 'B', 26)
        pdf.cell(330, 10, 'Production System Report', border=0, align='C', ln=2)
        pdf.cell(40, 10, '', 0, 1)
        pdf.set_font("Times", size=10)
        line_height = pdf.font_size * 3.2
        col_width = pdf.epw / 14

        def render_table_header():
            pdf.set_font(style="B")
            for col_name in TABLE_COL_NAMES:
                pdf.multi_cell(col_width, line_height, col_name, border=1, align='C', ln=3,
                               max_line_height=pdf.font_size)
            pdf.ln(line_height)
            pdf.set_font(style="")

        render_table_header()

        lh_list = []
        use_default_height = 0
        for row in data:
            for datum in row:
                dd = str(datum)
                word_list = dd.split()
                number_of_words = len(word_list)
                if number_of_words > 2:
                    use_default_height = 1
                    new_line_height = pdf.font_size * (number_of_words / 2)
            if not use_default_height:
                lh_list.append(line_height)
            else:
                lh_list.append(new_line_height)
                use_default_height = 0

        for j, row in enumerate(data):
            line_height = lh_list[j]
            if pdf.will_page_break(line_height):
                render_table_header()
            for col_num in range(len(row)):
                line_height = lh_list[j]
                pdf.multi_cell(col_width, line_height, f"{row[col_num]}", border=1, align='C', ln=3,
                               max_line_height=pdf.font_size)
            pdf.ln(line_height)
        pdf.output('ProductionSystem.pdf')
        return FileResponse(open('ProductionSystem.pdf', 'rb'), as_attachment=True, content_type='application/pdf')

    @staticmethod
    def GetProductionExcelList(request):

        data = []
        status = request.query_params['status']
        if status == 'Total':
            data = ProductionSystemStatus.objects.all().order_by('-id').values_list('system', 'organization', 'set_id',
                                                                                    'blt_status', 'pre_hil_status',
                                                                                    'vibaration_status',
                                                                                    'post_hil_status',
                                                                                    'fgt_status',
                                                                                    'final_integration_status',
                                                                                    'bhd_status', 'fqm_status',
                                                                                    'qm_certification_status',
                                                                                    'remarks')
        else:
            data = ProductionSystemStatus.objects.filter(qm_certification_status=status).order_by('-id').values_list(
                'system', 'organization', 'set_id',
                'blt_status', 'pre_hil_status',
                'vibaration_status',
                'post_hil_status',
                'fgt_status',
                'final_integration_status',
                'bhd_status', 'fqm_status',
                'qm_certification_status',
                'remarks')
        response = HttpResponse(content_type='application/ms-excel')
        response['Content-Disposition'] = 'attachment; filename="DocumentExcelSheet.xls"'
        work_book = xlwt.Workbook(encoding='utf-8')
        work_sheet = work_book.add_sheet('ExcelSheet')
        row_num = 0
        font_style = xlwt.XFStyle()
        font_style.font.bold = True
        TABLE_COL_NAMES = (
            "System", "Organization", "Set ID", "BLT Status", "Pre-Hil Status", "Vibration Status", "Post-Hil Status",
            "FGT Status", "Final Integration Status", "BHD Status", "FQM Status", "QM Certification Status", "Remarks")

        for col_num in range(len(TABLE_COL_NAMES)):
            work_sheet.write(row_num, col_num, TABLE_COL_NAMES[col_num], font_style)
        font_style = xlwt.XFStyle()

        for row in data:
            row_num += 1
            for col_num in range(len(row)):
                work_sheet.write(row_num, col_num, str(row[col_num]), font_style)
        work_book.save(response)
        return response

    @staticmethod
    def DeleteProdSys(request):

        try:
            id = request.query_params['id']
            delete = ProductionSystemStatus.objects.filter(id=id).delete()

            return JsonResponse({'status': 'True', 'message': "Record Deleted"},
                                status=200)

        except Exception as e:
            print(e)
            return JsonResponse({'status': 'False', "message": "Doc Not Deleted"}, status=500)
            pass

    @staticmethod
    def DeleteFlightSys(request):

        try:
            id = request.query_params['id']
            delete = FlightSystemStatus.objects.filter(id=id).delete()
            return JsonResponse({'status': 'True', 'message': "Record Deleted"},
                                status=200)

        except Exception as e:
            print(e)
            return JsonResponse({'status': 'False', "message": "Doc Not Deleted"}, status=500)
            pass

    @staticmethod
    def DeleteRelifingSys(request):

        try:
            id = request.query_params['id']
            delete = RelifingSystemStatus.objects.filter(id=id).delete()
            return JsonResponse({'status': 'True', 'message': "Record Deleted"},
                                status=200)

        except Exception as e:
            print(e)
            return JsonResponse({'status': 'False', "message": "Doc Not Deleted"}, status=500)
            pass

    @staticmethod
    def AddFlightStatus(request):
        is_active = 0
        if request['is_active'] == 'true':
            is_active = 1
        flightModel = FlightSystemStatus()

        # try:
        id = request['id']
        if id == '0':
            flightModel.system = request['system']
            flightModel.organization = request['organization']
            flightModel.set_id = request['set_id']
            flightModel.blt_date = request['blt_date']
            flightModel.testing_date = request['testing_date']
            flightModel.sys_type = request['sys_type']
            flightModel.blt_status = request['blt_status']
            flightModel.blt_remarks = request['blt_remarks']
            flightModel.pre_hil_date = request['pre_hil_date']
            flightModel.pre_hil_status = request['pre_hil_status']
            flightModel.pre_hil_remarks = request['pre_hil_remarks']
            flightModel.vibaration_date = request['vibaration_date']
            flightModel.vibaration_status = request['vibaration_status']
            flightModel.vibaration_remarks = request['vibaration_remarks']
            flightModel.post_hil_date = request['post_hil_date']
            flightModel.post_hil_status = request['post_hil_status']
            flightModel.post_hil_remarks = request['post_hil_remarks']
            flightModel.fgt_date = request['fgt_date']
            flightModel.fgt_status = request['fgt_status']
            flightModel.fgt_remarks = request['fgt_remarks']
            flightModel.final_integration_date = request['final_integration_date']
            flightModel.final_integration_status = request['final_integration_status']
            flightModel.final_integration_remarks = request['final_integration_remarks']
            flightModel.bhd_date = request['bhd_date']
            flightModel.bhd_status = request['bhd_status']
            flightModel.bhd_remarks = request['bhd_remarks']
            flightModel.fqm_date = request['fqm_date']
            flightModel.fqm_status = request['fqm_status']
            flightModel.fqm_remarks = request['fqm_remarks']
            flightModel.qm_certification_date = request['qm_certification_date']
            flightModel.qm_certification_status = request['qm_certification_status']
            flightModel.qm_certification_remarks = request['qm_certification_remarks']
            flightModel.attachment = request['attachment']
            flightModel.remarks = request['remarks']
            flightModel.cgbalancing_date = request['cgbalancing_date']
            flightModel.cgbalancing_date_status = request['cgbalancing_date_status']
            flightModel.cgbalancing_date_remarks = request['cgbalancing_date_remarks']
            flightModel.launchact_date = request['launchact_date']
            flightModel.launchact_status = request['launchact_status']
            flightModel.launchact_remarks = request['launchact_remarks']
            flightModel.incapsulation_date = request['incapsulation_date']
            flightModel.incapsulation_status = request['incapsulation_status']
            flightModel.incapsulation_remarks = request['incapsulation_remarks']
            flightModel.sys_align_Date = request['sys_align_Date']
            flightModel.sys_align_status = request['sys_align_status']
            flightModel.sys_align_remarks = request['sys_align_remarks']
            flightModel.testing_type = request['testing_type']
            flightModel.emp_proofing = request['emp_proofing']
            flightModel.func_tst = request['func_tst']
            flightModel.func_tst_dummy_bird = request['func_tst_dummy_bird']
            flightModel.road_test = request['road_test']
            flightModel.post_road_test = request['post_road_test']
            flightModel.integrated_operation = request['integrated_operation']
            flightModel.rain_test = request['rain_test']
            flightModel.pre_user_inspection = request['pre_user_inspection']
            flightModel.final_integrated_testing = request['final_integrated_testing']
            flightModel.load_unload_on_mlv_hlf = request['load_unload_on_mlv_hlf']

            flightModel.isActive = is_active
            flightModel.save()
            return JsonResponse({'status': 'True', 'message': "Production Status Added Successfully!"},
                                status=200)
        else:
            get_prod = FlightSystemStatus.objects.filter(id=id).first()
            if get_prod is not None:

                if request['blt_status'] != get_prod.blt_status or request[
                    'pre_hil_status'] != get_prod.pre_hil_status or request[
                    'vibaration_status'] != get_prod.vibaration_status or request[
                    'post_hil_status'] != get_prod.post_hil_status or request[
                    'fgt_status'] != get_prod.fgt_status or request[
                    'final_integration_status'] != get_prod.final_integration_status or request[
                    'bhd_status'] != get_prod.bhd_status or request['fqm_status'] != get_prod.fqm_status or request[
                    'qm_certification_status'] != get_prod.qm_certification_status or \
                        request['incapsulation_status'] != get_prod.incapsulation_status or \
                        request['sys_align_status'] != get_prod.sys_align_status:
                    print("add data in flight history")
                    FlightHistoryModal = FlightSystemStatusHistory()
                    FlightHistoryModal.f_id = get_prod.id
                    FlightHistoryModal.system = get_prod.system
                    FlightHistoryModal.organization = get_prod.organization
                    FlightHistoryModal.set_id = get_prod.set_id
                    FlightHistoryModal.blt_date = get_prod.blt_date
                    FlightHistoryModal.blt_status = get_prod.blt_status
                    FlightHistoryModal.blt_remarks = get_prod.blt_remarks
                    FlightHistoryModal.testing_date = get_prod.testing_date
                    FlightHistoryModal.sys_type = get_prod.sys_type
                    FlightHistoryModal.pre_hil_date = get_prod.pre_hil_date
                    FlightHistoryModal.pre_hil_status = get_prod.pre_hil_status
                    FlightHistoryModal.pre_hil_remarks = get_prod.pre_hil_remarks
                    FlightHistoryModal.vibaration_date = get_prod.vibaration_date
                    FlightHistoryModal.vibaration_status = get_prod.vibaration_status
                    FlightHistoryModal.vibaration_remarks = get_prod.vibaration_remarks
                    FlightHistoryModal.post_hil_date = get_prod.post_hil_date
                    FlightHistoryModal.post_hil_status = get_prod.post_hil_status
                    FlightHistoryModal.post_hil_remarks = get_prod.post_hil_remarks
                    FlightHistoryModal.fgt_date = get_prod.fgt_date
                    FlightHistoryModal.fgt_status = get_prod.fgt_status
                    FlightHistoryModal.fgt_remarks = get_prod.fgt_remarks
                    FlightHistoryModal.final_integration_date = get_prod.final_integration_date
                    FlightHistoryModal.final_integration_status = get_prod.final_integration_status
                    FlightHistoryModal.final_integration_remarks = get_prod.final_integration_remarks
                    FlightHistoryModal.bhd_date = get_prod.bhd_date
                    FlightHistoryModal.bhd_status = get_prod.bhd_status
                    FlightHistoryModal.bhd_remarks = get_prod.bhd_remarks
                    FlightHistoryModal.fqm_date = get_prod.fqm_date
                    FlightHistoryModal.fqm_status = get_prod.fqm_status
                    FlightHistoryModal.fqm_remarks = get_prod.fqm_remarks
                    FlightHistoryModal.qm_certification_date = get_prod.qm_certification_date
                    FlightHistoryModal.qm_certification_status = get_prod.qm_certification_status
                    FlightHistoryModal.qm_certification_remarks = get_prod.qm_certification_remarks
                    FlightHistoryModal.attachment = get_prod.attachment
                    FlightHistoryModal.remarks = get_prod.remarks
                    FlightHistoryModal.isActive = get_prod.isActive
                    FlightHistoryModal.cgbalancing_date = get_prod.cgbalancing_date
                    FlightHistoryModal.cgbalancing_date_status = get_prod.cgbalancing_date_status
                    FlightHistoryModal.cgbalancing_date_remarks = get_prod.cgbalancing_date_remarks
                    FlightHistoryModal.launchact_date = get_prod.launchact_date
                    FlightHistoryModal.launchact_status = get_prod.launchact_status
                    FlightHistoryModal.launchact_remarks = get_prod.launchact_remarks
                    FlightHistoryModal.incapsulation_date = get_prod.incapsulation_date
                    FlightHistoryModal.incapsulation_status = get_prod.incapsulation_status
                    FlightHistoryModal.incapsulation_remarks = get_prod.incapsulation_remarks
                    FlightHistoryModal.sys_align_Date = get_prod.sys_align_Date
                    FlightHistoryModal.sys_align_status = get_prod.sys_align_status
                    FlightHistoryModal.sys_align_remarks = get_prod.sys_align_remarks
                    FlightHistoryModal.testing_type = get_prod.testing_type
                    FlightHistoryModal.emp_proofing = get_prod.emp_proofing
                    FlightHistoryModal.func_tst = get_prod.func_tst
                    FlightHistoryModal.func_tst_dummy_bird = get_prod.func_tst_dummy_bird
                    FlightHistoryModal.road_test = get_prod.road_test
                    FlightHistoryModal.post_road_test = get_prod.post_road_test
                    FlightHistoryModal.integrated_operation = get_prod.integrated_operation
                    FlightHistoryModal.rain_test = get_prod.rain_test
                    FlightHistoryModal.pre_user_inspection = get_prod.pre_user_inspection
                    FlightHistoryModal.final_integrated_testing = get_prod.final_integrated_testing
                    FlightHistoryModal.load_unload_on_mlv_hlf = get_prod.load_unload_on_mlv_hlf

                    FlightHistoryModal.save()

                    # history saved
                get_prod.system = request['system']
                get_prod.organization = request['organization']
                get_prod.set_id = request['set_id']
                get_prod.blt_date = request['blt_date']
                get_prod.blt_status = request['blt_status']
                get_prod.blt_remarks = request['blt_remarks']
                get_prod.testing_date = request['testing_date']
                get_prod.sys_type = request['sys_type']
                get_prod.pre_hil_date = request['pre_hil_date']
                get_prod.pre_hil_status = request['pre_hil_status']
                get_prod.pre_hil_remarks = request['pre_hil_remarks']
                get_prod.vibaration_date = request['vibaration_date']
                get_prod.vibaration_status = request['vibaration_status']
                get_prod.vibaration_remarks = request['vibaration_remarks']
                get_prod.post_hil_date = request['post_hil_date']
                get_prod.post_hil_status = request['post_hil_status']
                get_prod.post_hil_remarks = request['post_hil_remarks']
                get_prod.fgt_date = request['fgt_date']
                get_prod.fgt_status = request['fgt_status']
                get_prod.fgt_remarks = request['fgt_remarks']
                get_prod.final_integration_date = request['final_integration_date']
                get_prod.final_integration_status = request['final_integration_status']
                get_prod.final_integration_remarks = request['final_integration_remarks']
                get_prod.bhd_date = request['bhd_date']
                get_prod.bhd_status = request['bhd_status']
                get_prod.bhd_remarks = request['bhd_remarks']
                get_prod.fqm_date = request['fqm_date']
                get_prod.fqm_status = request['fqm_status']
                get_prod.fqm_remarks = request['fqm_remarks']
                get_prod.qm_certification_date = request['qm_certification_date']
                get_prod.qm_certification_status = request['qm_certification_status']
                get_prod.qm_certification_remarks = request['qm_certification_remarks']
                if request['attachment'] != '':
                    get_prod.attachment = request['attachment']
                get_prod.remarks = request['remarks']
                get_prod.isActive = is_active
                get_prod.cgbalancing_date = request['cgbalancing_date']
                get_prod.cgbalancing_date_status = request['cgbalancing_date_status']
                get_prod.cgbalancing_date_remarks = request['cgbalancing_date_remarks']
                get_prod.launchact_date = request['launchact_date']
                get_prod.launchact_status = request['launchact_status']
                get_prod.launchact_remarks = request['launchact_remarks']
                get_prod.incapsulation_date = request['incapsulation_date']
                get_prod.incapsulation_status = request['incapsulation_status']
                get_prod.incapsulation_remarks = request['incapsulation_remarks']
                get_prod.sys_align_Date = request['sys_align_Date']
                get_prod.sys_align_status = request['sys_align_status']
                get_prod.sys_align_remarks = request['sys_align_remarks']
                get_prod.testing_type = request['testing_type']
                get_prod.emp_proofing = request['emp_proofing']
                get_prod.func_tst = request['func_tst']
                get_prod.func_tst_dummy_bird = request['func_tst_dummy_bird']
                get_prod.road_test = request['road_test']
                get_prod.post_road_test = request['post_road_test']
                get_prod.integrated_operation = request['integrated_operation']
                get_prod.rain_test = request['rain_test']
                get_prod.pre_user_inspection = request['pre_user_inspection']
                get_prod.final_integrated_testing = request['final_integrated_testing']
                get_prod.load_unload_on_mlv_hlf = request['load_unload_on_mlv_hlf']
                get_prod.isActive = is_active
                get_prod.save()
        return JsonResponse({'status': 'True', 'message': "Production Status Updated Successfully!"},
                            status=200)

    # except Exception as e:
    #     return JsonResponse({'status': 'False', "message": "Status Not Saved"}, status=500)

    @staticmethod
    def GetFlightList(request):
        # try:
        org = request.query_params['selected_org']
        year = request.query_params['selected_year']
        type = request.query_params['selected_type']
        system = request.query_params['selected_system']
        filter_objects = Q()

        def get_filter(field_name, filter_condition, filter_value):
            # thanks to the below post
            # https://stackoverflow.com/questions/310732/in-django-how-does-one-filter-a-queryset-with-dynamic-field-lookups
            # the idea to this below logic is very similar to that in the above mentioned post
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

        # create dynamic filter
        if year != '':
            filter_objects &= get_filter(
                'testing_date__year', 'equal',
                year)
        if org != '':
            filter_objects &= get_filter(
                'organization', 'equal',
                org)
        if type != '':
            filter_objects &= get_filter(
                'sys_type', 'equal',
                type)
        if system != '':
            filter_objects &= get_filter(
                'system', 'equal',
                system)
        data = FlightSystemStatus.objects.filter(filter_objects)
        serializer = FlightSystemSerialzer(data, many=True)
        return JsonResponse({'message': 'Welcome to Home Page', 'data': serializer.data}, status=200)

    # except:
    #     return JsonResponse({'message': 'Sorry! No Task found.'}, status=500)

    # API for PDF generator with table and Text_wraping

    @staticmethod
    def GetFlightListHistory(request):
        try:
            id = request.query_params['id']
            data = FlightSystemStatusHistory.objects.filter(f_id=id)
            serializer = FlightSystemSerialzer(data, many=True)
            return JsonResponse({'status': 'true', 'data': serializer.data}, status=200)

        except Exception as e:
            print(e)
            return JsonResponse({'status': 'false'}, status=200)

    @staticmethod
    def GetFlightPDFList(request):
        TABLE_COL_NAMES = (
            "System", "Organization", "Set ID", "BLT Status", "Pre-Hil Status", "Vibration Status", "Post-Hil Status",
            "FGT Status", "Final Integration Status", "BHD Status", "FQM Status", "QM Certification Status",
            "Remarks")
        data = []
        status = request.query_params['status']
        if status == 'Total':
            data = FlightSystemStatus.objects.all().order_by('-id').values_list('system', 'organization', 'set_id',
                                                                                'blt_status', 'pre_hil_status',
                                                                                'vibaration_status', 'post_hil_status',
                                                                                'fgt_status',
                                                                                'final_integration_status',
                                                                                'bhd_status', 'fqm_status',
                                                                                'qm_certification_status',
                                                                                'remarks')
        else:
            data = FlightSystemStatus.objects.filter(qm_certification_status=status).order_by('-id').values_list(
                'system', 'organization', 'set_id',
                'blt_status', 'pre_hil_status',
                'vibaration_status', 'post_hil_status',
                'fgt_status',
                'final_integration_status',
                'bhd_status', 'fqm_status',
                'qm_certification_status',
                'remarks')

        pdf = FPDF('L', 'mm', 'Legal')
        pdf.add_page()
        pdf.set_font('courier', 'B', 26)
        pdf.cell(330, 10, 'Flight System Report', border=0, align='C', ln=2)
        pdf.cell(40, 10, '', 0, 1)
        pdf.set_font("Times", size=10)
        line_height = pdf.font_size * 4.5
        col_width = pdf.epw / 14

        def render_table_header():
            pdf.set_font(style="B")
            for col_name in TABLE_COL_NAMES:
                pdf.multi_cell(col_width, line_height, col_name, border=1, align='C', ln=3,
                               max_line_height=pdf.font_size)
            pdf.ln(line_height)
            pdf.set_font(style="")

        render_table_header()

        lh_list = []
        use_default_height = 0
        for row in data:
            for datum in row:
                dd = str(datum)
                word_list = dd.split()
                number_of_words = len(word_list)
                if number_of_words > 2:
                    use_default_height = 1
                    new_line_height = pdf.font_size * 2 * (number_of_words / 2)
            if not use_default_height:
                lh_list.append(line_height)
            else:
                lh_list.append(new_line_height)
                use_default_height = 0

        for j, row in enumerate(data):
            line_height = lh_list[j]
            if pdf.will_page_break(line_height):
                render_table_header()
            for col_num in range(len(row)):
                line_height = lh_list[j]
                pdf.multi_cell(col_width, line_height, f"{row[col_num]}", border=1, align='C', ln=3,
                               max_line_height=pdf.font_size)
            pdf.ln(line_height)
        pdf.output('FlightSystem.pdf')
        return FileResponse(open('FlightSystem.pdf', 'rb'), as_attachment=True, content_type='application/pdf')

    @staticmethod
    def GetFlightExcelList(request):
        response = HttpResponse(content_type='application/ms-excel')
        response['Content-Disposition'] = 'attachment; filename="DocumentExcelSheet.xls"'
        work_book = xlwt.Workbook(encoding='utf-8')
        work_sheet = work_book.add_sheet('ExcelSheet')
        row_num = 0
        font_style = xlwt.XFStyle()
        font_style.font.bold = True
        TABLE_COL_NAMES = (
            "System", "Organization", "Set ID", "BLT Status", "Pre-Hil Status", "Vibration Status", "Post-Hil Status",
            "FGT Status", "Final Integration Status", "BHD Status", "FQM Status", "QM Certification Status",
            "Remarks")

        for col_num in range(len(TABLE_COL_NAMES)):
            work_sheet.write(row_num, col_num, TABLE_COL_NAMES[col_num], font_style)
        font_style = xlwt.XFStyle()
        data = []
        status = request.query_params['status']
        if status == 'Total':
            data = FlightSystemStatus.objects.all().order_by('-id').values_list('system', 'organization', 'set_id',
                                                                                'blt_status', 'pre_hil_status',
                                                                                'vibaration_status', 'post_hil_status',
                                                                                'fgt_status',
                                                                                'final_integration_status',
                                                                                'bhd_status', 'fqm_status',
                                                                                'qm_certification_status',
                                                                                'remarks')
        else:
            data = FlightSystemStatus.objects.filter(qm_certification_status=status).order_by('-id').values_list(
                'system', 'organization', 'set_id',
                'blt_status', 'pre_hil_status',
                'vibaration_status', 'post_hil_status',
                'fgt_status',
                'final_integration_status',
                'bhd_status', 'fqm_status',
                'qm_certification_status',
                'remarks')
        for row in data:
            row_num += 1
            for col_num in range(len(row)):
                work_sheet.write(row_num, col_num, str(row[col_num]), font_style)
        work_book.save(response)
        return response

    @staticmethod
    def AddRelifingStatus(request):
        refilModel = RelifingSystemStatus()
        is_active = 0
        if request['is_active'] == 'true':
            is_active = 1
        try:
            id = request['id']
            if id == '0':
                refilModel.system = request['system']
                refilModel.organization = request['organization']
                refilModel.set_id = request['set_id']
                refilModel.blt_date = request['blt_date']
                refilModel.testing_date = request['testing_date']
                refilModel.sys_type = request['sys_type']
                refilModel.blt_status = request['blt_status']
                refilModel.blt_remarks = request['blt_remarks']
                refilModel.pre_hil_date = request['pre_hil_date']
                refilModel.pre_hil_status = request['pre_hil_status']
                refilModel.pre_hil_remarks = request['pre_hil_remarks']
                refilModel.vibaration_date = request['vibaration_date']
                refilModel.vibaration_status = request['vibaration_status']
                refilModel.vibaration_remarks = request['vibaration_remarks']
                refilModel.post_hil_date = request['post_hil_date']
                refilModel.post_hil_status = request['post_hil_status']
                refilModel.post_hil_remarks = request['post_hil_remarks']
                refilModel.fgt_date = request['fgt_date']
                refilModel.fgt_status = request['fgt_status']
                refilModel.fgt_remarks = request['fgt_remarks']
                refilModel.final_integration_date = request['final_integration_date']
                refilModel.final_integration_status = request['final_integration_status']
                refilModel.final_integration_remarks = request['final_integration_remarks']
                refilModel.bhd_date = request['bhd_date']
                refilModel.bhd_status = request['bhd_status']
                refilModel.bhd_remarks = request['bhd_remarks']
                refilModel.fqm_date = request['fqm_date']
                refilModel.fqm_status = request['fqm_status']
                refilModel.fqm_remarks = request['fqm_remarks']
                refilModel.qm_certification_date = request['qm_certification_date']
                refilModel.qm_certification_status = request['qm_certification_status']
                refilModel.qm_certification_remarks = request['qm_certification_remarks']
                refilModel.attachment = request['attachment']
                refilModel.remarks = request['remarks']
                refilModel.cgbalancing_date = request['cgbalancing_date']
                refilModel.cgbalancing_date_status = request['cgbalancing_date_status']
                refilModel.cgbalancing_date_remarks = request['cgbalancing_date_remarks']
                refilModel.enduser_date = request['enduser_date']
                refilModel.enduser_status = request['enduser_status']
                refilModel.enduser_remarks = request['enduser_remarks']
                refilModel.incapsulation_date = request['incapsulation_date']
                refilModel.incapsulation_status = request['incapsulation_status']
                refilModel.incapsulation_remarks = request['incapsulation_remarks']
                refilModel.sys_align_Date = request['sys_align_Date']
                refilModel.sys_align_status = request['sys_align_status']
                refilModel.sys_align_remarks = request['sys_align_remarks']
                refilModel.testing_type = request['testing_type']
                refilModel.emp_proofing = request['emp_proofing']
                refilModel.func_tst = request['func_tst']
                refilModel.func_tst_dummy_bird = request['func_tst_dummy_bird']
                refilModel.road_test = request['road_test']
                refilModel.post_road_test = request['post_road_test']
                refilModel.integrated_operation = request['integrated_operation']
                refilModel.rain_test = request['rain_test']
                refilModel.pre_user_inspection = request['pre_user_inspection']
                refilModel.final_integrated_testing = request['final_integrated_testing']
                refilModel.load_unload_on_mlv_hlf = request['load_unload_on_mlv_hlf']

                refilModel.isActive = is_active
                refilModel.save()
                return JsonResponse({'status': 'True', 'message': "Production Status Added Successfully!"},
                                    status=200)
            else:
                get_prod = RelifingSystemStatus.objects.filter(id=id).first()
                if get_prod is not None:
                    if request['blt_status'] != get_prod.blt_status or request[
                        'pre_hil_status'] != get_prod.pre_hil_status or request[
                        'vibaration_status'] != get_prod.vibaration_status or request[
                        'post_hil_status'] != get_prod.post_hil_status or request[
                        'fgt_status'] != get_prod.fgt_status or request[
                        'final_integration_status'] != get_prod.final_integration_status or request[
                        'bhd_status'] != get_prod.bhd_status or request['fqm_status'] != get_prod.fqm_status or request[
                        'qm_certification_status'] != get_prod.qm_certification_status or \
                            request['incapsulation_status'] != get_prod.incapsulation_status or \
                            request['sys_align_status'] != get_prod.sys_align_status:
                        print("add data in flight history")
                        RelifingHistoryModal = RelifingSystemStatusHistory()
                        RelifingHistoryModal.r_id = get_prod.id
                        RelifingHistoryModal.system = get_prod.system
                        RelifingHistoryModal.organization = get_prod.organization
                        RelifingHistoryModal.set_id = get_prod.set_id
                        RelifingHistoryModal.blt_date = get_prod.blt_date
                        RelifingHistoryModal.blt_status = get_prod.blt_status
                        RelifingHistoryModal.blt_remarks = get_prod.blt_remarks
                        RelifingHistoryModal.testing_date = get_prod.testing_date
                        RelifingHistoryModal.sys_type = get_prod.sys_type
                        RelifingHistoryModal.pre_hil_date = get_prod.pre_hil_date
                        RelifingHistoryModal.pre_hil_status = get_prod.pre_hil_status
                        RelifingHistoryModal.pre_hil_remarks = get_prod.pre_hil_remarks
                        RelifingHistoryModal.vibaration_date = get_prod.vibaration_date
                        RelifingHistoryModal.vibaration_status = get_prod.vibaration_status
                        RelifingHistoryModal.vibaration_remarks = get_prod.vibaration_remarks
                        RelifingHistoryModal.post_hil_date = get_prod.post_hil_date
                        RelifingHistoryModal.post_hil_status = get_prod.post_hil_status
                        RelifingHistoryModal.post_hil_remarks = get_prod.post_hil_remarks
                        RelifingHistoryModal.fgt_date = get_prod.fgt_date
                        RelifingHistoryModal.fgt_status = get_prod.fgt_status
                        RelifingHistoryModal.fgt_remarks = get_prod.fgt_remarks
                        RelifingHistoryModal.final_integration_date = get_prod.final_integration_date
                        RelifingHistoryModal.final_integration_status = get_prod.final_integration_status
                        RelifingHistoryModal.final_integration_remarks = get_prod.final_integration_remarks
                        RelifingHistoryModal.bhd_date = get_prod.bhd_date
                        RelifingHistoryModal.bhd_status = get_prod.bhd_status
                        RelifingHistoryModal.bhd_remarks = get_prod.bhd_remarks
                        RelifingHistoryModal.fqm_date = get_prod.fqm_date
                        RelifingHistoryModal.fqm_status = get_prod.fqm_status
                        RelifingHistoryModal.fqm_remarks = get_prod.fqm_remarks
                        RelifingHistoryModal.qm_certification_date = get_prod.qm_certification_date
                        RelifingHistoryModal.qm_certification_status = get_prod.qm_certification_status
                        RelifingHistoryModal.qm_certification_remarks = get_prod.qm_certification_remarks
                        RelifingHistoryModal.attachment = get_prod.attachment
                        RelifingHistoryModal.remarks = get_prod.remarks
                        RelifingHistoryModal.isActive = get_prod.isActive
                        RelifingHistoryModal.cgbalancing_date = get_prod.cgbalancing_date
                        RelifingHistoryModal.cgbalancing_date_status = get_prod.cgbalancing_date_status
                        RelifingHistoryModal.cgbalancing_date_remarks = get_prod.cgbalancing_date_remarks
                        RelifingHistoryModal.enduser_date = get_prod.enduser_date
                        RelifingHistoryModal.enduser_status = get_prod.enduser_status
                        RelifingHistoryModal.enduser_remarks = get_prod.enduser_remarks
                        RelifingHistoryModal.incapsulation_date = get_prod.incapsulation_date
                        RelifingHistoryModal.incapsulation_status = get_prod.incapsulation_status
                        RelifingHistoryModal.incapsulation_remarks = get_prod.incapsulation_remarks
                        RelifingHistoryModal.sys_align_Date = get_prod.sys_align_Date
                        RelifingHistoryModal.sys_align_status = get_prod.sys_align_status
                        RelifingHistoryModal.sys_align_remarks = get_prod.sys_align_remarks
                        RelifingHistoryModal.testing_type = get_prod.testing_type
                        RelifingHistoryModal.emp_proofing = get_prod.emp_proofing
                        RelifingHistoryModal.func_tst = get_prod.func_tst
                        RelifingHistoryModal.func_tst_dummy_bird = get_prod.func_tst_dummy_bird
                        RelifingHistoryModal.road_test = get_prod.road_test
                        RelifingHistoryModal.post_road_test = get_prod.post_road_test
                        RelifingHistoryModal.integrated_operation = get_prod.integrated_operation
                        RelifingHistoryModal.rain_test = get_prod.rain_test
                        RelifingHistoryModal.pre_user_inspection = get_prod.pre_user_inspection
                        RelifingHistoryModal.final_integrated_testing = get_prod.final_integrated_testing
                        RelifingHistoryModal.load_unload_on_mlv_hlf = get_prod.load_unload_on_mlv_hlf

                        RelifingHistoryModal.save()

                        # history saved
                    get_prod.system = request['system']
                    get_prod.organization = request['organization']
                    get_prod.set_id = request['set_id']
                    get_prod.blt_date = request['blt_date']
                    get_prod.blt_status = request['blt_status']
                    get_prod.blt_remarks = request['blt_remarks']
                    get_prod.testing_date = request['testing_date']
                    get_prod.sys_type = request['sys_type']
                    get_prod.pre_hil_date = request['pre_hil_date']
                    get_prod.pre_hil_status = request['pre_hil_status']
                    get_prod.pre_hil_remarks = request['pre_hil_remarks']
                    get_prod.vibaration_date = request['vibaration_date']
                    get_prod.vibaration_status = request['vibaration_status']
                    get_prod.vibaration_remarks = request['vibaration_remarks']
                    get_prod.post_hil_date = request['post_hil_date']
                    get_prod.post_hil_status = request['post_hil_status']
                    get_prod.post_hil_remarks = request['post_hil_remarks']
                    get_prod.fgt_date = request['fgt_date']
                    get_prod.fgt_status = request['fgt_status']
                    get_prod.fgt_remarks = request['fgt_remarks']
                    get_prod.final_integration_date = request['final_integration_date']
                    get_prod.final_integration_status = request['final_integration_status']
                    get_prod.final_integration_remarks = request['final_integration_remarks']
                    get_prod.bhd_date = request['bhd_date']
                    get_prod.bhd_status = request['bhd_status']
                    get_prod.bhd_remarks = request['bhd_remarks']
                    get_prod.fqm_date = request['fqm_date']
                    get_prod.fqm_status = request['fqm_status']
                    get_prod.fqm_remarks = request['fqm_remarks']
                    get_prod.qm_certification_date = request['qm_certification_date']
                    get_prod.qm_certification_status = request['qm_certification_status']
                    get_prod.qm_certification_remarks = request['qm_certification_remarks']
                    if request['attachment'] != '':
                        get_prod.attachment = request['attachment']
                    get_prod.remarks = request['remarks']
                    get_prod.isActive = is_active
                    get_prod.cgbalancing_date = request['cgbalancing_date']
                    get_prod.cgbalancing_date_status = request['cgbalancing_date_status']
                    get_prod.cgbalancing_date_remarks = request['cgbalancing_date_remarks']
                    get_prod.enduser_date = request['enduser_date']
                    get_prod.enduser_status = request['enduser_status']
                    get_prod.enduser_remarks = request['enduser_remarks']
                    get_prod.incapsulation_date = request['incapsulation_date']
                    get_prod.incapsulation_status = request['incapsulation_status']
                    get_prod.incapsulation_remarks = request['incapsulation_remarks']
                    get_prod.sys_align_Date = request['sys_align_Date']
                    get_prod.sys_align_status = request['sys_align_status']
                    get_prod.sys_align_remarks = request['sys_align_remarks']
                    get_prod.testing_type = request['testing_type']
                    get_prod.emp_proofing = request['emp_proofing']
                    get_prod.func_tst = request['func_tst']
                    get_prod.func_tst_dummy_bird = request['func_tst_dummy_bird']
                    get_prod.road_test = request['road_test']
                    get_prod.post_road_test = request['post_road_test']
                    get_prod.integrated_operation = request['integrated_operation']
                    get_prod.rain_test = request['rain_test']
                    get_prod.pre_user_inspection = request['pre_user_inspection']
                    get_prod.final_integrated_testing = request['final_integrated_testing']
                    get_prod.load_unload_on_mlv_hlf = request['load_unload_on_mlv_hlf']

                    get_prod.save()
            return JsonResponse({'status': 'True', 'message': "Production Status Updated Successfully!"},
                                status=200)
        except Exception as e:
            return JsonResponse({'status': 'False', "message": "Status Not Saved"}, status=500)

    @staticmethod
    def GetRelifingList(request):
        try:
            org = request.query_params['selected_org']
            year = request.query_params['selected_year']
            type = request.query_params['selected_type']
            system = request.query_params['selected_system']
            filter_objects = Q()

            def get_filter(field_name, filter_condition, filter_value):
                # thanks to the below post
                # https://stackoverflow.com/questions/310732/in-django-how-does-one-filter-a-queryset-with-dynamic-field-lookups
                # the idea to this below logic is very similar to that in the above mentioned post
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

            # create dynamic filter
            if year != '':
                filter_objects &= get_filter(
                    'testing_date__year', 'equal',
                    year)
            if org != '':
                filter_objects &= get_filter(
                    'organization', 'equal',
                    org)
            if type != '':
                filter_objects &= get_filter(
                    'sys_type', 'equal',
                    type)
            if system != '':
                filter_objects &= get_filter(
                    'system', 'equal',
                    system)
            # if status == 'Total':
            #     data = RelifingSystemStatus.objects.all()
            #     serializer = RelifingSystemSerialzer(data, many=True)
            #     return JsonResponse({'message': 'Welcome to Home Page', 'data': serializer.data}, status=200)
            # else:
            data = RelifingSystemStatus.objects.filter(filter_objects)
            serializer = RelifingSystemSerialzer(data, many=True)
            return JsonResponse({'message': 'Welcome to Home Page', 'data': serializer.data}, status=200)
        except:
            return JsonResponse({'message': 'Sorry! No Task found.'}, status=200)

    @staticmethod
    def GetRelifingHistoryList(request):
        try:
            id = request.query_params['id']
            data = RelifingSystemStatusHistory.objects.filter(r_id=id)
            serializer = RelifingSystemSerialzer(data, many=True)
            return JsonResponse({'message': 'Welcome to Home Page', 'data': serializer.data}, status=200)

        except:
            return JsonResponse({'message': 'Sorry! No Task found.'}, status=200)

    # API for PDF generator with table and Text_wraping

    @staticmethod
    def GetRelifingPDFList(request):

        data = []
        status = request.query_params['status']
        if status == 'Total':
            data = RelifingSystemStatus.objects.all().order_by('-id').values_list('system', 'organization', 'set_id',
                                                                                  'blt_status', 'pre_hil_status',
                                                                                  'vibaration_status',
                                                                                  'post_hil_status',
                                                                                  'fgt_status',
                                                                                  'final_integration_status',
                                                                                  'bhd_status', 'fqm_status',
                                                                                  'qm_certification_status', 'remarks')
        else:
            data = RelifingSystemStatus.objects.filter(qm_certification_status=status).order_by('-id').values_list(
                'system', 'organization', 'set_id',
                'blt_status', 'pre_hil_status',
                'vibaration_status',
                'post_hil_status',
                'fgt_status',
                'final_integration_status',
                'bhd_status', 'fqm_status',
                'qm_certification_status', 'remarks')

        TABLE_COL_NAMES = (
            "System", "Organization", "Set ID", "BLT Status", "Pre-Hil Status", "Vibration Status", "Post-Hil Status",
            "FGT Status", "Final Integration Status", "BHD Status", "FQM Status", "QM Certification Status",
            "Remarks")

        pdf = FPDF('L', 'mm', 'Legal')
        pdf.add_page()
        pdf.set_font('courier', 'B', 26)
        pdf.cell(330, 10, 'Relifing Report', border=0, align='C', ln=2)
        pdf.cell(40, 10, '', 0, 1)
        pdf.set_font("Times", size=10)
        line_height = pdf.font_size * 3.2
        col_width = pdf.epw / 14

        def render_table_header():
            pdf.set_font(style="B")
            for col_name in TABLE_COL_NAMES:
                pdf.multi_cell(col_width, line_height, col_name, border=1, align='C', ln=3,
                               max_line_height=pdf.font_size)
            pdf.ln(line_height)
            pdf.set_font(style="")

        render_table_header()

        lh_list = []
        use_default_height = 0
        for row in data:
            for datum in row:
                dd = str(datum)
                word_list = dd.split()
                number_of_words = len(word_list)
                if number_of_words > 2:
                    use_default_height = 1
                    new_line_height = pdf.font_size * 2 * (number_of_words / 2)
            if not use_default_height:
                lh_list.append(line_height)
            else:
                lh_list.append(new_line_height)
                use_default_height = 0

        for j, row in enumerate(data):
            line_height = lh_list[j]
            if pdf.will_page_break(line_height):
                render_table_header()
            for col_num in range(len(row)):
                line_height = lh_list[j]
                pdf.multi_cell(col_width, line_height, f"{row[col_num]}", border=1, align='C', ln=3,
                               max_line_height=pdf.font_size)
            pdf.ln(line_height)
        pdf.output('RelifingReport.pdf')
        return FileResponse(open('RelifingReport.pdf', 'rb'), as_attachment=True, content_type='application/pdf')

    @staticmethod
    def GetRelifingExcelList(request):
        data = []
        status = request.query_params['status']
        if status == 'Total':
            data = RelifingSystemStatus.objects.all().order_by('-id').values_list('system', 'organization', 'set_id',
                                                                                  'blt_status', 'pre_hil_status',
                                                                                  'vibaration_status',
                                                                                  'post_hil_status',
                                                                                  'fgt_status',
                                                                                  'final_integration_status',
                                                                                  'bhd_status', 'fqm_status',
                                                                                  'qm_certification_status', 'remarks')
        else:
            data = RelifingSystemStatus.objects.filter(qm_certification_status=status).order_by('-id').values_list(
                'system', 'organization', 'set_id',
                'blt_status', 'pre_hil_status',
                'vibaration_status',
                'post_hil_status',
                'fgt_status',
                'final_integration_status',
                'bhd_status', 'fqm_status',
                'qm_certification_status', 'remarks')

        response = HttpResponse(content_type='application/ms-excel')
        response['Content-Disposition'] = 'attachment; filename="DocumentExcelSheet.xls"'
        work_book = xlwt.Workbook(encoding='utf-8')
        work_sheet = work_book.add_sheet('ExcelSheet')
        row_num = 0
        font_style = xlwt.XFStyle()
        font_style.font.bold = True
        TABLE_COL_NAMES = (
            "System", "Organization", "Set ID", "BLT Status", "Pre-Hil Status", "Vibration Status", "Post-Hil Status",
            "FGT Status", "Final Integration Status", "BHD Status", "FQM Status", "QM Certification Status",
            "Remarks")

        for col_num in range(len(TABLE_COL_NAMES)):
            work_sheet.write(row_num, col_num, TABLE_COL_NAMES[col_num], font_style)
        font_style = xlwt.XFStyle()

        for row in data:
            row_num += 1
            for col_num in range(len(row)):
                work_sheet.write(row_num, col_num, str(row[col_num]), font_style)
        work_book.save(response)
        return response

    @staticmethod
    def getSystemMonitoringDashboardCount(request):
        try:

            def get_filter(field_name, filter_condition, filter_value):
                # thanks to the below post
                # https://stackoverflow.com/questions/310732/in-django-how-does-one-filter-a-queryset-with-dynamic-field-lookups
                # the idea to this below logic is very similar to that in the above mentioned post
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

            filter_objects = Q()

            data = []
            prod_blt_ok = 0
            prod_blt_observation = 0
            prod_blt_up = 0
            prod_blt_hault = 0
            year = request.query_params['selected_year']
            org = request.query_params['selected_org']
            type = request.query_params['selected_type']
            system = request.query_params['selected_system']

            # create dynamic filter
            if year != '':
                filter_objects &= get_filter(
                    'testing_date__year', 'equal',
                    year)
            if org != '':
                filter_objects &= get_filter(
                    'organization', 'equal',
                    org)
            if type != '':
                filter_objects &= get_filter(
                    'sys_type', 'equal',
                    type)
            if system != '':
                filter_objects &= get_filter(
                    'system', 'equal',
                    system)

            # production system count
            prod_blt_oklist = ProductionSystemStatus.objects.filter(filter_objects, blt_status='OK(same stage)')
            prod_blt_oklistNext = ProductionSystemStatus.objects.filter(filter_objects, blt_status='OK(next stage)')
            prod_blt_observationlist = ProductionSystemStatus.objects.filter(filter_objects,
                                                                             blt_status='Observation(same stage)')
            prod_blt_observationlist_next = ProductionSystemStatus.objects.filter(filter_objects,
                                                                                  blt_status='Observation(next stage)')
            prod_blt_uplist = ProductionSystemStatus.objects.filter(filter_objects, blt_status='Under process')
            prod_blt_haultlist = ProductionSystemStatus.objects.filter(filter_objects, blt_status='Halt')

            prod_prehil_oklist = ProductionSystemStatus.objects.filter(filter_objects, pre_hil_status='OK(same stage)')
            prod_prehil_oklistNext = ProductionSystemStatus.objects.filter(filter_objects,
                                                                           pre_hil_status='OK(next stage)')
            prod_prehil_observationlist = ProductionSystemStatus.objects.filter(filter_objects,
                                                                                pre_hil_status='Observation(same stage)')
            prod_prehil_observationlistNext = ProductionSystemStatus.objects.filter(filter_objects,
                                                                                    pre_hil_status='Observation(next stage)')
            prod_prehil_uplist = ProductionSystemStatus.objects.filter(filter_objects, pre_hil_status='Under process')
            prod_prehil_haultlist = ProductionSystemStatus.objects.filter(filter_objects, pre_hil_status='Halt')

            prod_posthil_oklist = ProductionSystemStatus.objects.filter(filter_objects,
                                                                        post_hil_status='OK(same stage)')
            prod_posthil_oklistNext = ProductionSystemStatus.objects.filter(filter_objects,
                                                                            post_hil_status='OK(next stage)')
            prod_posthil_observationlist = ProductionSystemStatus.objects.filter(filter_objects,
                                                                                 post_hil_status='Observation(same stage)')
            prod_posthil_observationlistNext = ProductionSystemStatus.objects.filter(filter_objects,
                                                                                     post_hil_status='Observation(next stage)')
            prod_posthil_uplist = ProductionSystemStatus.objects.filter(filter_objects, post_hil_status='Under process')
            prod_posthil_haultlist = ProductionSystemStatus.objects.filter(filter_objects, post_hil_status='Halt')

            prod_finalintegration_oklist = ProductionSystemStatus.objects.filter(filter_objects,
                                                                                 final_integration_status='OK(same stage)')
            prod_finalintegration_oklistNext = ProductionSystemStatus.objects.filter(filter_objects,
                                                                                     final_integration_status='OK(next stage)')

            prod_finalintegration_observationlist = ProductionSystemStatus.objects.filter(filter_objects,
                                                                                          final_integration_status='Observation(same stage)')

            prod_finalintegration_observationlistNext = ProductionSystemStatus.objects.filter(filter_objects,
                                                                                              final_integration_status='Observation(next stage)')

            prod_finalintegration_completelist = ProductionSystemStatus.objects.filter(filter_objects,
                                                                                       final_integration_status='Completed')
            prod_finalintegration_haultlist = ProductionSystemStatus.objects.filter(filter_objects,
                                                                                    final_integration_status='Halt')

            prod_vibration_oklist = ProductionSystemStatus.objects.filter(filter_objects,
                                                                          vibaration_status='OK(same stage)')
            prod_vibration_oklistNext = ProductionSystemStatus.objects.filter(filter_objects,
                                                                              vibaration_status='OK(next stage)')

            prod_vibration_observationlist = ProductionSystemStatus.objects.filter(filter_objects,
                                                                                   vibaration_status='Observation(same stage)')
            prod_vibration_observationlistNext = ProductionSystemStatus.objects.filter(filter_objects,
                                                                                       vibaration_status='Observation(next stage)')
            prod_vibration_uplist = ProductionSystemStatus.objects.filter(filter_objects,
                                                                          vibaration_status='Under process')
            prod_vibration_haultlist = ProductionSystemStatus.objects.filter(filter_objects, vibaration_status='Halt')

            prod_cg_oklist = ProductionSystemStatus.objects.filter(filter_objects,
                                                                   cgbalancing_date_status='OK(same stage)')

            prod_cg_oklistNext = ProductionSystemStatus.objects.filter(filter_objects,
                                                                       cgbalancing_date_status='OK(next stage)')

            prod_cg_observationlist = ProductionSystemStatus.objects.filter(filter_objects,
                                                                            cgbalancing_date_status='Observation(same stage)')
            prod_cg_observationlistNext = ProductionSystemStatus.objects.filter(filter_objects,
                                                                                cgbalancing_date_status='Observation(next stage)')
            prod_cg_uplist = ProductionSystemStatus.objects.filter(filter_objects,
                                                                   cgbalancing_date_status='Under process')
            prod_cg_haultlist = ProductionSystemStatus.objects.filter(filter_objects, cgbalancing_date_status='Halt')

            prod_fgt_oklist = ProductionSystemStatus.objects.filter(filter_objects, fgt_status='OK(same stage)')
            prod_fgt_oklistNext = ProductionSystemStatus.objects.filter(filter_objects, fgt_status='OK(next stage)')
            prod_fgt_observationlist = ProductionSystemStatus.objects.filter(filter_objects,
                                                                             fgt_status='Observation(same stage)')
            prod_fgt_observationlistNext = ProductionSystemStatus.objects.filter(filter_objects,
                                                                                 fgt_status='Observation(next stage)')
            prod_fgt_uplist = ProductionSystemStatus.objects.filter(filter_objects,
                                                                    fgt_status='Under process')
            prod_fgt_haultlist = ProductionSystemStatus.objects.filter(filter_objects, fgt_status='Halt')

            prod_bhd_notsubmit = ProductionSystemStatus.objects.filter(filter_objects,
                                                                       bhd_status='BHD not submitted for QM audit')
            prod_bhd_inprocess = ProductionSystemStatus.objects.filter(filter_objects,
                                                                       bhd_status='Audit in-process')
            prod_bhd_qmforwarded = ProductionSystemStatus.objects.filter(filter_objects,
                                                                         bhd_status='QM observations forwarded')

            prod_fqm_planned = ProductionSystemStatus.objects.filter(filter_objects,
                                                                     fqm_status='Planned')

            prod_fqm_conducted = ProductionSystemStatus.objects.filter(filter_objects,
                                                                       fqm_status='Conducted')

            prod_qmc_issued = ProductionSystemStatus.objects.filter(filter_objects,
                                                                    qm_certification_status='QM certification issued')

            prod_qmc_in_process = ProductionSystemStatus.objects.filter(filter_objects,
                                                                        qm_certification_status='QM certification in-process')

            prod_enduser_oklist = ProductionSystemStatus.objects.filter(filter_objects, enduser_status='OK(same stage)')
            prod_enduser_oklistNext = ProductionSystemStatus.objects.filter(filter_objects,
                                                                            enduser_status='OK(next stage)')
            prod_enduser_observationlist = ProductionSystemStatus.objects.filter(filter_objects,
                                                                                 enduser_status='Observation(same stage)')
            prod_enduser_observationlistNext = ProductionSystemStatus.objects.filter(filter_objects,
                                                                                     enduser_status='Observation(next stage)')
            prod_enduser_uplist = ProductionSystemStatus.objects.filter(filter_objects,
                                                                        enduser_status='Completed')
            prod_enduser_haultlist = ProductionSystemStatus.objects.filter(filter_objects, enduser_status='Halt')




            # flight system count
            flight_blt_oklist = FlightSystemStatus.objects.filter(filter_objects, blt_status='OK(same stage)')
            flight_blt_oklistNext = FlightSystemStatus.objects.filter(filter_objects, blt_status='OK(next stage)')
            flight_blt_observationlist = FlightSystemStatus.objects.filter(filter_objects,
                                                                           blt_status='Observation(same stage)')
            flight_blt_observationlist_next = FlightSystemStatus.objects.filter(filter_objects,
                                                                                blt_status='Observation(next stage)')
            flight_blt_uplist = FlightSystemStatus.objects.filter(filter_objects, blt_status='Under process')
            flight_blt_haultlist = FlightSystemStatus.objects.filter(filter_objects, blt_status='Halt')

            flight_prehil_oklist = FlightSystemStatus.objects.filter(filter_objects, pre_hil_status='OK(same stage)')
            flight_prehil_oklistNext = FlightSystemStatus.objects.filter(filter_objects,
                                                                         pre_hil_status='OK(next stage)')
            flight_prehil_observationlist = FlightSystemStatus.objects.filter(filter_objects,
                                                                              pre_hil_status='Observation(same stage)')
            flight_prehil_observationlistNext = FlightSystemStatus.objects.filter(filter_objects,
                                                                                  pre_hil_status='Observation(next stage)')
            flight_prehil_uplist = FlightSystemStatus.objects.filter(filter_objects, pre_hil_status='Under process')
            flight_prehil_haultlist = FlightSystemStatus.objects.filter(filter_objects, pre_hil_status='Halt')

            flight_posthil_oklist = FlightSystemStatus.objects.filter(filter_objects, post_hil_status='OK(same stage)')
            flight_posthil_oklistNext = FlightSystemStatus.objects.filter(filter_objects,
                                                                          post_hil_status='OK(next stage)')
            flight_posthil_observationlist = FlightSystemStatus.objects.filter(filter_objects,
                                                                               post_hil_status='Observation(same stage)')
            flight_posthil_observationlistNext = FlightSystemStatus.objects.filter(filter_objects,
                                                                                   post_hil_status='Observation(next stage)')
            flight_posthil_uplist = FlightSystemStatus.objects.filter(filter_objects, post_hil_status='Under process')
            flight_posthil_haultlist = FlightSystemStatus.objects.filter(filter_objects, post_hil_status='Halt')

            flight_finalintegration_oklist = FlightSystemStatus.objects.filter(filter_objects,
                                                                               final_integration_status='OK(same stage)')
            flight_finalintegration_oklistNext = FlightSystemStatus.objects.filter(filter_objects,
                                                                                   final_integration_status='OK(next stage)')

            flight_finalintegration_observationlist = FlightSystemStatus.objects.filter(filter_objects,
                                                                                        final_integration_status='Observation(same stage)')
            flight_finalintegration_observationlistNext = FlightSystemStatus.objects.filter(filter_objects,
                                                                                            final_integration_status='Observation(next stage)')
            flight_finalintegration_completelist = FlightSystemStatus.objects.filter(filter_objects,
                                                                                     final_integration_status='Completed')
            flight_finalintegration_haultlist = FlightSystemStatus.objects.filter(filter_objects,
                                                                                  final_integration_status='Halt')

            flight_vibration_oklist = FlightSystemStatus.objects.filter(filter_objects,
                                                                        vibaration_status='OK(same stage)')
            flight_vibration_oklistNext = FlightSystemStatus.objects.filter(filter_objects,
                                                                            vibaration_status='OK(next stage)')
            flight_vibration_observationlist = FlightSystemStatus.objects.filter(filter_objects,
                                                                                 vibaration_status='Observation(same stage)')
            flight_vibration_observationlistNext = FlightSystemStatus.objects.filter(filter_objects,
                                                                                     vibaration_status='Observation(next stage)')
            flight_vibration_uplist = FlightSystemStatus.objects.filter(filter_objects,
                                                                        vibaration_status='Under process')
            flight_vibration_haultlist = FlightSystemStatus.objects.filter(filter_objects, vibaration_status='Halt')

            flight_cg_oklist = FlightSystemStatus.objects.filter(filter_objects,
                                                                 cgbalancing_date_status='OK(same stage)')
            flight_cg_oklistNext = FlightSystemStatus.objects.filter(filter_objects,
                                                                     cgbalancing_date_status='OK(next stage)')
            flight_cg_observationlist = FlightSystemStatus.objects.filter(filter_objects,
                                                                          cgbalancing_date_status='Observation(same stage)')
            flight_cg_observationlistNext = FlightSystemStatus.objects.filter(filter_objects,
                                                                              cgbalancing_date_status='Observation(next stage)')
            flight_cg_uplist = FlightSystemStatus.objects.filter(filter_objects,
                                                                 cgbalancing_date_status='Under process')
            flight_cg_haultlist = FlightSystemStatus.objects.filter(filter_objects, cgbalancing_date_status='Halt')

            flight_fgt_oklist = FlightSystemStatus.objects.filter(filter_objects, fgt_status='OK(same stage)')
            flight_fgt_oklistNext = FlightSystemStatus.objects.filter(filter_objects, fgt_status='OK(next stage)')
            flight_fgt_observationlist = FlightSystemStatus.objects.filter(filter_objects,
                                                                           fgt_status='Observation(same stage)')
            flight_fgt_observationlistNext = FlightSystemStatus.objects.filter(filter_objects,
                                                                               fgt_status='Observation(next stage)')
            flight_fgt_uplist = FlightSystemStatus.objects.filter(filter_objects,
                                                                  fgt_status='Under process')
            flight_fgt_haultlist = FlightSystemStatus.objects.filter(filter_objects, fgt_status='Halt')

            flight_bhd_notsubmit = FlightSystemStatus.objects.filter(filter_objects,
                                                                     bhd_status='BHD not submitted for QM audit')
            flight_bhd_inprocess = FlightSystemStatus.objects.filter(filter_objects,
                                                                     bhd_status='Audit in-process')
            flight_bhd_qmforwarded = FlightSystemStatus.objects.filter(filter_objects,
                                                                       bhd_status='QM observations forwarded')

            flight_fqm_planned = FlightSystemStatus.objects.filter(filter_objects,
                                                                   fqm_status='Planned')

            flight_fqm_conducted = FlightSystemStatus.objects.filter(filter_objects,
                                                                     fqm_status='Conducted')

            flight_qmc_issued = FlightSystemStatus.objects.filter(filter_objects,
                                                                  qm_certification_status='QM certification issued')

            flight_qmc_in_process = FlightSystemStatus.objects.filter(filter_objects,
                                                                      qm_certification_status='QM certification in-process')

            flight_launch_oklist = FlightSystemStatus.objects.filter(filter_objects, launchact_status='OK(same stage)')
            flight_launch_oklistNext = FlightSystemStatus.objects.filter(filter_objects,
                                                                         launchact_status='OK(next stage)')
            flight_launch_observationlist = FlightSystemStatus.objects.filter(filter_objects,
                                                                              launchact_status='Observation(same stage)')
            flight_launch_observationlistNext = FlightSystemStatus.objects.filter(filter_objects,
                                                                                  launchact_status='Observation(next stage)')
            flight_launch_uplist = FlightSystemStatus.objects.filter(filter_objects,
                                                                     launchact_status='Completed')
            flight_launch_haultlist = FlightSystemStatus.objects.filter(filter_objects, launchact_status='Halt')


            # relifing system count
            relifing_blt_oklist = RelifingSystemStatus.objects.filter(filter_objects, blt_status='OK(same stage)')
            relifing_blt_oklistNext = RelifingSystemStatus.objects.filter(filter_objects, blt_status='OK(next stage)')
            relifing_blt_observationlist = RelifingSystemStatus.objects.filter(filter_objects,
                                                                               blt_status='Observation(same stage)')
            relifing_blt_observationlistNext = RelifingSystemStatus.objects.filter(filter_objects,
                                                                                   blt_status='Observation(next stage)')
            relifing_blt_uplist = RelifingSystemStatus.objects.filter(filter_objects, blt_status='Under process')
            relifing_blt_haultlist = RelifingSystemStatus.objects.filter(filter_objects, blt_status='Halt')

            relifing_prehil_oklist = RelifingSystemStatus.objects.filter(filter_objects,
                                                                         pre_hil_status='OK(same stage)')
            relifing_prehil_oklistNext = RelifingSystemStatus.objects.filter(filter_objects,
                                                                             pre_hil_status='OK(next stage)')
            relifing_prehil_observationlist = RelifingSystemStatus.objects.filter(filter_objects,
                                                                                  pre_hil_status='Observation(same stage)')
            relifing_prehil_observationlistNext = RelifingSystemStatus.objects.filter(filter_objects,
                                                                                      pre_hil_status='Observation(next stage)')
            relifing_prehil_uplist = RelifingSystemStatus.objects.filter(filter_objects, pre_hil_status='Under process')
            relifing_prehil_haultlist = RelifingSystemStatus.objects.filter(filter_objects, pre_hil_status='Halt')

            relifing_posthil_oklist = RelifingSystemStatus.objects.filter(filter_objects,
                                                                          post_hil_status='OK(same stage)')
            relifing_posthil_oklistNext = RelifingSystemStatus.objects.filter(filter_objects,
                                                                              post_hil_status='OK(next stage)')
            relifing_posthil_observationlist = RelifingSystemStatus.objects.filter(filter_objects,
                                                                                   post_hil_status='Observation(same stage)')
            relifing_posthil_observationlistNext = RelifingSystemStatus.objects.filter(filter_objects,
                                                                                       post_hil_status='Observation(next stage)')
            relifing_posthil_uplist = RelifingSystemStatus.objects.filter(filter_objects,
                                                                          post_hil_status='Under process')
            relifing_posthil_haultlist = RelifingSystemStatus.objects.filter(filter_objects, post_hil_status='Halt')

            relifing_finalintegration_oklist = RelifingSystemStatus.objects.filter(filter_objects,
                                                                                   final_integration_status='OK(same stage)')
            relifing_finalintegration_oklistNext = RelifingSystemStatus.objects.filter(filter_objects,
                                                                                       final_integration_status='OK(next stage)')
            relifing_finalintegration_observationlist = RelifingSystemStatus.objects.filter(filter_objects,
                                                                                            final_integration_status='Observation(same stage)')
            relifing_finalintegration_observationlistNext = RelifingSystemStatus.objects.filter(filter_objects,
                                                                                                final_integration_status='Observation(next stage)')
            relifing_finalintegration_completelist = RelifingSystemStatus.objects.filter(filter_objects,
                                                                                         final_integration_status='Completed')
            relifing_finalintegration_haultlist = RelifingSystemStatus.objects.filter(filter_objects,
                                                                                      final_integration_status='Halt')

            relifing_vibration_oklist = RelifingSystemStatus.objects.filter(filter_objects,
                                                                            vibaration_status='OK(same stage)')
            relifing_vibration_oklistNext = RelifingSystemStatus.objects.filter(filter_objects,
                                                                                vibaration_status='OK(next stage)')
            relifing_vibration_observationlist = RelifingSystemStatus.objects.filter(filter_objects,
                                                                                     vibaration_status='Observation(same stage)')
            relifing_vibration_observationlistNext = RelifingSystemStatus.objects.filter(filter_objects,
                                                                                         vibaration_status='Observation(next stage)')
            relifing_vibration_uplist = RelifingSystemStatus.objects.filter(filter_objects,
                                                                            vibaration_status='Under process')
            relifing_vibration_haultlist = RelifingSystemStatus.objects.filter(filter_objects, vibaration_status='Halt')

            relifing_cg_oklist = RelifingSystemStatus.objects.filter(filter_objects,
                                                                     cgbalancing_date_status='OK(same stage)')
            relifing_cg_oklistNext = RelifingSystemStatus.objects.filter(filter_objects,
                                                                         cgbalancing_date_status='OK(next stage)')
            relifing_cg_observationlist = RelifingSystemStatus.objects.filter(filter_objects,
                                                                              cgbalancing_date_status='Observation(same stage)')
            relifing_cg_observationlistNext = RelifingSystemStatus.objects.filter(filter_objects,
                                                                                  cgbalancing_date_status='Observation(next stage)')
            relifing_cg_uplist = RelifingSystemStatus.objects.filter(filter_objects,
                                                                     cgbalancing_date_status='Under process')
            relifing_cg_haultlist = RelifingSystemStatus.objects.filter(filter_objects, cgbalancing_date_status='Halt')

            relifing_fgt_oklist = RelifingSystemStatus.objects.filter(filter_objects, fgt_status='OK(same stage)')
            relifing_fgt_oklistNext = RelifingSystemStatus.objects.filter(filter_objects, fgt_status='OK(next stage)')
            relifing_fgt_observationlist = RelifingSystemStatus.objects.filter(filter_objects,
                                                                               fgt_status='Observation(same stage)')
            relifing_fgt_observationlistNext = RelifingSystemStatus.objects.filter(filter_objects,
                                                                                   fgt_status='Observation(next stage)')
            relifing_fgt_uplist = RelifingSystemStatus.objects.filter(filter_objects,
                                                                      fgt_status='Under process')
            relifing_fgt_haultlist = RelifingSystemStatus.objects.filter(filter_objects, fgt_status='Halt')

            relifing_bhd_notsubmit = RelifingSystemStatus.objects.filter(filter_objects,
                                                                         bhd_status='BHD not submitted for QM audit')
            relifing_bhd_inprocess = RelifingSystemStatus.objects.filter(filter_objects,
                                                                         bhd_status='Audit in-process')
            relifing_bhd_qmforwarded = RelifingSystemStatus.objects.filter(filter_objects,
                                                                           bhd_status='QM observations forwarded')

            relifing_fqm_planned = RelifingSystemStatus.objects.filter(filter_objects,
                                                                       fqm_status='Planned')

            relifing_fqm_conducted = RelifingSystemStatus.objects.filter(filter_objects,
                                                                         fqm_status='Conducted')

            relifing_qmc_issued = RelifingSystemStatus.objects.filter(filter_objects,
                                                                      qm_certification_status='QM certification issued')

            relifing_qmc_in_process = RelifingSystemStatus.objects.filter(filter_objects,
                                                                          qm_certification_status='QM certification in-process')

            relifing_enduser_oklist = RelifingSystemStatus.objects.filter(filter_objects,
                                                                          enduser_status='OK(same stage)')
            relifing_enduser_oklistNext = RelifingSystemStatus.objects.filter(filter_objects,
                                                                              enduser_status='OK(next stage)')
            relifing_enduser_observationlist = RelifingSystemStatus.objects.filter(filter_objects,
                                                                                   enduser_status='Observation(same stage)')
            relifing_enduser_observationlistNext = RelifingSystemStatus.objects.filter(filter_objects,
                                                                                       enduser_status='Observation(next stage)')
            relifing_enduser_uplist = RelifingSystemStatus.objects.filter(filter_objects,
                                                                          enduser_status='Completed')
            relifing_enduser_haultlist = RelifingSystemStatus.objects.filter(filter_objects, enduser_status='Halt')

            dist = {
                'prod_blt_ok': prod_blt_oklist.count(),
                'prod_blt_ok_next': prod_blt_oklistNext.count(),
                'prod_blt_observation': prod_blt_observationlist.count(),
                'prod_blt_observation_next': prod_blt_observationlist_next.count(),
                'prod_blt_up': prod_blt_uplist.count(),
                'prod_blt_hault': prod_blt_haultlist.count(),
                'flight_blt_oklist': flight_blt_oklist.count(),
                'flight_blt_oklist_next': flight_blt_oklistNext.count(),
                'flight_blt_observationlist': flight_blt_observationlist.count(),
                'flight_blt_observationlist_next': flight_blt_observationlist_next.count(),
                'flight_blt_uplist': flight_blt_uplist.count(),
                'flight_blt_haultlist': flight_blt_haultlist.count(),
                'relifing_blt_oklist': relifing_blt_oklist.count(),
                'relifing_blt_oklist_next': relifing_blt_oklistNext.count(),
                'relifing_blt_observationlist': relifing_blt_observationlist.count(),
                'relifing_blt_observationlist_next': relifing_blt_observationlistNext.count(),
                'relifing_blt_uplist': relifing_blt_uplist.count(),
                'relifing_blt_haultlist': relifing_blt_haultlist.count(),
                'prod_prehil_oklist': prod_prehil_oklist.count(),
                'prod_prehil_oklist_next': prod_prehil_oklistNext.count(),
                'prod_prehil_observationlist': prod_prehil_observationlist.count(),
                'prod_prehil_observationlist_next': prod_prehil_observationlistNext.count(),
                'prod_prehil_uplist': prod_prehil_uplist.count(),
                'prod_prehil_haultlist': prod_prehil_haultlist.count(),
                'flight_prehil_oklist': flight_prehil_oklist.count(),
                'flight_prehil_oklist_next': flight_prehil_oklistNext.count(),
                'flight_prehil_observationlist': flight_prehil_observationlist.count(),
                'flight_prehil_observationlist_next': flight_prehil_observationlistNext.count(),
                'flight_prehil_uplist': flight_prehil_uplist.count(),
                'flight_prehil_haultlist': flight_prehil_haultlist.count(),
                'relifing_prehil_oklist': relifing_prehil_oklist.count(),
                'relifing_prehil_oklist_next': relifing_prehil_oklistNext.count(),
                'relifing_prehil_observationlist': relifing_prehil_observationlist.count(),
                'relifing_prehil_observationlist_next': relifing_prehil_observationlistNext.count(),
                'relifing_prehil_uplist': relifing_prehil_uplist.count(),
                'relifing_prehil_haultlist': relifing_prehil_haultlist.count(),
                'prod_posthil_oklist': prod_posthil_oklist.count(),
                'prod_posthil_oklist_next': prod_posthil_oklistNext.count(),
                'prod_posthil_observationlist': prod_posthil_observationlist.count(),
                'prod_posthil_observationlist_next': prod_posthil_observationlistNext.count(),
                'prod_posthil_uplist': prod_posthil_uplist.count(),
                'prod_posthil_haultlist': prod_posthil_haultlist.count(),
                'flight_posthil_oklist': flight_posthil_oklist.count(),
                'flight_posthil_oklist_next': flight_posthil_oklistNext.count(),
                'flight_posthil_observationlist': flight_posthil_observationlist.count(),
                'flight_posthil_observationlist_next': flight_posthil_observationlistNext.count(),
                'flight_posthil_uplist': flight_posthil_uplist.count(),
                'flight_posthil_haultlist': flight_posthil_haultlist.count(),
                'relifing_posthil_oklist': relifing_posthil_oklist.count(),
                'relifing_posthil_oklist_next': relifing_posthil_oklistNext.count(),
                'relifing_posthil_observationlist': relifing_posthil_observationlist.count(),
                'relifing_posthil_observationlist_next': relifing_posthil_observationlistNext.count(),
                'relifing_posthil_uplist': relifing_posthil_uplist.count(),
                'relifing_posthil_haultlist': relifing_posthil_haultlist.count(),
                'prod_finalintegration_oklist': prod_finalintegration_oklist.count(),
                'prod_finalintegration_oklist_next': prod_finalintegration_oklistNext.count(),
                'prod_finalintegration_observationlist': prod_finalintegration_observationlist.count(),
                'prod_finalintegration_observationlist_next': prod_finalintegration_observationlistNext.count(),
                'prod_finalintegration_completelist': prod_finalintegration_completelist.count(),
                'prod_finalintegration_haultlist': prod_finalintegration_haultlist.count(),
                'flight_finalintegration_oklist': flight_finalintegration_oklist.count(),
                'flight_finalintegration_oklist_next': flight_finalintegration_oklistNext.count(),
                'flight_finalintegration_observationlist': flight_finalintegration_observationlist.count(),
                'flight_finalintegration_observationlist_next': flight_finalintegration_observationlistNext.count(),
                'flight_finalintegration_completelist': flight_finalintegration_completelist.count(),
                'flight_finalintegration_haultlist': flight_finalintegration_haultlist.count(),
                'relifing_finalintegration_oklist': relifing_finalintegration_oklist.count(),
                'relifing_finalintegration_oklist_next': relifing_finalintegration_oklistNext.count(),
                'relifing_finalintegration_observationlist': relifing_finalintegration_observationlist.count(),
                'relifing_finalintegration_observationlist_next': relifing_finalintegration_observationlistNext.count(),
                'relifing_finalintegration_completelist': relifing_finalintegration_completelist.count(),
                'relifing_finalintegration_haultlist': relifing_finalintegration_haultlist.count(),
                'relifing_vibration_oklist': relifing_vibration_oklist.count(),
                'relifing_vibration_oklist_next': relifing_vibration_oklistNext.count(),
                'relifing_vibration_observationlist': relifing_vibration_observationlist.count(),
                'relifing_vibration_observationlist_next': relifing_vibration_observationlistNext.count(),
                'relifing_vibration_uplist': relifing_vibration_uplist.count(),
                'relifing_vibration_haultlist': relifing_vibration_haultlist.count(),
                'flight_vibration_oklist': flight_vibration_oklist.count(),
                'flight_vibration_oklist_next': flight_vibration_oklistNext.count(),
                'flight_vibration_observationlist': flight_vibration_observationlist.count(),
                'flight_vibration_observationlist_next': flight_vibration_observationlistNext.count(),
                'flight_vibration_uplist': flight_vibration_uplist.count(),
                'flight_vibration_haultlist': flight_vibration_haultlist.count(),
                'prod_vibration_oklist': prod_vibration_oklist.count(),
                'prod_vibration_oklist_next': prod_vibration_oklistNext.count(),
                'prod_vibration_observationlist': prod_vibration_observationlist.count(),
                'prod_vibration_observationlist_next': prod_vibration_observationlistNext.count(),
                'prod_vibration_uplist': prod_vibration_uplist.count(),
                'prod_vibration_haultlist': prod_vibration_haultlist.count(),
                'prod_cg_oklist': prod_cg_oklist.count(),
                'prod_cg_oklist_next': prod_cg_oklistNext.count(),
                'prod_cg_observationlist': prod_cg_observationlist.count(),
                'prod_cg_observationlist_next': prod_cg_observationlistNext.count(),
                'prod_cg_uplist': prod_cg_uplist.count(),
                'prod_cg_haultlist': prod_cg_haultlist.count(),
                'flight_cg_oklist': flight_cg_oklist.count(),
                'flight_cg_oklist_next': flight_cg_oklistNext.count(),
                'flight_cg_observationlist': flight_cg_observationlist.count(),
                'flight_cg_observationlist_next': flight_cg_observationlistNext.count(),
                'flight_cg_uplist': flight_cg_uplist.count(),
                'flight_cg_haultlist': flight_cg_haultlist.count(),
                'relifing_cg_oklist': relifing_cg_oklist.count(),
                'relifing_cg_oklist_next': relifing_cg_oklistNext.count(),
                'relifing_cg_observationlist': relifing_cg_observationlist.count(),
                'relifing_cg_observationlist_next': relifing_cg_observationlistNext.count(),
                'relifing_cg_uplist': relifing_cg_uplist.count(),
                'relifing_cg_haultlist': relifing_cg_haultlist.count(),
                'prod_fgt_oklist': prod_fgt_oklist.count(),
                'prod_fgt_oklist_next': prod_fgt_oklistNext.count(),
                'prod_fgt_observationlist': prod_fgt_observationlist.count(),
                'prod_fgt_observationlist_next': prod_fgt_observationlistNext.count(),
                'prod_fgt_uplist': prod_fgt_uplist.count(),
                'prod_fgt_haultlist': prod_fgt_haultlist.count(),
                'flight_fgt_oklist': flight_fgt_oklist.count(),
                'flight_fgt_oklist_next': flight_fgt_oklistNext.count(),
                'flight_fgt_observationlist': flight_fgt_observationlist.count(),
                'flight_fgt_observationlist_next': flight_fgt_observationlistNext.count(),
                'flight_fgt_uplist': flight_fgt_uplist.count(),
                'flight_fgt_haultlist': flight_fgt_haultlist.count(),
                'relifing_fgt_oklist': relifing_fgt_oklist.count(),
                'relifing_fgt_oklist_next': relifing_fgt_oklistNext.count(),
                'relifing_fgt_observationlist': relifing_fgt_observationlist.count(),
                'relifing_fgt_observationlist_next': relifing_fgt_observationlistNext.count(),
                'relifing_fgt_uplist': relifing_fgt_uplist.count(),
                'relifing_fgt_haultlist': relifing_fgt_haultlist.count(),
                'prod_bhd_notsubmit': prod_bhd_notsubmit.count(),
                'prod_bhd_inprocess': prod_bhd_inprocess.count(),
                'prod_bhd_qmforwarded': prod_bhd_qmforwarded.count(),
                'flight_bhd_notsubmit': flight_bhd_notsubmit.count(),
                'flight_bhd_inprocess': flight_bhd_inprocess.count(),
                'flight_bhd_qmforwarded': flight_bhd_qmforwarded.count(),
                'relifing_bhd_notsubmit': relifing_bhd_notsubmit.count(),
                'relifing_bhd_inprocess': relifing_bhd_inprocess.count(),
                'relifing_bhd_qmforwarded': relifing_bhd_qmforwarded.count(),
                'prod_fqm_planned': prod_fqm_planned.count(),
                'prod_fqm_conducted': prod_fqm_conducted.count(),
                'flight_fqm_planned': flight_fqm_planned.count(),
                'flight_fqm_conducted': flight_fqm_conducted.count(),
                'relifing_fqm_planned': relifing_fqm_planned.count(),
                'relifing_fqm_conducted': relifing_fqm_conducted.count(),
                'prod_qmc_issued': prod_qmc_issued.count(),
                'prod_qmc_in_process': prod_qmc_in_process.count(),
                'flight_qmc_issued': flight_qmc_issued.count(),
                'flight_qmc_in_process': flight_qmc_in_process.count(),
                'relifing_qmc_issued': relifing_qmc_issued.count(),
                'relifing_qmc_in_process': relifing_qmc_in_process.count(),
                'prod_enduser_oklist': prod_enduser_oklist.count(),
                'prod_enduser_oklist_next': prod_enduser_oklistNext.count(),
                'prod_enduser_observationlist': prod_enduser_observationlist.count(),
                'prod_enduser_observationlist_next': prod_enduser_observationlistNext.count(),
                'prod_enduser_uplist': prod_enduser_uplist.count(),
                'prod_enduser_haultlist': prod_enduser_haultlist.count(),
                'flight_launch_oklist': flight_launch_oklist.count(),
                'flight_launch_oklist_next': flight_launch_oklistNext.count(),
                'flight_launch_observationlist': flight_launch_observationlist.count(),
                'flight_launch_observationlist_next': flight_launch_observationlistNext.count(),
                'flight_launch_uplist': flight_launch_uplist.count(),
                'flight_launch_haultlist': flight_launch_haultlist.count(),
                'relifing_enduser_oklist': relifing_enduser_oklist.count(),
                'relifing_enduser_oklist_next': relifing_enduser_oklistNext.count(),
                'relifing_enduser_observationlist': relifing_enduser_observationlist.count(),
                'relifing_enduser_observationlist_next': relifing_enduser_observationlistNext.count(),
                'relifing_enduser_uplist': relifing_enduser_uplist.count(),
                'relifing_enduser_haultlist': relifing_enduser_haultlist.count()

            }

            return JsonResponse({'message': 'true', 'data': dist}, status=200)
        except Exception as e:
            print(e)
            return JsonResponse({'message': 'Sorry! No Task found.'}, status=500)
