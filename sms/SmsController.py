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
                prodModel.pre_hil_date = request['pre_hil_date']
                prodModel.pre_hil_status = request['pre_hil_status']
                prodModel.vibaration_date = request['vibaration_date']
                prodModel.vibaration_status = request['vibaration_status']
                prodModel.post_hil_date = request['post_hil_date']
                prodModel.post_hil_status = request['post_hil_status']
                prodModel.fgt_date = request['fgt_date']
                prodModel.fgt_status = request['fgt_status']
                prodModel.final_integration_date = request['final_integration_date']
                prodModel.final_integration_status = request['final_integration_status']
                prodModel.bhd_date = request['bhd_date']
                prodModel.bhd_status = request['bhd_status']
                prodModel.fqm_date = request['fqm_date']
                prodModel.fqm_status = request['fqm_status']
                prodModel.qm_certification_date = request['qm_certification_date']
                prodModel.qm_certification_status = request['qm_certification_status']
                prodModel.attachment = request['attachment']
                prodModel.remarks = request['remarks']
                prodModel.cgbalancing_date = request['cgbalancing_date']
                prodModel.cgbalancing_date_status = request['cgbalancing_date_status']
                prodModel.enduser_date = request['enduser_date']
                prodModel.enduser_status = request['enduser_status']
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
                        'qm_certification_status'] != get_prod.qm_certification_status:
                        print("add data in production history")
                        ProdHistoryModal = ProductionSystemStatusHistory()
                        ProdHistoryModal.prod_id = get_prod.id
                        ProdHistoryModal.system = get_prod.system
                        ProdHistoryModal.organization = get_prod.organization
                        ProdHistoryModal.set_id = get_prod.set_id
                        ProdHistoryModal.blt_date = get_prod.blt_date
                        ProdHistoryModal.blt_status = get_prod.blt_status
                        ProdHistoryModal.testing_date = get_prod.testing_date
                        ProdHistoryModal.sys_type = get_prod.sys_type
                        ProdHistoryModal.pre_hil_date = get_prod.pre_hil_date
                        ProdHistoryModal.pre_hil_status = get_prod.pre_hil_status
                        ProdHistoryModal.vibaration_date = get_prod.vibaration_date
                        ProdHistoryModal.vibaration_status = get_prod.vibaration_status
                        ProdHistoryModal.post_hil_date = get_prod.post_hil_date
                        ProdHistoryModal.post_hil_status = get_prod.post_hil_status
                        ProdHistoryModal.fgt_date = get_prod.fgt_date
                        ProdHistoryModal.fgt_status = get_prod.fgt_status
                        ProdHistoryModal.final_integration_date = get_prod.final_integration_date
                        ProdHistoryModal.final_integration_status = get_prod.final_integration_status
                        ProdHistoryModal.bhd_date = get_prod.bhd_date
                        ProdHistoryModal.bhd_status = get_prod.bhd_status
                        ProdHistoryModal.fqm_date = get_prod.fqm_date
                        ProdHistoryModal.fqm_status = get_prod.fqm_status
                        ProdHistoryModal.qm_certification_date = get_prod.qm_certification_date
                        ProdHistoryModal.qm_certification_status = get_prod.qm_certification_status
                        ProdHistoryModal.attachment = get_prod.attachment
                        ProdHistoryModal.remarks = get_prod.remarks
                        ProdHistoryModal.isActive = get_prod.isActive
                        ProdHistoryModal.cgbalancing_date = get_prod.cgbalancing_date
                        ProdHistoryModal.cgbalancing_date_status = get_prod.cgbalancing_date_status
                        ProdHistoryModal.enduser_date = get_prod.enduser_date
                        ProdHistoryModal.enduser_status = get_prod.enduser_status
                        ProdHistoryModal.save()

                        # history saved
                    get_prod.system = request['system']
                    get_prod.organization = request['organization']
                    get_prod.set_id = request['set_id']
                    get_prod.blt_date = request['blt_date']
                    get_prod.blt_status = request['blt_status']
                    get_prod.testing_date = request['testing_date']
                    get_prod.sys_type = request['sys_type']
                    get_prod.pre_hil_date = request['pre_hil_date']
                    get_prod.pre_hil_status = request['pre_hil_status']
                    get_prod.vibaration_date = request['vibaration_date']
                    get_prod.vibaration_status = request['vibaration_status']
                    get_prod.post_hil_date = request['post_hil_date']
                    get_prod.post_hil_status = request['post_hil_status']
                    get_prod.fgt_date = request['fgt_date']
                    get_prod.fgt_status = request['fgt_status']
                    get_prod.final_integration_date = request['final_integration_date']
                    get_prod.final_integration_status = request['final_integration_status']
                    get_prod.bhd_date = request['bhd_date']
                    get_prod.bhd_status = request['bhd_status']
                    get_prod.fqm_date = request['fqm_date']
                    get_prod.fqm_status = request['fqm_status']
                    get_prod.qm_certification_date = request['qm_certification_date']
                    get_prod.qm_certification_status = request['qm_certification_status']
                    if request['attachment'] != '':
                        get_prod.attachment = request['attachment']
                    get_prod.remarks = request['remarks']
                    get_prod.isActive = is_active
                    get_prod.cgbalancing_date = request['cgbalancing_date']
                    get_prod.cgbalancing_date_status = request['cgbalancing_date_status']
                    get_prod.enduser_date = request['enduser_date']
                    get_prod.enduser_status = request['enduser_status']
                    get_prod.save()
            return JsonResponse({'status': 'True', 'message': "Production Status Updated Successfully!"},
                                status=200)
        except Exception as e:
            return JsonResponse({'status': 'False', "message": "Status Not Saved"}, status=500)

    @staticmethod
    def GetProductionList(request):
        try:
            status = request.query_params['status']
            if status == 'Total':
                data = ProductionSystemStatus.objects.all()
                serializer = ProductionSystemSerialzer(data, many=True)
                return JsonResponse({'status': 'true', 'data': serializer.data}, status=200)
            else:
                data = ProductionSystemStatus.objects.filter(qm_certification_status=status)
                serializer = ProductionSystemSerialzer(data, many=True)
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

        try:
            id = request['id']
            if id == '0':
                flightModel.system = request['system']
                flightModel.organization = request['organization']
                flightModel.set_id = request['set_id']
                flightModel.blt_date = request['blt_date']
                flightModel.blt_status = request['blt_status']
                flightModel.pre_hil_date = request['pre_hil_date']
                flightModel.pre_hil_status = request['pre_hil_status']
                flightModel.vibaration_date = request['vibaration_date']
                flightModel.vibaration_status = request['vibaration_status']
                flightModel.post_hil_date = request['post_hil_date']
                flightModel.post_hil_status = request['post_hil_status']
                flightModel.fgt_date = request['fgt_date']
                flightModel.fgt_status = request['fgt_status']
                flightModel.final_integration_date = request['final_integration_date']
                flightModel.final_integration_status = request['final_integration_status']
                flightModel.bhd_date = request['bhd_date']
                flightModel.bhd_status = request['bhd_status']
                flightModel.fqm_date = request['fqm_date']
                flightModel.fqm_status = request['fqm_status']
                flightModel.qm_certification_date = request['qm_certification_date']
                flightModel.qm_certification_status = request['qm_certification_status']
                flightModel.attachment = request['attachment']
                flightModel.remarks = request['remarks']
                flightModel.isActive = is_active
                flightModel.testing_date = request['testing_date']
                flightModel.sys_type = request['sys_type']
                flightModel.cgbalancing_date = request['cgbalancing_date']
                flightModel.cgbalancing_date_status = request['cgbalancing_date_status']
                flightModel.launchact_date = request['launchact_date']
                flightModel.launchact_status = request['launchact_status']
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
                        'qm_certification_status'] != get_prod.qm_certification_status:
                        print("add data in flight history")
                        FlightHistoryModal = FlightSystemStatusHistory()
                        FlightHistoryModal.f_id = get_prod.id
                        FlightHistoryModal.system = get_prod.system
                        FlightHistoryModal.organization = get_prod.organization
                        FlightHistoryModal.set_id = get_prod.set_id
                        FlightHistoryModal.blt_date = get_prod.blt_date
                        FlightHistoryModal.blt_status = get_prod.blt_status
                        FlightHistoryModal.pre_hil_date = get_prod.pre_hil_date
                        FlightHistoryModal.pre_hil_status = get_prod.pre_hil_status
                        FlightHistoryModal.vibaration_date = get_prod.vibaration_date
                        FlightHistoryModal.vibaration_status = get_prod.vibaration_status
                        FlightHistoryModal.post_hil_date = get_prod.post_hil_date
                        FlightHistoryModal.post_hil_status = get_prod.post_hil_status
                        FlightHistoryModal.fgt_date = get_prod.fgt_date
                        FlightHistoryModal.fgt_status = get_prod.fgt_status
                        FlightHistoryModal.final_integration_date = get_prod.final_integration_date
                        FlightHistoryModal.final_integration_status = get_prod.final_integration_status
                        FlightHistoryModal.bhd_date = get_prod.bhd_date
                        FlightHistoryModal.bhd_status = get_prod.bhd_status
                        FlightHistoryModal.fqm_date = get_prod.fqm_date
                        FlightHistoryModal.fqm_status = get_prod.fqm_status
                        FlightHistoryModal.qm_certification_date = get_prod.qm_certification_date
                        FlightHistoryModal.qm_certification_status = get_prod.qm_certification_status
                        FlightHistoryModal.attachment = get_prod.attachment
                        FlightHistoryModal.remarks = get_prod.remarks
                        FlightHistoryModal.isActive = get_prod.isActive
                        FlightHistoryModal.testing_date = get_prod.testing_date
                        FlightHistoryModal.sys_type = get_prod.sys_type
                        FlightHistoryModal.cgbalancing_date = get_prod.cgbalancing_date
                        FlightHistoryModal.cgbalancing_date_status = get_prod.cgbalancing_date_status
                        FlightHistoryModal.launchact_date = get_prod.launchact_date
                        FlightHistoryModal.launchact_status = get_prod.launchact_status
                        FlightHistoryModal.save()

                        # history saved
                    get_prod.system = request['system']
                    get_prod.organization = request['organization']
                    get_prod.set_id = request['set_id']
                    get_prod.blt_date = request['blt_date']
                    get_prod.blt_status = request['blt_status']
                    get_prod.pre_hil_date = request['pre_hil_date']
                    get_prod.pre_hil_status = request['pre_hil_status']
                    get_prod.vibaration_date = request['vibaration_date']
                    get_prod.vibaration_status = request['vibaration_status']
                    get_prod.post_hil_date = request['post_hil_date']
                    get_prod.post_hil_status = request['post_hil_status']
                    get_prod.fgt_date = request['fgt_date']
                    get_prod.fgt_status = request['fgt_status']
                    get_prod.final_integration_date = request['final_integration_date']
                    get_prod.final_integration_status = request['final_integration_status']
                    get_prod.bhd_date = request['bhd_date']
                    get_prod.bhd_status = request['bhd_status']
                    get_prod.fqm_date = request['fqm_date']
                    get_prod.fqm_status = request['fqm_status']
                    get_prod.qm_certification_date = request['qm_certification_date']
                    get_prod.qm_certification_status = request['qm_certification_status']
                    get_prod.testing_date = request['testing_date']
                    get_prod.sys_type = request['sys_type']
                    if request['attachment'] != '':
                        get_prod.attachment = request['attachment']
                    get_prod.remarks = request['remarks']
                    get_prod.cgbalancing_date = request['cgbalancing_date']
                    get_prod.cgbalancing_date_status = request['cgbalancing_date_status']
                    get_prod.launchact_date = request['launchact_date']
                    get_prod.launchact_status = request['launchact_status']
                    get_prod.isActive = is_active
                    get_prod.save()
            return JsonResponse({'status': 'True', 'message': "Production Status Updated Successfully!"},
                                status=200)
        except Exception as e:
            return JsonResponse({'status': 'False', "message": "Status Not Saved"}, status=500)

    @staticmethod
    def GetFlightList(request):
        try:
            status = request.query_params['status']
            if status == 'Total':
                data = FlightSystemStatus.objects.all()
                serializer = FlightSystemSerialzer(data, many=True)
            else:
                data = FlightSystemStatus.objects.filter(qm_certification_status=status)
                serializer = FlightSystemSerialzer(data, many=True)
            return JsonResponse({'message': 'Welcome to Home Page', 'data': serializer.data}, status=200)
        except:
            return JsonResponse({'message': 'Sorry! No Task found.'}, status=500)

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
                refilModel.blt_status = request['blt_status']
                refilModel.pre_hil_date = request['pre_hil_date']
                refilModel.pre_hil_status = request['pre_hil_status']
                refilModel.vibaration_date = request['vibaration_date']
                refilModel.vibaration_status = request['vibaration_status']
                refilModel.post_hil_date = request['post_hil_date']
                refilModel.post_hil_status = request['post_hil_status']
                refilModel.fgt_date = request['fgt_date']
                refilModel.fgt_status = request['fgt_status']
                refilModel.final_integration_date = request['final_integration_date']
                refilModel.final_integration_status = request['final_integration_status']
                refilModel.bhd_date = request['bhd_date']
                refilModel.bhd_status = request['bhd_status']
                refilModel.fqm_date = request['fqm_date']
                refilModel.fqm_status = request['fqm_status']
                refilModel.qm_certification_date = request['qm_certification_date']
                refilModel.qm_certification_status = request['qm_certification_status']
                refilModel.attachment = request['attachment']
                refilModel.remarks = request['remarks']
                refilModel.isActive = is_active
                refilModel.testing_date = request['testing_date']
                refilModel.sys_type = request['sys_type']
                refilModel.cgbalancing_date = request['cgbalancing_date']
                refilModel.cgbalancing_date_status = request['cgbalancing_date_status']
                refilModel.enduser_date = request['enduser_date']
                refilModel.enduser_status = request['enduser_status']
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
                        'qm_certification_status'] != get_prod.qm_certification_status:
                        print("add data in flight history")
                        RelifingHistoryModal = RelifingSystemStatusHistory()
                        RelifingHistoryModal.r_id = get_prod.id
                        RelifingHistoryModal.system = get_prod.system
                        RelifingHistoryModal.organization = get_prod.organization
                        RelifingHistoryModal.set_id = get_prod.set_id
                        RelifingHistoryModal.blt_date = get_prod.blt_date
                        RelifingHistoryModal.blt_status = get_prod.blt_status
                        RelifingHistoryModal.pre_hil_date = get_prod.pre_hil_date
                        RelifingHistoryModal.pre_hil_status = get_prod.pre_hil_status
                        RelifingHistoryModal.vibaration_date = get_prod.vibaration_date
                        RelifingHistoryModal.vibaration_status = get_prod.vibaration_status
                        RelifingHistoryModal.post_hil_date = get_prod.post_hil_date
                        RelifingHistoryModal.post_hil_status = get_prod.post_hil_status
                        RelifingHistoryModal.fgt_date = get_prod.fgt_date
                        RelifingHistoryModal.fgt_status = get_prod.fgt_status
                        RelifingHistoryModal.final_integration_date = get_prod.final_integration_date
                        RelifingHistoryModal.final_integration_status = get_prod.final_integration_status
                        RelifingHistoryModal.bhd_date = get_prod.bhd_date
                        RelifingHistoryModal.bhd_status = get_prod.bhd_status
                        RelifingHistoryModal.fqm_date = get_prod.fqm_date
                        RelifingHistoryModal.fqm_status = get_prod.fqm_status
                        RelifingHistoryModal.qm_certification_date = get_prod.qm_certification_date
                        RelifingHistoryModal.qm_certification_status = get_prod.qm_certification_status
                        RelifingHistoryModal.attachment = get_prod.attachment
                        RelifingHistoryModal.remarks = get_prod.remarks
                        RelifingHistoryModal.isActive = get_prod.isActive
                        RelifingHistoryModal.cgbalancing_date = get_prod.cgbalancing_date
                        RelifingHistoryModal.cgbalancing_date_status = get_prod.cgbalancing_date_status
                        RelifingHistoryModal.enduser_date = get_prod.enduser_date
                        RelifingHistoryModal.enduser_status = get_prod.enduser_status
                        RelifingHistoryModal.save()

                        # history saved
                    get_prod.system = request['system']
                    get_prod.organization = request['organization']
                    get_prod.set_id = request['set_id']
                    get_prod.blt_date = request['blt_date']
                    get_prod.blt_status = request['blt_status']
                    get_prod.pre_hil_date = request['pre_hil_date']
                    get_prod.pre_hil_status = request['pre_hil_status']
                    get_prod.vibaration_date = request['vibaration_date']
                    get_prod.vibaration_status = request['vibaration_status']
                    get_prod.post_hil_date = request['post_hil_date']
                    get_prod.post_hil_status = request['post_hil_status']
                    get_prod.fgt_date = request['fgt_date']
                    get_prod.fgt_status = request['fgt_status']
                    get_prod.final_integration_date = request['final_integration_date']
                    get_prod.final_integration_status = request['final_integration_status']
                    get_prod.bhd_date = request['bhd_date']
                    get_prod.bhd_status = request['bhd_status']
                    get_prod.fqm_date = request['fqm_date']
                    get_prod.fqm_status = request['fqm_status']
                    get_prod.qm_certification_date = request['qm_certification_date']
                    get_prod.qm_certification_status = request['qm_certification_status']
                    if request['attachment'] != '':
                        get_prod.attachment = request['attachment']
                    get_prod.remarks = request['remarks']
                    get_prod.cgbalancing_date = request['cgbalancing_date']
                    get_prod.cgbalancing_date_status = request['cgbalancing_date_status']
                    get_prod.enduser_date = request['enduser_date']
                    get_prod.enduser_status = request['enduser_status']
                    get_prod.isActive = is_active
                    get_prod.save()
            return JsonResponse({'status': 'True', 'message': "Production Status Updated Successfully!"},
                                status=200)
        except Exception as e:
            return JsonResponse({'status': 'False', "message": "Status Not Saved"}, status=500)

    @staticmethod
    def GetRelifingList(request):
        try:
            status = request.query_params['status']
            if status == 'Total':
                data = RelifingSystemStatus.objects.all()
                serializer = RelifingSystemSerialzer(data, many=True)
                return JsonResponse({'message': 'Welcome to Home Page', 'data': serializer.data}, status=200)
            else:
                data = RelifingSystemStatus.objects.filter(qm_certification_status=status)
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
            prod_blt_oklist = ProductionSystemStatus.objects.filter(filter_objects, blt_status='OK')
            prod_blt_observationlist = ProductionSystemStatus.objects.filter(filter_objects, blt_status='Observation')
            prod_blt_uplist = ProductionSystemStatus.objects.filter(filter_objects, blt_status='Under process')
            prod_blt_haultlist = ProductionSystemStatus.objects.filter(filter_objects, blt_status='Halt')

            prod_prehil_oklist = ProductionSystemStatus.objects.filter(filter_objects, pre_hil_status='OK')
            prod_prehil_observationlist = ProductionSystemStatus.objects.filter(filter_objects,
                                                                                pre_hil_status='Observation')
            prod_prehil_uplist = ProductionSystemStatus.objects.filter(filter_objects, pre_hil_status='Under process')
            prod_prehil_haultlist = ProductionSystemStatus.objects.filter(filter_objects, pre_hil_status='Halt')

            prod_posthil_oklist = ProductionSystemStatus.objects.filter(filter_objects, post_hil_status='OK')
            prod_posthil_observationlist = ProductionSystemStatus.objects.filter(filter_objects,
                                                                                 post_hil_status='Observation')
            prod_posthil_uplist = ProductionSystemStatus.objects.filter(filter_objects, post_hil_status='Under process')
            prod_posthil_haultlist = ProductionSystemStatus.objects.filter(filter_objects, post_hil_status='Halt')

            prod_finalintegration_oklist = ProductionSystemStatus.objects.filter(filter_objects,
                                                                                 final_integration_status='OK')
            prod_finalintegration_observationlist = ProductionSystemStatus.objects.filter(filter_objects,
                                                                                          final_integration_status='Observation')
            prod_finalintegration_completelist = ProductionSystemStatus.objects.filter(filter_objects,
                                                                                       final_integration_status='Completed')
            prod_finalintegration_haultlist = ProductionSystemStatus.objects.filter(filter_objects,
                                                                                    final_integration_status='Halt')

            prod_vibration_oklist = ProductionSystemStatus.objects.filter(filter_objects, vibaration_status='OK')
            prod_vibration_observationlist = ProductionSystemStatus.objects.filter(filter_objects,
                                                                                   vibaration_status='Observation')
            prod_vibration_uplist = ProductionSystemStatus.objects.filter(filter_objects,
                                                                          vibaration_status='Under process')
            prod_vibration_haultlist = ProductionSystemStatus.objects.filter(filter_objects, vibaration_status='Halt')

            # flight system count
            flight_blt_oklist = FlightSystemStatus.objects.filter(filter_objects, blt_status='OK')
            flight_blt_observationlist = FlightSystemStatus.objects.filter(filter_objects, blt_status='Observation')
            flight_blt_uplist = FlightSystemStatus.objects.filter(filter_objects, blt_status='Under process')
            flight_blt_haultlist = FlightSystemStatus.objects.filter(filter_objects, blt_status='Halt')

            flight_prehil_oklist = FlightSystemStatus.objects.filter(filter_objects, pre_hil_status='OK')
            flight_prehil_observationlist = FlightSystemStatus.objects.filter(filter_objects,
                                                                              pre_hil_status='Observation')
            flight_prehil_uplist = FlightSystemStatus.objects.filter(filter_objects, pre_hil_status='Under process')
            flight_prehil_haultlist = FlightSystemStatus.objects.filter(filter_objects, pre_hil_status='Halt')

            flight_posthil_oklist = FlightSystemStatus.objects.filter(filter_objects, post_hil_status='OK')
            flight_posthil_observationlist = FlightSystemStatus.objects.filter(filter_objects,
                                                                               post_hil_status='Observation')
            flight_posthil_uplist = FlightSystemStatus.objects.filter(filter_objects, post_hil_status='Under process')
            flight_posthil_haultlist = FlightSystemStatus.objects.filter(filter_objects, post_hil_status='Halt')

            flight_finalintegration_oklist = FlightSystemStatus.objects.filter(filter_objects,
                                                                               final_integration_status='OK')
            flight_finalintegration_observationlist = FlightSystemStatus.objects.filter(filter_objects,
                                                                                        final_integration_status='Observation')
            flight_finalintegration_completelist = FlightSystemStatus.objects.filter(filter_objects,
                                                                                     final_integration_status='Completed')
            flight_finalintegration_haultlist = FlightSystemStatus.objects.filter(filter_objects,
                                                                                  final_integration_status='Halt')

            flight_vibration_oklist = FlightSystemStatus.objects.filter(filter_objects, vibaration_status='OK')
            flight_vibration_observationlist = FlightSystemStatus.objects.filter(filter_objects,
                                                                                 vibaration_status='Observation')
            flight_vibration_uplist = FlightSystemStatus.objects.filter(filter_objects,
                                                                        vibaration_status='Under process')
            flight_vibration_haultlist = FlightSystemStatus.objects.filter(filter_objects, vibaration_status='Halt')

            # relifing system count
            relifing_blt_oklist = RelifingSystemStatus.objects.filter(filter_objects, blt_status='OK')
            relifing_blt_observationlist = RelifingSystemStatus.objects.filter(filter_objects, blt_status='Observation')
            relifing_blt_uplist = RelifingSystemStatus.objects.filter(filter_objects, blt_status='Under process')
            relifing_blt_haultlist = RelifingSystemStatus.objects.filter(filter_objects, blt_status='Halt')

            relifing_prehil_oklist = RelifingSystemStatus.objects.filter(filter_objects, pre_hil_status='OK')
            relifing_prehil_observationlist = RelifingSystemStatus.objects.filter(filter_objects,
                                                                                  pre_hil_status='Observation')
            relifing_prehil_uplist = RelifingSystemStatus.objects.filter(filter_objects, pre_hil_status='Under process')
            relifing_prehil_haultlist = RelifingSystemStatus.objects.filter(filter_objects, pre_hil_status='Halt')

            relifing_posthil_oklist = RelifingSystemStatus.objects.filter(filter_objects, post_hil_status='OK')
            relifing_posthil_observationlist = RelifingSystemStatus.objects.filter(filter_objects,
                                                                                   post_hil_status='Observation')
            relifing_posthil_uplist = RelifingSystemStatus.objects.filter(filter_objects,
                                                                          post_hil_status='Under process')
            relifing_posthil_haultlist = RelifingSystemStatus.objects.filter(filter_objects, post_hil_status='Halt')

            relifing_finalintegration_oklist = RelifingSystemStatus.objects.filter(filter_objects,
                                                                                   final_integration_status='OK')
            relifing_finalintegration_observationlist = RelifingSystemStatus.objects.filter(filter_objects,
                                                                                            final_integration_status='Observation')
            relifing_finalintegration_completelist = RelifingSystemStatus.objects.filter(filter_objects,
                                                                                         final_integration_status='Completed')
            relifing_finalintegration_haultlist = RelifingSystemStatus.objects.filter(filter_objects,
                                                                                      final_integration_status='Halt')

            relifing_vibration_oklist = RelifingSystemStatus.objects.filter(filter_objects, vibaration_status='OK')
            relifing_vibration_observationlist = RelifingSystemStatus.objects.filter(filter_objects,
                                                                                     vibaration_status='Observation')
            relifing_vibration_uplist = RelifingSystemStatus.objects.filter(filter_objects,
                                                                            vibaration_status='Under process')
            relifing_vibration_haultlist = RelifingSystemStatus.objects.filter(filter_objects, vibaration_status='Halt')

            # if year is not '':
            #     prod_blt_oklist = prod_blt_oklist.filter(testing_date__year=year)
            #     prod_blt_observationlist = prod_blt_observationlist.filter(testing_date__year=year)
            #     prod_blt_uplist = prod_blt_uplist.filter(testing_date__year=year)
            #     prod_blt_haultlist = prod_blt_haultlist.filter(testing_date__year=year)
            #
            #     prod_prehil_oklist = prod_prehil_oklist.filter(testing_date__year=year)
            #     prod_prehil_observationlist = prod_prehil_observationlist.filter(testing_date__year=year)
            #     prod_prehil_uplist = prod_prehil_uplist.filter(testing_date__year=year)
            #     prod_prehil_haultlist = prod_prehil_haultlist.filter(testing_date__year=year)
            #
            #     prod_posthil_oklist = prod_posthil_oklist.filter(testing_date__year=year)
            #     prod_posthil_observationlist = prod_posthil_observationlist.filter(testing_date__year=year)
            #     prod_posthil_uplist = prod_posthil_uplist.filter(testing_date__year=year)
            #     prod_posthil_haultlist = prod_posthil_haultlist.filter(testing_date__year=year)
            #
            #     prod_finalintegration_oklist = prod_finalintegration_oklist.filter(testing_date__year=year)
            #     prod_finalintegration_observationlist = prod_finalintegration_observationlist.filter(
            #         testing_date__year=year)
            #     prod_finalintegration_completelist = prod_finalintegration_completelist.filter(testing_date__year=year)
            #     prod_finalintegration_haultlist = prod_finalintegration_haultlist.filter(testing_date__year=year)
            #
            #     # flight system
            #     flight_blt_oklist = flight_blt_oklist.filter(testing_date__year=year)
            #     flight_blt_observationlist = flight_blt_observationlist.filter(testing_date__year=year)
            #     flight_blt_uplist = flight_blt_uplist.filter(testing_date__year=year)
            #     flight_blt_haultlist = flight_blt_haultlist.filter(testing_date__year=year)
            #
            #     flight_prehil_oklist = flight_prehil_oklist.filter(testing_date__year=year)
            #     flight_prehil_observationlist = flight_prehil_observationlist.filter(
            #         testing_date__year=year)
            #     flight_prehil_uplist = flight_prehil_uplist.filter(testing_date__year=year)
            #     flight_prehil_haultlist = flight_prehil_haultlist.filter(testing_date__year=year)
            #
            #     flight_posthil_oklist = flight_posthil_oklist.filter(testing_date__year=year)
            #     flight_posthil_observationlist = flight_posthil_observationlist.filter(testing_date__year=year)
            #     flight_posthil_uplist = flight_posthil_uplist.filter(testing_date__year=year)
            #     flight_posthil_haultlist = flight_posthil_haultlist.filter(testing_date__year=year)
            #
            #     flight_finalintegration_oklist = flight_finalintegration_oklist.filter(testing_date__year=year)
            #     flight_finalintegration_observationlist = flight_finalintegration_observationlist.filter(
            #         testing_date__year=year)
            #     flight_finalintegration_completelist = flight_finalintegration_completelist.filter(
            #         testing_date__year=year)
            #     flight_finalintegration_haultlist = flight_finalintegration_haultlist.filter(
            #         testing_date__year=year)
            #
            #     # relifing system
            #     relifing_blt_oklist = relifing_blt_oklist.filter(testing_date__year=year)
            #     relifing_blt_observationlist = relifing_blt_observationlist.filter(testing_date__year=year)
            #     relifing_blt_uplist = relifing_blt_uplist.filter(testing_date__year=year)
            #     relifing_blt_haultlist = relifing_blt_haultlist.filter(testing_date__year=year)
            #
            #     relifing_prehil_oklist = relifing_prehil_oklist.filter(testing_date__year=year)
            #     relifing_prehil_observationlist = relifing_prehil_observationlist.filter(
            #         testing_date__year=year)
            #     relifing_prehil_uplist = relifing_prehil_uplist.filter(testing_date__year=year)
            #     relifing_prehil_haultlist = relifing_prehil_haultlist.filter(testing_date__year=year)
            #
            #     relifing_posthil_oklist = relifing_posthil_oklist.filter(testing_date__year=year)
            #     relifing_posthil_observationlist = relifing_posthil_observationlist.filter(
            #         testing_date__year=year)
            #     relifing_posthil_uplist = relifing_posthil_uplist.filter(testing_date__year=year)
            #     relifing_posthil_haultlist = relifing_posthil_haultlist.filter(testing_date__year=year)
            #
            #     relifing_finalintegration_oklist = relifing_finalintegration_oklist.filter(testing_date__year=year)
            #     relifing_finalintegration_observationlist = relifing_finalintegration_observationlist.filter(
            #         testing_date__year=year)
            #     relifing_finalintegration_completelist = relifing_finalintegration_completelist.filter(
            #         testing_date__year=year)
            #     relifing_finalintegration_haultlist = relifing_finalintegration_haultlist.filter(
            #         testing_date__year=year)
            #
            # if org is not '':
            #     prod_blt_oklist = prod_blt_oklist.filter(organization=org)
            #     prod_blt_observationlist = prod_blt_observationlist.filter(organization=org)
            #     prod_blt_uplist = prod_blt_uplist.filter(organization=org)
            #     prod_blt_haultlist = prod_blt_haultlist.filter(organization=org)
            #
            #     prod_prehil_oklist = prod_prehil_oklist.filter(organization=org)
            #     prod_prehil_observationlist = prod_prehil_observationlist.filter(organization=org)
            #     prod_prehil_uplist = prod_prehil_uplist.filter(organization=org)
            #     prod_prehil_haultlist = prod_prehil_haultlist.filter(organization=org)
            #
            #     prod_posthil_oklist = prod_posthil_oklist.filter(organization=org)
            #     prod_posthil_observationlist = prod_posthil_observationlist.filter(organization=org)
            #     prod_posthil_uplist = prod_posthil_uplist.filter(organization=org)
            #     prod_posthil_haultlist = prod_posthil_haultlist.filter(organization=org)
            #
            #     prod_finalintegration_oklist = prod_finalintegration_oklist.filter(organization=org)
            #     prod_finalintegration_observationlist = prod_finalintegration_observationlist.filter(
            #         organization=org)
            #     prod_finalintegration_completelist = prod_finalintegration_completelist.filter(
            #         organization=org)
            #     prod_finalintegration_haultlist = prod_finalintegration_haultlist.filter(
            #         organization=org)
            #
            #     # flight system
            #     flight_blt_oklist = flight_blt_oklist.filter(organization=org)
            #     flight_blt_observationlist = flight_blt_observationlist.filter(organization=org)
            #     flight_blt_uplist = flight_blt_uplist.filter(organization=org)
            #     flight_blt_haultlist = flight_blt_haultlist.filter(organization=org)
            #
            #     flight_prehil_oklist = flight_prehil_oklist.filter(organization=org)
            #     flight_prehil_observationlist = flight_prehil_observationlist.filter(
            #         organization=org)
            #     flight_prehil_uplist = flight_prehil_uplist.filter(organization=org)
            #     flight_prehil_haultlist = flight_prehil_haultlist.filter(organization=org)
            #
            #     flight_posthil_oklist = flight_posthil_oklist.filter(organization=org)
            #     flight_posthil_observationlist = flight_posthil_observationlist.filter(organization=org)
            #     flight_posthil_uplist = flight_posthil_uplist.filter(organization=org)
            #     flight_posthil_haultlist = flight_posthil_haultlist.filter(organization=org)
            #
            #     flight_finalintegration_oklist = flight_finalintegration_oklist.filter(organization=org)
            #     flight_finalintegration_observationlist = flight_finalintegration_observationlist.filter(
            #         organization=org)
            #     flight_finalintegration_completelist = flight_finalintegration_completelist.filter(
            #         organization=org)
            #     flight_finalintegration_haultlist = flight_finalintegration_haultlist.filter(
            #         organization=org)
            #
            #     # relifing system
            #     relifing_blt_oklist = relifing_blt_oklist.filter(organization=org)
            #     relifing_blt_observationlist = relifing_blt_observationlist.filter(organization=org)
            #     relifing_blt_uplist = relifing_blt_uplist.filter(organization=org)
            #     relifing_blt_haultlist = relifing_blt_haultlist.filter(organization=org)
            #
            #     relifing_prehil_oklist = relifing_prehil_oklist.filter(organization=org)
            #     relifing_prehil_observationlist = relifing_prehil_observationlist.filter(
            #         organization=org)
            #     relifing_prehil_uplist = relifing_prehil_uplist.filter(organization=org)
            #     relifing_prehil_haultlist = relifing_prehil_haultlist.filter(organization=org)
            #
            #     relifing_posthil_oklist = relifing_posthil_oklist.filter(organization=org)
            #     relifing_posthil_observationlist = relifing_posthil_observationlist.filter(
            #         organization=org)
            #     relifing_posthil_uplist = relifing_posthil_uplist.filter(organization=org)
            #     relifing_posthil_haultlist = relifing_posthil_haultlist.filter(organization=org)
            #
            #     relifing_finalintegration_oklist = relifing_finalintegration_oklist.filter(
            #         organization=org)
            #     relifing_finalintegration_observationlist = relifing_finalintegration_observationlist.filter(
            #         organization=org)
            #     relifing_finalintegration_completelist = relifing_finalintegration_completelist.filter(
            #         organization=org)
            #     relifing_finalintegration_haultlist = relifing_finalintegration_haultlist.filter(
            #         organization=org)
            #
            # if type is not '':
            #     prod_blt_oklist = prod_blt_oklist.filter(sys_type=type)
            #     prod_blt_observationlist = prod_blt_observationlist.filter(sys_type=type)
            #     prod_blt_uplist = prod_blt_uplist.filter(sys_type=type)
            #     prod_blt_haultlist = prod_blt_haultlist.filter(sys_type=type)
            #
            #     prod_prehil_oklist = prod_prehil_oklist.filter(sys_type=type)
            #     prod_prehil_observationlist = prod_prehil_observationlist.filter(sys_type=type)
            #     prod_prehil_uplist = prod_prehil_uplist.filter(sys_type=type)
            #     prod_prehil_haultlist = prod_prehil_haultlist.filter(sys_type=type)
            #
            #     prod_posthil_oklist = prod_posthil_oklist.filter(sys_type=type)
            #     prod_posthil_observationlist = prod_posthil_observationlist.filter(sys_type=type)
            #     prod_posthil_uplist = prod_posthil_uplist.filter(sys_type=type)
            #     prod_posthil_haultlist = prod_posthil_haultlist.filter(sys_type=type)
            #
            #     prod_finalintegration_oklist = prod_finalintegration_oklist.filter(sys_type=type)
            #     prod_finalintegration_observationlist = prod_finalintegration_observationlist.filter(
            #         sys_type=type)
            #     prod_finalintegration_completelist = prod_finalintegration_completelist.filter(
            #         sys_type=type)
            #     prod_finalintegration_haultlist = prod_finalintegration_haultlist.filter(
            #         sys_type=type)
            #
            #     # flight system
            #     flight_blt_oklist = flight_blt_oklist.filter(sys_type=type)
            #     flight_blt_observationlist = flight_blt_observationlist.filter(sys_type=type)
            #     flight_blt_uplist = flight_blt_uplist.filter(sys_type=type)
            #     flight_blt_haultlist = flight_blt_haultlist.filter(sys_type=type)
            #
            #     flight_prehil_oklist = flight_prehil_oklist.filter(sys_type=type)
            #     flight_prehil_observationlist = flight_prehil_observationlist.filter(
            #         sys_type=type)
            #     flight_prehil_uplist = flight_prehil_uplist.filter(sys_type=type)
            #     flight_prehil_haultlist = flight_prehil_haultlist.filter(sys_type=type)
            #
            #     flight_posthil_oklist = flight_posthil_oklist.filter(sys_type=type)
            #     flight_posthil_observationlist = flight_posthil_observationlist.filter(sys_type=type)
            #     flight_posthil_uplist = flight_posthil_uplist.filter(sys_type=type)
            #     flight_posthil_haultlist = flight_posthil_haultlist.filter(sys_type=type)
            #
            #     flight_finalintegration_oklist = flight_finalintegration_oklist.filter(sys_type=type)
            #     flight_finalintegration_observationlist = flight_finalintegration_observationlist.filter(
            #         sys_type=type)
            #     flight_finalintegration_completelist = flight_finalintegration_completelist.filter(
            #         sys_type=type)
            #     flight_finalintegration_haultlist = flight_finalintegration_haultlist.filter(
            #         sys_type=type)
            #
            #     # relifing system
            #     relifing_blt_oklist = relifing_blt_oklist.filter(sys_type=type)
            #     relifing_blt_observationlist = relifing_blt_observationlist.filter(sys_type=type)
            #     relifing_blt_uplist = relifing_blt_uplist.filter(sys_type=type)
            #     relifing_blt_haultlist = relifing_blt_haultlist.filter(sys_type=type)
            #
            #     relifing_prehil_oklist = relifing_prehil_oklist.filter(sys_type=type)
            #     relifing_prehil_observationlist = relifing_prehil_observationlist.filter(
            #         sys_type=type)
            #     relifing_prehil_uplist = relifing_prehil_uplist.filter(sys_type=type)
            #     relifing_prehil_haultlist = relifing_prehil_haultlist.filter(sys_type=type)
            #
            #     relifing_posthil_oklist = relifing_posthil_oklist.filter(sys_type=type)
            #     relifing_posthil_observationlist = relifing_posthil_observationlist.filter(
            #         sys_type=type)
            #     relifing_posthil_uplist = relifing_posthil_uplist.filter(sys_type=type)
            #     relifing_posthil_haultlist = relifing_posthil_haultlist.filter(sys_type=type)
            #
            #     relifing_finalintegration_oklist = relifing_finalintegration_oklist.filter(
            #         sys_type=type)
            #     relifing_finalintegration_observationlist = relifing_finalintegration_observationlist.filter(
            #         sys_type=type)
            #     relifing_finalintegration_completelist = relifing_finalintegration_completelist.filter(
            #         sys_type=type)
            #     relifing_finalintegration_haultlist = relifing_finalintegration_haultlist.filter(
            #         sys_type=type)
            #
            # if system is not '':
            #     prod_blt_oklist = prod_blt_oklist.filter(system=system)
            #     prod_blt_observationlist = prod_blt_observationlist.filter(system=system)
            #     prod_blt_uplist = prod_blt_uplist.filter(system=system)
            #     prod_blt_haultlist = prod_blt_haultlist.filter(system=system)
            #
            #     prod_prehil_oklist = prod_prehil_oklist.filter(system=system)
            #     prod_prehil_observationlist = prod_prehil_observationlist.filter(system=system)
            #     prod_prehil_uplist = prod_prehil_uplist.filter(system=system)
            #     prod_prehil_haultlist = prod_prehil_haultlist.filter(system=system)
            #
            #     prod_posthil_oklist = prod_posthil_oklist.filter(system=system)
            #     prod_posthil_observationlist = prod_posthil_observationlist.filter(system=system)
            #     prod_posthil_uplist = prod_posthil_uplist.filter(system=system)
            #     prod_posthil_haultlist = prod_posthil_haultlist.filter(system=system)
            #
            #     prod_finalintegration_oklist = prod_finalintegration_oklist.filter(system=system)
            #     prod_finalintegration_observationlist = prod_finalintegration_observationlist.filter(
            #         system=system)
            #     prod_finalintegration_completelist = prod_finalintegration_completelist.filter(
            #         system=system)
            #     prod_finalintegration_haultlist = prod_finalintegration_haultlist.filter(
            #         system=system)
            #
            #     # flight system
            #     flight_blt_oklist = flight_blt_oklist.filter(system=system)
            #     flight_blt_observationlist = flight_blt_observationlist.filter(system=system)
            #     flight_blt_uplist = flight_blt_uplist.filter(system=system)
            #     flight_blt_haultlist = flight_blt_haultlist.filter(system=system)
            #
            #     flight_prehil_oklist = flight_prehil_oklist.filter(system=system)
            #     flight_prehil_observationlist = flight_prehil_observationlist.filter(
            #         system=system)
            #     flight_prehil_uplist = flight_prehil_uplist.filter(system=system)
            #     flight_prehil_haultlist = flight_prehil_haultlist.filter(system=system)
            #
            #     flight_posthil_oklist = flight_posthil_oklist.filter(system=system)
            #     flight_posthil_observationlist = flight_posthil_observationlist.filter(system=system)
            #     flight_posthil_uplist = flight_posthil_uplist.filter(system=system)
            #     flight_posthil_haultlist = flight_posthil_haultlist.filter(system=system)
            #
            #     flight_finalintegration_oklist = flight_finalintegration_oklist.filter(system=system)
            #     flight_finalintegration_observationlist = flight_finalintegration_observationlist.filter(
            #         system=system)
            #     flight_finalintegration_completelist = flight_finalintegration_completelist.filter(
            #         system=system)
            #     flight_finalintegration_haultlist = flight_finalintegration_haultlist.filter(
            #         system=system)
            #
            #     # relifing system
            #     relifing_blt_oklist = relifing_blt_oklist.filter(system=system)
            #     relifing_blt_observationlist = relifing_blt_observationlist.filter(system=system)
            #     relifing_blt_uplist = relifing_blt_uplist.filter(system=system)
            #     relifing_blt_haultlist = relifing_blt_haultlist.filter(system=system)
            #
            #     relifing_prehil_oklist = relifing_prehil_oklist.filter(system=system)
            #     relifing_prehil_observationlist = relifing_prehil_observationlist.filter(
            #         system=system)
            #     relifing_prehil_uplist = relifing_prehil_uplist.filter(system=system)
            #     relifing_prehil_haultlist = relifing_prehil_haultlist.filter(system=system)
            #
            #     relifing_posthil_oklist = relifing_posthil_oklist.filter(system=system)
            #     relifing_posthil_observationlist = relifing_posthil_observationlist.filter(
            #         system=system)
            #     relifing_posthil_uplist = relifing_posthil_uplist.filter(system=system)
            #     relifing_posthil_haultlist = relifing_posthil_haultlist.filter(system=system)
            #
            #     relifing_finalintegration_oklist = relifing_finalintegration_oklist.filter(
            #         system=system)
            #     relifing_finalintegration_observationlist = relifing_finalintegration_observationlist.filter(
            #         system=system)
            #     relifing_finalintegration_completelist = relifing_finalintegration_completelist.filter(
            #         system=system)
            #     relifing_finalintegration_haultlist = relifing_finalintegration_haultlist.filter(
            #         system=system)

            dist = {
                'prod_blt_ok': prod_blt_oklist.count(),
                'prod_blt_observation': prod_blt_observationlist.count(),
                'prod_blt_up': prod_blt_uplist.count(),
                'prod_blt_hault': prod_blt_haultlist.count(),
                'flight_blt_oklist': flight_blt_oklist.count(),
                'flight_blt_observationlist': flight_blt_observationlist.count(),
                'flight_blt_uplist': flight_blt_uplist.count(),
                'flight_blt_haultlist': flight_blt_haultlist.count(),
                'relifing_blt_oklist': relifing_blt_oklist.count(),
                'relifing_blt_observationlist': relifing_blt_observationlist.count(),
                'relifing_blt_uplist': relifing_blt_uplist.count(),
                'relifing_blt_haultlist': relifing_blt_haultlist.count(),
                'prod_prehil_oklist': prod_prehil_oklist.count(),
                'prod_prehil_observationlist': prod_prehil_observationlist.count(),
                'prod_prehil_uplist': prod_prehil_uplist.count(),
                'prod_prehil_haultlist': prod_prehil_haultlist.count(),
                'flight_prehil_oklist': flight_prehil_oklist.count(),
                'flight_prehil_observationlist': flight_prehil_observationlist.count(),
                'flight_prehil_uplist': flight_prehil_uplist.count(),
                'flight_prehil_haultlist': flight_prehil_haultlist.count(),
                'relifing_prehil_oklist': relifing_prehil_oklist.count(),
                'relifing_prehil_observationlist': relifing_prehil_observationlist.count(),
                'relifing_prehil_uplist': relifing_prehil_uplist.count(),
                'relifing_prehil_haultlist': relifing_prehil_haultlist.count(),
                'prod_posthil_oklist': prod_posthil_oklist.count(),
                'prod_posthil_observationlist': prod_posthil_observationlist.count(),
                'prod_posthil_uplist': prod_posthil_uplist.count(),
                'prod_posthil_haultlist': prod_posthil_haultlist.count(),
                'flight_posthil_oklist': flight_posthil_oklist.count(),
                'flight_posthil_observationlist': flight_posthil_observationlist.count(),
                'flight_posthil_uplist': flight_posthil_uplist.count(),
                'flight_posthil_haultlist': flight_posthil_haultlist.count(),
                'relifing_posthil_oklist': relifing_posthil_oklist.count(),
                'relifing_posthil_observationlist': relifing_posthil_observationlist.count(),
                'relifing_posthil_uplist': relifing_posthil_uplist.count(),
                'relifing_posthil_haultlist': relifing_posthil_haultlist.count(),
                'prod_finalintegration_oklist': prod_finalintegration_oklist.count(),
                'prod_finalintegration_observationlist': prod_finalintegration_observationlist.count(),
                'prod_finalintegration_completelist': prod_finalintegration_completelist.count(),
                'prod_finalintegration_haultlist': prod_finalintegration_haultlist.count(),
                'flight_finalintegration_oklist': flight_finalintegration_oklist.count(),
                'flight_finalintegration_observationlist': flight_finalintegration_observationlist.count(),
                'flight_finalintegration_completelist': flight_finalintegration_completelist.count(),
                'flight_finalintegration_haultlist': flight_finalintegration_haultlist.count(),
                'relifing_finalintegration_oklist': relifing_finalintegration_oklist.count(),
                'relifing_finalintegration_observationlist': relifing_finalintegration_observationlist.count(),
                'relifing_finalintegration_completelist': relifing_finalintegration_completelist.count(),
                'relifing_finalintegration_haultlist': relifing_finalintegration_haultlist.count(),
                'relifing_vibration_oklist': relifing_vibration_oklist.count(),
                'relifing_vibration_observationlist': relifing_vibration_observationlist.count(),
                'relifing_vibration_uplist': relifing_vibration_uplist.count(),
                'relifing_vibration_haultlist': relifing_vibration_haultlist.count(),
                'flight_vibration_oklist': flight_vibration_oklist.count(),
                'flight_vibration_observationlist': flight_vibration_observationlist.count(),
                'flight_vibration_uplist': flight_vibration_uplist.count(),
                'flight_vibration_haultlist': flight_vibration_haultlist.count(),
                'prod_vibration_oklist': prod_vibration_oklist.count(),
                'prod_vibration_observationlist': prod_vibration_observationlist.count(),
                'prod_vibration_uplist': prod_vibration_uplist.count(),
                'prod_vibration_haultlist': prod_vibration_haultlist.count()
            }
            return JsonResponse({'message': 'true', 'data': dist}, status=200)
        except Exception as e:
            print(e)
            return JsonResponse({'message': 'Sorry! No Task found.'}, status=500)
