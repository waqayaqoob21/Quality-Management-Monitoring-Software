from importlib import import_module
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
import fitz #pip install PyMuPDF Pillow
import io
import os
from PIL import Image
from pytesseract import pytesseract

from sms.models import FlightSystemStatus, ProductionSystemStatus, RelifingSystemStatus
from sms.serializers import FlightSystemSerialzer, ProductionSystemSerialzer, RelifingSystemSerialzer # install tesseract-ocr-w64-setup-v5.2.0.20220712.exe (64 bit) resp. 
                                    # from https://github.com/UB-Mannheim/tesseract/wiki 
                                    # pip install pytesseract

class SmsController:

    @staticmethod
    def AddProductionStatus(request):
        prodModel = ProductionSystemStatus()

        try:
            id = request['id']
            if id == '0':
                prodModel.system = request['system']
                prodModel.organization = request['organization']
                prodModel.set_id = request['set_id']
                prodModel.blt_status = request['blt_status']
                prodModel.pre_hil_status = request['pre_hil_status']
                prodModel.vibaration_status = request['vibaration_status']
                prodModel.post_hil_status = request['post_hil_status']
                prodModel.fgt_status = request['fgt_status']
                prodModel.final_integration_st = request['final_integration_st']
                prodModel.bhd_status = request['bhd_status']
                prodModel.fqm_status = request['fqm_status']
                prodModel.qm_certification_st = request['qm_certification_st']
                prodModel.attachment = request['attachment']
                prodModel.remarks = request['remarks']
                prodModel.save()
                return JsonResponse({'status': 'True', 'message': "Production Status Added Successfully!"},
                                status=200)
            else:
                get_prod = prodModel.objects.filter(id=id).first()
                if get_prod is not None:
                    prodModel.system = request['system']
                    prodModel.organization = request['organization']
                    prodModel.set_id = request['set_id']
                    prodModel.blt_status = request['blt_status']
                    prodModel.pre_hil_status = request['pre_hil_status']
                    prodModel.vibaration_status = request['vibaration_status']
                    prodModel.post_hil_status = request['post_hil_status']
                    prodModel.fgt_status = request['fgt_status']
                    prodModel.final_integration_st = request['final_integration_st']
                    prodModel.bhd_status = request['bhd_status']
                    prodModel.fqm_status = request['fqm_status']
                    prodModel.qm_certification_st = request['qm_certification_st']
                    prodModel.attachment = request['attachment']
                    prodModel.remarks = request['remarks']
                    prodModel.save()
            return JsonResponse({'status': 'True', 'message': "Production Status Updated Successfully!"},
                                status=200)
        except Exception as e:
            return JsonResponse({'status': 'False', "message": "Status Not Saved"}, status=500)


    @staticmethod
    def GetProductionList(request):
        try:   
            data = ProductionSystemStatus.objects.all()
            serializer = ProductionSystemSerialzer(data, many=True)
            return JsonResponse({'message':'Welcome to Home Page','Task List': serializer.data}, status=200)
        except:
            return JsonResponse({'message':'Sorry! No Task found.'}, status=200)

    # API for PDF generator with table and Text_wraping

    @staticmethod
    def GetProductionPDFList(request):
        print("fuck")
        TABLE_COL_NAMES = ("System", "Organization", "Set ID", "BLT Status","Pre-Hil Status", "Vibration Status", "Post-Hil Status","FGT Status", "Final Integration Status","BHD Status", "FQM Status","QM Certification Status","Attachment","Remarks")
        data = ProductionSystemStatus.objects.all().order_by('-id').values_list('system','organization','set_id','blt_status','pre_hil_status','vibaration_status','post_hil_status','fgt_status','final_integration_st','bhd_status','fqm_status','qm_certification_st','attachment','remarks')

        pdf = FPDF('L', 'mm', 'Legal')
        pdf.add_page()
        pdf.set_font('courier', 'B', 26)
        pdf.cell(330, 10, 'Final Report', border=0,align='C', ln=2)
        pdf.cell(40, 10, '',0,1)
        pdf.set_font("Times", size=10)
        line_height = pdf.font_size * 3.2
        col_width = pdf.epw / 14
    
        def render_table_header():
            pdf.set_font(style="B") 
            for col_name in TABLE_COL_NAMES:
                pdf.multi_cell(col_width, line_height, col_name, border=1,align='C', ln=3, max_line_height=pdf.font_size)
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
                if number_of_words>2:
                    use_default_height = 1
                    new_line_height = pdf.font_size * (number_of_words/2)
            if not use_default_height:
                lh_list.append(line_height)
            else:
                lh_list.append(new_line_height)
                use_default_height = 0

        for j,row in enumerate(data):
            line_height = lh_list[j] 
            if pdf.will_page_break(line_height):
                render_table_header()
            for col_num in range(len(row)):
                line_height = lh_list[j] 
                pdf.multi_cell(col_width, line_height, f"{row[col_num]}", border=1,align='C',ln=3, 
                max_line_height=pdf.font_size)
            pdf.ln(line_height)
        pdf.output('report.pdf')
        return FileResponse(open('report.pdf', 'rb'), as_attachment=True, content_type='application/pdf')

    @staticmethod
    def GetProductionExcelList(request):
        response = HttpResponse(content_type = 'application/ms-excel')
        response['Content-Disposition'] = 'attachment; filename="DocumentExcelSheet.xls"'
        work_book = xlwt.Workbook(encoding='utf-8')
        work_sheet = work_book.add_sheet('ExcelSheet')
        row_num = 0
        font_style = xlwt.XFStyle()
        font_style.font.bold = True
        TABLE_COL_NAMES = ("System", "Organization", "Set ID", "BLT Status","Pre-Hil Status", "Vibration Status", "Post-Hil Status","FGT Status", "Final Integration Status","BHD Status", "FQM Status","QM Certification Status","Attachment","Remarks")

        for col_num in range(len(TABLE_COL_NAMES)):
            work_sheet.write(row_num,col_num,TABLE_COL_NAMES[col_num], font_style)
        font_style = xlwt.XFStyle()
        data = ProductionSystemStatus.objects.all().order_by('-id').values_list('system','organization','set_id','blt_status','pre_hil_status','vibaration_status','post_hil_status','fgt_status','final_integration_st','bhd_status','fqm_status','qm_certification_st','attachment','remarks')

        for row in data:
            row_num += 1
            for col_num in range(len(row)):
                work_sheet.write(row_num,col_num,str(row[col_num]), font_style)
        work_book.save(response)
        return response
    


    @staticmethod
    def AddFlightStatus(request):
        flightModel = FlightSystemStatus()

        try:
            id = request['id']
            if id == '0':
                flightModel.system = request['system']
                flightModel.organization = request['organization']
                flightModel.set_id = request['set_id']
                flightModel.blt_status = request['blt_status']
                flightModel.pre_hil_status = request['pre_hil_status']
                flightModel.vibaration_status = request['vibaration_status']
                flightModel.post_hil_status = request['post_hil_status']
                flightModel.fgt_status = request['fgt_status']
                flightModel.final_integration_st = request['final_integration_st']
                flightModel.bhd_status = request['bhd_status']
                flightModel.fqm_status = request['fqm_status']
                flightModel.qm_certification_st = request['qm_certification_st']
                flightModel.attachment = request['attachment']
                flightModel.remarks = request['remarks']
                flightModel.save()
                return JsonResponse({'status': 'True', 'message': "Production Status Added Successfully!"},
                                status=200)
            else:
                get_prod = flightModel.objects.filter(id=id).first()
                if get_prod is not None:
                    flightModel.system = request['system']
                    flightModel.organization = request['organization']
                    flightModel.set_id = request['set_id']
                    flightModel.blt_status = request['blt_status']
                    flightModel.pre_hil_status = request['pre_hil_status']
                    flightModel.vibaration_status = request['vibaration_status']
                    flightModel.post_hil_status = request['post_hil_status']
                    flightModel.fgt_status = request['fgt_status']
                    flightModel.final_integration_st = request['final_integration_st']
                    flightModel.bhd_status = request['bhd_status']
                    flightModel.fqm_status = request['fqm_status']
                    flightModel.qm_certification_st = request['qm_certification_st']
                    flightModel.attachment = request['attachment']
                    flightModel.remarks = request['remarks']
                    flightModel.save()
            return JsonResponse({'status': 'True', 'message': "Production Status Updated Successfully!"},
                                status=200)
        except Exception as e:
            return JsonResponse({'status': 'False', "message": "Status Not Saved"}, status=500)


    @staticmethod
    def GetFlightList(request):
        try:   
            data = FlightSystemStatus.objects.all()
            serializer = FlightSystemSerialzer(data, many=True)
            return JsonResponse({'message':'Welcome to Home Page','Task List': serializer.data}, status=200)
        except:
            return JsonResponse({'message':'Sorry! No Task found.'}, status=200)

    # API for PDF generator with table and Text_wraping

    @staticmethod
    def GetFlightPDFList(request):
        TABLE_COL_NAMES = ("System", "Organization", "Set ID", "BLT Status","Pre-Hil Status", "Vibration Status", "Post-Hil Status","FGT Status", "Final Integration Status","BHD Status", "FQM Status","QM Certification Status","Attachment","Remarks")
        data = FlightSystemStatus.objects.all().order_by('-id').values_list('system','organization','set_id','blt_status','pre_hil_status','vibaration_status','post_hil_status','fgt_status','final_integration_st','bhd_status','fqm_status','qm_certification_st','attachment','remarks')

        pdf = FPDF('L', 'mm', 'Legal')
        pdf.add_page()
        pdf.set_font('courier', 'B', 26)
        pdf.cell(330, 10, 'Final Report', border=0,align='C', ln=2)
        pdf.cell(40, 10, '',0,1)
        pdf.set_font("Times", size=10)
        line_height = pdf.font_size * 3.2
        col_width = pdf.epw / 14
    
        def render_table_header():
            pdf.set_font(style="B") 
            for col_name in TABLE_COL_NAMES:
                pdf.multi_cell(col_width, line_height, col_name, border=1,align='C', ln=3, max_line_height=pdf.font_size)
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
                if number_of_words>2:
                    use_default_height = 1
                    new_line_height = pdf.font_size * (number_of_words/2)
            if not use_default_height:
                lh_list.append(line_height)
            else:
                lh_list.append(new_line_height)
                use_default_height = 0

        for j,row in enumerate(data):
            line_height = lh_list[j] 
            if pdf.will_page_break(line_height):
                render_table_header()
            for col_num in range(len(row)):
                line_height = lh_list[j] 
                pdf.multi_cell(col_width, line_height, f"{row[col_num]}", border=1,align='C',ln=3, 
                max_line_height=pdf.font_size)
            pdf.ln(line_height)
        pdf.output('report.pdf')
        return FileResponse(open('report.pdf', 'rb'), as_attachment=True, content_type='application/pdf')

    @staticmethod
    def GetFlightExcelList(request):
        response = HttpResponse(content_type = 'application/ms-excel')
        response['Content-Disposition'] = 'attachment; filename="DocumentExcelSheet.xls"'
        work_book = xlwt.Workbook(encoding='utf-8')
        work_sheet = work_book.add_sheet('ExcelSheet')
        row_num = 0
        font_style = xlwt.XFStyle()
        font_style.font.bold = True
        TABLE_COL_NAMES = ("System", "Organization", "Set ID", "BLT Status","Pre-Hil Status", "Vibration Status", "Post-Hil Status","FGT Status", "Final Integration Status","BHD Status", "FQM Status","QM Certification Status","Attachment","Remarks")

        for col_num in range(len(TABLE_COL_NAMES)):
            work_sheet.write(row_num,col_num,TABLE_COL_NAMES[col_num], font_style)
        font_style = xlwt.XFStyle()
        data = FlightSystemStatus.objects.all().order_by('-id').values_list('system','organization','set_id','blt_status','pre_hil_status','vibaration_status','post_hil_status','fgt_status','final_integration_st','bhd_status','fqm_status','qm_certification_st','attachment','remarks')

        for row in data:
            row_num += 1
            for col_num in range(len(row)):
                work_sheet.write(row_num,col_num,str(row[col_num]), font_style)
        work_book.save(response)
        return response



    
    @staticmethod
    def AddRelifingStatus(request):
        refilModel = RelifingSystemStatus()

        try:
            id = request['id']
            if id == '0':
                refilModel.system = request['system']
                refilModel.organization = request['organization']
                refilModel.set_id = request['set_id']
                refilModel.blt_status = request['blt_status']
                refilModel.pre_hil_status = request['pre_hil_status']
                refilModel.vibaration_status = request['vibaration_status']
                refilModel.post_hil_status = request['post_hil_status']
                refilModel.fgt_status = request['fgt_status']
                refilModel.final_integration_st = request['final_integration_st']
                refilModel.bhd_status = request['bhd_status']
                refilModel.fqm_status = request['fqm_status']
                refilModel.qm_certification_st = request['qm_certification_st']
                refilModel.attachment = request['attachment']
                refilModel.remarks = request['remarks']
                refilModel.save()
                return JsonResponse({'status': 'True', 'message': "Production Status Added Successfully!"},
                                status=200)
            else:
                get_prod = refilModel.objects.filter(id=id).first()
                if get_prod is not None:
                    refilModel.system = request['system']
                    refilModel.organization = request['organization']
                    refilModel.set_id = request['set_id']
                    refilModel.blt_status = request['blt_status']
                    refilModel.pre_hil_status = request['pre_hil_status']
                    refilModel.vibaration_status = request['vibaration_status']
                    refilModel.post_hil_status = request['post_hil_status']
                    refilModel.fgt_status = request['fgt_status']
                    refilModel.final_integration_st = request['final_integration_st']
                    refilModel.bhd_status = request['bhd_status']
                    refilModel.fqm_status = request['fqm_status']
                    refilModel.qm_certification_st = request['qm_certification_st']
                    refilModel.attachment = request['attachment']
                    refilModel.remarks = request['remarks']
                    refilModel.save()
            return JsonResponse({'status': 'True', 'message': "Production Status Updated Successfully!"},
                                status=200)
        except Exception as e:
            return JsonResponse({'status': 'False', "message": "Status Not Saved"}, status=500)


    @staticmethod
    def GetRelifingList(request):
        try:   
            data = RelifingSystemStatus.objects.all()
            serializer = RelifingSystemSerialzer(data, many=True)
            return JsonResponse({'message':'Welcome to Home Page','Task List': serializer.data}, status=200)
        except:
            return JsonResponse({'message':'Sorry! No Task found.'}, status=200)

    # API for PDF generator with table and Text_wraping

    @staticmethod
    def GetRelifingPDFList(request):
        TABLE_COL_NAMES = ("System", "Organization", "Set ID", "BLT Status","Pre-Hil Status", "Vibration Status", "Post-Hil Status","FGT Status", "Final Integration Status","BHD Status", "FQM Status","QM Certification Status","Attachment","Remarks")
        data = RelifingSystemStatus.objects.all().order_by('-id').values_list('system','organization','set_id','blt_status','pre_hil_status','vibaration_status','post_hil_status','fgt_status','final_integration_st','bhd_status','fqm_status','qm_certification_st','attachment','remarks')

        pdf = FPDF('L', 'mm', 'Legal')
        pdf.add_page()
        pdf.set_font('courier', 'B', 26)
        pdf.cell(330, 10, 'Final Report', border=0,align='C', ln=2)
        pdf.cell(40, 10, '',0,1)
        pdf.set_font("Times", size=10)
        line_height = pdf.font_size * 3.2
        col_width = pdf.epw / 14
    
        def render_table_header():
            pdf.set_font(style="B") 
            for col_name in TABLE_COL_NAMES:
                pdf.multi_cell(col_width, line_height, col_name, border=1,align='C', ln=3, max_line_height=pdf.font_size)
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
                if number_of_words>2:
                    use_default_height = 1
                    new_line_height = pdf.font_size * (number_of_words/2)
            if not use_default_height:
                lh_list.append(line_height)
            else:
                lh_list.append(new_line_height)
                use_default_height = 0

        for j,row in enumerate(data):
            line_height = lh_list[j] 
            if pdf.will_page_break(line_height):
                render_table_header()
            for col_num in range(len(row)):
                line_height = lh_list[j] 
                pdf.multi_cell(col_width, line_height, f"{row[col_num]}", border=1,align='C',ln=3, 
                max_line_height=pdf.font_size)
            pdf.ln(line_height)
        pdf.output('report.pdf')
        return FileResponse(open('report.pdf', 'rb'), as_attachment=True, content_type='application/pdf')

    @staticmethod
    def GetRelifingExcelList(request):
        response = HttpResponse(content_type = 'application/ms-excel')
        response['Content-Disposition'] = 'attachment; filename="DocumentExcelSheet.xls"'
        work_book = xlwt.Workbook(encoding='utf-8')
        work_sheet = work_book.add_sheet('ExcelSheet')
        row_num = 0
        font_style = xlwt.XFStyle()
        font_style.font.bold = True
        TABLE_COL_NAMES = ("System", "Organization", "Set ID", "BLT Status","Pre-Hil Status", "Vibration Status", "Post-Hil Status","FGT Status", "Final Integration Status","BHD Status", "FQM Status","QM Certification Status","Attachment","Remarks")

        for col_num in range(len(TABLE_COL_NAMES)):
            work_sheet.write(row_num,col_num,TABLE_COL_NAMES[col_num], font_style)
        font_style = xlwt.XFStyle()
        data = RelifingSystemStatus.objects.all().order_by('-id').values_list('system','organization','set_id','blt_status','pre_hil_status','vibaration_status','post_hil_status','fgt_status','final_integration_st','bhd_status','fqm_status','qm_certification_st','attachment','remarks')

        for row in data:
            row_num += 1
            for col_num in range(len(row)):
                work_sheet.write(row_num,col_num,str(row[col_num]), font_style)
        work_book.save(response)
        return response