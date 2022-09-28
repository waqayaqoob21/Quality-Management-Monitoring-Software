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

from mpm.models import ActiveMotors
from mpm.serializer import ActiveMotorSerializer
from qms.models import CespAudit, QmsAudit
from qms.serializer import CespAuditSerializer, QmsAuditSerializer


class QmsController:
    @staticmethod
    def AddQmsAudit(request):
        qmsModel = QmsAudit()
        try:
            id = request['id']
            if id == '0':
                qmsModel.audit_id = request['audit_id']
                qmsModel.Organization = request['Organization']
                qmsModel.site = request['site']
                qmsModel.certification_status = request['certification_status']
                qmsModel.previous_standard = request['previous_standard']
                qmsModel.certification_validity = request['certification_validity']
                qmsModel.audit_type = request['audit_type']
                qmsModel.planned_date = request['planned_date']
                qmsModel.audit_start_date = request['audit_start_date']
                qmsModel.audit_close_date = request['audit_close_date']
                qmsModel.audit_status = request['audit_status']
                qmsModel.standard = request['standard']
                qmsModel.remarks = request['remarks']
                qmsModel.save()
                return JsonResponse({'status': 'True', 'message': "QMS Audit Created Successfully!"},
                                status=200)
            else:
                get_qms = QmsAudit.objects.filter(id=id).first()
                if get_qms is not None:
                    get_qms.audit_id = request['audit_id']
                    get_qms.Organization = request['Organization']
                    get_qms.site = request['site']
                    get_qms.certification_status = request['certification_status']
                    get_qms.previous_standard = request['previous_standard']
                    get_qms.certification_validity = request['certification_validity']
                    get_qms.audit_type = request['audit_type']
                    get_qms.planned_date = request['planned_date']
                    get_qms.audit_start_date = request['audit_start_date']
                    get_qms.audit_close_date = request['audit_close_date']
                    get_qms.audit_status = request['audit_status']
                    get_qms.standard = request['standard']
                    get_qms.remarks = request['remarks']
                    get_qms.save()
                    return JsonResponse({'status': 'True', 'message': "QMS Audit Updated Successfully!"},
                                status=200)
        except Exception as e:
            return JsonResponse({'status': 'False', "message": "QMS Audit Not Saved"}, status=500)
            
    @staticmethod
    def GetQmsAuditList(request):
        try:
            data = QmsAudit.objects.all().order_by('id')
            serializer = QmsAuditSerializer(data, many=True)
            return JsonResponse({'data':serializer.data}, safe=False, status=200)
        except:
            return JsonResponse({'message': 'Sorry! No QMS Audit found.'}, status=500)

    @staticmethod
    def DeleteQmsAudit(request, pk):
        try:
            qms = QmsAudit.objects.get(id=pk)
            qms.delete()
            return JsonResponse({'message': 'QMS Audit has been deleted'}, status=200)
        except:
            return JsonResponse({'message': 'Sorry! No QMS Audit found.'}, status=500)

    @staticmethod
    def GetQmsAuditPDFList(request):
        TABLE_COL_NAMES = (
            "Audit ID", "Organization", "Site", "Certification Status", "Previous Standard", "Certification Validity",
            "Audit Type", "Planned Date","Audit Start Date", "Audit Close Date", "Audit Status", "Standard", "Remarks")
        data = QmsAudit.objects.all().values_list('audit_id', 'Organization', 'site', 'certification_status',
                                                  'previous_standard', 'certification_validity', 'audit_type',
                                                  'planned_date',
                                                  'audit_start_date', 'audit_close_date', 'audit_status', 'standard',
                                                  'remarks')

        pdf = FPDF('L', 'mm', 'Legal')
        pdf.add_page()
        pdf.set_font('courier', 'B', 26)
        pdf.cell(330, 10, 'Final Report', border=0, align='C', ln=2)
        pdf.cell(40, 10, '', 0, 1)
        pdf.set_font("Times", size=10)
        line_height = pdf.font_size * 2.5
        col_width = pdf.epw / 13

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
                    new_line_height = pdf.font_size*2 * (number_of_words / 2)
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
        pdf.output('report.pdf')
        return FileResponse(open('report.pdf', 'rb'), as_attachment=True, content_type='application/pdf')

    @staticmethod
    def GetQmsAuditExcelList(request):
        response = HttpResponse(content_type='application/ms-excel')
        response['Content-Disposition'] = 'attachment; filename="DocumentExcelSheet.xls"'
        work_book = xlwt.Workbook(encoding='utf-8')
        work_sheet = work_book.add_sheet('ExcelSheet')
        row_num = 0
        font_style = xlwt.XFStyle()
        font_style.font.bold = True
        TABLE_COL_NAMES = (
        "Audit ID", "Organization", "Site", "Certification Status", "Previous Standard", "Certification Validity", "Audit Type","Planned Date",
        "Audit Start Date","Audit Close Date","Audit Status","Standard","Remarks")
        for col_num in range(len(TABLE_COL_NAMES)):
            work_sheet.write(row_num, col_num, TABLE_COL_NAMES[col_num], font_style)
        font_style = xlwt.XFStyle()

        data = QmsAudit.objects.all().values_list('audit_id','Organization','site','certification_status',
                                                                                  'previous_standard','certification_validity','audit_type','planned_date',
                                                                                  'audit_start_date','audit_close_date','audit_status','standard','remarks')
        for row in data:
            row_num += 1
            for col_num in range(len(row)):
                work_sheet.write(row_num, col_num, str(row[col_num]), font_style)
        work_book.save(response)
        return response

#=========================CeSP Audit==================================== 
    @staticmethod
    def AddCespAudit(request):
        cespModel = CespAudit()
        try:
            id = request['id']
            if id == '0':
                cespModel.audit_id = request['audit_id']
                cespModel.Organization = request['Organization']
                cespModel.site = request['site']
                cespModel.certification_status = request['certification_status']
                cespModel.previous_standard = request['previous_standard']
                cespModel.certification_validity = request['certification_validity']
                cespModel.audit_type = request['audit_type']
                cespModel.planned_date = request['planned_date']
                cespModel.audit_start_date = request['audit_start_date']
                cespModel.audit_close_date = request['audit_close_date']
                cespModel.audit_status = request['audit_status']
                cespModel.standard = request['standard']
                cespModel.remarks = request['remarks']
                cespModel.save()
                return JsonResponse({'status': 'True', 'message': "CeSP Audit Created Successfully!"},
                                status=200)
            else:
                get_cesp = CespAudit.objects.filter(id=id).first()
                if get_cesp is not None:
                    get_cesp.audit_id = request['audit_id']
                    get_cesp.Organization = request['Organization']
                    get_cesp.site = request['site']
                    get_cesp.certification_status = request['certification_status']
                    get_cesp.previous_standard = request['previous_standard']
                    get_cesp.certification_validity = request['certification_validity']
                    get_cesp.audit_type = request['audit_type']
                    get_cesp.planned_date = request['planned_date']
                    get_cesp.audit_start_date = request['audit_start_date']
                    get_cesp.audit_close_date = request['audit_close_date']
                    get_cesp.audit_status = request['audit_status']
                    get_cesp.standard = request['standard']
                    get_cesp.remarks = request['remarks']
                    get_cesp.save()
                    return JsonResponse({'status': 'True', 'message': "CeSP Audit Updated Successfully!"},
                                status=200)
        except Exception as e:
            return JsonResponse({'status': 'False', "message": "CeSP Audit Not Saved"}, status=500)
            
    @staticmethod
    def GetCespAuditList(request):
        try:
            data = CespAudit.objects.all().order_by('id')
            serializer = CespAuditSerializer(data, many=True)
            return JsonResponse({'data':serializer.data}, safe=False, status=200)
        except:
            return JsonResponse({'message': 'Sorry! No Audit found.'}, status=500)

    @staticmethod
    def DeleteCespAudit(request, pk):
        try:
            cesp = CespAudit.objects.get(id=pk)
            cesp.delete()
            return JsonResponse({'message': 'CeSP Audit has been deleted'}, status=200)
        except:
            return JsonResponse({'message': 'Sorry! No Audit found.'}, status=500)

    @staticmethod
    def GetCespAuditPDFList(request):
        TABLE_COL_NAMES = (
            "Audit ID", "Organization", "Site", "Certification Status", "Previous Standard", "Certification Validity",
            "Audit Type", "Planned Date","Audit Start Date", "Audit Close Date", "Audit Status", "Standard", "Remarks")
        data = CespAudit.objects.all().values_list('audit_id', 'Organization', 'site', 'certification_status',
                                                  'previous_standard', 'certification_validity', 'audit_type',
                                                  'planned_date','audit_start_date', 'audit_close_date', 'audit_status',
                                                  'standard','remarks')

        pdf = FPDF('L', 'mm', 'Legal')
        pdf.add_page()
        pdf.set_font('courier', 'B', 26)
        pdf.cell(330, 10, 'Final Report', border=0, align='C', ln=2)
        pdf.cell(40, 10, '', 0, 1)
        pdf.set_font("Times", size=10)
        line_height = pdf.font_size * 2.5
        col_width = pdf.epw / 13

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
                    new_line_height = pdf.font_size* 2 * (number_of_words / 2)
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
        pdf.output('report.pdf')
        return FileResponse(open('report.pdf', 'rb'), as_attachment=True, content_type='application/pdf')

    @staticmethod
    def GetCespAuditExcelList(request):
        response = HttpResponse(content_type='application/ms-excel')
        response['Content-Disposition'] = 'attachment; filename="DocumentExcelSheet.xls"'
        work_book = xlwt.Workbook(encoding='utf-8')
        work_sheet = work_book.add_sheet('ExcelSheet')
        row_num = 0
        font_style = xlwt.XFStyle()
        font_style.font.bold = True
        TABLE_COL_NAMES = (
            "Audit ID", "Organization", "Site", "Certification Status", "Previous Standard", "Certification Validity",
            "Audit Type", "Planned Date", "Audit Start Date", "Audit Close Date", "Audit Status", "Standard", "Remarks")
        for col_num in range(len(TABLE_COL_NAMES)):
            work_sheet.write(row_num, col_num, TABLE_COL_NAMES[col_num], font_style)
        font_style = xlwt.XFStyle()


        data = CespAudit.objects.all().values_list('audit_id', 'Organization', 'site', 'certification_status',
                                                  'previous_standard', 'certification_validity', 'audit_type',
                                                  'planned_date', 'audit_start_date', 'audit_close_date',
                                                  'audit_status',
                                                  'standard', 'remarks')
        for row in data:
            row_num += 1
            for col_num in range(len(row)):
                work_sheet.write(row_num, col_num, str(row[col_num]), font_style)
        work_book.save(response)
        return response