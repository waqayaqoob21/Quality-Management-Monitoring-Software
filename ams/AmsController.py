from django.http import JsonResponse
from ams.serializer import *
from ams.models import *
from fpdf import FPDF
import xlwt
from datetime import datetime, timedelta
from django.http import FileResponse
from django.http import HttpResponse

class AmsController:

    @staticmethod
    def AddTask(request):
        taskModel = TaskSummary()
        try:
            taskModel.task_name = request['task_name']
            taskModel.assigned_by = request['assigned_by']
            taskModel.assigned_date = request['assigned_date']
            taskModel.assigned_to = request['assigned_to']
            taskModel.target_date = request['target_date']
            taskModel.status = request['status']
            taskModel.remarks = request['remarks']
            taskModel.save()
            return JsonResponse({'Success': "Task Created Successfully"},status=200)
        except Exception as e:
            print(e)
            return JsonResponse({"Error": "Task Not Saved"}, status=500)

    @staticmethod
    def EditTask(request): 
        try:
            task = TaskSummary.objects.get(id=request['id'])
            task.task_name = request['task_name']
            task.assigned_by = request['assigned_by']
            task.assigned_date = request['assigned_date']
            task.assigned_to = request['assigned_to']
            task.target_date = request['target_date']
            task.status = request['status']
            task.remarks = request['remarks']
            task.save()
            return JsonResponse({'Success':'Task Updated Successfully!'}, status=201)
        except:
            return JsonResponse({'Error':'Task Could Not Update Successfully!'}, status=400)

    @staticmethod
    def GetTaskList(request):
        try:   
            data = TaskSummary.objects.all()
            serializer = TaskSummarySerialzer(data, many=True)
            return JsonResponse({'message':'Welcome to Home Page','Task List': serializer.data}, status=200)
        except:
            return JsonResponse({'message':'Sorry! No Task found.'}, status=200)

       # API for PDF generator with table and Text_wraping

    @staticmethod
    def GetDocumentPDFList(request):
        TABLE_COL_NAMES = ("Task Name", "Assigned By", "Assigned Date", "Assigned To", "Target Date", "Status","Remarks")
        data = TaskSummary.objects.filter(due_date__lt=datetime.today()).order_by('-id').values_list('task_name','assigned_by','assigned_date','assigned_to','target_date','status','remarks')

        pdf = FPDF('L', 'mm', 'Legal')
        pdf.add_page()
        pdf.set_font('courier', 'B', 26)
        pdf.cell(330, 10, 'Final Report', border=0,align='C', ln=2)
        pdf.cell(40, 10, '',0,1)
        pdf.set_font("Times", size=10)
        line_height = pdf.font_size * 2.5
        col_width = pdf.epw / 12
    
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
    def GetDocumentExcelList(request):
        response = HttpResponse(content_type = 'application/ms-excel')
        response['Content-Disposition'] = 'attachment; filename="DocumentExcelSheet.xls"'
        work_book = xlwt.Workbook(encoding='utf-8')
        work_sheet = work_book.add_sheet('ExcelSheet')
        row_num = 0
        font_style = xlwt.XFStyle()
        font_style.font.bold = True
        TABLE_COL_NAMES = ("Task Name", "Assigned By", "Assigned Date","Assigned To", "Target Date","Status","Remarks")
        for col_num in range(len(TABLE_COL_NAMES)):
            work_sheet.write(row_num,col_num,TABLE_COL_NAMES[col_num], font_style)
        font_style = xlwt.XFStyle()

        data = TaskSummary.objects.filter(due_date__lt=datetime.today()).values_list('task_name','assigned_by','assigned_date','assigned_to','target_date','status','remarks')
        for row in data:
            row_num += 1
            for col_num in range(len(row)):
                work_sheet.write(row_num,col_num,str(row[col_num]), font_style)
        work_book.save(response)
        return response