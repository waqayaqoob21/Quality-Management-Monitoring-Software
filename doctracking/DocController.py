import base64
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta
from django.db import connection
from django.http import JsonResponse
from passlib.utils.compat import izip
from django.http import FileResponse
from django.http import HttpResponse
from fpdf import FPDF
import xlwt
from doctracking.serializer import DocListSerializer
from usermanagement.models import doctracking

from django.core.mail import EmailMessage
from django.conf import settings
class DocController:

    @staticmethod
    def AddDocument(request):

        try:
            id = request['id']
            if id == '0':
                docModal = doctracking()
                docModal.doc_name = request['doc_name']
                docModal.doc_type = request['doc_type']
                docModal.sender = request['sender']
                docModal.receive_date = request['receive_date']
                docModal.marked_to = request['marked_to']
                docModal.marked_date = request['marked_date']
                docModal.due_date = request['due_date']
                docModal.task_date = request['task_date']
                docModal.status = request['status']
                docModal.sent_to = request['sent_to']
                docModal.sent_date = request['sent_date']
                if request['isActive'] == 'true':
                    docModal.isActive = 1
                else:
                    docModal.isActive = 0
                docModal.remarks = request['remarks']
                docModal.attachement = request['attachement']
                docModal.save()
            else:
                get_doc = doctracking.objects.filter(id=id).first()
                if get_doc is not None:
                    get_doc.doc_name = request['doc_name']
                    get_doc.doc_type = request['doc_type']
                    get_doc.sender = request['sender']
                    get_doc.receive_date = request['receive_date']
                    get_doc.marked_to = request['marked_to']
                    get_doc.marked_date = request['marked_date']
                    get_doc.due_date = request['due_date']
                    get_doc.task_date = request['task_date']
                    get_doc.status = request['status']
                    get_doc.sent_to = request['sent_to']
                    get_doc.sent_date = request['sent_date']
                    if request['isActive'] == 'true':
                        get_doc.isActive = 1
                    else:
                        get_doc.isActive = 0
                    get_doc.remarks = request['remarks']
                    if request['attachement'] != '':
                        get_doc.attachement = request['attachement']
                    get_doc.save()
                    print("here")
            # print("Device Event Detail saved........", serializer.data)
            return JsonResponse({'status': 'True', 'message': "Record Created Successfully"},
                                status=200)

        except Exception as e:
            print(e)
            return JsonResponse({'status': 'False', "message": "Doc Not Saved"}, status=500)
            pass

    @staticmethod
    def GetDocumentList(request):

        try:

            docList = doctracking.objects.all().order_by('-id')
            serializer = DocListSerializer(docList, many=True)
            print(serializer)
            return JsonResponse({'status': 'True', 'data': serializer.data},
                                status=200)

        except Exception as e:
            print(e)
            return JsonResponse({'status': 'False', "message": "Internal Server Error"}, status=500)
            pass

    @staticmethod
    def GetPendingDocumentList(request):

        try:

            today_date = datetime.today()
            current_date = datetime.strptime(str(today_date.date()), "%Y-%m-%d") + relativedelta(hours=today_date.hour,
                                                                                                 minutes=today_date.minute,
                                                                                                 seconds=today_date.second,
                                                                                                 microseconds=today_date.microsecond)
            doc_list = []
            cursor = connection.cursor()
            query = "select *   " \
                    " from usermanagement_doctracking doc " \
                    " WHERE doc.due_date < '{0}'  ;".format(
                current_date)
            cursor.execute(query)
            col_names = [col[0] for col in cursor.description]
            for row in cursor.fetchall():
                row_dict = dict(izip(col_names, row))
                doc_list.append(row_dict)
            # docList = doctracking.objects.all().order_by('-id')
            serializer = DocListSerializer(doc_list, many=True)
            print(serializer)
            return JsonResponse({'status': 'True', 'data': serializer.data},
                                status=200)

        except Exception as e:
            print(e)
            return JsonResponse({'status': 'False', "message": "Internal Server Error"}, status=500)
            pass


# Document List API by Waqar
    @staticmethod
    def GetDocumentList(request):
        try:   
            data = doctracking.objects.filter(due_date__lt=datetime.today()).values()
            serializer = DocListSerializer(data, many=True)
            return JsonResponse({'message':'Welcome to Home Page','Data List': serializer.data}, status=200)
        except:
            return JsonResponse({'message':'Sorry! No list found.'}, status=204)

    # API for PDF generator with table and Text_wraping

    @staticmethod
    def GetDocumentPDFList(request):
        TABLE_COL_NAMES = ("Document Name", "Document Type", "Sender", "Receive Date", "Marked To", "Marked Date", "Due Date", "Task Date", "Status", "Sent To", "Sent Date","Remarks")
        data = doctracking.objects.filter(due_date__lt=datetime.today()).order_by('-id').values_list('doc_name','doc_type','sender','receive_date','marked_to','marked_date','due_date','task_date','status','sent_to','sent_date','remarks')

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
        TABLE_COL_NAMES = ("Document Name", "Document Type", "Sender", "Receive Date", "Marked To", "Marked Date", "Due Date", "Task Date", "Status", "Sent To", "Sent Date","Remarks")
        for col_num in range(len(TABLE_COL_NAMES)):
            work_sheet.write(row_num,col_num,TABLE_COL_NAMES[col_num], font_style)
        font_style = xlwt.XFStyle()

        data = doctracking.objects.filter(due_date__lt=datetime.today()).values_list('doc_name','doc_type','sender','receive_date','marked_to','marked_date','due_date','task_date','status','sent_to','sent_date','remarks')
        for row in data:
            row_num += 1
            for col_num in range(len(row)):
                work_sheet.write(row_num,col_num,str(row[col_num]), font_style)
        work_book.save(response)
        return response

# API for PDF and EXCEL Sheet generator with table and Text_wraping
    @staticmethod
    def GetDocumentEmailList(request):
        TABLE_COL_NAMES = ("Document Name", "Document Type", "Sender", "Receive Date", "Marked To", "Marked Date", "Due Date", "Task Date", "Status", "Sent To", "Sent Date","Remarks")
        data = doctracking.objects.filter(due_date__lt=datetime.today()).order_by('-id').values_list('doc_name','doc_type','sender','receive_date','marked_to','marked_date','due_date','task_date','status','sent_to','sent_date','remarks')

        # -----------------Excel Sheet Code Starts----------------------
        wb = xlwt.Workbook(encoding='utf-8')
        ws = wb.add_sheet('ExcelSheet')
        row_num = 0
        font_style = xlwt.XFStyle()
        font_style.font.bold = True
        for col_num in range(len(TABLE_COL_NAMES)):
            ws.write(row_num,col_num,TABLE_COL_NAMES[col_num], font_style)
        font_style = xlwt.XFStyle()
        # -----------------Excel Sheet Code Ends--------------------

        # -----------------PDF File Code Starts----------------------
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
        # -----------------PDF File Code Ends----------------------
       
        # -----------------File Generating Code Starts----------------------
        for j,row in enumerate(data):
            row_num += 1
            line_height = lh_list[j] 
            if pdf.will_page_break(line_height):
                render_table_header()
            for col_num in range(len(row)):
                ws.write(row_num,col_num,f"{row[col_num]}", font_style)
                line_height = lh_list[j] 
                pdf.multi_cell(col_width, line_height, f"{row[col_num]}", border=1,align='C',ln=3, 
                max_line_height=pdf.font_size)
            pdf.ln(line_height)
        pdf.output('pdf_report.pdf', 'F')
        wb.save('excel_report.xls')
        pdf_file_report =  FileResponse(open('pdf_report.pdf', 'rb'), as_attachment=True, content_type='application/pdf')
        excel_file_report =  FileResponse(open('excel_report.xls', 'rb'), as_attachment=True, content_type='application/ms-excel')
        send_action_email(excel_file_report.getvalue(), pdf_file_report.getvalue())
        return JsonResponse({'Success':'Email has been Sent Successfully'}, status=200)

def send_action_email(excel_file_report,pdf_file_report):
        try:
            body = f'Hi Sir! Please find attached files below:'
            subject = 'Record Files'
            from_email= settings.EMAIL_HOST_USER
            to_email = ['waqaryaqoob6@gmail.com']
            email = EmailMessage(subject, body, from_email, to_email)
            email.content_subtype='html'
            email.attach('report.xls', excel_file_report, 'application/ms-excel')
            email.attach('report.pdf', pdf_file_report, 'applicaiton/pdf')
            email.send()
        except:
            return JsonResponse({'Error':'Email Could not Send'}, status=400)