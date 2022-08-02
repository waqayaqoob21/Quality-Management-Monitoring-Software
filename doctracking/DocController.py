import base64
from datetime import datetime, timedelta

from dateutil.relativedelta import relativedelta
from django.db import connection
from django.http import JsonResponse
from passlib.utils.compat import izip
import textwrap
from django.http import FileResponse

from doctracking.serializer import DocListSerializer
from usermanagement.models import doctracking


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


    @staticmethod
    def GetDocumentPDFList(request):
        TABLE_COL_NAMES = ("Document Name", "Document Type", "Sender", "Receive Date", "Marked To", "Marked Date", "Due Date", "Task Date", "Status", "Sent To", "Sent Date","Remarks")
        data = doctracking.objects.filter(due_date__lt=datetime.today()).values('doc_name','doc_type','sender','receive_date','marked_to','marked_date','due_date','task_date','status','sent_to','sent_date','remarks')

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
                word_list = datum.split()
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
            for key, datum in row.items():
                line_height = lh_list[j] 
                pdf.multi_cell(col_width, line_height, f"{datum}", border=1,align='C',ln=3, 
                max_line_height=pdf.font_size)
            pdf.ln(line_height)
        pdf.output('report.pdf')
        return FileResponse(open('report.pdf', 'rb'), as_attachment=True, content_type='application/pdf')
