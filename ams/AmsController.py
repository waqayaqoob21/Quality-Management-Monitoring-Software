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
from PIL import Image
from pytesseract import pytesseract # install tesseract-ocr-w64-setup-v5.2.0.20220712.exe (64 bit) resp. 
                                    # from https://github.com/UB-Mannheim/tesseract/wiki 
                                    # pip install pytesseract

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
        data = TaskSummary.objects.all().order_by('-id').values_list('task_name','assigned_by','assigned_date','assigned_to','target_date','status','remarks')

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
    
   @staticmethod
    def GetStudentList(request):
        try:
            data = StdModel.objects.all().order_by('id')
            serializer = StdSerializer(data, many=True)
            return JsonResponse(serializer.data, safe=False, status=201)
        except:
            return JsonResponse({'message': 'Sorry! No student found.'}, status=400)

    @staticmethod
    def DeleteStudent(request, pk):
        try:
            student = StdModel.objects.get(id=pk)
            student.delete()
            return JsonResponse({'message': 'Student has been deleted'}, status=201)
        except:
            return JsonResponse({'message': 'Sorry! No student found.'}, status=400)
            
    @staticmethod
    def OcrPDF(request):
        ocrModel = OcrDataModel()
        fs = FileSystemStorage()
        attachment = request['uploaded_file']
        target_path = 'C:/Users/viCky/DjnagoAngularAPIs/djangoangularapi/files/'
        fs.save(target_path + attachment.name, attachment)
        with fitz.open(target_path + attachment.name) as my_pdf_file:
            # loop through every page
            for page_number in range(1, len(my_pdf_file) + 1):
                # acess individual page
                page = my_pdf_file[page_number - 1]
                # accesses all images of the page
                images = page.get_images(full=False)
                # check if images are there
                if images:
                    print(f"There are {len(images)} image/s on page number {page_number}[+]")
                else:
                    print(f"There are No image/s on page number {page_number}[!]")
                # loop through all images present in the page
                for image_number, image in enumerate(page.get_images(), start=1):
                    # access image xerf
                    xref_value = image[0]
                    # extract image information
                    base_image = my_pdf_file.extract_image(xref_value)
                    # access the image itself
                    image_bytes = base_image["image"]
                    # get image extension
                    ext = base_image["ext"]
                    # load image
                    image = Image.open(io.BytesIO(image_bytes))
                    # save image locally
                    image.save(open(f"files/Page{page_number}Image{image_number}.{ext}", "wb"))
        os.remove("files/" + attachment.name)
        path_to_tesseract = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
        image_path = r"files/Page1Image1.png"
        img = Image.open(image_path)
        pytesseract.tesseract_cmd = path_to_tesseract
        text = pytesseract.image_to_string(img)
        data= text.split("\n")
        final_list = [y for x in data for y in x.split(':')]
        for i in range(len(final_list)):
            ocrModel.organization = "NATIONAL DEVELOPMENT COMPLEX (NDC)"
            ocrModel.senior_directorate = "Waqar Afzal"

            if 'Group & Site' in final_list[i]:
                data = []
                data = final_list[i+1].split(" D")
                print(data[0])
                ocrModel.site = data[0]

            if final_list[i]=="Test Report No":
                ocrModel.test_report_no = final_list[i+1]

            if final_list[i]=="Job Card No":
                ocrModel.job_card_no = final_list[i+1]


            if final_list[i] == "Test Date":
                date = final_list[i+1]
                match_date = re.search(r'\d{2}-\d{2}-\d{2}', date)
                test_date = datetime.strptime(match_date.group(), '%d-%m-%y').date()
                ocrModel.test_report_date = test_date

            if 'Product Name' in final_list[i]:
                    data =[]
                    data = final_list[i+1].split(" ")
                    ocrModel.product_name = data[1]

            if 'ID No' in final_list[i]:
                id_data = final_list[i+1]
                ocrModel.id_no = id_data

            if final_list[i]=="Lot No & Lot Size":
                ocrModel.lot_no_lot_size = final_list[i+1]

            if 'Test Name' in final_list[i]:
                data = []
                data = final_list[i].split("e ")
                ocrModel.test_name = data[1]

            if 'Test Type' in final_list[i]:
                data = []
                data = final_list[i].split("e ")
                ocrModel.test_type = data[1]

            if 'Qualification Standard' in final_list[i]:
                data = []
                data = final_list[i].split("|")
                ocrModel.qualification_standard = data[1]

            if 'Test Specifications' in final_list[i]:
                data = final_list[i].partition('Specifications')[2]
                ocrModel.test_specifications = data

            if 'Results' in final_list[i]:
                data = final_list[i].partition(' e ')[2]
                ocrModel.results = data

            if 'Remarks' in final_list[i]:
                if final_list[i+1]=="":
                    ocrModel.remarks = "No remarks are given."
                else:
                    ocrModel.remarks = final_list[i+1]
        ocrModel.save()

        return JsonResponse({'Success': 'OCR has been completed Successfully!'}, status=200)

    @staticmethod
    def getOcrData(request):
        try:
            data = OcrDataModel.objects.all().order_by('id')
            serializer = OcrDataSerializer(data, many=True)
            return JsonResponse({'Success': 'OCR has been completed Successfully!', 'data': serializer.data},
                                status=200)
        except:
            return JsonResponse({'Success': 'OCR data', 'data': serializer.data},
                                status=200)