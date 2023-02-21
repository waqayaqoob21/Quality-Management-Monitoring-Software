import os
from datetime import datetime, timedelta

import xlwt
from django.db.models import F, Q
from django.http import JsonResponse, FileResponse, HttpResponse
from fpdf import FPDF

from ams.serializer import *
from ams.models import *
import fitz  # pip install PyMuPDF Pillow
import io
from PIL import Image
from pytesseract import pytesseract  # install tesseract-ocr-w64-setup-v5.2.0.20220712.exe (64 bit) resp.


# from https://github.com/UB-Mannheim/tesseract/wiki
# pip install pytesseract

class AmsController:

    @staticmethod
    def AddTask(request):
        taskModel = TaskSummary()
        try:
            id = request['id']
            if id == '0':
                taskModel.task_name = request['task_name']
                taskModel.assigned_by = request['assigned_by']
                taskModel.assigned_date = request['assigned_date']
                taskModel.assigned_to = request['assigned_to']
                taskModel.target_date = request['target_date']
                if request['task_date'] != '':
                    taskModel.task_date = request['task_date']
                taskModel.status = request['status']
                taskModel.remarks = request['remarks']
                taskModel.follow_up = request['follow_up']
                taskModel.save()
                return JsonResponse({'Success': "Task Created Successfully"}, status=200)
            else:
                get_task = TaskSummary.objects.filter(id=id).first()
                if get_task is not None:
                    if get_task.status != request['status'] or get_task.follow_up != request['follow_up'] or str(
                            get_task.assigned_date) != request['assigned_date'] or str(
                        get_task.target_date) != request['target_date'] or str(get_task.task_date) != \
                            request['remarks'] or get_task.status != request['remarks']:
                        # add entry in history
                        taskHistoryModal = TaskSummaryHistory()
                        taskHistoryModal.task_id = get_task.id
                        taskHistoryModal.task_name = get_task.task_name
                        taskHistoryModal.assigned_by = get_task.assigned_by
                        taskHistoryModal.assigned_date = get_task.assigned_date
                        taskHistoryModal.assigned_to = get_task.assigned_to
                        taskHistoryModal.target_date = get_task.target_date
                        if request['task_date'] != '':
                            taskHistoryModal.task_date = get_task.task_date

                        taskHistoryModal.status = get_task.status
                        taskHistoryModal.remarks = get_task.remarks
                        taskHistoryModal.follow_up = get_task.follow_up
                        taskHistoryModal.save()
                task = TaskSummary.objects.get(id=request['id'])
                task.task_name = request['task_name']
                task.assigned_by = request['assigned_by']
                task.assigned_date = request['assigned_date']
                task.assigned_to = request['assigned_to']
                task.target_date = request['target_date']
                if request['task_date'] != '':
                    task.task_date = request['task_date']
                else:
                    task.task_date = None
                task.status = request['status']
                task.remarks = request['remarks']
                task.follow_up = request['follow_up']
                task.save()
                return JsonResponse({'Success': 'Task Updated Successfully!'}, status=200)
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
            return JsonResponse({'Success': 'Task Updated Successfully!'}, status=201)
        except:
            return JsonResponse({'Error': 'Task Could Not Update Successfully!'}, status=400)

    @staticmethod
    def GetTaskList(request):
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

            status = request.query_params['status']
            selected_group = request.query_params['group']
            selected_year = request.query_params['year']
            assignFrom = request.query_params.get('assign_date_from')
            assignTo = request.query_params.get('assign_date_to')
            data = []
            total_filter_objects = Q()
            all_filter_objects = Q()
            if selected_group != '':
                all_filter_objects &= get_filter(
                    'assigned_to', 'equal',
                    selected_group)
            if status == '!Completed':

                total_filter_objects &= get_filter(
                    'status', 'not_equal',
                    'Task Completed')
                data = TaskSummary.objects.filter(total_filter_objects);
                serializer = TaskSummarySerialzer(data, many=True)
                return JsonResponse({'message': 'Welcome to Home Page', 'data': serializer.data}, status=200)
            else:
                total_filter_objects &= get_filter(
                    'status', 'equal',
                    status)
            if status == 'current_overdue':
                current_over_due_filter = Q()
                current_over_due_filter &= get_filter(
                    'status', 'equal',
                    'Task in-process')
                data = TaskSummary.objects.filter(current_over_due_filter,
                                                  assigned_date__year=selected_year)
                result = []
                if data is not None:
                    for item in data:
                        if item.task_date is None:
                            item.task_date = datetime.today() + timedelta(hours=5)
                            item.task_date = item.task_date.date()
                        if item.target_date < item.task_date:
                            result.append(item)
                serializer = TaskSummarySerialzer(result, many=True)
                return JsonResponse({'message': 'true', 'data': serializer.data}, status=200)

            if status == 'total_current_year':
                current_total_filter = Q()
                if selected_group != '':
                    current_total_filter &= get_filter(
                        'assigned_to', 'equal',
                        selected_group)
                data = TaskSummary.objects.filter(current_total_filter, assigned_date__year=selected_year)
                serializer = TaskSummarySerialzer(data, many=True)
                return JsonResponse({'message': 'true', 'data': serializer.data}, status=200)
            if status == 'total_overdue':
                current_over_due_filter = Q()
                current_over_due_filter &= get_filter(
                    'status', 'equal',
                    'Task in-process')
                data = TaskSummary.objects.filter(current_over_due_filter, task_date__gt=F('target_date'))
                serializer = TaskSummarySerialzer(data, many=True)
                return JsonResponse({'message': 'true', 'data': serializer.data}, status=200)
            if status == 'all_tasks':
                data = TaskSummary.objects.filter(all_filter_objects)
                if assignFrom != '' and assignTo != '':
                    data = data.filter(assigned_date__gte=assignFrom,
                                       assigned_date__lte=assignTo).order_by('-id')
                serializer = TaskSummarySerialzer(data, many=True)
                return JsonResponse({'message': 'true', 'data': serializer.data}, status=200)
            if status == 'total_TaskCompleted':
                all_filter_objects &= get_filter(
                    'status', 'equal',
                    'Task Completed')
                data = TaskSummary.objects.filter(all_filter_objects)
                serializer = TaskSummarySerialzer(data, many=True)
                return JsonResponse({'message': 'true', 'data': serializer.data}, status=200)

            if status == 'total_TaskNotCompleted':
                all_filter_objects &= get_filter(
                    'status', 'not_equal',
                    'Task Completed')
                data = TaskSummary.objects.filter(all_filter_objects)
                serializer = TaskSummarySerialzer(data, many=True)
                return JsonResponse({'message': 'true', 'data': serializer.data}, status=200)

            if status == 'all' and selected_year is not '':
                data = TaskSummary.objects.filter(assigned_date__year=selected_year)
                serializer = TaskSummarySerialzer(data, many=True)
                return JsonResponse({'message': 'true', 'data': serializer.data}, status=200)
            if selected_year is not '' and selected_group is '':
                data = TaskSummary.objects.filter(assigned_date__year=selected_year,
                                                  status=status)
            elif selected_year is '' and selected_group is not '':
                data = TaskSummary.objects.filter(assigned_to=selected_group,
                                                  status=status)

            elif selected_year is not '' and selected_group is not '':
                data = TaskSummary.objects.filter(assigned_to=selected_group,
                                                  assigned_date__year=selected_year,
                                                  status=status)
            if status =='Task in-process':
                result =[]
                if data is not None:
                    for item in data:
                        if item.task_date is None:
                            item.task_date = datetime.today() + timedelta(hours=5)
                            item.task_date = item.task_date.date()
                        if item.task_date < item.target_date:
                            result.append(item)
                    serializer = TaskSummarySerialzer(result, many=True)
                    return JsonResponse({'message': 'true', 'data': serializer.data}, status=200)


            serializer = TaskSummarySerialzer(data, many=True)
            return JsonResponse({'message': 'Welcome to Home Page', 'data': serializer.data}, status=200)
        except Exception as e:
            return JsonResponse({'message': 'Sorry! No Task found.'}, status=500)

    @staticmethod
    def GetTaskListHistory(request):
        try:
            id = request.query_params['id']
            data = TaskSummaryHistory.objects.filter(task_id=id)

            serializer = TaskSummarySerialzer(data, many=True)
            return JsonResponse({'message': 'Welcome to Home Page', 'data': serializer.data}, status=200)
        except:
            return JsonResponse({'message': 'Sorry! No Task found.'}, status=200)

    @staticmethod
    def GetDocumentPDFList(request):

        status = request.query_params['status']
        type = request.query_params['type']
        data = []
        if type == "group":
            data = TaskSummary.objects.filter(assigned_to=status).order_by('-id').values_list('task_name',
                                                                                              'assigned_by',
                                                                                              'assigned_date',
                                                                                              'assigned_to',
                                                                                              'target_date',
                                                                                              'task_date',
                                                                                              'status',
                                                                                              'remarks')
        if type == "status":
            if status == "Total Tasks":
                data = TaskSummary.objects.all().order_by('-id').values_list('task_name',
                                                                             'assigned_by',
                                                                             'assigned_date',
                                                                             'assigned_to',
                                                                             'target_date',
                                                                             'task_date',
                                                                             'status',
                                                                             'remarks')
            else:
                data = TaskSummary.objects.filter(status=status).order_by('-id').values_list('task_name',
                                                                                             'assigned_by',
                                                                                             'assigned_date',
                                                                                             'assigned_to',
                                                                                             'target_date',
                                                                                             'task_date',
                                                                                             'status',
                                                                                             'remarks')
        TABLE_COL_NAMES = (
            "Task Name", "Assigned By", "Assigned Date", "Assigned To", "Target Date", "Task Completion Date", "Status",
            "Remarks")
        # data = TaskSummary.objects.all().order_by('-id').values_list('task_name',
        #                                                              'assigned_by',
        #                                                              'assigned_date',
        #                                                              'assigned_to',
        #                                                              'target_date',
        #                                                              'task_date',
        #                                                              'status',
        #                                                              'remarks')

        pdf = FPDF('p', 'mm', 'Legal')
        pdf.add_page()
        pdf.set_font('courier', 'B', 26)
        pdf.cell(330, 10, 'AMS Report', border=0, align='C', ln=2)
        pdf.cell(40, 10, '', 0, 1)
        pdf.set_font("Times", size=10)
        line_height = pdf.font_size * 2.5
        col_width = pdf.epw / 8

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
        pdf.output('ams-report.pdf')
        return FileResponse(open('ams-report.pdf', 'rb'), as_attachment=True, content_type='application/pdf')

    @staticmethod
    def GetDocumentExcelList(request):
        response = HttpResponse(content_type='application/ms-excel')
        response['Content-Disposition'] = 'attachment; filename="DocumentExcelSheet.xls"'
        work_book = xlwt.Workbook(encoding='utf-8')
        work_sheet = work_book.add_sheet('ExcelSheet')
        row_num = 0
        font_style = xlwt.XFStyle()
        font_style.font.bold = True
        TABLE_COL_NAMES = (
            "Task Name", "Assigned By", "Assigned Date", "Assigned To", "Target Date", "Task Completion Date", "Status",
            "Remarks")
        for col_num in range(len(TABLE_COL_NAMES)):
            work_sheet.write(row_num, col_num, TABLE_COL_NAMES[col_num], font_style)
        font_style = xlwt.XFStyle()
        status = request.query_params['status']
        type = request.query_params['type']
        data = []
        if type == "group":
            data = TaskSummary.objects.filter(assigned_to=status).order_by('-id').values_list('task_name',
                                                                                              'assigned_by',
                                                                                              'assigned_date',
                                                                                              'assigned_to',
                                                                                              'target_date',
                                                                                              'task_date',
                                                                                              'status',
                                                                                              'remarks')
        if type == "status":
            if status == "Total Tasks":
                data = TaskSummary.objects.all().order_by('-id').values_list('task_name',
                                                                             'assigned_by',
                                                                             'assigned_date',
                                                                             'assigned_to',
                                                                             'target_date',
                                                                             'task_date',
                                                                             'status',
                                                                             'remarks')
            else:
                data = TaskSummary.objects.filter(status=status).order_by('-id').values_list('task_name',
                                                                                             'assigned_by',
                                                                                             'assigned_date',
                                                                                             'assigned_to',
                                                                                             'target_date',
                                                                                             'task_date',
                                                                                             'status',
                                                                                             'remarks')
        # data = TaskSummary.objects.all().values_list('task_name', 'assigned_by', 'assigned_date', 'assigned_to',
        #                                              'target_date', 'task_date', 'status', 'remarks')
        for row in data:
            row_num += 1
            for col_num in range(len(row)):
                work_sheet.write(row_num, col_num, str(row[col_num]), font_style)
        work_book.save(response)
        return response

    @staticmethod
    def OcrPDF(request):
        ocrModel = OcrDataModel()

        filename = "test.pdf"
        attachment = request['uploaded_file']
        # open file
        with fitz.open(attachment) as my_pdf_file:
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
                    image.save(open(f"Page{page_number}Image{image_number}.{ext}", "wb"))

        path_to_tesseract = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
        image_path = r"Page1Image1.png"
        img = Image.open(image_path)
        pytesseract.tesseract_cmd = path_to_tesseract
        text = pytesseract.image_to_string(img)
        data = text.split("\n")
        final_list = [y for x in data for y in x.split(':')]
        for i in range(len(final_list)):
            ocrModel.organization = "NATIONAL DEVELOPMENT COMPLEX (NDC)"
            ocrModel.senior_directorate = "Waqar Afzal"
            ocrModel.site = "13A"
            if final_list[i] == "Test Report No":
                ocrModel.test_report_no = final_list[i + 1]
                print(ocrModel.test_report_no)

            if final_list[i] == "Job Card No":
                ocrModel.job_card_no = final_list[i + 1]

            ocrModel.test_report_date = "2022-09-03"
            if final_list[i] == "Product Name":
                ocrModel.product_name = final_list[i + 1]

            if final_list[i] == "ID No":
                ocrModel.id_no = final_list[i + 1]

            if final_list[i] == "Lot No & Lot Size":
                ocrModel.lot_no_lot_size = final_list[i + 1]

            ocrModel.test_name = "Random Vibration"
            ocrModel.test_type = "General"

            if final_list[i] == "Qualification Standard":
                ocrModel.qualification_standard = final_list[i + 1]
            ocrModel.test_specifications = "1400 Hz 0016 g%/Hz"
            ocrModel.results = "Qualified"
            ocrModel.remarks = "Not given"
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
