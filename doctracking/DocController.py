import base64
import textwrap
from datetime import datetime, timedelta

from dateutil.relativedelta import relativedelta
from django.db import connection
from django.db.models import F, Q
from django.http import JsonResponse, FileResponse, HttpResponse
from fpdf import FPDF
import xlwt

from ams.models import TaskSummary
from doctracking.models import doctrackingHistory
from doctracking.serializer import DocListSerializer
from mpm.models import ActiveMotors
from qms.models import QmsAudit, CespAudit
from sms.models import ProductionSystemStatus, FlightSystemStatus, RelifingSystemStatus
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
                if request['task_date'] != "":
                    docModal.task_date = request['task_date']
                docModal.status = request['status']
                docModal.sent_to = request['sent_to']
                if request['sent_date'] != "":
                    docModal.sent_date = request['sent_date']
                docModal.product_sr_no = request['product_sr_no']
                docModal.tracking_id = request['tracking_id']

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
                    if str(get_doc.receive_date.date()) != request['receive_date'] or str(get_doc.marked_date.date()) != \
                            request['marked_date'] or str(get_doc.due_date.date()) != request['due_date']:
                        # add entry in history table
                        docHistoryModal = doctrackingHistory()
                        docHistoryModal.doc_name = get_doc.doc_name
                        docHistoryModal.doc_type = get_doc.doc_type
                        docHistoryModal.sender = get_doc.sender
                        docHistoryModal.receive_date = get_doc.receive_date
                        docHistoryModal.marked_to = get_doc.marked_to
                        docHistoryModal.marked_date = get_doc.marked_date
                        docHistoryModal.due_date = get_doc.due_date
                        docHistoryModal.task_date = get_doc.task_date
                        docHistoryModal.status = get_doc.status
                        docHistoryModal.sent_to = get_doc.sent_to
                        docHistoryModal.sent_date = get_doc.sent_date
                        docHistoryModal.product_sr_no = get_doc.product_sr_no
                        docHistoryModal.tracking_id = get_doc.tracking_id
                        docHistoryModal.remarks = get_doc.remarks
                        docHistoryModal.doc_id = id
                        docHistoryModal.save()
                        # save data in history end

                get_doc.doc_name = request['doc_name']
                get_doc.doc_type = request['doc_type']
                get_doc.sender = request['sender']
                get_doc.receive_date = request['receive_date']
                get_doc.marked_to = request['marked_to']
                get_doc.marked_date = request['marked_date']
                get_doc.due_date = request['due_date']
                if request['task_date'] != "":
                    get_doc.task_date = request['task_date']
                else:
                    get_doc.task_date = None
                get_doc.status = request['status']
                get_doc.sent_to = request['sent_to']
                if request['sent_date'] != "":
                    get_doc.sent_date = request['sent_date']
                else:
                    get_doc.sent_date = None
                get_doc.product_sr_no = request['product_sr_no']
                get_doc.tracking_id = request['tracking_id']
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
    def DeleteDocument(request):

        try:
            id = request.query_params['id']
            delete = doctracking.objects.filter(id=id).delete()

            return JsonResponse({'status': 'True', 'message': "Record Deleted"},
                                status=200)

        except Exception as e:
            print(e)
            return JsonResponse({'status': 'False', "message": "Doc Not Deleted"}, status=500)
            pass

    @staticmethod
    def DeleteTask(request):

        try:
            id = request.query_params['id']
            delete = TaskSummary.objects.filter(id=id).delete()

            return JsonResponse({'status': 'True', 'message': "Record Deleted"},
                                status=200)

        except Exception as e:
            print(e)
            return JsonResponse({'status': 'False', "message": "Task Not Deleted"}, status=500)
            pass

    @staticmethod
    def GetDocumentList(request):

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

            current_status = request.query_params.get('current_status')
            current_year = request.query_params.get('year')
            current_org = request.query_params.get('org')
            current_type = request.query_params.get('type')
            receiveFrom = request.query_params.get('receive_date_from')
            receiveTo = request.query_params.get('receive_date_to')
            total_filter_objects = Q()
            if current_org != '' and current_org != 'All':
                total_filter_objects &= get_filter(
                    'sender', 'equal',
                    current_org)
            current_filter_objects = Q()
            if current_org != '' and current_org != 'All':
                current_filter_objects &= get_filter(
                    'sender', 'equal',
                    current_org)
            if current_status == 'Total Documents':
                docList = doctracking.objects.all().order_by('-id')
                if current_year != '':
                    docList = docList.filter(receive_date__year=current_year)
                if current_org != '':
                    docList = docList.filter(sender=current_org)
                # if current_type != '':
                current_filter_objects = Q()
                if current_type == 'document' or current_type == '':
                    current_filter_objects &= get_filter(
                        'doc_type', 'not_equal',
                        'BHD')
                else:
                    current_filter_objects &= get_filter(
                        'doc_type', 'equal',
                        'BHD')
                docList = docList.filter(current_filter_objects)
                # docList = docList.filter(sender=current_org)
                serializer = DocListSerializer(docList, many=True)
                # print(serializer)
                return JsonResponse({'status': 'True', 'data': serializer.data},
                                    status=200)
            if current_status == 'total':

                if current_type == 'document' or current_type == '':
                    total_filter_objects &= get_filter(
                        'doc_type', 'not_equal',
                        'BHD')
                    total_filter_objects &= get_filter(
                        'status', 'not_equal',
                        'QM Certificate issued')
                    docList = doctracking.objects.filter(total_filter_objects).order_by('-id')

                    if receiveFrom != '' and receiveTo != '':
                        docList = docList.filter(receive_date__gte=receiveFrom,
                                                             receive_date__lte=receiveTo).order_by('-id')
                        # serializer = DocListSerializer(docList, many=True)
                        # return JsonResponse({'status': 'True', 'data': serializer.data},
                        #                     status=200)
                    serializer = DocListSerializer(docList, many=True)
                    return JsonResponse({'status': 'True', 'data': serializer.data},
                                        status=200)
                else:

                    total_filter_objects &= get_filter(
                        'doc_type', 'equal',
                        'BHD')
                    docList = doctracking.objects.filter(total_filter_objects).order_by('-id')
                    if receiveFrom != '' and receiveTo != '':
                        docList = docList.filter(receive_date__gte=receiveFrom,
                                                             receive_date__lte=receiveTo).order_by('-id')
                        # serializer = DocListSerializer(docList, many=True)
                        # return JsonResponse({'status': 'True', 'data': serializer.data},
                        #                     status=200)
                    serializer = DocListSerializer(docList, many=True)
                    return JsonResponse({'status': 'True', 'data': serializer.data},
                                        status=200)

            if current_status == 'totalApproved':
                total_filter_objects &= get_filter(
                    'status', 'equal',
                    'Approved')
                if current_type == 'document' or current_type == '':
                    total_filter_objects &= get_filter(
                        'doc_type', 'not_equal',
                        'BHD')
                else:
                    total_filter_objects &= get_filter(
                        'doc_type', 'equal',
                        'BHD')
                docList = doctracking.objects.filter(total_filter_objects).order_by('-id')
                serializer = DocListSerializer(docList, many=True)
                return JsonResponse({'status': 'True', 'data': serializer.data},
                                    status=200)
            if current_status == 'totalQMIssued':
                docList = doctracking.objects.filter(doc_type='BHD', status='QM Certificate issued').order_by('-id')
                serializer = DocListSerializer(docList, many=True)
                return JsonResponse({'status': 'True', 'data': serializer.data},
                                    status=200)
            if current_status == 'totaloverdue':
                total_filter_document_totaloverdue = Q()
                if current_type == 'document' or current_type == '':

                    total_filter_document_totaloverdue &= get_filter(
                        'doc_type', 'not_equal',
                        'BHD')
                    total_filter_document_totaloverdue &= get_filter(
                        'status', 'equal',
                        'Audit in-process')
                    # total_filter_objects &= get_filter(
                    #     'doc_type', 'not_equal',
                    #     'BHD')
                    # total_filter_objects &= get_filter(
                    #     'status', 'equal',
                    #     'Audit in-process')
                    if current_org != '' and current_org != 'All':
                        total_filter_document_totaloverdue &= get_filter(
                            'sender', 'equal',
                            current_org)
                else:
                    total_filter_document_totaloverdue &= get_filter(
                        'doc_type', 'equal',
                        'BHD')
                    total_filter_document_totaloverdue &= get_filter(
                        'status', 'equal',
                        'Audit in-process')
                    if current_org != '' and current_org != 'All':
                        total_filter_document_totaloverdue &= get_filter(
                            'sender', 'equal',
                            current_org)

                total_over_due_all = doctracking.objects.filter(total_filter_document_totaloverdue).order_by('-id')
                list = []
                if total_over_due_all is not None:
                    for item in total_over_due_all:
                        if item.task_date is None:
                            item.task_date = datetime.today() + timedelta(hours=5)
                        if item.due_date.date() < item.task_date.date():
                            list.append(item)

                docList = list  # doctracking.objects.filter(total_filter_document_totaloverdue,
                #                          due_date__lt=F('task_date')).order_by('-id')

                serializer = DocListSerializer(docList, many=True)
                return JsonResponse({'status': 'True', 'data': serializer.data},
                                    status=200)
            if current_status == 'totalnotApprovedQM':
                total_filter_objects &= get_filter(
                    'status', 'not_equal',
                    'Approved')

                if current_type == 'document' or current_type == '':
                    total_filter_objects &= get_filter(
                        'doc_type', 'not_equal',
                        'BHD')
                    total_filter_objects &= get_filter(
                        'status', 'equal',
                        'Audit in-process')
                else:
                    total_filter_objects &= get_filter(
                        'doc_type', 'equal',
                        'BHD')
                    total_filter_objects &= get_filter(
                        'status', 'equal',
                        'Audit in-process')
                if current_org != '' and current_org != 'All':
                    total_filter_objects &= get_filter(
                        'sender', 'equal',
                        current_org)
                docList = doctracking.objects.filter(total_filter_objects).order_by('-id')
                serializer = DocListSerializer(docList, many=True)
                return JsonResponse({'status': 'True', 'data': serializer.data},
                                    status=200)
            if current_status == 'totalnotApprovedGM':

                if current_type == 'document' or current_type == '':
                    total_filter_objects &= get_filter(
                        'doc_type', 'not_equal',
                        'BHD')
                    total_filter_objects &= get_filter(
                        'status', 'equal',
                        'QM Observations Forwarded')
                    total_filter_objects |= get_filter(
                        'status', 'equal',
                        'QM Observations Repeated')
                else:
                    total_filter_objects &= get_filter(
                        'doc_type', 'equal',
                        'BHD')
                    total_filter_objects &= get_filter(
                        'status', 'equal',
                        'QM Observations Forwarded')
                    total_filter_objects |= get_filter(
                        'status', 'equal',
                        'QM Observations Repeated')
                total_doc_type_objects = Q()
                if current_type == 'document' or current_type == '':
                    total_doc_type_objects &= get_filter(
                        'doc_type', 'not_equal',
                        'BHD')
                else:
                    total_doc_type_objects &= get_filter(
                        'doc_type', 'equal',
                        'BHD')
                docList = doctracking.objects.filter(total_filter_objects).order_by('-id')
                docList = docList.filter(total_doc_type_objects)
                serializer = DocListSerializer(docList, many=True)
                return JsonResponse({'status': 'True', 'data': serializer.data},
                                    status=200)
            if current_status == 'current_doc_approved':

                if current_type == 'document' or current_type == '':
                    current_filter_objects &= get_filter(
                        'doc_type', 'not_equal',
                        'BHD')
                    current_filter_objects &= get_filter(
                        'status', 'not_equal',
                        'QM Certificate issued')
                    current_filter_objects &= get_filter(
                        'status', 'equal',
                        'Approved')
                if current_org != '' and current_org != 'All':
                    current_filter_objects &= get_filter(
                        'sender', 'equal',
                        current_org)

                else:
                    current_filter_objects &= get_filter(
                        'doc_type', 'equal',
                        'BHD')
                    current_filter_objects &= get_filter(
                        'status', 'not_equal',
                        'Approved')
                    if current_org != '' and current_org != 'All':
                        current_filter_objects &= get_filter(
                            'sender', 'equal',
                            current_org)
                docList = doctracking.objects.filter(current_filter_objects, receive_date__year=current_year).order_by(
                    '-id')
                serializer = DocListSerializer(docList, many=True)
                return JsonResponse({'status': 'True', 'data': serializer.data},
                                    status=200)
            if current_status == 'current_doc_total':
                if current_type == 'document' or current_type == '':
                    current_filter_objects &= get_filter(
                        'doc_type', 'not_equal',
                        'BHD')
                    current_filter_objects &= get_filter(
                        'status', 'not_equal',
                        'QM Certificate issued')


                else:
                    current_filter_objects &= get_filter(
                        'doc_type', 'equal',
                        'BHD')
                    current_filter_objects &= get_filter(
                        'status', 'not_equal',
                        'Approved')
                docList = doctracking.objects.filter(current_filter_objects, receive_date__year=current_year).order_by(
                    '-id')
                serializer = DocListSerializer(docList, many=True)
                return JsonResponse({'status': 'True', 'data': serializer.data},
                                    status=200)
            if current_status == 'current_cer_issued':
                current_filter_objects &= get_filter(
                    'doc_type', 'equal',
                    'BHD')
                current_filter_objects &= get_filter(
                    'status', 'equal',
                    'QM Certificate issued')
                docList = doctracking.objects.filter(current_filter_objects, receive_date__year=current_year).order_by(
                    '-id')
                serializer = DocListSerializer(docList, many=True)
                return JsonResponse({'status': 'True', 'data': serializer.data},
                                    status=200)
            if current_status == 'current_QM_Observations_Forwarded':
                if current_type == 'document' or current_type == '':
                    current_filter_objects &= get_filter(
                        'doc_type', 'not_equal',
                        'BHD')
                    current_filter_objects &= get_filter(
                        'status', 'equal',
                        'QM Observations Forwarded')

                else:
                    current_filter_objects &= get_filter(
                        'doc_type', 'equal',
                        'BHD')
                    current_filter_objects &= get_filter(
                        'status', 'equal',
                        'QM Observations Forwarded')
                docList = doctracking.objects.filter(current_filter_objects, receive_date__year=current_year).order_by(
                    '-id')
                serializer = DocListSerializer(docList, many=True)
                return JsonResponse({'status': 'True', 'data': serializer.data},
                                    status=200)
            if current_status == 'current_QM_Observations_Repeated':
                if current_type == 'document' or current_type == '':
                    current_filter_objects &= get_filter(
                        'doc_type', 'not_equal',
                        'BHD')
                    current_filter_objects &= get_filter(
                        'status', 'equal',
                        'QM Observations Repeated')

                else:
                    current_filter_objects &= get_filter(
                        'doc_type', 'equal',
                        'BHD')
                    current_filter_objects &= get_filter(
                        'status', 'equal',
                        'QM Observations Repeated')
                docList = doctracking.objects.filter(current_filter_objects, receive_date__year=current_year).order_by(
                    '-id')
                serializer = DocListSerializer(docList, many=True)
                return JsonResponse({'status': 'True', 'data': serializer.data},
                                    status=200)
            if current_status == 'current_Audit_in_process':
                if current_type == 'document' or current_type == '':
                    current_filter_objects &= get_filter(
                        'doc_type', 'not_equal',
                        'BHD')
                    current_filter_objects &= get_filter(
                        'status', 'equal',
                        'Audit in-process')

                else:
                    current_filter_objects &= get_filter(
                        'doc_type', 'equal',
                        'BHD')
                    current_filter_objects &= get_filter(
                        'status', 'equal',
                        'Audit in-process')
                docList = doctracking.objects.filter(current_filter_objects, receive_date__year=current_year).order_by(
                    '-id')
                serializer = DocListSerializer(docList, many=True)
                return JsonResponse({'status': 'True', 'data': serializer.data},
                                    status=200)
            else:
                if current_status == 'overdue':
                    total_filter_document_totaloverdue = Q()
                    if current_type == 'document' or current_type == '':
                        total_filter_document_totaloverdue &= get_filter(
                            'doc_type', 'not_equal',
                            'BHD')
                        total_filter_document_totaloverdue &= get_filter(
                            'status', 'equal',
                            'Audit in-process')
                    else:
                        total_filter_document_totaloverdue &= get_filter(
                            'doc_type', 'equal',
                            'BHD')
                        total_filter_document_totaloverdue &= get_filter(
                            'status', 'equal',
                            'Audit in-process')
                        # total_filter_document_totaloverdue &= get_filter(
                        #     'status', 'not_equal',
                        #     'QM Certificate issued')
                        # total_filter_document_totaloverdue &= get_filter(
                        #     'status', 'not_equal',
                        #     'Approved')

                    total_over_due_current = doctracking.objects.filter(total_filter_document_totaloverdue,
                                                                        receive_date__year=current_year).order_by('-id')
                    list = []
                    if total_over_due_current is not None:
                        for item in total_over_due_current:
                            if item.task_date is None:
                                item.task_date = datetime.today() + timedelta(hours=5)
                            if item.due_date.date() < item.task_date.date():
                                list.append(item)
                    docList = list  # doctracking.objects.filter(total_filter_document_totaloverdue,
                    #                          receive_date__year=current_year,
                    #                         due_date__lt=F('task_date')).order_by('-id')

                    serializer = DocListSerializer(docList, many=True)
                    return JsonResponse({'status': 'True', 'data': serializer.data},
                                        status=200)
                else:
                    total_filter_objects = Q()
                    if current_status == '!Approved':

                        total_filter_objects &= get_filter(
                            'status', 'not_equal',
                            'Approved')
                    else:

                        total_filter_objects &= get_filter(
                            'status', 'equal',
                            current_status)
                    docList = doctracking.objects.filter(total_filter_objects).order_by('-id')
                    if current_year != '':
                        docList = docList.filter(receive_date__year=current_year)
                    if current_org != '':
                        docList = docList.filter(sender=current_org)
                    current_filter_objects = Q()
                    current_filter_objects &= get_filter(
                        'doc_type', 'equal',
                        current_type)

                    docList = docList.filter(current_filter_objects)
                    serializer = DocListSerializer(docList, many=True)
                    return JsonResponse({'status': 'True', 'data': serializer.data},
                                        status=200)

        except Exception as e:
            print(e)
            return JsonResponse({'status': 'False', "message": "Internal Server Error"}, status=500)
            pass

    @staticmethod
    def GetDocumentHistory(request):

        try:
            doc_id = request.query_params.get('id')

            docList = doctrackingHistory.objects.filter(doc_id=doc_id).order_by('-id')
            serializer = DocListSerializer(docList, many=True)
            return JsonResponse({'status': 'True', 'data': serializer.data},
                                status=200)

        except Exception as e:
            print(e)
            return JsonResponse({'status': 'False', "message": "Internal Server Error"}, status=500)
            pass

    @staticmethod
    def GetPendingDocument(request):

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
            # data = doctracking.objects.filter(due_date__lt=datetime.today()).values()
            serializer = DocListSerializer(doc_list, many=True)
            print(serializer)
            return JsonResponse({'status': 'True', 'data': serializer.data},
                                status=200)

        except Exception as e:
            print(e)
            return JsonResponse({'status': 'False', "message": "Internal Server Error"}, status=500)
            pass

    @staticmethod
    def GetDocumentPDFList(request):
        TABLE_COL_NAMES = (
            "Document Name", "Document Type", "Sender", "Receive Date", "Marked To", "Marked Date", "Due Date",
            "Task Date",
            "Status", "Sent To", "Sent Date", "Remarks")

        current_status = request.query_params.get('current_status')
        data = []
        if current_status == 'Total Documents':
            data = doctracking.objects.all().order_by('-id').values_list('doc_name',
                                                                         'doc_type',
                                                                         'sender',
                                                                         'receive_date',
                                                                         'marked_to',
                                                                         'marked_date',
                                                                         'due_date',
                                                                         'task_date',
                                                                         'status',
                                                                         'sent_to',
                                                                         'sent_date',
                                                                         'remarks')
        else:
            data = doctracking.objects.filter(status=current_status).order_by('-id').values_list('doc_name',
                                                                                                 'doc_type',
                                                                                                 'sender',
                                                                                                 'receive_date',
                                                                                                 'marked_to',
                                                                                                 'marked_date',
                                                                                                 'due_date',
                                                                                                 'task_date',
                                                                                                 'status',
                                                                                                 'sent_to',
                                                                                                 'sent_date',
                                                                                                 'remarks')

        pdf = FPDF('L', 'mm', 'Legal')
        pdf.add_page()
        pdf.set_font('courier', 'B', 26)
        pdf.cell(330, 10, 'QMS Report', border=0, align='C', ln=2)
        pdf.cell(40, 10, '', 0, 1)
        pdf.set_font("Times", size=10)
        line_height = pdf.font_size * 2.5
        col_width = pdf.epw / 12

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
                lh_list.append(20)
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
        pdf.output('QMS_Report.pdf')
        return FileResponse(open('QMS_Report.pdf', 'rb'), as_attachment=True, content_type='application/pdf')

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
            "Document Name", "Document Type", "Sender", "Receive Date", "Marked To", "Marked Date", "Due Date",
            "Task Date",
            "Status", "Sent To", "Sent Date", "Remarks")
        for col_num in range(len(TABLE_COL_NAMES)):
            work_sheet.write(row_num, col_num, TABLE_COL_NAMES[col_num], font_style)
        font_style = xlwt.XFStyle()

        current_status = request.query_params.get('current_status')
        data = []
        if current_status == 'Total Documents':
            data = doctracking.objects.all().values_list('doc_name', 'doc_type', 'sender',
                                                         'receive_date', 'marked_to',
                                                         'marked_date', 'due_date',
                                                         'task_date', 'status', 'sent_to',
                                                         'sent_date', 'remarks')
        else:
            data = doctracking.objects.filter(status=current_status).values_list('doc_name', 'doc_type', 'sender',
                                                                                 'receive_date', 'marked_to',
                                                                                 'marked_date', 'due_date',
                                                                                 'task_date', 'status', 'sent_to',
                                                                                 'sent_date', 'remarks')
        for row in data:
            row_num += 1
            for col_num in range(len(row)):
                work_sheet.write(row_num, col_num, str(row[col_num]), font_style)
        work_book.save(response)
        return response

    @staticmethod
    def GetDashboardCount(request):
        try:
            DataCount = []
            qm_total_doc = doctracking.objects.all().count()

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
            # data = doctracking.objects.filter(due_date__lt=datetime.today()).values()
            serializer = DocListSerializer(doc_list, many=True)

            docList = doctracking.objects.all().order_by('-id')
            serializer = DocListSerializer(docList, many=True)

            autt_inProcess = doctracking.objects.filter(status='Audit in-process').count()
            AuditCompleted = doctracking.objects.filter(status='Audit completed').count()
            qm_certificateIssued = doctracking.objects.filter(status='QM Certificate issued').count()
            am_observation_forwarded = doctracking.objects.filter(status='QM observations forwarded').count()
            am_observation_forwarded_unsettled = doctracking.objects.filter(
                status='QM observations forwarded (un-settled)').count()
            total_tasks = TaskSummary.objects.all().count()
            ams_completed = TaskSummary.objects.filter(status='Completed').count()
            ams_not_completed = TaskSummary.objects.filter(status='Not Completed').count()
            ams_group1 = TaskSummary.objects.filter(assigned_to='ADG (QM)').count()
            ams_group2 = TaskSummary.objects.filter(assigned_to='DDG (Cert)').count()
            ams_group3 = TaskSummary.objects.filter(assigned_to='DDG (Process)').count()
            ams_group4 = TaskSummary.objects.filter(assigned_to='DDG (System)').count()
            prodSysCount = ProductionSystemStatus.objects.all().count()
            flightSysCount = FlightSystemStatus.objects.all().count()
            refilingSysCount = RelifingSystemStatus.objects.all().count()

            ams_group5 = TaskSummary.objects.filter(assigned_to='Dir (QMS)').count()
            ams_group6 = TaskSummary.objects.filter(assigned_to='Dir (Reliablity)').count()
            ams_group7 = TaskSummary.objects.filter(assigned_to='Dir (SQA)').count()
            ams_group8 = TaskSummary.objects.filter(assigned_to='Dir (CeSP)').count()
            active_motors = ActiveMotors.objects.filter().count()
            cesp_audit = CespAudit.objects.all().count()
            prodSysQmsInProcess = ProductionSystemStatus.objects.filter(
                qm_certification_status='QM certification in-process').count()
            prodSysQmsIssued = ProductionSystemStatus.objects.filter(
                qm_certification_status='QM certification issued').count()

            flightSysQmsInProcess = FlightSystemStatus.objects.filter(
                qm_certification_status='QM certification in-process').count()
            flightSysQmsInIssued = FlightSystemStatus.objects.filter(
                qm_certification_status='QM certification issued').count()
            qms_audit = QmsAudit.objects.all().count()

            relifingSysQmsInProcess = RelifingSystemStatus.objects.filter(
                qm_certification_status='QM certification in-process').count()
            relifingSysQmsInIssued = RelifingSystemStatus.objects.filter(
                qm_certification_status='QM certification issued').count()
            dist = {
                'qm_total_doc': qm_total_doc,
                'qm_pending_doc': serializer.data.__len__(),
                'qm_audit_inprocess_doc': autt_inProcess,
                'qm_completed_doc': AuditCompleted,
                'qm_certificateIssued': qm_certificateIssued,
                'am_observation_forwarded': am_observation_forwarded,
                'am_observation_forwarded_unsettled': am_observation_forwarded_unsettled,
                'ams_total': total_tasks,
                'ams_completed': ams_completed,
                'ams_not_completed': ams_not_completed,
                'ams_group1': ams_group1,
                'ams_group2': ams_group2,
                'ams_group3': ams_group3,
                'ams_group4': ams_group4,
                'prodSysCount': prodSysCount,
                'flightSysCount': flightSysCount,
                'ams_group5': ams_group5,
                'ams_group6': ams_group6,
                'ams_group7': ams_group7,
                'ams_group8': ams_group8,
                'refilingSysCount': refilingSysCount,
                'active_motors': active_motors,
                'cesp_audit': cesp_audit,
                'qms_audit': qms_audit,
                'prodSysQmsInProcess': prodSysQmsInProcess,
                'prodSysQmsIssued': prodSysQmsIssued,
                'flightSysQmsInProcess': flightSysQmsInProcess,
                'flightSysQmsInIssued': flightSysQmsInIssued,
                'relifingSysQmsInProcess': relifingSysQmsInProcess,
                'relifingSysQmsInIssued': relifingSysQmsInIssued

            }

            DataCount.append(dist)
            return JsonResponse({'status': 'True', 'data': DataCount},
                                status=200)

        except Exception as e:
            print(e)
            return JsonResponse({'status': 'False', "message": "Internal Server Error"}, status=500)
            pass

    @staticmethod
    def GetDocTrackingDashboardCount(request, self=None):
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

            selected_year = request.query_params.get('selected_year')
            selected_org = request.query_params.get('selected_org')
            selected_type = request.query_params.get('selected_type')
            if selected_type == '':
                selected_type = 'document'
            typeQuery = Q()
            filter_objects = Q()
            total_filter_objects = Q()
            total_filter_objects &= get_filter(
                'status', 'not_equal',
                'Approved')
            total_filter_objects &= get_filter(
                'status', 'equal',
                'Audit in-process')
            if selected_type == 'document' or selected_type == '':
                total_filter_objects &= get_filter(
                    'doc_type', 'not_equal',
                    'BHD')
            else:
                total_filter_objects &= get_filter(
                    'doc_type', 'equal',
                    'BHD')
            if selected_org != '' and selected_org != 'All':
                total_filter_objects &= get_filter(
                    'sender', 'equal',
                    selected_org)

            total_filter_objectsGN = Q()
            total_filter_objectsGN &= get_filter(
                'status', 'equal',
                'QM Observations Forwarded')
            total_filter_objectsGN |= get_filter(
                'status', 'equal',
                'QM Observations Repeated')
            if selected_type == 'document' or selected_type == '':
                total_filter_objectsGN &= get_filter(
                    'doc_type', 'not_equal',
                    'BHD')
            else:
                total_filter_objectsGN &= get_filter(
                    'doc_type', 'equal',
                    'BHD')
            if selected_org != '' and selected_org != 'All':
                total_filter_objectsGN &= get_filter(
                    'sender', 'equal',
                    selected_org)

            overdueDoc_filter_objects = Q()
            overdueDoc_filter_objects &= get_filter(
                'status', 'not_equal',
                'Approved')

            overdueBHd_filter_objects = Q()
            overdueBHd_filter_objects &= get_filter(
                'status', 'equal',
                'Audit in-process')
            # overdueBHd_filter_objects &= get_filter(
            #     'status', 'not_equal',
            #     'QM Certificate issued')

            if selected_type == 'BHD':
                filter_objects &= get_filter(
                    'doc_type', 'equal',
                    'BHD')
            else:
                filter_objects &= get_filter(
                    'doc_type', 'not_equal',
                    'BHD')

            total_filter_bhd = Q()
            total_filter_bhd &= get_filter(
                'status', 'not_equal',
                'Approved')
            total_filter_bhd &= get_filter(
                'status', 'equal',
                'Audit in-process')
            if selected_org != '' and selected_org != 'All':
                total_filter_bhd &= get_filter(
                    'sender', 'equal',
                    selected_org)

            total_filter_bhd_NG = Q()
            total_filter_bhd_NG &= get_filter(
                'status', 'not_equal',
                'Approved')
            total_filter_bhd_NG &= get_filter(
                'status', 'equal',
                'QM Observations Forwarded')
            total_filter_bhd_NG |= get_filter(
                'status', 'equal',
                'QM Observations Repeated')
            if selected_org != '' and selected_org != 'All':
                total_filter_bhd_NG &= get_filter(
                    'sender', 'equal',
                    selected_org)

            total_filter_document = Q()
            total_filter_document &= get_filter(
                'doc_type', 'not_equal',
                'BHD')
            total_filter_document &= get_filter(
                'status', 'equal',
                'Audit in-process')
            if selected_org != '' and selected_org != 'All':
                total_filter_document &= get_filter(
                    'sender', 'equal',
                    selected_org)

            total_type_bhd = Q()
            total_type_bhd &= get_filter(
                'doc_type', 'equal',
                'BHD')
            total_type_bhd &= get_filter(
                'status', 'equal',
                'Audit in-process')
            # total_type_bhd &= get_filter(
            #     'status', 'not_equal',
            #     'Approved')
            if selected_org != '' and selected_org != 'All':
                total_type_bhd &= get_filter(
                    'sender', 'equal',
                    selected_org)

            total_doc_documents = Q()
            total_doc_documents &= get_filter(
                'status', 'not_equal',
                'QM Certificate issued')
            if selected_type == 'document' or selected_type == '':
                total_doc_documents &= get_filter(
                    'doc_type', 'not_equal',
                    'BHD')
            else:
                total_doc_documents &= get_filter(
                    'doc_type', 'equal',
                    'BHD')
            if selected_org != '' and selected_org != 'All':
                total_doc_documents &= get_filter(
                    'sender', 'equal',
                    selected_org)

            total_approved_doc_documents = Q()
            total_approved_doc_documents &= get_filter(
                'status', 'not_equal',
                'QM Certificate issued')
            if selected_type == 'document' or selected_type == '':
                total_approved_doc_documents &= get_filter(
                    'doc_type', 'not_equal',
                    'BHD')
            else:
                total_approved_doc_documents &= get_filter(
                    'doc_type', 'equal',
                    'BHD')
            if selected_org != '' and selected_org != 'All':
                total_approved_doc_documents &= get_filter(
                    'sender', 'equal',
                    selected_org)

            doc_count = 0
            current_year_count = 0
            doc_approved = 0
            am_observation_forwarded = 0
            am_observation_repeated = 0
            audit_inProcess = 0
            over_due_doc = 0
            over_due_doc_current = 0
            qm_certification_issued = 0
            total_qm_repeated = 0
            total_qm_forwarded = 0
            total_qm_issued = 0
            total_auditinprocess = 0
            doc_not_approvedListQM = 0
            doc_not_approvedListNG = 0
            total_bhd_count = 0
            total_bhd_approved = 0
            total_bhd_qm_issued = 0
            over_due_bhd_total = 0
            over_due_bhd = 0
            ATP_count = 0
            total_doc_approvedList = doctracking.objects.filter(status='Approved')
            over_due_doc_total_all = doctracking.objects.filter(total_filter_document)
            count_all = 0
            if over_due_doc_total_all is not None:
                for item in over_due_doc_total_all:
                    if item.task_date is None:
                        item.task_date = datetime.today() + timedelta(hours=5)
                    if item.due_date.date() < item.task_date.date():
                        count_all = count_all + 1

            over_due_bhd_total_all = doctracking.objects.filter(total_type_bhd)
            count_all_bhds = 0
            if over_due_bhd_total_all is not None:
                for item in over_due_bhd_total_all:
                    if item.task_date is None:
                        item.task_date = datetime.today() + timedelta(hours=5)
                    if item.due_date.date() < item.task_date.date():
                        count_all_bhds = count_all_bhds + 1
            over_due_doc_total = count_all  # over_due_doc_total_all.filter(due_date__lt=F('task_date')).count()
            over_due_bhd_total = count_all_bhds  # doctracking.objects.filter(total_type_bhd, due_date__lt=F('task_date')).count()
            total_bhd_count_filter = Q()
            total_bhd_count_filter &= get_filter(
                'doc_type', 'equal',
                'BHD')
            if selected_org != '':
                total_bhd_count_filter &= get_filter(
                    'sender', 'equal',
                    selected_org)

            total_bhd_count = doctracking.objects.filter(total_bhd_count_filter).count()

            total_bhd_approved_filter = Q()
            total_bhd_approved_filter &= get_filter(
                'doc_type', 'equal',
                'BHD')
            if selected_org != '':
                total_bhd_approved_filter &= get_filter(
                    'sender', 'equal',
                    selected_org)
            if selected_type != '':
                total_bhd_approved_filter &= get_filter(
                    'status', 'equal',
                    'Approved')
            total_bhd_approved = doctracking.objects.filter(total_bhd_approved_filter).count()

            total_bhd_qm_issued_filter = Q()
            total_bhd_qm_issued_filter &= get_filter(
                'doc_type', 'equal',
                'BHD')
            if selected_org != '':
                total_bhd_qm_issued_filter &= get_filter(
                    'sender', 'equal',
                    selected_org)
            if selected_type == 'BHD':
                total_bhd_qm_issued_filter &= get_filter(
                    'status', 'equal',
                    'QM Certificate issued')

            total_bhd_qm_issued = doctracking.objects.filter(total_bhd_qm_issued_filter).count()
            total_bhd_underprocessQM = doctracking.objects.filter(total_filter_bhd, doc_type='BHD').count()
            total_bhd_underprocessNG = doctracking.objects.filter(total_filter_bhd_NG, doc_type='BHD').count()
            doc_count = doctracking.objects.filter(total_doc_documents).count()
            # doc_approved = doctracking.objects.filter(status='Approved').count()
            total_doc_approved = doctracking.objects.filter(total_approved_doc_documents, status='Approved').count()
            total_qm_repeated = doctracking.objects.filter(status='QM Observations Repeated').count()
            total_qm_forwarded = doctracking.objects.filter(status='QM Observations Forwarded').count()
            total_qm_issued = doctracking.objects.filter(status='QM Certificate issued').count()
            total_auditinprocess = doctracking.objects.filter(status='Audit in-process').count()
            doc_approvedList = doctracking.objects.filter(filter_objects, status='Approved')
            doc_not_approvedListQM = doctracking.objects.filter(total_filter_objects).count()
            doc_not_approvedListNG = doctracking.objects.filter(total_filter_objectsGN).count()
            over_due = doctracking.objects.filter(overdueDoc_filter_objects)
            current_over_document_filter = Q()
            current_over_document_filter &= get_filter(
                'status', 'equal',
                'Audit in-process')
            current_over_document_filter &= get_filter(
                'doc_type', 'not_equal',
                'BHD')

            count_current_doc = 0
            current_over_due_doc_all = over_due.filter(current_over_document_filter, receive_date__year=selected_year)
            if current_over_due_doc_all is not None:
                for item in current_over_due_doc_all:
                    if item.task_date is None:
                        item.task_date = datetime.today() + timedelta(hours=5)
                    if item.due_date.date() < item.task_date.date():
                        count_current_doc = count_current_doc + 1
            over_due_document = count_current_doc  # over_due.filter(current_over_document_filter, receive_date__year=selected_year,
            #                due_date__lt=F('task_date')).count()

            count_current_bhd = 0
            current_over_due_bhd_all = over_due.filter(overdueBHd_filter_objects, receive_date__year=selected_year,
                                                       doc_type='BHD')
            if current_over_due_bhd_all is not None:
                for item in current_over_due_bhd_all:
                    if item.task_date is None:
                        item.task_date = datetime.today() + timedelta(hours=5)
                    if item.due_date.date() < item.task_date.date():
                        count_current_bhd = count_current_bhd + 1
            over_due_bhd = count_current_bhd  # over_due.filter(overdueBHd_filter_objects, receive_date__year=selected_year,
            #               due_date__lt=F('task_date'),
            #               doc_type='BHD').count()

            current_filter_objects = Q()
            if selected_type == 'document' or selected_type == '':
                current_filter_objects &= get_filter(
                    'doc_type', 'not_equal',
                    'BHD')
                current_filter_objects &= get_filter(
                    'status', 'not_equal',
                    'QM Certificate issued')


            else:
                current_filter_objects &= get_filter(
                    'doc_type', 'equal',
                    'BHD')
                current_filter_objects &= get_filter(
                    'status', 'not_equal',
                    'Approved')
            if selected_year is '':
                # doc_approved = doctracking.objects.filter(status='Approved').count()
                # doc_approvedList = doctracking.objects.filter(status='Approved')
                # am_observation_repeated = doctracking.objects.filter(status='QM Observations Repeated').count()
                # am_observation_forwarded = doctracking.objects.filter(status='QM Observations Forwarded').count()

                # audit_inProcess = doctracking.objects.filter(status='Audit in-process').count()
                # qm_certification_issued = doctracking.objects.filter(status='QM Certificate issued').count()

                if selected_type is not '' and selected_org is '':
                    current_year_count = doctracking.objects.filter(filter_objects).count()
                    doc_approved = doctracking.objects.filter(filter_objects, status='Approved').count()
                    doc_approvedList = doctracking.objects.filter(filter_objects, status='Approved')
                    am_observation_forwarded = doctracking.objects.filter(filter_objects,
                                                                          status='QM Observations Forwarded').count()
                    am_observation_repeated = doctracking.objects.filter(filter_objects,
                                                                         status='QM Observations Repeated').count()
                    audit_inProcess = doctracking.objects.filter(filter_objects, status='Audit in-process').count()
                    qm_certification_issued = doctracking.objects.filter(filter_objects,
                                                                         status='QM Certificate issued').count()
                    et_type = doctracking.objects.filter(filter_objects, ).count()
                elif selected_type is '' and selected_org is not '':
                    current_year_count = doctracking.objects.filter(sender=selected_org).count()
                    doc_approved = doctracking.objects.filter(status='Approved', sender=selected_org).count()
                    doc_approvedList = doctracking.objects.filter(status='Approved', sender=selected_org)
                    am_observation_forwarded = doctracking.objects.filter(status='QM Observations Forwarded',
                                                                          sender=selected_org).count()
                    am_observation_repeated = doctracking.objects.filter(status='QM Observations Repeated',
                                                                         sender=selected_org).count()
                    audit_inProcess = doctracking.objects.filter(status='Audit in-process',
                                                                 sender=selected_org).count()
                    qm_certification_issued = doctracking.objects.filter(status='QM Certificate issued',
                                                                         sender=selected_org).count()
                elif selected_type is not '' and selected_org is not '':
                    current_year_count = doctracking.objects.filter(filter_objects, sender=selected_org).count()
                    doc_approved = doctracking.objects.filter(filter_objects, status='Approved',
                                                              sender=selected_org).count()
                    doc_approvedList = doctracking.objects.filter(filter_objects, status='Approved',
                                                                  sender=selected_org)
                    am_observation_forwarded = doctracking.objects.filter(filter_objects,
                                                                          status='QM Observations Forwarded').count()
                    am_observation_repeated = doctracking.objects.filter(filter_objects,
                                                                         status='QM Observations Repeated').count()
                    audit_inProcess = doctracking.objects.filter(filter_objects, status='Audit in-process').count()
                    qm_certification_issued = doctracking.objects.filter(filter_objects,
                                                                         status='QM Certificate issued').count()


            else:

                if selected_type is '' and selected_org is '':
                    current_year_count = doctracking.objects.filter(current_filter_objects,
                                                                    receive_date__year=selected_year).count()
                    doc_approved = doctracking.objects.filter(status='Approved',
                                                              receive_date__year=selected_year).count()
                    doc_approvedList = doctracking.objects.filter(current_filter_objects, status='Approved',
                                                                  receive_date__year=selected_year)
                    am_observation_forwarded = doctracking.objects.filter(current_filter_objects,
                                                                          status='QM Observations Forwarded',
                                                                          receive_date__year=selected_year).count()
                    am_observation_repeated = doctracking.objects.filter(current_filter_objects,
                                                                         status='QM Observations Repeated',
                                                                         receive_date__year=selected_year).count()
                    audit_inProcess = doctracking.objects.filter(current_filter_objects, status='Audit in-process',
                                                                 receive_date__year=selected_year).count()
                    qm_certification_issued = doctracking.objects.filter(current_filter_objects,
                                                                         status='QM Certificate issued',
                                                                         receive_date__year=selected_year).count()

                elif selected_org is not '' and selected_type is '':
                    if selected_org == 'All':
                        current_year_count = doctracking.objects.filter(current_filter_objects,
                                                                        receive_date__year=selected_year).count()
                        doc_approved = doctracking.objects.filter(current_filter_objects, status='Approved',
                                                                  receive_date__year=selected_year).count()
                        doc_approvedList = doctracking.objects.filter(status='Approved',
                                                                      receive_date__year=selected_year)
                        am_observation_forwarded = doctracking.objects.filter(current_filter_objects,
                                                                              status='QM Observation Forwarded',
                                                                              receive_date__year=selected_year).count()
                        am_observation_repeated = doctracking.objects.filter(current_filter_objects,
                                                                             status='QM Observations Repeated',
                                                                             receive_date__year=selected_year).count()
                        audit_inProcess = doctracking.objects.filter(current_filter_objects, status='Audit in-process',
                                                                     receive_date__year=selected_year).count()
                        qm_certification_issued = doctracking.objects.filter(current_filter_objects,
                                                                             status='QM Certificate issued',
                                                                             receive_date__year=selected_type).count()

                    else:
                        current_year_count = doctracking.objects.filter(current_filter_objects,
                                                                        receive_date__year=selected_year,
                                                                        sender=selected_org).count()
                        doc_approved = doctracking.objects.filter(status='Approved',
                                                                  receive_date__year=selected_year,
                                                                  sender=selected_org).count()
                        doc_approvedList = doctracking.objects.filter(current_filter_objects, status='Approved',
                                                                      receive_date__year=selected_year,
                                                                      sender=selected_org)
                        am_observation_forwarded = doctracking.objects.filter(current_filter_objects,
                                                                              status='QM Observations Forwarded',
                                                                              receive_date__year=selected_year,
                                                                              sender=selected_org).count()
                        am_observation_repeated = doctracking.objects.filter(current_filter_objects,
                                                                             status='QM Observations Repeated',
                                                                             receive_date__year=selected_year,
                                                                             sender=selected_org).count()
                        audit_inProcess = doctracking.objects.filter(current_filter_objects, status='Audit in-process',
                                                                     receive_date__year=selected_year,
                                                                     sender=selected_org).count()
                        qm_certification_issued = doctracking.objects.filter(current_filter_objects,
                                                                             status='QM Certificate issued',
                                                                             receive_date__year=selected_year,
                                                                             sender=selected_org).count()

                elif selected_org is '' and selected_type is not '':
                    if selected_org == 'All':
                        current_year_count = doctracking.objects.filter(current_filter_objects, filter_objects,
                                                                        receive_date__year=selected_year).count()
                        doc_approved = doctracking.objects.filter(filter_objects,
                                                                  status='Approved',
                                                                  receive_date__year=selected_year).count()
                        doc_approvedList = doctracking.objects.filter(current_filter_objects, filter_objects,
                                                                      status='Approved',
                                                                      receive_date__year=selected_year)
                        am_observation_forwarded = doctracking.objects.filter(current_filter_objects, filter_objects,
                                                                              status='QM Observations Forwarded',
                                                                              receive_date__year=selected_year).count()
                        am_observation_repeated = doctracking.objects.filter(filter_objects,
                                                                             status='QM Observations Repeated',
                                                                             receive_date__year=selected_year).count()
                        audit_inProcess = doctracking.objects.filter(current_filter_objects, filter_objects,
                                                                     status='Audit in-process',
                                                                     receive_date__year=selected_year).count()
                        qm_certification_issued = doctracking.objects.filter(current_filter_objects, filter_objects,
                                                                             status='QM Certificate issued',
                                                                             receive_date__year=selected_year).count()

                    else:
                        current_year_count = doctracking.objects.filter(current_filter_objects, filter_objects,
                                                                        receive_date__year=selected_year).count()
                        doc_approved = doctracking.objects.filter(filter_objects,
                                                                  status='Approved',
                                                                  receive_date__year=selected_year).count()
                        doc_approvedList = doctracking.objects.filter(current_filter_objects, filter_objects,
                                                                      status='Approved',
                                                                      receive_date__year=selected_year)
                        am_observation_forwarded = doctracking.objects.filter(current_filter_objects, filter_objects,
                                                                              status='QM Observations Forwarded',
                                                                              receive_date__year=selected_year).count()
                        am_observation_repeated = doctracking.objects.filter(current_filter_objects, filter_objects,
                                                                             status='QM Observations Repeated',
                                                                             receive_date__year=selected_year).count()
                        audit_inProcess = doctracking.objects.filter(current_filter_objects, filter_objects,
                                                                     status='Audit in-process',
                                                                     receive_date__year=selected_year).count()
                        qm_certification_issued = doctracking.objects.filter(current_filter_objects, filter_objects,
                                                                             status='QM Certificate issued',
                                                                             receive_date__year=selected_year).count()


                elif selected_type is not '' and selected_org is not '':
                    if selected_org == 'All':
                        current_year_count = doctracking.objects.filter(current_filter_objects, filter_objects,
                                                                        receive_date__year=selected_year).count()
                        doc_approved = doctracking.objects.filter(filter_objects, status='Approved',
                                                                  receive_date__year=selected_year).count()
                        doc_approvedList = doctracking.objects.filter(current_filter_objects, filter_objects,
                                                                      status='Approved',
                                                                      receive_date__year=selected_year)
                        am_observation_forwarded = doctracking.objects.filter(current_filter_objects, filter_objects,
                                                                              status='QM Observations Forwarded',
                                                                              receive_date__year=selected_year).count()
                        am_observation_repeated = doctracking.objects.filter(current_filter_objects, filter_objects,
                                                                             status='QM Observations Repeated',
                                                                             receive_date__year=selected_year).count()
                        audit_inProcess = doctracking.objects.filter(current_filter_objects, filter_objects,
                                                                     status='Audit in-process',
                                                                     receive_date__year=selected_year).count()

                        qm_certification_issued = doctracking.objects.filter(current_filter_objects, filter_objects,
                                                                             status='QM Certificate issued',
                                                                             receive_date__year=selected_year).count()
                    else:
                        current_year_count = doctracking.objects.filter(current_filter_objects, filter_objects,
                                                                        receive_date__year=selected_year,
                                                                        sender=selected_org).count()
                        doc_approved = doctracking.objects.filter(filter_objects, status='Approved',
                                                                  receive_date__year=selected_year,
                                                                  sender=selected_org).count()
                        doc_approvedList = doctracking.objects.filter(current_filter_objects, filter_objects,
                                                                      status='Approved',
                                                                      receive_date__year=selected_year,
                                                                      sender=selected_org)
                        am_observation_forwarded = doctracking.objects.filter(current_filter_objects, filter_objects,
                                                                              status='QM Observations Forwarded',
                                                                              receive_date__year=selected_year,
                                                                              sender=selected_org).count()
                        am_observation_repeated = doctracking.objects.filter(current_filter_objects, filter_objects,
                                                                             status='QM Observations Repeated',
                                                                             receive_date__year=selected_year,
                                                                             sender=selected_org).count()
                        audit_inProcess = doctracking.objects.filter(current_filter_objects, filter_objects,
                                                                     status='Audit in-process',
                                                                     receive_date__year=selected_year,
                                                                     sender=selected_org).count()
                        qm_certification_issued = doctracking.objects.filter(current_filter_objects, filter_objects,
                                                                             status='QM Certificate issued',
                                                                             receive_date__year=selected_year,
                                                                             sender=selected_org).count()

            doc_approved = doc_approvedList.filter(current_filter_objects).count()
            QAC_count = doc_approvedList.filter(doc_type='Qualification & Acceptance Criteria').count()
            QFTP_count = doc_approvedList.filter(doc_type='QFTP').count()
            TDP_count = doc_approvedList.filter(doc_type='TDP').count()
            SOP_count = doc_approvedList.filter(doc_type='SOP').count()
            Guidelines_count = doc_approvedList.filter(doc_type='Guidelines').count()
            others_count = doc_approvedList.filter(doc_type='Others/Misc').count()
            ATP_count = doc_approvedList.filter(doc_type='ATP').count()
            # Total document
            total_QAC_count = total_doc_approvedList.filter(doc_type='Qualification & Acceptance Criteria').count()
            total_QFTP_count = total_doc_approvedList.filter(doc_type='QFTP').count()
            total_TDP_count = total_doc_approvedList.filter(doc_type='TDP').count()
            total_SOP_count = total_doc_approvedList.filter(doc_type='SOP').count()
            total_Guidelines_count = total_doc_approvedList.filter(doc_type='Guidelines').count()
            total_others_count = total_doc_approvedList.filter(doc_type='Others/Misc').count()

            # SST_count = doc_approvedList.filter(doc_type='Structural Strength Testing(SST)').count()
            # TDP_count = doc_approvedList.filter(doc_type='Technical Data Pack (TDP)').count()
            dist = {
                'total_bhd_underprocessQM': total_bhd_underprocessQM,
                'total_bhd_underprocessNG': total_bhd_underprocessNG,
                'total_bhd_approved': total_bhd_approved,
                'total_bhd_qm_issued': total_bhd_qm_issued,
                'doc_count': doc_count,
                'current_year_count': current_year_count,
                'current_doc_approved': doc_approved,
                'am_observation_forwarded': am_observation_forwarded,
                'am_observation_repeated': am_observation_repeated,
                'audit_inProcess': audit_inProcess,
                'QFTP_count': QFTP_count,
                'TDP_count': TDP_count,
                'SOP_count': SOP_count,
                'Guidelines_count': Guidelines_count,
                'others_count': others_count,
                'ATP_count': ATP_count,
                'QAC_count': QAC_count,
                # 'SST_count': SST_count,
                # 'TDP_count': TDP_count,
                'over_due_count': over_due_doc_total,
                'qm_certification_issued': qm_certification_issued,
                'total_doc': doc_count,
                'total_doc_approved': total_doc_approved,
                'total_QAC_count': total_QAC_count,
                'total_QFTP_count': total_QFTP_count,
                'total_TDP_count': total_TDP_count,
                'total_SOP_count': total_SOP_count,
                'total_Guidelines_count': total_Guidelines_count,
                'total_others_count': total_others_count,
                'total_qm_repeated': total_qm_repeated,
                'total_qm_forwarded': total_qm_forwarded,
                'total_qm_issued': total_qm_issued,
                'total_auditinprocess': total_auditinprocess,
                'doc_not_approvedListQM': doc_not_approvedListQM,
                'doc_not_approvedListNG': doc_not_approvedListNG,
                'over_due_document': over_due_document,
                'total_bhd_count': total_bhd_count,
                'over_due_bhd_total': over_due_bhd_total,
                'over_due_bhd': over_due_bhd

            }

            # DataCount.append(dist)
            return JsonResponse({'status': 'True', 'data': dist},
                                status=200)

        except Exception as e:
            print(e)
            return JsonResponse({'status': 'False', "message": "Internal Server Error"}, status=500)
            pass

    @staticmethod
    def getTaskMonitoringDashboardCount(request):
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

            selected_year = request.query_params.get('selected_year')
            selected_group = request.query_params.get('selected_group')
            task_completed = 0
            task_inprocess = 0
            task_follow_up = 0
            total_task_completed = 0
            total_task_inprocess = 0
            total_task_follow_up = 0
            total_tasks = 0
            doc_not_completedList = 0
            total_over_due = 0
            current_year = datetime.today().year

            total_filter_objects = Q()
            total_filter_objects &= get_filter(
                'status', 'not_equal',
                'Task Completed')
            if selected_group != '':
                total_filter_objects &= get_filter(
                    'assigned_to', 'equal',
                    selected_group)

            all_filter_objects = Q()
            if selected_group != '':
                all_filter_objects &= get_filter(
                    'assigned_to', 'equal',
                    selected_group)
            current_over_due_filter = Q()
            current_over_due_filter &= get_filter(
                'status', 'equal',
                'Task in-process')
            current_year_task = TaskSummary.objects.filter(all_filter_objects,
                                                           assigned_date__year=selected_year).count()
            doc_not_completedList = TaskSummary.objects.filter(total_filter_objects).count()
            total_task_inprocess = TaskSummary.objects.filter(all_filter_objects, status='Task in-process').count()
            total_task_follow_up = TaskSummary.objects.filter(all_filter_objects, status='Task follow-up').count()
            total_over_due = TaskSummary.objects.filter(current_over_due_filter, task_date__gt=F('target_date')).count()
            current_over_due = TaskSummary.objects.filter(current_over_due_filter, task_date__gt=F('target_date'),
                                                          assigned_date__year=selected_year).count()

            total_tasks = TaskSummary.objects.filter(all_filter_objects).count()
            total_task_completed = TaskSummary.objects.filter(all_filter_objects, status='Task Completed').count()
            if selected_year is not '' and selected_group is '':
                task_completed = TaskSummary.objects.filter(assigned_date__year=selected_year,
                                                            status='Task Completed').count()
                task_inprocess = TaskSummary.objects.filter(assigned_date__year=selected_year,
                                                            status='Task in-process').count()
                task_follow_up = TaskSummary.objects.filter(assigned_date__year=selected_year,
                                                            status='Task follow-up').count()
            elif selected_year is '' and selected_group is not '':
                task_completed = TaskSummary.objects.filter(assigned_to=selected_group,
                                                            status='Task Completed').count()
                task_inprocess = TaskSummary.objects.filter(assigned_to=selected_group,
                                                            status='Task in-process').count()
                task_follow_up = TaskSummary.objects.filter(assigned_to=selected_group,
                                                            status='Task follow-up').count()
            elif selected_year is not '' and selected_group is not '':
                task_completed = TaskSummary.objects.filter(assigned_to=selected_group,
                                                            assigned_date__year=selected_year,
                                                            status='Task Completed').count()
                task_inprocess = TaskSummary.objects.filter(assigned_to=selected_group,
                                                            assigned_date__year=selected_year,
                                                            status='Task in-process').count()
                task_follow_up = TaskSummary.objects.filter(assigned_to=selected_group,
                                                            assigned_date__year=selected_year,
                                                            status='Task follow-up').count()
            dist = {
                'task_completed': task_completed,
                'task_inprocess': task_inprocess,
                'task_follow_up': task_follow_up,
                'total_over_due': total_over_due,
                'total_tasks': total_tasks,
                'total_current_year': current_year_task,
                'total_task_completed': total_task_completed,
                'total_task_inprocess': total_task_inprocess,
                'total_task_follow_up': total_task_follow_up,
                'doc_not_completedList': doc_not_completedList,
                'current_over_due': current_over_due
            }

            # DataCount.append(dist)
            return JsonResponse({'status': 'True', 'data': dist},
                                status=200)

        except Exception as e:
            print(e)
            return JsonResponse({'status': 'False', "message": "Internal Server Error"}, status=500)
            pass
