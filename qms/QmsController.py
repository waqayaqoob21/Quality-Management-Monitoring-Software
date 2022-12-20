from django.http import JsonResponse
from fpdf import FPDF
import xlwt
from datetime import datetime, timedelta
from django.http import FileResponse
from django.http import HttpResponse
from django.core.files.storage import FileSystemStorage
from django.db.models import F, Q
import itertools
from qms.models import *
from qms.serializer import *
from datetime import date

class QmsController:
    @staticmethod
    def AddQmsAudit(request):
        qmsModel = QmsAudit()
        try:
            id = request['id']
            if id == '0':
                qmsModel.audit_id = 0  # request['audit_id']
                qmsModel.Organization = request['Organization']
                qmsModel.site = request['site']
                qmsModel.setup = request['setup']
                qmsModel.certification_status = request['certification_status']
                qmsModel.previous_standard = request['previous_standard']
                qmsModel.certification_validity_date = request['certification_validity_date']
                if request['certification_validity_rescheduling_date'] != '':
                    qmsModel.certification_validity_rescheduling_date = request['certification_validity_rescheduling_date']
                qmsModel.audit_type = request['audit_type']
                qmsModel.planned_date = request['planned_date']
                qmsModel.audit_start_date = request['audit_start_date']
                qmsModel.audit_close_date = request['audit_close_date']
                qmsModel.audit_status = request['audit_status']
                qmsModel.standard = request['standard']
                qmsModel.remarks = request['remarks']
                qmsModel.certification_setup = request['certification_setup']
                qmsModel.save()
                return JsonResponse({'status': 'True', 'message': "QMS Audit Created Successfully!"},
                                    status=200)
            else:
                get_qms = QmsAudit.objects.filter(id=id).first()
                if get_qms is not None:
                    if get_qms.audit_status != request['audit_status'] or str(get_qms.certification_validity_date.date()) != \
                            request['certification_validity_date'] or str(get_qms.certification_validity_rescheduling_date.date()) != \
                            request['certification_validity_rescheduling_date'] or str(get_qms.planned_date.date()) != \
                            request['planned_date'] or  str(get_qms.audit_start_date.date()) != \
                            request['audit_start_date'] or str(get_qms.audit_close_date.date()) != request['audit_close_date']:

                        auditHistory = QmsAuditHistory()
                        auditHistory.audit_id = get_qms.audit_id
                        auditHistory.Organization = get_qms.Organization
                        auditHistory.site = get_qms.site
                        auditHistory.setup = get_qms.setup
                        auditHistory.certification_status = get_qms.certification_status
                        auditHistory.previous_standard = get_qms.previous_standard
                        auditHistory.certification_validity_date = get_qms.certification_validity_date
                        if request['certification_validity_rescheduling_date'] != '':
                            auditHistory.certification_validity_rescheduling_date = get_qms.certification_validity_rescheduling_date
                        auditHistory.audit_type = get_qms.audit_type
                        auditHistory.planned_date = get_qms.planned_date
                        auditHistory.audit_start_date = get_qms.audit_start_date
                        auditHistory.audit_close_date = get_qms.audit_close_date
                        auditHistory.audit_status = get_qms.audit_status
                        auditHistory.standard = get_qms.standard
                        auditHistory.remarks = get_qms.remarks
                        auditHistory.certification_setup = get_qms.certification_setup
                        auditHistory.qms_audit_id = id
                        auditHistory.save()

                get_qms.audit_id = 0  # request['audit_id']
                get_qms.Organization = request['Organization']
                get_qms.site = request['site']
                get_qms.setup = request['setup']
                get_qms.certification_status = request['certification_status']
                get_qms.previous_standard = request['previous_standard']
                get_qms.certification_validity_date = request['certification_validity_date']
                if request['certification_validity_rescheduling_date'] != '':
                    get_qms.certification_validity_rescheduling_date = request['certification_validity_rescheduling_date']
                get_qms.audit_type = request['audit_type']
                get_qms.planned_date = request['planned_date']
                get_qms.audit_start_date = request['audit_start_date']
                get_qms.audit_close_date = request['audit_close_date']
                get_qms.audit_status = request['audit_status']
                get_qms.standard = request['standard']
                get_qms.remarks = request['remarks']
                get_qms.certification_setup = request['certification_setup']
                get_qms.save()
                return JsonResponse({'status': 'True', 'message': "QMS Audit Updated Successfully!"},
                                    status=200)
        except Exception as e:
            return JsonResponse({'status': 'False', "message": "QMS Audit Not Saved"}, status=500)

    @staticmethod
    def GetQmsAuditList(request, self=None):
        try:
            current_status = request.query_params.get('audit_status')
            current_year = request.query_params.get('year')
            current_org = request.query_params.get('organization')
            current_stand = request.query_params.getlist('standard')
            current_setup = request.query_params.getlist('setup')
            today = date.today()
            standItems = ""
            setupItems = ""
            noVal = ['']
            undefined = ['undefined']

            if current_stand != noVal and current_stand != undefined:
                for item in current_stand:
                    standItems = item.split(',')
            if current_setup != noVal and current_setup != undefined:
                for item in current_setup:
                    setupItems = item.split(',')

            dataList = []
            def get_filter(field_name, filter_condition, filter_value):
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
                        '{0}__ne'.format(field_name): filter_value
                    }
                    return ~Q(**kwargs)
                if filter_condition.strip() == "less_than":
                    kwargs = {
                        '{0}__lt'.format(field_name): filter_value
                    }
                    return ~Q(**kwargs)
                if filter_condition.strip() == "greater_than":
                    kwargs = {
                        '{0}__gt'.format(field_name): filter_value
                    }
                    return ~Q(**kwargs)
            typeQuery = Q()
            filter_objects = Q()

            if current_year != '' :
                filter_objects &= get_filter(
                    'planned_date__year', 'equal',
                    current_year)

            if current_status != ''and current_status != 'Overdue' and current_status != 'Under Process':
                filter_objects &= get_filter(
                    'audit_status', 'equal',
                    current_status)

            # if current_status == 'Overdue':
            #     filter_objects &= get_filter(
            #     'certification_validity_date', 'greater_than',today)

            if current_status == 'Under Process':
                filter_objects &= get_filter(
                'audit_status', 'not_equal','Completed')

            if current_org != '':
                filter_objects &= get_filter(
                    'Organization', 'equal',
                    current_org)

            # if len(standItems) >0:
            #     for item in standItems:
            #         filter_objects &= get_filter(
            #             'standard', 'contains',
            #             item)

            # if len(standItems) >0:
            #     for item in standItems:
            #         filter_objects &= get_filter(
            #             'standard', 'greater_than',
            #             item)


            dataList = QmsAudit.objects.filter(filter_objects)
            if current_status == 'Overdue':
                dataList = dataList.filter(~Q(audit_status = 'Completed'), certification_validity_date__lt = today)

            if current_status == 'Total Audits':
                data = QmsAudit.objects.all().order_by('-id')
                serializer = QmsAuditSerializer(data, many=True)
                return JsonResponse({'status': 'True', 'data': serializer.data},
                                    status=200)

            elif len(standItems) > 0 and len(setupItems)>0:
                temp =[]
                for (set,stand) in itertools.zip_longest (setupItems,standItems):
                    temp.extend(dataList.filter(standard= stand,setup = set))
                dataList = temp
            elif len(setupItems)>0 and len(standItems) == 0:
                temp =[]
                for set in setupItems:
                    temp.extend(dataList.filter(setup = set))
                dataList = temp
            elif len(setupItems)==0 and len(standItems) >0:
                temp =[]
                for stand in standItems:
                    temp.extend(dataList.filter(standard=stand))
                dataList = temp


            serializer = QmsAuditSerializer(dataList, many=True)
            return JsonResponse({'status': 'True', 'data': serializer.data},
                                status=200)
        except Exception as e:
            print(e)
            return JsonResponse({'status': 'False', "message": "Internal Server Error"}, status=500)

    @staticmethod
    def GetQmsAuditListCount(request, self=None):
        try:
            selected_year = request.query_params.get('selected_year')
            selected_org = request.query_params.get('selected_organization')
            selected_setup = request.query_params.getlist('selected_setup')
            selected_standard = request.query_params.getlist('selected_standard')
            today= date.today()
            standItems = ""
            setupItems = ""
            noVal = ['']
            if selected_standard != noVal:
                for item in selected_standard:
                    standItems = item.split(',')
            if selected_setup != noVal:
                for item in selected_setup:
                    setupItems = item.split(',')
            total_audits = 0
            current_year_audits = 0
            audit_in_process = 0
            audit_completed = 0
            inprocess_overdue = 0
            total_audit_schedule = 0
            audit_remaining = 0
            audit_under_process = 0
            currYearTotalAudits = 0
            currYearAuditInprocess = 0
            currYearAuditCompleted = 0
            currYearOverdue = 0
            currYearAuditScheduled = 0
            currYearAuditRemaining = 0

            total_audits = QmsAudit.objects.count()

            if selected_year == '':
                if selected_org == '' and selected_standard == noVal and selected_setup == noVal:
                    audit_completed = QmsAudit.objects.filter(audit_status='Completed').count()
                    audit_under_process = QmsAudit.objects.filter(~Q(audit_status='Completed')).count()
                    inprocess_overdue = QmsAudit.objects.filter(~Q(audit_status='Completed'), certification_validity_date__lt=today).count()


                elif selected_org == '' and selected_standard == noVal and selected_setup != noVal:
                    for set in setupItems:
                        tot_audit_under_process = QmsAudit.objects.filter(~Q(audit_status='Completed'),setup=set).count()
                        tot_audit_completed = QmsAudit.objects.filter(setup=set, audit_status='Completed').count()
                        tot_inprocess_overdue = QmsAudit.objects.filter(~Q(audit_status='Completed'),setup=set,
                                                                        certification_validity_date__lt=today).count()
                        audit_under_process += tot_audit_under_process
                        audit_completed += tot_audit_completed
                        inprocess_overdue += tot_inprocess_overdue

                elif selected_org == '' and selected_standard != noVal and selected_setup == noVal:
                    for stand in standItems:
                        tot_audit_under_process = QmsAudit.objects.filter(~Q(audit_status='Completed'),
                                                                          standard=stand).count()
                        tot_audit_completed = QmsAudit.objects.filter(standard=stand,
                                                                      audit_status='Completed').count()
                        tot_inprocess_overdue = QmsAudit.objects.filter(~Q(audit_status='Completed'),standard=stand,
                                                                        certification_validity_date__lt=today).count()
                        audit_under_process += tot_audit_under_process
                        audit_completed += tot_audit_completed
                        inprocess_overdue += tot_inprocess_overdue

                elif selected_org == '' and selected_standard != noVal and selected_setup != noVal:
                    for (stand, set) in itertools.zip_longest(standItems, setupItems):
                        tot_audit_under_process = QmsAudit.objects.filter(~Q(audit_status='Completed'),
                                                                          standard=stand, setup=set).count()
                        tot_audit_completed = QmsAudit.objects.filter(standard=stand, setup=set,
                                                                      audit_status='Completed').count()
                        tot_inprocess_overdue = QmsAudit.objects.filter(~Q(audit_status='Completed'),standard=stand,
                                                                        setup=set,certification_validity_date__lt=today).count()
                        audit_under_process += tot_audit_under_process
                        audit_completed += tot_audit_completed
                        inprocess_overdue += tot_inprocess_overdue

                elif selected_org != '' and selected_standard == noVal and selected_setup == noVal:
                    audit_under_process = QmsAudit.objects.filter(~Q(audit_status='Completed'),
                                                                  Organization=selected_org).count()
                    audit_completed = QmsAudit.objects.filter(Organization=selected_org,
                                                              audit_status='Completed').count()
                    inprocess_overdue = QmsAudit.objects.filter(~Q(audit_status='Completed'),Organization=selected_org,
                                                                certification_validity_date__lt=today).count()

                elif selected_org != '' and selected_standard == noVal and selected_setup != noVal:
                    for set in setupItems:
                        tot_audit_under_process = QmsAudit.objects.filter(~Q(audit_status='Completed'),Organization=selected_org,
                                                                          setup=set).count()
                        tot_audit_completed = QmsAudit.objects.filter(Organization=selected_org, setup=set,
                                                                      audit_status='Completed').count()
                        tot_inprocess_overdue = QmsAudit.objects.filter(~Q(audit_status='Completed'),Organization=selected_org, setup=set,
                                                                        certification_validity_date__lt=today).count()
                        audit_under_process += tot_audit_under_process
                        audit_completed += tot_audit_completed
                        inprocess_overdue += tot_inprocess_overdue

                elif selected_org != '' and selected_standard != noVal and selected_setup == noVal:
                    for stand in standItems:
                        tot_audit_under_process = QmsAudit.objects.filter(~Q(audit_status='Completed'),Organization=selected_org,
                                                                          standard=stand).count()
                        tot_audit_completed = QmsAudit.objects.filter(Organization=selected_org, standard=stand,
                                                                      audit_status='Completed').count()
                        tot_inprocess_overdue = QmsAudit.objects.filter(~Q(audit_status='Completed'),Organization=selected_org, standard=stand,
                                                                        certification_validity_date__lt=today).count()

                        audit_under_process += tot_audit_under_process
                        audit_completed += tot_audit_completed
                        inprocess_overdue += tot_inprocess_overdue

                else:
                    for (stand, set) in itertools.zip_longest(standItems, setupItems):
                        tot_audit_under_process = QmsAudit.objects.filter(~Q(audit_status='Completed'),Organization=selected_org,
                                                                          standard=stand,setup=set).count()
                        tot_audit_completed = QmsAudit.objects.filter(Organization=selected_org, standard=stand,
                                                                      setup=set, audit_status='Completed').count()
                        tot_inprocess_overdue = QmsAudit.objects.filter(~Q(audit_status='Completed'),Organization=selected_org, standard=stand,
                                                                        setup=set,certification_validity_date__lt=today).count()
                        audit_under_process += tot_audit_under_process
                        audit_completed += tot_audit_completed
                        inprocess_overdue += tot_inprocess_overdue

            else:
                if selected_org == '' and selected_standard == noVal and selected_setup ==noVal:
                    currYearTotalAudits = QmsAudit.objects.filter(planned_date__year=selected_year).count()
                    currYearAuditInprocess = QmsAudit.objects.filter(planned_date__year=selected_year,
                                                               audit_status='In process').count()
                    currYearAuditScheduled = QmsAudit.objects.filter(planned_date__year=selected_year,
                                                                   audit_status='Scheduled').count()
                    currYearAuditCompleted = QmsAudit.objects.filter(planned_date__year=selected_year,
                                                              audit_status='Completed').count()
                    currYearAuditRemaining = QmsAudit.objects.filter(planned_date__year=selected_year,
                                                              audit_status='Remaining').count()
                    currYearOverdue = QmsAudit.objects.filter(~Q(audit_status='Completed'),planned_date__year=selected_year
                                                                ,certification_validity_date__lt = today).count()

                elif selected_org == '' and selected_standard == noVal and selected_setup != noVal:
                    for set in setupItems:
                        tot_current_year_audits = QmsAudit.objects.filter(planned_date__year=selected_year,
                                                                      setup=set).count()
                        tot_audit_in_process = QmsAudit.objects.filter(planned_date__year=selected_year,
                                                                   setup=set, audit_status='In process').count()
                        tot_audit_schedule = QmsAudit.objects.filter(planned_date__year=selected_year,
                                                                       setup=set,
                                                                       audit_status='Scheduled').count()
                        tot_audit_completed = QmsAudit.objects.filter(planned_date__year=selected_year,
                                                                  setup=set, audit_status='Completed').count()
                        tot_audit_remaining = QmsAudit.objects.filter(planned_date__year=selected_year,
                                                                  setup=set, audit_status='Remaining').count()
                        tot_inprocess_overdue = QmsAudit.objects.filter(~Q(audit_status='Completed'),planned_date__year=selected_year,
                                                                    setup=set,certification_validity_date__lt = today).count()
                        currYearTotalAudits += tot_current_year_audits
                        currYearAuditInprocess += tot_audit_in_process
                        currYearAuditScheduled += tot_audit_schedule
                        currYearAuditCompleted += tot_audit_completed
                        currYearAuditRemaining += tot_audit_remaining
                        currYearOverdue += tot_inprocess_overdue

                elif selected_org == '' and selected_standard != noVal and selected_setup == noVal:
                    for stand in standItems:
                        tot_current_year_audits = QmsAudit.objects.filter(planned_date__year=selected_year,
                                                                      standard=stand).count()
                        tot_audit_in_process = QmsAudit.objects.filter(planned_date__year=selected_year,
                                                                   standard=stand,
                                                                   audit_status='In process').count()
                        tot_audit_schedule = QmsAudit.objects.filter(planned_date__year=selected_year,
                                                                       standard=stand,
                                                                       audit_status='Scheduled').count()
                        tot_audit_completed = QmsAudit.objects.filter(planned_date__year=selected_year,
                                                                  standard=stand,
                                                                  audit_status='Completed').count()
                        tot_audit_remaining = QmsAudit.objects.filter(planned_date__year=selected_year,
                                                                  standard=stand,
                                                                  audit_status='Remaining').count()
                        tot_inprocess_overdue = QmsAudit.objects.filter(~Q(audit_status='Completed'),planned_date__year=selected_year,
                                                                    standard=stand,
                                                                    certification_validity_date__lt = today).count()
                        currYearTotalAudits += tot_current_year_audits
                        currYearAuditInprocess += tot_audit_in_process
                        currYearAuditScheduled += tot_audit_schedule
                        currYearAuditCompleted += tot_audit_completed
                        currYearAuditRemaining += tot_audit_remaining
                        currYearOverdue += tot_inprocess_overdue

                elif selected_org == '' and selected_standard != noVal and selected_setup != noVal:
                    for (stand,set) in itertools.zip_longest(standItems,setupItems):
                        tot_current_year_audits = QmsAudit.objects.filter(planned_date__year=selected_year,
                                                                      standard=stand,
                                                                      setup=set).count()
                        tot_audit_in_process = QmsAudit.objects.filter(planned_date__year=selected_year,
                                                                   standard=stand, setup=set,
                                                                   audit_status='In process').count()
                        tot_audit_schedule = QmsAudit.objects.filter(planned_date__year=selected_year,
                                                                       standard=stand, setup=set,
                                                                       audit_status='Scheduled').count()
                        tot_audit_completed = QmsAudit.objects.filter(planned_date__year=selected_year,
                                                                  standard=stand, setup=set,
                                                                  audit_status='Completed').count()
                        tot_audit_remaining = QmsAudit.objects.filter(planned_date__year=selected_year,
                                                                  standard=stand, setup=set,
                                                                  audit_status='Remaining').count()
                        tot_inprocess_overdue = QmsAudit.objects.filter(~Q(audit_status='Completed'),planned_date__year=selected_year,
                                                                    standard=stand, setup=set,
                                                                    certification_validity_date__lt = today).count()
                        currYearTotalAudits += tot_current_year_audits
                        currYearAuditInprocess += tot_audit_in_process
                        currYearAuditScheduled += tot_audit_schedule
                        currYearAuditCompleted += tot_audit_completed
                        currYearAuditRemaining += tot_audit_remaining
                        currYearOverdue += tot_inprocess_overdue

                elif selected_org != '' and selected_standard == noVal and selected_setup == noVal:
                    currYearTotalAudits = QmsAudit.objects.filter(planned_date__year=selected_year,
                                                                  Organization=selected_org, ).count()
                    currYearAuditInprocess = QmsAudit.objects.filter(planned_date__year=selected_year,
                                                               Organization=selected_org,
                                                               audit_status='In process').count()
                    currYearAuditScheduled = QmsAudit.objects.filter(planned_date__year=selected_year,
                                                                   Organization=selected_org,
                                                                   audit_status='Scheduled').count()
                    currYearAuditCompleted = QmsAudit.objects.filter(planned_date__year=selected_year,
                                                              Organization=selected_org,
                                                              audit_status='Completed').count()
                    currYearAuditRemaining = QmsAudit.objects.filter(planned_date__year=selected_year,
                                                              Organization=selected_org,
                                                              audit_status='Remaining').count()
                    currYearOverdue = QmsAudit.objects.filter(~Q(audit_status='Completed'),planned_date__year=selected_year,
                                                                Organization=selected_org,
                                                                certification_validity_date__lt = today).count()

                elif selected_org != '' and selected_standard == noVal and selected_setup != noVal:
                    for set in setupItems:
                        tot_current_year_audits = QmsAudit.objects.filter(planned_date__year=selected_year,
                                                                      Organization=selected_org,
                                                                      setup=set).count()
                        tot_audit_in_process = QmsAudit.objects.filter(planned_date__year=selected_year,
                                                                   Organization=selected_org, setup=set,
                                                                   audit_status='In process').count()
                        tot_audit_schedule = QmsAudit.objects.filter(planned_date__year=selected_year,
                                                                       Organization=selected_org, setup=set,
                                                                       audit_status='Scheduled').count()
                        tot_audit_completed = QmsAudit.objects.filter(planned_date__year=selected_year,
                                                                  Organization=selected_org, setup=set,
                                                                  audit_status='Completed').count()
                        tot_audit_remaining = QmsAudit.objects.filter(planned_date__year=selected_year,
                                                                  Organization=selected_org, setup=set,
                                                                  audit_status='Remaining').count()
                        tot_inprocess_overdue = QmsAudit.objects.filter(~Q(audit_status='Completed'),planned_date__year=selected_year,
                                                                    Organization=selected_org, setup=set,
                                                                    certification_validity_date__lt = today).count()
                        currYearTotalAudits += tot_current_year_audits
                        currYearAuditInprocess += tot_audit_in_process
                        currYearAuditScheduled += tot_audit_schedule
                        currYearAuditCompleted += tot_audit_completed
                        currYearAuditRemaining += tot_audit_remaining
                        currYearOverdue += tot_inprocess_overdue

                elif selected_org != '' and selected_standard != noVal and selected_setup == noVal:
                    for stand in standItems:
                        tot_current_year_audits = QmsAudit.objects.filter(planned_date__year=selected_year,
                                                                      Organization=selected_org,
                                                                      standard=stand).count()
                        tot_audit_in_process = QmsAudit.objects.filter(planned_date__year=selected_year,
                                                                   Organization=selected_org, standard=stand,
                                                                   audit_status='In process').count()
                        tot_audit_schedule = QmsAudit.objects.filter(planned_date__year=selected_year,
                                                                       Organization=selected_org,
                                                                       standard=stand,
                                                                       audit_status='Scheduled').count()
                        tot_audit_completed = QmsAudit.objects.filter(planned_date__year=selected_year,
                                                                  Organization=selected_org, standard=stand,
                                                                  audit_status='Completed').count()
                        tot_audit_remaining = QmsAudit.objects.filter(planned_date__year=selected_year,
                                                                  Organization=selected_org, standard=stand,
                                                                  audit_status='Remaining').count()
                        tot_inprocess_overdue = QmsAudit.objects.filter(~Q(audit_status='Completed'),planned_date__year=selected_year,
                                                                    Organization=selected_org, standard=stand,
                                                                    certification_validity_date__lt = today).count()
                        currYearTotalAudits += tot_current_year_audits
                        currYearAuditInprocess += tot_audit_in_process
                        currYearAuditScheduled += tot_audit_schedule
                        currYearAuditCompleted += tot_audit_completed
                        currYearAuditRemaining += tot_audit_remaining
                        currYearOverdue += tot_inprocess_overdue

                else:
                    for (stand, set) in itertools.zip_longest(standItems, setupItems):
                        tot_current_year_audits = QmsAudit.objects.filter(planned_date__year=selected_year,
                                                                      Organization=selected_org,
                                                                      standard=stand,
                                                                      setup=set).count()
                        tot_audit_in_process = QmsAudit.objects.filter(planned_date__year=selected_year,
                                                                   Organization=selected_org, standard=stand,
                                                                   setup=set, audit_status='In process').count()
                        tot_audit_schedule = QmsAudit.objects.filter(planned_date__year=selected_year,
                                                                       Organization=selected_org,
                                                                       standard=stand,
                                                                       setup=set,
                                                                       audit_status='Scheduled').count()
                        tot_audit_completed = QmsAudit.objects.filter(planned_date__year=selected_year,
                                                                  Organization=selected_org, standard=stand,
                                                                  setup=set, audit_status='Completed').count()
                        tot_audit_remaining = QmsAudit.objects.filter(planned_date__year=selected_year,
                                                                  Organization=selected_org, standard=stand,
                                                                  setup=set, audit_status='Remaining').count()
                        tot_inprocess_overdue = QmsAudit.objects.filter(~Q(audit_status='Completed'),planned_date__year=selected_year,
                                                                    Organization=selected_org, standard=stand,
                                                                    setup=set, certification_validity_date__lt = today).count()
                        currYearTotalAudits += tot_current_year_audits
                        currYearAuditInprocess += tot_audit_in_process
                        currYearAuditScheduled += tot_audit_schedule
                        currYearAuditCompleted += tot_audit_completed
                        currYearAuditRemaining += tot_audit_remaining
                        currYearOverdue += tot_inprocess_overdue
                currYeardist = {
                    'currYearTotalAudits': currYearTotalAudits,
                    'currYearAuditInprocess': currYearAuditInprocess,
                    'currYearAuditCompleted': currYearAuditCompleted,
                    'currYearAuditRemaining': currYearAuditRemaining,
                    'currYearAuditScheduled': currYearAuditScheduled,
                    'currYearOverdue': currYearOverdue,
                }
                return JsonResponse({'status': 'True', 'data': currYeardist},
                                    status=200)

            dist = {
                'total_audits': total_audits,
                'audit_completed': audit_completed,
                'audit_under_process': audit_under_process,
                'inprocess_overdue': inprocess_overdue,
            }
            return JsonResponse({'status': 'True', 'data': dist},
                                status=200)

        except Exception as e:
            print(e)
            return JsonResponse({'status': 'False', "message": "Internal Server Error"}, status=500)
            pass

    @staticmethod
    def qmsTrainingSchedule(request):
        qms_obj = QmsTrainingSchedule()
        id = request['id']
        if id == '0':
            qms_obj.sr_no = request['sr_no']
            qms_obj.training_type = request['training_type']
            qms_obj.training_start_date = request['training_start_date']
            qms_obj.training_end_date = request['training_end_date']
            qms_obj.trainer = request['trainer']
            qms_obj.standards = request['standards']
            qms_obj.organizations = request['organizations']
            qms_obj.setups = request['setups']
            qms_obj.save()
            return JsonResponse({'Success': 'Training Schedule inserted Successfully!'})
        else:
            get_obj = QmsTrainingSchedule.objects.filter(id=id).first()
            get_obj.sr_no = request['sr_no']
            get_obj.training_type = request['training_type']
            get_obj.training_start_date = request['training_start_date']
            get_obj.training_end_date = request['training_end_date']
            get_obj.trainer = request['trainer']
            get_obj.standards = request['standards']
            get_obj.organizations = request['organizations']
            get_obj.setups = request['setups']
            get_obj.save()
            return JsonResponse({'Success': 'Training Schedule Updated Successfully!'})

    @staticmethod
    def qmsTrainingScheduleList(request):
        selected_year = request.query_params.get('selected_year')
        selected_org = request.query_params.get('selected_organization')
        selected_standard = request.query_params.getlist('selected_standard')
        selected_setup = request.query_params.getlist('selected_setup')
        standItems = ""
        setupItems = ""
        noVal = ['']
        if selected_standard != noVal:
            for item in selected_standard:
                standItems = item.split(',')
        if selected_setup != noVal:
            for item in selected_setup:
                setupItems = item.split(',')
        dataList = []
        if selected_year == '':
            if selected_org == '' and selected_standard == noVal and selected_setup == noVal:
                data = QmsTrainingSchedule.objects.all()
                serializer = QmsTrainingScheduleSerializer(data, many=True)
                return JsonResponse({'data': serializer.data}, safe=False, status=200)

            elif selected_org == '' and selected_standard == noVal and selected_setup != noVal:
                for set in setupItems:
                    trainingList = QmsTrainingSchedule.objects.filter(setups=set).values()
                    dataList.extend(list(trainingList))


            elif selected_org == '' and selected_standard != noVal  and selected_setup == noVal:
                for stand in standItems:
                    trainingList = QmsTrainingSchedule.objects.filter(standards=stand).values()
                    dataList.extend(list(trainingList))


            elif selected_org == '' and selected_standard != noVal  and selected_setup != noVal :
                for (stand,set) in itertools.zip_longest(standItems,setupItems):
                        trainingList = QmsTrainingSchedule.objects.filter(standards=stand, setups=set).values()
                        dataList.extend(list(trainingList))

            elif selected_org != '' and selected_standard == noVal and selected_setup == noVal:
                dataList = QmsTrainingSchedule.objects.filter(organizations=selected_org).values()


            elif selected_org != '' and selected_standard == noVal and selected_setup != noVal:
                for set in setupItems:
                    trainingList = QmsTrainingSchedule.objects.filter(organizations=selected_org, setups=set).values()
                    dataList.extend(list(trainingList))

            elif selected_org != '' and selected_standard != noVal and selected_setup == noVal:
                for stand in standItems:
                    trainingList = QmsTrainingSchedule.objects.filter(organizations=selected_org,
                                                          standards=stand).values()
                    dataList.extend(list(trainingList))

            else:
                for (stand,set) in itertools.zip_longest(standItems,setupItems):
                    trainingList = QmsTrainingSchedule.objects.filter(organizations=selected_org, standards=stand,
                                                          setups=set).values()
                    dataList.extend(list(trainingList))


        else:
            if selected_org == '' and selected_standard == noVal and selected_setup == noVal:
                data = QmsTrainingSchedule.objects.filter(training_end_date__year = selected_year).values()
                serializer = QmsTrainingScheduleSerializer(data, many=True)
                return JsonResponse({'data': serializer.data}, safe=False, status=200)

            elif selected_org == '' and selected_standard == noVal and selected_setup != noVal:
                for set in setupItems:
                    trainingList = QmsTrainingSchedule.objects.filter(training_end_date__year = selected_year,setups=set).values()
                    dataList.extend(list(trainingList))


            elif selected_org == '' and selected_standard != noVal and selected_setup == noVal:
                for stand in standItems:
                    trainingList = QmsTrainingSchedule.objects.filter(training_end_date__year = selected_year,standards=stand).values()
                    dataList.extend(list(trainingList))


            elif selected_org == '' and selected_standard != noVal and selected_setup != noVal:
                for (stand,set) in itertools.zip_longest(standItems, setupItems):
                    trainingList = QmsTrainingSchedule.objects.filter(training_end_date__year = selected_year,standards=stand,
                                                                      setups=set).values()
                    dataList.extend(list(trainingList))


            elif selected_org != '' and selected_standard == noVal and selected_setup == noVal:
                dataList = QmsTrainingSchedule.objects.filter(training_end_date__year = selected_year,organizations=selected_org).values()


            elif selected_org != '' and selected_standard == noVal and selected_setup != noVal:
                for set in setupItems:
                    trainingList = QmsTrainingSchedule.objects.filter(training_end_date__year = selected_year,organizations=selected_org,
                                                                      setups=set).values()
                    dataList.extend(list(trainingList))

            elif selected_org != '' and selected_standard != noVal and selected_setup == noVal:
                for stand in standItems:
                    trainingList = QmsTrainingSchedule.objects.filter(training_end_date__year = selected_year,organizations=selected_org,
                                                                      standards=stand).values()
                    dataList.extend(list(trainingList))

            else:
                for (stand,set) in itertools.zip_longest(standItems, setupItems):
                    trainingList = QmsTrainingSchedule.objects.filter(training_end_date__year = selected_year,organizations=selected_org,
                                                                      standards=stand,
                                                                      setups=set).values()
                    dataList.extend(list(trainingList))

        serializer = QmsTrainingScheduleSerializer(dataList, many=True)
        return JsonResponse({'data': serializer.data}, safe=False, status=200)


    @staticmethod
    def qmsAuditScheduled(request):
        aud_obj = QmsAuditScheduled()
        id = request['id']
        if id == '0':
            aud_obj.sr_no = request['sr_no']
            aud_obj.organization = request['organization']
            aud_obj.standard = request['standard']
            aud_obj.setup = request['setup']
            aud_obj.audit_due_date = request['audit_due_date']
            aud_obj.audit_done = request['audit_done']
            aud_obj.followup_done = request['followup_done']
            aud_obj.certifification_validity_date = request['certifification_validity_date']
            if request['certifification_validity_rescheduling_date'] != '':
                aud_obj.certifification_validity_rescheduling_date = request['certifification_validity_rescheduling_date']
            aud_obj.category = request['category']
            aud_obj.status = request['status']
            aud_obj.next_due_date = request['next_due_date']
            aud_obj.remarks = request['remarks']
            aud_obj.save()
            return JsonResponse({'Success': 'Training Schedule inserted Successfully!'})
        else:
            get_obj = QmsTrainingSchedule.objects.filter(id=id).first()
            get_obj.sr_no = request['sr_no']
            get_obj.organization = request['organization']
            get_obj.standard = request['standard']
            get_obj.setup = request['setup']
            get_obj.audit_due_date = request['audit_due_date']
            get_obj.audit_done = request['audit_done']
            get_obj.followup_done = request['followup_done']
            get_obj.certifification_validity_date = request['certifification_validity_date']
            get_obj.certifification_validity_rescheduling_date = request['certifification_validity_rescheduling_date']
            get_obj.category = request['category']
            get_obj.status = request['status']
            get_obj.next_due_date = request['next_due_date']
            get_obj.remarks = request['remarks']
            get_obj.save()
            return JsonResponse({'Success': 'Training Schedule Updated Successfully!'})

    @staticmethod
    def qmsAuditScheduledList(request):
        selected_year = request.query_params.get('selected_year')
        selected_org = request.query_params.get('selected_organization')
        selected_standard = request.query_params.getlist('selected_standard')
        selected_setup = request.query_params.getlist('selected_setup')
        standItems = ""
        setupItems = ""
        noVal = ['']
        if selected_standard != noVal:
            for item in selected_standard:
                standItems = item.split(',')
        if selected_setup != noVal:
            for item in selected_setup:
                setupItems = item.split(',')
        dataList = []
        if selected_year == '':
            if selected_org == '' and selected_standard == noVal and selected_setup == noVal:
                data = QmsAuditScheduled.objects.all()
                serializer = QmsAuditScheduleSerializer(data, many=True)
                return JsonResponse({'data': serializer.data}, safe=False, status=200)

            elif selected_org == '' and selected_standard == noVal and selected_setup != noVal:
                for set in setupItems:
                    trainingList = QmsAuditScheduled.objects.filter(setup=set).values()
                    dataList.extend(list(trainingList))


            elif selected_org == '' and selected_standard != noVal and selected_setup == noVal:
                for stand in standItems:
                    trainingList = QmsAuditScheduled.objects.filter(standard=stand).values()
                    dataList.extend(list(trainingList))


            elif selected_org == '' and selected_standard != noVal and selected_setup != noVal:
                for (stand, set) in itertools.zip_longest(standItems, setupItems):
                    trainingList = QmsAuditScheduled.objects.filter(standard=stand, setup=set).values()
                    dataList.extend(list(trainingList))

            elif selected_org != '' and selected_standard == noVal and selected_setup == noVal:
                dataList = QmsAuditScheduled.objects.filter(organization=selected_org).values()


            elif selected_org != '' and selected_standard == noVal and selected_setup != noVal:
                for set in setupItems:
                    trainingList = QmsAuditScheduled.objects.filter(organization=selected_org, setup=set).values()
                    dataList.extend(list(trainingList))

            elif selected_org != '' and selected_standard != noVal and selected_setup == noVal:
                for stand in standItems:
                    trainingList = QmsAuditScheduled.objects.filter(organization=selected_org,
                                                                      standard=stand).values()
                    dataList.extend(list(trainingList))

            else:
                for (stand, set) in itertools.zip_longest(standItems, setupItems):
                    trainingList = QmsAuditScheduled.objects.filter(organization=selected_org, standard=stand,
                                                                      setup=set).values()
                    dataList.extend(list(trainingList))


        else:
            if selected_org == '' and selected_standard == noVal and selected_setup == noVal:
                data = QmsAuditScheduled.objects.filter(training_end_date__year=selected_year).values()
                serializer = QmsAuditScheduleSerializer(data, many=True)
                return JsonResponse({'data': serializer.data}, safe=False, status=200)

            elif selected_org == '' and selected_standard == noVal and selected_setup != noVal:
                for set in setupItems:
                    trainingList = QmsAuditScheduled.objects.filter(training_end_date__year=selected_year,
                                                                      setup=set).values()
                    dataList.extend(list(trainingList))


            elif selected_org == '' and selected_standard != noVal and selected_setup == noVal:
                for stand in standItems:
                    trainingList = QmsAuditScheduled.objects.filter(training_end_date__year=selected_year,
                                                                      standard=stand).values()
                    dataList.extend(list(trainingList))


            elif selected_org == '' and selected_standard != noVal and selected_setup != noVal:
                for (stand, set) in itertools.zip_longest(standItems, setupItems):
                    trainingList = QmsAuditScheduled.objects.filter(training_end_date__year=selected_year,
                                                                      standard=stand,
                                                                      setup=set).values()
                    dataList.extend(list(trainingList))


            elif selected_org != '' and selected_standard == noVal and selected_setup == noVal:
                dataList = QmsAuditScheduled.objects.filter(training_end_date__year=selected_year,
                                                              organization=selected_org).values()


            elif selected_org != '' and selected_standard == noVal and selected_setup != noVal:
                for set in setupItems:
                    trainingList = QmsAuditScheduled.objects.filter(training_end_date__year=selected_year,
                                                                      organization=selected_org,
                                                                      setup=set).values()
                    dataList.extend(list(trainingList))

            elif selected_org != '' and selected_standard != noVal and selected_setup == noVal:
                for stand in standItems:
                    trainingList = QmsAuditScheduled.objects.filter(training_end_date__year=selected_year,
                                                                      organization=selected_org,
                                                                      standard=stand).values()
                    dataList.extend(list(trainingList))

            else:
                for (stand, set) in itertools.zip_longest(standItems, setupItems):
                    trainingList = QmsAuditScheduled.objects.filter(training_end_date__year=selected_year,
                                                                      organization=selected_org,
                                                                      standard=stand,
                                                                      setup=set).values()
                    dataList.extend(list(trainingList))

        serializer = QmsAuditScheduleSerializer(dataList, many=True)
        return JsonResponse({'data': serializer.data}, safe=False, status=200)

    @staticmethod
    def GetQmsAuditHistory(request):

        try:
            audit_id = request.query_params.get('id')

            docList = QmsAuditHistory.objects.filter(qms_audit_id=audit_id).order_by('-id')
            serializer = QmsAuditSerializer(docList, many=True)
            return JsonResponse({'status': 'True', 'data': serializer.data},
                                status=200)

        except Exception as e:
            print(e)
            return JsonResponse({'status': 'False', "message": "Internal Server Error"}, status=500)
            pass


    @staticmethod
    def DeleteQmsAudit(request):
        try:
            auditId = request.query_params['id']
            qms = QmsAudit.objects.get(id=auditId)
            qms.delete()
            return JsonResponse({'message': 'QMS Audit has been deleted'}, status=200)
        except:
            return JsonResponse({'message': 'Sorry! No QMS Audit found.'}, status=500)

    @staticmethod
    def GetQmsAuditPDFList(request):
        TABLE_COL_NAMES = (
            "Audit ID", "Organization", "Site", "Certification Status", "Certification Validity ",
            "Certification Rescheduling", "Audit Type", "Planned Date", "Audit Start Date", "Audit Close Date", "Audit Status", "Standard", "Remarks")
        data = QmsAudit.objects.all().values_list('audit_id', 'Organization', 'site', 'certification_status',
                                                  'certification_validity_date', 'certification_validity_rescheduling_date','audit_type',
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
            "Audit ID", "Organization", "Site", "Certification Status", "Previous Standard", "Certification Validity",
            "Certification Rescheduling","Audit Type", "Planned Date",
            "Audit Start Date", "Audit Close Date", "Audit Status", "Standard", "Remarks")
        for col_num in range(len(TABLE_COL_NAMES)):
            work_sheet.write(row_num, col_num, TABLE_COL_NAMES[col_num], font_style)
        font_style = xlwt.XFStyle()

        data = QmsAudit.objects.all().values_list('audit_id', 'Organization', 'site', 'certification_status',
                                                  'previous_standard', 'certification_validity_date', 'certification_validity_rescheduling_date',
                                                  'audit_type',
                                                  'planned_date',
                                                  'audit_start_date', 'audit_close_date', 'audit_status', 'standard',
                                                  'remarks')
        for row in data:
            row_num += 1
            for col_num in range(len(row)):
                work_sheet.write(row_num, col_num, str(row[col_num]), font_style)
        work_book.save(response)
        return response

    # =========================CeSP Audit====================================
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
            return JsonResponse({'data': serializer.data}, safe=False, status=200)
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
            "Audit Type", "Planned Date", "Audit Start Date", "Audit Close Date", "Audit Status", "Standard", "Remarks")
        data = CespAudit.objects.all().values_list('audit_id', 'Organization', 'site', 'certification_status',
                                                   'previous_standard', 'certification_validity', 'audit_type',
                                                   'planned_date', 'audit_start_date', 'audit_close_date',
                                                   'audit_status',
                                                   'standard', 'remarks')

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
