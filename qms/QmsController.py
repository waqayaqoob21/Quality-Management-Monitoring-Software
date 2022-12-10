from django.http import JsonResponse
from fpdf import FPDF
import xlwt
from datetime import datetime, timedelta
from django.http import FileResponse
from django.http import HttpResponse
from django.core.files.storage import FileSystemStorage
from django.db.models import F, Q

from qms.models import *
from qms.serializer import *


class QmsController:
    @staticmethod
    def AddQmsAudit(request):
            print("thi is add qms controller")
            qmsModel = QmsAudit()
        # try:
            id = request['id']
            if id == '0':
                qmsModel.audit_id = 0 #request['audit_id']
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
                qmsModel.certification_setup = request['certification_setup']
                qmsModel.save()
                return JsonResponse({'status': 'True', 'message': "QMS Audit Created Successfully!"},
                                status=200)
            else:
                get_qms = QmsAudit.objects.filter(id=id).first()
                if get_qms is not None:
                    get_qms.audit_id = 0 #request['audit_id']
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
                    get_qms.certification_setup = request['certification_setup']
                    get_qms.save()
                    return JsonResponse({'status': 'True', 'message': "QMS Audit Updated Successfully!"},
                                status=200)
        # except Exception as e:
        #     return JsonResponse({'status': 'False', "message": "QMS Audit Not Saved"}, status=500)

    @staticmethod
    def GetQmsAuditListCount(request, self=None):
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
            selected_org = request.query_params.get('selected_organization')
            selected_setup = request.query_params.get('selected_setup')
            selected_standard = request.query_params.get('selected_standard')

            typeQuery = Q()
            filter_objects = Q()
            if selected_standard == 'PA':
                filter_objects &= get_filter('standard', 'equal','PA')
            elif selected_standard == 'IS O9001 : 2015':
                filter_objects &= get_filter('standard', 'equal','IS O9001 : 2015')
            elif selected_standard == 'AS9100 : Rev D':
                filter_objects &= get_filter('standard', 'equal','AS9100 : Rev D')
            elif selected_standard == 'ISO17025 : 2017':
                filter_objects &= get_filter('standard', 'equal','ISO17025 : 2017')
            else:
                filter_objects &= get_filter('standard', 'equal','ISO17020 : 2012')

            total_audits = 0
            current_year_audits = 0
            audit_in_process = 0
            audit_completed = 0
            inprocess_overdue = 0
            total_audit_schedule = 0
            audit_remaining = 0

            inprocess_overdue = QmsAudit.objects.filter(audit_close_date__lt=F('planned_date')).count()
            total_audits = QmsAudit.objects.count()

            if selected_year is '':
                current_year_audits = QmsAudit.objects.filter(filter_objects).count()
                audit_in_process = QmsAudit.objects.filter(audit_status='In process').count()
                audit_completed = QmsAudit.objects.filter(audit_status='Completed').count()
                total_audit_schedule = QmsAudit.objects.filter(audit_status='Scheduled').count()
                audit_remaining = QmsAudit.objects.filter(audit_status='Remaining').count()

                if selected_standard is not '' and selected_org is '' and selected_setup is '':
                    current_year_audits = QmsAudit.objects.filter(filter_objects).count()
                    audit_in_process = QmsAudit.objects.filter(filter_objects, audit_status='In process').count()
                    total_audit_schedule = QmsAudit.objects.filter(filter_objects, audit_status='Scheduled').count()
                    audit_completed = QmsAudit.objects.filter(filter_objects, audit_status='Completed').count()
                    audit_remaining = QmsAudit.objects.filter(filter_objects, audit_status='Remaining').count()

                elif selected_standard is '' and selected_org is not '' and selected_setup is not '':
                    current_year_audits = QmsAudit.objects.filter(Organization=selected_org, certification_setup = selected_setup).count()
                    audit_in_process = QmsAudit.objects.filter(audit_status='In process',Organization=selected_org,certification_setup = selected_setup).count()
                    total_audit_schedule = QmsAudit.objects.filter(audit_status='Scheduled',Organization=selected_org,certification_setup = selected_setup).count()
                    audit_completed = QmsAudit.objects.filter(audit_status='Completed',Organization=selected_org,certification_setup = selected_setup).count()
                    audit_remaining = QmsAudit.objects.filter(audit_status='Remaining',Organization=selected_org,certification_setup = selected_setup).count()

                elif selected_standard is not '' and selected_org is not '' and selected_setup is '':
                    current_year_audits = QmsAudit.objects.filter(filter_objects,Organization=selected_org).count()
                    audit_in_process = QmsAudit.objects.filter(filter_objects,audit_status='In process',Organization=selected_org).count()
                    total_audit_schedule = QmsAudit.objects.filter(filter_objects,Organization=selected_org, audit_status='Scheduled').count()
                    audit_completed = QmsAudit.objects.filter(filter_objects,Organization=selected_org, audit_status='Completed').count()
                    audit_remaining = QmsAudit.objects.filter(filter_objects,Organization=selected_org, audit_status='Remaining').count()

                elif selected_standard is '' and selected_org is not '' and selected_setup is '':
                    current_year_audits = QmsAudit.objects.filter(Organization=selected_org).count()
                    audit_in_process = QmsAudit.objects.filter(audit_status='In process',Organization=selected_org).count()
                    total_audit_schedule = QmsAudit.objects.filter(audit_status='Scheduled',Organization=selected_org).count()
                    audit_completed = QmsAudit.objects.filter(audit_status='Completed',Organization=selected_org).count()
                    audit_remaining = QmsAudit.objects.filter(audit_status='Remaining',Organization=selected_org).count()

                elif selected_standard is not '' and selected_org is not ''and selected_setup is not '':
                    current_year_audits = QmsAudit.objects.filter(filter_objects,Organization=selected_org,certification_setup = selected_setup).count()
                    audit_in_process = QmsAudit.objects.filter(filter_objects, audit_status='In process',Organization=selected_org,certification_setup = selected_setup).count()
                    total_audit_schedule = QmsAudit.objects.filter(filter_objects, audit_status='Scheduled',Organization=selected_org,certification_setup = selected_setup).count()
                    audit_completed = QmsAudit.objects.filter(filter_objects, audit_status='Completed',Organization=selected_org,certification_setup = selected_setup).count()
                    audit_remaining = QmsAudit.objects.filter(filter_objects, audit_status='Remaining',Organization=selected_org,certification_setup = selected_setup).count()
            else:
                if selected_standard is '' and selected_org is '' and selected_setup is '':
                    current_year_audits = QmsAudit.objects.filter(audit_close_date__year=selected_year,Organization=selected_org).count()
                    audit_in_process = QmsAudit.objects.filter(audit_status='In process',audit_close_date__year=selected_year,Organization=selected_org).count()
                    total_audit_schedule = QmsAudit.objects.filter(audit_status='Scheduled',audit_close_date__year=selected_year,Organization=selected_org).count()
                    audit_completed = QmsAudit.objects.filter(audit_status='Completed',audit_close_date__year=selected_year,Organization=selected_org).count()
                    audit_remaining = QmsAudit.objects.filter(audit_status='Remaining',audit_close_date__year=selected_year,Organization=selected_org).count()
                elif selected_standard is not '' and selected_org is '' and selected_setup is '':
                    if selected_org == 'All Nocs':
                        current_year_audits = QmsAudit.objects.filter(filter_objects,
                                                                      audit_close_date__year=selected_year).count()
                        audit_in_process = QmsAudit.objects.filter(filter_objects, audit_status='In process',
                                                                   audit_close_date__year=selected_year).count()
                        total_audit_schedule = QmsAudit.objects.filter(filter_objects, audit_status='Scheduled',
                                                                       audit_close_date__year=selected_year).count()
                        audit_completed = QmsAudit.objects.filter(filter_objects, audit_status='Completed',
                                                                  audit_close_date__year=selected_year).count()
                        audit_remaining = QmsAudit.objects.filter(filter_objects, audit_status='Remaining',
                                                                  audit_close_date__year=selected_year).count()
                    else:
                        current_year_audits = QmsAudit.objects.filter(filter_objects,
                                                                      audit_close_date__year=selected_year).count()
                        audit_in_process = QmsAudit.objects.filter(filter_objects, audit_status='In process',
                                                                   audit_close_date__year=selected_year).count()
                        total_audit_schedule = QmsAudit.objects.filter(filter_objects, audit_status='Scheduled',
                                                                       audit_close_date__year=selected_year).count()
                        audit_completed = QmsAudit.objects.filter(filter_objects, audit_status='Completed',
                                                                  audit_close_date__year=selected_year).count()
                        audit_remaining = QmsAudit.objects.filter(filter_objects, audit_status='Remaining',
                                                                  audit_close_date__year=selected_year).count()

                elif selected_standard is '' and selected_org is not '' and selected_setup is '':
                    if selected_org == 'All Nocs':
                        current_year_audits = QmsAudit.objects.filter(audit_close_date__year=selected_year,Organization=selected_org).count()
                        audit_in_process = QmsAudit.objects.filter(audit_status='In process',audit_close_date__year=selected_year,Organization=selected_org).count()
                        total_audit_schedule = QmsAudit.objects.filter(audit_status='Scheduled',audit_close_date__year=selected_year,Organization=selected_org).count()
                        audit_completed = QmsAudit.objects.filter(audit_status='Completed',audit_close_date__year=selected_year,Organization=selected_org).count()
                        audit_remaining = QmsAudit.objects.filter(audit_status='Remaining',audit_close_date__year=selected_year,Organization=selected_org).count()

                    else:
                        current_year_audits = QmsAudit.objects.filter(audit_close_date__year=selected_year,Organization=selected_org).count()
                        audit_in_process = QmsAudit.objects.filter(audit_status='Audit in-process',audit_close_date__year=selected_year,Organization=selected_org).count()
                        total_audit_schedule = QmsAudit.objects.filter(audit_status='Scheduled',audit_close_date__year=selected_year,Organization=selected_org).count()
                        audit_completed = QmsAudit.objects.filter(audit_status='Completed',audit_close_date__year=selected_year,Organization=selected_org).count()
                        audit_remaining = QmsAudit.objects.filter(audit_status='Remaining',audit_close_date__year=selected_year,Organization=selected_org).count()

                elif selected_standard is'' and selected_org is '' and selected_setup is not '':
                    if selected_org == 'All Nocs':
                        current_year_audits = QmsAudit.objects.filter(audit_close_date__year=selected_year,certification_setup = selected_setup).count()
                        audit_in_process = QmsAudit.objects.filter( audit_status='In process',audit_close_date__year=selected_year,certification_setup = selected_setup).count()
                        total_audit_schedule = QmsAudit.objects.filter(audit_status='Scheduled',audit_close_date__year=selected_year,certification_setup = selected_setup).count()
                        audit_completed = QmsAudit.objects.filter(audit_status='Completed',audit_close_date__year=selected_year,certification_setup = selected_setup).count()
                        audit_remaining = QmsAudit.objects.filter(audit_status='Remaining',audit_close_date__year=selected_year,certification_setup = selected_setup).count()

                    else:
                        current_year_audits = QmsAudit.objects.filter(filter_objects,audit_close_date__year=selected_year,Organization=selected_org,certification_setup = selected_setup).count()
                        audit_in_process = QmsAudit.objects.filter(filter_objects, audit_status='In process',audit_close_date__year=selected_year,certification_setup = selected_setup).count()
                        total_audit_schedule = QmsAudit.objects.filter(filter_objects,audit_status='Scheduled',audit_close_date__year=selected_year,certification_setup = selected_setup).count()
                        audit_completed = QmsAudit.objects.filter(filter_objects, audit_status='Completed',audit_close_date__year=selected_year,certification_setup = selected_setup).count()
                        audit_remaining = QmsAudit.objects.filter(filter_objects, audit_status='Remaining',audit_close_date__year=selected_year,certification_setup = selected_setup).count()

                elif selected_standard is not '' and selected_org is not '' and selected_setup is not '':
                    if selected_org == 'All Nocs':
                        current_year_audits = QmsAudit.objects.filter(filter_objects,
                                                                        audit_close_date__year=selected_year,certification_setup = selected_setup,Organization=selected_org).count()
                        audit_in_process = QmsAudit.objects.filter(filter_objects, audit_status='In process',
                                                                     audit_close_date__year=selected_year,certification_setup = selected_setup,Organization=selected_org).count()
                        total_audit_schedule = QmsAudit.objects.filter(filter_objects,
                                                                        audit_status='Scheduled',
                                                                     audit_close_date__year=selected_year,certification_setup = selected_setup,Organization=selected_org).count()
                        audit_completed = QmsAudit.objects.filter(filter_objects, audit_status='Completed',
                                                                     audit_close_date__year=selected_year,certification_setup = selected_setup,Organization=selected_org).count()
                        audit_remaining = QmsAudit.objects.filter(filter_objects, audit_status='Remaining',
                                                                     audit_close_date__year=selected_year,certification_setup = selected_setup,Organization=selected_org).count()

                    else:
                        current_year_audits = QmsAudit.objects.filter(filter_objects,
                                                                        audit_close_date__year=selected_year,
                                                                        Organization=selected_org,certification_setup = selected_setup).count()
                        audit_in_process = QmsAudit.objects.filter(filter_objects, audit_status='In process',
                                                                     audit_close_date__year=selected_year,
                                                                     Organization=selected_org,certification_setup = selected_setup).count()
                        total_audit_schedule = QmsAudit.objects.filter(filter_objects,
                                                                        audit_status='Scheduled',
                                                                     Organization=selected_org,certification_setup = selected_setup).count()
                        audit_completed = QmsAudit.objects.filter(filter_objects, audit_status='Completed',
                                                                     Organization=selected_org,certification_setup = selected_setup).count()
                        audit_remaining = QmsAudit.objects.filter(filter_objects, audit_status='Remaining',
                                                                     Organization=selected_org,certification_setup = selected_setup).count()


            dist = {
                'total_audits': total_audits,
                'current_year_audits': current_year_audits,
                'audit_in_process': audit_in_process,
                'audit_completed':audit_completed,
                'audit_remaining':audit_remaining,
                'total_audit_schedule': total_audit_schedule,
                'inprocess_overdue':inprocess_overdue,

            }

            # DataCount.append(dist)
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
        selected_setup = request.query_params.get('selected_setup')
        selected_standard = request.query_params.get('selected_standard')

        if selected_year is '':
            if selected_org is   '' and selected_standard == '0'  and selected_setup == '0':
                data = QmsTrainingSchedule.objects.all()
                serializer = QmsTrainingScheduleSerializer(data, many=True)
                return JsonResponse({'data': serializer.data}, safe=False, status=200)

            elif selected_org is '' and selected_standard == '0' and selected_setup is not'':
                data = QmsTrainingSchedule.objects.filter(setups=selected_setup)
                serializer = QmsTrainingScheduleSerializer(data, many=True)
                return JsonResponse({'data': serializer.data}, safe=False, status=200)

            elif selected_org is '' and selected_standard is  not'' and selected_setup == '0':
                data = QmsTrainingSchedule.objects.filter(standards=selected_standard)
                serializer = QmsTrainingScheduleSerializer(data, many=True)
                return JsonResponse({'data': serializer.data}, safe=False, status=200)

            elif selected_org is '' and selected_standard  is not'' and selected_setup is not'':
                data = QmsTrainingSchedule.objects.filter(standards=selected_standard,setups=selected_setup)
                serializer = QmsTrainingScheduleSerializer(data, many=True)
                return JsonResponse({'data': serializer.data}, safe=False, status=200)

            elif selected_org is not '' and selected_standard == '0' and selected_setup == '0':
                data = QmsTrainingSchedule.objects.filter(organizations=selected_org)
                serializer = QmsTrainingScheduleSerializer(data, many=True)
                return JsonResponse({'data': serializer.data}, safe=False, status=200)

            elif selected_org is not '' and selected_standard == '0' and selected_setup is not'':
                data = QmsTrainingSchedule.objects.filter(organizations=selected_org,setups=selected_setup)
                serializer = QmsTrainingScheduleSerializer(data, many=True)
                return JsonResponse({'data': serializer.data}, safe=False, status=200)

            elif selected_org is not '' and selected_standard is not '' and selected_setup == '0':
                data = QmsTrainingSchedule.objects.filter(organizations=selected_org,standards=selected_standard)
                serializer = QmsTrainingScheduleSerializer(data, many=True)
                return JsonResponse({'data': serializer.data}, safe=False, status=200)

            else:
                data = QmsTrainingSchedule.objects.filter(organizations=selected_org,standards=selected_standard,setups=selected_setup)
                serializer = QmsTrainingScheduleSerializer(data, many=True)
                return JsonResponse({'data': serializer.data}, safe=False, status=200)

        else:
                if selected_org is '' and selected_standard == '0' and selected_setup == '0':
                    data = QmsTrainingSchedule.objects.filter(training_end_date__year=selected_year)
                    serializer = QmsTrainingScheduleSerializer(data, many=True)
                    return  JsonResponse({'data':serializer.data}, safe=False, status=200)

                elif selected_org is '' and selected_standard == '0' and selected_setup is not'':
                    data = QmsTrainingSchedule.objects.filter(training_end_date__year=selected_year,setups=selected_setup)
                    serializer = QmsTrainingScheduleSerializer(data, many=True)
                    return  JsonResponse({'data':serializer.data}, safe=False, status=200)

                elif selected_org is '' and selected_standard is not '' and selected_setup == '0':
                    data = QmsTrainingSchedule.objects.filter(training_end_date__year=selected_year,standards=selected_standard)
                    serializer = QmsTrainingScheduleSerializer(data, many=True)
                    return  JsonResponse({'data':serializer.data}, safe=False, status=200)

                elif selected_org is '' and selected_standard is not '' and selected_setup is not'':
                    data = QmsTrainingSchedule.objects.filter(training_end_date__year=selected_year,standards=selected_standard,setups=selected_setup)
                    serializer = QmsTrainingScheduleSerializer(data, many=True)
                    return  JsonResponse({'data':serializer.data}, safe=False, status=200)

                elif selected_org is not '' and selected_standard =='0' and selected_setup == '0':
                    data = QmsTrainingSchedule.objects.filter(training_end_date__year = selected_year,organizations=selected_org)
                    serializer = QmsTrainingScheduleSerializer(data, many=True)
                    return  JsonResponse({'data':serializer.data}, safe=False, status=200)

                elif selected_org is not '' and selected_standard =='0' and selected_setup is not'':
                    data = QmsTrainingSchedule.objects.filter(training_end_date__year = selected_year,organizations=selected_org,setups=selected_setup)
                    serializer = QmsTrainingScheduleSerializer(data, many=True)
                    return  JsonResponse({'data':serializer.data}, safe=False, status=200)

                elif selected_org is not '' and selected_standard is not '' and selected_setup == '0':
                    data = QmsTrainingSchedule.objects.filter(training_end_date__year = selected_year,organizations=selected_org,standards=selected_standard)
                    serializer = QmsTrainingScheduleSerializer(data, many=True)
                    return  JsonResponse({'data':serializer.data}, safe=False, status=200)

                else:
                    data = QmsTrainingSchedule.objects.filter(training_end_date__year = selected_year,organizations=selected_org,standards=selected_standard,setups=selected_setup)
                    serializer = QmsTrainingScheduleSerializer(data, many=True)
                    return  JsonResponse({'data':serializer.data}, safe=False, status=200)

        return  JsonResponse({'sorry','Data not found'}, safe=False, status=200)


    @staticmethod
    def qmsAuditScheduled(request):
        aud_obj = QmsAuditScheduled()
        id = request['id']
        if id == '0':
            aud_obj.sr_no = request['sr_no']
            aud_obj.organization = request['organization']
            aud_obj.standard = request['standard']
            aud_obj.audit_due_date = request['audit_due_date']
            aud_obj.audit_done = request['audit_done']
            aud_obj.followup_done = request['followup_done']
            aud_obj.certifification_validity_date = request['certifification_validity_date']
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
        selected_setup = request.query_params.get('selected_setup')
        selected_standard = request.query_params.get('selected_standard')

        if selected_year is '':
            if selected_org is '' and selected_standard == '0' and selected_setup == '0':
                data = QmsAuditScheduled.objects.all()
                serializer = QmsAuditScheduleSerializer(data, many=True)
                return JsonResponse({'data': serializer.data}, safe=False, status=200)

            elif selected_org is '' and selected_standard == '0' and selected_setup is not '':
                data = QmsAuditScheduled.objects.filter(setups=selected_setup)
                serializer = QmsAuditScheduleSerializer(data, many=True)
                return JsonResponse({'data': serializer.data}, safe=False, status=200)

            elif selected_org is '' and selected_standard is not '' and selected_setup == '0':
                data = QmsAuditScheduled.objects.filter(standards=selected_standard)
                serializer = QmsAuditScheduleSerializer(data, many=True)
                return JsonResponse({'data': serializer.data}, safe=False, status=200)

            elif selected_org is '' and selected_standard is not '' and selected_setup is not '':
                data = QmsAuditScheduled.objects.filter(standards=selected_standard, setups=selected_setup)
                serializer = QmsAuditScheduleSerializer(data, many=True)
                return JsonResponse({'data': serializer.data}, safe=False, status=200)

            elif selected_org is not '' and selected_standard == '0' and selected_setup == '0':
                data = QmsAuditScheduled.objects.filter(organizations=selected_org)
                serializer = QmsAuditScheduleSerializer(data, many=True)
                return JsonResponse({'data': serializer.data}, safe=False, status=200)

            elif selected_org is not '' and selected_standard == '0' and selected_setup is not '':
                data = QmsAuditScheduled.objects.filter(organizations=selected_org, setups=selected_setup)
                serializer = QmsAuditScheduleSerializer(data, many=True)
                return JsonResponse({'data': serializer.data}, safe=False, status=200)

            elif selected_org is not '' and selected_standard is not '' and selected_setup == '0':
                data = QmsAuditScheduled.objects.filter(organizations=selected_org, standards=selected_standard)
                serializer = QmsAuditScheduleSerializer(data, many=True)
                return JsonResponse({'data': serializer.data}, safe=False, status=200)

            else:
                data = QmsAuditScheduled.objects.filter(organizations=selected_org, standards=selected_standard,
                                                          setups=selected_setup)
                serializer = QmsAuditScheduleSerializer(data, many=True)
                return JsonResponse({'data': serializer.data}, safe=False, status=200)

        else:
            if selected_org is '' and selected_standard == '0' and selected_setup == '0':
                data = QmsAuditScheduled.objects.filter(training_end_date__year=selected_year)
                serializer = QmsAuditScheduleSerializer(data, many=True)
                return JsonResponse({'data': serializer.data}, safe=False, status=200)

            elif selected_org is '' and selected_standard == '0' and selected_setup is not '':
                data = QmsAuditScheduled.objects.filter(training_end_date__year=selected_year, setups=selected_setup)
                serializer = QmsAuditScheduleSerializer(data, many=True)
                return JsonResponse({'data': serializer.data}, safe=False, status=200)

            elif selected_org is '' and selected_standard is not '' and selected_setup == '0':
                data = QmsAuditScheduled.objects.filter(training_end_date__year=selected_year,
                                                          standards=selected_standard)
                serializer = QmsAuditScheduleSerializer(data, many=True)
                return JsonResponse({'data': serializer.data}, safe=False, status=200)

            elif selected_org is '' and selected_standard is not '' and selected_setup is not '':
                data = QmsAuditScheduled.objects.filter(training_end_date__year=selected_year,
                                                          standards=selected_standard, setups=selected_setup)
                serializer = QmsAuditScheduleSerializer(data, many=True)
                return JsonResponse({'data': serializer.data}, safe=False, status=200)

            elif selected_org is not '' and selected_standard == '0' and selected_setup == '0':
                data = QmsAuditScheduled.objects.filter(training_end_date__year=selected_year,
                                                          organizations=selected_org)
                serializer = QmsAuditScheduleSerializer(data, many=True)
                return JsonResponse({'data': serializer.data}, safe=False, status=200)

            elif selected_org is not '' and selected_standard == '0' and selected_setup is not '':
                data = QmsAuditScheduled.objects.filter(training_end_date__year=selected_year,
                                                          organizations=selected_org, setups=selected_setup)
                serializer = QmsAuditScheduleSerializer(data, many=True)
                return JsonResponse({'data': serializer.data}, safe=False, status=200)

            elif selected_org is not '' and selected_standard is not '' and selected_setup == '0':
                data = QmsAuditScheduled.objects.filter(training_end_date__year=selected_year,
                                                          organizations=selected_org, standards=selected_standard)
                serializer = QmsAuditScheduleSerializer(data, many=True)
                return JsonResponse({'data': serializer.data}, safe=False, status=200)

            else:
                data = QmsAuditScheduled.objects.filter(training_end_date__year=selected_year,
                                                          organizations=selected_org, standards=selected_standard,
                                                          setups=selected_setup)
                serializer = QmsAuditScheduleSerializer(data, many=True)
                return JsonResponse({'data': serializer.data}, safe=False, status=200)

        return JsonResponse({'sorry', 'Data not found'}, safe=False, status=200)




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