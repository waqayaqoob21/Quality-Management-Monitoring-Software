from django.urls import path,re_path
from .views import *

urlpatterns = [
        path('addqms/', AddQmsAuditAPIVIEW.as_view(), name='addqms'),
        path('editqms/', AddQmsAuditAPIVIEW.as_view(), name='editqms'),
        path('qmslist/', GetQmsAuditListCountAPIVIEW.as_view(), name='qmslist'),
        path('qmsauditlist/',GetQmsAuditListAPIVIEW.as_view(), name='qmsauditlist'),
        path('deleteqmsaudit/', DeletetQmsAuditAPIVIEW.as_view(), name='deleteqmsaudit'),
        path('qmsauditistory/', GetQmsAuditHistoryAPIVIEW.as_view(), name='qmsauditistory'),
        # re_path(r'^deleteqms/(?P<pk>[0-9]+)$',DeletetQmsAuditAPIVIEW.as_view(), name='deleteqms'),
        path('qmspdf/', GetQmsAuditPDFAPIVIEW.as_view(), name='qmspdf'),
        path('qmsexcel/', GetQmsAuditExcelAPIVIEW.as_view(), name='qmsexcel'),
        path('addqmstrainingschedule/', AddQmsTrainingScheduleAPIVIEW.as_view(), name='qmstrainingschedule'),
        path('qmstrainingschedulelist/', GetQmsTrainingScheduleListAPIVIEW.as_view(), name='qmstrainingschedulelist'),
        path('deleteqmstrainingscheduled/', DeletetQmsTrainingScheduleAPIVIEW.as_view(), name='deleteqmstrainingscheduled'),

        path('addqmsauditscheduled/', addQmsAuditScheduledAPIVIEW.as_view(), name='addqmsauditscheduled'),
        path('qmsauditscheduledlist/', GetQmsAuditScheduledListAPIVIEW.as_view(), name='qmsauditscheduledlist'),
        path('deleteqmsauditscheduled/', DeletetQmsAuditScheduleAPIVIEW.as_view(), name='deleteqmsauditscheduled'),

        path('addcesp/', AddCespAuditAPIVIEW.as_view(), name='addcesp'),
        path('editcesp/', AddCespAuditAPIVIEW.as_view(), name='editcesp'),
        path('cesplist/', GetCespAuditListAPIVIEW.as_view(), name='cesplist'),
        path('cesplistcount/', GetCespAuditListCountAPIVIEW.as_view(), name='cesplistcount'),
        path('deletecespaudit/', DeletetCespAuditAPIVIEW.as_view(), name='deletecespaudit'),
        path('cespauditistory/', GetCespAuditHistoryAPIVIEW.as_view(), name='cespauditistory'),

        path('addcesptraining/', AddCespTrainingCalendarAPIVIEW.as_view(), name='addcesptraining'),
        path('cesptraining/', GetCespTraingingCalendarListAPIVIEW.as_view(), name='cesptraining'),
        path('deletecesptraining/', DeletetCespTraingingCalendarAPIVIEW.as_view(), name='deletecesptraining'),

        path('cespschedule/', GetCespAuditSchduleListAPIVIEW.as_view(), name='cespschedule'),
        path('addcespschedule/', AddCespAuditScheduleAPIVIEW.as_view(), name='addcespschedule'),
        path('deletecespauditschedule/', DeleteCespAuditScheduledAPIView.as_view(), name='deletecespauditschedule'),

        # re_path(r'^deletecesp/(?P<pk>[0-9]+)$',DeletetCespAuditAPIVIEW.as_view(), name='deletecesp'),
        path('cesppdf/', GetCespAuditPDFAPIVIEW.as_view(), name='cesppdf'),
        path('cespexcel/', GetCespAuditExcelAPIVIEW.as_view(), name='cespexcel'),


        path('getnotificationsqms/', GetNotificationsQMSAPIVIEW.as_view(), name='getnotificationsqms'),
        path('getnotificationscesp/', GetNotificationsCeSPAPIVIEW.as_view(), name='getnotificationscesp'),

        path('getnotificationsdetailqms/', GetNotificationsDetailQMSAPIVIEW.as_view(), name='getnotificationsdetailqms'),
        path('getnotificationsdetailcesp/', GetNotificationsDetailCeSPAPIVIEW.as_view(), name='getnotificationsdetailcesp'),
        path('getnotificationcount/', GetNotificationsCountAPIVIEW.as_view(),name='getnotificationcount'),


        path('addauditschedule/', AddAuditSchduleAPIVIEW.as_view(), name='addauditschedule'),
        path('getauditschedule/', GetAuditSchduleAPIVIEW.as_view(), name='getauditschedule'),
]