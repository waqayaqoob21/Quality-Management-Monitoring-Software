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
        path('addqmstrainingschedule/', GetQmsTrainingScheduleAPIVIEW.as_view(), name='qmstrainingschedule'),
        path('qmstrainingschedulelist/', GetQmsTrainingScheduleListAPIVIEW.as_view(), name='qmstrainingschedulelist'),
        path('addqmsauditscheduled/', GetQmsAuditScheduledAPIVIEW.as_view(), name='addqmsauditscheduled'),
        path('qmsauditscheduledlist/', GetQmsAuditScheduledListAPIVIEW.as_view(), name='qmsauditscheduledlist'),


        path('addcesp/', AddCespAuditAPIVIEW.as_view(), name='addcesp'),
        path('editcesp/', AddCespAuditAPIVIEW.as_view(), name='editcesp'),
        path('cesplist/', GetCespAuditListAPIVIEW.as_view(), name='cesplist'),
        path('cesplistcount/', GetCespAuditListCountAPIVIEW.as_view(), name='cesplistcount'),
        path('deletecespaudit/', DeletetCespAuditAPIVIEW.as_view(), name='deletecespaudit'),
        path('cespauditistory/', GetCespAuditHistoryAPIVIEW.as_view(), name='cespauditistory'),
        path('cesptraining/', GetCespTraingingCalendarListAPIVIEW.as_view(), name='cesptraining'),
        path('cespschedule/', GetCespAuditSchduleListAPIVIEW.as_view(), name='cespschedule'),
        path('addcesptraining/', AddCespTrainingCalendarAPIVIEW.as_view(), name='addcesptraining'),
        path('addcespschedule/', AddCespAuditScheduleAPIVIEW.as_view(), name='addcespschedule'),
        # re_path(r'^deletecesp/(?P<pk>[0-9]+)$',DeletetCespAuditAPIVIEW.as_view(), name='deletecesp'),
        path('cesppdf/', GetCespAuditPDFAPIVIEW.as_view(), name='cesppdf'),
        path('cespexcel/', GetCespAuditExcelAPIVIEW.as_view(), name='cespexcel'),
   
]