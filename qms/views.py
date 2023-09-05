from rest_framework.permissions import AllowAny
from rest_framework.views import APIView

from qms.QmsController import *

qms_obj = QmsController()
# Create your views here.

class AddQmsAuditAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def post(self, request):
        print(request.data)
        result = qms_obj.AddQmsAudit(request.data)
        return result
class DeletetQmsAuditAPIVIEW(APIView):
    permission_classes = [AllowAny]
    def get(self, request):
        result = qms_obj.DeleteQmsAudit(request)
        print(result)
        return result

class GetQmsAuditListAPIVIEW(APIView):
        permission_classes = [AllowAny]

        def get(self, request):
            result = qms_obj.GetQmsAuditList(request)
            return result


class GetQmsAuditHistoryAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def get(self, request):
        result = qms_obj.GetQmsAuditHistory(request)
        return result

class GetQmsAuditListCountAPIVIEW(APIView):
    permission_classes = [AllowAny]
    def get(self, request):
        result = qms_obj.GetQmsAuditListCount(request)
        return result

class GetQmsAuditPDFAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def get(self, request):
        result = qms_obj.GetQmsAuditPDFList(request.data)
        return result

class GetNotificationsQMSAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def get(self, request):
        result = qms_obj.GetNotificationsQMS(request.data)
        return result

class GetNotificationsCeSPAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def get(self, request):
        result = qms_obj.GetNotificationsCeSP(request.data)
        return result

class GetNotificationsDetailQMSAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def get(self, request):
        result = qms_obj.GetNotificationsDetailQMS(request)
        return result

class GetNotificationsDetailCeSPAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def get(self, request):
        result = qms_obj.GetNotificationsDetailCeSP(request)
        return result

class GetNotificationsCountAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def get(self, request):
        result = qms_obj.getNotificationCount(request.data)
        return result
class GetQmsAuditExcelAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def get(self, request):
        result = qms_obj.GetQmsAuditExcelList(request.data)
        return result

class AddQmsTrainingScheduleAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def post(self, request):
        result = qms_obj.AddQmsTrainingSchedule(request.data)
        return result

class GetQmsTrainingScheduleListAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def get(self, request):
        result = qms_obj.qmsTrainingScheduleList(request)
        return result
class DeletetQmsTrainingScheduleAPIVIEW(APIView):
    permission_classes = [AllowAny]
    def get(self, request):
        result = qms_obj.DeleteQmsTrainingScheduled(request)
        return result
class addQmsAuditScheduledAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def post(self, request):
        result = qms_obj.AddQmsAuditScheduled(request.data)
        return result

class GetQmsAuditScheduledListAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def get(self, request):
        result = qms_obj.qmsAuditScheduledList(request)
        return result


class DeletetQmsAuditScheduleAPIVIEW(APIView):
    permission_classes = [AllowAny]
    def get(self, request):
        result = qms_obj.DeleteQmsAuditSchedule(request)
        print(result)
        return result
# ======================CeSP Audit===============

class AddCespAuditAPIVIEW(APIView):
    permission_classes = [AllowAny]
    def post(self, request):
        result = qms_obj.AddCespAudit(request.data)
        return result

class DeletetCespAuditAPIVIEW(APIView):
    permission_classes = [AllowAny]
    def get(self, request):
        # print(request.data)
        result = qms_obj.DeleteCespAudit(request)
        print(result)
        return result
class GetCespAuditListCountAPIVIEW(APIView):
    permission_classes = [AllowAny]
    def get(self, request):
        result = qms_obj.GetCespAuditListCount(request)
        return result
class GetCespAuditListAPIVIEW(APIView):
    permission_classes = [AllowAny]
    def get(self, request):
        result = qms_obj.GetCespAuditList(request)
        return result
class AddCespTrainingCalendarAPIVIEW(APIView):
    permission_classes = [AllowAny]
    def post(self, request):
        result = qms_obj.AddCespTrainingCalendar(request.data)
        return result
class GetCespTraingingCalendarListAPIVIEW(APIView):
    permission_classes = [AllowAny]
    def get(self, request):
        result = qms_obj.GetCespTrainingCalendarList(request.data)
        return result

class DeletetCespTraingingCalendarAPIVIEW(APIView):
    permission_classes = [AllowAny]
    def get(self, request):
        # print(request.data)
        result = qms_obj.DeleteCespTrainingCalendar(request)
        print(result)
        return result
class AddCespAuditScheduleAPIVIEW(APIView):
    permission_classes = [AllowAny]
    def post(self, request):
        result = qms_obj.AddCespAuditSchedule(request.data)
        return result
class GetCespAuditSchduleListAPIVIEW(APIView):
    permission_classes = [AllowAny]
    def get(self, request):
        result = qms_obj.GetCespAuditScheduledList(request)
        return result
class GetCespAuditHistoryAPIVIEW(APIView):
    permission_classes = [AllowAny]
    def get(self, request):
        result = qms_obj.GetCespAuditHistory(request)
        return result

class GetCespAuditPDFAPIVIEW(APIView):
    permission_classes = [AllowAny]
    def get(self, request):
        result = qms_obj.GetCespAuditPDFList(request.data)
        return result

class GetCespAuditExcelAPIVIEW(APIView):
    permission_classes = [AllowAny]
    def get(self, request):
        result = qms_obj.GetCespAuditExcelList(request.data)
        return result

class DeleteCespAuditScheduledAPIView(APIView):
    permission_classes = [AllowAny]
    def get(self, request):
        result = qms_obj.DeleteCespAuditScheduled(request)
        print(result)
        return result


class AddAuditSchduleAPIVIEW(APIView):
    permission_classes = [AllowAny]
    def post(self, request):
        result = qms_obj.addAuditSchdule(request.data)
        return result


class GetAuditSchduleAPIVIEW(APIView):
    permission_classes = [AllowAny]
    def get(self, request):
        result = qms_obj.getAuditSchedule(request)
        print(result)
        return result