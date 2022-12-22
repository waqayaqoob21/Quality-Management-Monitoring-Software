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

class GetQmsAuditExcelAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def get(self, request):
        result = qms_obj.GetQmsAuditExcelList(request.data)
        return result

class GetQmsTrainingScheduleAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def post(self, request):
        result = qms_obj.qmsTrainingSchedule(request.data)
        return result

class GetQmsTrainingScheduleListAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def get(self, request):
        result = qms_obj.qmsTrainingScheduleList(request)
        return result

class GetQmsAuditScheduledAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def post(self, request):
        result = qms_obj.qmsAuditScheduled(request.data)
        return result

class GetQmsAuditScheduledListAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def get(self, request):
        result = qms_obj.qmsAuditScheduledList(request)
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
    print("Hello this training calendar")
    permission_classes = [AllowAny]
    def get(self, request):
        result = qms_obj.GetCespTrainingCalendarList(request.data)
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