from rest_framework.permissions import AllowAny
from rest_framework.views import APIView

from qms.QmsController import *

qms_obj = QmsController()
# Create your views here.

class AddQmsAuditAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def post(self, request):
        result = qms_obj.AddQmsAudit(request.data)
        return result
class DeletetQmsAuditAPIVIEW(APIView):
    permission_classes = [AllowAny]
    def delete(self, request,pk):
        result = qms_obj.DeleteQmsAudit(request.data, pk)
        print(result)
        return result
class GetQmsAuditListAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def get(self, request):
        result = qms_obj.GetQmsAuditList(request.data)
        return result

# ======================CeSP Audit===============

class AddCespAuditAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def post(self, request):
        result = qms_obj.AddCespAudit(request.data)
        return result
class DeletetCespAuditAPIVIEW(APIView):
    permission_classes = [AllowAny]
    def delete(self, request,pk):
        result = qms_obj.DeleteCespAudit(request.data, pk)
        print(result)
        return result
class GetCespAuditListAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def get(self, request):
        result = qms_obj.GetCespAuditList(request.data)
        return result