from rest_framework.permissions import AllowAny
from rest_framework.views import APIView

from mpm.MpmController import *

motor_obj = MpmController()
# Create your views here.

class AddActiveMotorAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def post(self, request):
        result = motor_obj.AddActiveMotor(request.data)
        return result

# class EditActiveMotorAPIVIEW(APIView):
#     permission_classes = [AllowAny]
#     def put(self, request):
#         result = motor_obj.EditActiveMotor(request.data)
#         print(result)
#         return result
class DeletetActiveMotorAPIVIEW(APIView):
    permission_classes = [AllowAny]
    def delete(self, request,pk):
        result = motor_obj.DeleteActiveMotor(request.data, pk)
        print(result)
        return result
class GetActiveMotorListAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def get(self, request):
        result = motor_obj.GetActiveMotorList(request.data)
        return result
class GetActiveMotorPDFAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def get(self, request):
        result = motor_obj.GetActiveMotorPDFList(request.data)
        return result
class GetActiveMotorExcelAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def get(self, request):
        result = motor_obj.GetActiveMotorExcelList(request.data)
        return result
