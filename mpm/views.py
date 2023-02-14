from rest_framework.permissions import AllowAny
from rest_framework.views import APIView

from mpm.MpmController import *

motor_obj = MpmController()


# Create your views here.

class AddActiveMotorAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def post(self, request):
        print(request.data)

        result = motor_obj.AddActiveMotor(request.data)
        return result


class AddLotIdsAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def post(self, request):
        print(request.data)

        result = motor_obj.AddLotIds(request.data)
        return result


class GetLotAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def get(self, request):

        result = motor_obj.GetLots(request)
        return result


# class EditActiveMotorAPIVIEW(APIView):
#     permission_classes = [AllowAny]
#     def put(self, request):
#         result = motor_obj.EditActiveMotor(request.data)
#         print(result)
#         return result
class AddLotIdsAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def post(self, request):
        print(request.data)

        result = motor_obj.AddLotIds(request.data)
        return result


class GetLotAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def get(self, request):

        result = motor_obj.GetLots(request)
        return result
class DeletetActiveMotorAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def get(self, request):
        result = motor_obj.DeleteActiveMotor(request)
        return result


class GetActiveMotorListAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def get(self, request):
        result = motor_obj.GetActiveMotorList(request)
        return result


class GetActiveMotorListCountAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def get(self, request):
        result = motor_obj.getProcessMonitoringDashboardCount(request)
        return result

class GetActiveMotorHistoryAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def get(self, request):
        result = motor_obj.GetActiveMotorHistoryList(request)
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
