from django.shortcuts import render

# Create your views here.
from rest_framework.permissions import AllowAny
from rest_framework.views import APIView

from sms.SmsController import *

task_obj = SmsController()


# Create your views here.

class AddProductionAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def post(self, request):
        result = task_obj.AddProductionStatus(request.data)
        return result

class ImportProductionCsvAPIVIEW(APIView):
    permission_classes = [AllowAny]
    def post(self, request):
        result = task_obj.ImportProductionCsv(request.data)
        return result
class ImportFlightCsvAPIVIEW(APIView):
    permission_classes = [AllowAny]
    def post(self, request):
        result = task_obj.ImportFlightCsv(request.data)
        return result
class ImportRelifingCsvAPIVIEW(APIView):
    permission_classes = [AllowAny]
    def post(self, request):
        result = task_obj.ImportRelifingCsv(request.data)
        return result
class EditProductionAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def put(self, request):
        result = task_obj.AddProductionStatus(request.data)
        print(result)
        return result

class GetTotalSystemsAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def get(self, request):
        result = task_obj.GetTotalSystems(request)
        return result


class GetProductionListAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def get(self, request):
        result = task_obj.GetProductionList(request)
        return result


class GetProductionObservationStatusAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def get(self, request):
        result = task_obj.GetProductionObservationStatus(request)
        return result

class GetProductionDateNoneListAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def get(self, request):
        result = task_obj.GetProdctionDateNone(request)
        return result

class GetFlightDateNoneListAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def get(self, request):
        result = task_obj.GetFlightDateNone(request)
        return result

class GetRelifingDateNoneListAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def get(self, request):
        result = task_obj.GetRelifingDateNone(request)
        return result
class GetProductionHistoryListAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def get(self, request):
        result = task_obj.GetProductionListHistory(request)
        return result


class GetProductionPDFAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def get(self, request):
        result = task_obj.GetProductionPDFList(request)
        return result


class GetProductionExcelAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def get(self, request):
        result = task_obj.GetProductionExcelList(request)
        return result


class DeleteProductionAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def get(self, request):
        result = task_obj.DeleteProdSys(request)
        return result

class DeleteProductionHistoryAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def get(self, request):
        result = task_obj.DeleteProdSysHistory(request)
        return result
# ===================================Flight System=======================
# ===================================Flight System=======================
# ===================================Flight System=======================

class AddFlightAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def post(self, request):
        result = task_obj.AddFlightStatus(request.data)
        return result


class EditFlightAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def put(self, request):
        result = task_obj.AddFlightStatus(request.data)
        print(result)
        return result


class GetFlightListAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def get(self, request):
        result = task_obj.GetFlightList(request)
        return result


class flightlistHistoryAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def get(self, request):
        result = task_obj.GetFlightListHistory(request)
        return result


class GetFlightPDFAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def get(self, request):
        result = task_obj.GetFlightPDFList(request)
        return result


class GetFlightExcelAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def get(self, request):
        result = task_obj.GetFlightExcelList(request)
        return result


class DeleteFlightAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def get(self, request):
        result = task_obj.DeleteFlightSys(request)
        return result

class DeleteFlightHistoryAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def get(self, request):
        result = task_obj.DeleteFlightSysHistory(request)
        return result
# ===================================Relfing System=======================
# ===================================Relfing System=======================
# ===================================Relfing System=======================

class AddRelifingAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def post(self, request):
        result = task_obj.AddRelifingStatus(request.data)
        return result


class EditRelifingAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def put(self, request):
        result = task_obj.AddRelifingStatus(request.data)
        print(result)
        return result


class GetRelifingListAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def get(self, request):
        result = task_obj.GetRelifingList(request)
        return result


class GetRelifingHistoryListAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def get(self, request):
        result = task_obj.GetRelifingHistoryList(request)
        return result


class GetRelifingPDFAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def get(self, request):
        result = task_obj.GetRelifingPDFList(request)
        return result


class GetRelifingExcelAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def get(self, request):
        result = task_obj.GetRelifingExcelList(request)
        return result

class DeleteRelifingAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def get(self, request):
        result = task_obj.DeleteRelifingSys(request)
        return result

class DeleteRelifingHistoryAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def get(self, request):
        result = task_obj.DeleteRelifingSysHistory(request)
        return result
class getSystemMonitoringDashboardCountAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def get(self, request):
        result = task_obj.getSystemMonitoringDashboardCount(request)
        return result

