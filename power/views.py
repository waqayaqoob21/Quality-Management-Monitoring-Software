from rest_framework.permissions import AllowAny
from rest_framework.views import APIView

from .powerController import *

powr_obj = powerController()


# Create your views here.


class GetPowerToBeUpdatedAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def get(self, request):
        print(request.data)
        result = powr_obj.getPowerToBeUpdated(request)
        return result


class AddPowerAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def post(self, request):
        print(request.data)
        result = powr_obj.addPower(request.data)
        return result


class GetPowerDashboardCountAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def get(self, request):
        print(request.data)
        result = powr_obj.getPowerDashboardCount(request)
        return result


class GetPowerListAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def get(self, request):
        print(request.data)
        result = powr_obj.getPowerList(request)
        return result


class DeletePowerAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def delete(self, request):
        result = powr_obj.deletePower(request)
        return result


class GetPowerHistoryAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def get(self, request):
        print(request.data)
        result = powr_obj.getPowerHistory(request)
        return result