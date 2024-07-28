from rest_framework.permissions import AllowAny
from rest_framework.views import APIView

from .reliabilityController import *

reliability_obj = reliabilityController()
# Create your views here.
class GetReliabilityToBeUpdatedAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def get(self, request):
        print(request.data)
        result = reliability_obj.getReliabilityToBeUpdated(request)
        return result
class AddReliabilityAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def post(self, request):
        print(request.data)
        result = reliability_obj.addReliability(request.data)
        return result


class GetReliabilityDashboardCountAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def get(self, request):
        print(request.data)
        result = reliability_obj.getReliabilityDashboardCount(request)
        return result


class GetReliabilityListAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def get(self, request):
        print(request.data)
        result = reliability_obj.getReliabilityList(request)
        return result


class DeleteReliabilityAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def delete(self, request):
        result = reliability_obj.deleteReliability(request)
        return result


class GetReliabilityHistoryAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def get(self, request):
        print(request.data)
        result = reliability_obj.getReliabilityHistory(request)
        return result
