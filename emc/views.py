from rest_framework.permissions import AllowAny
from rest_framework.views import APIView

from .emcController import *

emc_obj = emcController()
# Create your views here.


class GetProductAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def get(self, request):
        print(request.data)
        result = emc_obj.getProduct(request)
        return result
class AddEmcAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def post(self, request):
        print(request.data)
        result = emc_obj.addEmc(request.data)
        return result

class GetEmcDashboardCountAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def get(self, request):
        print(request.data)
        result = emc_obj.getEmcDashboardCount(request)
        return result
class GetEmcListAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def get(self, request):
        print(request.data)
        result = emc_obj.getEmcList(request)
        return result
    
class DeleteEmcAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def delete(self, request):
        result = emc_obj.deleteEmc(request)
        return result

class GetEmcHistoryAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def get(self, request):
        print(request.data)
        result = emc_obj.getEmcHistory(request)
        return result