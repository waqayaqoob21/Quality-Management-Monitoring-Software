from django.shortcuts import render

# Create your views here.
from rest_framework.permissions import AllowAny
from rest_framework.views import APIView

from doctracking.DocController import DocController

doc_obj = DocController()


class AddDocAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def post(self, request):
        result = doc_obj.AddDocument(request.data)
        return result


class DeleteDocAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def get(self, request):
        result = doc_obj.DeleteDocument(request)
        return result

class DeleteCertificateAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def get(self, request):
        result = doc_obj.DeleteCertificate(request)
        return result

class DeleteTaskAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def get(self, request):
        result = doc_obj.DeleteTask(request)
        return result


class docTracingListAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def get(self, request):
        result = doc_obj.GetDocumentList(request)
        return result


class docTracingHistoryAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def get(self, request):
        result = doc_obj.GetDocumentHistory(request)
        return result


class docPendingListAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def get(self, request):
        result = doc_obj.GetPendingDocumentList(request)
        return result


class GetDocumentsPDFAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def get(self, request):
        result = doc_obj.GetDocumentPDFList(request)
        return result


class GetDocumentsExcelAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def get(self, request):
        result = doc_obj.GetDocumentExcelList(request)
        return result

class GenerateCertificateAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def post(self, request):
        result = doc_obj.generateCertificate(request.data)
        return result

class certificateListAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def get(self, request):
         result = doc_obj.GetCertificateList(request)
         return result

class GetCertificateSerialNumberAPIVIEW(APIView):
        permission_classes = [AllowAny]
        def get(self, request):
            result = doc_obj.GetCertificateSerialNumber(request)
            return result
        
class GetDashboardAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def get(self, request):
        result = doc_obj.GetDashboardCount(request)
        return result


class GetDocTrackingDashboardCountAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def get(self, request):
        result = doc_obj.GetDocTrackingDashboardCount(request)
        return result


class getTaskMonitoringDashboardCountAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def get(self, request):
        result = doc_obj.getTaskMonitoringDashboardCount(request)
        return result

class DeleteDocumentHistoryAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def get(self, request):
        result = doc_obj.deleteDocumentHistory(request)
        return result