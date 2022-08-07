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


class docTracingListAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def get(self, request):
        result = doc_obj.GetDocumentList(request)
        return result


class docPendingListAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def get(self, request):
        result = doc_obj.GetPendingDocumentList(request)
        return result

# Waqar==============================
class GetDocumentListAPIVIEW(APIView):
    permission_classes = [AllowAny]

    def get(self, request):
        result = doc_obj.GetDocumentList(request)
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

class GetDocumentsEmailAPIVIEW(APIView):
    permission_classes = [AllowAny]
    def get(self, request):
        result = doc_obj.GetDocumentEmailList(request)
        return result