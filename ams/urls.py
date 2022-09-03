from django.urls import path
from .views import *

urlpatterns = [
        path('addtask/', AddTaskAPIVIEW.as_view(), name='addtask'),
        path('edittask/', EidtTaskAPIVIEW.as_view(), name='edittask'),
        path('tasklist/', GetTaskListAPIVIEW.as_view(), name='tasklist'),
        path('getpdf/', GetDocumentPDFAPIVIEW.as_view(), name='getpdf'),
        path('getexcel/', GetDocumentExcelAPIVIEW.as_view(), name='getexcel'),
        path('ocr/', OcrPdfAPIVIEW.as_view(), name='ocr'),
        path('getocrdata/', GetOcrDataAPIVIEW.as_view(), name='getocrdata'),


]