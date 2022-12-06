from django.urls import path,re_path
from .views import *

urlpatterns = [
        path('addmotor/', AddActiveMotorAPIVIEW.as_view(), name='addmotor'),
        path('editmotor/', AddActiveMotorAPIVIEW.as_view(), name='editmotor'),
        path('motorlist/', GetActiveMotorListAPIVIEW.as_view(), name='motorlist'),
        re_path(r'^deletemotor/(?P<pk>[0-9]+)$',DeletetActiveMotorAPIVIEW.as_view(), name='deletemotor'),
        path('getpdf/', GetActiveMotorPDFAPIVIEW.as_view(), name='getpdf'),
        path('getexcel/', GetActiveMotorExcelAPIVIEW.as_view(), name='getexcel'),


]