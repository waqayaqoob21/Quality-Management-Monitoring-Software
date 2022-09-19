from django.urls import path
from .views import *

urlpatterns = [
  

        path('addprod/', AddProductionAPIVIEW.as_view(), name='addprod'),
        path('editprod/', EditProductionAPIVIEW.as_view(), name='editprod'),
        path('prodlist/', GetProductionListAPIVIEW.as_view(), name='prodlist'),
        path('getprodpdf/', GetProductionPDFAPIVIEW.as_view(), name='getprodpdf'),
        path('getprodexl/', GetProductionExcelAPIVIEW.as_view(), name='getprodexl'),


        path('addflight/', AddFlightAPIVIEW.as_view(), name='addflight'),
        path('editflight/', EditFlightAPIVIEW.as_view(), name='editflight'),
        path('flightlist/', GetFlightListAPIVIEW.as_view(), name='flightlist'),
        path('getflightpdf/', GetFlightPDFAPIVIEW.as_view(), name='getflightpdf'),
        path('getflightexl/', GetFlightExcelAPIVIEW.as_view(), name='getflightexl'),


        path('addrelifing/', AddRelifingAPIVIEW.as_view(), name='addrelifing'),
        path('editrelifing/', EditRelifingAPIVIEW.as_view(), name='editrelifing'),
        path('relifinglist/', GetRelifingListAPIVIEW.as_view(), name='relifinglist'),
        path('getrelifingpdf/', GetRelifingPDFAPIVIEW.as_view(), name='getrelifingpdf'),
        path('getrelifingexl/', GetRelifingExcelAPIVIEW.as_view(), name='getrelifingexl'),

]