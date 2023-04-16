from django.urls import path
from .views import *

urlpatterns = [
  

        path('addprod/', AddProductionAPIVIEW.as_view(), name='addprod'),
        path('editprod/', EditProductionAPIVIEW.as_view(), name='editprod'),
        path('prodlist/', GetProductionListAPIVIEW.as_view(), name='prodlist'),
        path('prodlistHistory/', GetProductionHistoryListAPIVIEW.as_view(), name='prodlistHistory'),
        path('getprodpdf/', GetProductionPDFAPIVIEW.as_view(), name='getprodpdf'),
        path('getprodexl/', GetProductionExcelAPIVIEW.as_view(), name='getprodexl'),
        path('deleteprodsys/', DeleteProductionAPIVIEW.as_view(), name='deleteprodsys'),



        path('addflight/', AddFlightAPIVIEW.as_view(), name='addflight'),
        path('editflight/', EditFlightAPIVIEW.as_view(), name='editflight'),
        path('flightlist/', GetFlightListAPIVIEW.as_view(), name='flightlist'),
        path('flightlistHistory/', flightlistHistoryAPIVIEW.as_view(), name='flightlistHistory'),
        path('getflightpdf/', GetFlightPDFAPIVIEW.as_view(), name='getflightpdf'),
        path('getflightexl/', GetFlightExcelAPIVIEW.as_view(), name='getflightexl'),
        path('deleteflightsys/', DeleteFlightAPIVIEW.as_view(), name='deleteflightsys'),

        path('addrelifing/', AddRelifingAPIVIEW.as_view(), name='addrelifing'),
        path('editrelifing/', EditRelifingAPIVIEW.as_view(), name='editrelifing'),
        path('relifinglist/', GetRelifingListAPIVIEW.as_view(), name='relifinglist'),
        path('relifinglistHistory/', GetRelifingHistoryListAPIVIEW.as_view(), name='relifinglistHistory'),
        path('getrelifingpdf/', GetRelifingPDFAPIVIEW.as_view(), name='getrelifingpdf'),
        path('getrelifingexl/', GetRelifingExcelAPIVIEW.as_view(), name='getrelifingexl'),
        path('deleterelifing/', DeleteRelifingAPIVIEW.as_view(), name='deleterelifing'),

        path('getSystemMonitoringDashboardCount/', getSystemMonitoringDashboardCountAPIVIEW.as_view(), name='getSystemMonitoringDashboardCount'),

        path('prodatenone/', GetProductionDateNoneListAPIVIEW.as_view(), name='prodatenone'),
        path('flightdatenone/', GetFlightDateNoneListAPIVIEW.as_view(), name='flightdatenone'),
        path('relifdatenone/', GetRelifingDateNoneListAPIVIEW.as_view(), name='relifdatenone'),

        path('importprodcsv/', ImportProductionCsvAPIVIEW.as_view(), name='importprodcsv'),
        path('importflightcsv/', ImportFlightCsvAPIVIEW.as_view(), name='importflightcsv'),
        path('importrelifcsv/', ImportRelifingCsvAPIVIEW.as_view(), name='importrelifcsv'),

        path('importmotorcsv/', ImportProductionCsvAPIVIEW.as_view(), name='importmotorcsv'),
        path('importbatteriescsv/', ImportProductionCsvAPIVIEW.as_view(), name='importbatteriescsv'),
        path('importpyrocsv/', ImportProductionCsvAPIVIEW.as_view(), name='importpyrocsv'),

]
