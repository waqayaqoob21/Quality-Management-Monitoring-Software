from django.urls import path,re_path
from .views import *

urlpatterns = [
        path('addPower/', AddPowerAPIVIEW.as_view(), name='addPower'),
        path('getPowerDashboardCount/', GetPowerDashboardCountAPIVIEW.as_view(), name='getPowerDashboardCount'),
        path('getPowerList/', GetPowerListAPIVIEW.as_view(), name='getPowerList'),
        path('getPowerToBeUpdated/', GetPowerToBeUpdatedAPIVIEW.as_view(), name='getPowerToBeUpdated'),
        path('deletePower/', DeletePowerAPIVIEW.as_view(), name='deletePower'),
        path('getPowerHistory/', GetPowerHistoryAPIVIEW.as_view(), name='getPowerHistory'),


]