from django.urls import path,re_path
from .views import *

urlpatterns = [
        path('addSqa/', AddSqaAPIVIEW.as_view(), name='addSqa'),
        path('getSqaDashboardCount/', GetSqaDashbaordCountAPIVIEW.as_view(), name='getSqaDashboardCount'),
        path('getSqaList/', GetSqaListAPIVIEW.as_view(), name='getSqaList'),
        path('deleteSqa/', DeleteSqaAPIVIEW.as_view(), name='deleteSqa'),
        path('getSqaHistory/', GetSqaHistoryAPIVIEW.as_view(), name='getSqaHistory'),
        path('getSqa/', GetSqaAPIVIEW.as_view(), name='getSqa'),
        path('getModuleNameIDList/', GetModuleNameIDListAPIVIEW.as_view(), name='getModuleNameIDList'),
        path('getSqaSummary/', GetSqaSummaryAPIVIEW.as_view(), name='getSqaSummary'),
]