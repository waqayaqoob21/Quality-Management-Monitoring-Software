from django.urls import path,re_path
from .views import *

urlpatterns = [
        path('addSqa/', AddSqaAPIVIEW.as_view(), name='addSqa'),
        path('getSqaDashboardCount/', GetSqaDashbaordCountAPIVIEW.as_view(), name='getSqaDashboardCount'),
        path('getSqaList/', GetSqaListAPIVIEW.as_view(), name='getSqaList'),
        path('deleteSqa/', DeleteSqaAPIVIEW.as_view(), name='deleteSqa'),
]