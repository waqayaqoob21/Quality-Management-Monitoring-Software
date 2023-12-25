from django.urls import path,re_path
from .views import *

urlpatterns = [
        path('addEmsEs/', AddEmsEsAPIVIEW.as_view(), name='addEmsEs'),
        path('getEmsList/', GetEmsEsListAPIVIEW.as_view(), name='getEmsList'),
]