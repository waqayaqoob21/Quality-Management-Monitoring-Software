from django.urls import path,re_path
from .views import *

urlpatterns = [
        path('addEmc/', AddEmcAPIVIEW.as_view(), name='addEmc'),
        path('getEmcList/', GetEmcListAPIVIEW.as_view(), name='getEmcList'),
        path('getProduct/', GetProductAPIVIEW.as_view(), name='getProduct'),
        path('deleteEmc/', DeleteEmcAPIVIEW.as_view(), name='deleteEmc'),
]