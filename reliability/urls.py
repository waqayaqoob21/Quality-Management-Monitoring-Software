from django.urls import path,re_path
from .views import *

urlpatterns = [
        path('addReliability/', AddReliabilityAPIVIEW.as_view(), name='addReliability'),
        path('getReliabilityList/', GetReliabilityListAPIVIEW.as_view(), name='getReliabilityList'),
]