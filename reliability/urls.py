from django.urls import path,re_path
from .views import *

urlpatterns = [
        path('addReliability/', AddReliabilityAPIVIEW.as_view(), name='addReliability'),
        path('getReliabilityDashboardCount/', GetReliabilityDashboardCountAPIVIEW.as_view(), name='getReliabilityDashboardCount'),
        path('getReliabilityList/', GetReliabilityListAPIVIEW.as_view(), name='getReliabilityList'),
        path('deleteReliability/', DeleteReliabilityAPIVIEW.as_view(), name='deleteReliability'),
        path('getReliabilityHistory/', GetReliabilityHistoryAPIVIEW.as_view(), name='getReliabilityHistory'),
]