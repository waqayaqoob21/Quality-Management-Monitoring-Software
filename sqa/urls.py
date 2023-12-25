from django.urls import path,re_path
from .views import *

urlpatterns = [
        path('addSqa/', AddSqaAPIVIEW.as_view(), name='addSqa'),
        path('getSqaList/', GetSqaListAPIVIEW.as_view(), name='getSqaList'),
]