from django.urls import path,re_path
from .views import *

urlpatterns = [
        path('addqms/', AddQmsAuditAPIVIEW.as_view(), name='addqms'),
        path('editqms/', AddQmsAuditAPIVIEW.as_view(), name='editqms'),
        path('qmslist/', GetQmsAuditListAPIVIEW.as_view(), name='qmslist'),
        re_path(r'^deleteqms/(?P<pk>[0-9]+)$',DeletetQmsAuditAPIVIEW.as_view(), name='deleteqms'),


        path('addcesp/', AddCespAuditAPIVIEW.as_view(), name='addcesp'),
        path('editcesp/', AddCespAuditAPIVIEW.as_view(), name='editcesp'),
        path('cesplist/', GetCespAuditListAPIVIEW.as_view(), name='cesplist'),
        re_path(r'^deletecesp/(?P<pk>[0-9]+)$',DeletetCespAuditAPIVIEW.as_view(), name='deletecesp'),
   
]