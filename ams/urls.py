from django.urls import path
from .views import *

urlpatterns = [
        path('addtask/', AddTaskAPIVIEW.as_view(), name='addtask'),
        path('edittask/', EidtTaskAPIVIEW.as_view(), name='edittask'),
        path('tasklist/', GetTaskListAPIVIEW.as_view(), name='tasklist'),
]