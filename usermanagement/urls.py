"""nescServices URL Configuration

The `urlpatterns` list routes URLs to views. For more information please see:
    https://docs.djangoproject.com/en/4.0/topics/http/urls/
Examples:
Function views
    1. Add an import:  from my_app import views
    2. Add a URL to urlpatterns:  path('', views.home, name='home')
Class-based views
    1. Add an import:  from other_app.views import Home
    2. Add a URL to urlpatterns:  path('', Home.as_view(), name='home')
Including another URLconf
    1. Import the include() function: from django.urls import include, path
    2. Add a URL to urlpatterns:  path('blog/', include('blog.urls'))
"""
# from django.conf.urls import url
from django.contrib import admin
from django.urls import path
from django.urls import path, include

from usermanagement.views import *

urlpatterns = [
    path('UserLogin/', UserLoginApiView.as_view(), name='UserLogin'),
    path('adduser/', AddUserAPIView.as_view(), name='adduser'),
    path('updateuser/', AddUserAPIView.as_view(), name='updateuser'),
    path('getuser/', GetUserListAPIVIEW.as_view(), name='getuser'),
    path('deleteuser/', DeleteUserAPIVIEW.as_view(), name='deleteuser'),
    path('usermotor/', GetUserMotorListAPIVIEW.as_view(), name='usermotor'),
    path('userprod/', GetUserProductionListAPIVIEW.as_view(), name='userprod'),
    path('userrelif/', GetUserRelifingListAPIVIEW.as_view(), name='userrelif'),
    path('userflight/', GetUserFlightListAPIVIEW.as_view(), name='userflight'),
    path('usercsv/', AddUserCSVAPIView.as_view(), name='usercsv'),

]
