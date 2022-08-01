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
from django.conf.urls import url
from django.contrib import admin
from django.urls import path
from django.urls import path, include

from doctracking.views import AddDocAPIVIEW, docTracingListAPIVIEW, docPendingListAPIVIEW,GetDocumentListAPIVIEW, GetDocumentsPDFAPIVIEW
from usermanagement.views import UserLoginApiView

urlpatterns = [
    path('addDoc/', AddDocAPIVIEW.as_view(), name='addDoc'),
    path('docTracingList/', docTracingListAPIVIEW.as_view(), name='docTracingList'),
    path('docPendingList/', docPendingListAPIVIEW.as_view(), name='docPendingList'),
    path('documentList/', GetDocumentListAPIVIEW.as_view(), name='documentList'),
    path('documentpdf/', GetDocumentsPDFAPIVIEW.as_view(), name='documentpdf'),


]
