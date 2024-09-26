from django.urls import path
from .views import *

urlpatterns = [
    path('sendMessage/', SendMessageAPIView.as_view(), name='sendMessage'),
    path('fetchMessages/', FetchMessagesAPIView.as_view(), name='fetchMessages'),
    path('users/', GetUsersListAPIView.as_view(), name='users'),
]
