from rest_framework_simplejwt.views import TokenObtainPairView
from rest_framework.permissions import AllowAny

from django.http import JsonResponse
from django.dispatch import receiver
from rest_framework.response import Response
from rest_framework.views import APIView
from rest_framework import status
from .serializers import *
from .models import *
from .chatController import *

chat_obj = chatContoller()


class GetUsersListAPIView(APIView):
    permission_classes = [AllowAny]

    def get(self, request):
        result = chat_obj.userList(request)
        return result

class SendMessageAPIView(APIView):
    permission_classes = [AllowAny]

    def post(self, request):
        result = chat_obj.sendMessage(request.data)
        return result

class FetchMessagesAPIView(APIView):
    permission_classes = [AllowAny]

    def get(self, request):
        result = chat_obj.fetchMessage(request)
        return result