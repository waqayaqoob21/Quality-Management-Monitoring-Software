from django.contrib.auth import authenticate, login, logout
from django.http import JsonResponse
from rest_framework.response import Response
from rest_framework import status
from .serializers import *
from .models import *
from usermanagement.serializer import *

class chatContoller:

    @staticmethod
    def userList(request):
        try:
            users = User.objects.all()
            serializer = UserSerializer(users, many=True)
            return JsonResponse({'success': True, 'Message':'Message sent successfully','data': serializer.data,'status':'200'},status=200)
        except Exception as e:
            print(e)
            return  JsonResponse({'success': False, 'Message':'Message could not send','status':'500'},status=500)

    @staticmethod
    def sendMessage(request):
        try:
            msgModel = Message()
            msgModel.user = request['sender']
            msgModel.receiver = request['receiver']
            msgModel.content = request['message']
            msgModel.save()
            latestMessage = Message.objects.last()
            serializer = MessageSerializer(latestMessage)
            return JsonResponse({'success': True, 'Message':'Message sent successfully','data': serializer.data,'status':'200'},status=200)
        except Exception as e:
            print(e)
            return  JsonResponse({'success': False, 'Message':'Message could not send','status':'500'},status=500)

    @staticmethod
    def fetchMessage(request):
        try:
            sender = request.query_params['sender']
            receiver = request.query_params['receiver']
            if receiver != '':
                chat = Message.objects.filter(receiver=receiver)
                serializer = MessageSerializer(chat, many=True)
                return JsonResponse(
                    {'success': True, 'Message': 'Message sent successfully', 'data': serializer.data, 'status': '200'},
                    status=200)
            else:
                return JsonResponse(
                    {'success': True, 'Message': 'No user selected', 'data': [], 'status': '200'},
                    status=200)
        except Exception as e:
            print(e)
            return  JsonResponse({'success': False, 'Message':'Message could not send','status':'500'},status=500)

    @staticmethod
    def logout(request):
        try:
            logout(request)  # This logs out the user
            return Response({"message": "Logout successful!"}, status=status.HTTP_200_OK)
        except Exception as e:
            print(e)
            return  JsonResponse({'success': False, 'Message':'Could not logout','status':'500'},status=500)