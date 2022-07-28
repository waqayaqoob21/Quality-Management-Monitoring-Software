from django.http import JsonResponse
from django.shortcuts import render

# Create your views here.


# Login API
from rest_framework_simplejwt.views import TokenObtainPairView

from usermanagement.serializer import MyTokenObtainPairSerializer


class UserLoginApiView(TokenObtainPairView):
    serializer_class = MyTokenObtainPairSerializer

    def post(self, request, *args, **kwargs):

        try:
            print("User Login Api Called")
            serializer = self.get_serializer(data=request.data)

            if serializer.is_valid():

                user_data = {'id': serializer.user.id, 'username': serializer.user.username,
                             'email': serializer.user.email, 'first_name': serializer.user.first_name,
                             'last_name': serializer.user.last_name, 'is_superuser': serializer.user.is_superuser,
                             }
                return JsonResponse(
                    {'message': "Login Successfully", 'Token': serializer.validated_data,
                     'data': user_data}, status=200)

            else:
                return JsonResponse({'message': "Invalid username or password"}, status=401)
        except Exception as e:
            print(e)
            return JsonResponse({'message': "Invalid username or password"}, status=401)
            pass
