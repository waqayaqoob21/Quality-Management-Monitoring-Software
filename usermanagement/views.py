# Login API
from rest_framework_simplejwt.views import TokenObtainPairView
from usermanagement.serializer import MyTokenObtainPairSerializer
from rest_framework.permissions import AllowAny
from rest_framework.views import APIView

from usermanagement.UserController import *

user_obj = UserController()

class AddUserAPIView(APIView):
    permission_classes = [AllowAny]
    def post(self, request):
        result = user_obj.AddUser(request.data)
        return result

class GetUserListAPIVIEW(APIView):
    permission_classes = [AllowAny]
    def get(self, request):
        result = user_obj.GetUserList(request)
        return result

class DeleteUserAPIVIEW(APIView):
    permission_classes = [AllowAny]
    def get(self, request):
        result = user_obj.DeleteUser(request)
        return result


class GetUserMotorListAPIVIEW(APIView):
    permission_classes = [AllowAny]
    def get(self, request):
        result = user_obj.getUserMotorList(request)
        return result

class GetUserProductionListAPIVIEW(APIView):
    permission_classes = [AllowAny]
    def get(self, request):
        result = user_obj.getUserProductionList(request)
        return result
class GetUserRelifingListAPIVIEW(APIView):
    permission_classes = [AllowAny]
    def get(self, request):
        result = user_obj.getUserRelifingList(request)
        return result
class GetUserFlightListAPIVIEW(APIView):
    permission_classes = [AllowAny]
    def get(self, request):
        result = user_obj.getUserFlightList(request)
        return result

  # ================================== User Login===============================================
class UserLoginApiView(TokenObtainPairView):
    serializer_class = MyTokenObtainPairSerializer


    def post(self, request, *args, **kwargs):

        try:
            print("User Login Api Called")
            serializer = self.get_serializer(data=request.data)

            if serializer.is_valid():
                id = str(serializer.user.id)
                userRolesData = UserRoles.objects.filter(user_id = id).first()
                if userRolesData is not None:
                    user_data = {'id': serializer.user.id, 'username': serializer.user.username,
                                 'email': serializer.user.email, 'first_name': serializer.user.first_name,
                                 'last_name': serializer.user.last_name, 'is_superuser': serializer.user.is_superuser,
                                 'prod_roles': userRolesData.prod_roles, 'relif_roles': userRolesData.relif_roles,
                                 'flight_roles': userRolesData.flight_roles, 'motor_roles': userRolesData.motor_roles
                                 }
                    return JsonResponse(
                        {'message': "Login Successfully", 'Token': serializer.validated_data,
                         'data': user_data}, status=200)
                else:
                    user_data = {'id': serializer.user.id, 'username': serializer.user.username,
                                 'email': serializer.user.email, 'first_name': serializer.user.first_name,
                                 'last_name': serializer.user.last_name, 'is_superuser': serializer.user.is_superuser
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
