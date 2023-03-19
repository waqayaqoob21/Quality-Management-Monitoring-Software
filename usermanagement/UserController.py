from django.http import JsonResponse
from django.contrib.auth.models import User
from datetime import date
from usermanagement.serializer import *
from django.db import connection
from django.contrib.auth.hashers import make_password, check_password


class UserController:

    @staticmethod
    def AddUser(request):
        userModel = User()

        try:
            uname = request['username']
            id = request['id']
            if id == '0':
                userModel.first_name = request['first_name']
                userModel.last_name = request['last_name']
                userModel.username = request['username']
                userModel.email = 'nescom@nescom.com'
                encryptedpassword = make_password(request['password'])
                userModel.password = encryptedpassword
                userModel.is_superuser = 'False'
                userModel.last_login = date.today()
                userModel.is_active = 'True'
                userModel.is_staff = 'True'
                userModel.save()
                username = User.objects.filter(username = uname).first()
                if username != '':
                    rolesModel = UserRoles()
                    rolesModel.user_id = username.id
                    rolesModel.prod_roles = request['prod_roles']
                    rolesModel.relif_roles = request['relif_roles']
                    rolesModel.flight_roles = request['flight_roles']
                    rolesModel.motor_roles = request['motor_roles']
                    rolesModel.save()
                return JsonResponse({'masssage': 'User Added Successfully!'},status=200)
            else:
                userModel = User.objects.filter(id=id).first()
                userModel.first_name = request['first_name']
                userModel.last_name = request['last_name']
                userModel.username = request['username']
                userModel.email = 'nescom@nescom.com'
                encryptedpassword = make_password(request['password'])
                userModel.password = encryptedpassword
                userModel.is_superuser = 'False'
                userModel.last_login = date.today()
                userModel.is_active = 'True'
                userModel.is_staff = 'True'
                userModel.save()
                u_id = str(id)
                userRole = UserRoles.objects.filter(user_id = u_id).first()
                if userRole is not None:
                    # userRoleModel = UserRoles()
                    userRole.prod_roles = request['prod_roles']
                    userRole.relif_roles = request['relif_roles']
                    userRole.flight_roles = request['flight_roles']
                    userRole.motor_roles = request['motor_roles']
                    userRole.save()
                return JsonResponse({'masssage': 'User Updated Successfully!'}, status=200)
        except:
            return JsonResponse({'masssage': 'User registration Failed!'}, status=200)


    @staticmethod
    def GetUserList(request, self=None):
        dataList = []
        try:
            data = User.objects
            doc_list = []
            cursor = connection.cursor()
            query = "SELECT au.id,au.first_name,au.last_name,au.username,au.password, "\
                    "ur.prod_roles,ur.flight_roles,ur.relif_roles,ur.motor_roles  " \
                    "FROM public.auth_user au " \
                    "FULL OUTER JOIN "\
                    "public.usermanagement_userroles ur ON CAST(ur.user_id AS INTEGER) = au.id;"

            cursor.execute(query)
            col_names = [col[0] for col in cursor.description]
            for row in cursor.fetchall():
                row_dict = dict(zip(col_names, row))
                doc_list.append(row_dict)
            # serializer = UserSerializer(data, many=True)
            return JsonResponse({'status': 'True', 'data': doc_list},
                                status=200)
        except:
            return JsonResponse({'massage', 'No record found!'}, status=201)