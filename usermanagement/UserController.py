from django.http import JsonResponse
from .models import User
from datetime import date
from usermanagement.serializer import *
from django.db import connection
from django.contrib.auth.hashers import make_password, check_password
from sms.models import  *
from sms.serializers import *
from mpm.models import *
from mpm.serializer import *
from django.db.models import F, Q
from csv import reader
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
                userModel.user_type = request['user_type']
                userModel.email = 'nescom@nescom.com'
                encryptedpassword = make_password(request['password'])
                userModel.password = encryptedpassword
                userModel.is_superuser = 'False'
                userModel.last_login = date.today()
                userModel.is_active = 'True'
                userModel.is_staff = 'True'
                userModel.save()
                username = User.objects.filter(username = uname).first()
                user_roles = UserRoles.objects.filter(user_id = username.id).first()
                if username != '':
                    rolesModel = UserRoles()
                    if user_roles is None:
                        rolesModel.user_id = username.id
                        rolesModel.prod_roles = request['prod_roles']
                        rolesModel.relif_roles = request['relif_roles']
                        rolesModel.flight_roles = request['flight_roles']
                        rolesModel.motor_roles = request['motor_roles']
                        rolesModel.battery_roles = request['battery_roles']
                        rolesModel.pyro_roles = request['pyro_roles']
                        rolesModel.bhd_roles = request['bhd_roles']
                        rolesModel.task_roles = request['task_roles']
                        rolesModel.qms_roles = request['qms_roles']
                        rolesModel.cesp_roles = request['cesp_roles']
                        rolesModel.save()
                return JsonResponse({'masssage': 'User Added Successfully!'},status=200)
            else:
                userModel = User.objects.filter(id=id).first()
                userModel.first_name = request['first_name']
                userModel.last_name = request['last_name']
                userModel.username = request['username']
                userModel.user_type = request['user_type']
                userModel.email = 'nescom@nescom.com'
                # encryptedpassword = make_password(request['password'])
                # userModel.password = encryptedpassword
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
                    userRole.battery_roles = request['battery_roles']
                    userRole.pyro_roles = request['pyro_roles']
                    userRole.bhd_roles = request['bhd_roles']
                    userRole.task_roles = request['task_roles']
                    userRole.qms_roles = request['qms_roles']
                    userRole.cesp_roles = request['cesp_roles']
                    userRole.save()
                return JsonResponse({'masssage': 'User Updated Successfully!'}, status=200)
        except:
            return JsonResponse({'masssage': 'User registration Failed!'}, status=200)


    @staticmethod
    def GetUserList(request, self=None):
        dataList = []
        try:
            doc_list = []
            cursor = connection.cursor()
            query = "SELECT au.id,au.first_name,au.last_name,au.username,au.password,au.user_type, "\
                    "ur.prod_roles,ur.flight_roles,ur.relif_roles,ur.motor_roles,ur.battery_roles,ur.pyro_roles,bhd_roles,task_roles,qms_roles,cesp_roles  " \
                    "FROM public.usermanagement_user au " \
                    "FULL OUTER JOIN "\
                    "public.usermanagement_userroles ur ON CAST(ur.user_id AS INTEGER) = au.id ORDER BY au.id DESC;"

            cursor.execute(query)
            col_names = [col[0] for col in cursor.description]
            for row in cursor.fetchall():
                row_dict = dict(zip(col_names, row))
                doc_list.append(row_dict)
            return JsonResponse({'status': 'True', 'data': doc_list},
                                status=200)
        except:
            return JsonResponse({'massage', 'No record found!'}, status=201)
    @staticmethod
    def DeleteUser(request):
        # try:
            userId = request.query_params['id']
            if userId != '':
                user = User.objects.get(id=userId)
                user.delete()
            if userId != '':
                role = UserRoles.objects.get(user_id = userId)
                role.delete()
            return JsonResponse({'message': 'User has been deleted'}, status=200)
        # except:
        #     return JsonResponse({'message': 'Sorry! No Audit found.'}, status=500)

    @staticmethod
    def getUserMotorList(request):
        try:
            userId = request.query_params['id']
            current_comp = request.query_params['component']
            comp = current_comp.split('[')
            component = comp[0]
            def get_filter(field_name, filter_condition, filter_value):
                if filter_condition.strip() == "contains":
                    kwargs = {
                        '{0}__icontains'.format(field_name): filter_value
                    }
                    return Q(**kwargs)

                if filter_condition.strip() == "not_equal":
                    kwargs = {
                        '{0}__iexact'.format(field_name): filter_value
                    }
                    return ~Q(**kwargs)

                if filter_condition.strip() == "starts_with":
                    kwargs = {
                        '{0}__istartswith'.format(field_name): filter_value
                    }
                    return Q(**kwargs)
                if filter_condition.strip() == "equal":
                    kwargs = {
                        '{0}__iexact'.format(field_name): filter_value
                    }
                    return Q(**kwargs)

                if filter_condition.strip() == "not_equal":
                    kwargs = {
                        '{0}__ne'.format(field_name): filter_value
                    }
                    return ~Q(**kwargs)
                if filter_condition.strip() == "less_than":
                    kwargs = {
                        '{0}__lt'.format(field_name): filter_value
                    }
                    return ~Q(**kwargs)
                if filter_condition.strip() == "greater_than":
                    kwargs = {
                        '{0}__gt'.format(field_name): filter_value
                    }
                    return ~Q(**kwargs)

            typeQuery = Q()
            filter_objects = Q()

            if component !='':
                filter_objects &= get_filter(
                    'component_type', 'equal', component)

            data = ActiveMotors.objects.filter(component_type = component, user_id=userId)
            serializer = ActiveMotorSerializer(data, many=True)
            return JsonResponse({'status': 'True', 'data': serializer.data},
                                    status=200)
        except:
            return JsonResponse({'message': 'Sorry! No Audit found.'}, status=500)
    @staticmethod
    def getUserProductionList(request):
        try:
            userId = request.query_params['id']
            data = ProductionSystemStatus.objects.filter(user_id=userId)
            serializer = ProductionSystemSerialzer(data, many=True)
            return JsonResponse({'status': 'True', 'data': serializer.data},
                                status=200)
        except:
            return JsonResponse({'message': 'Sorry! No Audit found.'}, status=500)

    @staticmethod
    def getUserRelifingList(request):
        try:
            userId = request.query_params['id']
            data = RelifingSystemStatus.objects.filter(user_id=userId)
            serializer = RelifingSystemSerialzer(data, many=True)
            return JsonResponse({'status': 'True', 'data': serializer.data},
                                status=200)
        except:
            return JsonResponse({'message': 'Sorry! No Audit found.'}, status=500)

    @staticmethod
    def getUserFlightList(request):
        try:
            userId = request.query_params['id']
            data = FlightSystemStatus.objects.filter(user_id=userId)
            serializer = FlightSystemSerialzer(data, many=True)
            print(serializer.data)
            return JsonResponse({'status': 'True', 'data': serializer.data},
                                status=200)
        except:
            return JsonResponse({'message': 'Sorry! No Audit found.'}, status=500)


    @staticmethod
    def AddUserCSV(request):
        userModel = User()
        # attachment = request['attachment']
        with open('users.csv', 'r') as csv_file:
            csvf = reader(csv_file)
            data = []
            for first_name, last_name, username, password, email, *__ in csvf:
                user = User(first_name=first_name,last_name=last_name,username=username,email=email )
                user.set_password(password)
                data.append(user)
            User.objects.bulk_create(data)
            return JsonResponse({'message': 'Sorry! No Audit found.'}, status=500)

        # try:
        #     uname = request['username']
        #     id = request['id']
        #     if id == '0':
        #         userModel.first_name = request['first_name']
        #         userModel.last_name = request['last_name']
        #         userModel.username = request['username']
        #         userModel.email = 'nescom@nescom.com'
        #         encryptedpassword = make_password(request['password'])
        #         userModel.password = encryptedpassword
        #         userModel.is_superuser = 'False'
        #         userModel.last_login = date.today()
        #         userModel.is_active = 'True'
        #         userModel.is_staff = 'True'
        #         userModel.save()