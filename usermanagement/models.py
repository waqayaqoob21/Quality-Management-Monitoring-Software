from django.db import models
from django.contrib.auth.models import AbstractUser
from .manager import UserManager



# Create User models here.
class User(AbstractUser):
    user_type = models.CharField(max_length=300)

    USERNAME_FIELD = 'username'
    REQUIRED_FIELDS = []
    objects = UserManager()

class UserRoles(models.Model):
    id = models.AutoField(primary_key=True)
    prod_roles = models.CharField(max_length=555)
    relif_roles = models.CharField(max_length=555)
    flight_roles = models.CharField(max_length=555)
    motor_roles = models.CharField(max_length=555)
    battery_roles = models.CharField(max_length=555, null=True)
    pyro_roles = models.CharField(max_length=555, null=True)
    bhd_roles = models.CharField(max_length=555, null=True)
    task_roles = models.CharField(max_length=555, null=True)
    qms_roles = models.CharField(max_length=555, null=True)
    cesp_roles = models.CharField(max_length=555, null=True)
    user_id = models.CharField(max_length=555)

class doctracking(models.Model):
    id = models.AutoField(primary_key=True)
    doc_name = models.CharField(max_length=500)
    doc_type = models.CharField(max_length=500)
    sender = models.CharField(max_length=500, null=True)
    receive_date = models.DateTimeField(auto_now_add=False, db_index=True, null=True)
    marked_to = models.CharField(max_length=500, null=True)
    marked_date = models.DateTimeField(auto_now_add=False, db_index=True, null=True)
    due_date = models.DateTimeField(auto_now_add=False, db_index=True, null=True)
    task_date = models.DateTimeField(auto_now_add=False, db_index=True, null=True)
    status = models.CharField(max_length=500, null=True)
    sent_to = models.CharField(max_length=500, null=True)
    sent_date = models.DateTimeField(auto_now_add=False, db_index=True, null=True)
    isActive = models.IntegerField(default=0, null=False)
    attachement = models.TextField()
    remarks = models.CharField(max_length=500, null=False)
    created_at = models.DateTimeField(auto_now_add=True, db_index=True)
    product_sr_no = models.CharField(max_length=500, null=True)
    tracking_id = models.CharField(max_length=500, null=True)
    certificate_number = models.CharField(max_length=500, null=True)
    last_meeting_date = models.DateTimeField(auto_now_add=False, db_index=True, null=True)
