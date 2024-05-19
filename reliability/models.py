from django.db import models

# Create your models here.
class Reliability(models.Model):
    id = models.AutoField(primary_key = True)
    system_type = models.CharField(max_length=300, null=True)
    module_name = models.CharField(max_length=300, null=True)
    estimation_type = models.CharField(max_length=300, null=True)
    estimation_method = models.CharField(max_length=300, null=True)
    request_date = models.DateTimeField(auto_now_add=False, null=True)
    due_date = models.DateTimeField(auto_now_add=False, null=True)
    completion_date = models.DateTimeField(auto_now_add=False, null=True)
    current_status = models.CharField(max_length=300, null=True)
    remarks = models.CharField(max_length=300, null=True)
    created_at = models.DateTimeField(auto_now_add=True, null=True)

class ReliabilityHistory(models.Model):
    id = models.AutoField(primary_key = True)
    reliability_id = models.IntegerField(null=True)
    system_type = models.CharField(max_length=300, null=True)
    module_name = models.CharField(max_length=300, null=True)
    estimation_type = models.CharField(max_length=300, null=True)
    estimation_method = models.CharField(max_length=300, null=True)
    request_date = models.DateTimeField(auto_now_add=False, null=True)
    due_date = models.DateTimeField(auto_now_add=False, null=True)
    completion_date = models.DateTimeField(auto_now_add=False, null=True)
    current_status = models.CharField(max_length=300, null=True)
    remarks = models.CharField(max_length=300, null=True)
    created_at = models.DateTimeField(auto_now_add=True, null=True)