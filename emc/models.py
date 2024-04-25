from django.db import models

# Create your models here.
class Emc(models.Model):
    id = models.AutoField(primary_key=True)
    organization = models.CharField(max_length=300, null=True)
    sys_type = models.CharField(max_length=300)
    sys_name = models.CharField(max_length=300)
    product = models.CharField(max_length=300, null=True)
    # product_tests = models.TextField()
    test = models.CharField(max_length=300,null=True)
    platform = models.CharField(max_length=300,null=True)
    test_requirements = models.CharField(max_length=300,null=True)
    test_conducted = models.CharField(max_length=300,null=True)
    test_compliance = models.CharField(max_length=300,null=True)
    selected_status = models.CharField(max_length=300,null=True)
    report_status = models.CharField(max_length=300,null=True)
    module = models.CharField(max_length=300,null=True)
    compliance_status = models.CharField(max_length=300,null=True)
    due_date = models.DateTimeField(auto_now_add=False)
    audit_completion_date = models.DateTimeField(auto_now_add=False)
    created_at = models.DateTimeField(auto_now_add=True)