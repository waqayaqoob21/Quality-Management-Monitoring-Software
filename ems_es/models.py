from django.db import models

# Create your models here.
class EmsEs(models.Model):
    id = models.AutoField(primary_key=True)
    sys_type = models.CharField(max_length=300)
    sys_name = models.CharField(max_length=300)
    product = models.CharField(max_length=300)
    test = models.CharField(max_length=300,null=True)
    platform = models.CharField(max_length=300)
    test_requirements = models.CharField(max_length=300)
    test_conducted = models.CharField(max_length=300)
    test_compliance = models.CharField(max_length=300)
    selected_status = models.CharField(max_length=300)
    report_status = models.CharField(max_length=300)
    module = models.CharField(max_length=300)
    compliance_status = models.CharField(max_length=300)