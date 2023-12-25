from django.db import models

# Create your models here.
class Reliability(models.Model):
    id = models.AutoField(primary_key = True)
    sys_type = models.CharField(max_length=300)
    system_name = models.CharField(max_length=300)
    items = models.CharField(max_length=300)
    age_estimation = models.CharField(max_length=300)
    age_method = models.CharField(max_length=300)
    age_remarks = models.CharField(max_length=300)
    reliability_estimation = models.CharField(max_length=300)
    reliability_estimation_method = models.CharField(max_length=300)
    reliability_estimation_remarks = models.CharField(max_length=300)