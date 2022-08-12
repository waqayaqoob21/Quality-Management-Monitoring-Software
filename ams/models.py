from django.db import models

# Create your models here.
class TaskSummary(models.Model):
    task = models.TextField()
    assigned_by = models.CharField(max_length=100)
    assigned_date_by = models.DateField()
    target_date_by = models.DateField()
    assigned_to = models.CharField(max_length=255)
    assigned_date_to = models.DateField()
    target_date_to = models.DateField()
    status = models.CharField(max_length=355)
    attachment = models.TextField()
    isActive = models.BooleanField()

