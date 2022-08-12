from django.db import models

# Create your models here.
class TaskSummary(models.Model):
    task_name = models.TextField()
    assigned_by = models.CharField(max_length=100)
    assigned_date = models.DateField()
    assigned_to = models.CharField(max_length=255)
    target_date = models.DateField()
    status = models.CharField(max_length=355)
    remarks = models.TextField()

