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

class OcrDataModel(models.Model):
    organization = models.CharField(max_length=300)
    senior_directorate = models.CharField(max_length=500)
    site = models.CharField(max_length=500)
    test_report_no = models.CharField(max_length=100)
    job_card_no = models.CharField(max_length=100)
    test_report_date = models.DateField()
    product_name = models.CharField(max_length=300)
    id_no = models.CharField(max_length=100)
    lot_no_lot_size = models.CharField(max_length=200)
    test_name = models.CharField(max_length=300)
    test_type = models.CharField(max_length=300)
    qualification_standard = models.CharField(max_length=300)
    test_specifications = models.CharField(max_length=300)
    results = models.CharField(max_length=300)
    remarks = models.TextField()
    Created_at = models.DateTimeField(auto_now_add=True)