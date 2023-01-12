from django.db import models

# Create your models here.


class doctrackingHistory(models.Model):
    id = models.AutoField(primary_key=True)
    doc_id = models.IntegerField(null=False)
    doc_name = models.CharField(max_length=500)
    doc_type = models.CharField(max_length=500)
    sender = models.CharField(max_length=100, null=True)
    receive_date = models.DateTimeField(auto_now_add=False, db_index=True, null=True)
    marked_to = models.CharField(max_length=100, null=True)
    marked_date = models.DateTimeField(auto_now_add=False, db_index=True, null=True)
    due_date = models.DateTimeField(auto_now_add=False, db_index=True, null=True)
    task_date = models.DateTimeField(auto_now_add=False, db_index=True, null=True)
    status = models.CharField(max_length=100, null=True)
    sent_to = models.CharField(max_length=100, null=True)
    sent_date = models.DateTimeField(auto_now_add=False, db_index=True, null=True)
    isActive = models.IntegerField(default=0, null=False)
    attachement = models.TextField()
    remarks = models.CharField(max_length=100, null=False)
    created_at = models.DateTimeField(auto_now_add=True, db_index=True)
    product_sr_no = models.CharField(max_length=100, null=True)
    tracking_id = models.CharField(max_length=500, null=True)