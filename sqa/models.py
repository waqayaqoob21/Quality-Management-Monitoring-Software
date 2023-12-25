from django.db import models

# Create your models here.

class Sqa(models.Model):
    id = models.AutoField(primary_key = True)
    sys_type = models.CharField(max_length=300)
    system_name = models.CharField(max_length=300)
    module_id = models.CharField(max_length=300)
    module_name = models.CharField(max_length=300)
    platform = models.CharField(max_length=300)
    software_version = models.CharField(max_length=300)
    sqa_certificate_no = models.CharField(max_length=300)
    selected_status = models.CharField(max_length=300)
    requestDate = models.DateTimeField(auto_now_add=False)
    audit_completion_date = models.DateTimeField(auto_now_add=False)
    DueDate = models.DateTimeField(auto_now_add=False)
    assertion_density = models.CharField(max_length=300)
    bugs = models.CharField(max_length=300)
    cyclomatic_complexity = models.CharField(max_length=300)
    code_coverage = models.CharField(max_length=300)
    functional_testing = models.CharField(max_length=300)
    unit_testing = models.CharField(max_length=300)
    static_testing = models.CharField(max_length=300)
    formal_testing = models.CharField(max_length=300)
    standard_compliance = models.CharField(max_length=300)
    remarks = models.TextField()