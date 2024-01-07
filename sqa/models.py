from django.db import models

# Create your models here.

class Sqa(models.Model):
    id = models.AutoField(primary_key = True)
    sys_type = models.CharField(max_length=300)
    system_name = models.CharField(max_length=300)
    module_name = models.CharField(max_length=300)
    module_id = models.CharField(max_length=300)
    software_version = models.CharField(max_length=300)
    sqa_certificate_no = models.CharField(max_length=300)
    request_date = models.DateTimeField(auto_now_add=False)
    urd = models.CharField(max_length=300, null=True)
    srs = models.CharField(max_length=300, null=True)
    sdd = models.CharField(max_length=300, null=True)
    rtm = models.CharField(max_length=300, null=True)
    stp = models.CharField(max_length=300, null=True)
    unit_test = models.CharField(max_length=300)
    static_analysis_report = models.CharField(max_length=300)
    assertion_density = models.CharField(max_length=300)
    eng_change_proposal = models.CharField(max_length=300, null=True)
    bugs_observation = models.CharField(max_length=300)
    cyclomatic_complexity = models.CharField(max_length=300)
    code_coverage = models.CharField(max_length=300)
    functional_testing = models.CharField(max_length=300)
    formal_testing = models.CharField(max_length=300)
    status = models.CharField(max_length=300)
    due_date = models.DateTimeField(auto_now_add=False)
    audit_completion_date = models.DateTimeField(auto_now_add=False)
    remarks = models.TextField()

