from django.db import models


# Create your models here.
class ActiveMotors(models.Model):
    id = models.AutoField(primary_key=True)
    types = models.CharField(max_length=300, null=True)
    system = models.CharField(max_length=300)
    motor_id = models.CharField(max_length=300)
    year = models.DateTimeField(auto_now_add=False, db_index=True, null=True)
    component_type = models.CharField(max_length=300, null=True)

    qualification_insulation_lining_propellant_rm = models.CharField(max_length=300)
    qualification_insulation_lining_propellant_rm_date = models.DateTimeField(auto_now_add=False, db_index=True, null=True)
    qualification_insulation_lining_propellant_rm_remarks = models.TextField()

    acceptance_casting = models.CharField(max_length=300)
    acceptance_casting_date = models.DateTimeField(auto_now_add=False, db_index=True, null=True)
    acceptance_casting_remarks = models.TextField()

    sandblasting = models.CharField(max_length=300)
    sandblasting_date = models.DateTimeField(auto_now_add=False, db_index=True, null=True)
    sandblasting_remarks = models.TextField()

    insulation = models.CharField(max_length=300)
    insulation_date = models.DateTimeField(auto_now_add=False, db_index=True, null=True)
    insulation_remarks = models.TextField()

    ut_rt_insulated_case = models.CharField(max_length=300)
    ut_rt_insulated_case_date =  models.DateTimeField(auto_now_add=False, db_index=True, null=True)
    ut_rt_insulated_case_remarks = models.TextField()

    acceptance_silver_material = models.CharField(max_length=300)
    acceptance_silver_material_date = models.DateTimeField(auto_now_add=False, db_index=True, null=True)
    acceptance_silver_material_remarks = models.TextField()

    silver_application = models.CharField(max_length=300)
    silver_application_date = models.DateTimeField(auto_now_add=False, db_index=True, null=True)
    silver_application_remarks = models.TextField()

    formulation_tailoring_liner_propellant = models.CharField(max_length=300)
    formulation_tailoring_liner_propellant_date = models.DateTimeField(auto_now_add=False, db_index=True, null=True)
    formulation_tailoring_liner_propellant_remarks = models.TextField()

    conditioning_raw_materials = models.CharField(max_length=300)
    conditioning_raw_materials_date = models.DateTimeField(auto_now_add=False, db_index=True, null=True)
    conditioning_raw_materials_remarks = models.TextField()

    lining = models.CharField(max_length=300)
    lining_date = models.DateTimeField(auto_now_add=False, db_index=True, null=True)
    lining_remarks = models.TextField()

    casting = models.CharField(max_length=300)
    casting_date = models.DateTimeField(auto_now_add=False, db_index=True, null=True)
    casting_remarks = models.TextField()

    curing = models.CharField(max_length=300)
    curing_date = models.DateTimeField(auto_now_add=False, db_index=True, null=True)
    curing_remarks = models.TextField()

    liner_mechanical_properties = models.CharField(max_length=300)
    liner_mechanical_properties_date = models.DateTimeField(auto_now_add=False, db_index=True, null=True)
    liner_mechanical_properties_remarks = models.TextField()

    propellant_mechanical_properties = models.CharField(max_length=300)
    propellant_mechanical_properties_date = models.DateTimeField(auto_now_add=False, db_index=True, null=True)
    propellant_mechanical_properties_remarks = models.TextField()

    interface_bond_strength = models.CharField(max_length=300)
    interface_bond_strength_date =  models.DateTimeField(auto_now_add=False, db_index=True, null=True)
    interface_bond_strength_remarks = models.TextField()

    propellant_burn_rate = models.CharField(max_length=300)
    propellant_burn_rate_date = models.DateTimeField(auto_now_add=False, db_index=True, null=True)
    propellant_burn_rate_remarks = models.TextField()

    trimming_Propellant_grain = models.CharField(max_length=300)
    trimming_Propellant_grain_date = models.DateTimeField(auto_now_add=False, db_index=True, null=True)
    trimming_Propellant_grain_remarks = models.TextField()

    mass_liner_insulation_propellant_srm = models.CharField(max_length=300)
    mass_liner_insulation_propellant_srm_date = models.DateTimeField(auto_now_add=False, db_index=True, null=True)
    mass_liner_insulation_propellant_srm_remarks = models.TextField()

    ut_endoscopy_rt_grain = models.CharField(max_length=300)
    ut_endoscopy_rt_grain_date = models.DateTimeField(auto_now_add=False, db_index=True, null=True)
    ut_endoscopy_rt_grain_remarks = models.TextField()

    ncr_status = models.CharField(max_length=300)
    ncr_status_remarks = models.TextField()
    overall_status = models.CharField(max_length=300)
    overall_remarks = models.TextField()
    attachments = models.TextField()
    Created_at = models.DateTimeField(auto_now_add=True)
