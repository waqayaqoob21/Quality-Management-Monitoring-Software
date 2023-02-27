from django.db import models


# Create your models here.
class ActiveMotors(models.Model):
    id = models.AutoField(primary_key=True)
    system_type = models.CharField(max_length=300, null=True)

    system_name = models.CharField(max_length=300)
    organization = models.CharField(max_length=300,null=True)
    testing_type = models.CharField(max_length=300, null=True)
    testing_date = models.DateTimeField(auto_now_add=False, null=True)
    component_type = models.CharField(max_length=300, null=True)
    motor_id = models.CharField(max_length=300)

    qualification_insulation_lining_propellant_rm = models.CharField(max_length=300)
    qualification_insulation_lining_propellant_rm_date = models.DateTimeField(auto_now_add=False, null=True)
    qualification_insulation_lining_propellant_rm_remarks = models.TextField()

    lining = models.CharField(max_length=300)
    lining_date = models.DateTimeField(auto_now_add=False, null=True)
    lining_remarks = models.TextField()

    propellant_mechanical_properties = models.CharField(max_length=300)
    propellant_mechanical_properties_date = models.DateTimeField(auto_now_add=False, null=True)
    propellant_mechanical_properties_remarks = models.TextField()

    sandblasting = models.CharField(max_length=300)
    sandblasting_date = models.DateTimeField(auto_now_add=False, null=True)
    sandblasting_remarks = models.TextField()

    insulation = models.CharField(max_length=300)
    insulation_date = models.DateTimeField(auto_now_add=False, null=True)
    insulation_remarks = models.TextField()

    ut_rt_insulated_case = models.CharField(max_length=300)
    ut_rt_insulated_case_date = models.DateTimeField(auto_now_add=False, null=True)
    ut_rt_insulated_case_remarks = models.TextField()

    acceptance_silver_material = models.CharField(max_length=300)
    acceptance_silver_material_date = models.DateTimeField(auto_now_add=False, null=True)
    acceptance_silver_material_remarks = models.TextField()

    silver_application = models.CharField(max_length=300)
    silver_application_date = models.DateTimeField(auto_now_add=False, null=True)
    silver_application_remarks = models.TextField()

    formulation_tailoring_liner_propellant = models.CharField(max_length=300)
    formulation_tailoring_liner_propellant_date = models.DateTimeField(auto_now_add=False, null=True)
    formulation_tailoring_liner_propellant_remarks = models.TextField()

    conditioning_raw_materials = models.CharField(max_length=300)
    conditioning_raw_materials_date = models.DateTimeField(auto_now_add=False, null=True)
    conditioning_raw_materials_remarks = models.TextField()

    conditioning_of_lining = models.CharField(max_length=300,null=True)
    conditioning_of_lining_date = models.DateTimeField(auto_now_add=False, null=True)
    conditioning_of_lining_remarks = models.TextField(null=True)

    casting = models.CharField(max_length=300)
    casting_date = models.DateTimeField(auto_now_add=False, null=True)
    casting_remarks = models.TextField()

    curing = models.CharField(max_length=300)
    curing_date = models.DateTimeField(auto_now_add=False, null=True)
    curing_remarks = models.TextField()

    ut_endoscopy_rt_grain = models.CharField(max_length=300)
    ut_endoscopy_rt_grain_date = models.DateTimeField(auto_now_add=False, null=True)
    ut_endoscopy_rt_grain_remarks = models.TextField()

    liner_mechanical_properties = models.CharField(max_length=300)
    liner_mechanical_properties_date = models.DateTimeField(auto_now_add=False, null=True)
    liner_mechanical_properties_remarks = models.TextField()

    mechanical_properties_liner = models.CharField(max_length=300, null=True)
    mechanical_properties_liner_date = models.DateTimeField(auto_now_add=False, null=True)
    mechanical_properties_liner_remarks = models.TextField(null=True)

    mechanical_properties_propellant = models.CharField(max_length=300, null=True)
    mechanical_properties_propellant_date = models.DateTimeField(auto_now_add=False, null=True)
    mechanical_properties_propellant_remarks = models.TextField(null=True)

    interface_bond_strength = models.CharField(max_length=300)
    interface_bond_strength_date = models.DateTimeField(auto_now_add=False, null=True)
    interface_bond_strength_remarks = models.TextField()

    propellant_burn_rate = models.CharField(max_length=300)
    propellant_burn_rate_date = models.DateTimeField(auto_now_add=False, null=True)
    propellant_burn_rate_remarks = models.TextField()

    mass_liner_insulation_propellant_srm = models.CharField(max_length=300)
    mass_liner_insulation_propellant_srm_date = models.DateTimeField(auto_now_add=False, null=True)
    mass_liner_insulation_propellant_srm_remarks = models.TextField()

    mass_liner = models.CharField(max_length=300, null=True)
    mass_liner_date = models.DateTimeField(auto_now_add=False, null=True)
    mass_liner_remarks = models.TextField(null=True)

    mass_insulation = models.CharField(max_length=300, null=True)
    mass_insulation_date = models.DateTimeField(auto_now_add=False, null=True)
    mass_insulation_remarks = models.TextField(null=True)

    mass_propellant = models.CharField(max_length=300, null=True)
    mass_propellant_date = models.DateTimeField(auto_now_add=False, null=True)
    mass_propellant_remarks = models.TextField(null=True)

    overall_qualification_status = models.CharField(max_length=300, null=True)
    overall_qualification_date = models.DateTimeField(auto_now_add=False, null=True)
    overall_qualification_remarks = models.TextField(null=True)

    bhd_status = models.CharField(max_length=300, null=True)
    bhd_date = models.DateTimeField(auto_now_add=False, null=True)
    bhd_remarks = models.TextField(null=True)

    qm_certification_status = models.CharField(max_length=300, null=True)
    qm_certification_date = models.DateTimeField(auto_now_add=False, null=True)
    qm_certification_remarks = models.TextField(null=True)

    acceptance_casting = models.CharField(max_length=300)
    acceptance_casting_date = models.DateTimeField(auto_now_add=False, null=True)
    acceptance_casting_remarks = models.TextField()

    trimming_Propellant_grain = models.CharField(max_length=300)
    trimming_Propellant_grain_date = models.DateTimeField(auto_now_add=False, null=True)
    trimming_Propellant_grain_remarks = models.TextField()
    # ====================================== Battries ==============================================
    battery_type = models.CharField(max_length=300,null=True)
    tb_type = models.CharField(max_length=300,null=True)
    battery_id = models.CharField(max_length=300,null=True)
    lot_id  = models.CharField(max_length=300,null=True)

    matiral_qualified_date = models.DateTimeField(auto_now_add=False,null=True)
    matiral_qualified_status = models.CharField(max_length=300,null=True)
    matiral_qualified_remarks = models.TextField(null=True)

    compo_manufacturing_date = models.DateTimeField(auto_now_add=False,null=True)
    compo_manufacturing_status = models.CharField(max_length=300,null=True)
    compo_manufacturing_remarks = models.TextField(null=True)

    powerpack_assembly_date = models.DateTimeField(auto_now_add=False,null=True)
    powerpack_assembly_status = models.CharField(max_length=300,null=True)
    powerpack_assembly_remarks = models.TextField(null=True)

    powerpack_testing_date = models.DateTimeField(auto_now_add=False,null=True)
    powerpack_testing_status = models.CharField(max_length=300,null=True)
    powerpack_testing_remarks = models.TextField(null=True)

    raw_material_inspection_date = models.DateTimeField(auto_now_add=False,null=True)
    raw_material_inspection_status = models.CharField(max_length=300,null=True)
    raw_material_inspection_remarks = models.TextField(null=True)

    pressing_electrode_date = models.DateTimeField(auto_now_add=False,null=True)
    pressing_electrode_status = models.CharField(max_length=300,null=True)
    pressing_electrode_remarks = models.TextField(null=True)

    formation_process_date = models.DateTimeField(auto_now_add=False,null=True)
    formation_process_status = models.CharField(max_length=300,null=True)
    formation_process_remarks = models.TextField(null=True)

    assembly_process_date = models.DateTimeField(auto_now_add=False,null=True)
    assembly_process_status = models.CharField(max_length=300,null=True)
    assembly_process_remarks = models.TextField(null=True)

    battery_testing_date = models.DateTimeField(auto_now_add=False,null=True)
    battery_testing_status = models.CharField(max_length=300,null=True)
    battery_testing_remarks = models.TextField(null=True)

    final_qualification_date = models.DateTimeField(auto_now_add=False,null=True)
    final_qualification_status = models.CharField(max_length=300,null=True)
    final_qualification_remarks = models.TextField(null=True)
# ====================================== Battries ==============================================

    # ====================================== Pyro Devices ==============================================

    pd_type = models.CharField(max_length=300, null=True)
    pd_id = models.CharField(max_length=300, null=True)

    qualification_raw_material_date = models.DateTimeField(auto_now_add=False, null=True)
    qualification_raw_material_status = models.CharField(max_length=300, null=True)
    qualification_raw_material_remarks = models.TextField(null=True)

    filling_date = models.DateTimeField(auto_now_add=False, null=True)
    filling_status = models.CharField(max_length=300, null=True)
    filling_remarks = models.TextField(null=True)

    assembling_integration_date = models.DateTimeField(auto_now_add=False, null=True)
    assembling_integration_status = models.CharField(max_length=300, null=True)
    assembling_integration_remarks = models.TextField(null=True)

    qualification_testing_date = models.DateTimeField(auto_now_add=False, null=True)
    qualification_testing_status = models.CharField(max_length=300, null=True)
    qualification_testing_remarks = models.TextField(null=True)

    performance_testing_date = models.DateTimeField(auto_now_add=False, null=True)
    performance_testing_status = models.CharField(max_length=300, null=True)
    performance_testing_remarks = models.TextField(null=True)

    # ====================================== Pyro Devices ==============================================

    ncr_status = models.CharField(max_length=300)
    ncr_status_remarks = models.TextField()
    overall_status = models.CharField(max_length=300)
    overall_remarks = models.TextField()
    attachments = models.TextField()
    Created_at = models.DateTimeField(auto_now_add=True)

class ActiveMotorsHistory(models.Model):
    id = models.AutoField(primary_key=True)
    active_motor_id = models.IntegerField(null=True)
    system_type = models.CharField(max_length=300, null=True)

    system_name = models.CharField(max_length=300)
    organization = models.CharField(max_length=300,null=True)
    testing_type = models.CharField(max_length=300, null=True)
    testing_date = models.DateTimeField(auto_now_add=False, null=True)
    component_type = models.CharField(max_length=300, null=True)
    motor_id = models.CharField(max_length=300)

    qualification_insulation_lining_propellant_rm = models.CharField(max_length=300)
    qualification_insulation_lining_propellant_rm_date = models.DateTimeField(auto_now_add=False, null=True)
    qualification_insulation_lining_propellant_rm_remarks = models.TextField()

    lining = models.CharField(max_length=300)
    lining_date = models.DateTimeField(auto_now_add=False, null=True)
    lining_remarks = models.TextField()

    propellant_mechanical_properties = models.CharField(max_length=300)
    propellant_mechanical_properties_date = models.DateTimeField(auto_now_add=False, null=True)
    propellant_mechanical_properties_remarks = models.TextField()

    sandblasting = models.CharField(max_length=300)
    sandblasting_date = models.DateTimeField(auto_now_add=False, null=True)
    sandblasting_remarks = models.TextField()

    insulation = models.CharField(max_length=300)
    insulation_date = models.DateTimeField(auto_now_add=False, null=True)
    insulation_remarks = models.TextField()

    ut_rt_insulated_case = models.CharField(max_length=300)
    ut_rt_insulated_case_date = models.DateTimeField(auto_now_add=False, null=True)
    ut_rt_insulated_case_remarks = models.TextField()

    acceptance_silver_material = models.CharField(max_length=300)
    acceptance_silver_material_date = models.DateTimeField(auto_now_add=False, null=True)
    acceptance_silver_material_remarks = models.TextField()

    silver_application = models.CharField(max_length=300)
    silver_application_date = models.DateTimeField(auto_now_add=False, null=True)
    silver_application_remarks = models.TextField()

    formulation_tailoring_liner_propellant = models.CharField(max_length=300)
    formulation_tailoring_liner_propellant_date = models.DateTimeField(auto_now_add=False, null=True)
    formulation_tailoring_liner_propellant_remarks = models.TextField()

    conditioning_raw_materials = models.CharField(max_length=300)
    conditioning_raw_materials_date = models.DateTimeField(auto_now_add=False, null=True)
    conditioning_raw_materials_remarks = models.TextField()

    conditioning_of_lining = models.CharField(max_length=300,null=True)
    conditioning_of_lining_date = models.DateTimeField(auto_now_add=False, null=True)
    conditioning_of_lining_remarks = models.TextField(null=True)

    casting = models.CharField(max_length=300)
    casting_date = models.DateTimeField(auto_now_add=False, null=True)
    casting_remarks = models.TextField()

    curing = models.CharField(max_length=300)
    curing_date = models.DateTimeField(auto_now_add=False, null=True)
    curing_remarks = models.TextField()

    ut_endoscopy_rt_grain = models.CharField(max_length=300)
    ut_endoscopy_rt_grain_date = models.DateTimeField(auto_now_add=False, null=True)
    ut_endoscopy_rt_grain_remarks = models.TextField()

    liner_mechanical_properties = models.CharField(max_length=300)
    liner_mechanical_properties_date = models.DateTimeField(auto_now_add=False, null=True)
    liner_mechanical_properties_remarks = models.TextField()

    mechanical_properties_liner = models.CharField(max_length=300, null=True)
    mechanical_properties_liner_date = models.DateTimeField(auto_now_add=False, null=True)
    mechanical_properties_liner_remarks = models.TextField(null=True)

    mechanical_properties_propellant = models.CharField(max_length=300, null=True)
    mechanical_properties_propellant_date = models.DateTimeField(auto_now_add=False, null=True)
    mechanical_properties_propellant_remarks = models.TextField(null=True)

    interface_bond_strength = models.CharField(max_length=300)
    interface_bond_strength_date = models.DateTimeField(auto_now_add=False, null=True)
    interface_bond_strength_remarks = models.TextField()

    propellant_burn_rate = models.CharField(max_length=300)
    propellant_burn_rate_date = models.DateTimeField(auto_now_add=False, null=True)
    propellant_burn_rate_remarks = models.TextField()

    mass_liner_insulation_propellant_srm = models.CharField(max_length=300)
    mass_liner_insulation_propellant_srm_date = models.DateTimeField(auto_now_add=False, null=True)
    mass_liner_insulation_propellant_srm_remarks = models.TextField()

    mass_liner = models.CharField(max_length=300, null=True)
    mass_liner_date = models.DateTimeField(auto_now_add=False, null=True)
    mass_liner_remarks = models.TextField(null=True)

    mass_insulation = models.CharField(max_length=300, null=True)
    mass_insulation_date = models.DateTimeField(auto_now_add=False, null=True)
    mass_insulation_remarks = models.TextField(null=True)

    mass_propellant = models.CharField(max_length=300, null=True)
    mass_propellant_date = models.DateTimeField(auto_now_add=False, null=True)
    mass_propellant_remarks = models.TextField(null=True)

    bhd_status = models.CharField(max_length=300, null=True)
    bhd_date = models.DateTimeField(auto_now_add=False, null=True)
    bhd_remarks = models.TextField(null=True)

    overall_qualification_status = models.CharField(max_length=300, null=True)
    overall_qualification_date = models.DateTimeField(auto_now_add=False, null=True)
    overall_qualification_remarks = models.TextField(null=True)

    qm_certification_status = models.CharField(max_length=300, null=True)
    qm_certification_date = models.DateTimeField(auto_now_add=False, null=True)
    qm_certification_remarks = models.TextField(null=True)

    acceptance_casting = models.CharField(max_length=300)
    acceptance_casting_date = models.DateTimeField(auto_now_add=False, null=True)
    acceptance_casting_remarks = models.TextField()

    trimming_Propellant_grain = models.CharField(max_length=300)
    trimming_Propellant_grain_date = models.DateTimeField(auto_now_add=False, null=True)
    trimming_Propellant_grain_remarks = models.TextField()
    # ====================================== Battries ==============================================
    battery_type = models.CharField(max_length=300,null=True)
    tb_type = models.CharField(max_length=300,null=True)
    battery_id = models.CharField(max_length=300,null=True)
    lot_id  = models.CharField(max_length=300,null=True)

    matiral_qualified_date = models.DateTimeField(auto_now_add=False,null=True)
    matiral_qualified_status = models.CharField(max_length=300,null=True)
    matiral_qualified_remarks = models.TextField(null=True)

    compo_manufacturing_date = models.DateTimeField(auto_now_add=False,null=True)
    compo_manufacturing_status = models.CharField(max_length=300,null=True)
    compo_manufacturing_remarks = models.TextField(null=True)

    powerpack_assembly_date = models.DateTimeField(auto_now_add=False,null=True)
    powerpack_assembly_status = models.CharField(max_length=300,null=True)
    powerpack_assembly_remarks = models.TextField(null=True)

    powerpack_testing_date = models.DateTimeField(auto_now_add=False,null=True)
    powerpack_testing_status = models.CharField(max_length=300,null=True)
    powerpack_testing_remarks = models.TextField(null=True)

    raw_material_inspection_date = models.DateTimeField(auto_now_add=False,null=True)
    raw_material_inspection_status = models.CharField(max_length=300,null=True)
    raw_material_inspection_remarks = models.TextField(null=True)

    pressing_electrode_date = models.DateTimeField(auto_now_add=False,null=True)
    pressing_electrode_status = models.CharField(max_length=300,null=True)
    pressing_electrode_remarks = models.TextField(null=True)

    formation_process_date = models.DateTimeField(auto_now_add=False,null=True)
    formation_process_status = models.CharField(max_length=300,null=True)
    formation_process_remarks = models.TextField(null=True)

    assembly_process_date = models.DateTimeField(auto_now_add=False,null=True)
    assembly_process_status = models.CharField(max_length=300,null=True)
    assembly_process_remarks = models.TextField(null=True)

    battery_testing_date = models.DateTimeField(auto_now_add=False,null=True)
    battery_testing_status = models.CharField(max_length=300,null=True)
    battery_testing_remarks = models.TextField(null=True)

    final_qualification_date = models.DateTimeField(auto_now_add=False,null=True)
    final_qualification_status = models.CharField(max_length=300,null=True)
    final_qualification_remarks = models.TextField(null=True)
# ====================================== Battries ==============================================

# ====================================== Pyro Devices ==============================================

    pd_type = models.CharField(max_length=300,null=True)
    pd_id = models.CharField(max_length=300,null=True)

    qualification_raw_material_date = models.DateTimeField(auto_now_add=False,null=True)
    qualification_raw_material_status = models.CharField(max_length=300,null=True)
    qualification_raw_material_remarks = models.TextField(null=True)

    filling_date = models.DateTimeField(auto_now_add=False,null=True)
    filling_status = models.CharField(max_length=300,null=True)
    filling_remarks = models.TextField(null=True)

    assembling_integration_date = models.DateTimeField(auto_now_add=False,null=True)
    assembling_integration_status = models.CharField(max_length=300,null=True)
    assembling_integration_remarks = models.TextField(null=True)

    qualification_testing_date = models.DateTimeField(auto_now_add=False,null=True)
    qualification_testing_status = models.CharField(max_length=300,null=True)
    qualification_testing_remarks = models.TextField(null=True)

    performance_testing_date = models.DateTimeField(auto_now_add=False,null=True)
    performance_testing_status = models.CharField(max_length=300,null=True)
    performance_testing_remarks = models.TextField(null=True)

# ====================================== Pyro Devices ==============================================

    ncr_status = models.CharField(max_length=300)
    ncr_status_remarks = models.TextField()
    overall_status = models.CharField(max_length=300)
    overall_remarks = models.TextField()
    attachments = models.TextField()
    Created_at = models.DateTimeField(auto_now_add=True)

class organization_lot_ids(models.Model):
    id = models.AutoField(primary_key=True)
    organization = models.CharField(max_length=100, null=False)
    lot_id_number = models.CharField(max_length=100, null=False)
    Created_at = models.DateTimeField(auto_now_add=True)