from rest_framework.serializers import Serializer
from rest_framework.fields import CharField, IntegerField, DateTimeField
from rest_framework import serializers
from mpm.models import organization_lot_ids

class ActiveMotorSerializer(Serializer):
    id = IntegerField()
    system_type = CharField()
    system_name = CharField()
    organization = CharField()
    creation_date = CharField()
    testing_date = CharField()
    testing_type = CharField()
    motor_id = CharField()
    component_type = CharField()
    qualification_insulation_lining_propellant_rm= CharField()
    qualification_insulation_lining_propellant_rm_date= DateTimeField()
    qualification_insulation_lining_propellant_rm_remarks = CharField()

    acceptance_casting = CharField()
    acceptance_casting_date =DateTimeField()
    acceptance_casting_remarks = CharField()

    sandblasting = CharField()
    sandblasting_date = DateTimeField()
    sandblasting_remarks = CharField()

    insulation = CharField()
    insulation_date = DateTimeField()
    insulation_remarks = CharField()

    ut_rt_insulated_case = CharField()
    ut_rt_insulated_case_date = DateTimeField()
    ut_rt_insulated_case_remarks = CharField()

    acceptance_silver_material = CharField()
    acceptance_silver_material_date = DateTimeField()
    acceptance_silver_material_remarks = CharField()

    silver_application = CharField()
    silver_application_date = DateTimeField()
    silver_application_remarks = CharField()

    formulation_tailoring_liner_propellant = CharField()
    formulation_tailoring_liner_propellant_date = DateTimeField()
    formulation_tailoring_liner_propellant_remarks = CharField()

    conditioning_raw_materials = CharField()
    conditioning_raw_materials_date = DateTimeField()
    conditioning_raw_materials_remarks = CharField()

    lining = CharField()
    lining_date = DateTimeField()
    lining_remarks = CharField()

    casting =  CharField()
    casting_date = DateTimeField()
    casting_remarks = CharField()

    curing = CharField()
    curing_date = DateTimeField()
    curing_remarks = CharField()

    liner_mechanical_properties = CharField()
    liner_mechanical_properties_date = DateTimeField()
    liner_mechanical_properties_remarks = CharField()

    propellant_mechanical_properties = CharField()
    propellant_mechanical_properties_date = DateTimeField()
    propellant_mechanical_properties_remarks = CharField()

    interface_bond_strength =  CharField()
    interface_bond_strength_date = DateTimeField()
    interface_bond_strength_remarks = CharField()

    propellant_burn_rate = CharField()
    propellant_burn_rate_date = DateTimeField()
    propellant_burn_rate_remarks = CharField()

    trimming_Propellant_grain =  CharField()
    trimming_Propellant_grain_date = DateTimeField()
    trimming_Propellant_grain_remarks = CharField()

    mass_liner_insulation_propellant_srm = CharField()
    mass_liner_insulation_propellant_srm_date = DateTimeField()
    mass_liner_insulation_propellant_srm_remarks = CharField()

    conditioning_of_lining = CharField()
    conditioning_of_lining_date = CharField()
    conditioning_of_lining_remarks = CharField()

    mechanical_properties_liner = CharField()
    mechanical_properties_liner_date = CharField()
    mechanical_properties_liner_remarks = CharField()

    mechanical_properties_propellant = CharField()
    mechanical_properties_propellant_date = CharField()
    mechanical_properties_propellant_remarks = CharField()

    mass_liner = CharField()
    mass_liner_date = CharField()
    mass_liner_remarks = CharField()

    mass_insulation = CharField()
    mass_insulation_date = CharField()
    mass_insulation_remarks = CharField()

    mass_propellant = CharField()
    mass_propellant_date = CharField()
    mass_propellant_remarks = CharField()

    overall_qualification_status = CharField()
    overall_qualification_date = CharField()
    overall_qualification_remarks = CharField()

    bhd_status = CharField()
    bhd_date = CharField()
    bhd_remarks = CharField()

    qm_certification_status = CharField()
    qm_certification_date = CharField()
    qm_certification_remarks = CharField()

    ut_endoscopy_rt_grain =  CharField()
    ut_endoscopy_rt_grain_date = DateTimeField()
    ut_endoscopy_rt_grain_remarks = CharField()


    battery_type = DateTimeField()
    tb_type = CharField()
    battery_id = CharField()
    lot_id = CharField()
    matiral_qualified_date = DateTimeField()
    matiral_qualified_status = CharField()
    matiral_qualified_remarks = CharField()

    compo_manufacturing_date = DateTimeField()
    compo_manufacturing_status = CharField()
    compo_manufacturing_remarks = CharField()

    powerpack_assembly_date = DateTimeField()
    powerpack_assembly_status = CharField()
    powerpack_assembly_remarks = CharField()

    powerpack_testing_date = DateTimeField()
    powerpack_testing_status = CharField()
    powerpack_testing_remarks = CharField()

    raw_material_inspection_date = DateTimeField()
    raw_material_inspection_status = CharField()
    raw_material_inspection_remarks = CharField()

    pressing_electrode_date = DateTimeField()
    pressing_electrode_status = CharField()
    pressing_electrode_remarks = CharField()

    formation_process_date = DateTimeField()
    formation_process_status = CharField()
    formation_process_remarks = CharField()

    assembly_process_date = DateTimeField()
    assembly_process_status = CharField()
    assembly_process_remarks = CharField()

    battery_testing_date = DateTimeField()
    battery_testing_status = CharField()
    battery_testing_remarks = CharField()

    final_qualification_date = DateTimeField()
    final_qualification_status = CharField()
    final_qualification_remarks = CharField()


    pd_type = CharField()
    pd_id = CharField()

    qualification_raw_material_date = DateTimeField()
    qualification_raw_material_status = CharField()
    qualification_raw_material_remarks = CharField()

    filling_date = DateTimeField()
    filling_status = CharField()
    filling_remarks = CharField()

    assembling_integration_date = DateTimeField()
    assembling_integration_status = CharField()
    assembling_integration_remarks = CharField()

    qualification_testing_date = DateTimeField()
    qualification_testing_status = CharField()
    qualification_testing_remarks = CharField()

    performance_testing_date = DateTimeField()
    performance_testing_status = CharField()
    performance_testing_remarks = CharField()

    ncr_status  = CharField()
    ncr_status_remarks = CharField()
    overall_status = CharField()
    overall_remarks = CharField()
    attachments =  CharField()
    Created_at = DateTimeField()



class LotIdsSerialzer(serializers.ModelSerializer):
    class Meta:
        model = organization_lot_ids
        fields = ['id', 'organization', 'lot_id_number','Created_at']