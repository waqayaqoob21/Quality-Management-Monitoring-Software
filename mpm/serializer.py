from rest_framework import serializers
from rest_framework.serializers import Serializer
from rest_framework.fields import CharField, IntegerField, DateTimeField

from mpm.models import organization_lot_ids


class ActiveMotorSerializer(Serializer):
    id = IntegerField()
    types = CharField()
    system = CharField()
    motor_id = CharField()
    year = DateTimeField()
    component_type = CharField()
    qualification_insulation_lining_propellant_rm = CharField()
    qualification_insulation_lining_propellant_rm_date = DateTimeField()
    qualification_insulation_lining_propellant_rm_remarks = CharField()

    acceptance_casting = CharField()
    acceptance_casting_date = DateTimeField()
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

    casting = CharField()
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

    interface_bond_strength = CharField()
    interface_bond_strength_date = DateTimeField()
    interface_bond_strength_remarks = CharField()

    propellant_burn_rate = CharField()
    propellant_burn_rate_date = DateTimeField()
    propellant_burn_rate_remarks = CharField()

    trimming_Propellant_grain = CharField()
    trimming_Propellant_grain_date = DateTimeField()
    trimming_Propellant_grain_remarks = CharField()

    mass_liner_insulation_propellant_srm = CharField()
    mass_liner_insulation_propellant_srm_date = DateTimeField()
    mass_liner_insulation_propellant_srm_remarks = CharField()

    ut_endoscopy_rt_grain = CharField()
    ut_endoscopy_rt_grain_date = DateTimeField()
    ut_endoscopy_rt_grain_remarks = CharField()

    ncr_status = CharField()
    ncr_status_remarks = CharField()
    overall_status = CharField()
    overall_remarks = CharField()
    attachments = CharField()
    Created_at = DateTimeField()


class LotIdsSerialzer(serializers.ModelSerializer):
    class Meta:
        model = organization_lot_ids
        fields = ['id', 'organization', 'lot_id_number','Created_at']
