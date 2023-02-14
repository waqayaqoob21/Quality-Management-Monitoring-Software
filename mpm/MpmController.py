from importlib import import_module
from django.http import JsonResponse
from ams.serializer import *
from ams.models import *
from fpdf import FPDF
import xlwt
from datetime import datetime, timedelta, date
from django.http import FileResponse
from django.http import HttpResponse
from django.core.files.storage import FileSystemStorage
from mpm.models import ActiveMotors, organization_lot_ids,ActiveMotorsHistory
from mpm.serializer import ActiveMotorSerializer, LotIdsSerialzer
import itertools
from django.db.models import F, Q


class MpmController:
    @staticmethod
    def AddActiveMotor(request):
        try:
            id = request['id']
            motorModel = ActiveMotors()
            if id == 0 or id == '0':
                motorModel.system_type = request['system_type']
                motorModel.system_name = request['system_name']
                motorModel.organization = request['organization']
                motorModel.testing_type = request['testing_type']
                motorModel.testing_date = request['testing_date']
                motorModel.motor_id = request['motor_id']
                motorModel.component_type = request['component_type']
                motorModel.qualification_insulation_lining_propellant_rm = request[
                    'qualification_insulation_lining_propellant_rm']
                motorModel.qualification_insulation_lining_propellant_rm_date = request[
                    'qualification_insulation_lining_propellant_rm_date']
                motorModel.qualification_insulation_lining_propellant_rm_remarks = request[
                    'qualification_insulation_lining_propellant_remarks']
                motorModel.acceptance_casting = request['Acceptance_of_casting']
                motorModel.acceptance_casting_date = request['Acceptance_of_casting_date']
                motorModel.acceptance_casting_remarks = request['Acceptance_of_casting_remarks']
                motorModel.sandblasting = request['Sand_blasting']
                motorModel.sandblasting_date = request['Sand_blasting_date']
                motorModel.sandblasting_remarks = request['Sand_blasting_remarks']

                motorModel.insulation = request['motor_Insulation']
                motorModel.insulation_date = request['motor_Insulation_date']
                motorModel.insulation_remarks = request['motor_Insulation_remarks']

                motorModel.ut_rt_insulated_case = request['UTandRTofInsulated_Case']
                motorModel.ut_rt_insulated_case_date = request['UTandRTofInsulated_Case_date']
                motorModel.ut_rt_insulated_case_remarks = request['UTandRTofInsulated_Case_Remarks']

                motorModel.acceptance_silver_material = request['AcceptanceofSilver_Material']
                motorModel.acceptance_silver_material_date = request['AcceptanceofSilver_Material_date']
                motorModel.acceptance_silver_material_remarks = request['AcceptanceofSilver_Material_Remarks']

                motorModel.silver_application = request['Silver_Application']
                motorModel.silver_application_date = request['Silver_Application_date']
                motorModel.silver_application_remarks = request['Silver_Application_Remarks']

                motorModel.formulation_tailoring_liner_propellant = request['Formulationtailoringoflinerand_propellant']
                motorModel.formulation_tailoring_liner_propellant_date = request['Formulationtailoringoflinerand_propellant_date']
                motorModel.formulation_tailoring_liner_propellant_remarks = request[
                    'Formulationtailoringoflinerand_propellant_remarks']

                motorModel.conditioning_raw_materials = request['ConditioningofRaw_Material']
                motorModel.conditioning_raw_materials_date = request['ConditioningofRaw_Material_date']
                motorModel.conditioning_raw_materials_remarks = request['ConditioningofRaw_Material_remarks']

                motorModel.lining = request['Lining_motor']
                motorModel.lining_date = request['Lining_motor_date']
                motorModel.lining_remarks = request['Lining_motor_remarks']

                motorModel.casting = request['Casting_motor']
                motorModel.casting_date = request['Casting_motor_date']
                motorModel.casting_remarks = request['Casting_motor_remarks']

                motorModel.curing = request['Curing_motor']
                motorModel.curing_date = request['Curing_motor_date']
                motorModel.curing_remarks = request['Curing_motor_remarks']

                motorModel.liner_mechanical_properties = request['Liner_Mechanical_Properties']
                motorModel.liner_mechanical_properties_date = request['Liner_Mechanical_Properties_date']
                motorModel.liner_mechanical_properties_remarks = request['Liner_Mechanical_Properties_Remarks']

                motorModel.propellant_mechanical_properties = request['Propellant_Mechanical_Properties']
                motorModel.propellant_mechanical_properties_date = request['Propellant_Mechanical_Properties_date']
                motorModel.propellant_mechanical_properties_remarks = request[
                    'Propellant_Mechanical_Properties_Remarks']

                motorModel.interface_bond_strength = request['Interfacebond_strength']
                motorModel.interface_bond_strength_date = request['Interfacebond_strength_date']
                motorModel.interface_bond_strength_remarks = request['Interfacebond_strength_Remarks']

                motorModel.propellant_burn_rate = request['Propellant_burn_rate']
                motorModel.propellant_burn_rate_date = request['Propellant_burn_rate_date']
                motorModel.propellant_burn_rate_remarks = request['Propellant_burn_rate_Remarks']

                motorModel.trimming_Propellant_grain = request['Trimming_of_propellant_grain']
                motorModel.trimming_Propellant_grain_date = request['Trimming_of_propellant_grain_date']
                motorModel.trimming_Propellant_grain_remarks = request['Trimming_of_propellant_grain_remarks']

                motorModel.mass_liner_insulation_propellant_srm = request['mass_liner_insulation_propellant_srm']
                motorModel.mass_liner_insulation_propellant_srm_date = request['mass_liner_insulation_propellant_srm_date']
                motorModel.mass_liner_insulation_propellant_srm_remarks = request[
                    'mass_liner_insulation_propellant_srm_remarks']
                motorModel.ut_endoscopy_rt_grain = request['ut_endoscopy_rt_grain']
                motorModel.ut_endoscopy_rt_grain_date = request['ut_endoscopy_rt_grain_date']
                motorModel.ut_endoscopy_rt_grain_remarks = request['ut_endoscopy_rt_grain_remarks']

                motorModel.battery_type = request['battery_type']
                motorModel.tb_type = request['tb_type']
                motorModel.battery_id = request['battery_id']
                motorModel.lot_id = request['lot_id']

                motorModel.matiral_qualified_date = request['matiral_qualified_date']
                motorModel.matiral_qualified_status = request['matiral_qualified_status']
                motorModel.matiral_qualified_remarks = request['matiral_qualified_remarks']

                motorModel.compo_manufacturing_date = request['compo_manufacturing_date']
                motorModel.compo_manufacturing_status = request['compo_manufacturing_status']
                motorModel.compo_manufacturing_remarks = request['compo_manufacturing_remarks']

                motorModel.powerpack_assembly_date = request['powerpack_assembly_date']
                motorModel.powerpack_assembly_status = request['powerpack_assembly_status']
                motorModel.powerpack_assembly_remarks = request['powerpack_assembly_remarks']

                motorModel.powerpack_testing_date = request['powerpack_testing_date']
                motorModel.powerpack_testing_status = request['powerpack_testing_status']
                motorModel.powerpack_testing_remarks = request['powerpack_testing_remarks']

                motorModel.raw_material_inspection_date = request['raw_material_inspection_date']
                motorModel.raw_material_inspection_status = request['raw_material_inspection_status']
                motorModel.raw_material_inspection_remarks = request['raw_material_inspection_remarks']

                motorModel.pressing_electrode_date = request['pressing_electrode_date']
                motorModel.pressing_electrode_status = request['pressing_electrode_status']
                motorModel.pressing_electrode_remarks = request['pressing_electrode_remarks']

                motorModel.formation_process_date = request['formation_process_date']
                motorModel.formation_process_status = request['formation_process_status']
                motorModel.formation_process_remarks = request['formation_process_remarks']

                motorModel.assembly_process_date = request['assembly_process_date']
                motorModel.assembly_process_status = request['assembly_process_status']
                motorModel.assembly_process_remarks = request['assembly_process_remarks']

                motorModel.battery_testing_date = request['battery_testing_date']
                motorModel.battery_testing_status = request['battery_testing_status']
                motorModel.battery_testing_remarks = request['battery_testing_remarks']

                motorModel.final_qualification_date = request['final_qualification_date']
                motorModel.final_qualification_status = request['final_qualification_status']
                motorModel.final_qualification_remarks = request['final_qualification_remarks']

                motorModel.pd_type = request['pd_type']
                motorModel.pd_id = request['pd_id']

                motorModel.qualification_raw_material_date = request['qualification_raw_material_date']
                motorModel.qualification_raw_material_status = request['qualification_raw_material_status']
                motorModel.qualification_raw_material_remarks = request['qualification_raw_material_remarks']

                motorModel.filling_date = request['filling_date']
                motorModel.filling_status = request['filling_status']
                motorModel.filling_remarks = request['filling_remarks']

                motorModel.assembling_integration_date = request['assembling_integration_date']
                motorModel.assembling_integration_status = request['assembling_integration_status']
                motorModel.assembling_integration_remarks = request['assembling_integration_remarks']

                motorModel.qualification_testing_date = request['qualification_testing_date']
                motorModel.qualification_testing_status = request['qualification_testing_status']
                motorModel.qualification_testing_remarks = request['qualification_testing_remarks']

                motorModel.performance_testing_date = request['performance_testing_date']
                motorModel.performance_testing_status = request['performance_testing_status']
                motorModel.performance_testing_remarks = request['performance_testing_remarks']

                motorModel.ncr_status = request['ncr_status']
                motorModel.ncr_status_remarks = request['ncr_status_ramarks']
                motorModel.overall_status = request['overall_status']
                motorModel.overall_remarks = request['overall_status_ramarks']

                if request['base64File'] != '':
                    motorModel.attachments = request['base64File']
                motorModel.save()
                return JsonResponse({'status': 'True', 'message': "Motor Created Successfully!"},
                                    status=200)
            else:
                get_motor = ActiveMotors.objects.filter(id=id).first()
                if get_motor is not None:
                    if request['qualification_insulation_lining_propellant_rm'] != get_motor.qualification_insulation_lining_propellant_rm \
                        or request['Acceptance_of_casting'] != get_motor.acceptance_casting  or  request['Sand_blasting'] != get_motor.sandblasting \
                        or request['motor_Insulation'] != get_motor.insulation or request['UTandRTofInsulated_Case'] != get_motor.ut_rt_insulated_case \
                        or request['UTandRTofInsulated_Case_date'] != get_motor.ut_rt_insulated_case_date or request['AcceptanceofSilver_Material'] != get_motor.acceptance_silver_material\
                        or request['Silver_Application'] != get_motor.silver_application or request[ 'Formulationtailoringoflinerand_propellant'] != get_motor.formulation_tailoring_liner_propellant \
                        or request['ConditioningofRaw_Material'] != get_motor.conditioning_raw_materials or request['Lining_motor'] != get_motor.lining\
                        or request['Casting_motor'] != get_motor.casting or request['Curing_motor'] != get_motor.curing \
                        or request['Liner_Mechanical_Properties'] != get_motor.liner_mechanical_properties or request['Propellant_Mechanical_Properties'] != get_motor.propellant_mechanical_properties \
                        or request['Interfacebond_strength'] != get_motor.interface_bond_strength or request['Propellant_burn_rate'] != get_motor.propellant_burn_rate \
                        or request['Trimming_of_propellant_grain'] != get_motor.trimming_Propellant_grain or request['mass_liner_insulation_propellant_srm']  != get_motor.mass_liner_insulation_propellant_srm \
                        or request['ut_endoscopy_rt_grain'] != get_motor.ut_endoscopy_rt_grain or request['ncr_status'] != get_motor.ncr_status \
                        or request['matiral_qualified_status'] != get_motor.matiral_qualified_status  or request['compo_manufacturing_status'] != get_motor.compo_manufacturing_status \
                        or request['powerpack_assembly_status'] != get_motor.powerpack_assembly_status  or request['powerpack_testing_status'] != get_motor.powerpack_testing_status \
                        or request['raw_material_inspection_status'] != get_motor.raw_material_inspection_status or request['pressing_electrode_status'] != get_motor.pressing_electrode_status \
                        or request['formation_process_status'] != get_motor.formation_process_status or request['assembly_process_status'] != get_motor.assembly_process_status \
                        or request['battery_testing_status'] != get_motor.battery_testing_status or request['final_qualification_status'] != get_motor.final_qualification_status \
                        or request['qualification_raw_material_status'] != get_motor.qualification_raw_material_status or request['filling_status'] != get_motor.filling_status \
                        or request['assembling_integration_status'] != get_motor.assembling_integration_status or request['qualification_testing_status'] != get_motor.qualification_testing_status \
                        or request['performance_testing_status'] != get_motor.performance_testing_status or request['qualification_testing_status'] != get_motor.qualification_testing_status:
                        MotorHistory = ActiveMotorsHistory()
                        MotorHistory.active_motor_id = get_motor.id
                        MotorHistory.system_type = get_motor.system_type
                        MotorHistory.system_name = get_motor.system_name
                        MotorHistory.organization = get_motor.organization
                        MotorHistory.testing_type = get_motor.testing_type
                        MotorHistory.testing_date = get_motor.testing_date
                        MotorHistory.component_type = get_motor.component_type
                        MotorHistory.motor_id = get_motor.motor_id
                        MotorHistory.qualification_insulation_lining_propellant_rm = get_motor.qualification_insulation_lining_propellant_rm
                        MotorHistory.qualification_insulation_lining_propellant_rm_date = get_motor.qualification_insulation_lining_propellant_rm_date
                        MotorHistory.qualification_insulation_lining_propellant_rm_remarks = get_motor.qualification_insulation_lining_propellant_rm_remarks
                        MotorHistory.acceptance_casting = get_motor.acceptance_casting
                        MotorHistory.acceptance_casting_date = get_motor.acceptance_casting_date
                        MotorHistory.acceptance_casting_remarks = get_motor.acceptance_casting_remarks

                        MotorHistory.sandblasting = get_motor.sandblasting
                        MotorHistory.sandblasting_date = get_motor.sandblasting_date
                        MotorHistory.sandblasting_remarks = get_motor.sandblasting_remarks

                        MotorHistory.insulation = get_motor.insulation
                        MotorHistory.insulation_date = get_motor.insulation_date
                        MotorHistory.insulation_remarks = get_motor.insulation_remarks

                        MotorHistory.ut_rt_insulated_case = get_motor.ut_rt_insulated_case
                        MotorHistory.ut_rt_insulated_case_date = get_motor.ut_rt_insulated_case_date
                        MotorHistory.ut_rt_insulated_case_remarks = get_motor.ut_rt_insulated_case_remarks

                        MotorHistory.acceptance_silver_material = get_motor.acceptance_silver_material
                        MotorHistory.acceptance_silver_material_date = get_motor.acceptance_silver_material_date
                        MotorHistory.acceptance_silver_material_remarks = get_motor.acceptance_silver_material_remarks

                        MotorHistory.silver_application = get_motor.silver_application
                        MotorHistory.silver_application_date = get_motor.silver_application_date
                        MotorHistory.silver_application_remarks = get_motor.silver_application_remarks

                        MotorHistory.formulation_tailoring_liner_propellant = get_motor.formulation_tailoring_liner_propellant
                        MotorHistory.formulation_tailoring_liner_propellant_date = get_motor.formulation_tailoring_liner_propellant_date
                        MotorHistory.formulation_tailoring_liner_propellant_remarks = get_motor.formulation_tailoring_liner_propellant_remarks

                        MotorHistory.conditioning_raw_materials = get_motor.conditioning_raw_materials
                        MotorHistory.conditioning_raw_materials_date = get_motor.conditioning_raw_materials_date
                        MotorHistory.conditioning_raw_materials_remarks = get_motor.conditioning_raw_materials_remarks

                        MotorHistory.lining = get_motor.lining
                        MotorHistory.lining_date = get_motor.lining_date
                        MotorHistory.lining_remarks = get_motor.lining_remarks

                        MotorHistory.casting = get_motor.casting
                        MotorHistory.casting_date = get_motor.casting_date
                        MotorHistory.casting_remarks = get_motor.casting_remarks

                        MotorHistory.curing = get_motor.curing
                        MotorHistory.curing_date = get_motor.curing_date
                        MotorHistory.curing_remarks = get_motor.curing_remarks

                        MotorHistory.liner_mechanical_properties = get_motor.liner_mechanical_properties
                        MotorHistory.liner_mechanical_properties_date = get_motor.liner_mechanical_properties_date
                        MotorHistory.liner_mechanical_properties_remarks = get_motor.liner_mechanical_properties_remarks

                        MotorHistory.propellant_mechanical_properties = get_motor.propellant_mechanical_properties
                        MotorHistory.propellant_mechanical_properties_date = get_motor.propellant_mechanical_properties_date
                        MotorHistory.propellant_mechanical_properties_remarks = get_motor.propellant_mechanical_properties_remarks

                        MotorHistory.interface_bond_strength = get_motor.interface_bond_strength
                        MotorHistory.interface_bond_strength_date = get_motor.interface_bond_strength_date
                        MotorHistory.interface_bond_strength_remarks = get_motor.interface_bond_strength_remarks

                        MotorHistory.propellant_burn_rate = get_motor.propellant_burn_rate
                        MotorHistory.propellant_burn_rate_date = get_motor.propellant_burn_rate_date
                        MotorHistory.propellant_burn_rate_remarks = get_motor.propellant_burn_rate_remarks

                        MotorHistory.trimming_Propellant_grain = get_motor.trimming_Propellant_grain
                        MotorHistory.trimming_Propellant_grain_date = get_motor.trimming_Propellant_grain_date
                        MotorHistory.trimming_Propellant_grain_remarks = get_motor.trimming_Propellant_grain_remarks

                        MotorHistory.mass_liner_insulation_propellant_srm = get_motor.mass_liner_insulation_propellant_srm
                        MotorHistory.mass_liner_insulation_propellant_srm_date = get_motor.mass_liner_insulation_propellant_srm_date
                        MotorHistory.mass_liner_insulation_propellant_srm_remarks = get_motor.mass_liner_insulation_propellant_srm_remarks
                        MotorHistory.ut_endoscopy_rt_grain = get_motor.ut_endoscopy_rt_grain
                        MotorHistory.ut_endoscopy_rt_grain_date = get_motor.ut_endoscopy_rt_grain_date
                        MotorHistory.ut_endoscopy_rt_grain_remarks = get_motor.ut_endoscopy_rt_grain_remarks

                        MotorHistory.battery_type = get_motor.battery_type
                        MotorHistory.tb_type = get_motor.tb_type
                        MotorHistory.battery_id = get_motor.battery_id
                        MotorHistory.lot_id = get_motor.lot_id

                        MotorHistory.matiral_qualified_date = get_motor.matiral_qualified_date
                        MotorHistory.matiral_qualified_status = get_motor.matiral_qualified_status
                        MotorHistory.matiral_qualified_remarks = get_motor.matiral_qualified_remarks

                        MotorHistory.compo_manufacturing_date = get_motor.compo_manufacturing_date
                        MotorHistory.compo_manufacturing_status = get_motor.compo_manufacturing_status
                        MotorHistory.compo_manufacturing_remarks = get_motor.compo_manufacturing_remarks

                        MotorHistory.powerpack_assembly_date = get_motor.powerpack_assembly_date
                        MotorHistory.powerpack_assembly_status = get_motor.powerpack_assembly_status
                        MotorHistory.powerpack_assembly_remarks = get_motor.powerpack_assembly_remarks

                        MotorHistory.powerpack_testing_date = get_motor.powerpack_testing_date
                        MotorHistory.powerpack_testing_status = get_motor.powerpack_testing_status
                        MotorHistory.powerpack_testing_remarks = get_motor.powerpack_testing_remarks

                        MotorHistory.raw_material_inspection_date = get_motor.raw_material_inspection_date
                        MotorHistory.raw_material_inspection_status = get_motor.raw_material_inspection_status
                        MotorHistory.raw_material_inspection_remarks = get_motor.raw_material_inspection_remarks

                        MotorHistory.pressing_electrode_date = get_motor.pressing_electrode_date
                        MotorHistory.pressing_electrode_status = get_motor.pressing_electrode_status
                        MotorHistory.pressing_electrode_remarks = get_motor.pressing_electrode_remarks

                        MotorHistory.formation_process_date = get_motor.formation_process_date
                        MotorHistory.formation_process_status = get_motor.formation_process_status
                        MotorHistory.formation_process_remarks = get_motor.formation_process_remarks

                        MotorHistory.assembly_process_date = get_motor.assembly_process_date
                        MotorHistory.assembly_process_status = get_motor.assembly_process_status
                        MotorHistory.assembly_process_remarks = get_motor.assembly_process_remarks

                        MotorHistory.battery_testing_date = get_motor.battery_testing_date
                        MotorHistory.battery_testing_status = get_motor.battery_testing_status
                        MotorHistory.battery_testing_remarks = get_motor.battery_testing_remarks

                        MotorHistory.final_qualification_date = get_motor.final_qualification_date
                        MotorHistory.final_qualification_status = get_motor.final_qualification_status
                        MotorHistory.final_qualification_remarks = get_motor.final_qualification_remarks

                        MotorHistory.pd_type = get_motor.pd_type
                        MotorHistory.pd_id = get_motor.pd_id

                        MotorHistory.qualification_raw_material_date = get_motor.qualification_raw_material_date
                        MotorHistory.qualification_raw_material_status = get_motor.qualification_raw_material_status
                        MotorHistory.qualification_raw_material_remarks = get_motor.qualification_raw_material_remarks

                        MotorHistory.filling_date = get_motor.filling_date
                        MotorHistory.filling_status = get_motor.filling_status
                        MotorHistory.filling_remarks = get_motor.filling_remarks

                        MotorHistory.assembling_integration_date = get_motor.assembling_integration_date
                        MotorHistory.assembling_integration_status = get_motor.assembling_integration_status
                        MotorHistory.assembling_integration_remarks = get_motor.assembling_integration_remarks

                        MotorHistory.qualification_testing_date = get_motor.qualification_testing_date
                        MotorHistory.qualification_testing_status = get_motor.qualification_testing_status
                        MotorHistory.qualification_testing_remarks = get_motor.qualification_testing_remarks

                        MotorHistory.performance_testing_date = get_motor.performance_testing_date
                        MotorHistory.performance_testing_status = get_motor.performance_testing_status
                        MotorHistory.performance_testing_remarks = get_motor.performance_testing_remarks

                        MotorHistory.ncr_status = get_motor.ncr_status
                        MotorHistory.ncr_status_remarks = get_motor.ncr_status_remarks
                        MotorHistory.overall_status = get_motor.overall_status
                        MotorHistory.overall_remarks = get_motor.overall_remarks
                        MotorHistory.attachments = get_motor.attachments
                        MotorHistory.save()

                    get_motor.system_type = request['system_type']
                    get_motor.system_name = request['system_name']
                    get_motor.organization = request['organization']
                    get_motor.testing_type = request['testing_type']
                    get_motor.testing_date = request['testing_date']
                    get_motor.component_type = request['component_type']
                    get_motor.motor_id = request['motor_id']
                    get_motor.qualification_insulation_lining_propellant_rm = request[
                        'qualification_insulation_lining_propellant_rm']
                    get_motor.qualification_insulation_lining_propellant_rm_date = request[
                        'qualification_insulation_lining_propellant_rm_date']
                    get_motor.qualification_insulation_lining_propellant_rm_remarks = request[
                        'qualification_insulation_lining_propellant_remarks']
                    get_motor.acceptance_casting = request['Acceptance_of_casting']
                    get_motor.acceptance_casting_date = request['Acceptance_of_casting_date']
                    get_motor.acceptance_casting_remarks = request['Acceptance_of_casting_remarks']
                    get_motor.sandblasting = request['Sand_blasting']
                    get_motor.sandblasting_date = request['Sand_blasting_date']
                    get_motor.sandblasting_remarks = request['Sand_blasting_remarks']

                    get_motor.insulation = request['motor_Insulation']
                    get_motor.insulation_date = request['motor_Insulation_date']
                    get_motor.insulation_remarks = request['motor_Insulation_remarks']

                    get_motor.ut_rt_insulated_case = request['UTandRTofInsulated_Case']
                    get_motor.ut_rt_insulated_case_date = request['UTandRTofInsulated_Case_date']
                    get_motor.ut_rt_insulated_case_remarks = request['UTandRTofInsulated_Case_Remarks']

                    get_motor.acceptance_silver_material = request['AcceptanceofSilver_Material']
                    get_motor.acceptance_silver_material_date = request['AcceptanceofSilver_Material_date']
                    get_motor.acceptance_silver_material_remarks = request['AcceptanceofSilver_Material_Remarks']

                    get_motor.silver_application = request['Silver_Application']
                    get_motor.silver_application_date = request['Silver_Application_date']
                    get_motor.silver_application_remarks = request['Silver_Application_Remarks']

                    get_motor.formulation_tailoring_liner_propellant = request[
                        'Formulationtailoringoflinerand_propellant']
                    get_motor.formulation_tailoring_liner_propellant_date = request[
                        'Formulationtailoringoflinerand_propellant_date']
                    get_motor.formulation_tailoring_liner_propellant_remarks = request[
                        'Formulationtailoringoflinerand_propellant_remarks']

                    get_motor.conditioning_raw_materials = request['ConditioningofRaw_Material']
                    get_motor.conditioning_raw_materials_date = request['ConditioningofRaw_Material_date']
                    get_motor.conditioning_raw_materials_remarks = request['ConditioningofRaw_Material_remarks']

                    get_motor.lining = request['Lining_motor']
                    get_motor.lining_date = request['Lining_motor_date']
                    get_motor.lining_remarks = request['Lining_motor_remarks']

                    get_motor.casting = request['Casting_motor']
                    get_motor.casting_date = request['Casting_motor_date']
                    get_motor.casting_remarks = request['Casting_motor_remarks']

                    get_motor.curing = request['Curing_motor']
                    get_motor.curing_date = request['Curing_motor_date']
                    get_motor.curing_remarks = request['Curing_motor_remarks']

                    get_motor.liner_mechanical_properties = request['Liner_Mechanical_Properties']
                    get_motor.liner_mechanical_properties_date = request['Liner_Mechanical_Properties_date']
                    get_motor.liner_mechanical_properties_remarks = request['Liner_Mechanical_Properties_Remarks']

                    get_motor.propellant_mechanical_properties = request['Propellant_Mechanical_Properties']
                    get_motor.propellant_mechanical_properties_date = request['Propellant_Mechanical_Properties_date']
                    get_motor.propellant_mechanical_properties_remarks = request[
                        'Propellant_Mechanical_Properties_Remarks']

                    get_motor.interface_bond_strength = request['Interfacebond_strength']
                    get_motor.interface_bond_strength_date = request['Interfacebond_strength_date']
                    get_motor.interface_bond_strength_remarks = request['Interfacebond_strength_Remarks']

                    get_motor.propellant_burn_rate = request['Propellant_burn_rate']
                    get_motor.propellant_burn_rate_date = request['Propellant_burn_rate_date']
                    get_motor.propellant_burn_rate_remarks = request['Propellant_burn_rate_Remarks']

                    get_motor.trimming_Propellant_grain = request['Trimming_of_propellant_grain']
                    get_motor.trimming_Propellant_grain_date = request['Trimming_of_propellant_grain_date']
                    get_motor.trimming_Propellant_grain_remarks = request['Trimming_of_propellant_grain_remarks']

                    get_motor.mass_liner_insulation_propellant_srm = request['mass_liner_insulation_propellant_srm']
                    get_motor.mass_liner_insulation_propellant_srm_date = request[
                        'mass_liner_insulation_propellant_srm_date']
                    get_motor.mass_liner_insulation_propellant_srm_remarks = request[
                        'mass_liner_insulation_propellant_srm_remarks']
                    get_motor.ut_endoscopy_rt_grain = request['ut_endoscopy_rt_grain']
                    get_motor.ut_endoscopy_rt_grain_date = request['ut_endoscopy_rt_grain_date']
                    get_motor.ut_endoscopy_rt_grain_remarks = request['ut_endoscopy_rt_grain_remarks']

                    get_motor.battery_type = request['battery_type']
                    get_motor.tb_type = request['tb_type']
                    get_motor.battery_id = request['battery_id']
                    get_motor.lot_id = request['lot_id']

                    get_motor.matiral_qualified_date = request['matiral_qualified_date']
                    get_motor.matiral_qualified_status = request['matiral_qualified_status']
                    get_motor.matiral_qualified_remarks = request['matiral_qualified_remarks']

                    get_motor.compo_manufacturing_date = request['compo_manufacturing_date']
                    get_motor.compo_manufacturing_status = request['compo_manufacturing_status']
                    get_motor.compo_manufacturing_remarks = request['compo_manufacturing_remarks']

                    get_motor.powerpack_assembly_date = request['powerpack_assembly_date']
                    get_motor.powerpack_assembly_status = request['powerpack_assembly_status']
                    get_motor.powerpack_assembly_remarks = request['powerpack_assembly_remarks']

                    get_motor.powerpack_testing_date = request['powerpack_testing_date']
                    get_motor.powerpack_testing_status = request['powerpack_testing_status']
                    get_motor.powerpack_testing_remarks = request['powerpack_testing_remarks']

                    get_motor.raw_material_inspection_date = request['raw_material_inspection_date']
                    get_motor.raw_material_inspection_status = request['raw_material_inspection_status']
                    get_motor.raw_material_inspection_remarks = request['raw_material_inspection_remarks']

                    get_motor.pressing_electrode_date = request['pressing_electrode_date']
                    get_motor.pressing_electrode_status = request['pressing_electrode_status']
                    get_motor.pressing_electrode_remarks = request['pressing_electrode_remarks']

                    get_motor.formation_process_date = request['formation_process_date']
                    get_motor.formation_process_status = request['formation_process_status']
                    get_motor.formation_process_remarks = request['formation_process_remarks']

                    get_motor.assembly_process_date = request['assembly_process_date']
                    get_motor.assembly_process_status = request['assembly_process_status']
                    get_motor.assembly_process_remarks = request['assembly_process_remarks']

                    get_motor.battery_testing_date = request['battery_testing_date']
                    get_motor.battery_testing_status = request['battery_testing_status']
                    get_motor.battery_testing_remarks = request['battery_testing_remarks']

                    get_motor.final_qualification_date = request['final_qualification_date']
                    get_motor.final_qualification_status = request['final_qualification_status']
                    get_motor.final_qualification_remarks = request['final_qualification_remarks']

                    get_motor.pd_type = request['pd_type']
                    get_motor.pd_id = request['pd_id']

                    get_motor.qualification_raw_material_date = request['qualification_raw_material_date']
                    get_motor.qualification_raw_material_status = request['qualification_raw_material_status']
                    get_motor.qualification_raw_material_remarks = request['qualification_raw_material_remarks']

                    get_motor.filling_date = request['filling_date']
                    get_motor.filling_status = request['filling_status']
                    get_motor.filling_remarks = request['filling_remarks']

                    get_motor.assembling_integration_date = request['assembling_integration_date']
                    get_motor.assembling_integration_status = request['assembling_integration_status']
                    get_motor.assembling_integration_remarks = request['assembling_integration_remarks']

                    get_motor.qualification_testing_date = request['qualification_testing_date']
                    get_motor.qualification_testing_status = request['qualification_testing_status']
                    get_motor.qualification_testing_remarks = request['qualification_testing_remarks']

                    get_motor.performance_testing_date = request['performance_testing_date']
                    get_motor.performance_testing_status = request['performance_testing_status']
                    get_motor.performance_testing_remarks = request['performance_testing_remarks']

                    get_motor.ncr_status = request['ncr_status']
                    get_motor.ncr_status_remarks = request['ncr_status_ramarks']
                    get_motor.overall_status = request['overall_status']
                    get_motor.overall_remarks = request['overall_status_ramarks']
                    if request['base64File'] != '':
                        get_motor.attachments = request['base64File']
                    get_motor.save()
                    return JsonResponse({'status': 'True', 'message': "Motor Created Successfully!"},
                                        status=200)
        except Exception as e:
            return JsonResponse({'status': 'False', "message": "Motor Not Saved"}, status=500)

    @staticmethod
    def GetActiveMotorList(request, self=None):
        try:
            current_year = request.query_params.get('year')
            current_sys = request.query_params.get('system')
            current_comp = request.query_params.get('component')
            current_org = request.query_params.get('selected_organization')
            current_bat = request.query_params.get('selected_battery')
            current_lot = request.query_params.get('selected_lot_id')
            system_type = request.query_params.get('selected_system_type')

            ParentStatus = request.query_params['parent_status']
            ChildStatus = request.query_params['child_status']


            today = date.today()
            dataList = []

            def get_filter(field_name, filter_condition, filter_value):
                if filter_condition.strip() == "contains":
                    kwargs = {
                        '{0}__icontains'.format(field_name): filter_value
                    }
                    return Q(**kwargs)

                if filter_condition.strip() == "not_equal":
                    kwargs = {
                        '{0}__iexact'.format(field_name): filter_value
                    }
                    return ~Q(**kwargs)

                if filter_condition.strip() == "starts_with":
                    kwargs = {
                        '{0}__istartswith'.format(field_name): filter_value
                    }
                    return Q(**kwargs)
                if filter_condition.strip() == "equal":
                    kwargs = {
                        '{0}__iexact'.format(field_name): filter_value
                    }
                    return Q(**kwargs)

                if filter_condition.strip() == "not_equal":
                    kwargs = {
                        '{0}__ne'.format(field_name): filter_value
                    }
                    return ~Q(**kwargs)
                if filter_condition.strip() == "less_than":
                    kwargs = {
                        '{0}__lt'.format(field_name): filter_value
                    }
                    return ~Q(**kwargs)
                if filter_condition.strip() == "greater_than":
                    kwargs = {
                        '{0}__gt'.format(field_name): filter_value
                    }
                    return ~Q(**kwargs)

            typeQuery = Q()
            filter_objects = Q()

            if current_year !='':
                filter_objects &= get_filter(
                    'testing_date__year', 'equal',
                    current_year)

            if system_type !='':
                filter_objects &= get_filter(
                    'system_type', 'equal', system_type)
            if current_org !='':
                filter_objects &= get_filter(
                    'organization', 'equal', current_org)
            if current_bat !='':
                filter_objects &= get_filter(
                    'battery_type', 'equal', current_bat)
            if current_lot !='':
                filter_objects &= get_filter(
                    'lot_id', 'equal', current_lot)

            if current_sys !='':
                filter_objects &= get_filter(
                    'system_name', 'equal',
                    current_sys)
            if current_comp !='':
                filter_objects &= get_filter(
                    'component_type', 'equal',
                    current_comp)

            dataList = ActiveMotors.objects.filter(filter_objects)
            if ParentStatus != '':
                ListItems = []
                if ParentStatus == 'Qualification of raw material of insulation':
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(qualification_insulation_lining_propellant_rm=ChildStatus)
                    if ChildStatus == 'Current Count':
                        for data in dataList:
                            if data.qualification_insulation_lining_propellant_rm == 'Under process' or data.qualification_insulation_lining_propellant_rm == 'Observation(same stage)' or data.qualification_insulation_lining_propellant_rm == 'Halt':
                                ListItems.append(data)
                elif ParentStatus == 'Acceptance of casting':
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(acceptance_casting=ChildStatus)
                    if ChildStatus == 'Current Count':
                        for data in dataList:
                            if data.acceptance_casting == 'Under process' or data.acceptance_casting == 'Observation(same stage)' or data.acceptance_casting == 'Halt':
                                ListItems.append(data)
                elif ParentStatus == 'Sandblasting':
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(sandblasting=ChildStatus)
                    if ChildStatus == 'Current Count':
                        for data in dataList:
                            if data.sandblasting == 'Under process' or data.sandblasting == 'Observation(same stage)' or data.sandblasting == 'Halt':
                                ListItems.append(data)
                elif ParentStatus == 'Insulation':
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(insulation=ChildStatus)
                    if ChildStatus == 'Current Count':
                        for data in dataList:
                            if data.insulation == 'Under process' or data.insulation == 'Observation(same stage)' or data.insulation == 'Halt':
                                ListItems.append(data)
                elif ParentStatus == 'UT and RT of Insulated Case':
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(ut_rt_insulated_case=ChildStatus)
                    if ChildStatus == 'Current Count':
                        for data in dataList:
                            if data.ut_rt_insulated_case == 'Under process' or data.ut_rt_insulated_case == 'Observation(same stage)' or data.ut_rt_insulated_case == 'Halt':
                                ListItems.append(data)
                elif ParentStatus == 'Acceptance of Silver Material':
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(acceptance_silver_material=ChildStatus)
                    if ChildStatus == 'Current Count':
                        for data in dataList:
                            if data.acceptance_silver_material == 'Under process' or data.acceptance_silver_material == 'Observation(same stage)' or data.acceptance_silver_material == 'Halt':
                                ListItems.append(data)
                elif ParentStatus == 'Silver Application':
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(silver_application=ChildStatus)
                    if ChildStatus == 'Current Count':
                        for data in dataList:
                            if data.silver_application == 'Under process' or data.silver_application == 'Observation(same stage)' or data.silver_application == 'Halt':
                                ListItems.append(data)
                elif ParentStatus == 'Formulation tailoring':
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(formulation_tailoring_liner_propellant=ChildStatus)
                    if ChildStatus == 'Current Count':
                        for data in dataList:
                            if data.formulation_tailoring_liner_propellant == 'Under process' or data.formulation_tailoring_liner_propellant == 'Observation(same stage)' or data.formulation_tailoring_liner_propellant == 'Halt':
                                ListItems.append(data)
                elif ParentStatus == 'Conditioning of Raw Material':
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(conditioning_raw_materials=ChildStatus)
                    if ChildStatus == 'Current Count':
                        for data in dataList:
                            if data.conditioning_raw_materials == 'Under process' or data.conditioning_raw_materials == 'Observation(same stage)' or data.conditioning_raw_materials == 'Halt':
                                ListItems.append(data)
                elif ParentStatus == 'Lining':
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(lining=ChildStatus)
                    if ChildStatus == 'Current Count':
                        for data in dataList:
                            if data.lining == 'Under process' or data.lining == 'Observation(same stage)' or data.lining == 'Halt':
                                ListItems.append(data)
                elif ParentStatus == 'Casting':
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(casting=ChildStatus)
                    if ChildStatus == 'Current Count':
                        for data in dataList:
                            if data.casting == 'Under process' or data.casting == 'Observation(same stage)' or data.casting == 'Halt':
                                ListItems.append(data)
                elif ParentStatus == 'Curing':
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(blt_status=ChildStatus)
                    if ChildStatus == 'Current Count':
                        for data in dataList:
                            if data.curing == 'Under process' or data.curing == 'Observation(same stage)' or data.curing == 'Halt':
                                ListItems.append(data)
                elif ParentStatus == 'Liner Mechanical Properties':
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(liner_mechanical_properties=ChildStatus)
                    if ChildStatus == 'Current Count':
                        for data in dataList:
                            if data.liner_mechanical_properties == 'Under process' or data.liner_mechanical_properties == 'Observation(same stage)' or data.liner_mechanical_properties == 'Halt':
                                ListItems.append(data)
                elif ParentStatus == 'Propellant Mechanical Properties':
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(propellant_mechanical_properties=ChildStatus)
                    if ChildStatus == 'Current Count':
                        for data in dataList:
                            if data.propellant_mechanical_properties == 'Under process' or data.propellant_mechanical_properties == 'Observation(same stage)' or data.propellant_mechanical_properties == 'Halt':
                                ListItems.append(data)
                elif ParentStatus == 'Interface bond strength':
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(interface_bond_strength=ChildStatus)
                    if ChildStatus == 'Current Count':
                        for data in dataList:
                            if data.interface_bond_strength == 'Under process' or data.interface_bond_strength == 'Observation(same stage)' or data.interface_bond_strength == 'Halt':
                                ListItems.append(data)
                elif ParentStatus == 'Propellant burn rate':
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(propellant_burn_rate=ChildStatus)
                    if ChildStatus == 'Current Count':
                        for data in dataList:
                            if data.propellant_burn_rate == 'Under process' or data.propellant_burn_rate == 'Observation(same stage)' or data.propellant_burn_rate == 'Halt':
                                ListItems.append(data)
                elif ParentStatus == 'Trimming of propellant grain':
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(trimming_Propellant_grain=ChildStatus)
                    if ChildStatus == 'Current Count':
                        for data in dataList:
                            if data.trimming_Propellant_grain == 'Under process' or data.trimming_Propellant_grain == 'Observation(same stage)' or data.trimming_Propellant_grain == 'Halt':
                                ListItems.append(data)
                elif ParentStatus == 'Mass of linear':
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(mass_liner_insulation_propellant_srm=ChildStatus)
                    if ChildStatus == 'Current Count':
                        for data in dataList:
                            if data.mass_liner_insulation_propellant_srm == 'Under process' or data.mass_liner_insulation_propellant_srm == 'Observation(same stage)' or data.mass_liner_insulation_propellant_srm == 'Halt':
                                ListItems.append(data)
                elif ParentStatus == 'UT, endoscopy and RT of grain':
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(ut_endoscopy_rt_grain=ChildStatus)
                    if ChildStatus == 'Current Count':
                        for data in dataList:
                            if data.ut_endoscopy_rt_grain == 'Under process' or data.ut_endoscopy_rt_grain == 'Observation(same stage)' or data.ut_endoscopy_rt_grain == 'Halt':
                                ListItems.append(data)
                elif ParentStatus == 'Material Qualified':
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(matiral_qualified_status=ChildStatus)
                    if ChildStatus == 'Current Count':
                        for data in dataList:
                            if data.matiral_qualified_status == 'Under process' or data.matiral_qualified_status == 'Observation(same stage)' or data.matiral_qualified_status == 'Halt':
                                ListItems.append(data)
                elif ParentStatus == 'Component Manufacturing':
                    if ChildStatus != 'Current Count':
                        # filter_objects &= get_filter('blt_status', 'equal', ChildStatus)
                        ListItems = dataList.filter(compo_manufacturing_status=ChildStatus)
                    if ChildStatus == 'Current Count':
                        for data in dataList:
                            if data.compo_manufacturing_status == 'Under process' or data.compo_manufacturing_status == 'Observation(same stage)' or data.compo_manufacturing_status == 'Halt':
                                ListItems.append(data)
                elif ParentStatus == 'Power Pack Assembly':
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(powerpack_assembly_status=ChildStatus)
                    if ChildStatus == 'Current Count':
                        for data in dataList:
                            if data.powerpack_assembly_status == 'Under process' or data.powerpack_assembly_status == 'Observation(same stage)' or data.powerpack_assembly_status == 'Halt':
                                ListItems.append(data)
                elif ParentStatus == 'Power Pack Testing':
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(powerpack_testing_status=ChildStatus)
                    if ChildStatus == 'Current Count':
                        for data in dataList:
                            if data.powerpack_testing_status == 'Under process' or data.powerpack_testing_status == 'Observation(same stage)' or data.powerpack_testing_status == 'Halt':
                                ListItems.append(data)
                elif ParentStatus == 'Raw Material Inspection':
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(raw_material_inspection_status=ChildStatus)
                    if ChildStatus == 'Current Count':
                        for data in dataList:
                            if data.raw_material_inspection_status == 'Under process' or data.raw_material_inspection_status == 'Observation(same stage)' or data.raw_material_inspection_status == 'Halt':
                                ListItems.append(data)
                elif ParentStatus == 'Pressing Electrode':
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(pressing_electrode_status=ChildStatus)
                    if ChildStatus == 'Current Count':
                        for data in dataList:
                            if data.pressing_electrode_status == 'Under process' or data.pressing_electrode_status == 'Observation(same stage)' or data.pressing_electrode_status == 'Halt':
                                ListItems.append(data)
                elif ParentStatus == 'Formation Process':
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(formation_process_status=ChildStatus)
                    if ChildStatus == 'Current Count':
                        for data in dataList:
                            if data.formation_process_status == 'Under process' or data.formation_process_status == 'Observation(same stage)' or data.formation_process_status == 'Halt':
                                ListItems.append(data)
                elif ParentStatus == 'Assembly Process':
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(assembly_process_status=ChildStatus)
                    if ChildStatus == 'Current Count':
                        for data in dataList:
                            if data.assembly_process_status == 'Under process' or data.assembly_process_status == 'Observation(same stage)' or data.assembly_process_status == 'Halt':
                                ListItems.append(data)
                elif ParentStatus == 'Battery Testing':
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(battery_testing_status=ChildStatus)
                    if ChildStatus == 'Current Count':
                        for data in dataList:
                            if data.battery_testing_status == 'Under process' or data.battery_testing_status == 'Observation(same stage)' or data.battery_testing_status == 'Halt':
                                ListItems.append(data)
                elif ParentStatus == 'Final Qualification':
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(final_qualification_status=ChildStatus)
                    if ChildStatus == 'Current Count':
                        for data in dataList:
                            if data.final_qualification_status == 'Under process' or data.final_qualification_status == 'Observation(same stage)' or data.final_qualification_status == 'Halt':
                                ListItems.append(data)

                elif ParentStatus == 'Qualification of Raw Material':
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(qualification_raw_material_status=ChildStatus)
                    if ChildStatus == 'Current Count':
                        for data in dataList:
                            if data.qualification_raw_material_status == 'Under process' or data.qualification_raw_material_status == 'Observation(same stage)' or data.qualification_raw_material_status == 'Halt':
                                ListItems.append(data)
                elif ParentStatus == 'Filling':
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(filling_status=ChildStatus)
                    if ChildStatus == 'Current Count':
                        for data in dataList:
                            if data.filling_status == 'Under process' or data.filling_status == 'Observation(same stage)' or data.filling_status == 'Halt':
                                ListItems.append(data)
                elif ParentStatus == 'Assembling/Ingegration':
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(assembling_integration_status=ChildStatus)
                    if ChildStatus == 'Current Count':
                        for data in dataList:
                            if data.assembling_integration_status == 'Under process' or data.assembling_integration_status == 'Observation(same stage)' or data.assembling_integration_status == 'Halt':
                                ListItems.append(data)
                elif ParentStatus == 'Qualification Testing':
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(qualification_testing_status=ChildStatus)
                    if ChildStatus == 'Current Count':
                        for data in dataList:
                            if data.qualification_testing_status == 'Under process' or data.qualification_testing_status == 'Observation(same stage)' or data.qualification_testing_status == 'Halt':
                                ListItems.append(data)
                elif ParentStatus == 'Performance Testing':
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(performance_testing_status=ChildStatus)
                    if ChildStatus == 'Current Count':
                        for data in dataList:
                            if data.performance_testing_status == 'Under process' or data.performance_testing_status == 'Observation(same stage)' or data.performance_testing_status == 'Halt':
                                ListItems.append(data)
                elif ParentStatus == 'NCR':
                        ListItems = dataList.filter(ncr_status=ChildStatus)

                elif ParentStatus == 'Overall':
                        ListItems = dataList.filter(overall_status=ChildStatus)

                serializer = ActiveMotorSerializer(ListItems, many=True)
                return JsonResponse({'message': 'Welcome to Home Page', 'data': serializer.data}, status=200)
            serializer = ActiveMotorSerializer(dataList, many=True)
            return JsonResponse({'status': 'True', 'data': serializer.data},
                                status=200)
        except Exception as e:
            print(e)
            return JsonResponse({'status': 'False', "message": "Internal Server Error"}, status=500)
    # @staticmethod
    # def GetActiveMotorListCount(request, self=None):
    #     try:
    #         current_year = request.query_params.get('selected_year')
    #         current_sys = request.query_params.get('selected_system')
    #         current_comp = request.query_params.get('selected_component')
    #         totCurrentYearSRM = 0
    #         totCurrentYearQualified = 0
    #         totCurrentYearObservation = 0
    #         totCurrentYearInprocess = 0
    #         totCurrentYearOverdue = 0
    #         if current_year !='':
    #             if current_sys =='' and current_comp == '':
    #                 totCurrentYearSRM = ActiveMotors.objects.filter(year__year = current_year).count()
    #                 totCurrentYearQualified = ActiveMotors.objects.filter(year__year = current_year,overall_status = "Qualified").count()
    #                 totCurrentYearInprocess = ActiveMotors.objects.filter(year__year = current_year,overall_status = "In Process").count()
    #                 totCurrentYearObservation = ActiveMotors.objects.filter(year__year = current_year,overall_status = "Observation").count()
    #                 totCurrentYearOverdue = ActiveMotors.objects.filter(year__year = current_year,overall_status = "Overdue").count()
    #
    #             elif current_sys =='' and current_comp !='':
    #                 totCurrentYearSRM = ActiveMotors.objects.filter(year__year=current_year, component_type = current_comp).count()
    #                 totCurrentYearQualified = ActiveMotors.objects.filter(year__year=current_year,
    #                                                                       overall_status="Qualified", component_type = current_comp).count()
    #                 totCurrentYearInprocess = ActiveMotors.objects.filter(year__year=current_year,
    #                                                                       overall_status="In Process", component_type = current_comp).count()
    #                 totCurrentYearObservation = ActiveMotors.objects.filter(year__year=current_year,
    #                                                                         overall_status="Observation", component_type = current_comp).count()
    #                 totCurrentYearOverdue = ActiveMotors.objects.filter(year__year=current_year,
    #                                                                     overall_status="Overdue", component_type = current_comp).count()
    #             elif current_sys !='' and current_comp =='':
    #                 totCurrentYearSRM = ActiveMotors.objects.filter(year__year=current_year, system = current_sys).count()
    #                 totCurrentYearQualified = ActiveMotors.objects.filter(year__year=current_year,
    #                                                                       overall_status="Qualified", system = current_sys).count()
    #                 totCurrentYearInprocess = ActiveMotors.objects.filter(year__year=current_year,
    #                                                                       overall_status="In Process", system = current_sys).count()
    #                 totCurrentYearObservation = ActiveMotors.objects.filter(year__year=current_year,
    #                                                                         overall_status="Observation", system = current_sys).count()
    #                 totCurrentYearOverdue = ActiveMotors.objects.filter(year__year=current_year,
    #                                                                     overall_status="Overdue", system = current_sys).count()
    #             else:
    #                 totCurrentYearSRM = ActiveMotors.objects.filter(year__year=current_year, system=current_sys).count()
    #                 totCurrentYearQualified = ActiveMotors.objects.filter(year__year=current_year,
    #                                                                       overall_status="Qualified",
    #                                                                       system=current_sys, component_type = current_comp).count()
    #                 totCurrentYearInprocess = ActiveMotors.objects.filter(year__year=current_year,
    #                                                                       overall_status="In Process",
    #                                                                       system=current_sys, component_type = current_comp).count()
    #                 totCurrentYearObservation = ActiveMotors.objects.filter(year__year=current_year,
    #                                                                         overall_status="Observation",
    #                                                                         system=current_sys, component_type = current_comp).count()
    #                 totCurrentYearOverdue = ActiveMotors.objects.filter(year__year=current_year,
    #                                                                     overall_status="Overdue", system=current_sys, component_type = current_comp).count()
    #         else:
    #             data = ActiveMotors.objects.all().count()
    #             return JsonResponse({'status': 'True', 'data': data},
    #                                 status=200)
    #         dist = {
    #
    #             'currYearTotalSrm': totCurrentYearSRM,
    #             'currYearSrmInprocess': totCurrentYearInprocess,
    #             'currYearSrmQualified': totCurrentYearQualified,
    #             'currYearSrmOverdue': totCurrentYearOverdue,
    #             'currYearSrmObservation': totCurrentYearObservation,
    #             'currYearSrmOverdue': totCurrentYearOverdue
    #
    #         }
    #         return JsonResponse({'status': 'True', 'data': dist},
    #                             status=200)
    #     except Exception as e:
    #         print(e)
    #         return JsonResponse({'status': 'False', "message": "Internal Server Error"}, status=500)
    @staticmethod
    def DeleteActiveMotor(request):
        # try:
            auditId = request.query_params['id']
            motor = ActiveMotors.objects.get(id=auditId)
            motor.delete()
            return JsonResponse({'message': 'Motor has been deleted'}, status=200)
        # except:
        #     return JsonResponse({'message': 'Sorry! No Motor found.'}, status=500)

    @staticmethod
    def GetActiveMotorHistoryList(request):
        try:
            id = request.query_params['id']
            data = ActiveMotorsHistory.objects.filter(active_motor_id=id)
            serializer = ActiveMotorSerializer(data, many=True)
            return JsonResponse({'message': 'Welcome to Home Page', 'data': serializer.data}, status=200)

        except:
            return JsonResponse({'message': 'Sorry! No Task found.'}, status=200)
    @staticmethod
    def GetActiveMotorPDFList(request):
        TABLE_COL_NAMES = ("Process/SRM",
                           "Acceptance of casting and raw materials", "Sandblasting",
                           " Insulation application and UT/RT status",
                           "Silver Acceptance and application", "Formulation tailoring",
                           "Conditioning of raw materials", "Lining ", "Casting UT, endoscopy and RT ",
                           "Liner Mechanical Properties",
                           "Propellant Mechanical Properties", "Interface bond strentgh", "Propellant burn rate",
                           "Mass of liner, Insulation, propellant and SRM", "Remarks")
        datum = list(ActiveMotors.objects.order_by('id').values_list('motor_id', 'acceptance_casting', 'sandblasting',
                                                                     'ut_rt_insulated_case',
                                                                     'acceptance_silver_material',
                                                                     'formulation_tailoring_liner_propellant',
                                                                     'conditioning_raw_materials', 'lining',
                                                                     'ut_endoscopy_rt_grain',
                                                                     'liner_mechanical_properties',
                                                                     'propellant_mechanical_properties',
                                                                     'interface_bond_strength', 'propellant_burn_rate',
                                                                     'mass_liner_insulation_propellant_srm',
                                                                     'overall_remarks'))

        table = list(zip(TABLE_COL_NAMES, *datum))
        pdf = FPDF("L", "mm", "Legal")
        # font = TTFont('venv/lib/python3.8/site-packages/fontTools/ttLib/ttFont.py')
        # pdf.add_font('Calibri', '', 'venv/lib/python3.8/site-packages/fontTools/ttLib/ttFont.py', uni=True)
        def page_title():
            pdf.set_font('courier', 'B', 26)
            today = date.today()
            pdf.cell(330, 10, 'Weekly Status of SRMs at CPS (NDC) dt ' + f"{today}", border=0, align='C', ln=2)
            # title = ActiveMotors.objects.values_list('system',flat=True)
            # pdf.cell(40, 10, 'A. ' + f"{title}", 0, 1)
            pdf.set_font('Times', 'BU', 15)
            pdf.cell(40, 10, 'A. SRMs for Ballistic System (SWS)', 0, 1)
            pdf.set_font("Times", size=11)

        def footer():
            pdf.set_y(-15)
            pdf.set_font('Times', 'B', 10)
            pdf.cell(330, 10, 'Page ' + str(pdf.page_no()) + ' of ' + '{nb}', 0, 0, 'R')

        c_h = pdf.font_size * 2.7
        c_w = pdf.epw / 5
        lh_list = []
        use_default_height = 0
        for row in table:
            for data_items in row:
                dd = str(data_items)
                word_list = dd.split()
                number_of_words = len(word_list)
                if number_of_words > 2:
                    use_default_height = 2
                    new_line_height = pdf.font_size * (number_of_words / 2)
            if not use_default_height:
                lh_list.append(c_h)
            else:
                lh_list.append(new_line_height)
                use_default_height = 0
        rows_per_page = 15
        cols_per_page = 5
        num_rows = len(table)
        num_cols = len(table[0])
        row_offset = 0
        while row_offset < num_rows:
            row_max = row_offset + rows_per_page
            if row_max > num_rows:
                row_max = num_rows
            col_offset = 0
            while col_offset < num_cols:
                col_max = col_offset + cols_per_page
                if col_max > num_cols:
                    col_max = num_cols
                pdf.add_page()
                pdf.t_margin = 25
                if col_max == 5:
                    page_title()
                    pdf.cell(40, 3, '', 0, 1)
                for i in range(row_offset, row_max):
                    c_h = lh_list[i]
                    for j in range(col_offset, col_max):
                        if i == 0 and j <= num_rows:
                            pdf.set_font('Times','')
                            pdf.multi_cell(c_w, c_h, table[i][j], border=1, align='C', ln=3,
                                           max_line_height=pdf.font_size)
                        elif i <= 14 and j == 0:
                            pdf.set_font('Times','B')
                            pdf.multi_cell(c_w, c_h, table[i][j], border=1, align='L', ln=3,
                                           max_line_height=pdf.font_size)

                        else:
                            cell_value = table[i][j]
                            pdf.set_font('Times','')
                            pdf.multi_cell(c_w, c_h, cell_value, border=1, align='L', ln=3,
                                           max_line_height=pdf.font_size)

                    pdf.ln(c_h)
                    pdf.set_auto_page_break(True, margin=1.5)

                col_offset += cols_per_page
                footer()

            row_offset += rows_per_page
        pdf.output("weekly-report.pdf", dest="F")
        return FileResponse(open('weekly-report.pdf', 'rb'), as_attachment=True, content_type='application/pdf')

    @staticmethod
    def AddLotIds(request):
        try:
            modal = organization_lot_ids()
            modal.organization = request['organization']
            modal.lot_id_number = request['lot_id_number']
            modal.save()
            return JsonResponse({'status': 'true', "message": "Record Added"}, status=200)
        except Exception as e:
            return JsonResponse({'status': 'False', "message": "Data Not Saved"}, status=500)

    @staticmethod
    def GetLots(request, self=None):
        try:
            data = organization_lot_ids.objects.all()
            serializers = LotIdsSerialzer(data,many=True)
            return JsonResponse({'status': 'True', 'data': serializers.data},
                                status=200)
        except Exception as e:
            print(e)
            return JsonResponse({'status': 'False', "message": "Internal Server Error"}, status=500)
    # @staticmethod
    # def GetActiveMotorPDFList(request):
    #     TABLE_COL_NAMES = (
    #         "Acceptance of casting and raw materials", "Sandblasting", " Insulation application and UT/RT status",
    #         "Silver Acceptance and application", "Formulation tailoring",
    #         "Conditioning of raw materials", "Lining ", "Casting UT, endoscopy and RT ", "Liner Mechanical Properties",
    #         "Propellant Mechanical Properties", "Interface bond strentgh", "Propellant burn rate",
    #         "Mass of liner, Insulation, propellant and SRM", "Remarks")
    #
    #     datum = list(ActiveMotors.objects.values_list('acceptance_casting', 'sandblasting', 'ut_rt_insulated_case',
    #                                                   'acceptance_silver_material',
    #                                                   'formulation_tailoring_liner_propellant',
    #                                                   'conditioning_raw_materials', 'lining',
    #                                                   'ut_endoscopy_rt_grain', 'liner_mechanical_properties',
    #                                                   'propellant_mechanical_properties',
    #                                                   'interface_bond_strength', 'propellant_burn_rate',
    #                                                   'mass_liner_insulation_propellant_srm', 'overall_remarks'))
    #     data = list(zip(TABLE_COL_NAMES, *datum))
    #     pdf = FPDF('L', 'mm', 'Legal')
    #     pdf.add_page()
    #     pdf.set_font('courier', 'B', 26)
    #     pdf.cell(330, 10, 'Active Motor Final Report', border=0, align='C', ln=2)
    #     pdf.cell(40, 10, '', 0, 1)
    #     pdf.set_font("arial", size=11)
    #     line_height = pdf.font_size * 4
    #     col_width = pdf.epw / 5
    #
    #     lh_list = []
    #     use_default_height = 0
    #     for row in data:
    #         for datum in row:
    #             dd = str(datum)
    #             word_list = dd.split()
    #             number_of_words = len(word_list)
    #             if number_of_words > 2:
    #                 use_default_height = 2
    #                 new_line_height = pdf.font_size * (number_of_words / 2)
    #         if not use_default_height:
    #             lh_list.append(line_height)
    #         else:
    #             lh_list.append(new_line_height)
    #             use_default_height = 0
    #
    #     for j, row in enumerate(data):
    #         line_height = lh_list[j]
    #         for col_num in range(len(row)):
    #             if pdf.get_x() < 300:
    #                 pdf.multi_cell(col_width, line_height, f"{row[col_num]}", border=1, align='L', ln=3,
    #                                max_line_height=pdf.font_size)
    #
    #                 # pdf.set_auto_page_break(True,margin=0)
    #             # if pdf.get_x() >299:
    #             #     pdf.multi_cell(col_width, line_height, f"{row[col_num]}", border=1, align='L', ln=3,
    #             #                max_line_height=pdf.font_size)
    #         pdf.set_auto_page_break(True, margin=0)
    #         pdf.ln(line_height)
    #     pdf.output('report.pdf')
    #     return FileResponse(open('report.pdf', 'rb'), as_attachment=True, content_type='application/pdf')

    @staticmethod
    def GetActiveMotorExcelList(request):
        response = HttpResponse(content_type='application/ms-excel')
        response['Content-Disposition'] = 'attachment; filename="DocumentExcelSheet.xls"'
        work_book = xlwt.Workbook(encoding='utf-8')
        work_sheet = work_book.add_sheet('ExcelSheet')
        row_num = 0
        font_style = xlwt.XFStyle()
        font_style.font.bold = True
        TABLE_COL_NAMES = ("System", "Motor ID", "Qualification of raw materials of insulation, lining and propellant",
                           "Qualification of raw materials of insulation, lining and propellant remarks",
                           "Acceptance of casting", "Acceptance of casting remarks", "Sandblasting",
                           "Sandblasting remarks", "Insulation", "Insulation remarks", "UT and RT of Insulated Case",
                           "UT and RT of Insulated Case remarks",
                           "Acceptance of Silver Material", "Acceptance of Silver Material remarks",
                           "Silver Application", "Silver Application remarks",
                           "Formulation tailoring of liner and propellant",
                           "Formulation tailoring of liner and propellant remarks",
                           "Conditioning of raw materials", "Conditioning of raw materials remarks", "Lining",
                           "Lining remarks", "Casting", "Casting remarks", "Curing", "Curing remarks",
                           "Liner Mechanical Properties", "Liner Mechanical Properties remarks",
                           "Propellant Mechanical Properties", "Propellant Mechanical Properties remarks",
                           "Interface bond strentgh", "Interface bond strentgh remarks", "Propellant burn rate",
                           "Propellant burn rate remarks", "Trimming of propellant grain",
                           "Trimming of propellant grain remarks", "Trimming of propellant grain remarks",
                           "Mass of liner, Insulation, propellant and SRM",
                           "Mass of liner, Insulation, propellant and SRM remarks", "UT, endoscopy and RT of grain",
                           "UT, endoscopy and RT of grain remarks", "NCR Status", "NCR Status remarks",
                           "Overall Status", "Overall remarks")
        for col_num in range(len(TABLE_COL_NAMES)):
            work_sheet.write(row_num, col_num, TABLE_COL_NAMES[col_num], font_style)
        font_style = xlwt.XFStyle()

        data = ActiveMotors.objects.all().order_by('-id').values_list('system', 'motor_id',
                                                                      'qualification_insulation_lining_propellant_rm',
                                                                      'qualification_insulation_lining_propellant_rm_remarks',
                                                                      'acceptance_casting',
                                                                      'acceptance_casting_remarks', 'sandblasting',
                                                                      'sandblasting_remarks', 'insulation',
                                                                      'insulation_remarks', 'ut_rt_insulated_case',
                                                                      'ut_rt_insulated_case_remarks',
                                                                      'acceptance_silver_material',
                                                                      'acceptance_silver_material_remarks',
                                                                      'silver_application',
                                                                      'silver_application_remarks',
                                                                      'formulation_tailoring_liner_propellant',
                                                                      'formulation_tailoring_liner_propellant_remarks',
                                                                      'conditioning_raw_materials',
                                                                      'conditioning_raw_materials_remarks', 'lining',
                                                                      'lining_remarks', 'casting', 'casting_remarks',
                                                                      'curing', 'curing_remarks',
                                                                      'liner_mechanical_properties',
                                                                      'liner_mechanical_properties_remarks',
                                                                      'propellant_mechanical_properties',
                                                                      'propellant_mechanical_properties_remarks',
                                                                      'interface_bond_strength',
                                                                      'interface_bond_strength_remarks',
                                                                      'propellant_burn_rate',
                                                                      'propellant_burn_rate_remarks',
                                                                      'trimming_Propellant_grain',
                                                                      'trimming_Propellant_grain_remarks',
                                                                      'mass_liner_insulation_propellant_srm',
                                                                      'mass_liner_insulation_propellant_srm_remarks',
                                                                      'ut_endoscopy_rt_grain',
                                                                      'ut_endoscopy_rt_grain_remarks', 'ncr_status',
                                                                      'ncr_status_remarks', 'overall_status',
                                                                      'overall_remarks')

        for row in data:
            row_num += 1
            for col_num in range(len(row)):
                work_sheet.write(row_num, col_num, str(row[col_num]), font_style)
        work_book.save(response)
        return response


    @staticmethod
    def getProcessMonitoringDashboardCount(request):
        # try:

            def get_filter(field_name, filter_condition, filter_value):
                # thanks to the below post
                # https://stackoverflow.com/questions/310732/in-django-how-does-one-filter-a-queryset-with-dynamic-field-lookups
                # the idea to this below logic is very similar to that in the above mentioned post
                if filter_condition.strip() == "contains":
                    kwargs = {
                        '{0}__icontains'.format(field_name): filter_value
                    }
                    return Q(**kwargs)

                if filter_condition.strip() == "not_equal":
                    kwargs = {
                        '{0}__iexact'.format(field_name): filter_value
                    }
                    return ~Q(**kwargs)

                if filter_condition.strip() == "starts_with":
                    kwargs = {
                        '{0}__istartswith'.format(field_name): filter_value
                    }
                    return Q(**kwargs)
                if filter_condition.strip() == "equal":
                    kwargs = {
                        '{0}__iexact'.format(field_name): filter_value
                    }
                    return Q(**kwargs)

                if filter_condition.strip() == "not_equal":
                    kwargs = {
                        '{0}__iexact'.format(field_name): filter_value
                    }

                    return ~Q(**kwargs)

            filter_objects = Q()

            data = []
            prod_blt_ok = 0
            prod_blt_observation = 0
            prod_blt_up = 0
            prod_blt_hault = 0
            year = request.query_params['selected_year']
            component = request.query_params['selected_component']
            current_org = request.query_params.get('selected_organization')
            current_bat = request.query_params.get('selected_battery')
            current_lot = request.query_params.get('selected_lot_id')
            system_type = request.query_params.get('selected_system_type')
            system = request.query_params['selected_system']

            # create dynamic filter
            if year != '':
                filter_objects &= get_filter(
                    'testing_date__year', 'equal',
                    year)

            if component != '':
                filter_objects &= get_filter(
                    'component_type', 'equal',
                    component)
            if system_type != '':
                filter_objects &= get_filter(
                    'system_type', 'equal',
                    system_type)
            if system != '':
                filter_objects &= get_filter(
                    'system_name', 'equal',
                    system)
            if current_org !='':
                filter_objects &= get_filter(
                    'organization', 'equal', current_org)
            if current_bat !='':
                filter_objects &= get_filter(
                    'battery_type', 'equal', current_bat)
            if current_lot !='':
                filter_objects &= get_filter(
                    'lot_id', 'equal', current_lot)
            # production system count
            qualification_insulation_lining_propellant_rm_OK = ActiveMotors.objects.filter(filter_objects, qualification_insulation_lining_propellant_rm='Ok')
            qualification_insulation_lining_propellant_rm_ObsSame = ActiveMotors.objects.filter(filter_objects,
                                                                             qualification_insulation_lining_propellant_rm='Observation(same stage)')
            qualification_insulation_lining_propellant_rm_ObsNext = ActiveMotors.objects.filter(filter_objects,
                                                                                    qualification_insulation_lining_propellant_rm='Observation(next stage)')
            qualification_insulation_lining_propellant_rm_Up = ActiveMotors.objects.filter(filter_objects, qualification_insulation_lining_propellant_rm='Under process')
            qualification_insulation_lining_propellant_rm_Hlt = ActiveMotors.objects.filter(filter_objects, qualification_insulation_lining_propellant_rm='Halt')

            acceptance_casting_Ok = ActiveMotors.objects.filter(filter_objects, acceptance_casting='Ok')
            acceptance_casting_ObsSame = ActiveMotors.objects.filter(filter_objects,
                                                                                acceptance_casting='Observation(same stage)')
            acceptance_casting_ObsNext = ActiveMotors.objects.filter(filter_objects,
                                                                                    acceptance_casting='Observation(next stage)')
            acceptance_casting_Up = ActiveMotors.objects.filter(filter_objects, acceptance_casting='Under process')
            acceptance_casting_Hlt = ActiveMotors.objects.filter(filter_objects, acceptance_casting='Halt')



            sandblasting_Ok = ActiveMotors.objects.filter(filter_objects, sandblasting='Ok')
            sandblasting_ObsSame = ActiveMotors.objects.filter(filter_objects,sandblasting='Observation(same stage)')
            sandblasting_ObsNext = ActiveMotors.objects.filter(filter_objects,sandblasting='Observation(next stage)')
            sandblasting_Up = ActiveMotors.objects.filter(filter_objects, sandblasting='Under process')
            sandblasting_Hlt = ActiveMotors.objects.filter(filter_objects, sandblasting='Halt')



            insulation_Ok = ActiveMotors.objects.filter(filter_objects, insulation='Ok')
            insulation_ObsSame = ActiveMotors.objects.filter(filter_objects,
                                                                                insulation='Observation(same stage)')
            insulation_ObsNext = ActiveMotors.objects.filter(filter_objects,
                                                                                    insulation='Observation(next stage)')
            insulation_Up = ActiveMotors.objects.filter(filter_objects, insulation='Under process')
            insulation_Hlt = ActiveMotors.objects.filter(filter_objects, insulation='Halt')


            ut_rt_insulated_case_Ok = ActiveMotors.objects.filter(filter_objects, ut_rt_insulated_case='Ok')
            ut_rt_insulated_case_ObsSame = ActiveMotors.objects.filter(filter_objects,
                                                                                ut_rt_insulated_case='Observation(same stage)')
            ut_rt_insulated_case_ObsNext = ActiveMotors.objects.filter(filter_objects,
                                                                                    ut_rt_insulated_case='Observation(next stage)')
            ut_rt_insulated_case_Up = ActiveMotors.objects.filter(filter_objects, ut_rt_insulated_case='Under process')
            ut_rt_insulated_case_Hlt = ActiveMotors.objects.filter(filter_objects, ut_rt_insulated_case='Halt')


            acceptance_silver_material_Ok = ActiveMotors.objects.filter(filter_objects, acceptance_silver_material='Ok')
            acceptance_silver_material_ObsSame = ActiveMotors.objects.filter(filter_objects,
                                                                                acceptance_silver_material='Observation(same stage)')
            acceptance_silver_material_ObsNext = ActiveMotors.objects.filter(filter_objects,
                                                                                    acceptance_silver_material='Observation(next stage)')
            acceptance_silver_material_Up = ActiveMotors.objects.filter(filter_objects, acceptance_silver_material='Under process')
            acceptance_silver_material_Hlt = ActiveMotors.objects.filter(filter_objects, acceptance_silver_material='Halt')


            silver_application_Ok = ActiveMotors.objects.filter(filter_objects, silver_application='Ok')
            silver_application_ObsSame = ActiveMotors.objects.filter(filter_objects,
                                                                                silver_application='Observation(same stage)')
            silver_application_ObsNext = ActiveMotors.objects.filter(filter_objects,
                                                                                    silver_application='Observation(next stage)')
            silver_application_Up = ActiveMotors.objects.filter(filter_objects, silver_application='Under process')
            silver_application_Hlt = ActiveMotors.objects.filter(filter_objects, silver_application='Halt')


            formulation_tailoring_liner_propellant_Ok = ActiveMotors.objects.filter(filter_objects, formulation_tailoring_liner_propellant='Ok')
            formulation_tailoring_liner_propellant_ObsSame = ActiveMotors.objects.filter(filter_objects,
                                                                                formulation_tailoring_liner_propellant='Observation(same stage)')
            formulation_tailoring_liner_propellant_ObsNext = ActiveMotors.objects.filter(filter_objects,
                                                                                    formulation_tailoring_liner_propellant='Observation(next stage)')
            formulation_tailoring_liner_propellant_Up = ActiveMotors.objects.filter(filter_objects, formulation_tailoring_liner_propellant='Under process')
            formulation_tailoring_liner_propellant_Hlt = ActiveMotors.objects.filter(filter_objects, formulation_tailoring_liner_propellant='Halt')


            conditioning_raw_materials_Ok = ActiveMotors.objects.filter(filter_objects, conditioning_raw_materials='Ok')
            conditioning_raw_materials_ObsSame = ActiveMotors.objects.filter(filter_objects,
                                                                                conditioning_raw_materials='Observation(same stage)')
            conditioning_raw_materials_ObsNext = ActiveMotors.objects.filter(filter_objects,
                                                                                    conditioning_raw_materials='Observation(next stage)')
            conditioning_raw_materials_Up = ActiveMotors.objects.filter(filter_objects, conditioning_raw_materials='Under process')
            conditioning_raw_materials_Hlt = ActiveMotors.objects.filter(filter_objects, conditioning_raw_materials='Halt')


            lining_Ok = ActiveMotors.objects.filter(filter_objects, lining='Ok')
            lining_ObsSame = ActiveMotors.objects.filter(filter_objects,
                                                                                lining='Observation(same stage)')
            lining_ObsNext = ActiveMotors.objects.filter(filter_objects,
                                                                                    lining='Observation(next stage)')
            lining_Up = ActiveMotors.objects.filter(filter_objects, lining='Under process')
            lining_Hlt = ActiveMotors.objects.filter(filter_objects, lining='Halt')


            casting_Ok = ActiveMotors.objects.filter(filter_objects, casting='Ok')
            casting_ObsSame = ActiveMotors.objects.filter(filter_objects,
                                                                                casting='Observation(same stage)')
            casting_ObsNext = ActiveMotors.objects.filter(filter_objects,
                                                                                    casting='Observation(next stage)')
            casting_Up = ActiveMotors.objects.filter(filter_objects, casting='Under process')
            casting_Hlt = ActiveMotors.objects.filter(filter_objects, casting='Halt')


            curing_Ok = ActiveMotors.objects.filter(filter_objects, curing='Ok')
            curing_ObsSame = ActiveMotors.objects.filter(filter_objects,
                                                                                curing='Observation(same stage)')
            curing_ObsNext = ActiveMotors.objects.filter(filter_objects,
                                                                                    curing='Observation(next stage)')
            curing_Up = ActiveMotors.objects.filter(filter_objects, curing='Under process')
            curing_Hlt = ActiveMotors.objects.filter(filter_objects, curing='Halt')


            liner_mechanical_properties_Ok = ActiveMotors.objects.filter(filter_objects, liner_mechanical_properties='Ok')
            liner_mechanical_properties_ObsSame = ActiveMotors.objects.filter(filter_objects,
                                                                                liner_mechanical_properties='Observation(same stage)')
            liner_mechanical_properties_ObsNext = ActiveMotors.objects.filter(filter_objects,
                                                                                    liner_mechanical_properties='Observation(next stage)')
            liner_mechanical_properties_Up = ActiveMotors.objects.filter(filter_objects, liner_mechanical_properties='Under process')
            liner_mechanical_properties_Hlt = ActiveMotors.objects.filter(filter_objects, liner_mechanical_properties='Halt')


            propellant_mechanical_properties_Ok = ActiveMotors.objects.filter(filter_objects, propellant_mechanical_properties='Ok')
            propellant_mechanical_properties_ObsSame = ActiveMotors.objects.filter(filter_objects,
                                                                                propellant_mechanical_properties='Observation(same stage)')
            propellant_mechanical_properties_ObsNext = ActiveMotors.objects.filter(filter_objects,
                                                                                    propellant_mechanical_properties='Observation(next stage)')
            propellant_mechanical_properties_Up = ActiveMotors.objects.filter(filter_objects, propellant_mechanical_properties='Under process')
            propellant_mechanical_properties_Hlt = ActiveMotors.objects.filter(filter_objects, propellant_mechanical_properties='Halt')


            interface_bond_strength_Ok = ActiveMotors.objects.filter(filter_objects, interface_bond_strength='Ok')
            interface_bond_strength_ObsSame = ActiveMotors.objects.filter(filter_objects,
                                                                                interface_bond_strength='Observation(same stage)')
            interface_bond_strength_ObsNext = ActiveMotors.objects.filter(filter_objects,
                                                                                    interface_bond_strength='Observation(next stage)')
            interface_bond_strength_Up = ActiveMotors.objects.filter(filter_objects, interface_bond_strength='Under process')
            interface_bond_strength_Hlt = ActiveMotors.objects.filter(filter_objects, interface_bond_strength='Halt')


            propellant_burn_rate_Ok = ActiveMotors.objects.filter(filter_objects, propellant_burn_rate='Ok')
            propellant_burn_rate_ObsSame = ActiveMotors.objects.filter(filter_objects,
                                                                                propellant_burn_rate='Observation(same stage)')
            propellant_burn_rate_ObsNext = ActiveMotors.objects.filter(filter_objects,
                                                                                    propellant_burn_rate='Observation(next stage)')
            propellant_burn_rate_Up = ActiveMotors.objects.filter(filter_objects, propellant_burn_rate='Under process')
            propellant_burn_rate_Hlt = ActiveMotors.objects.filter(filter_objects, propellant_burn_rate='Halt')


            trimming_Propellant_grain_Ok = ActiveMotors.objects.filter(filter_objects, trimming_Propellant_grain='Ok')
            trimming_Propellant_grain_ObsSame = ActiveMotors.objects.filter(filter_objects,
                                                                                trimming_Propellant_grain='Observation(same stage)')
            trimming_Propellant_grain_ObsNext = ActiveMotors.objects.filter(filter_objects,
                                                                                    trimming_Propellant_grain='Observation(next stage)')
            trimming_Propellant_grain_Up = ActiveMotors.objects.filter(filter_objects, trimming_Propellant_grain='Under process')
            trimming_Propellant_grain_Hlt = ActiveMotors.objects.filter(filter_objects, trimming_Propellant_grain='Halt')


            mass_liner_insulation_propellant_srm_Ok = ActiveMotors.objects.filter(filter_objects, mass_liner_insulation_propellant_srm='Ok')
            mass_liner_insulation_propellant_srm_ObsSame = ActiveMotors.objects.filter(filter_objects,
                                                                                mass_liner_insulation_propellant_srm='Observation(same stage)')
            mass_liner_insulation_propellant_srm_ObsNext = ActiveMotors.objects.filter(filter_objects,
                                                                                    mass_liner_insulation_propellant_srm='Observation(next stage)')
            mass_liner_insulation_propellant_srm_Up = ActiveMotors.objects.filter(filter_objects, mass_liner_insulation_propellant_srm='Under process')
            mass_liner_insulation_propellant_srm_Hlt = ActiveMotors.objects.filter(filter_objects, mass_liner_insulation_propellant_srm='Halt')


            ut_endoscopy_rt_grain_Ok = ActiveMotors.objects.filter(filter_objects, ut_endoscopy_rt_grain='Ok')
            ut_endoscopy_rt_grain_ObsSame = ActiveMotors.objects.filter(filter_objects,
                                                                                ut_endoscopy_rt_grain='Observation(same stage)')
            ut_endoscopy_rt_grain_ObsNext = ActiveMotors.objects.filter(filter_objects,
                                                                                    ut_endoscopy_rt_grain='Observation(next stage)')
            ut_endoscopy_rt_grain_Up = ActiveMotors.objects.filter(filter_objects, ut_endoscopy_rt_grain='Under process')
            ut_endoscopy_rt_grain_Hlt = ActiveMotors.objects.filter(filter_objects, ut_endoscopy_rt_grain='Halt')


            matiral_qualified_status_Ok = ActiveMotors.objects.filter(filter_objects, matiral_qualified_status='Ok')
            matiral_qualified_status_ObsSame = ActiveMotors.objects.filter(filter_objects,
                                                                                matiral_qualified_status='Observation(same stage)')
            matiral_qualified_status_ObsNext = ActiveMotors.objects.filter(filter_objects,
                                                                                    matiral_qualified_status='Observation(next stage)')
            matiral_qualified_status_Up = ActiveMotors.objects.filter(filter_objects, matiral_qualified_status='Under process')
            matiral_qualified_status_Hlt = ActiveMotors.objects.filter(filter_objects, matiral_qualified_status='Halt')


            compo_manufacturing_status_Ok = ActiveMotors.objects.filter(filter_objects, compo_manufacturing_status='Ok')
            compo_manufacturing_status_ObsSame = ActiveMotors.objects.filter(filter_objects,
                                                                                compo_manufacturing_status='Observation(same stage)')
            compo_manufacturing_status_ObsNext = ActiveMotors.objects.filter(filter_objects,
                                                                                    compo_manufacturing_status='Observation(next stage)')
            compo_manufacturing_status_Up = ActiveMotors.objects.filter(filter_objects, compo_manufacturing_status='Under process')
            compo_manufacturing_status_Hlt = ActiveMotors.objects.filter(filter_objects, compo_manufacturing_status='Halt')


            powerpack_assembly_status_Ok = ActiveMotors.objects.filter(filter_objects, powerpack_assembly_status='Ok')
            powerpack_assembly_status_ObsSame = ActiveMotors.objects.filter(filter_objects,
                                                                                powerpack_assembly_status='Observation(same stage)')
            powerpack_assembly_status_ObsNext = ActiveMotors.objects.filter(filter_objects,
                                                                                    powerpack_assembly_status='Observation(next stage)')
            powerpack_assembly_status_Up = ActiveMotors.objects.filter(filter_objects, powerpack_assembly_status='Under process')
            powerpack_assembly_status_Hlt = ActiveMotors.objects.filter(filter_objects, powerpack_assembly_status='Halt')


            powerpack_testing_status_Ok = ActiveMotors.objects.filter(filter_objects, powerpack_testing_status='Ok')
            powerpack_testing_status_ObsSame = ActiveMotors.objects.filter(filter_objects,
                                                                                powerpack_testing_status='Observation(same stage)')
            powerpack_testing_status_ObsNext = ActiveMotors.objects.filter(filter_objects,
                                                                                    powerpack_testing_status='Observation(next stage)')
            powerpack_testing_status_Up = ActiveMotors.objects.filter(filter_objects, powerpack_testing_status='Under process')
            powerpack_testing_status_Hlt = ActiveMotors.objects.filter(filter_objects, powerpack_testing_status='Halt')



            raw_material_inspection_status_Ok = ActiveMotors.objects.filter(filter_objects, raw_material_inspection_status='Ok')
            raw_material_inspection_status_ObsSame = ActiveMotors.objects.filter(filter_objects,
                                                                                raw_material_inspection_status='Observation(same stage)')
            raw_material_inspection_status_ObsNext = ActiveMotors.objects.filter(filter_objects,
                                                                                    raw_material_inspection_status='Observation(next stage)')
            raw_material_inspection_status_Up = ActiveMotors.objects.filter(filter_objects, raw_material_inspection_status='Under process')
            raw_material_inspection_status_Hlt = ActiveMotors.objects.filter(filter_objects, raw_material_inspection_status='Halt')



            pressing_electrode_status_Ok = ActiveMotors.objects.filter(filter_objects, pressing_electrode_status='Ok')
            pressing_electrode_status_ObsSame = ActiveMotors.objects.filter(filter_objects,
                                                                                pressing_electrode_status='Observation(same stage)')
            pressing_electrode_status_ObsNext = ActiveMotors.objects.filter(filter_objects,
                                                                                    pressing_electrode_status='Observation(next stage)')
            pressing_electrode_status_Up = ActiveMotors.objects.filter(filter_objects, pressing_electrode_status='Under process')
            pressing_electrode_status_Hlt = ActiveMotors.objects.filter(filter_objects, pressing_electrode_status='Halt')



            formation_process_status_Ok = ActiveMotors.objects.filter(filter_objects, formation_process_status='Ok')
            formation_process_status_ObsSame = ActiveMotors.objects.filter(filter_objects,
                                                                                formation_process_status='Observation(same stage)')
            formation_process_status_ObsNext = ActiveMotors.objects.filter(filter_objects,
                                                                                    formation_process_status='Observation(next stage)')
            formation_process_status_Up = ActiveMotors.objects.filter(filter_objects, formation_process_status='Under process')
            formation_process_status_Hlt = ActiveMotors.objects.filter(filter_objects, formation_process_status='Halt')



            assembly_process_status_Ok = ActiveMotors.objects.filter(filter_objects, assembly_process_status='Ok')
            assembly_process_status_ObsSame = ActiveMotors.objects.filter(filter_objects,
                                                                                assembly_process_status='Observation(same stage)')
            assembly_process_status_ObsNext = ActiveMotors.objects.filter(filter_objects,
                                                                                    assembly_process_status='Observation(next stage)')
            assembly_process_status_Up = ActiveMotors.objects.filter(filter_objects, assembly_process_status='Under process')
            assembly_process_status_Hlt = ActiveMotors.objects.filter(filter_objects, assembly_process_status='Halt')


            battery_testing_status_Ok = ActiveMotors.objects.filter(filter_objects, battery_testing_status='Ok')
            battery_testing_status_ObsSame = ActiveMotors.objects.filter(filter_objects,
                                                                                battery_testing_status='Observation(same stage)')
            battery_testing_status_ObsNext = ActiveMotors.objects.filter(filter_objects,
                                                                                    battery_testing_status='Observation(next stage)')
            battery_testing_status_Up = ActiveMotors.objects.filter(filter_objects, battery_testing_status='Under process')
            battery_testing_status_Hlt = ActiveMotors.objects.filter(filter_objects, battery_testing_status='Halt')



            final_qualification_status_Ok = ActiveMotors.objects.filter(filter_objects, final_qualification_status='Ok')
            final_qualification_status_ObsSame = ActiveMotors.objects.filter(filter_objects,
                                                                                final_qualification_status='Observation(same stage)')
            final_qualification_status_ObsNext = ActiveMotors.objects.filter(filter_objects,
                                                                                    final_qualification_status='Observation(next stage)')
            final_qualification_status_Up = ActiveMotors.objects.filter(filter_objects, final_qualification_status='Under process')
            final_qualification_status_Hlt = ActiveMotors.objects.filter(filter_objects, final_qualification_status='Halt')

# ==========================Pyro Devices ==========================
            qualification_raw_material_status_Ok = ActiveMotors.objects.filter(filter_objects, qualification_raw_material_status='Ok')
            qualification_raw_material_status_ObsSame = ActiveMotors.objects.filter(filter_objects,
                                                                            qualification_raw_material_status='Observation(same stage)')
            qualification_raw_material_status_ObsNext = ActiveMotors.objects.filter(filter_objects,
                                                                            qualification_raw_material_status='Observation(next stage)')
            qualification_raw_material_status_Up = ActiveMotors.objects.filter(filter_objects,
                                                                       qualification_raw_material_status='Under process')
            qualification_raw_material_status_Hlt = ActiveMotors.objects.filter(filter_objects, qualification_raw_material_status='Halt')


            filling_status_Ok = ActiveMotors.objects.filter(filter_objects, filling_status='Ok')
            filling_status_ObsSame = ActiveMotors.objects.filter(filter_objects,
                                                                           filling_status='Observation(same stage)')
            filling_status_ObsNext = ActiveMotors.objects.filter(filter_objects,
                                                                           filling_status='Observation(next stage)')
            filling_status_Up = ActiveMotors.objects.filter(filter_objects, filling_status='Under process')
            filling_status_Hlt = ActiveMotors.objects.filter(filter_objects, filling_status='Halt')


            assembling_integration_status_Ok = ActiveMotors.objects.filter(filter_objects, assembling_integration_status='Ok')
            assembling_integration_status_ObsSame = ActiveMotors.objects.filter(filter_objects,
                                                                          assembling_integration_status='Observation(same stage)')
            assembling_integration_status_ObsNext = ActiveMotors.objects.filter(filter_objects,
                                                                          assembling_integration_status='Observation(next stage)')
            assembling_integration_status_Up = ActiveMotors.objects.filter(filter_objects, assembling_integration_status='Under process')
            assembling_integration_status_Hlt = ActiveMotors.objects.filter(filter_objects, assembling_integration_status='Halt')


            qualification_testing_status_Ok = ActiveMotors.objects.filter(filter_objects, qualification_testing_status='Ok')
            qualification_testing_status_ObsSame = ActiveMotors.objects.filter(filter_objects,
                                                                         qualification_testing_status='Observation(same stage)')
            qualification_testing_status_ObsNext = ActiveMotors.objects.filter(filter_objects,
                                                                         qualification_testing_status='Observation(next stage)')
            qualification_testing_status_Up = ActiveMotors.objects.filter(filter_objects, qualification_testing_status='Under process')
            qualification_testing_status_Hlt = ActiveMotors.objects.filter(filter_objects, qualification_testing_status='Halt')


            performance_testing_status_Ok = ActiveMotors.objects.filter(filter_objects, performance_testing_status='Ok')
            performance_testing_status_ObsSame = ActiveMotors.objects.filter(filter_objects,
                                                                             performance_testing_status='Observation(same stage)')
            performance_testing_status_ObsNext = ActiveMotors.objects.filter(filter_objects,
                                                                             performance_testing_status='Observation(next stage)')
            performance_testing_status_Up = ActiveMotors.objects.filter(filter_objects,
                                                                        performance_testing_status='Under process')
            performance_testing_status_Hlt = ActiveMotors.objects.filter(filter_objects, performance_testing_status='Halt')


            # ncr_ok = ActiveMotors.objects.filter(filter_objects,ncr_status='Ok')
            # ncr_obs = ActiveMotors.objects.filter(filter_objects,ncr_status='Observation')
            # ncr_inprocess = ActiveMotors.objects.filter(filter_objects,ncr_status='In-process')
            #
            # overall_ok = ActiveMotors.objects.filter(filter_objects, overall_status='Ok')
            # overall_obs = ActiveMotors.objects.filter(filter_objects, overall_status='Observation')
            # overall_inprocess = ActiveMotors.objects.filter(filter_objects, overall_status='In-process')


    # ===============================Pyro Devices =========================
            ncr_status_Ok = ActiveMotors.objects.filter(filter_objects, ncr_status='Ok')
            ncr_status_Obs = ActiveMotors.objects.filter(filter_objects,ncr_status='Observation')
            ncr_status_Inpro = ActiveMotors.objects.filter(filter_objects,ncr_status='In-process')

            overall_status_Ok = ActiveMotors.objects.filter(filter_objects, overall_status='Ok')
            overall_status_Obs = ActiveMotors.objects.filter(filter_objects,overall_status='Observation')
            overall_status_Inpro = ActiveMotors.objects.filter(filter_objects,overall_status='In-process')

            dist = {
                'qualification_insulation_lining_propellant_rm_OK' : qualification_insulation_lining_propellant_rm_OK.count(),
                'qualification_insulation_lining_propellant_rm_ObsSame' : qualification_insulation_lining_propellant_rm_ObsSame.count(),
                'qualification_insulation_lining_propellant_rm_ObsNext' : qualification_insulation_lining_propellant_rm_ObsNext.count(),
                'qualification_insulation_lining_propellant_rm_Up' : qualification_insulation_lining_propellant_rm_Up.count(),
                'qualification_insulation_lining_propellant_rm_Hlt' : qualification_insulation_lining_propellant_rm_Hlt.count(),

                'acceptance_casting_Ok' : acceptance_casting_Ok.count(),
                'acceptance_casting_ObsSame' : acceptance_casting_ObsSame.count(),
                'acceptance_casting_ObsNext' : acceptance_casting_ObsNext.count(),
                'acceptance_casting_Up' : acceptance_casting_Up.count(),
                'acceptance_casting_Hlt' : acceptance_casting_Hlt.count(),

                'sandblasting_Ok' : sandblasting_Ok.count(),
                'sandblasting_ObsSame' : sandblasting_ObsSame.count(),
                'sandblasting_ObsNext' : sandblasting_ObsNext.count(),
                'sandblasting_Up' : sandblasting_Up.count(),
                'sandblasting_Hlt' : sandblasting_Hlt.count(),

                'insulation_Ok' : insulation_Ok.count(),
                'insulation_ObsSame' : insulation_ObsSame.count(),
                'insulation_ObsNext' : insulation_ObsNext.count(),
                'insulation_Up' : insulation_Up.count(),
                'insulation_Hlt' : insulation_Hlt.count(),

                'ut_rt_insulated_case_Ok' : ut_rt_insulated_case_Ok.count(),
                'ut_rt_insulated_case_ObsSame' : ut_rt_insulated_case_ObsSame.count(),
                'ut_rt_insulated_case_ObsNext' : ut_rt_insulated_case_ObsNext.count(),
                'ut_rt_insulated_case_Up' : ut_rt_insulated_case_Up.count(),
                'ut_rt_insulated_case_Hlt' : ut_rt_insulated_case_Hlt.count(),

                'acceptance_silver_material_Ok' : acceptance_silver_material_Ok.count(),
                'acceptance_silver_material_ObsSame' : acceptance_silver_material_ObsSame.count(),
                'acceptance_silver_material_ObsNext' : acceptance_silver_material_ObsNext.count(),
                'acceptance_silver_material_Up' : acceptance_silver_material_Up.count(),
                'acceptance_silver_material_Hlt' : acceptance_silver_material_Hlt.count(),

                'silver_application_Ok' : silver_application_Ok.count(),
                'silver_application_ObsSame' : silver_application_ObsSame.count(),
                'silver_application_ObsNext' : silver_application_ObsNext.count(),
                'silver_application_Up' : silver_application_Up.count(),
                'silver_application_Hlt' : silver_application_Hlt.count(),

                'formulation_tailoring_liner_propellant_Ok' : formulation_tailoring_liner_propellant_Ok.count(),
                'formulation_tailoring_liner_propellant_ObsSame' : formulation_tailoring_liner_propellant_ObsSame.count(),
                'formulation_tailoring_liner_propellant_ObsNext' : formulation_tailoring_liner_propellant_ObsNext.count(),
                'formulation_tailoring_liner_propellant_Up' : formulation_tailoring_liner_propellant_Up.count(),
                'formulation_tailoring_liner_propellant_Hlt' : formulation_tailoring_liner_propellant_Hlt.count(),

                'conditioning_raw_materials_Ok' : conditioning_raw_materials_Ok.count(),
                'conditioning_raw_materials_ObsSame' : conditioning_raw_materials_ObsSame.count(),
                'conditioning_raw_materials_ObsNext' : conditioning_raw_materials_ObsNext.count(),
                'conditioning_raw_materials_Up' : conditioning_raw_materials_Up.count(),
                'conditioning_raw_materials_Hlt' : conditioning_raw_materials_Hlt.count(),

                'lining_Ok' : lining_Ok.count(),
                'lining_ObsSame' : lining_ObsSame.count(),
                'lining_ObsNext' : lining_ObsNext.count(),
                'lining_Up' : lining_Up.count(),
                'lining_Hlt' : lining_Hlt.count(),

                'casting_Ok' : casting_Ok.count(),
                'casting_ObsSame' : casting_ObsSame.count(),
                'casting_ObsNext' : casting_ObsNext.count(),
                'casting_Up' : casting_Up.count(),
                'casting_Hlt' : casting_Hlt.count(),

                'curing_Ok' : curing_Ok.count(),
                'curing_ObsSame' : curing_ObsSame.count(),
                'curing_ObsNext' : curing_ObsNext.count(),
                'curing_Up' : curing_Up.count(),
                'curing_Hlt' : curing_Hlt.count(),

                'liner_mechanical_properties_Ok' : liner_mechanical_properties_Ok.count(),
                'liner_mechanical_properties_ObsSame' : liner_mechanical_properties_ObsSame.count(),
                'liner_mechanical_properties_ObsNext' : liner_mechanical_properties_ObsNext.count(),
                'liner_mechanical_properties_Up' : liner_mechanical_properties_Up.count(),
                'liner_mechanical_properties_Hlt' : liner_mechanical_properties_Hlt.count(),




                'propellant_mechanical_properties_Ok' : propellant_mechanical_properties_Ok.count(),
                'propellant_mechanical_properties_ObsSame' : propellant_mechanical_properties_ObsSame.count(),
                'propellant_mechanical_properties_ObsNext' : propellant_mechanical_properties_ObsNext.count(),
                'propellant_mechanical_properties_Up' : propellant_mechanical_properties_Up.count(),
                'propellant_mechanical_properties_Hlt' : propellant_mechanical_properties_Hlt.count(),

                'interface_bond_strength_Ok' :   interface_bond_strength_Ok.count(),
                'interface_bond_strength_ObsSame' : interface_bond_strength_ObsSame.count(),
                'interface_bond_strength_ObsNext' : interface_bond_strength_ObsNext.count(),
                'interface_bond_strength_Up' : interface_bond_strength_Up.count(),
                'interface_bond_strength_Hlt' : interface_bond_strength_Hlt.count(),


                'propellant_burn_rate_Ok' : propellant_burn_rate_Ok.count(),
                'propellant_burn_rate_ObsSame' : propellant_burn_rate_ObsSame.count(),
                'propellant_burn_rate_ObsNext' : propellant_burn_rate_ObsNext.count(),
                'propellant_burn_rate_Up' : propellant_burn_rate_Up.count(),
                'propellant_burn_rate_Hlt' : propellant_burn_rate_Hlt.count(),




                'trimming_Propellant_grain_Ok' : trimming_Propellant_grain_Ok.count(),
                'trimming_Propellant_grain_ObsSame' : trimming_Propellant_grain_ObsSame.count(),
                'trimming_Propellant_grain_ObsNext' : trimming_Propellant_grain_ObsNext.count(),
                'trimming_Propellant_grain_Up' : trimming_Propellant_grain_Up.count(),
                'trimming_Propellant_grain_Hlt' : trimming_Propellant_grain_Hlt.count(),




                'mass_liner_insulation_propellant_srm_Ok' : mass_liner_insulation_propellant_srm_Ok.count(),
                'mass_liner_insulation_propellant_srm_ObsSame' : mass_liner_insulation_propellant_srm_ObsSame.count(),
                'mass_liner_insulation_propellant_srm_ObsNext' : mass_liner_insulation_propellant_srm_ObsNext.count(),
                'mass_liner_insulation_propellant_srm_Up' : mass_liner_insulation_propellant_srm_Up.count(),
                'mass_liner_insulation_propellant_srm_Hlt' : mass_liner_insulation_propellant_srm_Hlt.count(),



                'ut_endoscopy_rt_grain_Ok' : ut_endoscopy_rt_grain_Ok.count(),
                'ut_endoscopy_rt_grain_ObsSame' : ut_endoscopy_rt_grain_ObsSame.count(),
                'ut_endoscopy_rt_grain_ObsNext' : ut_endoscopy_rt_grain_ObsNext.count(),
                'ut_endoscopy_rt_grain_Up' : ut_endoscopy_rt_grain_Up.count(),
                'ut_endoscopy_rt_grain_Hlt' : ut_endoscopy_rt_grain_Hlt.count(),


                'matiral_qualified_status_Ok' : matiral_qualified_status_Ok.count(),
                'matiral_qualified_status_ObsSame' : matiral_qualified_status_ObsSame.count(),
                'matiral_qualified_status_ObsNext' : matiral_qualified_status_ObsNext.count(),
                'matiral_qualified_status_Up' : matiral_qualified_status_Up.count(),
                'matiral_qualified_status_Hlt' : matiral_qualified_status_Hlt.count(),



                'compo_manufacturing_status_Ok' : compo_manufacturing_status_Ok.count(),
                'compo_manufacturing_status_ObsSame' : compo_manufacturing_status_ObsSame.count(),
                'compo_manufacturing_status_ObsNext' : compo_manufacturing_status_ObsNext.count(),
                'compo_manufacturing_status_Up' : compo_manufacturing_status_Up.count(),
                'compo_manufacturing_status_Hlt' : compo_manufacturing_status_Hlt.count(),



                'powerpack_assembly_status_Ok' : powerpack_assembly_status_Ok.count(),
                'powerpack_assembly_status_ObsSame' : powerpack_assembly_status_ObsSame.count(),
                'powerpack_assembly_status_ObsNext' : powerpack_assembly_status_ObsNext.count(),
                'powerpack_assembly_status_Up' : powerpack_assembly_status_Up.count(),
                'powerpack_assembly_status_Hlt' : powerpack_assembly_status_Hlt.count(),



                'powerpack_testing_status_Ok' :  powerpack_testing_status_Ok.count(),
                'powerpack_testing_status_ObsSame' : powerpack_testing_status_ObsSame.count(),
                'powerpack_testing_status_ObsNext' : powerpack_testing_status_ObsNext.count(),
                'powerpack_testing_status_Up' : powerpack_testing_status_Up.count(),
                'powerpack_testing_status_Hlt' : powerpack_testing_status_Hlt.count(),



                'raw_material_inspection_status_Ok' : raw_material_inspection_status_Ok.count(),
                'raw_material_inspection_status_ObsSame' : raw_material_inspection_status_ObsSame.count(),
                'raw_material_inspection_status_ObsNext' : raw_material_inspection_status_ObsNext.count(),
                'raw_material_inspection_status_Up' : raw_material_inspection_status_Up.count(),
                'raw_material_inspection_status_Hlt' : raw_material_inspection_status_Hlt.count(),



                'pressing_electrode_status_Ok' : pressing_electrode_status_Ok.count(),
                'pressing_electrode_status_ObsSame' : pressing_electrode_status_ObsSame.count(),
                'pressing_electrode_status_ObsNext' : pressing_electrode_status_ObsNext.count(),
                'pressing_electrode_status_Up' : pressing_electrode_status_Up.count(),
                'pressing_electrode_status_Hlt' : pressing_electrode_status_Hlt.count(),



                'formation_process_status_Ok' : formation_process_status_Ok.count(),
                'formation_process_status_ObsSame' : formation_process_status_ObsSame.count(),
                'formation_process_status_ObsNext' : formation_process_status_ObsNext.count(),
                'formation_process_status_Up' : formation_process_status_Up.count(),
                'formation_process_status_Hlt' : formation_process_status_Hlt.count(),



                'assembly_process_status_Ok' : assembly_process_status_Ok.count(),
                'assembly_process_status_ObsSame' : assembly_process_status_ObsSame.count(),
                'assembly_process_status_ObsNext' : assembly_process_status_ObsNext.count(),
                'assembly_process_status_Up' : assembly_process_status_Up.count(),
                'assembly_process_status_Hlt' : assembly_process_status_Hlt.count(),


                'battery_testing_status_Ok' : battery_testing_status_Ok.count(),
                'battery_testing_status_ObsSame' : battery_testing_status_ObsSame.count(),
                'battery_testing_status_ObsNext' : battery_testing_status_ObsNext.count(),
                'battery_testing_status_Up' : battery_testing_status_Up.count(),
                'battery_testing_status_Hlt' : battery_testing_status_Hlt.count(),

                'final_qualification_status_Ok' : final_qualification_status_Ok.count(),
                'final_qualification_status_ObsSame' : final_qualification_status_ObsSame.count(),
                'final_qualification_status_ObsNext' : final_qualification_status_ObsNext.count(),
                'final_qualification_status_Up' : final_qualification_status_Up.count(),
                'final_qualification_status_Hlt' : final_qualification_status_Hlt.count(),


                'qualification_raw_material_status_Ok' : qualification_raw_material_status_Ok.count(),
                'qualification_raw_material_status_ObsSame' : qualification_raw_material_status_ObsSame.count(),
                'qualification_raw_material_status_ObsNext' : qualification_raw_material_status_ObsNext.count(),
                'qualification_raw_material_status_Up' : qualification_raw_material_status_Up.count(),
                'qualification_raw_material_status_Hlt' : qualification_raw_material_status_Hlt.count(),

                'filling_status_Ok' : filling_status_Ok.count(),
                'filling_status_ObsSame' : filling_status_ObsSame.count(),
                'filling_status_ObsNext' : filling_status_ObsNext.count(),
                'filling_status_Up' : filling_status_Up.count(),
                'filling_status_Hlt' : filling_status_Hlt.count(),

                'assembling_integration_status_Ok' : assembling_integration_status_Ok.count(),
                'assembling_integration_status_ObsSame' : assembling_integration_status_ObsSame.count(),
                'assembling_integration_status_ObsNext' : assembling_integration_status_ObsNext.count(),
                'assembling_integration_status_Up' : assembling_integration_status_Up.count(),
                'assembling_integration_status_Hlt' : assembling_integration_status_Hlt.count(),

                'qualification_testing_status_Ok' : qualification_testing_status_Ok.count(),
                'qualification_testing_status_ObsSame' : qualification_testing_status_ObsSame.count(),
                'qualification_testing_status_ObsNext' : qualification_testing_status_ObsNext.count(),
                'qualification_testing_status_Up' : qualification_testing_status_Up.count(),
                'qualification_testing_status_Hlt' : qualification_testing_status_Hlt.count(),

                'performance_testing_status_Ok' : performance_testing_status_Ok.count(),
                'performance_testing_status_ObsSame' : performance_testing_status_ObsSame.count(),
                'performance_testing_status_ObsNext' : performance_testing_status_ObsNext.count(),
                'performance_testing_status_Up' : performance_testing_status_Up.count(),
                'performance_testing_status_Hlt' : performance_testing_status_Hlt.count(),


                'ncr_status_Ok' : ncr_status_Ok.count(),
                'ncr_status_Obs' : ncr_status_Obs.count(),
                'ncr_status_Inpro' : ncr_status_Inpro.count(),

                'overall_status_Ok' : overall_status_Ok.count(),
                'overall_status_Obs' : overall_status_Obs.count(),
                'overall_status_Inpro' : overall_status_Inpro.count(),


            }
            return JsonResponse({'message': 'true', 'data': dist}, status=200)
        # except Exception as e:
        #     print(e)
        #     return JsonResponse({'message': 'Sorry! No Task found.'}, status=500)
