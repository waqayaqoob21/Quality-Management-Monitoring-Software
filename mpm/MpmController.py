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
from csv import reader
import os
import csv
from dateutil.relativedelta import relativedelta
from django.db.models.functions import *

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
                motorModel.creation_date = request['creation_date']
                motorModel.testing_date = request['testing_date']
                motorModel.motor_id = request['motor_id']
                motorModel.component_type = request['component_type']
                motorModel.qualification_insulation_lining_propellant_rm = request[
                    'qualification_insulation_lining_propellant_rm']
                if request['qualification_insulation_lining_propellant_rm_date'] != '':
                    motorModel.qualification_insulation_lining_propellant_rm_date = request[
                    'qualification_insulation_lining_propellant_rm_date']
                motorModel.qualification_insulation_lining_propellant_rm_remarks = request[
                    'qualification_insulation_lining_propellant_remarks']
                motorModel.acceptance_casting = request['Acceptance_of_casting']
                if request['Acceptance_of_casting_date'] != '':
                    motorModel.acceptance_casting_date = request['Acceptance_of_casting_date']
                motorModel.acceptance_casting_remarks = request['Acceptance_of_casting_remarks']
                motorModel.sandblasting = request['Sand_blasting']
                if request['Sand_blasting_date'] != '':
                    motorModel.sandblasting_date = request['Sand_blasting_date']
                motorModel.sandblasting_remarks = request['Sand_blasting_remarks']

                motorModel.insulation = request['motor_Insulation']
                if request['motor_Insulation_date'] != '':
                    motorModel.insulation_date = request['motor_Insulation_date']
                motorModel.insulation_remarks = request['motor_Insulation_remarks']

                motorModel.ut_rt_insulated_case = request['UTandRTofInsulated_Case']
                if request['UTandRTofInsulated_Case_date'] != '':
                    motorModel.ut_rt_insulated_case_date = request['UTandRTofInsulated_Case_date']
                motorModel.ut_rt_insulated_case_remarks = request['UTandRTofInsulated_Case_Remarks']

                motorModel.acceptance_silver_material = request['AcceptanceofSilver_Material']
                if request['AcceptanceofSilver_Material_date'] != '':
                    motorModel.acceptance_silver_material_date = request['AcceptanceofSilver_Material_date']
                motorModel.acceptance_silver_material_remarks = request['AcceptanceofSilver_Material_Remarks']

                motorModel.silver_application = request['Silver_Application']
                if request['Silver_Application_date'] != '':
                    motorModel.silver_application_date = request['Silver_Application_date']
                motorModel.silver_application_remarks = request['Silver_Application_Remarks']

                motorModel.formulation_tailoring_liner_propellant = request['Formulationtailoringoflinerand_propellant']
                if request['Formulationtailoringoflinerand_propellant_date'] != '':
                    motorModel.formulation_tailoring_liner_propellant_date = request['Formulationtailoringoflinerand_propellant_date']
                motorModel.formulation_tailoring_liner_propellant_remarks = request[
                    'Formulationtailoringoflinerand_propellant_remarks']

                motorModel.conditioning_raw_materials = request['ConditioningofRaw_Material']
                if request['ConditioningofRaw_Material_date'] != '':
                    motorModel.conditioning_raw_materials_date = request['ConditioningofRaw_Material_date']
                motorModel.conditioning_raw_materials_remarks = request['ConditioningofRaw_Material_remarks']

                motorModel.lining = request['Lining_motor']
                if request['Lining_motor_date'] != '':
                    motorModel.lining_date = request['Lining_motor_date']
                motorModel.lining_remarks = request['Lining_motor_remarks']

                motorModel.casting = request['Casting_motor']
                if request['Casting_motor_date'] != '':
                    motorModel.casting_date = request['Casting_motor_date']
                motorModel.casting_remarks = request['Casting_motor_remarks']

                motorModel.curing = request['Curing_motor']
                if request['Curing_motor_date'] != '':
                    motorModel.curing_date = request['Curing_motor_date']
                motorModel.curing_remarks = request['Curing_motor_remarks']

                motorModel.liner_mechanical_properties = request['Liner_Mechanical_Properties']
                if request['Liner_Mechanical_Properties_date'] != '':
                    motorModel.liner_mechanical_properties_date = request['Liner_Mechanical_Properties_date']
                motorModel.liner_mechanical_properties_remarks = request['Liner_Mechanical_Properties_Remarks']

                motorModel.propellant_mechanical_properties = request['Propellant_Mechanical_Properties']
                if request['Propellant_Mechanical_Properties_date'] != '':
                    motorModel.propellant_mechanical_properties_date = request['Propellant_Mechanical_Properties_date']
                motorModel.propellant_mechanical_properties_remarks = request[
                    'Propellant_Mechanical_Properties_Remarks']

                motorModel.interface_bond_strength = request['Interfacebond_strength']
                if request['Interfacebond_strength_date'] != '':
                    motorModel.interface_bond_strength_date = request['Interfacebond_strength_date']
                motorModel.interface_bond_strength_remarks = request['Interfacebond_strength_Remarks']

                motorModel.propellant_burn_rate = request['Propellant_burn_rate']
                if request['Propellant_burn_rate_date'] != '':
                    motorModel.propellant_burn_rate_date = request['Propellant_burn_rate_date']
                motorModel.propellant_burn_rate_remarks = request['Propellant_burn_rate_Remarks']

                motorModel.trimming_Propellant_grain = request['Trimming_of_propellant_grain']
                if request['Trimming_of_propellant_grain_date'] != '':
                    motorModel.trimming_Propellant_grain_date = request['Trimming_of_propellant_grain_date']
                motorModel.trimming_Propellant_grain_remarks = request['Trimming_of_propellant_grain_remarks']

                motorModel.mass_liner_insulation_propellant_srm = request['mass_liner_insulation_propellant_srm']
                if request['mass_liner_insulation_propellant_srm_date'] != '':
                    motorModel.mass_liner_insulation_propellant_srm_date = request['mass_liner_insulation_propellant_srm_date']
                motorModel.mass_liner_insulation_propellant_srm_remarks = request[
                    'mass_liner_insulation_propellant_srm_remarks']

                motorModel.ut_endoscopy_rt_grain = request['ut_endoscopy_rt_grain']
                if request['ut_endoscopy_rt_grain_date'] != '':
                    motorModel.ut_endoscopy_rt_grain_date = request['ut_endoscopy_rt_grain_date']
                motorModel.ut_endoscopy_rt_grain_remarks = request['ut_endoscopy_rt_grain_remarks']

                motorModel.conditioning_of_lining = request['conditioning_of_lining']
                if request['conditioning_of_lining_date'] != '':
                    motorModel.conditioning_of_lining_date = request['conditioning_of_lining_date']
                motorModel.conditioning_of_lining_remarks = request['conditioning_of_lining_remarks']

                motorModel.mechanical_properties_liner = request['mechanical_properties_liner']
                if request['mechanical_properties_liner_date'] != '':
                    motorModel.mechanical_properties_liner_date = request['mechanical_properties_liner_date']
                motorModel.mechanical_properties_liner_remarks = request['mechanical_properties_liner_remarks']

                motorModel.mechanical_properties_propellant = request['mechanical_properties_propellant']
                if request['mechanical_properties_propellant_date'] != '':
                    motorModel.mechanical_properties_propellant_date = request['mechanical_properties_propellant_date']
                motorModel.mechanical_properties_propellant_remarks = request['mechanical_properties_propellant_remarks']

                motorModel.mass_liner = request['mass_liner']
                if request['mass_liner_date'] != '':
                    motorModel.mass_liner_date = request['mass_liner_date']
                motorModel.mass_liner_remarks = request['mass_liner_remarks']

                motorModel.mass_insulation = request['mass_insulation']
                if request['mass_insulation_date'] != '':
                    motorModel.mass_insulation_date = request['mass_insulation_date']
                motorModel.mass_insulation_remarks = request['mass_insulation_remarks']

                motorModel.mass_propellant = request['mass_propellant']
                if request['mass_propellant_date'] != '':
                    motorModel.mass_propellant_date = request['mass_propellant_date']
                motorModel.mass_propellant_remarks = request['mass_propellant_remarks']

                motorModel.overall_qualification_status = request['overall_qualification_status']
                if request['overall_qualification_date'] != '':
                    motorModel.overall_qualification_date = request['overall_qualification_date']
                motorModel.overall_qualification_remarks = request['overall_qualification_remarks']

                motorModel.bhd_status = request['bhd_status']
                if request['bhd_date'] != '':
                    motorModel.bhd_date = request['bhd_date']
                motorModel.bhd_remarks = request['bhd_remarks']

                motorModel.qm_certification_status = request['qm_certification_status']
                if request['qm_certification_date'] != '':
                    motorModel.qm_certification_date = request['qm_certification_date']
                motorModel.qm_certification_remarks = request['qm_certification_remarks']



                motorModel.battery_type = request['battery_type']
                motorModel.tb_type = request['tb_type']
                motorModel.battery_id = request['battery_id']
                motorModel.lot_id = request['lot_id']

                if request['matiral_qualified_date'] != '':
                    motorModel.matiral_qualified_date = request['matiral_qualified_date']
                motorModel.matiral_qualified_status = request['matiral_qualified_status']
                motorModel.matiral_qualified_remarks = request['matiral_qualified_remarks']

                if request['compo_manufacturing_date'] != '':
                    motorModel.compo_manufacturing_date = request['compo_manufacturing_date']
                motorModel.compo_manufacturing_status = request['compo_manufacturing_status']
                motorModel.compo_manufacturing_remarks = request['compo_manufacturing_remarks']

                if request['powerpack_assembly_date'] != '':
                    motorModel.powerpack_assembly_date = request['powerpack_assembly_date']
                motorModel.powerpack_assembly_status = request['powerpack_assembly_status']
                motorModel.powerpack_assembly_remarks = request['powerpack_assembly_remarks']

                if request['powerpack_testing_date'] != '':
                    motorModel.powerpack_testing_date = request['powerpack_testing_date']
                motorModel.powerpack_testing_status = request['powerpack_testing_status']
                motorModel.powerpack_testing_remarks = request['powerpack_testing_remarks']

                if request['raw_material_inspection_date'] != '':
                    motorModel.raw_material_inspection_date = request['raw_material_inspection_date']
                motorModel.raw_material_inspection_status = request['raw_material_inspection_status']
                motorModel.raw_material_inspection_remarks = request['raw_material_inspection_remarks']

                if request['pressing_electrode_date'] != '':
                    motorModel.pressing_electrode_date = request['pressing_electrode_date']
                motorModel.pressing_electrode_status = request['pressing_electrode_status']
                motorModel.pressing_electrode_remarks = request['pressing_electrode_remarks']

                if request['formation_process_date'] != '':
                    motorModel.formation_process_date = request['formation_process_date']
                motorModel.formation_process_status = request['formation_process_status']
                motorModel.formation_process_remarks = request['formation_process_remarks']

                if request['assembly_process_date'] != '':
                    motorModel.assembly_process_date = request['assembly_process_date']
                motorModel.assembly_process_status = request['assembly_process_status']
                motorModel.assembly_process_remarks = request['assembly_process_remarks']

                if request['battery_testing_date'] != '':
                    motorModel.battery_testing_date = request['battery_testing_date']
                motorModel.battery_testing_status = request['battery_testing_status']
                motorModel.battery_testing_remarks = request['battery_testing_remarks']

                if request['final_qualification_date'] != '':
                    motorModel.final_qualification_date = request['final_qualification_date']
                motorModel.final_qualification_status = request['final_qualification_status']
                motorModel.final_qualification_remarks = request['final_qualification_remarks']

                motorModel.pd_type = request['pd_type']
                motorModel.pd_id = request['pd_id']

                if request['qualification_raw_material_date'] != '':
                    motorModel.qualification_raw_material_date = request['qualification_raw_material_date']
                motorModel.qualification_raw_material_status = request['qualification_raw_material_status']
                motorModel.qualification_raw_material_remarks = request['qualification_raw_material_remarks']

                if request['filling_date'] != '':
                    motorModel.filling_date = request['filling_date']
                motorModel.filling_status = request['filling_status']
                motorModel.filling_remarks = request['filling_remarks']

                if request['assembling_integration_date'] != '':
                    motorModel.assembling_integration_date = request['assembling_integration_date']
                motorModel.assembling_integration_status = request['assembling_integration_status']
                motorModel.assembling_integration_remarks = request['assembling_integration_remarks']

                if request['qualification_testing_date'] != '':
                    motorModel.qualification_testing_date = request['qualification_testing_date']
                motorModel.qualification_testing_status = request['qualification_testing_status']
                motorModel.qualification_testing_remarks = request['qualification_testing_remarks']

                if request['performance_testing_date'] != '':
                    motorModel.performance_testing_date = request['performance_testing_date']
                motorModel.performance_testing_status = request['performance_testing_status']
                motorModel.performance_testing_remarks = request['performance_testing_remarks']

                motorModel.ncr_status = request['ncr_status']
                motorModel.ncr_status_remarks = request['ncr_status_ramarks']
                motorModel.overall_status = request['overall_status']
                motorModel.overall_remarks = request['overall_status_ramarks']
                if request['user_id'] != '':
                    motorModel.user_id = request['user_id']
                else:
                    motorModel.user_id = None
                if request['base64File'] != '':
                    motorModel.attachments = request['base64File']
                motorModel.save()
                return JsonResponse({'status': 'True', 'message': "Motor Created Successfully!"},
                                    status=200)
            else:
                get_motor = ActiveMotors.objects.filter(id=id).first()
                if get_motor is not None:
                    if request['qualification_insulation_lining_propellant_rm'] != get_motor.qualification_insulation_lining_propellant_rm \
                        or request['qualification_insulation_lining_propellant_remarks'] != get_motor.qualification_insulation_lining_propellant_rm \
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
                        or request['performance_testing_status'] != get_motor.performance_testing_status or request['qualification_testing_status'] != get_motor.qualification_testing_status \
                        or request['qualification_insulation_lining_propellant_rm_date'] != get_motor.qualification_insulation_lining_propellant_rm_date \
                        or request['Acceptance_of_casting_date'] != get_motor.acceptance_casting or request['Sand_blasting_date'] != get_motor.sandblasting_date \
                        or request['motor_Insulation_date'] != get_motor.insulation_date  \
                        or request['UTandRTofInsulated_Case_date'] != get_motor.ut_rt_insulated_case_date \
                        or request['AcceptanceofSilver_Material_date'] != get_motor.acceptance_silver_material_date \
                        or request['Silver_Application_date'] != get_motor.silver_application_date \
                        or request['Formulationtailoringoflinerand_propellant_date'] != get_motor.formulation_tailoring_liner_propellant_date \
                        or request['ConditioningofRaw_Material_date'] != get_motor.conditioning_raw_materials_date \
                        or request['Lining_motor_date'] != get_motor.lining or request['Casting_motor'] != get_motor.casting_date \
                        or request['Curing_motor_date'] != get_motor.curing_date \
                        or request['Liner_Mechanical_Properties_date'] != get_motor.liner_mechanical_properties_date \
                        or request['Propellant_Mechanical_Properties_date'] != get_motor.propellant_mechanical_properties_date \
                        or request['Interfacebond_strength_date'] != get_motor.interface_bond_strength_date \
                        or request['Propellant_burn_rate_date'] != get_motor.propellant_burn_rate_date \
                        or request['Trimming_of_propellant_grain_date'] != get_motor.trimming_Propellant_grain_date \
                        or request['mass_liner_insulation_propellant_srm_date'] != get_motor.mass_liner_insulation_propellant_srm_date \
                        or request['ut_endoscopy_rt_grain_date'] != get_motor.ut_endoscopy_rt_grain_date \
                        or request['ncr_status_ramarks'] != get_motor.ncr_status_remarks \
                        or request['matiral_qualified_date'] != get_motor.matiral_qualified_date \
                        or request['compo_manufacturing_date'] != get_motor.compo_manufacturing_date \
                        or request['powerpack_assembly_date'] != get_motor.powerpack_assembly_date \
                        or request['powerpack_testing_date'] != get_motor.powerpack_testing_date \
                        or request['raw_material_inspection_date'] != get_motor.raw_material_inspection_date \
                        or request['pressing_electrode_date'] != get_motor.pressing_electrode_date \
                        or request['formation_process_date'] != get_motor.formation_process_date \
                        or request['assembly_process_date'] != get_motor.assembly_process_date \
                        or request['battery_testing_date'] != get_motor.battery_testing_date \
                        or request['final_qualification_date'] != get_motor.final_qualification_date \
                        or request['qualification_raw_material_date'] != get_motor.qualification_raw_material_date \
                        or request['filling_date'] != get_motor.filling_date \
                        or request['assembling_integration_date'] != get_motor.assembling_integration_date \
                        or request['qualification_testing_date'] != get_motor.qualification_testing_date \
                        or request['performance_testing_date'] != get_motor.performance_testing_date \
                            or request['Acceptance_of_casting_remarks'] != get_motor.acceptance_casting_remarks or request['Sand_blasting_remarks'] != get_motor.sandblasting_remarks \
                            or request['motor_Insulation_remarks'] != get_motor.insulation_remarks \
                            or request['UTandRTofInsulated_Case_Remarks'] != get_motor.ut_rt_insulated_case_remarks \
                            or request['AcceptanceofSilver_Material_Remarks'] != get_motor.acceptance_silver_material_remarks \
                            or request['Silver_Application_Remarks'] != get_motor.silver_application_remarks \
                            or request['Formulationtailoringoflinerand_propellant_remarks'] != get_motor.formulation_tailoring_liner_propellant_remarks \
                            or request['ConditioningofRaw_Material_remarks'] != get_motor.conditioning_raw_materials_remarks \
                            or request['conditioning_of_lining_remarks'] != get_motor.conditioning_of_lining_remarks \
                            or request['Lining_motor_remarks'] != get_motor.lining or request['Casting_motor_remarks'] != get_motor.casting_remarks \
                            or request['Curing_motor_remarks'] != get_motor.curing_remarks \
                            or request['Liner_Mechanical_Properties_Remarks'] != get_motor.liner_mechanical_properties_remarks \
                            or request['Propellant_Mechanical_Properties_Remarks'] != get_motor.propellant_mechanical_properties_remarks \
                            or request['Interfacebond_strength_Remarks'] != get_motor.interface_bond_strength_remarks \
                            or request['Propellant_burn_rate_Remarks'] != get_motor.propellant_burn_rate_remarks \
                            or request['Trimming_of_propellant_grain_remarks'] != get_motor.trimming_Propellant_grain_remarks \
                            or request['mass_liner_insulation_propellant_srm_remarks'] != get_motor.mass_liner_insulation_propellant_srm_remarks \
                            or request['ut_endoscopy_rt_grain_remarks'] != get_motor.ut_endoscopy_rt_grain_remarks \
                            or request['ncr_status_ramarks'] != get_motor.ncr_status_remarks \
                            or request['matiral_qualified_remarks'] != get_motor.matiral_qualified_remarks \
                            or request['compo_manufacturing_remarks'] != get_motor.compo_manufacturing_remarks \
                            or request['powerpack_assembly_remarks'] != get_motor.powerpack_assembly_remarks \
                            or request['powerpack_testing_remarks'] != get_motor.powerpack_testing_remarks \
                            or request['raw_material_inspection_remarks'] != get_motor.raw_material_inspection_remarks \
                            or request['pressing_electrode_remarks'] != get_motor.pressing_electrode_remarks \
                            or request['formation_process_remarks'] != get_motor.formation_process_remarks \
                            or request['assembly_process_remarks'] != get_motor.assembly_process_remarks \
                            or request['battery_testing_remarks'] != get_motor.battery_testing_remarks \
                            or request['final_qualification_remarks'] != get_motor.final_qualification_remarks \
                            or request['qualification_raw_material_remarks'] != get_motor.qualification_raw_material_remarks \
                            or request['filling_remarks'] != get_motor.filling_remarks \
                            or request['assembling_integration_remarks'] != get_motor.assembling_integration_remarks \
                            or request['qualification_testing_remarks'] != get_motor.qualification_testing_remarks \
                            or request['performance_testing_remarks'] != get_motor.performance_testing_remarks \
                            or request['ncr_status_ramarks'] != get_motor.ncr_status_remarks \
                            or request['matiral_qualified_remarks'] != get_motor.matiral_qualified_remarks \
                            or request['compo_manufacturing_remarks'] != get_motor.compo_manufacturing_remarks \
                            or request['powerpack_assembly_remarks'] != get_motor.powerpack_assembly_remarks \
                            or request['powerpack_testing_remarks'] != get_motor.powerpack_testing_remarks \
                            or request['mechanical_properties_liner'] != get_motor.mechanical_properties_liner \
                            or request['mechanical_properties_liner_date'] != get_motor.mechanical_properties_liner_date \
                            or request['mechanical_properties_liner_remarks'] != get_motor.mechanical_properties_liner_remarks \
                            or request['mechanical_properties_propellant'] != get_motor.mechanical_properties_propellant \
                            or request['mechanical_properties_propellant_remarks'] != get_motor.mechanical_properties_propellant_remarks \
                            or request['mechanical_properties_propellant_remarks'] != get_motor.mechanical_properties_propellant_remarks \
                            or request['mass_liner'] != get_motor.mass_liner \
                            or request['mass_liner_date'] != get_motor.mass_liner_date \
                            or request['mass_liner_remarks'] != get_motor.mass_liner_remarks \
                            or request['mass_insulation'] != get_motor.mass_insulation \
                            or request['mass_insulation_date'] != get_motor.mass_insulation_date \
                            or request['mass_insulation_remarks'] != get_motor.mass_insulation_remarks \
                            or request['mass_propellant'] != get_motor.mass_propellant \
                            or request['mass_propellant_date'] != get_motor.mass_propellant_date \
                            or request['mass_propellant_remarks'] != get_motor.mass_propellant_remarks \
                            or request['overall_qualification_status'] != get_motor.overall_qualification_status \
                            or request['overall_qualification_date'] != get_motor.overall_qualification_date \
                            or request['overall_qualification_remarks'] != get_motor.overall_qualification_remarks \
                            or request['bhd_status'] != get_motor.bhd_status \
                            or request['bhd_date'] != get_motor.bhd_date \
                            or request['bhd_remarks'] != get_motor.bhd_remarks \
                            or request['qm_certification_status'] != get_motor.qm_certification_status \
                            or request['qm_certification_date'] != get_motor.qm_certification_date \
                            or request['qm_certification_remarks'] != get_motor.qm_certification_remarks:

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

                        MotorHistory.conditioning_of_lining = get_motor.conditioning_of_lining
                        MotorHistory.conditioning_of_lining_date = get_motor.conditioning_of_lining_date
                        MotorHistory.conditioning_of_lining_remarks = get_motor.conditioning_of_lining_remarks

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

                        MotorHistory.mechanical_properties_liner = get_motor.mechanical_properties_liner
                        MotorHistory.mechanical_properties_liner_date = get_motor.mechanical_properties_liner_date
                        MotorHistory.mechanical_properties_liner_remarks = get_motor.mechanical_properties_liner_remarks

                        MotorHistory.mechanical_properties_propellant = get_motor.mechanical_properties_propellant
                        MotorHistory.mechanical_properties_propellant_date = get_motor.mechanical_properties_propellant_date
                        MotorHistory.mechanical_properties_propellant_remarks = get_motor.mechanical_properties_propellant_remarks

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

                        MotorHistory.mass_liner = get_motor.mass_liner
                        MotorHistory.mass_liner_date = get_motor.mass_liner_date
                        MotorHistory.mass_liner_remarks = get_motor.mass_liner_remarks

                        MotorHistory.mass_insulation = get_motor.mass_insulation
                        MotorHistory.mass_insulation_date = get_motor.mass_insulation_date
                        MotorHistory.mass_insulation_remarks = get_motor.mass_insulation_remarks

                        MotorHistory.mass_propellant = get_motor.mass_propellant
                        MotorHistory.mass_propellant_date = get_motor.mass_propellant_date
                        MotorHistory.mass_propellant_remarks = get_motor.mass_propellant_remarks

                        MotorHistory.overall_qualification_status = get_motor.overall_qualification_status
                        MotorHistory.overall_qualification_date = get_motor.overall_qualification_date
                        MotorHistory.overall_qualification_remarks = get_motor.overall_qualification_remarks

                        MotorHistory.bhd_status = get_motor.bhd_status
                        MotorHistory.bhd_date = get_motor.bhd_date
                        MotorHistory.bhd_remarks = get_motor.bhd_remarks

                        MotorHistory.qm_certification_status = get_motor.qm_certification_status
                        MotorHistory.qm_certification_date = get_motor.qm_certification_date
                        MotorHistory.qm_certification_remarks = get_motor.qm_certification_remarks
                        
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
                        MotorHistory.user_id = get_motor.user_id
                        MotorHistory.save()

                    get_motor.system_type = request['system_type']
                    get_motor.system_name = request['system_name']
                    get_motor.organization = request['organization']
                    get_motor.testing_type = request['testing_type']
                    get_motor.creation_date = request['creation_date']
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

                    get_motor.conditioning_of_lining = request['conditioning_of_lining']
                    get_motor.conditioning_of_lining_date = request['conditioning_of_lining_date']
                    get_motor.conditioning_of_lining_remarks = request['conditioning_of_lining_remarks']

                    get_motor.mechanical_properties_liner = request['mechanical_properties_liner']
                    get_motor.mechanical_properties_liner_date = request['mechanical_properties_liner_date']
                    get_motor.mechanical_properties_liner_remarks = request['mechanical_properties_liner_remarks']

                    get_motor.mechanical_properties_propellant = request['mechanical_properties_propellant']
                    get_motor.mechanical_properties_propellant_date = request['mechanical_properties_propellant_date']
                    get_motor.mechanical_properties_propellant_remarks = request[
                        'mechanical_properties_propellant_remarks']

                    get_motor.mass_liner = request['mass_liner']
                    get_motor.mass_liner_date = request['mass_liner_date']
                    get_motor.mass_liner_remarks = request['mass_liner_remarks']

                    get_motor.mass_insulation = request['mass_insulation']
                    get_motor.mass_insulation_date = request['mass_insulation_date']
                    get_motor.mass_insulation_remarks = request['mass_insulation_remarks']

                    get_motor.mass_propellant = request['mass_propellant']
                    get_motor.mass_propellant_date = request['mass_propellant_date']
                    get_motor.mass_propellant_remarks = request['mass_propellant_remarks']

                    get_motor.overall_qualification_status = request['overall_qualification_status']
                    get_motor.overall_qualification_date = request['overall_qualification_date']
                    get_motor.overall_qualification_remarks = request['overall_qualification_remarks']

                    get_motor.bhd_status = request['bhd_status']
                    get_motor.bhd_date = request['bhd_date']
                    get_motor.bhd_remarks = request['bhd_remarks']

                    get_motor.qm_certification_status = request['qm_certification_status']
                    get_motor.qm_certification_date = request['qm_certification_date']
                    get_motor.qm_certification_remarks = request['qm_certification_remarks']

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
                    if request['user_id'] != '':
                        get_motor.user_id = request['user_id']
                    else:
                        get_motor.user_id = None
                    if request['base64File'] != '':
                        get_motor.attachments = request['base64File']
                    get_motor.save()
                    return JsonResponse({'status': 'True', 'message': "Motor Created Successfully!"},
                                        status=200)
        except Exception as e:
            return JsonResponse({'status': 'False', "message": "Motor Not Saved"}, status=500)


    @staticmethod
    def ImporProcessCsv(request):
        # try:
            id = request['id']
            importedCsvFile = request["csv_file"]
            system_type = request["component_type"]
            if not os.path.isdir('imported_files'):
                os.mkdir('imported_files')
            path = "imported_files/"
            fs = FileSystemStorage(location=path)
            fs.save(importedCsvFile.name, importedCsvFile)
            file_path = path + importedCsvFile.name
            user_id = 0
            if request['user_id'] != '':
                user_id = request['user_id']
            if id == '0':
                file = open(file_path)
                csvf = csv.reader(file)
                next(csvf, None)
                data = []
                if system_type == "Motor":
                    for component_type,system_name,sys_type,testing_type,organization,motor_id,\
                            qualification_insulation_lining_propellant_rm_date,qualification_insulation_lining_propellant_rm,qualification_insulation_lining_propellant_rm_remarks,\
                            lining_date,lining_status,lining_remarks,propellant_mechanical_properties_date,propellant_mechanical_properties_status,propellant_mechanical_properties_remarks,\
                            sandblasting_date,sandblasting_status,sandblasting_remarks,insulation_application_date,insulation_application_status,insulation_application_remarks,\
                            ut_rt_insulated_case_date,ut_rt_insulated_case_status,ut_rt_insulated_case_remarks,sliver_acceptance_date,sliver_acceptance_status,\
                            sliver_acceptance_remarks,sliver_application_date,sliver_application_status,sliver_application_remarks,\
                            formulation_tailoring_date,formulation_tailoring_status,formulation_tailoring_remarks,\
                            conditioning_raw_materials_date,conditioning_raw_materials_status,conditioning_raw_materials_remarks,\
                            conditioning_of_lining_date,conditioning_of_lining_status,conditioning_of_lining_remarks,casting_date,casting_status,casting_remarks,\
                            curing_date,curing_status,curing_remarks,ut_endoscopy_rt_grain_date,ut_endoscopy_rt_grain_status,ut_endoscopy_rt_grain_remarks,\
                            mechanical_properties_liner_date,mechanical_properties_liner_status,mechanical_properties_liner_remarks,\
                            mechanical_properties_propellant_date,mechanical_properties_propellant_status,mechanical_properties_propellant_remarks,\
                            interface_bond_strength_date,interface_bond_strength_status,interface_bond_strength_remarks,\
                            propellant_burn_rate_date,propellant_burn_rate_status,propellant_burn_rate_remarks,mass_liner_date,mass_liner_status,mass_liner_remarks,\
                            mass_insulation_date,mass_insulation_status,mass_insulation_remarks,mass_propellant_date,mass_propellant_status,mass_propellant_remarks,\
                            overall_qualification_date,overall_qualification_status,overall_qualification_remarks,\
                            bhd_date,bhd_status,bhd_remarks,qm_certification_date,qm_certification_status,qm_certification_remarks, *__ in csvf:

                        motor_system = ActiveMotors(component_type= component_type,system_name=system_name,system_type=sys_type,testing_type=testing_type,organization=organization,motor_id=motor_id,
                            qualification_insulation_lining_propellant_rm_date= datetime.strptime(qualification_insulation_lining_propellant_rm_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),qualification_insulation_lining_propellant_rm= qualification_insulation_lining_propellant_rm,qualification_insulation_lining_propellant_rm_remarks= qualification_insulation_lining_propellant_rm_remarks,
                            lining_date = datetime.strptime(lining_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),lining = lining_status,lining_remarks= lining_remarks,
                            propellant_mechanical_properties_date = datetime.strptime(propellant_mechanical_properties_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),propellant_mechanical_properties= propellant_mechanical_properties_status,propellant_mechanical_properties_remarks= propellant_mechanical_properties_remarks,
                            sandblasting_date = datetime.strptime(sandblasting_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),sandblasting = sandblasting_status,sandblasting_remarks = sandblasting_remarks,
                            insulation_date = datetime.strptime(insulation_application_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),insulation= insulation_application_status,insulation_remarks= insulation_application_remarks,
                            ut_rt_insulated_case_date = datetime.strptime(ut_rt_insulated_case_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),ut_rt_insulated_case = ut_rt_insulated_case_status,ut_rt_insulated_case_remarks= ut_rt_insulated_case_remarks,
                            acceptance_silver_material_date = datetime.strptime(sliver_acceptance_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),acceptance_silver_material= sliver_acceptance_status,acceptance_silver_material_remarks= sliver_acceptance_remarks,
                            silver_application_date = datetime.strptime(sliver_application_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),silver_application= sliver_application_status,silver_application_remarks= sliver_application_remarks,
                            formulation_tailoring_liner_propellant_date = datetime.strptime(formulation_tailoring_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),formulation_tailoring_liner_propellant= formulation_tailoring_status,formulation_tailoring_liner_propellant_remarks= formulation_tailoring_remarks,
                            conditioning_raw_materials_date = datetime.strptime(conditioning_raw_materials_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),conditioning_raw_materials = conditioning_raw_materials_status,conditioning_raw_materials_remarks = conditioning_raw_materials_remarks,
                            conditioning_of_lining_date = datetime.strptime(conditioning_of_lining_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),conditioning_of_lining= conditioning_of_lining_status,conditioning_of_lining_remarks = conditioning_of_lining_remarks,
                            casting_date = datetime.strptime(casting_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),casting= casting_status,casting_remarks = casting_remarks,
                            curing_date = datetime.strptime(curing_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),curing= curing_status,curing_remarks = curing_remarks,
                            ut_endoscopy_rt_grain_date = datetime.strptime(ut_endoscopy_rt_grain_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),ut_endoscopy_rt_grain = ut_endoscopy_rt_grain_status,ut_endoscopy_rt_grain_remarks = ut_endoscopy_rt_grain_remarks,
                            liner_mechanical_properties_date = datetime.strptime(mechanical_properties_liner_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),liner_mechanical_properties = mechanical_properties_liner_status,liner_mechanical_properties_remarks = mechanical_properties_liner_remarks,
                            mechanical_properties_propellant_date = datetime.strptime(mechanical_properties_propellant_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),mechanical_properties_propellant = mechanical_properties_propellant_status,mechanical_properties_propellant_remarks = mechanical_properties_propellant_remarks,
                            interface_bond_strength_date = datetime.strptime(interface_bond_strength_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),interface_bond_strength = interface_bond_strength_status,interface_bond_strength_remarks = interface_bond_strength_remarks,
                            propellant_burn_rate_date = datetime.strptime(propellant_burn_rate_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),propellant_burn_rate = propellant_burn_rate_status,propellant_burn_rate_remarks = propellant_burn_rate_remarks,
                            mass_liner_insulation_propellant_srm_date = datetime.strptime(mass_liner_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),mass_liner_insulation_propellant_srm = mass_liner_status,mass_liner_insulation_propellant_srm_remarks = mass_liner_remarks,
                            mass_insulation_date = datetime.strptime(mass_insulation_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),mass_insulation = mass_insulation_status,mass_insulation_remarks = mass_insulation_remarks,
                            mass_propellant_date = datetime.strptime(mass_propellant_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),mass_propellant = mass_propellant_status,mass_propellant_remarks = mass_propellant_remarks,
                            overall_qualification_date = datetime.strptime(overall_qualification_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),overall_qualification_status = overall_qualification_status,overall_qualification_remarks = overall_qualification_remarks,
                            bhd_date = datetime.strptime(bhd_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),bhd_status = bhd_status,bhd_remarks = bhd_remarks,qm_certification_date = datetime.strptime(qm_certification_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),qm_certification_status = qm_certification_status,qm_certification_remarks = qm_certification_remarks ,user_id= user_id)

                        data.append(motor_system)
                    ActiveMotors.objects.bulk_create(data)
                    os.remove(file_path)
                elif system_type == "Thermal Batteries":
                    for component_type,	organization,	battery_type,	tb_type,	battery_id,	lot_id,	matiral_qualified_date,	matiral_qualified_status,	matiral_qualified_remarks,\
                            compo_manufacturing_date,compo_manufacturing_status,	compo_manufacturing_remarks,	powerpack_assembly_date,	powerpack_assembly_status,	powerpack_assembly_remarks,\
                            powerpack_testing_date,	powerpack_testing_status,	powerpack_testing_remarks,	assembly_process_date,	assembly_process_status,	assembly_process_remarks,\
                            battery_testing_date,	battery_testing_status,	battery_testing_remarks,	final_qualification_date,	final_qualification_status,	final_qualification_remarks, *__ in csvf:

                        motor_system = ActiveMotors(component_type=component_type,organization=organization,battery_type=battery_type,tb_type=tb_type,battery_id= battery_id,lot_id = lot_id,
                            matiral_qualified_date= datetime.strptime(matiral_qualified_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),matiral_qualified_status=matiral_qualified_status,matiral_qualified_remarks=matiral_qualified_remarks,
                            compo_manufacturing_date= datetime.strptime(compo_manufacturing_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),compo_manufacturing_status=compo_manufacturing_status,compo_manufacturing_remarks=compo_manufacturing_remarks,
                            powerpack_assembly_date= datetime.strptime(powerpack_assembly_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),powerpack_assembly_status=powerpack_assembly_status,powerpack_assembly_remarks=powerpack_assembly_remarks,
                            powerpack_testing_date= datetime.strptime(powerpack_testing_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"), powerpack_testing_status=powerpack_testing_status,powerpack_testing_remarks=powerpack_testing_remarks,
                            assembly_process_date=  datetime.strptime(assembly_process_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),assembly_process_status=assembly_process_status,assembly_process_remarks=assembly_process_remarks,
                            battery_testing_date=  datetime.strptime(battery_testing_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"), battery_testing_status= battery_testing_status, battery_testing_remarks= battery_testing_remarks,
                            final_qualification_date=  datetime.strptime(final_qualification_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"), final_qualification_status= final_qualification_status, final_qualification_remarks=final_qualification_remarks,user_id= user_id)
                        data.append(motor_system)
                    ActiveMotors.objects.bulk_create(data)
                    os.remove(file_path)

                elif system_type =="Zinc Batteries":
                    for component_type,	organization,	battery_type,	battery_id,	lot_id,	raw_material_inspection_date,	raw_material_inspection_status,	raw_material_inspection_remarks,\
                            pressing_electrode_date,	pressing_electrode_status,	pressing_electrode_remarks,\
                            formation_process_date,	formation_process_status,	formation_process_remarks,	assembly_process_date,	assembly_process_status,	assembly_process_remarks,\
                            battery_testing_date,	battery_testing_status,	battery_testing_remarks,	final_qualification_date,	final_qualification_status,	final_qualification_remarks,*__ in csvf:


                        motor_system = ActiveMotors(component_type=component_type,organization=organization,battery_type=battery_type,battery_id= battery_id,lot_id = lot_id,
                            raw_material_inspection_date= datetime.strptime(raw_material_inspection_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),raw_material_inspection_status=raw_material_inspection_status,raw_material_inspection_remarks=raw_material_inspection_remarks,
                            pressing_electrode_date= datetime.strptime(pressing_electrode_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),pressing_electrode_status=pressing_electrode_status,pressing_electrode_remarks=pressing_electrode_remarks,
                            formation_process_date= datetime.strptime(formation_process_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),formation_process_status=formation_process_status,formation_process_remarks=formation_process_remarks,
                            assembly_process_date=  datetime.strptime(assembly_process_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),assembly_process_status=assembly_process_status,assembly_process_remarks=assembly_process_remarks,
                            battery_testing_date=  datetime.strptime(battery_testing_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"), battery_testing_status= battery_testing_status, battery_testing_remarks= battery_testing_remarks,
                            final_qualification_date=  datetime.strptime(final_qualification_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"), final_qualification_status= final_qualification_status, final_qualification_remarks=final_qualification_remarks,user_id= user_id)
                        data.append(motor_system)
                    ActiveMotors.objects.bulk_create(data)
                    os.remove(file_path)

                elif system_type == "Pyro Devices":
                    for component_type,	organization,	pyro_device_type,	pyro_device_id,	lot_id,	qualification_raw_material_date,	qualification_raw_material_status,	qualification_raw_material_remarks,	\
                        filling_date,	filling_status,	filling_remarks,	assembling_integration_date,	assembling_integration_status,	assembling_integration_remarks,\
                        qualification_testing_date,	qualification_testing_status,	qualification_testing_remarks,	\
                        performance_testing_date,performance_testing_status,	performance_testing_remarks, *__ in csvf:

                        motor_system = ActiveMotors(component_type=component_type,organization=organization,pd_type=pyro_device_type,pd_id=pyro_device_id,lot_id = lot_id,
                            qualification_raw_material_date= datetime.strptime(qualification_raw_material_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),qualification_raw_material_status=qualification_raw_material_status,qualification_raw_material_remarks=qualification_raw_material_remarks,
                            filling_date= datetime.strptime(filling_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),filling_status=filling_status,filling_remarks=filling_remarks,
                            assembling_integration_date= datetime.strptime(assembling_integration_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),assembling_integration_status=assembling_integration_status,assembling_integration_remarks=assembling_integration_remarks,
                            qualification_testing_date= datetime.strptime(qualification_testing_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"), qualification_testing_status=qualification_testing_status,qualification_testing_remarks=qualification_testing_remarks,
                            performance_testing_date=  datetime.strptime(performance_testing_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),performance_testing_status=performance_testing_status,performance_testing_remarks=performance_testing_remarks,user_id= user_id)
                        data.append(motor_system)
                    ActiveMotors.objects.bulk_create(data)
                    os.remove(file_path)

                return JsonResponse({'message': 'Production Status Updated Successfully!'}, status=200)

            else:
                return JsonResponse({'status': 'False', "message": "Status Not Saved"}, status=500)
        # except Exception as e:
        #     return JsonResponse({'status': 'False', "message": "Status Not Saved"}, status=500)



    @staticmethod
    def GetActiveMotorListDateNone(request, self=None):
            dataList = ActiveMotors.objects.all()
            for data in dataList:
                if data.qualification_insulation_lining_propellant_rm == 'None':
                    data.qualification_insulation_lining_propellant_rm_date = None

                if data.lining == 'None':
                    data.lining_date = None

                if data.propellant_mechanical_properties == 'None':
                    data.propellant_mechanical_properties_date = None

                if data.acceptance_casting == 'None':
                    data.acceptance_casting_date = None

                if data.sandblasting == 'None':
                    data.sandblasting_date = None

                if data.insulation == 'None':
                    data.insulation_date = None

                if data.ut_rt_insulated_case == 'None':
                    data.ut_rt_insulated_case_date = None

                if data.acceptance_silver_material == 'None':
                    data.acceptance_silver_material_date = None

                if data.silver_application == 'None':
                    data.silver_application_date = None

                if data.formulation_tailoring_liner_propellant == 'None':
                    data.formulation_tailoring_liner_propellant_date = None

                if data.conditioning_raw_materials == 'None':
                    data.conditioning_raw_materials_date = None

                if data.conditioning_of_lining == 'None':
                    data.conditioning_of_lining_date = None

                if data.casting == 'None':
                    data.casting_date = None

                if data.curing == 'None':
                    data.curing_date = None

                if  data.ut_endoscopy_rt_grain == 'None':
                    data.ut_endoscopy_rt_grain_date = None

                if data.liner_mechanical_properties == 'None':
                    data.liner_mechanical_properties_date = None

                if data.mechanical_properties_propellant == 'None':
                    data.mechanical_properties_propellant_date = None

                if data.interface_bond_strength == 'None':
                    data.interface_bond_strength_date = None

                if data.propellant_burn_rate == 'None':
                    data.propellant_burn_rate_date = None

                if data.trimming_Propellant_grain == 'None':
                    data.trimming_Propellant_grain_date = None

                if data.mass_liner_insulation_propellant_srm == 'None':
                    data.mass_liner_insulation_propellant_srm_date = None

                if data.mass_insulation == 'None':
                    data.mass_insulation_date = None

                if data.mass_propellant == 'None':
                    data.mass_propellant_date = None

                if data.bhd_status == 'None' :
                    data.bhd_date = None

                if data.qm_certification_status == 'None':
                    data.qm_certification_date = None

                if data.overall_qualification_status == 'None':
                    data.overall_qualification_date = None

                if data.matiral_qualified_status == 'None':
                    data.matiral_qualified_date = None

                if data.compo_manufacturing_status == 'None':
                    data.compo_manufacturing_date = None

                if data.powerpack_assembly_status == 'None':
                    data.powerpack_assembly_date = None

                if data.powerpack_testing_status == 'None':
                    data.powerpack_testing_date = None

                if data.raw_material_inspection_status == 'None':
                    data.raw_material_inspection_date = None

                if data.pressing_electrode_status == 'None':
                    data.pressing_electrode_date = None

                if data.formation_process_status == 'None':
                    data.formation_process_date = None

                if data.assembly_process_status == 'None':
                    data.assembly_process_date = None

                if data.battery_testing_status == 'None':
                    data.battery_testing_date = None

                if data.final_qualification_status == 'None':
                    data.final_qualification_date = None

                if data.qualification_raw_material_status == 'None':
                    data.qualification_raw_material_date = None

                if data.filling_status == 'None':
                    data.filling_date = None

                if data.assembling_integration_status == 'None':
                    data.assembling_integration_date = None

                if data.qualification_testing_status == 'None':
                    data.qualification_testing_date = None

                if data.performance_testing_status == 'None':
                    data.performance_testing_date = None
                data.save()
            data = ActiveMotors.objects.all()
            serializer = ActiveMotorSerializer(data, many=True)
            return JsonResponse({'status': 'True', 'data': serializer.data}, status=200)


    @staticmethod
    def GetActiveMotorHistoryListDateNone(request, self=None):
            dataList = ActiveMotorsHistory.objects.all()
            for data in dataList:
                if data.qualification_insulation_lining_propellant_rm == 'None':
                    data.qualification_insulation_lining_propellant_rm_date = None

                if data.lining == 'None':
                    data.lining_date = None

                if data.propellant_mechanical_properties == 'None':
                    data.propellant_mechanical_properties_date = None

                if data.acceptance_casting == 'None':
                    data.acceptance_casting_date = None

                if data.sandblasting == 'None':
                    data.sandblasting_date = None

                if data.insulation == 'None':
                    data.insulation_date = None

                if data.ut_rt_insulated_case == 'None':
                    data.ut_rt_insulated_case_date = None

                if data.acceptance_silver_material == 'None':
                    data.acceptance_silver_material_date = None

                if data.silver_application == 'None':
                    data.silver_application_date = None

                if data.formulation_tailoring_liner_propellant == 'None':
                    data.formulation_tailoring_liner_propellant_date = None

                if data.conditioning_raw_materials == 'None':
                    data.conditioning_raw_materials_date = None

                if data.conditioning_of_lining == 'None':
                    data.conditioning_of_lining_date = None

                if data.casting == 'None':
                    data.casting_date = None

                if data.curing == 'None':
                    data.curing_date = None

                if  data.ut_endoscopy_rt_grain == 'None':
                    data.ut_endoscopy_rt_grain_date = None

                if data.liner_mechanical_properties == 'None':
                    data.liner_mechanical_properties_date = None

                if data.mechanical_properties_propellant == 'None':
                    data.mechanical_properties_propellant_date = None

                if data.interface_bond_strength == 'None':
                    data.interface_bond_strength_date = None

                if data.propellant_burn_rate == 'None':
                    data.propellant_burn_rate_date = None

                if data.trimming_Propellant_grain == 'None':
                    data.trimming_Propellant_grain_date = None

                if data.mass_liner_insulation_propellant_srm == 'None':
                    data.mass_liner_insulation_propellant_srm_date = None

                if data.mass_insulation == 'None':
                    data.mass_insulation_date = None

                if data.mass_propellant == 'None':
                    data.mass_propellant_date = None

                if data.bhd_status == 'None' :
                    data.bhd_date = None

                if data.qm_certification_status == 'None':
                    data.qm_certification_date = None

                if data.overall_qualification_status == 'None':
                    data.overall_qualification_date = None

                if data.matiral_qualified_status == 'None':
                    data.matiral_qualified_date = None

                if data.compo_manufacturing_status == 'None':
                    data.compo_manufacturing_date = None

                if data.powerpack_assembly_status == 'None':
                    data.powerpack_assembly_date = None

                if data.powerpack_testing_status == 'None':
                    data.powerpack_testing_date = None

                if data.raw_material_inspection_status == 'None':
                    data.raw_material_inspection_date = None

                if data.pressing_electrode_status == 'None':
                    data.pressing_electrode_date = None

                if data.formation_process_status == 'None':
                    data.formation_process_date = None

                if data.assembly_process_status == 'None':
                    data.assembly_process_date = None

                if data.battery_testing_status == 'None':
                    data.battery_testing_date = None

                if data.final_qualification_status == 'None':
                    data.final_qualification_date = None

                if data.qualification_raw_material_status == 'None':
                    data.qualification_raw_material_date = None

                if data.filling_status == 'None':
                    data.filling_date = None

                if data.assembling_integration_status == 'None':
                    data.assembling_integration_date = None

                if data.qualification_testing_status == 'None':
                    data.qualification_testing_date = None

                if data.performance_testing_status == 'None':
                    data.performance_testing_date = None
                data.save()
            data = ActiveMotorsHistory.objects.all()
            serializer = ActiveMotorSerializer(data, many=True)
            return JsonResponse({'status': 'True', 'data': serializer.data}, status=200)
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
            pd_type = request.query_params.get('selected_pdType')


            ParentStatus = request.query_params['parent_status']
            ChildStatus = request.query_params['child_status']


            today = date.today()
            # dataList = []

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
            srm_objects = Q()
            # if current_year !='':
            #     filter_objects &= get_filter(
            #         'creation_date__year', 'equal',
            #         current_year)

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
            if pd_type !='':
                filter_objects &= get_filter(
                    'pd_type', 'equal',
                    pd_type)

            # create dynamic filter
            if current_comp != '':
                srm_objects &= get_filter(
                    'component_type', 'equal',
                    current_comp)
            if system_type != '':
                srm_objects &= get_filter(
                    'system_type', 'equal',
                    system_type)
            if current_sys != '':
                srm_objects &= get_filter(
                    'system_name', 'equal',
                    current_sys)
            if pd_type != '':
                srm_objects &= get_filter(
                    'pd_type', 'equal',
                    pd_type)
            if current_org !='':
                srm_objects &= get_filter(
                    'organization', 'equal', current_org)
            if current_bat !='':
                srm_objects &= get_filter(
                    'battery_type', 'equal', current_bat)
            if current_lot !='':
                srm_objects &= get_filter(
                    'lot_id', 'equal', current_lot)
            if current_comp == 'SRMs':
                srm_objects &= get_filter(
                    'motor_id', 'not_equal',
                    '')
            if ParentStatus =='Motor' and ChildStatus == 'Total System':
                srm_total_count = ActiveMotors.objects.filter(srm_objects)
                serializer = ActiveMotorSerializer(srm_total_count, many=True)
                return JsonResponse({'status': 'True', 'data': serializer.data},
                                    status=200)
            dataList = ActiveMotors.objects.filter(filter_objects)
            if ParentStatus != '':
                ListItems = []
                if ParentStatus == 'Qualification of raw material of insulation':
                    dataList = dataList.extra(
                        select={'year': 'extract (year from qualification_insulation_lining_propellant_rm_date)',
                                'month': 'extract (month from qualification_insulation_lining_propellant_rm_date)',
                                'day': 'extract (day from qualification_insulation_lining_propellant_rm_date)'},
                        order_by=['month', 'day', '-year'])
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(qualification_insulation_lining_propellant_rm_date__year = current_year,
                                                    qualification_insulation_lining_propellant_rm=ChildStatus)

                    if ChildStatus == 'Current Count':
                        for data in dataList:
                            if data.qualification_insulation_lining_propellant_rm_date.strftime("%Y") == current_year and (data.qualification_insulation_lining_propellant_rm == 'Under process' or data.qualification_insulation_lining_propellant_rm == 'Observation(same stage)' or data.qualification_insulation_lining_propellant_rm == 'Halt'):
                                ListItems.append(data)
                elif ParentStatus == 'Lining':
                    dataList = dataList.extra(
                        select={'year': 'extract (year from lining_date)',
                                'month': 'extract (month from lining_date)',
                                'day': 'extract (day from lining_date)'},
                        order_by=['month', 'day', '-year'])
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(lining_date__year = current_year, lining=ChildStatus)
                    if ChildStatus == 'Current Count':
                        for data in dataList:
                            if data.lining_date.strftime("%Y") == current_year and (data.lining == 'Under process' or data.lining == 'Observation(same stage)' or data.lining == 'Halt'):
                                ListItems.append(data)
                elif ParentStatus == 'Qualification of Propellant':
                    dataList = dataList.extra(
                        select={'year': 'extract (year from propellant_mechanical_properties_date)',
                                'month': 'extract (month from propellant_mechanical_properties_date)',
                                'day': 'extract (day from propellant_mechanical_properties_date)'},
                        order_by=['month', 'day', '-year'])
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(propellant_mechanical_properties_date__year = current_year, propellant_mechanical_properties=ChildStatus)
                    if ChildStatus == 'Current Count':
                        for data in dataList:
                            if data.propellant_mechanical_properties_date.strftime("%Y") == current_year and (data.propellant_mechanical_properties == 'Under process' or data.propellant_mechanical_properties == 'Observation(same stage)' or data.propellant_mechanical_properties == 'Halt'):
                                ListItems.append(data)
                elif ParentStatus == 'Acceptance of casting':
                    dataList = dataList.extra(
                        select={'year': 'extract (year from acceptance_casting_date)',
                                'month': 'extract (month from acceptance_casting_date)',
                                'day': 'extract (day from acceptance_casting_date)'},
                        order_by=['month', 'day', '-year'])
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(acceptance_casting_date__year = current_year, acceptance_casting=ChildStatus)
                    if ChildStatus == 'Current Count':
                        for data in dataList:
                            if data.acceptance_casting_date.strftime("%Y") == current_year and (data.acceptance_casting == 'Under process' or data.acceptance_casting == 'Observation(same stage)' or data.acceptance_casting == 'Halt'):
                                ListItems.append(data)
                elif ParentStatus == 'Sandblasting':
                    dataList = dataList.extra(
                        select={'year': 'extract (year from sandblasting_date)',
                                'month': 'extract (month from sandblasting_date)',
                                'day': 'extract (day from sandblasting_date)'},
                        order_by=['month', 'day', '-year'])
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(sandblasting_date__year = current_year, sandblasting=ChildStatus)
                    if ChildStatus == 'Current Count':
                        for data in dataList:
                            if data.sandblasting_date.strftime("%Y") == current_year and (data.sandblasting == 'Under process' or data.sandblasting == 'Observation(same stage)' or data.sandblasting == 'Halt'):
                                ListItems.append(data)
                elif ParentStatus == 'Insulation':
                    dataList = dataList.extra(
                        select={'year': 'extract (year from insulation_date)',
                                'month': 'extract (month from insulation_date)',
                                'day': 'extract (day from insulation_date)'},
                        order_by=['month', 'day', '-year'])
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(insulation_date__year = current_year, insulation=ChildStatus)
                    if ChildStatus == 'Current Count':
                        for data in dataList:
                            if data.insulation_date.strftime("%Y") == current_year and (data.insulation == 'Under process' or data.insulation == 'Observation(same stage)' or data.insulation == 'Halt'):
                                ListItems.append(data)
                elif ParentStatus == 'UT and RT of Insulated Case':
                    dataList = dataList.extra(
                        select={'year': 'extract (year from ut_rt_insulated_case_date)',
                                'month': 'extract (month from ut_rt_insulated_case_date)',
                                'day': 'extract (day from ut_rt_insulated_case_date)'},
                        order_by=['month', 'day', '-year'])
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(ut_rt_insulated_case_date__year = current_year, ut_rt_insulated_case=ChildStatus)
                    if ChildStatus == 'Current Count':
                        for data in dataList:
                            if data.ut_rt_insulated_case_date.strftime("%Y") == current_year and (data.ut_rt_insulated_case == 'Under process' or data.ut_rt_insulated_case == 'Observation(same stage)' or data.ut_rt_insulated_case == 'Halt'):
                                ListItems.append(data)
                elif ParentStatus == 'Acceptance of Sliver Material':
                    dataList = dataList.extra(
                        select={'year': 'extract (year from acceptance_silver_material_date)',
                                'month': 'extract (month from acceptance_silver_material_date)',
                                'day': 'extract (day from acceptance_silver_material_date)'},
                        order_by=['month', 'day', '-year'])
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(acceptance_silver_material_date__year = current_year, acceptance_silver_material=ChildStatus)
                    if ChildStatus == 'Current Count':
                        for data in dataList:
                            if data.acceptance_silver_material_date.strftime("%Y") == current_year and (data.acceptance_silver_material == 'Under process' or data.acceptance_silver_material == 'Observation(same stage)' or data.acceptance_silver_material == 'Halt'):
                                ListItems.append(data)
                elif ParentStatus == 'Silver Application':
                    dataList = dataList.extra(
                        select={'year': 'extract (year from silver_application_date)',
                                'month': 'extract (month from silver_application_date)',
                                'day': 'extract (day from silver_application_date)'},
                        order_by=['month', 'day', '-year'])
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(silver_application_date__year = current_year, silver_application=ChildStatus)
                    if ChildStatus == 'Current Count':
                        for data in dataList:
                            if data.silver_application_date.strftime("%Y") == current_year and (data.silver_application == 'Under process' or data.silver_application == 'Observation(same stage)' or data.silver_application == 'Halt'):
                                ListItems.append(data)
                elif ParentStatus == 'Formulation tailoring':
                    dataList = dataList.extra(
                        select={'year': 'extract (year from formulation_tailoring_liner_propellant_date)',
                                'month': 'extract (month from formulation_tailoring_liner_propellant_date)',
                                'day': 'extract (day from formulation_tailoring_liner_propellant_date)'},
                        order_by=['month', 'day', '-year'])
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(formulation_tailoring_liner_propellant_date__year = current_year, formulation_tailoring_liner_propellant=ChildStatus)
                    if ChildStatus == 'Current Count':
                        for data in dataList:
                            if data.formulation_tailoring_liner_propellant_date.strftime("%Y") == current_year and (data.formulation_tailoring_liner_propellant == 'Under process' or data.formulation_tailoring_liner_propellant == 'Observation(same stage)' or data.formulation_tailoring_liner_propellant == 'Halt'):
                                ListItems.append(data)
                elif ParentStatus == 'Conditioning of Raw Material':
                    dataList = dataList.extra(
                        select={'year': 'extract (year from conditioning_raw_materials_date)',
                                'month': 'extract (month from conditioning_raw_materials_date)',
                                'day': 'extract (day from conditioning_raw_materials_date)'},
                        order_by=['month', 'day', '-year'])
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(conditioning_raw_materials_date__year = current_year, conditioning_raw_materials=ChildStatus)
                    if ChildStatus == 'Current Count':
                        for data in dataList:
                            if data.conditioning_raw_materials_date.strftime("%Y") == current_year and (data.conditioning_raw_materials == 'Under process' or data.conditioning_raw_materials == 'Observation(same stage)' or data.conditioning_raw_materials == 'Halt'):
                                ListItems.append(data)
                elif ParentStatus == 'Conditioning of Lining':
                    dataList = dataList.extra(
                        select={'year': 'extract (year from conditioning_of_lining_date)',
                                'month': 'extract (month from conditioning_of_lining_date)',
                                'day': 'extract (day from conditioning_of_lining_date)'},
                        order_by=['month', 'day', '-year'])
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(conditioning_of_lining_date__year = current_year,conditioning_of_lining=ChildStatus)
                    if ChildStatus == 'Current Count':
                        for data in dataList:
                            if data.conditioning_of_lining_date.strftime("%Y") == current_year and  (data.conditioning_of_lining == 'Under process' or data.conditioning_of_lining == 'Observation(same stage)' or data.conditioning_of_lining == 'Halt'):
                                ListItems.append(data)
                elif ParentStatus == 'Casing':
                    dataList = dataList.extra(
                        select={'year': 'extract (year from casting_date)',
                                'month': 'extract (month from casting_date)',
                                'day': 'extract (day from casting_date)'},
                        order_by=['month', 'day', '-year'])
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(casting_date__year = current_year, casting=ChildStatus)
                    if ChildStatus == 'Current Count':
                        for data in dataList:
                            if data.casting_date.strftime("%Y") == current_year and (data.casting == 'Under process' or data.casting == 'Observation(same stage)' or data.casting == 'Halt'):
                                ListItems.append(data)
                elif ParentStatus == 'Curing':
                    dataList = dataList.extra(
                        select={'year': 'extract (year from curing_date)',
                                'month': 'extract (month from curing_date)',
                                'day': 'extract (day from curing_date)'},
                        order_by=['month', 'day', '-year'])
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(curing_date__year = current_year, curing=ChildStatus)
                    if ChildStatus == 'Current Count':
                        for data in dataList:
                            if data.curing_date.strftime("%Y") == current_year and (data.curing == 'Under process' or data.curing == 'Observation(same stage)' or data.curing == 'Halt'):
                                ListItems.append(data)
                elif ParentStatus == 'UT, endoscopy and RT of grain':
                    dataList = dataList.extra(
                        select={'year': 'extract (year from ut_endoscopy_rt_grain_date)',
                                'month': 'extract (month from ut_endoscopy_rt_grain_date)',
                                'day': 'extract (day from ut_endoscopy_rt_grain_date)'},
                        order_by=['month', 'day', '-year'])
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(ut_endoscopy_rt_grain_date__year = current_year,ut_endoscopy_rt_grain=ChildStatus)
                    if ChildStatus == 'Current Count':
                        for data in dataList:
                            if data.ut_endoscopy_rt_grain_date.strftime("%Y") == current_year and (data.ut_endoscopy_rt_grain == 'Under process' or data.ut_endoscopy_rt_grain == 'Observation(same stage)' or data.ut_endoscopy_rt_grain == 'Halt'):
                                ListItems.append(data)

                elif ParentStatus == 'Mechanical Properties of Liner':
                    dataList = dataList.extra(
                        select={'year': 'extract (year from liner_mechanical_properties_date)',
                                'month': 'extract (month from liner_mechanical_properties_date)',
                                'day': 'extract (day from liner_mechanical_properties_date)'},
                        order_by=['month', 'day', '-year'])
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(liner_mechanical_properties_date__year = current_year, liner_mechanical_properties=ChildStatus)
                    if ChildStatus == 'Current Count':
                        for data in dataList:
                            if data.liner_mechanical_properties_date.strftime("%Y") == current_year and (data.liner_mechanical_properties == 'Under process' or data.liner_mechanical_properties == 'Observation(same stage)' or data.liner_mechanical_properties == 'Halt'):
                                ListItems.append(data)
                elif ParentStatus == 'Mechanical Properties of Propellant':
                    dataList = dataList.extra(
                        select={'year': 'extract (year from mechanical_properties_propellant_date)',
                                'month': 'extract (month from mechanical_properties_propellant_date)',
                                'day': 'extract (day from mechanical_properties_propellant_date)'},
                        order_by=['month', 'day', '-year'])
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(mechanical_properties_propellant_date__year = current_year, mechanical_properties_propellant=ChildStatus)
                    if ChildStatus == 'Current Count':
                        for data in dataList:
                            if data.mechanical_properties_propellant_date.strftime("%Y") == current_year and (data.mechanical_properties_propellant == 'Under process' or data.mechanical_properties_propellant == 'Observation(same stage)' or data.mechanical_properties_propellant == 'Halt'):
                                ListItems.append(data)
                elif ParentStatus == 'Interface bond strength':
                    dataList = dataList.extra(
                        select={'year': 'extract (year from interface_bond_strength_date)',
                                'month': 'extract (month from interface_bond_strength_date)',
                                'day': 'extract (day from interface_bond_strength_date)'},
                        order_by=['month', 'day', '-year'])
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(interface_bond_strength_date__year = current_year, interface_bond_strength=ChildStatus)
                    if ChildStatus == 'Current Count':
                        for data in dataList:
                            if data.interface_bond_strength_date.strftime("%Y") == current_year and (data.interface_bond_strength == 'Under process' or data.interface_bond_strength == 'Observation(same stage)' or data.interface_bond_strength == 'Halt'):
                                ListItems.append(data)
                elif ParentStatus == 'Propellant burn rate':
                    dataList = dataList.extra(
                        select={'year': 'extract (year from propellant_burn_rate_date)',
                                'month': 'extract (month from propellant_burn_rate_date)',
                                'day': 'extract (day from propellant_burn_rate_date)'},
                        order_by=['month', 'day', '-year'])
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(propellant_burn_rate_date__year = current_year, propellant_burn_rate=ChildStatus)
                    if ChildStatus == 'Current Count':
                        for data in dataList:
                            if data.propellant_burn_rate_date.strftime("%Y") == current_year and (data.propellant_burn_rate == 'Under process' or data.propellant_burn_rate == 'Observation(same stage)' or data.propellant_burn_rate == 'Halt'):
                                ListItems.append(data)
                elif ParentStatus == 'Trimming of propellant grain':
                    dataList = dataList.extra(
                        select={'year': 'extract (year from trimming_Propellant_grain_date)',
                                'month': 'extract (month from trimming_Propellant_grain_date)',
                                'day': 'extract (day from trimming_Propellant_grain_date)'},
                        order_by=['month', 'day', '-year'])
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(trimming_Propellant_grain_date__year = current_year, trimming_Propellant_grain=ChildStatus)
                    if ChildStatus == 'Current Count':
                        for data in dataList:
                            if data.trimming_Propellant_grain_date.strftime("%Y") == current_year and (data.trimming_Propellant_grain == 'Under process' or data.trimming_Propellant_grain == 'Observation(same stage)' or data.trimming_Propellant_grain == 'Halt'):
                                ListItems.append(data)
                elif ParentStatus == 'Mass of liner':
                    dataList = dataList.extra(
                        select={'year': 'extract (year from mass_liner_insulation_propellant_srm_date)',
                                'month': 'extract (month from mass_liner_insulation_propellant_srm_date)',
                                'day': 'extract (day from mass_liner_insulation_propellant_srm_date)'},
                        order_by=['month', 'day', '-year'])
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(mass_liner_insulation_propellant_srm_date__year = current_year, mass_liner_insulation_propellant_srm=ChildStatus)
                    if ChildStatus == 'Current Count':
                        for data in dataList:
                            if data.mass_liner_insulation_propellant_srm_date.strftime("%Y") == current_year and (data.mass_liner_insulation_propellant_srm == 'Under process' or data.mass_liner_insulation_propellant_srm == 'Observation(same stage)' or data.mass_liner_insulation_propellant_srm == 'Halt'):
                                ListItems.append(data)
                elif ParentStatus == 'Mass of Insulation':
                    dataList = dataList.extra(
                        select={'year': 'extract (year from mass_insulation_date)',
                                'month': 'extract (month from mass_insulation_date)',
                                'day': 'extract (day from mass_insulation_date)'},
                        order_by=['month', 'day', '-year'])
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(mass_insulation_date__year = current_year, mass_insulation=ChildStatus)
                    if ChildStatus == 'Current Count':
                        for data in dataList:
                            if data.mass_insulation_date.strftime("%Y") == current_year and (data.mass_insulation == 'Under process' or data.mass_insulation == 'Observation(same stage)' or data.mass_insulation == 'Halt'):
                                ListItems.append(data)
                elif ParentStatus == 'Mass of Propellant':
                    dataList = dataList.extra(
                        select={'year': 'extract (year from mass_propellant_date)',
                                'month': 'extract (month from mass_propellant_date)',
                                'day': 'extract (day from mass_propellant_date)'},
                        order_by=['month', 'day', '-year'])
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(mass_propellant_date__year = current_year, mass_propellant=ChildStatus)
                    if ChildStatus == 'Current Count':
                        for data in dataList:
                            if data.mass_propellant_date.strftime("%Y") == current_year and (data.mass_propellant == 'Under process' or data.mass_propellant == 'Observation(same stage)' or data.mass_propellant == 'Halt'):
                                ListItems.append(data)
                elif ParentStatus == 'BHD Status':
                        ListItems = dataList.filter(bhd_date__year = current_year, bhd_status=ChildStatus).extra(
                        select={'year': 'extract (year from bhd_date)',
                                'month': 'extract (month from bhd_date)',
                                'day': 'extract (day from bhd_date)'},
                        order_by=['month', 'day', '-year'])

                elif ParentStatus == 'QM Certification Status':
                        ListItems = dataList.filter(qm_certification_date__year = current_year, qm_certification_status=ChildStatus).extra(
                        select={'year': 'extract (year from qm_certification_date)',
                                'month': 'extract (month from qm_certification_date)',
                                'day': 'extract (day from qm_certification_date)'},
                        order_by=['month', 'day', '-year'])

                elif ParentStatus == 'Overall Qualification':
                        ListItems = dataList.filter(overall_qualification_date__year = current_year,overall_qualification_status=ChildStatus).extra(
                        select={'year': 'extract (year from overall_qualification_date)',
                                'month': 'extract (month from overall_qualification_date)',
                                'day': 'extract (day from overall_qualification_date)'},
                        order_by=['month', 'day', '-year'])


                elif ParentStatus == 'Material Qualified':
                    dataList = dataList.extra(
                        select={'year': 'extract (year from matiral_qualified_date)',
                                'month': 'extract (month from matiral_qualified_date)',
                                'day': 'extract (day from matiral_qualified_date)'},
                        order_by=['month', 'day', '-year'])
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(matiral_qualified_date__year = current_year,matiral_qualified_status=ChildStatus)
                    if ChildStatus == 'Current Count':
                        for data in dataList:
                            if data.matiral_qualified_date.strftime("%Y") == current_year and (data.matiral_qualified_status == 'Under process' or data.matiral_qualified_status == 'Observation(same stage)' or data.matiral_qualified_status == 'Halt'):
                                ListItems.append(data)
                elif ParentStatus == 'Component Manufacturing':
                    dataList = dataList.extra(
                        select={'year': 'extract (year from compo_manufacturing_date)',
                                'month': 'extract (month from compo_manufacturing_date)',
                                'day': 'extract (day from compo_manufacturing_date)'},
                        order_by=['month', 'day', '-year'])
                    if ChildStatus != 'Current Count':
                        # filter_objects &= get_filter('blt_status', 'equal', ChildStatus)
                        ListItems = dataList.filter(compo_manufacturing_date__year = current_year, compo_manufacturing_status=ChildStatus)
                    if ChildStatus == 'Current Count':
                        for data in dataList:
                            if data.compo_manufacturing_date.strftime("%Y") == current_year and (data.compo_manufacturing_status == 'Under process' or data.compo_manufacturing_status == 'Observation(same stage)' or data.compo_manufacturing_status == 'Halt'):
                                ListItems.append(data)
                elif ParentStatus == 'Power Pack Assembly':
                    dataList = dataList.extra(
                        select={'year': 'extract (year from powerpack_assembly_date)',
                                'month': 'extract (month from powerpack_assembly_date)',
                                'day': 'extract (day from powerpack_assembly_date)'},
                        order_by=['month', 'day', '-year'])
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(powerpack_assembly_date__year = current_year, powerpack_assembly_status=ChildStatus)
                    if ChildStatus == 'Current Count':
                        for data in dataList:
                            if data.powerpack_assembly_date.strftime("%Y") == current_year and (data.powerpack_assembly_status == 'Under process' or data.powerpack_assembly_status == 'Observation(same stage)' or data.powerpack_assembly_status == 'Halt'):
                                ListItems.append(data)
                elif ParentStatus == 'Power Pack Testing':
                    dataList = dataList.extra(
                        select={'year': 'extract (year from powerpack_testing_date)',
                                'month': 'extract (month from powerpack_testing_date)',
                                'day': 'extract (day from powerpack_testing_date)'},
                        order_by=['month', 'day', '-year'])
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(powerpack_testing_status=ChildStatus)
                    if ChildStatus == 'Current Count':
                        for data in dataList:
                            if data.powerpack_testing_date.strftime("%Y") == current_year and (data.powerpack_testing_status == 'Under process' or data.powerpack_testing_status == 'Observation(same stage)' or data.powerpack_testing_status == 'Halt'):
                                ListItems.append(data)
                elif ParentStatus == 'Raw Material Inspection':
                    dataList = dataList.extra(
                        select={'year': 'extract (year from raw_material_inspection_date)',
                                'month': 'extract (month from raw_material_inspection_date)',
                                'day': 'extract (day from raw_material_inspection_date)'},
                        order_by=['month', 'day', '-year'])
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(raw_material_inspection_date__year = current_year, raw_material_inspection_status=ChildStatus)
                    if ChildStatus == 'Current Count':
                        for data in dataList:
                            if data.raw_material_inspection_date.strftime("%Y") == current_year and (data.raw_material_inspection_status == 'Under process' or data.raw_material_inspection_status == 'Observation(same stage)' or data.raw_material_inspection_status == 'Halt'):
                                ListItems.append(data)
                elif ParentStatus == 'Pressing Electrode':
                    dataList = dataList.extra(
                        select={'year': 'extract (year from pressing_electrode_date)',
                                'month': 'extract (month from pressing_electrode_date)',
                                'day': 'extract (day from pressing_electrode_date)'},
                        order_by=['month', 'day', '-year'])
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(pressing_electrode_date__year = current_year, pressing_electrode_status=ChildStatus)
                    if ChildStatus == 'Current Count':
                        for data in dataList:
                            if data.pressing_electrode_date.strftime("%Y") == current_year and (data.pressing_electrode_status == 'Under process' or data.pressing_electrode_status == 'Observation(same stage)' or data.pressing_electrode_status == 'Halt'):
                                ListItems.append(data)
                elif ParentStatus == 'Formation Process':
                    dataList = dataList.extra(
                        select={'year': 'extract (year from formation_process_date)',
                                'month': 'extract (month from formation_process_date)',
                                'day': 'extract (day from formation_process_date)'},
                        order_by=['month', 'day', '-year'])
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(formation_process_date__year = current_year, formation_process_status=ChildStatus)
                    if ChildStatus == 'Current Count':
                        for data in dataList:
                            if data.formation_process_date.strftime("%Y") == current_year and (data.formation_process_status == 'Under process' or data.formation_process_status == 'Observation(same stage)' or data.formation_process_status == 'Halt'):
                                ListItems.append(data)
                elif ParentStatus == 'Assembly Process':
                    dataList = dataList.extra(
                        select={'year': 'extract (year from assembly_process_date)',
                                'month': 'extract (month from assembly_process_date)',
                                'day': 'extract (day from assembly_process_date)'},
                        order_by=['month', 'day', '-year'])
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(assembly_process_date__year = current_year, assembly_process_status=ChildStatus)
                    if ChildStatus == 'Current Count':
                        for data in dataList:
                            if data.assembly_process_date.strftime("%Y") == current_year and (data.assembly_process_status == 'Under process' or data.assembly_process_status == 'Observation(same stage)' or data.assembly_process_status == 'Halt'):
                                ListItems.append(data)
                elif ParentStatus == 'Battery Testing':
                    dataList = dataList.extra(
                        select={'year': 'extract (year from battery_testing_date)',
                                'month': 'extract (month from battery_testing_date)',
                                'day': 'extract (day from battery_testing_date)'},
                        order_by=['month', 'day', '-year'])
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(battery_testing_date__year = current_year, battery_testing_status=ChildStatus)
                    if ChildStatus == 'Current Count':
                        for data in dataList:
                            if data.battery_testing_date.strftime("%Y") == current_year and (data.battery_testing_status == 'Under process' or data.battery_testing_status == 'Observation(same stage)' or data.battery_testing_status == 'Halt'):
                                ListItems.append(data)
                elif ParentStatus == 'Final Qualification':
                    dataList = dataList.extra(
                        select={'year': 'extract (year from final_qualification_date)',
                                'month': 'extract (month from final_qualification_date)',
                                'day': 'extract (day from final_qualification_date)'},
                        order_by=['month', 'day', '-year'])
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(final_qualification_date__year = current_year, final_qualification_status=ChildStatus)
                    if ChildStatus == 'Current Count':
                        for data in dataList:
                            if data.final_qualification_date.strftime("%Y") == current_year and (data.final_qualification_status == 'Under process' or data.final_qualification_status == 'Observation(same stage)' or data.final_qualification_status == 'Halt'):
                                ListItems.append(data)

                elif ParentStatus == 'Qualification of Raw Material':
                    dataList = dataList.extra(
                        select={'year': 'extract (year from qualification_raw_material_date)',
                                'month': 'extract (month from qualification_raw_material_date)',
                                'day': 'extract (day from qualification_raw_material_date)'},
                        order_by=['month', 'day', '-year'])
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(qualification_raw_material_date__year =  current_year, qualification_raw_material_status=ChildStatus)
                    if ChildStatus == 'Current Count':
                        for data in dataList:
                            if data.qualification_raw_material_date.strftime("%Y") ==  current_year and (data.qualification_raw_material_status == 'Under process' or data.qualification_raw_material_status == 'Observation(same stage)' or data.qualification_raw_material_status == 'Halt'):
                                ListItems.append(data)
                elif ParentStatus == 'Filling':
                    dataList = dataList.extra(
                        select={'year': 'extract (year from filling_date)',
                                'month': 'extract (month from filling_date)',
                                'day': 'extract (day from filling_date)'},
                        order_by=['month', 'day', '-year'])
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(filling_date__year = current_year, filling_status=ChildStatus)
                    if ChildStatus == 'Current Count':
                        for data in dataList:
                            if data.filling_date.strftime("%Y")== current_year and (data.filling_status == 'Under process' or data.filling_status == 'Observation(same stage)' or data.filling_status == 'Halt'):
                                ListItems.append(data)
                elif ParentStatus == 'Assembling/Integration':
                    dataList = dataList.extra(
                        select={'year': 'extract (year from assembling_integration_date)',
                                'month': 'extract (month from assembling_integration_date)',
                                'day': 'extract (day from assembling_integration_date)'},
                        order_by=['month', 'day', '-year'])
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(assembling_integration_date__year = current_year, assembling_integration_status=ChildStatus)
                    if ChildStatus == 'Current Count':
                        for data in dataList:
                            if data.assembling_integration_date.strftime("%Y") == current_year and (data.assembling_integration_status == 'Under process' or data.assembling_integration_status == 'Observation(same stage)' or data.assembling_integration_status == 'Halt'):
                                ListItems.append(data)
                elif ParentStatus == 'Qualification Testing':
                    dataList = dataList.extra(
                        select={'year': 'extract (year from qualification_testing_date)',
                                'month': 'extract (month from qualification_testing_date)',
                                'day': 'extract (day from qualification_testing_date)'},
                        order_by=['month', 'day', '-year'])
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(qualification_testing_date__year = current_year, qualification_testing_status=ChildStatus)
                    if ChildStatus == 'Current Count':
                        for data in dataList:
                            if data.qualification_testing_date.strftime("%Y") == current_year and (data.qualification_testing_status == 'Under process' or data.qualification_testing_status == 'Observation(same stage)' or data.qualification_testing_status == 'Halt'):
                                ListItems.append(data)
                elif ParentStatus == 'Performance Testing':
                    dataList = dataList.extra(
                        select={'year': 'extract (year from performance_testing_date)',
                                'month': 'extract (month from performance_testing_date)',
                                'day': 'extract (day from performance_testing_date)'},
                        order_by=['month', 'day', '-year'])
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(performance_testing_date__year = current_year, performance_testing_status=ChildStatus)
                    if ChildStatus == 'Current Count':
                        for data in dataList:
                            if data.performance_testing_date.strftime("%Y") == current_year and (data.performance_testing_status == 'Under process' or data.performance_testing_status == 'Observation(same stage)' or data.performance_testing_status == 'Halt'):
                                ListItems.append(data)
                elif ParentStatus == 'NCR':
                        ListItems = dataList.filter(ncr_status=ChildStatus)

                elif ParentStatus == 'Overall':
                        ListItems = dataList.filter(overall_status=ChildStatus)

                serializer = ActiveMotorSerializer(ListItems, many=True)
                return JsonResponse({'status': 'True', 'data': serializer.data}, status=200)

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
        try:
            auditId = request.query_params['id']
            motor = ActiveMotors.objects.get(id=auditId)
            motor.delete()
            return JsonResponse({'message': 'Motor has been deleted'}, status=200)
        except:
            return JsonResponse({'message': 'Sorry! No Motor found.'}, status=500)

    @staticmethod
    def GetActiveMotorHistoryList(request):
        try:
            id = request.query_params['id']
            ParentStatus = request.query_params['parent_status']
            print(ParentStatus)
            dataList = ActiveMotorsHistory.objects.filter(active_motor_id=id)
            if ParentStatus != '':
                if ParentStatus == 'Qualification of raw material of insulation':
                    dataList = dataList.extra(
                        select={'year': 'extract (year from qualification_insulation_lining_propellant_rm_date)',
                                'month': 'extract (month from qualification_insulation_lining_propellant_rm_date)',
                                'day': 'extract (day from qualification_insulation_lining_propellant_rm_date)'},
                        order_by=['month', 'day', '-year'])

                elif ParentStatus == 'Lining':
                    dataList = dataList.extra(
                        select={'year': 'extract (year from lining_date)',
                                'month': 'extract (month from lining_date)',
                                'day': 'extract (day from lining_date)'},
                        order_by=['month', 'day', '-year'])
                elif ParentStatus == 'Qualification of Propellant':
                    dataList = dataList.extra(
                        select={'year': 'extract (year from propellant_mechanical_properties_date)',
                                'month': 'extract (month from propellant_mechanical_properties_date)',
                                'day': 'extract (day from propellant_mechanical_properties_date)'},
                        order_by=['month', 'day', '-year'])

                elif ParentStatus == 'Acceptance of casting':
                    dataList = dataList.extra(
                        select={'year': 'extract (year from acceptance_casting_date)',
                                'month': 'extract (month from acceptance_casting_date)',
                                'day': 'extract (day from acceptance_casting_date)'},
                        order_by=['month', 'day', '-year'])

                elif ParentStatus == 'Sandblasting':
                    dataList = dataList.extra(
                        select={'year': 'extract (year from sandblasting_date)',
                                'month': 'extract (month from sandblasting_date)',
                                'day': 'extract (day from sandblasting_date)'},
                        order_by=['month', 'day', '-year'])

                elif ParentStatus == 'Insulation':
                    dataList = dataList.extra(
                        select={'year': 'extract (year from insulation_date)',
                                'month': 'extract (month from insulation_date)',
                                'day': 'extract (day from insulation_date)'},
                        order_by=['month', 'day', '-year'])

                elif ParentStatus == 'UT and RT of Insulated Case':
                    dataList = dataList.extra(
                        select={'year': 'extract (year from ut_rt_insulated_case_date)',
                                'month': 'extract (month from ut_rt_insulated_case_date)',
                                'day': 'extract (day from ut_rt_insulated_case_date)'},
                        order_by=['month', 'day', '-year'])

                elif ParentStatus == 'Acceptance of Sliver Material':
                    dataList = dataList.extra(
                        select={'year': 'extract (year from acceptance_silver_material_date)',
                                'month': 'extract (month from acceptance_silver_material_date)',
                                'day': 'extract (day from acceptance_silver_material_date)'},
                        order_by=['month', 'day', '-year'])

                elif ParentStatus == 'Silver Application':
                    dataList = dataList.extra(
                        select={'year': 'extract (year from silver_application_date)',
                                'month': 'extract (month from silver_application_date)',
                                'day': 'extract (day from silver_application_date)'},
                        order_by=['month', 'day', '-year'])

                elif ParentStatus == 'Formulation tailoring':
                    dataList = dataList.extra(
                        select={'year': 'extract (year from formulation_tailoring_liner_propellant_date)',
                                'month': 'extract (month from formulation_tailoring_liner_propellant_date)',
                                'day': 'extract (day from formulation_tailoring_liner_propellant_date)'},
                        order_by=['month', 'day', '-year'])

                elif ParentStatus == 'Conditioning of Raw Material':
                    dataList = dataList.extra(
                        select={'year': 'extract (year from conditioning_raw_materials_date)',
                                'month': 'extract (month from conditioning_raw_materials_date)',
                                'day': 'extract (day from conditioning_raw_materials_date)'},
                        order_by=['month', 'day', '-year'])

                elif ParentStatus == 'Conditioning of Lining':
                    dataList = dataList.extra(
                        select={'year': 'extract (year from conditioning_of_lining_date)',
                                'month': 'extract (month from conditioning_of_lining_date)',
                                'day': 'extract (day from conditioning_of_lining_date)'},
                        order_by=['month', 'day', '-year'])

                elif ParentStatus == 'Casing':
                    dataList = dataList.extra(
                        select={'year': 'extract (year from casting_date)',
                                'month': 'extract (month from casting_date)',
                                'day': 'extract (day from casting_date)'},
                        order_by=['month', 'day', '-year'])

                elif ParentStatus == 'Curing':
                    dataList = dataList.extra(
                        select={'year': 'extract (year from curing_date)',
                                'month': 'extract (month from curing_date)',
                                'day': 'extract (day from curing_date)'},
                        order_by=['month', 'day', '-year'])

                elif ParentStatus == 'UT, endoscopy and RT of grain':
                    dataList = dataList.extra(
                        select={'year': 'extract (year from ut_endoscopy_rt_grain_date)',
                                'month': 'extract (month from ut_endoscopy_rt_grain_date)',
                                'day': 'extract (day from ut_endoscopy_rt_grain_date)'},
                        order_by=['month', 'day', '-year'])


                elif ParentStatus == 'Mechanical Properties of Liner':
                    dataList = dataList.extra(
                        select={'year': 'extract (year from liner_mechanical_properties_date)',
                                'month': 'extract (month from liner_mechanical_properties_date)',
                                'day': 'extract (day from liner_mechanical_properties_date)'},
                        order_by=['month', 'day', '-year'])

                elif ParentStatus == 'Mechanical Properties of Propellant':
                    dataList = dataList.extra(
                        select={'year': 'extract (year from mechanical_properties_propellant_date)',
                                'month': 'extract (month from mechanical_properties_propellant_date)',
                                'day': 'extract (day from mechanical_properties_propellant_date)'},
                        order_by=['month', 'day', '-year'])

                elif ParentStatus == 'Interface bond strength':
                    dataList = dataList.extra(
                        select={'year': 'extract (year from interface_bond_strength_date)',
                                'month': 'extract (month from interface_bond_strength_date)',
                                'day': 'extract (day from interface_bond_strength_date)'},
                        order_by=['month', 'day', '-year'])

                elif ParentStatus == 'Propellant burn rate':
                    dataList = dataList.extra(
                        select={'year': 'extract (year from propellant_burn_rate_date)',
                                'month': 'extract (month from propellant_burn_rate_date)',
                                'day': 'extract (day from propellant_burn_rate_date)'},
                        order_by=['month', 'day', '-year'])

                elif ParentStatus == 'Trimming of propellant grain':
                    dataList = dataList.extra(
                        select={'year': 'extract (year from trimming_Propellant_grain_date)',
                                'month': 'extract (month from trimming_Propellant_grain_date)',
                                'day': 'extract (day from trimming_Propellant_grain_date)'},
                        order_by=['month', 'day', '-year'])

                elif ParentStatus == 'Mass of liner':
                    dataList = dataList.extra(
                        select={'year': 'extract (year from mass_liner_insulation_propellant_srm_date)',
                                'month': 'extract (month from mass_liner_insulation_propellant_srm_date)',
                                'day': 'extract (day from mass_liner_insulation_propellant_srm_date)'},
                        order_by=['month', 'day', '-year'])

                elif ParentStatus == 'Mass of Insulation':
                    dataList = dataList.extra(
                        select={'year': 'extract (year from mass_insulation_date)',
                                'month': 'extract (month from mass_insulation_date)',
                                'day': 'extract (day from mass_insulation_date)'},
                        order_by=['month', 'day', '-year'])

                elif ParentStatus == 'Mass of Propellant':
                    dataList = dataList.extra(
                        select={'year': 'extract (year from mass_propellant_date)',
                                'month': 'extract (month from mass_propellant_date)',
                                'day': 'extract (day from mass_propellant_date)'},
                        order_by=['month', 'day', '-year'])

                elif ParentStatus == 'BHD Status':
                    dataList = dataList.filter(bhd_date__year=current_year, bhd_status=ChildStatus).extra(
                        select={'year': 'extract (year from bhd_date)',
                                'month': 'extract (month from bhd_date)',
                                'day': 'extract (day from bhd_date)'},
                        order_by=['month', 'day', '-year'])

                elif ParentStatus == 'QM Certification Status':
                    dataList = dataList.filter(qm_certification_date__year=current_year,
                                                qm_certification_status=ChildStatus).extra(
                        select={'year': 'extract (year from qm_certification_date)',
                                'month': 'extract (month from qm_certification_date)',
                                'day': 'extract (day from qm_certification_date)'},
                        order_by=['month', 'day', '-year'])

                elif ParentStatus == 'Overall Qualification':
                    dataList = dataList.filter(overall_qualification_date__year=current_year,
                                                overall_qualification_status=ChildStatus).extra(
                        select={'year': 'extract (year from overall_qualification_date)',
                                'month': 'extract (month from overall_qualification_date)',
                                'day': 'extract (day from overall_qualification_date)'},
                        order_by=['month', 'day', '-year'])


                elif ParentStatus == 'Material Qualified':
                    dataList = dataList.extra(
                        select={'year': 'extract (year from matiral_qualified_date)',
                                'month': 'extract (month from matiral_qualified_date)',
                                'day': 'extract (day from matiral_qualified_date)'},
                        order_by=['month', 'day', '-year'])

                elif ParentStatus == 'Component Manufacturing':
                    dataList = dataList.extra(
                        select={'year': 'extract (year from compo_manufacturing_date)',
                                'month': 'extract (month from compo_manufacturing_date)',
                                'day': 'extract (day from compo_manufacturing_date)'},
                        order_by=['month', 'day', '-year'])

                elif ParentStatus == 'Power Pack Assembly':
                    dataList = dataList.extra(
                        select={'year': 'extract (year from powerpack_assembly_date)',
                                'month': 'extract (month from powerpack_assembly_date)',
                                'day': 'extract (day from powerpack_assembly_date)'},
                        order_by=['month', 'day', '-year'])

                elif ParentStatus == 'Power Pack Testing':
                    dataList = dataList.extra(
                        select={'year': 'extract (year from powerpack_testing_date)',
                                'month': 'extract (month from powerpack_testing_date)',
                                'day': 'extract (day from powerpack_testing_date)'},
                        order_by=['month', 'day', '-year'])

                elif ParentStatus == 'Raw Material Inspection':
                    dataList = dataList.extra(
                        select={'year': 'extract (year from raw_material_inspection_date)',
                                'month': 'extract (month from raw_material_inspection_date)',
                                'day': 'extract (day from raw_material_inspection_date)'},
                        order_by=['month', 'day', '-year'])

                elif ParentStatus == 'Pressing Electrode':
                    dataList = dataList.extra(
                        select={'year': 'extract (year from pressing_electrode_date)',
                                'month': 'extract (month from pressing_electrode_date)',
                                'day': 'extract (day from pressing_electrode_date)'},
                        order_by=['month', 'day', '-year'])

                elif ParentStatus == 'Formation Process':
                    dataList = dataList.extra(
                        select={'year': 'extract (year from formation_process_date)',
                                'month': 'extract (month from formation_process_date)',
                                'day': 'extract (day from formation_process_date)'},
                        order_by=['month', 'day', '-year'])

                elif ParentStatus == 'Assembly Process':
                    dataList = dataList.extra(
                        select={'year': 'extract (year from assembly_process_date)',
                                'month': 'extract (month from assembly_process_date)',
                                'day': 'extract (day from assembly_process_date)'},
                        order_by=['month', 'day', '-year'])

                elif ParentStatus == 'Battery Testing':
                    dataList = dataList.extra(
                        select={'year': 'extract (year from battery_testing_date)',
                                'month': 'extract (month from battery_testing_date)',
                                'day': 'extract (day from battery_testing_date)'},
                        order_by=['month', 'day', '-year'])

                elif ParentStatus == 'Final Qualification':
                    dataList = dataList.extra(
                        select={'year': 'extract (year from final_qualification_date)',
                                'month': 'extract (month from final_qualification_date)',
                                'day': 'extract (day from final_qualification_date)'},
                        order_by=['month', 'day', '-year'])


                elif ParentStatus == 'Qualification of Raw Material':
                    dataList = dataList.extra(
                        select={'year': 'extract (year from qualification_raw_material_date)',
                                'month': 'extract (month from qualification_raw_material_date)',
                                'day': 'extract (day from qualification_raw_material_date)'},
                        order_by=['month', 'day', '-year'])

                elif ParentStatus == 'Filling':
                    dataList = dataList.extra(
                        select={'year': 'extract (year from filling_date)',
                                'month': 'extract (month from filling_date)',
                                'day': 'extract (day from filling_date)'},
                        order_by=['month', 'day', '-year'])

                elif ParentStatus == 'Assembling/Integration':
                    dataList = dataList.extra(
                        select={'year': 'extract (year from assembling_integration_date)',
                                'month': 'extract (month from assembling_integration_date)',
                                'day': 'extract (day from assembling_integration_date)'},
                        order_by=['month', 'day', '-year'])

                elif ParentStatus == 'Qualification Testing':
                    dataList = dataList.extra(
                        select={'year': 'extract (year from qualification_testing_date)',
                                'month': 'extract (month from qualification_testing_date)',
                                'day': 'extract (day from qualification_testing_date)'},
                        order_by=['month', 'day', '-year'])

                elif ParentStatus == 'Performance Testing':
                    dataList = dataList.extra(
                        select={'year': 'extract (year from performance_testing_date)',
                                'month': 'extract (month from performance_testing_date)',
                                'day': 'extract (day from performance_testing_date)'},
                        order_by=['month', 'day', '-year'])

            serializer = ActiveMotorSerializer(dataList, many=True)
            return JsonResponse({'message': 'Welcome to Home Page', 'data': serializer.data}, status=200)

        except:
            return JsonResponse({'message': 'Sorry! No Task found.'}, status=200)

    @staticmethod
    def DeleteActiveMotorHistory(request):
        try:
            id = request.query_params['id']
            history = ActiveMotorsHistory.objects.filter(id=id)
            history.delete()
            return JsonResponse({'status': 'True', 'message': "Record Deleted"},
                                status=200)

        except Exception as e:
            print(e)
            return JsonResponse({'status': 'False', "message": "Doc Not Deleted"}, status=500)
            pass
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
            id = request['id']
            if id == '0':
                modal.organization = request['organization']
                modal.lot_id_number = request['lot_id_number']
                modal.save()
            else:
                get_lot = organization_lot_ids.objects.filter(id=id).first()
                get_lot.organization = request['organization']
                get_lot.lot_id_number = request['lot_id_number']
                get_lot.save()
            return JsonResponse({'status': 'true', "message": "Record Added"}, status=200)
        except Exception as e:
            return JsonResponse({'status': 'False', "message": "Data Not Saved"}, status=500)
    @staticmethod
    def DeleteLotId(request):
        try:
            lotId = request.query_params['id']
            motor = organization_lot_ids.objects.get(id=lotId)
            motor.delete()
            return JsonResponse({'message': 'Lot has been deleted'}, status=200)
        except:
            return JsonResponse({'message': 'Sorry! No Lot found.'}, status=500)
    @staticmethod
    def GetLots(request):
        try:
            data = organization_lot_ids.objects.all().order_by('-id')
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
            srm_objects = Q()
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
            pd_type = request.query_params['selected_pdType']
            # parent_status = request.query_params['parent_status']
            # child_status = request.query_params['child_status']
            # create dynamic filter
            if component != '' and component == 'SRMs':
                srm_objects &= get_filter(
                    'component_type', 'equal',
                    component)
            if system_type != '':
                srm_objects &= get_filter(
                    'system_type', 'equal',
                    system_type)
            if system != '':
                srm_objects &= get_filter(
                    'system_name', 'equal',
                    system)
            if pd_type != '':
                srm_objects &= get_filter(
                    'pd_type', 'equal',
                    pd_type)
            if current_org !='':
                srm_objects &= get_filter(
                    'organization', 'equal', current_org)
            if current_bat !='':
                srm_objects &= get_filter(
                    'battery_type', 'equal', current_bat)
            if current_lot !='':
                srm_objects &= get_filter(
                    'lot_id', 'equal', current_lot)
            if component == 'SRMs':
                srm_objects &= get_filter(
                    'motor_id', 'not_equal',
                    '')

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
            if pd_type != '':
                filter_objects &= get_filter(
                    'pd_type', 'equal',
                    pd_type)
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
            MotorTotalCount =  ActiveMotors.objects.filter(srm_objects).count()
            qualification_insulation_lining_propellant_rm_OK = ActiveMotors.objects.filter(filter_objects, qualification_insulation_lining_propellant_rm_date__year = year,qualification_insulation_lining_propellant_rm='Ok')
            qualification_insulation_lining_propellant_rm_ObsSame = ActiveMotors.objects.filter(filter_objects,
                                                                             qualification_insulation_lining_propellant_rm_date__year = year,qualification_insulation_lining_propellant_rm='Observation(same stage)')
            qualification_insulation_lining_propellant_rm_ObsNext = ActiveMotors.objects.filter(filter_objects,
                                                                                    qualification_insulation_lining_propellant_rm_date__year = year,qualification_insulation_lining_propellant_rm='Observation(next stage)')
            qualification_insulation_lining_propellant_rm_Up = ActiveMotors.objects.filter(filter_objects, qualification_insulation_lining_propellant_rm_date__year = year,qualification_insulation_lining_propellant_rm='Under process')
            qualification_insulation_lining_propellant_rm_Hlt = ActiveMotors.objects.filter(filter_objects, qualification_insulation_lining_propellant_rm_date__year = year,qualification_insulation_lining_propellant_rm='Halt')

            acceptance_casting_Ok = ActiveMotors.objects.filter(filter_objects, acceptance_casting_date__year= year,acceptance_casting='Ok')
            acceptance_casting_ObsSame = ActiveMotors.objects.filter(filter_objects,
                                                                                acceptance_casting_date__year=year,acceptance_casting='Observation(same stage)')
            acceptance_casting_ObsNext = ActiveMotors.objects.filter(filter_objects,
                                                                                    acceptance_casting_date__year=year,acceptance_casting='Observation(next stage)')
            acceptance_casting_Up = ActiveMotors.objects.filter(filter_objects, acceptance_casting_date__year=year,acceptance_casting='Under process')
            acceptance_casting_Hlt = ActiveMotors.objects.filter(filter_objects, acceptance_casting_date__year=year,acceptance_casting='Halt')



            sandblasting_Ok = ActiveMotors.objects.filter(filter_objects,sandblasting_date__year = year, sandblasting='Ok')
            sandblasting_ObsSame = ActiveMotors.objects.filter(filter_objects,sandblasting_date__year = year,sandblasting='Observation(same stage)')
            sandblasting_ObsNext = ActiveMotors.objects.filter(filter_objects,sandblasting_date__year = year,sandblasting='Observation(next stage)')
            sandblasting_Up = ActiveMotors.objects.filter(filter_objects,sandblasting_date__year = year, sandblasting='Under process')
            sandblasting_Hlt = ActiveMotors.objects.filter(filter_objects,sandblasting_date__year = year, sandblasting='Halt')



            insulation_Ok = ActiveMotors.objects.filter(filter_objects,insulation_date__year = year, insulation='Ok')
            insulation_ObsSame = ActiveMotors.objects.filter(filter_objects,insulation_date__year = year,
                                                                                insulation='Observation(same stage)')
            insulation_ObsNext = ActiveMotors.objects.filter(filter_objects,insulation_date__year = year,
                                                                                    insulation='Observation(next stage)')
            insulation_Up = ActiveMotors.objects.filter(filter_objects,insulation_date__year = year, insulation='Under process')
            insulation_Hlt = ActiveMotors.objects.filter(filter_objects,insulation_date__year = year, insulation='Halt')


            ut_rt_insulated_case_Ok = ActiveMotors.objects.filter(filter_objects,ut_rt_insulated_case_date__year = year, ut_rt_insulated_case='Ok')
            ut_rt_insulated_case_ObsSame = ActiveMotors.objects.filter(filter_objects,ut_rt_insulated_case_date__year = year,
                                                                                ut_rt_insulated_case='Observation(same stage)')
            ut_rt_insulated_case_ObsNext = ActiveMotors.objects.filter(filter_objects,ut_rt_insulated_case_date__year = year,
                                                                                    ut_rt_insulated_case='Observation(next stage)')
            ut_rt_insulated_case_Up = ActiveMotors.objects.filter(filter_objects,ut_rt_insulated_case_date__year = year, ut_rt_insulated_case='Under process')
            ut_rt_insulated_case_Hlt = ActiveMotors.objects.filter(filter_objects,ut_rt_insulated_case_date__year = year, ut_rt_insulated_case='Halt')


            acceptance_silver_material_Ok = ActiveMotors.objects.filter(filter_objects,acceptance_silver_material_date__year = year, acceptance_silver_material='Ok')
            acceptance_silver_material_ObsSame = ActiveMotors.objects.filter(filter_objects,acceptance_silver_material_date__year = year,
                                                                                acceptance_silver_material='Observation(same stage)')
            acceptance_silver_material_ObsNext = ActiveMotors.objects.filter(filter_objects,acceptance_silver_material_date__year = year,
                                                                                    acceptance_silver_material='Observation(next stage)')
            acceptance_silver_material_Up = ActiveMotors.objects.filter(filter_objects,acceptance_silver_material_date__year = year, acceptance_silver_material='Under process')
            acceptance_silver_material_Hlt = ActiveMotors.objects.filter(filter_objects,acceptance_silver_material_date__year = year, acceptance_silver_material='Halt')


            silver_application_Ok = ActiveMotors.objects.filter(filter_objects,silver_application_date__year = year, silver_application='Ok')
            silver_application_ObsSame = ActiveMotors.objects.filter(filter_objects,silver_application_date__year = year,
                                                                                silver_application='Observation(same stage)')
            silver_application_ObsNext = ActiveMotors.objects.filter(filter_objects,silver_application_date__year = year,
                                                                                    silver_application='Observation(next stage)')
            silver_application_Up = ActiveMotors.objects.filter(filter_objects,silver_application_date__year = year, silver_application='Under process')
            silver_application_Hlt = ActiveMotors.objects.filter(filter_objects, silver_application_date__year = year,silver_application='Halt')


            formulation_tailoring_liner_propellant_Ok = ActiveMotors.objects.filter(filter_objects,formulation_tailoring_liner_propellant_date__year = year, formulation_tailoring_liner_propellant='Ok')
            formulation_tailoring_liner_propellant_ObsSame = ActiveMotors.objects.filter(filter_objects,formulation_tailoring_liner_propellant_date__year = year,
                                                                                formulation_tailoring_liner_propellant='Observation(same stage)')
            formulation_tailoring_liner_propellant_ObsNext = ActiveMotors.objects.filter(filter_objects,formulation_tailoring_liner_propellant_date__year = year,
                                                                                    formulation_tailoring_liner_propellant='Observation(next stage)')
            formulation_tailoring_liner_propellant_Up = ActiveMotors.objects.filter(filter_objects, formulation_tailoring_liner_propellant_date__year = year,formulation_tailoring_liner_propellant='Under process')
            formulation_tailoring_liner_propellant_Hlt = ActiveMotors.objects.filter(filter_objects,formulation_tailoring_liner_propellant_date__year = year, formulation_tailoring_liner_propellant='Halt')


            conditioning_raw_materials_Ok = ActiveMotors.objects.filter(filter_objects,conditioning_raw_materials_date__year = year, conditioning_raw_materials='Ok')
            conditioning_raw_materials_ObsSame = ActiveMotors.objects.filter(filter_objects,conditioning_raw_materials_date__year = year,
                                                                                conditioning_raw_materials='Observation(same stage)')
            conditioning_raw_materials_ObsNext = ActiveMotors.objects.filter(filter_objects,conditioning_raw_materials_date__year = year,
                                                                                    conditioning_raw_materials='Observation(next stage)')
            conditioning_raw_materials_Up = ActiveMotors.objects.filter(filter_objects,conditioning_raw_materials_date__year = year, conditioning_raw_materials='Under process')
            conditioning_raw_materials_Hlt = ActiveMotors.objects.filter(filter_objects,conditioning_raw_materials_date__year = year, conditioning_raw_materials='Halt')


            lining_Ok = ActiveMotors.objects.filter(filter_objects,lining_date__year = year, lining='Ok')
            lining_ObsSame = ActiveMotors.objects.filter(filter_objects,lining_date__year = year,
                                                                                lining='Observation(same stage)')
            lining_ObsNext = ActiveMotors.objects.filter(filter_objects,lining_date__year = year,
                                                                                    lining='Observation(next stage)')
            lining_Up = ActiveMotors.objects.filter(filter_objects,lining_date__year = year,  lining='Under process')
            lining_Hlt = ActiveMotors.objects.filter(filter_objects,lining_date__year = year,  lining='Halt')


            casting_Ok = ActiveMotors.objects.filter(filter_objects,casting_date__year = year, casting='Ok')
            casting_ObsSame = ActiveMotors.objects.filter(filter_objects,casting_date__year = year,
                                                                                casting='Observation(same stage)')
            casting_ObsNext = ActiveMotors.objects.filter(filter_objects,casting_date__year = year,
                                                                                    casting='Observation(next stage)')
            casting_Up = ActiveMotors.objects.filter(filter_objects, casting_date__year = year,casting='Under process')
            casting_Hlt = ActiveMotors.objects.filter(filter_objects, casting_date__year = year,casting='Halt')


            curing_Ok = ActiveMotors.objects.filter(filter_objects,curing_date__year = year, curing='Ok')
            curing_ObsSame = ActiveMotors.objects.filter(filter_objects,curing_date__year = year,
                                                                                curing='Observation(same stage)')
            curing_ObsNext = ActiveMotors.objects.filter(filter_objects,curing_date__year = year,
                                                                                    curing='Observation(next stage)')
            curing_Up = ActiveMotors.objects.filter(filter_objects,curing_date__year = year, curing='Under process')
            curing_Hlt = ActiveMotors.objects.filter(filter_objects,curing_date__year = year, curing='Halt')


            liner_mechanical_properties_Ok = ActiveMotors.objects.filter(filter_objects,liner_mechanical_properties_date__year= year, liner_mechanical_properties='Ok')
            liner_mechanical_properties_ObsSame = ActiveMotors.objects.filter(filter_objects,liner_mechanical_properties_date__year= year,
                                                                                liner_mechanical_properties='Observation(same stage)')
            liner_mechanical_properties_ObsNext = ActiveMotors.objects.filter(filter_objects,liner_mechanical_properties_date__year= year,
                                                                                    liner_mechanical_properties='Observation(next stage)')
            liner_mechanical_properties_Up = ActiveMotors.objects.filter(filter_objects,liner_mechanical_properties_date__year= year, liner_mechanical_properties='Under process')
            liner_mechanical_properties_Hlt = ActiveMotors.objects.filter(filter_objects,liner_mechanical_properties_date__year= year, liner_mechanical_properties='Halt')


            propellant_mechanical_properties_Ok = ActiveMotors.objects.filter(filter_objects,propellant_mechanical_properties_date__year=year, propellant_mechanical_properties='Ok')
            propellant_mechanical_properties_ObsSame = ActiveMotors.objects.filter(filter_objects,propellant_mechanical_properties_date__year=year,
                                                                                propellant_mechanical_properties='Observation(same stage)')
            propellant_mechanical_properties_ObsNext = ActiveMotors.objects.filter(filter_objects,propellant_mechanical_properties_date__year=year,
                                                                                    propellant_mechanical_properties='Observation(next stage)')
            propellant_mechanical_properties_Up = ActiveMotors.objects.filter(filter_objects,propellant_mechanical_properties_date__year=year, propellant_mechanical_properties='Under process')
            propellant_mechanical_properties_Hlt = ActiveMotors.objects.filter(filter_objects,propellant_mechanical_properties_date__year=year, propellant_mechanical_properties='Halt')


            interface_bond_strength_Ok = ActiveMotors.objects.filter(filter_objects,interface_bond_strength_date__year = year, interface_bond_strength='Ok')
            interface_bond_strength_ObsSame = ActiveMotors.objects.filter(filter_objects,interface_bond_strength_date__year = year,
                                                                                interface_bond_strength='Observation(same stage)')
            interface_bond_strength_ObsNext = ActiveMotors.objects.filter(filter_objects,interface_bond_strength_date__year = year,
                                                                                    interface_bond_strength='Observation(next stage)')
            interface_bond_strength_Up = ActiveMotors.objects.filter(filter_objects,interface_bond_strength_date__year = year, interface_bond_strength='Under process')
            interface_bond_strength_Hlt = ActiveMotors.objects.filter(filter_objects,interface_bond_strength_date__year = year, interface_bond_strength='Halt')


            propellant_burn_rate_Ok = ActiveMotors.objects.filter(filter_objects,propellant_burn_rate_date__year=year, propellant_burn_rate='Ok')
            propellant_burn_rate_ObsSame = ActiveMotors.objects.filter(filter_objects,propellant_burn_rate_date__year=year,
                                                                                propellant_burn_rate='Observation(same stage)')
            propellant_burn_rate_ObsNext = ActiveMotors.objects.filter(filter_objects,propellant_burn_rate_date__year=year,
                                                                                    propellant_burn_rate='Observation(next stage)')
            propellant_burn_rate_Up = ActiveMotors.objects.filter(filter_objects, propellant_burn_rate_date__year=year, propellant_burn_rate='Under process')
            propellant_burn_rate_Hlt = ActiveMotors.objects.filter(filter_objects,propellant_burn_rate_date__year=year,  propellant_burn_rate='Halt')


            trimming_Propellant_grain_Ok = ActiveMotors.objects.filter(filter_objects,trimming_Propellant_grain_date__year=year, trimming_Propellant_grain='Ok')
            trimming_Propellant_grain_ObsSame = ActiveMotors.objects.filter(filter_objects,trimming_Propellant_grain_date__year=year,
                                                                                trimming_Propellant_grain='Observation(same stage)')
            trimming_Propellant_grain_ObsNext = ActiveMotors.objects.filter(filter_objects,trimming_Propellant_grain_date__year=year,
                                                                                    trimming_Propellant_grain='Observation(next stage)')
            trimming_Propellant_grain_Up = ActiveMotors.objects.filter(filter_objects,trimming_Propellant_grain_date__year=year, trimming_Propellant_grain='Under process')
            trimming_Propellant_grain_Hlt = ActiveMotors.objects.filter(filter_objects,trimming_Propellant_grain_date__year=year, trimming_Propellant_grain='Halt')


            mass_liner_insulation_propellant_srm_Ok = ActiveMotors.objects.filter(filter_objects,mass_liner_insulation_propellant_srm_date__year=year, mass_liner_insulation_propellant_srm='Ok')
            mass_liner_insulation_propellant_srm_ObsSame = ActiveMotors.objects.filter(filter_objects,mass_liner_insulation_propellant_srm_date__year=year,
                                                                                mass_liner_insulation_propellant_srm='Observation(same stage)')
            mass_liner_insulation_propellant_srm_ObsNext = ActiveMotors.objects.filter(filter_objects,mass_liner_insulation_propellant_srm_date__year=year,
                                                                                    mass_liner_insulation_propellant_srm='Observation(next stage)')
            mass_liner_insulation_propellant_srm_Up = ActiveMotors.objects.filter(filter_objects,mass_liner_insulation_propellant_srm_date__year=year, mass_liner_insulation_propellant_srm='Under process')
            mass_liner_insulation_propellant_srm_Hlt = ActiveMotors.objects.filter(filter_objects,mass_liner_insulation_propellant_srm_date__year=year, mass_liner_insulation_propellant_srm='Halt')


            ut_endoscopy_rt_grain_Ok = ActiveMotors.objects.filter(filter_objects,ut_endoscopy_rt_grain_date__year = year, ut_endoscopy_rt_grain='Ok')
            ut_endoscopy_rt_grain_ObsSame = ActiveMotors.objects.filter(filter_objects,ut_endoscopy_rt_grain_date__year = year,
                                                                                ut_endoscopy_rt_grain='Observation(same stage)')
            ut_endoscopy_rt_grain_ObsNext = ActiveMotors.objects.filter(filter_objects,ut_endoscopy_rt_grain_date__year = year,
                                                                                    ut_endoscopy_rt_grain='Observation(next stage)')
            ut_endoscopy_rt_grain_Up = ActiveMotors.objects.filter(filter_objects,ut_endoscopy_rt_grain_date__year = year, ut_endoscopy_rt_grain='Under process')
            ut_endoscopy_rt_grain_Hlt = ActiveMotors.objects.filter(filter_objects,ut_endoscopy_rt_grain_date__year = year, ut_endoscopy_rt_grain='Halt')

            conditioning_of_lining_Ok = ActiveMotors.objects.filter(filter_objects, conditioning_of_lining_date__year = year, conditioning_of_lining='Ok')
            conditioning_of_lining_ObsSame = ActiveMotors.objects.filter(filter_objects,conditioning_of_lining_date__year = year,
                                                                        conditioning_of_lining='Observation(same stage)')
            conditioning_of_lining_ObsNext = ActiveMotors.objects.filter(filter_objects,conditioning_of_lining_date__year = year,
                                                                        conditioning_of_lining='Observation(next stage)')
            conditioning_of_lining_Up = ActiveMotors.objects.filter(filter_objects,conditioning_of_lining_date__year = year, conditioning_of_lining='Under process')
            conditioning_of_lining_Hlt = ActiveMotors.objects.filter(filter_objects,conditioning_of_lining_date__year = year, conditioning_of_lining='Halt')


            mechanical_properties_liner_Ok = ActiveMotors.objects.filter(filter_objects,mechanical_properties_liner_date__year = year, mechanical_properties_liner='Ok')
            mechanical_properties_liner_ObsSame = ActiveMotors.objects.filter(filter_objects,mechanical_properties_liner_date__year = year,
                                                                         mechanical_properties_liner='Observation(same stage)')
            mechanical_properties_liner_ObsNext = ActiveMotors.objects.filter(filter_objects,mechanical_properties_liner_date__year = year,
                                                                         mechanical_properties_liner='Observation(next stage)')
            mechanical_properties_liner_Up = ActiveMotors.objects.filter(filter_objects,mechanical_properties_liner_date__year = year, mechanical_properties_liner='Under process')
            mechanical_properties_liner_Hlt = ActiveMotors.objects.filter(filter_objects,mechanical_properties_liner_date__year = year, mechanical_properties_liner='Halt')


            mechanical_properties_propellant_Ok = ActiveMotors.objects.filter(filter_objects,mechanical_properties_propellant_date__year=year, mechanical_properties_propellant='Ok')
            mechanical_properties_propellant_ObsSame = ActiveMotors.objects.filter(filter_objects,mechanical_properties_propellant_date__year=year,
                                                                              mechanical_properties_propellant='Observation(same stage)')
            mechanical_properties_propellant_ObsNext = ActiveMotors.objects.filter(filter_objects,mechanical_properties_propellant_date__year=year,
                                                                              mechanical_properties_propellant='Observation(next stage)')
            mechanical_properties_propellant_Up = ActiveMotors.objects.filter(filter_objects,mechanical_properties_propellant_date__year=year,
                                                                         mechanical_properties_propellant='Under process')
            mechanical_properties_propellant_Hlt = ActiveMotors.objects.filter(filter_objects,mechanical_properties_propellant_date__year=year, mechanical_properties_propellant='Halt')


            mass_liner_Ok = ActiveMotors.objects.filter(filter_objects,mass_liner_date__year = year, mass_liner='Ok')
            mass_liner_ObsSame = ActiveMotors.objects.filter(filter_objects,mass_liner_date__year = year,
                                                                              mass_liner='Observation(same stage)')
            mass_liner_ObsNext = ActiveMotors.objects.filter(filter_objects,mass_liner_date__year = year,
                                                                              mass_liner='Observation(next stage)')
            mass_liner_Up = ActiveMotors.objects.filter(filter_objects,mass_liner_date__year = year,
                                                                         mass_liner='Under process')
            mass_liner_Hlt = ActiveMotors.objects.filter(filter_objects,mass_liner_date__year = year, mass_liner='Halt')


            mass_insulation_Ok = ActiveMotors.objects.filter(filter_objects,mass_insulation_date__year = year, mass_insulation='Ok')
            mass_insulation_ObsSame = ActiveMotors.objects.filter(filter_objects,mass_insulation_date__year = year,
                                                                              mass_insulation='Observation(same stage)')
            mass_insulation_ObsNext = ActiveMotors.objects.filter(filter_objects,mass_insulation_date__year = year,
                                                                              mass_insulation='Observation(next stage)')
            mass_insulation_Up = ActiveMotors.objects.filter(filter_objects,mass_insulation_date__year = year,
                                                                         mass_insulation='Under process')
            mass_insulation_Hlt = ActiveMotors.objects.filter(filter_objects,mass_insulation_date__year = year, mass_insulation='Halt')


            mass_propellant_Ok = ActiveMotors.objects.filter(filter_objects,mass_propellant_date__year = year, mass_propellant='Ok')
            mass_propellant_ObsSame = ActiveMotors.objects.filter(filter_objects,mass_propellant_date__year = year,
                                                                              mass_propellant='Observation(same stage)')
            mass_propellant_ObsNext = ActiveMotors.objects.filter(filter_objects,mass_propellant_date__year = year,
                                                                              mass_propellant='Observation(next stage)')
            mass_propellant_Up = ActiveMotors.objects.filter(filter_objects,mass_propellant_date__year = year,
                                                                         mass_propellant='Under process')
            mass_propellant_Hlt = ActiveMotors.objects.filter(filter_objects,mass_propellant_date__year = year, mass_propellant='Halt')


            overall_qualification_status_ObsNcr = ActiveMotors.objects.filter(filter_objects,overall_qualification_date__year=year,
                                                                              overall_qualification_status='Observation/NCR')
            overall_qualification_status_Inprocess = ActiveMotors.objects.filter(filter_objects,overall_qualification_date__year=year,
                                                                              overall_qualification_status='in-process')



            bhd_status_Ok = ActiveMotors.objects.filter(filter_objects,bhd_date__year = year, bhd_status='Ok')
            bhd_status_NotSubmit = ActiveMotors.objects.filter(filter_objects,bhd_date__year = year, bhd_status='Not Submitted')
            bhd_status_ObsForwarded = ActiveMotors.objects.filter(filter_objects,bhd_date__year = year,
                                                                              bhd_status='QM Observations Forwarded')
            bhd_status_Inprocess = ActiveMotors.objects.filter(filter_objects,bhd_date__year = year,
                                                                              bhd_status='Audit in-process')


            qm_certification_status_Issued = ActiveMotors.objects.filter(filter_objects,qm_certification_date__year=year, qm_certification_status='QM certificate issued')
            qm_certification_status_ObsForwarded = ActiveMotors.objects.filter(filter_objects,qm_certification_date__year=year,
                                                                              qm_certification_status='QM Observations Forwarded')
            qm_certification_status_Inprocess = ActiveMotors.objects.filter(filter_objects,qm_certification_date__year=year,
                                                                              qm_certification_status='Audit in-process')


            matiral_qualified_status_Ok = ActiveMotors.objects.filter(filter_objects,matiral_qualified_date__year = year, matiral_qualified_status='Ok')
            matiral_qualified_status_ObsSame = ActiveMotors.objects.filter(filter_objects,matiral_qualified_date__year = year,
                                                                                matiral_qualified_status='Observation(same stage)')
            matiral_qualified_status_ObsNext = ActiveMotors.objects.filter(filter_objects,matiral_qualified_date__year = year,
                                                                                    matiral_qualified_status='Observation(next stage)')
            matiral_qualified_status_Up = ActiveMotors.objects.filter(filter_objects,matiral_qualified_date__year = year, matiral_qualified_status='Under process')
            matiral_qualified_status_Hlt = ActiveMotors.objects.filter(filter_objects,matiral_qualified_date__year = year, matiral_qualified_status='Halt')


            compo_manufacturing_status_Ok = ActiveMotors.objects.filter(filter_objects,compo_manufacturing_date__year = year, compo_manufacturing_status='Ok')
            compo_manufacturing_status_ObsSame = ActiveMotors.objects.filter(filter_objects,compo_manufacturing_date__year = year,
                                                                                compo_manufacturing_status='Observation(same stage)')
            compo_manufacturing_status_ObsNext = ActiveMotors.objects.filter(filter_objects,compo_manufacturing_date__year = year,
                                                                                    compo_manufacturing_status='Observation(next stage)')
            compo_manufacturing_status_Up = ActiveMotors.objects.filter(filter_objects,compo_manufacturing_date__year = year, compo_manufacturing_status='Under process')
            compo_manufacturing_status_Hlt = ActiveMotors.objects.filter(filter_objects,compo_manufacturing_date__year = year, compo_manufacturing_status='Halt')


            powerpack_assembly_status_Ok = ActiveMotors.objects.filter(filter_objects,powerpack_assembly_date__year = year, powerpack_assembly_status='Ok')
            powerpack_assembly_status_ObsSame = ActiveMotors.objects.filter(filter_objects,powerpack_assembly_date__year = year,
                                                                                powerpack_assembly_status='Observation(same stage)')
            powerpack_assembly_status_ObsNext = ActiveMotors.objects.filter(filter_objects,powerpack_assembly_date__year = year,
                                                                                    powerpack_assembly_status='Observation(next stage)')
            powerpack_assembly_status_Up = ActiveMotors.objects.filter(filter_objects,powerpack_assembly_date__year = year, powerpack_assembly_status='Under process')
            powerpack_assembly_status_Hlt = ActiveMotors.objects.filter(filter_objects,powerpack_assembly_date__year = year, powerpack_assembly_status='Halt')


            powerpack_testing_status_Ok = ActiveMotors.objects.filter(filter_objects,powerpack_testing_date__year = year, powerpack_testing_status='Ok')
            powerpack_testing_status_ObsSame = ActiveMotors.objects.filter(filter_objects,powerpack_testing_date__year = year,
                                                                                powerpack_testing_status='Observation(same stage)')
            powerpack_testing_status_ObsNext = ActiveMotors.objects.filter(filter_objects,powerpack_testing_date__year = year,
                                                                                    powerpack_testing_status='Observation(next stage)')
            powerpack_testing_status_Up = ActiveMotors.objects.filter(filter_objects,powerpack_testing_date__year = year, powerpack_testing_status='Under process')
            powerpack_testing_status_Hlt = ActiveMotors.objects.filter(filter_objects,powerpack_testing_date__year = year, powerpack_testing_status='Halt')



            raw_material_inspection_status_Ok = ActiveMotors.objects.filter(filter_objects,raw_material_inspection_date__year = year, raw_material_inspection_status='Ok')
            raw_material_inspection_status_ObsSame = ActiveMotors.objects.filter(filter_objects,raw_material_inspection_date__year = year,
                                                                                raw_material_inspection_status='Observation(same stage)')
            raw_material_inspection_status_ObsNext = ActiveMotors.objects.filter(filter_objects,raw_material_inspection_date__year = year,
                                                                                    raw_material_inspection_status='Observation(next stage)')
            raw_material_inspection_status_Up = ActiveMotors.objects.filter(filter_objects,raw_material_inspection_date__year = year, raw_material_inspection_status='Under process')
            raw_material_inspection_status_Hlt = ActiveMotors.objects.filter(filter_objects,raw_material_inspection_date__year = year, raw_material_inspection_status='Halt')



            pressing_electrode_status_Ok = ActiveMotors.objects.filter(filter_objects,pressing_electrode_date__year = year, pressing_electrode_status='Ok')
            pressing_electrode_status_ObsSame = ActiveMotors.objects.filter(filter_objects,pressing_electrode_date__year = year,
                                                                                pressing_electrode_status='Observation(same stage)')
            pressing_electrode_status_ObsNext = ActiveMotors.objects.filter(filter_objects,pressing_electrode_date__year = year,
                                                                                    pressing_electrode_status='Observation(next stage)')
            pressing_electrode_status_Up = ActiveMotors.objects.filter(filter_objects,pressing_electrode_date__year = year, pressing_electrode_status='Under process')
            pressing_electrode_status_Hlt = ActiveMotors.objects.filter(filter_objects,pressing_electrode_date__year = year, pressing_electrode_status='Halt')



            formation_process_status_Ok = ActiveMotors.objects.filter(filter_objects,formation_process_date__year = year, formation_process_status='Ok')
            formation_process_status_ObsSame = ActiveMotors.objects.filter(filter_objects,formation_process_date__year = year,
                                                                                formation_process_status='Observation(same stage)')
            formation_process_status_ObsNext = ActiveMotors.objects.filter(filter_objects,formation_process_date__year = year,
                                                                                    formation_process_status='Observation(next stage)')
            formation_process_status_Up = ActiveMotors.objects.filter(filter_objects,formation_process_date__year = year, formation_process_status='Under process')
            formation_process_status_Hlt = ActiveMotors.objects.filter(filter_objects,formation_process_date__year = year, formation_process_status='Halt')



            assembly_process_status_Ok = ActiveMotors.objects.filter(filter_objects,assembly_process_date__year = year, assembly_process_status='Ok')
            assembly_process_status_ObsSame = ActiveMotors.objects.filter(filter_objects,assembly_process_date__year = year,
                                                                                assembly_process_status='Observation(same stage)')
            assembly_process_status_ObsNext = ActiveMotors.objects.filter(filter_objects,assembly_process_date__year = year,
                                                                                    assembly_process_status='Observation(next stage)')
            assembly_process_status_Up = ActiveMotors.objects.filter(filter_objects,assembly_process_date__year = year, assembly_process_status='Under process')
            assembly_process_status_Hlt = ActiveMotors.objects.filter(filter_objects,assembly_process_date__year = year, assembly_process_status='Halt')


            battery_testing_status_Ok = ActiveMotors.objects.filter(filter_objects,battery_testing_date__year = year, battery_testing_status='Ok')
            battery_testing_status_ObsSame = ActiveMotors.objects.filter(filter_objects,battery_testing_date__year = year,
                                                                                battery_testing_status='Observation(same stage)')
            battery_testing_status_ObsNext = ActiveMotors.objects.filter(filter_objects,battery_testing_date__year = year,
                                                                                    battery_testing_status='Observation(next stage)')
            battery_testing_status_Up = ActiveMotors.objects.filter(filter_objects,battery_testing_date__year = year, battery_testing_status='Under process')
            battery_testing_status_Hlt = ActiveMotors.objects.filter(filter_objects,battery_testing_date__year = year, battery_testing_status='Halt')



            final_qualification_status_Ok = ActiveMotors.objects.filter(filter_objects,final_qualification_date__year = year, final_qualification_status='Ok')
            final_qualification_status_ObsSame = ActiveMotors.objects.filter(filter_objects,final_qualification_date__year = year,
                                                                                final_qualification_status='Observation(same stage)')
            final_qualification_status_ObsNext = ActiveMotors.objects.filter(filter_objects,final_qualification_date__year = year,
                                                                                    final_qualification_status='Observation(next stage)')
            final_qualification_status_Up = ActiveMotors.objects.filter(filter_objects,final_qualification_date__year = year, final_qualification_status='Under process')
            final_qualification_status_Hlt = ActiveMotors.objects.filter(filter_objects,final_qualification_date__year = year, final_qualification_status='Halt')

# ==========================Pyro Devices ==========================
            qualification_raw_material_status_Ok = ActiveMotors.objects.filter(filter_objects,qualification_raw_material_date__year = year, qualification_raw_material_status='Ok')
            qualification_raw_material_status_ObsSame = ActiveMotors.objects.filter(filter_objects,qualification_raw_material_date__year = year,
                                                                            qualification_raw_material_status='Observation(same stage)')
            qualification_raw_material_status_ObsNext = ActiveMotors.objects.filter(filter_objects,qualification_raw_material_date__year = year,
                                                                            qualification_raw_material_status='Observation(next stage)')
            qualification_raw_material_status_Up = ActiveMotors.objects.filter(filter_objects,qualification_raw_material_date__year = year,
                                                                       qualification_raw_material_status='Under process')
            qualification_raw_material_status_Hlt = ActiveMotors.objects.filter(filter_objects,qualification_raw_material_date__year = year, qualification_raw_material_status='Halt')


            filling_status_Ok = ActiveMotors.objects.filter(filter_objects,filling_date__year = year, filling_status='Ok')
            filling_status_ObsSame = ActiveMotors.objects.filter(filter_objects,filling_date__year = year,
                                                                           filling_status='Observation(same stage)')
            filling_status_ObsNext = ActiveMotors.objects.filter(filter_objects,filling_date__year = year,
                                                                           filling_status='Observation(next stage)')
            filling_status_Up = ActiveMotors.objects.filter(filter_objects,filling_date__year = year, filling_status='Under process')
            filling_status_Hlt = ActiveMotors.objects.filter(filter_objects,filling_date__year = year, filling_status='Halt')


            assembling_integration_status_Ok = ActiveMotors.objects.filter(filter_objects,assembling_integration_date__year = year, assembling_integration_status='Ok')
            assembling_integration_status_ObsSame = ActiveMotors.objects.filter(filter_objects,assembling_integration_date__year = year,
                                                                          assembling_integration_status='Observation(same stage)')
            assembling_integration_status_ObsNext = ActiveMotors.objects.filter(filter_objects,assembling_integration_date__year = year,
                                                                          assembling_integration_status='Observation(next stage)')
            assembling_integration_status_Up = ActiveMotors.objects.filter(filter_objects,assembling_integration_date__year = year, assembling_integration_status='Under process')
            assembling_integration_status_Hlt = ActiveMotors.objects.filter(filter_objects,assembling_integration_date__year = year, assembling_integration_status='Halt')


            qualification_testing_status_Ok = ActiveMotors.objects.filter(filter_objects,qualification_testing_date__year = year, qualification_testing_status='Ok')
            qualification_testing_status_ObsSame = ActiveMotors.objects.filter(filter_objects,qualification_testing_date__year = year,
                                                                         qualification_testing_status='Observation(same stage)')
            qualification_testing_status_ObsNext = ActiveMotors.objects.filter(filter_objects,qualification_testing_date__year = year,
                                                                         qualification_testing_status='Observation(next stage)')
            qualification_testing_status_Up = ActiveMotors.objects.filter(filter_objects,qualification_testing_date__year = year, qualification_testing_status='Under process')
            qualification_testing_status_Hlt = ActiveMotors.objects.filter(filter_objects,qualification_testing_date__year = year, qualification_testing_status='Halt')


            performance_testing_status_Ok = ActiveMotors.objects.filter(filter_objects,performance_testing_date__year = year, performance_testing_status='Ok')
            performance_testing_status_ObsSame = ActiveMotors.objects.filter(filter_objects,performance_testing_date__year = year,
                                                                             performance_testing_status='Observation(same stage)')
            performance_testing_status_ObsNext = ActiveMotors.objects.filter(filter_objects,performance_testing_date__year = year,
                                                                             performance_testing_status='Observation(next stage)')
            performance_testing_status_Up = ActiveMotors.objects.filter(filter_objects,performance_testing_date__year = year,
                                                                        performance_testing_status='Under process')
            performance_testing_status_Hlt = ActiveMotors.objects.filter(filter_objects,performance_testing_date__year = year, performance_testing_status='Halt')


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

                'conditioning_of_lining_Ok' : conditioning_of_lining_Ok.count(),
                'conditioning_of_lining_ObsSame' : conditioning_of_lining_ObsSame.count(),
                'conditioning_of_lining_ObsNext' : conditioning_of_lining_ObsNext.count(),
                'conditioning_of_lining_Up' : conditioning_of_lining_Up.count(),
                'conditioning_of_lining_Hlt' : conditioning_of_lining_Hlt.count(),

                'mechanical_properties_liner_Ok' : mechanical_properties_liner_Ok.count(),
                'mechanical_properties_liner_ObsSame' : mechanical_properties_liner_ObsSame.count(),
                'mechanical_properties_liner_ObsNext' : mechanical_properties_liner_ObsNext.count(),
                'mechanical_properties_liner_Up' : mechanical_properties_liner_Up.count(),
                'mechanical_properties_liner_Hlt' : mechanical_properties_liner_Hlt.count(),

                'mechanical_properties_propellant_Ok' : mechanical_properties_propellant_Ok.count(),
                'mechanical_properties_propellant_ObsSame' : mechanical_properties_propellant_ObsSame.count(),
                'mechanical_properties_propellant_ObsNext' : mechanical_properties_propellant_ObsNext.count(),
                'mechanical_properties_propellant_Up' : mechanical_properties_propellant_Up.count(),
                'mechanical_properties_propellant_Hlt' : mechanical_properties_propellant_Hlt.count(),

                'mass_liner_Ok' : mass_liner_Ok.count(),
                'mass_liner_ObsSame' : mass_liner_ObsSame.count(),
                'mass_liner_ObsNext' : mass_liner_ObsNext.count(),
                'mass_liner_Up' : mass_liner_Up.count(),
                'mass_liner_Hlt' : mass_liner_Hlt.count(),

                'mass_insulation_Ok' : mass_insulation_Ok.count(),
                'mass_insulation_ObsSame' : mass_insulation_ObsSame.count(),
                'mass_insulation_ObsNext' : mass_insulation_ObsNext.count(),
                'mass_insulation_Up' : mass_insulation_Up.count(),
                'mass_insulation_Hlt' : mass_insulation_Hlt.count(),

                'mass_propellant_Ok' : mass_propellant_Ok.count(),
                'mass_propellant_ObsSame' : mass_propellant_ObsSame.count(),
                'mass_propellant_ObsNext' : mass_propellant_ObsNext.count(),
                'mass_propellant_Up' : mass_propellant_Up.count(),
                'mass_propellant_Hlt' : mass_propellant_Hlt.count(),

                'overall_qualification_status_ObsNcr' : overall_qualification_status_ObsNcr.count(),
                'overall_qualification_status_Inprocess' : overall_qualification_status_Inprocess.count(),

                'bhd_status_Ok' : bhd_status_Ok.count(),
                'bhd_status_NotSubmit' : bhd_status_NotSubmit.count(),
                'bhd_status_ObsForwarded' : bhd_status_ObsForwarded.count(),
                'bhd_status_Inprocess' : bhd_status_Inprocess.count(),

                'qm_certification_status_Issued' : qm_certification_status_Issued.count(),
                'qm_certification_status_ObsForwarded' : qm_certification_status_ObsForwarded.count(),
                'qm_certification_status_Inprocess' : qm_certification_status_Inprocess.count(),

                'matiral_qualified_status_Ok': matiral_qualified_status_Ok.count(),
                'matiral_qualified_status_ObsSame': matiral_qualified_status_ObsSame.count(),
                'matiral_qualified_status_ObsNext': matiral_qualified_status_ObsNext.count(),
                'matiral_qualified_status_Up': matiral_qualified_status_Up.count(),
                'matiral_qualified_status_Hlt': matiral_qualified_status_Hlt.count(),

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
                'MotorTotalCount' : MotorTotalCount

            }
            return JsonResponse({'message': 'true', 'data': dist}, status=200)
        # except Exception as e:
        #     print(e)
        #     return JsonResponse({'message': 'Sorry! No Task found.'}, status=500)
