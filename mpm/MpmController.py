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
from mpm.models import ActiveMotors
from mpm.serializer import ActiveMotorSerializer
import itertools
from django.db.models import F, Q


class MpmController:
    @staticmethod
    def AddActiveMotor(request):
            motorModel = ActiveMotors()
        # try:
            id = request['id']
            if id == 0 or id == '0':
                motorModel.types = request['types']
                motorModel.system = request['system_Name']
                motorModel.motor_id = request['motor_id']
                motorModel.year = request['year']
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
                    get_motor.types = request['types']
                    get_motor.system = request['system_Name']
                    get_motor.motor_id = request['motor_id']
                    get_motor.year = request['year']
                    get_motor.component_type = request['component_type']
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
                    get_motor.ncr_status = request['ncr_status']
                    get_motor.ncr_status_remarks = request['ncr_status_ramarks']
                    get_motor.overall_status = request['overall_status']
                    get_motor.overall_remarks = request['overall_status_ramarks']
                    if request['base64File'] != '':
                        get_motor.attachments = request['base64File']
                    get_motor.save()
                    return JsonResponse({'status': 'True', 'message': "Motor Created Successfully!"},
                                        status=200)
        # except Exception as e:
        #     return JsonResponse({'status': 'False', "message": "Motor Not Saved"}, status=500)

    @staticmethod
    def GetActiveMotorList(request, self=None):
        try:
            current_status = request.query_params.get('status')
            current_year = request.query_params.get('year')
            current_sys = request.query_params.get('system')
            current_comp = request.query_params.get('component')
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
                    'year__year', 'equal',
                    current_year)

            if current_status !='':
                filter_objects &= get_filter(
                    'overall_status', 'equal', current_status)

            if current_sys !='':
                filter_objects &= get_filter(
                    'system', 'equal',
                    current_sys)
            if current_comp !='':
                filter_objects &= get_filter(
                    'component_type', 'equal',
                    current_comp)

            dataList = ActiveMotors.objects.filter(filter_objects)
            if current_year == '':
                dataList = ActiveMotors.objects.all()

            if current_status == 'Current Year SRM':
                dataList = ActiveMotors.objects.filter(year__year=current_year)

            serializer = ActiveMotorSerializer(dataList, many=True)
            return JsonResponse({'status': 'True', 'data': serializer.data},
                                status=200)
        except Exception as e:
            print(e)
            return JsonResponse({'status': 'False', "message": "Internal Server Error"}, status=500)
    @staticmethod
    def GetActiveMotorListCount(request, self=None):
        try:
            current_year = request.query_params.get('selected_year')
            current_sys = request.query_params.get('selected_system')
            current_comp = request.query_params.get('selected_component')
            totCurrentYearSRM = 0
            totCurrentYearQualified = 0
            totCurrentYearObservation = 0
            totCurrentYearInprocess = 0
            totCurrentYearOverdue = 0
            if current_year !='':
                if current_sys =='' and current_comp == '':
                    totCurrentYearSRM = ActiveMotors.objects.filter(year__year = current_year).count()
                    totCurrentYearQualified = ActiveMotors.objects.filter(year__year = current_year,overall_status = "Qualified").count()
                    totCurrentYearInprocess = ActiveMotors.objects.filter(year__year = current_year,overall_status = "In Process").count()
                    totCurrentYearObservation = ActiveMotors.objects.filter(year__year = current_year,overall_status = "Observation").count()
                    totCurrentYearOverdue = ActiveMotors.objects.filter(year__year = current_year,overall_status = "Overdue").count()

                elif current_sys =='' and current_comp !='':
                    totCurrentYearSRM = ActiveMotors.objects.filter(year__year=current_year, component_type = current_comp).count()
                    totCurrentYearQualified = ActiveMotors.objects.filter(year__year=current_year,
                                                                          overall_status="Qualified", component_type = current_comp).count()
                    totCurrentYearInprocess = ActiveMotors.objects.filter(year__year=current_year,
                                                                          overall_status="In Process", component_type = current_comp).count()
                    totCurrentYearObservation = ActiveMotors.objects.filter(year__year=current_year,
                                                                            overall_status="Observation", component_type = current_comp).count()
                    totCurrentYearOverdue = ActiveMotors.objects.filter(year__year=current_year,
                                                                        overall_status="Overdue", component_type = current_comp).count()
                elif current_sys !='' and current_comp =='':
                    totCurrentYearSRM = ActiveMotors.objects.filter(year__year=current_year, system = current_sys).count()
                    totCurrentYearQualified = ActiveMotors.objects.filter(year__year=current_year,
                                                                          overall_status="Qualified", system = current_sys).count()
                    totCurrentYearInprocess = ActiveMotors.objects.filter(year__year=current_year,
                                                                          overall_status="In Process", system = current_sys).count()
                    totCurrentYearObservation = ActiveMotors.objects.filter(year__year=current_year,
                                                                            overall_status="Observation", system = current_sys).count()
                    totCurrentYearOverdue = ActiveMotors.objects.filter(year__year=current_year,
                                                                        overall_status="Overdue", system = current_sys).count()
                else:
                    totCurrentYearSRM = ActiveMotors.objects.filter(year__year=current_year, system=current_sys).count()
                    totCurrentYearQualified = ActiveMotors.objects.filter(year__year=current_year,
                                                                          overall_status="Qualified",
                                                                          system=current_sys, component_type = current_comp).count()
                    totCurrentYearInprocess = ActiveMotors.objects.filter(year__year=current_year,
                                                                          overall_status="In Process",
                                                                          system=current_sys, component_type = current_comp).count()
                    totCurrentYearObservation = ActiveMotors.objects.filter(year__year=current_year,
                                                                            overall_status="Observation",
                                                                            system=current_sys, component_type = current_comp).count()
                    totCurrentYearOverdue = ActiveMotors.objects.filter(year__year=current_year,
                                                                        overall_status="Overdue", system=current_sys, component_type = current_comp).count()
            else:
                data = ActiveMotors.objects.all().count()
                return JsonResponse({'status': 'True', 'data': data},
                                    status=200)
            dist = {

                'currYearTotalSrm': totCurrentYearSRM,
                'currYearSrmInprocess': totCurrentYearInprocess,
                'currYearSrmQualified': totCurrentYearQualified,
                'currYearSrmOverdue': totCurrentYearOverdue,
                'currYearSrmObservation': totCurrentYearObservation,
                'currYearSrmOverdue': totCurrentYearOverdue

            }
            return JsonResponse({'status': 'True', 'data': dist},
                                status=200)
        except Exception as e:
            print(e)
            return JsonResponse({'status': 'False', "message": "Internal Server Error"}, status=500)
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
