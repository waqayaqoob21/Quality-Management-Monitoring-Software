from importlib import import_module
from django.http import JsonResponse
from ams.serializer import *
from ams.models import *
from fpdf import FPDF
import xlwt
from datetime import datetime, timedelta
from django.http import FileResponse
from django.http import HttpResponse
from django.core.files.storage import FileSystemStorage

from mpm.models import ActiveMotors
from mpm.serializer import ActiveMotorSerializer


class MpmController:
    @staticmethod
    def AddActiveMotor(request):
        motorModel = ActiveMotors()
        try:
            id = request['id']
            if id == '0':
                motorModel.system = request['system']
                motorModel.motor_id = request['motor_id']
                motorModel.qualification_insulation_lining_propellant_rm = request['qualification_insulation_lining_propellant_rm']
                motorModel.qualification_insulation_lining_propellant_rm_remarks = request['qualification_insulation_lining_propellant_rm_remarks']
                motorModel.acceptance_casting = request['acceptance_casting']
                motorModel.acceptance_casting_remarks = request['acceptance_casting_remarks']
                motorModel.sandblasting = request['sandblasting']
                motorModel.sandblasting_remarks = request['sandblasting_remarks']
                motorModel.insulation = request['insulation']
                motorModel.insulation_remarks = request['insulation_remarks']
                motorModel.ut_rt_insulated_case = request['ut_rt_insulated_case']
                motorModel.ut_rt_insulated_case_remarks = request['ut_rt_insulated_case_remarks']
                motorModel.acceptance_silver_material = request['acceptance_silver_material']
                motorModel.acceptance_silver_material_remarks = request['acceptance_silver_material_remarks']
                motorModel.silver_application = request['silver_application']
                motorModel.silver_application_remarks = request['silver_application_remarks']
                motorModel.formulation_tailoring_liner_propellant = request['formulation_tailoring_liner_propellant']
                motorModel.formulation_tailoring_liner_propellant_remarks = request['formulation_tailoring_liner_propellant_remarks']
                motorModel.conditioning_raw_materials = request['conditioning_raw_materials']
                motorModel.conditioning_raw_materials_remarks = request['conditioning_raw_materials_remarks']
                motorModel.lining = request['lining']
                motorModel.lining_remarks = request['lining_remarks']
                motorModel.casting = request['casting']
                motorModel.casting_remarks = request['casting_remarks']
                motorModel.curing = request['curing']
                motorModel.curing_remarks = request['curing_remarks']
                motorModel.liner_mechanical_properties = request['liner_mechanical_properties']
                motorModel.liner_mechanical_properties_remarks = request['liner_mechanical_properties_remarks']
                motorModel.propellant_mechanical_properties = request['propellant_mechanical_properties']
                motorModel.propellant_mechanical_properties_remarks = request['propellant_mechanical_properties_remarks']
                motorModel.interface_bond_strength = request['interface_bond_strength']
                motorModel.interface_bond_strength_remarks = request['interface_bond_strength_remarks']
                motorModel.propellant_burn_rate = request['propellant_burn_rate']
                motorModel.propellant_burn_rate_remarks = request['propellant_burn_rate_remarks']
                motorModel.trimming_Propellant_grain = request['trimming_Propellant_grain']
                motorModel.trimming_Propellant_grain_remarks = request['trimming_Propellant_grain_remarks']
                motorModel.mass_liner_insulation_propellant_srm = request['mass_liner_insulation_propellant_srm']
                motorModel.mass_liner_insulation_propellant_srm_remarks = request['mass_liner_insulation_propellant_srm_remarks']
                motorModel.ut_endoscopy_rt_grain = request['ut_endoscopy_rt_grain']
                motorModel.ut_endoscopy_rt_grain_remarks = request['ut_endoscopy_rt_grain_remarks']
                motorModel.ncr_status = request['ncr_status']
                motorModel.ncr_status_remarks = request['ncr_status_remarks']
                motorModel.overall_status = request['overall_status']
                motorModel.overall_remarks = request['overall_remarks']
                if request['attachments'] != '':
                        motorModel.attachments = request['attachments']
                motorModel.save()
                return JsonResponse({'status': 'True', 'message': "Motor Created Successfully!"},
                                status=200)
            else:
                get_motor = ActiveMotors.objects.filter(id=id).first()
                if get_motor is not None:
                    get_motor.system = request['system']
                    get_motor.motor_id = request['motor_id']
                    get_motor.qualification_insulation_lining_propellant_rm = request['qualification_insulation_lining_propellant_rm']
                    get_motor.qualification_insulation_lining_propellant_rm_remarks = request['qualification_insulation_lining_propellant_rm_remarks']
                    get_motor.acceptance_casting = request['acceptance_casting']
                    get_motor.acceptance_casting_remarks = request['acceptance_casting_remarks']
                    get_motor.sandblasting = request['sandblasting']
                    get_motor.sandblasting_remarks = request['sandblasting_remarks']
                    get_motor.insulation = request['insulation']
                    get_motor.insulation_remarks = request['insulation_remarks']
                    get_motor.ut_rt_insulated_case = request['ut_rt_insulated_case']
                    get_motor.ut_rt_insulated_case_remarks = request['ut_rt_insulated_case_remarks']
                    get_motor.acceptance_silver_material = request['acceptance_silver_material']
                    get_motor.acceptance_silver_material_remarks = request['acceptance_silver_material_remarks']
                    get_motor.silver_application = request['silver_application']
                    get_motor.silver_application_remarks = request['silver_application_remarks']
                    get_motor.formulation_tailoring_liner_propellant = request['formulation_tailoring_liner_propellant']
                    get_motor.formulation_tailoring_liner_propellant_remarks = request['formulation_tailoring_liner_propellant_remarks']
                    get_motor.conditioning_raw_materials = request['conditioning_raw_materials']
                    get_motor.conditioning_raw_materials_remarks = request['conditioning_raw_materials_remarks']
                    get_motor.lining = request['lining']
                    get_motor.lining_remarks = request['lining_remarks']
                    get_motor.casting = request['casting']
                    get_motor.casting_remarks = request['casting_remarks']
                    get_motor.curing = request['curing']
                    get_motor.curing_remarks = request['curing_remarks']
                    get_motor.liner_mechanical_properties = request['liner_mechanical_properties']
                    get_motor.liner_mechanical_properties_remarks = request['liner_mechanical_properties_remarks']
                    get_motor.propellant_mechanical_properties = request['propellant_mechanical_properties']
                    get_motor.propellant_mechanical_properties_remarks = request['propellant_mechanical_properties_remarks']
                    get_motor.interface_bond_strength = request['interface_bond_strength']
                    get_motor.interface_bond_strength_remarks = request['interface_bond_strength_remarks']
                    get_motor.propellant_burn_rate = request['propellant_burn_rate']
                    get_motor.propellant_burn_rate_remarks = request['propellant_burn_rate_remarks']
                    get_motor.trimming_Propellant_grain = request['trimming_Propellant_grain']
                    get_motor.trimming_Propellant_grain_remarks = request['trimming_Propellant_grain_remarks']
                    get_motor.mass_liner_insulation_propellant_srm = request['mass_liner_insulation_propellant_srm']
                    get_motor.mass_liner_insulation_propellant_srm_remarks = request['mass_liner_insulation_propellant_srm_remarks']
                    get_motor.ut_endoscopy_rt_grain = request['ut_endoscopy_rt_grain']
                    get_motor.ut_endoscopy_rt_grain_remarks = request['ut_endoscopy_rt_grain_remarks']
                    get_motor.ncr_status = request['ncr_status']
                    get_motor.ncr_status_remarks = request['ncr_status_remarks']
                    get_motor.overall_status = request['overall_status']
                    get_motor.overall_remarks = request['overall_remarks']
                    if request['attachments'] != '':
                        get_motor.attachments = request['attachments']
                    get_motor.save()
                    return JsonResponse({'status': 'True', 'message': "Motor Updated Successfully!"},
                                status=200)
        except Exception as e:
            return JsonResponse({'status': 'False', "message": "Motor Not Saved"}, status=500)
            
    @staticmethod
    def GetActiveMotorList(request):
        try:
            data = ActiveMotors.objects.all().order_by('id')
            serializer = ActiveMotorSerializer(data, many=True)
            return JsonResponse(serializer.data, safe=False, status=200)
        except:
            return JsonResponse({'message': 'Sorry! No student found.'}, status=500)

    @staticmethod
    def DeleteActiveMotor(request, pk):
        try:
            motor = ActiveMotors.objects.get(id=pk)
            motor.delete()
            return JsonResponse({'message': 'Motor has been deleted'}, status=200)
        except:
            return JsonResponse({'message': 'Sorry! No Motor found.'}, status=500)

    @staticmethod
    def GetActiveMotorPDFList(request):
        TABLE_COL_NAMES = ("System", "Motor ID", "Qualification of raw materials of insulation, lining and propellant","Qualification of raw materials of insulation, lining and propellant remarks",
         "Acceptance of casting","Acceptance of casting remarks", "Sandblasting", "Sandblasting remarks","Insulation","Insulation remarks","UT and RT of Insulated Case","UT and RT of Insulated Case remarks",
         "Acceptance of Silver Material","Acceptance of Silver Material remarks","Silver Application","Silver Application remarks","Formulation tailoring of liner and propellant","Formulation tailoring of liner and propellant remarks",
         "Conditioning of raw materials","Conditioning of raw materials remarks","Lining","Lining remarks","Casting","Casting remarks","Curing","Curing remarks","Liner Mechanical Properties","Liner Mechanical Properties remarks",
         "Propellant Mechanical Properties","Propellant Mechanical Properties remarks","Interface bond strentgh","Interface bond strentgh remarks","Propellant burn rate","Propellant burn rate remarks","Trimming of propellant grain",
         "Trimming of propellant grain remarks","Trimming of propellant grain remarks","Mass of liner, Insulation, propellant and SRM","Mass of liner, Insulation, propellant and SRM remarks","UT, endoscopy and RT of grain",
         "UT, endoscopy and RT of grain remarks","NCR Status","NCR Status remarks", "Overall Status","Overall remarks")
        data = ActiveMotors.objects.all().order_by('-id').values_list('system','motor_id','qualification_insulation_lining_propellant_rm','qualification_insulation_lining_propellant_rm_remarks',
                                                                        'acceptance_casting','acceptance_casting_remarks','sandblasting','sandblasting_remarks','insulation', 'insulation_remarks','ut_rt_insulated_case',
                                                                        'ut_rt_insulated_case_remarks','acceptance_silver_material','acceptance_silver_material_remarks','silver_application','silver_application_remarks',
                                                                        'formulation_tailoring_liner_propellant', 'formulation_tailoring_liner_propellant_remarks','conditioning_raw_materials',
                                                                        'conditioning_raw_materials_remarks', 'lining', 'lining_remarks','casting','casting_remarks',
                                                                        'curing','curing_remarks','liner_mechanical_properties','liner_mechanical_properties_remarks',
                                                                        'propellant_mechanical_properties','propellant_mechanical_properties_remarks','interface_bond_strength',
                                                                        'interface_bond_strength_remarks','propellant_burn_rate','propellant_burn_rate_remarks','trimming_Propellant_grain',
                                                                        'trimming_Propellant_grain_remarks','mass_liner_insulation_propellant_srm','mass_liner_insulation_propellant_srm_remarks',
                                                                        'ut_endoscopy_rt_grain','ut_endoscopy_rt_grain_remarks','ncr_status','ncr_status_remarks','overall_status','overall_remarks')

        pdf = FPDF('L', 'mm', 'Legal')
        pdf.add_page()
        pdf.page_no()
        pdf.set_font('courier', 'B', 26)
        pdf.cell(330, 10, 'Final Report', border=0,align='C', ln=2)
        pdf.cell(40, 10, '',0,1)
        pdf.set_font("Times", size=10)
        line_height = pdf.font_size * 14
        col_width = pdf.epw / 24
        count_col = 0;
        def column_headings():
            pdf.set_font(style="B") 
            for col_name in TABLE_COL_NAMES:
                pdf.multi_cell(col_width, line_height,col_name, border=1,align='L', ln=3, max_line_height=pdf.font_size)
            pdf.ln(line_height)
            pdf.set_font(style="")
        column_headings()
        lh_list = []
        use_default_height = 0 
        for row in data:
            for datum in row:
                dd = str(datum)
                word_list = dd.split()
                number_of_words = len(word_list)
                if number_of_words>2:
                    use_default_height = 1
                    new_line_height = pdf.font_size * (number_of_words/2)
            if not use_default_height:
                lh_list.append(line_height)
            else:
                lh_list.append(new_line_height)
                use_default_height = 0
        # count = 0;
        for j,row in enumerate(data):
            line_height = lh_list[j] 
            if pdf.will_page_break(line_height):
                    print(line_height)
                    column_headings()
            for col_num in range(len(row)):
            
                # line_height = lh_list[j] 
                if col_num<=22:
                    pdf.multi_cell(col_width, line_height, f"{row[col_num]}", border=1,align='L',ln=3, max_line_height=pdf.font_size)
                if col_num == 23:
                    # print("Second",col_num)
                    pdf.add_page()
                if col_num >=23:
                    # print("third",col_num)
                    pdf.multi_cell(col_width, line_height, f"{row[col_num]}", border=1,align='L',ln=3, max_line_height=pdf.font_size)
                # count_col += 1;
            pdf.ln(line_height)
        pdf.output('report.pdf')
        return FileResponse(open('report.pdf', 'rb'), as_attachment=True, content_type='application/pdf')

    @staticmethod
    def GetActiveMotorExcelList(request):
        response = HttpResponse(content_type = 'application/ms-excel')
        response['Content-Disposition'] = 'attachment; filename="DocumentExcelSheet.xls"'
        work_book = xlwt.Workbook(encoding='utf-8')
        work_sheet = work_book.add_sheet('ExcelSheet')
        row_num = 0
        font_style = xlwt.XFStyle()
        font_style.font.bold = True
        TABLE_COL_NAMES = ("System", "Motor ID", "Qualification of raw materials of insulation, lining and propellant","Qualification of raw materials of insulation, lining and propellant remarks",
         "Acceptance of casting","Acceptance of casting remarks", "Sandblasting", "Sandblasting remarks","Insulation","Insulation remarks","UT and RT of Insulated Case","UT and RT of Insulated Case remarks",
         "Acceptance of Silver Material","Acceptance of Silver Material remarks","Silver Application","Silver Application remarks","Formulation tailoring of liner and propellant","Formulation tailoring of liner and propellant remarks",
         "Conditioning of raw materials","Conditioning of raw materials remarks","Lining","Lining remarks","Casting","Casting remarks","Curing","Curing remarks","Liner Mechanical Properties","Liner Mechanical Properties remarks",
         "Propellant Mechanical Properties","Propellant Mechanical Properties remarks","Interface bond strentgh","Interface bond strentgh remarks","Propellant burn rate","Propellant burn rate remarks","Trimming of propellant grain",
         "Trimming of propellant grain remarks","Trimming of propellant grain remarks","Mass of liner, Insulation, propellant and SRM","Mass of liner, Insulation, propellant and SRM remarks","UT, endoscopy and RT of grain",
         "UT, endoscopy and RT of grain remarks","NCR Status","NCR Status remarks", "Overall Status","Overall remarks")
        for col_num in range(len(TABLE_COL_NAMES)):
            work_sheet.write(row_num,col_num,TABLE_COL_NAMES[col_num], font_style)
        font_style = xlwt.XFStyle()

        data = ActiveMotors.objects.all().order_by('-id').values_list('system','motor_id','qualification_insulation_lining_propellant_rm','qualification_insulation_lining_propellant_rm_remarks',
                                                                        'acceptance_casting','acceptance_casting_remarks','sandblasting','sandblasting_remarks','insulation', 'insulation_remarks','ut_rt_insulated_case',
                                                                        'ut_rt_insulated_case_remarks','acceptance_silver_material','acceptance_silver_material_remarks','silver_application','silver_application_remarks',
                                                                        'formulation_tailoring_liner_propellant', 'formulation_tailoring_liner_propellant_remarks','conditioning_raw_materials',
                                                                        'conditioning_raw_materials_remarks', 'lining', 'lining_remarks','casting','casting_remarks',
                                                                        'curing','curing_remarks','liner_mechanical_properties','liner_mechanical_properties_remarks',
                                                                        'propellant_mechanical_properties','propellant_mechanical_properties_remarks','interface_bond_strength',
                                                                        'interface_bond_strength_remarks','propellant_burn_rate','propellant_burn_rate_remarks','trimming_Propellant_grain',
                                                                        'trimming_Propellant_grain_remarks','mass_liner_insulation_propellant_srm','mass_liner_insulation_propellant_srm_remarks',
                                                                        'ut_endoscopy_rt_grain','ut_endoscopy_rt_grain_remarks','ncr_status','ncr_status_remarks','overall_status','overall_remarks')

        for row in data:
            row_num += 1
            for col_num in range(len(row)):
                work_sheet.write(row_num,col_num,str(row[col_num]), font_style)
        work_book.save(response)
        return response