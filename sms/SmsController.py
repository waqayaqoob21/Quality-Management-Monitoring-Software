from importlib import import_module

from django.db.models import Q
from django.http import JsonResponse
from ams.serializer import *
from ams.models import *
from fpdf import FPDF
import xlwt
from datetime import datetime, timedelta
from django.http import FileResponse
from django.http import HttpResponse
from django.core.files.storage import FileSystemStorage
import re
import fitz  # pip install PyMuPDF Pillow
import io
import os
from PIL import Image
from pytesseract import pytesseract
from csv import reader
import os
import csv
from django.db.models.functions import *
import json
from dateutil.relativedelta import relativedelta
from sms.models import FlightSystemStatus, ProductionSystemStatus, RelifingSystemStatus, ProductionSystemStatusHistory, \
    FlightSystemStatusHistory, RelifingSystemStatusHistory
from sms.serializers import FlightSystemSerialzer, ProductionSystemSerialzer, \
    RelifingSystemSerialzer  # install tesseract-ocr-w64-setup-v5.2.0.20220712.exe (64 bit) resp.
from django.db import connection

# from https://github.com/UB-Mannheim/tesseract/wiki
# pip install pytesseract

class SmsController:

    @staticmethod
    def AddProductionStatus(request):
        is_active = 0
        prodModel = ProductionSystemStatus()
        if request['is_active'] == 'true':
            is_active = 1
        try:
            id = request['id']
            if id == '0':
                prodModel.system = request['system']
                prodModel.organization = request['organization']
                prodModel.set_id = request['set_id']
                if request['blt_date'] != '':
                    prodModel.blt_date = request['blt_date']
                prodModel.testing_date = request['testing_date']
                prodModel.sys_type = request['sys_type']
                prodModel.blt_status = request['blt_status']
                prodModel.blt_remarks = request['blt_remarks']
                if request['pre_hil_date'] != '':
                    prodModel.pre_hil_date = request['pre_hil_date']
                prodModel.pre_hil_status = request['pre_hil_status']
                prodModel.pre_hil_remarks = request['pre_hil_remarks']
                if request['vibaration_date'] != '':
                    prodModel.vibaration_date = request['vibaration_date']
                prodModel.vibaration_status = request['vibaration_status']
                prodModel.vibaration_remarks = request['vibaration_remarks']
                if request['post_hil_date'] != '':
                    prodModel.post_hil_date = request['post_hil_date']
                prodModel.post_hil_status = request['post_hil_status']
                prodModel.post_hil_remarks = request['post_hil_remarks']
                if request['fgt_date'] != '':
                    prodModel.fgt_date = request['fgt_date']
                prodModel.fgt_status = request['fgt_status']
                prodModel.fgt_remarks = request['fgt_remarks']
                if request['final_integration_date'] != '':
                    prodModel.final_integration_date = request['final_integration_date']
                prodModel.final_integration_status = request['final_integration_status']
                prodModel.final_integration_remarks = request['final_integration_remarks']
                print(prodModel.final_integration_remarks)
                if request['bhd_date'] != '':
                    prodModel.bhd_date = request['bhd_date']
                    prodModel.bhd_date = request['bhd_date']
                prodModel.bhd_status = request['bhd_status']
                prodModel.bhd_remarks = request['bhd_remarks']
                if request['fqm_date'] != '':
                    prodModel.fqm_date = request['fqm_date']
                prodModel.fqm_status = request['fqm_status']
                prodModel.fqm_remarks = request['fqm_remarks']
                if request['qm_certification_date'] != '':
                    prodModel.qm_certification_date = request['qm_certification_date']
                prodModel.qm_certification_status = request['qm_certification_status']
                prodModel.qm_certification_remarks = request['qm_certification_remarks']
                prodModel.attachment = request['attachment']
                prodModel.remarks = request['remarks']
                if request['cgbalancing_date'] != '':
                    prodModel.cgbalancing_date = request['cgbalancing_date']
                prodModel.cgbalancing_date_status = request['cgbalancing_date_status']
                prodModel.cgbalancing_date_remarks = request['cgbalancing_date_remarks']
                if request['enduser_date'] != '':
                    prodModel.enduser_date = request['enduser_date']
                prodModel.enduser_status = request['enduser_status']
                prodModel.enduser_remarks = request['enduser_remarks']
                if request['incapsulation_date'] != '':
                    prodModel.incapsulation_date = request['incapsulation_date']
                prodModel.incapsulation_status = request['incapsulation_status']
                prodModel.incapsulation_remarks = request['incapsulation_remarks']
                if request['sys_align_Date'] != '':
                    prodModel.sys_align_Date = request['sys_align_Date']
                prodModel.sys_align_status = request['sys_align_status']
                prodModel.sys_align_remarks = request['sys_align_remarks']
                prodModel.testing_type = request['testing_type']
                prodModel.emp_proofing = request['emp_proofing']
                prodModel.func_tst = request['func_tst']
                prodModel.func_tst_dummy_bird = request['func_tst_dummy_bird']
                prodModel.road_test = request['road_test']
                prodModel.post_road_test = request['post_road_test']
                prodModel.integrated_operation = request['integrated_operation']
                prodModel.rain_test = request['rain_test']
                prodModel.pre_user_inspection = request['pre_user_inspection']
                prodModel.final_integrated_testing = request['final_integrated_testing']
                prodModel.load_unload_on_mlv_hlf = request['load_unload_on_mlv_hlf']

                if request['emp_proofing_date'] != '':
                    prodModel.emp_proofing_date = request['emp_proofing_date']
                prodModel.emp_proofing_remarks = request['emp_proofing_remarks']
                if request['func_tst_date'] != '':
                    prodModel.func_tst_date = request['func_tst_date']
                prodModel.func_tst_remarks = request['func_tst_remarks']
                if request['func_tst_dummy_bird_date'] != '':
                    prodModel.func_tst_dummy_bird_date = request['func_tst_dummy_bird_date']
                prodModel.func_tst_dummy_bird_remarks = request['func_tst_dummy_bird_remarks']
                if request['road_test_date'] != '':
                    prodModel.road_test_date = request['road_test_date']
                prodModel.road_test_remarks = request['road_test_remarks']
                if request['post_road_test_date'] != '':
                    prodModel.post_road_test_date = request['post_road_test_date']
                prodModel.post_road_test_remarks = request['post_road_test_remarks']
                if request['integrated_operation_date'] != '':
                    prodModel.integrated_operation_date = request['integrated_operation_date']
                prodModel.integrated_operation_remarks = request['integrated_operation_remarks']
                if request['rain_test_date'] != '':
                    prodModel.rain_test_date = request['rain_test_date']
                prodModel.rain_test_remarks = request['rain_test_remarks']
                if request['pre_user_inspection_date'] != '':
                    prodModel.pre_user_inspection_date = request['pre_user_inspection_date']
                prodModel.pre_user_inspection_remarks = request['pre_user_inspection_remarks']
                if request['final_integrated_testing_date'] != '':
                    prodModel.final_integrated_testing_date = request['final_integrated_testing_date']
                prodModel.final_integrated_testing_remarks = request['final_integrated_testing_remarks']
                if request['load_unload_on_mlv_hlf_date'] != '':
                    prodModel.load_unload_on_mlv_hlf_date = request['load_unload_on_mlv_hlf_date']
                prodModel.load_unload_on_mlv_hlf_remarks = request['load_unload_on_mlv_hlf_remarks']
                if request['user_id'] != '':
                    prodModel.user_id = request['user_id']
                else:
                    prodModel.user_id = None
                prodModel.isActive = is_active
                prodModel.save()
                return JsonResponse({'status': 'True', 'message': "Production Status Added Successfully!"},
                                    status=200)
            else:
                get_prod = ProductionSystemStatus.objects.filter(id=id).first()
                if get_prod is not None:
                    if request['blt_status'] != get_prod.blt_status \
                            or request['blt_date'] != get_prod.blt_date \
                            or request['blt_remarks'] != get_prod.blt_remarks \
                            or request['pre_hil_status'] != get_prod.pre_hil_status \
                            or request['pre_hil_date'] != get_prod.pre_hil_date \
                            or request['pre_hil_remarks'] != get_prod.pre_hil_remarks \
                            or request['vibaration_status'] != get_prod.vibaration_status \
                            or request['vibaration_date'] != get_prod.vibaration_date \
                            or request['vibaration_remarks'] != get_prod.vibaration_remarks \
                            or request['post_hil_status'] != get_prod.post_hil_status \
                            or request['post_hil_date'] != get_prod.post_hil_date \
                            or request['post_hil_remarks'] != get_prod.post_hil_remarks \
                            or request['fgt_status'] != get_prod.fgt_status \
                            or request['fgt_date'] != get_prod.fgt_date \
                            or request['fgt_remarks'] != get_prod.fgt_remarks \
                            or request['final_integration_status'] != get_prod.final_integration_status \
                            or request['final_integration_date'] != get_prod.final_integration_date \
                            or request['final_integration_remarks'] != get_prod.final_integration_remarks \
                            or request['bhd_status'] != get_prod.bhd_status \
                            or request['bhd_date'] != get_prod.bhd_date \
                            or request['bhd_remarks'] != get_prod.bhd_remarks \
                            or request['fqm_status'] != get_prod.fqm_status \
                            or request['fqm_date'] != get_prod.fqm_date \
                            or request['fqm_remarks'] != get_prod.fqm_remarks \
                            or request['qm_certification_status'] != get_prod.qm_certification_status \
                            or request['qm_certification_date'] != get_prod.qm_certification_date \
                            or request['qm_certification_remarks'] != get_prod.qm_certification_remarks \
                            or request['incapsulation_status'] != get_prod.incapsulation_status \
                            or request['incapsulation_date'] != get_prod.incapsulation_date \
                            or request['incapsulation_remarks'] != get_prod.incapsulation_remarks \
                            or request['sys_align_status'] != get_prod.sys_align_status \
                            or request['sys_align_Date'] != get_prod.sys_align_Date \
                            or request['sys_align_remarks'] != get_prod.sys_align_remarks \
                            or request['emp_proofing'] != get_prod.emp_proofing \
                            or request['emp_proofing_date'] != get_prod.emp_proofing_date \
                            or request['emp_proofing_remarks'] != get_prod.emp_proofing_remarks \
                            or request['func_tst'] != get_prod.func_tst \
                            or request['func_tst_date'] != get_prod.func_tst_date \
                            or request['func_tst_remarks'] != get_prod.func_tst_remarks \
                            or request['func_tst_dummy_bird'] != get_prod.func_tst_dummy_bird \
                            or request['func_tst_dummy_bird_date'] != get_prod.func_tst_dummy_bird_date \
                            or request['func_tst_dummy_bird_remarks'] != get_prod.func_tst_dummy_bird_remarks \
                            or request['road_test'] != get_prod.road_test \
                            or request['road_test_date'] != get_prod.road_test_date \
                            or request['road_test_remarks'] != get_prod.road_test_remarks \
                            or request['post_road_test'] != get_prod.post_road_test \
                            or request['post_road_test_date'] != get_prod.post_road_test_date \
                            or request['post_road_test_remarks'] != get_prod.post_road_test_remarks \
                            or request['integrated_operation'] != get_prod.integrated_operation \
                            or request['integrated_operation_date'] != get_prod.integrated_operation_date \
                            or request['integrated_operation_remarks'] != get_prod.integrated_operation_remarks \
                            or request['rain_test'] != get_prod.rain_test \
                            or request['rain_test_date'] != get_prod.rain_test_date \
                            or request['rain_test_remarks'] != get_prod.rain_test_remarks \
                            or request['pre_user_inspection'] != get_prod.pre_user_inspection \
                            or request['pre_user_inspection_date'] != get_prod.pre_user_inspection_date \
                            or request['pre_user_inspection_remarks'] != get_prod.pre_user_inspection_remarks \
                            or request['final_integrated_testing'] != get_prod.final_integrated_testing \
                            or request['final_integrated_testing_date'] != get_prod.final_integrated_testing_date \
                            or request['final_integrated_testing_remarks'] != get_prod.final_integrated_testing_remarks \
                            or request['load_unload_on_mlv_hlf'] != get_prod.load_unload_on_mlv_hlf \
                            or request['load_unload_on_mlv_hlf_date'] != get_prod.load_unload_on_mlv_hlf_date \
                            or request['load_unload_on_mlv_hlf_remarks'] != get_prod.load_unload_on_mlv_hlf_remarks \
                            or request['launchact_status'] != get_prod.launchact_status \
                            or request['enduser_status'] != get_prod.enduser_status \
                            or request['enduser_date'] != get_prod.enduser_date \
                            or request['enduser_remarks'] != get_prod.enduser_remarks:
                        ProdHistoryModal = ProductionSystemStatusHistory()
                        ProdHistoryModal.prod_id = get_prod.id
                        ProdHistoryModal.system = get_prod.system
                        ProdHistoryModal.organization = get_prod.organization
                        ProdHistoryModal.set_id = get_prod.set_id
                        if request['blt_date'] != '':
                            ProdHistoryModal.blt_date = get_prod.blt_date
                        ProdHistoryModal.blt_status = get_prod.blt_status
                        ProdHistoryModal.blt_remarks = get_prod.blt_remarks
                        ProdHistoryModal.testing_date = get_prod.testing_date
                        ProdHistoryModal.sys_type = get_prod.sys_type
                        if request['pre_hil_date'] != '':
                            ProdHistoryModal.pre_hil_date = get_prod.pre_hil_date
                        ProdHistoryModal.pre_hil_status = get_prod.pre_hil_status
                        ProdHistoryModal.pre_hil_remarks = get_prod.pre_hil_remarks
                        if request['vibaration_date'] != '':
                            ProdHistoryModal.vibaration_date = get_prod.vibaration_date
                        ProdHistoryModal.vibaration_status = get_prod.vibaration_status
                        ProdHistoryModal.vibaration_remarks = get_prod.vibaration_remarks
                        if request['post_hil_date'] != '':
                            ProdHistoryModal.post_hil_date = get_prod.post_hil_date
                        ProdHistoryModal.post_hil_status = get_prod.post_hil_status
                        ProdHistoryModal.post_hil_remarks = get_prod.post_hil_remarks
                        if request['fgt_date'] != '':
                            ProdHistoryModal.fgt_date = get_prod.fgt_date
                        ProdHistoryModal.fgt_status = get_prod.fgt_status
                        ProdHistoryModal.fgt_remarks = get_prod.fgt_remarks
                        if request['final_integration_date'] != '':
                            ProdHistoryModal.final_integration_date = get_prod.final_integration_date
                        ProdHistoryModal.final_integration_status = get_prod.final_integration_status
                        ProdHistoryModal.final_integration_remarks = get_prod.final_integration_remarks
                        if request['bhd_date'] != '':
                            ProdHistoryModal.bhd_date = get_prod.bhd_date
                        ProdHistoryModal.bhd_status = get_prod.bhd_status
                        ProdHistoryModal.bhd_remarks = get_prod.bhd_remarks
                        if request['fqm_date'] != '':
                            ProdHistoryModal.fqm_date = get_prod.fqm_date
                        ProdHistoryModal.fqm_status = get_prod.fqm_status
                        ProdHistoryModal.fqm_remarks = get_prod.fqm_remarks
                        if request['qm_certification_date'] != '':
                            ProdHistoryModal.qm_certification_date = get_prod.qm_certification_date
                        ProdHistoryModal.qm_certification_status = get_prod.qm_certification_status
                        ProdHistoryModal.qm_certification_remarks = get_prod.qm_certification_remarks
                        ProdHistoryModal.attachment = get_prod.attachment
                        ProdHistoryModal.remarks = get_prod.remarks
                        ProdHistoryModal.isActive = get_prod.isActive
                        if request['cgbalancing_date'] != '':
                            ProdHistoryModal.cgbalancing_date = get_prod.cgbalancing_date
                        ProdHistoryModal.cgbalancing_date_status = get_prod.cgbalancing_date_status
                        ProdHistoryModal.cgbalancing_date_remarks = get_prod.cgbalancing_date_remarks
                        if request['enduser_date'] != '':
                            ProdHistoryModal.enduser_date = get_prod.enduser_date
                        ProdHistoryModal.enduser_status = get_prod.enduser_status
                        ProdHistoryModal.enduser_remarks = get_prod.enduser_remarks
                        if request['incapsulation_date'] != '':
                            ProdHistoryModal.incapsulation_date = get_prod.incapsulation_date
                        ProdHistoryModal.incapsulation_status = get_prod.incapsulation_status
                        ProdHistoryModal.incapsulation_remarks = get_prod.incapsulation_remarks
                        if request['sys_align_Date'] != '':
                            ProdHistoryModal.sys_align_Date = get_prod.sys_align_Date
                        ProdHistoryModal.sys_align_status = get_prod.sys_align_status
                        ProdHistoryModal.sys_align_remarks = get_prod.sys_align_remarks
                        ProdHistoryModal.testing_type = get_prod.testing_type
                        ProdHistoryModal.emp_proofing = get_prod.emp_proofing
                        ProdHistoryModal.func_tst = get_prod.func_tst
                        ProdHistoryModal.func_tst_dummy_bird = get_prod.func_tst_dummy_bird
                        ProdHistoryModal.road_test = get_prod.road_test
                        ProdHistoryModal.post_road_test = get_prod.post_road_test
                        ProdHistoryModal.integrated_operation = get_prod.integrated_operation
                        ProdHistoryModal.rain_test = get_prod.rain_test
                        ProdHistoryModal.pre_user_inspection = get_prod.pre_user_inspection
                        ProdHistoryModal.final_integrated_testing = get_prod.final_integrated_testing
                        ProdHistoryModal.load_unload_on_mlv_hlf = get_prod.load_unload_on_mlv_hlf
                        if request['emp_proofing_date'] != '':
                            ProdHistoryModal.emp_proofing_date = get_prod.emp_proofing_date
                        ProdHistoryModal.emp_proofing_remarks = get_prod.emp_proofing_remarks
                        if request['func_tst_date'] != '':
                            ProdHistoryModal.func_tst_date = get_prod.func_tst_date
                        ProdHistoryModal.func_tst_remarks = get_prod.func_tst_remarks
                        if request['func_tst_dummy_bird_date'] != '':
                            ProdHistoryModal.func_tst_dummy_bird_date = get_prod.func_tst_dummy_bird_date
                        ProdHistoryModal.func_tst_dummy_bird_remarks = get_prod.func_tst_dummy_bird_remarks
                        if request['road_test_date'] != '':
                            ProdHistoryModal.road_test_date = get_prod.road_test_date
                        ProdHistoryModal.road_test_remarks = get_prod.road_test_remarks
                        if request['post_road_test_date'] != '':
                            ProdHistoryModal.post_road_test_date = get_prod.post_road_test_date
                        ProdHistoryModal.post_road_test_remarks = get_prod.post_road_test_remarks
                        if request['integrated_operation_date'] != '':
                            ProdHistoryModal.integrated_operation_date = get_prod.integrated_operation_date
                        ProdHistoryModal.integrated_operation_remarks = get_prod.integrated_operation_remarks
                        if request['rain_test_date'] != '':
                            ProdHistoryModal.rain_test_date = get_prod.rain_test_date
                        ProdHistoryModal.rain_test_remarks = get_prod.rain_test_remarks
                        if request['pre_user_inspection_date'] != '':
                            ProdHistoryModal.pre_user_inspection_date = get_prod.pre_user_inspection_date
                        ProdHistoryModal.pre_user_inspection_remarks = get_prod.pre_user_inspection_remarks
                        if request['final_integrated_testing_date'] != '':
                            ProdHistoryModal.final_integrated_testing_date = get_prod.final_integrated_testing_date
                        ProdHistoryModal.final_integrated_testing_remarks = get_prod.final_integrated_testing_remarks
                        if request['load_unload_on_mlv_hlf_date'] != '':
                            ProdHistoryModal.load_unload_on_mlv_hlf_date = get_prod.load_unload_on_mlv_hlf_date
                        ProdHistoryModal.load_unload_on_mlv_hlf_remarks = get_prod.load_unload_on_mlv_hlf_remarks
                        if request['user_id'] != '':
                            ProdHistoryModal.user_id = get_prod.user_id
                        else:
                            ProdHistoryModal.user_id = None
                        ProdHistoryModal.save()

                        # history saved
                    get_prod.system = request['system']
                    get_prod.organization = request['organization']
                    get_prod.set_id = request['set_id']
                    if request['blt_date'] != '':
                        get_prod.blt_date = request['blt_date']
                    else:
                        get_prod.blt_date = None
                    get_prod.testing_date = request['testing_date']
                    get_prod.sys_type = request['sys_type']
                    get_prod.blt_status = request['blt_status']
                    get_prod.blt_remarks = request['blt_remarks']
                    if request['pre_hil_date'] != '':
                        get_prod.pre_hil_date = request['pre_hil_date']
                    else:
                        get_prod.pre_hil_date = None
                    get_prod.pre_hil_status = request['pre_hil_status']
                    get_prod.pre_hil_remarks = request['pre_hil_remarks']
                    if request['vibaration_date'] != '':
                        get_prod.vibaration_date = request['vibaration_date']
                    else:
                        get_prod.vibaration_date = None
                    get_prod.vibaration_status = request['vibaration_status']
                    get_prod.vibaration_remarks = request['vibaration_remarks']
                    if request['post_hil_date'] != '':
                        get_prod.post_hil_date = request['post_hil_date']
                    else:
                        get_prod.post_hil_date = None
                    get_prod.post_hil_status = request['post_hil_status']
                    get_prod.post_hil_remarks = request['post_hil_remarks']
                    if request['fgt_date'] != '':
                        get_prod.fgt_date = request['fgt_date']
                    else:
                        get_prod.fgt_date = None
                    get_prod.fgt_status = request['fgt_status']
                    get_prod.fgt_remarks = request['fgt_remarks']
                    if request['final_integration_date'] != '':
                        get_prod.final_integration_date = request['final_integration_date']
                    else:
                        get_prod.final_integration_date = None
                    get_prod.final_integration_status = request['final_integration_status']
                    get_prod.final_integration_remarks = request['final_integration_remarks']
                    if request['bhd_date'] != '':
                        get_prod.bhd_date = request['bhd_date']
                    else:
                        get_prod.bhd_date = None
                    get_prod.bhd_status = request['bhd_status']
                    get_prod.bhd_remarks = request['bhd_remarks']
                    if request['fqm_date'] != '':
                        get_prod.fqm_date = request['fqm_date']
                    else:
                        get_prod.fqm_date = None
                    get_prod.fqm_status = request['fqm_status']
                    get_prod.fqm_remarks = request['fqm_remarks']
                    if request['qm_certification_date'] != '':
                        get_prod.qm_certification_date = request['qm_certification_date']
                    else:
                        get_prod.qm_certification_date = None
                    get_prod.qm_certification_status = request['qm_certification_status']
                    get_prod.qm_certification_remarks = request['qm_certification_remarks']
                    get_prod.attachment = request['attachment']
                    get_prod.remarks = request['remarks']
                    if request['cgbalancing_date'] != '':
                        get_prod.cgbalancing_date = request['cgbalancing_date']
                    else:
                        get_prod.cgbalancing_date = None
                    get_prod.cgbalancing_date_status = request['cgbalancing_date_status']
                    get_prod.cgbalancing_date_remarks = request['cgbalancing_date_remarks']
                    if request['enduser_date'] != '':
                        get_prod.enduser_date = request['enduser_date']
                    get_prod.enduser_status = request['enduser_status']
                    get_prod.enduser_remarks = request['enduser_remarks']
                    if request['incapsulation_date'] != '':
                        get_prod.incapsulation_date = request['incapsulation_date']
                    get_prod.incapsulation_status = request['incapsulation_status']
                    get_prod.incapsulation_remarks = request['incapsulation_remarks']
                    if request['sys_align_Date'] != '':
                        get_prod.sys_align_Date = request['sys_align_Date']
                    get_prod.sys_align_status = request['sys_align_status']
                    get_prod.sys_align_remarks = request['sys_align_remarks']
                    get_prod.testing_type = request['testing_type']
                    get_prod.emp_proofing = request['emp_proofing']
                    get_prod.func_tst = request['func_tst']
                    get_prod.func_tst_dummy_bird = request['func_tst_dummy_bird']
                    get_prod.road_test = request['road_test']
                    get_prod.integrated_operation = request['integrated_operation']
                    get_prod.rain_test = request['rain_test']
                    get_prod.pre_user_inspection = request['pre_user_inspection']
                    get_prod.final_integrated_testing = request['final_integrated_testing']
                    get_prod.load_unload_on_mlv_hlf = request['load_unload_on_mlv_hlf']

                    if request['emp_proofing_date'] != '':
                        get_prod.emp_proofing_date = request['emp_proofing_date']
                    get_prod.emp_proofing_remarks = request['emp_proofing_remarks']
                    if request['func_tst_date'] != '':
                        get_prod.func_tst_date = request['func_tst_date']
                    get_prod.func_tst_remarks = request['func_tst_remarks']
                    if request['func_tst_dummy_bird_date'] != '':
                        get_prod.func_tst_dummy_bird_date = request['func_tst_dummy_bird_date']
                    get_prod.func_tst_dummy_bird_remarks = request['func_tst_dummy_bird_remarks']
                    if request['road_test_date'] != '':
                        get_prod.road_test_date = request['road_test_date']
                    get_prod.road_test_remarks = request['road_test_remarks']
                    if request['post_road_test_date'] != '':
                        get_prod.post_road_test_date = request['post_road_test_date']
                    get_prod.post_road_test = request['post_road_test']
                    get_prod.post_road_test_remarks = request['post_road_test_remarks']
                    if request['integrated_operation_date'] != '':
                        get_prod.integrated_operation_date = request['integrated_operation_date']
                    get_prod.integrated_operation_remarks = request['integrated_operation_remarks']
                    if request['rain_test_date'] != '':
                        get_prod.rain_test_date = request['rain_test_date']
                    get_prod.rain_test_remarks = request['rain_test_remarks']
                    if request['pre_user_inspection_date'] != '':
                        get_prod.pre_user_inspection_date = request['pre_user_inspection_date']
                    get_prod.pre_user_inspection_remarks = request['pre_user_inspection_remarks']
                    if request['final_integrated_testing_date'] != '':
                        get_prod.final_integrated_testing_date = request['final_integrated_testing_date']
                    get_prod.final_integrated_testing_remarks = request['final_integrated_testing_remarks']
                    if request['load_unload_on_mlv_hlf_date'] != '':
                        get_prod.load_unload_on_mlv_hlf_date = request['load_unload_on_mlv_hlf_date']
                    get_prod.load_unload_on_mlv_hlf_remarks = request['load_unload_on_mlv_hlf_remarks']
                    if request['user_id'] !='':
                        get_prod.user_id = request['user_id']
                    else:
                        get_prod.user_id = None
                    get_prod.isActive = is_active
                    get_prod.save()
            return JsonResponse({'status': 'True', 'message': "Production Status Updated Successfully!"},
                                status=200)
        except Exception as e:
            return JsonResponse({'status': 'False', "message": "Status Not Saved"}, status=500)

    @staticmethod
    def ImportProductionCsv(request):
        try:
            id = request['id']
            importedCsvFile = request["csv_file"]
            system_type = request["system_type"]
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
                if system_type == "Ballistic":
                    for system, sys_type, testing_type, organization, set_id, blt_date, blt_status, blt_remarks, pre_hil_date, pre_hil_status, pre_hil_remarks,\
                            vibration_date, vibration_status, vibration_remarks, cg_balancing_date, cg_balancing_status, cg_balancing_remarks, post_hil_date, post_hil_status, post_hil_remarks,\
                            final_integrated_testing_date, final_integrated_testing_status, final_integrated_testing_remarks, bhd_date, bhd_status, bhd_remarks, fqm_date, fqm_status, fqm_remarks,\
                            qm_certification_date, qm_certification_status, qm_certification_remarks, end_user_date, end_user_status, end_user_remarks, remarks, *__ in csvf:

                        prod_system = ProductionSystemStatus(system=system,sys_type=sys_type,testing_type=testing_type,organization=organization,set_id=set_id,blt_date= datetime.strptime(blt_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),blt_status=blt_status,blt_remarks=blt_remarks,
                            pre_hil_date= datetime.strptime(pre_hil_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),pre_hil_status=pre_hil_status,pre_hil_remarks=pre_hil_remarks,vibaration_date= datetime.strptime(vibration_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),vibaration_status=vibration_status,vibaration_remarks=vibration_remarks,
                            cgbalancing_date=  datetime.strptime(cg_balancing_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),cgbalancing_date_status=cg_balancing_status,cgbalancing_date_remarks=cg_balancing_remarks,post_hil_date= datetime.strptime(post_hil_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),post_hil_status=post_hil_status,post_hil_remarks=post_hil_remarks,
                            final_integrated_testing_date= datetime.strptime(final_integrated_testing_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),final_integrated_testing=final_integrated_testing_status,final_integrated_testing_remarks=final_integrated_testing_remarks,
                            bhd_date= datetime.strptime(bhd_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),bhd_status=bhd_status,bhd_remarks=bhd_remarks,fqm_date= datetime.strptime(fqm_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),fqm_status=fqm_status,fqm_remarks=fqm_remarks,
                            qm_certification_date= datetime.strptime(qm_certification_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),qm_certification_status=qm_certification_status,qm_certification_remarks=qm_certification_remarks,
                            enduser_date= datetime.strptime(end_user_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),enduser_status=end_user_status, enduser_remarks=end_user_remarks,remarks=remarks,user_id= user_id)
                        data.append(prod_system)
                    ProductionSystemStatus.objects.bulk_create(data)
                    os.remove(file_path)
                elif system_type == "Cruise":
                    for system, sys_type, testing_type, organization, set_id, blt_date, blt_status, blt_remarks, pre_hil_date, pre_hil_status, pre_hil_remarks, vibration_date, vibration_status, vibration_remarks,\
                            cg_balancing_date, cg_balancing_status, cg_balancing_remarks, post_hil_date, post_hil_status, post_hil_remarks,sys_align_date, sys_align_status, sys_align_remarks, fgt_date, fgt_status, fgt_remarks,\
                            encapsulation_date, encapsulation_status, encapsulation_remarks, final_integrated_testing_date, final_integrated_testing_status, final_integrated_testing_remarks, bhd_date, bhd_status, bhd_remarks,\
                            fqm_date, fqm_status, fqm_remarks, qm_certification_date, qm_certification_status, qm_certification_remarks, end_user_date, end_user_status, end_user_remarks, remarks, *__ in csvf:
                        prod_system = ProductionSystemStatus(system=system,sys_type=sys_type,testing_type=testing_type,organization=organization,
                            set_id=set_id,blt_date= datetime.strptime(blt_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),blt_status=blt_status,blt_remarks=blt_remarks,pre_hil_date= datetime.strptime(pre_hil_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),pre_hil_status=pre_hil_status,
                            pre_hil_remarks=pre_hil_remarks,vibaration_date= datetime.strptime(vibration_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),vibaration_status=vibration_status,vibaration_remarks=vibration_remarks,
                            cgbalancing_date=  datetime.strptime(cg_balancing_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),cgbalancing_date_status=cg_balancing_status,cgbalancing_date_remarks=cg_balancing_remarks,post_hil_date= datetime.strptime(post_hil_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),post_hil_status=post_hil_status,post_hil_remarks=post_hil_remarks,
                            sys_align_Date=  datetime.strptime(sys_align_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"), sys_align_status= sys_align_status, sys_align_remarks= sys_align_remarks,
                            fgt_date=  datetime.strptime(fgt_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"), fgt_status= fgt_status, fgt_remarks=fgt_remarks, incapsulation_date=  datetime.strptime(encapsulation_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"), incapsulation_status=encapsulation_status, incapsulation_remarks=encapsulation_remarks,
                            final_integrated_testing_date= datetime.strptime(final_integrated_testing_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),final_integrated_testing=final_integrated_testing_status,final_integrated_testing_remarks=final_integrated_testing_remarks,
                            bhd_date= datetime.strptime(bhd_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),bhd_status=bhd_status,bhd_remarks=bhd_remarks,fqm_date= datetime.strptime(fqm_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),fqm_status=fqm_status,fqm_remarks=fqm_remarks,
                            qm_certification_date= datetime.strptime(qm_certification_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),qm_certification_status=qm_certification_status,qm_certification_remarks=qm_certification_remarks,enduser_date= datetime.strptime(end_user_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),enduser_status=end_user_status, enduser_remarks=end_user_remarks,remarks=remarks,user_id= user_id)
                        data.append(prod_system)
                    ProductionSystemStatus.objects.bulk_create(data)
                    os.remove(file_path)

                elif system_type =="Launch System":
                    for system, sys_type, testing_type, organization, set_id, blt_date, blt_status, blt_remarks,emp_proofing_date,emp_proofing_status,emp_proofing_remarks,\
                        func_test_date,func_test_status,func_test_remarks,func_test_dummy_date,func_test_dummy_status,func_test_dummy_remarks,road_test_date,road_test_status,road_test_remarks,\
                        post_road_test_date,post_road_test_status,post_road_test_remarks,integrated_operation_date,integrated_operation_status,integrated_operation_remarks,\
                        rain_test_date,rain_test_status,rain_test_remarks,pre_user_inspection_date,pre_user_inspection_status,\
                        pre_user_inspection_remarks,final_integrated_testing_date, final_integrated_testing_status, final_integrated_testing_remarks, bhd_date, bhd_status, bhd_remarks, fqm_date, fqm_status, fqm_remarks, qm_certification_date, qm_certification_status, qm_certification_remarks, end_user_date, end_user_status, end_user_remarks, remarks, *__ in csvf:

                        prod_system = ProductionSystemStatus(system=system,sys_type=sys_type,testing_type=testing_type,organization=organization,set_id=set_id,blt_date= datetime.strptime(blt_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),blt_status=blt_status,blt_remarks=blt_remarks,
                            emp_proofing_date=  datetime.strptime(emp_proofing_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),emp_proofing= emp_proofing_status,emp_proofing_remarks= emp_proofing_remarks,
                            func_tst_date=  datetime.strptime(func_test_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),func_tst= func_test_status,func_tst_remarks= func_test_remarks,
                            func_tst_dummy_bird= func_test_dummy_status,func_tst_dummy_bird_date= datetime.strptime(func_test_dummy_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),func_tst_dummy_bird_remarks= func_test_dummy_remarks,
                            road_test= road_test_status,road_test_date=  datetime.strptime(road_test_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),road_test_remarks= road_test_remarks,
                            post_road_test_date=  datetime.strptime(post_road_test_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),post_road_test= post_road_test_status,post_road_test_remarks= post_road_test_remarks,
                            integrated_operation=integrated_operation_status, integrated_operation_date= datetime.strptime(integrated_operation_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),integrated_operation_remarks= integrated_operation_remarks,
                            rain_test=rain_test_status, rain_test_date=  datetime.strptime(rain_test_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),rain_test_remarks= rain_test_remarks,
                            pre_user_inspection_date=  datetime.strptime(pre_user_inspection_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),pre_user_inspection= pre_user_inspection_status, pre_user_inspection_remarks= pre_user_inspection_remarks,
                            final_integrated_testing_date= datetime.strptime(final_integrated_testing_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),final_integrated_testing=final_integrated_testing_status,final_integrated_testing_remarks=final_integrated_testing_remarks,
                            bhd_date= datetime.strptime(bhd_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),bhd_status=bhd_status,bhd_remarks=bhd_remarks,fqm_date= datetime.strptime(fqm_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),fqm_status=fqm_status,fqm_remarks=fqm_remarks,
                            qm_certification_date= datetime.strptime(qm_certification_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),qm_certification_status=qm_certification_status,qm_certification_remarks=qm_certification_remarks,
                            enduser_date= datetime.strptime(end_user_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),enduser_status=end_user_status, enduser_remarks=end_user_remarks,remarks=remarks,user_id= user_id)
                        data.append(prod_system)
                    ProductionSystemStatus.objects.bulk_create(data)
                    os.remove(file_path)

                elif system_type == "Launch Tube":
                    for system, sys_type, testing_type, organization, set_id, blt_date, blt_status, blt_remarks, pre_user_inspection_date,pre_user_inspection_status,pre_user_inspection_remarks,\
                            load_unload_date,load_unload_status,load_unload_remarks,encapsulation_date,encapsulation_status,encapsulation_remarks, final_integrated_testing_date, final_integrated_testing_status, final_integrated_testing_remarks,\
                            bhd_date, bhd_status, bhd_remarks, fqm_date, fqm_status, fqm_remarks, qm_certification_date, qm_certification_status, qm_certification_remarks, end_user_date, end_user_status, end_user_remarks, remarks, *__ in csvf:

                        prod_system = ProductionSystemStatus(system=system,sys_type=sys_type,testing_type=testing_type,organization=organization,set_id=set_id,blt_date= datetime.strptime(blt_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),blt_status=blt_status,blt_remarks=blt_remarks,
                            pre_user_inspection_date=  datetime.strptime(pre_user_inspection_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),pre_user_inspection= pre_user_inspection_status,pre_user_inspection_remarks= pre_user_inspection_remarks,
                            load_unload_on_mlv_hlf_date=  datetime.strptime(load_unload_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),load_unload_on_mlv_hlf= load_unload_status,load_unload_on_mlv_hlf_remarks=load_unload_remarks,incapsulation_date= datetime.strptime(encapsulation_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),incapsulation_status= encapsulation_status,incapsulation_remarks= encapsulation_remarks,
                            final_integrated_testing_date= datetime.strptime(final_integrated_testing_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),final_integrated_testing=final_integrated_testing_status,final_integrated_testing_remarks=final_integrated_testing_remarks,bhd_date= datetime.strptime(bhd_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),bhd_status=bhd_status,bhd_remarks=bhd_remarks,
                            fqm_date= datetime.strptime(fqm_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),fqm_status=fqm_status,fqm_remarks=fqm_remarks,qm_certification_date= datetime.strptime(qm_certification_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),qm_certification_status=qm_certification_status,qm_certification_remarks=qm_certification_remarks,enduser_date= datetime.strptime(end_user_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),enduser_status=end_user_status, enduser_remarks=end_user_remarks,remarks=remarks,user_id= user_id)
                        data.append(prod_system)
                    ProductionSystemStatus.objects.bulk_create(data)
                    os.remove(file_path)

                elif system_type == "Air Launch":
                    for system, sys_type, testing_type, organization, set_id, blt_date, blt_status, blt_remarks, pre_hil_date, pre_hil_status, pre_hil_remarks, vibration_date, vibration_status, vibration_remarks,\
                            cg_balancing_date, cg_balancing_status, cg_balancing_remarks, post_hil_date, post_hil_status, post_hil_remarks, sys_align_date, sys_align_status, sys_align_remarks, bhd_date, bhd_status, bhd_remarks,\
                            fqm_date, fqm_status, fqm_remarks, qm_certification_date, qm_certification_status, qm_certification_remarks, end_user_date, end_user_status, end_user_remarks, remarks, *__ in csvf:

                        prod_system = ProductionSystemStatus(system=system,sys_type=sys_type,testing_type=testing_type,organization=organization,set_id=set_id,blt_date= datetime.strptime(blt_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),blt_status=blt_status,blt_remarks=blt_remarks,
                            pre_hil_date= datetime.strptime(pre_hil_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),pre_hil_status=pre_hil_status,pre_hil_remarks=pre_hil_remarks,vibaration_date= datetime.strptime(vibration_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),vibaration_status=vibration_status,vibaration_remarks=vibration_remarks,
                            cgbalancing_date=  datetime.strptime(cg_balancing_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),cgbalancing_date_status=cg_balancing_status,cgbalancing_date_remarks=cg_balancing_remarks,post_hil_date= datetime.strptime(post_hil_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),post_hil_status=post_hil_status,post_hil_remarks=post_hil_remarks,
                            sys_align_Date=  datetime.strptime(sys_align_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"), sys_align_status= sys_align_status, sys_align_remarks= sys_align_remarks,bhd_date= datetime.strptime(bhd_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),bhd_status=bhd_status,bhd_remarks=bhd_remarks,fqm_date= datetime.strptime(fqm_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),fqm_status=fqm_status,fqm_remarks=fqm_remarks,
                            qm_certification_date= datetime.strptime(qm_certification_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),qm_certification_status=qm_certification_status,qm_certification_remarks=qm_certification_remarks,enduser_date= datetime.strptime(end_user_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),enduser_status=end_user_status, enduser_remarks=end_user_remarks,remarks=remarks,user_id= user_id)
                        data.append(prod_system)
                    ProductionSystemStatus.objects.bulk_create(data)
                    os.remove(file_path)
                return JsonResponse({'message': 'Production Status Updated Successfully!'}, status=200)

            else:
                return JsonResponse({'status': 'False', "message": "Status Not Saved"}, status=500)
        except Exception as e:
            return JsonResponse({'status': 'False', "message": "Status Not Saved"}, status=500)


    @staticmethod
    def ImportFlightCsv(request):
        # try:
            id = request['id']
            # importedCsvFile = request["csv_file"]
            system_type = request["system_type"]
            CsvFileData = request["CsvFileData"]
            csv_data = json.loads(CsvFileData)
            # if not os.path.isdir('imported_files'):
            #     os.mkdir('imported_files')
            # path = "imported_files/"
            # fs = FileSystemStorage(location=path)
            # fs.save(importedCsvFile.name, importedCsvFile)
            # file_path = path + importedCsvFile.name
            user_id = 0
            if request['user_id'] != '':
                user_id = request['user_id']
            if id == '0':
                # file = open(file_path)
                # csvf = csv.reader(file)
                # next(csvf, None)
                if system_type == "Ballistic":

                    # bhd_remarks, fqm_remarks, qm_certification_remarks, launchact_remarks
                    for data in csv_data:
                        prodModel = FlightSystemStatus()
                        prodModel.system = data["system"]
                        prodModel.organization = data['organization']
                        prodModel.set_id = data['set_id']
                        prodModel.blt_date =  datetime.strptime(data['blt_date'], '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f")
                        prodModel.testing_type = data['testing_type']
                        prodModel.sys_type = data['sys_type']
                        prodModel.blt_status = data['blt_status']
                        prodModel.blt_remarks = json.dumps(data["blt_remarks"])
                        prodModel.pre_hil_date = datetime.strptime(data['pre_hil_date'], '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f")
                        prodModel.pre_hil_status = data['pre_hil_status']
                        prodModel.pre_hil_remarks = json.dumps(data['pre_hil_remarks'])
                        prodModel.vibaration_date = datetime.strptime(data['vibaration_date'], '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f")
                        prodModel.vibaration_status = data['vibaration_status']
                        prodModel.vibaration_remarks = json.dumps(data['vibaration_remarks'])
                        prodModel.cgbalancing_date = datetime.strptime(data['cgbalancing_date'], '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f")
                        prodModel.cgbalancing_date_status = data['cgbalancing_date_status']
                        prodModel.cgbalancing_date_remarks = json.dumps(data['cgbalancing_date_remarks'])
                        prodModel.post_hil_date = datetime.strptime(data['post_hil_date'], '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f")
                        prodModel.post_hil_status = data['post_hil_status']
                        prodModel.post_hil_remarks = json.dumps(data['post_hil_remarks'])
                        prodModel.final_integration_date = datetime.strptime(data['final_integration_date'], '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f")
                        prodModel.final_integration_status = data['final_integration_status']
                        prodModel.final_integration_remarks = json.dumps(data['final_integration_remarks'])
                        prodModel.bhd_date = datetime.strptime(data['bhd_date'], '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f")
                        prodModel.bhd_status = data['bhd_status']
                        # prodModel.bhd_remarks = json.dumps(data['bhd_remarks'])
                        prodModel.fqm_date =  datetime.strptime(data['fqm_date'], '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f")
                        prodModel.fqm_status = data['fqm_status']
                        # prodModel.fqm_remarks = json.dumps(data['fqm_remarks'])
                        prodModel.qm_certification_date = datetime.strptime(data['qm_certification_date'], '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f")
                        prodModel.qm_certification_status = data['qm_certification_status']
                        # prodModel.qm_certification_remarks = json.dumps(data['qm_certification_remarks'])

                        prodModel.save()

                    # for system, sys_type, testing_type, organization, set_id, blt_date, blt_status, blt_remarks, pre_hil_date, pre_hil_status, pre_hil_remarks,\
                    #         vibration_date, vibration_status, vibration_remarks, cg_balancing_date, cg_balancing_status, cg_balancing_remarks, post_hil_date, post_hil_status, post_hil_remarks,\
                    #         final_integration_date, final_integration_status, final_integration_remarks,bhd_date, bhd_status, fqm_date, fqm_status, \
                    #         qm_certification_date, qm_certification_status,launchact_date, launchact_status in csv_data:
                    #
                    #     flight_system = FlightSystemStatus(system=system,sys_type=sys_type,testing_type=testing_type,organization=organization,set_id=set_id,blt_date= datetime.strptime(blt_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),blt_status=blt_status,blt_remarks=blt_remarks,
                    #         pre_hil_date= datetime.strptime(pre_hil_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),pre_hil_status=pre_hil_status,pre_hil_remarks=pre_hil_remarks,vibaration_date= datetime.strptime(vibration_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),vibaration_status=vibration_status,vibaration_remarks=vibration_remarks,
                    #         cgbalancing_date=  datetime.strptime(cg_balancing_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),cgbalancing_date_status=cg_balancing_status,cgbalancing_date_remarks=cg_balancing_remarks,post_hil_date= datetime.strptime(post_hil_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),post_hil_status=post_hil_status,post_hil_remarks=post_hil_remarks,
                    #         final_integration_date= datetime.strptime(final_integration_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),final_integration_status=final_integration_status,final_integration_remarks=final_integration_remarks,
                    #         bhd_date = datetime.strptime(bhd_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"), bhd_status = bhd_status,
                    #         fqm_date = datetime.strptime(fqm_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"), fqm_status = fqm_status,
                    #         qm_certification_date = datetime.strptime(qm_certification_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"), qm_certification_status = qm_certification_status,
                    #         launchact_date = datetime.strptime(launchact_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"), launchact_status = launchact_status,user_id= user_id)
                    #     data.append(flight_system)
                    # FlightSystemStatus.objects.bulk_create(data)
                    # os.remove(file_path)
                elif system_type == "Cruise":
                    for system, sys_type, testing_type, organization, set_id, blt_date, blt_status, blt_remarks, pre_hil_date, pre_hil_status, pre_hil_remarks, vibration_date, vibration_status, vibration_remarks,\
                            cg_balancing_date, cg_balancing_status, cg_balancing_remarks, post_hil_date, post_hil_status, post_hil_remarks,sys_align_date, sys_align_status, sys_align_remarks, fgt_date, fgt_status, fgt_remarks,\
                            encapsulation_date, encapsulation_status, encapsulation_remarks, final_integrated_testing_date, final_integrated_testing_status, final_integrated_testing_remarks, bhd_date, bhd_status, bhd_remarks,\
                            fqm_date, fqm_status, fqm_remarks, qm_certification_date, qm_certification_status, qm_certification_remarks, \
                            launchact_date, launchact_status, launchact_remarks, remarks, *__ in csvf:
                        flight_system = FlightSystemStatus(system=system,sys_type=sys_type,testing_type=testing_type,organization=organization,
                            set_id=set_id,blt_date= datetime.strptime(blt_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),blt_status=blt_status,blt_remarks=blt_remarks,pre_hil_date= datetime.strptime(pre_hil_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),pre_hil_status=pre_hil_status,
                            pre_hil_remarks=pre_hil_remarks,vibaration_date= datetime.strptime(vibration_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),vibaration_status=vibration_status,vibaration_remarks=vibration_remarks,
                            cgbalancing_date=  datetime.strptime(cg_balancing_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),cgbalancing_date_status=cg_balancing_status,cgbalancing_date_remarks=cg_balancing_remarks,post_hil_date= datetime.strptime(post_hil_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),post_hil_status=post_hil_status,post_hil_remarks=post_hil_remarks,
                            sys_align_Date=  datetime.strptime(sys_align_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"), sys_align_status= sys_align_status, sys_align_remarks= sys_align_remarks,
                            fgt_date=  datetime.strptime(fgt_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"), fgt_status= fgt_status, fgt_remarks=fgt_remarks, incapsulation_date=  datetime.strptime(encapsulation_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"), incapsulation_status=encapsulation_status, incapsulation_remarks=encapsulation_remarks,
                            final_integrated_testing_date= datetime.strptime(final_integrated_testing_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),final_integrated_testing=final_integrated_testing_status,final_integrated_testing_remarks=final_integrated_testing_remarks,
                            bhd_date= datetime.strptime(bhd_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),bhd_status=bhd_status,bhd_remarks=bhd_remarks,fqm_date= datetime.strptime(fqm_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),fqm_status=fqm_status,fqm_remarks=fqm_remarks,
                            qm_certification_date= datetime.strptime(qm_certification_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),qm_certification_status=qm_certification_status,qm_certification_remarks=qm_certification_remarks,
                            launchact_date= datetime.strptime(launchact_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),launchact_status=launchact_status, launchact_remarks=launchact_remarks,remarks=remarks,user_id= user_id)
                        data.append(flight_system)
                    FlightSystemStatus.objects.bulk_create(data)
                    os.remove(file_path)

                elif system_type =="Launch System":
                    for system, sys_type, testing_type, organization, set_id, blt_date, blt_status, blt_remarks,emp_proofing_date,emp_proofing_status,emp_proofing_remarks,\
                        func_test_date,func_test_status,func_test_remarks,func_test_dummy_date,func_test_dummy_status,func_test_dummy_remarks,road_test_date,road_test_status,road_test_remarks,\
                        post_road_test_date,post_road_test_status,post_road_test_remarks,integrated_operation_date,integrated_operation_status,integrated_operation_remarks,\
                        rain_test_date,rain_test_status,rain_test_remarks,pre_user_inspection_date,pre_user_inspection_status,\
                        pre_user_inspection_remarks,final_integrated_testing_date, final_integrated_testing_status, final_integrated_testing_remarks,\
                        bhd_date, bhd_status, bhd_remarks, fqm_date, fqm_status, fqm_remarks, qm_certification_date, qm_certification_status, qm_certification_remarks,\
                        launchact_date, launchact_status, launchact_remarks, remarks, *__ in csvf:

                        flight_system = FlightSystemStatus(system=system,sys_type=sys_type,testing_type=testing_type,organization=organization,set_id=set_id,blt_date= datetime.strptime(blt_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),blt_status=blt_status,blt_remarks=blt_remarks,
                            emp_proofing_date=  datetime.strptime(emp_proofing_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),emp_proofing= emp_proofing_status,emp_proofing_remarks= emp_proofing_remarks,
                            func_tst_date=  datetime.strptime(func_test_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),func_tst= func_test_status,func_tst_remarks= func_test_remarks,
                            func_tst_dummy_bird= func_test_dummy_status,func_tst_dummy_bird_date= datetime.strptime(func_test_dummy_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),func_tst_dummy_bird_remarks= func_test_dummy_remarks,
                            road_test= road_test_status,road_test_date=  datetime.strptime(road_test_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),road_test_remarks= road_test_remarks,
                            post_road_test_date=  datetime.strptime(post_road_test_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),post_road_test= post_road_test_status,post_road_test_remarks= post_road_test_remarks,
                            integrated_operation=integrated_operation_status, integrated_operation_date= datetime.strptime(integrated_operation_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),integrated_operation_remarks= integrated_operation_remarks,
                            rain_test=rain_test_status, rain_test_date=  datetime.strptime(rain_test_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),rain_test_remarks= rain_test_remarks,
                            pre_user_inspection_date=  datetime.strptime(pre_user_inspection_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),pre_user_inspection= pre_user_inspection_status, pre_user_inspection_remarks= pre_user_inspection_remarks,
                            final_integrated_testing_date= datetime.strptime(final_integrated_testing_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),final_integrated_testing=final_integrated_testing_status,final_integrated_testing_remarks=final_integrated_testing_remarks,
                            bhd_date= datetime.strptime(bhd_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),bhd_status=bhd_status,bhd_remarks=bhd_remarks,fqm_date= datetime.strptime(fqm_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),fqm_status=fqm_status,fqm_remarks=fqm_remarks,
                            qm_certification_date= datetime.strptime(qm_certification_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),qm_certification_status=qm_certification_status,qm_certification_remarks=qm_certification_remarks,
                            launchact_date= datetime.strptime(launchact_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),launchact_status=launchact_status, launchact_remarks=launchact_remarks,remarks=remarks,user_id= user_id)
                        data.append(flight_system)
                    FlightSystemStatus.objects.bulk_create(data)
                    os.remove(file_path)

                elif system_type == "Launch Tube":
                    for system, sys_type, testing_type, organization, set_id, blt_date, blt_status, blt_remarks, pre_user_inspection_date,pre_user_inspection_status,pre_user_inspection_remarks,\
                            load_unload_date,load_unload_status,load_unload_remarks,encapsulation_date,encapsulation_status,encapsulation_remarks, final_integrated_testing_date, final_integrated_testing_status, final_integrated_testing_remarks,\
                            bhd_date, bhd_status, bhd_remarks, fqm_date, fqm_status, fqm_remarks, qm_certification_date, qm_certification_status, qm_certification_remarks, launchact_date, launchact_status, launchact_remarks, remarks, *__ in csvf:

                        flight_system = FlightSystemStatus(system=system,sys_type=sys_type,testing_type=testing_type,organization=organization,set_id=set_id,blt_date= datetime.strptime(blt_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),blt_status=blt_status,blt_remarks=blt_remarks,
                            pre_user_inspection_date=  datetime.strptime(pre_user_inspection_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),pre_user_inspection= pre_user_inspection_status,pre_user_inspection_remarks= pre_user_inspection_remarks,
                            load_unload_on_mlv_hlf_date=  datetime.strptime(load_unload_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),load_unload_on_mlv_hlf= load_unload_status,load_unload_on_mlv_hlf_remarks=load_unload_remarks,incapsulation_date= datetime.strptime(encapsulation_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),incapsulation_status= encapsulation_status,incapsulation_remarks= encapsulation_remarks,
                            final_integrated_testing_date= datetime.strptime(final_integrated_testing_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),final_integrated_testing=final_integrated_testing_status,final_integrated_testing_remarks=final_integrated_testing_remarks,bhd_date= datetime.strptime(bhd_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),bhd_status=bhd_status,bhd_remarks=bhd_remarks,
                            fqm_date= datetime.strptime(fqm_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),fqm_status=fqm_status,fqm_remarks=fqm_remarks,qm_certification_date= datetime.strptime(qm_certification_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),qm_certification_status=qm_certification_status,qm_certification_remarks=qm_certification_remarks,
                            launchact_date= datetime.strptime(launchact_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),launchact_status=launchact_status, launchact_remarks=launchact_remarks,remarks=remarks,user_id= user_id)
                        data.append(flight_system)
                    FlightSystemStatus.objects.bulk_create(data)
                    os.remove(file_path)

                elif system_type == "Air Launch":
                    for system, sys_type, testing_type, organization, set_id, blt_date, blt_status, blt_remarks, pre_hil_date, pre_hil_status, pre_hil_remarks, vibration_date, vibration_status, vibration_remarks,\
                            cg_balancing_date, cg_balancing_status, cg_balancing_remarks, post_hil_date, post_hil_status, post_hil_remarks, sys_align_date, sys_align_status, sys_align_remarks, bhd_date, bhd_status, bhd_remarks,\
                            fqm_date, fqm_status, fqm_remarks, qm_certification_date, qm_certification_status, qm_certification_remarks, launchact_date, launchact_status, launchact_remarks, remarks, *__ in csvf:

                        flight_system = FlightSystemStatus(system=system,sys_type=sys_type,testing_type=testing_type,organization=organization,set_id=set_id,blt_date= datetime.strptime(blt_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),blt_status=blt_status,blt_remarks=blt_remarks,
                            pre_hil_date= datetime.strptime(pre_hil_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),pre_hil_status=pre_hil_status,pre_hil_remarks=pre_hil_remarks,vibaration_date= datetime.strptime(vibration_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),vibaration_status=vibration_status,vibaration_remarks=vibration_remarks,
                            cgbalancing_date=  datetime.strptime(cg_balancing_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),cgbalancing_date_status=cg_balancing_status,cgbalancing_date_remarks=cg_balancing_remarks,post_hil_date= datetime.strptime(post_hil_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),post_hil_status=post_hil_status,post_hil_remarks=post_hil_remarks,
                            sys_align_Date=  datetime.strptime(sys_align_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"), sys_align_status= sys_align_status, sys_align_remarks= sys_align_remarks,bhd_date= datetime.strptime(bhd_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),bhd_status=bhd_status,bhd_remarks=bhd_remarks,fqm_date= datetime.strptime(fqm_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),fqm_status=fqm_status,fqm_remarks=fqm_remarks,
                            qm_certification_date= datetime.strptime(qm_certification_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),qm_certification_status=qm_certification_status,qm_certification_remarks=qm_certification_remarks,launchact_date= datetime.strptime(launchact_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),launchact_status=launchact_status, launchact_remarks=launchact_remarks,remarks=remarks,user_id= user_id)
                        data.append(flight_system)
                    FlightSystemStatus.objects.bulk_create(data)
                    os.remove(file_path)
                return JsonResponse({'message': 'Production Status Updated Successfully!'}, status=200)

            else:
                return JsonResponse({'status': 'False', "message": "Status Not Saved"}, status=500)
        # except Exception as e:
        #     return JsonResponse({'status': 'False', "message": "Status Not Saved"}, status=500)


    @staticmethod
    def ImportRelifingCsv(request):
        # try:
            id = request['id']
            importedCsvFile = request["csv_file"]
            system_type = request["system_type"]
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
                if system_type == "Ballistic":
                    for system, sys_type, testing_type, organization, set_id, blt_date, blt_status, blt_remarks, pre_hil_date, pre_hil_status, pre_hil_remarks,\
                            vibration_date, vibration_status, vibration_remarks, cg_balancing_date, cg_balancing_status, cg_balancing_remarks, post_hil_date, post_hil_status, post_hil_remarks,\
                            final_integrated_testing_date, final_integrated_testing_status, final_integrated_testing_remarks, bhd_date, bhd_status, bhd_remarks, fqm_date, fqm_status, fqm_remarks,\
                            qm_certification_date, qm_certification_status, qm_certification_remarks, end_user_date, end_user_status, end_user_remarks, remarks, *__ in csvf:

                        relif_system = RelifingSystemStatus(system=system,sys_type=sys_type,testing_type=testing_type,organization=organization,set_id=set_id,blt_date= datetime.strptime(blt_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),blt_status=blt_status,blt_remarks=blt_remarks,
                            pre_hil_date= datetime.strptime(pre_hil_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),pre_hil_status=pre_hil_status,pre_hil_remarks=pre_hil_remarks,vibaration_date= datetime.strptime(vibration_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),vibaration_status=vibration_status,vibaration_remarks=vibration_remarks,
                            cgbalancing_date=  datetime.strptime(cg_balancing_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),cgbalancing_date_status=cg_balancing_status,cgbalancing_date_remarks=cg_balancing_remarks,post_hil_date= datetime.strptime(post_hil_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),post_hil_status=post_hil_status,post_hil_remarks=post_hil_remarks,
                            final_integrated_testing_date= datetime.strptime(final_integrated_testing_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),final_integrated_testing=final_integrated_testing_status,final_integrated_testing_remarks=final_integrated_testing_remarks,
                            bhd_date= datetime.strptime(bhd_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),bhd_status=bhd_status,bhd_remarks=bhd_remarks,fqm_date= datetime.strptime(fqm_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),fqm_status=fqm_status,fqm_remarks=fqm_remarks,
                            qm_certification_date= datetime.strptime(qm_certification_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),qm_certification_status=qm_certification_status,qm_certification_remarks=qm_certification_remarks,
                            enduser_date= datetime.strptime(end_user_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),enduser_status=end_user_status, enduser_remarks=end_user_remarks,remarks=remarks,user_id= user_id)
                        data.append(relif_system)
                    RelifingSystemStatus.objects.bulk_create(data)
                    os.remove(file_path)
                elif system_type == "Cruise":
                    for system, sys_type, testing_type, organization, set_id, blt_date, blt_status, blt_remarks, pre_hil_date, pre_hil_status, pre_hil_remarks, vibration_date, vibration_status, vibration_remarks,\
                            cg_balancing_date, cg_balancing_status, cg_balancing_remarks, post_hil_date, post_hil_status, post_hil_remarks,sys_align_date, sys_align_status, sys_align_remarks, fgt_date, fgt_status, fgt_remarks,\
                            encapsulation_date, encapsulation_status, encapsulation_remarks, final_integrated_testing_date, final_integrated_testing_status, final_integrated_testing_remarks, bhd_date, bhd_status, bhd_remarks,\
                            fqm_date, fqm_status, fqm_remarks, qm_certification_date, qm_certification_status, qm_certification_remarks, end_user_date, end_user_status, end_user_remarks, remarks, *__ in csvf:
                        relif_system = RelifingSystemStatus(system=system,sys_type=sys_type,testing_type=testing_type,organization=organization,
                            set_id=set_id,blt_date= datetime.strptime(blt_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),blt_status=blt_status,blt_remarks=blt_remarks,pre_hil_date= datetime.strptime(pre_hil_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),pre_hil_status=pre_hil_status,
                            pre_hil_remarks=pre_hil_remarks,vibaration_date= datetime.strptime(vibration_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),vibaration_status=vibration_status,vibaration_remarks=vibration_remarks,
                            cgbalancing_date=  datetime.strptime(cg_balancing_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),cgbalancing_date_status=cg_balancing_status,cgbalancing_date_remarks=cg_balancing_remarks,post_hil_date= datetime.strptime(post_hil_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),post_hil_status=post_hil_status,post_hil_remarks=post_hil_remarks,
                            sys_align_Date=  datetime.strptime(sys_align_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"), sys_align_status= sys_align_status, sys_align_remarks= sys_align_remarks,
                            fgt_date=  datetime.strptime(fgt_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"), fgt_status= fgt_status, fgt_remarks=fgt_remarks, incapsulation_date=  datetime.strptime(encapsulation_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"), incapsulation_status=encapsulation_status, incapsulation_remarks=encapsulation_remarks,
                            final_integrated_testing_date= datetime.strptime(final_integrated_testing_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),final_integrated_testing=final_integrated_testing_status,final_integrated_testing_remarks=final_integrated_testing_remarks,
                            bhd_date= datetime.strptime(bhd_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),bhd_status=bhd_status,bhd_remarks=bhd_remarks,fqm_date= datetime.strptime(fqm_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),fqm_status=fqm_status,fqm_remarks=fqm_remarks,
                            qm_certification_date= datetime.strptime(qm_certification_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),qm_certification_status=qm_certification_status,qm_certification_remarks=qm_certification_remarks,enduser_date= datetime.strptime(end_user_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),enduser_status=end_user_status, enduser_remarks=end_user_remarks,remarks=remarks,user_id= user_id)
                        data.append(relif_system)
                    RelifingSystemStatus.objects.bulk_create(data)
                    os.remove(file_path)

                elif system_type =="Launch System":
                    for system, sys_type, testing_type, organization, set_id, blt_date, blt_status, blt_remarks,emp_proofing_date,emp_proofing_status,emp_proofing_remarks,\
                        func_test_date,func_test_status,func_test_remarks,func_test_dummy_date,func_test_dummy_status,func_test_dummy_remarks,road_test_date,road_test_status,road_test_remarks,\
                        post_road_test_date,post_road_test_status,post_road_test_remarks,integrated_operation_date,integrated_operation_status,integrated_operation_remarks,\
                        rain_test_date,rain_test_status,rain_test_remarks,pre_user_inspection_date,pre_user_inspection_status,\
                        pre_user_inspection_remarks,final_integrated_testing_date, final_integrated_testing_status, final_integrated_testing_remarks, bhd_date, bhd_status, bhd_remarks, fqm_date, fqm_status, fqm_remarks, qm_certification_date, qm_certification_status, qm_certification_remarks, end_user_date, end_user_status, end_user_remarks, remarks, *__ in csvf:

                        relif_system = RelifingSystemStatus(system=system,sys_type=sys_type,testing_type=testing_type,organization=organization,set_id=set_id,blt_date= datetime.strptime(blt_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),blt_status=blt_status,blt_remarks=blt_remarks,
                            emp_proofing_date=  datetime.strptime(emp_proofing_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),emp_proofing= emp_proofing_status,emp_proofing_remarks= emp_proofing_remarks,
                            func_tst_date=  datetime.strptime(func_test_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),func_tst= func_test_status,func_tst_remarks= func_test_remarks,
                            func_tst_dummy_bird= func_test_dummy_status,func_tst_dummy_bird_date= datetime.strptime(func_test_dummy_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),func_tst_dummy_bird_remarks= func_test_dummy_remarks,
                            road_test= road_test_status,road_test_date=  datetime.strptime(road_test_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),road_test_remarks= road_test_remarks,
                            post_road_test_date=  datetime.strptime(post_road_test_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),post_road_test= post_road_test_status,post_road_test_remarks= post_road_test_remarks,
                            integrated_operation=integrated_operation_status, integrated_operation_date= datetime.strptime(integrated_operation_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),integrated_operation_remarks= integrated_operation_remarks,
                            rain_test=rain_test_status, rain_test_date=  datetime.strptime(rain_test_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),rain_test_remarks= rain_test_remarks,
                            pre_user_inspection_date=  datetime.strptime(pre_user_inspection_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),pre_user_inspection= pre_user_inspection_status, pre_user_inspection_remarks= pre_user_inspection_remarks,
                            final_integrated_testing_date= datetime.strptime(final_integrated_testing_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),final_integrated_testing=final_integrated_testing_status,final_integrated_testing_remarks=final_integrated_testing_remarks,
                            bhd_date= datetime.strptime(bhd_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),bhd_status=bhd_status,bhd_remarks=bhd_remarks,fqm_date= datetime.strptime(fqm_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),fqm_status=fqm_status,fqm_remarks=fqm_remarks,
                            qm_certification_date= datetime.strptime(qm_certification_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),qm_certification_status=qm_certification_status,qm_certification_remarks=qm_certification_remarks,
                            enduser_date= datetime.strptime(end_user_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),enduser_status=end_user_status, enduser_remarks=end_user_remarks,remarks=remarks,user_id= user_id)
                        data.append(relif_system)
                    RelifingSystemStatus.objects.bulk_create(data)
                    os.remove(file_path)

                elif system_type == "Launch Tube":
                    for system, sys_type, testing_type, organization, set_id, blt_date, blt_status, blt_remarks, pre_user_inspection_date,pre_user_inspection_status,pre_user_inspection_remarks,\
                            load_unload_date,load_unload_status,load_unload_remarks,encapsulation_date,encapsulation_status,encapsulation_remarks, final_integrated_testing_date, final_integrated_testing_status, final_integrated_testing_remarks,\
                            bhd_date, bhd_status, bhd_remarks, fqm_date, fqm_status, fqm_remarks, qm_certification_date, qm_certification_status, qm_certification_remarks, end_user_date, end_user_status, end_user_remarks, remarks, *__ in csvf:

                        relif_system = RelifingSystemStatus(system=system,sys_type=sys_type,testing_type=testing_type,organization=organization,set_id=set_id,blt_date= datetime.strptime(blt_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),blt_status=blt_status,blt_remarks=blt_remarks,
                            pre_user_inspection_date=  datetime.strptime(pre_user_inspection_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),pre_user_inspection= pre_user_inspection_status,pre_user_inspection_remarks= pre_user_inspection_remarks,
                            load_unload_on_mlv_hlf_date=  datetime.strptime(load_unload_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),load_unload_on_mlv_hlf= load_unload_status,load_unload_on_mlv_hlf_remarks=load_unload_remarks,incapsulation_date= datetime.strptime(encapsulation_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),incapsulation_status= encapsulation_status,incapsulation_remarks= encapsulation_remarks,
                            final_integrated_testing_date= datetime.strptime(final_integrated_testing_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),final_integrated_testing=final_integrated_testing_status,final_integrated_testing_remarks=final_integrated_testing_remarks,bhd_date= datetime.strptime(bhd_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),bhd_status=bhd_status,bhd_remarks=bhd_remarks,
                            fqm_date= datetime.strptime(fqm_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),fqm_status=fqm_status,fqm_remarks=fqm_remarks,qm_certification_date= datetime.strptime(qm_certification_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),qm_certification_status=qm_certification_status,qm_certification_remarks=qm_certification_remarks,enduser_date= datetime.strptime(end_user_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),enduser_status=end_user_status, enduser_remarks=end_user_remarks,remarks=remarks,user_id= user_id)
                        data.append(relif_system)
                    RelifingSystemStatus.objects.bulk_create(data)
                    os.remove(file_path)

                elif system_type == "Air Launch":
                    for system, sys_type, testing_type, organization, set_id, blt_date, blt_status, blt_remarks, pre_hil_date, pre_hil_status, pre_hil_remarks, vibration_date, vibration_status, vibration_remarks,\
                            cg_balancing_date, cg_balancing_status, cg_balancing_remarks, post_hil_date, post_hil_status, post_hil_remarks, sys_align_date, sys_align_status, sys_align_remarks, bhd_date, bhd_status, bhd_remarks,\
                            fqm_date, fqm_status, fqm_remarks, qm_certification_date, qm_certification_status, qm_certification_remarks, end_user_date, end_user_status, end_user_remarks, remarks, *__ in csvf:

                        relif_system = RelifingSystemStatus(system=system,sys_type=sys_type,testing_type=testing_type,organization=organization,set_id=set_id,blt_date= datetime.strptime(blt_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),blt_status=blt_status,blt_remarks=blt_remarks,
                            pre_hil_date= datetime.strptime(pre_hil_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),pre_hil_status=pre_hil_status,pre_hil_remarks=pre_hil_remarks,vibaration_date= datetime.strptime(vibration_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),vibaration_status=vibration_status,vibaration_remarks=vibration_remarks,
                            cgbalancing_date=  datetime.strptime(cg_balancing_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),cgbalancing_date_status=cg_balancing_status,cgbalancing_date_remarks=cg_balancing_remarks,post_hil_date= datetime.strptime(post_hil_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),post_hil_status=post_hil_status,post_hil_remarks=post_hil_remarks,
                            sys_align_Date=  datetime.strptime(sys_align_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"), sys_align_status= sys_align_status, sys_align_remarks= sys_align_remarks,bhd_date= datetime.strptime(bhd_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),bhd_status=bhd_status,bhd_remarks=bhd_remarks,fqm_date= datetime.strptime(fqm_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),fqm_status=fqm_status,fqm_remarks=fqm_remarks,
                            qm_certification_date= datetime.strptime(qm_certification_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),qm_certification_status=qm_certification_status,qm_certification_remarks=qm_certification_remarks,enduser_date= datetime.strptime(end_user_date, '%m-%d-%Y').strftime("%Y-%m-%dT%H:%M:%S.%f"),enduser_status=end_user_status, enduser_remarks=end_user_remarks,remarks=remarks,user_id= user_id)
                        data.append(relif_system)
                    RelifingSystemStatus.objects.bulk_create(data)
                    os.remove(file_path)
                return JsonResponse({'message': 'Production Status Updated Successfully!'}, status=200)

            else:
                return JsonResponse({'status': 'False', "message": "Status Not Saved"}, status=500)
        # except Exception as e:
        #     return JsonResponse({'status': 'False', "message": "Status Not Saved"}, status=500)

    @staticmethod
    def GetProdctionDateNone(request):
        dataList = ProductionSystemStatusHistory.objects.all()

        for data in dataList:
            if data.blt_status == 'None':
                data.blt_date = None

            if data.emp_proofing == 'None':
                data.emp_proofing_date = None

            if data.func_tst == 'None':
                data.func_tst_date = None

            if data.func_tst_dummy_bird == 'None':
                data.func_tst_dummy_bird_date = None

            if data.road_test == 'None':
                data.road_test_date = None

            if data.post_road_test == 'None':
                data.post_road_test_date = None

            if data.integrated_operation == 'None':
                data.integrated_operation_date = None

            if data.rain_test == 'None':
                data.rain_test_date = None

            if data.pre_user_inspection == 'None':
                data.pre_user_inspection_date = None

            if data.final_integration_status == 'None':
                data.final_integration_date = None

            if data.load_unload_on_mlv_hlf == 'None':
                data.load_unload_on_mlv_hlf_date = None

            if data.pre_hil_status == 'None':
                data.pre_hil_date = None

            if data.vibaration_status == 'None':
                data.vibaration_date = None

            if data.cgbalancing_date_status == 'None':
                data.cgbalancing_date = None

            if data.cgbalancing_date_status == 'None':
                data.cgbalancing_date = None

            if data.post_hil_status == 'None':
                data.post_hil_date = None

            if data.sys_align_status == 'None':
                data.sys_align_Date = None

            if data.incapsulation_status == 'None':
                data.incapsulation_date = None

            if data.final_integration_status == 'None':
                data.final_integration_date = None

            if data.final_integration_status == 'None':
                data.final_integration_date = None

            if data.fgt_status == 'None':
                data.fgt_date =None

            if data.bhd_status == 'None':
                data.bhd_date = None

            if data.fqm_status == 'None':
                data.fqm_date = None

            if data.qm_certification_status == 'None':
                data.qm_certification_date = None

            if data.enduser_status == 'None' or data.enduser_status == '':
                data.enduser_date = None
            data.save()
        ListItem = ProductionSystemStatusHistory.objects.all()
        serializer = ProductionSystemSerialzer(ListItem, many=True)
        return JsonResponse({'message': 'Welcome to Home Page', 'data': serializer.data}, status=200)
    @staticmethod
    def GetProductionList(request):
        try:
            org = request.query_params['selected_org']
            year = request.query_params['selected_year']
            type = request.query_params['selected_type']
            system = request.query_params['selected_system']
            ParentStatus = request.query_params['selected_status']
            ChildStatus = request.query_params['child_status']
            # print(ParentStatus)
            # print(ChildStatus)
            filter_objects = Q()

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

            # create dynamic filter
            # if year != '':
            #     filter_objects &= get_filter(
            #         'testing_date__year', 'equal',
            #         year)
            if org != '':
                filter_objects &= get_filter(
                    'organization', 'equal',
                    org)
            if type != '':
                filter_objects &= get_filter(
                    'sys_type', 'equal',
                    type)
            if system != '':
                filter_objects &= get_filter(
                    'system', 'equal',
                    system)

            setid_filter = Q()
            setid_filter &= get_filter(
                    'set_id', 'not_equal',
                    '')
            if org != '':
                setid_filter &= get_filter(
                    'organization', 'equal',
                    org)
            if type != '':
                setid_filter &= get_filter(
                    'sys_type', 'equal',
                    type)
            if system != '':
                setid_filter &= get_filter(
                    'system', 'equal',
                    system)

            # else:
            dataList = ProductionSystemStatus.objects.filter(filter_objects)
            if ParentStatus != '':
                ListItems = []
                totCountProduction = ProductionSystemStatus.objects.filter(setid_filter)
                for set in totCountProduction:
                    count_id = 0
                    if set.blt_status != 'None' and set.set_id != '' and set.blt_date is not None:
                        if count_id == 0 and set.blt_date.strftime("%Y") == year:
                            count_id = set.id
                            ListItems.append(set)
                    if set.pre_hil_status != 'None' and set.set_id != '' and set.pre_hil_date is not None:
                        if count_id == 0 and set.pre_hil_date.strftime("%Y") == year:
                            count_id = set.id
                            ListItems.append(set)
                    if set.vibaration_status != 'None' and set.set_id != '' and set.vibaration_date is not None:
                        if count_id == 0 and set.vibaration_date.strftime("%Y") == year:
                            count_id = set.id
                            ListItems.append(set)
                    if set.post_hil_status != 'None' and set.set_id != '' and set.post_hil_date is not None:
                        if count_id == 0 and set.post_hil_date.strftime("%Y") == year:
                            count_id = set.id
                            ListItems.append(set)
                    if set.fgt_status != 'None' and set.set_id != '' and set.fgt_date is not None:
                        if count_id == 0 and set.fgt_date.strftime("%Y") == year:
                            count_id = set.id
                            ListItems.append(set)

                    if set.final_integration_status != 'None' and set.set_id != '' and set.final_integration_date is not None:
                        if count_id == 0 and set.final_integration_date.strftime("%Y") == year:
                            count_id = set.id
                            ListItems.append(set)
                    if set.bhd_status != 'None' and set.set_id != '' and set.bhd_date is not None:
                        if count_id == 0 and set.bhd_date.strftime("%Y") == year:
                            count_id = set.id
                            ListItems.append(set)
                    if set.fqm_status != 'None' and set.set_id != '' and set.fqm_date is not None:
                        if count_id == 0 and set.fqm_date.strftime("%Y") == year:
                            count_id = set.id
                            ListItems.append(set)
                    if set.qm_certification_status != 'None' and set.set_id != '' and set.qm_certification_date is not None:
                        if count_id == 0 and set.qm_certification_date.strftime("%Y") == year:
                            count_id = set.id
                            ListItems.append(set)
                    if set.cgbalancing_date_status != 'None' and set.set_id != '' and set.cgbalancing_date is not None:
                        if count_id == 0 and set.cgbalancing_date.strftime("%Y") == year:
                            count_id = set.id
                            ListItems.append(set)

                    if set.enduser_status != 'None' and set.set_id != '' and set.enduser_date is not None:
                        if count_id == 0 and set.enduser_date.strftime("%Y") == year:
                            count_id = set.id
                            ListItems.append(set)
                    if set.incapsulation_status != 'None' and set.set_id != '' and set.incapsulation_date is not None:
                        if count_id == 0 and set.incapsulation_date.strftime("%Y") == year:
                            count_id = set.id
                            ListItems.append(set)
                    if set.sys_align_status != 'None' and set.set_id != '' and set.sys_align_Date is not None:
                        if count_id == 0 and set.sys_align_Date.strftime("%Y") == year:
                            count_id = set.id
                            ListItems.append(set)
                    if set.emp_proofing != 'None' and set.set_id != '' and set.emp_proofing_date is not None:
                        if count_id == 0 and set.emp_proofing_date.strftime("%Y") == year:
                            count_id = set.id
                            ListItems.append(set)
                    if set.func_tst != 'None' and set.set_id != '' and set.func_tst_date is not None:
                        if count_id == 0 and set.func_tst_date.strftime("%Y") == year:
                            count_id = set.id
                            ListItems.append(set)

                    if set.func_tst_dummy_bird != 'None' and set.set_id != '' and set.func_tst_dummy_bird_date is not None:
                        if count_id == 0 and set.func_tst_dummy_bird_date.strftime("%Y") == year:
                            count_id = set.id
                            ListItems.append(set)
                    if set.road_test != 'None' and set.set_id != '' and set.road_test_date is not None:
                        if count_id == 0 and set.road_test_date.strftime("%Y") == year:
                            count_id = set.id
                            ListItems.append(set)
                    if set.post_road_test != 'None' and set.set_id != '' and set.post_road_test_date is not None:
                        if count_id == 0 and set.post_road_test_date.strftime("%Y") == year:
                            count_id = set.id
                            ListItems.append(set)
                    if set.integrated_operation != 'None' and set.set_id != '' and set.integrated_operation_date is not None:
                        if count_id == 0 and set.integrated_operation_date.strftime("%Y") == year:
                            count_id = set.id
                            ListItems.append(set)
                    if set.rain_test != 'None' and set.set_id != '' and set.rain_test_date is not None:
                        if count_id == 0 and set.rain_test_date.strftime("%Y") == year:
                            count_id = set.id
                            ListItems.append(set)

                    if set.pre_user_inspection != 'None' and set.set_id != '' and set.pre_user_inspection_date is not None:
                        if count_id == 0 and set.pre_user_inspection_date.strftime("%Y") == year:
                            count_id = set.id
                            ListItems.append(set)
                    if set.final_integrated_testing != 'None' and set.set_id != '' and set.final_integrated_testing_date is not None:
                        if count_id == 0 and set.final_integrated_testing_date.strftime("%Y") == year:
                            count_id = set.id
                            ListItems.append(set)
                    if set.load_unload_on_mlv_hlf != 'None' and set.set_id != '' and set.load_unload_on_mlv_hlf_date is not None:
                        if count_id == 0 and set.load_unload_on_mlv_hlf_date.strftime("%Y") == year:
                            count_id = set.id
                            ListItems.append(set)

                if ParentStatus == 'BLT':
                    dataList = dataList.extra(
                          select={
                          'year': 'extract (year from blt_date)',
                          'month': 'extract (month from blt_date)',
                          'day': 'extract (day from blt_date)'},
                           order_by=['month','day','-year']
                            )
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(blt_date__year = year,blt_status = ChildStatus)
                    if ChildStatus == 'Current Count':
                        for data  in dataList:
                            if data.blt_date.strftime("%Y") == year and (data.blt_status == 'Under process' or data.blt_status=='Observation(same stage)' or data.blt_status=='Halt'):
                                ListItems.append(data)

                if ParentStatus == 'EMP Proofing':
                    dataList = dataList.extra(
                          select={'year': 'extract (year from emp_proofing_date)',
                          'month': 'extract (month from emp_proofing_date)',
                          'day': 'extract (day from emp_proofing_date)'},
                           order_by=['month','day','-year'])
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(emp_proofing_date__year = year,emp_proofing = ChildStatus)

                    if ChildStatus == 'Current Count':
                        for data  in dataList:
                            if data.emp_proofing_date.strftime("%Y") == year and  (data.emp_proofing == 'Under process' or data.emp_proofing=='Observation(same stage)' or data.emp_proofing=='Halt'):
                                ListItems.append(data)

                if ParentStatus == 'Functional Test W/O Dummy Bird':
                    dataList = dataList.extra(
                          select={'year': 'extract (year from func_tst_date)',
                          'month': 'extract (month from func_tst_date)',
                          'day': 'extract (day from func_tst_date)'},
                           order_by=['month','day','-year'])
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(func_tst_date__year = year, func_tst = ChildStatus)

                    if ChildStatus == 'Current Count':
                        for data  in dataList:
                            if data.func_tst_date.strftime("%Y") == year and (data.func_tst == 'Under process' or data.func_tst=='Observation(same stage)' or data.func_tst=='Halt'):
                                ListItems.append(data)

                if ParentStatus == 'Functional Test With Dummy Bird':
                    dataList = dataList.extra(
                          select={'year': 'extract (year from func_tst_dummy_bird_date)',
                          'month': 'extract (month from func_tst_dummy_bird_date)',
                          'day': 'extract (day from func_tst_dummy_bird_date)'},
                           order_by=['month','day','-year'])
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(func_tst_dummy_bird_date__year = year, func_tst_dummy_bird = ChildStatus)

                    if ChildStatus == 'Current Count':
                        for data  in dataList:
                            if data.func_tst_dummy_bird_date.strftime("%Y") == year and (data.func_tst_dummy_bird == 'Under process' or data.func_tst_dummy_bird=='Observation(same stage)' or data.func_tst_dummy_bird=='Halt'):
                                ListItems.append(data)

                if ParentStatus == 'Road Test':
                    dataList = dataList.extra(
                          select={'year': 'extract (year from road_test_date)',
                          'month': 'extract (month from road_test_date)',
                          'day': 'extract (day from road_test_date)'},
                           order_by=['month','day','-year'])
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(road_test_date__year = year, road_test = ChildStatus)

                    if ChildStatus == 'Current Count':
                        for data  in dataList:
                            if data.road_test_date.strftime("%Y") == year and (data.road_test == 'Under process' or data.road_test=='Observation(same stage)' or data.road_test=='Halt'):
                                ListItems.append(data)

                if ParentStatus == 'Post Road Test':
                    dataList = dataList.extra(
                          select={'year': 'extract (year from post_road_test_date)',
                          'month': 'extract (month from post_road_test_date)',
                          'day': 'extract (day from post_road_test_date)'},
                           order_by=['month','day','-year'])
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(post_road_test_date__year = year, post_road_test = ChildStatus)

                    if ChildStatus == 'Current Count':
                        for data  in dataList:
                            if data.post_road_test_date.strftime("%Y") == year and (data.post_road_test == 'Under process' or data.post_road_test=='Observation(same stage)' or data.post_road_test=='Halt'):
                                ListItems.append(data)

                if ParentStatus == 'Integrated Operation':
                    dataList = dataList.extra(
                          select={'year': 'extract (year from integrated_operation_date)',
                          'month': 'extract (month from integrated_operation_date)',
                          'day': 'extract (day from integrated_operation_date)'},
                           order_by=['month','day','-year'])
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(integrated_operation_date__year = year, integrated_operation = ChildStatus)

                    if ChildStatus == 'Current Count':
                        for data  in dataList:
                            if data.integrated_operation_date.strftime("%Y") == year and (data.integrated_operation == 'Under process' or data.integrated_operation=='Observation(same stage)' or data.integrated_operation=='Halt'):
                                ListItems.append(data)

                if ParentStatus == 'Rain Test':
                    dataList = dataList.extra(
                          select={'year': 'extract (year from rain_test_date)',
                          'month': 'extract (month from rain_test_date)',
                          'day': 'extract (day from rain_test_date)'},
                           order_by=['month','day','-year'])
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(rain_test_date__year = year, rain_test = ChildStatus)

                    if ChildStatus == 'Current Count':
                        for data  in dataList:
                            if data.rain_test_date.strftime("%Y") == year and (data.rain_test == 'Under process' or data.rain_test=='Observation(same stage)' or data.rain_test=='Halt'):
                                ListItems.append(data)

                if ParentStatus == 'Pre User Inspection':
                    dataList = dataList.extra(
                          select={'year': 'extract (year from pre_user_inspection_date)',
                          'month': 'extract (month from pre_user_inspection_date)',
                          'day': 'extract (day from pre_user_inspection_date)'},
                           order_by=['month','day','-year'])
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(pre_user_inspection_date__year = year, pre_user_inspection = ChildStatus)

                    if ChildStatus == 'Current Count':
                        for data  in dataList:
                            if data.pre_user_inspection_date.strftime("%Y") == year and (data.pre_user_inspection == 'Under process' or data.pre_user_inspection=='Observation(same stage)' or data.pre_user_inspection=='Halt'):
                                ListItems.append(data)

                if ParentStatus == 'Final Integration':
                    dataList = dataList.extra(
                          select={'year': 'extract (year from final_integration_date)',
                          'month': 'extract (month from final_integration_date)',
                          'day': 'extract (day from final_integration_date)'},
                           order_by=['month','day','-year'])
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(final_integration_date__year = year, final_integration_status = ChildStatus)

                    if ChildStatus == 'Current Count':
                        for data  in dataList:
                            if data.final_integration_date.strftime("%Y") == year and (data.final_integration_status == 'Under process' or data.final_integration_status=='Observation(same stage)' or data.final_integration_status=='Halt'):
                                ListItems.append(data)

                if ParentStatus == 'Loading/Unloading on MLV/HLF':
                    dataList = dataList.extra(
                          select={'year': 'extract (year from load_unload_on_mlv_hlf_date)',
                          'month': 'extract (month from load_unload_on_mlv_hlf_date)',
                          'day': 'extract (day from load_unload_on_mlv_hlf_date)'},
                           order_by=['month','day','-year'])
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(load_unload_on_mlv_hlf_date__year = year, load_unload_on_mlv_hlf = ChildStatus)

                    if ChildStatus == 'Current Count':
                        for data  in dataList:
                            if data.load_unload_on_mlv_hlf_date.strftime("%Y") == year and (data.load_unload_on_mlv_hlf == 'Under process' or data.load_unload_on_mlv_hlf=='Observation(same stage)' or data.load_unload_on_mlv_hlf=='Halt'):
                                ListItems.append(data)

                if ParentStatus == 'Pre-HIL':
                    dataList = dataList.extra(
                          select={'year': 'extract (year from pre_hil_date)',
                          'month': 'extract (month from pre_hil_date)',
                          'day': 'extract (day from pre_hil_date)'},
                           order_by=['month','day','-year'])
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(pre_hil_date__year = year, pre_hil_status = ChildStatus)

                    if ChildStatus == 'Current Count':
                        for data  in dataList:
                            if data.pre_hil_date.strftime("%Y") == year and (data.pre_hil_status == 'Under process' or data.pre_hil_status=='Observation(same stage)' or data.pre_hil_status=='Halt'):
                                ListItems.append(data)

                if ParentStatus == 'Vibration':
                    dataList = dataList.extra(
                          select={'year': 'extract (year from vibaration_date)',
                          'month': 'extract (month from vibaration_date)',
                          'day': 'extract (day from vibaration_date)'},
                           order_by=['month','day','-year'])
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(vibaration_date__year = year, vibaration_status = ChildStatus)

                    if ChildStatus == 'Current Count':
                        for data  in dataList:
                            if data.vibaration_date.strftime("%Y") == year and (data.vibaration_status == 'Under process' or data.vibaration_status=='Observation(same stage)' or data.vibaration_status=='Halt'):
                                ListItems.append(data)

                if ParentStatus == 'CG Balancing':
                    dataList = dataList.extra(
                          select={'year': 'extract (year from cgbalancing_date)',
                          'month': 'extract (month from cgbalancing_date)',
                          'day': 'extract (day from cgbalancing_date)'},
                           order_by=['month','day','-year'])
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(cgbalancing_date__year = year, cgbalancing_date_status = ChildStatus)

                    if ChildStatus == 'Current Count':
                        for data  in dataList:
                            if data.cgbalancing_date.strftime("%Y") == year and (data.cgbalancing_date_status == 'Under process' or data.cgbalancing_date_status=='Observation(same stage)' or data.cgbalancing_date_status=='Halt'):
                                ListItems.append(data)

                if ParentStatus == 'Post-HIL':
                    dataList = dataList.extra(
                          select={'year': 'extract (year from post_hil_date)',
                          'month': 'extract (month from post_hil_date)',
                          'day': 'extract (day from post_hil_date)'},
                           order_by=['month','day','-year'])
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(post_hil_date__year = year, post_hil_status = ChildStatus)

                    if ChildStatus == 'Current Count':
                        for data  in dataList:
                            if data.post_hil_date.strftime("%Y") == year and (data.post_hil_status == 'Under process' or data.post_hil_status=='Observation(same stage)' or data.post_hil_status=='Halt'):
                                ListItems.append(data)

                if ParentStatus == 'System Alignment':
                    # dataList = dataList.extra(
                    #       select={'year': 'extract (year from sys_align_Date)',
                    #       'month': 'extract (month from sys_align_Date)',
                    #       'day': 'extract (day from sys_align_Date)'},
                    #        order_by=['month','day','-year'])
                    dataList = dataList.annotate(sys_align_Date__month=Extract('sys_align_Date', 'month'),
                        sys_align_Date__year=Extract('sys_align_Date', 'year'),
                        sys_align_Date__day=Extract('sys_align_Date', 'day')).order_by('sys_align_Date__month','sys_align_Date__day','-sys_align_Date__year')

                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(sys_align_Date__year = year, sys_align_status = ChildStatus)

                    if ChildStatus == 'Current Count':
                        for data  in dataList:
                            if data.sys_align_Date.strftime("%Y") == year and (data.sys_align_status == 'Under process' or data.sys_align_status=='Observation(same stage)' or data.sys_align_status=='Halt'):
                                ListItems.append(data)

                if ParentStatus == 'Incapsulation':
                    dataList = dataList.extra(
                          select={'year': 'extract (year from incapsulation_date)',
                          'month': 'extract (month from incapsulation_date)',
                          'day': 'extract (day from incapsulation_date)'},
                           order_by=['month','day','-year'])
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(incapsulation_date__year = year, incapsulation_status = ChildStatus)

                    if ChildStatus == 'Current Count':
                        for data  in dataList:
                            if data.incapsulation_date.strftime("%Y") == year and (data.incapsulation_status == 'Under process' or data.incapsulation_status=='Observation(same stage)' or data.incapsulation_status=='Halt'):
                                ListItems.append(data)

                if ParentStatus == 'Final Integrated Testing':
                    dataList = dataList.extra(
                          select={'year': 'extract (year from final_integration_date)',
                          'month': 'extract (month from final_integration_date)',
                          'day': 'extract (day from final_integration_date)'},
                           order_by=['month','day','-year'])
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(final_integration_date__year = year, final_integration_status = ChildStatus)

                    if ChildStatus == 'Current Count':
                        for data  in dataList:
                            if data.final_integration_date.strftime("%Y") == year and (data.final_integration_status == 'Under process' or data.final_integration_status=='Observation(same stage)' or data.final_integration_status=='Halt'):
                                ListItems.append(data)

                if ParentStatus == 'FGT Status':
                    dataList = dataList.extra(
                          select={'year': 'extract (year from fgt_date)',
                          'month': 'extract (month from fgt_date)',
                          'day': 'extract (day from fgt_date)'},
                           order_by=['month','day','-year'])
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(fgt_date__year = year, fgt_status = ChildStatus)

                    if ChildStatus == 'Current Count':
                        for data  in dataList:
                            if data.fgt_date.strftime("%Y") == year and (data.fgt_status == 'Under process' or data.fgt_status=='Observation(same stage)' or data.fgt_status=='Halt'):
                                ListItems.append(data)

                if ParentStatus == 'BHD Status':
                    dataList = dataList.extra(
                          select={'year': 'extract (year from bhd_date)',
                          'month': 'extract (month from bhd_date)',
                          'day': 'extract (day from bhd_date)'},
                           order_by=['month','day','-year'])
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(bhd_date__year = year, bhd_status = ChildStatus)

                    if ChildStatus == 'Current Count':
                        for data  in dataList:
                            if data.bhd_date.strftime("%Y") == year and (data.bhd_status == 'Ok' or data.bhd_status=='Not Submitted)' or data.bhd_status=='QM Observations Forwarded'):
                                ListItems.append(data)

                if ParentStatus == 'FQM Status':
                    dataList = dataList.extra(
                          select={'year': 'extract (year from fqm_date)',
                          'month': 'extract (month from fqm_date)',
                          'day': 'extract (day from fqm_date)'},
                           order_by=['month','day','-year'])
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(fqm_date__year = year, fqm_status = ChildStatus)

                    if ChildStatus == 'Current Count':
                        for data  in dataList:
                            if data.fqm_date.strftime("%Y") == year and (data.fqm_status=='Not Conducted)'):
                                ListItems.append(data)

                if ParentStatus == 'QM Certification Status':
                    dataList = dataList.extra(
                          select={'year': 'extract (year from qm_certification_date)',
                          'month': 'extract (month from qm_certification_date)',
                          'day': 'extract (day from qm_certification_date)'},
                           order_by=['month','day','-year'])
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(qm_certification_date__year = year, qm_certification_status = ChildStatus)

                    if ChildStatus == 'Current Count':
                        for data  in dataList:
                            if data.qm_certification_date.strftime("%Y") == year and (data.qm_certification_status == 'QM certificate issued' or data.qm_certification_status=='QM Observations Forwarded)'):
                                ListItems.append(data)

                if ParentStatus == 'Launch/ End User':
                    dataList = dataList.extra(
                          select={'year': 'extract (year from enduser_date)',
                          'month': 'extract (month from enduser_date)',
                          'day': 'extract (day from enduser_date)'},
                           order_by=['month','day','-year'])
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(enduser_date__year = year, enduser_status = ChildStatus)
                    else:
                        for data  in dataList:
                            if data.enduser_date.strftime("%Y") == year and (data.enduser_status == 'Ok' or data.enduser_status=='Observation' or data.enduser_status=='Halt'):
                                ListItems.append(data)

                serializer = ProductionSystemSerialzer(ListItems, many=True)
                return JsonResponse({'message': 'Welcome to Home Page', 'data': serializer.data}, status=200)
        except Exception as e:
            print(e)
            return JsonResponse({'status': 'false'}, status=200)

    @staticmethod
    def GetProductionObservationStatus(request):
        try:
            org = request.query_params['selected_org']
            year = request.query_params['selected_year']
            type = request.query_params['selected_type']
            system = request.query_params['selected_system']
            SelectedStatus = request.query_params['selected_status']
            sys_name = request.query_params['sys_name']
            child_status = request.query_params['child_status']
            filter_objects = Q()

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

            # create dynamic filter

            if org != '':
                filter_objects &= get_filter(
                    'organization', 'equal',
                    org)
            if type != '':
                filter_objects &= get_filter(
                    'sys_type', 'equal',
                    type)
            if system != '':
                filter_objects &= get_filter(
                    'system', 'equal',
                    system)


            dataList = ""
            if sys_name == 'Production':
                dataList = ProductionSystemStatus.objects.filter(filter_objects)
            if sys_name == 'Flight':
                dataList = FlightSystemStatus.objects.filter(filter_objects)
            if sys_name == 'Relifing':
                dataList = RelifingSystemStatus.objects.filter(filter_objects)
            if SelectedStatus != '':
                ListItems = []
                if SelectedStatus == 'BLT':
                    for data  in dataList:
                        if data.blt_date is not None and data.blt_date.strftime("%Y") == year and data.blt_status==child_status:
                            ListItems.append(data)
                if SelectedStatus == 'Pre-HIL':
                    for data  in dataList:
                        if data.pre_hil_date is not None and data.pre_hil_date.strftime("%Y") == year and data.pre_hil_status==child_status:
                            ListItems.append(data)

                if SelectedStatus == 'Vibration':
                    for data  in dataList:
                        if data.vibaration_date is not None and data.vibaration_date.strftime("%Y") == year and data.vibaration_status==child_status:
                            ListItems.append(data)

                if SelectedStatus == 'CG Balancing':
                    for data  in dataList:
                        if data.cgbalancing_date is not None and data.cgbalancing_date.strftime("%Y") == year and data.cgbalancing_date_status==child_status:
                            ListItems.append(data)

                if SelectedStatus == 'EMP Proofing':
                    for data  in dataList:
                        if data.emp_proofing_date is not None and data.emp_proofing_date.strftime("%Y") == year and data.emp_proofing==child_status:
                            ListItems.append(data)

                if SelectedStatus == 'Functional Test W/O Dummy Bird':
                    for data in dataList:
                        if data.func_tst_date is not None and data.func_tst_date.strftime(
                                "%Y") == year and data.func_tst ==child_status:
                            ListItems.append(data)

                if SelectedStatus == 'Functional Test With Dummy Bird':
                    for data  in dataList:
                        if data.func_tst_dummy_bird_date is not None and data.func_tst_dummy_bird_date.strftime("%Y") == year and data.func_tst_dummy_bird==child_status:
                            ListItems.append(data)

                if SelectedStatus == 'Road Test':
                    for data  in dataList:
                        if data.road_test_date is not None and data.road_test_date.strftime("%Y") == year and data.road_test==child_status:
                            ListItems.append(data)

                if SelectedStatus == 'Post Road Test':
                    for data  in dataList:
                        if data.post_road_test_date is not None and data.post_road_test_date.strftime("%Y") == year and data.post_road_test==child_status:
                            ListItems.append(data)

                if SelectedStatus == 'Integrated Operation':
                    for data in dataList:
                        if data.integrated_operation_date is not None and data.integrated_operation_date.strftime(
                                "%Y") == year and data.integrated_operation ==child_status:
                            ListItems.append(data)

                if SelectedStatus == 'Rain Test':
                    for data in dataList:
                        if data.rain_test_date is not None and data.rain_test_date.strftime(
                                "%Y") == year and data.rain_test ==child_status:
                            ListItems.append(data)

                if SelectedStatus == 'Pre User Inspection':
                    for data  in dataList:
                        if data.pre_user_inspection_date is not None and data.pre_user_inspection_date.strftime("%Y") == year and data.pre_user_inspection==child_status:
                            ListItems.append(data)

                if SelectedStatus == 'Final Integration':
                    for data in dataList:
                        if data.final_integration_date is not None and data.final_integration_date.strftime(
                                "%Y") == year and data.final_integration_status == child_status:
                            ListItems.append(data)

                if SelectedStatus == 'Loading/Unloading on MLV/HLF':
                    for data  in dataList:
                        if data.load_unload_on_mlv_hlf_date is not None and data.load_unload_on_mlv_hlf_date.strftime("%Y") == year and data.load_unload_on_mlv_hlf==child_status:
                            ListItems.append(data)

                if SelectedStatus == 'Post-HIL':
                    for data  in dataList:
                        if data.post_hil_date is not None and data.post_hil_date.strftime("%Y") == year and data.post_hil_status==child_status:
                            ListItems.append(data)

                if SelectedStatus == 'System Alignment':
                    for data  in dataList:
                        if data.sys_align_Date is not None and data.sys_align_Date.strftime("%Y") == year and data.sys_align_status==child_status:
                            ListItems.append(data)

                if SelectedStatus == 'Incapsulation':
                    for data  in dataList:
                        if data.incapsulation_date is not None and data.incapsulation_date.strftime("%Y") == year and data.incapsulation_status==child_status:
                            ListItems.append(data)

                if SelectedStatus == 'Final Integrated Testing':
                    for data  in dataList:
                        if data.final_integrated_testing_date is not None and data.final_integrated_testing_date.strftime("%Y") == year and data.final_integrated_testing==child_status:
                            ListItems.append(data)

                if SelectedStatus == 'FGT Status':
                    for data  in dataList:
                        if data.fgt_date is not None and data.fgt_date.strftime("%Y") == year and data.fgt_status==child_status:
                            ListItems.append(data)
                print(ListItems)
                serializer = ""
                if sys_name == 'Production':
                    serializer = ProductionSystemSerialzer(ListItems, many=True)
                if sys_name == 'Flight':
                    serializer = FlightSystemSerialzer(ListItems, many=True)
                if sys_name == 'Relifing':
                    serializer = RelifingSystemSerialzer(ListItems, many=True)
                print(serializer.data)
                return JsonResponse({'message': 'Welcome to Home Page', 'data': serializer.data}, status=200)

        except Exception as e:
            print(e)
            return JsonResponse({'status': 'false'}, status=200)
    @staticmethod
    def GetProductionListHistory(request):
        try:
            id = request.query_params['id']
            ParentStatus = request.query_params['parent_status']
            dataList = ProductionSystemStatusHistory.objects.filter(prod_id=id)
            if ParentStatus == 'BLT':
                dataList = dataList.extra(
                    select={
                        'year': 'extract (year from blt_date)',
                        'month': 'extract (month from blt_date)',
                        'day': 'extract (day from blt_date)'},
                    order_by=['month', 'day', '-year'])


            if ParentStatus == 'EMP Proofing':
                dataList = dataList.extra(
                    select={'year': 'extract (year from emp_proofing_date)',
                            'month': 'extract (month from emp_proofing_date)',
                            'day': 'extract (day from emp_proofing_date)'},
                    order_by=['month', 'day', '-year'])

            if ParentStatus == 'Functional Test W/O Dummy Bird':
                dataList = dataList.extra(
                    select={'year': 'extract (year from func_tst_date)',
                            'month': 'extract (month from func_tst_date)',
                            'day': 'extract (day from func_tst_date)'},
                    order_by=['month', 'day', '-year'])

            if ParentStatus == 'Functional Test With Dummy Bird':
                dataList = dataList.extra(
                    select={'year': 'extract (year from func_tst_dummy_bird_date)',
                            'month': 'extract (month from func_tst_dummy_bird_date)',
                            'day': 'extract (day from func_tst_dummy_bird_date)'},
                    order_by=['month', 'day', '-year'])

            if ParentStatus == 'Road Test':
                dataList = dataList.extra(
                    select={'year': 'extract (year from road_test_date)',
                            'month': 'extract (month from road_test_date)',
                            'day': 'extract (day from road_test_date)'},
                    order_by=['month', 'day', '-year'])

            if ParentStatus == 'Post Road Test':
                dataList = dataList.extra(
                    select={'year': 'extract (year from post_road_test_date)',
                            'month': 'extract (month from post_road_test_date)',
                            'day': 'extract (day from post_road_test_date)'},
                    order_by=['month', 'day', '-year'])

            if ParentStatus == 'Integrated Operation':
                dataList = dataList.extra(
                    select={'year': 'extract (year from integrated_operation_date)',
                            'month': 'extract (month from integrated_operation_date)',
                            'day': 'extract (day from integrated_operation_date)'},
                    order_by=['month', 'day', '-year'])

            if ParentStatus == 'Rain Test':
                dataList = dataList.extra(
                    select={'year': 'extract (year from rain_test_date)',
                            'month': 'extract (month from rain_test_date)',
                            'day': 'extract (day from rain_test_date)'},
                    order_by=['month', 'day', '-year'])

            if ParentStatus == 'Pre User Inspection':
                dataList = dataList.extra(
                    select={'year': 'extract (year from pre_user_inspection_date)',
                            'month': 'extract (month from pre_user_inspection_date)',
                            'day': 'extract (day from pre_user_inspection_date)'},
                    order_by=['month', 'day', '-year'])

            if ParentStatus == 'Final Integration':
                dataList = dataList.extra(
                    select={'year': 'extract (year from final_integration_date)',
                            'month': 'extract (month from final_integration_date)',
                            'day': 'extract (day from final_integration_date)'},
                    order_by=['month', 'day', '-year'])

            if ParentStatus == 'Loading/Unloading on MLV/HLF':
                dataList = dataList.extra(
                    select={'year': 'extract (year from load_unload_on_mlv_hlf_date)',
                            'month': 'extract (month from load_unload_on_mlv_hlf_date)',
                            'day': 'extract (day from load_unload_on_mlv_hlf_date)'},
                    order_by=['month', 'day', '-year'])

            if ParentStatus == 'Pre-HIL':
                dataList = dataList.extra(
                    select={'year': 'extract (year from pre_hil_date)',
                            'month': 'extract (month from pre_hil_date)',
                            'day': 'extract (day from pre_hil_date)'},
                    order_by=['month', 'day', '-year'])

            if ParentStatus == 'Vibration':
                dataList = dataList.extra(
                    select={'year': 'extract (year from vibaration_date)',
                            'month': 'extract (month from vibaration_date)',
                            'day': 'extract (day from vibaration_date)'},
                    order_by=['month', 'day', '-year'])

            if ParentStatus == 'CG Balancing':
                dataList = dataList.extra(
                    select={'year': 'extract (year from cgbalancing_date)',
                            'month': 'extract (month from cgbalancing_date)',
                            'day': 'extract (day from cgbalancing_date)'},
                    order_by=['month', 'day', '-year'])

            if ParentStatus == 'Post-HIL':
                dataList = dataList.extra(
                    select={'year': 'extract (year from post_hil_date)',
                            'month': 'extract (month from post_hil_date)',
                            'day': 'extract (day from post_hil_date)'},
                    order_by=['month', 'day', '-year'])

            if ParentStatus == 'System Alignment':
                dataList = dataList.extra(
                    select={'year': 'extract (year from sys_align_Date)',
                            'month': 'extract (month from sys_align_Date)',
                            'day': 'extract (day from sys_align_Date)'},
                    order_by=['month', 'day', '-year'])


            if ParentStatus == 'Incapsulation':
                dataList = dataList.extra(
                    select={'year': 'extract (year from incapsulation_date)',
                            'month': 'extract (month from incapsulation_date)',
                            'day': 'extract (day from incapsulation_date)'},
                    order_by=['month', 'day', '-year'])


            if ParentStatus == 'Final Integrated Testing':
                dataList = dataList.extra(
                    select={'year': 'extract (year from final_integration_date)',
                            'month': 'extract (month from final_integration_date)',
                            'day': 'extract (day from final_integration_date)'},
                    order_by=['month', 'day', '-year'])


            if ParentStatus == 'FGT Status':
                dataList = dataList.extra(
                    select={'year': 'extract (year from fgt_date)',
                            'month': 'extract (month from fgt_date)',
                            'day': 'extract (day from fgt_date)'},
                    order_by=['month', 'day', '-year'])

            if ParentStatus == 'BHD Status':
                dataList = dataList.extra(
                    select={'year': 'extract (year from bhd_date)',
                            'month': 'extract (month from bhd_date)',
                            'day': 'extract (day from bhd_date)'},
                    order_by=['month', 'day', '-year'])

            if ParentStatus == 'FQM Status':
                dataList = dataList.extra(
                    select={'year': 'extract (year from fqm_date)',
                            'month': 'extract (month from fqm_date)',
                            'day': 'extract (day from fqm_date)'},
                    order_by=['month', 'day', '-year'])


            if ParentStatus == 'QM Certification Status':
                dataList = dataList.extra(
                    select={'year': 'extract (year from qm_certification_date)',
                            'month': 'extract (month from qm_certification_date)',
                            'day': 'extract (day from qm_certification_date)'},
                    order_by=['month', 'day', '-year'])

            if ParentStatus == 'Launch/ End User':
                dataList = dataList.extra(
                    select={'year': 'extract (year from enduser_date)',
                            'month': 'extract (month from enduser_date)',
                            'day': 'extract (day from enduser_date)'},
                    order_by=['month', 'day', '-year'])

            serializer = ProductionSystemSerialzer(dataList, many=True)
            return JsonResponse({'status': 'true', 'data': serializer.data}, status=200)

        except Exception as e:
            print(e)
            return JsonResponse({'status': 'false'}, status=200)

    # API for PDF generator with table and Text_wraping

    @staticmethod
    def GetProductionPDFList(request):
        data = []
        status = request.query_params['status']
        if status == 'Total':
            data = ProductionSystemStatus.objects.all().order_by('-id').values_list('system', 'organization', 'set_id',
                                                                                    'blt_status', 'pre_hil_status',
                                                                                    'vibaration_status',
                                                                                    'post_hil_status',
                                                                                    'fgt_status',
                                                                                    'final_integration_status',
                                                                                    'bhd_status', 'fqm_status',
                                                                                    'qm_certification_status',
                                                                                    'remarks')
        else:
            data = ProductionSystemStatus.objects.filter(qm_certification_status=status).order_by('-id').values_list(
                'system', 'organization', 'set_id',
                'blt_status', 'pre_hil_status',
                'vibaration_status',
                'post_hil_status',
                'fgt_status',
                'final_integration_status',
                'bhd_status', 'fqm_status',
                'qm_certification_status',
                'remarks')
        TABLE_COL_NAMES = (
            "System", "Organization", "Set ID", "BLT Status", "Pre-Hil Status", "Vibration Status", "Post-Hil Status",
            "FGT Status", "Final Integration Status", "BHD Status", "FQM Status", "QM Certification Status", "Remarks")

        pdf = FPDF('L', 'mm', 'Legal')
        pdf.add_page()
        pdf.set_font('courier', 'B', 26)
        pdf.cell(330, 10, 'Production System Report', border=0, align='C', ln=2)
        pdf.cell(40, 10, '', 0, 1)
        pdf.set_font("Times", size=10)
        line_height = pdf.font_size * 3.2
        col_width = pdf.epw / 14

        def render_table_header():
            pdf.set_font(style="B")
            for col_name in TABLE_COL_NAMES:
                pdf.multi_cell(col_width, line_height, col_name, border=1, align='C', ln=3,
                               max_line_height=pdf.font_size)
            pdf.ln(line_height)
            pdf.set_font(style="")

        render_table_header()

        lh_list = []
        use_default_height = 0
        for row in data:
            for datum in row:
                dd = str(datum)
                word_list = dd.split()
                number_of_words = len(word_list)
                if number_of_words > 2:
                    use_default_height = 1
                    new_line_height = pdf.font_size * (number_of_words / 2)
            if not use_default_height:
                lh_list.append(line_height)
            else:
                lh_list.append(new_line_height)
                use_default_height = 0

        for j, row in enumerate(data):
            line_height = lh_list[j]
            if pdf.will_page_break(line_height):
                render_table_header()
            for col_num in range(len(row)):
                line_height = lh_list[j]
                pdf.multi_cell(col_width, line_height, f"{row[col_num]}", border=1, align='C', ln=3,
                               max_line_height=pdf.font_size)
            pdf.ln(line_height)
        pdf.output('ProductionSystem.pdf')
        return FileResponse(open('ProductionSystem.pdf', 'rb'), as_attachment=True, content_type='application/pdf')

    @staticmethod
    def GetProductionExcelList(request):

        data = []
        status = request.query_params['status']
        if status == 'Total':
            data = ProductionSystemStatus.objects.all().order_by('-id').values_list('system', 'organization', 'set_id',
                                                                                    'blt_status', 'pre_hil_status',
                                                                                    'vibaration_status',
                                                                                    'post_hil_status',
                                                                                    'fgt_status',
                                                                                    'final_integration_status',
                                                                                    'bhd_status', 'fqm_status',
                                                                                    'qm_certification_status',
                                                                                    'remarks')
        else:
            data = ProductionSystemStatus.objects.filter(qm_certification_status=status).order_by('-id').values_list(
                'system', 'organization', 'set_id',
                'blt_status', 'pre_hil_status',
                'vibaration_status',
                'post_hil_status',
                'fgt_status',
                'final_integration_status',
                'bhd_status', 'fqm_status',
                'qm_certification_status',
                'remarks')
        response = HttpResponse(content_type='application/ms-excel')
        response['Content-Disposition'] = 'attachment; filename="DocumentExcelSheet.xls"'
        work_book = xlwt.Workbook(encoding='utf-8')
        work_sheet = work_book.add_sheet('ExcelSheet')
        row_num = 0
        font_style = xlwt.XFStyle()
        font_style.font.bold = True
        TABLE_COL_NAMES = (
            "System", "Organization", "Set ID", "BLT Status", "Pre-Hil Status", "Vibration Status", "Post-Hil Status",
            "FGT Status", "Final Integration Status", "BHD Status", "FQM Status", "QM Certification Status", "Remarks")

        for col_num in range(len(TABLE_COL_NAMES)):
            work_sheet.write(row_num, col_num, TABLE_COL_NAMES[col_num], font_style)
        font_style = xlwt.XFStyle()

        for row in data:
            row_num += 1
            for col_num in range(len(row)):
                work_sheet.write(row_num, col_num, str(row[col_num]), font_style)
        work_book.save(response)
        return response

    @staticmethod
    def DeleteProdSys(request):

        try:
            id = request.query_params['id']
            delete = ProductionSystemStatus.objects.filter(id=id).delete()

            return JsonResponse({'status': 'True', 'message': "Record Deleted"},
                                status=200)

        except Exception as e:
            print(e)
            return JsonResponse({'status': 'False', "message": "Doc Not Deleted"}, status=500)
            pass

    @staticmethod
    def DeleteFlightSys(request):

        try:
            id = request.query_params['id']
            delete = FlightSystemStatus.objects.filter(id=id).delete()
            return JsonResponse({'status': 'True', 'message': "Record Deleted"},
                                status=200)

        except Exception as e:
            print(e)
            return JsonResponse({'status': 'False', "message": "Doc Not Deleted"}, status=500)
            pass

    @staticmethod
    def DeleteRelifingSys(request):

        try:
            id = request.query_params['id']
            delete = RelifingSystemStatus.objects.filter(id=id).delete()
            return JsonResponse({'status': 'True', 'message': "Record Deleted"},
                                status=200)

        except Exception as e:
            print(e)
            return JsonResponse({'status': 'False', "message": "Doc Not Deleted"}, status=500)
            pass

    @staticmethod
    def AddFlightStatus(request):
        is_active = 0
        if request['is_active'] == 'true':
            is_active = 1
        flightModel = FlightSystemStatus()

        # try:
        id = request['id']
        if id == '0':
            flightModel.system = request['system']
            flightModel.organization = request['organization']
            flightModel.set_id = request['set_id']
            if request['blt_date'] != '':
                flightModel.blt_date = request['blt_date']
            flightModel.testing_date = request['testing_date']
            flightModel.sys_type = request['sys_type']
            flightModel.blt_status = request['blt_status']
            flightModel.blt_remarks = request['blt_remarks']
            if request['pre_hil_date'] != '':
                flightModel.pre_hil_date = request['pre_hil_date']
            flightModel.pre_hil_status = request['pre_hil_status']
            flightModel.pre_hil_remarks = request['pre_hil_remarks']
            if request['vibaration_date'] != '':
                flightModel.vibaration_date = request['vibaration_date']
            flightModel.vibaration_status = request['vibaration_status']
            flightModel.vibaration_remarks = request['vibaration_remarks']
            if request['post_hil_date'] != '':
                flightModel.post_hil_date = request['post_hil_date']
            flightModel.post_hil_status = request['post_hil_status']
            flightModel.post_hil_remarks = request['post_hil_remarks']
            if request['fgt_date'] != '':
                flightModel.fgt_date = request['fgt_date']
            flightModel.fgt_status = request['fgt_status']
            flightModel.fgt_remarks = request['fgt_remarks']
            if request['final_integration_date'] != '':
                flightModel.final_integration_date = request['final_integration_date']
            flightModel.final_integration_status = request['final_integration_status']
            flightModel.final_integration_remarks = request['final_integration_remarks']
            if request['bhd_date'] != '':
                flightModel.bhd_date = request['bhd_date']
            flightModel.bhd_status = request['bhd_status']
            flightModel.bhd_remarks = request['bhd_remarks']
            if request['fqm_date'] != '':
                flightModel.fqm_date = request['fqm_date']
            flightModel.fqm_status = request['fqm_status']
            flightModel.fqm_remarks = request['fqm_remarks']
            if request['qm_certification_date'] != '':
                flightModel.qm_certification_date = request['qm_certification_date']
            flightModel.qm_certification_status = request['qm_certification_status']
            flightModel.qm_certification_remarks = request['qm_certification_remarks']
            flightModel.attachment = request['attachment']
            flightModel.remarks = request['remarks']
            if request['cgbalancing_date'] != '':
                flightModel.cgbalancing_date = request['cgbalancing_date']
            flightModel.cgbalancing_date_status = request['cgbalancing_date_status']
            flightModel.cgbalancing_date_remarks = request['cgbalancing_date_remarks']
            if request['launchact_date'] != '':
                flightModel.launchact_date = request['launchact_date']
            flightModel.launchact_status = request['launchact_status']
            flightModel.launchact_remarks = request['launchact_remarks']
            if request['incapsulation_date'] != '':
                flightModel.incapsulation_date = request['incapsulation_date']
            flightModel.incapsulation_status = request['incapsulation_status']
            flightModel.incapsulation_remarks = request['incapsulation_remarks']
            if request['sys_align_Date'] != '':
                flightModel.sys_align_Date = request['sys_align_Date']
            flightModel.sys_align_status = request['sys_align_status']
            flightModel.sys_align_remarks = request['sys_align_remarks']
            flightModel.testing_type = request['testing_type']
            flightModel.emp_proofing = request['emp_proofing']
            flightModel.func_tst = request['func_tst']
            flightModel.func_tst_dummy_bird = request['func_tst_dummy_bird']
            flightModel.road_test = request['road_test']
            flightModel.post_road_test = request['post_road_test']
            flightModel.integrated_operation = request['integrated_operation']
            flightModel.rain_test = request['rain_test']
            flightModel.pre_user_inspection = request['pre_user_inspection']
            flightModel.final_integrated_testing = request['final_integrated_testing']
            flightModel.load_unload_on_mlv_hlf = request['load_unload_on_mlv_hlf']
            if request['emp_proofing_date'] != '':
                flightModel.emp_proofing_date = request['emp_proofing_date']
            flightModel.emp_proofing_remarks = request['emp_proofing_remarks']
            if request['func_tst_date'] != '':
                flightModel.func_tst_date = request['func_tst_date']
            flightModel.func_tst_remarks = request['func_tst_remarks']
            if request['func_tst_dummy_bird_date'] != '':
                flightModel.func_tst_dummy_bird_date = request['func_tst_dummy_bird_date']
            flightModel.func_tst_dummy_bird_remarks = request['func_tst_dummy_bird_remarks']
            if request['road_test_date'] != '':
                flightModel.road_test_date = request['road_test_date']
            flightModel.road_test_remarks = request['road_test_remarks']
            if request['post_road_test_date'] != '':
                flightModel.post_road_test_date = request['post_road_test_date']
            flightModel.post_road_test_remarks = request['post_road_test_remarks']
            if request['integrated_operation_date'] != '':
                flightModel.integrated_operation_date = request['integrated_operation_date']
            flightModel.integrated_operation_remarks = request['integrated_operation_remarks']
            if request['rain_test_date'] != '':
                flightModel.rain_test_date = request['rain_test_date']
            flightModel.rain_test_remarks = request['rain_test_remarks']
            if request['pre_user_inspection_date'] != '':
                flightModel.pre_user_inspection_date = request['pre_user_inspection_date']
            flightModel.pre_user_inspection_remarks = request['pre_user_inspection_remarks']
            if request['final_integrated_testing_date'] != '':
                flightModel.final_integrated_testing_date = request['final_integrated_testing_date']
            flightModel.final_integrated_testing_remarks = request['final_integrated_testing_remarks']
            if request['load_unload_on_mlv_hlf_date'] != '':
                flightModel.load_unload_on_mlv_hlf_date = request['load_unload_on_mlv_hlf_date']
            flightModel.load_unload_on_mlv_hlf_remarks = request['load_unload_on_mlv_hlf_remarks']
            if request['user_id'] != '':
                flightModel.user_id = request['user_id']
            else:
                flightModel.user_id = None
            flightModel.isActive = is_active
            flightModel.save()
            return JsonResponse({'status': 'True', 'message': "Production Status Added Successfully!"},
                                status=200)
        else:
            get_prod = FlightSystemStatus.objects.filter(id=id).first()
            if get_prod is not None:

                if request['blt_status'] != get_prod.blt_status \
                        or request['blt_date'] != get_prod.blt_date \
                        or request['blt_remarks'] != get_prod.blt_remarks \
                        or request['pre_hil_status'] != get_prod.pre_hil_status\
                        or request['pre_hil_date'] != get_prod.pre_hil_date\
                        or request['pre_hil_remarks'] != get_prod.pre_hil_remarks\
                        or request['vibaration_status'] != get_prod.vibaration_status \
                        or request['vibaration_date'] != get_prod.vibaration_date \
                        or request['vibaration_remarks'] != get_prod.vibaration_remarks \
                        or request['post_hil_status'] != get_prod.post_hil_status \
                        or request['post_hil_date'] != get_prod.post_hil_date \
                        or request['post_hil_remarks'] != get_prod.post_hil_remarks \
                        or request['fgt_status'] != get_prod.fgt_status \
                        or request['fgt_date'] != get_prod.fgt_date \
                        or request['fgt_remarks'] != get_prod.fgt_remarks \
                        or request['final_integration_status'] != get_prod.final_integration_status \
                        or request['final_integration_date'] != get_prod.final_integration_date \
                        or request['final_integration_remarks'] != get_prod.final_integration_remarks \
                        or request['bhd_status'] != get_prod.bhd_status \
                        or request['bhd_date'] != get_prod.bhd_date \
                        or request['bhd_remarks'] != get_prod.bhd_remarks \
                        or request['fqm_status'] != get_prod.fqm_status\
                        or request['fqm_date'] != get_prod.fqm_date \
                        or request['fqm_remarks'] != get_prod.fqm_remarks\
                        or request['qm_certification_status'] != get_prod.qm_certification_status  \
                        or request['qm_certification_date'] != get_prod.qm_certification_date \
                        or request['qm_certification_remarks'] != get_prod.qm_certification_remarks  \
                        or request['incapsulation_status'] != get_prod.incapsulation_status  \
                        or request['incapsulation_date'] != get_prod.incapsulation_date \
                        or request['incapsulation_remarks'] != get_prod.incapsulation_remarks  \
                        or request['sys_align_status'] != get_prod.sys_align_status  \
                        or request['sys_align_Date'] != get_prod.sys_align_Date \
                        or request['sys_align_remarks'] != get_prod.sys_align_remarks  \
                        or request['emp_proofing'] != get_prod.emp_proofing \
                        or request['emp_proofing_date'] != get_prod.emp_proofing_date \
                        or request['emp_proofing_remarks'] != get_prod.emp_proofing_remarks \
                        or request['func_tst'] != get_prod.func_tst \
                        or request['func_tst_date'] != get_prod.func_tst_date \
                        or request['func_tst_remarks'] != get_prod.func_tst_remarks \
                        or request['func_tst_dummy_bird'] != get_prod.func_tst_dummy_bird \
                        or request['func_tst_dummy_bird_date'] != get_prod.func_tst_dummy_bird_date \
                        or request['func_tst_dummy_bird_remarks'] != get_prod.func_tst_dummy_bird_remarks \
                        or request['road_test'] != get_prod.road_test \
                        or request['road_test_date'] != get_prod.road_test_date \
                        or request['road_test_remarks'] != get_prod.road_test_remarks \
                        or request['post_road_test'] != get_prod.post_road_test \
                        or request['post_road_test_date'] != get_prod.post_road_test_date \
                        or request['post_road_test_remarks'] != get_prod.post_road_test_remarks \
                        or request['integrated_operation'] != get_prod.integrated_operation \
                        or request['integrated_operation_date'] != get_prod.integrated_operation_date \
                        or request['integrated_operation_remarks'] != get_prod.integrated_operation_remarks \
                        or request['rain_test'] != get_prod.rain_test \
                        or request['rain_test_date'] != get_prod.rain_test_date \
                        or request['rain_test_remarks'] != get_prod.rain_test_remarks \
                        or request['pre_user_inspection'] != get_prod.pre_user_inspection \
                        or request['pre_user_inspection_date'] != get_prod.pre_user_inspection_date \
                        or request['pre_user_inspection_remarks'] != get_prod.pre_user_inspection_remarks \
                        or request['final_integrated_testing'] != get_prod.final_integrated_testing \
                        or request['final_integrated_testing_date'] != get_prod.final_integrated_testing_date \
                        or request['final_integrated_testing_remarks'] != get_prod.final_integrated_testing_remarks \
                        or request['load_unload_on_mlv_hlf'] != get_prod.load_unload_on_mlv_hlf \
                        or request['load_unload_on_mlv_hlf_date'] != get_prod.load_unload_on_mlv_hlf_date \
                        or request['load_unload_on_mlv_hlf_remarks'] != get_prod.load_unload_on_mlv_hlf_remarks \
                        or request['launchact_status'] != get_prod.launchact_status  \
                        or request['launchact_remarks'] != get_prod.launchact_remarks  \
                        or request['launchact_date'] != get_prod.launchact_date:


                    print("add data in flight history")
                    FlightHistoryModal = FlightSystemStatusHistory()
                    FlightHistoryModal.f_id = get_prod.id
                    FlightHistoryModal.system = get_prod.system
                    FlightHistoryModal.organization = get_prod.organization
                    FlightHistoryModal.set_id = get_prod.set_id
                    if request['blt_date'] != '':
                        FlightHistoryModal.blt_date = get_prod.blt_date
                    FlightHistoryModal.blt_status = get_prod.blt_status
                    FlightHistoryModal.blt_remarks = get_prod.blt_remarks
                    FlightHistoryModal.testing_date = get_prod.testing_date
                    FlightHistoryModal.sys_type = get_prod.sys_type
                    if request['pre_hil_date'] != '':
                        FlightHistoryModal.pre_hil_date = get_prod.pre_hil_date
                    FlightHistoryModal.pre_hil_status = get_prod.pre_hil_status
                    FlightHistoryModal.pre_hil_remarks = get_prod.pre_hil_remarks
                    if request['vibaration_date'] != '':
                        FlightHistoryModal.vibaration_date = get_prod.vibaration_date
                    FlightHistoryModal.vibaration_status = get_prod.vibaration_status
                    FlightHistoryModal.vibaration_remarks = get_prod.vibaration_remarks
                    if request['post_hil_date'] != '':
                        FlightHistoryModal.post_hil_date = get_prod.post_hil_date
                    FlightHistoryModal.post_hil_status = get_prod.post_hil_status
                    FlightHistoryModal.post_hil_remarks = get_prod.post_hil_remarks
                    if request['fgt_date'] != '':
                        FlightHistoryModal.fgt_date = get_prod.fgt_date
                    FlightHistoryModal.fgt_status = get_prod.fgt_status
                    FlightHistoryModal.fgt_remarks = get_prod.fgt_remarks
                    if request['final_integration_date'] != '':
                        FlightHistoryModal.final_integration_date = get_prod.final_integration_date
                    FlightHistoryModal.final_integration_status = get_prod.final_integration_status
                    FlightHistoryModal.final_integration_remarks = get_prod.final_integration_remarks
                    if request['bhd_date'] != '':
                        FlightHistoryModal.bhd_date = get_prod.bhd_date
                    FlightHistoryModal.bhd_status = get_prod.bhd_status
                    FlightHistoryModal.bhd_remarks = get_prod.bhd_remarks
                    if request['fqm_date'] != '':
                        FlightHistoryModal.fqm_date = get_prod.fqm_date
                    FlightHistoryModal.fqm_status = get_prod.fqm_status
                    FlightHistoryModal.fqm_remarks = get_prod.fqm_remarks
                    if request['qm_certification_date'] != '':
                        FlightHistoryModal.qm_certification_date = get_prod.qm_certification_date
                    FlightHistoryModal.qm_certification_status = get_prod.qm_certification_status
                    FlightHistoryModal.qm_certification_remarks = get_prod.qm_certification_remarks
                    FlightHistoryModal.attachment = get_prod.attachment
                    FlightHistoryModal.remarks = get_prod.remarks
                    FlightHistoryModal.isActive = get_prod.isActive
                    if request['cgbalancing_date'] != '':
                        FlightHistoryModal.cgbalancing_date = get_prod.cgbalancing_date
                    FlightHistoryModal.cgbalancing_date_status = get_prod.cgbalancing_date_status
                    FlightHistoryModal.cgbalancing_date_remarks = get_prod.cgbalancing_date_remarks
                    if request['launchact_date'] != '':
                        FlightHistoryModal.launchact_date = get_prod.launchact_date
                    FlightHistoryModal.launchact_status = get_prod.launchact_status
                    FlightHistoryModal.launchact_remarks = get_prod.launchact_remarks
                    if request['incapsulation_date'] != '':
                        FlightHistoryModal.incapsulation_date = get_prod.incapsulation_date
                    FlightHistoryModal.incapsulation_status = get_prod.incapsulation_status
                    FlightHistoryModal.incapsulation_remarks = get_prod.incapsulation_remarks
                    if request['sys_align_Date'] != '':
                        FlightHistoryModal.sys_align_Date = get_prod.sys_align_Date
                    FlightHistoryModal.sys_align_status = get_prod.sys_align_status
                    FlightHistoryModal.sys_align_remarks = get_prod.sys_align_remarks
                    FlightHistoryModal.testing_type = get_prod.testing_type
                    FlightHistoryModal.emp_proofing = get_prod.emp_proofing
                    FlightHistoryModal.func_tst = get_prod.func_tst
                    FlightHistoryModal.func_tst_dummy_bird = get_prod.func_tst_dummy_bird
                    FlightHistoryModal.road_test = get_prod.road_test
                    FlightHistoryModal.post_road_test = get_prod.post_road_test
                    FlightHistoryModal.integrated_operation = get_prod.integrated_operation
                    FlightHistoryModal.rain_test = get_prod.rain_test
                    FlightHistoryModal.pre_user_inspection = get_prod.pre_user_inspection
                    FlightHistoryModal.final_integrated_testing = get_prod.final_integrated_testing
                    FlightHistoryModal.load_unload_on_mlv_hlf = get_prod.load_unload_on_mlv_hlf
                    if request['emp_proofing_date'] != '':
                        FlightHistoryModal.emp_proofing_date = get_prod.emp_proofing_date
                    FlightHistoryModal.emp_proofing_remarks = get_prod.emp_proofing_remarks
                    if request['func_tst_date'] != '':
                        FlightHistoryModal.func_tst_date = get_prod.func_tst_date
                    FlightHistoryModal.func_tst_remarks = get_prod.func_tst_remarks
                    if request['func_tst_dummy_bird_date'] != '':
                        FlightHistoryModal.func_tst_dummy_bird_date = get_prod.func_tst_dummy_bird_date
                    FlightHistoryModal.func_tst_dummy_bird_remarks = get_prod.func_tst_dummy_bird_remarks
                    if request['road_test_date'] != '':
                        FlightHistoryModal.road_test_date = get_prod.road_test_date
                    FlightHistoryModal.road_test_remarks = get_prod.road_test_remarks
                    if request['post_road_test_date'] != '':
                        FlightHistoryModal.post_road_test_date = get_prod.post_road_test_date
                    FlightHistoryModal.post_road_test_remarks = get_prod.post_road_test_remarks
                    if request['integrated_operation_date'] != '':
                        FlightHistoryModal.integrated_operation_date = get_prod.integrated_operation_date
                    FlightHistoryModal.integrated_operation_remarks = get_prod.integrated_operation_remarks
                    if request['rain_test_date'] != '':
                        FlightHistoryModal.rain_test_date = get_prod.rain_test_date
                    FlightHistoryModal.rain_test_remarks = get_prod.rain_test_remarks
                    if request['pre_user_inspection_date'] != '':
                        FlightHistoryModal.pre_user_inspection_date = get_prod.pre_user_inspection_date
                    FlightHistoryModal.pre_user_inspection_remarks = get_prod.pre_user_inspection_remarks
                    if request['final_integrated_testing_date'] != '':
                        FlightHistoryModal.final_integrated_testing_date = get_prod.final_integrated_testing_date
                    FlightHistoryModal.final_integrated_testing_remarks = get_prod.final_integrated_testing_remarks
                    if request['load_unload_on_mlv_hlf_date'] != '':
                        FlightHistoryModal.load_unload_on_mlv_hlf_date = get_prod.load_unload_on_mlv_hlf_date
                    FlightHistoryModal.load_unload_on_mlv_hlf_remarks = get_prod.load_unload_on_mlv_hlf_remarks
                    if request['user_id'] != '':
                        FlightHistoryModal.user_id = get_prod.user_id
                    else:
                        FlightHistoryModal.user_id = None
                    FlightHistoryModal.save()

                    # history saved
                get_prod.system = request['system']
                get_prod.organization = request['organization']
                get_prod.set_id = request['set_id']
                if request['blt_date'] != '':
                    get_prod.blt_date = request['blt_date']
                get_prod.testing_date = request['testing_date']
                get_prod.sys_type = request['sys_type']
                get_prod.blt_status = request['blt_status']
                get_prod.blt_remarks = request['blt_remarks']
                if request['pre_hil_date'] != '':
                    get_prod.pre_hil_date = request['pre_hil_date']
                get_prod.pre_hil_status = request['pre_hil_status']
                get_prod.pre_hil_remarks = request['pre_hil_remarks']
                if request['vibaration_date'] != '':
                    get_prod.vibaration_date = request['vibaration_date']
                get_prod.vibaration_status = request['vibaration_status']
                get_prod.vibaration_remarks = request['vibaration_remarks']
                if request['post_hil_date'] != '':
                    get_prod.post_hil_date = request['post_hil_date']
                get_prod.post_hil_status = request['post_hil_status']
                get_prod.post_hil_remarks = request['post_hil_remarks']
                if request['fgt_date'] != '':
                    get_prod.fgt_date = request['fgt_date']
                get_prod.fgt_status = request['fgt_status']
                get_prod.fgt_remarks = request['fgt_remarks']
                if request['final_integration_date'] != '':
                    get_prod.final_integration_date = request['final_integration_date']
                get_prod.final_integration_status = request['final_integration_status']
                get_prod.final_integration_remarks = request['final_integration_remarks']
                if request['bhd_date'] != '':
                    get_prod.bhd_date = request['bhd_date']
                get_prod.bhd_status = request['bhd_status']
                get_prod.bhd_remarks = request['bhd_remarks']
                if request['fqm_date'] != '':
                    get_prod.fqm_date = request['fqm_date']
                get_prod.fqm_status = request['fqm_status']
                get_prod.fqm_remarks = request['fqm_remarks']
                if request['qm_certification_date'] != '':
                    get_prod.qm_certification_date = request['qm_certification_date']
                get_prod.qm_certification_status = request['qm_certification_status']
                get_prod.qm_certification_remarks = request['qm_certification_remarks']
                get_prod.attachment = request['attachment']
                get_prod.remarks = request['remarks']
                if request['cgbalancing_date'] != '':
                    get_prod.cgbalancing_date = request['cgbalancing_date']
                get_prod.cgbalancing_date_status = request['cgbalancing_date_status']
                get_prod.cgbalancing_date_remarks = request['cgbalancing_date_remarks']
                if request['launchact_date'] != '':
                    get_prod.launchact_date = request['launchact_date']
                get_prod.launchact_status = request['launchact_status']
                get_prod.launchact_remarks = request['launchact_remarks']
                if request['incapsulation_date'] != '':
                    get_prod.incapsulation_date = request['incapsulation_date']
                get_prod.incapsulation_status = request['incapsulation_status']
                get_prod.incapsulation_remarks = request['incapsulation_remarks']
                if request['sys_align_Date'] != '':
                    get_prod.sys_align_Date = request['sys_align_Date']
                get_prod.sys_align_status = request['sys_align_status']
                get_prod.sys_align_remarks = request['sys_align_remarks']
                get_prod.testing_type = request['testing_type']
                get_prod.emp_proofing = request['emp_proofing']
                get_prod.func_tst = request['func_tst']
                get_prod.func_tst_dummy_bird = request['func_tst_dummy_bird']
                get_prod.road_test = request['road_test']
                get_prod.post_road_test = request['post_road_test']
                get_prod.integrated_operation = request['integrated_operation']
                get_prod.rain_test = request['rain_test']
                get_prod.pre_user_inspection = request['pre_user_inspection']
                get_prod.final_integrated_testing = request['final_integrated_testing']
                get_prod.load_unload_on_mlv_hlf = request['load_unload_on_mlv_hlf']
                if request['emp_proofing_date'] != '':
                    get_prod.emp_proofing_date = request['emp_proofing_date']
                get_prod.emp_proofing_remarks = request['emp_proofing_remarks']
                if request['func_tst_date'] != '':
                    get_prod.func_tst_date = request['func_tst_date']
                get_prod.func_tst_remarks = request['func_tst_remarks']
                if request['func_tst_dummy_bird_date'] != '':
                    get_prod.func_tst_dummy_bird_date = request['func_tst_dummy_bird_date']
                get_prod.func_tst_dummy_bird_remarks = request['func_tst_dummy_bird_remarks']
                if request['road_test_date'] != '':
                    get_prod.road_test_date = request['road_test_date']
                get_prod.road_test_remarks = request['road_test_remarks']
                if request['post_road_test_date'] != '':
                    get_prod.post_road_test_date = request['post_road_test_date']
                get_prod.post_road_test_remarks = request['post_road_test_remarks']
                if request['integrated_operation_date'] != '':
                    get_prod.integrated_operation_date = request['integrated_operation_date']
                get_prod.integrated_operation_remarks = request['integrated_operation_remarks']
                if request['rain_test_date'] != '':
                    get_prod.rain_test_date = request['rain_test_date']
                get_prod.rain_test_remarks = request['rain_test_remarks']
                if request['pre_user_inspection_date'] != '':
                    get_prod.pre_user_inspection_date = request['pre_user_inspection_date']
                get_prod.pre_user_inspection_remarks = request['pre_user_inspection_remarks']
                if request['final_integrated_testing_date'] != '':
                    get_prod.final_integrated_testing_date = request['final_integrated_testing_date']
                get_prod.final_integrated_testing_remarks = request['final_integrated_testing_remarks']
                if request['load_unload_on_mlv_hlf_date'] != '':
                    get_prod.load_unload_on_mlv_hlf_date = request['load_unload_on_mlv_hlf_date']
                get_prod.load_unload_on_mlv_hlf_remarks = request['load_unload_on_mlv_hlf_remarks']
                if request['user_id'] != '':
                    get_prod.user_id = request['user_id']
                else:
                    get_prod.user_id = None
                get_prod.isActive = is_active

                get_prod.save()
        return JsonResponse({'status': 'True', 'message': "Production Status Updated Successfully!"},
                            status=200)

    # except Exception as e:
    #     return JsonResponse({'status': 'False', "message": "Status Not Saved"}, status=500)

    @staticmethod
    def GetFlightDateNone(request):
        dataList = FlightSystemStatusHistory.objects.all()
        ListItems = []
        for data in dataList:
            if data.blt_status == 'None':
                data.blt_date = None

            if data.emp_proofing == 'None':
                data.emp_proofing_date = None

            if data.func_tst == 'None':
                data.func_tst_date = None

            if data.func_tst_dummy_bird == 'None':
                data.func_tst_dummy_bird_date = None

            if data.road_test == 'None':
                data.road_test_date = None

            if data.post_road_test == 'None':
                data.post_road_test_date = None

            if data.integrated_operation == 'None':
                data.integrated_operation_date = None

            if data.rain_test == 'None':
                data.rain_test_date = None

            if data.pre_user_inspection == 'None':
                data.pre_user_inspection_date = None

            if data.final_integration_status == 'None':
                data.final_integration_date = None

            if data.load_unload_on_mlv_hlf == 'None':
                data.load_unload_on_mlv_hlf_date = None

            if data.pre_hil_status == 'None':
                data.pre_hil_date = None

            if data.vibaration_status == 'None':
                data.vibaration_date = None

            if data.cgbalancing_date_status == 'None':
                data.cgbalancing_date = None

            if data.cgbalancing_date_status == 'None':
                data.cgbalancing_date = None

            if data.post_hil_status == 'None':
                data.post_hil_date = None

            if data.sys_align_status == 'None':
                data.sys_align_Date = None

            if data.incapsulation_status == 'None':
                data.incapsulation_date = None

            if data.final_integration_status == 'None':
                data.final_integration_date = None

            if data.final_integration_status == 'None':
                data.final_integration_date = None

            if data.fgt_status == 'None':
                data.fgt_date =None

            if data.bhd_status == 'None':
                data.bhd_date = None

            if data.fqm_status == 'None':
                data.fqm_date = None

            if data.qm_certification_status == 'None':
                data.qm_certification_date = None

            if data.launchact_status == 'None':
                data.launchact_date = None
            data.save()
        my_list = FlightSystemStatusHistory.objects.all()
        serializer = FlightSystemSerialzer(my_list, many=True)
        return JsonResponse({'message': 'Welcome to Home Page', 'data': serializer.data}, status=200)
    @staticmethod
    def GetFlightList(request):
        try:
            org = request.query_params['selected_org']
            year = request.query_params['selected_year']
            type = request.query_params['selected_type']
            system = request.query_params['selected_system']
            ParentStatus = request.query_params['selected_status']
            ChildStatus = request.query_params['child_status']
            # print(ParentStatus)
            # print(ChildStatus)
            filter_objects = Q()

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

            # create dynamic filter
            # if year != '':
            #     filter_objects &= get_filter(
            #         'testing_date__year', 'equal',
            #         year)
            if org != '':
                filter_objects &= get_filter(
                    'organization', 'equal',
                    org)
            if type != '':
                filter_objects &= get_filter(
                    'sys_type', 'equal',
                    type)
            if system != '':
                filter_objects &= get_filter(
                    'system', 'equal',
                    system)

            setid_filter = Q()
            setid_filter &= get_filter(
                    'set_id', 'not_equal',
                    '')
            if org != '':
                setid_filter &= get_filter(
                    'organization', 'equal',
                    org)
            if type != '':
                setid_filter &= get_filter(
                    'sys_type', 'equal',
                    type)
            if system != '':
                setid_filter &= get_filter(
                    'system', 'equal',
                    system)
            dataList = FlightSystemStatus.objects.filter(filter_objects)
            if ParentStatus != '':
                ListItems = []
                totCountProduction = FlightSystemStatus.objects.filter(setid_filter)
                for set in totCountProduction:
                    count_id = 0
                    if set.blt_status != 'None' and set.set_id != '' and set.blt_date is not None:
                        if count_id == 0 and set.blt_date.strftime("%Y") == year:
                            count_id = set.id
                            ListItems.append(set)
                    if set.pre_hil_status != 'None' and set.set_id != '' and set.pre_hil_date is not None:
                        if count_id == 0 and set.pre_hil_date.strftime("%Y") == year:
                            count_id = set.id
                            ListItems.append(set)
                    if set.vibaration_status != 'None' and set.set_id != '' and set.vibaration_date is not None:
                        if count_id == 0 and set.vibaration_date.strftime("%Y") == year:
                            count_id = set.id
                            ListItems.append(set)
                    if set.post_hil_status != 'None' and set.set_id != '' and set.post_hil_date is not None:
                        if count_id == 0 and set.post_hil_date.strftime("%Y") == year:
                            count_id = set.id
                            ListItems.append(set)
                    if set.fgt_status != 'None' and set.set_id != '' and set.fgt_date is not None:
                        if count_id == 0 and set.fgt_date.strftime("%Y") == year:
                            count_id = set.id
                            ListItems.append(set)

                    if set.final_integration_status != 'None' and set.set_id != '' and set.final_integration_date is not None:
                        if count_id == 0 and set.final_integration_date.strftime("%Y") == year:
                            count_id = set.id
                            ListItems.append(set)
                    if set.bhd_status != 'None' and set.set_id != '' and set.bhd_date is not None:
                        if count_id == 0 and set.bhd_date.strftime("%Y") == year:
                            count_id = set.id
                            ListItems.append(set)
                    if set.fqm_status != 'None' and set.set_id != '' and set.fqm_date is not None:
                        if count_id == 0 and set.fqm_date.strftime("%Y") == year:
                            count_id = set.id
                            ListItems.append(set)
                    if set.qm_certification_status != 'None' and set.set_id != '' and set.qm_certification_date is not None:
                        if count_id == 0 and set.qm_certification_date.strftime("%Y") == year:
                            count_id = set.id
                            ListItems.append(set)
                    if set.cgbalancing_date_status != 'None' and set.set_id != '' and set.cgbalancing_date is not None:
                        if count_id == 0 and set.cgbalancing_date.strftime("%Y") == year:
                            count_id = set.id
                            ListItems.append(set)

                    if set.launchact_status != 'None' and set.set_id != '' and set.launchact_date is not None:
                        if count_id == 0 and set.launchact_date.strftime("%Y") == year:
                            count_id = set.id
                            ListItems.append(set)
                    if set.incapsulation_status != 'None' and set.set_id != '' and set.incapsulation_date is not None:
                        if count_id == 0 and set.incapsulation_date.strftime("%Y") == year:
                            count_id = set.id
                            ListItems.append(set)
                    if set.sys_align_status != 'None' and set.set_id != '' and set.sys_align_Date is not None:
                        if count_id == 0 and set.sys_align_Date.strftime("%Y") == year:
                            count_id = set.id
                            ListItems.append(set)
                    if set.emp_proofing != 'None' and set.set_id != '' and set.emp_proofing_date is not None:
                        if count_id == 0 and set.emp_proofing_date.strftime("%Y") == year:
                            count_id = set.id
                            ListItems.append(set)
                    if set.func_tst != 'None' and set.set_id != '' and set.func_tst_date is not None:
                        if count_id == 0 and set.func_tst_date.strftime("%Y") == year:
                            count_id = set.id
                            ListItems.append(set)

                    if set.func_tst_dummy_bird != 'None' and set.set_id != '' and set.func_tst_dummy_bird_date is not None:
                        if count_id == 0 and set.func_tst_dummy_bird_date.strftime("%Y") == year:
                            count_id = set.id
                            ListItems.append(set)
                    if set.road_test != 'None' and set.set_id != '' and set.road_test_date is not None:
                        if count_id == 0 and set.road_test_date.strftime("%Y") == year:
                            count_id = set.id
                            ListItems.append(set)
                    if set.post_road_test != 'None' and set.set_id != '' and set.post_road_test_date is not None:
                        if count_id == 0 and set.post_road_test_date.strftime("%Y") == year:
                            count_id = set.id
                            ListItems.append(set)
                    if set.integrated_operation != 'None' and set.set_id != '' and set.integrated_operation_date is not None:
                        if count_id == 0 and set.integrated_operation_date.strftime("%Y") == year:
                            count_id = set.id
                            ListItems.append(set)
                    if set.rain_test != 'None' and set.set_id != '' and set.rain_test_date is not None:
                        if count_id == 0 and set.rain_test_date.strftime("%Y") == year:
                            count_id = set.id
                            ListItems.append(set)

                    if set.pre_user_inspection != 'None' and set.set_id != '' and set.pre_user_inspection_date is not None:
                        if count_id == 0 and set.pre_user_inspection_date.strftime("%Y") == year:
                            count_id = set.id
                            ListItems.append(set)
                    if set.final_integrated_testing != 'None' and set.set_id != '' and set.final_integrated_testing_date is not None:
                        if count_id == 0 and set.final_integrated_testing_date.strftime("%Y") == year:
                            count_id = set.id
                            ListItems.append(set)
                    if set.load_unload_on_mlv_hlf != 'None' and set.set_id != '' and set.load_unload_on_mlv_hlf_date is not None:
                        if count_id == 0 and set.load_unload_on_mlv_hlf_date.strftime("%Y") == year:
                            count_id = set.id
                            ListItems.append(set)

                if ParentStatus == 'BLT':
                    dataList = dataList.extra(
                        select={
                            'year': 'extract (year from blt_date)',
                            'month': 'extract (month from blt_date)',
                            'day': 'extract (day from blt_date)'},
                        order_by=['month', 'day', '-year']
                    )
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(blt_date__year=year, blt_status=ChildStatus)
                    if ChildStatus == 'Current Count':
                        for data in dataList:
                            if data.blt_date.strftime("%Y") == year and (
                                    data.blt_status == 'Under process' or data.blt_status == 'Observation(same stage)' or data.blt_status == 'Halt'):
                                ListItems.append(data)

                if ParentStatus == 'EMP Proofing':
                    dataList = dataList.extra(
                        select={'year': 'extract (year from emp_proofing_date)',
                                'month': 'extract (month from emp_proofing_date)',
                                'day': 'extract (day from emp_proofing_date)'},
                        order_by=['month', 'day', '-year'])
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(emp_proofing_date__year=year, emp_proofing=ChildStatus)

                    if ChildStatus == 'Current Count':
                        for data in dataList:
                            if data.emp_proofing_date.strftime("%Y") == year and (
                                    data.emp_proofing == 'Under process' or data.emp_proofing == 'Observation(same stage)' or data.emp_proofing == 'Halt'):
                                ListItems.append(data)

                if ParentStatus == 'Functional Test W/O Dummy Bird':
                    dataList = dataList.extra(
                        select={'year': 'extract (year from func_tst_date)',
                                'month': 'extract (month from func_tst_date)',
                                'day': 'extract (day from func_tst_date)'},
                        order_by=['month', 'day', '-year'])
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(func_tst_date__year=year, func_tst=ChildStatus)

                    if ChildStatus == 'Current Count':
                        for data in dataList:
                            if data.func_tst_date.strftime("%Y") == year and (
                                    data.func_tst == 'Under process' or data.func_tst == 'Observation(same stage)' or data.func_tst == 'Halt'):
                                ListItems.append(data)

                if ParentStatus == 'Functional Test With Dummy Bird':
                    dataList = dataList.extra(
                        select={'year': 'extract (year from func_tst_dummy_bird_date)',
                                'month': 'extract (month from func_tst_dummy_bird_date)',
                                'day': 'extract (day from func_tst_dummy_bird_date)'},
                        order_by=['month', 'day', '-year'])
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(func_tst_dummy_bird_date__year=year,
                                                    func_tst_dummy_bird=ChildStatus)

                    if ChildStatus == 'Current Count':
                        for data in dataList:
                            if data.func_tst_dummy_bird_date.strftime("%Y") == year and (
                                    data.func_tst_dummy_bird == 'Under process' or data.func_tst_dummy_bird == 'Observation(same stage)' or data.func_tst_dummy_bird == 'Halt'):
                                ListItems.append(data)

                if ParentStatus == 'Road Test':
                    dataList = dataList.extra(
                        select={'year': 'extract (year from road_test_date)',
                                'month': 'extract (month from road_test_date)',
                                'day': 'extract (day from road_test_date)'},
                        order_by=['month', 'day', '-year'])
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(road_test_date__year=year, road_test=ChildStatus)

                    if ChildStatus == 'Current Count':
                        for data in dataList:
                            if data.road_test_date.strftime("%Y") == year and (
                                    data.road_test == 'Under process' or data.road_test == 'Observation(same stage)' or data.road_test == 'Halt'):
                                ListItems.append(data)

                if ParentStatus == 'Post Road Test':
                    dataList = dataList.extra(
                        select={'year': 'extract (year from post_road_test_date)',
                                'month': 'extract (month from post_road_test_date)',
                                'day': 'extract (day from post_road_test_date)'},
                        order_by=['month', 'day', '-year'])
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(post_road_test_date__year=year, post_road_test=ChildStatus)

                    if ChildStatus == 'Current Count':
                        for data in dataList:
                            if data.post_road_test_date.strftime("%Y") == year and (
                                    data.post_road_test == 'Under process' or data.post_road_test == 'Observation(same stage)' or data.post_road_test == 'Halt'):
                                ListItems.append(data)

                if ParentStatus == 'Integrated Operation':
                    dataList = dataList.extra(
                        select={'year': 'extract (year from integrated_operation_date)',
                                'month': 'extract (month from integrated_operation_date)',
                                'day': 'extract (day from integrated_operation_date)'},
                        order_by=['month', 'day', '-year'])
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(integrated_operation_date__year=year,
                                                    integrated_operation=ChildStatus)

                    if ChildStatus == 'Current Count':
                        for data in dataList:
                            if data.integrated_operation_date.strftime("%Y") == year and (
                                    data.integrated_operation == 'Under process' or data.integrated_operation == 'Observation(same stage)' or data.integrated_operation == 'Halt'):
                                ListItems.append(data)

                if ParentStatus == 'Rain Test':
                    dataList = dataList.extra(
                        select={'year': 'extract (year from rain_test_date)',
                                'month': 'extract (month from rain_test_date)',
                                'day': 'extract (day from rain_test_date)'},
                        order_by=['month', 'day', '-year'])
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(rain_test_date__year=year, rain_test=ChildStatus)

                    if ChildStatus == 'Current Count':
                        for data in dataList:
                            if data.rain_test_date.strftime("%Y") == year and (
                                    data.rain_test == 'Under process' or data.rain_test == 'Observation(same stage)' or data.rain_test == 'Halt'):
                                ListItems.append(data)

                if ParentStatus == 'Pre User Inspection':
                    dataList = dataList.extra(
                        select={'year': 'extract (year from pre_user_inspection_date)',
                                'month': 'extract (month from pre_user_inspection_date)',
                                'day': 'extract (day from pre_user_inspection_date)'},
                        order_by=['month', 'day', '-year'])
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(pre_user_inspection_date__year=year,
                                                    pre_user_inspection=ChildStatus)

                    if ChildStatus == 'Current Count':
                        for data in dataList:
                            if data.pre_user_inspection_date.strftime("%Y") == year and (
                                    data.pre_user_inspection == 'Under process' or data.pre_user_inspection == 'Observation(same stage)' or data.pre_user_inspection == 'Halt'):
                                ListItems.append(data)

                if ParentStatus == 'Final Integration':
                    dataList = dataList.extra(
                        select={'year': 'extract (year from final_integration_date)',
                                'month': 'extract (month from final_integration_date)',
                                'day': 'extract (day from final_integration_date)'},
                        order_by=['month', 'day', '-year'])
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(final_integration_date__year=year,
                                                    final_integration_status=ChildStatus)

                    if ChildStatus == 'Current Count':
                        for data in dataList:
                            if data.final_integration_date.strftime("%Y") == year and (
                                    data.final_integration_status == 'Under process' or data.final_integration_status == 'Observation(same stage)' or data.final_integration_status == 'Halt'):
                                ListItems.append(data)

                if ParentStatus == 'Loading/Unloading on MLV/HLF':
                    dataList = dataList.extra(
                        select={'year': 'extract (year from load_unload_on_mlv_hlf_date)',
                                'month': 'extract (month from load_unload_on_mlv_hlf_date)',
                                'day': 'extract (day from load_unload_on_mlv_hlf_date)'},
                        order_by=['month', 'day', '-year'])
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(load_unload_on_mlv_hlf_date__year=year,
                                                    load_unload_on_mlv_hlf=ChildStatus)

                    if ChildStatus == 'Current Count':
                        for data in dataList:
                            if data.load_unload_on_mlv_hlf_date.strftime("%Y") == year and (
                                    data.load_unload_on_mlv_hlf == 'Under process' or data.load_unload_on_mlv_hlf == 'Observation(same stage)' or data.load_unload_on_mlv_hlf == 'Halt'):
                                ListItems.append(data)

                if ParentStatus == 'Pre-HIL':
                    dataList = dataList.extra(
                        select={'year': 'extract (year from pre_hil_date)',
                                'month': 'extract (month from pre_hil_date)',
                                'day': 'extract (day from pre_hil_date)'},
                        order_by=['month', 'day', '-year'])
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(pre_hil_date__year=year, pre_hil_status=ChildStatus)

                    if ChildStatus == 'Current Count':
                        for data in dataList:
                            if data.pre_hil_date.strftime("%Y") == year and (
                                    data.pre_hil_status == 'Under process' or data.pre_hil_status == 'Observation(same stage)' or data.pre_hil_status == 'Halt'):
                                ListItems.append(data)

                if ParentStatus == 'Vibration':
                    dataList = dataList.extra(
                        select={'year': 'extract (year from vibaration_date)',
                                'month': 'extract (month from vibaration_date)',
                                'day': 'extract (day from vibaration_date)'},
                        order_by=['month', 'day', '-year'])
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(vibaration_date__year=year, vibaration_status=ChildStatus)

                    if ChildStatus == 'Current Count':
                        for data in dataList:
                            if data.vibaration_date.strftime("%Y") == year and (
                                    data.vibaration_status == 'Under process' or data.vibaration_status == 'Observation(same stage)' or data.vibaration_status == 'Halt'):
                                ListItems.append(data)

                if ParentStatus == 'CG Balancing':
                    dataList = dataList.extra(
                        select={'year': 'extract (year from cgbalancing_date)',
                                'month': 'extract (month from cgbalancing_date)',
                                'day': 'extract (day from cgbalancing_date)'},
                        order_by=['month', 'day', '-year'])
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(cgbalancing_date__year=year, cgbalancing_date_status=ChildStatus)

                    if ChildStatus == 'Current Count':
                        for data in dataList:
                            if data.cgbalancing_date.strftime("%Y") == year and (
                                    data.cgbalancing_date_status == 'Under process' or data.cgbalancing_date_status == 'Observation(same stage)' or data.cgbalancing_date_status == 'Halt'):
                                ListItems.append(data)

                if ParentStatus == 'Post-HIL':
                    dataList = dataList.extra(
                        select={'year': 'extract (year from post_hil_date)',
                                'month': 'extract (month from post_hil_date)',
                                'day': 'extract (day from post_hil_date)'},
                        order_by=['month', 'day', '-year'])
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(post_hil_date__year=year, post_hil_status=ChildStatus)

                    if ChildStatus == 'Current Count':
                        for data in dataList:
                            if data.post_hil_date.strftime("%Y") == year and (
                                    data.post_hil_status == 'Under process' or data.post_hil_status == 'Observation(same stage)' or data.post_hil_status == 'Halt'):
                                ListItems.append(data)

                if ParentStatus == 'System Alignment':
                    # dataList = dataList.extra(
                    #     select={'year': 'extract (year from sys_align_Date)',
                    #             'month': 'extract (month from sys_align_Date)',
                    #             'day': 'extract (day from sys_align_Date)'},
                    #     order_by=['month', 'day', '-year'])
                    dataList = dataList.annotate(sys_align_Date__month=Extract('sys_align_Date', 'month'),
                        sys_align_Date__year=Extract('sys_align_Date', 'year'),
                        sys_align_Date__day=Extract('sys_align_Date', 'day')).order_by('sys_align_Date__month','sys_align_Date__day','-sys_align_Date__year')
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(sys_align_Date__year=year, sys_align_status=ChildStatus)

                    if ChildStatus == 'Current Count':
                        for data in dataList:
                            if data.sys_align_Date.strftime("%Y") == year and (
                                    data.sys_align_status == 'Under process' or data.sys_align_status == 'Observation(same stage)' or data.sys_align_status == 'Halt'):
                                ListItems.append(data)

                if ParentStatus == 'Incapsulation':
                    dataList = dataList.extra(
                        select={'year': 'extract (year from incapsulation_date)',
                                'month': 'extract (month from incapsulation_date)',
                                'day': 'extract (day from incapsulation_date)'},
                        order_by=['month', 'day', '-year'])
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(incapsulation_date__year=year, incapsulation_status=ChildStatus)

                    if ChildStatus == 'Current Count':
                        for data in dataList:
                            if data.incapsulation_date.strftime("%Y") == year and (
                                    data.incapsulation_status == 'Under process' or data.incapsulation_status == 'Observation(same stage)' or data.incapsulation_status == 'Halt'):
                                ListItems.append(data)

                if ParentStatus == 'Final Integrated Testing':
                    dataList = dataList.extra(
                        select={'year': 'extract (year from final_integration_date)',
                                'month': 'extract (month from final_integration_date)',
                                'day': 'extract (day from final_integration_date)'},
                        order_by=['month', 'day', '-year'])
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(final_integration_date__year=year,
                                                    final_integration_status=ChildStatus)

                    if ChildStatus == 'Current Count':
                        for data in dataList:
                            if data.final_integration_date.strftime("%Y") == year and (
                                    data.final_integration_status == 'Under process' or data.final_integration_status == 'Observation(same stage)' or data.final_integration_status == 'Halt'):
                                ListItems.append(data)

                if ParentStatus == 'FGT Status':
                    dataList = dataList.extra(
                        select={'year': 'extract (year from fgt_date)',
                                'month': 'extract (month from fgt_date)',
                                'day': 'extract (day from fgt_date)'},
                        order_by=['month', 'day', '-year'])
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(fgt_date__year=year, fgt_status=ChildStatus)

                    if ChildStatus == 'Current Count':
                        for data in dataList:
                            if data.fgt_date.strftime("%Y") == year and (
                                    data.fgt_status == 'Under process' or data.fgt_status == 'Observation(same stage)' or data.fgt_status == 'Halt'):
                                ListItems.append(data)

                if ParentStatus == 'BHD Status':
                    dataList = dataList.extra(
                        select={'year': 'extract (year from bhd_date)',
                                'month': 'extract (month from bhd_date)',
                                'day': 'extract (day from bhd_date)'},
                        order_by=['month', 'day', '-year'])
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(bhd_date__year=year, bhd_status=ChildStatus)

                    if ChildStatus == 'Current Count':
                        for data in dataList:
                            if data.bhd_date.strftime("%Y") == year and (
                                    data.bhd_status == 'Ok' or data.bhd_status == 'Not Submitted)' or data.bhd_status == 'QM Observations Forwarded'):
                                ListItems.append(data)

                if ParentStatus == 'FQM Status':
                    dataList = dataList.extra(
                        select={'year': 'extract (year from fqm_date)',
                                'month': 'extract (month from fqm_date)',
                                'day': 'extract (day from fqm_date)'},
                        order_by=['month', 'day', '-year'])
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(fqm_date__year=year, fqm_status=ChildStatus)

                    if ChildStatus == 'Current Count':
                        for data in dataList:
                            if data.fqm_date.strftime("%Y") == year and (data.fqm_status == 'Not Conducted)'):
                                ListItems.append(data)

                if ParentStatus == 'QM Certification Status':
                    dataList = dataList.extra(
                        select={'year': 'extract (year from qm_certification_date)',
                                'month': 'extract (month from qm_certification_date)',
                                'day': 'extract (day from qm_certification_date)'},
                        order_by=['month', 'day', '-year'])
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(qm_certification_date__year=year,
                                                    qm_certification_status=ChildStatus)

                    if ChildStatus == 'Current Count':
                        for data in dataList:
                            if data.qm_certification_date.strftime("%Y") == year and (
                                    data.qm_certification_status == 'QM certificate issued' or data.qm_certification_status == 'QM Observations Forwarded)'):
                                ListItems.append(data)

                if ParentStatus == 'Launch/ End User':
                    dataList = dataList.extra(
                        select={'year': 'extract (year from launchact_date)',
                                'month': 'extract (month from launchact_date)',
                                'day': 'extract (day from launchact_date)'},
                        order_by=['month', 'day', '-year'])
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(launchact_date__year = year, launchact_status=ChildStatus)
                    else:
                        for data in dataList:
                            if data.launchact_date.strftime("%Y") == year and (data.launchact_status == 'Ok' or data.launchact_status == 'Observation' or data.launchact_status == 'Halt'):
                                ListItems.append(data)

                serializer = FlightSystemSerialzer(ListItems, many=True)
                print(serializer.data)
                return JsonResponse({'message': 'Welcome to Home Page', 'data': serializer.data}, status=200)


        # serializer = FlightSystemSerialzer(dataList, many=True)
        # return JsonResponse({'message': 'Welcome to Home Page', 'data': serializer.data}, status=200)

        except Exception as e:
            print(e)
            return JsonResponse({'message': 'Sorry! No Task found.'}, status=500)

    # API for PDF generator with table and Text_wraping

    @staticmethod
    def GetFlightListHistory(request):
        try:
            id = request.query_params['id']
            ParentStatus = request.query_params['parent_status']
            dataList = FlightSystemStatusHistory.objects.filter(f_id=id)
            if ParentStatus == 'BLT':
                dataList = dataList.extra(
                    select={
                        'year': 'extract (year from blt_date)',
                        'month': 'extract (month from blt_date)',
                        'day': 'extract (day from blt_date)'},
                    order_by=['month', 'day', '-year'])


            if ParentStatus == 'EMP Proofing':
                dataList = dataList.extra(
                    select={'year': 'extract (year from emp_proofing_date)',
                            'month': 'extract (month from emp_proofing_date)',
                            'day': 'extract (day from emp_proofing_date)'},
                    order_by=['month', 'day', '-year'])

            if ParentStatus == 'Functional Test W/O Dummy Bird':
                dataList = dataList.extra(
                    select={'year': 'extract (year from func_tst_date)',
                            'month': 'extract (month from func_tst_date)',
                            'day': 'extract (day from func_tst_date)'},
                    order_by=['month', 'day', '-year'])

            if ParentStatus == 'Functional Test With Dummy Bird':
                dataList = dataList.extra(
                    select={'year': 'extract (year from func_tst_dummy_bird_date)',
                            'month': 'extract (month from func_tst_dummy_bird_date)',
                            'day': 'extract (day from func_tst_dummy_bird_date)'},
                    order_by=['month', 'day', '-year'])

            if ParentStatus == 'Road Test':
                dataList = dataList.extra(
                    select={'year': 'extract (year from road_test_date)',
                            'month': 'extract (month from road_test_date)',
                            'day': 'extract (day from road_test_date)'},
                    order_by=['month', 'day', '-year'])

            if ParentStatus == 'Post Road Test':
                dataList = dataList.extra(
                    select={'year': 'extract (year from post_road_test_date)',
                            'month': 'extract (month from post_road_test_date)',
                            'day': 'extract (day from post_road_test_date)'},
                    order_by=['month', 'day', '-year'])

            if ParentStatus == 'Integrated Operation':
                dataList = dataList.extra(
                    select={'year': 'extract (year from integrated_operation_date)',
                            'month': 'extract (month from integrated_operation_date)',
                            'day': 'extract (day from integrated_operation_date)'},
                    order_by=['month', 'day', '-year'])

            if ParentStatus == 'Rain Test':
                dataList = dataList.extra(
                    select={'year': 'extract (year from rain_test_date)',
                            'month': 'extract (month from rain_test_date)',
                            'day': 'extract (day from rain_test_date)'},
                    order_by=['month', 'day', '-year'])

            if ParentStatus == 'Pre User Inspection':
                dataList = dataList.extra(
                    select={'year': 'extract (year from pre_user_inspection_date)',
                            'month': 'extract (month from pre_user_inspection_date)',
                            'day': 'extract (day from pre_user_inspection_date)'},
                    order_by=['month', 'day', '-year'])

            if ParentStatus == 'Final Integration':
                dataList = dataList.extra(
                    select={'year': 'extract (year from final_integration_date)',
                            'month': 'extract (month from final_integration_date)',
                            'day': 'extract (day from final_integration_date)'},
                    order_by=['month', 'day', '-year'])

            if ParentStatus == 'Loading/Unloading on MLV/HLF':
                dataList = dataList.extra(
                    select={'year': 'extract (year from load_unload_on_mlv_hlf_date)',
                            'month': 'extract (month from load_unload_on_mlv_hlf_date)',
                            'day': 'extract (day from load_unload_on_mlv_hlf_date)'},
                    order_by=['month', 'day', '-year'])

            if ParentStatus == 'Pre-HIL':
                dataList = dataList.extra(
                    select={'year': 'extract (year from pre_hil_date)',
                            'month': 'extract (month from pre_hil_date)',
                            'day': 'extract (day from pre_hil_date)'},
                    order_by=['month', 'day', '-year'])

            if ParentStatus == 'Vibration':
                dataList = dataList.extra(
                    select={'year': 'extract (year from vibaration_date)',
                            'month': 'extract (month from vibaration_date)',
                            'day': 'extract (day from vibaration_date)'},
                    order_by=['month', 'day', '-year'])

            if ParentStatus == 'CG Balancing':
                dataList = dataList.extra(
                    select={'year': 'extract (year from cgbalancing_date)',
                            'month': 'extract (month from cgbalancing_date)',
                            'day': 'extract (day from cgbalancing_date)'},
                    order_by=['month', 'day', '-year'])

            if ParentStatus == 'Post-HIL':
                dataList = dataList.extra(
                    select={'year': 'extract (year from post_hil_date)',
                            'month': 'extract (month from post_hil_date)',
                            'day': 'extract (day from post_hil_date)'},
                    order_by=['month', 'day', '-year'])

            if ParentStatus == 'System Alignment':
                dataList = dataList.extra(
                    select={'year': 'extract (year from sys_align_Date)',
                            'month': 'extract (month from sys_align_Date)',
                            'day': 'extract (day from sys_align_Date)'},
                    order_by=['month', 'day', '-year'])


            if ParentStatus == 'Incapsulation':
                dataList = dataList.extra(
                    select={'year': 'extract (year from incapsulation_date)',
                            'month': 'extract (month from incapsulation_date)',
                            'day': 'extract (day from incapsulation_date)'},
                    order_by=['month', 'day', '-year'])


            if ParentStatus == 'Final Integrated Testing':
                dataList = dataList.extra(
                    select={'year': 'extract (year from final_integration_date)',
                            'month': 'extract (month from final_integration_date)',
                            'day': 'extract (day from final_integration_date)'},
                    order_by=['month', 'day', '-year'])


            if ParentStatus == 'FGT Status':
                dataList = dataList.extra(
                    select={'year': 'extract (year from fgt_date)',
                            'month': 'extract (month from fgt_date)',
                            'day': 'extract (day from fgt_date)'},
                    order_by=['month', 'day', '-year'])

            if ParentStatus == 'BHD Status':
                dataList = dataList.extra(
                    select={'year': 'extract (year from bhd_date)',
                            'month': 'extract (month from bhd_date)',
                            'day': 'extract (day from bhd_date)'},
                    order_by=['month', 'day', '-year'])

            if ParentStatus == 'FQM Status':
                dataList = dataList.extra(
                    select={'year': 'extract (year from fqm_date)',
                            'month': 'extract (month from fqm_date)',
                            'day': 'extract (day from fqm_date)'},
                    order_by=['month', 'day', '-year'])


            if ParentStatus == 'QM Certification Status':
                dataList = dataList.extra(
                    select={'year': 'extract (year from qm_certification_date)',
                            'month': 'extract (month from qm_certification_date)',
                            'day': 'extract (day from qm_certification_date)'},
                    order_by=['month', 'day', '-year'])

            if ParentStatus == 'Launch/ End User':
                dataList = dataList.extra(
                    select={'year': 'extract (year from launchact_date)',
                            'month': 'extract (month from launchact_date)',
                            'day': 'extract (day from launchact_date)'},
                    order_by=['month', 'day', '-year'])
            serializer = FlightSystemSerialzer(dataList, many=True)
            return JsonResponse({'status': 'true', 'data': serializer.data}, status=200)

        except Exception as e:
            print(e)
            return JsonResponse({'status': 'false'}, status=200)

    @staticmethod
    def GetFlightPDFList(request):
        TABLE_COL_NAMES = (
            "System", "Organization", "Set ID", "BLT Status", "Pre-Hil Status", "Vibration Status", "Post-Hil Status",
            "FGT Status", "Final Integration Status", "BHD Status", "FQM Status", "QM Certification Status",
            "Remarks")
        data = []
        status = request.query_params['status']
        if status == 'Total':
            data = FlightSystemStatus.objects.all().order_by('-id').values_list('system', 'organization', 'set_id',
                                                                                'blt_status', 'pre_hil_status',
                                                                                'vibaration_status', 'post_hil_status',
                                                                                'fgt_status',
                                                                                'final_integration_status',
                                                                                'bhd_status', 'fqm_status',
                                                                                'qm_certification_status',
                                                                                'remarks')
        else:
            data = FlightSystemStatus.objects.filter(qm_certification_status=status).order_by('-id').values_list(
                'system', 'organization', 'set_id',
                'blt_status', 'pre_hil_status',
                'vibaration_status', 'post_hil_status',
                'fgt_status',
                'final_integration_status',
                'bhd_status', 'fqm_status',
                'qm_certification_status',
                'remarks')

        pdf = FPDF('L', 'mm', 'Legal')
        pdf.add_page()
        pdf.set_font('courier', 'B', 26)
        pdf.cell(330, 10, 'Flight System Report', border=0, align='C', ln=2)
        pdf.cell(40, 10, '', 0, 1)
        pdf.set_font("Times", size=10)
        line_height = pdf.font_size * 4.5
        col_width = pdf.epw / 14

        def render_table_header():
            pdf.set_font(style="B")
            for col_name in TABLE_COL_NAMES:
                pdf.multi_cell(col_width, line_height, col_name, border=1, align='C', ln=3,
                               max_line_height=pdf.font_size)
            pdf.ln(line_height)
            pdf.set_font(style="")

        render_table_header()

        lh_list = []
        use_default_height = 0
        for row in data:
            for datum in row:
                dd = str(datum)
                word_list = dd.split()
                number_of_words = len(word_list)
                if number_of_words > 2:
                    use_default_height = 1
                    new_line_height = pdf.font_size * 2 * (number_of_words / 2)
            if not use_default_height:
                lh_list.append(line_height)
            else:
                lh_list.append(new_line_height)
                use_default_height = 0

        for j, row in enumerate(data):
            line_height = lh_list[j]
            if pdf.will_page_break(line_height):
                render_table_header()
            for col_num in range(len(row)):
                line_height = lh_list[j]
                pdf.multi_cell(col_width, line_height, f"{row[col_num]}", border=1, align='C', ln=3,
                               max_line_height=pdf.font_size)
            pdf.ln(line_height)
        pdf.output('FlightSystem.pdf')
        return FileResponse(open('FlightSystem.pdf', 'rb'), as_attachment=True, content_type='application/pdf')

    @staticmethod
    def GetFlightExcelList(request):
        response = HttpResponse(content_type='application/ms-excel')
        response['Content-Disposition'] = 'attachment; filename="DocumentExcelSheet.xls"'
        work_book = xlwt.Workbook(encoding='utf-8')
        work_sheet = work_book.add_sheet('ExcelSheet')
        row_num = 0
        font_style = xlwt.XFStyle()
        font_style.font.bold = True
        TABLE_COL_NAMES = (
            "System", "Organization", "Set ID", "BLT Status", "Pre-Hil Status", "Vibration Status", "Post-Hil Status",
            "FGT Status", "Final Integration Status", "BHD Status", "FQM Status", "QM Certification Status",
            "Remarks")

        for col_num in range(len(TABLE_COL_NAMES)):
            work_sheet.write(row_num, col_num, TABLE_COL_NAMES[col_num], font_style)
        font_style = xlwt.XFStyle()
        data = []
        status = request.query_params['status']
        if status == 'Total':
            data = FlightSystemStatus.objects.all().order_by('-id').values_list('system', 'organization', 'set_id',
                                                                                'blt_status', 'pre_hil_status',
                                                                                'vibaration_status', 'post_hil_status',
                                                                                'fgt_status',
                                                                                'final_integration_status',
                                                                                'bhd_status', 'fqm_status',
                                                                                'qm_certification_status',
                                                                                'remarks')
        else:
            data = FlightSystemStatus.objects.filter(qm_certification_status=status).order_by('-id').values_list(
                'system', 'organization', 'set_id',
                'blt_status', 'pre_hil_status',
                'vibaration_status', 'post_hil_status',
                'fgt_status',
                'final_integration_status',
                'bhd_status', 'fqm_status',
                'qm_certification_status',
                'remarks')
        for row in data:
            row_num += 1
            for col_num in range(len(row)):
                work_sheet.write(row_num, col_num, str(row[col_num]), font_style)
        work_book.save(response)
        return response

    @staticmethod
    def AddRelifingStatus(request):
        refilModel = RelifingSystemStatus()
        is_active = 0
        if request['is_active'] == 'true':
            is_active = 1
        try:
            id = request['id']
            if id == '0':
                refilModel.system = request['system']
                refilModel.organization = request['organization']
                refilModel.set_id = request['set_id']
                if request['blt_date'] != '':
                    refilModel.blt_date = request['blt_date']
                refilModel.testing_date = request['testing_date']
                refilModel.sys_type = request['sys_type']
                refilModel.blt_status = request['blt_status']
                refilModel.blt_remarks = request['blt_remarks']
                if request['pre_hil_date'] != '':
                    refilModel.pre_hil_date = request['pre_hil_date']
                refilModel.pre_hil_status = request['pre_hil_status']
                refilModel.pre_hil_remarks = request['pre_hil_remarks']
                if request['vibaration_date'] != '':
                    refilModel.vibaration_date = request['vibaration_date']
                refilModel.vibaration_status = request['vibaration_status']
                refilModel.vibaration_remarks = request['vibaration_remarks']
                if request['post_hil_date'] != '':
                    refilModel.post_hil_date = request['post_hil_date']
                refilModel.post_hil_status = request['post_hil_status']
                refilModel.post_hil_remarks = request['post_hil_remarks']
                if request['fgt_date'] != '':
                    refilModel.fgt_date = request['fgt_date']
                refilModel.fgt_status = request['fgt_status']
                refilModel.fgt_remarks = request['fgt_remarks']
                if request['final_integration_date'] != '':
                    refilModel.final_integration_date = request['final_integration_date']
                refilModel.final_integration_status = request['final_integration_status']
                refilModel.final_integration_remarks = request['final_integration_remarks']
                if request['bhd_date'] != '':
                    refilModel.bhd_date = request['bhd_date']
                refilModel.bhd_status = request['bhd_status']
                refilModel.bhd_remarks = request['bhd_remarks']
                if request['fqm_date'] != '':
                    refilModel.fqm_date = request['fqm_date']
                refilModel.fqm_status = request['fqm_status']
                refilModel.fqm_remarks = request['fqm_remarks']
                if request['qm_certification_date'] != '':
                    refilModel.qm_certification_date = request['qm_certification_date']
                refilModel.qm_certification_status = request['qm_certification_status']
                refilModel.qm_certification_remarks = request['qm_certification_remarks']
                refilModel.attachment = request['attachment']
                refilModel.remarks = request['remarks']
                if request['cgbalancing_date'] != '':
                    refilModel.cgbalancing_date = request['cgbalancing_date']
                refilModel.cgbalancing_date_status = request['cgbalancing_date_status']
                refilModel.cgbalancing_date_remarks = request['cgbalancing_date_remarks']
                if request['enduser_date'] != '':
                    refilModel.enduser_date = request['enduser_date']
                refilModel.enduser_status = request['enduser_status']
                refilModel.enduser_remarks = request['enduser_remarks']
                if request['incapsulation_date'] != '':
                    refilModel.incapsulation_date = request['incapsulation_date']
                refilModel.incapsulation_status = request['incapsulation_status']
                refilModel.incapsulation_remarks = request['incapsulation_remarks']
                if request['sys_align_Date'] != '':
                    refilModel.sys_align_Date = request['sys_align_Date']
                refilModel.sys_align_status = request['sys_align_status']
                refilModel.sys_align_remarks = request['sys_align_remarks']
                refilModel.testing_type = request['testing_type']
                refilModel.emp_proofing = request['emp_proofing']
                refilModel.func_tst = request['func_tst']
                refilModel.func_tst_dummy_bird = request['func_tst_dummy_bird']
                refilModel.road_test = request['road_test']
                refilModel.post_road_test = request['post_road_test']
                refilModel.integrated_operation = request['integrated_operation']
                refilModel.rain_test = request['rain_test']
                refilModel.pre_user_inspection = request['pre_user_inspection']
                refilModel.final_integrated_testing = request['final_integrated_testing']
                refilModel.load_unload_on_mlv_hlf = request['load_unload_on_mlv_hlf']
                if request['emp_proofing_date'] != '':
                    refilModel.emp_proofing_date = request['emp_proofing_date']
                refilModel.emp_proofing_remarks = request['emp_proofing_remarks']
                if request['func_tst_date'] != '':
                    refilModel.func_tst_date = request['func_tst_date']
                refilModel.func_tst_remarks = request['func_tst_remarks']
                if request['func_tst_dummy_bird_date'] != '':
                    refilModel.func_tst_dummy_bird_date = request['func_tst_dummy_bird_date']
                refilModel.func_tst_dummy_bird_remarks = request['func_tst_dummy_bird_remarks']
                if request['road_test_date'] != '':
                    refilModel.road_test_date = request['road_test_date']
                refilModel.road_test_remarks = request['road_test_remarks']
                if request['post_road_test_date'] != '':
                    refilModel.post_road_test_date = request['post_road_test_date']
                refilModel.post_road_test_remarks = request['post_road_test_remarks']
                if request['integrated_operation_date'] != '':
                    refilModel.integrated_operation_date = request['integrated_operation_date']
                refilModel.integrated_operation_remarks = request['integrated_operation_remarks']
                if request['rain_test_date'] != '':
                    refilModel.rain_test_date = request['rain_test_date']
                refilModel.rain_test_remarks = request['rain_test_remarks']
                if request['pre_user_inspection_date'] != '':
                    refilModel.pre_user_inspection_date = request['pre_user_inspection_date']
                refilModel.pre_user_inspection_remarks = request['pre_user_inspection_remarks']
                if request['final_integrated_testing_date'] != '':
                    refilModel.final_integrated_testing_date = request['final_integrated_testing_date']
                refilModel.final_integrated_testing_remarks = request['final_integrated_testing_remarks']
                if request['load_unload_on_mlv_hlf_date'] != '':
                    refilModel.load_unload_on_mlv_hlf_date = request['load_unload_on_mlv_hlf_date']
                refilModel.load_unload_on_mlv_hlf_remarks = request['load_unload_on_mlv_hlf_remarks']
                if request['user_id'] != '':
                    refilModel.user_id = request['user_id']
                else:
                    refilModel.user_id = None
                refilModel.isActive = is_active
                refilModel.save()
                return JsonResponse({'status': 'True', 'message': "Production Status Added Successfully!"},
                                    status=200)
            else:
                get_prod = RelifingSystemStatus.objects.filter(id=id).first()
                if get_prod is not None:
                    if request['blt_status'] != get_prod.blt_status \
                            or request['blt_date'] != get_prod.blt_date \
                            or request['blt_remarks'] != get_prod.blt_remarks \
                            or request['pre_hil_status'] != get_prod.pre_hil_status \
                            or request['pre_hil_date'] != get_prod.pre_hil_date \
                            or request['pre_hil_remarks'] != get_prod.pre_hil_remarks \
                            or request['vibaration_status'] != get_prod.vibaration_status \
                            or request['vibaration_date'] != get_prod.vibaration_date \
                            or request['vibaration_remarks'] != get_prod.vibaration_remarks \
                            or request['post_hil_status'] != get_prod.post_hil_status \
                            or request['post_hil_date'] != get_prod.post_hil_date \
                            or request['post_hil_remarks'] != get_prod.post_hil_remarks \
                            or request['fgt_status'] != get_prod.fgt_status \
                            or request['fgt_date'] != get_prod.fgt_date \
                            or request['fgt_remarks'] != get_prod.fgt_remarks \
                            or request['final_integration_status'] != get_prod.final_integration_status \
                            or request['final_integration_date'] != get_prod.final_integration_date \
                            or request['final_integration_remarks'] != get_prod.final_integration_remarks \
                            or request['bhd_status'] != get_prod.bhd_status \
                            or request['bhd_date'] != get_prod.bhd_date \
                            or request['bhd_remarks'] != get_prod.bhd_remarks \
                            or request['fqm_status'] != get_prod.fqm_status \
                            or request['fqm_date'] != get_prod.fqm_date \
                            or request['fqm_remarks'] != get_prod.fqm_remarks \
                            or request['qm_certification_status'] != get_prod.qm_certification_status \
                            or request['qm_certification_date'] != get_prod.qm_certification_date \
                            or request['qm_certification_remarks'] != get_prod.qm_certification_remarks \
                            or request['incapsulation_status'] != get_prod.incapsulation_status \
                            or request['incapsulation_date'] != get_prod.incapsulation_date \
                            or request['incapsulation_remarks'] != get_prod.incapsulation_remarks \
                            or request['sys_align_status'] != get_prod.sys_align_status \
                            or request['sys_align_Date'] != get_prod.sys_align_Date \
                            or request['sys_align_remarks'] != get_prod.sys_align_remarks \
                            or request['emp_proofing'] != get_prod.emp_proofing \
                            or request['emp_proofing_date'] != get_prod.emp_proofing_date \
                            or request['emp_proofing_remarks'] != get_prod.emp_proofing_remarks \
                            or request['func_tst'] != get_prod.func_tst \
                            or request['func_tst_date'] != get_prod.func_tst_date \
                            or request['func_tst_remarks'] != get_prod.func_tst_remarks \
                            or request['func_tst_dummy_bird'] != get_prod.func_tst_dummy_bird \
                            or request['func_tst_dummy_bird_date'] != get_prod.func_tst_dummy_bird_date \
                            or request['func_tst_dummy_bird_remarks'] != get_prod.func_tst_dummy_bird_remarks \
                            or request['road_test'] != get_prod.road_test \
                            or request['road_test_date'] != get_prod.road_test_date \
                            or request['road_test_remarks'] != get_prod.road_test_remarks \
                            or request['post_road_test'] != get_prod.post_road_test \
                            or request['post_road_test_date'] != get_prod.post_road_test_date \
                            or request['post_road_test_remarks'] != get_prod.post_road_test_remarks \
                            or request['integrated_operation'] != get_prod.integrated_operation \
                            or request['integrated_operation_date'] != get_prod.integrated_operation_date \
                            or request['integrated_operation_remarks'] != get_prod.integrated_operation_remarks \
                            or request['rain_test'] != get_prod.rain_test \
                            or request['rain_test_date'] != get_prod.rain_test_date \
                            or request['rain_test_remarks'] != get_prod.rain_test_remarks \
                            or request['pre_user_inspection'] != get_prod.pre_user_inspection \
                            or request['pre_user_inspection_date'] != get_prod.pre_user_inspection_date \
                            or request['pre_user_inspection_remarks'] != get_prod.pre_user_inspection_remarks \
                            or request['final_integrated_testing'] != get_prod.final_integrated_testing \
                            or request['final_integrated_testing_date'] != get_prod.final_integrated_testing_date \
                            or request['final_integrated_testing_remarks'] != get_prod.final_integrated_testing_remarks \
                            or request['load_unload_on_mlv_hlf'] != get_prod.load_unload_on_mlv_hlf \
                            or request['load_unload_on_mlv_hlf_date'] != get_prod.load_unload_on_mlv_hlf_date \
                            or request['load_unload_on_mlv_hlf_remarks'] != get_prod.load_unload_on_mlv_hlf_remarks \
                            or request['launchact_status'] != get_prod.launchact_status \
                            or request['enduser_status'] != get_prod.enduser_status \
                            or request['enduser_date'] != get_prod.enduser_date \
                            or request['enduser_remarks'] != get_prod.enduser_remarks:
                        print("add data in flight history")
                        RelifingHistoryModal = RelifingSystemStatusHistory()
                        RelifingHistoryModal.r_id = get_prod.id
                        RelifingHistoryModal.system = get_prod.system
                        RelifingHistoryModal.organization = get_prod.organization
                        RelifingHistoryModal.set_id = get_prod.set_id
                        if request['blt_date'] != '':
                            RelifingHistoryModal.blt_date = get_prod.blt_date
                        RelifingHistoryModal.blt_status = get_prod.blt_status
                        RelifingHistoryModal.blt_remarks = get_prod.blt_remarks
                        RelifingHistoryModal.testing_date = get_prod.testing_date
                        RelifingHistoryModal.sys_type = get_prod.sys_type
                        if request['pre_hil_date'] != '':
                            RelifingHistoryModal.pre_hil_date = get_prod.pre_hil_date
                        RelifingHistoryModal.pre_hil_status = get_prod.pre_hil_status
                        RelifingHistoryModal.pre_hil_remarks = get_prod.pre_hil_remarks
                        if request['vibaration_date'] != '':
                            RelifingHistoryModal.vibaration_date = get_prod.vibaration_date
                        RelifingHistoryModal.vibaration_status = get_prod.vibaration_status
                        RelifingHistoryModal.vibaration_remarks = get_prod.vibaration_remarks
                        if request['post_hil_date'] != '':
                            RelifingHistoryModal.post_hil_date = get_prod.post_hil_date
                        RelifingHistoryModal.post_hil_status = get_prod.post_hil_status
                        RelifingHistoryModal.post_hil_remarks = get_prod.post_hil_remarks
                        if request['fgt_date'] != '':
                            RelifingHistoryModal.fgt_date = get_prod.fgt_date
                        RelifingHistoryModal.fgt_status = get_prod.fgt_status
                        RelifingHistoryModal.fgt_remarks = get_prod.fgt_remarks
                        if request['final_integration_date'] != '':
                            RelifingHistoryModal.final_integration_date = get_prod.final_integration_date
                        RelifingHistoryModal.final_integration_status = get_prod.final_integration_status
                        RelifingHistoryModal.final_integration_remarks = get_prod.final_integration_remarks
                        if request['bhd_date'] != '':
                            RelifingHistoryModal.bhd_date = get_prod.bhd_date
                        RelifingHistoryModal.bhd_status = get_prod.bhd_status
                        RelifingHistoryModal.bhd_remarks = get_prod.bhd_remarks
                        if request['fqm_date'] != '':
                            RelifingHistoryModal.fqm_date = get_prod.fqm_date
                        RelifingHistoryModal.fqm_status = get_prod.fqm_status
                        RelifingHistoryModal.fqm_remarks = get_prod.fqm_remarks
                        if request['qm_certification_date'] != '':
                            RelifingHistoryModal.qm_certification_date = get_prod.qm_certification_date
                        RelifingHistoryModal.qm_certification_status = get_prod.qm_certification_status
                        RelifingHistoryModal.qm_certification_remarks = get_prod.qm_certification_remarks
                        RelifingHistoryModal.attachment = get_prod.attachment
                        RelifingHistoryModal.remarks = get_prod.remarks
                        RelifingHistoryModal.isActive = get_prod.isActive
                        if request['cgbalancing_date'] != '':
                            RelifingHistoryModal.cgbalancing_date = get_prod.cgbalancing_date
                        RelifingHistoryModal.cgbalancing_date_status = get_prod.cgbalancing_date_status
                        RelifingHistoryModal.cgbalancing_date_remarks = get_prod.cgbalancing_date_remarks
                        if request['enduser_date'] != '':
                            RelifingHistoryModal.enduser_date = get_prod.enduser_date
                        RelifingHistoryModal.enduser_status = get_prod.enduser_status
                        RelifingHistoryModal.enduser_remarks = get_prod.enduser_remarks
                        if request['incapsulation_date'] != '':
                            RelifingHistoryModal.incapsulation_date = get_prod.incapsulation_date
                        RelifingHistoryModal.incapsulation_status = get_prod.incapsulation_status
                        RelifingHistoryModal.incapsulation_remarks = get_prod.incapsulation_remarks
                        if request['sys_align_Date'] != '':
                            RelifingHistoryModal.sys_align_Date = get_prod.sys_align_Date
                        RelifingHistoryModal.sys_align_status = get_prod.sys_align_status
                        RelifingHistoryModal.sys_align_remarks = get_prod.sys_align_remarks
                        RelifingHistoryModal.testing_type = get_prod.testing_type
                        RelifingHistoryModal.emp_proofing = get_prod.emp_proofing
                        RelifingHistoryModal.func_tst = get_prod.func_tst
                        RelifingHistoryModal.func_tst_dummy_bird = get_prod.func_tst_dummy_bird
                        RelifingHistoryModal.road_test = get_prod.road_test
                        RelifingHistoryModal.post_road_test = get_prod.post_road_test
                        RelifingHistoryModal.integrated_operation = get_prod.integrated_operation
                        RelifingHistoryModal.rain_test = get_prod.rain_test
                        RelifingHistoryModal.pre_user_inspection = get_prod.pre_user_inspection
                        RelifingHistoryModal.final_integrated_testing = get_prod.final_integrated_testing
                        RelifingHistoryModal.load_unload_on_mlv_hlf = get_prod.load_unload_on_mlv_hlf
                        if request['emp_proofing_date'] != '':
                            RelifingHistoryModal.emp_proofing_date = get_prod.emp_proofing_date
                        RelifingHistoryModal.emp_proofing_remarks = get_prod.emp_proofing_remarks
                        if request['func_tst_date'] != '':
                            RelifingHistoryModal.func_tst_date = get_prod.func_tst_date
                        RelifingHistoryModal.func_tst_remarks = get_prod.func_tst_remarks
                        if request['func_tst_dummy_bird_date'] != '':
                            RelifingHistoryModal.func_tst_dummy_bird_date = get_prod.func_tst_dummy_bird_date
                        RelifingHistoryModal.func_tst_dummy_bird_remarks = get_prod.func_tst_dummy_bird_remarks
                        if request['road_test_date'] != '':
                            RelifingHistoryModal.road_test_date = get_prod.road_test_date
                        RelifingHistoryModal.road_test_remarks = get_prod.road_test_remarks
                        if request['post_road_test_date'] != '':
                            RelifingHistoryModal.post_road_test_date = get_prod.post_road_test_date
                        RelifingHistoryModal.post_road_test_remarks = get_prod.post_road_test_remarks
                        if request['integrated_operation_date'] != '':
                            RelifingHistoryModal.integrated_operation_date = get_prod.integrated_operation_date
                        RelifingHistoryModal.integrated_operation_remarks = get_prod.integrated_operation_remarks
                        if request['rain_test_date'] != '':
                            RelifingHistoryModal.rain_test_date = get_prod.rain_test_date
                        RelifingHistoryModal.rain_test_remarks = get_prod.rain_test_remarks
                        if request['pre_user_inspection_date'] != '':
                            RelifingHistoryModal.pre_user_inspection_date = get_prod.pre_user_inspection_date
                        RelifingHistoryModal.pre_user_inspection_remarks = get_prod.pre_user_inspection_remarks
                        if request['final_integrated_testing_date'] != '':
                            RelifingHistoryModal.final_integrated_testing_date = get_prod.final_integrated_testing_date
                        RelifingHistoryModal.final_integrated_testing_remarks = get_prod.final_integrated_testing_remarks
                        if request['load_unload_on_mlv_hlf_date'] != '':
                            RelifingHistoryModal.load_unload_on_mlv_hlf_date = get_prod.load_unload_on_mlv_hlf_date
                        RelifingHistoryModal.load_unload_on_mlv_hlf_remarks = get_prod.load_unload_on_mlv_hlf_remarks
                        if request['user_id'] != '':
                            RelifingHistoryModal.user_id = get_prod.user_id
                        else:
                            RelifingHistoryModal.user_id = None
                        RelifingHistoryModal.save()

                        # history saved
                    get_prod.system = request['system']
                    get_prod.organization = request['organization']
                    get_prod.set_id = request['set_id']
                    if request['blt_date'] != '':
                        get_prod.blt_date = request['blt_date']
                    get_prod.testing_date = request['testing_date']
                    get_prod.sys_type = request['sys_type']
                    get_prod.blt_status = request['blt_status']
                    get_prod.blt_remarks = request['blt_remarks']
                    if request['pre_hil_date'] != '':
                        get_prod.pre_hil_date = request['pre_hil_date']
                    get_prod.pre_hil_status = request['pre_hil_status']
                    get_prod.pre_hil_remarks = request['pre_hil_remarks']
                    if request['vibaration_date'] != '':
                        get_prod.vibaration_date = request['vibaration_date']
                    get_prod.vibaration_status = request['vibaration_status']
                    get_prod.vibaration_remarks = request['vibaration_remarks']
                    if request['post_hil_date'] != '':
                        get_prod.post_hil_date = request['post_hil_date']
                    get_prod.post_hil_status = request['post_hil_status']
                    get_prod.post_hil_remarks = request['post_hil_remarks']
                    if request['fgt_date'] != '':
                        get_prod.fgt_date = request['fgt_date']
                    get_prod.fgt_status = request['fgt_status']
                    get_prod.fgt_remarks = request['fgt_remarks']
                    if request['final_integration_date'] != '':
                        get_prod.final_integration_date = request['final_integration_date']
                    get_prod.final_integration_status = request['final_integration_status']
                    get_prod.final_integration_remarks = request['final_integration_remarks']
                    if request['bhd_date'] != '':
                        get_prod.bhd_date = request['bhd_date']
                    get_prod.bhd_status = request['bhd_status']
                    get_prod.bhd_remarks = request['bhd_remarks']
                    if request['fqm_date'] != '':
                        get_prod.fqm_date = request['fqm_date']
                    get_prod.fqm_status = request['fqm_status']
                    get_prod.fqm_remarks = request['fqm_remarks']
                    if request['qm_certification_date'] != '':
                        get_prod.qm_certification_date = request['qm_certification_date']
                    get_prod.qm_certification_status = request['qm_certification_status']
                    get_prod.qm_certification_remarks = request['qm_certification_remarks']
                    get_prod.attachment = request['attachment']
                    get_prod.remarks = request['remarks']
                    if request['cgbalancing_date'] != '':
                        get_prod.cgbalancing_date = request['cgbalancing_date']
                    get_prod.cgbalancing_date_status = request['cgbalancing_date_status']
                    get_prod.cgbalancing_date_remarks = request['cgbalancing_date_remarks']
                    if request['enduser_date'] != '':
                        get_prod.enduser_date = request['enduser_date']
                    get_prod.enduser_status = request['enduser_status']
                    get_prod.enduser_remarks = request['enduser_remarks']
                    if request['incapsulation_date'] != '':
                        get_prod.incapsulation_date = request['incapsulation_date']
                    get_prod.incapsulation_status = request['incapsulation_status']
                    get_prod.incapsulation_remarks = request['incapsulation_remarks']
                    if request['sys_align_Date'] != '':
                        get_prod.sys_align_Date = request['sys_align_Date']
                    get_prod.sys_align_status = request['sys_align_status']
                    get_prod.sys_align_remarks = request['sys_align_remarks']
                    get_prod.testing_type = request['testing_type']
                    get_prod.emp_proofing = request['emp_proofing']
                    get_prod.func_tst = request['func_tst']
                    get_prod.func_tst_dummy_bird = request['func_tst_dummy_bird']
                    get_prod.road_test = request['road_test']
                    get_prod.post_road_test = request['post_road_test']
                    get_prod.integrated_operation = request['integrated_operation']
                    get_prod.rain_test = request['rain_test']
                    get_prod.pre_user_inspection = request['pre_user_inspection']
                    get_prod.final_integrated_testing = request['final_integrated_testing']
                    get_prod.load_unload_on_mlv_hlf = request['load_unload_on_mlv_hlf']
                    if request['emp_proofing_date'] != '':
                        get_prod.emp_proofing_date = request['emp_proofing_date']
                    get_prod.emp_proofing_remarks = request['emp_proofing_remarks']
                    if request['func_tst_date'] != '':
                        get_prod.func_tst_date = request['func_tst_date']
                    get_prod.func_tst_remarks = request['func_tst_remarks']
                    if request['func_tst_dummy_bird_date'] != '':
                        get_prod.func_tst_dummy_bird_date = request['func_tst_dummy_bird_date']
                    get_prod.func_tst_dummy_bird_remarks = request['func_tst_dummy_bird_remarks']
                    if request['road_test_date'] != '':
                        get_prod.road_test_date = request['road_test_date']
                    get_prod.road_test_remarks = request['road_test_remarks']
                    if request['post_road_test_date'] != '':
                        get_prod.post_road_test_date = request['post_road_test_date']
                    get_prod.post_road_test_remarks = request['post_road_test_remarks']
                    if request['integrated_operation_date'] != '':
                        get_prod.integrated_operation_date = request['integrated_operation_date']
                    get_prod.integrated_operation_remarks = request['integrated_operation_remarks']
                    if request['rain_test_date'] != '':
                        get_prod.rain_test_date = request['rain_test_date']
                    get_prod.rain_test_remarks = request['rain_test_remarks']
                    if request['pre_user_inspection_date'] != '':
                        get_prod.pre_user_inspection_date = request['pre_user_inspection_date']
                    get_prod.pre_user_inspection_remarks = request['pre_user_inspection_remarks']
                    if request['final_integrated_testing_date'] != '':
                        get_prod.final_integrated_testing_date = request['final_integrated_testing_date']
                    get_prod.final_integrated_testing_remarks = request['final_integrated_testing_remarks']
                    if request['load_unload_on_mlv_hlf_date'] != '':
                        get_prod.load_unload_on_mlv_hlf_date = request['load_unload_on_mlv_hlf_date']
                    get_prod.load_unload_on_mlv_hlf_remarks = request['load_unload_on_mlv_hlf_remarks']
                    if request['user_id'] != '':
                        get_prod.user_id = request['user_id']
                    else:
                        get_prod.user_id = None
                    get_prod.isActive = is_active
                    get_prod.save()
            return JsonResponse({'status': 'True', 'message': "Production Status Updated Successfully!"},
                                status=200)
        except Exception as e:
            return JsonResponse({'status': 'False', "message": "Status Not Saved"}, status=500)


    @staticmethod
    def GetRelifingDateNone(request):
        dataList = RelifingSystemStatusHistory.objects.all()
        ListItems = []
        for data in dataList:
            if data.blt_status == 'None':
                data.blt_date = None

            if data.emp_proofing == 'None':
                data.emp_proofing_date = None

            if data.func_tst == 'None':
                data.func_tst_date = None

            if data.func_tst_dummy_bird == 'None':
                data.func_tst_dummy_bird_date = None

            if data.road_test == 'None':
                data.oad_test_date = None

            if data.post_road_test == 'None':
                data.post_road_test_date = None

            if data.integrated_operation == 'None':
                data.integrated_operation_date = None

            if data.rain_test == 'None':
                data.rain_test_date = None

            if data.pre_user_inspection == 'None':
                data.pre_user_inspection_date = None

            if data.final_integration_status == 'None':
                data.final_integration_date = None

            if data.load_unload_on_mlv_hlf == 'None':
                data.load_unload_on_mlv_hlf_date = None

            if data.pre_hil_status == 'None':
                data.pre_hil_date = None

            if data.vibaration_status == 'None':
                data.vibaration_date = None

            if data.cgbalancing_date_status == 'None':
                data.cgbalancing_date = None

            if data.cgbalancing_date_status == 'None':
                data.cgbalancing_date = None

            if data.post_hil_status == 'None':
                data.post_hil_date = None

            if data.sys_align_status == 'None':
                data.sys_align_Date = None

            if data.incapsulation_status == 'None':
                data.incapsulation_date = None

            if data.final_integration_status == 'None':
                data.final_integration_date = None

            if data.final_integration_status == 'None':
                data.final_integration_date = None

            if data.fgt_status == 'None':
                data.fgt_date =None

            if data.bhd_status == 'None':
                data.bhd_date = None


            if data.fqm_status == 'None':
                data.fqm_date = None

            if data.qm_certification_status == 'None':
                data.qm_certification_date = None

            if data.enduser_status == 'None':
                data.enduser_date = None
            data.save()
        my_list = RelifingSystemStatusHistory.objects.all()
        serializer = RelifingSystemSerialzer(my_list, many=True)
        return JsonResponse({'message': 'Welcome to Home Page', 'data': serializer.data}, status=200)
    @staticmethod
    def GetRelifingList(request):
        # try:
            org = request.query_params['selected_org']
            year = request.query_params['selected_year']
            type = request.query_params['selected_type']
            system = request.query_params['selected_system']
            ParentStatus = request.query_params['selected_status']
            ChildStatus = request.query_params['child_status']
            # print(ParentStatus)
            # print(ChildStatus)
            # setupItems = status.split(':')
            # ParentStatus = setupItems[0]
            # ChildStatus = setupItems[1]
            filter_objects = Q()

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

            # create dynamic filter
            # if year != '':
            #     filter_objects &= get_filter(
            #         'testing_date__year', 'equal',
            #         year)
            if org != '':
                filter_objects &= get_filter(
                    'organization', 'equal',
                    org)
            if type != '':
                filter_objects &= get_filter(
                    'sys_type', 'equal',
                    type)
            if system != '':
                filter_objects &= get_filter(
                    'system', 'equal',
                    system)

            setid_filter = Q()
            setid_filter &= get_filter(
                    'set_id', 'not_equal',
                    '')
            if org != '':
                setid_filter &= get_filter(
                    'organization', 'equal',
                    org)
            if type != '':
                setid_filter &= get_filter(
                    'sys_type', 'equal',
                    type)
            if system != '':
                setid_filter &= get_filter(
                    'system', 'equal',
                    system)
            dataList = RelifingSystemStatus.objects.filter(filter_objects)

            if ParentStatus != '':
                ListItems = []
                if ChildStatus == 'Total System':
                    # listVlaue = RelifingSystemStatus.objects.filter(setid_filter).values()
                    totCountRelif = RelifingSystemStatus.objects.filter(setid_filter)
                    relif_sys_count = 0
                    for set in totCountRelif:
                        count_id = 0
                        if set.blt_status != 'None' and set.set_id != '' and set.blt_date is not None:
                            if count_id == 0 and set.blt_date.strftime("%Y") == year:
                                count_id = set.id
                                ListItems.append(set)
                        if set.pre_hil_status != 'None' and set.set_id != '' and set.pre_hil_date is not None:
                            if count_id == 0 and set.pre_hil_date.strftime("%Y") == year:
                                count_id = set.id
                                ListItems.append(set)
                        if set.vibaration_status != 'None' and set.set_id != '' and set.vibaration_date is not None:
                            if count_id == 0 and set.vibaration_date.strftime("%Y") == year:
                                count_id = set.id
                                ListItems.append(set)
                        if set.post_hil_status != 'None' and set.set_id != '' and set.post_hil_date is not None:
                            if count_id == 0 and set.post_hil_date.strftime("%Y") == year:
                                count_id = set.id
                                ListItems.append(set)
                        if set.fgt_status != 'None' and set.set_id != '' and set.fgt_date is not None:
                            if count_id == 0 and set.fgt_date.strftime("%Y") == year:
                                count_id = set.id
                                ListItems.append(set)

                        if set.final_integration_status != 'None' and set.set_id != '' and set.final_integration_date is not None:
                            if count_id == 0 and set.final_integration_date.strftime("%Y") == year:
                                count_id = set.id
                                ListItems.append(set)
                        if set.bhd_status != 'None' and set.set_id != '' and set.bhd_date is not None:
                            if count_id == 0 and set.bhd_date.strftime("%Y") == year:
                                count_id = set.id
                                ListItems.append(set)
                        if set.fqm_status != 'None' and set.set_id != '' and set.fqm_date is not None:
                            if count_id == 0 and set.fqm_date.strftime("%Y") == year:
                                count_id = set.id
                                ListItems.append(set)
                        if set.qm_certification_status != 'None' and set.set_id != '' and set.qm_certification_date is not None:
                            if count_id == 0 and set.qm_certification_date.strftime("%Y") == year:
                                count_id = set.id
                                ListItems.append(set)
                        if set.cgbalancing_date_status != 'None' and set.set_id != '' and set.cgbalancing_date is not None:
                            if count_id == 0 and set.cgbalancing_date.strftime("%Y") == year:
                                count_id = set.id
                                ListItems.append(set)

                        if set.enduser_status != 'None' and set.set_id != '' and set.enduser_date is not None:
                            if count_id == 0 and set.enduser_date.strftime("%Y") == year:
                                count_id = set.id
                                ListItems.append(set)
                        if set.incapsulation_status != 'None' and set.set_id != '' and set.incapsulation_date is not None:
                            if count_id == 0 and set.incapsulation_date.strftime("%Y") == year:
                                count_id = set.id
                                ListItems.append(set)
                        if set.sys_align_status != 'None' and set.set_id != '' and set.sys_align_Date is not None:
                            if count_id == 0 and set.sys_align_Date.strftime("%Y") == year:
                                count_id = set.id
                                ListItems.append(set)
                        if set.emp_proofing != 'None' and set.set_id != '' and set.emp_proofing_date is not None:
                            if count_id == 0 and set.emp_proofing_date.strftime("%Y") == year:
                                count_id = set.id
                                ListItems.append(set)
                        if set.func_tst != 'None' and set.set_id != '' and set.func_tst_date is not None:
                            if count_id == 0 and set.func_tst_date.strftime("%Y") == year:
                                count_id = set.id
                                ListItems.append(set)

                        if set.func_tst_dummy_bird != 'None' and set.set_id != '' and set.func_tst_dummy_bird_date is not None:
                            if count_id == 0 and set.func_tst_dummy_bird_date.strftime("%Y") == year:
                                count_id = set.id
                                ListItems.append(set)
                        if set.road_test != 'None' and set.set_id != '' and set.road_test_date is not None:
                            if count_id == 0 and set.road_test_date.strftime("%Y") == year:
                                count_id = set.id
                                ListItems.append(set)
                        if set.post_road_test != 'None' and set.set_id != '' and set.post_road_test_date is not None:
                            if count_id == 0 and set.post_road_test_date.strftime("%Y") == year:
                                count_id = set.id
                                ListItems.append(set)
                        if set.integrated_operation != 'None' and set.set_id != '' and set.integrated_operation_date is not None:
                            if count_id == 0 and set.integrated_operation_date.strftime("%Y") == year:
                                count_id = set.id
                                ListItems.append(set)
                        if set.rain_test != 'None' and set.set_id != '' and set.rain_test_date is not None:
                            if count_id == 0 and set.rain_test_date.strftime("%Y") == year:
                                count_id = set.id
                                ListItems.append(set)

                        if set.pre_user_inspection != 'None' and set.set_id != '' and set.pre_user_inspection_date is not None:
                            if count_id == 0 and set.pre_user_inspection_date.strftime("%Y") == year:
                                count_id = set.id
                                ListItems.append(set)
                        if set.final_integrated_testing != 'None' and set.set_id != '' and set.final_integrated_testing_date is not None:
                            if count_id == 0 and set.final_integrated_testing_date.strftime("%Y") == year:
                                count_id = set.id
                                ListItems.append(set)
                        if set.load_unload_on_mlv_hlf != 'None' and set.set_id != '' and set.load_unload_on_mlv_hlf_date is not None:
                            if count_id == 0 and set.load_unload_on_mlv_hlf_date.strftime("%Y") == year:
                                count_id = set.id
                                ListItems.append(set)

                if ParentStatus == 'BLT':
                    dataList = dataList.extra(
                        select={
                            'year': 'extract (year from blt_date)',
                            'month': 'extract (month from blt_date)',
                            'day': 'extract (day from blt_date)'},
                        order_by=['month', 'day', '-year']
                    )
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(blt_date__year=year, blt_status=ChildStatus)
                    if ChildStatus == 'Current Count':
                        for data in dataList:
                            if data.blt_date.strftime("%Y") == year and (
                                    data.blt_status == 'Under process' or data.blt_status == 'Observation(same stage)' or data.blt_status == 'Halt'):
                                ListItems.append(data)

                if ParentStatus == 'EMP Proofing':
                    dataList = dataList.extra(
                        select={'year': 'extract (year from emp_proofing_date)',
                                'month': 'extract (month from emp_proofing_date)',
                                'day': 'extract (day from emp_proofing_date)'},
                        order_by=['month', 'day', '-year'])
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(emp_proofing_date__year=year, emp_proofing=ChildStatus)

                    if ChildStatus == 'Current Count':
                        for data in dataList:
                            if data.emp_proofing_date.strftime("%Y") == year and (
                                    data.emp_proofing == 'Under process' or data.emp_proofing == 'Observation(same stage)' or data.emp_proofing == 'Halt'):
                                ListItems.append(data)

                if ParentStatus == 'Functional Test W/O Dummy Bird':
                    dataList = dataList.extra(
                        select={'year': 'extract (year from func_tst_date)',
                                'month': 'extract (month from func_tst_date)',
                                'day': 'extract (day from func_tst_date)'},
                        order_by=['month', 'day', '-year'])
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(func_tst_date__year=year, func_tst=ChildStatus)

                    if ChildStatus == 'Current Count':
                        for data in dataList:
                            if data.func_tst_date.strftime("%Y") == year and (
                                    data.func_tst == 'Under process' or data.func_tst == 'Observation(same stage)' or data.func_tst == 'Halt'):
                                ListItems.append(data)

                if ParentStatus == 'Functional Test With Dummy Bird':
                    dataList = dataList.extra(
                        select={'year': 'extract (year from func_tst_dummy_bird_date)',
                                'month': 'extract (month from func_tst_dummy_bird_date)',
                                'day': 'extract (day from func_tst_dummy_bird_date)'},
                        order_by=['month', 'day', '-year'])
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(func_tst_dummy_bird_date__year=year,
                                                    func_tst_dummy_bird=ChildStatus)

                    if ChildStatus == 'Current Count':
                        for data in dataList:
                            if data.func_tst_dummy_bird_date.strftime("%Y") == year and (
                                    data.func_tst_dummy_bird == 'Under process' or data.func_tst_dummy_bird == 'Observation(same stage)' or data.func_tst_dummy_bird == 'Halt'):
                                ListItems.append(data)

                if ParentStatus == 'Road Test':
                    dataList = dataList.extra(
                        select={'year': 'extract (year from road_test_date)',
                                'month': 'extract (month from road_test_date)',
                                'day': 'extract (day from road_test_date)'},
                        order_by=['month', 'day', '-year'])
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(road_test_date__year=year, road_test=ChildStatus)

                    if ChildStatus == 'Current Count':
                        for data in dataList:
                            if data.road_test_date.strftime("%Y") == year and (
                                    data.road_test == 'Under process' or data.road_test == 'Observation(same stage)' or data.road_test == 'Halt'):
                                ListItems.append(data)

                if ParentStatus == 'Post Road Test':
                    dataList = dataList.extra(
                        select={'year': 'extract (year from post_road_test_date)',
                                'month': 'extract (month from post_road_test_date)',
                                'day': 'extract (day from post_road_test_date)'},
                        order_by=['month', 'day', '-year'])
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(post_road_test_date__year=year, post_road_test=ChildStatus)

                    if ChildStatus == 'Current Count':
                        for data in dataList:
                            if data.post_road_test_date.strftime("%Y") == year and (
                                    data.post_road_test == 'Under process' or data.post_road_test == 'Observation(same stage)' or data.post_road_test == 'Halt'):
                                ListItems.append(data)

                if ParentStatus == 'Integrated Operation':
                    dataList = dataList.extra(
                        select={'year': 'extract (year from integrated_operation_date)',
                                'month': 'extract (month from integrated_operation_date)',
                                'day': 'extract (day from integrated_operation_date)'},
                        order_by=['month', 'day', '-year'])
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(integrated_operation_date__year=year,
                                                    integrated_operation=ChildStatus)

                    if ChildStatus == 'Current Count':
                        for data in dataList:
                            if data.integrated_operation_date.strftime("%Y") == year and (
                                    data.integrated_operation == 'Under process' or data.integrated_operation == 'Observation(same stage)' or data.integrated_operation == 'Halt'):
                                ListItems.append(data)

                if ParentStatus == 'Rain Test':
                    dataList = dataList.extra(
                        select={'year': 'extract (year from rain_test_date)',
                                'month': 'extract (month from rain_test_date)',
                                'day': 'extract (day from rain_test_date)'},
                        order_by=['month', 'day', '-year'])
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(rain_test_date__year=year, rain_test=ChildStatus)

                    if ChildStatus == 'Current Count':
                        for data in dataList:
                            if data.rain_test_date.strftime("%Y") == year and (
                                    data.rain_test == 'Under process' or data.rain_test == 'Observation(same stage)' or data.rain_test == 'Halt'):
                                ListItems.append(data)

                if ParentStatus == 'Pre User Inspection':
                    dataList = dataList.extra(
                        select={'year': 'extract (year from pre_user_inspection_date)',
                                'month': 'extract (month from pre_user_inspection_date)',
                                'day': 'extract (day from pre_user_inspection_date)'},
                        order_by=['month', 'day', '-year'])
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(pre_user_inspection_date__year=year,
                                                    pre_user_inspection=ChildStatus)

                    if ChildStatus == 'Current Count':
                        for data in dataList:
                            if data.pre_user_inspection_date.strftime("%Y") == year and (
                                    data.pre_user_inspection == 'Under process' or data.pre_user_inspection == 'Observation(same stage)' or data.pre_user_inspection == 'Halt'):
                                ListItems.append(data)

                if ParentStatus == 'Final Integration':
                    dataList = dataList.extra(
                        select={'year': 'extract (year from final_integration_date)',
                                'month': 'extract (month from final_integration_date)',
                                'day': 'extract (day from final_integration_date)'},
                        order_by=['month', 'day', '-year'])
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(final_integration_date__year=year,
                                                    final_integration_status=ChildStatus)

                    if ChildStatus == 'Current Count':
                        for data in dataList:
                            if data.final_integration_date.strftime("%Y") == year and (
                                    data.final_integration_status == 'Under process' or data.final_integration_status == 'Observation(same stage)' or data.final_integration_status == 'Halt'):
                                ListItems.append(data)

                if ParentStatus == 'Loading/Unloading on MLV/HLF':
                    dataList = dataList.extra(
                        select={'year': 'extract (year from load_unload_on_mlv_hlf_date)',
                                'month': 'extract (month from load_unload_on_mlv_hlf_date)',
                                'day': 'extract (day from load_unload_on_mlv_hlf_date)'},
                        order_by=['month', 'day', '-year'])
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(load_unload_on_mlv_hlf_date__year=year,
                                                    load_unload_on_mlv_hlf=ChildStatus)

                    if ChildStatus == 'Current Count':
                        for data in dataList:
                            if data.load_unload_on_mlv_hlf_date.strftime("%Y") == year and (
                                    data.load_unload_on_mlv_hlf == 'Under process' or data.load_unload_on_mlv_hlf == 'Observation(same stage)' or data.load_unload_on_mlv_hlf == 'Halt'):
                                ListItems.append(data)

                if ParentStatus == 'Pre-HIL':
                    dataList = dataList.extra(
                        select={'year': 'extract (year from pre_hil_date)',
                                'month': 'extract (month from pre_hil_date)',
                                'day': 'extract (day from pre_hil_date)'},
                        order_by=['month', 'day', '-year'])
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(pre_hil_date__year=year, pre_hil_status=ChildStatus)

                    if ChildStatus == 'Current Count':
                        for data in dataList:
                            if data.pre_hil_date.strftime("%Y") == year and (
                                    data.pre_hil_status == 'Under process' or data.pre_hil_status == 'Observation(same stage)' or data.pre_hil_status == 'Halt'):
                                ListItems.append(data)

                if ParentStatus == 'Vibration':
                    dataList = dataList.extra(
                        select={'year': 'extract (year from vibaration_date)',
                                'month': 'extract (month from vibaration_date)',
                                'day': 'extract (day from vibaration_date)'},
                        order_by=['month', 'day', '-year'])
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(vibaration_date__year=year, vibaration_status=ChildStatus)

                    if ChildStatus == 'Current Count':
                        for data in dataList:
                            if data.vibaration_date.strftime("%Y") == year and (
                                    data.vibaration_status == 'Under process' or data.vibaration_status == 'Observation(same stage)' or data.vibaration_status == 'Halt'):
                                ListItems.append(data)

                if ParentStatus == 'CG Balancing':
                    dataList = dataList.extra(
                        select={'year': 'extract (year from cgbalancing_date)',
                                'month': 'extract (month from cgbalancing_date)',
                                'day': 'extract (day from cgbalancing_date)'},
                        order_by=['month', 'day', '-year'])
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(cgbalancing_date__year=year, cgbalancing_date_status=ChildStatus)

                    if ChildStatus == 'Current Count':
                        for data in dataList:
                            if data.cgbalancing_date.strftime("%Y") == year and (
                                    data.cgbalancing_date_status == 'Under process' or data.cgbalancing_date_status == 'Observation(same stage)' or data.cgbalancing_date_status == 'Halt'):
                                ListItems.append(data)

                if ParentStatus == 'Post-HIL':
                    dataList = dataList.extra(
                        select={'year': 'extract (year from post_hil_date)',
                                'month': 'extract (month from post_hil_date)',
                                'day': 'extract (day from post_hil_date)'},
                        order_by=['month', 'day', '-year'])
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(post_hil_date__year=year, post_hil_status=ChildStatus)

                    if ChildStatus == 'Current Count':
                        for data in dataList:
                            if data.post_hil_date.strftime("%Y") == year and (
                                    data.post_hil_status == 'Under process' or data.post_hil_status == 'Observation(same stage)' or data.post_hil_status == 'Halt'):
                                ListItems.append(data)

                if ParentStatus == 'System Alignment':
                    # dataList = dataList.extra(
                    #     select={'year': 'extract (year from sys_align_Date)',
                    #             'month': 'extract (month from sys_align_Date)',
                    #             'day': 'extract (day from sys_align_Date)'},
                    #     order_by=['month', 'day', '-year'])
                    dataList = dataList.annotate(sys_align_Date__month=Extract('sys_align_Date', 'month'),
                        sys_align_Date__year=Extract('sys_align_Date', 'year'),
                        sys_align_Date__day=Extract('sys_align_Date', 'day')).order_by('sys_align_Date__month','sys_align_Date__day','-sys_align_Date__year')
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(sys_align_Date__year=year, sys_align_status=ChildStatus)

                    if ChildStatus == 'Current Count':
                        for data in dataList:
                            if data.sys_align_Date.strftime("%Y") == year and (
                                    data.sys_align_status == 'Under process' or data.sys_align_status == 'Observation(same stage)' or data.sys_align_status == 'Halt'):
                                ListItems.append(data)

                if ParentStatus == 'Incapsulation':
                    dataList = dataList.extra(
                        select={'year': 'extract (year from incapsulation_date)',
                                'month': 'extract (month from incapsulation_date)',
                                'day': 'extract (day from incapsulation_date)'},
                        order_by=['month', 'day', '-year'])
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(incapsulation_date__year=year, incapsulation_status=ChildStatus)

                    if ChildStatus == 'Current Count':
                        for data in dataList:
                            if data.incapsulation_date.strftime("%Y") == year and (
                                    data.incapsulation_status == 'Under process' or data.incapsulation_status == 'Observation(same stage)' or data.incapsulation_status == 'Halt'):
                                ListItems.append(data)

                if ParentStatus == 'Final Integrated Testing':
                    dataList = dataList.extra(
                        select={'year': 'extract (year from final_integration_date)',
                                'month': 'extract (month from final_integration_date)',
                                'day': 'extract (day from final_integration_date)'},
                        order_by=['month', 'day', '-year'])
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(final_integration_date__year=year,
                                                    final_integration_status=ChildStatus)

                    if ChildStatus == 'Current Count':
                        for data in dataList:
                            if data.final_integration_date.strftime("%Y") == year and (
                                    data.final_integration_status == 'Under process' or data.final_integration_status == 'Observation(same stage)' or data.final_integration_status == 'Halt'):
                                ListItems.append(data)

                if ParentStatus == 'FGT Status':
                    dataList = dataList.extra(
                        select={'year': 'extract (year from fgt_date)',
                                'month': 'extract (month from fgt_date)',
                                'day': 'extract (day from fgt_date)'},
                        order_by=['month', 'day', '-year'])
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(fgt_date__year=year, fgt_status=ChildStatus)

                    if ChildStatus == 'Current Count':
                        for data in dataList:
                            if data.fgt_date.strftime("%Y") == year and (
                                    data.fgt_status == 'Under process' or data.fgt_status == 'Observation(same stage)' or data.fgt_status == 'Halt'):
                                ListItems.append(data)

                if ParentStatus == 'BHD Status':
                    dataList = dataList.extra(
                        select={'year': 'extract (year from bhd_date)',
                                'month': 'extract (month from bhd_date)',
                                'day': 'extract (day from bhd_date)'},
                        order_by=['month', 'day', '-year'])
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(bhd_date__year=year, bhd_status=ChildStatus)

                    if ChildStatus == 'Current Count':
                        for data in dataList:
                            if data.bhd_date.strftime("%Y") == year and (
                                    data.bhd_status == 'Ok' or data.bhd_status == 'Not Submitted)' or data.bhd_status == 'QM Observations Forwarded'):
                                ListItems.append(data)

                if ParentStatus == 'FQM Status':
                    dataList = dataList.extra(
                        select={'year': 'extract (year from fqm_date)',
                                'month': 'extract (month from fqm_date)',
                                'day': 'extract (day from fqm_date)'},
                        order_by=['month', 'day', '-year'])
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(fqm_date__year=year, fqm_status=ChildStatus)

                    if ChildStatus == 'Current Count':
                        for data in dataList:
                            if data.fqm_date.strftime("%Y") == year and (data.fqm_status == 'Not Conducted)'):
                                ListItems.append(data)

                if ParentStatus == 'QM Certification Status':
                    dataList = dataList.extra(
                        select={'year': 'extract (year from qm_certification_date)',
                                'month': 'extract (month from qm_certification_date)',
                                'day': 'extract (day from qm_certification_date)'},
                        order_by=['month', 'day', '-year'])
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(qm_certification_date__year=year,
                                                    qm_certification_status=ChildStatus)

                    if ChildStatus == 'Current Count':
                        for data in dataList:
                            if data.qm_certification_date.strftime("%Y") == year and (
                                    data.qm_certification_status == 'QM certificate issued' or data.qm_certification_status == 'QM Observations Forwarded)'):
                                ListItems.append(data)

                if ParentStatus == 'Launch/ End User':
                    dataList = dataList.extra(
                        select={'year': 'extract (year from enduser_date)',
                                'month': 'extract (month from enduser_date)',
                                'day': 'extract (day from enduser_date)'},
                        order_by=['month', 'day', '-year'])
                    if ChildStatus != 'Current Count':
                        ListItems = dataList.filter(enduser_date__year = year, enduser_status = ChildStatus)
                    else:
                        for data  in dataList:
                            if data.enduser_date.strftime("%Y") == year and (data.enduser_status == 'Ok' or data.enduser_status=='Observation' or data.enduser_status=='Halt'):
                                ListItems.append(data)

                serializer = RelifingSystemSerialzer(ListItems, many=True)
                return JsonResponse({'message': 'Welcome to Home Page', 'data': serializer.data }, status=200)
            # serializer = RelifingSystemSerialzer(dataList, many=True)
            # return JsonResponse({'message': 'Welcome to Home Page', 'data': serializer.data}, status=200)
        # except:
        #     return JsonResponse({'message': 'Sorry! No Task found.'}, status=200)

    @staticmethod
    def GetRelifingHistoryList(request):
        try:
            id = request.query_params['id']
            ParentStatus = request.query_params['parent_status']
            dataList = RelifingSystemStatusHistory.objects.filter(r_id=id)
            if ParentStatus == 'BLT':
                dataList = dataList.extra(
                    select={
                        'year': 'extract (year from blt_date)',
                        'month': 'extract (month from blt_date)',
                        'day': 'extract (day from blt_date)'},
                    order_by=['month', 'day', '-year'])


            if ParentStatus == 'EMP Proofing':
                dataList = dataList.extra(
                    select={'year': 'extract (year from emp_proofing_date)',
                            'month': 'extract (month from emp_proofing_date)',
                            'day': 'extract (day from emp_proofing_date)'},
                    order_by=['month', 'day', '-year'])

            if ParentStatus == 'Functional Test W/O Dummy Bird':
                dataList = dataList.extra(
                    select={'year': 'extract (year from func_tst_date)',
                            'month': 'extract (month from func_tst_date)',
                            'day': 'extract (day from func_tst_date)'},
                    order_by=['month', 'day', '-year'])

            if ParentStatus == 'Functional Test With Dummy Bird':
                dataList = dataList.extra(
                    select={'year': 'extract (year from func_tst_dummy_bird_date)',
                            'month': 'extract (month from func_tst_dummy_bird_date)',
                            'day': 'extract (day from func_tst_dummy_bird_date)'},
                    order_by=['month', 'day', '-year'])

            if ParentStatus == 'Road Test':
                dataList = dataList.extra(
                    select={'year': 'extract (year from road_test_date)',
                            'month': 'extract (month from road_test_date)',
                            'day': 'extract (day from road_test_date)'},
                    order_by=['month', 'day', '-year'])

            if ParentStatus == 'Post Road Test':
                dataList = dataList.extra(
                    select={'year': 'extract (year from post_road_test_date)',
                            'month': 'extract (month from post_road_test_date)',
                            'day': 'extract (day from post_road_test_date)'},
                    order_by=['month', 'day', '-year'])

            if ParentStatus == 'Integrated Operation':
                dataList = dataList.extra(
                    select={'year': 'extract (year from integrated_operation_date)',
                            'month': 'extract (month from integrated_operation_date)',
                            'day': 'extract (day from integrated_operation_date)'},
                    order_by=['month', 'day', '-year'])

            if ParentStatus == 'Rain Test':
                dataList = dataList.extra(
                    select={'year': 'extract (year from rain_test_date)',
                            'month': 'extract (month from rain_test_date)',
                            'day': 'extract (day from rain_test_date)'},
                    order_by=['month', 'day', '-year'])

            if ParentStatus == 'Pre User Inspection':
                dataList = dataList.extra(
                    select={'year': 'extract (year from pre_user_inspection_date)',
                            'month': 'extract (month from pre_user_inspection_date)',
                            'day': 'extract (day from pre_user_inspection_date)'},
                    order_by=['month', 'day', '-year'])

            if ParentStatus == 'Final Integration':
                dataList = dataList.extra(
                    select={'year': 'extract (year from final_integration_date)',
                            'month': 'extract (month from final_integration_date)',
                            'day': 'extract (day from final_integration_date)'},
                    order_by=['month', 'day', '-year'])

            if ParentStatus == 'Loading/Unloading on MLV/HLF':
                dataList = dataList.extra(
                    select={'year': 'extract (year from load_unload_on_mlv_hlf_date)',
                            'month': 'extract (month from load_unload_on_mlv_hlf_date)',
                            'day': 'extract (day from load_unload_on_mlv_hlf_date)'},
                    order_by=['month', 'day', '-year'])

            if ParentStatus == 'Pre-HIL':
                dataList = dataList.extra(
                    select={'year': 'extract (year from pre_hil_date)',
                            'month': 'extract (month from pre_hil_date)',
                            'day': 'extract (day from pre_hil_date)'},
                    order_by=['month', 'day', '-year'])

            if ParentStatus == 'Vibration':
                dataList = dataList.extra(
                    select={'year': 'extract (year from vibaration_date)',
                            'month': 'extract (month from vibaration_date)',
                            'day': 'extract (day from vibaration_date)'},
                    order_by=['month', 'day', '-year'])

            if ParentStatus == 'CG Balancing':
                dataList = dataList.extra(
                    select={'year': 'extract (year from cgbalancing_date)',
                            'month': 'extract (month from cgbalancing_date)',
                            'day': 'extract (day from cgbalancing_date)'},
                    order_by=['month', 'day', '-year'])

            if ParentStatus == 'Post-HIL':
                dataList = dataList.extra(
                    select={'year': 'extract (year from post_hil_date)',
                            'month': 'extract (month from post_hil_date)',
                            'day': 'extract (day from post_hil_date)'},
                    order_by=['month', 'day', '-year'])

            if ParentStatus == 'System Alignment':
                dataList = dataList.extra(
                    select={'year': 'extract (year from sys_align_Date)',
                            'month': 'extract (month from sys_align_Date)',
                            'day': 'extract (day from sys_align_Date)'},
                    order_by=['month', 'day', '-year'])


            if ParentStatus == 'Incapsulation':
                dataList = dataList.extra(
                    select={'year': 'extract (year from incapsulation_date)',
                            'month': 'extract (month from incapsulation_date)',
                            'day': 'extract (day from incapsulation_date)'},
                    order_by=['month', 'day', '-year'])


            if ParentStatus == 'Final Integrated Testing':
                dataList = dataList.extra(
                    select={'year': 'extract (year from final_integration_date)',
                            'month': 'extract (month from final_integration_date)',
                            'day': 'extract (day from final_integration_date)'},
                    order_by=['month', 'day', '-year'])


            if ParentStatus == 'FGT Status':
                dataList = dataList.extra(
                    select={'year': 'extract (year from fgt_date)',
                            'month': 'extract (month from fgt_date)',
                            'day': 'extract (day from fgt_date)'},
                    order_by=['month', 'day', '-year'])

            if ParentStatus == 'BHD Status':
                dataList = dataList.extra(
                    select={'year': 'extract (year from bhd_date)',
                            'month': 'extract (month from bhd_date)',
                            'day': 'extract (day from bhd_date)'},
                    order_by=['month', 'day', '-year'])

            if ParentStatus == 'FQM Status':
                dataList = dataList.extra(
                    select={'year': 'extract (year from fqm_date)',
                            'month': 'extract (month from fqm_date)',
                            'day': 'extract (day from fqm_date)'},
                    order_by=['month', 'day', '-year'])


            if ParentStatus == 'QM Certification Status':
                dataList = dataList.extra(
                    select={'year': 'extract (year from qm_certification_date)',
                            'month': 'extract (month from qm_certification_date)',
                            'day': 'extract (day from qm_certification_date)'},
                    order_by=['month', 'day', '-year'])

            if ParentStatus == 'Launch/ End User':
                dataList = dataList.extra(
                    select={'year': 'extract (year from enduser_date)',
                            'month': 'extract (month from enduser_date)',
                            'day': 'extract (day from enduser_date)'},
                    order_by=['month', 'day', '-year'])
            serializer = RelifingSystemSerialzer(dataList, many=True)
            return JsonResponse({'message': 'Welcome to Home Page', 'data': serializer.data}, status=200)

        except:
            return JsonResponse({'message': 'Sorry! No Task found.'}, status=200)

    # API for PDF generator with table and Text_wraping

    @staticmethod
    def GetRelifingPDFList(request):

        data = []
        status = request.query_params['status']
        if status == 'Total':
            data = RelifingSystemStatus.objects.all().order_by('-id').values_list('system', 'organization', 'set_id',
                                                                                  'blt_status', 'pre_hil_status',
                                                                                  'vibaration_status',
                                                                                  'post_hil_status',
                                                                                  'fgt_status',
                                                                                  'final_integration_status',
                                                                                  'bhd_status', 'fqm_status',
                                                                                  'qm_certification_status', 'remarks')
        else:
            data = RelifingSystemStatus.objects.filter(qm_certification_status=status).order_by('-id').values_list(
                'system', 'organization', 'set_id',
                'blt_status', 'pre_hil_status',
                'vibaration_status',
                'post_hil_status',
                'fgt_status',
                'final_integration_status',
                'bhd_status', 'fqm_status',
                'qm_certification_status', 'remarks')

        TABLE_COL_NAMES = (
            "System", "Organization", "Set ID", "BLT Status", "Pre-Hil Status", "Vibration Status", "Post-Hil Status",
            "FGT Status", "Final Integration Status", "BHD Status", "FQM Status", "QM Certification Status",
            "Remarks")

        pdf = FPDF('L', 'mm', 'Legal')
        pdf.add_page()
        pdf.set_font('courier', 'B', 26)
        pdf.cell(330, 10, 'Relifing Report', border=0, align='C', ln=2)
        pdf.cell(40, 10, '', 0, 1)
        pdf.set_font("Times", size=10)
        line_height = pdf.font_size * 3.2
        col_width = pdf.epw / 14

        def render_table_header():
            pdf.set_font(style="B")
            for col_name in TABLE_COL_NAMES:
                pdf.multi_cell(col_width, line_height, col_name, border=1, align='C', ln=3,
                               max_line_height=pdf.font_size)
            pdf.ln(line_height)
            pdf.set_font(style="")

        render_table_header()

        lh_list = []
        use_default_height = 0
        for row in data:
            for datum in row:
                dd = str(datum)
                word_list = dd.split()
                number_of_words = len(word_list)
                if number_of_words > 2:
                    use_default_height = 1
                    new_line_height = pdf.font_size * 2 * (number_of_words / 2)
            if not use_default_height:
                lh_list.append(line_height)
            else:
                lh_list.append(new_line_height)
                use_default_height = 0

        for j, row in enumerate(data):
            line_height = lh_list[j]
            if pdf.will_page_break(line_height):
                render_table_header()
            for col_num in range(len(row)):
                line_height = lh_list[j]
                pdf.multi_cell(col_width, line_height, f"{row[col_num]}", border=1, align='C', ln=3,
                               max_line_height=pdf.font_size)
            pdf.ln(line_height)
        pdf.output('RelifingReport.pdf')
        return FileResponse(open('RelifingReport.pdf', 'rb'), as_attachment=True, content_type='application/pdf')

    @staticmethod
    def GetRelifingExcelList(request):
        data = []
        status = request.query_params['status']
        if status == 'Total':
            data = RelifingSystemStatus.objects.all().order_by('-id').values_list('system', 'organization', 'set_id',
                                                                                  'blt_status', 'pre_hil_status',
                                                                                  'vibaration_status',
                                                                                  'post_hil_status',
                                                                                  'fgt_status',
                                                                                  'final_integration_status',
                                                                                  'bhd_status', 'fqm_status',
                                                                                  'qm_certification_status', 'remarks')
        else:
            data = RelifingSystemStatus.objects.filter(qm_certification_status=status).order_by('-id').values_list(
                'system', 'organization', 'set_id',
                'blt_status', 'pre_hil_status',
                'vibaration_status',
                'post_hil_status',
                'fgt_status',
                'final_integration_status',
                'bhd_status', 'fqm_status',
                'qm_certification_status', 'remarks')

        response = HttpResponse(content_type='application/ms-excel')
        response['Content-Disposition'] = 'attachment; filename="DocumentExcelSheet.xls"'
        work_book = xlwt.Workbook(encoding='utf-8')
        work_sheet = work_book.add_sheet('ExcelSheet')
        row_num = 0
        font_style = xlwt.XFStyle()
        font_style.font.bold = True
        TABLE_COL_NAMES = (
            "System", "Organization", "Set ID", "BLT Status", "Pre-Hil Status", "Vibration Status", "Post-Hil Status",
            "FGT Status", "Final Integration Status", "BHD Status", "FQM Status", "QM Certification Status",
            "Remarks")

        for col_num in range(len(TABLE_COL_NAMES)):
            work_sheet.write(row_num, col_num, TABLE_COL_NAMES[col_num], font_style)
        font_style = xlwt.XFStyle()

        for row in data:
            row_num += 1
            for col_num in range(len(row)):
                work_sheet.write(row_num, col_num, str(row[col_num]), font_style)
        work_book.save(response)
        return response

    @staticmethod
    def DeleteProdSysHistory(request):
        try:
            id = request.query_params['id']
            print(id)
            history = ProductionSystemStatusHistory.objects.filter(id=id)
            history.delete()
            return JsonResponse({'status': 'True', 'message': "Record Deleted"},
                                status=200)

        except Exception as e:
            print(e)
            return JsonResponse({'status': 'False', "message": "Doc Not Deleted"}, status=500)
            pass

    @staticmethod
    def DeleteFlightSysHistory(request):
        try:
            id = request.query_params['id']
            history = FlightSystemStatusHistory.objects.filter(id=id)
            history.delete()
            return JsonResponse({'status': 'True', 'message': "Record Deleted"},
                                status=200)

        except Exception as e:
            print(e)
            return JsonResponse({'status': 'False', "message": "Doc Not Deleted"}, status=500)
            pass

    @staticmethod
    def DeleteRelifingSysHistory(request):
        try:
            id = request.query_params['id']
            delete = RelifingSystemStatusHistory.objects.filter(id=id).delete()
            return JsonResponse({'status': 'True', 'message': "Record Deleted"},
                                status=200)

        except Exception as e:
            print(e)
            return JsonResponse({'status': 'False', "message": "Doc Not Deleted"}, status=500)
            pass
    @staticmethod
    def getSystemMonitoringDashboardCount(request):
        try:

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
            org = request.query_params['selected_org']
            type = request.query_params['selected_type']
            system = request.query_params['selected_system']

            # create dynamic filter

            if org != '':
                filter_objects &= get_filter(
                    'organization', 'equal',
                    org)
            if type != '':
                filter_objects &= get_filter(
                    'sys_type', 'equal',
                    type)
            if system != '':
                filter_objects &= get_filter(
                    'system', 'equal',
                    system)
            setid_filter = Q()
            setid_filter &= get_filter(
                    'set_id', 'not_equal',
                    '')
            if org != '':
                setid_filter &= get_filter(
                    'organization', 'equal',
                    org)
            if type != '':
                setid_filter &= get_filter(
                    'sys_type', 'equal',
                    type)
            if system != '':
                setid_filter &= get_filter(
                    'system', 'equal',
                    system)
            # if year != '':
            #     setid_filter &= get_filter(
            #         'set_id', 'not_equal',
            #         '')
            # production system count
            # prod_blt_count = ProductionSystemStatus.objects.filter(setid_filter, blt_date__year=year).count()
            totCount = ProductionSystemStatus.objects.filter(setid_filter)
            sys_count = 0
            for set in totCount:
                count_id = 0
                if set.blt_status != 'None' and set.set_id != '' and set.blt_date is not None:
                    if count_id == 0 and set.blt_date.strftime("%Y") == year:
                        count_id = set.id
                        sys_count += 1
                if set.pre_hil_status != 'None' and set.set_id != '' and set.pre_hil_date is not None:
                    if count_id == 0 and set.pre_hil_date.strftime("%Y") == year:
                        count_id = set.id
                        sys_count += 1
                if set.vibaration_status != 'None' and set.set_id != '' and set.vibaration_date is not None:
                    if count_id == 0 and set.vibaration_date.strftime("%Y") == year:
                        count_id = set.id
                        sys_count += 1
                if set.post_hil_status != 'None' and set.set_id != '' and set.post_hil_date is not None:
                    if count_id == 0 and set.post_hil_date.strftime("%Y") == year:
                        count_id = set.id
                        sys_count += 1
                if set.fgt_status != 'None' and set.set_id != '' and set.fgt_date is not None:
                    if count_id == 0 and set.fgt_date.strftime("%Y") == year:
                        count_id = set.id
                        sys_count += 1

                if set.final_integration_status != 'None' and set.set_id != '' and set.final_integration_date is not None:
                    if count_id == 0 and set.final_integration_date.strftime("%Y") == year:
                        count_id = set.id
                        sys_count += 1
                if set.bhd_status != 'None' and set.set_id != '' and set.bhd_date is not None:
                    if count_id == 0 and set.bhd_date.strftime("%Y") == year:
                        count_id = set.id
                        sys_count += 1
                if set.fqm_status != 'None' and set.set_id != '' and set.fqm_date is not None:
                    if count_id == 0 and set.fqm_date.strftime("%Y") == year:
                        count_id = set.id
                        sys_count += 1
                if set.qm_certification_status != 'None' and set.set_id != '' and set.qm_certification_date is not None:
                    if count_id == 0 and set.qm_certification_date.strftime("%Y") == year:
                        count_id = set.id
                        sys_count += 1
                if set.cgbalancing_date_status != 'None' and set.set_id != '' and set.cgbalancing_date is not None:
                    if count_id == 0 and set.cgbalancing_date.strftime("%Y") == year:
                        count_id = set.id
                        sys_count += 1

                if set.enduser_status != 'None' and set.set_id != '' and set.enduser_date is not None:
                    if count_id == 0 and set.enduser_date.strftime("%Y") == year:
                        count_id = set.id
                        sys_count += 1
                if set.incapsulation_status != 'None' and set.set_id != '' and set.incapsulation_date is not None:
                    if count_id == 0 and set.incapsulation_date.strftime("%Y") == year:
                        count_id = set.id
                        sys_count += 1
                if set.sys_align_status != 'None' and set.set_id != '' and set.sys_align_Date is not None:
                    if count_id == 0 and set.sys_align_Date.strftime("%Y") == year:
                        count_id = set.id
                        sys_count += 1
                if set.emp_proofing != 'None' and set.set_id != '' and set.emp_proofing_date is not None:
                    if count_id == 0 and set.emp_proofing_date.strftime("%Y") == year:
                        count_id = set.id
                        sys_count += 1
                if set.func_tst != 'None' and set.set_id != '' and set.func_tst_date is not None:
                    if count_id == 0 and set.func_tst_date.strftime("%Y") == year:
                        count_id = set.id
                        sys_count += 1

                if set.func_tst_dummy_bird != 'None' and set.set_id != '' and set.func_tst_dummy_bird_date is not None:
                    if count_id == 0 and set.func_tst_dummy_bird_date.strftime("%Y") == year:
                        count_id = set.id
                        sys_count += 1
                if set.road_test != 'None' and set.set_id != '' and set.road_test_date is not None:
                    if count_id == 0 and set.road_test_date.strftime("%Y") == year:
                        count_id = set.id
                        sys_count += 1
                if set.post_road_test != 'None' and set.set_id != '' and set.post_road_test_date is not None:
                    if count_id == 0 and set.post_road_test_date.strftime("%Y") == year:
                        count_id = set.id
                        sys_count += 1
                if set.integrated_operation != 'None' and set.set_id != '' and set.integrated_operation_date is not None:
                    if count_id == 0 and set.integrated_operation_date.strftime("%Y") == year:
                        count_id = set.id
                        sys_count += 1
                if set.rain_test != 'None' and set.set_id != '' and set.rain_test_date is not None:
                    if count_id == 0 and set.rain_test_date.strftime("%Y") == year:
                        count_id = set.id
                        sys_count += 1

                if set.pre_user_inspection != 'None' and set.set_id != '' and set.pre_user_inspection_date is not None:
                    if count_id == 0 and set.pre_user_inspection_date.strftime("%Y") == year:
                        count_id = set.id
                        sys_count += 1
                if set.final_integrated_testing != 'None' and set.set_id != '' and set.final_integrated_testing_date is not None:
                    if count_id == 0 and set.final_integrated_testing_date.strftime("%Y") == year:
                        count_id = set.id
                        sys_count += 1
                if set.load_unload_on_mlv_hlf != 'None' and set.set_id != '' and set.load_unload_on_mlv_hlf_date is not None:
                    if count_id == 0 and set.load_unload_on_mlv_hlf_date.strftime("%Y") == year:
                        count_id = set.id
                        sys_count += 1


            prodTotalCount = sys_count
            # print(prod_blt_count)
            prod_blt_oklist = ProductionSystemStatus.objects.filter(filter_objects, blt_date__year=year,blt_status='Ok')
            prod_blt_oklistNext = ProductionSystemStatus.objects.filter(filter_objects, blt_date__year=year,blt_status='OK(next stage)')
            prod_blt_observationlist = ProductionSystemStatus.objects.filter(filter_objects,
                                                                             blt_date__year=year,blt_status='Observation(same stage)')
            prod_blt_observationlist_next = ProductionSystemStatus.objects.filter(filter_objects,
                                                                                  blt_date__year=year,blt_status='Observation(next stage)')
            prod_blt_uplist = ProductionSystemStatus.objects.filter(filter_objects, blt_date__year=year,blt_status='Under process')
            prod_blt_haultlist = ProductionSystemStatus.objects.filter(filter_objects, blt_status='Halt')

            prod_prehil_count = ProductionSystemStatus.objects.filter(setid_filter, pre_hil_date__year=year).count()
            print(prod_prehil_count)

            prod_prehil_oklist = ProductionSystemStatus.objects.filter(filter_objects, pre_hil_date__year=year,pre_hil_status='Ok')
            prod_prehil_oklistNext = ProductionSystemStatus.objects.filter(filter_objects,
                                                                           pre_hil_date__year=year,pre_hil_status='OK(next stage)')
            prod_prehil_observationlist = ProductionSystemStatus.objects.filter(filter_objects,
                                                                                pre_hil_date__year=year,pre_hil_status='Observation(same stage)')
            prod_prehil_observationlistNext = ProductionSystemStatus.objects.filter(filter_objects,
                                                                                    pre_hil_date__year=year,pre_hil_status='Observation(next stage)')
            prod_prehil_uplist = ProductionSystemStatus.objects.filter(filter_objects, pre_hil_date__year=year,pre_hil_status='Under process')
            prod_prehil_haultlist = ProductionSystemStatus.objects.filter(filter_objects, pre_hil_date__year=year,pre_hil_status='Halt')


            prod_posthil_count = ProductionSystemStatus.objects.filter(setid_filter, post_hil_date__year=year).count()
            print(prod_posthil_count)

            prod_posthil_oklist = ProductionSystemStatus.objects.filter(filter_objects,
                                                                        post_hil_date__year=year,post_hil_status='Ok')
            prod_posthil_oklistNext = ProductionSystemStatus.objects.filter(filter_objects,
                                                                            post_hil_date__year=year,post_hil_status='OK(next stage)')
            prod_posthil_observationlist = ProductionSystemStatus.objects.filter(filter_objects,
                                                                                 post_hil_date__year=year,post_hil_status='Observation(same stage)')
            prod_posthil_observationlistNext = ProductionSystemStatus.objects.filter(filter_objects,
                                                                                     post_hil_date__year=year,post_hil_status='Observation(next stage)')
            prod_posthil_uplist = ProductionSystemStatus.objects.filter(filter_objects, post_hil_date__year=year,post_hil_status='Under process')
            prod_posthil_haultlist = ProductionSystemStatus.objects.filter(filter_objects, post_hil_date__year=year,post_hil_status='Halt')

            prod_finalintegration_count = ProductionSystemStatus.objects.filter(setid_filter, final_integration_date__year=year).count()
            print(prod_finalintegration_count)

            prod_finalintegration_oklist = ProductionSystemStatus.objects.filter(filter_objects,
                                                                                 final_integration_date__year=year,final_integration_status='Ok')
            prod_finalintegration_oklistNext = ProductionSystemStatus.objects.filter(filter_objects,
                                                                                     final_integration_date__year=year,final_integration_status='OK(next stage)')

            prod_finalintegration_observationlist = ProductionSystemStatus.objects.filter(filter_objects,
                                                                                          final_integration_date__year=year,final_integration_status='Observation(same stage)')

            prod_finalintegration_observationlistNext = ProductionSystemStatus.objects.filter(filter_objects,
                                                                                              final_integration_date__year=year,final_integration_status='Observation(next stage)')

            prod_finalintegration_completelist = ProductionSystemStatus.objects.filter(filter_objects,
                                                                                       final_integration_date__year=year,final_integration_status='Under process')
            prod_finalintegration_haultlist = ProductionSystemStatus.objects.filter(filter_objects,
                                                                                    final_integration_date__year=year,final_integration_status='Halt')

            prod_vibration_count = ProductionSystemStatus.objects.filter(setid_filter, vibaration_date__year=year).count()
            print(prod_vibration_count)

            prod_vibration_oklist = ProductionSystemStatus.objects.filter(filter_objects,
                                                                          vibaration_date__year=year,vibaration_status='Ok')
            prod_vibration_oklistNext = ProductionSystemStatus.objects.filter(filter_objects,
                                                                              vibaration_date__year=year,vibaration_status='OK(next stage)')

            prod_vibration_observationlist = ProductionSystemStatus.objects.filter(filter_objects,
                                                                                   vibaration_date__year=year,vibaration_status='Observation(same stage)')
            prod_vibration_observationlistNext = ProductionSystemStatus.objects.filter(filter_objects,
                                                                                       vibaration_date__year=year,vibaration_status='Observation(next stage)')
            prod_vibration_uplist = ProductionSystemStatus.objects.filter(filter_objects,
                                                                          vibaration_date__year=year,vibaration_status='Under process')
            prod_vibration_haultlist = ProductionSystemStatus.objects.filter(filter_objects, vibaration_date__year=year,vibaration_status='Halt')


            prod_cg_count = ProductionSystemStatus.objects.filter(setid_filter, cgbalancing_date__year=year).count()
            print(prod_cg_count)

            prod_cg_oklist = ProductionSystemStatus.objects.filter(filter_objects,
                                                                   cgbalancing_date__year=year,cgbalancing_date_status='Ok')

            prod_cg_oklistNext = ProductionSystemStatus.objects.filter(filter_objects,
                                                                       cgbalancing_date__year=year,cgbalancing_date_status='OK(next stage)')

            prod_cg_observationlist = ProductionSystemStatus.objects.filter(filter_objects,
                                                                            cgbalancing_date__year=year,cgbalancing_date_status='Observation(same stage)')
            prod_cg_observationlistNext = ProductionSystemStatus.objects.filter(filter_objects,
                                                                                cgbalancing_date__year=year,cgbalancing_date_status='Observation(next stage)')
            prod_cg_uplist = ProductionSystemStatus.objects.filter(filter_objects,
                                                                   cgbalancing_date__year=year,cgbalancing_date_status='Under process')
            prod_cg_haultlist = ProductionSystemStatus.objects.filter(filter_objects, cgbalancing_date__year=year,cgbalancing_date_status='Halt')

            prod_fgt_count = ProductionSystemStatus.objects.filter(setid_filter, fgt_date__year=year).count()
            prod_fgt_oklist = ProductionSystemStatus.objects.filter(filter_objects, fgt_date__year=year,fgt_status='Ok')
            prod_fgt_oklistNext = ProductionSystemStatus.objects.filter(filter_objects, fgt_date__year=year,fgt_status='OK(next stage)')
            prod_fgt_observationlist = ProductionSystemStatus.objects.filter(filter_objects,
                                                                             fgt_date__year=year,fgt_status='Observation(same stage)')
            prod_fgt_observationlistNext = ProductionSystemStatus.objects.filter(filter_objects,
                                                                                 fgt_date__year=year,fgt_status='Observation(next stage)')
            prod_fgt_uplist = ProductionSystemStatus.objects.filter(filter_objects,
                                                                    fgt_date__year=year,fgt_status='Under process')
            prod_fgt_haultlist = ProductionSystemStatus.objects.filter(filter_objects, fgt_date__year=year,fgt_status='Halt')

            prod_bhd_count = ProductionSystemStatus.objects.filter(setid_filter, bhd_date__year=year).count()
            prod_bhd_ok = ProductionSystemStatus.objects.filter(filter_objects,
                                                                       bhd_date__year=year,bhd_status='Ok')
            prod_bhd_submitted = ProductionSystemStatus.objects.filter(filter_objects,
                                                                       bhd_date__year=year,bhd_status='Submitted')
            prod_bhd_not_submitted = ProductionSystemStatus.objects.filter(filter_objects,
                                                                       bhd_date__year=year,bhd_status='Not Submitted')
            prod_bhd_qm_forwarded = ProductionSystemStatus.objects.filter(filter_objects,
                                                                         bhd_date__year=year,bhd_status='QM Observations Forwarded')

            prod_bhd_qm_repeated = ProductionSystemStatus.objects.filter(filter_objects,
                                                                         bhd_date__year=year,bhd_status='QM Observations Repeated')

            prod_bhd_inprocess= ProductionSystemStatus.objects.filter(filter_objects,
                                                                         bhd_date__year=year,bhd_status='Audit in-process')

            prod_fqm_count = ProductionSystemStatus.objects.filter(setid_filter, fqm_date__year=year).count()
            prod_fqm_planned = ProductionSystemStatus.objects.filter(filter_objects,
                                                                     fqm_date__year=year,fqm_status='Planned')

            prod_fqm_conducted = ProductionSystemStatus.objects.filter(filter_objects,
                                                                       fqm_date__year=year,fqm_status='Conducted')

            prod_qmc_count = ProductionSystemStatus.objects.filter(setid_filter, qm_certification_date__year=year).count()
            prod_qmc_issued = ProductionSystemStatus.objects.filter(filter_objects,
                                                                    qm_certification_date__year=year,qm_certification_status='QM certificate issued')

            prod_qmc_in_process = ProductionSystemStatus.objects.filter(filter_objects,
                                                                        qm_certification_date__year=year,qm_certification_status='Audit in-process')
            prod_qmc_obs_forwarded = ProductionSystemStatus.objects.filter(filter_objects,
                                                                        qm_certification_date__year=year,qm_certification_status='QM Observations Forwarded')


            prod_enduser_count = ProductionSystemStatus.objects.filter(setid_filter, enduser_date__year=year).count()

            prod_enduser_oklist = ProductionSystemStatus.objects.filter(filter_objects, enduser_date__year=year,enduser_status='Ok')
            prod_enduser_oklistNext = ProductionSystemStatus.objects.filter(filter_objects,
                                                                            enduser_date__year=year,enduser_status='OK(next stage)')
            prod_enduser_observationlist = ProductionSystemStatus.objects.filter(filter_objects,
                                                                                 enduser_date__year=year,enduser_status='Observation')
            prod_enduser_observationlistNext = ProductionSystemStatus.objects.filter(filter_objects,
                                                                                     enduser_date__year=year,enduser_status='Observation(next stage)')
            prod_enduser_uplist = ProductionSystemStatus.objects.filter(filter_objects,
                                                                        enduser_date__year=year,enduser_status='Completed')
            prod_enduser_haultlist = ProductionSystemStatus.objects.filter(filter_objects, enduser_date__year=year,enduser_status='Halt')


            prod_sys_align_count = ProductionSystemStatus.objects.filter(setid_filter, sys_align_Date__year=year).count()

            prod_sys_align_oklist = ProductionSystemStatus.objects.filter(filter_objects, sys_align_Date__year=year,sys_align_status='Ok')
            prod_sys_align_oklistNext = ProductionSystemStatus.objects.filter(filter_objects,
                                                                            sys_align_Date__year=year,sys_align_status='OK(next stage)')
            prod_sys_align_observationlist = ProductionSystemStatus.objects.filter(filter_objects,
                                                                                 sys_align_Date__year=year,sys_align_status='Observation(same stage)')
            prod_sys_align_observationlistNext = ProductionSystemStatus.objects.filter(filter_objects,
                                                                                     sys_align_Date__year=year,sys_align_status='Observation(next stage)')
            prod_sys_align_uplist = ProductionSystemStatus.objects.filter(filter_objects,
                                                                        sys_align_Date__year=year,sys_align_status='Under process')
            prod_sys_align_haultlist = ProductionSystemStatus.objects.filter(filter_objects, sys_align_Date__year=year,sys_align_status='Halt')

            prod_incapsulation_count = ProductionSystemStatus.objects.filter(setid_filter, incapsulation_date__year=year).count()
            prod_incapsulation_oklist = ProductionSystemStatus.objects.filter(filter_objects, incapsulation_date__year=year,incapsulation_status='Ok')
            prod_incapsulation_oklistNext = ProductionSystemStatus.objects.filter(filter_objects,
                                                                            incapsulation_date__year=year,incapsulation_status='OK(next stage)')
            prod_incapsulation_observationlist = ProductionSystemStatus.objects.filter(filter_objects,
                                                                                 incapsulation_date__year=year,incapsulation_status='Observation(same stage)')
            prod_incapsulation_observationlistNext = ProductionSystemStatus.objects.filter(filter_objects,
                                                                                     incapsulation_date__year=year,incapsulation_status='Observation(next stage)')
            prod_incapsulation_uplist = ProductionSystemStatus.objects.filter(filter_objects,
                                                                        incapsulation_date__year=year,incapsulation_status='Under process')
            prod_incapsulation_haultlist = ProductionSystemStatus.objects.filter(filter_objects, incapsulation_date__year=year,incapsulation_status='Halt')

            prod_emp_proofing_count = ProductionSystemStatus.objects.filter(setid_filter, emp_proofing_date__year=year).count()

            prod_emp_proofing_oklist = ProductionSystemStatus.objects.filter(filter_objects, emp_proofing_date__year=year,emp_proofing='Ok')
            prod_emp_proofing_oklistNext= ProductionSystemStatus.objects.filter(filter_objects,emp_proofing_date__year=year,emp_proofing='OK(next stage)')
            prod_emp_proofing_observationlist  = ProductionSystemStatus.objects.filter(filter_objects,emp_proofing_date__year=year,emp_proofing='Observation(same stage)')
            prod_emp_proofing_observationlistNext = ProductionSystemStatus.objects.filter(filter_objects,emp_proofing_date__year=year,emp_proofing='Observation(next stage)')
            prod_emp_proofing_uplist = ProductionSystemStatus.objects.filter(filter_objects,emp_proofing_date__year=year,emp_proofing='Under process')
            prod_emp_proofing_haultlist = ProductionSystemStatus.objects.filter(filter_objects,emp_proofing_date__year=year,emp_proofing='Halt')

            prod_func_tst_count = ProductionSystemStatus.objects.filter(setid_filter, func_tst_date__year=year).count()

            prod_func_tst_oklistNext_oklist  = ProductionSystemStatus.objects.filter(filter_objects,func_tst_date__year=year,func_tst='Ok')
            prod_func_tst_oklistNext_oklistNext  = ProductionSystemStatus.objects.filter(filter_objects,func_tst_date__year=year,func_tst='OK(next stage)')
            prod_func_tst_oklistNext_observationlist   = ProductionSystemStatus.objects.filter(filter_objects,func_tst_date__year=year,func_tst='Observation(same stage)')
            prod_func_tst_oklistNext_observationlistNext  = ProductionSystemStatus.objects.filter(filter_objects,func_tst_date__year=year,func_tst='Observation(next stage)')
            prod_func_tst_oklistNext_uplist  = ProductionSystemStatus.objects.filter(filter_objects,func_tst_date__year=year,func_tst='Under process')
            prod_func_tst_oklistNext_haultlist  = ProductionSystemStatus.objects.filter(filter_objects,func_tst_date__year=year,func_tst='Halt')

            prod_func_tst_dummy_bird_count = ProductionSystemStatus.objects.filter(setid_filter, func_tst_dummy_bird_date__year=year).count()

            prod_func_tst_dummy_bird_oklist = ProductionSystemStatus.objects.filter(filter_objects,func_tst_dummy_bird_date__year=year,func_tst_dummy_bird='Ok')
            prod_func_tst_dummy_bird_oklistNext = ProductionSystemStatus.objects.filter(filter_objects,func_tst_dummy_bird_date__year=year,func_tst_dummy_bird='OK(next stage)')
            prod_func_tst_dummy_bird_observationlist  = ProductionSystemStatus.objects.filter(filter_objects,func_tst_dummy_bird_date__year=year,func_tst_dummy_bird='Observation(same stage)')
            prod_func_tst_dummy_bird_observationlistNext = ProductionSystemStatus.objects.filter(filter_objects,func_tst_dummy_bird_date__year=year,func_tst_dummy_bird='Observation(next stage)')
            prod_func_tst_dummy_bird_uplist = ProductionSystemStatus.objects.filter(filter_objects,func_tst_dummy_bird_date__year=year,func_tst_dummy_bird='Under process')
            prod_func_tst_dummy_bird_haultlist = ProductionSystemStatus.objects.filter(filter_objects,func_tst_dummy_bird_date__year=year,func_tst_dummy_bird='Halt')

            prod_road_test_count = ProductionSystemStatus.objects.filter(setid_filter, road_test_date__year=year).count()

            prod_road_test_oklist = ProductionSystemStatus.objects.filter(filter_objects,road_test_date__year=year,road_test='Ok')
            prod_road_test_oklistNext = ProductionSystemStatus.objects.filter(filter_objects,road_test_date__year=year,road_test='OK(next stage)')
            prod_road_test_observationlist  = ProductionSystemStatus.objects.filter(filter_objects,road_test_date__year=year,road_test='Observation(same stage)')
            prod_road_test_observationlistNext = ProductionSystemStatus.objects.filter(filter_objects,road_test_date__year=year,road_test='Observation(next stage)')
            prod_road_test_uplist = ProductionSystemStatus.objects.filter(filter_objects,road_test_date__year=year,road_test='Under process')
            prod_road_test_haultlist = ProductionSystemStatus.objects.filter(filter_objects,road_test_date__year=year,road_test='Halt')

            prod_post_road_test_count = ProductionSystemStatus.objects.filter(setid_filter, post_road_test_date__year=year).count()

            prod_post_road_test_oklist = ProductionSystemStatus.objects.filter(filter_objects,post_road_test_date__year=year,post_road_test='Ok')
            prod_post_road_test_oklistNext = ProductionSystemStatus.objects.filter(filter_objects,post_road_test_date__year=year,post_road_test='OK(next stage)')
            prod_post_road_test_observationlist  = ProductionSystemStatus.objects.filter(filter_objects,post_road_test_date__year=year,post_road_test='Observation(same stage)')
            prod_post_road_test_observationlistNext = ProductionSystemStatus.objects.filter(filter_objects,post_road_test_date__year=year,post_road_test='Observation(next stage)')
            prod_post_road_test_uplist = ProductionSystemStatus.objects.filter(filter_objects,post_road_test_date__year=year,post_road_test='Under process')
            prod_post_road_test_haultlist = ProductionSystemStatus.objects.filter(filter_objects,post_road_test_date__year=year,post_road_test='Halt')

            prod_integrated_operation_count = ProductionSystemStatus.objects.filter(setid_filter, integrated_operation_date__year=year).count()

            prod_integrated_operation_oklist = ProductionSystemStatus.objects.filter(filter_objects,integrated_operation_date__year=year,integrated_operation='Ok')
            prod_integrated_operation_oklistNext = ProductionSystemStatus.objects.filter(filter_objects,integrated_operation_date__year=year,integrated_operation='OK(next stage)')
            prod_integrated_operation_observationlist  = ProductionSystemStatus.objects.filter(filter_objects,integrated_operation_date__year=year,integrated_operation='Observation(same stage)')
            prod_integrated_operation_observationlistNext = ProductionSystemStatus.objects.filter(filter_objects,integrated_operation_date__year=year,integrated_operation='Observation(next stage)')
            prod_integrated_operation_uplist = ProductionSystemStatus.objects.filter(filter_objects,integrated_operation_date__year=year,integrated_operation='Under process')
            prod_integrated_operation_haultlist = ProductionSystemStatus.objects.filter(filter_objects,integrated_operation_date__year=year,integrated_operation='Halt')

            prod_rain_test_count = ProductionSystemStatus.objects.filter(setid_filter, rain_test_date__year=year).count()

            prod_rain_test_oklist = ProductionSystemStatus.objects.filter(filter_objects,rain_test_date__year=year,rain_test='Ok')
            prod_rain_test_oklistNext = ProductionSystemStatus.objects.filter(filter_objects,rain_test_date__year=year,rain_test='OK(next stage)')
            prod_rain_test_observationlist  = ProductionSystemStatus.objects.filter(filter_objects,rain_test_date__year=year,rain_test='Observation(same stage)')
            prod_rain_test_observationlistNext = ProductionSystemStatus.objects.filter(filter_objects,rain_test_date__year=year,rain_test='Observation(next stage)')
            prod_rain_test_uplist = ProductionSystemStatus.objects.filter(filter_objects,rain_test_date__year=year,rain_test='Under process')
            prod_rain_test_haultlist = ProductionSystemStatus.objects.filter(filter_objects,rain_test_date__year=year,rain_test='Halt')

            prod_pre_user_inspection_count = ProductionSystemStatus.objects.filter(setid_filter, pre_user_inspection_date__year=year).count()

            prod_pre_user_inspection_oklist = ProductionSystemStatus.objects.filter(filter_objects,pre_user_inspection_date__year=year,pre_user_inspection='Ok')
            prod_pre_user_inspection_oklistNext = ProductionSystemStatus.objects.filter(filter_objects,pre_user_inspection_date__year=year,pre_user_inspection='OK(next stage)')
            prod_pre_user_inspection_observationlist  = ProductionSystemStatus.objects.filter(filter_objects,pre_user_inspection_date__year=year,pre_user_inspection='Observation(same stage)')
            prod_pre_user_inspection_observationlistNext = ProductionSystemStatus.objects.filter(filter_objects,pre_user_inspection_date__year=year,pre_user_inspection='Observation(next stage)')
            prod_pre_user_inspection_uplist = ProductionSystemStatus.objects.filter(filter_objects,pre_user_inspection_date__year=year,pre_user_inspection='Under process')
            prod_pre_user_inspection_haultlist = ProductionSystemStatus.objects.filter(filter_objects,pre_user_inspection_date__year=year,pre_user_inspection='Halt')

            prod_final_integrated_testing_count = ProductionSystemStatus.objects.filter(setid_filter, final_integrated_testing_date__year=year).count()

            prod_final_integrated_testing_oklist = ProductionSystemStatus.objects.filter(filter_objects,final_integrated_testing_date__year=year,final_integrated_testing='Ok')
            prod_final_integrated_testing_oklistNext = ProductionSystemStatus.objects.filter(filter_objects,final_integrated_testing_date__year=year,final_integrated_testing='OK(next stage)')
            prod_final_integrated_testing_observationlist  = ProductionSystemStatus.objects.filter(filter_objects,final_integrated_testing_date__year=year,final_integrated_testing='Observation(same stage)')
            prod_final_integrated_testing_observationlistNext = ProductionSystemStatus.objects.filter(filter_objects,final_integrated_testing_date__year=year,final_integrated_testing='Observation(next stage)')
            prod_final_integrated_testing_uplist = ProductionSystemStatus.objects.filter(filter_objects,final_integrated_testing_date__year=year,final_integrated_testing='Under process')
            prod_final_integrated_testing_haultlist = ProductionSystemStatus.objects.filter(filter_objects,final_integrated_testing_date__year=year,final_integrated_testing='Halt')

            prod_load_unload_on_mlv_hlf_count = ProductionSystemStatus.objects.filter(setid_filter, load_unload_on_mlv_hlf_date__year=year).count()

            prod_load_unload_on_mlv_hlf_oklist = ProductionSystemStatus.objects.filter(filter_objects,load_unload_on_mlv_hlf_date__year=year,load_unload_on_mlv_hlf='Ok')
            prod_load_unload_on_mlv_hlf_oklistNext = ProductionSystemStatus.objects.filter(filter_objects,load_unload_on_mlv_hlf_date__year=year,load_unload_on_mlv_hlf='OK(next stage)')
            prod_load_unload_on_mlv_hlf_observationlist  = ProductionSystemStatus.objects.filter(filter_objects,load_unload_on_mlv_hlf_date__year=year,load_unload_on_mlv_hlf='Observation(same stage)')
            prod_load_unload_on_mlv_hlf_observationlistNext = ProductionSystemStatus.objects.filter(filter_objects,load_unload_on_mlv_hlf_date__year=year,load_unload_on_mlv_hlf='Observation(next stage)')
            prod_load_unload_on_mlv_hlf_uplist = ProductionSystemStatus.objects.filter(filter_objects,load_unload_on_mlv_hlf_date__year=year,load_unload_on_mlv_hlf='Under process')
            prod_load_unload_on_mlv_hlf_haultlist = ProductionSystemStatus.objects.filter(filter_objects,load_unload_on_mlv_hlf_date__year=year,load_unload_on_mlv_hlf='Halt')


            # prodTotalCount = prod_blt_count+prod_prehil_count+prod_vibration_count+prod_bhd_count+prod_posthil_count+prod_vibration_count+prod_vibration_count+prod_cg_count+prod_qmc_count+prod_fgt_count+prod_fqm_count+prod_emp_proofing_count+prod_enduser_count+prod_finalintegration_count+prod_road_test_count+prod_post_road_test_count+prod_integrated_operation_count+prod_rain_test_count+prod_sys_align_count+prod_pre_user_inspection_count+prod_load_unload_on_mlv_hlf_count+prod_final_integrated_testing_count+prod_func_tst_count+prod_func_tst_dummy_bird_count+prod_incapsulation_count

            # flight system count
            # flight_blt_count = FlightSystemStatus.objects.filter(setid_filter, blt_date__year=year).count()
            totCountFlight = FlightSystemStatus.objects.filter(setid_filter)
            # totCount = ProductionSystemStatus.objects.filter(setid_filter)
            flight_sys_count = 0
            for set in totCountFlight:
                count_id = 0
                if set.blt_status != 'None' and set.set_id != '' and set.blt_date is not None:
                    if count_id == 0 and set.blt_date.strftime("%Y") == year:
                        count_id = set.id
                        flight_sys_count += 1
                if set.pre_hil_status != 'None' and set.set_id != '' and set.pre_hil_date is not None:
                    if count_id == 0 and set.pre_hil_date.strftime("%Y") == year:
                        count_id = set.id
                        flight_sys_count += 1
                if set.vibaration_status != 'None' and set.set_id != '' and set.vibaration_date is not None:
                    if count_id == 0 and set.vibaration_date.strftime("%Y") == year:
                        count_id = set.id
                        flight_sys_count += 1
                if set.post_hil_status != 'None' and set.set_id != '' and set.post_hil_date is not None:
                    if count_id == 0 and set.post_hil_date.strftime("%Y") == year:
                        count_id = set.id
                        flight_sys_count += 1
                if set.fgt_status != 'None' and set.set_id != '' and set.fgt_date is not None:
                    if count_id == 0 and set.fgt_date.strftime("%Y") == year:
                        count_id = set.id
                        flight_sys_count += 1

                if set.final_integration_status != 'None' and set.set_id != '' and set.final_integration_date is not None:
                    if count_id == 0 and set.final_integration_date.strftime("%Y") == year:
                        count_id = set.id
                        flight_sys_count += 1
                if set.bhd_status != 'None' and set.set_id != '' and set.bhd_date is not None:
                    if count_id == 0 and set.bhd_date.strftime("%Y") == year:
                        count_id = set.id
                        flight_sys_count += 1
                if set.fqm_status != 'None' and set.set_id != '' and set.fqm_date is not None:
                    if count_id == 0 and set.fqm_date.strftime("%Y") == year:
                        count_id = set.id
                        flight_sys_count += 1
                if set.qm_certification_status != 'None' and set.set_id != '' and set.qm_certification_date is not None:
                    if count_id == 0 and set.qm_certification_date.strftime("%Y") == year:
                        count_id = set.id
                        flight_sys_count += 1
                if set.cgbalancing_date_status != 'None' and set.set_id != '' and set.cgbalancing_date is not None:
                    if count_id == 0 and set.cgbalancing_date.strftime("%Y") == year:
                        count_id = set.id
                        flight_sys_count += 1

                if set.launchact_status != 'None' and set.set_id != '' and set.launchact_date is not None:
                    if count_id == 0 and set.launchact_date.strftime("%Y") == year:
                        count_id = set.id
                        sys_count += 1
                if set.incapsulation_status != 'None' and set.set_id != '' and set.incapsulation_date is not None:
                    if count_id == 0 and set.incapsulation_date.strftime("%Y") == year:
                        count_id = set.id
                        flight_sys_count += 1
                if set.sys_align_status != 'None' and set.set_id != '' and set.sys_align_Date is not None:
                    if count_id == 0 and set.sys_align_Date.strftime("%Y") == year:
                        count_id = set.id
                        flight_sys_count += 1
                if set.emp_proofing != 'None' and set.set_id != '' and set.emp_proofing_date is not None:
                    if count_id == 0 and set.emp_proofing_date.strftime("%Y") == year:
                        count_id = set.id
                        flight_sys_count += 1
                if set.func_tst != 'None' and set.set_id != '' and set.func_tst_date is not None:
                    if count_id == 0 and set.func_tst_date.strftime("%Y") == year:
                        count_id = set.id
                        flight_sys_count += 1

                if set.func_tst_dummy_bird != 'None' and set.set_id != '' and set.func_tst_dummy_bird_date is not None:
                    if count_id == 0 and set.func_tst_dummy_bird_date.strftime("%Y") == year:
                        count_id = set.id
                        flight_sys_count += 1
                if set.road_test != 'None' and set.set_id != '' and set.road_test_date is not None:
                    if count_id == 0 and set.road_test_date.strftime("%Y") == year:
                        count_id = set.id
                        flight_sys_count += 1
                if set.post_road_test != 'None' and set.set_id != '' and set.post_road_test_date is not None:
                    if count_id == 0 and set.post_road_test_date.strftime("%Y") == year:
                        count_id = set.id
                        flight_sys_count += 1
                if set.integrated_operation != 'None' and set.set_id != '' and set.integrated_operation_date is not None:
                    if count_id == 0 and set.integrated_operation_date.strftime("%Y") == year:
                        count_id = set.id
                        flight_sys_count += 1
                if set.rain_test != 'None' and set.set_id != '' and set.rain_test_date is not None:
                    if count_id == 0 and set.rain_test_date.strftime("%Y") == year:
                        count_id = set.id
                        flight_sys_count += 1

                if set.pre_user_inspection != 'None' and set.set_id != '' and set.pre_user_inspection_date is not None:
                    if count_id == 0 and set.pre_user_inspection_date.strftime("%Y") == year:
                        count_id = set.id
                        flight_sys_count += 1
                if set.final_integrated_testing != 'None' and set.set_id != '' and set.final_integrated_testing_date is not None:
                    if count_id == 0 and set.final_integrated_testing_date.strftime("%Y") == year:
                        count_id = set.id
                        flight_sys_count += 1
                if set.load_unload_on_mlv_hlf != 'None' and set.set_id != '' and set.load_unload_on_mlv_hlf_date is not None:
                    if count_id == 0 and set.load_unload_on_mlv_hlf_date.strftime("%Y") == year:
                        count_id = set.id
                        flight_sys_count += 1

            flightTotalCount = flight_sys_count
            flight_blt_oklist = FlightSystemStatus.objects.filter(filter_objects, blt_date__year=year,blt_status='Ok')
            flight_blt_oklistNext = FlightSystemStatus.objects.filter(filter_objects, blt_date__year=year,blt_status='OK(next stage)')
            flight_blt_observationlist = FlightSystemStatus.objects.filter(filter_objects,
                                                                           blt_date__year=year,blt_status='Observation(same stage)')
            flight_blt_observationlist_next = FlightSystemStatus.objects.filter(filter_objects,
                                                                                blt_date__year=year,blt_status='Observation(next stage)')
            flight_blt_uplist = FlightSystemStatus.objects.filter(filter_objects, blt_date__year=year,blt_status='Under process')
            flight_blt_haultlist = FlightSystemStatus.objects.filter(filter_objects, blt_date__year=year,blt_status='Halt')

            flight_prehil_count = FlightSystemStatus.objects.filter(setid_filter, pre_hil_date__year=year).count()

            flight_prehil_oklist = FlightSystemStatus.objects.filter(filter_objects, pre_hil_date__year=year,pre_hil_status='Ok')
            flight_prehil_oklistNext = FlightSystemStatus.objects.filter(filter_objects,
                                                                         pre_hil_date__year=year,pre_hil_status='OK(next stage)')
            flight_prehil_observationlist = FlightSystemStatus.objects.filter(filter_objects,
                                                                              pre_hil_date__year=year,pre_hil_status='Observation(same stage)')
            flight_prehil_observationlistNext = FlightSystemStatus.objects.filter(filter_objects,
                                                                                  pre_hil_date__year=year,pre_hil_status='Observation(next stage)')
            flight_prehil_uplist = FlightSystemStatus.objects.filter(filter_objects, pre_hil_date__year=year,pre_hil_status='Under process')
            flight_prehil_haultlist = FlightSystemStatus.objects.filter(filter_objects, pre_hil_date__year=year,pre_hil_status='Halt')

            flight_posthil_count = FlightSystemStatus.objects.filter(setid_filter, post_hil_date__year=year).count()

            flight_posthil_oklist = FlightSystemStatus.objects.filter(filter_objects, post_hil_date__year=year,post_hil_status='Ok')
            flight_posthil_oklistNext = FlightSystemStatus.objects.filter(filter_objects,
                                                                          post_hil_date__year=year,post_hil_status='OK(next stage)')
            flight_posthil_observationlist = FlightSystemStatus.objects.filter(filter_objects,
                                                                               post_hil_date__year=year,post_hil_status='Observation(same stage)')
            flight_posthil_observationlistNext = FlightSystemStatus.objects.filter(filter_objects,
                                                                                   post_hil_date__year=year,post_hil_status='Observation(next stage)')
            flight_posthil_uplist = FlightSystemStatus.objects.filter(filter_objects, post_hil_date__year=year,post_hil_status='Under process')
            flight_posthil_haultlist = FlightSystemStatus.objects.filter(filter_objects, post_hil_date__year=year,post_hil_status='Halt')

            flight_finalintegration_count = FlightSystemStatus.objects.filter(setid_filter, final_integration_date__year=year).count()

            flight_finalintegration_oklist = FlightSystemStatus.objects.filter(filter_objects,
                                                                               final_integration_date__year=year,final_integration_status='Ok')
            flight_finalintegration_oklistNext = FlightSystemStatus.objects.filter(filter_objects,
                                                                                   final_integration_date__year=year,final_integration_status='OK(next stage)')

            flight_finalintegration_observationlist = FlightSystemStatus.objects.filter(filter_objects,
                                                                                        final_integration_date__year=year,final_integration_status='Observation(same stage)')
            flight_finalintegration_observationlistNext = FlightSystemStatus.objects.filter(filter_objects,
                                                                                            final_integration_date__year=year,final_integration_status='Observation(next stage)')
            flight_finalintegration_completelist = FlightSystemStatus.objects.filter(filter_objects,
                                                                                     final_integration_date__year=year,final_integration_status='Under process')
            flight_finalintegration_haultlist = FlightSystemStatus.objects.filter(filter_objects,
                                                                                  final_integration_date__year=year,final_integration_status='Halt')

            flight_vibration_count = FlightSystemStatus.objects.filter(setid_filter, vibaration_date__year=year).count()

            flight_vibration_oklist = FlightSystemStatus.objects.filter(filter_objects,
                                                                        vibaration_date__year=year,vibaration_status='Ok')
            flight_vibration_oklistNext = FlightSystemStatus.objects.filter(filter_objects,
                                                                            vibaration_date__year=year,vibaration_status='OK(next stage)')
            flight_vibration_observationlist = FlightSystemStatus.objects.filter(filter_objects,
                                                                                 vibaration_date__year=year,vibaration_status='Observation(same stage)')
            flight_vibration_observationlistNext = FlightSystemStatus.objects.filter(filter_objects,
                                                                                     vibaration_date__year=year,vibaration_status='Observation(next stage)')
            flight_vibration_uplist = FlightSystemStatus.objects.filter(filter_objects,
                                                                        vibaration_date__year=year,vibaration_status='Under process')
            flight_vibration_haultlist = FlightSystemStatus.objects.filter(filter_objects, vibaration_date__year=year,vibaration_status='Halt')


            flight_cg_count = FlightSystemStatus.objects.filter(setid_filter, cgbalancing_date__year=year).count()

            flight_cg_oklist = FlightSystemStatus.objects.filter(filter_objects,
                                                                 cgbalancing_date__year=year,cgbalancing_date_status='Ok')
            flight_cg_oklistNext = FlightSystemStatus.objects.filter(filter_objects,
                                                                     cgbalancing_date__year=year,cgbalancing_date_status='OK(next stage)')
            flight_cg_observationlist = FlightSystemStatus.objects.filter(filter_objects,
                                                                          cgbalancing_date__year=year,cgbalancing_date_status='Observation(same stage)')
            flight_cg_observationlistNext = FlightSystemStatus.objects.filter(filter_objects,
                                                                              cgbalancing_date__year=year,cgbalancing_date_status='Observation(next stage)')
            flight_cg_uplist = FlightSystemStatus.objects.filter(filter_objects,
                                                                 cgbalancing_date__year=year,cgbalancing_date_status='Under process')
            flight_cg_haultlist = FlightSystemStatus.objects.filter(filter_objects, cgbalancing_date__year=year,cgbalancing_date_status='Halt')

            flight_fgt_count = FlightSystemStatus.objects.filter(setid_filter, fgt_date__year=year).count()

            flight_fgt_oklist = FlightSystemStatus.objects.filter(filter_objects, fgt_date__year=year,fgt_status='Ok')
            flight_fgt_oklistNext = FlightSystemStatus.objects.filter(filter_objects, fgt_date__year=year,fgt_status='OK(next stage)')
            flight_fgt_observationlist = FlightSystemStatus.objects.filter(filter_objects,
                                                                           fgt_date__year=year,fgt_status='Observation(same stage)')
            flight_fgt_observationlistNext = FlightSystemStatus.objects.filter(filter_objects,
                                                                               fgt_date__year=year,fgt_status='Observation(next stage)')
            flight_fgt_uplist = FlightSystemStatus.objects.filter(filter_objects,
                                                                  fgt_date__year=year,fgt_status='Under process')
            flight_fgt_haultlist = FlightSystemStatus.objects.filter(filter_objects, fgt_date__year=year,fgt_status='Halt')

            flight_bhd_count = FlightSystemStatus.objects.filter(setid_filter, bhd_date__year=year).count()

            flight_bhd_not_submit = FlightSystemStatus.objects.filter(filter_objects,
                                                                     bhd_date__year=year,bhd_status='Not Submitted')
            flight_bhd_inprocess = FlightSystemStatus.objects.filter(filter_objects,
                                                                     bhd_date__year=year,bhd_status='Audit in-process')
            flight_bhd_qm_forwarded = FlightSystemStatus.objects.filter(filter_objects,
                                                                       bhd_date__year=year,bhd_status='QM Observations Forwarded')

            flight_bhd_qm_repeated = FlightSystemStatus.objects.filter(filter_objects,
                                                                       bhd_date__year=year,bhd_status='QM Observations Repeated')

            flight_bhd_ok = FlightSystemStatus.objects.filter(filter_objects,
                                                                       bhd_date__year=year,bhd_status='Ok')
            flight_bhd_submitted = FlightSystemStatus.objects.filter(filter_objects,
                                                                       bhd_date__year=year,bhd_status='Submitted')

            flight_fqm_count = FlightSystemStatus.objects.filter(setid_filter, fqm_date__year=year).count()

            flight_fqm_planned = FlightSystemStatus.objects.filter(filter_objects,
                                                                   fqm_date__year=year,fqm_status='Planned')

            flight_fqm_conducted = FlightSystemStatus.objects.filter(filter_objects,
                                                                     fqm_date__year=year,fqm_status='Conducted')

            flight_qmc_count = FlightSystemStatus.objects.filter(setid_filter, qm_certification_date__year=year).count()

            flight_qmc_issued = FlightSystemStatus.objects.filter(filter_objects,
                                                                  qm_certification_date__year=year,qm_certification_status='QM certificate issued')

            flight_qmc_in_process = FlightSystemStatus.objects.filter(filter_objects,
                                                                      qm_certification_date__year=year,qm_certification_status='Audit in-process')
            flight_qmc_obs_forwarded = FlightSystemStatus.objects.filter(filter_objects,
                                                                        qm_certification_date__year=year,qm_certification_status='QM Observations Forwarded')


            flight_launch_count = FlightSystemStatus.objects.filter(setid_filter, launchact_date__year=year).count()

            flight_launch_oklist = FlightSystemStatus.objects.filter(filter_objects, launchact_date__year=year,launchact_status='Ok')
            flight_launch_oklistNext = FlightSystemStatus.objects.filter(filter_objects,
                                                                         launchact_date__year=year,launchact_status='OK(next stage)')
            flight_launch_observationlist = FlightSystemStatus.objects.filter(filter_objects,
                                                                              launchact_date__year=year,launchact_status='Observation')
            flight_launch_observationlistNext = FlightSystemStatus.objects.filter(filter_objects,
                                                                                  launchact_date__year=year,launchact_status='Observation(next stage)')
            flight_launch_uplist = FlightSystemStatus.objects.filter(filter_objects,
                                                                     launchact_date__year=year,launchact_status='Completed')
            flight_launch_haultlist = FlightSystemStatus.objects.filter(filter_objects, launchact_date__year=year,launchact_status='Halt')


            flight_sys_align_count = FlightSystemStatus.objects.filter(setid_filter, sys_align_Date__year=year).count()

            flight_sys_align_oklist = FlightSystemStatus.objects.filter(filter_objects, sys_align_Date__year=year,sys_align_status='Ok')
            flight_sys_align_oklistNext = FlightSystemStatus.objects.filter(filter_objects,
                                                                         sys_align_Date__year=year,sys_align_status='OK(next stage)')
            flight_sys_align_observationlist = FlightSystemStatus.objects.filter(filter_objects,
                                                                              sys_align_Date__year=year,sys_align_status='Observation(same stage)')
            flight_sys_align_observationlistNext = FlightSystemStatus.objects.filter(filter_objects,
                                                                                  sys_align_Date__year=year,sys_align_status='Observation(next stage)')
            flight_sys_align_uplist = FlightSystemStatus.objects.filter(filter_objects,
                                                                     sys_align_Date__year=year,sys_align_status='Under process')
            flight_sys_align_haultlist = FlightSystemStatus.objects.filter(filter_objects, sys_align_Date__year=year,sys_align_status='Halt')


            flight_incapsulation_count = FlightSystemStatus.objects.filter(setid_filter, incapsulation_date__year=year).count()

            flight_incapsulation_oklist = FlightSystemStatus.objects.filter(filter_objects, incapsulation_date__year=year,incapsulation_status='Ok')
            flight_incapsulation_oklistNext = FlightSystemStatus.objects.filter(filter_objects,
                                                                         incapsulation_date__year=year,incapsulation_status='OK(next stage)')
            flight_incapsulation_observationlist = FlightSystemStatus.objects.filter(filter_objects,
                                                                              incapsulation_date__year=year,incapsulation_status='Observation(same stage)')
            flight_incapsulation_observationlistNext = FlightSystemStatus.objects.filter(filter_objects,
                                                                                  incapsulation_date__year=year,incapsulation_status='Observation(next stage)')
            flight_incapsulation_uplist = FlightSystemStatus.objects.filter(filter_objects,
                                                                     incapsulation_date__year=year,incapsulation_status='Under process')
            flight_incapsulation_haultlist = FlightSystemStatus.objects.filter(filter_objects, incapsulation_date__year=year,incapsulation_status='Halt')

            flight_emp_proofing_count = FlightSystemStatus.objects.filter(setid_filter, emp_proofing_date__year=year).count()

            flight_emp_proofing_oklist = FlightSystemStatus.objects.filter(filter_objects,emp_proofing_date__year=year,emp_proofing='Ok')
            flight_emp_proofing_oklistNext= FlightSystemStatus.objects.filter(filter_objects,emp_proofing_date__year=year,emp_proofing='OK(next stage)')
            flight_emp_proofing_observationlist  = FlightSystemStatus.objects.filter(filter_objects,emp_proofing_date__year=year,emp_proofing='Observation(same stage)')
            flight_emp_proofing_observationlistNext = FlightSystemStatus.objects.filter(filter_objects,emp_proofing_date__year=year,emp_proofing='Observation(next stage)')
            flight_emp_proofing_uplist = FlightSystemStatus.objects.filter(filter_objects,emp_proofing_date__year=year,emp_proofing='Under process')
            flight_emp_proofing_haultlist = FlightSystemStatus.objects.filter(filter_objects,emp_proofing_date__year=year,emp_proofing='Halt')

            flight_func_tst_count = FlightSystemStatus.objects.filter(setid_filter, func_tst_date__year=year).count()

            flight_func_tst_oklistNext_oklist  = FlightSystemStatus.objects.filter(filter_objects,func_tst_date__year=year,func_tst='Ok')
            flight_func_tst_oklistNext_oklistNext  = FlightSystemStatus.objects.filter(filter_objects,func_tst_date__year=year,func_tst='OK(next stage)')
            flight_func_tst_oklistNext_observationlist   = FlightSystemStatus.objects.filter(filter_objects,func_tst_date__year=year,func_tst='Observation(same stage)')
            flight_func_tst_oklistNext_observationlistNext  = FlightSystemStatus.objects.filter(filter_objects,func_tst_date__year=year,func_tst='Observation(next stage)')
            flight_func_tst_oklistNext_uplist  = FlightSystemStatus.objects.filter(filter_objects,func_tst_date__year=year,func_tst='Under process')
            flight_func_tst_oklistNext_haultlist  = FlightSystemStatus.objects.filter(filter_objects,func_tst_date__year=year,func_tst='Halt')

            flight_func_tst_dummy_bird_count = FlightSystemStatus.objects.filter(setid_filter, func_tst_dummy_bird_date__year=year).count()

            flight_func_tst_dummy_bird_oklist = FlightSystemStatus.objects.filter(filter_objects,func_tst_dummy_bird_date__year=year,func_tst_dummy_bird='Ok')
            flight_func_tst_dummy_bird_oklistNext = FlightSystemStatus.objects.filter(filter_objects,func_tst_dummy_bird_date__year=year,func_tst_dummy_bird='OK(next stage)')
            flight_func_tst_dummy_bird_observationlist  = FlightSystemStatus.objects.filter(filter_objects,func_tst_dummy_bird_date__year=year,func_tst_dummy_bird='Observation(same stage)')
            flight_func_tst_dummy_bird_observationlistNext = FlightSystemStatus.objects.filter(filter_objects,func_tst_dummy_bird_date__year=year,func_tst_dummy_bird='Observation(next stage)')
            flight_func_tst_dummy_bird_uplist = FlightSystemStatus.objects.filter(filter_objects,func_tst_dummy_bird_date__year=year,func_tst_dummy_bird='Under process')
            flight_func_tst_dummy_bird_haultlist = FlightSystemStatus.objects.filter(filter_objects,func_tst_dummy_bird_date__year=year,func_tst_dummy_bird='Halt')

            flight_road_test_count = FlightSystemStatus.objects.filter(setid_filter, road_test_date__year=year).count()

            flight_road_test_oklist = FlightSystemStatus.objects.filter(filter_objects,road_test_date__year=year,road_test='Ok')
            flight_road_test_oklistNext = FlightSystemStatus.objects.filter(filter_objects,road_test_date__year=year,road_test='OK(next stage)')
            flight_road_test_observationlist  = FlightSystemStatus.objects.filter(filter_objects,road_test_date__year=year,road_test='Observation(same stage)')
            flight_road_test_observationlistNext = FlightSystemStatus.objects.filter(filter_objects,road_test_date__year=year,road_test='Observation(next stage)')
            flight_road_test_uplist = FlightSystemStatus.objects.filter(filter_objects,road_test_date__year=year,road_test='Under process')
            flight_road_test_haultlist = FlightSystemStatus.objects.filter(filter_objects,road_test_date__year=year,road_test='Halt')


            flight_post_road_test_count = FlightSystemStatus.objects.filter(setid_filter, post_road_test_date__year=year).count()

            flight_post_road_test_oklist = FlightSystemStatus.objects.filter(filter_objects,post_road_test_date__year=year,post_road_test='Ok')
            flight_post_road_test_oklistNext = FlightSystemStatus.objects.filter(filter_objects,post_road_test_date__year=year,post_road_test='OK(next stage)')
            flight_post_road_test_observationlist  = FlightSystemStatus.objects.filter(filter_objects,post_road_test_date__year=year,post_road_test='Observation(same stage)')
            flight_post_road_test_observationlistNext = FlightSystemStatus.objects.filter(filter_objects,post_road_test_date__year=year,post_road_test='Observation(next stage)')
            flight_post_road_test_uplist = FlightSystemStatus.objects.filter(filter_objects,post_road_test_date__year=year,post_road_test='Under process')
            flight_post_road_test_haultlist = FlightSystemStatus.objects.filter(filter_objects,post_road_test_date__year=year,post_road_test='Halt')


            flight_integrated_operation_count = FlightSystemStatus.objects.filter(setid_filter, integrated_operation_date__year=year).count()

            flight_integrated_operation_oklist = FlightSystemStatus.objects.filter(filter_objects,integrated_operation_date__year=year,integrated_operation='Ok')
            flight_integrated_operation_oklistNext = FlightSystemStatus.objects.filter(filter_objects,integrated_operation_date__year=year,integrated_operation='OK(next stage)')
            flight_integrated_operation_observationlist  = FlightSystemStatus.objects.filter(filter_objects,integrated_operation_date__year=year,integrated_operation='Observation(same stage)')
            flight_integrated_operation_observationlistNext = FlightSystemStatus.objects.filter(filter_objects,integrated_operation_date__year=year,integrated_operation='Observation(next stage)')
            flight_integrated_operation_uplist = FlightSystemStatus.objects.filter(filter_objects,integrated_operation_date__year=year,integrated_operation='Under process')
            flight_integrated_operation_haultlist = FlightSystemStatus.objects.filter(filter_objects,integrated_operation_date__year=year,integrated_operation='Halt')


            flight_rain_test_count = FlightSystemStatus.objects.filter(setid_filter, rain_test_date__year=year).count()

            flight_rain_test_oklist = FlightSystemStatus.objects.filter(filter_objects,rain_test_date__year=year,rain_test='Ok')
            flight_rain_test_oklistNext = FlightSystemStatus.objects.filter(filter_objects,rain_test_date__year=year,rain_test='OK(next stage)')
            flight_rain_test_observationlist  = FlightSystemStatus.objects.filter(filter_objects,rain_test_date__year=year,rain_test='Observation(same stage)')
            flight_rain_test_observationlistNext = FlightSystemStatus.objects.filter(filter_objects,rain_test_date__year=year,rain_test='Observation(next stage)')
            flight_rain_test_uplist = FlightSystemStatus.objects.filter(filter_objects,rain_test_date__year=year,rain_test='Under process')
            flight_rain_test_haultlist = FlightSystemStatus.objects.filter(filter_objects,rain_test_date__year=year,rain_test='Halt')


            flight_pre_user_inspection_count = FlightSystemStatus.objects.filter(setid_filter, pre_user_inspection_date__year=year).count()

            flight_pre_user_inspection_oklist = FlightSystemStatus.objects.filter(filter_objects,pre_user_inspection_date__year=year,pre_user_inspection='Ok')
            flight_pre_user_inspection_oklistNext = FlightSystemStatus.objects.filter(filter_objects,pre_user_inspection_date__year=year,pre_user_inspection='OK(next stage)')
            flight_pre_user_inspection_observationlist  = FlightSystemStatus.objects.filter(filter_objects,pre_user_inspection_date__year=year,pre_user_inspection='Observation(same stage)')
            flight_pre_user_inspection_observationlistNext = FlightSystemStatus.objects.filter(filter_objects,pre_user_inspection_date__year=year,pre_user_inspection='Observation(next stage)')
            flight_pre_user_inspection_uplist = FlightSystemStatus.objects.filter(filter_objects,pre_user_inspection_date__year=year,pre_user_inspection='Under process')
            flight_pre_user_inspection_haultlist = FlightSystemStatus.objects.filter(filter_objects,pre_user_inspection_date__year=year,pre_user_inspection='Halt')


            flight_final_integrated_testing_count = FlightSystemStatus.objects.filter(setid_filter, final_integrated_testing_date__year=year).count()

            flight_final_integrated_testing_oklist = FlightSystemStatus.objects.filter(filter_objects,final_integrated_testing_date__year=year,final_integrated_testing='Ok')
            flight_final_integrated_testing_oklistNext = FlightSystemStatus.objects.filter(filter_objects,final_integrated_testing_date__year=year,final_integrated_testing='OK(next stage)')
            flight_final_integrated_testing_observationlist  = FlightSystemStatus.objects.filter(filter_objects,final_integrated_testing_date__year=year,final_integrated_testing='Observation(same stage)')
            flight_final_integrated_testing_observationlistNext = FlightSystemStatus.objects.filter(filter_objects,final_integrated_testing_date__year=year,final_integrated_testing='Observation(next stage)')
            flight_final_integrated_testing_uplist = FlightSystemStatus.objects.filter(filter_objects,final_integrated_testing_date__year=year,final_integrated_testing='Under process')
            flight_final_integrated_testing_haultlist = FlightSystemStatus.objects.filter(filter_objects,final_integrated_testing_date__year=year,final_integrated_testing='Halt')


            flight_load_unload_on_mlv_hlf_count = FlightSystemStatus.objects.filter(setid_filter, load_unload_on_mlv_hlf_date__year=year).count()

            flight_load_unload_on_mlv_hlf_oklist = FlightSystemStatus.objects.filter(filter_objects,load_unload_on_mlv_hlf_date__year=year,load_unload_on_mlv_hlf='Ok')
            flight_load_unload_on_mlv_hlf_oklistNext = FlightSystemStatus.objects.filter(filter_objects,load_unload_on_mlv_hlf_date__year=year,load_unload_on_mlv_hlf='OK(next stage)')
            flight_load_unload_on_mlv_hlf_observationlist  = FlightSystemStatus.objects.filter(filter_objects,load_unload_on_mlv_hlf_date__year=year,load_unload_on_mlv_hlf='Observation(same stage)')
            flight_load_unload_on_mlv_hlf_observationlistNext = FlightSystemStatus.objects.filter(filter_objects,load_unload_on_mlv_hlf_date__year=year,load_unload_on_mlv_hlf='Observation(next stage)')
            flight_load_unload_on_mlv_hlf_uplist = FlightSystemStatus.objects.filter(filter_objects,load_unload_on_mlv_hlf_date__year=year,load_unload_on_mlv_hlf='Under process')
            flight_load_unload_on_mlv_hlf_haultlist = FlightSystemStatus.objects.filter(filter_objects,load_unload_on_mlv_hlf_date__year=year,load_unload_on_mlv_hlf='Halt')

            # flightTotalCount = flight_blt_count+flight_prehil_count+flight_posthil_count+flight_vibration_count+flight_cg_count+flight_fgt_count+flight_bhd_count+flight_qmc_count+flight_fqm_count+flight_sys_align_count+flight_incapsulation_count+flight_launch_count+flight_finalintegration_count+flight_emp_proofing_count+flight_road_test_count+flight_post_road_test_count+flight_final_integrated_testing_count+flight_rain_test_count+flight_pre_user_inspection_count+flight_func_tst_count+flight_func_tst_dummy_bird_count+flight_integrated_operation_count+flight_load_unload_on_mlv_hlf_count
            # relifing system count

            # relifing_blt_count = RelifingSystemStatus.objects.filter(setid_filter, blt_date__year=year).count()
            totCountRelif = RelifingSystemStatus.objects.filter(setid_filter)
            # totCount = ProductionSystemStatus.objects.filter(setid_filter)
            relif_sys_count = 0
            for set in totCountRelif:
                count_id = 0
                if set.blt_status != 'None' and set.set_id != '' and set.blt_date is not None:
                    if count_id == 0 and set.blt_date.strftime("%Y") == year:
                        count_id = set.id
                        relif_sys_count += 1
                if set.pre_hil_status != 'None' and set.set_id != '' and set.pre_hil_date is not None:
                    if count_id == 0 and set.pre_hil_date.strftime("%Y") == year:
                        count_id = set.id
                        relif_sys_count += 1
                if set.vibaration_status != 'None' and set.set_id != '' and set.vibaration_date is not None:
                    if count_id == 0 and set.vibaration_date.strftime("%Y") == year:
                        count_id = set.id
                        relif_sys_count += 1
                if set.post_hil_status != 'None' and set.set_id != '' and set.post_hil_date is not None:
                    if count_id == 0 and set.post_hil_date.strftime("%Y") == year:
                        count_id = set.id
                        relif_sys_count += 1
                if set.fgt_status != 'None' and set.set_id != '' and set.fgt_date is not None:
                    if count_id == 0 and set.fgt_date.strftime("%Y") == year:
                        count_id = set.id
                        relif_sys_count += 1

                if set.final_integration_status != 'None' and set.set_id != '' and set.final_integration_date is not None:
                    if count_id == 0 and set.final_integration_date.strftime("%Y") == year:
                        count_id = set.id
                        relif_sys_count += 1
                if set.bhd_status != 'None' and set.set_id != '' and set.bhd_date is not None:
                    if count_id == 0 and set.bhd_date.strftime("%Y") == year:
                        count_id = set.id
                        relif_sys_count += 1
                if set.fqm_status != 'None' and set.set_id != '' and set.fqm_date is not None:
                    if count_id == 0 and set.fqm_date.strftime("%Y") == year:
                        count_id = set.id
                        relif_sys_count += 1
                if set.qm_certification_status != 'None' and set.set_id != '' and set.qm_certification_date is not None:
                    if count_id == 0 and set.qm_certification_date.strftime("%Y") == year:
                        count_id = set.id
                        relif_sys_count += 1
                if set.cgbalancing_date_status != 'None' and set.set_id != '' and set.cgbalancing_date is not None:
                    if count_id == 0 and set.cgbalancing_date.strftime("%Y") == year:
                        count_id = set.id
                        relif_sys_count += 1

                if set.enduser_status != 'None' and set.set_id != '' and set.enduser_date is not None:
                    if count_id == 0 and set.enduser_date.strftime("%Y") == year:
                        count_id = set.id
                        relif_sys_count += 1
                if set.incapsulation_status != 'None' and set.set_id != '' and set.incapsulation_date is not None:
                    if count_id == 0 and set.incapsulation_date.strftime("%Y") == year:
                        count_id = set.id
                        relif_sys_count += 1
                if set.sys_align_status != 'None' and set.set_id != '' and set.sys_align_Date is not None:
                    if count_id == 0 and set.sys_align_Date.strftime("%Y") == year:
                        count_id = set.id
                        relif_sys_count += 1
                if set.emp_proofing != 'None' and set.set_id != '' and set.emp_proofing_date is not None:
                    if count_id == 0 and set.emp_proofing_date.strftime("%Y") == year:
                        count_id = set.id
                        relif_sys_count += 1
                if set.func_tst != 'None' and set.set_id != '' and set.func_tst_date is not None:
                    if count_id == 0 and set.func_tst_date.strftime("%Y") == year:
                        count_id = set.id
                        relif_sys_count += 1

                if set.func_tst_dummy_bird != 'None' and set.set_id != '' and set.func_tst_dummy_bird_date is not None:
                    if count_id == 0 and set.func_tst_dummy_bird_date.strftime("%Y") == year:
                        count_id = set.id
                        relif_sys_count += 1
                if set.road_test != 'None' and set.set_id != '' and set.road_test_date is not None:
                    if count_id == 0 and set.road_test_date.strftime("%Y") == year:
                        count_id = set.id
                        relif_sys_count += 1
                if set.post_road_test != 'None' and set.set_id != '' and set.post_road_test_date is not None:
                    if count_id == 0 and set.post_road_test_date.strftime("%Y") == year:
                        count_id = set.id
                        relif_sys_count += 1
                if set.integrated_operation != 'None' and set.set_id != '' and set.integrated_operation_date is not None:
                    if count_id == 0 and set.integrated_operation_date.strftime("%Y") == year:
                        count_id = set.id
                        relif_sys_count += 1
                if set.rain_test != 'None' and set.set_id != '' and set.rain_test_date is not None:
                    if count_id == 0 and set.rain_test_date.strftime("%Y") == year:
                        count_id = set.id
                        relif_sys_count += 1

                if set.pre_user_inspection != 'None' and set.set_id != '' and set.pre_user_inspection_date is not None:
                    if count_id == 0 and set.pre_user_inspection_date.strftime("%Y") == year:
                        count_id = set.id
                        relif_sys_count += 1
                if set.final_integrated_testing != 'None' and set.set_id != '' and set.final_integrated_testing_date is not None:
                    if count_id == 0 and set.final_integrated_testing_date.strftime("%Y") == year:
                        count_id = set.id
                        relif_sys_count += 1
                if set.load_unload_on_mlv_hlf != 'None' and set.set_id != '' and set.load_unload_on_mlv_hlf_date is not None:
                    if count_id == 0 and set.load_unload_on_mlv_hlf_date.strftime("%Y") == year:
                        count_id = set.id
                        relif_sys_count += 1
            relifTotalCount = relif_sys_count
            relifing_blt_oklist = RelifingSystemStatus.objects.filter(filter_objects, blt_date__year=year,blt_status='Ok')
            relifing_blt_oklistNext = RelifingSystemStatus.objects.filter(filter_objects, blt_date__year=year,blt_status='OK(next stage)')
            relifing_blt_observationlist = RelifingSystemStatus.objects.filter(filter_objects,
                                                                               blt_date__year=year,blt_status='Observation(same stage)')
            relifing_blt_observationlistNext = RelifingSystemStatus.objects.filter(filter_objects,
                                                                                   blt_date__year=year,blt_status='Observation(next stage)')
            relifing_blt_uplist = RelifingSystemStatus.objects.filter(filter_objects, blt_date__year=year,blt_status='Under process')
            relifing_blt_haultlist = RelifingSystemStatus.objects.filter(filter_objects, blt_date__year=year,blt_status='Halt')


            relifing_prehil_count = RelifingSystemStatus.objects.filter(setid_filter, pre_hil_date__year=year).count()

            relifing_prehil_oklist = RelifingSystemStatus.objects.filter(filter_objects,
                                                                         pre_hil_date__year=year,pre_hil_status='Ok')
            relifing_prehil_oklistNext = RelifingSystemStatus.objects.filter(filter_objects,
                                                                             pre_hil_date__year=year,pre_hil_status='OK(next stage)')
            relifing_prehil_observationlist = RelifingSystemStatus.objects.filter(filter_objects,
                                                                                  pre_hil_date__year=year,pre_hil_status='Observation(same stage)')
            relifing_prehil_observationlistNext = RelifingSystemStatus.objects.filter(filter_objects,
                                                                                      pre_hil_date__year=year,pre_hil_status='Observation(next stage)')
            relifing_prehil_uplist = RelifingSystemStatus.objects.filter(filter_objects, pre_hil_date__year=year,pre_hil_status='Under process')
            relifing_prehil_haultlist = RelifingSystemStatus.objects.filter(filter_objects, pre_hil_date__year=year,pre_hil_status='Halt')


            relifing_posthil_count = RelifingSystemStatus.objects.filter(setid_filter, post_hil_date__year=year).count()

            relifing_posthil_oklist = RelifingSystemStatus.objects.filter(filter_objects,
                                                                          post_hil_date__year=year,post_hil_status='Ok')
            relifing_posthil_oklistNext = RelifingSystemStatus.objects.filter(filter_objects,
                                                                              post_hil_date__year=year,post_hil_status='OK(next stage)')
            relifing_posthil_observationlist = RelifingSystemStatus.objects.filter(filter_objects,
                                                                                   post_hil_date__year=year,post_hil_status='Observation(same stage)')
            relifing_posthil_observationlistNext = RelifingSystemStatus.objects.filter(filter_objects,
                                                                                       post_hil_date__year=year,post_hil_status='Observation(next stage)')
            relifing_posthil_uplist = RelifingSystemStatus.objects.filter(filter_objects,
                                                                          post_hil_date__year=year,post_hil_status='Under process')
            relifing_posthil_haultlist = RelifingSystemStatus.objects.filter(filter_objects, post_hil_date__year=year,post_hil_status='Halt')


            relifing_finalintegration_count = RelifingSystemStatus.objects.filter(setid_filter, final_integration_date__year=year).count()

            relifing_finalintegration_oklist = RelifingSystemStatus.objects.filter(filter_objects,
                                                                                   final_integration_date__year=year,final_integration_status='Ok')
            relifing_finalintegration_oklistNext = RelifingSystemStatus.objects.filter(filter_objects,
                                                                                       final_integration_date__year=year,final_integration_status='OK(next stage)')
            relifing_finalintegration_observationlist = RelifingSystemStatus.objects.filter(filter_objects,
                                                                                            final_integration_date__year=year,final_integration_status='Observation(same stage)')
            relifing_finalintegration_observationlistNext = RelifingSystemStatus.objects.filter(filter_objects,
                                                                                                final_integration_date__year=year,final_integration_status='Observation(next stage)')
            relifing_finalintegration_completelist = RelifingSystemStatus.objects.filter(filter_objects,
                                                                                         final_integration_date__year=year,final_integration_status='Under process')
            relifing_finalintegration_haultlist = RelifingSystemStatus.objects.filter(filter_objects,
                                                                                      final_integration_date__year=year,final_integration_status='Halt')


            relifing_vibration_count = RelifingSystemStatus.objects.filter(setid_filter, vibaration_date__year=year).count()

            relifing_vibration_oklist = RelifingSystemStatus.objects.filter(filter_objects,
                                                                            vibaration_date__year=year,vibaration_status='Ok')
            relifing_vibration_oklistNext = RelifingSystemStatus.objects.filter(filter_objects,
                                                                                vibaration_date__year=year,vibaration_status='OK(next stage)')
            relifing_vibration_observationlist = RelifingSystemStatus.objects.filter(filter_objects,
                                                                                     vibaration_date__year=year,vibaration_status='Observation(same stage)')
            relifing_vibration_observationlistNext = RelifingSystemStatus.objects.filter(filter_objects,
                                                                                         vibaration_date__year=year,vibaration_status='Observation(next stage)')
            relifing_vibration_uplist = RelifingSystemStatus.objects.filter(filter_objects,
                                                                            vibaration_date__year=year,vibaration_status='Under process')
            relifing_vibration_haultlist = RelifingSystemStatus.objects.filter(filter_objects, vibaration_date__year=year,vibaration_status='Halt')


            relifing_cg_count = RelifingSystemStatus.objects.filter(setid_filter, cgbalancing_date__year=year).count()

            relifing_cg_oklist = RelifingSystemStatus.objects.filter(filter_objects,
                                                                     cgbalancing_date__year=year,cgbalancing_date_status='Ok')
            relifing_cg_oklistNext = RelifingSystemStatus.objects.filter(filter_objects,
                                                                         cgbalancing_date__year=year,cgbalancing_date_status='OK(next stage)')
            relifing_cg_observationlist = RelifingSystemStatus.objects.filter(filter_objects,
                                                                              cgbalancing_date__year=year,cgbalancing_date_status='Observation(same stage)')
            relifing_cg_observationlistNext = RelifingSystemStatus.objects.filter(filter_objects,
                                                                                  cgbalancing_date__year=year,cgbalancing_date_status='Observation(next stage)')
            relifing_cg_uplist = RelifingSystemStatus.objects.filter(filter_objects,
                                                                     cgbalancing_date__year=year,cgbalancing_date_status='Under process')
            relifing_cg_haultlist = RelifingSystemStatus.objects.filter(filter_objects, cgbalancing_date__year=year,cgbalancing_date_status='Halt')


            relifing_fgt_count = RelifingSystemStatus.objects.filter(setid_filter, fgt_date__year=year).count()

            relifing_fgt_oklist = RelifingSystemStatus.objects.filter(filter_objects, fgt_date__year=year,fgt_status='Ok')
            relifing_fgt_oklistNext = RelifingSystemStatus.objects.filter(filter_objects, fgt_date__year=year,fgt_status='OK(next stage)')
            relifing_fgt_observationlist = RelifingSystemStatus.objects.filter(filter_objects,
                                                                               fgt_date__year=year,fgt_status='Observation(same stage)')
            relifing_fgt_observationlistNext = RelifingSystemStatus.objects.filter(filter_objects,
                                                                                   fgt_date__year=year,fgt_status='Observation(next stage)')
            relifing_fgt_uplist = RelifingSystemStatus.objects.filter(filter_objects,
                                                                      fgt_date__year=year,fgt_status='Under process')
            relifing_fgt_haultlist = RelifingSystemStatus.objects.filter(filter_objects, fgt_date__year=year,fgt_status='Halt')


            relifing_bhd_count = RelifingSystemStatus.objects.filter(setid_filter, bhd_date__year=year).count()

            relifing_bhd_not_submit = RelifingSystemStatus.objects.filter(filter_objects,
                                                                         bhd_date__year=year,bhd_status='Not Submitted')
            relifing_bhd_inprocess = RelifingSystemStatus.objects.filter(filter_objects,
                                                                         bhd_date__year=year,bhd_status='Audit in-process')
            relifing_bhd_qm_forwarded = RelifingSystemStatus.objects.filter(filter_objects,
                                                                           bhd_date__year=year,bhd_status='QM Observations Forwarded')

            relifing_bhd_qm_repeated = RelifingSystemStatus.objects.filter(filter_objects,
                                                                           bhd_date__year=year,bhd_status='QM Observations Repeated')

            relifing_bhd_ok = RelifingSystemStatus.objects.filter(filter_objects,
                                                                       bhd_date__year=year,bhd_status='Ok')
            relifing_bhd_submitted = RelifingSystemStatus.objects.filter(filter_objects,
                                                                       bhd_date__year=year,bhd_status='Submitted')


            relifing_fqm_count = RelifingSystemStatus.objects.filter(setid_filter, fqm_date__year=year).count()

            relifing_fqm_planned = RelifingSystemStatus.objects.filter(filter_objects,
                                                                       fqm_date__year=year,fqm_status='Planned')

            relifing_fqm_conducted = RelifingSystemStatus.objects.filter(filter_objects,
                                                                         fqm_date__year=year,fqm_status='Conducted')


            relifing_qmc_count = RelifingSystemStatus.objects.filter(setid_filter, qm_certification_date__year=year).count()

            relifing_qmc_issued = RelifingSystemStatus.objects.filter(filter_objects,
                                                                  qm_certification_date__year=year,qm_certification_status='QM certificate issued')

            relifing_qmc_in_process = RelifingSystemStatus.objects.filter(filter_objects,
                                                                      qm_certification_date__year=year,qm_certification_status='Audit in-process')
            relifing_qmc_obs_forwarded = RelifingSystemStatus.objects.filter(filter_objects,
                                                                        qm_certification_date__year=year,qm_certification_status='QM Observations Forwarded')


            relifing_enduser_count = RelifingSystemStatus.objects.filter(setid_filter, enduser_date__year=year).count()

            relifing_enduser_oklist = RelifingSystemStatus.objects.filter(filter_objects,
                                                                          enduser_date__year=year,enduser_status='Ok')
            relifing_enduser_oklistNext = RelifingSystemStatus.objects.filter(filter_objects,
                                                                              enduser_date__year=year,enduser_status='OK(next stage)')
            relifing_enduser_observationlist = RelifingSystemStatus.objects.filter(filter_objects,
                                                                                   enduser_date__year=year,enduser_status='Observation')
            relifing_enduser_observationlistNext = RelifingSystemStatus.objects.filter(filter_objects,
                                                                                       enduser_date__year=year,enduser_status='Observation(next stage)')
            relifing_enduser_uplist = RelifingSystemStatus.objects.filter(filter_objects,
                                                                          enduser_date__year=year,enduser_status='Completed')
            relifing_enduser_haultlist = RelifingSystemStatus.objects.filter(filter_objects, enduser_date__year=year,enduser_status='Halt')



            relifing_sys_align_count = RelifingSystemStatus.objects.filter(setid_filter, sys_align_Date__year=year).count()

            relifing_sys_align_oklist = RelifingSystemStatus.objects.filter(filter_objects,
                                                                          sys_align_Date__year=year,sys_align_status='Ok')
            relifing_sys_align_oklistNext = RelifingSystemStatus.objects.filter(filter_objects,
                                                                              sys_align_Date__year=year,sys_align_status='OK(next stage)')
            relifing_sys_align_observationlist = RelifingSystemStatus.objects.filter(filter_objects,
                                                                                   sys_align_Date__year=year,sys_align_status='Observation(same stage)')
            relifing_sys_align_observationlistNext = RelifingSystemStatus.objects.filter(filter_objects,
                                                                                       sys_align_Date__year=year,sys_align_status='Observation(next stage)')
            relifing_sys_align_uplist = RelifingSystemStatus.objects.filter(filter_objects,
                                                                          sys_align_Date__year=year,sys_align_status='Under process')
            relifing_sys_align_haultlist = RelifingSystemStatus.objects.filter(filter_objects, sys_align_Date__year=year,sys_align_status='Halt')


            relifing_incapsulation_count = RelifingSystemStatus.objects.filter(setid_filter, incapsulation_date__year=year).count()

            relifing_incapsulation_oklist = RelifingSystemStatus.objects.filter(filter_objects,
                                                                          incapsulation_date__year=year,incapsulation_status='Ok')
            relifing_incapsulation_oklistNext = RelifingSystemStatus.objects.filter(filter_objects,
                                                                              incapsulation_date__year=year,incapsulation_status='OK(next stage)')
            relifing_incapsulation_observationlist = RelifingSystemStatus.objects.filter(filter_objects,
                                                                                   incapsulation_date__year=year,incapsulation_status='Observation(same stage)')
            relifing_incapsulation_observationlistNext = RelifingSystemStatus.objects.filter(filter_objects,
                                                                                       incapsulation_date__year=year,incapsulation_status='Observation(next stage)')
            relifing_incapsulation_uplist = RelifingSystemStatus.objects.filter(filter_objects,
                                                                          incapsulation_date__year=year,incapsulation_status='Under process')
            relifing_incapsulation_haultlist = RelifingSystemStatus.objects.filter(filter_objects, incapsulation_date__year=year,incapsulation_status='Halt')


            relifing_emp_proofing_count = RelifingSystemStatus.objects.filter(setid_filter, emp_proofing_date__year=year).count()

            relifing_emp_proofing_oklist = RelifingSystemStatus.objects.filter(filter_objects,emp_proofing_date__year=year,emp_proofing='Ok')
            relifing_emp_proofing_oklistNext= RelifingSystemStatus.objects.filter(filter_objects,emp_proofing_date__year=year,emp_proofing='OK(next stage)')
            relifing_emp_proofing_observationlist  = RelifingSystemStatus.objects.filter(filter_objects,emp_proofing_date__year=year,emp_proofing='Observation(same stage)')
            relifing_emp_proofing_observationlistNext = RelifingSystemStatus.objects.filter(filter_objects,emp_proofing_date__year=year,emp_proofing='Observation(next stage)')
            relifing_emp_proofing_uplist = RelifingSystemStatus.objects.filter(filter_objects,emp_proofing_date__year=year,emp_proofing='Under process')
            relifing_emp_proofing_haultlist = RelifingSystemStatus.objects.filter(filter_objects,emp_proofing_date__year=year,emp_proofing='Halt')


            relifing_func_tst_count = RelifingSystemStatus.objects.filter(setid_filter, func_tst_date__year=year).count()

            relifing_func_tst_oklistNext_oklist  = RelifingSystemStatus.objects.filter(filter_objects,func_tst_date__year=year,func_tst='Ok')
            relifing_func_tst_oklistNext_oklistNext  = RelifingSystemStatus.objects.filter(filter_objects,func_tst_date__year=year,func_tst='OK(next stage)')
            relifing_func_tst_oklistNext_observationlist   = RelifingSystemStatus.objects.filter(filter_objects,func_tst_date__year=year,func_tst='Observation(same stage)')
            relifing_func_tst_oklistNext_observationlistNext  = RelifingSystemStatus.objects.filter(filter_objects,func_tst_date__year=year,func_tst='Observation(next stage)')
            relifing_func_tst_oklistNext_uplist  = RelifingSystemStatus.objects.filter(filter_objects,func_tst_date__year=year,func_tst='Under process')
            relifing_func_tst_oklistNext_haultlist  = RelifingSystemStatus.objects.filter(filter_objects,func_tst_date__year=year,func_tst='Halt')


            relifing_func_tst_dummy_bird_count = RelifingSystemStatus.objects.filter(setid_filter, func_tst_dummy_bird_date__year=year).count()

            relifing_func_tst_dummy_bird_oklist = RelifingSystemStatus.objects.filter(filter_objects,func_tst_dummy_bird_date__year=year,func_tst_dummy_bird='Ok')
            relifing_func_tst_dummy_bird_oklistNext = RelifingSystemStatus.objects.filter(filter_objects,func_tst_dummy_bird_date__year=year,func_tst_dummy_bird='OK(next stage)')
            relifing_func_tst_dummy_bird_observationlist  = RelifingSystemStatus.objects.filter(filter_objects,func_tst_dummy_bird_date__year=year,func_tst_dummy_bird='Observation(same stage)')
            relifing_func_tst_dummy_bird_observationlistNext = RelifingSystemStatus.objects.filter(filter_objects,func_tst_dummy_bird_date__year=year,func_tst_dummy_bird='Observation(next stage)')
            relifing_func_tst_dummy_bird_uplist = RelifingSystemStatus.objects.filter(filter_objects,func_tst_dummy_bird_date__year=year,func_tst_dummy_bird='Under process')
            relifing_func_tst_dummy_bird_haultlist = RelifingSystemStatus.objects.filter(filter_objects,func_tst_dummy_bird_date__year=year,func_tst_dummy_bird='Halt')


            relifing_road_test_count = RelifingSystemStatus.objects.filter(setid_filter, road_test_date__year=year).count()

            relifing_road_test_oklist = RelifingSystemStatus.objects.filter(filter_objects,road_test_date__year=year,road_test='Ok')
            relifing_road_test_oklistNext = RelifingSystemStatus.objects.filter(filter_objects,road_test_date__year=year,road_test='OK(next stage)')
            relifing_road_test_observationlist  = RelifingSystemStatus.objects.filter(filter_objects,road_test_date__year=year,road_test='Observation(same stage)')
            relifing_road_test_observationlistNext = RelifingSystemStatus.objects.filter(filter_objects,road_test_date__year=year,road_test='Observation(next stage)')
            relifing_road_test_uplist = RelifingSystemStatus.objects.filter(filter_objects,road_test_date__year=year,road_test='Under process')
            relifing_road_test_haultlist = RelifingSystemStatus.objects.filter(filter_objects,road_test_date__year=year,road_test='Halt')


            relifing_post_road_test_count = RelifingSystemStatus.objects.filter(setid_filter, post_road_test_date__year=year).count()

            relifing_post_road_test_oklist = RelifingSystemStatus.objects.filter(filter_objects,post_road_test_date__year=year,post_road_test='Ok')
            relifing_post_road_test_oklistNext = RelifingSystemStatus.objects.filter(filter_objects,post_road_test_date__year=year,post_road_test='OK(next stage)')
            relifing_post_road_test_observationlist  = RelifingSystemStatus.objects.filter(filter_objects,post_road_test_date__year=year,post_road_test='Observation(same stage)')
            relifing_post_road_test_observationlistNext = RelifingSystemStatus.objects.filter(filter_objects,post_road_test_date__year=year,post_road_test='Observation(next stage)')
            relifing_post_road_test_uplist = RelifingSystemStatus.objects.filter(filter_objects,post_road_test_date__year=year,post_road_test='Under process')
            relifing_post_road_test_haultlist = RelifingSystemStatus.objects.filter(filter_objects,post_road_test_date__year=year,post_road_test='Halt')


            relifing_integrated_operation_count = RelifingSystemStatus.objects.filter(setid_filter, integrated_operation_date__year=year).count()

            relifing_integrated_operation_oklist = RelifingSystemStatus.objects.filter(filter_objects,integrated_operation_date__year=year,integrated_operation='Ok')
            relifing_integrated_operation_oklistNext = RelifingSystemStatus.objects.filter(filter_objects,integrated_operation_date__year=year,integrated_operation='OK(next stage)')
            relifing_integrated_operation_observationlist  = RelifingSystemStatus.objects.filter(filter_objects,integrated_operation_date__year=year,integrated_operation='Observation(same stage)')
            relifing_integrated_operation_observationlistNext = RelifingSystemStatus.objects.filter(filter_objects,integrated_operation_date__year=year,integrated_operation='Observation(next stage)')
            relifing_integrated_operation_uplist = RelifingSystemStatus.objects.filter(filter_objects,integrated_operation_date__year=year,integrated_operation='Under process')
            relifing_integrated_operation_haultlist = RelifingSystemStatus.objects.filter(filter_objects,integrated_operation_date__year=year,integrated_operation='Halt')


            relifing_rain_test_count = RelifingSystemStatus.objects.filter(setid_filter, rain_test_date__year=year).count()

            relifing_rain_test_oklist = RelifingSystemStatus.objects.filter(filter_objects,rain_test_date__year=year,rain_test='Ok')
            relifing_rain_test_oklistNext = RelifingSystemStatus.objects.filter(filter_objects,rain_test_date__year=year,rain_test='OK(next stage)')
            relifing_rain_test_observationlist  = RelifingSystemStatus.objects.filter(filter_objects,rain_test_date__year=year,rain_test='Observation(same stage)')
            relifing_rain_test_observationlistNext = RelifingSystemStatus.objects.filter(filter_objects,rain_test_date__year=year,rain_test='Observation(next stage)')
            relifing_rain_test_uplist = RelifingSystemStatus.objects.filter(filter_objects,rain_test_date__year=year,rain_test='Under process')
            relifing_rain_test_haultlist = RelifingSystemStatus.objects.filter(filter_objects,rain_test_date__year=year,rain_test='Halt')


            relifing_pre_user_inspection_count = RelifingSystemStatus.objects.filter(setid_filter, pre_user_inspection_date__year=year).count()

            relifing_pre_user_inspection_oklist = RelifingSystemStatus.objects.filter(filter_objects,pre_user_inspection_date__year=year,pre_user_inspection='Ok')
            relifing_pre_user_inspection_oklistNext = RelifingSystemStatus.objects.filter(filter_objects,pre_user_inspection_date__year=year,pre_user_inspection='OK(next stage)')
            relifing_pre_user_inspection_observationlist  = RelifingSystemStatus.objects.filter(filter_objects,pre_user_inspection_date__year=year,pre_user_inspection='Observation(same stage)')
            relifing_pre_user_inspection_observationlistNext = RelifingSystemStatus.objects.filter(filter_objects,pre_user_inspection_date__year=year,pre_user_inspection='Observation(next stage)')
            relifing_pre_user_inspection_uplist = RelifingSystemStatus.objects.filter(filter_objects,pre_user_inspection_date__year=year,pre_user_inspection='Under process')
            relifing_pre_user_inspection_haultlist = RelifingSystemStatus.objects.filter(filter_objects,pre_user_inspection_date__year=year,pre_user_inspection='Halt')


            relifing_final_integrated_testing_count = RelifingSystemStatus.objects.filter(setid_filter, final_integrated_testing_date__year=year).count()

            relifing_final_integrated_testing_oklist = RelifingSystemStatus.objects.filter(filter_objects,final_integrated_testing_date__year=year,final_integrated_testing='Ok')
            relifing_final_integrated_testing_oklistNext = RelifingSystemStatus.objects.filter(filter_objects,final_integrated_testing_date__year=year,final_integrated_testing='OK(next stage)')
            relifing_final_integrated_testing_observationlist  = RelifingSystemStatus.objects.filter(filter_objects,final_integrated_testing_date__year=year,final_integrated_testing='Observation(same stage)')
            relifing_final_integrated_testing_observationlistNext = RelifingSystemStatus.objects.filter(filter_objects,final_integrated_testing_date__year=year,final_integrated_testing='Observation(next stage)')
            relifing_final_integrated_testing_uplist = RelifingSystemStatus.objects.filter(filter_objects,final_integrated_testing_date__year=year,final_integrated_testing='Under process')
            relifing_final_integrated_testing_haultlist = RelifingSystemStatus.objects.filter(filter_objects,final_integrated_testing_date__year=year,final_integrated_testing='Halt')


            relifing_load_unload_on_mlv_hlf_count = RelifingSystemStatus.objects.filter(setid_filter, load_unload_on_mlv_hlf_date__year=year).count()

            relifing_load_unload_on_mlv_hlf_oklist = RelifingSystemStatus.objects.filter(filter_objects,load_unload_on_mlv_hlf_date__year=year,load_unload_on_mlv_hlf='Ok')
            relifing_load_unload_on_mlv_hlf_oklistNext = RelifingSystemStatus.objects.filter(filter_objects,load_unload_on_mlv_hlf_date__year=year,load_unload_on_mlv_hlf='OK(next stage)')
            relifing_load_unload_on_mlv_hlf_observationlist  = RelifingSystemStatus.objects.filter(filter_objects,load_unload_on_mlv_hlf_date__year=year,load_unload_on_mlv_hlf='Observation(same stage)')
            relifing_load_unload_on_mlv_hlf_observationlistNext = RelifingSystemStatus.objects.filter(filter_objects,load_unload_on_mlv_hlf_date__year=year,load_unload_on_mlv_hlf='Observation(next stage)')
            relifing_load_unload_on_mlv_hlf_uplist = RelifingSystemStatus.objects.filter(filter_objects,load_unload_on_mlv_hlf_date__year=year,load_unload_on_mlv_hlf='Under process')
            relifing_load_unload_on_mlv_hlf_haultlist = RelifingSystemStatus.objects.filter(filter_objects,load_unload_on_mlv_hlf_date__year=year,load_unload_on_mlv_hlf='Halt')

            # relifTotalCount = relifing_blt_count+relifing_prehil_count+relifing_posthil_count+relifing_vibration_count+relifing_cg_count+relifing_finalintegration_count+relifing_fgt_count+relifing_qmc_count+relifing_fqm_count+relifing_bhd_count+relifing_sys_align_count+relifing_incapsulation_count+relifing_emp_proofing_count+relifing_road_test_count+relifing_post_road_test_count+relifing_integrated_operation_count+relifing_rain_test_count+relifing_func_tst_count+relifing_func_tst_dummy_bird_count+relifing_final_integrated_testing_count+relifing_pre_user_inspection_count+relifing_enduser_count+relifing_load_unload_on_mlv_hlf_count

            dist = {
                'prod_blt_ok': prod_blt_oklist.count(),
                'prod_blt_ok_next': prod_blt_oklistNext.count(),
                'prod_blt_observation': prod_blt_observationlist.count(),
                'prod_blt_observation_next': prod_blt_observationlist_next.count(),
                'prod_blt_up': prod_blt_uplist.count(),
                'prod_blt_hault': prod_blt_haultlist.count(),
                'flight_blt_oklist': flight_blt_oklist.count(),
                'flight_blt_oklist_next': flight_blt_oklistNext.count(),
                'flight_blt_observationlist': flight_blt_observationlist.count(),
                'flight_blt_observationlist_next': flight_blt_observationlist_next.count(),
                'flight_blt_uplist': flight_blt_uplist.count(),
                'flight_blt_haultlist': flight_blt_haultlist.count(),
                'relifing_blt_oklist': relifing_blt_oklist.count(),
                'relifing_blt_oklist_next': relifing_blt_oklistNext.count(),
                'relifing_blt_observationlist': relifing_blt_observationlist.count(),
                'relifing_blt_observationlist_next': relifing_blt_observationlistNext.count(),
                'relifing_blt_uplist': relifing_blt_uplist.count(),
                'relifing_blt_haultlist': relifing_blt_haultlist.count(),
                'prod_prehil_oklist': prod_prehil_oklist.count(),
                'prod_prehil_oklist_next': prod_prehil_oklistNext.count(),
                'prod_prehil_observationlist': prod_prehil_observationlist.count(),
                'prod_prehil_observationlist_next': prod_prehil_observationlistNext.count(),
                'prod_prehil_uplist': prod_prehil_uplist.count(),
                'prod_prehil_haultlist': prod_prehil_haultlist.count(),
                'flight_prehil_oklist': flight_prehil_oklist.count(),
                'flight_prehil_oklist_next': flight_prehil_oklistNext.count(),
                'flight_prehil_observationlist': flight_prehil_observationlist.count(),
                'flight_prehil_observationlist_next': flight_prehil_observationlistNext.count(),
                'flight_prehil_uplist': flight_prehil_uplist.count(),
                'flight_prehil_haultlist': flight_prehil_haultlist.count(),
                'relifing_prehil_oklist': relifing_prehil_oklist.count(),
                'relifing_prehil_oklist_next': relifing_prehil_oklistNext.count(),
                'relifing_prehil_observationlist': relifing_prehil_observationlist.count(),
                'relifing_prehil_observationlist_next': relifing_prehil_observationlistNext.count(),
                'relifing_prehil_uplist': relifing_prehil_uplist.count(),
                'relifing_prehil_haultlist': relifing_prehil_haultlist.count(),
                'prod_posthil_oklist': prod_posthil_oklist.count(),
                'prod_posthil_oklist_next': prod_posthil_oklistNext.count(),
                'prod_posthil_observationlist': prod_posthil_observationlist.count(),
                'prod_posthil_observationlist_next': prod_posthil_observationlistNext.count(),
                'prod_posthil_uplist': prod_posthil_uplist.count(),
                'prod_posthil_haultlist': prod_posthil_haultlist.count(),
                'flight_posthil_oklist': flight_posthil_oklist.count(),
                'flight_posthil_oklist_next': flight_posthil_oklistNext.count(),
                'flight_posthil_observationlist': flight_posthil_observationlist.count(),
                'flight_posthil_observationlist_next': flight_posthil_observationlistNext.count(),
                'flight_posthil_uplist': flight_posthil_uplist.count(),
                'flight_posthil_haultlist': flight_posthil_haultlist.count(),
                'relifing_posthil_oklist': relifing_posthil_oklist.count(),
                'relifing_posthil_oklist_next': relifing_posthil_oklistNext.count(),
                'relifing_posthil_observationlist': relifing_posthil_observationlist.count(),
                'relifing_posthil_observationlist_next': relifing_posthil_observationlistNext.count(),
                'relifing_posthil_uplist': relifing_posthil_uplist.count(),
                'relifing_posthil_haultlist': relifing_posthil_haultlist.count(),
                'prod_finalintegration_oklist': prod_finalintegration_oklist.count(),
                'prod_finalintegration_oklist_next': prod_finalintegration_oklistNext.count(),
                'prod_finalintegration_observationlist': prod_finalintegration_observationlist.count(),
                'prod_finalintegration_observationlist_next': prod_finalintegration_observationlistNext.count(),
                'prod_finalintegration_completelist': prod_finalintegration_completelist.count(),
                'prod_finalintegration_haultlist': prod_finalintegration_haultlist.count(),
                'flight_finalintegration_oklist': flight_finalintegration_oklist.count(),
                'flight_finalintegration_oklist_next': flight_finalintegration_oklistNext.count(),
                'flight_finalintegration_observationlist': flight_finalintegration_observationlist.count(),
                'flight_finalintegration_observationlist_next': flight_finalintegration_observationlistNext.count(),
                'flight_finalintegration_completelist': flight_finalintegration_completelist.count(),
                'flight_finalintegration_haultlist': flight_finalintegration_haultlist.count(),
                'relifing_finalintegration_oklist': relifing_finalintegration_oklist.count(),
                'relifing_finalintegration_oklist_next': relifing_finalintegration_oklistNext.count(),
                'relifing_finalintegration_observationlist': relifing_finalintegration_observationlist.count(),
                'relifing_finalintegration_observationlist_next': relifing_finalintegration_observationlistNext.count(),
                'relifing_finalintegration_completelist': relifing_finalintegration_completelist.count(),
                'relifing_finalintegration_haultlist': relifing_finalintegration_haultlist.count(),
                'relifing_vibration_oklist': relifing_vibration_oklist.count(),
                'relifing_vibration_oklist_next': relifing_vibration_oklistNext.count(),
                'relifing_vibration_observationlist': relifing_vibration_observationlist.count(),
                'relifing_vibration_observationlist_next': relifing_vibration_observationlistNext.count(),
                'relifing_vibration_uplist': relifing_vibration_uplist.count(),
                'relifing_vibration_haultlist': relifing_vibration_haultlist.count(),
                'flight_vibration_oklist': flight_vibration_oklist.count(),
                'flight_vibration_oklist_next': flight_vibration_oklistNext.count(),
                'flight_vibration_observationlist': flight_vibration_observationlist.count(),
                'flight_vibration_observationlist_next': flight_vibration_observationlistNext.count(),
                'flight_vibration_uplist': flight_vibration_uplist.count(),
                'flight_vibration_haultlist': flight_vibration_haultlist.count(),
                'prod_vibration_oklist': prod_vibration_oklist.count(),
                'prod_vibration_oklist_next': prod_vibration_oklistNext.count(),
                'prod_vibration_observationlist': prod_vibration_observationlist.count(),
                'prod_vibration_observationlist_next': prod_vibration_observationlistNext.count(),
                'prod_vibration_uplist': prod_vibration_uplist.count(),
                'prod_vibration_haultlist': prod_vibration_haultlist.count(),
                'prod_cg_oklist': prod_cg_oklist.count(),
                'prod_cg_oklist_next': prod_cg_oklistNext.count(),
                'prod_cg_observationlist': prod_cg_observationlist.count(),
                'prod_cg_observationlist_next': prod_cg_observationlistNext.count(),
                'prod_cg_uplist': prod_cg_uplist.count(),
                'prod_cg_haultlist': prod_cg_haultlist.count(),
                'flight_cg_oklist': flight_cg_oklist.count(),
                'flight_cg_oklist_next': flight_cg_oklistNext.count(),
                'flight_cg_observationlist': flight_cg_observationlist.count(),
                'flight_cg_observationlist_next': flight_cg_observationlistNext.count(),
                'flight_cg_uplist': flight_cg_uplist.count(),
                'flight_cg_haultlist': flight_cg_haultlist.count(),
                'relifing_cg_oklist': relifing_cg_oklist.count(),
                'relifing_cg_oklist_next': relifing_cg_oklistNext.count(),
                'relifing_cg_observationlist': relifing_cg_observationlist.count(),
                'relifing_cg_observationlist_next': relifing_cg_observationlistNext.count(),
                'relifing_cg_uplist': relifing_cg_uplist.count(),
                'relifing_cg_haultlist': relifing_cg_haultlist.count(),
                'prod_fgt_oklist': prod_fgt_oklist.count(),
                'prod_fgt_oklist_next': prod_fgt_oklistNext.count(),
                'prod_fgt_observationlist': prod_fgt_observationlist.count(),
                'prod_fgt_observationlist_next': prod_fgt_observationlistNext.count(),
                'prod_fgt_uplist': prod_fgt_uplist.count(),
                'prod_fgt_haultlist': prod_fgt_haultlist.count(),
                'flight_fgt_oklist': flight_fgt_oklist.count(),
                'flight_fgt_oklist_next': flight_fgt_oklistNext.count(),
                'flight_fgt_observationlist': flight_fgt_observationlist.count(),
                'flight_fgt_observationlist_next': flight_fgt_observationlistNext.count(),
                'flight_fgt_uplist': flight_fgt_uplist.count(),
                'flight_fgt_haultlist': flight_fgt_haultlist.count(),
                'relifing_fgt_oklist': relifing_fgt_oklist.count(),
                'relifing_fgt_oklist_next': relifing_fgt_oklistNext.count(),
                'relifing_fgt_observationlist': relifing_fgt_observationlist.count(),
                'relifing_fgt_observationlist_next': relifing_fgt_observationlistNext.count(),
                'relifing_fgt_uplist': relifing_fgt_uplist.count(),
                'relifing_fgt_haultlist': relifing_fgt_haultlist.count(),

                'prod_bhd_ok': prod_bhd_ok.count(),
                'prod_bhd_submitted': prod_bhd_submitted.count(),
                'prod_bhd_not_submitted' : prod_bhd_not_submitted.count(),
                'prod_bhd_qm_forwarded' : prod_bhd_qm_forwarded.count(),
                'prod_bhd_qm_repeated' : prod_bhd_qm_repeated.count(),
                'prod_bhd_inprocess' : prod_bhd_inprocess.count(),

                'flight_bhd_ok': flight_bhd_ok.count(),
                'flight_bhd_submitted': flight_bhd_submitted.count(),
                'flight_bhd_not_submit' : flight_bhd_not_submit.count(),
                'flight_bhd_inprocess' : flight_bhd_inprocess.count(),
                'flight_bhd_qm_forwarded' : flight_bhd_qm_forwarded.count(),
                'flight_bhd_qm_repeated': flight_bhd_qm_repeated.count(),

                'relifing_bhd_ok': relifing_bhd_ok.count(),
                'relifing_bhd_submitted': relifing_bhd_submitted.count(),
                'relifing_bhd_not_submit' : relifing_bhd_not_submit.count(),
                'relifing_bhd_inprocess' : relifing_bhd_inprocess.count(),
                'relifing_bhd_qm_forwarded' : relifing_bhd_qm_forwarded.count(),
                'relifing_bhd_qm_repeated': relifing_bhd_qm_repeated.count(),

                'prod_fqm_planned': prod_fqm_planned.count(),
                'prod_fqm_conducted': prod_fqm_conducted.count(),
                'flight_fqm_planned': flight_fqm_planned.count(),
                'flight_fqm_conducted': flight_fqm_conducted.count(),
                'relifing_fqm_planned': relifing_fqm_planned.count(),
                'relifing_fqm_conducted': relifing_fqm_conducted.count(),
                'prod_qmc_issued': prod_qmc_issued.count(),
                'prod_qmc_in_process': prod_qmc_in_process.count(),
                'prod_qmc_obs_forwarded': prod_qmc_obs_forwarded.count(),
                'flight_qmc_issued': flight_qmc_issued.count(),
                'flight_qmc_in_process': flight_qmc_in_process.count(),
                'flight_qmc_obs_forwarded': flight_qmc_obs_forwarded.count(),
                'relifing_qmc_issued': relifing_qmc_issued.count(),
                'relifing_qmc_in_process': relifing_qmc_in_process.count(),
                'relifing_qmc_obs_forwarded': relifing_qmc_obs_forwarded.count(),
                'prod_enduser_oklist': prod_enduser_oklist.count(),
                'prod_enduser_oklist_next': prod_enduser_oklistNext.count(),
                'prod_enduser_observationlist': prod_enduser_observationlist.count(),
                'prod_enduser_observationlist_next': prod_enduser_observationlistNext.count(),
                'prod_enduser_uplist': prod_enduser_uplist.count(),
                'prod_enduser_haultlist': prod_enduser_haultlist.count(),
                'flight_launch_oklist': flight_launch_oklist.count(),
                'flight_launch_oklist_next': flight_launch_oklistNext.count(),
                'flight_launch_observationlist': flight_launch_observationlist.count(),
                'flight_launch_observationlist_next': flight_launch_observationlistNext.count(),
                'flight_launch_uplist': flight_launch_uplist.count(),
                'flight_launch_haultlist': flight_launch_haultlist.count(),
                'relifing_enduser_oklist': relifing_enduser_oklist.count(),
                'relifing_enduser_oklist_next': relifing_enduser_oklistNext.count(),
                'relifing_enduser_observationlist': relifing_enduser_observationlist.count(),
                'relifing_enduser_observationlist_next': relifing_enduser_observationlistNext.count(),
                'relifing_enduser_uplist': relifing_enduser_uplist.count(),
                'relifing_enduser_haultlist': relifing_enduser_haultlist.count(),
                'prod_sys_align_oklist' : prod_sys_align_oklist.count(),
                'prod_sys_align_oklistNext' : prod_sys_align_oklistNext.count(),
                'prod_sys_align_observationlist' : prod_sys_align_observationlist.count(),
                'prod_sys_align_observationlistNext' : prod_sys_align_observationlistNext.count(),
                'prod_sys_align_uplist' : prod_sys_align_uplist.count(),
                'prod_sys_align_haultlist' : prod_sys_align_haultlist.count(),
                'prod_incapsulation_oklist' : prod_incapsulation_oklist.count(),
                'prod_incapsulation_oklistNext' : prod_incapsulation_oklistNext.count(),
                'prod_incapsulation_observationlist' : prod_incapsulation_observationlist.count(),
                'prod_incapsulation_observationlistNext' : prod_incapsulation_observationlistNext.count(),
                'prod_incapsulation_uplist' : prod_incapsulation_uplist.count(),
                'prod_incapsulation_haultlist' : prod_incapsulation_haultlist.count(),
                'flight_sys_align_oklist' : flight_sys_align_oklist.count(),
                'flight_sys_align_oklistNext' : flight_sys_align_oklistNext.count(),
                'flight_sys_align_observationlist' : flight_sys_align_observationlist.count(),
                'flight_sys_align_observationlistNext' : flight_sys_align_observationlistNext.count(),
                'flight_sys_align_uplist' : flight_sys_align_uplist.count(),
                'flight_sys_align_haultlist' : flight_sys_align_haultlist.count(),
                'flight_incapsulation_oklist' : flight_incapsulation_oklist.count(),
                'flight_incapsulation_oklistNext' : flight_incapsulation_oklistNext.count(),
                'flight_incapsulation_observationlist' : flight_incapsulation_observationlist.count(),
                'flight_incapsulation_observationlistNext' : flight_incapsulation_observationlistNext.count(),
                'flight_incapsulation_uplist' : flight_incapsulation_uplist.count(),
                'flight_incapsulation_haultlist' : flight_incapsulation_haultlist.count(),
                'relifing_sys_align_oklist' : relifing_sys_align_oklist.count(),
                'relifing_sys_align_oklistNext' : relifing_sys_align_oklistNext.count(),
                'relifing_sys_align_observationlist' : relifing_sys_align_observationlist.count(),
                'relifing_sys_align_observationlistNext' : relifing_sys_align_observationlistNext.count(),
                'relifing_sys_align_uplist' : relifing_sys_align_uplist.count(),
                'relifing_sys_align_haultlist' : relifing_sys_align_haultlist.count(),
                'relifing_incapsulation_oklist' : relifing_incapsulation_oklist.count(),
                'relifing_incapsulation_oklistNext' : relifing_incapsulation_oklistNext.count(),
                'relifing_incapsulation_observationlist' : relifing_incapsulation_observationlist.count(),
                'relifing_incapsulation_observationlistNext' : relifing_incapsulation_observationlistNext.count(),
                'relifing_incapsulation_uplist' : relifing_incapsulation_uplist.count(),
                'relifing_incapsulation_haultlist' : relifing_incapsulation_haultlist.count(),

                'prod_emp_proofing_oklist' : prod_emp_proofing_oklist.count(),
                'prod_emp_proofing_oklistNext': prod_emp_proofing_oklistNext.count(),
                'prod_emp_proofing_observationlist'  : prod_emp_proofing_observationlist.count(),
                'prod_emp_proofing_observationlistNext' : prod_emp_proofing_observationlistNext.count(),
                'prod_emp_proofing_uplist' : prod_emp_proofing_uplist.count(),
                'prod_emp_proofing_haultlist' : prod_emp_proofing_haultlist.count(),

                'prod_func_tst_oklistNext_oklist'  : prod_func_tst_oklistNext_oklist.count(),
                'prod_func_tst_oklistNext_oklistNext'  : prod_func_tst_oklistNext_oklistNext.count(),
                'prod_func_tst_oklistNext_observationlist'   : prod_func_tst_oklistNext_observationlist.count(),
                'prod_func_tst_oklistNext_observationlistNext'  : prod_func_tst_oklistNext_observationlistNext.count(),
                'prod_func_tst_oklistNext_uplist'  : prod_func_tst_oklistNext_uplist.count(),
                'prod_func_tst_oklistNext_haultlist'  : prod_func_tst_oklistNext_haultlist.count(),

                'prod_func_tst_dummy_bird_oklist' : prod_func_tst_dummy_bird_oklist.count(),
                'prod_func_tst_dummy_bird_oklistNext' : prod_func_tst_dummy_bird_oklistNext.count(),
                'prod_func_tst_dummy_bird_observationlist'  : prod_func_tst_dummy_bird_observationlist.count(),
                'prod_func_tst_dummy_bird_observationlistNext' : prod_func_tst_dummy_bird_observationlistNext.count(),
                'prod_func_tst_dummy_bird_uplist' : prod_func_tst_dummy_bird_uplist.count(),
                'prod_func_tst_dummy_bird_haultlist' : prod_func_tst_dummy_bird_haultlist.count(),

                'prod_road_test_oklist' : prod_road_test_oklist.count(),
                'prod_road_test_oklistNext' : prod_road_test_oklistNext.count(),
                'prod_road_test_observationlist'  : prod_road_test_observationlist.count(),
                'prod_road_test_observationlistNext' : prod_road_test_observationlistNext.count(),
                'prod_road_test_uplist' : prod_road_test_uplist.count(),
                'prod_road_test_haultlist' : prod_road_test_haultlist.count(),

                'prod_post_road_test_oklist' : prod_post_road_test_oklist.count(),
                'prod_post_road_test_oklistNext' : prod_post_road_test_oklistNext.count(),
                'prod_post_road_test_observationlist'  : prod_post_road_test_observationlist.count(),
                'prod_post_road_test_observationlistNext' : prod_post_road_test_observationlistNext.count(),
                'prod_post_road_test_uplist' : prod_post_road_test_uplist.count(),
                'prod_post_road_test_haultlist' : prod_post_road_test_haultlist.count(),

                'prod_integrated_operation_oklist' : prod_integrated_operation_oklist.count(),
                'prod_integrated_operation_oklistNext' : prod_integrated_operation_oklistNext.count(),
                'prod_integrated_operation_observationlist'  : prod_integrated_operation_observationlist.count(),
                'prod_integrated_operation_observationlistNext' : prod_integrated_operation_observationlistNext.count(),
                'prod_integrated_operation_uplist' : prod_integrated_operation_uplist.count(),
                'prod_integrated_operation_haultlist' : prod_integrated_operation_haultlist.count(),

                'prod_rain_test_oklist' : prod_rain_test_oklist.count(),
                'prod_rain_test_oklistNext' : prod_rain_test_oklistNext.count(),
                'prod_rain_test_observationlist'  : prod_rain_test_observationlist.count(),
                'prod_rain_test_observationlistNext' : prod_rain_test_observationlistNext.count(),
                'prod_rain_test_uplist' : prod_rain_test_uplist.count(),
                'prod_rain_test_haultlist' : prod_rain_test_haultlist.count(),

                'prod_pre_user_inspection_oklist' : prod_pre_user_inspection_oklist.count(),
                'prod_pre_user_inspection_oklistNext' : prod_pre_user_inspection_oklistNext.count(),
                'prod_pre_user_inspection_observationlist'  : prod_pre_user_inspection_observationlist.count(),
                'prod_pre_user_inspection_observationlistNext' : prod_pre_user_inspection_observationlistNext.count(),
                'prod_pre_user_inspection_uplist' : prod_pre_user_inspection_uplist.count(),
                'prod_pre_user_inspection_haultlist' : prod_pre_user_inspection_haultlist.count(),

                'prod_final_integrated_testing_oklist' : prod_final_integrated_testing_oklist.count(),
                'prod_final_integrated_testing_oklistNext' : prod_final_integrated_testing_oklistNext.count(),
                'prod_final_integrated_testing_observationlist'  : prod_final_integrated_testing_observationlist.count(),
                'prod_final_integrated_testing_observationlistNext' : prod_final_integrated_testing_observationlistNext.count(),
                'prod_final_integrated_testing_uplist' : prod_final_integrated_testing_uplist.count(),
                'prod_final_integrated_testing_haultlist' : prod_final_integrated_testing_haultlist.count(),

                'prod_load_unload_on_mlv_hlf_oklist' : prod_load_unload_on_mlv_hlf_oklist.count(),
                'prod_load_unload_on_mlv_hlf_oklistNext' : prod_load_unload_on_mlv_hlf_oklistNext.count(),
                'prod_load_unload_on_mlv_hlf_observationlist'  : prod_load_unload_on_mlv_hlf_observationlist.count(),
                'prod_load_unload_on_mlv_hlf_observationlistNext' : prod_load_unload_on_mlv_hlf_observationlistNext.count(),
                'prod_load_unload_on_mlv_hlf_uplist' : prod_load_unload_on_mlv_hlf_uplist.count(),
                'prod_load_unload_on_mlv_hlf_haultlist' : prod_load_unload_on_mlv_hlf_haultlist.count(),

                'relifing_emp_proofing_oklist' : relifing_emp_proofing_oklist.count(),
                'relifing_emp_proofing_oklistNext': relifing_emp_proofing_oklistNext.count(),
                'relifing_emp_proofing_observationlist'  : relifing_emp_proofing_observationlist.count(),
                'relifing_emp_proofing_observationlistNext' : relifing_emp_proofing_observationlistNext.count(),
                'relifing_emp_proofing_uplist' : relifing_emp_proofing_uplist.count(),
                'relifing_emp_proofing_haultlist' : relifing_emp_proofing_haultlist.count(),

                'relifing_func_tst_oklistNext_oklist'  : relifing_func_tst_oklistNext_oklist.count(),
                'relifing_func_tst_oklistNext_oklistNext'  : relifing_func_tst_oklistNext_oklistNext.count(),
                'relifing_func_tst_oklistNext_observationlist'   : relifing_func_tst_oklistNext_observationlist.count(),
                'relifing_func_tst_oklistNext_observationlistNext'  : relifing_func_tst_oklistNext_observationlistNext.count(),
                'relifing_func_tst_oklistNext_uplist'  : relifing_func_tst_oklistNext_uplist.count(),
                'relifing_func_tst_oklistNext_haultlist'  : relifing_func_tst_oklistNext_haultlist.count(),

                'relifing_func_tst_dummy_bird_oklist' : relifing_func_tst_dummy_bird_oklist.count(),
                'relifing_func_tst_dummy_bird_oklistNext' : relifing_func_tst_dummy_bird_oklistNext.count(),
                'relifing_func_tst_dummy_bird_observationlist'  : relifing_func_tst_dummy_bird_observationlist.count(),
                'relifing_func_tst_dummy_bird_observationlistNext' : relifing_func_tst_dummy_bird_observationlistNext.count(),
                'relifing_func_tst_dummy_bird_uplist' : relifing_func_tst_dummy_bird_uplist.count(),
                'relifing_func_tst_dummy_bird_haultlist' : relifing_func_tst_dummy_bird_haultlist.count(),

                'relifing_road_test_oklist' : relifing_road_test_oklist.count(),
                'relifing_road_test_oklistNext' : relifing_road_test_oklistNext.count(),
                'relifing_road_test_observationlist'  : relifing_road_test_observationlist.count(),
                'relifing_road_test_observationlistNext' : relifing_road_test_observationlistNext.count(),
                'relifing_road_test_uplist' : relifing_road_test_uplist.count(),
                'relifing_road_test_haultlist' : relifing_road_test_haultlist.count(),

                'relifing_post_road_test_oklist' : relifing_post_road_test_oklist.count(),
                'relifing_post_road_test_oklistNext' : relifing_post_road_test_oklistNext.count(),
                'relifing_post_road_test_observationlist'  : relifing_post_road_test_observationlist.count(),
                'relifing_post_road_test_observationlistNext' : relifing_post_road_test_observationlistNext.count(),
                'relifing_post_road_test_uplist' : relifing_post_road_test_uplist.count(),
                'relifing_post_road_test_haultlist' : relifing_post_road_test_haultlist.count(),

                'relifing_integrated_operation_oklist' : relifing_integrated_operation_oklist.count(),
                'relifing_integrated_operation_oklistNext' : relifing_integrated_operation_oklistNext.count(),
                'relifing_integrated_operation_observationlist'  : relifing_integrated_operation_observationlist.count(),
                'relifing_integrated_operation_observationlistNext' : relifing_integrated_operation_observationlistNext.count(),
                'relifing_integrated_operation_uplist' : relifing_integrated_operation_uplist.count(),
                'relifing_integrated_operation_haultlist' : relifing_integrated_operation_haultlist.count(),

                'relifing_rain_test_oklist' : relifing_rain_test_oklist.count(),
                'relifing_rain_test_oklistNext' : relifing_rain_test_oklistNext.count(),
                'relifing_rain_test_observationlist'  : relifing_rain_test_observationlist.count(),
                'relifing_rain_test_observationlistNext' : relifing_rain_test_observationlistNext.count(),
                'relifing_rain_test_uplist' : relifing_rain_test_uplist.count(),
                'relifing_rain_test_haultlist' : relifing_rain_test_haultlist.count(),

                'relifing_pre_user_inspection_oklist' : relifing_pre_user_inspection_oklist.count(),
                'relifing_pre_user_inspection_oklistNext' : relifing_pre_user_inspection_oklistNext.count(),
                'relifing_pre_user_inspection_observationlist'  : relifing_pre_user_inspection_observationlist.count(),
                'relifing_pre_user_inspection_observationlistNext' : relifing_pre_user_inspection_observationlistNext.count(),
                'relifing_pre_user_inspection_uplist' : relifing_pre_user_inspection_uplist.count(),
                'relifing_pre_user_inspection_haultlist' : relifing_pre_user_inspection_haultlist.count(),

                'relifing_final_integrated_testing_oklist' : relifing_final_integrated_testing_oklist.count(),
                'relifing_final_integrated_testing_oklistNext' : relifing_final_integrated_testing_oklistNext.count(),
                'relifing_final_integrated_testing_observationlist'  : relifing_final_integrated_testing_observationlist.count(),
                'relifing_final_integrated_testing_observationlistNext' : relifing_final_integrated_testing_observationlistNext.count(),
                'relifing_final_integrated_testing_uplist' : relifing_final_integrated_testing_uplist.count(),
                'relifing_final_integrated_testing_haultlist' : relifing_final_integrated_testing_haultlist.count(),

                'relifing_load_unload_on_mlv_hlf_oklist' : relifing_load_unload_on_mlv_hlf_oklist.count(),
                'relifing_load_unload_on_mlv_hlf_oklistNext' : relifing_load_unload_on_mlv_hlf_oklistNext.count(),
                'relifing_load_unload_on_mlv_hlf_observationlist'  : relifing_load_unload_on_mlv_hlf_observationlist.count(),
                'relifing_load_unload_on_mlv_hlf_observationlistNext' : relifing_load_unload_on_mlv_hlf_observationlistNext.count(),
                'relifing_load_unload_on_mlv_hlf_uplist' : relifing_load_unload_on_mlv_hlf_uplist.count(),
                'relifing_load_unload_on_mlv_hlf_haultlist' : relifing_load_unload_on_mlv_hlf_haultlist.count(),

                'flight_emp_proofing_oklist' : flight_emp_proofing_oklist.count(),
                'flight_emp_proofing_oklistNext': flight_emp_proofing_oklistNext.count(),
                'flight_emp_proofing_observationlist'  : flight_emp_proofing_observationlist.count(),
                'flight_emp_proofing_observationlistNext' : flight_emp_proofing_observationlistNext.count(),
                'flight_emp_proofing_uplist' : flight_emp_proofing_uplist.count(),
                'flight_emp_proofing_haultlist' : flight_emp_proofing_haultlist.count(),

                'flight_func_tst_oklistNext_oklist'  : flight_func_tst_oklistNext_oklist.count(),
                'flight_func_tst_oklistNext_oklistNext'  : flight_func_tst_oklistNext_oklistNext.count(),
                'flight_func_tst_oklistNext_observationlist'   : flight_func_tst_oklistNext_observationlist.count(),
                'flight_func_tst_oklistNext_observationlistNext'  : flight_func_tst_oklistNext_observationlistNext.count(),
                'flight_func_tst_oklistNext_uplist'  : flight_func_tst_oklistNext_uplist.count(),
                'flight_func_tst_oklistNext_haultlist'  : flight_func_tst_oklistNext_haultlist.count(),

                'flight_func_tst_dummy_bird_oklist' : flight_func_tst_dummy_bird_oklist.count(),
                'flight_func_tst_dummy_bird_oklistNext' : flight_func_tst_dummy_bird_oklistNext.count(),
                'flight_func_tst_dummy_bird_observationlist'  : flight_func_tst_dummy_bird_observationlist.count(),
                'flight_func_tst_dummy_bird_observationlistNext' : flight_func_tst_dummy_bird_observationlistNext.count(),
                'flight_func_tst_dummy_bird_uplist' : flight_func_tst_dummy_bird_uplist.count(),
                'flight_func_tst_dummy_bird_haultlist' : flight_func_tst_dummy_bird_haultlist.count(),

                'flight_road_test_oklist' : flight_road_test_oklist.count(),
                'flight_road_test_oklistNext' : flight_road_test_oklistNext.count(),
                'flight_road_test_observationlist'  : flight_road_test_observationlist.count(),
                'flight_road_test_observationlistNext' : flight_road_test_observationlistNext.count(),
                'flight_road_test_uplist' : flight_road_test_uplist.count(),
                'flight_road_test_haultlist' : flight_road_test_haultlist.count(),

                'flight_post_road_test_oklist' : flight_post_road_test_oklist.count(),
                'flight_post_road_test_oklistNext' : flight_post_road_test_oklistNext.count(),
                'flight_post_road_test_observationlist'  : flight_post_road_test_observationlist.count(),
                'flight_post_road_test_observationlistNext' : flight_post_road_test_observationlistNext.count(),
                'flight_post_road_test_uplist' : flight_post_road_test_uplist.count(),
                'flight_post_road_test_haultlist' : flight_post_road_test_haultlist.count(),

                'flight_integrated_operation_oklist' : flight_integrated_operation_oklist.count(),
                'flight_integrated_operation_oklistNext' : flight_integrated_operation_oklistNext.count(),
                'flight_integrated_operation_observationlist'  : flight_integrated_operation_observationlist.count(),
                'flight_integrated_operation_observationlistNext' : flight_integrated_operation_observationlistNext.count(),
                'flight_integrated_operation_uplist' : flight_integrated_operation_uplist.count(),
                'flight_integrated_operation_haultlist' : flight_integrated_operation_haultlist.count(),

                'flight_rain_test_oklist' : flight_rain_test_oklist.count(),
                'flight_rain_test_oklistNext' : flight_rain_test_oklistNext.count(),
                'flight_rain_test_observationlist'  : flight_rain_test_observationlist.count(),
                'flight_rain_test_observationlistNext' : flight_rain_test_observationlistNext.count(),
                'flight_rain_test_uplist' : flight_rain_test_uplist.count(),
                'flight_rain_test_haultlist' : flight_rain_test_haultlist.count(),

                'flight_pre_user_inspection_oklist' : flight_pre_user_inspection_oklist.count(),
                'flight_pre_user_inspection_oklistNext' : flight_pre_user_inspection_oklistNext.count(),
                'flight_pre_user_inspection_observationlist'  : flight_pre_user_inspection_observationlist.count(),
                'flight_pre_user_inspection_observationlistNext' : flight_pre_user_inspection_observationlistNext.count(),
                'flight_pre_user_inspection_uplist' : flight_pre_user_inspection_uplist.count(),
                'flight_pre_user_inspection_haultlist' : flight_pre_user_inspection_haultlist.count(),

                'flight_final_integrated_testing_oklist' : flight_final_integrated_testing_oklist.count(),
                'flight_final_integrated_testing_oklistNext' : flight_final_integrated_testing_oklistNext.count(),
                'flight_final_integrated_testing_observationlist'  : flight_final_integrated_testing_observationlist.count(),
                'flight_final_integrated_testing_observationlistNext' : flight_final_integrated_testing_observationlistNext.count(),
                'flight_final_integrated_testing_uplist' : flight_final_integrated_testing_uplist.count(),
                'flight_final_integrated_testing_haultlist' : flight_final_integrated_testing_haultlist.count(),

                'flight_load_unload_on_mlv_hlf_oklist' : flight_load_unload_on_mlv_hlf_oklist.count(),
                'flight_load_unload_on_mlv_hlf_oklistNext' : flight_load_unload_on_mlv_hlf_oklistNext.count(),
                'flight_load_unload_on_mlv_hlf_observationlist'  : flight_load_unload_on_mlv_hlf_observationlist.count(),
                'flight_load_unload_on_mlv_hlf_observationlistNext' : flight_load_unload_on_mlv_hlf_observationlistNext.count(),
                'flight_load_unload_on_mlv_hlf_uplist' : flight_load_unload_on_mlv_hlf_uplist.count(),
                'flight_load_unload_on_mlv_hlf_haultlist' : flight_load_unload_on_mlv_hlf_haultlist.count(),
                'prodTotalCount': prodTotalCount,
                'flightTotalCount': flightTotalCount,
                'relifTotalCount': relifTotalCount
            }

            return JsonResponse({'message': 'true', 'data': dist}, status=200)
        except Exception as e:
            print(e)
            return JsonResponse({'message': 'Sorry! No Task found.'}, status=500)
