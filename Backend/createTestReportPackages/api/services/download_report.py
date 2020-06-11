import base64

from flask import send_from_directory, send_file
from flask_restplus import fields
import os
import uuid
import pythoncom
import win32com.client
from pdf2image import convert_from_path
import cv2
import numpy as np
from PIL import Image
import pickle
import fitz
from shutil import copyfile

from createTestReportPackages.api.api_response import pdf_response
from createTestReportPackages.parser import CONFIG

wdFormatPDF = 17

URL = '/test-report-generator/download_report'

request_fields = {
    "report_name": fields.String('Name of the report to be downloaded', required=True),
    "report_type": fields.String('Type of report to be downloaded', required=True)
}


def func(request_json):
    report_name = request_json["report_name"]
    report_type = request_json["report_type"]
    report_name = os.path.splitext(report_name)[0]
    if report_type=="FullReport":
        save_document_path = os.path.join(CONFIG["tempFolder"], str(report_name), str(report_name)+"_img.pdf")
    else:
        save_document_path = os.path.join(CONFIG["tempFolder"], str(report_name), str(report_name)+".pdf")
    print(save_document_path)
    file_obj = open(save_document_path, "rb").read()
    b64_pdf = base64.b64encode(file_obj).decode("utf-8")
    # response = send_file(CONFIG["tempFolder"], str(report_name), as_attachment=True)
    return pdf_response(b64_pdf, 200)
