import pandas as pd
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
from reportlab.pdfgen import canvas
from PyPDF2 import PdfFileWriter, PdfFileReader
from createTestReportPackages.model.MailService import MailService
from createTestReportPackages.parser import CONFIG
import datetime
import json
from createTestReportPackages.model import ReportsData
from createTestReportPackages.utils.helper_utilities import write_report_in_dir

wdFormatPDF = 17

URL = '/test-report-generator/create-report-from-doc'

request_fields = {
    "report_file_path": fields.String('File path of the report', required=True),
    "test_engineer_name": fields.String('Test engineer name', required=True),
    "report_file_name": fields.String('Report file name', required=True)
}


def doc_to_pdf(in_file, out_file):
    pythoncom.CoInitialize()
    word = win32com.client.Dispatch('Word.Application')
    doc = word.Documents.Open(in_file)
    doc.SaveAs(out_file, FileFormat=wdFormatPDF)
    doc.Close()


def make_text_pdf_with_watermark(out_file, img_pdf_path, test_engineer_name, test_engineer_stamp, approving_authority):
    file_obj_text_pdf = PdfFileReader(open(out_file, "rb"))
    X = float(file_obj_text_pdf.getPage(0).mediabox[2])
    Y = float(file_obj_text_pdf.getPage(0).mediabox[3])
    approving_authorit_stamp_coordinates, approving_authority_stamp = get_location_for_approving_authority(out_file, X, Y, approving_authority)
    BDH_stamp_coordinates, BDH_stamp = get_stamp_location_for_BDH(out_file, X, Y)
    test_engineer_stamp_coordinates = get_stamp_location_for_test_engineer(out_file, test_engineer_name, X, Y)
    img = cv2.imread(test_engineer_stamp, cv2.IMREAD_UNCHANGED)
    height = img.shape[0]
    width = img.shape[1]
    width_te_sign = 60
    dwh = width / width_te_sign
    height_te_sign = height / dwh
    page_count_text_pdf = file_obj_text_pdf.getNumPages()
    document_name = os.path.splitext(out_file)[0]
    # save_document_path = CONFIG["tempFolder"] + str(document_name)
    save_document_path = str(document_name)
    watermark_file_name = os.path.join(save_document_path, document_name) + 'signs_watermark.pdf'
    # watermark_file_name = 'letterhead_to_add.pdf'
    c = canvas.Canvas(watermark_file_name)
    watermark = PdfFileReader(open("Letterheadsite2.pdf", "rb"))
    output_file_wm = PdfFileWriter()
    for i in range(page_count_text_pdf):
        if i in approving_authorit_stamp_coordinates:
            coordinates = approving_authorit_stamp_coordinates[i]
            c.drawImage(approving_authority_stamp, coordinates["X"], Y - coordinates["Y"], 60, 60)
        if i in BDH_stamp_coordinates:
            coordinates = BDH_stamp_coordinates[i]
            c.drawImage(BDH_stamp, coordinates["X"], Y - coordinates["Y"], 150, 60)
        if i in test_engineer_stamp_coordinates:
            coordinates = test_engineer_stamp_coordinates[i]
            c.drawImage(test_engineer_stamp, coordinates["X"], Y - coordinates["Y"], width_te_sign, height_te_sign)
        c.showPage()
        output_file_wm.addPage(watermark.getPage(0))
    c.save()
    multiple_img_path =  os.path.join(save_document_path, document_name) + '_Letter_multiple.pdf'
    with open(multiple_img_path, "wb") as outputStream:
        output_file_wm.write(outputStream)
    output_file = PdfFileWriter()

    watermark = PdfFileReader(open(multiple_img_path, "rb"))
    input_file = PdfFileReader(open(watermark_file_name, "rb"))
    for page_number in range(page_count_text_pdf):
        print ("Watermarking page {} of {}".format(page_number, page_count_text_pdf))
        # merge the watermark with the page
        input_page = watermark.getPage(page_number)
        input_page.mergePage(input_file.getPage(page_number))
        input_page.mergePage(file_obj_text_pdf.getPage(page_number))
        # add page from input file to output document
        output_file.addPage(input_page)

    # finally, write "output" to document-output.pdf
    with open(img_pdf_path, "wb") as outputStream:
        output_file.write(outputStream)

def pdf_to_image_pdf(out_file, img_pdf_path, test_engineer_name, test_engineer_dict, approving_authority):
    pages = convert_from_path(out_file, 200)
    X = pages[0].size[0]
    Y = pages[0].size[1]
    new_pixmap_list = get_all_pixmap(out_file, X, Y, approving_authority)
    BDH_atamp_pixmap = get_pix_map_for_BDH(out_file, X, Y)
    test_engineer_sign = get_pix_map_for_test_engineer(out_file, test_engineer_name, test_engineer_dict, X, Y)
    letterhead = get_letterhead("letterhead.pickle")
    newPage = []
    for i, page in enumerate(pages):
        open_cv_image = np.array(page)
        pixmap = new_pixmap_list[i]
        for pixel in pixmap.keys():
            open_cv_image[pixel[0], pixel[1]] = pixmap[pixel]
        pixmap = BDH_atamp_pixmap[i]
        for pixel in pixmap.keys():
            open_cv_image[pixel[0], pixel[1]] = pixmap[pixel]
        pixmap = test_engineer_sign[i]
        for pixel in pixmap.keys():
            open_cv_image[pixel[0], pixel[1]] = pixmap[pixel]
        pixmap = letterhead["pixmap"]
        for pixel in pixmap.keys():
            try:
                open_cv_image[pixel[0], pixel[1]] = pixmap[pixel]
            except Exception as e:
                continue
        scale_percent = 60  # percent of original size
        width = int(open_cv_image.shape[1] * scale_percent / 100)
        height = int(open_cv_image.shape[0] * scale_percent / 100)
        dim = (width, height)
        # resize image
        resized = cv2.resize(open_cv_image, dim, interpolation=cv2.INTER_AREA)
        newPage.append(Image.fromarray(resized))
    newPage[0].save(img_pdf_path, save_all=True, append_images=newPage[1:])

def get_pix_map_for_test_engineer(out_file, name, test_engineer_dict, X, Y):
    stamp = get_stamp(test_engineer_dict[name])
    pixmap = stamp["pixmap"]
    stamp_height = stamp["height"]
    stamp_widht = stamp["width"]
    doc = fitz.open(out_file)
    page = doc[0]
    scalex = X / page.MediaBox[2]
    scaley = Y / page.MediaBox[3]
    print(scalex, scaley)
    new_pixmap_list = []
    for page in doc:
        text = name
        text_instances = page.searchFor(text)
        if (text_instances):
            new_pix_map = {}
            dx = int((text_instances[0][0] * scalex) - (1 * stamp_widht) / 6)
            dy = int((text_instances[0][1] * scaley) - (stamp_height * 1.1))
            for key in pixmap.keys():
                new_pix_map[key[0] + dy, key[1] + dx] = pixmap[key]
            new_pixmap_list.append(new_pix_map)
        else:
            new_pixmap_list.append({})
    return new_pixmap_list



def get_pix_map_for_forsign_wiht_location(out_file, name ,text_to_replace, X, Y, place= ""):
    stamp = get_stamp(name)
    pixmap = stamp["pixmap"]
    stamp_height = stamp["height"]
    stamp_widht = stamp["width"]
    doc = fitz.open(out_file)
    page = doc[0]
    scalex = X / page.MediaBox[2]
    scaley = Y / page.MediaBox[3]
    print(scalex, scaley)
    new_pixmap_list = []
    for page in doc:
        text = text_to_replace
        text_instances = page.searchFor(text)

        if (text_instances):
            new_pix_map = {}
            dx = int((text_instances[0][0] * scalex) - (1 * stamp_widht) / 6)
            dy = int((text_instances[0][1] * scaley) - (stamp_height * 1.1))
            if place == "right":
                dx = int((text_instances[0][0] * scalex) + stamp_widht*0.4)
                dy = int(0.7 * stamp_height)
                if text_to_replace == "Date :":
                    dy = int((text_instances[0][1] * scaley) - (stamp_height * 0.5))
            for key in pixmap.keys():
                new_pix_map[key[0] + dy, key[1] + dx] = pixmap[key]
            new_pixmap_list.append(new_pix_map)
        else:
            new_pixmap_list.append({})
    return new_pixmap_list


def get_pix_map_for_BDH(out_file, X, Y):
    stamp = get_stamp("stampBDH.pickle")
    pixmap = stamp["pixmap"]
    stamp_height = stamp["height"]
    stamp_widht = stamp["width"]
    doc = fitz.open(out_file)
    page = doc[0]
    scalex = X / page.MediaBox[2]
    scaley = Y / page.MediaBox[3]
    print(scalex, scaley)
    new_pixmap_list = []
    for page in doc:
        text = "Sailesh Chandra Srivastava"
        text_instances = page.searchFor(text)
        if (text_instances):
            new_pix_map = {}
            dx = int((text_instances[0][0] * scalex) - (1 * stamp_widht) / 6)
            dy = int((text_instances[0][1] * scaley) - stamp_height)
            for key in pixmap.keys():
                new_pix_map[key[0] + dy, key[1] + dx] = pixmap[key]
            new_pixmap_list.append(new_pix_map)
        else:
            new_pixmap_list.append({})
    return new_pixmap_list



def get_stamp_location_for_BDH(out_file, X, Y):
    doc = fitz.open(out_file)
    page = doc[0]
    scalex = X / page.MediaBox[2]
    scaley = Y / page.MediaBox[3]
    print(scalex, scaley)
    coordinate_list = {}
    BDH_stamp = os.path.join("signs","stampBDH1.png")
    for i, page in enumerate(doc):
        text = "Sailesh Chandra Srivastava"
        text_instances = page.searchFor(text)
        if (text_instances):
            coordinate_list[i] = {"X":text_instances[0][0],"Y":text_instances[0][1]}
    return coordinate_list, BDH_stamp


def get_location_for_approving_authority(out_file, X, Y, approving_authority):
    if approving_authority == "Zahid Raza":
        name_list = ["Zahid Raza", "Approving Authority","(Signature of Authorized person"]
        approving_authority_stamp = os.path.join("signs","Zahid.png")
    elif approving_authority == "Avishek Kumar":
        name_list = ["Avishek Kumar", "Approving Authority", "(Signature of Authorized person"]
        approving_authority_stamp = os.path.join("signs", "Avishek.png")
    else:
        name_list = ["Shashank Raghubanshi", "Approving Authority", "(Signature of Authorized person"]
        approving_authority_stamp = os.path.join("signs", "ShashankRaghubanshiSign.png")
    doc = fitz.open(out_file)
    page = doc[0]
    scalex = X / page.MediaBox[2]
    scaley = Y / page.MediaBox[3]
    print(scalex, scaley)
    coordinate_list = {}
    for i, page in enumerate(doc):
        text_instances = []
        for text in name_list:
            text_instances += page.searchFor(text)
        if (text_instances):
            coordinate_list[i] = {"X":text_instances[0][0],"Y":text_instances[0][1]}
        # else:
        #     text1 = page.getText(output='dict')
        #     maxY = 0
        #     for box in text1['blocks']:
        #         boxY = box['bbox'][3]
        #         if maxY < boxY:
        #             try:
        #                 if not box['lines'][0]['spans'][0]['text'].startswith("TRF No."):
        #                     maxY = boxY
        #             except Exception as e:
        #                 maxY = boxY
        #     lastRowHeight = maxY
        #     coordinate_list[i] = {"X":X-100, "Y":lastRowHeight-35}
    return coordinate_list, approving_authority_stamp


def get_stamp_location_for_test_engineer(out_file, name, X, Y):
    doc = fitz.open(out_file)
    page = doc[0]
    scalex = X / page.MediaBox[2]
    scaley = Y / page.MediaBox[3]
    print(scalex, scaley)
    coordinate_list = {}
    for i, page in enumerate(doc):
        text = name
        text_instances = page.searchFor(text)
        if (text_instances):
            coordinate_list[i] = {"X":text_instances[0][0],"Y":text_instances[0][1]}
    return coordinate_list


def get_all_pixmap(out_file, X, Y, approving_authority):
    if approving_authority == "Zahid Raza":
        stamp = get_stamp("ZahidnewSign.pickle")
        name_list = ["Zahid Raza", "Approving Authority","(Signature of Authorized person"]
        xx = 1
    else:
        stamp = get_stamp("ShashankRaghubanshiSign.pickle")
        name_list = ["Shashank Raghubanshi", "Approving Authority", "(Signature of Authorized person"]
        xx = 0
    pixmap = stamp["pixmap"]
    stamp_height = stamp["height"]
    stamp_widht = stamp["width"]
    doc = fitz.open(out_file)
    page = doc[0]
    scalex = X / page.MediaBox[2]
    scaley = Y / page.MediaBox[3]
    print(scalex, scaley)
    new_pixmap_list = []
    for page in doc:
        text_instances = []
        for text in name_list:
            text_instances += page.searchFor(text)
        new_pix_map = {}
        if (text_instances):
            dx = int(int(text_instances[0][0] * scalex) - (xx * stamp_height) / 4)
            dy = int((text_instances[0][1] * scaley) - (3.5 * stamp_height) / 4)
        else:
            text1 = page.getText(output='dict')
            maxY = 0
            for box in text1['blocks']:
                boxY = box['bbox'][3]
                if maxY < boxY:
                    try:
                        if not box['lines'][0]['spans'][0]['text'].startswith("TRF No."):
                            maxY = boxY
                    except Exception as e:
                        maxY = boxY
            lastRowHeight = maxY
            dx = X - int(1.2 * stamp_widht)
            new_y = int((lastRowHeight * scaley))
            if new_y >= Y - stamp_height:
                dy = Y - int(1.2 * stamp_height)
            else:
                dy = new_y
        for key in pixmap.keys():
            new_pix_map[key[0] + dy, key[1] + dx] = pixmap[key]
        new_pixmap_list.append(new_pix_map)
    return new_pixmap_list


def get_letterhead(stamp_name):
    create_stamp = False
    pixmap = {}
    stamp_height = 0
    stamp_widht = 0
    if create_stamp:
        pages = convert_from_path(r"C:\Users\aditya.verma\Desktop\letterhead.pdf", 200)
        stamp = np.array(pages[0])
        for i in range(stamp.shape[0]):
            for j in range(stamp.shape[1]):
                if not np.array_equal(stamp[i][j], [255, 255, 255]):
                    if i > stamp_height:
                        stamp_height = i
                    if j > stamp_widht:
                        stamp_widht = j
                    pixmap[(i, j)] = stamp[i][j]
        with open(stamp_name, 'wb') as handle:
            stamp = {
                "pixmap": pixmap,
                "height": stamp_height,
                "width": stamp_widht
            }
            pickle.dump(stamp, handle, protocol=pickle.HIGHEST_PROTOCOL)
            return stamp
    else:
        with open(stamp_name, 'rb') as handle:
            return pickle.load(handle)


def get_stamp(stamp_name, pdfPath=""):
    print(stamp_name, pdfPath)
    create_stamp = False
    pixmap = {}
    stamp_height = 0
    stamp_widht = 0
    if create_stamp:
        pages = convert_from_path(pdfPath, 200)
        stamp = np.array(pages[0])
        img = cv2.cvtColor(stamp, cv2.COLOR_BGR2GRAY)
        ret, binarized_image = cv2.threshold(img, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
        for i in range(binarized_image.shape[0]):
            for j in range(binarized_image.shape[1]):
                if binarized_image[i][j] != 255:
                    if i > stamp_height:
                        stamp_height = i
                    if j > stamp_widht:
                        stamp_widht = j
                    pixmap[(i, j)] = stamp[i][j]
        with open(stamp_name, 'wb') as handle:
            stamp = {
                "pixmap": pixmap,
                "height": stamp_height,
                "width": stamp_widht
            }
            pickle.dump(stamp, handle, protocol=pickle.HIGHEST_PROTOCOL)
            return stamp
    else:
        with open(stamp_name, 'rb') as handle:
            return pickle.load(handle)


def func(request_json):
    test_engineer_dict_for_text_pdf = {
        "Ankit Kumar": "Ankit Kumar.png",
        "Aviral mishra": "Aviral.png",
        "Avishek Kumar": "Avishek.png",
        "Gaurav Kumar": "GauravGoswami.png",
        "Kajal Jha": "KajalJha.png",
        "Kaushal Kumar": "Kaushal.png",
        "Mohit Singh": "Mohit.png",
        "Parth": "Parth Singh.png",
        "Tushant Rajvanshi": "tushantSign.png"
    }
    test_engineer_name = request_json.form["test_engineer_name"]
    report_file_name = request_json.form["report_file_name"]
    approving_authority = request_json.form["approving_authority"]
    word_file = request_json.files["report_docx"]
    print(test_engineer_name, word_file, report_file_name)
    document_name = os.path.splitext(report_file_name)[0]
    save_document_path = CONFIG["tempFolder"] + str(document_name)
    os.makedirs(save_document_path)
    in_file = os.path.join(save_document_path, document_name) + '.docx'
    word_file.save(in_file)
    out_file = os.path.join(save_document_path, document_name) + '.pdf'
    img_pdf_path = os.path.join(save_document_path, document_name) + '_img.pdf'
    doc_to_pdf(in_file, out_file)
    status_repost = {"ReportName": document_name, "UploadDate": datetime.datetime.now().strftime("%m/%d/%Y, %H:%M:%S"),
                     "TestEngineerName": test_engineer_name,
                     "ApprovedSatatus1": "unapproved", "ApprovedSatatus2": "unapproved",
                     "ApprovedSatatus3": "unapproved",
                     "ReportApprovedSatatus": "unapproved",
                     "RejectReportMessage": ""}
    report_data = ReportsData.ReportsData()
    pd.to_pickle(report_data.REPORT_FILE_DATA_FRAME.append(status_repost, ignore_index=True), "reportFileDataframe.pkl")
    status_repost = json.dumps(status_repost)
    write_report_in_dir(document_name, status_repost)
    test_engineer_sign = os.path.join("signs",test_engineer_dict_for_text_pdf[test_engineer_name])
    make_text_pdf_with_watermark(out_file, img_pdf_path, test_engineer_name, test_engineer_sign, approving_authority)    # get_stamp(test_engineer_dict["Zahid Raza"], r"C:\Users\aditya.verma\Desktop\ZahidSign.pdf")
    mail_data = {
        "to": CONFIG["MailTo"],
        "Subject": f"New Report Uploaded {document_name}",
        "body": f"Hi <br/> <b> {test_engineer_name} </b> has uploaded new test report named <b> {document_name} </b> please view it and take necessary action. <br/> Click here to view the report {CONFIG['AppURL']}",
    }
    # mailService = MailService()
    # mailServicelService.send_mail(mail_data)

if __name__ =="__main__":
    out_file = "D:/temp1/AU2-220204-001Report/AU2-220204-001Report.pdf"
    img_pdf_path = "test_img.pdf"
    test_engineer_name = "Parth"
    test_engineer_dict_for_text_pdf = {
        "Ankit Kumar": "Ankit Kumar.png",
        "Aviral mishra": "Aviral.png",
        "Avishek Kumar": "Avishek.png",
        "Gaurav Kumar": "GauravGoswami.png",
        "Kajal Jha": "KajalJha.png",
        "Kaushal Kumar": "Kaushal.png",
        "Mohit Singh": "Mohit.png",
        "Parth": "Parth Singh.png",
        "Tushant Rajvanshi": "tushantSign.png"
    }
    approving_authority = "Avishek"
    # pdf_to_image_pdf(out_file, img_pdf_path, test_engineer_name, test_engineer_dict, approving_authority)
    test_engineer_sign = os.path.join("signs",test_engineer_dict_for_text_pdf[test_engineer_name])
    make_text_pdf_with_watermark(out_file, img_pdf_path, test_engineer_name, test_engineer_sign, approving_authority)
