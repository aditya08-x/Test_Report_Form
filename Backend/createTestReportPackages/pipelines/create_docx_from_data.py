from pdf2image import convert_from_path
import os
import shutil
import pythoncom
import win32com.client
import uuid
from createTestReportPackages.parser import CONFIG
from PIL import Image

wdFormatPDF = 17


def copytree(src, dst, symlinks=False, ignore=None):
    for item in os.listdir(src):
        s = os.path.join(src, item)
        d = os.path.join(dst, item)
        if os.path.isdir(s):
            shutil.copytree(s, d, symlinks, ignore)
        else:
            shutil.copy2(s, d)


def replace_placeholders(xml_file_path, form_data):
    print(form_data)
    f = open(xml_file_path, 'r', encoding="UTF-8")
    xml = f.read()
    for key in form_data.keys():
        xml = xml.replace(key, form_data[key])
    f = open(xml_file_path, "w", encoding="UTF-8")
    f.write(xml)
    f.close()


def zip_and_convert_to_docx(copy_template_path, save_document_path):
    shutil.make_archive(os.path.join(save_document_path, CONFIG["saveDocumentName"]), 'zip', copy_template_path)
    os.rename(os.path.join(save_document_path, CONFIG["saveDocumentName"] + '.zip'),
              os.path.join(save_document_path, CONFIG["saveDocumentName"] + '.docx'))


def doc_to_pdf(in_file, out_file):
    pythoncom.CoInitialize()
    word = win32com.client.Dispatch('Word.Application')
    doc = word.Documents.Open(in_file)
    doc.SaveAs(out_file, FileFormat=wdFormatPDF)
    doc.Close()


def pdf_to_image_pdf(out_file, img_pdf_path):
    pages = convert_from_path(out_file, 200)
    pages[0].save(img_pdf_path, save_all=True, append_images=pages[1:])


def create_pdf(form_data):
    template_path = CONFIG["unZippedDocxPath"]
    current_process_temp_path = CONFIG["tempFolder"] + str(uuid.uuid4())
    copy_template_path = current_process_temp_path + CONFIG["copyFolderName"]
    copytree(template_path, copy_template_path)
    word_folder_path_for_template = copy_template_path + '/word'
    for filename in os.listdir(word_folder_path_for_template):
        if not filename.endswith('.xml'): continue
        fullname = os.path.join(word_folder_path_for_template, filename)
        replace_placeholders(fullname, form_data)
    img = Image.open(r"C:\Users\aditya.verma\Desktop\New folder (3)\word\media\image15.png")
    media_folder_path_for_template = word_folder_path_for_template + "/media"
    im1 = img.save(media_folder_path_for_template + "//image1.png")
    save_document_path = os.path.join(current_process_temp_path, 'document')
    zip_and_convert_to_docx(copy_template_path, save_document_path)
    in_file = os.path.join(save_document_path, CONFIG["saveDocumentName"]) + '.docx'
    out_file = os.path.join(save_document_path, CONFIG["saveDocumentName"]) + '.pdf'
    img_pdf_path = os.path.join(save_document_path, CONFIG["saveDocumentName"]) + '_img.pdf'
    doc_to_pdf(in_file, out_file)
    pdf_to_image_pdf(out_file, img_pdf_path)
