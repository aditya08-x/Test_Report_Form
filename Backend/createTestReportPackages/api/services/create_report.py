import ast
import threading
import json
from time import time
from flask_restplus import fields
from createTestReportPackages.api import api_response
from createTestReportPackages.parser import CONFIG
from createTestReportPackages.utils import platform_exception as exp
from createTestReportPackages.utils.platform_logging import tagging_logger
from createTestReportPackages.pipelines import create_docx_from_data

URL = '/test-report-generator/create-report'

request_fields = {
    "form_data": fields.String('key and values from the form', required=True)}


def func(request_json):
    """
    API entrypoint of upload-image

    :param request_json: api payload json
    :return:text response
    """
    start_time = time()
    form_data = request_json['form_data']
    try:
        tagging_logger.info(f"API START : UPLOAD IMAGE :: TEMPLATE_ID - ")
        # form_data = type(ast.literal_eval(form_data))
        form_data = json.loads(form_data.replace("'",'"'))
        create_docx_from_data.create_pdf(form_data)

        # ocr_thread = threading.Thread(target=create_docx_from_data.create_pdf,
        #                               args=(form_data))
        # ocr_thread.start()
        response = api_response.text_response("Successfully uploaded image", api_response.HTTP_SUCCESS_STATUS_CODE)
        end_time = time()
        tagging_logger.info(
            f"API SUCCESSFUL : UPLOAD IMAGE :: TEMPLATE_ID -  : Total Time - {end_time - start_time}")
        return response
    except exp.PlatformException as e:
        response = api_response.text_response(e.custom_message, api_response.HTTP_FAILURE_STATUS_CODE)
        end_time = time()
        tagging_logger.exception(
            f"API FAILED : UPLOAD IMAGE :: TEMPLATE_ID -  : ERROR - {e.custom_message} : "
            f"Total Time - {end_time - start_time}")
        return response

    except Exception as e:
        response = api_response.text_response(str(e), api_response.HTTP_FAILURE_STATUS_CODE)
        end_time = time()
        tagging_logger.exception(
            f"API FAILED : UPLOAD IMAGE :: TEMPLATE_ID -  : ERROR - {str(e)} : "
            f"Total Time - {end_time - start_time}")
        return response
