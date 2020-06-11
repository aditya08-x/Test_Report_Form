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
from createTestReportPackages.model import ReportsData

URL = '/test-report-generator/get-report-list'


def func():
    """
    API entrypoint of upload-image

    :param request_json: api payload json
    :return:text response
    """
    start_time = time()
    try:
        tagging_logger.info(f"API START : UPLOAD IMAGE :: TEMPLATE_ID - ")
        report_data = ReportsData.ReportsData()
        response_data = report_data.REPORT_FILE_DATA_FRAME.to_json(orient='records')
        # response_data = [
        #     {
        #         "ReportName": "document1",
        #         "UploadDate": "10/11/2109",
        #         "TestEngineerName": "kaushal",
        #         "ApprovedSatatus1": "Approved",
        #         "ApprovedSatatus2": "NoAction",
        #         "ApprovedSatatus3": "Approved",
        #         "ReportApprovedSatatus": "NotApproved"
        #     }]
        response = api_response.json_response(response_data, api_response.HTTP_SUCCESS_STATUS_CODE)
        end_time = time()
        tagging_logger.info(
            f"API SUCCESSFUL :: Total Time - {end_time - start_time}")
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
