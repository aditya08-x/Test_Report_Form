import ast
import threading
import json
from time import time
from flask_restplus import fields
from createTestReportPackages.api import api_response
from createTestReportPackages.model.MailService import MailService
from createTestReportPackages.parser import CONFIG
from createTestReportPackages.utils import platform_exception as exp
from createTestReportPackages.utils.helper_utilities import write_report_in_dir
from createTestReportPackages.utils.platform_logging import tagging_logger
from createTestReportPackages.pipelines import create_docx_from_data
from createTestReportPackages.model import ReportsData

URL = '/test-report-generator/update-status'

request_fields = {
    "report_name": fields.String('Name of the report to be updated', required=True),
    "access_code": fields.String('access code of the authority', required=True),
    "authority_name": fields.String('Name of the approving authority', required=True),
    "update_type": fields.String('Accept or Reject', required=True),
    "reject_message": fields.String('Reason for Rejection of report', required=True)
}


def func(request_json):
    """
    API entrypoint of upload-image

    :param request_json: api payload json
    :return:text response
    """
    start_time = time()
    try:
        tagging_logger.info(f"API START : UPLOAD IMAGE :: TEMPLATE_ID - ")
        report_name = request_json["report_name"]
        access_code = request_json["access_code"]
        authority_name = request_json["authority_name"]
        update_type = request_json["update_type"]
        reject_message = request_json["reject_message"]
        authority = {
            'ApprovedSatatus1': 'Testing Engineer',
            'ApprovedSatatus2': 'Head of Department',
            'ApprovedSatatus3': 'Business Development Head'
        }
        if authority_name == "ApprovedSatatus1" and access_code == "1234" or authority_name == "ApprovedSatatus2" and access_code == "3456" or authority_name == "ApprovedSatatus3" and access_code == "7910":
            report_data = ReportsData.ReportsData()
            if update_type == "Accept":
                report_data.REPORT_FILE_DATA_FRAME.loc[
                    report_data.REPORT_FILE_DATA_FRAME.ReportName == report_name, authority_name] = "Approved"

                new_row = report_data.REPORT_FILE_DATA_FRAME.loc[
                    report_data.REPORT_FILE_DATA_FRAME.ReportName == report_name].reset_index()
                ready_to_download = ""
                if new_row['ApprovedSatatus1'][0] == new_row['ApprovedSatatus2'][0] == new_row['ApprovedSatatus3'][
                    0] == "Approved":
                    report_data.REPORT_FILE_DATA_FRAME.loc[
                        report_data.REPORT_FILE_DATA_FRAME.ReportName == report_name, "ReportApprovedSatatus"] = "Approved"
                    ready_to_download = "And the report is ready to download"
                mail_data = {
                    "to": CONFIG["MailTo"],
                    "Subject": f"Report accepted by {authority[authority_name]}",
                    "body": f"Hi <br> <b> {authority[authority_name]} </b> has approved the report <b> {report_name} </b> " + ready_to_download,
                }
            else:
                report_data.REPORT_FILE_DATA_FRAME.loc[
                    report_data.REPORT_FILE_DATA_FRAME.ReportName == report_name, authority_name] = "Rejected"
                report_data.REPORT_FILE_DATA_FRAME.loc[
                    report_data.REPORT_FILE_DATA_FRAME.ReportName == report_name, "ReportApprovedSatatus"] = "Rejected"
                report_data.REPORT_FILE_DATA_FRAME.loc[
                    report_data.REPORT_FILE_DATA_FRAME.ReportName == report_name, "RejectReportMessage"] = reject_message
                mail_data = {
                    "to": CONFIG["MailTo"],
                    "Subject": f"Report Rejected by {authority[authority_name]}",
                    "body": f"Hi <br> <b> {authority[authority_name]} </b> has Rejected the report <b> {report_name} </b> <br/> the reason for rejection is {reject_message} ",
                }
            mailService = MailService()
            mailService.send_mail(mail_data)
            new_row = report_data.REPORT_FILE_DATA_FRAME.loc[
                report_data.REPORT_FILE_DATA_FRAME.ReportName == report_name].reset_index()
            import pandas as pd
            pd.to_pickle(report_data.REPORT_FILE_DATA_FRAME,
                         "reportFileDataframe.pkl")
            status_report = json.dumps(json.loads(new_row.to_json(orient='records'))[0])
            write_report_in_dir(report_name, status_report)
            response = api_response.text_response("Status Updated", api_response.HTTP_SUCCESS_STATUS_CODE)
        else:
            response = api_response.text_response("Wrong Access Code", api_response.HTTP_SUCCESS_STATUS_CODE)
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
