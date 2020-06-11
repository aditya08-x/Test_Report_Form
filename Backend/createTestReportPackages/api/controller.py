from flask import Flask, Blueprint, request
from flask_restplus import Resource, Api
from flask_cors import CORS
from werkzeug.datastructures import FileStorage

from createTestReportPackages.api.api_response import HTTP_SUCCESS_STATUS_CODE, HTTP_FAILURE_STATUS_CODE
from createTestReportPackages.api.services import health_check, create_report, create_report_from_docx, download_report, \
    get_report_list, update_status

app = Flask(__name__)
CORS(app)
blueprint = Blueprint('Test Report Generator API', __name__)
api = Api(blueprint,
          title='Test Report Generator API',
          description="Test Report Generator REST API",
          version="1.0.0",
          default="Test Report Generator",
          default_label="",
          doc="/test-report-generator/swagger",
          validate=True)

app.register_blueprint(blueprint)

create_report_model = api.model('create_report_model', create_report.request_fields)
download_report_model = api.model('download_report_model', download_report.request_fields)
update_status_model = api.model('update_status_model', update_status.request_fields)
create_report_from_docx_model = api.model('create_report_from_docx_model', create_report_from_docx.request_fields)


@api.route(create_report.URL)
class CreateReport(Resource):
    @api.expect(create_report_model)
    @api.response(HTTP_SUCCESS_STATUS_CODE, 'Success')
    @api.response(HTTP_FAILURE_STATUS_CODE, 'Internal Server Error Message')
    def post(self):
        response = create_report.func(api.payload)
        return response


@api.route(get_report_list.URL)
class GetReportList(Resource):
    @api.response(HTTP_SUCCESS_STATUS_CODE, 'Success')
    @api.response(HTTP_FAILURE_STATUS_CODE, 'Internal Server Error Message')
    def get(self):
        response = get_report_list.func()
        return response


parser = api.parser()
parser.add_argument('test_engineer_name', type=str, help='Some param', location='form')
parser.add_argument('report_docx', type=FileStorage, location='files')


@api.route(create_report_from_docx.URL)
class CreateReportFromDocx(Resource):
    @api.expect(parser)
    @api.response(HTTP_SUCCESS_STATUS_CODE, 'Success')
    @api.response(HTTP_FAILURE_STATUS_CODE, 'Internal Server Error Message')
    def post(self):
        response = create_report_from_docx.func(request)
        return response


@api.route(download_report.URL)
class DownloadReport(Resource):
    @api.expect(download_report_model)
    @api.response(HTTP_SUCCESS_STATUS_CODE, 'Success')
    @api.response(HTTP_FAILURE_STATUS_CODE, 'Internal Server Error Message')
    def post(self):
        response = download_report.func(api.payload)
        return response


@api.route(update_status.URL)
class UpdateStatus(Resource):
    @api.expect(update_status_model)
    @api.response(HTTP_SUCCESS_STATUS_CODE, 'Success')
    @api.response(HTTP_FAILURE_STATUS_CODE, 'Internal Server Error Message')
    def post(self):
        response = update_status.func(api.payload)
        return response


@api.route(health_check.URL)
class HealthCheck(Resource):
    @api.response(HTTP_SUCCESS_STATUS_CODE, 'Return List of Env')
    @api.response(HTTP_FAILURE_STATUS_CODE, 'Internal Server Error Message')
    def get(self):
        response = health_check.func()
        return response
