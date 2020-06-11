from flask import make_response

HTTP_SUCCESS_STATUS_CODE = 200
HTTP_FAILURE_STATUS_CODE = 500


def text_response(text, status_code):
    """
    Create text response for api

    :param text:
    :param status_code: status code for http response
    :return: text response
    """
    response = make_response(text, status_code)
    response.headers['Content-Type'] = 'text/plain'
    return response


def zip_response(zip_file, status_code):
    """
    Create zip response for api

    :param zip_file: zip file
    :param status_code: status code for http response
    :return: zip response
    """
    response = make_response(zip_file, status_code)
    response.headers['Content-Type'] = 'application/zip'
    return response


def pdf_response(pdf_file, status_code):
    """
    Create image response for api

    :param base64_image: base64 string of image
    :param status_code: status code for http response
    :return: image response
    """
    encoded = "data:application/pdf;base64," + pdf_file
    response = make_response(encoded, status_code)
    response.headers['Content-Type'] = 'application/pdf'
    return response


def image_response(base64_image, status_code):
    """
    Create image response for api

    :param base64_image: base64 string of image
    :param status_code: status code for http response
    :return: image response
    """
    encoded = "data:image/jpeg;base64," + base64_image
    response = make_response(encoded, status_code)
    response.headers['Content-Type'] = 'application/octet-stream'
    return response


def json_response(json, status_code):
    """
    Create json response for api

    :param json: json or dictionary containing data to return
    :param status_code: status code for http response
    :return: json response
    """
    response = make_response(json, status_code)
    response.headers['Content-Type'] = 'application/json'
    return response
