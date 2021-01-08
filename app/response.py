from flask import jsonify

def response(is_success, status_code, message, result=None):
    json_data = {
        "is_success" : is_success,
        "message": message,
        "result": result
    }
    response_obj = jsonify(json_data)
    response_obj.status_code = status_code
    return response_obj

def parameter_error(parameter, message=None):
    if not message:
        message = f"parameter error: {parameter}"
    return response(False, 400, message)

def success(result=None):
    return response(True, 200, "success", result)

def recognize_error():
    return response(False, 400, "recognize fail")