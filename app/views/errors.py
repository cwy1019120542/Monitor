from flask import Blueprint
from ..response import response

errors_blueprint = Blueprint('errors', __name__)

@errors_blueprint.app_errorhandler(400)
def error_400(e):
    return response(False, 400, "bad request")

@errors_blueprint.app_errorhandler(404)
def error_404(e):
    return response(False, 404, "not found")

@errors_blueprint.app_errorhandler(405)
def error_405(e):
    return response(False, 405, "method not allowed")

@errors_blueprint.app_errorhandler(500)
def error_500(e):
    return response(False, 500, "server error")
