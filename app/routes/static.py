from flask import Blueprint, send_from_directory

static_bp = Blueprint('static', __name__)


@static_bp.route('/static/processed_files/<filename>')
def serve_processed_file(filename):
    return send_from_directory('app/static/processed_files', filename)