from flask import Blueprint, request
from app.services.whatsapp_service import WhatsAppService

whatsapp_bp = Blueprint("whatsapp", __name__)

@whatsapp_bp.route("/", methods=["POST"])
def receive_whatsapp_message():
    payload = request.form
    response = WhatsAppService.process_message(payload)
    return response
