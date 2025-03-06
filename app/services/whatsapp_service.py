from typing import Any, Optional, Dict
from datetime import datetime
import requests
from flask import Response
from twilio.twiml.messaging_response import MessagingResponse
import os

from app.config import Config
from app.utils.logger import Logger

logging = Logger().get_logger()


class WhatsAppService:
    """Handles WhatsApp message processing and file handling."""

    @staticmethod
    def process_message(payload: Dict[str, Any]) -> Response:
        """Process incoming WhatsApp messages and handle file uploads if necessary."""
        try:
            message_body = payload.get("Body", "").strip().lower()
            media_url = payload.get("MediaUrl0")
            media_type = payload.get("MediaContentType0")
            user_info = WhatsAppService.get_user_info(payload)

            response = MessagingResponse()

            if "optimize my bids" in message_body:
                return WhatsAppService.handle_file_upload(response, user_info, media_url, media_type, "bids")

            elif "create new ppc campaign" in message_body:
                return WhatsAppService.handle_file_upload(response, user_info, media_url, media_type, "campaigns")

            else:
                response.message("Invalid command. Try 'Optimize my bids' or 'Create new PPC campaign'.")

            return Response(str(response), content_type="application/xml")

        except Exception as e:
            logging.error(f"Error processing WhatsApp message: {str(e)}", exc_info=True)
            response = MessagingResponse()
            response.message("An error occurred. Please try again later.")
            return Response(str(response), content_type="application/xml")

    @staticmethod
    def handle_file_upload(
        response: MessagingResponse, user_info: Dict[str, str], media_url: Optional[str], media_type: Optional[str], folder_type: str
    ) -> Response:
        """Handles file uploads and saves the file dynamically."""
        if not media_url or not media_type:
            response.message("Please upload a valid file.")
            return Response(str(response), content_type="application/xml")

        file_extension = WhatsAppService.get_file_extension(media_type)
        if not file_extension:
            response.message("Unsupported file type. Please upload an Excel file.")
            return Response(str(response), content_type="application/xml")

        file_name = WhatsAppService.generate_file_name(user_info["name"], folder_type, file_extension)
        downloaded_file = WhatsAppService.download_file(media_url, file_name)

        if downloaded_file:
            response.message(f"Your file has been uploaded successfully.")
        else:
            response.message("Failed to download the file. Please try again.")

        return Response(str(response), content_type="application/xml")

    @staticmethod
    def get_user_info(payload: Dict[str, Any]) -> Dict[str, str]:
        """Extract user details from Twilio's payload."""
        return {
            "name": payload.get("ProfileName", "User").replace(" ", "_"),  # Prevents spaces in filenames
            "phone_number": payload.get("From", "unknown"),
        }

    @staticmethod
    def generate_file_name(username: str, folder_type: str, file_extension: str) -> str:
        """Generate a unique filename and ensure the correct directory exists."""
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        directory = f"app/data/uploads/{folder_type}/"
        os.makedirs(directory, exist_ok=True)  # Ensure directory exists
        return os.path.join(directory, f"{username}_{folder_type}_{timestamp}{file_extension}")

    @staticmethod
    def get_file_extension(media_type: str) -> Optional[str]:
        """Extract file extension from media type."""
        extension_map = {
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": ".xlsx",
            "application/vnd.ms-excel": ".xls",
            "text/csv": ".csv",
            "application/pdf": ".pdf"
        }
        return extension_map.get(media_type, None)

    @staticmethod
    def download_file(media_url: str, file_name: str) -> Optional[str]:
        """Download file from Twilio's media URL with authentication."""
        try:
            response = requests.get(media_url, auth=(Config.TWILIO_ACCOUNT_SID, Config.TWILIO_AUTH_TOKEN), timeout=10)
            response.raise_for_status()  # Raises an HTTPError for bad responses

            with open(file_name, "wb") as file:
                file.write(response.content)

            logging.info(f"File downloaded successfully: {file_name}")
            return file_name

        except requests.exceptions.RequestException as e:
            logging.error(f"Error downloading file: {e}", exc_info=True)
            return None
