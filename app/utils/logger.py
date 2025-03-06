import logging
import os

from app.config import Config



class Logger:
    def __init__(self, log_file=Config.LOG_FILE, log_level=Config.LOG_LEVEL):
        self.logger = logging.getLogger("Whatsapp PPC Bot")
        self.logger.setLevel(log_level)

        # Ensure the log directory exists
        os.makedirs(os.path.dirname(log_file), exist_ok=True)

        # Formatter for logs
        formatter = logging.Formatter(
            "%(asctime)s - %(name)s - %(levelname)s - %(message)s"
        )

        # **1. Stream Handler (Logs to Console)**
        stream_handler = logging.StreamHandler()
        stream_handler.setLevel(log_level)  # Show all messages in console
        stream_handler.setFormatter(formatter)

        # **2. File Handler (Logs to `app.log`)**
        file_handler = logging.FileHandler(log_file)
        file_handler.setLevel(log_level)  # Save only INFO+ messages to file
        file_handler.setFormatter(formatter)

        # Add handlers to logger
        self.logger.addHandler(stream_handler)
        self.logger.addHandler(file_handler)

    def get_logger(self):
        """Return the configured logger instance."""
        return self.logger