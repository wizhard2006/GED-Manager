"""
utils/logger.py
Système de logs et historique des actions GED-Manager
"""

import os
from datetime import datetime


class Logger:
    def __init__(self, log_file="ged_manager.log"):
        self.log_file = log_file

    def _write(self, level, message):
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        line = f"[{timestamp}] [{level}] {message}"
        with open(self.log_file, "a", encoding="utf-8") as f:
            f.write(line + "\n")
        print(line)

    def info(self, message):
        self._write("INFO", message)

    def error(self, message):
        self._write("ERROR", message)

    def warning(self, message):
        self._write("WARNING", message)

    def success(self, message):
        self._write("SUCCESS", message)

    def log_action(self, action, filename, source, destination, keyword=""):
        message = (
            f"Action={action} | Fichier={filename} | "
            f"De={source} | Vers={destination} | Mot-clé={keyword}"
        )
        self._write("ACTION", message)
