"""
utils/config.py
Gestion de la configuration de l'application GED-Manager
"""

import configparser
import os

CONFIG_FILE = "ged_config.ini"

DEFAULTS = {
    "paths": {
        "ged_root": "D:\\",
        "quarantine_folder": "D:\\_A_Classer",
        "tesseract_path": r"C:\Program Files\Tesseract-OCR\tesseract.exe",
        "scanner_output_folder": "",
    },
    "scanner": {
        "dpi": "300",
        "color_mode": "Color",
        "output_format": "PDF",
    },
    "app": {
        "log_file": "ged_manager.log",
        "db_file": "ged_manager.db",
        "language": "fra+eng",
    },
}


class Config:
    def __init__(self):
        self.config = configparser.ConfigParser()
        self._load()

    def _load(self):
        if os.path.exists(CONFIG_FILE):
            self.config.read(CONFIG_FILE, encoding="utf-8")
        else:
            # Créer avec les valeurs par défaut
            for section, values in DEFAULTS.items():
                self.config[section] = values
            self._save()

    def _save(self):
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            self.config.write(f)

    def get(self, section, key):
        try:
            return self.config[section][key]
        except KeyError:
            return DEFAULTS.get(section, {}).get(key, "")

    def set(self, section, key, value):
        if section not in self.config:
            self.config[section] = {}
        self.config[section][key] = value
        self._save()

    @property
    def ged_root(self):
        return self.get("paths", "ged_root")

    @property
    def quarantine_folder(self):
        return self.get("paths", "quarantine_folder")

    @property
    def tesseract_path(self):
        return self.get("paths", "tesseract_path")

    @property
    def scanner_output_folder(self):
        return self.get("paths", "scanner_output_folder")

    @property
    def db_file(self):
        return self.get("app", "db_file")

    @property
    def log_file(self):
        return self.get("app", "log_file")

    @property
    def ocr_language(self):
        return self.get("app", "language")
