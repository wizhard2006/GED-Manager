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
    "classification": {
        # Seuil de confiance pour le mode hybride (0 = désactivé, validation toujours demandée)
        "auto_classify_threshold": "0",
    },
    "ocr": {
        # PSM 6 = bloc de texte uniforme (mieux pour factures/courriers)
        # Valeurs possibles : 3 (auto), 4 (colonne unique), 6 (bloc uniforme)
        # Pour revenir à l'ancien comportement : mettre 3
        "tesseract_psm": "6",
        # Prétraitement OpenCV : débruitage + binarisation + deskew
        # false = désactivé (comportement original garanti)
        # true  = activé (tester sur quelques docs avant de valider)
        "enhanced_preprocessing": "false",
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

    @property
    def tesseract_psm(self):
        try:
            return int(self.get("ocr", "tesseract_psm"))
        except (ValueError, TypeError):
            return 6

    @property
    def enhanced_preprocessing(self):
        return self.get("ocr", "enhanced_preprocessing").lower() == "true"

    @property
    def auto_classify_threshold(self):
        """Seuil de confiance pour classement auto en mode masse.
        0 = toujours demander validation (Mode 2 pur).
        Valeur typique quand actifé : 10-20.
        """
        try:
            return int(self.get("classification", "auto_classify_threshold"))
        except (ValueError, TypeError):
            return 0
