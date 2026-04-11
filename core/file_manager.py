"""
core/file_manager.py
Gestion des fichiers : déplacement, copie, dédoublonnage, nommage
"""

import os
import shutil
import hashlib
from datetime import datetime
from db.database import Database
from utils.logger import Logger


class FileManager:
    def __init__(self, db: Database, logger: Logger, ged_root: str, quarantine_folder: str):
        self.db = db
        self.logger = logger
        self.ged_root = ged_root
        self.quarantine_folder = quarantine_folder

    def compute_hash(self, file_path: str) -> str:
        """Calcule le hash SHA256 d'un fichier pour dédoublonnage"""
        hasher = hashlib.sha256()
        with open(file_path, "rb") as f:
            for chunk in iter(lambda: f.read(65536), b""):
                hasher.update(chunk)
        return hasher.hexdigest()

    def is_duplicate(self, file_path: str) -> bool:
        """Vérifie si ce fichier existe déjà dans l'historique (même hash)"""
        file_hash = self.compute_hash(file_path)
        history = self.db.get_history(limit=5000)
        for entry in history:
            if entry.get("file_hash") == file_hash:
                return True
        return False

    def build_destination(self, target_folder: str, filename: str) -> str:
        """
        Construit le chemin de destination complet.
        target_folder est relatif à ged_root (ex: 'Automobile/Citroen').
        """
        dest_dir = os.path.join(self.ged_root, target_folder)
        os.makedirs(dest_dir, exist_ok=True)
        return os.path.join(dest_dir, self._safe_filename(filename))

    def build_quarantine_path(self, filename: str) -> str:
        os.makedirs(self.quarantine_folder, exist_ok=True)
        return os.path.join(self.quarantine_folder, self._safe_filename(filename))

    def _safe_filename(self, filename: str) -> str:
        """Évite les conflits de nom en ajoutant un timestamp si nécessaire"""
        base, ext = os.path.splitext(filename)
        return f"{base}_{datetime.now().strftime('%Y%m%d_%H%M%S')}{ext}"

    def move_file(self, source_path: str, target_folder: str,
                  keyword: str = "", detected_keywords: str = "") -> str:
        """
        Déplace le fichier vers target_folder (relatif à ged_root).
        Retourne le chemin de destination.
        """
        filename = os.path.basename(source_path)
        file_hash = self.compute_hash(source_path)
        destination = self.build_destination(target_folder, filename)

        shutil.move(source_path, destination)

        self.db.add_history(
            filename=filename,
            source_path=source_path,
            destination_path=destination,
            detected_keywords=detected_keywords,
            matched_keyword=keyword,
            file_hash=file_hash,
            action="moved",
        )
        self.logger.log_action("moved", filename, source_path, destination, keyword)
        return destination

    def copy_file(self, source_path: str, target_folder: str,
                  keyword: str = "", detected_keywords: str = "") -> str:
        """
        Copie le fichier (utile pour conserver l'original lors d'un import).
        Retourne le chemin de destination.
        """
        filename = os.path.basename(source_path)
        file_hash = self.compute_hash(source_path)
        destination = self.build_destination(target_folder, filename)

        shutil.copy2(source_path, destination)

        self.db.add_history(
            filename=filename,
            source_path=source_path,
            destination_path=destination,
            detected_keywords=detected_keywords,
            matched_keyword=keyword,
            file_hash=file_hash,
            action="copied",
        )
        self.logger.log_action("copied", filename, source_path, destination, keyword)
        return destination

    def send_to_quarantine(self, source_path: str,
                           detected_keywords: str = "", confidence: float = 0.0) -> str:
        """Déplace le fichier vers la quarantaine"""
        filename = os.path.basename(source_path)
        destination = self.build_quarantine_path(filename)

        shutil.move(source_path, destination)

        self.db.add_quarantine(
            filename=filename,
            file_path=destination,
            detected_keywords=detected_keywords,
            confidence=confidence,
        )
        self.db.add_history(
            filename=filename,
            source_path=source_path,
            destination_path=destination,
            detected_keywords=detected_keywords,
            matched_keyword="",
            file_hash=self.compute_hash(destination),
            action="quarantine",
        )
        self.logger.log_action("quarantine", filename, source_path, destination)
        return destination
