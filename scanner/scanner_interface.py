"""
scanner/scanner_interface.py
Interface scanner pour CANON DR-C125 (Windows).
Stratégie 1 : pilotage direct via WIA (Windows Image Acquisition)
Stratégie 2 : surveillance d'un dossier de sortie scanner
"""

import os
import sys
import time
import glob
import subprocess
from utils.logger import Logger


class ScannerInterface:
    """
    Gestion de l'acquisition scanner sous Windows.
    Deux modes supportés :
      - WIA  : pilotage direct du scanner (Option A)
      - WATCH: surveillance dossier de sortie (Option B fallback)
    """

    def __init__(self, logger: Logger, output_folder: str = ""):
        self.logger = logger
        self.output_folder = output_folder

    # ── OPTION A : WIA (Windows Image Acquisition) ───────────────────────────

    def scan_via_wia(self) -> str | None:
        """
        Lance un scan via WIA (Windows) et retourne le chemin du PDF généré.
        Nécessite le module pywin32 (win32com).
        """
        try:
            import win32com.client
            wia = win32com.client.Dispatch("WIA.CommonDialog")
            device = wia.ShowSelectDevice()
            if device is None:
                self.logger.warning("Aucun scanner sélectionné.")
                return None

            image = wia.ShowAcquireImage(
                Device=device,
                DeviceType=1,          # Scanner
                Intent=4,             # Couleur
                Bias=0,
                AlwaysSelectDevice=False,
                UseCommonUI=True,
                CancelError=True
            )
            if image is None:
                return None

            # Sauvegarder en TIFF puis convertir en PDF
            temp_tiff = os.path.join(os.environ.get("TEMP", "."), "ged_scan_tmp.tif")
            image.SaveFile(temp_tiff)
            pdf_path = self._tiff_to_pdf(temp_tiff)
            self.logger.info(f"Scan WIA réussi → {pdf_path}")
            return pdf_path

        except ImportError:
            self.logger.warning("pywin32 non disponible, impossible d'utiliser WIA.")
            return None
        except Exception as e:
            self.logger.error(f"Erreur WIA : {e}")
            return None

    def _tiff_to_pdf(self, tiff_path: str) -> str:
        """Convertit un TIFF en PDF via Pillow"""
        from PIL import Image
        pdf_path = tiff_path.replace(".tif", ".pdf")
        img = Image.open(tiff_path)
        img.save(pdf_path, "PDF", resolution=300)
        return pdf_path

    # ── OPTION B : Surveillance dossier de sortie scanner ────────────────────

    def watch_for_new_scan(self, folder: str, timeout: int = 120) -> str | None:
        """
        Surveille un dossier pendant `timeout` secondes.
        Retourne le chemin du nouveau fichier PDF/image créé, ou None.
        """
        if not folder or not os.path.isdir(folder):
            self.logger.error(f"Dossier de sortie invalide : {folder}")
            return None

        before = set(glob.glob(os.path.join(folder, "*")))
        self.logger.info(f"Surveillance du dossier : {folder} (timeout {timeout}s)...")

        start = time.time()
        while time.time() - start < timeout:
            time.sleep(1)
            after = set(glob.glob(os.path.join(folder, "*")))
            new_files = after - before
            pdf_or_img = [
                f for f in new_files
                if f.lower().endswith((".pdf", ".jpg", ".jpeg", ".tif", ".png"))
            ]
            if pdf_or_img:
                new_file = sorted(pdf_or_img)[-1]
                self.logger.success(f"Nouveau fichier détecté : {new_file}")
                return new_file

        self.logger.warning("Timeout dépassé, aucun nouveau fichier détecté.")
        return None

    # ── Point d'entrée unifié ─────────────────────────────────────────────────

    def acquire(self, mode: str = "watch") -> str | None:
        """
        Lance l'acquisition selon le mode :
          - 'wia'   : pilotage direct (Option A)
          - 'watch' : surveillance dossier (Option B)
        """
        if mode == "wia":
            result = self.scan_via_wia()
            if result:
                return result
            # Fallback automatique sur watch
            self.logger.warning("Bascule sur mode surveillance dossier.")

        return self.watch_for_new_scan(self.output_folder)
