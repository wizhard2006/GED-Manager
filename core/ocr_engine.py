"""
core/ocr_engine.py
Moteur OCR : extraction de texte depuis PDF ou image.
- Résolution 3x (~216 DPI) pour meilleure qualité
- Retourne le texte complet (toutes pages) pour les métadonnées
- Le classifier se charge de n'utiliser que la 1ère page pour classer
- Prétraitement OpenCV optionnel (activable dans ged_config.ini)
"""

import os
import pytesseract
from PIL import Image
import fitz  # PyMuPDF
import numpy as np


def _preprocess_image_cv2(pil_img):
    """
    Prétraitement OpenCV pour améliorer la reconnaissance OCR.
    Opérations effectuées EN MÉMOIRE — le fichier original n'est jamais modifié.

    Étapes :
    1. Conversion en niveaux de gris
    2. Débruitage (filtre bilatéral — préserve les contours)
    3. Binarisation adaptative (seuillage local — gère les variations d'éclairage)
    4. Deskew léger (correction d'inclinaison jusqu'à ±5°)
    """
    try:
        import cv2

        # PIL -> numpy array
        img = np.array(pil_img)

        # 1. Niveaux de gris
        if len(img.shape) == 3:
            gray = cv2.cvtColor(img, cv2.COLOR_RGB2GRAY)
        else:
            gray = img

        # 2. Débruitage (préserve les bords du texte)
        denoised = cv2.bilateralFilter(gray, d=9, sigmaColor=75, sigmaSpace=75)

        # 3. Binarisation adaptative
        binary = cv2.adaptiveThreshold(
            denoised,
            255,
            cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
            cv2.THRESH_BINARY,
            blockSize=31,
            C=10,
        )

        # 4. Deskew : correction inclinaison légère
        coords = np.column_stack(np.where(binary < 128))  # pixels sombres
        if len(coords) > 100:
            angle = cv2.minAreaRect(coords.astype(np.float32))[-1]
            # minAreaRect retourne un angle entre -90 et 0
            if angle < -45:
                angle = 90 + angle
            # Ne corriger que les petites inclinaisons (évite les faux positifs)
            if abs(angle) < 5:
                h, w = binary.shape
                center = (w // 2, h // 2)
                M = cv2.getRotationMatrix2D(center, angle, 1.0)
                binary = cv2.warpAffine(
                    binary, M, (w, h),
                    flags=cv2.INTER_CUBIC,
                    borderMode=cv2.BORDER_REPLICATE,
                )

        # numpy -> PIL
        return Image.fromarray(binary)

    except Exception as e:
        # En cas d'erreur OpenCV, on retourne l'image originale sans planter
        print(f"[WARN] Prétraitement OpenCV échoué, image originale utilisée : {e}")
        return pil_img


class OCREngine:
    def __init__(self, tesseract_path=None, language="fra+eng",
                 enhanced_preprocessing=False, tesseract_psm=6):
        if tesseract_path and os.path.exists(tesseract_path):
            pytesseract.pytesseract.tesseract_cmd = tesseract_path
        self.language = language
        self.enhanced_preprocessing = enhanced_preprocessing
        self.tesseract_psm = tesseract_psm

    def _ocr_image(self, pil_img) -> str:
        """
        Lance Tesseract sur une image PIL.
        Applique le prétraitement OpenCV si enhanced_preprocessing=True.
        Le fichier original n'est jamais modifié — tout se passe en mémoire.
        """
        if self.enhanced_preprocessing:
            pil_img = _preprocess_image_cv2(pil_img)

        config = f"--psm {self.tesseract_psm}"
        return pytesseract.image_to_string(pil_img, lang=self.language, config=config)

    def extract_text_from_pdf(self, pdf_path):
        """
        Extrait le texte d'un PDF.
        Tente d'abord l'extraction native (PDF texte),
        puis bascule sur OCR si le texte est vide/insuffisant.
        Les pages sont séparées par \\f pour permettre l'extraction
        de la première page uniquement par le classifier.
        """
        text = self._extract_native(pdf_path)
        if len(text.strip()) > 50:
            return text
        # Fallback OCR sur images
        return self._extract_ocr(pdf_path)

    def _extract_native(self, pdf_path):
        """Extraire le texte natif du PDF (PDF texte, non image)"""
        try:
            doc = fitz.open(pdf_path)
            pages = []
            for page in doc:
                pages.append(page.get_text())
            doc.close()
            # Séparer les pages par \f (form feed) pour extraction 1ère page
            return "\f".join(pages)
        except Exception:
            return ""

    def _extract_ocr(self, pdf_path):
        """
        Extraire le texte via OCR (PDF image / scanné).
        Résolution 3x (~216 DPI) pour meilleure qualité de reconnaissance.
        """
        try:
            doc = fitz.open(pdf_path)
            pages = []
            for page_num in range(len(doc)):
                page = doc[page_num]
                # 3x = ~216 DPI — meilleure qualité, légèrement plus lent
                mat = fitz.Matrix(3.0, 3.0)
                pix = page.get_pixmap(matrix=mat)
                img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                text = self._ocr_image(img)
                pages.append(text)
            doc.close()
            # Séparer les pages par \f pour extraction 1ère page
            return "\f".join(pages)
        except Exception as e:
            return f"[ERREUR OCR] {str(e)}"

    def extract_text_from_image(self, image_path):
        """Extraire le texte d'une image"""
        try:
            img = Image.open(image_path)
            return self._ocr_image(img)
        except Exception as e:
            return f"[ERREUR OCR image] {str(e)}"

    def extract_text(self, file_path):
        """Point d'entrée universel"""
        ext = os.path.splitext(file_path)[1].lower()
        if ext == ".pdf":
            return self.extract_text_from_pdf(file_path)
        elif ext in [".jpg", ".jpeg", ".png", ".tiff", ".bmp"]:
            return self.extract_text_from_image(file_path)
        else:
            return ""
