"""
core/ocr_engine.py
Moteur OCR : extraction de texte depuis PDF ou image.
- Résolution 3x (~216 DPI) pour meilleure qualité
- Retourne le texte complet (toutes pages) pour les métadonnées
- Le classifier se charge de n'utiliser que la 1ère page pour classer
"""

import os
import pytesseract
from PIL import Image
import fitz  # PyMuPDF


class OCREngine:
    def __init__(self, tesseract_path=None, language="fra+eng"):
        if tesseract_path and os.path.exists(tesseract_path):
            pytesseract.pytesseract.tesseract_cmd = tesseract_path
        self.language = language

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
                text = pytesseract.image_to_string(img, lang=self.language)
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
            return pytesseract.image_to_string(img, lang=self.language)
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
