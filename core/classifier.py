"""
core/classifier.py
Classification des documents par détection de mots-clés.
- Analyse uniquement la PREMIÈRE PAGE pour la classification
- Mots-clés chargés depuis keywords.json (éditable par l'utilisateur)
- Mots-clés utilisateur (mappages DB) prioritaires
"""

import re
import json
import os
import unicodedata
from db.database import Database


def normalize(s: str) -> str:
    """Normalise une chaine : minuscules + suppression des accents.
    Permet une comparaison insensible a la casse ET aux accents.
    Ex: 'Société' -> 'societe', 'EDF' -> 'edf'
    """
    return unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("ascii").lower()

KEYWORDS_FILE = os.path.join(os.path.dirname(os.path.dirname(__file__)), "keywords.json")
TYPES_FILE = os.path.join(os.path.dirname(os.path.dirname(__file__)), "document_types.json")


def load_keywords() -> dict:
    """Charge les mots-clés depuis keywords.json"""
    if not os.path.exists(KEYWORDS_FILE):
        return {}
    try:
        with open(KEYWORDS_FILE, encoding="utf-8") as f:
            data = json.load(f)
        # Aplatir toutes les catégories en un seul dict {keyword_normalise: folder}
        flat = {}
        for category, entries in data.items():
            if category == "_notice":
                continue
            if isinstance(entries, dict):
                for kw, folder in entries.items():
                    flat[normalize(kw)] = folder
        return flat
    except Exception as e:
        print(f"[WARN] Impossible de charger keywords.json : {e}")
        return {}


class Classifier:
    def __init__(self, db: Database):
        self.db = db
        self.default_keywords = load_keywords()

    def reload_keywords(self):
        """Recharger keywords.json (utile après modification par l'utilisateur)"""
        self.default_keywords = load_keywords()

    def extract_first_page_text(self, full_text: str) -> str:
        """
        Extrait le texte de la première page uniquement.
        Les pages sont séparées par un saut de page (\\f) ou
        on prend les 3000 premiers caractères si pas de séparateur.
        """
        if "\f" in full_text:
            return full_text.split("\f")[0]
        # Fallback : prendre les 3000 premiers caractères (~1 page)
        return full_text[:3000]

    def classify(self, text: str, first_page_only: bool = True) -> list[dict]:
        """
        Analyse le texte et retourne une liste de correspondances
        triées par score décroissant.
        Chaque entrée : {'keyword': ..., 'folder': ..., 'score': ...}

        first_page_only=True : classification sur la 1ère page uniquement
        """
        # Utiliser uniquement la première page pour la classification
        analysis_text = self.extract_first_page_text(text) if first_page_only else text
        text_norm = normalize(analysis_text)  # texte normalisé (sans accents, minuscules)

        scores = {}

        # 1. Mappages utilisateur (priorité maximale, score x10)
        user_mappings = self.db.get_all_mappings()
        for mapping in user_mappings:
            kw = normalize(mapping["keyword"])
            if kw in text_norm:
                count = text_norm.count(kw)
                scores[kw] = {
                    "keyword": kw,
                    "folder": mapping["target_folder"],
                    "score": count * 10,
                    "source": "user",
                }

        # 2. Mots-clés par défaut (keywords.json, score x5)
        # Trier par longueur décroissante : les mots-clés les plus longs
        # sont plus spécifiques et doivent être testés en premier
        sorted_kws = sorted(self.default_keywords.items(),
                            key=lambda x: len(x[0]), reverse=True)

        for kw, folder in sorted_kws:
            if kw not in scores and kw in text_norm:
                count = text_norm.count(kw)
                scores[kw] = {
                    "keyword": kw,
                    "folder": folder,
                    "score": count * 5,
                    "source": "default",
                }

        # Trier par score décroissant
        results = sorted(scores.values(), key=lambda x: x["score"], reverse=True)
        return results

    def top_suggestion(self, text: str):
        """Retourne la meilleure suggestion ou None"""
        results = self.classify(text)
        return results[0] if results else None

    def extract_metadata(self, text: str) -> dict:
        """
        Extraire des métadonnées simples du texte :
        dates, montants, numéros de documents
        (appliqué sur le texte complet, pas seulement la 1ère page)
        """
        meta = {}

        # Dates (formats courants FR)
        dates = re.findall(
            r'\b(\d{1,2}[/\-\.]\d{1,2}[/\-\.]\d{2,4})\b', text
        )
        if dates:
            meta["dates"] = dates[:5]

        # Montants (€)
        montants = re.findall(
            r'(\d[\d\s]*[,\.]\d{2})\s*€|€\s*(\d[\d\s]*[,\.]\d{2})', text
        )
        if montants:
            meta["montants"] = [m[0] or m[1] for m in montants[:5]]

        # Numéro de facture / contrat
        numeros = re.findall(
            r'(?:n°|num[eé]ro|facture|contrat|r[eé]f)[.\s:]*([A-Z0-9\-]{4,20})',
            text, re.IGNORECASE
        )
        if numeros:
            meta["numeros"] = numeros[:3]

        return meta

    def detect_type(self, text: str) -> str:
        """
        Detecte le type de document (Facture, Contrat, Releve, etc.)
        a partir des mots-cles de detection_keywords dans document_types.json.
        Retourne le type detecte ou "Autre" par defaut.
        """
        if not os.path.exists(TYPES_FILE):
            return "Autre"
        try:
            with open(TYPES_FILE, encoding="utf-8") as f:
                data = json.load(f)
            detection = data.get("detection_keywords", {})
            first_page = normalize(self.extract_first_page_text(text))
            for doc_type, keywords in detection.items():
                for kw in keywords:
                    if normalize(kw) in first_page:
                        return doc_type
        except Exception:
            pass
        return "Autre"

