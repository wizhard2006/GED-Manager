"""
core/classifier.py
Classification des documents par détection de mots-clés et d'origines.
Utilise les mappages enregistrés + une liste de mots-clés par défaut.
"""

import re
from db.database import Database


# Mots-clés de référence par catégorie (exemples de départ, extensibles)
DEFAULT_KEYWORDS = {
    # Énergie / Utilities
    "engie": "Energie/Engie",
    "edf": "Energie/EDF",
    "eni gaz": "Energie/ENI",
    "direct energie": "Energie/DirectEnergie",
    "total energies": "Energie/TotalEnergies",
    "gaz de france": "Energie/GDF",
    "saur": "Eau/SAUR",
    "veolia eau": "Eau/Veolia",
    "suez": "Eau/Suez",

    # Télécoms
    "orange": "Telecom/Orange",
    "sfr": "Telecom/SFR",
    "bouygues telecom": "Telecom/Bouygues",
    "free": "Telecom/Free",
    "la poste mobile": "Telecom/LaPosteMobile",
    "sosh": "Telecom/Sosh",

    # Banque
    "credit agricole": "Banque/CreditAgricole",
    "bnp paribas": "Banque/BNP",
    "bnp": "Banque/BNP",
    "societe generale": "Banque/SocieteGenerale",
    "lcl": "Banque/LCL",
    "caisse d'epargne": "Banque/CaisseEpargne",
    "banque postale": "Banque/BanquePostale",
    "credit mutuel": "Banque/CreditMutuel",
    "boursorama": "Banque/Boursorama",
    "hello bank": "Banque/HelloBank",
    "ing": "Banque/ING",

    # Assurance
    "maif": "Assurance/MAIF",
    "macif": "Assurance/MACIF",
    "matmut": "Assurance/MATMUT",
    "axa": "Assurance/AXA",
    "allianz": "Assurance/Allianz",
    "groupama": "Assurance/Groupama",
    "covea": "Assurance/Covea",
    "generali": "Assurance/Generali",
    "mma": "Assurance/MMA",
    "maaf": "Assurance/MAAF",

    # Automobile
    "citroen": "Automobile/Citroen",
    "renault": "Automobile/Renault",
    "peugeot": "Automobile/Peugeot",
    "volkswagen": "Automobile/Volkswagen",
    "toyota": "Automobile/Toyota",
    "controle technique": "Automobile/ControleTechnique",
    "dekra": "Automobile/ControleTechnique",
    "autovision": "Automobile/ControleTechnique",
    "carglass": "Automobile/Carglass",
    "norauto": "Automobile/Norauto",
    "feu vert": "Automobile/FeuVert",
    "midas": "Automobile/Midas",

    # Santé
    "cpam": "Sante/CPAM",
    "ameli": "Sante/CPAM",
    "secu": "Sante/SecuriteSociale",
    "mutuelle": "Sante/Mutuelle",
    "mgen": "Sante/MGEN",
    "harmonie mutuelle": "Sante/HarmonieMutuelle",
    "alan": "Sante/Alan",
    "malakoff": "Sante/MalakoffHumanis",

    # Impôts / Admin
    "direction generale des finances": "Admin/Impots",
    "impots.gouv": "Admin/Impots",
    "tresor public": "Admin/Impots",
    "prefecture": "Admin/Prefecture",
    "caf": "Admin/CAF",
    "pole emploi": "Admin/PoleEmploi",
    "france travail": "Admin/FranceTravail",
    "urssaf": "Admin/URSSAF",
    "rsi": "Admin/RSI",

    # Logement
    "edf habitation": "Logement/EDF",
    "appartement": "Logement/Divers",
    "bail": "Logement/Bail",
    "syndic": "Logement/Syndic",
    "notaire": "Logement/Notaire",

    # Grande distribution / Divers
    "amazon": "Achats/Amazon",
    "ebay": "Achats/Ebay",
    "fnac": "Achats/Fnac",
    "darty": "Achats/Darty",
    "leroy merlin": "Achats/LeroyMerlin",
}


class Classifier:
    def __init__(self, db: Database):
        self.db = db

    def classify(self, text: str) -> list[dict]:
        """
        Analyse le texte et retourne une liste de correspondances triées
        par score décroissant.
        Chaque entrée : {'keyword': ..., 'folder': ..., 'score': ...}
        """
        text_lower = text.lower()
        scores = {}

        # 1. Chercher dans les mappages utilisateur (priorité maximale)
        user_mappings = self.db.get_all_mappings()
        for mapping in user_mappings:
            kw = mapping["keyword"].lower()
            if kw in text_lower:
                count = text_lower.count(kw)
                scores[kw] = {
                    "keyword": kw,
                    "folder": mapping["target_folder"],
                    "score": count * 10,  # priorité haute
                    "source": "user",
                }

        # 2. Chercher dans les mots-clés par défaut
        for kw, folder in DEFAULT_KEYWORDS.items():
            if kw.lower() in text_lower:
                if kw not in scores:
                    count = text_lower.count(kw.lower())
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
        """
        meta = {}

        # Dates (formats courants FR)
        dates = re.findall(
            r'\b(\d{1,2}[/\-\.]\d{1,2}[/\-\.]\d{2,4})\b', text
        )
        if dates:
            meta["dates"] = dates[:5]  # max 5

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
