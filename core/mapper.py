"""
core/mapper.py
Système de mappage : associe un mot-clé détecté à un dossier de destination.
Gère l'apprentissage progressif (1ère fois → demande, fois suivantes → auto).
"""

from db.database import Database


class Mapper:
    def __init__(self, db: Database):
        self.db = db

    def get_folder(self, keyword: str):
        """Retourne le dossier associé à ce mot-clé, ou None si inconnu"""
        return self.db.get_mapping(keyword)

    def learn(self, keyword: str, target_folder: str):
        """Enregistrer ou mettre à jour un mappage"""
        self.db.add_mapping(keyword, target_folder)

    def forget(self, keyword: str):
        """Supprimer un mappage"""
        self.db.delete_mapping(keyword)

    def all_mappings(self) -> list[dict]:
        """Retourner tous les mappages connus"""
        return self.db.get_all_mappings()

    def resolve(self, candidates: list[dict]) -> tuple[str | None, str | None]:
        """
        À partir d'une liste de candidats classifiés,
        retourne (keyword, target_folder) si un mappage est connu,
        sinon (None, None).
        """
        for candidate in candidates:
            kw = candidate["keyword"]
            folder = self.get_folder(kw)
            if folder:
                return kw, folder
        return None, None
