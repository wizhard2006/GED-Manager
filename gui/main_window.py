"""
gui/main_window.py
Interface graphique principale de GED-Manager (PySimpleGUI)
"""

import os
import sys
import threading
import PySimpleGUI as sg

from utils.config import Config
from utils.logger import Logger
from db.database import Database
from core.ocr_engine import OCREngine
from core.classifier import Classifier, TYPES_FILE
from core.file_manager import FileManager
from core.mapper import Mapper
from scanner.scanner_interface import ScannerInterface


sg.theme("LightBlue2")
FONT_MAIN = ("Helvetica", 11)
FONT_TITLE = ("Helvetica", 14, "bold")
FONT_MONO = ("Courier New", 10)


class GEDManagerApp:
    def __init__(self):
        self.config = Config()
        self.logger = Logger(self.config.log_file)
        self.db = Database(self.config.db_file)
        self.ocr = OCREngine(
            tesseract_path=self.config.tesseract_path,
            language=self.config.ocr_language,
        )
        self.classifier = Classifier(self.db)
        self.mapper = Mapper(self.db)
        self.file_manager = FileManager(
            db=self.db,
            logger=self.logger,
            ged_root=self.config.ged_root,
            quarantine_folder=self.config.quarantine_folder,
        )
        self.scanner = ScannerInterface(
            logger=self.logger,
            output_folder=self.config.scanner_output_folder,
        )

    # ── FENÊTRE PRINCIPALE ──────────────────────────────────────────────────

    def run(self):
        layout = [
            [sg.Text("GED-Manager", font=FONT_TITLE, justification="center",
                     expand_x=True)],
            [sg.HorizontalSeparator()],
            [
                sg.Button("📄  Fichier existant", key="-FICHIER-",
                          size=(22, 2), font=FONT_MAIN),
                sg.Button("🖨️  Scanner", key="-SCANNER-",
                          size=(22, 2), font=FONT_MAIN),
            ],
            [sg.HorizontalSeparator()],
            [
                sg.Button("🔍  Rechercher", key="-RECHERCHE-",
                          size=(22, 2), font=FONT_MAIN),
                sg.Button("📋  Quarantaine", key="-QUARANTAINE-",
                          size=(22, 2), font=FONT_MAIN),
            ],
            [
                sg.Button("🗂️  Mappages", key="-MAPPAGES-",
                          size=(22, 2), font=FONT_MAIN),
                sg.Button("⚙️  Paramètres", key="-PARAMS-",
                          size=(22, 2), font=FONT_MAIN),
            ],
            [
                sg.Button("🔑  Mots-clés", key="-KEYWORDS-",
                          size=(22, 2), font=FONT_MAIN),
            ],
            [sg.HorizontalSeparator()],
            [sg.Multiline("Prêt.\n", key="-LOG-", size=(58, 8),
                          font=FONT_MONO, disabled=True, autoscroll=True)],
            [sg.Button("Quitter", key="-QUIT-", size=(12, 1))],
        ]

        self.window = sg.Window(
            "GED-Manager — Gestion Électronique de Documents",
            layout,
            finalize=True,
            resizable=True,
        )
        self._log("Application démarrée.")

        while True:
            event, values = self.window.read()
            if event in (sg.WIN_CLOSED, "-QUIT-"):
                break
            elif event == "-FICHIER-":
                self._action_fichier()
            elif event == "-SCANNER-":
                self._action_scanner()
            elif event == "-RECHERCHE-":
                self._action_recherche()
            elif event == "-QUARANTAINE-":
                self._action_quarantaine()
            elif event == "-MAPPAGES-":
                self._action_mappages()
            elif event == "-PARAMS-":
                self._action_params()
            elif event == "-KEYWORDS-":
                self._action_keywords()
            elif event == "-HISTORIQUE-":
                self._action_historique()
            elif event == "-OPENLOG-":
                self._action_open_log()
            elif event == "-SCAN-DONE-":
                # Résultat du scan asynchrone
                result = values["-SCAN-DONE-"]
                if result and isinstance(result, str) and os.path.exists(result):
                    self._log(f"Scan terminé : {os.path.basename(result)}")
                    self._traiter_document(result)
                elif result is None:
                    self._log("Scan annulé ou aucun fichier détecté.")
                    sg.popup_error(
                        "Aucun fichier récupéré du scanner.\n"
                        "Vérifiez que le scanner est allumé et le dossier de sortie configuré."
                    )
                else:
                    self._log(f"Erreur scanner : {result}")
                    sg.popup_error(f"Erreur lors du scan :\n{result}")

        self.window.close()
        self.db.close()

    def _log(self, msg):
        self.window["-LOG-"].update(
            self.window["-LOG-"].get() + msg + "\n"
        )

    # ── SÉLECTION FICHIER ───────────────────────────────────────────────────

    def _action_fichier(self):
        path = sg.popup_get_file(
            "Sélectionnez un document à analyser",
            file_types=(("PDF Files", "*.pdf"), ("Images", "*.jpg *.jpeg *.png *.tiff")),
            title="Ouvrir un document",
        )
        if path:
            self._traiter_document(path)

    # ── SCANNER ─────────────────────────────────────────────────────────────

    def _action_scanner(self):
        choix = sg.popup_yes_no(
            "Voulez-vous utiliser le mode pilotage direct du scanner ?\n\n"
            "OUI  → Pilotage automatique (WIA)\n"
            "NON  → Surveillance dossier de sortie du scanner",
            title="Mode scanner",
        )
        mode = "wia" if choix == "Yes" else "watch"

        if mode == "watch" and not self.config.scanner_output_folder:
            folder = sg.popup_get_folder(
                "Indiquez le dossier où votre scanner dépose les fichiers :",
                title="Dossier de sortie scanner",
            )
            if folder:
                self.config.set("paths", "scanner_output_folder", folder)
                self.scanner.output_folder = folder
            else:
                return

        self._log("Lancement de l'acquisition scanner...")
        self.window.perform_long_operation(
            lambda: self.scanner.acquire(mode), "-SCAN-DONE-"
        )

    # ── TRAITEMENT DOCUMENT ─────────────────────────────────────────────────

    def _traiter_document(self, file_path: str):
        if not os.path.exists(file_path):
            sg.popup_error(f"Fichier introuvable : {file_path}")
            return

        self._log(f"Analyse de : {os.path.basename(file_path)}")
        filename = os.path.basename(file_path)

        # Dédoublonnage
        if self.file_manager.is_duplicate(file_path):
            rep = sg.popup_yes_no(
                f"Ce fichier semble déjà avoir été traité.\n"
                f"Voulez-vous quand même continuer ?",
                title="Doublon détecté",
            )
            if rep != "Yes":
                self._log("Annulé (doublon).")
                return

        # OCR
        self._log("OCR en cours...")
        text = self.ocr.extract_text(file_path)
        # Diagnostic : afficher les 300 premiers caracteres de la 1ere page
        first_page = text.split("\f")[0] if "\f" in text else text[:3000]
        preview = first_page[:300].replace("\n", " ").strip()
        self._log(f"[OCR] Extrait : {preview[:200]}...")
        if not text.strip():
            self._log("Aucun texte extrait — envoi en quarantaine.")
            dest = self.file_manager.send_to_quarantine(
                file_path, detected_keywords="AUCUN_TEXTE"
            )
            sg.popup(f"Aucun texte reconnu.\nDocument placé en quarantaine :\n{dest}")
            return

        # Classification
        candidates = self.classifier.classify(text)
        detected_kws = ", ".join([c["keyword"] for c in candidates[:5]])
        self._log(f"Mots-clés détectés : {detected_kws or 'aucun'}")

        # Vérifier mappage existant
        matched_kw, known_folder = self.mapper.resolve(candidates)
        if known_folder:
            rep = sg.popup_yes_no(
                f"Document reconnu !\n\n"
                f"Mot-clé : {matched_kw}\n"
                f"Dossier : {known_folder}\n\n"
                f"Confirmer le classement ?",
                title="Classification automatique",
            )
            if rep == "Yes":
                final_path = self._propose_rename(file_path, matched_kw, text)
                if final_path is None:
                    self._log("Classement annule.")
                    return
                dest = self.file_manager.move_file(
                    final_path, known_folder, matched_kw, detected_kws
                )
                self._log(f"Deplace vers : {dest}")
                sg.popup(f"Document classe !\n\n-> {dest}", title="Succes")
                return
            else:
                # L'utilisateur veut choisir manuellement
                pass

        # Proposition de classification manuelle
        self._proposer_classification(file_path, candidates, text, detected_kws)

    def _proposer_classification(self, file_path, candidates, text, detected_kws):
        """Fenêtre de classification manuelle avec suggestions"""
        filename = os.path.basename(file_path)
        suggestions = [f"{c['keyword']} → {c['folder']}" for c in candidates[:3]]
        if not suggestions:
            suggestions = ["(aucune suggestion)"]

        # Texte OCR 1ere page pour affichage
        first_page_text = text.split("\f")[0] if "\f" in text else text[:3000]
        ocr_preview = first_page_text[:1500]

        layout = [
            [sg.Text(f"Document : {filename}", font=FONT_MAIN)],
            [sg.Text("Mots-clés détectés :", font=FONT_MAIN)],
            [sg.Text(detected_kws or "aucun", font=FONT_MONO)],
            [sg.HorizontalSeparator()],
            [sg.Text("Texte reconnu par OCR (1ère page) :", font=FONT_MAIN)],
            [sg.Multiline(ocr_preview, size=(60, 6), disabled=True,
                          font=FONT_MONO, autoscroll=False)],
            [sg.HorizontalSeparator()],
            [sg.Text("Suggestions :", font=FONT_MAIN)],
            [sg.Listbox(values=suggestions, size=(55, 4), key="-SUGGESTIONS-",
                        enable_events=True)],
            [sg.HorizontalSeparator()],
            [sg.Text("Ou saisissez manuellement :", font=FONT_MAIN)],
            [sg.Text("Mot-clé / Origine :"),
             sg.Input(key="-KW-", size=(30, 1))],
            [sg.Text("Dossier destination\n(ex: Automobile/Citroen) :"),
             sg.Input(key="-FOLDER-", size=(30, 1)),
             sg.FolderBrowse("Parcourir", initial_folder=self.config.ged_root,
                             target="-FOLDER-")],
            [sg.Checkbox("Mémoriser ce mappage", default=True, key="-LEARN-")],
            [sg.HorizontalSeparator()],
            [
                sg.Button("✅ Classer", key="-OK-"),
                sg.Button("📥 Quarantaine", key="-QUAR-"),
                sg.Button("❌ Annuler", key="-CANCEL-"),
            ],
        ]

        win = sg.Window("Classifier ce document", layout, modal=True, finalize=True)

        while True:
            ev, vals = win.read()
            if ev in (sg.WIN_CLOSED, "-CANCEL-"):
                win.close()
                self._log("Classement annulé.")
                return
            elif ev == "-SUGGESTIONS-" and vals["-SUGGESTIONS-"]:
                selected = vals["-SUGGESTIONS-"][0]
                if "→" in selected:
                    kw, folder = selected.split("→", 1)
                    win["-KW-"].update(kw.strip())
                    win["-FOLDER-"].update(folder.strip())
            elif ev == "-QUAR-":
                win.close()
                dest = self.file_manager.send_to_quarantine(
                    file_path, detected_kws
                )
                self._log(f"📥 Mis en quarantaine : {dest}")
                sg.popup(f"Document placé en quarantaine :\n{dest}")
                return
            elif ev == "-OK-":
                kw = vals["-KW-"].strip()
                folder = vals["-FOLDER-"].strip()
                # Si dossier absolu → rendre relatif à ged_root
                if os.path.isabs(folder) and folder.startswith(self.config.ged_root):
                    folder = os.path.relpath(folder, self.config.ged_root)
                if not folder:
                    sg.popup_error("Veuillez indiquer un dossier de destination.")
                    continue
                win.close()
                if kw and vals["-LEARN-"]:
                    self.mapper.learn(kw, folder)
                    self._log(f"🔖 Mappage enregistré : '{kw}' → '{folder}'")
                # Proposer renommage avant deplacement
                final_path = self._propose_rename(file_path, kw or folder.split("/")[-1], text)
                if final_path is None:
                    self._log("Classement annule.")
                    return
                dest = self.file_manager.move_file(
                    final_path, folder, kw, detected_kws
                )
                self._log(f"Classe dans : {dest}")
                sg.popup(f"Document classe !\n\n-> {dest}", title="Succes")
                return

    # ── RECHERCHE ──────────────────────────────────────────────────────────


    def _propose_rename(self, file_path: str, societe: str, text: str) -> str:
        """
        Fenetre de renommage propose avant deplacement.
        Retourne le nouveau chemin (renomme dans le meme dossier source)
        ou le chemin original si annule.
        """
        import json
        from datetime import date

        # Charger la liste des types
        types_list = ["Facture", "Contrat", "Courrier", "Releve", "Attestation",
                      "Devis", "Rapport", "Avis", "Bulletin", "Quittance",
                      "Certificat", "Convention", "Avenant", "Ticket", "Autre"]
        if os.path.exists(TYPES_FILE):
            try:
                with open(TYPES_FILE, encoding="utf-8") as f:
                    data = json.load(f)
                types_list = data.get("types", types_list)
            except Exception:
                pass

        # Detecter le type automatiquement
        detected_type = self.classifier.detect_type(text)
        default_idx = types_list.index(detected_type) if detected_type in types_list else 0

        # Date du jour par defaut
        today = date.today().strftime("%Y-%m-%d")

        # Nom propose
        ext = os.path.splitext(file_path)[1]
        proposed_name = f"{societe.capitalize()}_{detected_type}_{today}{ext}"

        layout = [
            [sg.Text("Renommer le document", font=FONT_TITLE)],
            [sg.Text("Nom original :"), sg.Text(os.path.basename(file_path), font=FONT_MONO)],
            [sg.HorizontalSeparator()],
            [sg.Text("Societe / Origine :", size=(18, 1)),
             sg.Input(societe.capitalize(), key="-SOC-", size=(25, 1))],
            [sg.Text("Type de document :", size=(18, 1)),
             sg.Combo(types_list, default_value=detected_type,
                      key="-TYPE-", size=(23, 1), readonly=True)],
            [sg.Text("Date (AAAA-MM-JJ) :", size=(18, 1)),
             sg.Input(today, key="-DATE-", size=(15, 1))],
            [sg.HorizontalSeparator()],
            [sg.Text("Nom final :"), sg.Text(proposed_name, key="-PREVIEW-",
                                              font=FONT_MONO, size=(45, 1))],
            [sg.HorizontalSeparator()],
            [
                sg.Button("Renommer et classer", key="-OK-"),
                sg.Button("Classer sans renommer", key="-KEEP-"),
                sg.Button("Annuler", key="-CANCEL-"),
            ],
        ]

        win = sg.Window("Renommage", layout, modal=True, finalize=True)

        def update_preview(vals):
            soc = vals["-SOC-"].strip().replace(" ", "_") or "Document"
            typ = vals["-TYPE-"] or "Autre"
            dat = vals["-DATE-"].strip() or today
            preview = f"{soc}_{typ}_{dat}{ext}"
            win["-PREVIEW-"].update(preview)
            return preview

        current_preview = proposed_name

        while True:
            ev, vals = win.read()
            if ev in (sg.WIN_CLOSED, "-CANCEL-"):
                win.close()
                return None  # Annulation complete
            elif ev in ("-SOC-", "-TYPE-", "-DATE-"):
                current_preview = update_preview(vals)
            elif ev == "-KEEP-":
                win.close()
                return file_path  # Garder nom original
            elif ev == "-OK-":
                current_preview = update_preview(vals)
                new_name = current_preview
                new_path = os.path.join(os.path.dirname(file_path), new_name)
                try:
                    os.rename(file_path, new_path)
                    self._log(f"Renomme : {os.path.basename(file_path)} -> {new_name}")
                    win.close()
                    return new_path
                except Exception as e:
                    sg.popup_error(f"Erreur lors du renommage :\n{e}")

        win.close()
        return file_path

    def _action_recherche(self):
        query = sg.popup_get_text(
            "Rechercher dans l'historique des documents :",
            title="Recherche",
        )
        if not query:
            return
        results = self.db.search_history(query)
        if not results:
            sg.popup("Aucun résultat trouvé.", title="Recherche")
            return

        rows = []
        for r in results:
            rows.append([
                r.get("filename", ""),
                r.get("destination_path", ""),
                r.get("processed_at", ""),
            ])

        layout = [
            [sg.Text(f"Résultats pour : {query}", font=FONT_MAIN)],
            [sg.Table(
                values=rows,
                headings=["Fichier", "Destination", "Date"],
                col_widths=[25, 30, 18],
                auto_size_columns=False,
                display_row_numbers=False,
                num_rows=min(15, len(rows)),
                key="-TABLE-",
                enable_events=True,
                justification="left",
            )],
            [sg.Button("Ouvrir le dossier", key="-OPEN-"),
             sg.Button("Fermer", key="-CLOSE-")],
        ]
        win = sg.Window("Résultats de recherche", layout, modal=True, finalize=True)

        while True:
            ev, vals = win.read()
            if ev in (sg.WIN_CLOSED, "-CLOSE-"):
                break
            elif ev == "-OPEN-" and vals["-TABLE-"]:
                idx = vals["-TABLE-"][0]
                dest = rows[idx][1]
                folder = os.path.dirname(dest)
                if os.path.isdir(folder):
                    os.startfile(folder)
        win.close()

    # ── QUARANTAINE ────────────────────────────────────────────────────────

    def _action_quarantaine(self):
        items = self.db.get_quarantine()
        if not items:
            sg.popup("La quarantaine est vide.", title="Quarantaine")
            return

        rows = [[i["filename"], i["file_path"], i["created_at"]] for i in items]
        layout = [
            [sg.Text("Documents en quarantaine", font=FONT_TITLE)],
            [sg.Table(
                values=rows,
                headings=["Fichier", "Chemin", "Date"],
                col_widths=[25, 35, 18],
                auto_size_columns=False,
                num_rows=min(10, len(rows)),
                key="-QTABLE-",
                enable_events=True,
                justification="left",
            )],
            [sg.Button("Classer ce document", key="-CLASSER-"),
             sg.Button("Fermer", key="-CLOSE-")],
        ]
        win = sg.Window("Quarantaine", layout, modal=True, finalize=True)

        while True:
            ev, vals = win.read()
            if ev in (sg.WIN_CLOSED, "-CLOSE-"):
                break
            elif ev == "-CLASSER-" and vals["-QTABLE-"]:
                idx = vals["-QTABLE-"][0]
                item = items[idx]
                win.close()
                self._traiter_document(item["file_path"])
                self.db.resolve_quarantine(item["id"])
                return
        win.close()

    # ── MAPPAGES ───────────────────────────────────────────────────────────

    def _action_mappages(self):
        mappings = self.db.get_all_mappings()
        rows = [[m["keyword"], m["target_folder"], m["updated_at"]] for m in mappings]

        layout = [
            [sg.Text("Mappages enregistrés", font=FONT_TITLE)],
            [sg.Table(
                values=rows if rows else [["—", "—", "—"]],
                headings=["Mot-clé", "Dossier", "Mis à jour"],
                col_widths=[22, 30, 18],
                auto_size_columns=False,
                num_rows=min(15, max(5, len(rows))),
                key="-MTABLE-",
                justification="left",
            )],
            [sg.Button("Supprimer", key="-DEL-"),
             sg.Button("Fermer", key="-CLOSE-")],
        ]
        win = sg.Window("Mappages", layout, modal=True, finalize=True)

        while True:
            ev, vals = win.read()
            if ev in (sg.WIN_CLOSED, "-CLOSE-"):
                break
            elif ev == "-DEL-" and vals["-MTABLE-"]:
                idx = vals["-MTABLE-"][0]
                kw = rows[idx][0]
                self.db.delete_mapping(kw)
                sg.popup(f"Mappage '{kw}' supprimé.")
                win.close()
                self._action_mappages()
                return
        win.close()

    # ── PARAMÈTRES ─────────────────────────────────────────────────────────


    # -- MOTS-CLES -----------------------------------------------------------


    def _action_historique(self):
        """Affiche les 50 derniers documents traites"""
        history = self.db.get_history(limit=50)
        if not history:
            sg.popup("Aucun document dans l'historique.", title="Historique")
            return

        rows = []
        for r in history:
            rows.append([
                r.get("processed_at", "")[:16],
                r.get("filename", ""),
                r.get("action", ""),
                r.get("destination_path", ""),
            ])

        layout = [
            [sg.Text("Historique des 50 derniers documents", font=FONT_TITLE)],
            [sg.Table(
                values=rows,
                headings=["Date", "Fichier", "Action", "Destination"],
                col_widths=[14, 25, 8, 35],
                auto_size_columns=False,
                num_rows=min(20, len(rows)),
                key="-HTABLE-",
                justification="left",
                enable_events=True,
            )],
            [sg.Button("Ouvrir le dossier", key="-HOPEN-"),
             sg.Button("Fermer", key="-HCLOSE-")],
        ]

        win = sg.Window("Historique", layout, modal=True, finalize=True)
        while True:
            ev, vals = win.read()
            if ev in (sg.WIN_CLOSED, "-HCLOSE-"):
                break
            elif ev == "-HOPEN-" and vals["-HTABLE-"]:
                idx = vals["-HTABLE-"][0]
                dest = rows[idx][3]
                folder = os.path.dirname(dest)
                if os.path.isdir(folder):
                    os.startfile(folder)
                else:
                    sg.popup_error(f"Dossier introuvable :\n{folder}")
        win.close()

    def _action_open_log(self):
        """Ouvre le fichier log dans le Bloc-notes"""
        log_path = os.path.abspath(self.config.log_file)
        if os.path.exists(log_path):
            os.startfile(log_path)
        else:
            sg.popup_error(f"Fichier log introuvable :\n{log_path}")

    def _action_keywords(self):
        """Fenetre de gestion des mots-cles (keywords.json) - v1.4 : edition + tri"""
        import json
        from core.classifier import KEYWORDS_FILE, TYPES_FILE

        # --- Fonctions utilitaires ---

        def load_flat(sort_col=0):
            """Charge et aplatit keywords.json, trie par colonne (0=mot-cle, 1=dossier, 2=categorie)."""
            if not os.path.exists(KEYWORDS_FILE):
                return []
            with open(KEYWORDS_FILE, encoding="utf-8") as f:
                data = json.load(f)
            flat = []
            for cat, entries in data.items():
                if cat == "_notice":
                    continue
                if isinstance(entries, dict):
                    for kw, folder in entries.items():
                        flat.append([kw, folder, cat])
            flat.sort(key=lambda x: x[sort_col].lower())
            return flat

        def save_entry(keyword, folder, category="Personnalise", old_keyword=None, old_category=None):
            """Ajoute ou met a jour une entree dans keywords.json.
            Si old_keyword/old_category sont fournis, supprime l'ancienne entree avant d'ajouter."""
            if not os.path.exists(KEYWORDS_FILE):
                data = {}
            else:
                with open(KEYWORDS_FILE, encoding="utf-8") as f:
                    data = json.load(f)
            # Supprimer l'ancienne entree si on est en mode edition
            if old_keyword and old_category:
                old_kw_lower = old_keyword.lower().strip()
                if old_category in data and old_kw_lower in data[old_category]:
                    del data[old_category][old_kw_lower]
            # Ajouter/mettre a jour
            if category not in data:
                data[category] = {}
            data[category][keyword.lower().strip()] = folder.strip()
            with open(KEYWORDS_FILE, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)

        def delete_entry(keyword, category):
            if not os.path.exists(KEYWORDS_FILE):
                return
            with open(KEYWORDS_FILE, encoding="utf-8") as f:
                data = json.load(f)
            if category in data and keyword.lower() in data[category]:
                del data[category][keyword.lower()]
            with open(KEYWORDS_FILE, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)

        # --- Etat initial ---
        sort_col = 0          # 0=Mot-cle, 1=Dossier, 2=Categorie
        edit_mode = False     # True quand on edite une ligne existante
        edit_orig_kw = None   # Mot-cle original avant edition
        edit_orig_cat = None  # Categorie originale avant edition

        rows = load_flat(sort_col)

        # --- Layout ---
        layout = [
            [sg.Text("Mots-cles de classification", font=FONT_TITLE)],
            [sg.Text(
                "Recherches dans la 1ere page du document pour detecter l'origine.",
                font=FONT_MAIN
            )],
            # Tri
            [
                sg.Text("Trier par :", font=FONT_MAIN),
                sg.Radio("Mot-cle", "SORT", default=True, key="-SORT0-", enable_events=True),
                sg.Radio("Dossier cible", "SORT", key="-SORT1-", enable_events=True),
                sg.Radio("Categorie", "SORT", key="-SORT2-", enable_events=True),
            ],
            [sg.Table(
                values=rows if rows else [["--", "--", "--"]],
                headings=["Mot-cle", "Dossier cible", "Categorie"],
                col_widths=[25, 28, 15],
                auto_size_columns=False,
                num_rows=min(15, max(6, len(rows))),
                key="-KTABLE-",
                justification="left",
                enable_events=True,
            )],
            [sg.HorizontalSeparator()],
            # Formulaire ajout / edition
            [sg.Text("Ajouter un mot-cle :", font=FONT_MAIN, key="-FORM-TITLE-")],
            [
                sg.Text("Mot-cle :", size=(13, 1)),
                sg.Input(key="-NKW-", size=(25, 1)),
            ],
            [
                sg.Text("Dossier cible :", size=(13, 1)),
                sg.Input(key="-NFOLDER-", size=(25, 1)),
                sg.FolderBrowse("Parcourir",
                                initial_folder=self.config.ged_root,
                                target="-NFOLDER-"),
            ],
            [
                sg.Text("Categorie :", size=(13, 1)),
                sg.Input("Personnalise", key="-NCAT-", size=(20, 1)),
            ],
            [sg.HorizontalSeparator()],
            [
                sg.Button("Ajouter", key="-KADD-"),
                sg.Button("Enregistrer modification", key="-KEDIT-", visible=False),
                sg.Button("Annuler edition", key="-KCANCEL-", visible=False),
                sg.Button("Supprimer selection", key="-KDEL-"),
            ],
            [sg.HorizontalSeparator()],
            # Section Types de documents
            [sg.Text("Types de documents (liste de renommage) :", font=FONT_MAIN)],
            [sg.Listbox(values=[], key="-TYPELIST-", size=(40, 4), enable_events=True)],
            [sg.Text("Ajouter :", size=(8, 1)),
             sg.Input(key="-NEWTYPE-", size=(20, 1)),
             sg.Button("Ajouter type", key="-ADDTYPE-"),
             sg.Button("Supprimer type", key="-DELTYPE-")],
            [sg.HorizontalSeparator()],
            [sg.Button("Fermer", key="-KCLOSE-")],
        ]

        win = sg.Window("Mots-cles", layout, modal=True, finalize=True)

        # --- Fonctions types ---
        def load_types():
            if not os.path.exists(TYPES_FILE):
                return []
            with open(TYPES_FILE, encoding="utf-8") as f:
                data = json.load(f)
            return data.get("types", [])

        def save_types(types_list):
            if not os.path.exists(TYPES_FILE):
                data = {}
            else:
                with open(TYPES_FILE, encoding="utf-8") as f:
                    data = json.load(f)
            data["types"] = types_list
            with open(TYPES_FILE, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)

        def reset_form():
            """Remet le formulaire en mode Ajouter."""
            nonlocal edit_mode, edit_orig_kw, edit_orig_cat
            edit_mode = False
            edit_orig_kw = None
            edit_orig_cat = None
            win["-FORM-TITLE-"].update("Ajouter un mot-cle :")
            win["-NKW-"].update("")
            win["-NFOLDER-"].update("")
            win["-NCAT-"].update("Personnalise")
            win["-KADD-"].update(visible=True)
            win["-KEDIT-"].update(visible=False)
            win["-KCANCEL-"].update(visible=False)

        current_types = load_types()
        win["-TYPELIST-"].update(values=current_types)

        # --- Boucle evenements ---
        while True:
            ev, vals = win.read()

            if ev in (sg.WIN_CLOSED, "-KCLOSE-"):
                self.classifier.reload_keywords()
                self._log("Mots-cles recharges.")
                break

            # --- Tri ---
            elif ev in ("-SORT0-", "-SORT1-", "-SORT2-"):
                sort_col = int(ev[-2])  # extrait 0, 1 ou 2
                rows = load_flat(sort_col)
                win["-KTABLE-"].update(values=rows if rows else [["--", "--", "--"]])

            # --- Clic sur une ligne : charger en edition ---
            elif ev == "-KTABLE-" and vals["-KTABLE-"]:
                idx = vals["-KTABLE-"][0]
                if 0 <= idx < len(rows):
                    kw_sel, folder_sel, cat_sel = rows[idx][0], rows[idx][1], rows[idx][2]
                    edit_mode = True
                    edit_orig_kw = kw_sel
                    edit_orig_cat = cat_sel
                    win["-FORM-TITLE-"].update(f"Modifier le mot-cle : '{kw_sel}'")
                    win["-NKW-"].update(kw_sel)
                    win["-NFOLDER-"].update(folder_sel)
                    win["-NCAT-"].update(cat_sel)
                    win["-KADD-"].update(visible=False)
                    win["-KEDIT-"].update(visible=True)
                    win["-KCANCEL-"].update(visible=True)

            # --- Ajouter ---
            elif ev == "-KADD-":
                kw = vals["-NKW-"].strip()
                folder = vals["-NFOLDER-"].strip()
                cat = vals["-NCAT-"].strip() or "Personnalise"
                if os.path.isabs(folder) and folder.startswith(self.config.ged_root):
                    folder = os.path.relpath(folder, self.config.ged_root)
                if not kw or not folder:
                    sg.popup_error("Mot-cle et dossier cible sont obligatoires.")
                    continue
                save_entry(kw, folder, cat)
                self._log(f"Mot-cle ajoute : '{kw}' -> '{folder}'")
                rows = load_flat(sort_col)
                win["-KTABLE-"].update(values=rows if rows else [["--", "--", "--"]])
                reset_form()

            # --- Enregistrer modification ---
            elif ev == "-KEDIT-":
                kw = vals["-NKW-"].strip()
                folder = vals["-NFOLDER-"].strip()
                cat = vals["-NCAT-"].strip() or "Personnalise"
                if os.path.isabs(folder) and folder.startswith(self.config.ged_root):
                    folder = os.path.relpath(folder, self.config.ged_root)
                if not kw or not folder:
                    sg.popup_error("Mot-cle et dossier cible sont obligatoires.")
                    continue
                save_entry(kw, folder, cat,
                           old_keyword=edit_orig_kw,
                           old_category=edit_orig_cat)
                self._log(f"Mot-cle modifie : '{edit_orig_kw}' -> '{kw}' / '{folder}'")
                rows = load_flat(sort_col)
                win["-KTABLE-"].update(values=rows if rows else [["--", "--", "--"]])
                reset_form()

            # --- Annuler edition ---
            elif ev == "-KCANCEL-":
                reset_form()

            # --- Supprimer ---
            elif ev == "-KDEL-" and vals["-KTABLE-"]:
                idx = vals["-KTABLE-"][0]
                if 0 <= idx < len(rows):
                    kw_del = rows[idx][0]
                    cat_del = rows[idx][2]
                    rep = sg.popup_yes_no(
                        f"Supprimer le mot-cle '{kw_del}' ?",
                        title="Confirmation"
                    )
                    if rep == "Yes":
                        delete_entry(kw_del, cat_del)
                        rows = load_flat(sort_col)
                        win["-KTABLE-"].update(values=rows if rows else [["--", "--", "--"]])
                        self._log(f"Mot-cle supprime : '{kw_del}'")
                        reset_form()

            # --- Types de documents ---
            elif ev == "-ADDTYPE-":
                new_type = vals["-NEWTYPE-"].strip()
                if new_type and new_type not in current_types:
                    current_types.append(new_type)
                    save_types(current_types)
                    win["-TYPELIST-"].update(values=current_types)
                    win["-NEWTYPE-"].update("")
                    self._log(f"Type ajoute : '{new_type}'")
            elif ev == "-DELTYPE-" and vals["-TYPELIST-"]:
                type_del = vals["-TYPELIST-"][0]
                current_types.remove(type_del)
                save_types(current_types)
                win["-TYPELIST-"].update(values=current_types)
                self._log(f"Type supprime : '{type_del}'")

        win.close()

    def _action_params(self):
        layout = [
            [sg.Text("Paramètres", font=FONT_TITLE)],
            [sg.Text("Dossier racine GED (ex: D:\\)"),
             sg.Input(self.config.ged_root, key="-ROOT-", size=(35, 1)),
             sg.FolderBrowse()],
            [sg.Text("Dossier quarantaine"),
             sg.Input(self.config.quarantine_folder, key="-QFOLDER-", size=(35, 1)),
             sg.FolderBrowse()],
            [sg.Text("Chemin Tesseract OCR"),
             sg.Input(self.config.tesseract_path, key="-TESS-", size=(35, 1)),
             sg.FileBrowse()],
            [sg.Text("Dossier sortie scanner"),
             sg.Input(self.config.scanner_output_folder, key="-SCANFOLDER-", size=(35, 1)),
             sg.FolderBrowse()],
            [sg.Text("Langue OCR (ex: fra+eng)"),
             sg.Input(self.config.ocr_language, key="-LANG-", size=(15, 1))],
            [sg.HorizontalSeparator()],
            [sg.Button("Enregistrer", key="-SAVE-"),
             sg.Button("Annuler", key="-CANCEL-")],
        ]
        win = sg.Window("Paramètres", layout, modal=True)

        while True:
            ev, vals = win.read()
            if ev in (sg.WIN_CLOSED, "-CANCEL-"):
                break
            elif ev == "-SAVE-":
                self.config.set("paths", "ged_root", vals["-ROOT-"])
                self.config.set("paths", "quarantine_folder", vals["-QFOLDER-"])
                self.config.set("paths", "tesseract_path", vals["-TESS-"])
                self.config.set("paths", "scanner_output_folder", vals["-SCANFOLDER-"])
                self.config.set("app", "language", vals["-LANG-"])
                # Mettre à jour les instances
                self.file_manager.ged_root = vals["-ROOT-"]
                self.file_manager.quarantine_folder = vals["-QFOLDER-"]
                self.scanner.output_folder = vals["-SCANFOLDER-"]
                self.ocr.language = vals["-LANG-"]
                sg.popup("Paramètres enregistrés.", title="Paramètres")
                break
        win.close()
