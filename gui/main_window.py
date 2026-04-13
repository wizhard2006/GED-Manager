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
            enhanced_preprocessing=self.config.enhanced_preprocessing,
            tesseract_psm=self.config.tesseract_psm,
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
                sg.Button("📦  Traitement en masse", key="-MASSE-",
                          size=(22, 2), font=FONT_MAIN),
            ],
            [
                sg.Button("📖  Manuel", key="-MANUEL-",
                          size=(22, 2), font=FONT_MAIN),
                sg.Button("📋  Changelog", key="-CHANGELOG-",
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
            elif event == "-MASSE-":
                self._action_masse()
            elif event == "-HISTORIQUE-":
                self._action_historique()
            elif event == "-MANUEL-":
                self._action_manuel()
            elif event == "-CHANGELOG-":
                self._action_changelog()
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

    def _action_masse(self):
        """Traitement en masse : file d'attente de documents, validation un par un."""

        # --- Sélection multi-fichiers ---
        files_str = sg.popup_get_file(
            "Sélectionnez les documents à traiter",
            title="Traitement en masse",
            multiple_files=True,
            file_types=(("PDF Files", "*.pdf"), ("Images", "*.jpg *.jpeg *.png *.tiff")),
        )
        if not files_str:
            return

        # PySimpleGUI retourne les fichiers séparés par ";"
        file_list = [f.strip() for f in files_str.split(";") if f.strip()]
        if not file_list:
            return

        total = len(file_list)
        self._log(f"Traitement en masse : {total} document(s) sélectionné(s).")

        # --- Statuts possibles ---
        STATUT = {"attente": "En attente", "ok": "Classé ✅", "quar": "Quarantaine 📥", "annule": "Annulé ❌"}

        # --- Construction de la liste de statuts ---
        statuts = [STATUT["attente"]] * total
        noms = [os.path.basename(f) for f in file_list]
        rows = [[noms[i], statuts[i]] for i in range(total)]

        layout = [
            [sg.Text(f"File d'attente — {total} document(s)", font=FONT_TITLE)],
            [sg.Table(
                values=rows,
                headings=["Fichier", "Statut"],
                col_widths=[45, 15],
                auto_size_columns=False,
                num_rows=min(15, max(5, total)),
                key="-MTABLE-",
                justification="left",
                font=FONT_MAIN,
            )],
            [sg.HorizontalSeparator()],
            [sg.ProgressBar(total, orientation="h", size=(40, 20), key="-MPROG-")],
            [sg.Text("Prêt.", key="-MSTATUS-", font=FONT_MAIN, size=(55, 1))],
            [sg.HorizontalSeparator()],
            [
                sg.Button("▶  Démarrer", key="-MSTART-"),
                sg.Button("Fermer", key="-MCLOSE-"),
            ],
        ]

        win = sg.Window("Traitement en masse", layout, modal=True, finalize=True)

        # Compteurs récapitulatif
        nb_classe = 0
        nb_quar = 0
        nb_annule = 0
        started = False

        while True:
            ev, vals = win.read(timeout=100)
            if ev in (sg.WIN_CLOSED, "-MCLOSE-"):
                break

            elif ev == "-MSTART-" and not started:
                started = True
                win["-MSTART-"].update(disabled=True)

                for i, file_path in enumerate(file_list):
                    nom = noms[i]
                    win["-MSTATUS-"].update(f"Traitement {i+1}/{total} : {nom}")
                    rows[i][1] = "En cours..."
                    win["-MTABLE-"].update(values=rows)
                    win.refresh()

                    if not os.path.exists(file_path):
                        rows[i][1] = "Introuvable ⚠️"
                        win["-MTABLE-"].update(values=rows)
                        nb_annule += 1
                        continue

                    # OCR
                    text = self.ocr.extract_text(file_path)
                    if not text.strip():
                        dest = self.file_manager.send_to_quarantine(
                            file_path, detected_keywords="AUCUN_TEXTE"
                        )
                        rows[i][1] = STATUT["quar"]
                        win["-MTABLE-"].update(values=rows)
                        self._log(f"[Masse] Quarantaine (aucun texte) : {nom}")
                        nb_quar += 1
                        win["-MPROG-"].update(i + 1)
                        win.refresh()
                        continue

                    # Classification
                    candidates = self.classifier.classify(text)
                    detected_kws = ", ".join([c["keyword"] for c in candidates[:5]])

                    # Mode hybride : si seuil actifé et score suffisant → auto
                    threshold = self.config.auto_classify_threshold
                    top = candidates[0] if candidates else None
                    auto_ok = threshold > 0 and top and top["score"] >= threshold

                    if auto_ok:
                        # Classement automatique silencieux
                        final_path = self._propose_rename(file_path, top["keyword"], text)
                        if final_path:
                            dest = self.file_manager.move_file(
                                final_path, top["folder"], top["keyword"], detected_kws
                            )
                            rows[i][1] = STATUT["ok"]
                            self._log(f"[Masse/Auto] Classé : {nom} → {dest}")
                            nb_classe += 1
                        else:
                            rows[i][1] = STATUT["annule"]
                            nb_annule += 1
                    else:
                        # Mode 2 : validation manuelle pour ce document
                        # On ferme temporairement la fenêtre de masse (hide)
                        win.hide()
                        self._proposer_classification_masse(
                            file_path, candidates, text, detected_kws,
                            callback_ok=lambda d, s: None,   # résultat via _last_masse
                        )
                        # Récupérer le résultat de la classification
                        result = getattr(self, "_masse_last_result", None)
                        win.un_hide()

                        if result == "classe":
                            rows[i][1] = STATUT["ok"]
                            nb_classe += 1
                        elif result == "quarantaine":
                            rows[i][1] = STATUT["quar"]
                            nb_quar += 1
                        else:
                            rows[i][1] = STATUT["annule"]
                            nb_annule += 1

                        self._masse_last_result = None

                    win["-MTABLE-"].update(values=rows)
                    win["-MPROG-"].update(i + 1)
                    win.refresh()

                # Récapitulatif final
                win["-MSTATUS-"].update(
                    f"Terminé — Classés : {nb_classe}  |  "
                    f"Quarantaine : {nb_quar}  |  Annulés : {nb_annule}"
                )
                win["-MSTART-"].update(visible=False)
                self._log(
                    f"[Masse] Terminé : {nb_classe} classé(s), "
                    f"{nb_quar} quarantaine, {nb_annule} annulé(s)."
                )

        win.close()

    def _proposer_classification_masse(self, file_path, candidates, text,
                                       detected_kws, callback_ok=None):
        """
        Variante de _proposer_classification pour le mode masse.
        Stocke le résultat dans self._masse_last_result :
          'classe' | 'quarantaine' | 'annule'
        """
        self._masse_last_result = "annule"
        filename = os.path.basename(file_path)
        suggestions = [f"{c['keyword']} → {c['folder']}" for c in candidates[:3]]
        if not suggestions:
            suggestions = ["(aucune suggestion)"]

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
            [sg.Text("Dossier destination :"),
             sg.Input(key="-FOLDER-", size=(30, 1)),
             sg.FolderBrowse("Parcourir", initial_folder=self.config.ged_root,
                             target="-FOLDER-")],
            [sg.Checkbox("Mémoriser ce mappage", default=True, key="-LEARN-")],
            [sg.HorizontalSeparator()],
            [
                sg.Button("✅ Classer", key="-OK-"),
                sg.Button("📥 Quarantaine", key="-QUAR-"),
                sg.Button("⏭ Passer (annuler)", key="-SKIP-"),
            ],
        ]

        win = sg.Window(
            f"Classer — {filename}", layout, modal=True, finalize=True
        )

        while True:
            ev, vals = win.read()
            if ev in (sg.WIN_CLOSED, "-SKIP-"):
                self._masse_last_result = "annule"
                break
            elif ev == "-SUGGESTIONS-" and vals["-SUGGESTIONS-"]:
                selected = vals["-SUGGESTIONS-"][0]
                if "→" in selected:
                    kw, folder = selected.split("→", 1)
                    win["-KW-"].update(kw.strip())
                    win["-FOLDER-"].update(folder.strip())
            elif ev == "-QUAR-":
                dest = self.file_manager.send_to_quarantine(file_path, detected_kws)
                self._log(f"[Masse] Quarantaine : {filename} → {dest}")
                self._masse_last_result = "quarantaine"
                break
            elif ev == "-OK-":
                kw = vals["-KW-"].strip()
                folder = vals["-FOLDER-"].strip()
                if os.path.isabs(folder) and folder.startswith(self.config.ged_root):
                    folder = os.path.relpath(folder, self.config.ged_root)
                if not folder:
                    sg.popup_error("Veuillez indiquer un dossier de destination.")
                    continue
                if kw and vals["-LEARN-"]:
                    self.mapper.learn(kw, folder)
                final_path = self._propose_rename(file_path, kw, text)
                if final_path is None:
                    self._masse_last_result = "annule"
                    break
                dest = self.file_manager.move_file(
                    final_path, folder, kw, detected_kws
                )
                self._log(f"[Masse] Classé : {filename} → {dest}")
                self._masse_last_result = "classe"
                break

        win.close()

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

    def _action_changelog(self):
        """Popup changelog versions GED-Manager"""
        CHANGELOG = """
GED-Manager — Historique des versions
════════════════════════════════════════════

v1.8 — Manuel, Changelog, Lanceur (actuelle)
  + Bouton Manuel : notice d'utilisation avec glossaire
  + Bouton Changelog : historique des versions
  + Lanceur GED-Manager.vbs (sans fenêtre noire)
  + Script creer_raccourci.bat pour icône sur le bureau

v1.7 — PSM Tesseract configurable depuis l'interface
  + Menu déroulant PSM dans Paramètres (5 modes présélectionnés)
  + Bouton ? avec aide complète (14 modes PSM commentés)
  + Checkbox prétraitement OpenCV dans Paramètres
  + Changements PSM pris en compte sans redémarrage

v1.6 — Amélioration OCR (PSM6 + OpenCV)
  + PSM 6 actif par défaut (meilleur pour factures/courriers)
  + Prétraitement OpenCV optionnel : débruitage, binarisation, deskew
  + Fichier original jamais modifié (traitement en mémoire uniquement)
  + Paramètres dans ged_config.ini (retour arrière facile)

v1.5 — Traitement en masse + règle sous-chemin
  + Bouton Traitement en masse : sélection multi-fichiers
  + File d'attente avec tableau de statut et barre de progression
  + Validation document par document (Classer / Quarantaine / Passer)
  + Récapitulatif final (classés / quarantaine / annulés)
  + Règle sous-chemin : dossier plus précis toujours privilégié
  + Seuil de confiance hybride configurable (désactivé par défaut)

v1.4.2 — Mots-clés véhicules
  + Catégorie Automobile renommée Vehicules
  + Mots-clés orientés vers D:\\SynologyDrive\\_Vehicules\\
  + Sous-dossiers par véhicule : BMW S1000R, Buell, Citroen_C3,
    Copen, DACIA Duster, Secma F16, Suzuki DR750, YUBA Fastrack

v1.4.1 — Normalisation accents et casse
  + Insensible aux accents ET à la casse dans tout l'OCR
  + 'Société' = 'societe' = 'SOCIETE' pour la classification
  + Correction encodage corrompu (SociÃ©tÃ© etc.)

v1.4 — Édition mots-clés + tri par colonne
  + Clic sur une ligne → chargement dans le formulaire d'édition
  + Boutons Enregistrer modification / Annuler édition
  + Radio buttons de tri : Mot-clé / Dossier cible / Catégorie
  + Correction corruption main_window.py (méthodes dupliquées)

v1.3 — Diagnostic et Historique
  + Texte OCR extrait visible dans la fenêtre de classification
  + Log de diagnostic OCR dans la zone principale
  + Bouton Historique : 50 derniers documents traités
  + Bouton Ouvrir log : ouvre ged_manager.log dans le Bloc-notes

v1.2 — Renommage automatique
  + Dialogue de renommage : Société_Type_Date.pdf
  + Détection automatique du type de document
  + Liste des types éditable depuis l'interface
  + Fichier document_types.json avec mots-clés de détection
  + Boutons : Renommer et classer / Classer sans renommer

v1.1 — OCR amélioré et mots-clés
  + Analyse sur la 1ère page uniquement (moins de faux positifs)
  + keywords.json : mots-clés éditables sans toucher au code
  + Résolution OCR 3x (~216 DPI)
  + Mots-clés plus spécifiques (suppression des 3 lettres)

v1.0 — MVP initial
  + Interface graphique PySimpleGUI
  + OCR Tesseract (fra+eng)
  + Classification par mots-clés → dossier
  + Mode apprentissage : mappage manuel mémorisé
  + Documents incertains → quarantaine
  + Scanner CANON DR-C125 (WIA + fallback dossier surveillé)
  + Vérification stabilité fichier scanner (anti-fichier partiel)
  + Recherche, Quarantaine, Mappages, Paramètres
  + Base SQLite : mappages + historique + quarantaine
  + Logs horodatés
"""
        sg.popup_scrolled(
            CHANGELOG,
            title="Changelog — GED-Manager",
            size=(62, 30),
            font=FONT_MONO,
        )

    def _action_manuel(self):
        """Manuel d'utilisation avec glossaire interactif"""

        SECTIONS = {
            "Modus operandi — Scanner et classer": """
MODUS OPERANDI — Scanner un document et le classer
══════════════════════════════════════════════════

1. Placez le document dans le chargeur du scanner
2. Cliquez sur [Scanner]
3. L'application attend le fichier scanné (mode dossier surveillé)
4. Une fois détecté, l'OCR extrait le texte automatiquement
5. Une fenêtre de classification s'ouvre :
   • Les mots-clés détectés sont affichés
   • Le texte OCR de la 1ère page est visible
   • Les suggestions de dossier sont proposées
6. Choisissez une suggestion ou saisissez manuellement
7. Cliquez sur [Renommer et classer] ou [Classer sans renommer]
8. Le document est déplacé dans le bon dossier
   Le mappage est mémorisé pour les prochains documents similaires

Conseil : scanner en N&B à 400 DPI pour une meilleure reconnaissance.
""",
            "Modus operandi — Fichier existant": """
MODUS OPERANDI — Classer un fichier existant
══════════════════════════════════════════════════

1. Cliquez sur [Fichier existant]
2. Sélectionnez le PDF ou l'image à classer
3. L'OCR analyse automatiquement le document
4. La fenêtre de classification s'ouvre
5. Suivez les étapes 5 à 8 du modus operandi Scanner
""",
            "Modus operandi — Traitement en masse": """
MODUS OPERANDI — Traiter plusieurs documents d'un coup
══════════════════════════════════════════════════

1. Cliquez sur [Traitement en masse]
2. Sélectionnez plusieurs fichiers (Ctrl+clic dans l'explorateur)
3. Une file d'attente s'affiche avec le statut de chaque document
4. Cliquez sur [Démarrer]
5. Pour chaque document, une fenêtre de classification s'ouvre
   Choisissez : Classer / Quarantaine / Passer
6. Le récapitulatif final indique :
   Classés ✅ | Quarantaine 📥 | Annulés ❌

Conseil : les documents non reconnus peuvent être envoyés en
quarantaine pour être traités plus tard.
""",
            "Fichier existant": """
FONCTION : Fichier existant
════════════════════════════

Permet de sélectionner un fichier PDF ou image déjà présent
sur votre ordinateur pour le classer dans la GED.

Formats acceptés : PDF, JPG, PNG, TIFF
Accès : bouton [Fichier existant] en haut de la fenêtre principale
""",
            "Scanner": """
FONCTION : Scanner
════════════════════

Lance l'acquisition depuis le scanner CANON DR-C125.
Deux modes automatiques :
  • WIA (direct) : tentative de scan direct
  • Dossier surveillé (fallback) : attend qu'un fichier
    apparaisse dans le dossier de sortie du scanner
    (configuré dans Paramètres)

Vérification de stabilité : le fichier est attendu stable
(taille identique 3x) avant traitement pour éviter les
fichiers partiels en cours de transfert.
""",
            "Rechercher": """
FONCTION : Rechercher
═══════════════════════

Recherche dans l'historique des documents classés.
Critères : nom de fichier, mot-clé détecté, dossier destination.
Accès : bouton [Rechercher]
""",
            "Quarantaine": """
FONCTION : Quarantaine
═════════════════════════

Affiche les documents en attente de classement manuel.
Un document est envoyé en quarantaine quand :
  • Aucun mot-clé n'est reconnu
  • Le texte OCR est vide
  • Vous avez cliqué [Quarantaine] lors de la classification

Depuis cette fenêtre, vous pouvez relancer la classification
de chaque document manuellement.
Dossier physique configuré dans Paramètres (D:\\_A_Classer par défaut).
""",
            "Mappages": """
FONCTION : Mappages
═════════════════════

Affiche et gère les associations apprises :
  mot-clé détecté → dossier de destination

Différence avec Mots-clés :
  • Mappages : appris automatiquement lors du classement manuel
    Stockés en base de données (ged_manager.db)
    Priorité maximale (score x10)
  • Mots-clés : liste de référence éditable (keywords.json)
    Priorité secondaire (score x5)
""",
            "Mots-clés": """
FONCTION : Mots-clés
═══════════════════════

Gère la liste de référence des mots-clés de classification.

Fonctionnalités :
  • Ajouter un mot-clé et son dossier cible
  • Cliquer une ligne pour éditer (modifier mot-clé, dossier, catégorie)
  • Supprimer un mot-clé sélectionné
  • Trier par colonne : Mot-clé / Dossier cible / Catégorie
  • Gérer la liste des types de documents (renommage)

Les comparaisons sont insensibles aux accents ET à la casse.
Modifications prises en compte à la fermeture de la fenêtre.
""",
            "Traitement en masse": """
FONCTION : Traitement en masse
══════════════════════════════════

Traite plusieurs documents en file d'attente.
Sélection : Ctrl+clic dans l'explorateur Windows.

Statuts possibles :
  En attente   — pas encore traité
  En cours...  — OCR + classification en cours
  Classé ✅     — déplacé dans le bon dossier
  Quarantaine 📥 — envoyé en attente de classement manuel
  Annulé ❌    — ignoré (Passer ou fichier introuvable)

Mode hybride (avancé) : si le score de confiance dépasse
le seuil configuré, le classement est automatique.
Désactivé par défaut (seuil = 0).
""",
            "Historique": """
FONCTION : Historique
══════════════════════

Affiche les 50 derniers documents traités avec :
  • Date et heure
  • Nom du fichier original
  • Dossier de destination
  • Mot-clé ayant déclenché le classement
""",
            "Paramètres": """
FONCTION : Paramètres
══════════════════════

Configure l'application :

  Chemins
  • Dossier racine GED (ex: D:\\)
  • Dossier quarantaine (ex: D:\\_A_Classer)
  • Chemin Tesseract OCR
  • Dossier de sortie scanner
  • Langue OCR (fra+eng recommandé)

  OCR avancé
  • PSM Tesseract : mode de segmentation de page
    (6 recommandé pour factures/courriers)
    Bouton [?] pour aide complète sur les 14 modes
  • Prétraitement OpenCV : débruitage + binarisation + deskew
    Tester sur quelques documents avant d'activer définitivement
    Fichier original jamais modifié (traitement en mémoire)

Sauvegardes dans ged_config.ini (modifiable aussi manuellement).
""",
            "Ouvrir log": """
FONCTION : Ouvrir log
══════════════════════

Ouvre le fichier ged_manager.log dans le Bloc-notes Windows.
Contient toutes les actions horodatées : classements, erreurs,
clés détectées, déplacements de fichiers.
Utile pour diagnostiquer un problème de classification.
""",
            "Renommage automatique": """
FONCTION : Renommage automatique
══════════════════════════════════

Propose un nom standardisé avant classement :
  Format : Société_Type_Date.pdf
  Exemple : EDF_Facture_2024-03.pdf

Le type est détecté automatiquement puis proposé modifiable.
La date est saisie manuellement (format libre).

Deux boutons au choix :
  [Renommer et classer] — applique le nouveau nom
  [Classer sans renommer] — conserve le nom original
""",
            "Classification et scores": """
FONCTION : Classification et scores
════════════════════════════════════

Le système calcule un score pour chaque mot-clé détecté :
  • Mappage appris (base données) : score x10 par occurrence
  • Mot-clé de référence (keywords.json) : score x5 par occurrence

Règle sous-chemin : si deux dossiers détectés dont l'un est
sous-dossier de l'autre, le plus précis est toujours privilégié.
Ex: _Vehicules + Citroen_C3 dans le même doc → Citroen_C3 gagne.

Analyse sur la 1ère page uniquement pour éviter les faux positifs.
Comparaison insensible aux accents et à la casse.
""",
        }

        glossaire = list(SECTIONS.keys())
        first_key = glossaire[0]

        layout = [
            [sg.Text("Manuel d'utilisation — GED-Manager", font=FONT_TITLE)],
            [sg.HorizontalSeparator()],
            [
                sg.Column(
                    [[sg.Listbox(
                        values=glossaire,
                        size=(28, 20),
                        key="-GLOSSAIRE-",
                        enable_events=True,
                        font=FONT_MAIN,
                    )]],
                    vertical_alignment="top",
                ),
                sg.VSeparator(),
                sg.Column(
                    [[sg.Multiline(
                        SECTIONS[first_key],
                        size=(48, 20),
                        disabled=True,
                        key="-CONTENU-",
                        font=FONT_MONO,
                        autoscroll=False,
                    )]],
                    vertical_alignment="top",
                ),
            ],
            [sg.HorizontalSeparator()],
            [sg.Button("Fermer", key="-MCLOSE-")],
        ]

        win = sg.Window("Manuel — GED-Manager", layout, modal=True, finalize=True)
        win["-GLOSSAIRE-"].update(set_to_index=[0])

        while True:
            ev, vals = win.read()
            if ev in (sg.WIN_CLOSED, "-MCLOSE-"):
                break
            elif ev == "-GLOSSAIRE-" and vals["-GLOSSAIRE-"]:
                selected = vals["-GLOSSAIRE-"][0]
                win["-CONTENU-"].update(SECTIONS.get(selected, ""))

        win.close()

    def _action_params(self):

        # --- Options PSM avec descriptions courtes ---
        PSM_OPTIONS = [
            "3  — Auto (détection complète)",
            "4  — Colonne unique",
            "6  — Bloc de texte uniforme (recommandé factures/courriers)",
            "11 — Texte épars (formulaires, étiquettes)",
            "12 — Texte épars + détection auto",
        ]
        PSM_HELP = (
            "Modes de segmentation de page Tesseract (PSM)\n"
            "═══════════════════════════════════════════════\n\n"
            "  0  Détecte l'orientation uniquement, ne lit pas\n"
            "  1  Auto + détection orientation\n"
            "  2  Auto sans découpage de page\n"
            "  3  Détection automatique complète  ← ancien défaut\n"
            "  4  Colonne de texte unique, tailles variables\n"
            "  5  Bloc vertical (texte vertical, japonais…)\n"
            "  6  Bloc de texte uniforme  ← recommandé GED\n"
            "  7  Ligne unique de texte\n"
            "  8  Un seul mot\n"
            "  9  Un seul mot dans un cercle\n"
            " 10  Un seul caractère\n"
            " 11  Texte épars, pas d'ordre  ← formulaires\n"
            " 12  Texte épars + détection auto\n"
            " 13  Ligne brute (ignore les heuristiques)\n\n"
            "Conseil : commencer par PSM 6 pour les documents\n"
            "courants, PSM 11 pour les formulaires administratifs."
        )

        # Trouver l'option courante dans la liste
        current_psm = str(self.config.tesseract_psm)
        default_psm_idx = 2  # PSM 6 par défaut
        for i, opt in enumerate(PSM_OPTIONS):
            if opt.startswith(current_psm + " ") or opt.startswith(current_psm + "  "):
                default_psm_idx = i
                break

        layout = [
            [sg.Text("Paramètres", font=FONT_TITLE)],
            [sg.HorizontalSeparator()],
            # Chemins
            [sg.Text("Dossier racine GED (ex: D:\\)", size=(30, 1)),
             sg.Input(self.config.ged_root, key="-ROOT-", size=(35, 1)),
             sg.FolderBrowse()],
            [sg.Text("Dossier quarantaine", size=(30, 1)),
             sg.Input(self.config.quarantine_folder, key="-QFOLDER-", size=(35, 1)),
             sg.FolderBrowse()],
            [sg.Text("Chemin Tesseract OCR", size=(30, 1)),
             sg.Input(self.config.tesseract_path, key="-TESS-", size=(35, 1)),
             sg.FileBrowse()],
            [sg.Text("Dossier sortie scanner", size=(30, 1)),
             sg.Input(self.config.scanner_output_folder, key="-SCANFOLDER-", size=(35, 1)),
             sg.FolderBrowse()],
            [sg.Text("Langue OCR (ex: fra+eng)", size=(30, 1)),
             sg.Input(self.config.ocr_language, key="-LANG-", size=(15, 1))],
            [sg.HorizontalSeparator()],
            # Section OCR avancé
            [sg.Text("Paramètres OCR avancés", font=FONT_TITLE)],
            [
                sg.Text("Mode segmentation Tesseract (PSM)", size=(30, 1)),
                sg.Combo(
                    PSM_OPTIONS,
                    default_value=PSM_OPTIONS[default_psm_idx],
                    key="-PSM-",
                    size=(42, 1),
                    readonly=True,
                ),
                sg.Button("?", key="-PSM-HELP-", size=(3, 1), tooltip="Aide sur les modes PSM"),
            ],
            [
                sg.Text("Prétraitement OpenCV", size=(30, 1)),
                sg.Checkbox(
                    "Activer (débruitage + binarisation + deskew)",
                    default=self.config.enhanced_preprocessing,
                    key="-PREPROC-",
                ),
            ],
            [sg.Text(
                "⚠  Tester le prétraitement sur quelques documents avant de valider.",
                font=("Helvetica", 9), text_color="orange"
            )],
            [sg.HorizontalSeparator()],
            [sg.Button("Enregistrer", key="-SAVE-"),
             sg.Button("Annuler", key="-CANCEL-")],
        ]

        win = sg.Window("Paramètres", layout, modal=True)

        while True:
            ev, vals = win.read()
            if ev in (sg.WIN_CLOSED, "-CANCEL-"):
                break
            elif ev == "-PSM-HELP-":
                sg.popup_scrolled(
                    PSM_HELP,
                    title="Aide — Modes PSM Tesseract",
                    size=(55, 22),
                    font=FONT_MONO,
                )
            elif ev == "-SAVE-":
                # Extraire le numéro PSM depuis la sélection
                psm_str = vals["-PSM-"].split("—")[0].strip()
                self.config.set("paths", "ged_root", vals["-ROOT-"])
                self.config.set("paths", "quarantine_folder", vals["-QFOLDER-"])
                self.config.set("paths", "tesseract_path", vals["-TESS-"])
                self.config.set("paths", "scanner_output_folder", vals["-SCANFOLDER-"])
                self.config.set("app", "language", vals["-LANG-"])
                self.config.set("ocr", "tesseract_psm", psm_str)
                self.config.set(
                    "ocr", "enhanced_preprocessing",
                    "true" if vals["-PREPROC-"] else "false"
                )
                # Mettre à jour les instances en mémoire
                self.file_manager.ged_root = vals["-ROOT-"]
                self.file_manager.quarantine_folder = vals["-QFOLDER-"]
                self.scanner.output_folder = vals["-SCANFOLDER-"]
                self.ocr.language = vals["-LANG-"]
                self.ocr.tesseract_psm = int(psm_str)
                self.ocr.enhanced_preprocessing = vals["-PREPROC-"]
                sg.popup("Paramètres enregistrés.", title="Paramètres")
                break
        win.close()
