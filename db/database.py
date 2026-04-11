"""
db/database.py
Base de données SQLite pour GED-Manager
Tables : mapping, document_history, quarantine
"""

import sqlite3
import os
from datetime import datetime


class Database:
    def __init__(self, db_file="ged_manager.db"):
        self.db_file = db_file
        self.conn = sqlite3.connect(db_file, check_same_thread=False)
        self.conn.row_factory = sqlite3.Row
        self._create_tables()

    def _create_tables(self):
        cursor = self.conn.cursor()

        cursor.executescript("""
            CREATE TABLE IF NOT EXISTS mapping (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                keyword TEXT NOT NULL,
                target_folder TEXT NOT NULL,
                created_at TEXT DEFAULT (datetime('now')),
                updated_at TEXT DEFAULT (datetime('now'))
            );

            CREATE UNIQUE INDEX IF NOT EXISTS idx_mapping_keyword
                ON mapping(keyword);

            CREATE TABLE IF NOT EXISTS document_history (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                filename TEXT NOT NULL,
                source_path TEXT,
                destination_path TEXT,
                detected_keywords TEXT,
                matched_keyword TEXT,
                file_hash TEXT,
                action TEXT,
                processed_at TEXT DEFAULT (datetime('now'))
            );

            CREATE TABLE IF NOT EXISTS quarantine (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                filename TEXT NOT NULL,
                file_path TEXT NOT NULL,
                detected_keywords TEXT,
                confidence REAL DEFAULT 0.0,
                status TEXT DEFAULT 'pending',
                created_at TEXT DEFAULT (datetime('now'))
            );
        """)
        self.conn.commit()

    # ── MAPPING ──────────────────────────────────────────────────────────────

    def add_mapping(self, keyword, target_folder):
        """Ajouter ou mettre à jour un mappage mot-clé → dossier"""
        cursor = self.conn.cursor()
        cursor.execute("""
            INSERT INTO mapping (keyword, target_folder, updated_at)
            VALUES (?, ?, datetime('now'))
            ON CONFLICT(keyword) DO UPDATE SET
                target_folder=excluded.target_folder,
                updated_at=datetime('now')
        """, (keyword.lower().strip(), target_folder))
        self.conn.commit()

    def get_mapping(self, keyword):
        """Récupérer le dossier cible pour un mot-clé"""
        cursor = self.conn.cursor()
        row = cursor.execute(
            "SELECT target_folder FROM mapping WHERE keyword = ?",
            (keyword.lower().strip(),)
        ).fetchone()
        return row["target_folder"] if row else None

    def get_all_mappings(self):
        """Récupérer tous les mappages"""
        cursor = self.conn.cursor()
        rows = cursor.execute(
            "SELECT keyword, target_folder, updated_at FROM mapping ORDER BY keyword"
        ).fetchall()
        return [dict(r) for r in rows]

    def delete_mapping(self, keyword):
        cursor = self.conn.cursor()
        cursor.execute("DELETE FROM mapping WHERE keyword = ?", (keyword.lower().strip(),))
        self.conn.commit()

    # ── HISTORIQUE ───────────────────────────────────────────────────────────

    def add_history(self, filename, source_path, destination_path,
                    detected_keywords, matched_keyword, file_hash, action):
        cursor = self.conn.cursor()
        cursor.execute("""
            INSERT INTO document_history
                (filename, source_path, destination_path, detected_keywords,
                 matched_keyword, file_hash, action)
            VALUES (?, ?, ?, ?, ?, ?, ?)
        """, (filename, source_path, destination_path,
              detected_keywords, matched_keyword, file_hash, action))
        self.conn.commit()

    def get_history(self, limit=100):
        cursor = self.conn.cursor()
        rows = cursor.execute(
            "SELECT * FROM document_history ORDER BY processed_at DESC LIMIT ?",
            (limit,)
        ).fetchall()
        return [dict(r) for r in rows]

    def search_history(self, query):
        cursor = self.conn.cursor()
        q = f"%{query}%"
        rows = cursor.execute("""
            SELECT * FROM document_history
            WHERE filename LIKE ? OR destination_path LIKE ?
               OR detected_keywords LIKE ?
            ORDER BY processed_at DESC
        """, (q, q, q)).fetchall()
        return [dict(r) for r in rows]

    # ── QUARANTAINE ───────────────────────────────────────────────────────────

    def add_quarantine(self, filename, file_path, detected_keywords, confidence=0.0):
        cursor = self.conn.cursor()
        cursor.execute("""
            INSERT INTO quarantine (filename, file_path, detected_keywords, confidence)
            VALUES (?, ?, ?, ?)
        """, (filename, file_path, detected_keywords, confidence))
        self.conn.commit()

    def get_quarantine(self):
        cursor = self.conn.cursor()
        rows = cursor.execute(
            "SELECT * FROM quarantine WHERE status='pending' ORDER BY created_at DESC"
        ).fetchall()
        return [dict(r) for r in rows]

    def resolve_quarantine(self, item_id):
        cursor = self.conn.cursor()
        cursor.execute(
            "UPDATE quarantine SET status='resolved' WHERE id=?", (item_id,)
        )
        self.conn.commit()

    def close(self):
        self.conn.close()
