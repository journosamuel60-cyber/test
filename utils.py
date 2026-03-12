"""
utils.py
Utilitaires généraux : logging, gestion des fichiers temporaires,
validation des données, formatage des scores.
"""

import hashlib
import json
import logging
import os
import time
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional

# ─────────────────────────────────────────────
# CONFIGURATION DU LOGGING
# ─────────────────────────────────────────────

def setup_logging(log_level: str = "INFO", log_file: Optional[str] = None):
    """Configure le système de logging."""
    handlers = [logging.StreamHandler()]
    if log_file:
        handlers.append(logging.FileHandler(log_file, encoding="utf-8"))

    logging.basicConfig(
        level=getattr(logging, log_level.upper(), logging.INFO),
        format="%(asctime)s | %(levelname)-8s | %(name)s | %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
        handlers=handlers
    )


# ─────────────────────────────────────────────
# JOURNAL DES TRAITEMENTS
# ─────────────────────────────────────────────

class ProcessingLog:
    """Journalise les traitements de fichiers."""

    def __init__(self, log_path: str = "processing_log.json"):
        self.log_path = log_path
        self.entries: List[Dict] = []
        self._load()

    def _load(self):
        if os.path.exists(self.log_path):
            try:
                with open(self.log_path, "r", encoding="utf-8") as f:
                    self.entries = json.load(f)
            except Exception:
                self.entries = []

    def add_entry(
        self,
        filename: str,
        status: str,
        fields_found: int,
        fields_total: int,
        errors: Optional[List[str]] = None,
        processing_time: Optional[float] = None
    ):
        entry = {
            "timestamp": datetime.now().isoformat(),
            "filename": filename,
            "status": status,
            "fields_found": fields_found,
            "fields_total": fields_total,
            "success_rate": f"{fields_found/fields_total*100:.0f}%" if fields_total > 0 else "0%",
            "errors": errors or [],
            "processing_time_sec": round(processing_time, 2) if processing_time else None
        }
        self.entries.append(entry)
        self._save()

    def _save(self):
        try:
            with open(self.log_path, "w", encoding="utf-8") as f:
                json.dump(self.entries, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    def get_summary(self) -> Dict:
        if not self.entries:
            return {}
        total = len(self.entries)
        success = sum(1 for e in self.entries if e["status"] == "success")
        return {
            "total_processed": total,
            "success": success,
            "errors": total - success,
            "success_rate": f"{success/total*100:.0f}%"
        }


# ─────────────────────────────────────────────
# VALIDATION & FORMATAGE
# ─────────────────────────────────────────────

def format_confidence(score: float) -> str:
    """Convertit un score (0-1) en label lisible avec émoji."""
    if score >= 0.8:
        return f"🟢 {score*100:.0f}%"
    elif score >= 0.5:
        return f"🟡 {score*100:.0f}%"
    elif score > 0:
        return f"🔴 {score*100:.0f}%"
    else:
        return "⚪ Non trouvé"


def count_found_fields(merged_results: dict) -> tuple:
    """Retourne (nb_trouvés, nb_total)."""
    total = len(merged_results)
    found = sum(1 for v in merged_results.values() if v.get("value", "Non trouvé") != "Non trouvé")
    return found, total


def safe_filename(name: str) -> str:
    """Sanitize un nom de fichier."""
    import re
    return re.sub(r'[<>:"/\\|?*]', '_', name)


def compute_file_hash(file_path: str) -> str:
    """Calcule le hash MD5 d'un fichier (pour détecter les doublons)."""
    h = hashlib.md5()
    with open(file_path, "rb") as f:
        for chunk in iter(lambda: f.read(8192), b""):
            h.update(chunk)
    return h.hexdigest()


# ─────────────────────────────────────────────
# GESTION DES FICHIERS TEMPORAIRES
# ─────────────────────────────────────────────

def save_uploaded_file(uploaded_file, tmp_dir: str = "/tmp") -> str:
    """
    Sauvegarde un fichier uploadé (Streamlit UploadedFile) sur disque.
    Retourne le chemin du fichier temporaire.
    """
    Path(tmp_dir).mkdir(parents=True, exist_ok=True)
    tmp_path = os.path.join(tmp_dir, safe_filename(uploaded_file.name))
    with open(tmp_path, "wb") as f:
        f.write(uploaded_file.getvalue())
    return tmp_path
