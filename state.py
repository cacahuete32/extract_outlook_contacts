"""Checkpoint d'extraction : mails déjà traités (EntryID) et watermark par dossier.

Corrections par rapport à la version précédente :
  * L'état peut être LIÉ au fichier de sortie : ExtractionState(output_dir, name="x")
    utilise extraction_state_x.json. Avant, tous les fichiers Excel d'un même
    répertoire partageaient un seul état : lancer un second export (autre nom)
    ignorait les mails « déjà traités » par le premier.
  * Migration transparente : si le fichier nommé n'existe pas encore mais que
    l'ancien extraction_state.json existe, il sert de point de départ (le
    fichier nommé est créé au premier save()).
  * Fichier d'état corrompu : conservé en .corrupt (avant, il était écrasé au
    prochain save) puis scan complet.
  * reset() remplace la remise à zéro manuelle des deux attributs.
  * Sans argument `name`, le comportement est identique à avant.
"""
import json
import os
import re
import logging
from typing import Dict, Optional, Set

logger = logging.getLogger(__name__)

STATE_VERSION = 1


class ExtractionState:
    """Mémorise les mails déjà traités (EntryID) et la date du dernier scan."""

    def __init__(self, output_dir: str = "output", name: Optional[str] = None):
        self.output_dir = output_dir
        safe = re.sub(r"[^\w.-]+", "_", name or "").strip("._")
        filename = f"extraction_state_{safe}.json" if safe else "extraction_state.json"
        self.path = os.path.join(output_dir, filename)
        # ancien fichier partagé : point de départ si le fichier nommé n'existe pas
        self._legacy_path = (os.path.join(output_dir, "extraction_state.json")
                             if safe else None)

        self.processed_mail_ids: Set[str] = set()
        self.folder_watermarks: Dict[str, str] = {}   # folder_path -> date ISO du dernier mail traité
        self._load()

    def _load(self):
        source = self.path
        if not os.path.exists(source):
            if self._legacy_path and os.path.exists(self._legacy_path):
                source = self._legacy_path
                logger.info(f"[STATE] Migration de l'ancien etat partage: {source}")
            else:
                return
        try:
            with open(source, "r", encoding="utf-8") as f:
                data = json.load(f)
            self.processed_mail_ids = set(data.get("mail_ids", []))
            self.folder_watermarks = dict(data.get("watermarks", {}))
            logger.info(f"[STATE] {len(self.processed_mail_ids)} mails deja traites charges")
        except Exception as e:
            logger.warning(f"[STATE] Fichier d'etat illisible ({e}) : conserve en .corrupt, scan complet")
            self.processed_mail_ids = set()
            self.folder_watermarks = {}
            try:
                os.replace(source, source + ".corrupt")
            except Exception:
                pass

    def save(self):
        tmp = self.path + ".tmp"
        try:
            os.makedirs(os.path.dirname(self.path) or ".", exist_ok=True)
            with open(tmp, "w", encoding="utf-8") as f:
                json.dump({
                    "version": STATE_VERSION,
                    "mail_ids": sorted(self.processed_mail_ids),
                    "watermarks": self.folder_watermarks,
                }, f)
            os.replace(tmp, self.path)   # écriture atomique
        except Exception as e:
            logger.warning(f"[STATE] Sauvegarde impossible: {e}")
            try:
                os.remove(tmp)
            except Exception:
                pass

    def reset(self):
        """Oublie tout (rescan complet). N'écrit rien tant que save() n'est pas appelé."""
        self.processed_mail_ids = set()
        self.folder_watermarks = {}

    def get_watermark(self, folder_path: str) -> Optional[str]:
        return self.folder_watermarks.get(folder_path)

    def mark_processed(self, mail_id: str, folder_path: str, received_iso: str):
        if mail_id:
            self.processed_mail_ids.add(mail_id)
        if folder_path and received_iso:
            current = self.folder_watermarks.get(folder_path)
            if not current or received_iso > current:
                self.folder_watermarks[folder_path] = received_iso