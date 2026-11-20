"""Accès à Outlook (COM) : dossiers, mails, pièces jointes, OCR.

Corrections par rapport à la version précédente :
  * Filtre de dates côté Python : le Restrict() dépendait de la locale
    Windows (format de date, AM/PM mélangé au 24h) et pouvait renvoyer des
    résultats faux SANS erreur. Les mails étant triés du plus récent au plus
    ancien, la lecture s'arrête au premier mail trop ancien.
  * Dossiers « système » reconnus par leur TYPE (DefaultItemType != courrier)
    et leur EntryID (Éléments envoyés, Corbeille...), plus par leur NOM : un
    dossier projet appelé « Archive », « Notes » ou « Contacts » n'est plus
    ignoré en silence. Le filtre par noms reste disponible, en option :
    OutlookConnector(skip_folder_names=SYSTEM_FOLDERS).
  * Un dossier dont tous les sous-dossiers sont inexploitables (calendrier...)
    est traité comme une feuille : ses mails ne sont plus perdus.
  * get_inline_images_text fonctionne enfin (la regex n'avait aucun groupe) ;
    les images sont indexées par Content-ID en un seul passage.
  * Pièces jointes : seuls les PDF/images exploitables sont enregistrés sur
    disque (avant, tout fichier était sauvegardé puis supprimé), limite de
    taille, 10 pages max, PDF chiffrés ignorés, noms de fichiers assainis.
  * OCR : image lue une seule fois (imdecode, compatible chemins non-ASCII),
    timeout Tesseract, repli sur « eng » si le pack « fra » est absent.
  * pypdf (successeur de PyPDF2) utilisé s'il est installé.
  * is_valid_email / parse_date délèguent à utils (plus de doublons qui
    divergent : info@ était accepté ici et rejeté ailleurs).
  * Dossier temporaire créé à la demande et supprimé par close() ;
    OutlookConnector est utilisable avec « with ».
  * root_name : chemin de la boîte de réception détecté à la connexion, donc
    correct quelle que soit la langue d'Outlook (« Bandeja de entrada »...).
  * Connexion : exception dédiée (OutlookConnectionError) avec la cause.
"""
import itertools
import logging
import os
import re
import shutil
import tempfile
import time
from datetime import datetime
from typing import Any, Generator, Iterable, List, Optional

import win32com.client

from utils import is_valid_email as _utils_is_valid_email
from utils import parse_date as _utils_parse_date

try:
    from pypdf import PdfReader
except ImportError:
    try:
        from PyPDF2 import PdfReader
    except ImportError:
        PdfReader = None

try:
    import pytesseract
    import cv2
    import numpy as np
    OCR_AVAILABLE = True
except ImportError:
    OCR_AVAILABLE = False
    pytesseract = None
    cv2 = None
    np = None

logger = logging.getLogger(__name__)

IMAGE_EXTENSIONS = ('.png', '.jpg', '.jpeg', '.bmp', '.tiff', '.tif')
OFFICE_EXTENSIONS = ('.doc', '.docx', '.xls', '.xlsx', '.ppt', '.pptx')

# Optionnel : à passer à OutlookConnector(skip_folder_names=SYSTEM_FOLDERS)
# pour retrouver l'ancien filtrage par noms.
SYSTEM_FOLDERS = {
    "éléments envoyés", "sent items", "brouillons", "drafts", "corbeille", "deleted items",
    "junk email", "courrier indésirable", "archive", "archives", "rss feeds", "recherches",
    "dossiers de recherche", "favoris", "notes", "journal", "tâches", "calendrier", "contacts",
    "problèmes de synchronisation", "sync issues",
    "historique des conversations", "conversation history",
}

# Codes OlDefaultFolders : Deleted, Outbox, Sent, Drafts, Junk
_SYSTEM_FOLDER_CODES = (3, 4, 5, 16, 23)
OL_FOLDER_INBOX = 6
OL_MAIL_ITEM = 0                      # DefaultItemType d'un dossier de courrier

PR_ATTACH_CONTENT_ID = "http://schemas.microsoft.com/mapi/proptag/0x3712001E"

OCR_JUNK_PATTERNS = (
    re.compile(r"^(https?|www|mailto|ftp)\s*[:.]?\s*$", re.IGNORECASE),
    re.compile(r"^[^a-zA-Z0-9]+$"),
    re.compile(r"^[A-Z]\.([A-Z]\.)+$"),
)
MIN_IMAGE_SIZE = 50
MAX_ATTACHMENT_BYTES = 15 * 1024 * 1024
PDF_MAX_PAGES = 10
PDF_MAX_CHARS = 2000
OCR_TIMEOUT_S = 30
INLINE_MAX_IMAGES = 10


class OutlookConnectionError(Exception):
    """Impossible de se connecter à Outlook."""


def _safe_get(obj, attr, default=None):
    """getattr qui absorbe aussi les com_error (pas seulement AttributeError)."""
    if obj is None:
        return default
    try:
        return getattr(obj, attr)
    except Exception:
        return default


def _remove_file(path: Optional[str]):
    if path:
        try:
            os.remove(path)
        except Exception:
            pass


def _sanitize_ocr_text(text: str, max_chars: int = 500) -> str:
    """Filtre le texte OCR/PDF avant de le renvoyer."""
    if not text:
        return ""
    kept: List[str] = []
    for line in text.splitlines():
        line = line.strip()
        if not line or len(line) < 3:
            continue
        if any(p.match(line) for p in OCR_JUNK_PATTERNS):
            continue
        alnum = sum(c.isalnum() or c in " .,'&-+/()@" for c in line)
        if alnum / len(line) < 0.6:
            continue
        kept.append(line)
        if sum(len(k) for k in kept) > max_chars:
            break
    return "\n".join(kept)


# ----------------------------------------------------------------------
# Dossier
# ----------------------------------------------------------------------

class OutlookFolder:
    # Renseigné par OutlookConnector à la connexion : chemin de la boîte de
    # réception, quelle que soit la langue d'Outlook.
    known_inbox_path: Optional[str] = None

    INBOX_NAMES = {
        "boîte de réception", "inbox", "bandeja de entrada", "posteingang",
        "posta in arrivo", "caixa de entrada", "postvak in", "entrada",
    }

    def __init__(self, name: str, path: str, folder: Any = None):
        self.name = name
        self.path = path
        self.folder = folder
        self.entry_id = _safe_get(folder, "EntryID")
        self.root_name = self._compute_root_name(name, path)

    @classmethod
    def _compute_root_name(cls, name: str, path: str) -> str:
        """Premier niveau sous la boîte de réception (= « famille client »)."""
        try:
            known = (cls.known_inbox_path or "").rstrip("\\")
            if known and path.lower().startswith(known.lower() + "\\"):
                rest = [p.strip() for p in path[len(known) + 1:].split("\\") if p.strip()]
                if rest:
                    return rest[0]

            parts = [p.strip() for p in path.split("\\") if p.strip()]
            for i, part in enumerate(parts):
                if part.lower() in cls.INBOX_NAMES:
                    if len(parts) > i + 1:
                        return parts[i + 1]
                    break
            if len(parts) >= 3:
                return parts[2]
        except Exception:
            pass
        return name


# ----------------------------------------------------------------------
# Connecteur
# ----------------------------------------------------------------------

class OutlookConnector:
    def __init__(self, ocr_inline_images: bool = False,
                 skip_folder_names: Iterable[str] = ()):
        self.outlook = None
        self.namespace = None
        self.mailbox = None                 # dossier Boîte de réception
        self.ocr_inline_images = ocr_inline_images
        self.skip_folder_names = {n.lower().strip() for n in skip_folder_names}
        self._temp_counter = itertools.count()
        self._temp_dir: Optional[str] = None
        self._system_ids: Optional[set] = None
        self._ocr_lang = "fra+eng"
        self._warned: set = set()
        self._connect()
        OutlookFolder.known_inbox_path = _safe_get(self.mailbox, "FolderPath")

    # -- cycle de vie ---------------------------------------------------
    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc, tb):
        self.close()
        return False

    def __del__(self):
        try:
            self.close()
        except Exception:
            pass

    def close(self):
        self.outlook = None
        self.namespace = None
        self.mailbox = None
        temp_dir = getattr(self, "_temp_dir", None)
        if temp_dir:
            self._temp_dir = None
            try:
                shutil.rmtree(temp_dir, ignore_errors=True)
                logger.info(f"[OUTLOOK] Dossier temporaire nettoye: {temp_dir}")
            except Exception:
                pass

    @property
    def temp_dir(self) -> str:
        """Créé à la demande (inutile si aucune pièce jointe / OCR)."""
        if not self._temp_dir:
            self._temp_dir = tempfile.mkdtemp(prefix="outlook_ocr_")
            logger.info(f"[OUTLOOK] Dossier temporaire pour OCR: {self._temp_dir}")
        return self._temp_dir

    def _warn_once(self, key: str, message: str):
        if key not in self._warned:
            self._warned.add(key)
            logger.warning(message)

    def _connect(self, max_retries: int = 3, retry_delay: int = 2):
        last_error: Optional[Exception] = None
        for attempt in range(1, max_retries + 1):
            try:
                self.outlook = win32com.client.Dispatch("Outlook.Application")
                self.namespace = self.outlook.GetNamespace("MAPI")
                self.mailbox = self.namespace.GetDefaultFolder(OL_FOLDER_INBOX)
                logger.info("[OUTLOOK] Connecté à Outlook")
                return
            except Exception as e:
                last_error = e
                logger.warning(f"[OUTLOOK] Échec {attempt}/{max_retries}: {e}")
                if attempt < max_retries:
                    time.sleep(retry_delay)
        raise OutlookConnectionError(
            f"Impossible de se connecter à Outlook ({last_error})") from last_error

    # -- chemins temporaires ---------------------------------------------
    def _unique_temp_path(self, file_name: str) -> str:
        safe = os.path.basename(file_name or "file")
        safe = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", safe)
        root, ext = os.path.splitext(safe)
        return os.path.join(
            self.temp_dir, f"{next(self._temp_counter)}_{root[:60]}{ext[:10]}")

    # ------------------------------------------------------------------
    # Dossiers
    # ------------------------------------------------------------------

    def get_inbox(self) -> Any:
        return self.namespace.GetDefaultFolder(OL_FOLDER_INBOX)

    def get_folder_by_entry_id(self, entry_id: str) -> Optional[Any]:
        try:
            return self.namespace.GetFolderFromID(entry_id)
        except Exception as e:
            logger.error(f"[OUTLOOK] Erreur get_folder_by_entry_id: {e}")
            return None

    def get_folder_hierarchy(self, folder: 'OutlookFolder') -> List[str]:
        return [p.strip() for p in folder.path.split('\\') if p.strip()]

    def _is_system_folder(self, folder_name: str) -> bool:
        """Filtre OPTIONNEL par nom (skip_folder_names)."""
        return folder_name.lower().strip() in self.skip_folder_names

    def _get_system_entry_ids(self) -> set:
        if self._system_ids is None:
            ids = set()
            for code in _SYSTEM_FOLDER_CODES:
                try:
                    ids.add(self.namespace.GetDefaultFolder(code).EntryID)
                except Exception:
                    pass
            self._system_ids = ids
        return self._system_ids

    def _should_skip_folder(self, folder: Any) -> bool:
        name = _safe_get(folder, "Name", "") or ""
        if self.skip_folder_names and self._is_system_folder(name):
            return True
        # Calendrier, contacts, tâches, notes, journal... : pas du courrier
        if _safe_get(folder, "DefaultItemType", OL_MAIL_ITEM) != OL_MAIL_ITEM:
            logger.debug(f"[OUTLOOK] Dossier ignore (pas du courrier): {name}")
            return True
        entry_id = _safe_get(folder, "EntryID")
        if entry_id and entry_id in self._get_system_entry_ids():
            return True
        return False

    def _eligible_subfolders(self, parent: Any) -> List[Any]:
        """Sous-dossiers directs exploitables (courrier, accessibles)."""
        out: List[Any] = []
        folders = _safe_get(parent, "Folders")
        if folders is None:
            return out
        try:
            for sub in folders:
                try:
                    if self._should_skip_folder(sub):
                        continue
                    _ = sub.Items.Count          # lève si inaccessible
                    out.append(sub)
                except Exception:
                    continue
        except Exception as e:
            logger.error(f"[OUTLOOK] Erreur lecture dossiers: {e}")
        return out

    def _is_leaf_folder(self, folder: Any) -> bool:
        """Feuille = aucun sous-dossier EXPLOITABLE."""
        try:
            return not self._eligible_subfolders(folder)
        except Exception:
            return True

    def _iter_leaves(self, folder: Any) -> Generator[Any, None, None]:
        children = self._eligible_subfolders(folder)
        if not children:
            yield folder
            return
        for child in children:
            yield from self._iter_leaves(child)

    def get_direct_subfolders(self, parent_folder: Any = None) -> List[OutlookFolder]:
        if parent_folder is None:
            parent_folder = self.get_inbox()
        return [OutlookFolder(sub.Name, sub.FolderPath, sub)
                for sub in self._eligible_subfolders(parent_folder)]

    def get_all_subfolders(self, parent_folder: Any = None) -> Generator[OutlookFolder, None, None]:
        if parent_folder is None:
            parent_folder = self.get_inbox()
        for sub in self._eligible_subfolders(parent_folder):
            yield OutlookFolder(sub.Name, sub.FolderPath, sub)
            yield from self.get_all_subfolders(sub)

    def get_leaf_folders(self, parent_folder: Any = None) -> Generator[OutlookFolder, None, None]:
        """Feuilles sous parent_folder (parent_folder lui-même n'est jamais renvoyé)."""
        if parent_folder is None:
            parent_folder = self.get_inbox()
        for sub in self._eligible_subfolders(parent_folder):
            for leaf in self._iter_leaves(sub):
                yield OutlookFolder(leaf.Name, leaf.FolderPath, leaf)

    # ------------------------------------------------------------------
    # Mails
    # ------------------------------------------------------------------

    def get_mails(self, folder: Any, date_from: Optional[datetime] = None,
                  date_to: Optional[datetime] = None,
                  limit: Optional[int] = None) -> Generator[Any, None, None]:
        """Renvoie les mails d'un dossier, du plus récent au plus ancien.

        Le filtre de dates est appliqué ICI (et non via Restrict, dépendant de
        la locale Windows). Tant que le tri Outlook a fonctionné, la lecture
        s'arrête au premier mail antérieur à date_from."""
        try:
            items = _safe_get(folder, "Items")
            if items is None:
                return
            sorted_ok = True
            try:
                items.Sort("[ReceivedTime]", True)
            except Exception:
                sorted_ok = False

            count = 0
            for item in items:
                try:
                    message_class = _safe_get(item, "MessageClass", "")
                    if message_class and not str(message_class).upper().startswith("IPM.NOTE"):
                        continue

                    if date_from or date_to:
                        received = _utils_parse_date(_safe_get(item, "ReceivedTime"))
                        if received:
                            if date_to and received > date_to:
                                continue
                            if date_from and received < date_from:
                                if sorted_ok:
                                    break          # tout le reste est plus ancien
                                continue
                except Exception:
                    continue

                yield item
                count += 1
                if limit is not None and count >= limit:
                    break
        except Exception as e:
            logger.error(f"[OUTLOOK] Erreur get_mails: {e}")

    # ------------------------------------------------------------------
    # Pièces jointes
    # ------------------------------------------------------------------

    @staticmethod
    def _pdf_text(path: str) -> str:
        with open(path, 'rb') as f:
            reader = PdfReader(f)
            if getattr(reader, "is_encrypted", False):
                try:
                    if not reader.decrypt(""):
                        return ""
                except Exception:
                    return ""
            chunks: List[str] = []
            total = 0
            for page in list(reader.pages)[:PDF_MAX_PAGES]:
                try:
                    page_text = page.extract_text() or ""
                except Exception:
                    page_text = ""
                if page_text:
                    chunks.append(page_text)
                    total += len(page_text)
                if total >= PDF_MAX_CHARS:
                    break
        return _sanitize_ocr_text("\n".join(chunks))

    def get_attachments_text(self, mail: Any, skip_pdf: bool = False,
                             skip_images: bool = False, skip_office: bool = False) -> str:
        can_pdf = PdfReader is not None and not skip_pdf
        can_img = OCR_AVAILABLE and not skip_images
        if PdfReader is None and not OCR_AVAILABLE:
            self._warn_once("no_extractors",
                            "[OUTLOOK] pypdf/PyPDF2 et pytesseract non installes.")
        if not (can_pdf or can_img):
            return ""

        attachments = _safe_get(mail, "Attachments")
        if attachments is None:
            return ""
        try:
            items = list(attachments)
        except Exception:
            return ""

        texts: List[str] = []
        for attachment in items:
            file_name = None
            temp_path = None
            try:
                file_name = _safe_get(attachment, "FileName")
                if not isinstance(file_name, str) or not file_name:
                    continue
                lowered = file_name.lower()
                is_pdf = lowered.endswith('.pdf')
                is_image = lowered.endswith(IMAGE_EXTENSIONS)

                # Rien n'est enregistré sur disque si le fichier n'est pas exploitable
                if is_pdf and not can_pdf:
                    continue
                if is_image and not can_img:
                    continue
                if not (is_pdf or is_image):
                    if lowered.endswith(OFFICE_EXTENSIONS) and not skip_office:
                        logger.debug(f"[OUTLOOK] Piece Office non extractible: {file_name}")
                    continue

                size = _safe_get(attachment, "Size", 0)
                if isinstance(size, int) and size > MAX_ATTACHMENT_BYTES:
                    logger.debug(f"[OUTLOOK] Piece jointe trop volumineuse ignoree: {file_name}")
                    continue

                temp_path = self._unique_temp_path(file_name)
                attachment.SaveAsFile(temp_path)
                text = (self._pdf_text(temp_path) if is_pdf
                        else self._extract_signature_text(temp_path))
                if text:
                    texts.append(text)
            except Exception as e:
                logger.warning(f"[OUTLOOK] Erreur piece jointe {file_name or '<inconnu>'}: {e}")
            finally:
                _remove_file(temp_path)
        return "\n".join(texts)

    # ------------------------------------------------------------------
    # OCR
    # ------------------------------------------------------------------

    @staticmethod
    def _load_image(path: str):
        """Lecture robuste (chemins non-ASCII) ; None si illisible."""
        try:
            data = np.fromfile(path, dtype=np.uint8)
            return cv2.imdecode(data, cv2.IMREAD_COLOR)
        except Exception:
            return None

    @staticmethod
    def _is_probable_tracking_image(image) -> bool:
        try:
            h, w = image.shape[0], image.shape[1]
            return h < MIN_IMAGE_SIZE or w < MIN_IMAGE_SIZE
        except Exception:
            return False

    def _ocr_image(self, image) -> str:
        if image.shape[0] > 2000 or image.shape[1] > 2000:
            image = cv2.resize(image, (0, 0), fx=0.5, fy=0.5)
        gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
        gray = cv2.medianBlur(gray, 3)
        thresh = cv2.adaptiveThreshold(
            gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 11, 2)
        if np.mean(thresh) < 128:
            thresh = cv2.bitwise_not(thresh)
        for _ in range(2):
            config = f"--oem 3 --psm 6 -l {self._ocr_lang}"
            try:
                return pytesseract.image_to_string(
                    thresh, config=config, timeout=OCR_TIMEOUT_S).strip()
            except pytesseract.TesseractError as e:
                # pack de langue manquant : repli unique sur l'anglais
                if "language" in str(e).lower() and self._ocr_lang != "eng":
                    self._warn_once("ocr_lang",
                                    "[OUTLOOK] Langue OCR 'fra' indisponible, repli sur 'eng'.")
                    self._ocr_lang = "eng"
                    continue
                raise
        return ""

    def _extract_text_from_image(self, image_path: str) -> str:
        if not OCR_AVAILABLE:
            return ""
        try:
            image = self._load_image(image_path)
            if image is None:
                return ""
            return self._ocr_image(image)
        except Exception as e:
            logger.warning(f"[OUTLOOK] Erreur OCR sur {image_path}: {e}")
            return ""

    def _extract_signature_text(self, image_path: str) -> str:
        if not OCR_AVAILABLE:
            return ""
        try:
            image = self._load_image(image_path)
            if image is None:
                return ""
            if self._is_probable_tracking_image(image):
                logger.debug(f"[OUTLOOK] Image trop petite ignoree: {image_path}")
                return ""
            return _sanitize_ocr_text(self._ocr_image(image))
        except Exception as e:
            logger.warning(f"[OUTLOOK] Erreur OCR signature {image_path}: {e}")
            return ""

    def get_inline_images_text(self, mail: Any) -> str:
        """OCR des images incrustées (cid:) du corps HTML."""
        if not OCR_AVAILABLE or not self.ocr_inline_images:
            return ""
        texts: List[str] = []
        try:
            html_body = _safe_get(mail, "HTMLBody")
            if not isinstance(html_body, str) or not html_body:
                return ""

            cids: List[str] = []
            for img_tag in re.findall(r'<img[^>]+>', html_body, re.IGNORECASE):
                m = re.search(r'src=["\']([^"\']+)["\']', img_tag, re.IGNORECASE)
                if m and m.group(1).lower().startswith("cid:"):
                    cid = m.group(1)[4:]
                    if cid not in cids:
                        cids.append(cid)
                if len(cids) >= INLINE_MAX_IMAGES:
                    break
            if not cids:
                return ""

            wanted = set(cids)
            for attachment in list(mail.Attachments):         # un seul passage
                try:
                    content_id = attachment.PropertyAccessor.GetProperty(PR_ATTACH_CONTENT_ID)
                except Exception:
                    continue
                if content_id not in wanted:
                    continue
                wanted.discard(content_id)
                temp_path = None
                try:
                    safe_cid = re.sub(r'[^A-Za-z0-9_.@-]', '_', content_id)[:80]
                    temp_path = self._unique_temp_path(f"{safe_cid}.png")
                    attachment.SaveAsFile(temp_path)
                    text = self._extract_signature_text(temp_path)
                    if text:
                        texts.append(text)
                except Exception as e:
                    logger.debug(f"[OUTLOOK] Image inline illisible ({content_id}): {e}")
                finally:
                    _remove_file(temp_path)
                if not wanted:
                    break
        except Exception as e:
            logger.warning(f"[OUTLOOK] Erreur extraction images inline: {e}")
        return "\n".join(texts)

    # ------------------------------------------------------------------
    # Compatibilité : délèguent à utils (une seule définition)
    # ------------------------------------------------------------------

    def is_valid_email(self, email: str) -> bool:
        return _utils_is_valid_email(email)

    def parse_date(self, date_obj) -> Optional[datetime]:
        return _utils_parse_date(date_obj)