"""Export Excel des contacts.

Corrections par rapport à la version précédente :
  * Fusion incrémentale CHAMP PAR CHAMP : une ligne connue (poste, téléphone...)
    n'est plus écrasée par une ligne plus pauvre issue du run courant. Une
    source moins fiable (en-tête) ne remplace pas une valeur plus fiable
    (Exchange). La dernière activité et le dernier sujet restent cohérents
    (on garde le mail le plus récent).
  * La clé de fusion (Email, Famille Client, Projet) est normalisée
    UNIQUEMENT pour comparer : les valeurs écrites gardent leur casse
    (« ClientA » ne devient plus « clienta »). Les doublons créés par
    l'ancien bug sont fusionnés à la lecture.
  * Fichier existant illisible = ERREUR (ExistingFileError) au lieu d'être
    écrasé silencieusement ; sauvegarde .bak avant toute écriture.
  * Écriture atomique (fichier temporaire puis os.replace) : un plantage ou
    un fichier verrouillé ne laisse jamais un classeur à moitié écrit.
    Fichier ouvert dans Excel -> PermissionError (la GUI bascule alors sur un
    fichier de secours).
  * Relecture en texte (dtype=str) : les zéros de tête des téléphones sont
    conservés, plus de NaN / Timestamp parasites.
  * Téléphone : zéro initial conservé, 00xx converti en +xx, numéros
    aberrants (> 15 chiffres) rejetés.
  * Poste : les postes en minuscules (« chef de projet ») ne sont plus supprimés.
  * Noms à un mot : nom de famille laissé vide (plus de Prénom == Nom) ;
    « DUPONT Jean » -> « Jean DUPONT ».
  * Robustesse Excel : caractères de contrôle supprimés (openpyxl plantait),
    valeurs commençant par « = » (sujets encodés « =?UTF-8?... ») écrites
    comme du texte et non comme des formules.
  * COHÉRENCE : 'x.com' testé avec frontière de mot ; nom de feuille Excel : le
    caractère antislash est de nouveau remplacé (la classe de caractères l'avait perdu).
  * EMAILS HORS PROJETS : feuille 'Emails hors projets' listant les mails des
    dossiers non-projet (parents/intermediaires) dont l'entreprise de
    l'expediteur est identifiee (signature ou domaine du mail) ; fusion sans
    doublons entre les lancements (cle Email + Sujet + Date).
"""
import os
import re
import shutil
import logging
from collections import Counter
from typing import Dict, List, Optional
from datetime import datetime

import pandas as pd
from openpyxl.utils import get_column_letter

from extractor import ContactInfo

logger = logging.getLogger(__name__)

EMAIL_RE = re.compile(r"^[\w.+-]+@[\w-]+\.[\w.-]+$")
URL_RE = re.compile(r"^https?://[^\s]+$", re.IGNORECASE)
VALID_DOMAIN_URL_RE = re.compile(
    r"^https?://([A-Za-z0-9-]+\.)+[A-Za-z]{2,}(/[^\s]*)?$"
)
JUNK_URL_KEYWORDS = (
    "googleusercontent", "googlegroups", "outlook.office", "track",
    "click", "signature", "proxy", "bookwithme",
    "teams.microsoft", "meet.google", "zoom.us", "webex",
    "calendly.com", "bit.ly", "tinyurl", "linkedin.com",
    "facebook.com", "twitter.com",
)
_JUNK_X_RE = re.compile(r"(?<![\w-])x\.com", re.IGNORECASE)      # frontière de mot
MAILTO_RE = re.compile(r"^mailto:", re.IGNORECASE)
DATE_LIKE_RE = re.compile(r"^\d{4}-\d{2}-\d{2}([ T]\d{2}:\d{2}(:\d{2})?)?$")
CONTROL_CHARS_RE = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]")

COLUMNS = [
    "Email", "Nom", "Prenom", "Nom de famille", "Poste", "Entreprise",
    "Telephone", "Famille Client", "Projet", "Derniere activite",
    "Dernier sujet", "Site web", "Adresse", "Source",
]
MERGE_KEY = ["Email", "Famille Client", "Projet"]
NAME_COLS = ("Nom", "Prenom", "Nom de famille")
NO_PROJECT_COLUMNS = ["Date", "Nom", "Email", "Entreprise", "Dossier", "Sujet"]
NO_PROJECT_SHEET = "Emails hors projets"
ACTIVITY_COLS = ("Derniere activite", "Dernier sujet")

SOURCE_LABELS = {
    "exchange": "Exchange (annuaire)",
    "signature": "Signature email",
    "body": "Corps du message",
    "header": "En-tete email",
}
SOURCE_DISPLAY_PRIORITY = ["exchange", "signature", "body", "header"]
_SOURCE_RANK = {"exchange": 2, "signature": 1, "body": 1, "header": 0}
_LABEL_TO_KEY = {v: k for k, v in SOURCE_LABELS.items()}


class ExistingFileError(RuntimeError):
    """Le fichier Excel existe mais n'a pas pu être relu : on refuse de l'écraser."""


# ----------------------------------------------------------------------
# Nettoyage de champs
# ----------------------------------------------------------------------

def _clean(value) -> str:
    if value is None:
        return ""
    # Un datetime COM (pywintypes) dans un champ texte = donnée parasite
    if isinstance(value, datetime) or not isinstance(value, str):
        return ""
    s = CONTROL_CHARS_RE.sub("", value).strip()
    s = s.strip("<>")
    s = MAILTO_RE.sub("", s).strip()
    s = re.sub(r"\s+", " ", s)
    return s


def _is_email(value: str) -> bool:
    return bool(value) and bool(EMAIL_RE.match(value))


def _is_url(value: str) -> bool:
    return bool(value) and bool(URL_RE.match(value))


def _is_junk_url(value: str) -> bool:
    if not _is_url(value):
        return False
    lowered = value.lower()
    if any(k in lowered for k in JUNK_URL_KEYWORDS) or _JUNK_X_RE.search(value):
        return True
    host = value.split("//", 1)[-1].split("/")[0].split("?")[0].lower().rstrip(".")
    labels = host.split(".")
    return len(labels) < 2 or not re.match(r"^[a-z]{2,}$", labels[-1])


def _normalize_name(value: str, full_name: bool = False) -> str:
    """'Medina, Cristina Martin' -> 'Cristina Martin Medina'.
    Pour un nom COMPLET uniquement : 'DUPONT Jean' -> 'Jean DUPONT'."""
    s = _clean(value)
    if "," in s:
        last, _, first = s.partition(",")
        last, first = last.strip(), first.strip()
        if last and first and 0 < len(first.split()) <= 3:
            return f"{first} {last}"
    if full_name:
        parts = s.split()
        if (len(parts) >= 2 and len(parts[0]) >= 2 and parts[0].isalpha()
                and parts[0].isupper()
                and not all(p.isupper() for p in parts[1:])):
            return " ".join(parts[1:] + [parts[0]])
        if len(parts) >= 2 and s.isupper():      # 'ALEXANDRE DUPONT' -> 'Alexandre Dupont'
            return s.title()
    return s


def _clean_name(value: str, email: str, full_name: bool = False) -> str:
    s = _normalize_name(value, full_name=full_name)
    if not s:
        return ""
    if _is_email(s) or _is_url(s) or s.lower() == email.lower():
        return ""
    if DATE_LIKE_RE.match(s):
        return ""
    return s


def _clean_job_title(value: str, email: str) -> str:
    s = _clean(value)
    if not s:
        return ""
    if _is_email(s) or _is_url(s):
        return ""
    if s.lower() in {"https", "http", "www", "com"}:
        return ""
    if DATE_LIKE_RE.match(s):
        return ""
    return s


def _clean_company(value: str, email: str) -> str:
    s = _clean(value)
    if not s:
        return ""
    lower = s.lower()
    if _is_email(s) or _is_url(s) or DATE_LIKE_RE.match(s):
        return ""
    forbidden = (
        "mailto:", "bookwithme", "outlook.office", "calendar",
        "video conference", "meetingtype", "anonymous&ep", "from:",
        "de :", "sent:", "envoyé", "https://", "http://", "@",
    )
    if any(f in lower for f in forbidden):
        return ""
    if len(s) > 80:
        return ""
    return s


def _clean_website(value: str, email: str) -> str:
    s = _clean(value)
    if not s:
        return ""
    if "<" in s:                       # Outlook : '<url> <url>'
        s = s.split("<")[0].strip().rstrip(">")
    if not s or not _is_url(s):
        return ""
    if not VALID_DOMAIN_URL_RE.match(s) or _is_junk_url(s):
        return ""
    return s


def _clean_address(value: str, email: str) -> str:
    s = _clean(value)
    if not s or _is_url(s) or _is_email(s) or DATE_LIKE_RE.match(s):
        return ""
    return s


def _clean_phone(value: str) -> str:
    """Garde le zéro initial des numéros nationaux ; 00xx -> +xx."""
    s = _clean(value)
    if not s:
        return ""
    if re.fullmatch(r"\d+\.0", s):
        s = s[:-2]
    digits = re.sub(r"[^\d]", "", s)
    if len(digits) < 6 or len(digits) > 15:
        return ""
    if s.startswith("+"):
        return "+" + digits
    if digits.startswith("00"):
        return "+" + digits[2:]
    return digits


def _contact_source_label(contact: ContactInfo) -> str:
    """Libellé « Source » à partir de _field_sources (meilleure source rencontrée)."""
    sources = getattr(contact, "_field_sources", None)
    if not isinstance(sources, dict) or not sources:
        return SOURCE_LABELS["header"]
    values = {v for v in sources.values() if isinstance(v, str) and v}
    for key in SOURCE_DISPLAY_PRIORITY:
        if key in values:
            return SOURCE_LABELS.get(key, key)
    return SOURCE_LABELS["header"]


def _source_rank(label: str) -> int:
    return _SOURCE_RANK.get(_LABEL_TO_KEY.get(label or "", ""), 0)


def _unique_sheet_name(base: str, seen: set) -> str:
    """Nom de feuille Excel valide et unique (31 caractères max)."""
    sheet = re.sub(r"[\\/*?:\[\]]", "-", base)[:31] or "Famille Client"
    candidate = sheet
    n = 2
    while candidate.lower() in seen:
        suffix = f"_{n}"
        candidate = sheet[: 31 - len(suffix)] + suffix
        n += 1
    seen.add(candidate.lower())
    return candidate


_ISO_FAST_RE = re.compile(r"^(\d{4}-\d{2}-\d{2})[ T](\d{2}:\d{2})(:\d{2})?")


def _norm_date_str(value: str) -> str:
    """Date -> 'AAAA-MM-JJ HH:MM:SS' (chaîne comparable), '' si illisible."""
    value = (value or "").strip()
    if not value:
        return ""
    m = _ISO_FAST_RE.match(value)
    if m:
        return f"{m.group(1)} {m.group(2)}{m.group(3) or ':00'}"
    if re.match(r"^\d{4}-\d{2}-\d{2}$", value):
        return value + " 00:00:00"
    ts = pd.to_datetime(value, errors="coerce")
    return "" if pd.isna(ts) else ts.strftime("%Y-%m-%d %H:%M:%S")


# ----------------------------------------------------------------------
# Exporteur
# ----------------------------------------------------------------------

class DataExporter:
    def __init__(self, contacts: List[ContactInfo], no_project_emails=None):
        self.contacts = contacts
        self.no_project_emails = list(no_project_emails or [])

    # -- construction d'une ligne ---------------------------------------
    def _build_row(self, contact: ContactInfo) -> Optional[dict]:
        email = _clean(getattr(contact, "email", "")).lower()

        row = {
            "Email": email,
            "Nom": _clean_name(getattr(contact, "name", ""), email, full_name=True),
            "Prenom": _clean_name(getattr(contact, "first_name", ""), email),
            "Nom de famille": _clean_name(getattr(contact, "last_name", ""), email),
            "Poste": _clean_job_title(getattr(contact, "job_title", ""), email),
            "Entreprise": _clean_company(getattr(contact, "company", ""), email),
            "Telephone": _clean_phone(getattr(contact, "phone", "")),
            "Famille Client": _clean(getattr(contact, "project", "")),
            "Projet": _clean(getattr(contact, "subfolder", "")),
            "Derniere activite": (
                contact.last_activity.strftime("%Y-%m-%d %H:%M:%S")
                if isinstance(getattr(contact, "last_activity", None), datetime)
                else ""
            ),
            "Dernier sujet": _clean(getattr(contact, "last_subject", "")),
            "Site web": _clean_website(getattr(contact, "website", ""), email),
            "Adresse": _clean_address(getattr(contact, "address", ""), email),
            "Source": _contact_source_label(contact),
        }

        # --- Reconstitution Nom / Prenom / Nom de famille ---
        if not row["Nom"] and (row["Prenom"] or row["Nom de famille"]):
            row["Nom"] = f'{row["Prenom"]} {row["Nom de famille"]}'.strip()
        if row["Nom"] and not (row["Prenom"] and row["Nom de famille"]):
            parts = row["Nom"].split(" ", 1)
            if len(parts) == 2:
                row["Prenom"] = row["Prenom"] or parts[0]
                row["Nom de famille"] = row["Nom de famille"] or parts[1]
            else:
                # Un seul mot (ex. 'Jdupont') : pas de nom de famille inventé
                row["Prenom"] = row["Prenom"] or row["Nom"]

        if not row["Email"] or not _is_email(row["Email"]):
            logger.warning("[EXPORTER] Ligne ignoree (email absent/invalid): %r", row)
            return None

        # --- Audit d'incohérences (log uniquement) ---
        if row["Prenom"] and row["Nom de famille"] and \
                row["Prenom"].lower() == row["Nom de famille"].lower():
            logger.warning(
                "[EXPORTER] INCOHERENCE Prenom == Nom de famille pour %s: %r / %r",
                row["Email"], row["Prenom"], row["Nom de famille"])
        if not row["Entreprise"]:
            logger.info("[EXPORTER] Entreprise absente pour %s (nom: %r)",
                        row["Email"], row["Nom"])
        if row["Nom de famille"] and row["Nom de famille"].lower() in (
                "com", "fr", "net", "org", "www", "http", "https"):
            logger.warning(
                "[EXPORTER] Nom de famille suspect (token web) pour %s: %r",
                row["Email"], row["Nom de famille"])

        return row

    # -- mise en forme --------------------------------------------------
    @staticmethod
    def _protect_formulas(worksheet):
        """Une chaîne commençant par '=' serait écrite comme une formule
        (sujets encodés '=?UTF-8?...') : on force le type texte."""
        for row in worksheet.iter_rows(min_row=2):
            for cell in row:
                v = cell.value
                if isinstance(v, str) and v.startswith("="):
                    cell.data_type = "s"

    @staticmethod
    def _format_sheet(worksheet, frame: pd.DataFrame):
        for i, col in enumerate(frame.columns, start=1):
            if frame[col].empty:
                max_len = len(col)
            else:
                max_len = max(int(frame[col].astype(str).str.len().max()), len(col))
            worksheet.column_dimensions[get_column_letter(i)].width = min(max_len + 2, 50)
        worksheet.freeze_panes = "A2"
        worksheet.auto_filter.ref = worksheet.dimensions
        DataExporter._protect_formulas(worksheet)

    # -- fichier existant -----------------------------------------------
    @staticmethod
    def _backup(output_path: str):
        try:
            shutil.copy2(output_path, output_path + ".bak")
        except Exception as e:
            logger.warning("[EXPORTER] Sauvegarde .bak impossible: %s", e)

    @staticmethod
    def _backup_timestamped(output_path: str) -> Optional[str]:
        """Copie de sécurité PERMANENTE (le .bak simple est écrasé à chaque run)."""
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        path = f"{output_path}.{stamp}.bak"
        try:
            shutil.copy2(output_path, path)
            logger.info("[EXPORTER] Sauvegarde de securite: %s", path)
            return path
        except Exception as e:
            logger.warning("[EXPORTER] Sauvegarde horodatee impossible: %s", e)
            return None

    @staticmethod
    def _load_existing(output_path: str) -> Optional[List[Dict[str, str]]]:
        """Relit la feuille 'Contacts' (tout en texte). None si le fichier a disparu ;
        PermissionError si verrouillé ; ExistingFileError si illisible."""
        try:
            old_df = pd.read_excel(output_path, sheet_name="Contacts",
                                   dtype=str, keep_default_na=False)
        except FileNotFoundError:
            return None
        except PermissionError:
            logger.warning("[EXPORTER] Fichier Excel verrouille (ouvert dans Excel ?)")
            raise
        except Exception as e:
            raise ExistingFileError(
                f"Le fichier existant '{output_path}' est illisible ou n'a pas de "
                f"feuille 'Contacts' ({e}). Il n'a pas ete modifie ; une copie est "
                f"dans '{output_path}.bak'. Deplacez-le ou lancez un rescan complet."
            ) from e

        old_df.columns = [str(c).strip() for c in old_df.columns]
        rows = old_df.to_dict("records")
        for r in rows:
            for col in COLUMNS:
                r.setdefault(col, "")
            r["Derniere activite"] = _norm_date_str(r.get("Derniere activite", ""))
        return rows

    @staticmethod
    def _load_existing_no_project(output_path: str) -> List[dict]:
        """Relit la feuille 'Emails hors projets' ; [] si absente ou illisible."""
        try:
            old_df = pd.read_excel(output_path, sheet_name=NO_PROJECT_SHEET,
                                   dtype=str, keep_default_na=False)
        except Exception:
            return []
        old_df.columns = [str(c).strip() for c in old_df.columns]
        rows = old_df.to_dict("records")
        for r in rows:
            for col in NO_PROJECT_COLUMNS:
                r.setdefault(col, "")
            r["Date"] = _norm_date_str(r.get("Date", ""))
        return rows

    @staticmethod
    def _np_key(row: dict) -> tuple:
        return (str(row.get("Email") or "").strip().lower(),
                str(row.get("Sujet") or "").strip().lower(),
                str(row.get("Date") or "").strip())

    def _merge_no_project_rows(self, old_rows: List[dict],
                               new_rows: List[dict]) -> List[dict]:
        """Anciennes lignes conservees ; pour chaque mail : la ligne la plus recente
        complete les champs vides."""
        result: List[dict] = []
        index: Dict[tuple, int] = {}
        for r in old_rows:
            k = self._np_key(r)
            if k in index:
                continue
            index[k] = len(result)
            result.append({c: str(r.get(c) or "") for c in NO_PROJECT_COLUMNS})
        for r in new_rows:
            k = self._np_key(r)
            row = {c: str(r.get(c) or "") for c in NO_PROJECT_COLUMNS}
            if k in index:
                base = result[index[k]]
                for c in NO_PROJECT_COLUMNS:
                    if row[c]:
                        base[c] = row[c]
            else:
                index[k] = len(result)
                result.append(row)
        return result

    # -- fusion ---------------------------------------------------------
    @staticmethod
    def _row_key(row: dict, fallback: int) -> tuple:
        key = tuple((str(row.get(c) or "")).strip().lower() for c in MERGE_KEY)
        # ligne sans email : jamais fusionnée avec une autre
        return key if key[0] else ("__row__", fallback)

    @staticmethod
    def _merge_pair(old: dict, new: dict) -> dict:
        """Fusionne deux versions de la même ligne : `new` est plus récente."""
        out = dict(old)
        new_wins = _source_rank(new.get("Source", "")) >= _source_rank(old.get("Source", ""))

        # Clés : on garde l'affichage de la ligne récente (restaure la casse)
        for col in MERGE_KEY:
            if new.get(col):
                out[col] = new[col]

        # Nom / Prénom / Nom de famille : bloc cohérent
        if new.get("Nom") and (new_wins or not old.get("Nom")):
            for col in NAME_COLS:
                out[col] = new.get(col, "")

        # Dernière activité + sujet : le mail le plus récent gagne
        old_date, new_date = old.get("Derniere activite", ""), new.get("Derniere activite", "")
        if new_date and new_date >= old_date:
            out["Derniere activite"] = new_date
            out["Dernier sujet"] = new.get("Dernier sujet") or old.get("Dernier sujet", "")
        elif not old_date and new.get("Dernier sujet") and not old.get("Dernier sujet"):
            out["Dernier sujet"] = new["Dernier sujet"]

        # Source : la plus fiable
        if _source_rank(new.get("Source", "")) >= _source_rank(old.get("Source", "")):
            out["Source"] = new.get("Source") or old.get("Source", "")

        # Autres champs : la valeur récente non vide gagne si sa source est
        # au moins aussi fiable ; sinon elle ne fait que combler un vide.
        skip = set(MERGE_KEY) | set(NAME_COLS) | set(ACTIVITY_COLS) | {"Source"}
        for col in COLUMNS:
            if col in skip:
                continue
            nv = str(new.get(col) or "")
            if nv and (new_wins or not old.get(col)):
                out[col] = nv
        return out

    def _merge_rows(self, old_rows: List[dict], new_rows: List[dict]) -> List[dict]:
        """Ordre des anciennes lignes conservé ; doublons (clé normalisée) fusionnés."""
        groups: Dict[tuple, dict] = {}
        for i, r in enumerate(old_rows):
            k = self._row_key(r, i)
            groups[k] = self._merge_pair(groups[k], r) if k in groups else dict(r)
        for i, r in enumerate(new_rows):
            k = self._row_key(r, -1 - i)
            groups[k] = self._merge_pair(groups[k], r) if k in groups else dict(r)
        return list(groups.values())

    # -- export ---------------------------------------------------------
    def export_to_excel(self, output_name, output_dir="output",
                        sheet_per_project=False, incremental=True, keep_backup=False):
        output_path = os.path.join(output_dir, f"{output_name}.xlsx")
        tmp_path = os.path.join(output_dir, f"{output_name}.tmp.xlsx")
        try:
            new_rows = []
            for contact in self.contacts:
                row = self._build_row(contact)
                if row:
                    new_rows.append(row)

            new_np_rows = [dict(r) for r in self.no_project_emails]
            if not new_rows:
                np_empty = not new_np_rows
                if np_empty and os.path.exists(output_path):
                    np_empty = not self._load_existing_no_project(output_path)
                if np_empty:
                    logger.warning("[EXPORTER] Aucun contact valide a exporter")
                    return None

            os.makedirs(output_dir, exist_ok=True)

            final_rows = new_rows
            final_np_rows = list(new_np_rows)
            if os.path.exists(output_path):
                self._backup(output_path)
                if keep_backup:                      # mode refresh : copie horodatée conservée
                    self._backup_timestamped(output_path)
                if incremental:
                    old_rows = self._load_existing(output_path)
                    if old_rows:
                        final_rows = self._merge_rows(old_rows, new_rows)
                        logger.info(
                            "[EXPORTER] Fusion: %d lignes existantes + %d nouvelles "
                            "-> %d totales", len(old_rows), len(new_rows), len(final_rows))
                np_old = self._load_existing_no_project(output_path)
                if np_old or final_np_rows:
                    if incremental:
                        final_np_rows = self._merge_no_project_rows(np_old, new_np_rows)
                    if final_np_rows:
                        logger.info("[EXPORTER] Emails hors projets : %d lignes",
                                    len(final_np_rows))

            columns = list(COLUMNS)
            for r in final_rows:                      # colonnes ajoutées par l'utilisateur
                for c in r:
                    if c not in columns:
                        columns.append(c)
            df = pd.DataFrame(final_rows, columns=columns).fillna("")

            for source, count in Counter(df["Source"]).items():
                if source:
                    logger.info("[EXPORTER] Source '%s': %d contacts", source, int(count))

            with pd.ExcelWriter(tmp_path, engine="openpyxl") as writer:
                df.to_excel(writer, index=False, sheet_name="Contacts")
                self._format_sheet(writer.sheets["Contacts"], df)

                if sheet_per_project:
                    seen = set()
                    for famille, group in df.groupby("Famille Client", sort=False):
                        sheet = _unique_sheet_name(str(famille), seen)
                        group.to_excel(writer, index=False, sheet_name=sheet)
                        self._format_sheet(writer.sheets[sheet], group)

                if final_np_rows:
                    df_np = pd.DataFrame(final_np_rows,
                                         columns=NO_PROJECT_COLUMNS).fillna("")
                    df_np.to_excel(writer, index=False, sheet_name=NO_PROJECT_SHEET)
                    self._format_sheet(writer.sheets[NO_PROJECT_SHEET], df_np)

            os.replace(tmp_path, output_path)   # atomique ; PermissionError si ouvert dans Excel
            return output_path

        except PermissionError:
            logger.error(
                "[EXPORTER] Permission refusee : fermez '%s' dans Excel "
                "avant de relancer l'extraction.", output_path)
            raise
        except ExistingFileError as e:
            logger.error("[EXPORTER] %s", e)
            raise
        except Exception as e:
            logger.exception("[EXPORTER] Erreur lors de l'export: %s", str(e))
            raise
        finally:
            if os.path.exists(tmp_path):
                try:
                    os.remove(tmp_path)
                except Exception:
                    pass
