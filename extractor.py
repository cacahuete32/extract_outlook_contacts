"""Extraction des contacts depuis les mails Outlook (Version optimisée, Thread-Safe, OCR & NLP Intégré).

Optimisations & Corrections :
  * Utilisation prioritaire de signature_parser.py (extract_signature_nlp) avec fallback sur utils.py.
  * Concaténation automatique du texte OCR (inline_images_text / attachments_text) dans le body.
  * Support amélioré des postes composés avec traits d'union (ex: Technico-commercial).
  * Traitement parallélisé via ThreadPoolExecutor et sécurisé par Lock (Thread-Safe).
  * RECHERCHE DE POSTE : _valid_job_title applique utils.is_plausible_job_title (sauf
    annuaire Exchange, de confiance) ; seuls les préfixes composés (Technico-, Co-,
    Vice-...) voient leur tiret normalisé (« Directeur / Gérant » n'est plus modifié).
  * Le texte OCR / des pièces jointes n'est plus collé après le corps pour chercher la
    signature (il devenait les « dernières lignes » analysées) : il ne sert qu'aux
    adresses email ; une signature en IMAGE est analysée à part, avec preuve.
  * _parse_signature_full analyse de nouveau le message ENTIER (hors citations).
  * LECTURE DES PIECES JOINTES ACTIVEE PAR DEFAUT : PDF + OCR (extraction
    parallele, voir outlook.py) et vCards .vcf jointes transformees en contacts
    complets (poste, societe, telephone, site, adresse).
  * Les SMTP sans annuaire Exchange sont resolvus via CreateRecipient +
    GetExchangeUser (carnet d'adresses) pour enrichir nom/poste/societe/telephone.
"""
import re
import logging
import traceback
from typing import List, Optional, Any, Dict, Tuple, Set
from datetime import datetime, timedelta
from concurrent.futures import ThreadPoolExecutor, as_completed
from threading import Lock

from outlook import OutlookFolder
from utils import (
    LEGAL_SUFFIX_RE,
    body_text_from_html,
    is_plausible_job_title,
    is_valid_email,
    normalize_phone,
    extract_signature,
    split_name,
    parse_date,
)

# Tentative d'importation du parseur avancé avec NLP / spaCy
try:
    from signature_parser import extract_signature_nlp
    HAS_NLP_PARSER = True
except ImportError:
    HAS_NLP_PARSER = False
    extract_signature_nlp = None

logger = logging.getLogger(__name__)

# ----------------------------------------------------------------------
# Constantes
# ----------------------------------------------------------------------

EMAIL_RE = re.compile(r"^[\w.+-]+@[\w-]+\.[\w.-]+$")
BODY_EMAIL_RE = re.compile(
    r"[A-Za-z0-9._%+-]+@[A-Za-z0-9-]+(?:\.[A-Za-z0-9-]+)*\.[A-Za-z]{2,}"
)
IMAGE_FAKE_TLDS = {"png", "jpg", "jpeg", "gif", "svg", "webp", "bmp", "ico", "tif", "tiff"}
URL_RE = re.compile(r"^https?://\S+$", re.IGNORECASE)
VALID_DOMAIN_RE = re.compile(
    r"^https?://([A-Za-z0-9-]+\.)+[A-Za-z]{2,}(/\S*)?$", re.IGNORECASE
)
JUNK_URL_KEYWORDS = (
    "googleusercontent", "googlegroups", "outlook.office",
    "bookwithme", "tracking", "signature", "proxy",
    "teams.microsoft", "meet.google", "zoom.us", "webex",
    "calendly.com", "bit.ly", "tinyurl", "linkedin.com",
    "facebook.com", "twitter.com",
)
_JUNK_X_RE = re.compile(r"(?<![\w-])x\.com", re.IGNORECASE)

PR_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E"

SOURCE_PRIORITY = {"header": 0, "signature": 1, "body": 1, "exchange": 2}
WATERMARK_SAFETY_DAYS = 1

QUOTE_CUT_MARKERS = (
    "-----message d'origine-----", "-----original message-----",
    "message d'origine", "original message", "message transfere", "message transféré",
    "forwarded message", "de :", "de:", "from:", "du :", "envoye :", "envoyé :",
    "sent:", "a :", "à :", "to:", "cc :", "cc:", "objet :", "subject:", "date:",
)
QUOTE_ATTRIBUTION_RE = re.compile(
    r"^(le .{5,100} a [ée]crit\s*:|on .{5,100} wrote:)\s*$", re.IGNORECASE
)

QUOTE_HEADER_RE = re.compile(
    r"^(objet|subject|asunto|de|from|envoy\u00e9|sent|date|cc|a|to)\s*[:：]",
    re.IGNORECASE,
)
PHONE_PREFIX_RE = re.compile(
    r"^(tel|t\u00e9l|t\u00e9l\u00e9phone|c|g|m|p|fax|mobile|portable)\s*[:.]\s*[\d+]",
    re.IGNORECASE,
)
COMPANY_SENTENCE_RE = re.compile(
    r"\b(please|lead time|ci-dessous|below|through|colleagues|confirm|"
    r"quotation|shipment|invoice|peux|demande|toujours|"
    r"r[ée]ponse|disponibles?)\b",
    re.IGNORECASE,
)
JOB_SENTENCE_RE = re.compile(
    r"\b(please|quotation|invoice|shipment|colleagues|confirm|lead time|regards)\b",
    re.IGNORECASE,
)
_BAD_COMPANY_START = "*:-\u2022\u25aa>#,.;'\"("
_ABBREV_RE = re.compile(r"\b[A-Za-z]{1,4}\.")

LEGAL_WORDS_RE = re.compile(
    r"\b(s\.?a\.?s?|sarl|sa|ltd|limited|llc|inc|gmbh|bv|nv|plc|group|groupe|"
    r"holding|corp|corporation|company|co|spa|srl|pty|ag|se)\b",
    re.IGNORECASE,
)

_SECOND_LEVEL_TLDS = {"co", "com", "org", "net", "gov", "edu", "ac", "or", "ne"}
FREE_MAIL_TOKENS = {
    "gmail", "googlemail", "hotmail", "outlook", "live", "msn", "yahoo", "ymail",
    "icloud", "me", "aol", "gmx", "protonmail", "proton", "orange", "free", "sfr",
    "wanadoo", "laposte", "neuf", "bbox", "numericable", "mail", "web", "yandex",
}

GENERIC_LOCALS = frozenset({
    "sales", "purchasing", "achats", "achat", "commercial", "billing", "orders",
    "order", "rfq", "info", "contact", "support", "accounting", "compta",
    "comptabilite", "procurement", "quotation", "quotations", "tender", "tenders",
    "service", "services", "hr", "rh", "admin", "office", "logistics",
    "logistique", "invoice", "invoices", "devis", "reception", "accueil", "team",
    "department", "departement", "dept", "supply", "chain", "customer", "sav",
    "after", "marketing", "finance", "direction", "secretariat",
})

SIGNOFF_RE = re.compile(
    r"^(cordialement|bien cordialement|tr[eè]s cordialement|cdt|"
    r"bien [àa] (vous|toi)|salutations|best regards|kind regards|regards|"
    r"best|thanks|thank you|merci)\b",
    re.IGNORECASE,
)
GREETING_RE = re.compile(
    r"^(bonjour|bonsoir|hello|hi|dear|salut|cher|ch[èe]re)\b", re.IGNORECASE
)
SIGNATURE_MAX_LINES = 15


# ----------------------------------------------------------------------
# Fonctions utilitaires de nettoyage
# ----------------------------------------------------------------------

def _strip_quoted_history(body: str) -> str:
    if not body:
        return ""
    kept: List[str] = []
    for line in body.splitlines():
        stripped = line.strip().replace("\xa0", " ")
        if stripped.startswith(">"):
            continue
        lowered = stripped.lower()
        segments = lowered.split("|") if "|" in lowered else [lowered]
        if any(seg.strip().startswith(m) for seg in segments
               for m in QUOTE_CUT_MARKERS):
            break
        if QUOTE_ATTRIBUTION_RE.match(stripped):
            break
        kept.append(line)
    return "\n".join(kept)


def _signature_text(body: str, max_lines: int = SIGNATURE_MAX_LINES) -> str:
    clean = _strip_quoted_history(body)
    lines = [l.strip() for l in clean.splitlines() if re.search(r"\w", l)]
    if not lines:
        return ""
    block = None
    for i, line in enumerate(lines):
        m = SIGNOFF_RE.match(line)
        if m and len(line.split()) <= 4:
            rest = line[m.end():].strip(" |\u2022\u25aa-")
            block = ([rest] if rest and re.search(r"\w", rest) else []) \
                + lines[i + 1:i + 1 + max_lines]
            break
    if block is None:
        block = lines[-max_lines:]
    block = [l for l in block if not GREETING_RE.match(l)]
    return "\n".join(block)


def _clean(value) -> str:
    if value is None:
        return ""
    if isinstance(value, datetime) or not isinstance(value, str):
        return ""
    s = value.strip()
    s = re.sub(r"^mailto:", "", s, flags=re.IGNORECASE).strip()
    s = s.strip("<>")
    s = re.sub(r"\s+", " ", s)
    return s


def _com_text(value) -> str:
    return value.strip() if isinstance(value, str) else ""


def _mail_body(mail) -> str:
    body = _com_text(_safe_attr(mail, "Body"))
    html = _com_text(_safe_attr(mail, "HTMLBody"))
    if html:
        try:
            rebuilt = body_text_from_html(html)
            if rebuilt:
                return rebuilt
        except Exception as e:
            logger.debug(f"[EXTRACTOR] Reconstruction HTML impossible: {e}")
    return body


def _safe_attr(obj, attr, default=None):
    try:
        return getattr(obj, attr)
    except Exception:
        return default


def _is_excluded_addr(addr: str, domains: List[str]) -> bool:
    if not addr or not isinstance(addr, str):
        return True
    addr = addr.lower().strip()
    for d in domains:
        if addr == d or addr.endswith("@" + d) or addr.endswith("." + d):
            return True
    return False


def _is_email_like(s: str) -> bool:
    return bool(s) and bool(EMAIL_RE.match(s))


def _is_url_like(s: str) -> bool:
    return bool(s) and bool(URL_RE.match(s))


def _normalize_person_name(s: str) -> str:
    s = _clean(s)
    if "," in s:
        last, _, first = s.partition(",")
        last, first = last.strip(), first.strip()
        if last and first:
            return f"{first} {last}"
    parts = s.split()
    if (len(parts) >= 2 and len(parts[0]) >= 2 and parts[0].isalpha()
            and parts[0].isupper()
            and not all(p.isupper() for p in parts[1:])):
        return " ".join(parts[1:] + [parts[0]])
    if len(parts) >= 2 and s.isupper():
        return s.title()
    return s


def _valid_name(s: str) -> str:
    s = _clean(_normalize_person_name(s))
    if not s:
        return ""
    if _is_email_like(s) or _is_url_like(s):
        return ""
    return s


# « Technico - commercial » -> « Technico-commercial » (préfixes composés uniquement ;
# les autres séparateurs ' - ' / ' / ' ne sont PAS modifiés)
_COMPOUND_PREFIX_RE = re.compile(r"(?i)\b(technico|co|vice|sous|ex|pr[ée]|auto|inter|multi)\s*[-–—]\s*(?=\w)")


def _valid_job_title(s: str, trusted: bool = False) -> str:
    """Valide un intitulé de poste. trusted=True (annuaire Exchange) : contrôles
    élémentaires uniquement. Sinon l'intitulé doit être un GROUPE NOMINAL plausible,
    jamais une phrase (voir utils.is_plausible_job_title)."""
    s = _clean(s)
    if not s:
        return ""
    s = _COMPOUND_PREFIX_RE.sub(r"\1-", s)
    if _is_email_like(s) or _is_url_like(s):
        return ""
    if trusted:                       # annuaire Exchange : titres parfois longs, jamais filtrés
        return s if len(s) <= 120 else ""
    if len(s) > 65 or len(s.split()) > 8:
        return ""
    if any(x in s for x in ("?", "!", ";")) or "." in _ABBREV_RE.sub("", s):
        return ""
    if re.match(r"^\d+\.", s):
        return ""
    if QUOTE_HEADER_RE.match(s) or PHONE_PREFIX_RE.match(s):
        return ""
    if JOB_SENTENCE_RE.search(s):
        return ""
    if not is_plausible_job_title(s):
        return ""
    return s


def _domain_label(email: str) -> str:
    try:
        parts = [p for p in email.split("@")[1].lower().split(".") if p]
    except Exception:
        return ""
    if len(parts) < 2:
        return parts[0] if parts else ""
    parts = parts[:-1]
    if len(parts) >= 2 and parts[-1] in _SECOND_LEVEL_TLDS:
        parts = parts[:-1]
    return parts[-1]


def _company_domain_token(email: str) -> str:
    return re.sub(r"[^a-z0-9]", "", _domain_label(email))


def _looks_like_project_code(s: str) -> bool:
    s = s.strip()
    if not s:
        return False
    if re.fullmatch(r"[\d.,°%]+", s):
        return True
    compact = re.sub(r"\s+", "", s)
    if re.fullmatch(r"[A-Z]{1,6}[\d]+[A-Z0-9]*", compact):
        return True
    if re.fullmatch(r"[A-Z0-9]+(?:[-_/.][A-Z0-9]+)+", compact):
        return True
    if re.match(r"^[A-Z]{1,4}[-\s]?\d", compact) and \
            len(re.sub(r"[^A-Za-z]", "", compact)) <= 6:
        return True
    if re.fullmatch(r"\d{2,}[A-Z]{1,3}", compact):
        return True
    return False


_PERSON_LIKE_RE = re.compile(r"^[A-ZÀ-ÖØ-Þ][a-zà-ÿ'-]+\s+[A-ZÀ-ÖØ-Þ][a-zà-ÿ'-]+$")
_PERSON_UPPER_RE = re.compile(
    r"^[A-ZÀ-ÖØ-Þ]+(?:[-'][A-ZÀ-ÖØ-Þ]+)?\s+[A-ZÀ-ÖØ-Þ]+(?:[-'][A-ZÀ-ÖØ-Þ]+)?$"
)


def _looks_like_person(s: str) -> bool:
    return bool(_PERSON_LIKE_RE.match(s)) or bool(_PERSON_UPPER_RE.match(s))


def _valid_company(s: str, email: str = "", trusted: bool = False) -> str:
    s = _clean(s)
    if not s:
        return ""

    if trusted:
        if _is_email_like(s) or _is_url_like(s) or len(s) > 80:
            return ""
        return s

    lower = s.lower()

    if s[:1] in _BAD_COMPANY_START:
        return ""
    if s[-1:] in ",:;":
        return ""
    if QUOTE_HEADER_RE.match(s) or PHONE_PREFIX_RE.match(s):
        return ""
    if any(c in s for c in "?…!"):
        return ""
    if COMPANY_SENTENCE_RE.search(lower):
        return ""
    if _looks_like_project_code(s):
        return ""
    if _is_email_like(s) or _is_url_like(s):
        return ""

    bad_tokens = (
        "bookwithme", "outlook.office", "calendar", "meetingtype",
        "video conference", "mailto:", "https://", "http://", "from:", "@",
    )
    if any(token in lower for token in bad_tokens):
        return ""

    if len(s) > 80 or len(s.split()) > 10:
        return ""
    if re.match(r"^\d{5}\s+[A-Za-zÀ-ÿ\- ]+$", s):
        return ""

    if _looks_like_person(s):
        company_token = re.sub(r"[^a-z0-9]", "", lower)
        domain_token = _company_domain_token(email)
        if not (domain_token and company_token
                and (domain_token in company_token or company_token in domain_token)):
            return ""

    if s.isupper() and len(s.split()) >= 3 and not LEGAL_SUFFIX_RE.search(s):
        compact = re.sub(r"[^a-z0-9]", "", lower)
        token = _company_domain_token(email)
        if not (token and len(token) >= 4 and (token in compact or compact in token)):
            return ""
    if any(x in lower for x in (
            "we pitch", "follow us", "connect with", "visit us", "we are hiring")):
        return ""
    if re.search(r"\b(rue|avenue|boulevard|street|road|calle|plaza|allee|chemin)\b", lower):
        return ""
    return s


def _valid_website(s: str) -> str:
    s = _clean(s)
    if not s:
        return ""
    if "<" in s:
        s = _clean(s.split("<")[0].strip().rstrip(">"))
        if not s:
            return ""
    if not VALID_DOMAIN_RE.match(s):
        return ""
    if any(k in s.lower() for k in JUNK_URL_KEYWORDS) or _JUNK_X_RE.search(s):
        return ""
    return s


def _valid_address(s: str) -> str:
    s = _clean(s)
    if not s or _is_url_like(s) or _is_email_like(s):
        return ""
    return s


def _company_from_email_safe(email: str) -> str:
    if not email or not isinstance(email, str) or "@" not in email:
        return ""
    label = _domain_label(email)
    token = re.sub(r"[^a-z0-9]", "", label)
    if not token or token in FREE_MAIL_TOKENS:
        return ""
    candidate = _clean(label)
    if len(candidate) < 3:
        return ""
    return candidate


def _company_key(s: str) -> str:
    low = LEGAL_WORDS_RE.sub(" ", (s or "").lower())
    return re.sub(r"[^a-z0-9]", "", low)


def _same_company(a: str, b: str) -> bool:
    ka, kb = _company_key(a), _company_key(b)
    if not ka or not kb:
        return False
    if ka == kb:
        return True
    shorter, longer = sorted((ka, kb), key=len)
    return len(shorter) >= 4 and shorter in longer


def _person_key(email: str, name: str = "") -> str:
    email = (email or "").lower().strip() if isinstance(email, str) else ""
    if email and "@" in email:
        tokens = [p for p in re.split(r"[^a-z0-9]+", email.split("@")[0]) if len(p) >= 2]
        if len(tokens) >= 2 and not any(t in GENERIC_LOCALS for t in tokens):
            return "".join(tokens)
    norm = _normalize_person_name(name if isinstance(name, str) else "")
    words = [re.sub(r"[^a-z0-9]", "", w) for w in norm.lower().split()]
    words = [w for w in words if w]
    if len(words) >= 2 and not any(w in GENERIC_LOCALS for w in words):
        return "".join(words)
    return ""


def _empty_sender() -> Dict[str, Any]:
    return {"email": "", "name": "", "job_title": "", "company": "",
            "phone": "", "is_exchange": False}


# ----------------------------------------------------------------------
# Modèles de données
# ----------------------------------------------------------------------

class ContactInfo:
    def __init__(self):
        self.email = ""
        self.name = ""
        self.first_name = ""
        self.last_name = ""
        self.job_title = ""
        self.company = ""
        self.phone = ""
        self.project = ""
        self.subfolder = ""
        self.parent = ""
        self.last_activity = None
        self.last_subject = ""
        self.website = ""
        self.address = ""
        self._field_sources: Dict[str, str] = {}
        self.company_history: List[Dict[str, Any]] = []


class EmailData:
    def __init__(self, folder_path, project_name, subfolder, sender_email,
                 sender_name, subject, body, received,
                 sender_job_title="", sender_company="", sender_phone="",
                 sender_is_exchange=False,
                 attachments_text="", inline_images_text="", recipients=None,
                 vcf_contacts=None, mail_id=""):  # vCards jointes -> contacts
        self.mail_id = mail_id
        self.folder_path = folder_path
        self.project_name = project_name
        self.subfolder = subfolder
        self.parent_name = project_name
        self.sender_email = sender_email
        self.sender_name = sender_name
        self.sender_job_title = sender_job_title
        self.sender_company = sender_company
        self.sender_phone = sender_phone
        self.sender_is_exchange = sender_is_exchange
        self.subject = subject
        self.body = body
        self.received = received
        self.attachments_text = attachments_text
        self.inline_images_text = inline_images_text
        self.recipients = recipients or []
        self.vcf_contacts = vcf_contacts or []


# ----------------------------------------------------------------------
# Converter : Mail COM -> EmailData
# ----------------------------------------------------------------------

class MailConverter:
    def __init__(self, outlook=None, include_recipients: bool = True,
                 exclude_domains: Optional[List[str]] = None,
                 read_attachment_text: bool = True,
                 skip_pdf: bool = False, skip_images: bool = False,
                 skip_office: bool = False):
        self.outlook = outlook
        self.namespace = _safe_attr(outlook, "namespace")
        self.include_recipients = include_recipients
        self.exclude_domains = exclude_domains or []
        self.read_attachment_text = read_attachment_text
        self.skip_pdf = skip_pdf
        self.skip_images = skip_images
        self.skip_office = skip_office
        self._sender_cache: Dict[str, Dict[str, Any]] = {}
        self._recipient_cache: Dict[str, Dict[str, Any]] = {}
        self._exchange_cache: Dict[str, Dict[str, Any]] = {}

    def _exchange_details(self, email: str) -> Dict[str, Any]:
        """Resout un SMTP via le carnet d'adresses (GAL / contacts locaux)
        meme quand GetExchangeUser n'est pas disponible directement."""
        empty: Dict[str, Any] = {"name": "", "job_title": "", "company": "",
                                 "phone": "", "is_exchange": False}
        if self.namespace is None or not email:
            return empty
        if email in self._exchange_cache:
            return self._exchange_cache[email]
        try:
            recipient = self.namespace.CreateRecipient(email)
            resolved = bool(_safe_attr(recipient, "Resolve")()) if recipient is not None else False
            if not resolved:
                self._exchange_cache[email] = empty
                return empty
            entry = _safe_attr(recipient, "AddressEntry")
            user = entry.GetExchangeUser() if entry is not None else None
            if user is None:
                self._exchange_cache[email] = empty
                return empty
            info = {"name": _clean(_safe_attr(user, "Name")),
                    "job_title": _clean(_safe_attr(user, "JobTitle")),
                    "company": _clean(_safe_attr(user, "CompanyName")),
                    "phone": (_clean(_safe_attr(user, "BusinessTelephoneNumber"))
                              or _clean(_safe_attr(user, "MobileTelephoneNumber"))),
                    "is_exchange": True}
        except Exception as e:
            logger.debug(f"[EXTRACTOR] Resolution carnet d'adresses impossible ({email}): {e}")
            info = empty
        self._exchange_cache[email] = info
        return info

    def _fill_missing_from_directory(self, info: Dict[str, Any]) -> None:
        """Complete les champs vides par ceux du carnet d'adresses."""
        if info.get("is_exchange"):
            return
        email = info.get("email") or ""
        if "@" not in email:
            return
        details = self._exchange_details(email)
        if not details.get("is_exchange"):
            return
        for f in ("name", "job_title", "company", "phone"):
            if not info.get(f) and details.get(f):
                info[f] = details[f]
        info["is_exchange"] = True

    def _sender(self, mail) -> Dict[str, Any]:
        raw_addr = _clean(_safe_attr(mail, "SenderEmailAddress"))
        raw_name = _clean(_safe_attr(mail, "SenderName"))
        key = (raw_addr or raw_name).lower()
        if key and key in self._sender_cache:
            return self._sender_cache[key]

        info = _empty_sender()
        info["email"], info["name"] = raw_addr, raw_name

        sender_obj = _safe_attr(mail, "Sender")
        if sender_obj is not None:
            exch = None
            try:
                exch = sender_obj.GetExchangeUser()
            except Exception:
                logger.debug("[EXTRACTOR] GetExchangeUser indisponible pour l'expediteur")
            if exch:
                info["email"] = _clean(_safe_attr(exch, "PrimarySmtpAddress")) or info["email"]
                info["name"] = _clean(_safe_attr(exch, "Name")) or info["name"]
                info["job_title"] = _clean(_safe_attr(exch, "JobTitle"))
                info["company"] = _clean(_safe_attr(exch, "CompanyName"))
                info["phone"] = (_clean(_safe_attr(exch, "BusinessTelephoneNumber"))
                                 or _clean(_safe_attr(exch, "MobileTelephoneNumber")))
                info["is_exchange"] = True
            if "@" not in info["email"]:
                try:
                    smtp = sender_obj.PropertyAccessor.GetProperty(PR_SMTP_ADDRESS)
                    info["email"] = _clean(smtp) or info["email"]
                except Exception:
                    pass

        self._fill_missing_from_directory(info)

        if key:
            self._sender_cache[key] = info
        return info

    def _resolve_recipient(self, recipient, addr_raw: str) -> Dict[str, Any]:
        """Resolve un destinataire : SMTP + poste, societe et telephone de
        l'annuaire Exchange quand ils sont disponibles."""
        name = _clean(_safe_attr(recipient, "Name"))
        email = ""
        try:
            email = _clean(recipient.PropertyAccessor.GetProperty(PR_SMTP_ADDRESS))
        except Exception:
            pass
        if not email:
            email = addr_raw
        info: Dict[str, Any] = {"email": (email or "").lower().strip(), "name": name,
                                "job_title": "", "company": "", "phone": "",
                                "is_exchange": False}
        try:
            entry = _safe_attr(recipient, "AddressEntry")
            user = entry.GetExchangeUser() if entry is not None else None
            if user:
                info["email"] = _clean(_safe_attr(user, "PrimarySmtpAddress")) or info["email"]
                info["name"] = _clean(_safe_attr(user, "Name")) or name
                info["job_title"] = _clean(_safe_attr(user, "JobTitle"))
                info["company"] = _clean(_safe_attr(user, "CompanyName"))
                info["phone"] = (_clean(_safe_attr(user, "BusinessTelephoneNumber"))
                                 or _clean(_safe_attr(user, "MobileTelephoneNumber")))
                info["is_exchange"] = True
        except Exception:
            pass
        self._fill_missing_from_directory(info)
        return info

    def _recipients(self, mail) -> List[Dict[str, Any]]:
        out: List[Dict[str, Any]] = []
        seen: Set[str] = set()
        recipients = _safe_attr(mail, "Recipients")
        if recipients is None:
            return out
        try:
            items = list(recipients)
        except Exception:
            return out
        for r in items:
            try:
                addr_raw = _clean(_safe_attr(r, "Address"))
                key = (addr_raw or _clean(_safe_attr(r, "Name"))).lower()
                if key and key in self._recipient_cache:
                    info = self._recipient_cache[key]
                else:
                    info = self._resolve_recipient(r, addr_raw)
                    if key:
                        self._recipient_cache[key] = info
                email = info.get("email", "")
                if (email and email not in seen and is_valid_email(email)
                        and not _is_excluded_addr(email, self.exclude_domains)):
                    seen.add(email)
                    out.append({
                        "email": email, "name": info.get("name", ""),
                        "job_title": info.get("job_title", "") or "",
                        "company": info.get("company", "") or "",
                        "phone": info.get("phone", "") or "",
                        "is_exchange": bool(info.get("is_exchange")),
                    })
            except Exception:
                continue
        return out

    def convert(self, mail, famille_client: str, projet: str, folder_path: str,
                mail_id: Optional[str], received: Optional[datetime]) -> EmailData:
        subject = _clean(_safe_attr(mail, "Subject"))
        body = _mail_body(mail)

        sender = self._sender(mail)
        sender_email = sender["email"].lower().strip()
        if sender_email and _is_excluded_addr(sender_email, self.exclude_domains):
            sender, sender_email = _empty_sender(), ""

        attachments_text = ""
        inline_images_text = ""
        if self.read_attachment_text and self.outlook:
            try:
                attachments_text = self.outlook.get_attachments_text(
                    mail, skip_pdf=self.skip_pdf, skip_images=self.skip_images,
                    skip_office=self.skip_office)
            except TypeError:
                attachments_text = self.outlook.get_attachments_text(mail)
            except Exception as e:
                logger.warning(f"[EXTRACTOR] Erreur pieces jointes: {e}")
            if not self.skip_images:
                try:
                    inline_images_text = self.outlook.get_inline_images_text(mail)
                except Exception:
                    pass

        vcf_contacts: List[Dict[str, str]] = []
        if self.outlook:
            try:
                vcf_contacts = self.outlook.get_vcf_contacts(mail)
            except Exception as e:
                logger.debug(f"[EXTRACTOR] Lecture vCard impossible: {e}")

        return EmailData(
            folder_path=folder_path,
            project_name=famille_client,
            subfolder=projet,
            sender_email=sender_email,
            sender_name=sender["name"],
            subject=subject,
            body=body,
            received=received,
            sender_job_title=sender["job_title"],
            sender_company=sender["company"],
            sender_phone=sender["phone"],
            sender_is_exchange=sender["is_exchange"],
            attachments_text=attachments_text,
            inline_images_text=inline_images_text,
            recipients=self._recipients(mail) if self.include_recipients else [],
            vcf_contacts=vcf_contacts,
            mail_id=mail_id or "",
        )


# ----------------------------------------------------------------------
# Extracteur Principal
# ----------------------------------------------------------------------

class ContactExtractor:
    def __init__(
        self,
        outlook_connector=None,
        process_attachments=True,
        skip_pdf=False,
        skip_images=False,
        skip_office=False,
        date_from=None,
        date_to=None,
        limit=None,
        include_recipients=True,
        exclude_domains=None,
        state=None,
        full_rescan=False,
        process_body_emails=True,
        include_intermediate=False,
        read_attachment_text=True,
        refresh=False,
        max_workers=4,
    ):
        self.outlook = outlook_connector
        self.contacts: List[ContactInfo] = []
        self._index: Dict[Tuple[str, str], ContactInfo] = {}
        self._person_index: Dict[str, List[ContactInfo]] = {}
        self.state = state
        self.full_rescan = full_rescan
        self.refresh = refresh
        self.process_body_emails = process_body_emails
        self.include_intermediate = include_intermediate
        self.max_workers = max_workers
        self.lock = Lock()

        if state is not None and not full_rescan:
            self.processed_mail_ids = state.processed_mail_ids
        else:
            self.processed_mail_ids = set()
        self._seen_ids: Set[str] = set()

        self._pending: List[Tuple[str, str, str]] = []
        self.incomplete_folders: Set[str] = set()
        self.errors: List[str] = []

        self.processed_folder_paths: Set[str] = set()
        self.process_attachments = process_attachments
        self.skip_pdf = skip_pdf
        self.skip_images = skip_images
        self.skip_office = skip_office
        self.date_from = date_from
        self.date_to = date_to
        self.limit = limit
        self.include_recipients = include_recipients
        self.exclude_domains = [d.lower() for d in (exclude_domains or [])]

        self.converter = MailConverter(
            outlook=outlook_connector,
            include_recipients=include_recipients,
            exclude_domains=self.exclude_domains,
            read_attachment_text=bool(read_attachment_text and process_attachments),
            skip_pdf=skip_pdf, skip_images=skip_images, skip_office=skip_office,
        )

    def _already_processed(self, mail_id: Optional[str]) -> bool:
        if not mail_id:
            return False
        if mail_id in self._seen_ids:
            return True
        if self.full_rescan or self.refresh:
            return False
        return mail_id in self.processed_mail_ids

    def _checkpoint(self, mail_id: Optional[str], folder_path: str,
                    received: Optional[datetime]):
        iso = received.strftime("%Y-%m-%d %H:%M:%S") if received else ""
        with self.lock:
            self._pending.append((mail_id or "", folder_path, iso))

    def finalize(self):
        if self.state is None:
            with self.lock:
                self._pending.clear()
            return
        with self.lock:
            global_partial = bool(self.limit) or "*" in self.incomplete_folders
            for mail_id, folder_path, iso in self._pending:
                if global_partial or folder_path in self.incomplete_folders:
                    if mail_id:
                        self.state.processed_mail_ids.add(mail_id)
                else:
                    self.state.mark_processed(mail_id, folder_path, iso)
            self.state.save()
            self._pending.clear()

    def _effective_date_from(self, folder_path: str) -> Optional[datetime]:
        if self.date_from:
            return self.date_from
        if self.state is not None and not (self.full_rescan or self.refresh):
            watermark = self.state.get_watermark(folder_path)
            if watermark:
                try:
                    wm = datetime.strptime(watermark, "%Y-%m-%d %H:%M:%S")
                    return wm - timedelta(days=WATERMARK_SAFETY_DAYS)
                except ValueError:
                    pass
        return None

    @staticmethod
    def _company_trusted(contact: ContactInfo) -> bool:
        return contact._field_sources.get("company") == "exchange"

    def _register_person(self, contact: ContactInfo):
        key = _person_key(contact.email, contact.name)
        if not key:
            return
        siblings = self._person_index.setdefault(key, [])
        if contact not in siblings:
            siblings.append(contact)

    @staticmethod
    def _upsert_history(c: ContactInfo, company: str, date: Optional[datetime]):
        clean = []
        for h in c.company_history:
            if (isinstance(h, dict) and isinstance(h.get("company"), str)
                    and h["company"]):
                clean.append({"company": h["company"], "date": h.get("date")})
        c.company_history = clean
        for h in clean:
            if _same_company(h["company"], company):
                if len(company) > len(h["company"]):
                    h["company"] = company
                if date and (not h.get("date") or date > h["date"]):
                    h["date"] = date
                return
        clean.append({"company": company, "date": date})

    def _append_company_history(self, contact: ContactInfo, company: str,
                                date: Optional[datetime], trusted: bool = False):
        if not isinstance(company, str) or not company:
            return
        contact_email = contact.email if isinstance(contact.email, str) else ""
        company = _valid_company(company, contact_email, trusted=trusted)
        if not company:
            return
        date = date if isinstance(date, datetime) else None

        key = _person_key(contact.email, contact.name)
        siblings = list(self._person_index.get(key, [])) if key else []
        if contact not in siblings:
            siblings.append(contact)
        for c in siblings:
            self._upsert_history(c, company, date)

    def _sync_person_history(self, contact: ContactInfo):
        key = _person_key(contact.email, contact.name)
        if not key:
            return
        siblings = self._person_index.get(key)
        if not siblings or len(siblings) <= 1:
            return
        merged: List[Dict[str, Any]] = []
        for c in siblings:
            for h in c.company_history:
                if not isinstance(h, dict):
                    continue
                comp = h.get("company")
                if not isinstance(comp, str) or not comp:
                    continue
                d = h.get("date")
                for m in merged:
                    if _same_company(m["company"], comp):
                        if len(comp) > len(m["company"]):
                            m["company"] = comp
                        if d and (not m["date"] or d > m["date"]):
                            m["date"] = d
                        break
                else:
                    merged.append({"company": comp, "date": d})
        merged.sort(key=lambda m: (m["date"] is not None, m["date"] or datetime.min))
        for c in siblings:
            c.company_history = [dict(m) for m in merged]

    def _set_field(self, contact: ContactInfo, field: str, value: str, source: str):
        if not value:
            return
        current = getattr(contact, field, "")
        current_source = contact._field_sources.get(field, "header")
        if (not current
                or SOURCE_PRIORITY.get(source, 0) >= SOURCE_PRIORITY.get(current_source, 0)):
            setattr(contact, field, value)
            contact._field_sources[field] = source

    def _merge_contact_data(
        self,
        existing: ContactInfo,
        new_email: str,
        new_name: str,
        new_job_title: str,
        new_phone: str,
        new_company: str,
        new_website: str,
        new_address: str,
        new_last_activity: Optional[datetime],
        new_last_subject: str,
        sources: Optional[Dict[str, str]] = None,
    ) -> ContactInfo:
        sources = sources or {}

        if new_email and is_valid_email(new_email):
            existing.email = new_email.lower().strip()

        self._set_field(existing, "name", _valid_name(new_name),
                        sources.get("name", "header"))
        if existing.name and not (existing.first_name and existing.last_name):
            first, last = split_name(_normalize_person_name(existing.name))
            if first and last:
                existing.first_name, existing.last_name = first, last
            else:
                existing.first_name = existing.first_name or (first or "")
                existing.last_name = existing.last_name or (last or "")

        self._set_field(existing, "job_title", _valid_job_title(
                            new_job_title, trusted=(sources.get("job_title") == "exchange")),
                        sources.get("job_title", "header"))

        company_source = sources.get("company", "header")
        company_clean = (_valid_company(new_company, existing.email,
                                        trusted=(company_source == "exchange"))
                         if isinstance(new_company, str) else "")
        self._set_field(existing, "company", company_clean, company_source)

        self._set_field(existing, "website", _valid_website(new_website),
                        sources.get("website", "signature"))
        self._set_field(existing, "address", _valid_address(new_address),
                        sources.get("address", "signature"))

        if not existing.company:
            existing.company = _company_from_email_safe(existing.email)

        if new_phone:
            new_phone = normalize_phone(new_phone)
            if new_phone:
                if not existing.phone:
                    existing.phone = new_phone
                elif len(re.sub(r"\D", "", new_phone)) > len(re.sub(r"\D", "", existing.phone)):
                    existing.phone = new_phone

        if new_last_activity and (not existing.last_activity
                                  or new_last_activity > existing.last_activity):
            existing.last_activity = new_last_activity
            existing.last_subject = new_last_subject

        activity = new_last_activity or existing.last_activity
        self._register_person(existing)
        if company_clean:
            self._append_company_history(existing, company_clean, activity,
                                         trusted=(company_source == "exchange"))
        self._append_company_history(
            existing, existing.company if isinstance(existing.company, str) else "",
            activity, trusted=self._company_trusted(existing))
        self._sync_person_history(existing)
        return existing

    @staticmethod
    def _guess_name_from_email(contact: ContactInfo):
        local = contact.email.split("@")[0]
        if "." in local and all(len(p) >= 2 for p in local.split(".")):
            guessed = " ".join(local.split(".")).title()
            contact.name = guessed
            contact.first_name, contact.last_name = split_name(guessed)
        elif local and len(local) >= 3:
            contact.name = local.title()
            contact.first_name = local.title()
            contact.last_name = ""

    def _validate_contact_coherence(self, contact: ContactInfo):
        if not contact.email or not is_valid_email(contact.email):
            return

        trusted = self._company_trusted(contact)

        contact.name = _valid_name(contact.name)
        contact.job_title = _valid_job_title(
            contact.job_title, trusted=(contact._field_sources.get("job_title") == "exchange"))
        contact.company = _valid_company(contact.company, contact.email, trusted=trusted)
        contact.website = _valid_website(contact.website)
        contact.address = _valid_address(contact.address)

        if not contact.name:
            self._guess_name_from_email(contact)

        if contact.company and not trusted and _looks_like_person(contact.company):
            compact = re.sub(r"[^a-z0-9]", "", contact.company.lower())
            token = _company_domain_token(contact.email)
            if not (token and compact and (token in compact or compact in token)):
                fallback = _company_from_email_safe(contact.email)
                if fallback:
                    contact.company = fallback

        if not contact.company:
            contact.company = _company_from_email_safe(contact.email)

        self._register_person(contact)
        self._append_company_history(
            contact, contact.company if isinstance(contact.company, str) else "",
            contact.last_activity, trusted=trusted)
        self._sync_person_history(contact)

    def _folders_to_scan(self, folder: OutlookFolder) -> List[OutlookFolder]:
        if not self.outlook:
            return []
        if self.include_intermediate:
            return [folder] + list(self.outlook.get_all_subfolders(folder.folder))
        if self.outlook._is_leaf_folder(folder.folder):
            return [folder]
        return list(self.outlook.get_leaf_folders(folder.folder))

    def extract_from_folder(self, folder: OutlookFolder):
        if folder.path in self.processed_folder_paths:
            return
        self.processed_folder_paths.add(folder.path)
        root_name = folder.root_name

        try:
            targets = self._folders_to_scan(folder)
        except Exception as e:
            self._folder_failed(folder.path, e)
            return

        for target in targets:
            if target.path != folder.path:
                if target.path in self.processed_folder_paths:
                    continue
                self.processed_folder_paths.add(target.path)
            try:
                self._process_folder(target, root_name)
            except Exception as e:
                self._folder_failed(target.path, e)

    def _folder_failed(self, path: str, exc: Exception):
        with self.lock:
            self.incomplete_folders.add(path)
            msg = f"Erreur dossier {path}: {exc}"
            self.errors.append(msg)
        logger.error("[EXTRACTOR] %s\n%s", msg, traceback.format_exc(limit=5))

    def _process_folder(self, folder: OutlookFolder, root_name: str):
        famille_client = root_name
        projet = folder.name
        effective = self._effective_date_from(folder.path)

        if self.outlook and effective:
            try:
                latest = self.outlook.folder_latest_received(folder.folder)
                if latest and latest < effective:
                    logger.info(f"[EXTRACTOR] Dossier ignore (rien de plus recent "
                                f"que {effective.strftime('%Y-%m-%d %H:%M')}): {folder.path}")
                    return
            except Exception as e:
                logger.debug(f"[EXTRACTOR] Verrouillage rapide impossible "
                             f"({folder.path}): {e}")

        if effective:
            logger.info(f"[EXTRACTOR] Traitement increment: {folder.path} "
                        f"depuis le {effective.strftime('%Y-%m-%d %H:%M')}")
        else:
            logger.info(f"[EXTRACTOR] Traitement complet: {folder.path} "
                        f"(Famille Client: {famille_client}, Projet: {projet})")

        email_data_batch: List[EmailData] = []
        mail_count = skipped = 0

        if self.outlook:
            for mail in self.outlook.get_mails(folder.folder, date_from=effective,
                                               date_to=self.date_to, limit=None):
                mail_id = _safe_attr(mail, "EntryID")
                if self._already_processed(mail_id):
                    skipped += 1
                    continue

                received = parse_date(_safe_attr(mail, "ReceivedTime"))
                if received:
                    if effective and received < effective:
                        continue
                    if self.date_to and received > self.date_to:
                        continue

                try:
                    email_data = self.converter.convert(
                        mail, famille_client, projet, folder.path, mail_id, received)
                    email_data_batch.append(email_data)
                    mail_count += 1

                    if self.limit and mail_count >= self.limit:
                        with self.lock:
                            self.incomplete_folders.add(folder.path)
                        break
                except Exception as e:
                    logger.error(f"[EXTRACTOR] Erreur conversion mail: {e}")

        if email_data_batch:
            self._process_email_batch_parallel(email_data_batch)

        logger.info(f"[EXTRACTOR] {mail_count} emails traites "
                    f"({skipped} deja connus) dans {folder.path}")

    def _process_email_batch_parallel(self, email_data_list: List[EmailData]):
        """Traite un lot d'emails en parallèle de manière thread-safe."""
        with ThreadPoolExecutor(max_workers=self.max_workers) as executor:
            futures = [
                executor.submit(self._process_email_data, ed)
                for ed in email_data_list
            ]
            for future in as_completed(futures):
                try:
                    future.result()
                except Exception as e:
                    logger.error(f"[EXTRACTOR] Erreur dans un thread de traitement: {e}")

    def _is_excluded(self, email: str) -> bool:
        return _is_excluded_addr(email, self.exclude_domains)

    def _add_contact(self, email: str, name: str, job_title: str, phone: str,
                     company: str, website: str, address: str, project: str,
                     parent: str, last_activity: Optional[datetime],
                     last_subject: str, subfolder: str = "",
                     sources: Optional[Dict[str, str]] = None) -> bool:
        try:
            if not isinstance(email, str):
                return True
            email = email.lower().strip()
            if not email or self._is_excluded(email):
                return True

            name = _normalize_person_name(name) if isinstance(name, str) else ""
            job_title = job_title if isinstance(job_title, str) else ""
            phone = phone if isinstance(phone, str) else ""
            company = company if isinstance(company, str) else ""
            website = website if isinstance(website, str) else ""
            address = address if isinstance(address, str) else ""
            project = project if isinstance(project, str) else ""
            parent = parent if isinstance(parent, str) else ""
            subfolder = subfolder if isinstance(subfolder, str) else ""
            last_subject = last_subject if isinstance(last_subject, str) else ""
            sources = sources if isinstance(sources, dict) else {}
            last_activity = last_activity if isinstance(last_activity, datetime) else None

            normalized_project = (subfolder or project).lower().strip()

            with self.lock:
                existing = self._index.get((email, normalized_project))

                if existing is not None:
                    self._merge_contact_data(
                        existing, email, name, job_title, phone, company,
                        website, address, last_activity, last_subject, sources=sources)
                    if subfolder:
                        existing.subfolder = subfolder
                    return True

                contact = ContactInfo()
                contact.email = email
                contact._field_sources = dict(sources)
                contact.name = _valid_name(name)
                contact.first_name, contact.last_name = (
                    split_name(contact.name) if contact.name else ("", ""))
                contact.job_title = _valid_job_title(
                    job_title, trusted=(sources.get("job_title") == "exchange"))
                contact.phone = normalize_phone(phone) if phone else ""
                contact.company = _valid_company(
                    company, email, trusted=(sources.get("company") == "exchange"))
                contact.website = _valid_website(website)
                contact.address = _valid_address(address)
                contact.project = project
                contact.subfolder = subfolder
                contact.parent = parent
                contact.last_activity = last_activity
                contact.last_subject = last_subject

                if not contact.company:
                    contact.company = _company_from_email_safe(email)
                if not contact.name:
                    self._guess_name_from_email(contact)

                self._validate_contact_coherence(contact)

                self.contacts.append(contact)
                self._index[(email, normalized_project)] = contact

                self._register_person(contact)
                self._append_company_history(contact, contact.company, contact.last_activity,
                                             trusted=self._company_trusted(contact))
                self._sync_person_history(contact)
            return True

        except Exception as e:
            logger.error(f"[EXTRACTOR] Erreur ajout contact: {e}")
            return False

    def _process_email_data(self, email_data: EmailData) -> bool:
        try:
            famille_client = email_data.project_name
            projet = email_data.subfolder
            received = email_data.received
            subject = email_data.subject
            
            # Signature : corps du mail SEUL. Le texte OCR / des pièces jointes collé à la
            # fin devenait les « dernières lignes » analysées (faux postes tirés d'un PDF).
            # Il ne sert qu'à repérer des adresses email ; une signature en IMAGE (OCR)
            # est analysée à part, avec preuve d'appartenance (voir _extract_sender_from_data).
            body = email_data.body or ""
            body_for_emails = body
            if email_data.inline_images_text:
                body_for_emails += "\n\n--- OCR Inline Image ---\n" + email_data.inline_images_text
            if email_data.attachments_text:
                body_for_emails += "\n\n--- Attachments Text ---\n" + email_data.attachments_text

            ok = True
            if email_data.sender_email:
                ok &= self._extract_sender_from_data(
                    email_data, famille_client, projet, received, subject, body)

            sender_email = (email_data.sender_email or "").lower().strip()
            for vcf in email_data.vcf_contacts:
                vcf_email = (vcf.get("email") or "").lower().strip()
                if not vcf_email or vcf_email == sender_email:
                    continue
                vcf_name = vcf.get("name") \
                    or " ".join(x for x in (vcf.get("first_name"),
                                            vcf.get("last_name")) if x)
                ok &= self._add_contact(
                    email=vcf_email, name=vcf_name,
                    job_title=vcf.get("job_title") or "",
                    phone=vcf.get("phone") or "",
                    company=vcf.get("company") or "",
                    website=vcf.get("website") or "",
                    address=vcf.get("address") or "",
                    project=famille_client, subfolder=projet, parent=famille_client,
                    last_activity=received, last_subject=subject,
                    sources=self._vcf_sources(vcf))

            if self.include_recipients and email_data.recipients:
                for recipient in email_data.recipients:
                    ok &= self._extract_recipient_from_data(
                        recipient, famille_client, projet, received, subject)

            ok &= self._extract_body_emails(body_for_emails, famille_client, projet, received, subject)

            if email_data.mail_id:
                with self.lock:
                    self._seen_ids.add(email_data.mail_id)
            self._checkpoint(email_data.mail_id, email_data.folder_path, received)

            return bool(ok)
        except Exception as e:
            logger.error(f"[EXTRACTOR] Erreur traitement EmailData: {e}")
            return False

    @staticmethod
    def _parse_signature(body: str, email: str = "") -> Optional[dict]:
        """Utilise signature_parser_nlp prioritairement ou fallback sur extract_signature."""
        if not body:
            return None
        try:
            if HAS_NLP_PARSER and extract_signature_nlp is not None:
                parsed = extract_signature_nlp(body, email_hint=email)
                if any(parsed.values()):
                    return parsed
            return extract_signature(_signature_text(body), email_hint=email) or None
        except Exception as e:
            logger.debug(f"[EXTRACTOR] Erreur parse signature: {e}")
            return None

    @staticmethod
    def _parse_signature_full(body: str, email: str = "") -> Optional[dict]:
        """Signature cherchée sur le MESSAGE ENTIER (hors citations), pas seulement sur
        les dernières lignes : signature placée en tête, tableau sans formule de politesse.
        Le résultat n'est utilisé qu'avec une preuve d'appartenance (require_proof)."""
        if not body:
            return None
        try:
            clean = _strip_quoted_history(body)
            if HAS_NLP_PARSER and extract_signature_nlp is not None:
                parsed = extract_signature_nlp(clean, email_hint=email, whole_message=True)
                if any(parsed.values()):
                    return parsed
            return extract_signature(clean, email_hint=email) or None
        except Exception:
            return None

    @staticmethod
    def _signature_belongs_to_sender(sig: Optional[dict], email: str,
                                     sender_name: str, require_proof: bool = False) -> bool:
        if not sig:
            return False
        email = (email or "").strip().lower()
        sig_email = (sig.get("email") or "").strip().lower()
        if sig_email and sig_email == email:
            return True

        def toks(s: str) -> Set[str]:
            return {re.sub(r"[^a-z0-9]", "", t) for t in s.lower().split() if len(t) >= 3} - {""}

        full = sig.get("full_name") or sig.get("name") or ""
        sender = _normalize_person_name(sender_name or "")
        ftok, stok = toks(full), toks(sender)
        if ftok and stok:
            fc = re.sub(r"[^a-z0-9]", "", full.lower())
            sc = re.sub(r"[^a-z0-9]", "", sender.lower())
            return all(t in sc for t in ftok) or all(t in fc for t in stok)

        if sig_email and email:
            s_local, _, s_dom = sig_email.partition("@")
            e_local, _, e_dom = email.partition("@")
            if s_dom != e_dom:
                return False
            s_tok = {t for t in re.split(r"[^a-z0-9]+", s_local) if len(t) >= 3}
            e_tok = {t for t in re.split(r"[^a-z0-9]+", e_local) if len(t) >= 3}
            return bool(s_tok & e_tok) or any(t in e_local for t in s_tok) \
                or any(t in s_local for t in e_tok) or (s_local and s_local in e_local) \
                or (e_local and e_local in s_local)
        return not require_proof

    @staticmethod
    def _vcf_sources(vcf: Dict[str, str]) -> Dict[str, str]:
        sources: Dict[str, str] = {}
        for f in ("job_title", "company", "phone", "website", "address"):
            if vcf.get(f):
                sources[f] = "signature"
        if vcf.get("name"):
            sources["name"] = "signature"
        return sources

    def _extract_sender_from_data(self, email_data: EmailData, famille_client: str,
                                  projet: str, received: Optional[datetime],
                                  subject: str, body: str) -> bool:
        try:
            email = (email_data.sender_email or "").lower().strip()
            if not email or not is_valid_email(email):
                return True

            raw = {
                "name": email_data.sender_name or "",
                "job_title": email_data.sender_job_title or "",
                "phone": email_data.sender_phone or "",
                "company": email_data.sender_company or "",
                "website": "",
                "address": "",
            }
            sources: Dict[str, str] = {
                "name": "exchange" if email_data.sender_is_exchange else "header"}
            for f in ("job_title", "phone", "company"):
                if raw[f]:
                    sources[f] = "exchange"

            matched_vcf: Optional[Dict[str, str]] = None
            for vcf in email_data.vcf_contacts:
                if (vcf.get("email") or "").lower().strip() == email:
                    matched_vcf = vcf
                    break
            if matched_vcf:
                vcf_name = matched_vcf.get("name") \
                    or " ".join(x for x in (matched_vcf.get("first_name"),
                                            matched_vcf.get("last_name")) if x)
                for f, v in (("name", vcf_name),
                             ("job_title", matched_vcf.get("job_title")),
                             ("company", matched_vcf.get("company")),
                             ("phone", matched_vcf.get("phone")),
                             ("website", matched_vcf.get("website")),
                             ("address", matched_vcf.get("address"))):
                    if v and not raw.get(f):
                        raw[f] = v
                        sources[f] = "signature"

            sig = self._parse_signature(body, email)
            if sig and not self._signature_belongs_to_sender(
                    sig, email, email_data.sender_name or ""):
                sig = None

            if sig:
                sig_name = sig.get("full_name") or sig.get("name")
                if sig_name and not raw["name"]:
                    raw["name"] = sig_name
                    sources["name"] = "signature"
                for f in ("job_title", "company", "phone", "website", "address"):
                    if sig.get(f) and not raw[f]:
                        raw[f] = sig[f]
                        sources[f] = "signature"

            if not any(raw[f] for f in ("job_title", "phone", "website", "address")):
                full_sig = self._parse_signature_full(body, email)
                if full_sig and self._signature_belongs_to_sender(
                        full_sig, email, email_data.sender_name or "", require_proof=True):
                    sig_name = full_sig.get("full_name") or full_sig.get("name")
                    if not raw["name"] and sig_name:
                        raw["name"] = sig_name
                        sources["name"] = "signature"
                    for f in ("job_title", "phone", "company", "website", "address"):
                        if full_sig.get(f) and not raw[f]:
                            raw[f] = full_sig[f]
                            sources[f] = "signature"

            # Signature en IMAGE (OCR) : analysée à part, jamais collée au corps
            ocr_text = getattr(email_data, "inline_images_text", "") or ""
            if ocr_text and not any(raw[f] for f in ("job_title", "phone", "website", "address")):
                ocr_sig = self._parse_signature_full(ocr_text, email)
                if ocr_sig and self._signature_belongs_to_sender(
                        ocr_sig, email, email_data.sender_name or "", require_proof=True):
                    for f in ("job_title", "phone", "company", "website", "address"):
                        if ocr_sig.get(f) and not raw[f]:
                            raw[f] = ocr_sig[f]
                            sources[f] = "signature"

            return self._add_contact(
                email=email, name=raw["name"], job_title=raw["job_title"],
                phone=raw["phone"], company=raw["company"],
                website=raw["website"], address=raw["address"],
                project=famille_client, subfolder=projet, parent=famille_client,
                last_activity=received, last_subject=subject, sources=sources)
        except Exception as e:
            logger.error(f"[EXTRACTOR] Erreur expediteur (data): {e}")
            return False

    def _extract_recipient_from_data(self, recipient: Dict, famille_client: str,
                                     projet: str, received: Optional[datetime],
                                     subject: str) -> bool:
        try:
            email = (recipient.get("email") or "").lower().strip()
            if not email or not is_valid_email(email):
                return True
            sources: Dict[str, str] = {"name": "header"}
            if recipient.get("is_exchange"):
                if recipient.get("name"):
                    sources["name"] = "exchange"
                for f in ("job_title", "company", "phone"):
                    if recipient.get(f):
                        sources[f] = "exchange"
            return self._add_contact(
                email=email, name=recipient.get("name") or "",
                job_title=recipient.get("job_title") or "",
                phone=recipient.get("phone") or "",
                company=recipient.get("company") or "",
                website="", address="",
                project=famille_client, subfolder=projet, parent=famille_client,
                last_activity=received, last_subject=subject,
                sources=sources)
        except Exception as e:
            logger.error(f"[EXTRACTOR] Erreur destinataire (data): {e}")
            return False

    def _extract_body_emails(self, body: str, famille_client: str, projet: str,
                             received: Optional[datetime], subject: str) -> bool:
        try:
            if not self.process_body_emails or not body or not isinstance(body, str):
                return True
            clean = _strip_quoted_history(body)
            if not clean:
                return True

            candidates: Set[str] = set()
            for addr in BODY_EMAIL_RE.findall(clean):
                addr = addr.lower().strip(" .,;:<>\"'()[]")
                if addr.count("@") != 1:
                    continue
                local, domain = addr.split("@")
                domain = domain.rstrip(".")
                addr = f"{local}@{domain}"
                if domain.rsplit(".", 1)[-1] in IMAGE_FAKE_TLDS:
                    continue
                if not is_valid_email(addr) or self._is_excluded(addr):
                    continue
                candidates.add(addr)

            ok = True
            for addr in sorted(candidates):
                ok &= self._add_contact(
                    email=addr, name="", job_title="", phone="", company="",
                    website="", address="",
                    project=famille_client, subfolder=projet, parent=famille_client,
                    last_activity=received, last_subject=subject,
                    sources={"name": "body"})
            return bool(ok)
        except Exception as e:
            logger.error(f"[EXTRACTOR] Erreur extraction corps: {e}")
            return False