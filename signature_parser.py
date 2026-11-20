"""Analyse de signatures d'emails (regex d'abord, spaCy en secours).

Optimisations & Corrections :
  * Chargement du modèle spaCy à la demande (Lazy Loading) avec désactivation des pipes inutiles.
  * Gestion simultanée des étiquettes de personnes FR (« PER ») et EN (« PERSON »).
  * Reconnaissance étendue des postes (support des traits d'union et mots composés).
  * Extraction robuste des numéros de téléphone, des sites web et des adresses postales.
  * Intégration transparente pour un appel direct via extract_signature_nlp(body).
  * RECHERCHE DE POSTE : validation sémantique et contexte de signature (voir
    utils.is_plausible_job_title / has_signature_context) ; la boucle « de secours »
    trop laxiste est soumise aux mêmes règles ; parmi plusieurs candidats, celui qui
    suit immédiatement un nom est préféré.
  * email_hint / whole_message : le domaine de l'expéditeur retrouve le nom de la
    société telle qu'écrite ; whole_message=True analyse tout le texte (repli).
"""
import logging
import re
from functools import lru_cache
from typing import Dict, List, Optional, Tuple

from utils import (
    LEGAL_SUFFIX_RE,
    clean_signature_lines,
    extract_address,
    extract_phone,
    find_company_line_from_domain,
    has_signature_context,
    is_name_like,
    is_plausible_job_title,
    get_company_from_email,
    is_valid_email,
    normalize_phone,
    split_job_line,
    split_name,
)

try:
    import spacy
except ImportError:
    spacy = None

try:
    import phonenumbers
    from phonenumbers import PhoneNumberFormat, PhoneNumberMatcher
except ImportError:
    phonenumbers = None
    PhoneNumberFormat = None
    PhoneNumberMatcher = None

logger = logging.getLogger(__name__)

# ----------------------------------------------------------------------
# Constantes
# ----------------------------------------------------------------------

DEFAULT_SPACY_MODELS = ("fr_core_news_sm", "en_core_web_sm")
PERSON_LABELS = {"PER", "PERSON"}
SIGNATURE_MAX_LINES = 15

QUOTE_CUT_MARKERS = (
    "-----message d'origine-----", "-----original message-----",
    "de :", "from:", "envoyé :", "envoye :", "sent:", "à :", "a :", "to:",
    "cc :", "objet :", "subject:", "date:",
)
QUOTE_ATTRIBUTION_RE = re.compile(
    r"^(le .{5,100} a [ée]crit\s*:|on .{5,100} wrote:)\s*$", re.IGNORECASE)

SIGNOFF_RE = re.compile(
    r"^(cordialement|bien cordialement|tr[eè]s cordialement|cdt|bien [àa] (vous|toi)|"
    r"sinc[eè]res salutations|salutations|respectueusement|best regards|kind regards|"
    r"regards|best|thanks|thank you|merci|saludos|atentamente|un saludo)\b",
    re.IGNORECASE)
GREETING_RE = re.compile(
    r"^(bonjour|bonsoir|hello|hi|dear|salut|cher|ch[èe]re|hola|buenos d[ií]as|estimad[oa])\b", re.IGNORECASE)
DEVICE_FOOTER_RE = re.compile(
    r"^(sent from my|envoy[ée] (depuis|de mon)|get outlook for|obtenez outlook|enviado desde)",
    re.IGNORECASE)

JOB_TITLES = [
    "CEO", "PDG", "Directeur", "Director", "Directrice", "Directora", "Chef", "Responsable",
    "Manager", "Gérant", "Gerente", "Ingénieur", "Ingeniero", "Ingeniera", "Engineer",
    "Consultant", "Adjoint", "Assistant", "Chargé", "Chargée", "Coordinateur",
    "Coordinatrice", "Technicien", "Técnico", "Technico-commercial", "Technico-commerciale",
    "Expert", "Spécialiste", "Analyste", "Développeur", "DevOps", "Architecte",
    "Chef de projet", "Project Manager", "Commercial", "Comercial", "Vendeur",
    "Marketing", "RH", "DRH", "Comptable", "Financier", "Juriste", "Avocat",
    "Président", "Vice-Président", "CTO", "CFO", "COO", "CIO", "CDO", "Jefe", "Jefa",
    "Comprador", "Buyer", "Head of", "Purchasing", "Procurement",
]
_JOB_UPPER = sorted({t for t in JOB_TITLES if t.isupper()}, key=len, reverse=True)
_JOB_WORDS = sorted({t for t in JOB_TITLES if not t.isupper()}, key=len, reverse=True)
_JOB_UPPER_RE = re.compile(r"\b(?:" + "|".join(map(re.escape, _JOB_UPPER)) + r")\b")
_JOB_WORD_RE = re.compile(r"\b(?:" + "|".join(map(re.escape, _JOB_WORDS)) + r")\b",
                          re.IGNORECASE)
_JOB_LABEL_RE = re.compile(
    r"^\s*(?:poste|fonction|job\s*title|title|cargo)\s*[:=]\s*(.+)", re.IGNORECASE | re.MULTILINE)
_JOB_SEPARATORS = ("|", "·", " — ", " – ", " - ")

_COMPANY_LABEL_RE = re.compile(
    r"^\s*(?:entreprise|company|société|societe|empresa)\s*[:=]\s*(.+)",
    re.IGNORECASE | re.MULTILINE)
JUNK_COMPANY_TOKENS = (
    "bookwithme", "outlook.office", "calendar", "video conference", "meetingtype",
    "anonymous&ep", "mailto:", "http://", "https://", "www.", "from:", "de :",
    "sent:", "envoyé", "subject:", "objet :", "click here",
)
FREE_MAIL_TOKENS = {
    "gmail", "googlemail", "hotmail", "outlook", "live", "msn", "yahoo", "ymail",
    "icloud", "me", "aol", "gmx", "protonmail", "proton", "orange", "free", "sfr",
    "wanadoo", "laposte", "neuf", "bbox", "numericable", "mail", "web", "yandex",
}

# --- Noms ---
_NAME_TOKEN = (r"(?:[A-ZÀ-ÖØ-Þ][a-zà-ÿ'']+(?:-[A-ZÀ-ÖØ-Þ][a-zà-ÿ'']+)*"
               r"|[A-ZÀ-ÖØ-Þ]{2,}(?:[-''][A-ZÀ-ÖØ-Þ]{2,})*|[A-ZÀ-ÖØ-Þ]\.)")
_NAME_PARTICLE = r"(?:de|du|des|del|della|di|da|van|von|der|den|la|le|el|al|dos|das)"
NAME_LINE_RE = re.compile(
    rf"^{_NAME_TOKEN}(?:\s+(?:{_NAME_PARTICLE}\s+)*{_NAME_TOKEN}){{1,3}}$")
_TITLES = {"mr", "mrs", "ms", "mme", "mlle", "m", "dr", "prof", "madame",
           "monsieur", "sr", "sra", "don", "doña"}

# --- Téléphone ---
_FAX_PREFIX_RE = re.compile(r"(?:fax|\bf)\s*[:.]?\s*$")
_FAX_CUT_RE = re.compile(r"(?im)(?:\bfax\b|\bf\s*:).*$")

# --- Web ---
_URL_RE = re.compile(r"(?:https?://|www\.)[^\s<>'\"]+\b", re.IGNORECASE)
_BARE_DOMAIN_RE = re.compile(
    r"(?<![\w@.-])(?:[A-Za-z0-9][A-Za-z0-9-]*\.)+"
    r"(?:com|es|fr|mx|net|org|eu|de|it|pt|uk|nl|be|ch|pl|info|biz|co|edu|gov|cat|io|us|ca)"
    r"(?![\w@-])", re.IGNORECASE)
JUNK_URL_KEYWORDS = ("googleusercontent", "googlegroups", "outlook.office",
                     "bookwithme", "tracking", "click", "proxy")
SOCIAL_HOSTS = ("linkedin.com", "facebook.com", "twitter.com", "x.com", "instagram.com",
                "youtube.com", "wa.me", "calendly.com")
_EMAIL_RE = re.compile(r"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}")

# --- Adresse ---
_STREET_RE = re.compile(
    r"\b(?:rue|avenue|avenida|boulevard|bd|street|allée|allee|chemin|calle|camino|plaza|avda|paseo|road|carrer|passeig|avinguda|ronda)\b", re.IGNORECASE)
_ZIP_CITY_RE = re.compile(r"\b\d{4,5}\s+[A-ZÀ-ÖØ-Þ][A-Za-zÀ-ÿ'\- ]+")
_ADDRESS_LABEL_RE = re.compile(r"^(?:adresse|address|direcci[oó]n)\s*[:=]\s*",
                               re.IGNORECASE)

# ----------------------------------------------------------------------
# Utilitaires de découpage et NLP
# ----------------------------------------------------------------------

def _strip_quoted_history(text: str) -> str:
    if not text:
        return ""
    kept: List[str] = []
    for line in text.splitlines():
        stripped = line.strip()
        if stripped.startswith(">"):
            continue
        if any(stripped.lower().startswith(m) for m in QUOTE_CUT_MARKERS):
            break
        if QUOTE_ATTRIBUTION_RE.match(stripped):
            break
        kept.append(line)
    return "\n".join(kept)

def _has_job_keyword(s: str) -> bool:
    return bool(_JOB_UPPER_RE.search(s) or _JOB_WORD_RE.search(s))

def _host(url: str) -> str:
    host = re.sub(r"^[a-z]+://", "", url, flags=re.IGNORECASE)
    return host.split("/")[0].split("?")[0].rstrip(".").lower()

def _is_plausible_url(url: str) -> bool:
    try:
        labels = _host(url).split(".")
        if len(labels) < 2 or any(len(l) < 2 for l in labels):
            return False
        return bool(re.match(r"^[a-z]{2,}$", labels[-1]))
    except Exception:
        return False

@lru_cache(maxsize=None)
def _load_nlp(models: Tuple[str, ...]):
    if spacy is None:
        logger.warning("[SIGNATURE_PARSER] spaCy non installe : regex uniquement.")
        return None
    for name in models:
        try:
            nlp = spacy.load(name)
        except OSError:
            continue
        for pipe in ("parser", "lemmatizer", "morphologizer", "attribute_ruler",
                     "tagger", "senter"):
            if pipe in nlp.pipe_names:
                try:
                    nlp.disable_pipe(pipe)
                except Exception:
                    pass
        logger.info(f"[SIGNATURE_PARSER] Modele spaCy charge: {name}")
        return nlp
    logger.warning("[SIGNATURE_PARSER] Aucun modele spaCy disponible : regex uniquement.")
    return None

class _LazyDoc:
    def __init__(self, nlp, text: str):
        self._nlp, self._text = nlp, text
        self._doc = None
        self._failed = False

    def get(self):
        if self._doc is None and self._nlp is not None and not self._failed:
            try:
                self._doc = self._nlp(self._text)
            except Exception as e:
                self._failed = True
                logger.debug(f"[SIGNATURE_PARSER] spaCy a echoue: {e}")
        return self._doc

# ----------------------------------------------------------------------
# Classe SignatureParser
# ----------------------------------------------------------------------

class SignatureParser:
    def __init__(self, default_region: str = "FR", use_nlp: bool = True,
                 spacy_models: Tuple[str, ...] = DEFAULT_SPACY_MODELS):
        self.default_region = default_region
        self.nlp = _load_nlp(tuple(spacy_models)) if use_nlp else None

    @staticmethod
    def _empty_result() -> Dict[str, Optional[str]]:
        return {
            "full_name": None, "first_name": None, "last_name": None,
            "job_title": None, "company": None, "phone": None,
            "email": None, "website": None, "address": None,
        }

    def extract_signature(self, text: str) -> Optional[str]:
        if not text:
            return None
        text = text.replace("\xa0", " ").replace("\u200b", "").replace("\r\n", "\n")
        lines = [l for l in _strip_quoted_history(text).split("\n") if l.strip()]

        for i, line in enumerate(lines):
            if DEVICE_FOOTER_RE.match(line.strip()):
                lines = lines[:i]
                break

        start = None
        for i, line in enumerate(lines):
            s = line.strip()
            if len(s) >= 2 and set(s) <= set("-_= "):
                start = i + 1
                break
            if len(s.split()) <= 4 and SIGNOFF_RE.match(s):
                start = i + 1
                break

        if start is None:
            block = clean_signature_lines("\n".join(lines))[-SIGNATURE_MAX_LINES:]
        else:
            block = clean_signature_lines("\n".join(lines[start:]))[:SIGNATURE_MAX_LINES]
        block = [l for l in block if not GREETING_RE.match(l.strip())]
        return "\n".join(block).strip() or None

    def parse_signature(self, signature: str, email_hint: str = "") -> Dict[str, Optional[str]]:
        if not signature:
            return self._empty_result()
        signature = self._clean_signature(signature)
        doc = _LazyDoc(self.nlp, signature)

        email = self._extract_email(signature)
        company_email = email or (email_hint or "").strip().lower() or None
        job_title, job_hint = self._extract_job_title_and_hint(signature)
        full_name = self._extract_name(signature, doc, exclude=(job_title,))
        company = self._extract_company(signature, doc, exclude=(full_name, job_title),
                                        email=company_email, hint=job_hint)
        result = {
            "full_name": full_name, "first_name": None, "last_name": None,
            "job_title": job_title, "company": company,
            "phone": self._extract_phone(signature), "email": email,
            "website": self._extract_website(signature),
            "address": self._extract_address(signature),
        }
        if full_name:
            result["first_name"], result["last_name"] = self._split_name(full_name)
        return result

    def parse_body(self, body: str, email_hint: str = "") -> Dict[str, Optional[str]]:
        block = self.extract_signature(body)
        return self.parse_signature(block, email_hint) if block else self._empty_result()

    @staticmethod
    def _clean_signature(text: str) -> str:
        return "\n".join(clean_signature_lines(text.replace("\u200b", "").replace("\r\n", "\n")))

    @staticmethod
    def _lines(text: str) -> List[str]:
        return [l.strip() for l in text.splitlines() if l.strip()]

    @staticmethod
    def _is_plausible_name(s: str) -> bool:
        s = s.strip()
        if not (4 <= len(s) <= 50):
            return False
        if any(ch.isdigit() for ch in s) or "@" in s or "://" in s:
            return False
        if not NAME_LINE_RE.match(s):
            return False
        if SIGNOFF_RE.match(s) or GREETING_RE.match(s):
            return False
        if _has_job_keyword(s) or LEGAL_SUFFIX_RE.search(s):
            return False
        return True

    def _extract_name_without_nlp(self, text: str, exclude=()) -> Optional[str]:
        ex = {e.lower() for e in exclude if e}
        for line in self._lines(text):
            low = line.lower()
            if "@" in line or low.startswith(("http", "www", "de :", "from:", ">")):
                continue
            if low in ex:
                continue
            if self._is_plausible_name(line):
                return line.title() if line.isupper() else line
        return None

    def _extract_name(self, text: str, doc=None, exclude=()) -> Optional[str]:
        name = self._extract_name_without_nlp(text, exclude)
        if name:
            return name
        if doc is None:
            doc = _LazyDoc(self.nlp, text)
        d = doc.get()
        if d is None:
            return None
        ex = {e.lower() for e in exclude if e}
        for ent in d.ents:
            if ent.label_ not in PERSON_LABELS:
                continue
            for segment in ent.text.split("\n"):
                cand = segment.strip()
                if cand and cand.lower() not in ex and self._is_plausible_name(cand):
                    return cand
        return None

    def _split_name(self, full_name: str) -> Tuple[str, str]:
        name = (full_name or "").strip()
        if not name or "@" in name or name.lower().startswith(("http", "www")):
            return ("", "")
        parts = name.split()
        while parts and parts[0].lower().rstrip(".") in _TITLES:
            parts = parts[1:]
        if not parts:
            return ("", "")
        name = " ".join(parts)
        if len(parts) >= 2 and LEGAL_SUFFIX_RE.search(name):
            return ("", name)
        return split_name(name)

    def _extract_job_title(self, text: str) -> Optional[str]:
        return self._extract_job_title_and_hint(text)[0]

    def _extract_job_title_and_hint(self, text: str) -> Tuple[Optional[str], Optional[str]]:
        """(intitulé complet, société probable de la même ligne) ;
        « Global Area Sales Manager - AX GROUP » -> ('Global Area Sales Manager', 'AX GROUP').

        Un intitulé doit (1) être un groupe nominal plausible (pas une phrase) et
        (2) être ENTOURÉ d'éléments de signature (nom, téléphone, email, société...).
        Parmi plusieurs candidats, celui qui suit immédiatement un nom est préféré."""
        lines = self._lines(text)

        m = _JOB_LABEL_RE.search(text)                  # « Poste : Directeur »
        if m:
            candidate = m.group(1).strip().strip(".,;")
            if 2 <= len(candidate) <= 60 and is_plausible_job_title(candidate):
                return candidate, None

        candidates = []                                   # (rang, position, titre, société)
        for i, line in enumerate(lines):
            if "@" in line or "://" in line or line.lower().startswith("www"):
                continue
            title, hint = split_job_line(line)            # liste de mots-clés de utils (FR/EN/ES)
            if not title:                                 # liste étendue locale (Gerente, Jefe...)
                loose = line.strip(" ,;:-–—")
                if (2 <= len(loose) <= 60 and _has_job_keyword(loose)
                        and is_plausible_job_title(loose)
                        and not re.search(r"\d{6,}", re.sub(r"[\s.\-]", "", loose))):
                    title, hint = loose, None
            if not title or not has_signature_context(lines, i):
                continue
            after_name = i > 0 and is_name_like(lines[i - 1])
            candidates.append((0 if after_name else 1, i, title, hint or None))
        if not candidates:
            return None, None
        _, _, title, hint = min(candidates)
        return title, hint

    @staticmethod
    def _looks_like_address(line: str) -> bool:
        return bool((_STREET_RE.search(line) and re.search(r"\d", line))
                    or _ZIP_CITY_RE.search(line))

    def _looks_like_junk_company(self, text: str) -> bool:
        if not text:
            return True
        lower = text.lower().strip()
        if any(token in lower for token in JUNK_COMPANY_TOKENS):
            return True
        if "@" in text or len(text) > 80 or len(text.split()) > 10:
            return True
        if re.fullmatch(r"[\d\s.,+()\-]+", text):
            return True
        return self._looks_like_address(text)

    def _extract_company(self, text: str, doc=None, exclude=(),
                         email: Optional[str] = None, hint: Optional[str] = None) -> Optional[str]:
        ex = {e.lower() for e in exclude if e}
        lines = self._lines(text)

        m = _COMPANY_LABEL_RE.search(text)
        if m:
            candidate = m.group(1).strip().strip(".,;")
            if 2 <= len(candidate) <= 60 and not self._looks_like_junk_company(candidate):
                return candidate

        email = email or self._extract_email(text)
        legal = None
        for line in lines:
            for seg in re.split(r"\s*[|·—–]\s*|\s+-\s+|\s+l\s+", line):
                cand = seg.strip().strip(",;:")
                if (3 <= len(cand) <= 60 and len(cand.split()) >= 2
                        and LEGAL_SUFFIX_RE.search(cand)
                        and cand.lower() not in ex
                        and not _JOB_WORD_RE.search(cand) and not _JOB_UPPER_RE.search(cand)
                        and not self._looks_like_junk_company(cand)):
                    legal = cand
                    break
            if legal:
                break

        domain_line = find_company_line_from_domain(lines, email or "", exclude=exclude)
        if legal:
            if domain_line and legal.isupper():
                key = lambda x: re.sub(r"[^a-z0-9]", "", x.lower())
                if key(legal).startswith(key(domain_line)):
                    return domain_line
            return legal

        if domain_line and not self._looks_like_junk_company(domain_line):
            return domain_line

        if hint and not self._looks_like_junk_company(hint):
            return hint

        if doc is None:
            doc = _LazyDoc(self.nlp, text)
        d = doc.get()
        if d is not None:
            for ent in d.ents:
                if ent.label_ != "ORG":
                    continue
                for segment in ent.text.split("\n"):
                    cand = segment.strip()
                    if (2 <= len(cand) <= 60
                            and cand.lower() not in ex
                            and not self._looks_like_junk_company(cand)
                            and not _has_job_keyword(cand)):
                        return cand

        if email:
            company = get_company_from_email(email)
            token = re.sub(r"[^a-z0-9]", "", company.lower()) if company else ""
            if (company and token not in FREE_MAIL_TOKENS and 2 <= len(company) <= 60
                    and not self._looks_like_junk_company(company)):
                return company
        return None

    def _extract_phone(self, text: str) -> Optional[str]:
        if phonenumbers is not None:
            try:
                for m in PhoneNumberMatcher(text, self.default_region):
                    line_start = text.rfind("\n", 0, m.start) + 1
                    prefix = text[line_start:m.start].lower()
                    if _FAX_PREFIX_RE.search(prefix):
                        continue
                    if phonenumbers.is_valid_number(m.number):
                        return phonenumbers.format_number(m.number, PhoneNumberFormat.E164)
            except Exception as e:
                logger.debug(f"[SIGNATURE_PARSER] phonenumbers a echoue: {e}")
        for line in self._lines(text):
            phone = normalize_phone(extract_phone(_FAX_CUT_RE.sub("", line)))
            if phone:
                return phone
        return None

    @staticmethod
    def _extract_email(text: str) -> Optional[str]:
        for match in _EMAIL_RE.finditer(text):
            email = match.group(0).lower().strip(".,;:")
            if is_valid_email(email):
                return email
        return None

    @staticmethod
    def _junk_url(url: str) -> bool:
        if any(k in url.lower() for k in JUNK_URL_KEYWORDS):
            return True
        if not _is_plausible_url(url):
            return True
        host = _host(url)
        return any(host == s or host.endswith("." + s) for s in SOCIAL_HOSTS)

    def _extract_website(self, text: str) -> Optional[str]:
        for match in _URL_RE.finditer(text):
            url = match.group(0).strip().rstrip(".,;)")
            if self._junk_url(url):
                continue
            return url if url.lower().startswith(("http://", "https://")) else "https://" + url
        for match in _BARE_DOMAIN_RE.finditer(text):
            url = match.group(0)
            if not self._junk_url(url):
                return "https://" + url
        return None

    def _extract_address(self, text: str) -> Optional[str]:
        return extract_address(self._lines(text)) or None


_default_parser: Optional[SignatureParser] = None


def extract_signature_nlp(body: str, default_region: str = "FR",
                          email_hint: str = "", whole_message: bool = False) -> Dict[str, str]:
    global _default_parser
    if _default_parser is None or _default_parser.default_region != default_region:
        _default_parser = SignatureParser(default_region=default_region)
    if whole_message:       # tout le texte est analysé (pas seulement le bloc final)
        parsed = _default_parser.parse_signature(body, email_hint)
    else:
        parsed = _default_parser.parse_body(body, email_hint)
    return {k: (v or "") for k, v in parsed.items()}