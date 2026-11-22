"""Fonctions utilitaires pures (emails, noms, téléphones, signatures, dates).

Corrections & Optimisations :
  * JOB_KEYWORDS étendu : postes français accentués et variantes courantes
    ('Technico-commercial', 'Ingénieur', 'Chargé d'Affaires', etc.).
  * split_job_line : apostrophes normalisées et détection des libellés métiers.
  * get_company_from_email : extraction robuste des domaines d'entreprise.
  * is_valid_email : filtrage strict des boîtes génériques et adresses automatiques.
  * split_name et normalize_person_name : gestion fine des prénoms, noms et particules.
  * extract_phone & normalize_phone : support multilingue (FR, US, ES, etc.).
  * body_text_from_html : reconstruction fidèle du texte HTML sans altérer les blocs de signature.
  * _signature_company_candidate : filtrage renforcé via une liste noire d'expressions de junk/slogans/en-têtes.
  * is_plausible_company : une entreprise ressemble à une raison sociale, jamais
    à une phrase du message (« This message and… », « Regards », slogans,
    fragments de bas de page) : pronoms/verbes/politesse FR-EN-ES rejets, sigles
    courts en MAJUSCULES exigent une corrélation avec le domaine email.
  * RECHERCHE DE POSTE (faux postes tirés de phrases du corps du message) :
      - is_plausible_job_title : un intitulé est un GROUPE NOMINAL court ; refuse les
        phrases (verbes, pronoms, politesse, article en tête, ponctuation de fin...).
      - has_signature_context : le poste doit être ENTOURÉ d'éléments de signature
        (nom, téléphone, email, société, site, adresse) dans les 3 lignes voisines, ou
        dans les autres cellules de sa ligne de tableau.
      - Regex invalide corrigée dans extract_signature (« [...’- ] ») : elle levait
        re.error sur toute ligne courte avec virgule et faisait perdre la signature.
      - LEGAL_SUFFIX_RE : liste explicite (plus de « s + n'importe quelle lettre ») ;
        SA / SE / SL / SAU / AG / BV / NV seulement en MAJUSCULES (« sa », « se » sont
        des mots courants). COMPANY_BLACKLIST_TERMS : mots entiers (« por » rejetait
        Corporation, Portugal, Support ; « cif » rejetait Pacific).
      - Intitulés espagnols ajoutés (directora, jefe, gerente, ingeniero...).
"""
import html as _html
import re
from typing import List, Optional, Tuple
from datetime import datetime
import logging

logger = logging.getLogger(__name__)

JOB_KEYWORDS = (
    "chef de projet", "responsable de projet", "directeur de projet",
    "directrice de projet", "chargé de projet", "chargée de projet",
    "chargé d'affaires", "chargée d'affaires", "chargé de mission",
    "chargée de mission", "ingénieur", "ingénieure", "ingénieur d'affaires",
    "ingénieur commercial", "ingénieure commerciale", "ingénieur avant-vente",
    "ingénieur d'études", "ingénieur méthodes", "ingénieur maintenance",
    "technico-commercial", "technico-commerciale", "avant-vente",
    "directeur commercial", "directrice commerciale", "directeur des ventes",
    "directrice des ventes", "directeur technique", "directrice technique",
    "directeur général", "directrice générale", "directeur des opérations",
    "directrice des opérations", "directeur financier", "directrice financière",
    "directeur industriel", "directrice industrielle",
    "directeur des ressources humaines", "directeur des achats",
    "directrice des achats", "directeur de filiale", "directeur de site",
    "directrice de site", "directeur associé", "directrice associée",
    "associé", "associée", "fondateur", "fondatrice",
    "président", "présidente", "président-directeur général",
    "directeur général délégué", "directeur général adjoint",
    "directrice générale adjointe", "gérant", "gérante",
    "co-gérant", "co-gérante",
    "responsable commercial", "responsable commerciale", "responsable technique",
    "responsable avant-vente", "responsable grands comptes",
    "responsable de secteur", "responsable de zone", "responsable d'agence",
    "responsable achats", "responsable logistique", "responsable maintenance",
    "responsable qualité", "responsable production", "responsable exploitation",
    "responsable service", "responsable marketing", "responsable communication",
    "responsable administratif", "responsable administrative",
    "responsable ressources humaines", "responsable développement",
    "responsable projets", "responsable projet", "responsable d'affaires",
    "responsable de business",
    "directeur", "directrice", "responsable", "gérant", "gérante",
    "président", "présidente", "associé", "associée",
    "ingenieur", "ingenieure", "chargé", "chargée", "développeur",
    "développeuse", "developpeur", "projeteur", "projeteuse",
    "technicien", "technicienne", "commercial", "commerciale",
    "acheteur", "acheteuse", "assistant", "assistante",
    "coordinateur", "coordinatrice", "consultant", "consultante",
    "architecte", "juriste", "comptable", "analyste",
    "contrôleur", "contrôleuse", "expert", "experte",
    "ceo", "cto", "cfo", "coo", "cdo", "pdg", "dg", "daf", "drh",
    "superviseur", "superintendent", "team lead", "product owner",
    "scrum master", "project manager", "project director", "program manager",
    "engineer", "sales engineer", "application engineer", "service engineer",
    "field service engineer", "design engineer", "project engineer",
    "process engineer", "mechanical engineer", "electrical engineer",
    "sales manager", "general manager", "country manager", "area manager",
    "regional manager", "district manager", "branch manager",
    "key account manager", "account manager", "business developer",
    "business development manager", "office manager", "hr",
    "proposal manager", "proposal engineer", "bid manager",
    "operations manager", "plant manager", "site manager",
    "quality manager", "qa manager", "maintenance manager",
    "purchasing manager", "procurement manager", "supply chain manager",
    "finance manager", "hr manager", "marketing manager",
    "technical director", "sales director", "managing director",
    "managing partner", "founder", "co-founder", "owner",
    "deputy manager", "department head", "head of",
    "vp", "svp", "evp", "vice president", "vice-president", "vice-président",
    "chief", "officer", "engineering", "specialist", "spécialiste",
    "estimator", "buyer", "consultant", "manager", "director",

    # --- Espagnol / catalan ---
    "directora", "jefe", "jefa", "jefe de proyecto", "jefa de proyecto", "gerente",
    "ingeniero", "ingeniera", "técnico", "técnica", "comprador", "compradora",
    "coordinador", "coordinadora", "administrador", "administradora", "presidente",
    "consultor", "consultora", "comercial", "responsable de compras",
)
JOB_LINE_RE = re.compile(
    r'\b(' + '|'.join(re.escape(k) for k in JOB_KEYWORDS if k) + r')\b',
    re.IGNORECASE,
)
SECOND_LEVEL_TLDS = {"co", "com", "org", "net", "gov", "edu", "ac", "or", "ne"}
IMAGE_FAKE_TLDS = {"png", "jpg", "jpeg", "gif", "svg", "webp", "bmp", "ico", "tif", "tiff"}
AUTOMATIC_PREFIXES = (
    "noreply", "donotreply", "mailerdaemon", "postmaster", "autoreply",
    "bounce", "mailer", "mailcontratacion", "suscripciones", "licitaciones",
    "servicios", "administracion", "administration", "nocontestar", "aviso",
    "notificacion", "notifications", "newsletter", "bulletin",
)
GENERIC_MAILBOXES = {
    "info", "contact", "support", "help", "service", "webmaster",
    "admin", "administrator", "abuse", "return", "error", "fail",
    "sales", "purchasing", "achats", "achat", "commercial", "billing",
    "orders", "order", "rfq", "quotation", "quotations", "tender", "tenders",
    "accounting", "compta", "comptabilite", "procurement", "devis", "invoices",
    "invoice", "logistics", "logistique", "reception", "accueil", "alumni",
    "finance", "marketing", "direction", "secretariat", "customerservice",
}
_EMAIL_PATTERN = re.compile(r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$')
# Les suffixes COURTS ET AMBIGUS (SA, SE, SL, SAU, AG, BV, NV : « sa », « se » sont des
# mots courants) ne comptent qu'en MAJUSCULES : (?-i:...) désactive la casse pour eux.
LEGAL_SUFFIX_RE = re.compile(
    r"\b((?-i:S\.?A\.?)|s\.?a\.?s\.?|s\.?a\.?r\.?l\.?|s\.?e\.?l\.?a\.?r\.?l\.?|s\.?n\.?c\.?|"
    r"s\.?c\.?[ips]\.?|s\.?c\.?a\.?|ltd\.?|limited|inc\.?|llc\.?|gmbh|"
    r"s\.?p\.?a\.?|s\.?r\.?l\.?|pvt\.?|pty\.?|plc\.?|(?-i:BV|N\.?V\.?|AG)|"
    r"group|corp(?:oration)?|company|holding|(?-i:S\.?E\.?)|co\.?\.?|kft|"
    r"(?-i:S\.?L\.?U?\.?)|(?-i:S\.?A\.?U\.?)|sasu|eurl)$",
    re.IGNORECASE,
)
SIGNATURE_COMPANY_JUNK_RE = re.compile(
    r"\b(rfq|rfa|dwg|p&id|warranty|fat|exw|in-house|purchasing|consignee|"
    r"urgent|request|attention|br|ptc|boilers|descrizione|objeto|art)\b|n[°º]",
    re.IGNORECASE,
)

COMPANY_BLACKLIST_TERMS = {
    "suivez", "we are", "all rights reserved", "upcoming events", 
    "this message and", "remorque", "citerne", "aviation", "te confirmo", 
    "objet", "subject", "fyi", "fya", "por", "cif", "avgas", 
    "confidential", "copyright", "sent from", "get outlook", "two turbines and"
}

_COMPANY_BLACKLIST_RE = re.compile(
    r"(?<!\w)(?:" + "|".join(re.escape(t) for t in sorted(COMPANY_BLACKLIST_TERMS, key=len, reverse=True))
    + r")(?!\w)", re.IGNORECASE)

QUOTE_CUT_RE = re.compile(
    r"^(-{3,}|_{3,}|"
    r"(objet|subject|asunto|de|from|envoy[ée]|sent|date|[àa]|to|cc)\s*[:：])",
    re.IGNORECASE,
)
COMMON_TLDS = {
    "com", "fr", "net", "org", "io", "eu", "de", "es", "it", "uk", "be", "ch",
    "nl", "us", "ca", "au", "jp", "cn", "in", "co", "biz", "info", "pt", "se",
    "no", "dk", "fi", "pl", "at", "ie", "cz", "ru", "br", "mx", "ae", "sg", "hk",
    "tr", "za", "kr", "tw", "nz", "ro", "hu", "gr",
}
NAME_PARTICLES = {
    "de", "du", "des", "del", "della", "di", "da", "van", "von", "der", "den",
    "ten", "ter", "la", "le", "el", "al", "bin", "ben", "dos", "das", "do",
    "zu", "af",
}
_LABEL_RE = re.compile(
    r"^(?:(?P<word>e-?mail|mail|courriel|t[ée]l(?:[ée]phone)?|mob(?:ile)?|"
    r"phone|portable|fax|web|site)\s*[:.]|(?P<ch>[etmfw])\s*:)\s*",
    re.IGNORECASE,
)
_ZIP_CITY_RE = re.compile(r"\b\d{4,5}\s+[A-ZÀ-ÖØ-Þ]")
_STREET_RE = re.compile(
    r"\b(rue|avenue|avenida|avinguda|boulevard|bd|street|road|calle|carrer|camino|"
    r"passeig|paseo|plaza|ronda|allee|allée|chemin|edif|n[°º]|p\.?\s?o\.?\s*box)\b",
    re.IGNORECASE,
)

_HTML_TAG_RE = re.compile(r"<[^>]+>")
_HTML_DROP_RE = re.compile(
    r"(?is)<(head|style|script|title|xml)\b.*?</\1\s*>|<!--.*?-->|<!\[[^\]]*\]>")
_HTML_INLINE_RE = re.compile(
    r"(?i)</?(?:span|b|i|u|em|strong|font|a|sup|sub|small|big|o:p|st1:[a-z]+)\b[^>]*>")
_O_PARA_RE = re.compile(r"<o:p>\s*</o:p>", re.IGNORECASE)
_BLOCK_CLOSE_RE = re.compile(r"(?i)</(?:tr|p|div|h[1-6]|li|table)>|<br\s*/?>")
_TD_CLOSE_RE = re.compile(r"(?i)</t[dh]>")
_TD_OPEN_RE = re.compile(r"(?i)<t[dh][^>]*>")


def body_text_from_html(html_body: str) -> str:
    if not html_body or not isinstance(html_body, str):
        return ""
    h = _HTML_DROP_RE.sub("", html_body)
    h = _O_PARA_RE.sub("", h)
    h = _HTML_INLINE_RE.sub("", h)
    h = _TD_CLOSE_RE.sub(" ", h)
    h = _TD_OPEN_RE.sub(" | ", h)
    h = _BLOCK_CLOSE_RE.sub("\n", h)
    h = _HTML_TAG_RE.sub(" ", h)
    h = _html.unescape(h)
    lines: List[str] = []
    for raw in h.split("\n"):
        line = raw.replace("\xa0", " ").replace("\u200b", "")
        line = re.sub(r"[ \t]+", " ", line).strip(" |")
        if line and re.search(r"\w", line):
            lines.append(line)
    return "\n".join(lines)


def is_valid_email(email: str) -> bool:
    if not email or not isinstance(email, str):
        return False
    email = email.strip().lower()
    if len(email) > 254 or not _EMAIL_PATTERN.match(email):
        return False
    local, domain = email.rsplit("@", 1)
    if len(local) > 64 or local.startswith(".") or local.endswith(".") \
            or ".." in local or ".." in domain:
        return False
    if domain.rsplit(".", 1)[-1] in IMAGE_FAKE_TLDS:
        return False
    compact = re.sub(r"[^a-z0-9]", "", local)
    if any(compact.startswith(p) for p in AUTOMATIC_PREFIXES):
        return False
    if compact in GENERIC_MAILBOXES:
        return False
    return True


def get_company_from_email(email: str) -> str:
    if not email or not isinstance(email, str) or '@' not in email:
        return ""
    parts = [p for p in email.split('@')[1].lower().split('.') if p]
    if len(parts) < 2:
        return parts[0] if parts else ""
    parts = parts[:-1]
    if len(parts) >= 2 and parts[-1] in SECOND_LEVEL_TLDS:
        parts = parts[:-1]
    return parts[-1]


def _is_upper_token(t: str, min_len: int = 2) -> bool:
    letters = re.sub(r"[^A-Za-zÀ-ÿ]", "", t)
    return len(letters) >= min_len and t.upper() == t


def normalize_person_name(full_name: str) -> str:
    if not full_name or not isinstance(full_name, str):
        return ""
    s = full_name.strip()
    if "," in s:
        last, _, first = s.partition(",")
        last, first = last.strip(), first.strip()
        if last and first and 0 < len(first.split()) <= 3:
            return f"{first} {last}"
    tokens = s.split()
    j = 0
    while j < len(tokens) and _is_upper_token(tokens[j], 3 if j == 0 else 2):
        j += 1
    if 0 < j < len(tokens):
        return " ".join(tokens[j:] + tokens[:j])
    return s


def split_name(full_name: str) -> Tuple[str, str]:
    if not full_name or not isinstance(full_name, str):
        return "", ""
    s = full_name.strip()
    if "," in s:
        last, _, first = s.partition(",")
        last, first = last.strip(), first.strip()
        if last and first and 0 < len(first.split()) <= 3:
            return first, last
        s = s.replace(",", " ")
    tokens = s.split()
    if not tokens:
        return "", ""
    if len(tokens) == 1:
        return tokens[0], ""
    i = len(tokens)
    while i > 0 and _is_upper_token(tokens[i - 1]):
        i -= 1
    if 0 < i < len(tokens):
        return " ".join(tokens[:i]), " ".join(tokens[i:])
    j = 0
    while j < len(tokens) and _is_upper_token(tokens[j], 3 if j == 0 else 2):
        j += 1
    if 0 < j < len(tokens):
        return " ".join(tokens[j:]), " ".join(tokens[:j])
    for idx in range(1, len(tokens) - 1):
        if tokens[idx].lower() in NAME_PARTICLES:
            return " ".join(tokens[:idx]), " ".join(tokens[idx:])
    return tokens[0], " ".join(tokens[1:])


def normalize_phone(phone: str) -> str:
    if not phone or not isinstance(phone, str):
        return ""
    s = re.sub(r"\(0\)", "", phone.strip())
    digits = re.sub(r"\D", "", s)
    has_plus = "+" in s
    if has_plus:
        cleaned = "+" + digits
    elif digits.startswith("00"):
        cleaned = "+" + digits[2:]
    else:
        cleaned = digits
    n = len(digits)
    if n < 6 or n > 15:
        return ""
    return cleaned


_PHONE_PATTERNS = (
    re.compile(r"\+\d{1,3}(?:[ \t.\-()]*\d{1,4}){2,6}"),
    re.compile(r"(?<![\d+])00\d{1,3}(?:[ \t.\-()]*\d{1,4}){2,6}"),
    re.compile(r"(?<![\d+])0\d(?:[ \t.\-()]?\d{2}){4}(?!\d)"),
    re.compile(r"(?<!\d)\(?\d{3}\)?[ \t.\-]\d{3}[ \t.\-]\d{4}(?!\d)"),
    re.compile(r"(?<!\d)[6-9]\d{2}[ \t.]\d{3}[ \t.]\d{3}(?!\d)"),
)


def extract_phone(text: str) -> str:
    if not text:
        return ""
    for pattern in _PHONE_PATTERNS:
        for m in pattern.finditer(text):
            candidate = m.group(0).strip(" .-()")
            n = len(re.sub(r"\D", "", candidate))
            if 8 <= n <= 15:
                return candidate
    return ""


def parse_date(date_obj) -> Optional[datetime]:
    if date_obj is None:
        return None
    try:
        if isinstance(date_obj, datetime):
            return date_obj.replace(tzinfo=None)
        if hasattr(date_obj, 'year'):
            return datetime(date_obj.year, date_obj.month, date_obj.day,
                            date_obj.hour, date_obj.minute,
                            getattr(date_obj, 'second', 0) or 0).replace(tzinfo=None)
        if isinstance(date_obj, str):
            from dateutil.parser import parse
            return parse(date_obj).replace(tzinfo=None)
    except Exception as e:
        logger.warning(f"[UTILS] Erreur de conversion de date: {str(e)}")
        return None
    return None


_EMOJI_RE = re.compile(
    "[\U0001F000-\U0001FAFF\u2600-\u27BF\u2B00-\u2BFF\uFE00-\uFE0F\u200d\u200b\u2060]")
_MD_LINK_RE = re.compile(r"\[([^\]]*)\]\(([^)]*)\)")
_ANGLE_MAIL_RE = re.compile(r"<mailto:[^>]*>", re.IGNORECASE)
_EMAIL_SEARCH_RE = re.compile(r"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}")
_BARE_DOMAIN_LINE_RE = re.compile(r"^(www\.)?([A-Za-z0-9-]+\.)+([A-Za-z]{2,})/?$")
_JUNK_WEBHOST_RE = re.compile(
    r"(?<![\w-])(teams\.microsoft|meet\.google|zoom\.us|webex\.|calendly\.com|"
    r"outlook\.office|bookwithme|googleusercontent|googlegroups|"
    r"bit\.ly|tinyurl|linkedin\.com|facebook\.com|twitter\.com|x\.com)",
    re.IGNORECASE)
_ADDRESS_LABEL_RE = re.compile(r"^(?:adresse|address|direcci[oó]n)\s*[:=]\s*", re.IGNORECASE)
_JOB_SPLIT_RE = re.compile(r"(\s*[|·—–]\s*|\s+-\s+|\s+l\s+|\s+/\s+)")
_COMPANY_SEP_TYPES = ("|", "·", "—", "–", "-")


def clean_signature_lines(text: str) -> List[str]:
    out: List[str] = []
    for raw in (text or "").split("\n"):
        line = raw.replace("\xa0", " ")
        line = _MD_LINK_RE.sub(r"\1", line)
        line = _ANGLE_MAIL_RE.sub("", line)
        line = _EMOJI_RE.sub(" ", line)
        line = re.sub(r"[ \t]+", " ", line).strip()
        if line and re.search(r"\w", line):
            out.append(line)
    return out


# ----------------------------------------------------------------------
# Intitulés de poste : validation sémantique et contexte de signature
# ----------------------------------------------------------------------
# Mots qui peuvent figurer DANS un intitulé (« Chef de projet », « Head of Sales »)
_TITLE_CONNECTORS = {
    "de", "du", "des", "d", "la", "le", "les", "l", "et", "and", "of", "for", "the",
    "y", "e", "del", "para", "en", "at", "in", "to", "on", "&",
}
# Mots qui n'apparaissent QUE dans une phrase (verbes, pronoms, politesse...)
_PROSE_STRONG = {
    # français
    "merci", "bonjour", "bonsoir", "salut", "cher", "chère", "cordialement", "veuillez",
    "vous", "nous", "je", "j", "tu", "il", "elle", "ils", "elles", "on", "me", "m",
    "mon", "ma", "mes", "ton", "ta", "tes", "notre", "nos", "votre", "vos", "leur", "leurs",
    "est", "sont", "suis", "es", "sera", "serait", "soit", "été", "ai", "avons", "avez",
    "ont", "peux", "peut", "peuvent", "pouvez", "pourriez", "voudrais", "souhaite",
    "souhaitez", "joindre", "contacter", "contactez", "appeler", "appelez", "envoyer",
    "envoyez", "transmettre", "transmets", "transférer", "transférez", "trouver", "voir",
    "voici", "voilà", "ci", "joint", "jointe", "demander", "confirmer", "confirme",
    "confirmons", "oui", "non", "svp", "stp", "asap", "pas", "ne", "n", "très", "aussi",
    "demain", "hier", "aujourd", "hui", "répondra", "reviens", "reviendrai",
    # anglais
    "please", "thanks", "thank", "hello", "hi", "dear", "regards", "i", "you", "we", "he",
    "she", "they", "it", "my", "your", "our", "their", "is", "are", "was", "were", "be",
    "been", "am", "will", "would", "shall", "should", "can", "could", "may", "might",
    "must", "has", "have", "had", "do", "does", "did", "ask", "call", "send", "see",
    "find", "spoke", "speak", "told", "said", "attached", "below", "above", "this",
    "that", "these", "those", "tomorrow", "yesterday", "today", "morning",
    # espagnol
    "hola", "gracias", "estimado", "estimada", "saludos", "favor", "adjunto", "usted",
    "ustedes", "nosotros", "yo", "tú", "él", "ella", "son", "soy", "puedo", "puede",
    "pueden", "hablaré", "hablar", "enviar", "envíe", "llamar", "mañana", "hoy", "ayer",
    "mi", "mis", "nuestro", "nuestra", "su", "sus",
}
# Petits mots de phrase : tolérés isolément, rejetés par deux
_PROSE_WEAK = {
    "un", "une", "ce", "cette", "ces", "au", "aux", "que", "qui", "quand", "si", "ou",
    "a", "à", "pour", "avec", "dans", "sur", "par", "sans", "chez", "with", "from", "by",
    "as", "if", "or", "but", "con", "por", "sobre", "una", "al", "como", "pero",
}
_ARTICLE_START = {"le", "la", "les", "l", "un", "une", "the", "a", "an", "el", "los", "las",
                  "ce", "cette", "ces", "this", "that", "mon", "ma", "mes", "notre", "nos",
                  "votre", "vos", "mi", "mis", "nuestro", "nuestra"}
_NAME_LIKE_RE = re.compile(
    r"^[A-ZÀ-ÖØ-Þ][A-Za-zÀ-ÿ'’-]*(?:\s+(?:(?:de|du|des|del|della|di|da|van|von|der|den|la|le|el|al|dos|das)\s+)*"
    r"[A-ZÀ-ÖØ-Þ][A-Za-zÀ-ÿ'’-]*){1,3}$")


# ------------------------------------------------------------------
# Entreprises : une raison sociale, jamais une phrase
# ------------------------------------------------------------------
# Mots qui n'apparaissent QUE dans une phrase du corps/une signature que l'on
# ne doit jamais prendre pour une entreprise (FR / EN / ES / IT).
_COMPANY_STRONG_PROSE = {
    # pronoms / demonstratifs
    "this", "that", "these", "those", "we", "our", "us", "you", "your", "yours",
    "they", "their", "them", "he", "she", "it", "its", "i", "me", "my",
    "je", "tu", "il", "elle", "ils", "elles", "on", "nous", "vous", "me", "te",
    "yo", "usted", "ustedes", "nosotros", "el", "ella",
    # verbes / auxiliaires
    "are", "was", "were", "is", "am", "be", "been", "being", "have", "has", "had",
    "will", "would", "shall", "should", "can", "could", "may", "might", "must",
    "do", "does", "did", "done", "going", "gonna", "want", "wanted", "hope",
    "est", "sont", "suis", "es", "sera", "serait", "peux", "peut", "peuvent",
    "pouvez", "voudrais", "souhaite", "souhaitez", "va", "vas", "vais", "allez",
    "son", "soy", "somos", "est\u00e1", "puede", "pueden", "hemos",
    "considered", "considers", "confirm", "confirmed", "confirms",
    "review", "reviewed", "correct", "corrected", "removed", "remove",
    "supply", "supplied", "damage", "damaged", "arrange", "arranged", "delete",
    "deleted", "apologise", "apologize", "apologies", "apology", "unfortunately",
    "unfortunatly", "consider", "consid\u00e8re", "considere", "possibilit\u00e9",
    "possibilite", "run", "runs", "risk", "saw", "see", "seen", "envoyer",
    "envoyez", "connait", "conna\u00eet", "member", "connexion",
    # politesse / mail
    "yes", "no", "not", "nor", "please", "thanks", "thank", "merci", "bonjour",
    "bonsoir", "bonne", "bon", "hello", "hi", "hola", "regards", "saludos",
    "salutations", "gracias", "muchas", "favor", "adjunto", "cordialement",
    "svp", "stp", "asap", "suivez", "enviado", "oggetto", "objet", "subject",
    "re:", "fw:", "fwd:", "notification", "param\u00e8tres", "parametres",
    "wetransfer", "whatsapp", "linkedin", "skype", "caution", "office",
    "mr", "mrs", "ms", "monsieur", "madame", "mme", "mlle", "test", "cde",
    "direction", "d\u00e9partement", "departement", "department",
    "administration", "administraci\u00f3n", "administracion",
    "reception", "recepci\u00f3n", "bureau", "syst\u00e8me", "board", "valve",
    "sales", "representative", "repr\u00e9sentant", "ligne", "mes", "france",
    "service", "communiqu\u00e9", "communique",
    # nombres crits
    "two", "three", "four", "five", "six", "deux", "trois", "quatre",
    "dos", "tres",
    # mois / jours
    "january", "february", "march", "april", "june", "july", "august",
    "september", "october", "november", "december",
    "janvier", "fevrier", "f\u00e9vrier", "mars", "avril", "mai", "juin",
    "juillet", "aout", "ao\u00fbt", "septembre", "octobre", "novembre", "decembre",
    "d\u00e9cembre", "lundi", "mardi", "mercredi", "jeudi", "vendredi", "samedi",
    "dimanche", "monday", "tuesday", "wednesday", "thursday", "friday",
    "saturday", "sunday",
}
# Petits mots utilitaires : tol\u00e9rs seul, rejets par deux ou en fin de nom
_COMPANY_WEAK_PROSE = {
    "the", "in", "on", "for", "with", "from", "about", "into", "over", "between",
    "of", "at", "to", "by", "as", "or", "and", "but", "if", "then", "than",
    "when", "where", "while", "le", "la", "les", "des", "du", "de", "un", "une",
    "au", "aux", "dans", "sur", "pour", "par", "avec", "sans", "chez", "en",
    "del", "al", "los", "las", "con", "por", "para", "per", "della", "delle",
    "und", "der", "die", "das", "van", "von",
}
# Phrases / fragments vus dans les signatures qui ne sont jamais des soci\u00e9t\u00e9s
_COMPANY_JUNK_PHRASES = (
    "this message", "all rights", "upcoming events", "two turbines", "we are",
    "follow us", "connect with", "visit us", "we are hiring", "we pitch",
    "bureau d", "dir co", "bon we", "prior written", "electronic", "entity named",
    "may be", "hereby", "confidential", "get outlook", "sent from",
    "j'ai", "j\u2019ai", "j\u2019", "qu'", "qu\u2019", "n'", "n\u2019",
)
_COMPANY_JUNK_PHRASE_RE = re.compile(
    r"(?<!\w)(?:" + "|".join(re.escape(t) for t in _COMPANY_JUNK_PHRASES) + r")(?!\w)",
    re.IGNORECASE)
_COMPANY_FILE_EXT_RE = re.compile(r"\.(docx?|xlsx?|pptx?|pdf|png|jpe?g|vcf|zip)$",
                                  re.IGNORECASE)
_COMPANY_UPPER_COMMON = {"system", "systems", "valve"}


def is_plausible_company(name: str) -> bool:
    """Une entreprise est une raison sociale ('SPIG S.p.A.', 'Kelvion'),
    jamais une phrase du message, un slogan, une formule de politesse,
    un fragment de footer ou un sigle court hors domaine."""
    s0 = (name or "").strip()
    if not s0:
        return False
    s = s0.strip(" \t.,;:!?+-\u2013\u2014|")
    if not s:
        return False
    if len(s) < 2 or len(s) > 80:
        return False
    if s[0] in "'\u2019\"" or s[-1] in "'\u2019\"":
        return False
    if "->" in s or "/" in s or "@" in s or _COMPANY_FILE_EXT_RE.search(s):
        return False
    if _COMPANY_JUNK_PHRASE_RE.search(s):
        return False
    if re.search(r"https?://|www\.", s, re.IGNORECASE):
        return False
    tokens = [t for t in re.split(r"[^0-9a-z\u00e0-\u00ff]+", s.lower()) if t]
    if not tokens:
        return False
    if len(tokens) == 1 and tokens[0] in _COMPANY_UPPER_COMMON:
        return False
    if tokens[0] in _COMPANY_STRONG_PROSE or tokens[0] in _COMPANY_WEAK_PROSE:
        return False
    if tokens[-1] in _COMPANY_STRONG_PROSE or tokens[-1] in _COMPANY_WEAK_PROSE:
        return False
    if any(t in _COMPANY_STRONG_PROSE for t in tokens):
        return False
    if sum(1 for t in tokens if t in _COMPANY_WEAK_PROSE) >= 2:
        return False
    if len(tokens) == 1 and tokens[0].isalpha() and 2 <= len(tokens[0]) <= 4:
        if any(c.isupper() for c in s):
            return False      # PT / PE / FAT / KKS / ExW / Mit... : pas sans le domaine
    if re.search(r"\d", s) and len(s) <= 4 and not s.isupper():
        return False          # x1 / v2... : code, pas une societe
    if any(t.isdigit() for t in tokens[-1:]):
        return False          # adresse postale ("Markkaweg 1"), "Option 1"
    return True


def is_plausible_job_title(title: str) -> bool:
    """Un intitulé de poste est un GROUPE NOMINAL court (« Chef de projet »,
    « Head of Sales »), jamais une phrase. Rejette : verbes, pronoms, politesse,
    phrases commençant par un article, ponctuation de fin de phrase, ':' '@' '?' '!' ';',
    plus de 7 mots, chiffres de téléphone."""
    t = (title or "").strip()
    if not (2 <= len(t) <= 60) or len(t.split()) > 7:
        return False
    if any(ch in t for ch in "?!;:@") or re.search(r"https?://|www\.", t, re.IGNORECASE):
        return False
    if re.search(r"[.,]$", t) or re.search(r"\d{3,}", t):
        return False
    tokens = [w for w in re.split(r"[^a-zà-ÿ&]+", t.lower().replace("’", "'")) if w]
    if not tokens:
        return False
    if tokens[0] in _ARTICLE_START:
        return False
    if any(w in _PROSE_STRONG for w in tokens):
        return False
    if sum(1 for w in tokens if w in _PROSE_WEAK) >= 2:
        return False
    return True


def is_name_like(line: str) -> bool:
    """'Sami Mighri', 'Alexandre DUPONT', 'Ana da Silva' : 2 à 4 mots capitalisés,
    sans chiffre ni mot de poste."""
    s = (line or "").strip()
    if not (2 <= len(s.split()) <= 5) or re.search(r"[0-9@:/|]", s):
        return False
    if JOB_LINE_RE.search(s):
        return False
    return bool(_NAME_LIKE_RE.match(s))


def _is_signature_anchor(line: str) -> bool:
    """La ligne ressemble-t-elle à un élément de signature (nom, téléphone, email,
    site, société à suffixe légal, adresse, libellé « Tél : ») ?"""
    s = (line or "").strip()
    if not s:
        return False
    if is_name_like(s) or extract_phone(s) or _EMAIL_SEARCH_RE.search(s):
        return True
    if re.match(r"^(https?://|www\.)", s, re.IGNORECASE):
        return True
    bare = _BARE_DOMAIN_LINE_RE.match(s)
    if bare and (bare.group(1) or bare.group(3).lower() in COMMON_TLDS):
        return True
    if (len(s.split()) >= 2 and LEGAL_SUFFIX_RE.search(s.strip(" ,;:"))
            and not JOB_LINE_RE.search(s)):
        return True
    if _strip_label(s)[0] is not None:
        return True
    return _looks_like_address_line(s)


def has_signature_context(lines: List[str], idx: int, window: int = 3) -> bool:
    """Un intitulé de poste n'est retenu que s'il est ENTOURÉ d'éléments de signature
    (dans les `window` lignes voisines) : c'est ce qui sépare une vraie signature d'une
    phrase du corps du message qui contient un mot comme « responsable » ou « manager »."""
    own = [seg.strip() for seg in lines[idx].split("|")]
    if len(own) > 1 and any(_is_signature_anchor(seg) and not JOB_LINE_RE.search(seg) for seg in own):
        return True                                   # ligne de tableau : 'Nom | Poste | +33 ...'
    lo, hi = max(0, idx - window), min(len(lines), idx + window + 1)
    return any(j != idx and _is_signature_anchor(lines[j]) for j in range(lo, hi))


def split_job_line(line: str) -> Tuple[str, str]:
    line = (line or "").replace("’", "'").replace("`", "'")
    if not JOB_LINE_RE.search(line):
        return "", ""
    if (not _JOB_SPLIT_RE.search(line) and len(line.split()) >= 2
            and LEGAL_SUFFIX_RE.search(line.strip(" ,;:"))):
        return "", ""
    parts = _JOB_SPLIT_RE.split(line)
    for i in range(0, len(parts), 2):
        seg = parts[i].strip(" ,;:-–—")
        if not JOB_LINE_RE.search(seg):
            continue
        if not (2 <= len(seg) <= 60) or len(seg.split()) > 8:
            return "", ""
        if "@" in seg or any(c in seg for c in "?!;") or seg.endswith("."):
            return "", ""
        if re.search(r"\+?\d[\d\s.-]{7,}", seg):
            return "", ""
        if not is_plausible_job_title(seg):
            continue
        hint = ""
        neighbours = []
        if i + 2 < len(parts):
            neighbours.append((parts[i + 1].strip(), parts[i + 2]))
        if i - 2 >= 0:
            neighbours.append((parts[i - 1].strip(), parts[i - 2]))
        for sep, cand in neighbours:
            cand = cand.strip(" ,;:-–—")
            if (sep in _COMPANY_SEP_TYPES and 2 <= len(cand) <= 40
                    and len(cand.split()) <= 4 and not JOB_LINE_RE.search(cand)
                    and "@" not in cand):
                hint = cand
                break
        return seg, hint
    return "", ""


def _detect_job_title(line: str) -> str:
    return split_job_line(line)[0]


def _signature_company_candidate(line: str) -> bool:
    s = line.strip().strip(",;:")
    if not s or len(s) > 80:
        return False
    lower_s = s.lower()
    if _COMPANY_BLACKLIST_RE.search(lower_s):
        return False
    if not is_plausible_company(s):
        return False
    tokens = s.split()
    if JOB_LINE_RE.search(s):
        if _JOB_SPLIT_RE.search(s) or not (len(tokens) >= 2 and LEGAL_SUFFIX_RE.search(s)):
            return False
    if SIGNATURE_COMPANY_JUNK_RE.search(s):
        return False
    if len(tokens) >= 2 and LEGAL_SUFFIX_RE.search(s):
        return True
    if any(c.isdigit() for c in s):
        return False
    if re.search(r"\s[-–—|/&]\s|[|:]", s):
        return False
    return s.isupper() and 2 <= len(tokens) <= 4


def _strip_label(line: str) -> Tuple[Optional[str], str]:
    m = _LABEL_RE.match(line)
    if not m:
        return None, line
    word = (m.group("word") or "").lower()
    ch = (m.group("ch") or "").lower()
    if word:
        if word.startswith(("e-mail", "email", "mail", "courriel")):
            key = "mail"
        elif word.startswith(("t", "mob", "phone", "portable")):
            key = "tel"
        elif word == "fax":
            key = "fax"
        else:
            key = "web"
    else:
        key = {"e": "mail", "t": "tel", "m": "tel", "f": "fax", "w": "web"}[ch]
    return key, line[m.end():].strip()


def _is_other_field(line: str) -> bool:
    if _strip_label(line)[0] is not None:
        return True
    if _EMAIL_SEARCH_RE.search(line) or re.match(r'^(https?://|www\.)', line, re.IGNORECASE):
        return True
    return bool(extract_phone(line))


def _looks_like_address_line(line: str) -> bool:
    if "©" in line or _is_other_field(line):
        return False
    return bool(_ZIP_CITY_RE.search(line)
                or (_STREET_RE.search(line) and re.search(r"\d", line)))


def extract_address(lines: List[str]) -> str:
    for i, raw in enumerate(lines):
        line = _ADDRESS_LABEL_RE.sub("", raw).strip()
        if not _looks_like_address_line(line):
            continue
        parts = [line]
        complete = bool(_ZIP_CITY_RE.search(line))
        zip_seen = False
        j = i + 1
        while not complete and j < len(lines) and len(parts) < 5:
            nxt = lines[j].strip()
            if _is_other_field(nxt):
                break
            if re.fullmatch(r"\d{4,6}", nxt) or _ZIP_CITY_RE.search(nxt):
                parts.append(nxt)
                zip_seen = True
                complete = bool(_ZIP_CITY_RE.search(nxt))
            elif (zip_seen and len(nxt.split()) <= 3
                    and re.fullmatch(r"[A-Za-zÀ-ÿ'’. -]+", nxt)
                    and not JOB_LINE_RE.search(nxt)):
                parts.append(nxt)
            else:
                break
            j += 1
        return ", ".join(parts)
    return ""


def find_company_line_from_domain(lines: List[str], email: str, exclude=()) -> str:
    label = re.sub(r"[^a-z0-9]", "", get_company_from_email(email or "").lower())
    if len(label) < 4:
        return ""
    ex = {e.lower() for e in exclude if e}
    for line in lines:
        cand = line.strip().strip(",;:")
        if not (2 <= len(cand) <= 60) or cand.lower() in ex:
            continue
        if len(cand.split()) > 5 or any(c in cand for c in "@|:/") \
                or _BARE_DOMAIN_LINE_RE.match(cand):
            continue
        if JOB_LINE_RE.search(cand) and not LEGAL_SUFFIX_RE.search(cand):
            continue
        if re.sub(r"[^a-z0-9]", "", cand.lower()).startswith(label):
            return cand
    return ""


def extract_signature(text: str, email_hint: str = "") -> dict:
    result = {
        "full_name": "", "first_name": "", "last_name": "",
        "job_title": "", "company": "", "phone": "",
        "email": "", "website": "", "address": ""
    }
    if not text:
        return result

    kept_lines = []
    for line in text.split('\n'):
        if QUOTE_CUT_RE.match(line.strip()):
            break
        kept_lines.append(line)
    lines = clean_signature_lines("\n".join(kept_lines))

    job_hint = ""
    units: List[Tuple[Optional[str], str]] = []
    for li, raw_line in enumerate(lines):
        label, line = _strip_label(raw_line)
        if label == "fax" or not line:
            continue
        segs = [s.strip() for s in line.split("|")] if "|" in line else [line]
        segs = [s for s in segs if s]
        if len(segs) > 1 and not result["job_title"]:
            title, hint = split_job_line(line)
            # contexte : un AUTRE segment de la ligne ou une ligne voisine est un élément de signature
            if title and (any(_is_signature_anchor(sg) and not JOB_LINE_RE.search(sg) for sg in segs)
                          or has_signature_context(lines, li)):
                result["job_title"] = title
                if hint and not result["company"]:
                    result["company"] = hint
                job_hint = hint
        if len(segs) == 1:
            units.append((label, segs[0]))
        else:
            for seg in segs:
                seg_label, rest = _strip_label(seg)
                if seg_label == "fax" or not rest:
                    continue
                units.append((seg_label, rest))

    flat_units = [u[1] for u in units]
    for ui, (label, line) in enumerate(units):
        if _EMAIL_PATTERN.match(line):
            if not result["email"]:
                result["email"] = line
            continue
        if label == "mail":
            m = _EMAIL_SEARCH_RE.search(line)
            if m and not result["email"]:
                result["email"] = m.group(0)
            continue

        phone = extract_phone(line)
        if phone:
            if not result["phone"]:
                result["phone"] = phone
            continue

        if "<" in line and "://" in line:
            line = line.split("<")[0].strip().rstrip(">")
            if not line:
                continue

        if re.match(r'^https?://', line, re.IGNORECASE):
            if not result["website"] and not _JUNK_WEBHOST_RE.search(line):
                result["website"] = line
            continue
        bare = _BARE_DOMAIN_LINE_RE.match(line)
        if bare and (bare.group(1) or bare.group(3).lower() in COMMON_TLDS):
            if not result["website"] and not _JUNK_WEBHOST_RE.search(line):
                result["website"] = "http://" + line.rstrip("/")
            continue

        if not result["job_title"]:
            title, hint = split_job_line(line)
            if title and has_signature_context(flat_units, ui):
                result["job_title"] = title
                if hint and not result["company"]:
                    result["company"] = hint
                job_hint = hint
                continue

        is_company = _signature_company_candidate(line)
        has_legal = bool(LEGAL_SUFFIX_RE.search(line.strip().strip(",;:")))

        if not result["company"] and is_company and has_legal:
            result["company"] = line.strip().strip(",;:")
            continue

        if not result["full_name"]:
            name_candidate = ""
            if ("," in line
                    and len(line.split()) <= 4
                    and not re.search(r'[0-9@]', line)
                    and re.match(
                        r"^[A-Za-zÀ-ÿ'’-]+,\s*[A-Za-zÀ-ÿ'’ -]+$", line)):
                name_candidate = normalize_person_name(line)
            elif (2 <= len(line.split()) <= 3
                    and not re.search(r'[0-9@]', line)
                    and re.match(
                        r"^[A-ZÀ-ÖØ-Þ][A-Za-zÀ-ÿ'’-]*"
                        r"(\s+[A-ZÀ-ÖØ-Þ][A-Za-zÀ-ÿ'’-]*)*$", line)):
                name_candidate = line

            if name_candidate:
                if name_candidate.isupper():
                    name_candidate = name_candidate.title()
                result["full_name"] = name_candidate
                result["first_name"], result["last_name"] = split_name(name_candidate)
                continue

        if not result["company"] and is_company:
            result["company"] = line.strip().strip(",;:")
            continue

    email = result["email"] or email_hint
    domain_line = find_company_line_from_domain(
        flat_units, email, exclude=(result["full_name"], result["job_title"]))
    company = result["company"]
    if domain_line:
        if not company:
            company = domain_line
        elif company.isupper():
            key = lambda x: re.sub(r"[^a-z0-9]", "", x.lower())
            if key(company).startswith(key(domain_line)):
                company = domain_line
    result["company"] = company or job_hint

    result["address"] = extract_address(flat_units)
    return result
