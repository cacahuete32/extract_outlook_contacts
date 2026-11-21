"""Revue de compte : changements d'entreprise, entreprises, statistiques.

Corrections par rapport à la version précédente :
  * Les helpers d'identité (_person_key, _same_company, _normalize_person_name)
    sont importés de extractor.py : une seule définition, plus de divergence.
    Conséquence : sales@a.com / sales@b.com (ou achats@, rfq@...) ne sont plus
    pris pour la même personne, et « Kelvion » / « Kelvion Thermal Solutions
    SAS » ne sont plus déclarés comme un changement d'entreprise (la première
    passe comparait encore avec lower() ==).
  * La revue n'est plus limitée au run courant. Elle s'appuie sur :
      - un historique persistant (<fichier>.history.json, écriture atomique)
        qui garde ce qu'on a vu à chaque exécution : un changement entre deux
        runs est enfin détecté ;
      - la feuille « Contacts » cumulée du fichier Excel (statistiques et
        entreprises sur l'ensemble, plus sur les seuls nouveaux mails).
    Si le fichier ou l'historique est illisible, la revue retombe sur le run.
  * « Personnes uniques » : les contacts sans clé d'identité comptent chacun
    pour une personne (avant, ils étaient tous regroupés sous la clé '').
  * Entreprises regroupées par variantes de nom (suffixes SAS/Ltd/Group...).
  * Écriture Excel atomique (copie temporaire puis os.replace) : un plantage
    ne corrompt plus le classeur ; chaînes commençant par « = » protégées.
  * Les dates sont toutes naïves (plus de bricolage UTC/aware).
  * TOUTE entreprise observee (historique persiste, run courant, Excel cumule)
    est revalidee par extractor._valid_company : les phrases, sigles et
    fragments stockes par les anciens runs disparaissent de la revue au
    prochain passage sans toucher manuellement au .history.json.
  * COHERENCE EMAIL : une observation doit recouper le domaine de l'adresse
    (TORAY sur fgoffaux@prindustries.fr = ecarte) ; deux entreprises du meme
    employeur sur la meme adresse (spig-gmab <-> SPIG S.p.A.) ne comptent
    plus comme un changement.

API inchangée : AccountReview(contacts), .append_to_excel(path), .gui_data().
Appeler append_to_excel() AVANT gui_data() pour obtenir la vue cumulée.
"""
import json
import logging
import os
import re
import shutil
from datetime import datetime
from typing import Any, Dict, List, Optional

import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

from extractor import (
    ContactInfo,
    _company_matches_email,
    _normalize_person_name,
    _person_key,
    _same_company,
    _valid_company,
)

logger = logging.getLogger(__name__)

CONTROL_CHARS_RE = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]")
HISTORY_VERSION = 1


# ----------------------------------------------------------------------
# Helpers
# ----------------------------------------------------------------------

def _s(value: Any) -> str:
    """Chaîne propre ; un datetime dans un champ texte est ignoré."""
    if value is None or isinstance(value, datetime):
        return ""
    if isinstance(value, str):
        s = value
    else:
        try:
            s = str(value)
        except Exception:
            return ""
    return CONTROL_CHARS_RE.sub("", s).strip()


def _naive(dt: Any) -> Optional[datetime]:
    """datetime naïf (accepte les pywintypes.datetime, sous-classe de datetime)."""
    if isinstance(dt, datetime):
        return dt.replace(tzinfo=None)
    return None


def _parse_dt(value: Any) -> Optional[datetime]:
    if isinstance(value, datetime):
        return _naive(value)
    s = _s(value)
    if not s:
        return None
    try:
        return datetime.fromisoformat(s[:19].replace("T", " "))
    except ValueError:
        pass
    try:
        from dateutil.parser import parse
        return _naive(parse(s))
    except Exception:
        return None


def _fmt_date(value: Any) -> str:
    dt = _parse_dt(value)
    return dt.strftime("%Y-%m-%d") if dt else ""


def _dt_key(dt: Optional[datetime]) -> datetime:
    return dt or datetime.min


def _person_id(email: str, name: str = "") -> str:
    """Clé de personne ; à défaut (boîte générique, nom inconnu) l'adresse
    elle-même : son historique propre est analysé, jamais fusionné avec un autre."""
    email = _s(email).lower()
    key = _person_key(email, name)
    if key:
        return key
    return f"email:{email}" if email else ""


# ----------------------------------------------------------------------
# Revue
# ----------------------------------------------------------------------

class AccountReview:
    """Génère une revue de compte à partir des contacts extraits."""

    def __init__(self, contacts: List[ContactInfo], persist: bool = True):
        self.contacts = contacts
        self.persist = persist
        self._persisted: Dict[str, Dict[str, Any]] = {}
        self._cumulative = False
        self._rows: List[Dict[str, Any]] = self._rows_from_contacts(contacts)
        self._reset_results()
        self._analyze()

    # -- données -------------------------------------------------------
    @staticmethod
    def _rows_from_contacts(contacts: List[ContactInfo]) -> List[Dict[str, Any]]:
        rows = []
        for c in contacts:
            email = _s(getattr(c, "email", "")).lower()
            if not email:
                continue
            rows.append({
                "email": email,
                "name": _normalize_person_name(_s(getattr(c, "name", ""))),
                "company": _s(getattr(c, "company", "")),
                "famille": _s(getattr(c, "project", "")),
                "projet": _s(getattr(c, "subfolder", "")) or _s(getattr(c, "project", "")),
                "date": _naive(getattr(c, "last_activity", None)),
            })
        return rows

    @staticmethod
    def _load_cumulative_rows(excel_path: str) -> Optional[List[Dict[str, Any]]]:
        """Relit la feuille « Contacts » (cumul de tous les runs)."""
        try:
            df = pd.read_excel(excel_path, sheet_name="Contacts",
                               dtype=str, keep_default_na=False)
        except Exception as e:
            logger.warning(f"[REVIEW] Feuille 'Contacts' illisible, revue limitee au run: {e}")
            return None
        rows = []
        for rec in df.to_dict("records"):
            email = _s(rec.get("Email")).lower()
            if not email:
                continue
            rows.append({
                "email": email,
                "name": _normalize_person_name(_s(rec.get("Nom"))),
                "company": _s(rec.get("Entreprise")),
                "famille": _s(rec.get("Famille Client")),
                "projet": _s(rec.get("Projet")) or _s(rec.get("Famille Client")),
                "date": _parse_dt(rec.get("Derniere activite")),
            })
        return rows or None

    # -- historique persistant -----------------------------------------
    @staticmethod
    def _history_path(excel_path: str) -> str:
        base = os.path.splitext(excel_path)[0]
        # le fichier de secours horodaté de la GUI partage l'historique du principal
        base = re.sub(r"_\d{8}_\d{6}$", "", base)
        return base + ".history.json"

    @staticmethod
    def _load_history(path: str) -> Dict[str, Dict[str, Any]]:
        if not os.path.exists(path):
            return {}
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
            persons = data.get("persons", {})
            if not isinstance(persons, dict):
                raise ValueError("format inattendu")
            return {k: v for k, v in persons.items()
                    if isinstance(v, dict) and isinstance(v.get("entries"), list)}
        except Exception as e:
            logger.warning(f"[REVIEW] Historique illisible ({e}) : conserve en .corrupt")
            try:
                os.replace(path, path + ".corrupt")
            except Exception:
                pass
            return {}

    def _save_history(self, path: str):
        persons = {}
        for key, tl in self._timelines.items():
            persons[key] = {
                "name": tl["name"],
                "entries": [
                    {"company": e["company"],
                     "date": e["date"].strftime("%Y-%m-%d %H:%M:%S") if e["date"] else "",
                     "email": e["email"]}
                    for e in tl["entries"]
                ],
            }
        tmp = path + ".tmp"
        try:
            os.makedirs(os.path.dirname(path) or ".", exist_ok=True)
            with open(tmp, "w", encoding="utf-8") as f:
                json.dump({"version": HISTORY_VERSION, "persons": persons}, f,
                          ensure_ascii=False)
            os.replace(tmp, path)               # écriture atomique
        except Exception as e:
            logger.warning(f"[REVIEW] Sauvegarde de l'historique impossible: {e}")
            try:
                os.remove(tmp)
            except Exception:
                pass

    def _enrich_from_disk(self, excel_path: str):
        """Ajoute l'historique persistant et le cumul du fichier, puis ré-analyse."""
        history_path = self._history_path(excel_path)
        self._persisted = self._load_history(history_path)
        rows = self._load_cumulative_rows(excel_path)
        if rows is not None:
            self._rows = rows
            self._cumulative = True
        self._reset_results()
        self._analyze()
        if self.persist:
            self._save_history(history_path)    # AVANT l'écriture du classeur (verrou Excel)

    # -- analyse -------------------------------------------------------
    def _reset_results(self):
        self._changes: List[Dict[str, Any]] = []
        self._companies: Dict[str, Dict[str, Any]] = {}
        self._stats: List[Dict[str, str]] = []
        self._timelines: Dict[str, Dict[str, Any]] = {}

    def _analyze(self):
        for step in (self._find_company_changes, self._organize_by_company,
                     self._compute_stats):
            try:
                step()
            except Exception:
                logger.exception(f"[REVIEW] Erreur dans {step.__name__}")

    def _gather_sightings(self) -> Dict[str, List[Dict[str, Any]]]:
        """Toutes les observations (entreprise, date, email) par personne."""
        out: Dict[str, List[Dict[str, Any]]] = {}

        def add(key, company, date, email, name):
            company = _valid_company(_s(company), _s(email))
            if not key or not company:
                return
            if not _company_matches_email(company, _s(email)):
                return
            out.setdefault(key, []).append({
                "company": company, "date": date,
                "email": _s(email), "name": name or "",
            })

        # 1) historique des runs précédents
        for key, person in self._persisted.items():
            for e in person.get("entries", []):
                if isinstance(e, dict):
                    add(key, e.get("company"), _parse_dt(e.get("date")),
                        e.get("email"), _s(person.get("name")))

        # 2) contacts de ce run (leur company_history + leur entreprise courante)
        for c in self.contacts:
            email = _s(getattr(c, "email", "")).lower()
            name = _normalize_person_name(_s(getattr(c, "name", "")))
            key = _person_id(email, name)
            for h in getattr(c, "company_history", None) or []:
                if isinstance(h, dict):
                    add(key, h.get("company"), _naive(h.get("date")), email, name)
            add(key, getattr(c, "company", ""), _naive(getattr(c, "last_activity", None)),
                email, name)

        # 3) lignes du fichier cumulé (amorçage + secours si l'historique est perdu)
        if self._cumulative:
            for r in self._rows:
                company = _valid_company(r["company"], r["email"])
                add(_person_id(r["email"], r["name"]), company, r["date"],
                    r["email"], r["name"])
        return out

    @staticmethod
    def _collapse(sightings: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        """Fusionne les variantes d'un même nom d'entreprise (date = la plus récente,
        nom = le plus complet) et trie chronologiquement."""
        entries: List[Dict[str, Any]] = []
        for s in sorted(sightings, key=lambda x: _dt_key(x["date"])):
            for e in entries:
                if _same_company(e["company"], s["company"]):
                    if len(s["company"]) > len(e["company"]):
                        e["company"] = s["company"]
                    if s["date"] and (not e["date"] or s["date"] >= e["date"]):
                        e["date"] = s["date"]
                        e["email"] = s["email"] or e["email"]
                        e["name"] = s["name"] or e["name"]
                    break
            else:
                entries.append(dict(s))
        entries.sort(key=lambda x: _dt_key(x["date"]))
        return entries

    def _find_company_changes(self):
        """Changement = deux entreprises DIFFÉRENTES successives pour une même personne
        (même adresse ou adresses différentes), sur tout l'historique connu."""
        seen = set()
        for key, sightings in self._gather_sightings().items():
            entries = self._collapse(sightings)
            if not entries:
                continue
            name = next((e["name"] for e in reversed(entries) if e["name"]), "")
            self._timelines[key] = {"name": name, "entries": entries}

            for i in range(1, len(entries)):
                old, new = entries[i - 1], entries[i]
                rec_key = (new["email"].lower(), old["company"].lower(),
                           new["company"].lower())
                if rec_key in seen:
                    continue
                same_employer = (old["email"] and old["email"] == new["email"]
                                 and _company_matches_email(old["company"], old["email"])
                                 and _company_matches_email(new["company"], new["email"]))
                if same_employer:
                    continue
                seen.add(rec_key)
                self._changes.append({
                    "Email": new["email"] or "Inconnu",
                    "Nom": name or "Inconnu",
                    "Ancienne entreprise": old["company"],
                    "Nouvelle entreprise": new["company"],
                    "Date": _fmt_date(new["date"]),
                })
        self._changes.sort(key=lambda c: c.get("Date") or "", reverse=True)

    def _organize_by_company(self):
        """Contacts et projets par entreprise (variantes de nom regroupées)."""
        agg: Dict[str, Dict[str, set]] = {}
        for r in self._rows:
            company = _s(r["company"])
            if not company or not r["email"]:
                continue
            a = agg.setdefault(company, {"emails": set(), "projects": set()})
            a["emails"].add(r["email"])
            if r["projet"]:
                a["projects"].add(r["projet"])

        clusters: List[Dict[str, Any]] = []
        for company in sorted(agg, key=lambda c: (-len(c), c.lower())):
            data = agg[company]
            for cl in clusters:
                if _same_company(cl["name"], company):
                    cl["emails"] |= data["emails"]
                    cl["projects"] |= data["projects"]
                    break
            else:
                clusters.append({"name": company, "emails": set(data["emails"]),
                                 "projects": set(data["projects"])})

        clusters.sort(key=lambda c: (-len(c["emails"]), c["name"].lower()))
        self._companies = {
            c["name"]: {
                "contacts": len(c["emails"]),
                "projets": ", ".join(sorted(c["projects"])),
            }
            for c in clusters
        }

    def _compute_stats(self):
        emails = {r["email"] for r in self._rows if r["email"]}
        persons = {_person_id(r["email"], r["name"]) or r["email"]
                   for r in self._rows if r["email"]}

        stats = [{"Statistique": "Contacts extraits (ce run)",
                  "Valeur": str(len(self.contacts))}]
        if self._cumulative:
            stats.append({"Statistique": "Contacts au total (fichier)",
                          "Valeur": str(len(self._rows))})
        stats += [
            {"Statistique": "Emails uniques", "Valeur": str(len(emails))},
            {"Statistique": "Personnes uniques (tous emails confondus)",
             "Valeur": str(len(persons))},
            {"Statistique": "Entreprises différentes", "Valeur": str(len(self._companies))},
            {"Statistique": "Changements d'entreprise", "Valeur": str(len(self._changes))},
        ]
        self._stats = stats

    # -- sorties -------------------------------------------------------
    def gui_data(self) -> Dict[str, Any]:
        """Données pour la GUI / la console (appeler après append_to_excel
        pour la vue cumulée)."""
        return {
            "stats": self._stats,
            "changes": self._changes,
            "companies": [
                (company, data["contacts"], data["projets"])
                for company, data in self._companies.items()
            ],
        }

    def append_to_excel(self, excel_path: str):
        """Ajoute / remplace les feuilles de revue dans le fichier Excel.

        Intègre d'abord l'historique persistant et le cumul du fichier ;
        écrit ensuite sur une copie temporaire remplacée atomiquement."""
        try:
            self._enrich_from_disk(excel_path)
        except Exception:
            logger.exception("[REVIEW] Vue cumulee impossible, revue limitee au run")

        sheets: Dict[str, pd.DataFrame] = {}
        if self._changes:
            sheets["Revue - Changements"] = pd.DataFrame(self._changes)
        if self._companies:
            sheets["Revue - Entreprises"] = pd.DataFrame([
                {"Entreprise": company, "Contacts": data["contacts"],
                 "Projets": data["projets"]}
                for company, data in self._companies.items()
            ])
        if self._stats:
            sheets["Revue - Statistiques"] = pd.DataFrame(self._stats)

        if not sheets:
            logger.info("[REVIEW] Aucune donnee de revue a exporter.")
            return

        tmp_path = excel_path + ".review.tmp.xlsx"
        try:
            shutil.copy2(excel_path, tmp_path)
            self._write_sheets(tmp_path, sheets)
            os.replace(tmp_path, excel_path)    # PermissionError si ouvert dans Excel
            logger.info(f"[REVIEW] Feuilles de revue ajoutées à {excel_path}")
        except PermissionError:
            logger.error("[REVIEW] Fichier Excel verrouille (fermez-le dans Excel "
                         "pendant l'extraction).")
            raise
        except Exception as e:
            logger.error(f"[REVIEW] Erreur lors de l'ajout des feuilles: {e}")
            raise
        finally:
            if os.path.exists(tmp_path):
                try:
                    os.remove(tmp_path)
                except Exception:
                    pass

    def _write_sheets(self, path: str, sheets: Dict[str, pd.DataFrame]):
        try:
            with pd.ExcelWriter(path, engine="openpyxl", mode="a",
                                if_sheet_exists="replace") as writer:
                for name, df in sheets.items():
                    df.to_excel(writer, sheet_name=name, index=False)
                    self._format_sheet(writer.sheets[name], df)
        except (TypeError, ValueError) as e:
            logger.warning(f"[REVIEW] if_sheet_exists indisponible ({e}), "
                           "repli openpyxl manuel.")
            self._append_sheets_openpyxl(path, sheets)

    @staticmethod
    def _append_sheets_openpyxl(excel_path: str, sheets: Dict[str, pd.DataFrame]):
        """Repli : écrit/remplace les feuilles directement via openpyxl."""
        wb = load_workbook(excel_path)
        for name, df in sheets.items():
            if name in wb.sheetnames:
                del wb[name]
            ws = wb.create_sheet(name)
            ws.append([str(c) for c in df.columns])
            for row in df.itertuples(index=False):
                ws.append(["" if v is None else v for v in row])
            AccountReview._format_sheet(ws, df)
        wb.save(excel_path)

    @staticmethod
    def _format_sheet(worksheet, df):
        """Largeur des colonnes, en-tête figé, chaînes « = » protégées."""
        for i, col in enumerate(df.columns, start=1):
            max_len = max(
                (int(df[col].astype(str).str.len().max()) if not df[col].empty else 0),
                len(str(col)),
            ) + 2
            worksheet.column_dimensions[get_column_letter(i)].width = min(max_len, 60)
        worksheet.freeze_panes = "A2"
        for row in worksheet.iter_rows(min_row=2):
            for cell in row:
                if isinstance(cell.value, str) and cell.value.startswith("="):
                    cell.data_type = "s"