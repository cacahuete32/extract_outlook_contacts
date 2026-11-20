"""Extraction des contacts Outlook par projet (ligne de commande / lancement de la GUI).

Corrections par rapport à la version précédente :
  * Checkpoint validé APRÈS un export réussi (finalize), jamais avant ; si
    l'export échoue, rien n'est marqué « traité ». Fichier Excel verrouillé :
    export de secours horodaté (comme la GUI) au lieu d'une « erreur fatale ».
  * Un dossier en erreur n'interrompt plus les suivants (extractor.errors) ;
    Ctrl+C exporte ce qui a été collecté et valide le checkpoint SANS faire
    avancer les watermarks.
  * --gui lancé AVANT toute lecture d'état (la GUI gère le sien) ; plus de
    connexion Outlook et d'état chargés pour rien.
  * Mode --refresh : relit tous les mails et fusionne dans l'Excel existant (complète
    les blancs des contacts déjà scannés) ; sauvegarde horodatée avant écriture.
  * Nouvelles options : --folders (sélection non interactive par nom),
    --no-body-emails, --include-intermediate, --attachments-text, --log-dir.
  * Sélection interactive : « 1-3 », « 1 - 3 », « 1,4-6 », « tous », « q » ;
    stdin fermé / non interactif géré proprement (message, pas de boucle infinie).
  * Checkpoint orphelin (Excel supprimé mais état présent) : proposition de
    rescan complet (automatique en mode non interactif).
  * État lié au nom du fichier de sortie (voir state.py).
  * Arguments validés par argparse (dates, limite, nom de fichier assaini).
  * Journal tournant (5 Mo x 3), bruit des bibliothèques réduit, imports lourds
    (pandas, win32com...) différés : --help et les erreurs d'arguments sont instantanés.
  * Code de sortie : 0 ok, 1 échec, 3 terminé avec erreurs de dossiers,
    130 interrompu (Ctrl+C).
"""
import argparse
import logging
import os
import re
import sys
from datetime import datetime
from logging.handlers import RotatingFileHandler
from pathlib import Path
from typing import List, Optional, Tuple

sys.path.insert(0, str(Path(__file__).parent))

logger = logging.getLogger(__name__)

EXIT_OK = 0
EXIT_FAILED = 1
EXIT_PARTIAL = 3
EXIT_INTERRUPTED = 130


# ----------------------------------------------------------------------
# Arguments
# ----------------------------------------------------------------------

def _date_arg(value: str) -> datetime:
    try:
        return datetime.strptime(value, "%Y-%m-%d")
    except ValueError:
        raise argparse.ArgumentTypeError(f"date invalide ({value}) : format attendu AAAA-MM-JJ")


def _non_negative_int(value: str) -> int:
    try:
        n = int(value)
    except ValueError:
        raise argparse.ArgumentTypeError(f"entier attendu ({value})")
    if n < 0:
        raise argparse.ArgumentTypeError("doit etre >= 0")
    return n


def _sanitize_output_name(name: str) -> str:
    name = os.path.basename((name or "").strip())
    if name.lower().endswith(".xlsx"):
        name = name[:-5]
    return re.sub(r'[<>:"/\\|?*]', "_", name) or "contacts_projets"


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(description="Extraction des contacts Outlook par projet")

    p.add_argument("--output", default="contacts_projets", help="Nom du fichier Excel de sortie")
    p.add_argument("--output-dir", default="output", help="Repertoire de sortie du fichier Excel")
    p.add_argument("--log-dir", default="logs", help="Repertoire des journaux")
    p.add_argument("--debug", action="store_true", help="Active les logs detailles")

    sel = p.add_argument_group("selection des dossiers")
    sel.add_argument("--all", action="store_true",
                     help="Traite tous les dossiers de la boite de reception")
    sel.add_argument("--folders", default="",
                     help="Dossiers a traiter (noms separes par des virgules, sans interaction)")
    sel.add_argument("--include-intermediate", action="store_true",
                     help="Lire aussi les mails des dossiers intermediaires (pas seulement les feuilles)")

    flt = p.add_argument_group("filtres")
    flt.add_argument("--date-from", type=_date_arg, default=None,
                     help="Emails recus depuis cette date (AAAA-MM-JJ)")
    flt.add_argument("--date-to", type=_date_arg, default=None,
                     help="Emails recus jusqu'a cette date (AAAA-MM-JJ)")
    flt.add_argument("--limit", type=_non_negative_int, default=0,
                     help="Nombre max de NOUVEAUX emails par dossier (0 = tous)")
    flt.add_argument("--exclude-domains", default="",
                     help="Domaines a exclure, separes par des virgules (ex: newsletter.fr,spam.com)")
    flt.add_argument("--no-recipients", action="store_true", help="Ne pas enregistrer les destinataires")
    flt.add_argument("--no-body-emails", action="store_true",
                     help="Ne pas extraire les adresses citees dans le corps des messages")

    att = p.add_argument_group("pieces jointes (experimental : le texte n'est pas encore exploite)")
    att.add_argument("--attachments-text", action="store_true",
                     help="Lire le texte des pieces jointes (PDF / OCR) ; lent")
    att.add_argument("--no-attachments", action="store_true", help="Ne pas analyser les pieces jointes")
    att.add_argument("--no-pdf", action="store_true", help="Ignorer les PDF")
    att.add_argument("--no-images", action="store_true", help="Ignorer les images")
    att.add_argument("--no-office", action="store_true", help="Ignorer Word, Excel et PowerPoint")
    att.add_argument("--ocr-inline", action="store_true",
                     help="OCR des images inline (deconseille ; necessite --attachments-text)")

    out = p.add_argument_group("sortie")
    out.add_argument("--sheet-per-project", action="store_true",
                     help="Ajouter une feuille Excel par famille client en plus de la feuille globale")
    out.add_argument("--full-rescan", action="store_true",
                     help="Ignorer le checkpoint, retraiter tous les emails et REMPLACER le fichier Excel")
    out.add_argument("--refresh", action="store_true",
                     help="Actualiser : relire TOUS les emails et fusionner dans le fichier Excel existant "
                          "(complete les informations manquantes des contacts deja scannes, "
                          "sans rien supprimer ; sauvegarde horodatee avant ecriture)")

    p.add_argument("--gui", action="store_true", help="Lancer l'interface graphique")
    return p


def _parse_args(argv=None) -> argparse.Namespace:
    parser = build_parser()
    args = parser.parse_args(argv)
    if args.refresh and args.full_rescan:
        parser.error("--refresh et --full-rescan sont incompatibles "
                     "(--refresh fusionne dans le fichier, --full-rescan le remplace)")
    if args.date_from and args.date_to and args.date_from > args.date_to:
        parser.error("--date-from est posterieur a --date-to")
    if args.date_to:
        args.date_to = args.date_to.replace(hour=23, minute=59, second=59)
    args.output = _sanitize_output_name(args.output)
    args.exclude_domains = [
        d.strip().lower().lstrip("@*") for d in args.exclude_domains.split(",") if d.strip()
    ]
    args.folder_names = [n.strip() for n in args.folders.split(",") if n.strip()]
    return args


def _configure_logging(debug: bool, log_dir: str):
    os.makedirs(log_dir, exist_ok=True)
    file_handler = RotatingFileHandler(
        os.path.join(log_dir, "extraction.log"),
        maxBytes=5 * 1024 * 1024, backupCount=3, encoding="utf-8")
    logging.basicConfig(
        level=logging.DEBUG if debug else logging.INFO,
        format="%(asctime)s - %(levelname)s - %(message)s",
        handlers=[logging.StreamHandler(), file_handler],
    )
    for noisy in ("PIL", "pytesseract", "matplotlib"):
        logging.getLogger(noisy).setLevel(logging.WARNING)


# ----------------------------------------------------------------------
# Sélection des dossiers
# ----------------------------------------------------------------------

def _parse_selection(selection: str, count: int) -> List[int]:
    """'1 3 5', '1-3', '1 - 3', '1,4-6' -> indices (base 1) triés.
    ValueError si le format est invalide, IndexError si hors bornes."""
    text = re.sub(r"\s*-\s*", "-", selection.strip())
    indices = set()
    for part in re.split(r"[\s,;]+", text):
        if not part:
            continue
        if "-" in part:
            a, b = part.split("-", 1)
            start, end = int(a), int(b)
            if start > end:
                start, end = end, start
            indices.update(range(start, end + 1))
        else:
            indices.add(int(part))
    if not indices:
        raise ValueError("selection vide")
    if any(i < 1 or i > count for i in indices):
        raise IndexError(f"Les indices doivent etre compris entre 1 et {count}")
    return sorted(indices)


def select_folders_to_process(outlook_connector):
    inbox = outlook_connector.get_inbox()
    direct_subfolders = outlook_connector.get_direct_subfolders(inbox)

    print("\n" + "=" * 80)
    print("SELECTION DES DOSSIERS A TRAITER")
    print("=" * 80)

    if not direct_subfolders:
        print("Aucun sous-dossier trouve.")
        return []

    if not sys.stdin or not sys.stdin.isatty():
        print("Mode non interactif : utilisez --all ou --folders \"A,B\".")
        return []

    for i, folder in enumerate(direct_subfolders, 1):
        print(f"{i}. {folder.name} ({folder.path})")

    while True:
        try:
            selection = input(
                "\nQuels dossiers traiter ? (ex: 1 3 5, 1-3, 'tous', 'q' pour quitter) : "
            ).strip()
        except EOFError:
            return []

        if selection.lower() in ("q", "quit", "quitter"):
            return []
        if selection.lower() in ("tous", "all", "*"):
            return direct_subfolders
        try:
            return [direct_subfolders[i - 1]
                    for i in _parse_selection(selection, len(direct_subfolders))]
        except IndexError as e:
            print(e)
        except ValueError:
            print("Format invalide. Exemple : 1 3 5 ou 1-3 ou tous")


def _select_by_names(outlook_connector, names: List[str]):
    """--folders : correspondance insensible à la casse ; None si un nom est inconnu."""
    available = outlook_connector.get_direct_subfolders(outlook_connector.get_inbox())
    by_name = {f.name.lower(): f for f in available}
    missing = [n for n in names if n.lower() not in by_name]
    if missing:
        print(f"Dossier(s) introuvable(s) : {', '.join(missing)}")
        print("Disponibles : " + ", ".join(f.name for f in available))
        return None
    selected, seen = [], set()
    for n in names:
        f = by_name[n.lower()]
        if f.path not in seen:
            seen.add(f.path)
            selected.append(f)
    return selected


# ----------------------------------------------------------------------
# Extraction
# ----------------------------------------------------------------------

def _check_orphan_state(args, state) -> bool:
    """Excel supprimé mais état présent : les mails « déjà traités » seraient
    ignorés et le nouveau fichier serait incomplet. Renvoie True si rescan complet."""
    out_path = os.path.join(args.output_dir, f"{args.output}.xlsx")
    if args.refresh or args.full_rescan or not state.processed_mail_ids or os.path.exists(out_path):
        return args.full_rescan                    # refresh relit tout de toute façon

    print(f"\nATTENTION : '{out_path}' n'existe pas, mais {len(state.processed_mail_ids)} "
          "emails sont marques comme deja traites.")
    interactive = bool(sys.stdin) and sys.stdin.isatty() and not args.all and not args.folder_names
    if interactive:
        try:
            answer = input("Faire un rescan complet ? [O/n] ").strip().lower()
        except EOFError:
            answer = "o"
        return answer in ("", "o", "oui", "y", "yes")
    print("Mode non interactif : rescan complet automatique.")
    return True


def _export(args, contacts) -> Tuple[Optional[str], bool]:
    """Exporte vers Excel. Renvoie (chemin, repli_utilise)."""
    from exporter import DataExporter
    exporter = DataExporter(contacts)
    try:
        path = exporter.export_to_excel(
            args.output, output_dir=args.output_dir,
            sheet_per_project=args.sheet_per_project,
            incremental=not args.full_rescan,
            keep_backup=args.refresh)
        return path, False
    except PermissionError:
        alt = f"{args.output}_{datetime.now():%Y%m%d_%H%M%S}"
        logger.warning(f"Fichier verrouille (ouvert dans Excel ?). Export de secours "
                       f"(nouveaux contacts uniquement) : {alt}.xlsx")
        path = exporter.export_to_excel(
            alt, output_dir=args.output_dir,
            sheet_per_project=args.sheet_per_project, incremental=False)
        return path, True


def run_extraction(args, outlook_connector, folders_to_process, state) -> Tuple[Optional[str], int]:
    """Renvoie (fichier_généré, code_de_sortie)."""
    from extractor import ContactExtractor
    from exporter import ExistingFileError
    from review import AccountReview

    extractor = ContactExtractor(
        outlook_connector=outlook_connector,
        process_attachments=not args.no_attachments,
        skip_pdf=args.no_pdf,
        skip_images=args.no_images,
        skip_office=args.no_office,
        date_from=args.date_from,
        date_to=args.date_to,
        limit=args.limit or None,
        include_recipients=not args.no_recipients,
        exclude_domains=args.exclude_domains,
        state=state,                              # mode incrémental
        full_rescan=args.full_rescan,             # repartir de zéro si demandé
        refresh=args.refresh,                     # tout relire, fusionner dans l'Excel
        process_body_emails=not args.no_body_emails,
        include_intermediate=args.include_intermediate,
        read_attachment_text=args.attachments_text,
    )

    interrupted = False
    try:
        for folder in folders_to_process:
            logger.info(f"Traitement dossier : {folder.name}")
            extractor.extract_from_folder(folder)   # une erreur de dossier n'arrête pas les autres
    except KeyboardInterrupt:
        interrupted = True
        extractor.incomplete_folders.add("*")       # lecture partielle : aucun watermark ne doit avancer
        print("\nInterruption (Ctrl+C) : export de ce qui a ete collecte...")

    errors = list(extractor.errors)

    # --- Aucun contact : le checkpoint reste valide (rien n'est perdu) ---
    if not extractor.contacts:
        print("Aucun contact extrait." + (" (erreurs, voir le journal)" if errors else ""))
        extractor.finalize()
        return None, EXIT_INTERRUPTED if interrupted else (EXIT_PARTIAL if errors else EXIT_OK)

    # --- Export : le checkpoint n'est validé QUE si l'export réussit ---
    try:
        output_file, fallback = _export(args, extractor.contacts)
    except ExistingFileError as e:
        print(f"ERREUR : {e}")
        print("Checkpoint NON valide : ces emails seront retraites au prochain lancement.")
        return None, EXIT_FAILED
    except Exception as e:
        logger.exception(f"Echec de l'export Excel : {e}")
        print(f"ERREUR : export impossible ({e}).")
        print("Checkpoint NON valide : ces emails seront retraites au prochain lancement.")
        return None, EXIT_FAILED

    extractor.finalize()

    # --- Revue de compte (secondaire : n'invalide jamais le fichier Contacts) ---
    review_data = None
    if output_file:
        try:
            review = AccountReview(extractor.contacts)
            review.append_to_excel(output_file)
            review_data = review.gui_data()            # APRES append_to_excel : vue cumulée
        except Exception as e:
            logger.exception(f"Erreur lors de la revue de compte : {e}")
            print(f"ATTENTION : la revue de compte a echoue ({e})")

    # --- Résumé console ---
    print("\n" + "=" * 80)
    print("RESUME")
    print("=" * 80)
    print(f"Dossiers traites : {len(folders_to_process)}" + (" (mode refresh)" if args.refresh else ""))
    print(f"Contacts extraits : {len(extractor.contacts)}")
    print(f"Fichier genere : {output_file}" + (" (fichier de secours)" if fallback else ""))
    if state is not None:
        print(f"Checkpoint : {state.path}")

    if review_data:
        for stat in review_data["stats"]:
            print(f"  - {stat['Statistique']} : {stat['Valeur']}")
        for ch in review_data["changes"][:20]:
            print(f"      * {ch.get('Nom', '')} ({ch.get('Email', '')}): "
                  f"{ch.get('Ancienne entreprise', '')} -> {ch.get('Nouvelle entreprise', '')} "
                  f"[{ch.get('Date', '')}]")

    if errors:
        print(f"ATTENTION : {len(errors)} dossier(s) en erreur (voir le journal) ; "
              "le fichier contient les contacts des autres dossiers.")
        for msg in errors[:5]:
            print(f"  - {msg}")
    if interrupted:
        print("ATTENTION : extraction interrompue ; le fichier est partiel.")
    print("=" * 80)

    if interrupted:
        return output_file, EXIT_INTERRUPTED
    return output_file, EXIT_PARTIAL if errors else EXIT_OK


# ----------------------------------------------------------------------
# Point d'entrée
# ----------------------------------------------------------------------

def main(argv=None) -> int:
    args = _parse_args(argv)
    os.makedirs(args.output_dir, exist_ok=True)
    _configure_logging(args.debug, args.log_dir)

    outlook = None
    try:
        # Imports lourds différés : --help et les erreurs d'arguments restent instantanés
        from outlook import OutlookConnector, OutlookConnectionError
        from state import ExtractionState

        logger.info("=" * 80)
        logger.info("DEMARRAGE EXTRACTION CONTACTS")
        logger.info("=" * 80)

        try:
            outlook = OutlookConnector(ocr_inline_images=args.ocr_inline)
        except OutlookConnectionError as e:
            logger.error(str(e))
            print(f"ERREUR : {e}")
            return EXIT_FAILED
        logger.info(f"Boite Outlook : {outlook.mailbox.Name}")

        # La GUI gère son propre état : rien à lire ici
        if args.gui:
            import tkinter as tk
            from gui import ContactExtractionGUI
            root = tk.Tk()
            app = ContactExtractionGUI(root, outlook)
            if app.ready:
                root.mainloop()
            return EXIT_OK

        if args.folder_names:
            folders_to_process = _select_by_names(outlook, args.folder_names)
            if folders_to_process is None:
                return EXIT_FAILED
        elif args.all:
            folders_to_process = outlook.get_direct_subfolders(outlook.get_inbox())
        else:
            folders_to_process = select_folders_to_process(outlook)

        if not folders_to_process:
            print("Aucun dossier selectionne.")
            return EXIT_OK

        # --- Checkpoint incrémental (lié au nom du fichier de sortie) ---
        state = ExtractionState(output_dir=args.output_dir, name=args.output)
        args.full_rescan = _check_orphan_state(args, state)
        if args.full_rescan:
            state.reset()
            logger.info("[STATE] Full rescan : checkpoint reinitialise")
        if args.refresh:
            logger.info("[STATE] Refresh : tous les emails seront relus et fusionnes dans le fichier existant")
        logger.info(f"[STATE] {len(state.processed_mail_ids)} emails deja traites "
                    f"({len(state.folder_watermarks)} dossiers avec watermark)")

        _, code = run_extraction(args, outlook, folders_to_process, state)
        return code

    except KeyboardInterrupt:
        print("\nInterrompu.")
        return EXIT_INTERRUPTED
    except Exception as e:
        logger.exception(f"Erreur fatale : {e}")
        print(f"ERREUR FATALE : {e} (voir {os.path.join(args.log_dir, 'extraction.log')})")
        return EXIT_FAILED
    finally:
        if outlook:
            try:
                outlook.close()
            except Exception:
                pass


if __name__ == "__main__":
    sys.exit(main())