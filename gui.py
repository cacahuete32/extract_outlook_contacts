"""Interface graphique d'extraction des contacts Outlook.

Améliorations intégrées :
  * Bouton Pause / Reprise fonctionnel et thread-safe.
  * Bouton d'ouverture rapide du fichier Excel généré.
  * Améliorations ergonomiques et visuelles (thème, dispositions, états de boutons).
  * Checkpoint validé APRÈS un export réussi.
  * Gestion robuste des files et de la file unique vers l'UI (`ui_queue`).
"""
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import os
import re
import sys
import shutil
import logging
import traceback
from dataclasses import dataclass, field
from pathlib import Path
from datetime import datetime, timedelta
from queue import Queue, Empty, Full
from threading import Thread, Event, Lock
from typing import Any, Dict, List, Optional, Set, Tuple

import pythoncom

sys.path.insert(0, str(Path(__file__).parent))

from outlook import OutlookConnector, OutlookFolder
from extractor import ContactExtractor, EmailData, MailConverter, PR_SMTP_ADDRESS, _mail_body
from exporter import DataExporter
from review import AccountReview
from state import ExtractionState
from utils import is_valid_email

WATERMARK_SAFETY_DAYS = 1
EMAIL_QUEUE_MAX = 200
PROGRESS_EVERY = 10
LOG_EVERY = 50
MAX_LOG_LINES = 5000
UI_MESSAGES_PER_TICK = 200


# ----------------------------------------------------------------------
# Utilitaires
# ----------------------------------------------------------------------

def _com_str(value) -> str:
    if isinstance(value, str):
        return value.strip()
    return ""


def _safe(obj, attr, default=None):
    try:
        return getattr(obj, attr)
    except Exception:
        return default


def _is_excluded(addr: str, domains: List[str]) -> bool:
    if not addr or not isinstance(addr, str):
        return True
    addr = addr.lower().strip()
    for d in domains:
        if addr == d or addr.endswith("@" + d) or addr.endswith("." + d):
            return True
    return False


class QueueLogHandler(logging.Handler):
    def __init__(self, queue: Queue):
        super().__init__()
        self.queue = queue

    def emit(self, record):
        try:
            self.queue.put(("LOG", record.levelname, self.format(record)))
        except Exception:
            pass


# ----------------------------------------------------------------------
# Configuration et contexte d'exécution
# ----------------------------------------------------------------------

@dataclass
class ExtractionConfig:
    output: str
    output_dir: str
    date_from: Optional[datetime]
    date_to: Optional[datetime]
    limit: int
    include_senders: bool
    include_recipients: bool
    exclude_domains: List[str]
    sheet_per_project: bool
    ocr_inline: bool
    debug: bool
    all_folders: bool
    no_attachments: bool
    no_pdf: bool
    no_images: bool
    no_office: bool
    full_rescan: bool
    refresh: bool
    extract_body_emails: bool


@dataclass
class RunContext:
    cfg: ExtractionConfig
    state: ExtractionState
    cancel_event: Event = field(default_factory=Event)
    loader_stop: Event = field(default_factory=Event)
    pause_event: Event = field(default_factory=Event)  # Activé quand en pause
    pause_lock: Lock = field(default_factory=Lock)
    email_queue: Queue = field(default_factory=lambda: Queue(maxsize=EMAIL_QUEUE_MAX))
    incomplete_folders: Set[str] = field(default_factory=set)
    total: int = 0
    processed: int = 0
    loading_done: bool = False
    loader_thread: Optional[Thread] = None


# ----------------------------------------------------------------------
# Lecture d'un mail Outlook -> EmailData (avec caches COM)
# ----------------------------------------------------------------------

def _empty_sender() -> Dict[str, Any]:
    return {"email": "", "name": "", "job_title": "", "company": "",
            "phone": "", "is_exchange": False}


class MailReader:
    def __init__(self, outlook: OutlookConnector, cfg: ExtractionConfig, warn):
        self.outlook = outlook
        self.cfg = cfg
        self._warn = warn
        self._resolver = MailConverter(outlook=outlook, include_recipients=False)
        self._sender_cache: Dict[str, Dict[str, Any]] = {}
        self._recipient_cache: Dict[str, Dict[str, Any]] = {}

    def _sender(self, mail) -> Dict[str, Any]:
        raw_addr = _com_str(_safe(mail, "SenderEmailAddress"))
        raw_name = _com_str(_safe(mail, "SenderName"))
        key = (raw_addr or raw_name).lower()
        if key and key in self._sender_cache:
            return self._sender_cache[key]

        info = _empty_sender()
        info["email"] = raw_addr
        info["name"] = raw_name

        sender_obj = _safe(mail, "Sender")
        if sender_obj is not None:
            exch = None
            try:
                exch = sender_obj.GetExchangeUser()
            except Exception:
                exch = None
            if exch:
                info["email"] = _com_str(_safe(exch, "PrimarySmtpAddress")) or info["email"]
                info["name"] = _com_str(_safe(exch, "Name")) or info["name"]
                info["job_title"] = _com_str(_safe(exch, "JobTitle"))
                info["company"] = _com_str(_safe(exch, "CompanyName"))
                info["phone"] = (_com_str(_safe(exch, "BusinessTelephoneNumber"))
                                 or _com_str(_safe(exch, "MobileTelephoneNumber")))
                info["is_exchange"] = True

            if "@" not in info["email"]:
                try:
                    smtp = sender_obj.PropertyAccessor.GetProperty(PR_SMTP_ADDRESS)
                    info["email"] = _com_str(smtp) or info["email"]
                except Exception:
                    pass

        self._resolver._fill_missing_from_directory(info)

        if key:
            self._sender_cache[key] = info
        return info

    def _recipients(self, mail) -> List[Dict[str, Any]]:
        out: List[Dict[str, Any]] = []
        seen: Set[str] = set()
        recipients = _safe(mail, "Recipients")
        if recipients is None:
            return out
        try:
            items = list(recipients)
        except Exception:
            return out
        for r in items:
            try:
                addr_raw = _com_str(_safe(r, "Address"))
                key = (addr_raw or _com_str(_safe(r, "Name"))).lower()
                if key and key in self._recipient_cache:
                    info = self._recipient_cache[key]
                else:
                    resolved = self._resolver._resolve_recipient(r, addr_raw)
                    info = {
                        "email": resolved["email"],
                        "name": resolved["name"],
                        "job_title": resolved.get("job_title", "") or "",
                        "company": resolved.get("company", "") or "",
                        "phone": resolved.get("phone", "") or "",
                        "is_exchange": bool(resolved.get("is_exchange")),
                    }
                    if key:
                        self._recipient_cache[key] = info
                email = info["email"]
                if (email and email not in seen and is_valid_email(email)
                        and not _is_excluded(email, self.cfg.exclude_domains)):
                    seen.add(email)
                    out.append(info)
            except Exception:
                continue
        return out

    def read(self, mail, folder: OutlookFolder, famille: str,
             mail_id: Optional[str], received: Optional[datetime]) -> EmailData:
        cfg = self.cfg
        subject = _com_str(_safe(mail, "Subject"))
        body = _mail_body(mail)

        sender = self._sender(mail) if cfg.include_senders else _empty_sender()
        sender_email = sender["email"].lower().strip()
        if sender_email and _is_excluded(sender_email, cfg.exclude_domains):
            sender = _empty_sender()
            sender_email = ""

        attachments_text = ""
        inline_images_text = ""
        if not cfg.no_attachments:
            try:
                attachments_text = self.outlook.get_attachments_text(
                    mail, skip_pdf=cfg.no_pdf, skip_images=cfg.no_images,
                    skip_office=cfg.no_office)
            except Exception as e:
                self._warn(f"Erreur pieces jointes: {e}")
        if not cfg.no_images and cfg.ocr_inline:
            try:
                inline_images_text = self.outlook.get_inline_images_text(mail)
            except Exception as e:
                self._warn(f"Erreur images inline: {e}")

        recipients = self._recipients(mail) if cfg.include_recipients else []
        vcf_contacts = []
        resolver = self._resolver
        try:
            vcf_contacts = self.outlook.get_vcf_contacts(mail)
        except Exception as e:
            self._warn(f"Erreur vCard: {e}")

        return EmailData(
            folder_path=folder.path,
            project_name=famille,
            subfolder=folder.name,
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
            recipients=recipients,
            vcf_contacts=vcf_contacts,
            mail_id=mail_id or "",
        )


# ----------------------------------------------------------------------
# Interface Graphique
# ----------------------------------------------------------------------

class ContactExtractionGUI:
    def __init__(self, root, outlook_connector=None):
        self.root = root
        self.ready = True
        self.root.title("Extraction Contacts Outlook - Projets")
        self.root.geometry("950x720")
        self.root.minsize(850, 620)

        self.output_var = tk.StringVar(value="contacts_projets")
        self.output_dir_var = tk.StringVar(value="output")
        self.date_from_var = tk.StringVar(value="")
        self.date_to_var = tk.StringVar(value="")
        self.limit_var = tk.IntVar(value=0)
        self.include_senders_var = tk.BooleanVar(value=True)
        self.include_recipients_var = tk.BooleanVar(value=True)
        self.exclude_domains_var = tk.StringVar(value="")
        self.sheet_per_project_var = tk.BooleanVar(value=False)
        self.ocr_inline_var = tk.BooleanVar(value=False)
        self.debug_var = tk.BooleanVar(value=False)
        self.all_folders_var = tk.BooleanVar(value=False)
        self.no_attachments_var = tk.BooleanVar(value=False)
        self.no_pdf_var = tk.BooleanVar(value=False)
        self.no_images_var = tk.BooleanVar(value=False)
        self.no_office_var = tk.BooleanVar(value=True)
        self.full_rescan_var = tk.BooleanVar(value=False)
        self.refresh_var = tk.BooleanVar(value=False)
        self.extract_body_emails_var = tk.BooleanVar(value=True)

        self.folder_vars: Dict[str, tk.BooleanVar] = {}
        self.folders_list: List[OutlookFolder] = []
        self.last_generated_file: Optional[str] = None

        self.ui_queue: Queue = Queue()
        self.extraction_in_progress = False
        self.run: Optional[RunContext] = None

        self.log_handler = QueueLogHandler(self.ui_queue)
        self.log_handler.setFormatter(logging.Formatter("%(message)s"))
        logging.getLogger().addHandler(self.log_handler)

        self.outlook = outlook_connector
        if not self.outlook:
            if not self._connect_outlook():
                self.ready = False
                return

        self._create_widgets()
        self._load_folders()
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self.root.after(100, self._process_queue)

    def _connect_outlook(self) -> bool:
        try:
            self.outlook = OutlookConnector()
            return True
        except Exception as e:
            messagebox.showerror(
                "Erreur de connexion",
                f"Impossible de se connecter a Outlook:\n{e}"
            )
            logging.getLogger().removeHandler(self.log_handler)
            self.root.destroy()
            return False

    def _create_widgets(self):
        style = ttk.Style()
        style.configure("Accent.TButton", foreground="green", font=("Arial", 10, "bold"))
        style.configure("Warning.TButton", foreground="darkorange", font=("Arial", 10, "bold"))

        main_frame = ttk.Frame(self.root, padding=12)
        main_frame.pack(fill=tk.BOTH, expand=True)

        notebook = ttk.Notebook(main_frame)
        notebook.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        folders_tab = ttk.Frame(notebook, padding=5)
        notebook.add(folders_tab, text="📁 Dossiers")
        options_tab = ttk.Frame(notebook, padding=5)
        notebook.add(options_tab, text="⚙️ Options")
        logs_tab = ttk.Frame(notebook, padding=5)
        notebook.add(logs_tab, text="📜 Logs")

        # --- Onglet dossiers
        folders_frame = ttk.LabelFrame(folders_tab, text=" Sélectionnez les dossiers à traiter ", padding=10)
        folders_frame.pack(fill=tk.BOTH, expand=True, pady=5)

        ttk.Button(folders_tab, text="🔄 Rafraîchir la liste des dossiers",
                   command=self._load_folders).pack(pady=5, anchor=tk.W)

        canvas = tk.Canvas(folders_frame, highlightthickness=0)
        scrollbar = ttk.Scrollbar(folders_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        scrollable_frame.bind(
            "<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        window_id = canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.bind("<Configure>", lambda e: canvas.itemconfigure(window_id, width=e.width))
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        self.folders_container = scrollable_frame

        def _on_wheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        canvas.bind("<Enter>", lambda e: canvas.bind_all("<MouseWheel>", _on_wheel))
        canvas.bind("<Leave>", lambda e: canvas.unbind_all("<MouseWheel>"))

        buttons_frame = ttk.Frame(folders_tab)
        buttons_frame.pack(fill=tk.X, pady=5)
        ttk.Button(buttons_frame, text="✅ Tout sélectionner",
                   command=self._select_all_folders).pack(side=tk.LEFT, padx=2)
        ttk.Button(buttons_frame, text="❌ Tout désélectionner",
                   command=self._deselect_all_folders).pack(side=tk.LEFT, padx=2)

        # --- Onglet options
        options_frame = ttk.LabelFrame(options_tab, text=" Paramètres d'extraction ", padding=10)
        options_frame.pack(fill=tk.BOTH, expand=True, pady=5)

        def _row(r, label, widget):
            ttk.Label(options_frame, text=label).grid(row=r, column=0, sticky=tk.W, pady=3)
            widget.grid(row=r, column=1, sticky=tk.W, pady=3)

        _row(0, "Nom du fichier Excel :",
             ttk.Entry(options_frame, textvariable=self.output_var, width=35))
        _row(1, "Répertoire de sortie :",
             ttk.Entry(options_frame, textvariable=self.output_dir_var, width=35))
        _row(2, "Emails reçus depuis (AAAA-MM-JJ) :",
             ttk.Entry(options_frame, textvariable=self.date_from_var, width=20))
        _row(3, "Emails reçus jusqu'à (AAAA-MM-JJ) :",
             ttk.Entry(options_frame, textvariable=self.date_to_var, width=20))
        _row(4, "Limite de NOUVEAUX emails par dossier (0 = tous) :",
             ttk.Spinbox(options_frame, from_=0, to=100000, textvariable=self.limit_var, width=10))
        _row(5, "Domaines à exclure (séparés par des virgules) :",
             ttk.Entry(options_frame, textvariable=self.exclude_domains_var, width=35))

        options = [
            ("Inclure les expéditeurs", self.include_senders_var),
            ("Inclure les destinataires", self.include_recipients_var),
            ("Mode debug (logs détaillés)", self.debug_var),
            ("Inclure aussi les mails des dossiers intermédiaires", self.all_folders_var),
            ("Une feuille Excel par famille client", self.sheet_per_project_var),
            ("OCR des images inline (logos - déconseillé)", self.ocr_inline_var),
            ("Ignorer les pièces jointes", self.no_attachments_var),
            ("Ignorer les PDF", self.no_pdf_var),
            ("Ignorer les images", self.no_images_var),
            ("Ignorer Word/Excel/PPT", self.no_office_var),
            ("Rescan complet (ignore le checkpoint, REMPLACE le fichier Excel)", self.full_rescan_var),
            ("Actualiser : relire tous les mails et FUSIONNER dans le fichier", self.refresh_var),
            ("Extraire les emails du corps des messages", self.extract_body_emails_var),
        ]
        for i, (text, var) in enumerate(options, start=6):
            ttk.Checkbutton(options_frame, text=text, variable=var,
                            onvalue=True, offvalue=False
                            ).grid(row=i, column=0, columnspan=2, sticky=tk.W, pady=2)

        # --- Onglet logs
        logs_frame = ttk.LabelFrame(logs_tab, text=" Journal des opérations ", padding=10)
        logs_frame.pack(fill=tk.BOTH, expand=True)
        self.log_text = scrolledtext.ScrolledText(
            logs_frame, wrap=tk.WORD, width=80, height=20,
            state=tk.DISABLED, font=("Consolas", 10))
        self.log_text.pack(fill=tk.BOTH, expand=True)
        self.log_text.tag_config("error", foreground="red")
        self.log_text.tag_config("warning", foreground="orange")
        self.log_text.tag_config("debug", foreground="gray")

        # --- Barre d'actions et boutons principaux
        action_frame = ttk.Frame(main_frame)
        action_frame.pack(fill=tk.X, pady=5)

        self.start_button = ttk.Button(
            action_frame, text="🚀 Lancer l'extraction",
            command=self._start_extraction, style="Accent.TButton")
        self.start_button.pack(side=tk.RIGHT, padx=3)

        self.pause_button = ttk.Button(
            action_frame, text="⏸ Mettre en pause",
            command=self._toggle_pause, state=tk.DISABLED, style="Warning.TButton")
        self.pause_button.pack(side=tk.RIGHT, padx=3)

        self.stop_button = ttk.Button(
            action_frame, text="⏹ Arrêter", command=self._cancel_extraction,
            state=tk.DISABLED)
        self.stop_button.pack(side=tk.RIGHT, padx=3)

        self.open_excel_button = ttk.Button(
            action_frame, text="📂 Ouvrir le fichier Excel",
            command=self._open_excel_file, state=tk.DISABLED)
        self.open_excel_button.pack(side=tk.LEFT, padx=3)

        ttk.Button(action_frame, text="❌ Quitter",
                   command=self._on_close).pack(side=tk.LEFT, padx=3)

        # --- Barre de progression
        self.progress_frame = ttk.Frame(main_frame)
        self.progress_frame.pack(fill=tk.X, pady=5)
        self.progress_label = ttk.Label(self.progress_frame, text="0/0 emails traités")
        self.progress_label.pack(side=tk.LEFT)
        self.progress = ttk.Progressbar(
            self.progress_frame, orient=tk.HORIZONTAL, length=300, mode='determinate')
        self.progress.pack(side=tk.LEFT, padx=10, fill=tk.X, expand=True)
        self.progress_frame.pack_forget()

    @staticmethod
    def _folder_key(folder: OutlookFolder) -> str:
        return folder.entry_id or folder.path

    def _load_folders(self):
        self._log_info("Chargement des dossiers Outlook...")
        try:
            for widget in self.folders_container.winfo_children():
                widget.destroy()
            self.folder_vars.clear()
            self.folders_list = []

            inbox = self.outlook.get_inbox()
            direct_subfolders = self.outlook.get_direct_subfolders(inbox)

            if not direct_subfolders:
                ttk.Label(self.folders_container,
                          text="Aucun sous-dossier trouvé dans la boîte de réception.",
                          foreground="red").pack(pady=10)
                return

            for folder in direct_subfolders:
                var = tk.BooleanVar(value=True)
                self.folder_vars[self._folder_key(folder)] = var
                self.folders_list.append(folder)

                frame = ttk.Frame(self.folders_container)
                frame.pack(fill=tk.X, pady=2)
                ttk.Checkbutton(frame, variable=var).pack(side=tk.LEFT, padx=5)
                ttk.Label(frame, text=f"{folder.name} ({folder.path})"
                          ).pack(side=tk.LEFT, fill=tk.X, expand=True)

            self._log_info(f"Trouvé {len(direct_subfolders)} dossiers.")
        except Exception as e:
            self._log_error(f"Erreur lors du chargement des dossiers : {e}")

    def _select_all_folders(self):
        for var in self.folder_vars.values():
            var.set(True)

    def _deselect_all_folders(self):
        for var in self.folder_vars.values():
            var.set(False)

    @staticmethod
    def _parse_date_field(value: str) -> Optional[datetime]:
        value = value.strip()
        if not value:
            return None
        return datetime.strptime(value, "%Y-%m-%d")

    def _build_config(self) -> Optional[ExtractionConfig]:
        try:
            date_from = self._parse_date_field(self.date_from_var.get())
            date_to = self._parse_date_field(self.date_to_var.get())
        except ValueError:
            messagebox.showwarning(
                "Date invalide", "Format de date attendu : AAAA-MM-JJ (ex : 2026-01-31).")
            return None
        if date_to:
            date_to = date_to.replace(hour=23, minute=59, second=59)
        if date_from and date_to and date_from > date_to:
            messagebox.showwarning(
                "Dates incohérentes", "La date de début est postérieure à la date de fin.")
            return None

        try:
            limit = max(0, int(self.limit_var.get()))
        except (tk.TclError, ValueError):
            messagebox.showwarning("Limite invalide", "La limite doit être un entier (0 = tous).")
            return None

        output = os.path.basename(self.output_var.get().strip())
        if output.lower().endswith(".xlsx"):
            output = output[:-5]
        output = re.sub(r'[<>:"/\\|?*]', "_", output) or "contacts_projets"
        output_dir = self.output_dir_var.get().strip() or "output"

        return ExtractionConfig(
            output=output,
            output_dir=output_dir,
            date_from=date_from,
            date_to=date_to,
            limit=limit,
            include_senders=self.include_senders_var.get(),
            include_recipients=self.include_recipients_var.get(),
            exclude_domains=[
                d.strip().lower().lstrip("@*")
                for d in self.exclude_domains_var.get().split(",") if d.strip()
            ],
            sheet_per_project=self.sheet_per_project_var.get(),
            ocr_inline=self.ocr_inline_var.get(),
            debug=self.debug_var.get(),
            all_folders=self.all_folders_var.get(),
            no_attachments=self.no_attachments_var.get(),
            no_pdf=self.no_pdf_var.get(),
            no_images=self.no_images_var.get(),
            no_office=self.no_office_var.get(),
            full_rescan=self.full_rescan_var.get(),
            refresh=self.refresh_var.get(),
            extract_body_emails=self.extract_body_emails_var.get(),
        )

    def _start_extraction(self):
        if self.extraction_in_progress:
            return

        cfg = self._build_config()
        if cfg is None:
            return

        if cfg.refresh and cfg.full_rescan:
            messagebox.showwarning(
                "Options incompatibles",
                "« Actualiser » fusionne dans le fichier existant, « Rescan complet » le remplace.\n"
                "Cochez l'une ou l'autre.")
            return

        selected = [f for f in self.folders_list
                    if self.folder_vars[self._folder_key(f)].get()]
        if not selected:
            messagebox.showwarning("Aucun dossier", "Veuillez sélectionner au moins un dossier.")
            return

        state = ExtractionState(output_dir=cfg.output_dir, name=cfg.output)

        out_path = os.path.join(cfg.output_dir, f"{cfg.output}.xlsx")
        self.last_generated_file = out_path

        if (not cfg.full_rescan and not cfg.refresh and state.processed_mail_ids
                and not os.path.exists(out_path)):
            if messagebox.askyesno(
                "Checkpoint sans fichier Excel",
                f"Le fichier '{out_path}' n'existe pas, mais {len(state.processed_mail_ids)} "
                "emails sont marqués comme déjà traités.\n\n"
                "Faire un rescan complet ?"):
                cfg.full_rescan = True
                self.full_rescan_var.set(True)

        if cfg.full_rescan:
            state.processed_mail_ids = set()
            state.folder_watermarks = {}
            self._log_info("Rescan complet demandé : checkpoint réinitialisé.")
        elif cfg.refresh:
            self._log_info("Mode actualisation : relecture et fusion dans le fichier existant.")
        else:
            self._log_info(
                f"Checkpoint: {len(state.processed_mail_ids)} emails déjà traités, "
                f"{len(state.folder_watermarks)} dossiers avec watermark.")

        logging.getLogger().setLevel(logging.DEBUG if cfg.debug else logging.INFO)
        self.log_handler.setLevel(logging.DEBUG if cfg.debug else logging.INFO)

        run = RunContext(cfg=cfg, state=state)
        self.run = run
        self.extraction_in_progress = True

        self._log_info("=" * 60)
        self._log_info("DÉMARRAGE DE L'EXTRACTION")

        self.start_button.config(state=tk.DISABLED)
        self.pause_button.config(state=tk.NORMAL, text="⏸ Mettre en pause")
        self.stop_button.config(state=tk.NORMAL)
        self.open_excel_button.config(state=tk.DISABLED)
        self.progress_frame.pack(fill=tk.X, pady=5)
        self.progress["maximum"] = 100
        self.progress["value"] = 0
        self.progress_label.config(text="0/0 emails traités")

        entry_ids = [f.entry_id for f in selected if f.entry_id]

        run.loader_thread = Thread(
            target=self._load_emails_to_queue, args=(entry_ids, run), daemon=True)
        run.loader_thread.start()
        Thread(target=self._worker_process_emails, args=(run,), daemon=True).start()

    def _toggle_pause(self):
        if not self.run:
            return
        if self.run.pause_event.is_set():
            # Reprise
            self.run.pause_event.clear()
            self.pause_button.config(text="⏸ Mettre en pause")
            self._log_info("▶ Extraction reprise.")
        else:
            # Pause
            self.run.pause_event.set()
            self.pause_button.config(text="▶ Reprendre")
            self._log_info("⏸ Extraction en pause...")

    def _cancel_extraction(self):
        if self.run and self.extraction_in_progress:
            self._log_warning("Arrêt demandé : finalisation de ce qui a déjà été traité...")
            if self.run.pause_event.is_set():
                self.run.pause_event.clear()  # Débloquer si en pause pour quitter proprement
            self.run.cancel_event.set()
            self.run.loader_stop.set()
            self.stop_button.config(state=tk.DISABLED)
            self.pause_button.config(state=tk.DISABLED)

    def _open_excel_file(self):
        if self.last_generated_file and os.path.exists(self.last_generated_file):
            try:
                if sys.platform == "win32":
                    os.startfile(os.path.abspath(self.last_generated_file))
                elif sys.platform == "darwin":
                    os.system(f"open '{self.last_generated_file}'")
                else:
                    os.system(f"xdg-open '{self.last_generated_file}'")
                self._log_info(f"Ouverture du fichier Excel : {self.last_generated_file}")
            except Exception as e:
                messagebox.showerror("Erreur", f"Impossible d'ouvrir le fichier :\n{e}")
        else:
            messagebox.showwarning("Introuvable", "Aucun fichier Excel valide n'a été trouvé.")

    @staticmethod
    def _effective_date_from(folder_path: str, run: RunContext) -> Optional[datetime]:
        cfg, state = run.cfg, run.state
        if cfg.date_from:
            return cfg.date_from
        if not (cfg.full_rescan or cfg.refresh):
            watermark = state.get_watermark(folder_path)
            if watermark:
                try:
                    wm = datetime.strptime(watermark, "%Y-%m-%d %H:%M:%S")
                    return wm - timedelta(days=WATERMARK_SAFETY_DAYS)
                except ValueError:
                    pass
        return None

    @staticmethod
    def _folders_to_scan(outlook: OutlookConnector, com_folder, parent: OutlookFolder,
                         cfg: ExtractionConfig) -> List[OutlookFolder]:
        if cfg.all_folders:
            return [parent] + list(outlook.get_all_subfolders(com_folder))
        leaves = list(outlook.get_leaf_folders(com_folder))
        return leaves or [parent]

    def _put_email(self, run: RunContext, item) -> bool:
        while not (run.loader_stop.is_set() or run.cancel_event.is_set()):
            while run.pause_event.is_set() and not (run.loader_stop.is_set() or run.cancel_event.is_set()):
                run.pause_event.wait(timeout=0.2)
            try:
                run.email_queue.put(item, timeout=0.5)
                return True
            except Full:
                continue
        return False

    def _post_progress(self, run: RunContext):
        self.ui_queue.put(("PROGRESS", run.processed, run.total, run.loading_done))

    def _scan_folder(self, outlook: OutlookConnector, reader: MailReader,
                     folder: OutlookFolder, famille: str, run: RunContext) -> bool:
        cfg, state = run.cfg, run.state

        effective = self._effective_date_from(folder.path, run)
        if effective:
            self._log_info(f"Incrément: {folder.name} depuis le {effective:%Y-%m-%d %H:%M}")
        else:
            self._log_info(f"Chargement des emails du dossier : {folder.name} ({folder.path})")

        new_count = skipped = errors = 0
        completed = True
        try:
            for mail in outlook.get_mails(folder.folder, date_from=effective,
                                          date_to=cfg.date_to, limit=None):
                while run.pause_event.is_set() and not (run.cancel_event.is_set() or run.loader_stop.is_set()):
                    run.pause_event.wait(timeout=0.2)
                if run.cancel_event.is_set() or run.loader_stop.is_set():
                    completed = False
                    break
                try:
                    mail_id = _safe(mail, "EntryID")
                    if (mail_id and not (cfg.full_rescan or cfg.refresh)
                            and mail_id in state.processed_mail_ids):
                        skipped += 1
                        continue

                    received = outlook.parse_date(_safe(mail, "ReceivedTime"))
                    if received:
                        if effective and received < effective:
                            continue
                        if cfg.date_to and received > cfg.date_to:
                            continue

                    email_data = reader.read(mail, folder, famille, mail_id, received)

                    if not self._put_email(run, email_data):
                        completed = False
                        break
                    new_count += 1
                    run.total += 1
                    if run.total % PROGRESS_EVERY == 0:
                        self._post_progress(run)

                    if cfg.limit and new_count >= cfg.limit:
                        completed = False
                        break
                except Exception as e:
                    errors += 1
                    self._log_warning(f"Mail ignoré dans {folder.name} : {e}")
                    continue
        except Exception as e:
            completed = False
            self._log_error(f"Erreur dossier {folder.path} : {e}\n"
                            + traceback.format_exc(limit=5))

        self._log_info(
            f"Dossier {folder.name} : {new_count} nouveaux emails "
            f"({skipped} déjà traités, {errors} en erreur)")
        return completed

    def _load_emails_to_queue(self, entry_ids: List[str], run: RunContext):
        outlook = None
        com_ready = False
        try:
            pythoncom.CoInitializeEx(pythoncom.COINIT_MULTITHREADED)
            com_ready = True
            outlook = OutlookConnector(ocr_inline_images=run.cfg.ocr_inline)
            reader = MailReader(outlook, run.cfg, self._log_warning)

            processed_paths: Set[str] = set()

            for entry_id in entry_ids:
                while run.pause_event.is_set() and not (run.cancel_event.is_set() or run.loader_stop.is_set()):
                    run.pause_event.wait(timeout=0.2)
                if run.cancel_event.is_set() or run.loader_stop.is_set():
                    break

                com_folder = outlook.get_folder_by_entry_id(entry_id)
                if not com_folder:
                    self._log_warning(f"Dossier introuvable (EntryID: {entry_id})")
                    continue

                parent = OutlookFolder(name=com_folder.Name,
                                       path=com_folder.FolderPath,
                                       folder=com_folder)
                famille = parent.root_name

                for folder in self._folders_to_scan(outlook, com_folder, parent, run.cfg):
                    while run.pause_event.is_set() and not (run.cancel_event.is_set() or run.loader_stop.is_set()):
                        run.pause_event.wait(timeout=0.2)
                    if folder.path in processed_paths:
                        continue
                    processed_paths.add(folder.path)

                    if not self._scan_folder(outlook, reader, folder, famille, run):
                        run.incomplete_folders.add(folder.path)
                    if run.cancel_event.is_set() or run.loader_stop.is_set():
                        break

            self._log_info(f"Total: {run.total} nouveaux emails chargés")

        except Exception as e:
            self._log_error(f"Erreur fatale lors du chargement: {e}\n"
                            + traceback.format_exc(limit=10))
            run.incomplete_folders.add("*")
        finally:
            run.loading_done = True
            self._post_progress(run)
            self._put_email(run, None)
            if outlook:
                try:
                    outlook.close()
                except Exception:
                    pass
            if com_ready:
                pythoncom.CoUninitialize()

    def _export(self, cfg: ExtractionConfig, contacts) -> Tuple[Optional[str], bool]:
        exporter = DataExporter(contacts)
        out_path = os.path.join(cfg.output_dir, f"{cfg.output}.xlsx")

        if os.path.exists(out_path):
            try:
                shutil.copy2(out_path, out_path + ".bak")
            except Exception as e:
                self._log_warning(f"Sauvegarde .bak impossible : {e}")

        try:
            path = exporter.export_to_excel(
                cfg.output, output_dir=cfg.output_dir,
                sheet_per_project=cfg.sheet_per_project,
                incremental=not cfg.full_rescan,
                keep_backup=cfg.refresh)
            return path, False
        except PermissionError:
            alt_name = f"{cfg.output}_{datetime.now():%Y%m%d_%H%M%S}"
            self._log_warning(
                f"'{out_path}' est verrouillé (ouvert dans Excel ?). "
                f"Export de secours : {alt_name}.xlsx")
            path = exporter.export_to_excel(
                alt_name, output_dir=cfg.output_dir,
                sheet_per_project=cfg.sheet_per_project, incremental=False)
            return path, True

    def _append_review(self, contacts, output_file: str):
        try:
            review = AccountReview(contacts)
            review.append_to_excel(output_file)
            self._log_info("Revue de compte ajoutée au fichier Excel.")
            gui_data = review.gui_data()
            for stat in gui_data["stats"]:
                self._log_info(f"  - {stat['Statistique']} : {stat['Valeur']}")
        except Exception as e:
            self._log_error(f"Erreur revue de compte : {e}\n" + traceback.format_exc(limit=8))

    def _commit_pending(self, run: RunContext,
                        pending: List[Tuple[str, str, str]]):
        state, cfg = run.state, run.cfg
        global_incomplete = "*" in run.incomplete_folders
        for mail_id, folder_path, received_iso in pending:
            partial = (cfg.limit > 0 or global_incomplete
                       or folder_path in run.incomplete_folders)
            if partial:
                if mail_id:
                    state.processed_mail_ids.add(mail_id)
            else:
                state.mark_processed(mail_id, folder_path, received_iso)
        state.save()
        self._log_info(f"Checkpoint enregistré ({len(pending)} emails).")

    def _worker_process_emails(self, run: RunContext):
        cfg = run.cfg
        pending: List[Tuple[str, str, str]] = []
        try:
            extractor = ContactExtractor(
                outlook_connector=None,
                process_attachments=not cfg.no_attachments,
                skip_pdf=cfg.no_pdf,
                skip_images=cfg.no_images,
                skip_office=cfg.no_office,
                include_recipients=cfg.include_recipients,
                exclude_domains=cfg.exclude_domains,
                state=run.state,
                full_rescan=cfg.full_rescan,
                process_body_emails=cfg.extract_body_emails,
                refresh=cfg.refresh,
            )

            while not run.cancel_event.is_set():
                while run.pause_event.is_set() and not run.cancel_event.is_set():
                    run.pause_event.wait(timeout=0.2)
                try:
                    item = run.email_queue.get(timeout=0.5)
                except Empty:
                    continue
                if item is None:
                    break

                try:
                    extractor._process_email_data(item)
                except Exception as e:
                    self._log_error(f"Erreur traitement email : {e}\n"
                                    + traceback.format_exc(limit=5))
                    continue

                received_iso = (item.received.strftime("%Y-%m-%d %H:%M:%S")
                                if item.received else "")
                pending.append((item.mail_id, item.folder_path, received_iso))
                run.processed += 1
                if run.processed % PROGRESS_EVERY == 0:
                    self._post_progress(run)
                if run.processed % LOG_EVERY == 0:
                    self._log_info(f"Traitement : {run.processed}/{run.total} emails...")

            run.loader_stop.set()
            if run.loader_thread:
                run.loader_thread.join(timeout=30)
            if run.cancel_event.is_set():
                run.incomplete_folders.add("*")
            self._post_progress(run)

            contacts = extractor.contacts
            output_file = None
            export_ok = True

            if contacts:
                try:
                    output_file, fallback = self._export(cfg, contacts)
                except Exception as e:
                    export_ok = False
                    self._log_error(f"Échec de l'export Excel : {e}\n"
                                    + traceback.format_exc(limit=8))
                    self.ui_queue.put((
                        "ERROR_BOX", "Export impossible",
                        f"L'export Excel a échoué :\n{e}\n\nLe checkpoint n'a PAS été validé."))

                if export_ok and output_file:
                    output_file = os.path.abspath(output_file)
                    self.last_generated_file = output_file
                    self._log_info(f"\nFichier généré : {output_file}")
                    self._log_info(f"Contacts extraits : {len(contacts)}")
                    self._append_review(contacts, output_file)
                    suffix = " (fichier de secours)" if fallback else ""
                    verb = "interrompue" if run.cancel_event.is_set() else "terminée"
                    self.ui_queue.put((
                        "INFO_BOX", f"Extraction {verb}",
                        f"Extraction {verb} !\n\nFichier : {output_file}{suffix}\n"
                        f"Contacts : {len(contacts)}"))
                elif export_ok:
                    self._log_warning("Aucune ligne valide à exporter.")
            else:
                self._log_warning("Aucun nouveau contact extrait.")

            if export_ok:
                self._commit_pending(run, pending)

        except Exception as e:
            self._log_error(f"Erreur dans le worker : {e}\n" + traceback.format_exc(limit=10))
        finally:
            run.loader_stop.set()
            self.ui_queue.put(("FINISHED",))

    def _process_queue(self):
        try:
            for _ in range(UI_MESSAGES_PER_TICK):
                try:
                    msg = self.ui_queue.get_nowait()
                except Empty:
                    break
                kind = msg[0]
                if kind == "LOG":
                    self._display_log(msg[1], msg[2])
                elif kind == "PROGRESS":
                    self._update_progress(msg[1], msg[2], msg[3])
                elif kind == "INFO_BOX":
                    messagebox.showinfo(msg[1], msg[2])
                elif kind == "ERROR_BOX":
                    messagebox.showerror(msg[1], msg[2])
                elif kind == "FINISHED":
                    self._reset_extraction_state()
        except Exception:
            pass
        self.root.after(100, self._process_queue)

    def _update_progress(self, processed: int, total: int, loading_done: bool):
        if total > 0:
            self.progress["value"] = min(100.0, processed / total * 100)
        suffix = "" if loading_done else " (lecture en cours...)"
        if self.run and self.run.pause_event.is_set():
            suffix = " [EN PAUSE]"
        self.progress_label.config(text=f"{processed}/{total} emails traités{suffix}")

    def _reset_extraction_state(self):
        self.extraction_in_progress = False
        self.start_button.config(state=tk.NORMAL)
        self.pause_button.config(state=tk.DISABLED, text="⏸ Mettre en pause")
        self.stop_button.config(state=tk.DISABLED)
        if self.last_generated_file and os.path.exists(self.last_generated_file):
            self.open_excel_button.config(state=tk.NORMAL)
        self.progress_frame.pack_forget()
        self._log_info("Extraction terminée.")
        self._log_info("=" * 60)

    def _log_info(self, message):
        self.ui_queue.put(("LOG", "INFO", message))

    def _log_warning(self, message):
        self.ui_queue.put(("LOG", "WARNING", message))

    def _log_error(self, message):
        self.ui_queue.put(("LOG", "ERROR", message))

    def _display_log(self, level: str, message: str):
        tag = {"ERROR": "error", "CRITICAL": "error",
               "WARNING": "warning", "DEBUG": "debug"}.get(level, "")
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, f"[{level}] {message}\n", tag)
        lines = int(self.log_text.index("end-1c").split(".")[0])
        if lines > MAX_LOG_LINES:
            self.log_text.delete("1.0", f"{lines - MAX_LOG_LINES}.0")
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)

    def _on_close(self):
        if self.extraction_in_progress:
            if not messagebox.askokcancel(
                    "Quitter",
                    "Une extraction est en cours.\nQuitter quand même ?"):
                return
            if self.run:
                self.run.pause_event.clear()
                self.run.cancel_event.set()
                self.run.loader_stop.set()
        try:
            logging.getLogger().removeHandler(self.log_handler)
        except Exception:
            pass
        try:
            if self.outlook:
                self.outlook.close()
        except Exception:
            pass
        self.root.destroy()


if __name__ == "__main__":
    root = tk.Tk()
    app = ContactExtractionGUI(root)
    if app.ready:
        root.mainloop()