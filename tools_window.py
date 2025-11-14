from __future__ import annotations
import os
import csv
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from datetime import datetime, timedelta
from typing import Dict, List

from .outlook_core import save_attachments, search_messages, mail_as_text
from .widgets_datepicker import DatePicker
from .attachments_window import AttachmentsWindow
from .excel_export import export_messages_to_xlsx
from .archiver import archive_messages

from .group_window import GroupManagerWindow
from .group_archiver import archive_by_groups
from .group_rules import load_rules, resolve_group
from .diag_window import DiagnoseWindow
from .log_utils import log_path
from .suggest_window import SuggestGroupsWindow
from .group_mover import move_by_groups  # NYTT

class OutlookToolsWindow(tk.Toplevel):
    def __init__(self, master, session):
        super().__init__(master)
        self.title("Outlook-verktøy – søk/les/vedlegg")
        self.geometry("1120x720")
        self.session = session

        self._stop_evt = threading.Event()
        self._results: List[Dict] = []
        self._id_to_ix: Dict[str, int] = {}
        self._row_iids: List[str] = []
        self._save_per_sender = tk.BooleanVar(value=False)

        self._build_ui()

    def _build_ui(self):
        frm = ttk.Frame(self, padding=10); frm.pack(fill="both", expand=True)

        r1 = ttk.Frame(frm); r1.pack(fill="x")
        ttk.Label(r1, text="Avsender (navn/epost):").grid(row=0, column=0, sticky="w")
        self.sender_var = tk.StringVar(); ttk.Entry(r1, textvariable=self.sender_var, width=40)\
            .grid(row=0, column=1, padx=(6,10))

        ttk.Label(r1, text="Fra (DD.MM.YYYY):").grid(row=0, column=2, sticky="w")
        monday = (datetime.now() - timedelta(days=datetime.now().weekday())).strftime("%d.%m.%Y")
        self.from_var = tk.StringVar(value=monday)
        ttk.Entry(r1, textvariable=self.from_var, width=12).grid(row=0, column=3, padx=(6,2))
        ttk.Button(r1, text="📅", width=3, command=lambda: DatePicker(self, self.from_var))\
            .grid(row=0, column=4, padx=(0,8))

        ttk.Label(r1, text="Til (DD.MM.YYYY):").grid(row=0, column=5, sticky="w")
        self.to_var = tk.StringVar(value=datetime.now().strftime("%d.%m.%Y"))
        ttk.Entry(r1, textvariable=self.to_var, width=12).grid(row=0, column=6, padx=(6,2))
        ttk.Button(r1, text="📅", width=3, command=lambda: DatePicker(self, self.to_var))\
            .grid(row=0, column=7, padx=(0,8))

        self.subfolders_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(r1, text="Inkluder undermapper", variable=self.subfolders_var)\
            .grid(row=0, column=8)
        ttk.Button(r1, text="Søk", command=self.on_search).grid(row=0, column=9, padx=(10,4))
        ttk.Button(r1, text="Avbryt", command=self.on_cancel).grid(row=0, column=10)

        r2 = ttk.Frame(frm); r2.pack(fill="x", pady=(6,4))
        for txt, fn in [("I dag", self.quick_today), ("Denne uken", self.quick_thisweek),
                        ("Siste 7", self.quick_7), ("Siste 30", self.quick_30)]:
            ttk.Button(r2, text=txt, command=fn).pack(side="left")
        ttk.Label(r2, text="   Emne inneholder:").pack(side="left", padx=(16,2))
        self.subj_var = tk.StringVar(); ttk.Entry(r2, textvariable=self.subj_var, width=30).pack(side="left")
        self.unread_var = tk.BooleanVar(value=False); self.attach_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(r2, text="Bare uleste", variable=self.unread_var).pack(side="left", padx=(16,0))
        ttk.Checkbutton(r2, text="Kun med vedlegg", variable=self.attach_var).pack(side="left")

        mid = ttk.Frame(frm); mid.pack(fill="both", expand=True, pady=(6,4))
        cols = ("dato","emne","fra","vedlegg","ulest")
        self.tree = ttk.Treeview(mid, columns=cols, show="headings", height=16)
        for c,t,w,a in [("dato","Dato",165,"w"),("emne","Emne",560,"w"),("fra","Fra",260,"w"),
                        ("vedlegg","#Vedlegg",80,"e"),("ulest","Ulest",60,"center")]:
            self.tree.heading(c, text=t); self.tree.column(c, width=w, anchor=a)
        self.tree.pack(side="left", fill="both", expand=True)
        sb = ttk.Scrollbar(mid, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscroll=sb.set); sb.pack(side="right", fill="y")
        self.tree.bind("<<TreeviewSelect>>", self.on_select)
        self.tree.bind("<Double-1>", lambda e: self.open_in_outlook())
        self.tree.bind("<Button-3>", self._popup)

        bot = ttk.Frame(frm); bot.pack(fill="both", expand=True)
        self.body = tk.Text(bot, wrap="word"); self.body.pack(side="left", fill="both", expand=True)
        sb2 = ttk.Scrollbar(bot, orient="vertical", command=self.body.yview)
        self.body.configure(yscroll=sb2.set); sb2.pack(side="right", fill="y")

        b = ttk.Frame(frm); b.pack(fill="x", pady=(6,0))
        ttk.Button(b, text="Åpne i Outlook", command=self.open_in_outlook).pack(side="left")
        ttk.Button(b, text="Vedlegg…", command=self.show_attachments).pack(side="left", padx=(6,0))
        ttk.Button(b, text="Kopier tekst", command=self.copy_text).pack(side="left", padx=(8,0))
        ttk.Button(b, text="Lagre e‑post (MSG)", command=lambda: self.save_message(fmt="msg")).pack(side="left", padx=(8,0))
        ttk.Button(b, text="Lagre e‑post (HTML)", command=lambda: self.save_message(fmt="html")).pack(side="left")
        ttk.Checkbutton(b, text="Vedlegg → undermappe pr. avsender", variable=self._save_per_sender)\
            .pack(side="left", padx=(16,0))
        ttk.Button(b, text="Lagre vedlegg (valgt…)", command=self.save_selected_attachments).pack(side="right")
        ttk.Button(b, text="Lagre vedlegg (alle…)", command=self.save_all_attachments).pack(side="right", padx=(0,8))

        b2 = ttk.Frame(frm); b2.pack(fill="x", pady=(6,0))
        ttk.Button(b2, text="Marker som lest (valgte)", command=lambda: self.mark_read(selected=True)).pack(side="left")
        ttk.Button(b2, text="Marker som lest (alle)", command=lambda: self.mark_read(selected=False)).pack(side="left", padx=(8,0))
        ttk.Button(b2, text="Flytt til mappe… (valgte)", command=lambda: self.move_to_folder(selected=True)).pack(side="left", padx=(16,0))
        ttk.Button(b2, text="Flytt til mappe… (alle)", command=lambda: self.move_to_folder(selected=False)).pack(side="left")
        ttk.Button(b2, text="Tørrflytt … (valgte)", command=lambda: self.simulate_move_to_folder(selected=True)).pack(side="left", padx=(8,0))

        # NYTT: Regelstyrt flytt (grupper)
        ttk.Button(b2, text="Tørrflytt ↑ grupper", command=self.dryrun_move_via_groups).pack(side="right", padx=(0,8))
        ttk.Button(b2, text="Flytt ↑ grupper", command=self.move_via_groups).pack(side="right", padx=(0,8))

        ttk.Button(b2, text="Eksporter CSV", command=self.export_csv).pack(side="right")
        ttk.Button(b2, text="Excel (.xlsx)", command=self.export_excel).pack(side="right", padx=(0,8))
        ttk.Button(b2, text="Arkiver vedlegg…", command=self.archive_all).pack(side="right", padx=(0,8))
        ttk.Button(b2, text="Tørrkjør via grupper …", command=self.dryrun_via_groups).pack(side="right", padx=(0,8))
        ttk.Button(b2, text="Arkiver via grupper …", command=self.archive_via_groups).pack(side="right", padx=(0,8))
        ttk.Button(b2, text="Grupper/Regler …", command=lambda: GroupManagerWindow(self)).pack(side="right", padx=(0,8))
        ttk.Button(b2, text="Foreslå grupper …", command=self.suggest_groups).pack(side="right", padx=(0,8))
        ttk.Button(b2, text="Diagnose+", command=lambda: DiagnoseWindow(self)).pack(side="right", padx=(0,8))
        ttk.Button(b2, text="Loggmappe", command=lambda: os.startfile(str(log_path().parent))).pack(side="right", padx=(0,8))
        ttk.Button(b2, text="Diagnose", command=self.show_diag).pack(side="right", padx=(0,8))

        status_row = ttk.Frame(frm); status_row.pack(fill="x", pady=(8,0))
        self.status = ttk.Label(status_row, text="Klar.")
        self.status.pack(side="left", fill="x", expand=True)
        self.pg = ttk.Progressbar(status_row, mode="indeterminate", length=180)
        self.pg.pack(side="right")

        self._menu = tk.Menu(self, tearoff=False)
        for label, cmd in [("Åpne i Outlook", self.open_in_outlook),
                           ("Vedlegg…", self.show_attachments),
                           ("Kopier tekst", self.copy_text),
                           ("Lagre vedlegg (valgt…)", self.save_selected_attachments),
                           ("Marker som lest (valgte)", lambda: self.mark_read(selected=True)),
                           ("Flytt til mappe… (valgte)", lambda: self.move_to_folder(selected=True))]:
            self._menu.add_command(label=label, command=cmd)

    # ---- (resten: søk, presentasjon, operasjoner, eksport, diagnose) UENDRET fra forrige versjon ----
    # For plass: metoder on_cancel, _parse_date, on_search, _populate, _get_item, on_select,
    # _popup, open_in_outlook, show_attachments, copy_text, _choose_dir, _target_dir,
    # save_selected_attachments, save_all_attachments, save_message, _for_each, mark_read,
    # move_to_folder, simulate_move_to_folder, export_csv, export_excel, archive_all,
    # archive_via_groups, dryrun_via_groups, suggest_groups, show_diag, _async_update_attachment_counts
    # (Du kan beholde disse som i forrige fil – her kommer kun de 2 nye metodene under.)

    # --- NYTT: Flytt via grupper (real) ---
    def move_via_groups(self):
        if not self._results:
            messagebox.showinfo("Flytt via grupper", "Ingen treff i listen."); return
        rules = load_rules()
        if not rules:
            if messagebox.askyesno("Flytt via grupper", "Ingen grupper definert. Åpne oppsett nå?"):
                GroupManagerWindow(self)
            return

        def work():
            import pythoncom; pythoncom.CoInitialize()
            try:
                summary, unassigned, nodest = move_by_groups(self.session, self._results, rules=rules, dry_run=False)
            finally:
                pythoncom.CoUninitialize()
            def done():
                lines = ["Flytt via grupper:"]
                for g, s in summary.items():
                    lines.append(f"• {g}: flyttet {s['moved']}, feil {s['errors']}")
                if nodest:
                    lines.append(f"\nGrupper uten flytt‑mappe: {', '.join(nodest)}")
                if unassigned:
                    lines.append(f"(Uten gruppe: {len(unassigned)} meldinger – ikke berørt)")
                messagebox.showinfo("Flytt via grupper", "\n".join(lines))
                self.status.config(text="Flytting fullført.")
            self.after(0, done)
        threading.Thread(target=work, daemon=True).start()

    # --- NYTT: Tørrflytt via grupper ---
    def dryrun_move_via_groups(self):
        if not self._results:
            messagebox.showinfo("Tørrflytt via grupper", "Ingen treff i listen."); return
        rules = load_rules()
        if not rules:
            if messagebox.askyesno("Tørrflytt via grupper", "Ingen grupper definert. Åpne oppsett nå?"):
                GroupManagerWindow(self)
            return

        def work():
            import pythoncom; pythoncom.CoInitialize()
            try:
                summary, unassigned, nodest = move_by_groups(self.session, self._results, rules=rules, dry_run=True)
            finally:
                pythoncom.CoUninitialize()
            def done():
                lines = ["(TØRRFLYTT) via grupper:"]
                for g, s in summary.items():
                    lines.append(f"• {g}: ville flyttet {s['moved']}, feil {s['errors']}")
                if nodest:
                    lines.append(f"\nGrupper uten flytt‑mappe: {', '.join(nodest)}")
                if unassigned:
                    lines.append(f"(Uten gruppe: {len(unassigned)} meldinger – ikke berørt)")
                messagebox.showinfo("Tørrflytt via grupper", "\n".join(lines))
                self.status.config(text="Tørrflytt fullført.")
            self.after(0, done)
        threading.Thread(target=work, daemon=True).start()
