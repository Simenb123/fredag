from __future__ import annotations
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
from datetime import datetime, timedelta
from typing import List

from .group_rules import GroupRule, load_rules, save_rules, default_rules_path, resolve_group
from .scheduler import install_weekly_task, delete_task
from .auto_archive import run_archive
from .outlook_core import get_session
from .retention import apply_retention

TASK_ARCHIVE = "Fredag_AutoArkiv"
TASK_RETENT  = "Fredag_Retention"

_HELP_TEMPLATE = (
    "Plassholdere i mål‑mal:\n"
    "  {year} {month2} {month_abbr} {sender} {domain} {subject_tag}\n"
    "Eksempel: {year}/{month2}_{month_abbr}/{domain}/{subject_tag}"
)

class GroupManagerWindow(tk.Toplevel):
    def __init__(self, master):
        super().__init__(master)
        self.title("Grupper / regler for arkivering")
        self.geometry("1020x780")
        self.rules: List[GroupRule] = load_rules()
        self._build()

    def _build(self):
        top = ttk.Frame(self, padding=10); top.pack(fill="both", expand=True)

        left = ttk.Frame(top); left.pack(side="left", fill="y")
        ttk.Label(left, text="Grupper").pack(anchor="w")
        self.lst = tk.Listbox(left, width=28, height=32)
        self.lst.pack(fill="y", expand=False)
        self.lst.bind("<<ListboxSelect>>", self._on_sel)

        btns = ttk.Frame(left); btns.pack(fill="x", pady=(6,0))
        ttk.Button(btns, text="Ny", command=self._new).pack(side="left")
        ttk.Button(btns, text="Slett", command=self._delete).pack(side="left", padx=(6,0))

        right = ttk.Frame(top, padding=(12,0)); right.pack(side="left", fill="both", expand=True)

        f = ttk.Frame(right); f.pack(fill="x")
        ttk.Label(f, text="Gruppenavn:").grid(row=0, column=0, sticky="w")
        self.var_name = tk.StringVar(); ttk.Entry(f, textvariable=self.var_name, width=40).grid(row=0, column=1, sticky="w")

        ttk.Label(f, text="Rotmappe (filsystem):").grid(row=1, column=0, sticky="w", pady=(6,0))
        self.var_dir = tk.StringVar(); ttk.Entry(f, textvariable=self.var_dir, width=60).grid(row=1, column=1, sticky="w", pady=(6,0))
        ttk.Button(f, text="Bla …", command=self._choose_dir).grid(row=1, column=2, padx=(6,0), pady=(6,0))

        grid = ttk.Frame(right); grid.pack(fill="x", pady=(12,4))
        ttk.Label(grid, text="Tillatte endelser (komma):").grid(row=0, column=0, sticky="w")
        self.var_exts = tk.StringVar(); ttk.Entry(grid, textvariable=self.var_exts, width=40).grid(row=0, column=1, sticky="w")
        ttk.Label(grid, text="Min KB:").grid(row=0, column=2, sticky="e", padx=(12,0))
        self.var_min = tk.StringVar(value="0"); ttk.Entry(grid, textvariable=self.var_min, width=6).grid(row=0, column=3, sticky="w")
        ttk.Label(grid, text="Max KB:").grid(row=0, column=4, sticky="e")
        self.var_max = tk.StringVar(value="0"); ttk.Entry(grid, textvariable=self.var_max, width=6).grid(row=0, column=5, sticky="w")

        ttk.Label(grid, text="Kategori (Outlook):").grid(row=1, column=0, sticky="w", pady=(6,0))
        self.var_cat = tk.StringVar(); ttk.Entry(grid, textvariable=self.var_cat, width=30).grid(row=1, column=1, sticky="w", pady=(6,0))
        ttk.Label(grid, text="Farge:").grid(row=1, column=2, sticky="e", padx=(12,0))
        self.var_cat_color = tk.StringVar(value="")
        ttk.Combobox(grid, values=["","red","orange","yellow","green","teal","blue","purple","maroon","gray","black"],
                     textvariable=self.var_cat_color, width=12, state="readonly").grid(row=1, column=3, sticky="w")

        ttk.Label(grid, text="Retention (dager, 0=behold):").grid(row=1, column=4, sticky="e")
        self.var_ret = tk.StringVar(value="0"); ttk.Entry(grid, textvariable=self.var_ret, width=8).grid(row=1, column=5, sticky="w")

        tmp = ttk.Frame(right); tmp.pack(fill="x", pady=(8,0))
        ttk.Label(tmp, text="Mål‑mal (relativ til rot):").grid(row=0, column=0, sticky="w")
        self.var_tpl = tk.StringVar()
        ttk.Entry(tmp, textvariable=self.var_tpl, width=64).grid(row=0, column=1, sticky="w")
        ttk.Button(tmp, text="Standard", command=lambda: self.var_tpl.set("")).grid(row=0, column=2, padx=(6,0))

        ttk.Label(tmp, text="Emne‑tag (regex):").grid(row=1, column=0, sticky="w", pady=(6,0))
        self.var_rx = tk.StringVar()
        ttk.Entry(tmp, textvariable=self.var_rx, width=64).grid(row=1, column=1, sticky="w", pady=(6,0))
        ttk.Label(tmp, text=_HELP_TEMPLATE).grid(row=2, column=0, columnspan=3, sticky="w", pady=(6,0))

        # --- Flytt-regler
        mv = ttk.LabelFrame(right, text="Flytt‑regler (Outlook)")
        mv.pack(fill="x", pady=(10,4))
        ttk.Label(mv, text="Flytt til mappe (sti):").grid(row=0, column=0, sticky="w")
        self.var_move = tk.StringVar()
        ttk.Entry(mv, textvariable=self.var_move, width=70).grid(row=0, column=1, sticky="w")
        ttk.Label(mv, text=r"Eksempel: \\Mailbox - Ola Nordmann\Arkiv\KundeX").grid(row=1, column=1, sticky="w", pady=(2,0))
        self.var_move_read = tk.BooleanVar(value=False)
        ttk.Checkbutton(mv, text="Marker som lest etter flytt", variable=self.var_move_read).grid(row=0, column=2, sticky="w", padx=(10,0))

        ttk.Label(right, text="Sendere (en per linje) – epost, @domene.no eller wildcard som *firma.no").pack(anchor="w", pady=(8,2))
        self.txt_senders = tk.Text(right, height=12); self.txt_senders.pack(fill="both", expand=True)

        act = ttk.Frame(right); act.pack(fill="x", pady=(6,0))
        ttk.Button(act, text="Lagre", command=self._save).pack(side="left")
        ttk.Button(act, text="Test match …", command=self._test).pack(side="left", padx=(6,0))
        ttk.Button(act, text="Tørrkjør (siste 7 dager)", command=lambda: self._archive_or_dry(dry=True)).pack(side="right")
        ttk.Button(act, text="Arkiver nå (siste 7 dager)", command=lambda: self._archive_or_dry(dry=False)).pack(side="right", padx=(0,6))

        plan = ttk.Frame(right); plan.pack(fill="x", pady=(10,0))
        ttk.Button(plan, text="Kjør retention nå", command=self._retention_now).pack(side="left")
        ttk.Button(plan, text="Planlegg Arkiv (fredag …)", command=lambda: self._schedule_dialog(kind='archive')).pack(side="right")
        ttk.Button(plan, text="Planlegg Retention (fredag …)", command=lambda: self._schedule_dialog(kind='retention')).pack(side="right", padx=(0,6))

        cfg = ttk.Frame(self, padding=(10,6)); cfg.pack(fill="x", pady=(4,8))
        from .config_io import export_config, import_config
        ttk.Button(cfg, text="Eksporter konfig …", command=lambda: self._export_cfg(export_config)).pack(side="left")
        ttk.Button(cfg, text="Importer konfig …", command=lambda: self._import_cfg(import_config)).pack(side="left", padx=(6,0))
        ttk.Button(cfg, text="Innstillinger …", command=self._open_settings).pack(side="left", padx=(6,0))

        self.status = ttk.Label(self, text=f"Lagringsfil: {default_rules_path()}")
        self.status.pack(fill="x", pady=(0,4))
        self._reload_list()

    def _open_settings(self):
        from .settings_window import SettingsWindow
        SettingsWindow(self)

    def _reload_list(self):
        self.lst.delete(0, "end")
        for r in self.rules:
            self.lst.insert("end", r.name)

    def _on_sel(self, _):
        i = self._sel_index();
        if i is None: return
        r = self.rules[i]
        self.var_name.set(r.name); self.var_dir.set(r.target_dir)
        self.var_exts.set(",".join(r.allowed_exts or []))
        self.var_min.set(str(r.min_kb or 0)); self.var_max.set(str(r.max_kb or 0))
        self.var_cat.set(r.category or ""); self.var_cat_color.set(r.category_color or "")
        self.var_ret.set(str(r.retention_days or 0))
        self.var_tpl.set(r.target_template or ""); self.var_rx.set(r.subject_tag_regex or "")
        self.var_move.set(r.move_to_folder_path or ""); self.var_move_read.set(bool(r.move_mark_read))
        self.txt_senders.delete("1.0","end"); self.txt_senders.insert("1.0", "\n".join(r.senders))

    def _sel_index(self):
        sel = self.lst.curselection()
        return sel[0] if sel else None

    def _new(self):
        self.rules.append(GroupRule(name="Ny gruppe", target_dir=str(Path.home()/ "Arkiv"), senders=[]))
        self._reload_list(); self.lst.selection_clear(0, "end"); self.lst.selection_set("end"); self.lst.event_generate("<<ListboxSelect>>")

    def _delete(self):
        i = self._sel_index()
        if i is None: return
        if not messagebox.askyesno("Slett", f"Slette gruppen «{self.rules[i].name}»?"): return
        self.rules.pop(i); self._reload_list(); save_rules(self.rules)

    def _choose_dir(self):
        p = filedialog.askdirectory(title="Velg rotmappe for gruppen")
        if p: self.var_dir.set(p)

    def _save(self):
        i = self._sel_index()
        if i is None: messagebox.showinfo("Lagre", "Velg en gruppe først."); return
        name = self.var_name.get().strip()
        if not name: messagebox.showerror("Lagre","Gruppen må ha navn."); return
        target = self.var_dir.get().strip()
        if not target: messagebox.showerror("Lagre","Gruppen må ha rotmappe."); return
        senders = [ln.strip().lower() for ln in self.txt_senders.get("1.0","end").splitlines() if ln.strip()]
        exts = [e.strip().lower().lstrip(".") for e in (self.var_exts.get() or "").split(",") if e.strip()]
        try:
            min_kb = int(self.var_min.get() or "0")
            max_kb = int(self.var_max.get() or "0")
            ret    = int(self.var_ret.get() or "0")
        except ValueError:
            messagebox.showerror("Lagre","Min/Max KB og Retention må være tall."); return
        self.rules[i] = GroupRule(
            name=name, target_dir=target, senders=senders,
            allowed_exts=exts, min_kb=min_kb, max_kb=max_kb,
            category=self.var_cat.get().strip(), category_color=self.var_cat_color.get().strip(),
            retention_days=ret, target_template=self.var_tpl.get().strip(),
            subject_tag_regex=self.var_rx.get().strip(),
            move_to_folder_path=self.var_move.get().strip(),
            move_mark_read=bool(self.var_move_read.get())
        )
        save_rules(self.rules); self._reload_list(); messagebox.showinfo("Lagre", "Gruppen er lagret.")

    def _test(self):
        if not self.rules:
            messagebox.showinfo("Test", "Ingen grupper definert."); return
        w = tk.Toplevel(self); w.title("Test avsender → gruppe"); w.geometry("420x160")
        frm = ttk.Frame(w, padding=10); frm.pack(fill="both", expand=True)
        ttk.Label(frm, text="Avsenders epost:").grid(row=0, column=0, sticky="w")
        ve = tk.StringVar(); ttk.Entry(frm, textvariable=ve, width=40).grid(row=0, column=1, sticky="w")
        ttk.Label(frm, text="Navn (valgfritt):").grid(row=1, column=0, sticky="w", pady=(6,0))
        vn = tk.StringVar(); ttk.Entry(frm, textvariable=vn, width=40).grid(row=1, column=1, sticky="w", pady=(6,0))
        out = ttk.Label(frm, text=""); out.grid(row=2, column=0, columnspan=2, sticky="w", pady=(10,0))
        def run():
            r = resolve_group(self.rules, ve.get().strip().lower(), vn.get().strip())
            out.config(text=f"→ Gruppe: {r.name if r else '(ingen)'}")
        ttk.Button(frm, text="Test", command=run).grid(row=3, column=1, sticky="e", pady=(10,0))

    def _archive_or_dry(self, dry: bool):
        try:
            session = get_session()
            f = (datetime.now() - timedelta(days=7)).date()
            t = datetime.now().date()
            summary, unassigned = run_archive(session, f, t, include_subfolders=True, only_attachments=True, dry_run=dry)
            msg = ["(TØRRKJØRING)" if dry else "Arkivering fullført:", ""]
            for g, s in summary.items():
                msg.append(f"• {g}: {'ville lagret' if dry else 'lagret'} {s['saved']}, hoppet {s['skipped']} (meldinger i gruppe: {s['msgs']})")
            if unassigned:
                msg.append(f"\nUten gruppe: {len(unassigned)} meldinger (ikke berørt)")
            messagebox.showinfo("Arkiv", "\n".join(msg))
        except Exception as e:
            messagebox.showerror("Arkiv", str(e))

    def _retention_now(self):
        try:
            summary = apply_retention(self.rules, dry_run=False)
            lines = ["Retention kjørt:\n"]
            for g, s in summary.items():
                lines.append(f"• {g}: slettet {s['deleted']}, beholdt {s['kept']}, feil {s['errors']}")
            if len(lines) == 1: lines.append("(Ingen grupper har retention definert.)")
            messagebox.showinfo("Retention", "\n".join(lines))
        except Exception as e:
            messagebox.showerror("Retention", str(e))

    def _schedule_dialog(self, kind: str):
        w = tk.Toplevel(self); w.title(f"Planlegg ukentlig ({'arkiv' if kind=='archive' else 'retention'})"); w.geometry("360x180")
        frm = ttk.Frame(w, padding=10); frm.pack(fill="both", expand=True)

        ttk.Label(frm, text="Tid (HH:MM):").grid(row=0, column=0, sticky="w")
        vt = tk.StringVar(value="16:10" if kind=="retention" else "16:00")
        ttk.Entry(frm, textvariable=vt, width=8).grid(row=0, column=1, sticky="w")

        info = ttk.Label(frm, text="Oppgaven kjører Python‑modul i din brukerkontekst.")
        info.grid(row=1, column=0, columnspan=2, sticky="w", pady=(10,0))

        def do_install():
            from pathlib import Path
            if kind == "archive":
                from . import auto_archive as mod
                script = Path(mod.__file__).resolve()
                args = '--from-days 7 --only-attachments --mail-report'
                ok, out = install_weekly_task(TASK_ARCHIVE, vt.get().strip(), script, args=args, day="FRI")
            else:
                from . import retention_job as mod
                script = Path(mod.__file__).resolve()
                args = '--mail-report'
                ok, out = install_weekly_task(TASK_RETENT, vt.get().strip(), script, args=args, day="FRI")
            messagebox.showinfo("Planlegg", ("OK: " if ok else "FEIL: ") + (out or ""))

        def do_remove():
            ok, out = delete_task(TASK_ARCHIVE if kind == "archive" else TASK_RETENT)
            messagebox.showinfo("Planlegg", ("OK: " if ok else "FEIL: ") + (out or ""))

        btns = ttk.Frame(frm); btns.grid(row=2, column=0, columnspan=2, pady=(12,0), sticky="e")
        ttk.Button(btns, text="Opprett/oppdater", command=do_install).pack(side="right")
        ttk.Button(btns, text="Fjern", command=do_remove).pack(side="right", padx=(0,8))

    def _export_cfg(self, fn):
        path = filedialog.asksaveasfilename(defaultextension=".zip", filetypes=[("ZIP","*.zip")],
                                            initialfile="fredag_konfig.zip", title="Eksporter konfig")
        if not path: return
        ok, msg = fn(path); messagebox.showinfo("Eksporter konfig", msg)

    def _import_cfg(self, fn):
        path = filedialog.askopenfilename(defaultextension=".zip", filetypes=[("ZIP","*.zip")],
                                          title="Importer konfig (ZIP)")
        if not path: return
        ok, msg = fn(path, backup_current=True)
        if ok:
            self.rules = load_rules(); self._reload_list()
        messagebox.showinfo("Importer konfig", msg)
