import csv
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime, timedelta
from typing import List, Dict, Optional

from .widgets_datepicker import DatePicker

_INPUTS = ("%d.%m.%Y", "%d.%m.%y", "%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y")
def _parse_date(s: str):
    s = (s or "").strip()
    for fmt in _INPUTS:
        try: return datetime.strptime(s, fmt).date()
        except ValueError: continue
    raise ValueError("Ugyldig dato. Bruk DD.MM.YYYY")

def _parse_time(s: str):
    s = (s or "").strip()
    for fmt in ("%H:%M", "%H.%M"):
        try: return datetime.strptime(s, fmt).time()
        except ValueError: continue
    raise ValueError("Ugyldig klokkeslett. Bruk HH:MM")

def _parse_dt(date_str: str, time_str: str) -> datetime:
    d = _parse_date(date_str); t = _parse_time(time_str)
    return datetime(d.year, d.month, d.day, t.hour, t.minute)

def _thread_load_events(d_from, d_to, include_subfolders: bool) -> (List[Dict], Optional[str]):
    try:
        import pythoncom, win32com.client
        pythoncom.CoInitialize()
        try:
            app = win32com.client.Dispatch("Outlook.Application")
            session = app.Session
        except Exception as e:
            return [], f"Klarte ikke å starte Outlook i tråd: {e}"

        def walk_subfolders(folder):
            yield folder
            if not include_subfolders: return
            try:
                subs = folder.Folders
                for i in range(1, subs.Count + 1):
                    sub = subs.Item(i)
                    for f in walk_subfolders(sub): yield f
            except Exception: return

        def calendar_roots():
            roots = []
            try:
                ds = session.DefaultStore
                if ds:
                    try: roots.append(ds.GetDefaultFolder(9))  # olFolderCalendar
                    except Exception: pass
                for i in range(1, session.Stores.Count + 1):
                    st = session.Stores.Item(i)
                    if ds and st.StoreID == ds.StoreID: continue
                    try: roots.append(st.GetDefaultFolder(9))
                    except Exception: continue
            except Exception:
                pass
            return roots or []

        start_dt = datetime.combine(d_from, datetime.min.time())
        end_dt   = datetime.combine(d_to,   datetime.max.time())
        start_s  = start_dt.strftime("%m/%d/%Y %I:%M %p")   # Restrict må bruke US‑format
        end_s    = end_dt.strftime("%m/%d/%Y %I:%M %p")

        results: List[Dict] = []
        for root in calendar_roots():
            for folder in walk_subfolders(root):
                try: items = folder.Items
                except Exception: continue
                try: items.IncludeRecurrences = True
                except Exception: pass
                try: items.Sort("[Start]")
                except Exception: pass

                restr = items.Restrict(f"[Start] >= '{start_s}' AND [Start] <= '{end_s}'")
                try: cnt = restr.Count
                except Exception: cnt = 0
                for i in range(1, cnt + 1):
                    try:
                        it = restr.Item(i)
                        if getattr(it, "Class", None) != 26:  # olAppointment
                            continue
                        results.append({
                            "eid": getattr(it, "EntryID", None),
                            "store": getattr(folder, "StoreID", None),
                            "start": getattr(it, "Start", None),
                            "end": getattr(it, "End", None),
                            "subject": getattr(it, "Subject", "") or "",
                            "location": getattr(it, "Location", "") or "",
                        })
                    except Exception:
                        continue
        return results, None
    except Exception as e:
        return [], f"Uventet feil i tråd: {e}"
    finally:
        try:
            import pythoncom; pythoncom.CoUninitialize()
        except Exception: pass

class CalendarWindow(tk.Toplevel):
    """Min kalender: vis avtaler, lag ny avtale/møte (Teams), eksporter CSV."""
    def __init__(self, master, session):
        super().__init__(master)
        self.title("Min kalender (Outlook)")
        self.geometry("980x640")
        self.session = session
        self._build_ui()

    def _build_ui(self):
        frm = ttk.Frame(self, padding=10); frm.pack(fill="both", expand=True)

        flt = ttk.Frame(frm); flt.pack(fill="x")
        ttk.Label(flt, text="Fra (DD.MM.YYYY):").grid(row=0, column=0, sticky="w")
        self.from_var = tk.StringVar(value=datetime.now().strftime("%d.%m.%Y"))
        ttk.Entry(flt, textvariable=self.from_var, width=12).grid(row=0, column=1, padx=(6,2))
        ttk.Button(flt, text="📅", width=3, command=lambda: DatePicker(self, self.from_var)).grid(row=0, column=2, padx=(0,8))

        ttk.Label(flt, text="Til (DD.MM.YYYY):").grid(row=0, column=3, sticky="w")
        self.to_var = tk.StringVar(value=(datetime.now() + timedelta(days=14)).strftime("%d.%m.%Y"))
        ttk.Entry(flt, textvariable=self.to_var, width=12).grid(row=0, column=4, padx=(6,2))
        ttk.Button(flt, text="📅", width=3, command=lambda: DatePicker(self, self.to_var)).grid(row=0, column=5, padx=(0,8))

        self.subfolders_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(flt, text="Inkluder undermapper", variable=self.subfolders_var).grid(row=0, column=6)
        ttk.Button(flt, text="Oppdater", command=self.refresh).grid(row=0, column=7, padx=6)
        ttk.Button(flt, text="Eksporter CSV", command=self.export_csv).grid(row=0, column=8, padx=6)

        mid = ttk.Frame(frm); mid.pack(fill="both", expand=True, pady=(8, 6))
        cols = ("start","slutt","emne","sted")
        self.tree = ttk.Treeview(mid, columns=cols, show="headings", height=14)
        for c, title, w, anchor in [("start","Start",160,"w"),("slutt","Slutt",160,"w"),
                                    ("emne","Emne",480,"w"),("sted","Sted",150,"w")]:
            self.tree.heading(c, text=title); self.tree.column(c, width=w, anchor=anchor)
        self.tree.pack(side="left", fill="both", expand=True)
        sb = ttk.Scrollbar(mid, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscroll=sb.set); sb.pack(side="right", fill="y")
        self.tree.bind("<<TreeviewSelect>>", self._on_select)

        bottom = ttk.Frame(frm); bottom.pack(fill="both", expand=True)
        self.details = tk.Text(bottom, wrap="word", height=6); self.details.pack(side="left", fill="both", expand=True)
        sb2 = ttk.Scrollbar(bottom, orient="vertical", command=self.details.yview)
        self.details.configure(yscroll=sb2.set); sb2.pack(side="right", fill="y")

        newf = ttk.LabelFrame(frm, text="Ny avtale/møte"); newf.pack(fill="x", pady=(8,0))
        ttk.Label(newf, text="Start:").grid(row=0, column=0, sticky="w")
        self.new_from_date = tk.StringVar(value=datetime.now().strftime("%d.%m.%Y"))
        self.new_from_time = tk.StringVar(value=(datetime.now().replace(second=0, microsecond=0) + timedelta(hours=1)).strftime("%H:%M"))
        ttk.Entry(newf, textvariable=self.new_from_date, width=12).grid(row=0, column=1, padx=(6,2))
        ttk.Button(newf, text="📅", width=3, command=lambda: DatePicker(self, self.new_from_date)).grid(row=0, column=2, padx=(0,8))
        ttk.Entry(newf, textvariable=self.new_from_time, width=6).grid(row=0, column=3, padx=(0,12))

        ttk.Label(newf, text="Slutt:").grid(row=0, column=4, sticky="w")
        self.new_to_date = tk.StringVar(value=datetime.now().strftime("%d.%m.%Y"))
        self.new_to_time = tk.StringVar(value=(datetime.now().replace(second=0, microsecond=0) + timedelta(hours=2)).strftime("%H:%M"))
        ttk.Entry(newf, textvariable=self.new_to_date, width=12).grid(row=0, column=5, padx=(6,2))
        ttk.Button(newf, text="📅", width=3, command=lambda: DatePicker(self, self.new_to_date)).grid(row=0, column=6, padx=(0,8))
        ttk.Entry(newf, textvariable=self.new_to_time, width=6).grid(row=0, column=7, padx=(0,12))

        ttk.Label(newf, text="Emne:").grid(row=1, column=0, sticky="w", pady=(6,0))
        self.new_subject = tk.StringVar(value="Ny avtale")
        ttk.Entry(newf, textvariable=self.new_subject, width=60).grid(row=1, column=1, columnspan=4, sticky="w", padx=(6,12), pady=(6,0))

        ttk.Label(newf, text="Sted:").grid(row=1, column=5, sticky="w", pady=(6,0))
        self.new_location = tk.StringVar(value="")
        ttk.Entry(newf, textvariable=self.new_location, width=30).grid(row=1, column=6, columnspan=2, sticky="w", padx=(6,12), pady=(6,0))

        ttk.Button(newf, text="Kopier tider fra valgt", command=self.copy_from_selected).grid(row=0, column=8, padx=6)
        ttk.Button(newf, text="Lag avtale", command=lambda: self.make_new(invite=False, teams=False)).grid(row=1, column=8, padx=6, pady=(6,0))
        ttk.Button(newf, text="Lag møte (Teams)", command=lambda: self.make_new(invite=True, teams=True)).grid(row=1, column=9, padx=6, pady=(6,0))

        self.status = ttk.Label(frm, text="Klar."); self.status.pack(fill="x", pady=(8, 0))
        self.results: List[Dict] = []; self._id_to_ix: Dict[str,int] = {}
        self.refresh()

    # ---- datahenting ----
    def refresh(self):
        try:
            d_from = _parse_date(self.from_var.get()); d_to = _parse_date(self.to_var.get())
        except Exception:
            messagebox.showerror("Dato", "Ugyldig dato. Bruk DD.MM.YYYY."); return
        if d_to < d_from:
            messagebox.showerror("Interval", "Til‑dato kan ikke være før Fra‑dato."); return

        self.tree.delete(*self.tree.get_children()); self.details.delete("1.0","end")
        self.status.config(text="Henter avtaler …")

        include_sub = self.subfolders_var.get()
        import threading
        def _run():
            res, err = _thread_load_events(d_from, d_to, include_sub)
            def _done():
                if err: self.status.config(text=f"Feil under henting: {err}")
                self._populate(res)
            self.after(0, _done)
        threading.Thread(target=_run, daemon=True).start()

    def _populate(self, res: List[Dict]):
        self.results = res; self._id_to_ix.clear()
        for r in res:
            start = r["start"].strftime("%d.%m.%Y %H:%M") if r["start"] else ""
            end   = r["end"].strftime("%d.%m.%Y %H:%M") if r["end"] else ""
            row_id = self.tree.insert("", "end", values=(start, end, r["subject"], r["location"]))
            self._id_to_ix[row_id] = len(self._id_to_ix)
        self.status.config(text=f"{len(res)} avtaler i valgt intervall.")

    def _get_item(self, r: Dict):
        try: return self.session.GetItemFromID(r.get("eid"), r.get("store"))
        except Exception: return None

    def _on_select(self, _evt):
        sel = self.tree.selection()
        if not sel: return
        r = self.results[self._id_to_ix[sel[0]]]
        it = self._get_item(r)
        self.details.delete("1.0","end")
        if not it:
            self.details.insert("1.0", "(Klarte ikke å åpne avtalen.)"); return
        lines = [
            f"Emne: {getattr(it,'Subject','')}",
            f"Start: {getattr(it,'Start','')}",
            f"Slutt: {getattr(it,'End','')}",
            f"Sted: {getattr(it,'Location','')}",
            f"Organizer: {getattr(it,'Organizer','')}",
            "",
            getattr(it,"Body","") or ""
        ]
        self.details.insert("1.0", "\n".join(lines))

    def open_in_outlook(self):
        sel = self.tree.selection()
        if not sel: return
        r = self.results[self._id_to_ix[sel[0]]]
        it = self._get_item(r)
        if not it:
            messagebox.showerror("Outlook", "Klarte ikke å åpne avtalen."); return
        try: it.Display()
        except Exception:
            messagebox.showerror("Outlook", "Klarte ikke å vise avtalen i Outlook.")

    # ---- opprettelse / eksport ----
    def copy_from_selected(self):
        sel = self.tree.selection()
        if not sel: return
        r = self.results[self._id_to_ix[sel[0]]]
        if r.get("start"):
            self.new_from_date.set(r["start"].strftime("%d.%m.%Y")); self.new_from_time.set(r["start"].strftime("%H:%M"))
        if r.get("end"):
            self.new_to_date.set(r["end"].strftime("%d.%m.%Y")); self.new_to_time.set(r["end"].strftime("%H:%M"))
        if r.get("subject"): self.new_subject.set(r["subject"])
        if r.get("location"): self.new_location.set(r["location"])

    def make_new(self, invite: bool, teams: bool):
        try:
            start = _parse_dt(self.new_from_date.get(), self.new_from_time.get())
            end   = _parse_dt(self.new_to_date.get(),   self.new_to_time.get())
        except Exception as e:
            messagebox.showerror("Tid", f"Ugyldig start/slutt: {e}"); return
        if end <= start:
            messagebox.showerror("Tid", "Slutt må være etter start."); return
        subject = (self.new_subject.get() or "").strip() or "Ny avtale"
        location = self.new_location.get() or ""
        try:
            import win32com.client
            app = win32com.client.Dispatch("Outlook.Application")
            apt = app.CreateItem(1)
            apt.Subject = subject; apt.Location = location
            apt.Start, apt.End = start, end
            apt.ReminderMinutesBeforeStart = 15
            if invite: apt.MeetingStatus = 1
            if teams:
                try:
                    apt.IsOnlineMeeting = True
                    apt.OnlineMeetingProvider = 7
                    if not apt.Location: apt.Location = "Microsoft Teams Møte"
                except Exception:
                    apt.Body = "⚠️ Kunne ikke sette Teams‑lenke automatisk – kontroller i Outlook.\n\n" + (apt.Body or "")
            apt.Display()
        except Exception:
            messagebox.showerror("Outlook", "Klarte ikke å opprette avtale/møte.")

    def export_csv(self):
        if not self.results:
            messagebox.showinfo("Eksport", "Ingen avtaler å eksportere."); return
        path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV", "*.csv")],
                                            initialfile="kalender.csv", title="Lagre CSV")
        if not path: return
        try:
            with open(path, "w", newline="", encoding="utf-8") as f:
                w = csv.writer(f, delimiter=";")
                w.writerow(["Start","Slutt","Emne","Sted"])
                for r in self.results:
                    start = r["start"].strftime("%Y-%m-%d %H:%M") if r["start"] else ""
                    end   = r["end"].strftime("%Y-%m-%d %H:%M") if r["end"] else ""
                    w.writerow([start, end, r["subject"], r["location"]])
            messagebox.showinfo("Eksport", f"Eksportert {len(self.results)} rader til:\n{path}")
        except Exception as e:
            messagebox.showerror("Eksport", f"Feil ved skriving av CSV: {e}")
