import os, re, html, threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from datetime import datetime, timedelta

# ---------- Outlook helpers (samme stil som i din app) ----------
def _try_import_outlook():
    try:
        import win32com.client
        return win32com.client
    except Exception:
        return None

win32 = _try_import_outlook()

def _get_outlook():
    if not win32:
        return None
    try:
        return win32.Dispatch("Outlook.Application")
    except Exception:
        return None

def _get_session(app):
    try:
        return app.Session
    except Exception:
        return None

PR_SENDER_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x5D01001E"
PR_SENDER_NAME         = "http://schemas.microsoft.com/mapi/proptag/0x0C1A001E"
PR_HEADERS             = "http://schemas.microsoft.com/mapi/proptag/0x007D001E"
FROM_REGEX = re.compile(r"^From:\s*(?P<disp>.*?)\s*<(?P<smtp>[^>]+)>", re.IGNORECASE | re.MULTILINE)

def _normalize_sender(mail):
    name, smtp = "", ""
    try:
        pa = mail.PropertyAccessor
    except Exception:
        pa = None
    if pa:
        try:
            v = pa.GetProperty(PR_SENDER_SMTP_ADDRESS)
            if v and "@" in v:
                smtp = v.strip().lower()
        except Exception:
            pass
    if not smtp:
        try:
            ae = mail.Sender
            if ae is not None:
                exu = ae.GetExchangeUser()
                if exu and exu.PrimarySmtpAddress:
                    smtp = exu.PrimarySmtpAddress.strip().lower()
                    if not name:
                        name = (exu.Name or "").strip()
        except Exception:
            pass
    if not smtp:
        try:
            ae = mail.Sender
            if ae is not None:
                smtp2 = ae.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x39FE001E")
                if smtp2 and "@" in smtp2:
                    smtp = smtp2.strip().lower()
        except Exception:
            pass
    if not smtp and pa:
        try:
            headers = pa.GetProperty(PR_HEADERS)
            if headers:
                m = FROM_REGEX.search(headers)
                if m:
                    smtp = m.group("smtp").strip().lower()
                    if not name:
                        name = (m.group("disp") or "").strip().strip('"')
        except Exception:
            pass
    if not smtp:
        try:
            raw = getattr(mail, "SenderEmailAddress", "") or ""
            if raw and "@" in raw:
                smtp = raw.strip().lower()
        except Exception:
            pass
    if not name and pa:
        try:
            n = pa.GetProperty(PR_SENDER_NAME)
            if n:
                name = n.strip()
        except Exception:
            pass
    if not name:
        try:
            name = (getattr(mail, "SenderName", "") or "").strip()
        except Exception:
            pass
    return name, smtp

def _msg_time(it):
    for attr in ("ReceivedTime", "CreationTime", "SentOn"):
        try:
            dt = getattr(it, attr, None)
            if dt:
                return dt
        except Exception:
            continue
    return None

def _walk_subfolders(folder, include_subfolders=True):
    yield folder
    if not include_subfolders:
        return
    try:
        subs = folder.Folders
        for i in range(1, subs.Count + 1):
            sub = subs.Item(i)
            for f in _walk_subfolders(sub, include_subfolders):
                yield f
    except Exception:
        return

# ---------- Søking ----------
def _find_messages(session, sender_query, since_date, include_subfolders=True, max_per_folder=5000, cap_total=2000):
    """Returnerer liste av dicts: {'item':ComObject,'dt':datetime,'from':'', 'from_email':'', 'subject':'', 'folder':'', 'attach':int}"""
    results = []
    try:
        inbox = session.GetDefaultFolder(6)  # olFolderInbox
    except Exception:
        return results

    q = (sender_query or "").strip().lower()
    def _matches(name, smtp):
        return (q in (smtp or "").lower()) or (q in (name or "").lower())

    for folder in _walk_subfolders(inbox, include_subfolders):
        try:
            items = folder.Items
        except Exception:
            continue
        try:
            items.Sort("[ReceivedTime]", True)
        except Exception:
            pass

        # Foretrukket iterasjon
        processed = 0
        try:
            it = items.GetFirst()
            seq_ok = True
        except Exception:
            it = None
            seq_ok = False

        if seq_ok:
            while it and processed < max_per_folder and len(results) < cap_total:
                try:
                    if getattr(it, "Class", None) == 43:
                        dt = _msg_time(it)
                        if dt and (since_date is None or dt.date() >= since_date):
                            name, smtp = _normalize_sender(it)
                            if not q or _matches(name, smtp):
                                sub = getattr(it, "Subject", "") or ""
                                fpath = getattr(folder, "FolderPath", "")
                                n_att = 0
                                try:
                                    n_att = getattr(it.Attachments, "Count", 0)
                                except Exception:
                                    n_att = 0
                                results.append({
                                    "item": it, "dt": dt, "from": name, "from_email": smtp,
                                    "subject": sub, "folder": fpath, "attach": n_att
                                })
                    processed += 1
                    it = items.GetNext()
                except Exception:
                    break
        else:
            total = getattr(items, "Count", 0)
            upto = min(max_per_folder, total)
            for idx in range(1, upto + 1):
                if len(results) >= cap_total:
                    break
                try:
                    it = items.Item(idx)
                    if getattr(it, "Class", None) != 43:
                        continue
                    dt = _msg_time(it)
                    if dt and (since_date is None or dt.date() >= since_date):
                        name, smtp = _normalize_sender(it)
                        if not q or _matches(name, smtp):
                            sub = getattr(it, "Subject", "") or ""
                            fpath = getattr(folder, "FolderPath", "")
                            n_att = 0
                            try:
                                n_att = getattr(it.Attachments, "Count", 0)
                            except Exception:
                                n_att = 0
                            results.append({
                                "item": it, "dt": dt, "from": name, "from_email": smtp,
                                "subject": sub, "folder": fpath, "attach": n_att
                            })
                except Exception:
                    continue
    return results

# ---------- Vedlegg ----------
_ILLEGAL = r'<>:"/\\|?*'
def _sanitize_filename(name):
    return "".join(('_' if c in _ILLEGAL else c) for c in name)

def _unique_path(base_dir, filename):
    base, ext = os.path.splitext(filename)
    cand = os.path.join(base_dir, filename)
    i = 1
    while os.path.exists(cand):
        cand = os.path.join(base_dir, f"{base} ({i}){ext}")
        i += 1
    return cand

def _save_attachments(mail, target_dir):
    saved = 0
    try:
        atts = mail.Attachments
        cnt = getattr(atts, "Count", 0)
        for i in range(1, cnt + 1):
            try:
                att = atts.Item(i)
                raw = getattr(att, "FileName", f"vedlegg_{i}")
                clean = _sanitize_filename(raw)
                path = _unique_path(target_dir, clean)
                att.SaveAsFile(path)
                saved += 1
            except Exception:
                continue
    except Exception:
        pass
    return saved

# ---------- Body-tekst (enkel) ----------
_TAG_RE = re.compile(r"<[^>]+>")
def _as_text(mail):
    """Prøv HTML -> tekst, ellers Body."""
    try:
        html_body = getattr(mail, "HTMLBody", None)
        if html_body:
            txt = _TAG_RE.sub("", html_body)
            txt = html.unescape(txt)
            return txt.strip()
    except Exception:
        pass
    try:
        return (getattr(mail, "Body", "") or "").strip()
    except Exception:
        return ""

# ---------- GUI ----------
class OutlookTools(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Outlook-verktøy – søk/les/vedlegg")
        self.geometry("980x620")

        self.app = _get_outlook()
        if not self.app:
            messagebox.showerror("Outlook", "Kunne ikke starte Outlook (klassisk).")
            self.destroy(); return
        self.session = _get_session(self.app)
        if not self.session:
            messagebox.showerror("Outlook", "Fikk ikke tak i Outlook-sesjonen.")
            self.destroy(); return

        self._build_ui()

    def _build_ui(self):
        frm = ttk.Frame(self, padding=10)
        frm.pack(fill="both", expand=True)

        # Søkefelter
        top = ttk.Frame(frm)
        top.pack(fill="x")

        ttk.Label(top, text="Avsender (navn/epost):").grid(row=0, column=0, sticky="w")
        self.sender_var = tk.StringVar()
        ttk.Entry(top, textvariable=self.sender_var, width=40).grid(row=0, column=1, padx=6)

        ttk.Label(top, text="Fra dato (YYYY-MM-DD):").grid(row=0, column=2, sticky="w")
        self.from_var = tk.StringVar(value=(datetime.now() - timedelta(days=datetime.now().weekday())).strftime("%Y-%m-%d"))
        ttk.Entry(top, textvariable=self.from_var, width=12).grid(row=0, column=3, padx=6)

        self.subfolders_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(top, text="Inkluder undermapper", variable=self.subfolders_var).grid(row=0, column=4, padx=6)

        ttk.Button(top, text="Søk", command=self.on_search).grid(row=0, column=5, padx=6)

        # Resultattabell
        mid = ttk.Frame(frm)
        mid.pack(fill="both", expand=True, pady=(8, 6))

        cols = ("dato", "emne", "fra", "vedlegg")
        self.tree = ttk.Treeview(mid, columns=cols, show="headings", height=16)
        self.tree.heading("dato", text="Dato")
        self.tree.heading("emne", text="Emne")
        self.tree.heading("fra", text="Fra")
        self.tree.heading("vedlegg", text="#Vedlegg")
        self.tree.column("dato", width=150, anchor="w")
        self.tree.column("emne", width=520, anchor="w")
        self.tree.column("fra", width=220, anchor="w")
        self.tree.column("vedlegg", width=80, anchor="e")
        self.tree.pack(side="left", fill="both", expand=True)

        sb = ttk.Scrollbar(mid, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscroll=sb.set)
        sb.pack(side="right", fill="y")

        self.tree.bind("<<TreeviewSelect>>", self.on_select)

        # Detalj/lesevisning
        bottom = ttk.Frame(frm)
        bottom.pack(fill="both", expand=True)

        self.body = tk.Text(bottom, wrap="word")
        self.body.pack(side="left", fill="both", expand=True)
        sb2 = ttk.Scrollbar(bottom, orient="vertical", command=self.body.yview)
        self.body.configure(yscroll=sb2.set)
        sb2.pack(side="right", fill="y")

        # Knapper
        btns = ttk.Frame(frm)
        btns.pack(fill="x", pady=(6, 0))
        ttk.Button(btns, text="Åpne i Outlook", command=self.open_in_outlook).pack(side="left")
        ttk.Button(btns, text="Last ned vedlegg (valgt…)", command=self.save_selected_attachments).pack(side="left", padx=8)
        ttk.Button(btns, text="Last ned vedlegg (alle treff…)", command=self.save_all_attachments).pack(side="left")

        self.status = ttk.Label(frm, text="Klar.")
        self.status.pack(fill="x", pady=(8, 0))

        self.results = []   # liste av dicts (se _find_messages)
        self._id_to_ix = {} # map tree item -> index

    def on_search(self):
        # Les inn parametre
        q = self.sender_var.get().strip()
        try:
            since = datetime.strptime(self.from_var.get().strip(), "%Y-%m-%d").date()
        except Exception:
            messagebox.showerror("Dato", "Ugyldig datoformat. Bruk YYYY-MM-DD.")
            return
        incl = self.subfolders_var.get()

        self.tree.delete(*self.tree.get_children())
        self.body.delete("1.0", "end")
        self.status.config(text="Søker… dette kan ta litt tid ved store postbokser.")

        def _run():
            res = _find_messages(self.session, q, since, include_subfolders=incl, max_per_folder=6000, cap_total=3000)
            # Oppdater GUI i hovedtråd
            self.after(0, lambda: self._populate(res))

        threading.Thread(target=_run, daemon=True).start()

    def _populate(self, res):
        self.results = res
        self._id_to_ix.clear()
        for r in res:
            d = r["dt"].strftime("%Y-%m-%d %H:%M")
            subj = (r["subject"] or "").replace("\r", " ").replace("\n", " ")
            row_id = self.tree.insert("", "end", values=(d, subj, f"{r['from']} <{r['from_email'] or ''}>", r["attach"]))
            self._id_to_ix[row_id] = len(self._id_to_ix)
        self.status.config(text=f"Fant {len(res)} e‑poster.")

    def on_select(self, _evt):
        sel = self.tree.selection()
        if not sel:
            return
        r = self.results[self._id_to_ix[sel[0]]]
        txt = _as_text(r["item"])
        self.body.delete("1.0", "end")
        self.body.insert("1.0", txt)

    def _choose_folder(self):
        d = filedialog.askdirectory(title="Velg mappe for vedlegg")
        return d if d else None

    def open_in_outlook(self):
        sel = self.tree.selection()
        if not sel: return
        r = self.results[self._id_to_ix[sel[0]]]
        try:
            r["item"].Display()
        except Exception:
            messagebox.showerror("Outlook", "Klarte ikke å åpne e‑posten i Outlook.")

    def save_selected_attachments(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("Vedlegg", "Velg en rad først.")
            return
        target = self._choose_folder()
        if not target:
            return
        saved_sum = 0
        for iid in sel:
            r = self.results[self._id_to_ix[iid]]
            saved_sum += _save_attachments(r["item"], target)
        messagebox.showinfo("Vedlegg", f"Lagret {saved_sum} vedlegg.")

    def save_all_attachments(self):
        if not self.results:
            messagebox.showinfo("Vedlegg", "Ingen treff å lagre fra.")
            return
        target = self._choose_folder()
        if not target:
            return
        saved_sum = 0
        for r in self.results:
            saved_sum += _save_attachments(r["item"], target)
        messagebox.showinfo("Vedlegg", f"Lagret {saved_sum} vedlegg fra {len(self.results)} e‑poster.")

if __name__ == "__main__":
    app = OutlookTools()
    try:
        app.mainloop()
    except Exception:
        pass
