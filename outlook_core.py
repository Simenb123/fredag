from __future__ import annotations
import os
import html as htmlmod
import re
from datetime import datetime, date
from typing import Dict, Iterator, List, Optional, Tuple, Callable

# --------- logging (valgfritt, faller stille tilbake) ----------
try:
    from .log_utils import get_logger  # type: ignore
    log = get_logger(__name__)
except Exception:  # pragma: no cover
    class _Null:
        def info(self, *a, **k): ...
        def warning(self, *a, **k): ...
        def error(self, *a, **k): ...
        def exception(self, *a, **k): ...
    log = _Null()

# ---------- Outlook bootstrap ----------
def have_outlook() -> bool:
    """Rask sjekk at pywin32/Outlook finnes (for helgesjekk_app)."""
    try:
        import win32com.client  # type: ignore
        _ = win32com.client
        return True
    except Exception:
        return False

def get_outlook():
    import win32com.client  # type: ignore
    return win32com.client.Dispatch("Outlook.Application")

def get_session(app=None):
    if app is None:
        app = get_outlook()
    return app.Session

# ---------- Avsender / tider / tekst ----------
SMTP_PROP = "http://schemas.microsoft.com/mapi/proptag/0x5D01001E"

def default_smtp(session) -> Optional[str]:
    try:
        store = session.DefaultStore
        owner = getattr(store, "Owner", None)
        smtp = getattr(owner, "PrimarySmtpAddress", None)
        if smtp:
            return str(smtp)
    except Exception:
        pass
    try:
        accs = session.Accounts
        if accs and accs.Count >= 1:
            return str(accs.Item(1).SmtpAddress)
    except Exception:
        pass
    return None

def normalize_sender(mail) -> Tuple[str, str]:
    """Returnerer (navn, smtp) – robust også for Exchange."""
    name, smtp = "", ""
    try:
        pa = mail.PropertyAccessor
        v = pa.GetProperty(SMTP_PROP)
        if v and "@" in v:
            smtp = str(v).strip().lower()
    except Exception:
        pass
    if not name:
        try:
            name = (getattr(mail, "SenderName", "") or "").strip()
        except Exception:
            pass
    if not smtp:
        try:
            raw = getattr(mail, "SenderEmailAddress", "") or ""
            if "@" in raw:
                smtp = raw.strip().lower()
        except Exception:
            pass
    return name, smtp

def msg_time(item) -> Optional[datetime]:
    for a in ("ReceivedTime", "CreationTime", "SentOn"):
        try:
            dt = getattr(item, a, None)
            if dt:
                return dt
        except Exception:
            continue
    return None

# ---------- Tekstudtrekk ----------
_RE_STYLE = re.compile(r"<style[^>]*>.*?</style>", re.I | re.S)
_RE_SCRIPT = re.compile(r"<script[^>]*>.*?</script>", re.I | re.S)
_RE_COM = re.compile(r"<!--.*?-->", re.S)

def _a_to_text(h: str) -> str:
    def rep(m):
        href = (m.group(1) or "").strip()
        txt = (m.group(2) or "").strip()
        if txt and href and href not in txt:
            return f"{htmlmod.unescape(txt)} ({htmlmod.unescape(href)})"
        return htmlmod.unescape(txt or href or "")
    return re.sub(r'<a[^>]+href=["\']?([^"\'>\s]+)[^>]*>(.*?)</a>', rep, h, flags=re.I | re.S)

def html_to_text(html: str) -> str:
    if not html:
        return ""
    h = _RE_STYLE.sub("", html)
    h = _RE_SCRIPT.sub("", h)
    h = _RE_COM.sub("", h)
    h = _a_to_text(h)
    h = re.sub(r"<\s*br\s*/?>", "\n", h, flags=re.I)
    h = re.sub(r"</\s*p\s*>", "\n\n", h, flags=re.I)
    h = re.sub(r"<[^>]+>", "", h)
    h = htmlmod.unescape(h).replace("\r", "")
    h = re.sub(r"[ \t]+\n", "\n", h)
    h = re.sub(r"\n{3,}", "\n\n", h)
    return h.strip()

def mail_as_text(mail) -> str:
    try:
        html = getattr(mail, "HTMLBody", None)
        if html:
            t = html_to_text(html)
            if t:
                return t
    except Exception:
        pass
    try:
        return (getattr(mail, "Body", "") or "").strip()
    except Exception:
        return ""

# ---------- Vedlegg ----------
_ILLEGAL_FS = r'<>:"/\\|?*'
def sanitize_filename(name: str) -> str:
    return "".join("_" if c in _ILLEGAL_FS else c for c in (name or ""))

def unique_path(dir_path: str, filename: str) -> str:
    base, ext = os.path.splitext(filename or "vedlegg")
    cand = os.path.join(dir_path, filename or "vedlegg")
    i = 1
    while os.path.exists(cand):
        cand = os.path.join(dir_path, f"{base} ({i}){ext}")
        i += 1
    return cand

def save_attachments(mail, target_dir: str) -> int:
    os.makedirs(target_dir, exist_ok=True)
    try:
        atts = mail.Attachments
        count = getattr(atts, "Count", 0)
    except Exception:
        return 0
    saved = 0
    for i in range(1, count + 1):
        try:
            att = atts.Item(i)
            name = sanitize_filename(getattr(att, "FileName", f"vedlegg_{i}") or f"vedlegg_{i}")
            path = unique_path(target_dir, name)
            att.SaveAsFile(path)
            saved += 1
        except Exception:
            log.exception("Feil ved lagring av vedlegg")
            continue
    return saved

# ---------- Mappevandring ----------
def walk_subfolders(folder, include_subfolders: bool) -> Iterator:
    yield folder
    if not include_subfolders:
        return
    try:
        subs = folder.Folders
        for i in range(1, subs.Count + 1):
            sub = subs.Item(i)
            for f in walk_subfolders(sub, True):
                yield f
    except Exception:
        return

# ---------- Restrict‑hjelpere ----------
def _fmt(dt: datetime) -> str:
    # US-format (12‑timers) kreves av Outlook Restrict/GetTable
    return dt.strftime("%m/%d/%Y %I:%M %p")

def _restrict_str(after: Optional[datetime], before: Optional[datetime],
                  only_unread: Optional[bool], has_attachments: Optional[bool]) -> str:
    clauses = []
    if after:
        clauses.append(f"[ReceivedTime] >= '{_fmt(after)}'")
    if before:
        clauses.append(f"[ReceivedTime] <= '{_fmt(before)}'")
    if only_unread is True:
        clauses.append("[UnRead] = True")
    if has_attachments is True:
        clauses.append("[HasAttachment] = True")
    return " AND ".join(clauses)

# ---------- Intern: GetTable‑motor ----------
def _search_via_gettable(session, inbox, flt_base: str, q_sender: str, q_subj: str,
                         include_subfolders: bool, cap_per_folder: int, cap_total: int,
                         stop_evt, progress: Optional[Callable[[str, int, int], None]]) -> Tuple[List[Dict], Optional[str], bool]:
    results: List[Dict] = []
    aborted = False

    for folder in walk_subfolders(inbox, include_subfolders):
        if stop_evt.is_set():
            aborted = True
            break
        if len(results) >= cap_total:
            break

        added_folder = 0
        try:
            tbl = folder.GetTable(flt_base) if flt_base else folder.GetTable()
            cols = tbl.Columns
            for col in ["[EntryID]", "[ReceivedTime]", "[Subject]", "[SenderName]",
                        "[SenderEmailAddress]", "[UnRead]", "[HasAttachment]"]:
                try: cols.Add(col)
                except Exception: pass
        except Exception as e:
            if progress:
                try: progress(getattr(folder, "FolderPath", ""), 0, len(results))
                except Exception: pass
            log.warning("GetTable feilet for %s: %s", getattr(folder, "FolderPath", ""), e)
            continue

        try:
            row = tbl.GetNextRow()
        except Exception:
            row = None

        while row and added_folder < cap_per_folder and len(results) < cap_total:
            if stop_evt.is_set():
                aborted = True
                break
            try:
                dt = row.Item("ReceivedTime")
                if not isinstance(dt, datetime):
                    row = tbl.GetNextRow();  continue
                subj = (row.Item("Subject") or "")
                from_name = (row.Item("SenderName") or "")
                from_raw = (row.Item("SenderEmailAddress") or "")
                unread = bool(row.Item("UnRead"))
                has_att = bool(row.Item("HasAttachment"))

                if q_subj and q_subj not in subj.lower():
                    row = tbl.GetNextRow();  continue
                if q_sender:
                    if q_sender not in (from_name or "").lower() and q_sender not in (from_raw or "").lower():
                        row = tbl.GetNextRow();  continue

                results.append({
                    "eid": row.Item("EntryID"),
                    "store": getattr(folder, "StoreID", None),
                    "dt": dt,
                    "from": from_name,
                    "from_email": from_raw.lower() if isinstance(from_raw, str) else "",
                    "subject": subj,
                    "folder": getattr(folder, "FolderPath", ""),
                    "attach": 1 if has_att else 0,  # hurtig indikator
                    "unread": unread,
                })
                added_folder += 1
                if progress and added_folder % 200 == 0:
                    try: progress(getattr(folder, "FolderPath", ""), added_folder, len(results))
                    except Exception: pass
            except Exception:
                pass
            try:
                row = tbl.GetNextRow()
            except Exception:
                break

        if progress:
            try: progress(getattr(folder, "FolderPath", ""), added_folder, len(results))
            except Exception: pass

    return results, None, aborted

# ---------- Intern: Items.Restrict‑motor (fallback) ----------
def _search_via_items(session, inbox, flt_base: str, q_sender: str, q_subj: str,
                      include_subfolders: bool, cap_per_folder: int, cap_total: int,
                      stop_evt, progress: Optional[Callable[[str, int, int], None]]) -> Tuple[List[Dict], Optional[str], bool]:
    results: List[Dict] = []
    aborted = False

    for folder in walk_subfolders(inbox, include_subfolders):
        if stop_evt.is_set():
            aborted = True
            break
        if len(results) >= cap_total:
            break

        added_folder = 0
        try:
            items = folder.Items
            items.Sort("[ReceivedTime]", True)
        except Exception:
            continue

        try:
            rset = items.Restrict(flt_base) if flt_base else items
        except Exception:
            rset = items

        try:
            it = rset.GetFirst()
            iter_by_next = True
        except Exception:
            it = None
            iter_by_next = False

        def accept(mail) -> bool:
            name, smtp = normalize_sender(mail)
            if q_sender and not (q_sender in (smtp or "") or q_sender in (name or "").lower()):
                return False
            subj = (getattr(mail, "Subject", "") or "")
            if q_subj and q_subj not in subj.lower():
                return False
            return True

        def capture(mail, folder_obj):
            try:
                dt = msg_time(mail)
                if not dt:
                    return
                name, smtp = normalize_sender(mail)
                n_att = 0
                try:
                    n_att = getattr(mail.Attachments, "Count", 0)
                except Exception:
                    n_att = 0
                results.append({
                    "eid": getattr(mail, "EntryID", None),
                    "store": getattr(folder_obj, "StoreID", None),
                    "dt": dt,
                    "from": name,
                    "from_email": smtp,
                    "subject": (getattr(mail, "Subject", "") or ""),
                    "folder": getattr(folder_obj, "FolderPath", ""),
                    "attach": n_att,
                    "unread": bool(getattr(mail, "UnRead", False)),
                })
            except Exception:
                log.exception("Feil under bygging av søkeresultat")

        if iter_by_next:
            while it and added_folder < cap_per_folder and len(results) < cap_total:
                if stop_evt.is_set():
                    aborted = True
                    break
                try:
                    if getattr(it, "Class", None) == 43:  # olMail
                        if accept(it):
                            capture(it, folder)
                            added_folder += 1
                            if progress and added_folder % 200 == 0:
                                try: progress(getattr(folder, "FolderPath", ""), added_folder, len(results))
                                except Exception: pass
                    it = rset.GetNext()
                except Exception:
                    break
        else:
            total = getattr(rset, "Count", 0)
            upto = min(cap_per_folder, total)
            for idx in range(1, upto + 1):
                if stop_evt.is_set():
                    aborted = True
                    break
                if len(results) >= cap_total:
                    break
                try:
                    it = rset.Item(idx)
                    if getattr(it, "Class", None) != 43:
                        continue
                    if accept(it):
                        capture(it, folder)
                        added_folder += 1
                        if progress and added_folder % 200 == 0:
                            try: progress(getattr(folder, "FolderPath", ""), added_folder, len(results))
                            except Exception: pass
                except Exception:
                    continue

        if progress:
            try: progress(getattr(folder, "FolderPath", ""), added_folder, len(results))
            except Exception: pass

    return results, None, aborted

# ---------- Offentlig API: auto‑valg + fallback ----------
def search_messages(session,
                    sender_query: str,
                    subject_contains: str,
                    after_date: date,
                    before_date: date,
                    include_subfolders: bool,
                    only_unread: bool,
                    only_attachments: bool,
                    cap_per_folder: int,
                    cap_total: int,
                    stop_evt,
                    progress: Optional[Callable[[str, int, int], None]] = None):
    """
    Returnerer (results, error, aborted). Bruker GetTable, faller tilbake til Items.Restrict ved behov.
    """
    results: List[Dict] = []
    aborted = False
    try:
        import pythoncom  # type: ignore
        pythoncom.CoInitialize()
        inbox = session.GetDefaultFolder(6)  # olFolderInbox

        q_sender = (sender_query or "").strip().lower()
        q_subj = (subject_contains or "").strip().lower()

        after_dt = datetime.combine(after_date, datetime.min.time()) if after_date else None
        before_dt = datetime.combine(before_date, datetime.max.time()) if before_date else None
        flt_base = _restrict_str(after_dt, before_dt, only_unread, only_attachments)

        if progress:
            try: progress("Metode: GetTable", 0, 0)
            except Exception: pass

        results, err, aborted = _search_via_gettable(
            session, inbox, flt_base, q_sender, q_subj, include_subfolders,
            cap_per_folder, cap_total, stop_evt, progress
        )

        # Fallback hvis 0 og ikke eksplisitt avbrutt/feil
        if (not results) and (not aborted):
            if progress:
                try: progress("Bytter til Items.Restrict (fallback)", 0, 0)
                except Exception: pass
            results, err, aborted = _search_via_items(
                session, inbox, flt_base, q_sender, q_subj, include_subfolders,
                cap_per_folder, cap_total, stop_evt, progress
            )

        return results, err, aborted
    except Exception as e:
        log.exception("Uventet feil i search_messages (auto)")
        return [], f"Uventet feil i søk: {e}", aborted
    finally:
        try:
            import pythoncom  # type: ignore
            pythoncom.CoUninitialize()
        except Exception:
            pass
