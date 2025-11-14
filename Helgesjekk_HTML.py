import tkinter as tk
from datetime import datetime, timedelta, time, date
import time as pytime
import re
import html as htmlmod  # for HTML-escaping

# =======================
#  KONFIG
# =======================
WEEKEND_CUTOFF = time(15, 15)   # fredag 15:15
DAGNAVN = ["mandag","tirsdag","onsdag","torsdag","fredag","lørdag","søndag"]

TOP_N_SENDERS = 25              # antall avsendere i e-posten
SEARCH_TIMEOUT_SEC = 12         # ventetid per AdvancedSearch-kall
MAX_PER_FOLDER = 6000           # maks meldinger per mappe ved manuell skanning
LOOKBACK_DAYS = 30              # sikkerhetsmargin bakover ved manuell skanning
FALLBACK_RECENT_N = 2000        # N siste i default Innboks (fallback), filtreres til uke
FALLBACK_EMAIL = None           # f.eks. "din@adresse.no" hvis profil ikke gir standardadresse

# =======================
#  OUTLOOK (klassisk)
# =======================
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

def _default_smtp(session):
    """Hent standard SMTP: DefaultStore->Owner->Primary/MAPI, deretter Accounts."""
    PR_SMTP = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E"
    try:
        store = session.DefaultStore
        if store:
            try:
                owner = store.GetOwner()
                ae = owner.AddressEntry if owner else None
                if ae:
                    try:
                        exu = ae.GetExchangeUser()
                        if exu and exu.PrimarySmtpAddress:
                            return exu.PrimarySmtpAddress
                    except Exception:
                        pass
                    try:
                        smtp = ae.PropertyAccessor.GetProperty(PR_SMTP)
                        if smtp and "@" in smtp:
                            return smtp
                    except Exception:
                        pass
            except Exception:
                pass
        # Fallbacks via Accounts
        try:
            for i in range(1, session.Accounts.Count + 1):
                acc = session.Accounts.Item(i)
                if getattr(acc, "SmtpAddress", None):
                    return acc.SmtpAddress
            for i in range(1, session.Accounts.Count + 1):
                acc = session.Accounts.Item(i)
                uname = getattr(acc, "UserName", "")
                if "@" in (uname or ""):
                    return uname
        except Exception:
            pass
    except Exception:
        pass
    return None

# =======================
#  AVSENDER-OPPSLAG
# =======================
PR_SENDER_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x5D01001E"
PR_SENDER_NAME         = "http://schemas.microsoft.com/mapi/proptag/0x0C1A001E"
PR_HEADERS             = "http://schemas.microsoft.com/mapi/proptag/0x007D001E"
FROM_REGEX = re.compile(r"^From:\s*(?P<disp>.*?)\s*<(?P<smtp>[^>]+)>", re.IGNORECASE | re.MULTILINE)

def _normalize_sender(mail):
    """Returner (display_name, smtp_lower) via robuste fallbacks."""
    name, smtp = "", ""
    try:
        pa = mail.PropertyAccessor
    except Exception:
        pa = None

    # 1) Direkte MAPI-felt
    if pa:
        try:
            v = pa.GetProperty(PR_SENDER_SMTP_ADDRESS)
            if v and "@" in v:
                smtp = v.strip().lower()
        except Exception:
            pass

    # 2) ExchangeUser Primary SMTP
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

    # 3) AddressEntry.PropertyAccessor -> 0x39FE
    if not smtp:
        try:
            ae = mail.Sender
            if ae is not None:
                smtp2 = ae.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x39FE001E")
                if smtp2 and "@" in smtp2:
                    smtp = smtp2.strip().lower()
        except Exception:
            pass

    # 4) From: i rå headere
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

    # 5) Siste utvei
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

# =======================
#  TID/HELGE-HJELPERE
# =======================
def day_name(d: date) -> str:
    return DAGNAVN[d.weekday()]

def next_friday_cutoff(now: datetime) -> datetime:
    days_to_fri = (4 - now.weekday()) % 7
    candidate = (now + timedelta(days=days_to_fri)).replace(
        hour=WEEKEND_CUTOFF.hour, minute=WEEKEND_CUTOFF.minute, second=0, microsecond=0
    )
    if candidate <= now:
        candidate += timedelta(days=7)
    return candidate

def next_monday_midnight(now: datetime) -> datetime:
    days_to_mon = (7 - now.weekday()) % 7
    candidate = (now + timedelta(days=days_to_mon)).replace(
        hour=0, minute=0, second=0, microsecond=0
    )
    if candidate <= now:
        candidate += timedelta(days=7)
    return candidate

def is_weekend(now: datetime) -> bool:
    wd = now.weekday()
    return wd in (5, 6) or (wd == 4 and now.time() >= WEEKEND_CUTOFF)

def days_until_friday(now: datetime) -> int:
    return (4 - now.weekday()) % 7

def _start_of_week_local() -> datetime:
    t = datetime.now()
    return (t - timedelta(days=t.weekday())).replace(hour=0, minute=0, second=0, microsecond=0)

def _msg_time(it):
    """Robust uthenting av meldingstid (ReceivedTime/CreationTime/SentOn)."""
    for attr in ("ReceivedTime", "CreationTime", "SentOn"):
        try:
            dt = getattr(it, attr, None)
            if dt:
                return dt
        except Exception:
            continue
    return None

# =======================
#  HENTING VIA OUTLOOK
# =======================
def _run_advanced_search(app, scope_str, query):
    """Kjør AdvancedSearch og returner liste_av_mailitems. scope_str må være sitert ('...')."""
    srch = app.AdvancedSearch(Scope=scope_str, Filter=query, SearchSubFolders=True, Tag=f"HS_{int(pytime.time()*1000)}")
    waited = 0.0
    while True:
        try:
            if getattr(srch, "Complete", False):
                break
        except Exception:
            break
        pytime.sleep(0.25)
        waited += 0.25
        if waited >= SEARCH_TIMEOUT_SEC:
            break

    results = []
    try:
        res = srch.Results
        cnt = getattr(res, "Count", 0)
        for i in range(1, cnt + 1):
            try:
                it = res.Item(i)
                if getattr(it, "Class", None) == 43:  # olMailItem
                    results.append(it)
            except Exception:
                continue
    except Exception:
        pass
    return results

def _build_inbox_scope_list(session):
    """Returner en komma-separert streng med **sitert** sti for alle Innboks-mapper."""
    parts = []
    olFolderInbox = 6
    try:
        ds = session.DefaultStore
        if ds:
            try:
                p = ds.GetDefaultFolder(olFolderInbox).FolderPath.replace("'", "''")
                parts.append(f"'{p}'")
            except Exception:
                pass
        for i in range(1, session.Stores.Count + 1):
            st = session.Stores.Item(i)
            if ds and st.StoreID == ds.StoreID:
                continue
            try:
                p = st.GetDefaultFolder(olFolderInbox).FolderPath.replace("'", "''")
                parts.append(f"'{p}'")
            except Exception:
                continue
    except Exception:
        pass
    return ", ".join(parts)

def _walk_subfolders(folder):
    yield folder
    try:
        subs = folder.Folders
        for i in range(1, subs.Count + 1):
            sub = subs.Item(i)
            for f in _walk_subfolders(sub):
                yield f
    except Exception:
        return

def _count_senders_in_folder(folder, sow_date, per_folder_limit=MAX_PER_FOLDER):
    counts = {}
    # Les items fra mappen
    try:
        items = folder.Items
    except Exception:
        return counts
    # Sorter (hvis mulig)
    try:
        items.Sort("[ReceivedTime]", True)
    except Exception:
        pass

    processed = 0
    # Foretrukket iterasjon: GetFirst/GetNext
    use_seq = True
    try:
        it = items.GetFirst()
    except Exception:
        use_seq = False
        it = None

    if use_seq:
        while it and processed < per_folder_limit:
            try:
                if getattr(it, "Class", None) == 43:
                    dt = _msg_time(it)
                    if dt and dt.date() >= sow_date:
                        name, smtp = _normalize_sender(it)
                        key = ((name or "")[:120], (smtp or "")[:200])
                        counts[key] = counts.get(key, 0) + 1
                processed += 1
                it = items.GetNext()
            except Exception:
                break
    else:
        # Fallback: indeksering
        total = 0
        try:
            total = items.Count
        except Exception:
            total = 0
        upto = min(per_folder_limit, total)
        for idx in range(1, upto + 1):
            try:
                it = items.Item(idx)
                if it and getattr(it, "Class", None) == 43:
                    dt = _msg_time(it)
                    if dt and dt.date() >= sow_date:
                        name, smtp = _normalize_sender(it)
                        key = ((name or "")[:120], (smtp or "")[:200])
                        counts[key] = counts.get(key, 0) + 1
            except Exception:
                continue

    return counts

def _weekly_sender_stats(session, top_n=TOP_N_SENDERS):
    """Avsenderstatistikk for inneværende uke. Skanner Innboks + undermapper."""
    sow_date = _start_of_week_local().date()
    counts = {}

    # Default Innboks
    try:
        inbox = session.GetDefaultFolder(6)  # olFolderInbox
    except Exception:
        inbox = None

    # Skann Innboks + undermapper
    if inbox is not None:
        for folder in _walk_subfolders(inbox):
            sub_counts = _count_senders_in_folder(folder, sow_date)
            if sub_counts:
                for k, v in sub_counts.items():
                    counts[k] = counts.get(k, 0) + v

    # Fallback: hvis helt tomt, ta de N siste i Innboks (fortsatt filtrert til uke)
    if not counts and inbox is not None:
        counts = _count_senders_in_folder(inbox, sow_date, per_folder_limit=FALLBACK_RECENT_N)

    ranked = sorted(((k[0], k[1], v) for k, v in counts.items()), key=lambda x: x[2], reverse=True)
    return ranked[:top_n]

# =======================
#  HTML-E-POST
# =======================
def _build_html(subject_text, status_text, stats):
    esc = htmlmod.escape
    timestamp = datetime.now().strftime("%d.%m.%Y %H:%M")
    rows_html = []

    if stats:
        for name, addr, cnt in stats:
            rows_html.append(
                f"<tr>"
                f"<td style='padding:8px;border:1px solid #e5e7eb'>{esc(name or '—')}</td>"
                f"<td style='padding:8px;border:1px solid #e5e7eb'>{esc(addr or '—')}</td>"
                f"<td style='padding:8px;border:1px solid #e5e7eb;text-align:right'>{cnt}</td>"
                f"</tr>"
            )
    else:
        rows_html.append(
            "<tr><td colspan='3' style='padding:10px;border:1px solid #e5e7eb'>"
            "Ingen e‑poster funnet denne uken.</td></tr>"
        )

    table_html = (
        "<table style='border-collapse:collapse;width:100%;margin-top:8px'>"
        "<thead>"
        "<tr style='background:#f3f4f6'>"
        "<th style='padding:8px;border:1px solid #e5e7eb;text-align:left'>Avsender</th>"
        "<th style='padding:8px;border:1px solid #e5e7eb;text-align:left'>Adresse</th>"
        "<th style='padding:8px;border:1px solid #e5e7eb;text-align:right'>Antall</th>"
        "</tr>"
        "</thead>"
        "<tbody>"
        + "".join(rows_html) +
        "</tbody></table>"
    )

    html = f"""
<!doctype html>
<html>
  <head><meta charset="utf-8"><title>{esc(subject_text)}</title></head>
  <body style="font-family:'Segoe UI', Arial, sans-serif; font-size:12pt; color:#111; line-height:1.35">
    <p style="margin:0 0 12px 0;">Hei!</p>
    <p style="margin:0 0 4px 0; font-weight:600;">Status fra Helgesjekk‑appen:</p>
    <p style="margin:0 0 12px 0;">{esc(status_text)}</p>
    <hr style="border:none;border-top:1px solid #e5e7eb; margin:12px 0;">
    <p style="margin:0 0 4px 0; font-weight:600;">Ukesoppsummering – avsendere (inneværende uke):</p>
    {table_html}
    <p style="margin-top:12px; font-size:10pt; color:#6b7280;">
      Generert {esc(timestamp)} · Topp {TOP_N_SENDERS} avsendere.
    </p>
  </body>
</html>
"""
    return html

# =======================
#  SEND E-POST
# =======================
def _send_mail(subject, status_text):
    if not win32:
        return False, "Outlook/pywin32 mangler. Installer pywin32 og bruk klassisk Outlook."
    app = _get_outlook()
    if not app:
        return False, "Klarte ikke å starte Outlook (klassisk)."
    session = _get_session(app)
    if not session:
        return False, "Klarte ikke å hente Outlook-sesjon."

    to_addr = _default_smtp(session) or FALLBACK_EMAIL
    if not to_addr:
        return False, "Fant ikke standard e-post i Outlook-profilen. Sett FALLBACK_EMAIL i koden."

    stats = _weekly_sender_stats(session, TOP_N_SENDERS)
    html = _build_html(subject, status_text, stats)

    m = app.CreateItem(0)  # olMailItem
    m.To = to_addr
    m.Subject = subject
    # HTML e‑post:
    try:
        m.BodyFormat = 2  # olFormatHTML
    except Exception:
        pass
    m.HTMLBody = html
    m.Send()
    return True, f"E-post sendt til {to_addr}."

# =======================
#  GUI + NEDTELLING
# =======================
countdown_job = None
_last_status_text = ""

def cancel_countdown():
    global countdown_job
    if countdown_job:
        try:
            root.after_cancel(countdown_job)
        except Exception:
            pass
        countdown_job = None

def start_countdown(target: datetime, prefix: str = ""):
    cancel_countdown()
    update_countdown(target, prefix)

def update_countdown(target: datetime, prefix: str):
    global countdown_job, _last_status_text
    now = datetime.now()
    remaining = target - now
    if remaining.total_seconds() <= 0:
        txt = "🎉 GOD HELG! 🎉" if is_weekend(datetime.now()) else "Helgen er over. God mandag!"
        try:
            svar_label.config(text=txt)
        except Exception:
            pass
        _last_status_text = txt
        countdown_job = None
        return

    total = int(remaining.total_seconds())
    days, rem = divmod(total, 86400)
    hours, rem = divmod(rem, 3600)
    minutes, seconds = divmod(rem, 60)
    if days > 0:
        txt = f"{prefix}{days}d {hours:02d}:{minutes:02d}:{seconds:02d}"
    else:
        txt = f"{prefix}{hours:02d}:{minutes:02d}:{seconds:02d}"

    try:
        svar_label.config(text=txt)
    except Exception:
        pass
    _last_status_text = txt
    countdown_job = root.after(1000, update_countdown, target, prefix)

def sjekk_helg_og_send():
    # 1) Oppdater GUI / start nedtelling
    now = datetime.now()
    cancel_countdown()

    if is_weekend(now):
        today = day_name(now.date())
        slutt = next_monday_midnight(now)
        prefix = f"Det er {today} – det er helg 🎉  Helgen varer i: "
        start_countdown(slutt, prefix)
    else:
        today = day_name(now.date())
        n = days_until_friday(now)
        if n == 0:
            info = "Det er fredag – nedtelling til helg: "
        elif n == 1:
            info = f"Det er {today} – 1 dag til helg. Nedtelling: "
        else:
            info = f"Det er {today} – {n} dager til helg. Nedtelling: "
        start_countdown(next_friday_cutoff(now), info)

    # 2) Send e-post (HTML)
    subject = f"Helgesjekk + ukesoppsummering – {datetime.now():%Y-%m-%d %H:%M}"
    ok, msg = _send_mail(subject, _last_status_text)
    try:
        feedback_label.config(text=msg, fg=("green" if ok else "red"))
    except Exception:
        pass

# =======================
#  GUI
# =======================
root = tk.Tk()
root.title("Helgesjekk (Outlook – ukesoppsummering, HTML)")
tk.Label(root, text="Helgesjekk", font=("TkDefaultFont", 13, "bold")).pack(padx=20, pady=(18, 8))
tk.Button(root, text="Sjekk helg", width=16, command=sjekk_helg_og_send).pack(pady=6)
svar_label = tk.Label(root, text="Trykk «Sjekk helg» – sender pen HTML‑epost for inneværende uke.", font=("TkDefaultFont", 11))
svar_label.pack(pady=(6, 6))
feedback_label = tk.Label(root, text="", font=("TkDefaultFont", 9))
feedback_label.pack(pady=(2, 14))
root.mainloop()
