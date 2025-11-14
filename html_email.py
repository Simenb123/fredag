from datetime import datetime
from html import escape
from .config import TOP_N_SENDERS

def build_html(subject_text: str, status_text: str, stats):
    ts = datetime.now().strftime("%d.%m.%Y %H:%M")
    rows = []
    if stats:
        for name, addr, cnt in stats:
            rows.append(
                f"<tr>"
                f"<td style='padding:6px 8px;border:1px solid #d0d7de'>{escape(name or '—')}</td>"
                f"<td style='padding:6px 8px;border:1px solid #d0d7de'>{escape(addr or '—')}</td>"
                f"<td style='padding:6px 8px;border:1px solid #d0d7de;text-align:right'>{cnt}</td>"
                f"</tr>"
            )
    else:
        rows.append("<tr><td colspan='3' style='padding:10px;border:1px solid #d0d7de'>Ingen e‑poster funnet denne uken.</td></tr>")

    table_html = (
        "<table style='border-collapse:collapse;width:100%;margin-top:6px'>"
        "<thead><tr style='background:#f2f4f7'>"
        "<th style='padding:8px;border:1px solid #d0d7de;text-align:left'>Avsender</th>"
        "<th style='padding:8px;border:1px solid #d0d7de;text-align:left'>Adresse</th>"
        "<th style='padding:8px;border:1px solid #d0d7de;text-align:right'>Antall</th>"
        "</tr></thead>"
        "<tbody>" + "".join(rows) + "</tbody></table>"
    )

    return f"""<!doctype html>
<html><head><meta charset="utf-8"><title>{escape(subject_text)}</title></head>
<body style="font-family:'Segoe UI', Arial, sans-serif; font-size:12pt; color:#111">
  <p style="margin:0 0 12px 0">Hei!</p>
  <p style="margin:0 0 4px 0; font-weight:600">Status fra Helgesjekk‑appen:</p>
  <p style="margin:0 0 12px 0">{escape(status_text)}</p>
  <hr style="border:none;border-top:1px solid #e5e7eb;margin:12px 0">
  <p style="margin:0 0 4px 0; font-weight:600">Ukesoppsummering – avsendere (inneværende uke):</p>
  {table_html}
  <p style="margin-top:12px; font-size:10pt; color:#6b7280">Generert {escape(ts)} · Topp {TOP_N_SENDERS} avsendere.</p>
</body></html>"""
