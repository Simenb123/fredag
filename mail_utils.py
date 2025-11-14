from __future__ import annotations

def send_html_mail(session, to_email: str, subject: str, html: str) -> tuple[bool, str]:
    """
    Enkelt HTML‑send via Outlook. Bruker standardkonto i profilen.
    """
    try:
        app = session.Application
        mail = app.CreateItem(0)  # olMailItem
        mail.To = to_email
        mail.Subject = subject
        mail.HTMLBody = html
        mail.Send()
        return True, "Sendt."
    except Exception as e:
        return False, f"Feil ved sending: {e}"
