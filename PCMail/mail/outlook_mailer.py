# mail/outlook_mailer.py

import os
import re
from pathlib import Path

import win32com.client

from PCMail.config.recipients import TO_RECIPIENTS
from statics.data import STRUCTURED_INVESTMENTS_EMAIL


def _get_signature_html() -> str:
    """Lees de standaard Outlook-handtekening uit het bestandssysteem."""
    sig_dir = Path(os.environ.get("APPDATA", "")) / "Microsoft" / "Signatures"
    if not sig_dir.exists():
        return ""
    htm_files = list(sig_dir.glob("*.htm"))
    if not htm_files:
        return ""
    sig_file = htm_files[0]
    try:
        content = sig_file.read_text(encoding="utf-8")
        body_match = re.search(r'<body[^>]*>(.*)</body>', content, re.DOTALL | re.IGNORECASE)
        if body_match:
            return body_match.group(1).strip()
        return content.strip()
    except Exception:
        return ""


def send_mail(subject, html_body, attachments):

    outlook = win32com.client.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)

    mail.To = TO_RECIPIENTS
    mail.CC = STRUCTURED_INVESTMENTS_EMAIL
    mail.Subject = subject

    sig_html = _get_signature_html()
    if sig_html:
        html_body = re.sub(
            r'</body>',
            f'<br>{sig_html}</body>',
            html_body,
            count=1,
            flags=re.IGNORECASE,
        )

    mail.HTMLBody = html_body

    for f in attachments:
        mail.Attachments.Add(f)

    mail.Display()
