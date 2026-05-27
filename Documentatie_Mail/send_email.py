# Documentatie_Mail/send_email.py
"""
Low-level helper voor het versturen van Outlook-mails.
Regelt de echte interactie met de Outlook-COM-interface.
"""
import os
import re
from pathlib import Path


def _get_signature_html() -> str:
    """Lees de standaard-Outlookhandtekening van het bestandssysteem."""
    sig_dir = Path(os.environ.get("APPDATA", "")) / "Microsoft" / "Signatures"
    if not sig_dir.exists():
        return ""
    htm_files = list(sig_dir.glob("*.htm"))
    if not htm_files:
        return ""
    sig_file = htm_files[0]
    try:
        content = sig_file.read_text(encoding="utf-8")
        # Haal alleen de body-inhoud eruit (handtekeningbestanden zijn fragmenten)
        body_match = re.search(r'<body[^>]*>(.*)</body>', content, re.DOTALL | re.IGNORECASE)
        if body_match:
            return body_match.group(1).strip()
        return content.strip()
    except Exception:
        return ""


def send_email(
        subject: str,
        html_body: str,
        to_names: list,
        cc_names: list = None
) -> None:
    """
    Verstuur een mail via Outlook met HTML-inhoud en handtekening.

    Beperkingen (Outlook 16.0.19929+ security-update):
    - Recipients.Add of ResolveAll kan NIET na Display (Operation aborted)
    - PropertyAccessor kan NIET na Display
    - WordEditor kan NIET na Display
    - mail.To/CC/Subject/HTMLBody zetten kan WEL vóór Display
    - HTMLBody en Attachments.Add zetten kan WEL na Display

    Aanpak: zet alles klaar (ook ontvangers via mail.To/CC-strings)
    vóór Display(). Lees de handtekening uit het bestandssysteem.
    """
    try:
        import win32com.client
    except ImportError as e:
        raise ImportError(
            "win32com not installed. Install with: pip install pywin32\n"
            "Then run: python -m pip install --upgrade --force-reinstall pywin32"
        ) from e

    if cc_names is None:
        cc_names = []

    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
    except Exception as e:
        raise RuntimeError(
            "Could not connect to Outlook. Make sure Outlook is installed and running."
        ) from e

    mail = outlook.CreateItem(0)
    mail.Subject = subject

    # Zet ontvangers via directe stringtoewijzing (NIET via Recipients.Add)
    # Outlook lost de namen op zodra de gebruiker met de mail werkt
    if to_names:
        mail.To = "; ".join(to_names)
    if cc_names:
        mail.CC = "; ".join(cc_names)

    # Voeg de handtekening uit het bestandssysteem toe
    sig_html = _get_signature_html()
    if sig_html:
        if re.search(r'</body>', html_body, re.IGNORECASE):
            html_body = re.sub(
                r'</body>',
                f'<br>{sig_html}</body>',
                html_body,
                count=1,
                flags=re.IGNORECASE,
            )
        else:
            html_body = html_body + f"<br>{sig_html}"

    # Zet HTMLBody VOOR Display — dat is de enige veilige volgorde
    mail.HTMLBody = html_body
    mail.Display()