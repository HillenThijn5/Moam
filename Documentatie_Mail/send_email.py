# Documentatie_Mail/send_email.py
"""
Low-level Outlook email sending utility.
Handles the actual interaction with Outlook COM interface.
"""


def send_email(
        subject: str,
        html_body: str,
        to_names: list,
        cc_names: list = None
) -> None:
    """
    Send an email via Outlook with HTML body and signature.
    Lazy imports win32com to avoid import errors on non-Windows systems.

    Args:
        subject: Email subject line
        html_body: HTML formatted email body
        to_names: List of recipient names (Outlook will resolve them)
        cc_names: List of CC recipient names
    """
    try:
        import win32com.client  # Lazy import
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

    # Add TO recipients by name
    for name in to_names:
        r = mail.Recipients.Add(name)
        r.Type = 1  # 1 = To

    # Add CC recipients by name
    for name in cc_names:
        r = mail.Recipients.Add(name)
        r.Type = 2  # 2 = CC

    # Force Outlook to resolve names
    mail.Recipients.ResolveAll()

    mail.Display()  # load signature
    signature = mail.HTMLBody
    mail.HTMLBody = html_body + signature