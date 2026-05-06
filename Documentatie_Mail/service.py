# Documentatie_Mail/service.py
"""
High-level service for sending Documentatie Mail.
Orchestrates building and sending the email.
"""

from Documentatie_Mail.email_builder import build_subject, build_body, derive_cc_from_trades
from Documentatie_Mail.send_email import send_email
from statics.data import DOCUMENTATIE_MAIL_RECIPIENTS

FIXED_TO = DOCUMENTATIE_MAIL_RECIPIENTS["to"]
FIXED_CC = DOCUMENTATIE_MAIL_RECIPIENTS["fixed_cc"]

def send_documentatie_mail(email_data: dict) -> None:
    # ✅ Validate required fields
    required_fields = ["product", "isin", "maturity", "currency", "issuer", "trades"]
    missing = [f for f in required_fields if not email_data.get(f)]
    if missing:
        raise ValueError(f"Missing required fields: {', '.join(missing)}")

    if not email_data.get("trades"):
        raise ValueError("At least one trade is required")

    # ✅ Always ensure TO recipients
    to_names = email_data.get("to") or FIXED_TO

    if not to_names:
        raise ValueError("Documentatie mail has no TO recipients")

    # ✅ Build CC list
    cc_names = []
    cc_names += email_data.get("cc", [])
    cc_names += FIXED_CC
    cc_names += derive_cc_from_trades(email_data.get("trades", []))

    # de-duplicate, keep clean
    cc_names = sorted(set(cc_names))

    subject = build_subject(email_data)
    body = build_body(email_data)

    send_email(
        subject=subject,
        html_body=body,
        to_names=to_names,
        cc_names=cc_names,
    )
