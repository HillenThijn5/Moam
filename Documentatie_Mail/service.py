# Documentatie_Mail/service.py
"""
Service op hoger niveau voor het versturen van de Documentatie Mail.
Stuurt het opbouwen en versturen van de mail aan.
"""

from Documentatie_Mail.email_builder import (
    build_subject, build_body, build_body_belgium,
    derive_cc_from_trades, is_belgian,
)
from Documentatie_Mail.send_email import send_email
from statics.data import DOCUMENTATIE_MAIL_RECIPIENTS

FIXED_TO = DOCUMENTATIE_MAIL_RECIPIENTS["to"]
FIXED_CC = DOCUMENTATIE_MAIL_RECIPIENTS["fixed_cc"]

BSMO_EMAIL = "bsmo@merciervanlanschot.be"

def send_documentatie_mail(email_data: dict) -> None:
    # ✅ Verplichte velden checken
    required_fields = ["product", "isin", "maturity", "currency", "issuer", "trades"]
    missing = [f for f in required_fields if not email_data.get(f)]
    if missing:
        raise ValueError(f"Missing required fields: {', '.join(missing)}")

    if not email_data.get("trades"):
        raise ValueError("At least one trade is required")

    # ✅ Zorg altijd voor TO-ontvangers
    to_names = email_data.get("to") or FIXED_TO

    if not to_names:
        raise ValueError("Documentatie mail has no TO recipients")

    # ✅ Bouw de CC-lijst op
    cc_names = []
    cc_names += email_data.get("cc", [])
    cc_names += FIXED_CC
    cc_names += derive_cc_from_trades(email_data.get("trades", []))

    # Belgische deals: voeg Business Support MID Office toe
    belgian = is_belgian(email_data)
    if belgian and BSMO_EMAIL not in cc_names:
        cc_names.append(BSMO_EMAIL)

    # haal dubbelen weg en hou het netjes
    cc_names = sorted(set(cc_names))

    subject = build_subject(email_data)
    body = build_body_belgium(email_data) if belgian else build_body(email_data)

    send_email(
        subject=subject,
        html_body=body,
        to_names=to_names,
        cc_names=cc_names,
    )
