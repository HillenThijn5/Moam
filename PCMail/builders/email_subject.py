"""
Onderwerpbouwer voor de PC-notificatiemail.
"""
from PCMail.models.inputdefinition import PCMailProduct


def build_email_subject(product: PCMailProduct) -> str:
    """
    Bouwt de volledige Outlook-onderwerpregel voor de PC-mail.
    Formaat: 'PC Product Manufacturing Group: NEW PRODUCT NOTIFICATION (...) - series N'
    """
    p = product
    base = "PC Product Manufacturing Group: NEW PRODUCT NOTIFICATION"
    is_frn = "FIXED RATE NOTE" in p.product.upper()

    if not is_frn and p.underlyings:
        ow_part = f" linked to {' / '.join(u.ticker for u in p.underlyings)} in {p.currency}"
    else:
        ow_part = ""

    middle = f"{p.currency} {p.issue_size}, {p.maturity} {p.product}{ow_part}"
    suffix = f" - series {p.series}" if p.issuer.upper() == "VLK" else f" - Issuer {p.issuer}"

    return f"{base} ({middle}){suffix}"
