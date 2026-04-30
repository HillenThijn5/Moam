# Documentatie_Mail/email_builder.py
"""
Build the email subject and HTML body for Documentatie Mail.
Handles product-specific content and formatting.
"""
from datetime import date
from statics.data import PRODUCT_TYPES, ADVISER_HELPERS, PRIIP_HUB_URL
from MarketingMail.product_title import UNDERLYING_ALIASES
from statics.outlook_user import get_sender_first_name


def build_subject(data: dict) -> str:
    """
    Build email subject line based on email data.
    Format: Documentatie {Issuer} {Product} {Underlyings} {StartYear} - {EndYear} – {ISIN}
    """
    today = data.get("today") or date.today()
    year_start = today.year % 100

    # Extract maturity years from format like "2Y", "3Y", etc.
    maturity_str = data["maturity"].replace("Y", "").replace("y", "")
    try:
        maturity_num = int(maturity_str)
    except ValueError:
        maturity_num = 0

    year_end = year_start + maturity_num

    # Join underlyings using aliases (e.g. SX5E → Eurozone)
    ow = " / ".join(
        UNDERLYING_ALIASES.get(u, u) for u in data.get("underlyings", [])
    )

    subject = f"Documentatie {data['issuer']} {data['product']}"
    if ow:
        subject += f" {ow}"
    subject += f" {year_start} - {year_end}"

    if data.get("isin"):
        subject += f" – {data['isin']}"

    return subject


def greeting(trades: list) -> str:
    """Generate appropriate greeting based on number of advisers."""
    if len(trades) == 1:
        first_name = trades[0]["adviser"].split(" ")[0]
        return f"Hi {first_name},"
    return "Hi all,"


def build_rtd_list(data: dict) -> str:
    """Build HTML list of trades/RTD entries."""
    rows = []
    for t in data["trades"]:
        price = t.get("price") or "100"
        rows.append(
            f"<li>{data['currency']} {t['amount']} @{price}% voor {t['adviser']}</li>"
        )

    return (
            "<ul style='padding-left:20px; margin-top:0; margin-bottom:0;'>"
            + "".join(rows)
            + "</ul>"
    )


def derive_cc_from_trades(trades: list) -> list:
    """Extract CC recipients from trade adviser data."""
    cc = set()
    for t in trades:
        helper = ADVISER_HELPERS.get(t["adviser"])
        if helper:
            cc.add(helper)
    return sorted(cc)


def get_product_info(product_name: str) -> dict:
    """Always return product info as dict with at least 'label'."""
    product_name_l = (product_name or "").lower()

    for key, info in PRODUCT_TYPES.items():
        key_l = str(key).lower()

        # Proper dict entry
        if isinstance(info, dict):
            label_l = str(info.get("label", "")).lower()
            if key_l in product_name_l or (label_l and label_l in product_name_l):
                return info

        # String entry → wrap immediately
        else:
            info_str = str(info)
            info_l = info_str.lower()
            if key_l in product_name_l or info_l in product_name_l:
                return {
                    "label": info_str,
                    "brochure_url": "",
                }

    # Default fallback to Index Garantie Note
    garantie = PRODUCT_TYPES.get("Index Garantie Note", {})
    if isinstance(garantie, dict):
        return garantie

    return {
        "label": str(garantie),
        "brochure_url": "",
    }


def build_body(data: dict) -> str:
    """Build complete HTML email body."""
    trades = data["trades"]
    adviser_count = len(trades)

    greet = greeting(trades)
    have_text = "Je hebt" if adviser_count == 1 else "Jullie hebben"

    # Build PRIIP URL
    eid_url = (
        f"{PRIIP_HUB_URL}"
        f"?documentType=pdf&identifier={data['isin']}"
        "&country=NL&language=NL&response=document"
    )

    rtd_list = build_rtd_list(data)
    vlk_code = data["vlk_code"]
    sender = get_sender_first_name() or "thijn"
    product = data["product"]


    # Get product-specific info
    product_info = get_product_info(product)
    product_name = product_info["label"]
    brochure = product_info["brochure_url"]

    # Build video HTML if available
    video_html = ""
    if product_info.get("video_url"):
        video_html = (
            f"Link naar {product_name} video: "
            f"<a href='{product_info['video_url']}'>{product_info['video_label']}</a><br>"
        )

    # Build complete HTML body
    body_html = f"""
<html>
<body style="font-family:Segoe UI; font-size:11pt">
<p>{greet}</p>

<p>{have_text} onlangs voor een klant het maatwerkproduct aangekocht.</p>

<p>
Hierbij de generieke {product_name} brochure, het inlegvel, de Final Terms
en een link naar het Essentiële-informatiedocument voor deze Maatwerk Note.
</p>

<p>
De Final Terms, inlegvel, de generieke {product_name} brochure en een link
naar het Essentiële-informatiedocument kunnen met de begeleidende brief
van Van Lanschot naar de klant worden gestuurd.
</p>

<div style="margin:0; padding:0;">
{video_html}
Link naar {product_name} brochure:
<a href="{brochure}">Brochure {product_name}</a><br>
Link naar het EID:
<a href="{eid_url}">{eid_url}</a>
</div>

<br>

<p>Graag onderstaande trades RTD inleggen:</p>
{rtd_list}

<br>

<p>De VL code is: <b>{vlk_code}</b></p>

<p>Groet,<br>{sender}</p>
</body>
</html>
"""

    return body_html