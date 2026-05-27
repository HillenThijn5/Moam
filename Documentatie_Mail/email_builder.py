# Documentatie_Mail/email_builder.py
"""
Bouw het onderwerp en de HTML-inhoud voor de Documentatie Mail.
Regelt productspecifieke inhoud en opmaak.
"""
from datetime import date, datetime
from statics.data import PRODUCT_TYPES, ADVISER_HELPERS, PRIIP_HUB_URL
from MarketingMail.product_title import UNDERLYING_ALIASES
from statics.outlook_user import get_sender_first_name

_DUTCH_MONTHS = [
    "Januari", "Februari", "Maart", "April", "Mei", "Juni",
    "Juli", "Augustus", "September", "Oktober", "November", "December",
]


def _fmt_date_dutch(date_str: str) -> str:
    """Zet DD/MM/YYYY om naar het Nederlandse formaat '6 Mei 2026'."""
    if not date_str:
        return ""
    try:
        dt = datetime.strptime(date_str.strip(), "%d/%m/%Y")
        return f"{dt.day} {_DUTCH_MONTHS[dt.month - 1]} {dt.year}"
    except ValueError:
        return date_str


def build_subject(data: dict) -> str:
    """
    Bouw de onderwerpregel op basis van de email-data.
    Formaat: Documentatie {Issuer} {Product} {Underlyings} {StartYear} - {EndYear} – {ISIN}
    """
    today = data.get("today") or date.today()
    year_start = today.year % 100

    # Haal het aantal looptijdjaren uit een formaat als "2Y", "3Y", enz.
    maturity_str = data["maturity"].replace("Y", "").replace("y", "")
    try:
        maturity_num = int(maturity_str)
    except ValueError:
        maturity_num = 0

    year_end = year_start + maturity_num

    # Voeg underlyings samen met aliassen (bijv. SX5E → Eurozone)
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
    """Maak een passende aanhef op basis van het aantal advisers."""
    if len(trades) == 1:
        first_name = trades[0]["adviser"].split(" ")[0]
        return f"Hi {first_name},"
    return "Hi all,"


def build_rtd_list(data: dict) -> str:
    """Bouw een HTML-lijst van trades/RTD-regels."""
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
    """Haal CC-ontvangers uit de adviser-data van trades."""
    cc = set()
    for t in trades:
        helper = ADVISER_HELPERS.get(t["adviser"])
        if helper:
            cc.add(helper)
    return sorted(cc)


def get_product_info(product_name: str) -> dict:
    """Geef productinfo altijd terug als dict met minstens 'label'."""
    product_name_l = (product_name or "").lower()

    for key, info in PRODUCT_TYPES.items():
        key_l = str(key).lower()

        # Netjes dict-item
        if isinstance(info, dict):
            label_l = str(info.get("label", "")).lower()
            if key_l in product_name_l or (label_l and label_l in product_name_l):
                return info

        # Stringwaarde → meteen inpakken
        else:
            info_str = str(info)
            info_l = info_str.lower()
            if key_l in product_name_l or info_l in product_name_l:
                return {
                    "label": info_str,
                    "brochure_url": "",
                }

    # Standaard terugvaloptie naar Index Garantie Note
    garantie = PRODUCT_TYPES.get("Index Garantie Note", {})
    if isinstance(garantie, dict):
        return garantie

    return {
        "label": str(garantie),
        "brochure_url": "",
    }


def build_body(data: dict) -> str:
    """Bouw de volledige HTML-inhoud van de mail."""
    trades = data["trades"]
    adviser_count = len(trades)

    greet = greeting(trades)
    have_text = "Je hebt" if adviser_count == 1 else "Jullie hebben"

    # Bouw de PRIIP-URL op
    eid_url = (
        f"{PRIIP_HUB_URL}"
        f"?documentType=pdf&identifier={data['isin']}"
        "&country=NL&language=NL&response=document"
    )

    rtd_list = build_rtd_list(data)
    vlk_code = data["vlk_code"]
    sender = get_sender_first_name() or "thijn"
    product = data["product"]


    # Haal productspecifieke info op
    product_info = get_product_info(product)
    product_name = product_info["label"]
    brochure = product_info["brochure_url"]

    # Bouw de video-HTML op als die er is
    video_html = ""
    if product_info.get("video_url"):
        video_html = (
            f"Link naar {product_name} video: "
            f"<a href='{product_info['video_url']}'>{product_info['video_label']}</a><br>"
        )

    # Bouw de volledige HTML-inhoud op
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


def build_body_belgium(data: dict) -> str:
    """Bouw de HTML-inhoud voor Belgische deals (Mercier VL)."""
    trades = data["trades"]
    greet_names = []
    for t in trades:
        first = t["adviser"].split(" ")[0]
        if first not in greet_names:
            greet_names.append(first)
    greet_names.append("Business Support Mid Office")
    greet = f"Hi {', '.join(greet_names)},"

    currency = data["currency"]
    maturity = data["maturity"]
    issuer = data["issuer"]
    product = data["product"]
    product_label = f"{currency} {maturity} {issuer} {product}"

    isin = data["isin"]
    vlk_code = data["vlk_code"]
    issue_date = data.get("issue_date", "")
    sender = get_sender_first_name() or "thijn"

    # Bouw tradedetails op — één blok per trade
    trade_blocks = []
    for t in trades:
        amount = t["amount"]
        price = t.get("price", "100")

        lines = [
            f"<b>{product_label}</b>",
            f"ISIN: {isin}",
            f"VL code: {vlk_code}",
            f"Nominaal bedrag: {amount}",
        ]

        # Toon de Reoffer-regel als price != 100%
        price_f = float(price) if price else 100.0
        if price_f != 100.0:
            vergoeding = round(100.0 - price_f, 2)
            # Formatteer zonder nullen aan het eind: 0.25 en niet 0.250
            vergoeding_str = f"{vergoeding:g}"
            lines.append(
                f"Reoffer: {price}% "
                f"(issue price voor eindklant = 100%; "
                f"dus {vergoeding_str}% upfront vergoeding voor Mercier VL)"
            )

        if issue_date:
            lines.append(f"Settlement date: &nbsp;{_fmt_date_dutch(issue_date)}")

        trade_blocks.append("<br>\n".join(lines))

    trades_html = "<br><br>\n".join(trade_blocks)

    body_html = f"""
<html>
<body style="font-family:Segoe UI; font-size:11pt">
<p>{greet}</p>

<p>Bijgevoegd de Final Terms.</p>

<p>Zouden jullie onderstaande trade willen inleggen:</p>

<p>
{trades_html}
</p>

<p>Groet,</p>
<p>{sender}</p>
</body>
</html>
"""
    return body_html


def is_belgian(data: dict) -> bool:
    """Kijk of een deal Belgisch is op basis van het client-veld."""
    return (data.get("client") or "").strip().lower() == "vl belgium"