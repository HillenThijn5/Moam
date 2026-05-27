# sharepoint/parser.py
"""
Zet ruwe SharePoint-rijen om naar GUI-klare veld-dicts.

Titelopmaak:
  "YYYYMMDD - {maturity}Y {product_code} {underlying1} {underlying2?}"

Opmaak van comments:
  "EUR {size} @{price}% voor {client}\\nHedged EUR {amount} BtB met {party} @{upfront}%"

Opmaak van 'Van Lanschot Lange_x' (lange productnaam):
  Trigger/AC:  "VGN VLK TP 100-90-60 EU 26-31"  → trigger=100, coupon_barrier=90, redemption=60
  IGN:         "GN VLK IGN 95-100p EU 26-31"     → protection=95, participation=100
  IGN Capped:  "GN VLK IGNC 95-100-120 EU 26-31" → protection=95, participation=100, cap=120
  MCN:         "VGN VLK MCN 50-50 EU 26-31"      → coupon_barrier=50, redemption=50
  FRN:         "GN VLK FRN 2.85% 26-29"          → premie=2.85
"""
from __future__ import annotations

import re
from datetime import datetime, timedelta
from statics.loader import resolve_adviser_name

# Koppelt productcodes uit SharePoint-titels aan productnamen voor PC Mail
PRODUCT_CODE_MAP: dict[str, str] = {
    "AC":              "Trigger Plus Note",
    "TN":              "Trigger Plus Note",
    "TP":              "Trigger Plus Note",
    "TPN":             "Trigger Plus Note",
    "TRIGGER":         "Trigger Plus Note",
    "IGN":             "Index Garantie Note",
    "IGNC":            "Index Garantie Note Capped",
    "MCN":             "Memory Coupon",
    "FRN":             "Fixed Rate Note",
    "FIXED RATE NOTE": "Fixed Rate Note",
}

JURISDICTION_TO_CLIENT: dict[str, str] = {
    "netherlands": "VL NL",
    "belgium":     "VL Belgium",
    "switzerland": "VL Switzerland",
}

# Koppelt issuer-afkortingen in de lange naam aan volledige waarden uit ISSUERS
ISSUER_CODE_MAP: dict[str, str] = {
    "VLK":    "VLK",
    "BNP":    "BNP Paribas Issuance",
    "BNPP":   "BNP Paribas Issuance",
    "UBS":    "UBS AG",
    "SG":     "Société Générale",
    "SOCGEN": "Société Générale",
}

# Valutacodes die nooit als onderliggende waarden mogen worden gezien
KNOWN_CURRENCIES: set[str] = {
    "EUR", "USD", "GBP", "CHF", "JPY", "SEK", "NOK", "DKK", "AUD", "CAD",
}


# ── Bedraghelpers ─────────────────────────────────────────────────────────────

def _parse_amount(s: str) -> float:
    """Zet '300k', '1mio', '1.5 mio', '295' om naar een float in EUR."""
    s = s.strip().replace(",", "").replace(" ", "")
    m = re.match(r"^([\d.]+)(k|mio)?$", s, re.IGNORECASE)
    if not m:
        return 0.0
    n = float(m.group(1))
    suffix = (m.group(2) or "").lower()
    if suffix == "mio":
        return n * 1_000_000
    elif suffix == "k":
        return n * 1_000
    else:
        # Zonder suffix: behandel als duizenden bij gebruikelijke dealgroottes (< 10 000)
        return n * 1_000 if n < 10_000 else n


def _sum_amount_expr(expr: str) -> float:
    """
    Tel alle numerieke bedragen op in expressies zoals '100K+ EUR 250K + EUR 200K'.
    Verwijdert eerst ingesloten valutacodes.
    """
    clean = re.sub(r"\b(?:EUR|USD|GBP|CHF|JPY|SEK|NOK|DKK)\b", "", expr, flags=re.IGNORECASE)
    parts = re.findall(r"[\d,.]+\s*(?:k|mio)?", clean, re.IGNORECASE)
    return sum(_parse_amount(p) for p in parts)


def _format_amount(eur: float) -> str:
    """Formatteer 1 950 000 → '1.95 mio', 300 000 → '300k'."""
    if eur <= 0:
        return ""
    if eur >= 1_000_000:
        val = eur / 1_000_000
        formatted = f"{val:.2f}".rstrip("0").rstrip(".")
        return f"{formatted} mio"
    val = eur / 1_000
    formatted = f"{val:.0f}"
    return f"{formatted}k"


# ── Titelparser ───────────────────────────────────────────────────────────────

def parse_title(title: str) -> dict:
    """
    Zet bijv. '20260504 - 7Y AC SX5E SPX USD' om naar:
      {"maturity": "7Y", "product": "Trigger Plus Note", "underlyings": ["SX5E", "SPX"]}
    Valutacodes (USD, EUR, …) worden uit de underlyings gefilterd.
    """
    result: dict = {"maturity": "", "product": "", "underlyings": []}
    title = title.strip('"').strip()

    m = re.match(r"^(\d{8})\s*-\s*(.+)$", title)
    if not m:
        return result

    rest = m.group(2).strip()

    # Looptijd: bijv. "5Y", "7Y"
    mat_m = re.match(r"^(\d+Y)\s+(.+)$", rest, re.IGNORECASE)
    if mat_m:
        result["maturity"] = mat_m.group(1).upper()
        rest = mat_m.group(2).strip()

    # Productcode — probeer eerst de langste meerwoordmatch (bijv. "Fixed Rate Note")
    tokens = rest.split()
    product = ""
    underlyings_start = 0

    for n in range(min(len(tokens), 4), 0, -1):
        candidate = " ".join(tokens[:n]).upper()
        if candidate in PRODUCT_CODE_MAP:
            product = PRODUCT_CODE_MAP[candidate]
            underlyings_start = n
            break

    if not product and tokens:
        product = PRODUCT_CODE_MAP.get(tokens[0].upper(), "")
        underlyings_start = 1

    result["product"] = product
    # Filter valutacodes weg zodat "SX5E USD" → ["SX5E"]
    result["underlyings"] = [
        t for t in tokens[underlyings_start:]
        if t.upper() not in KNOWN_CURRENCIES
    ]
    return result


# ── Parser voor lange naam-parameters ─────────────────────────────────────────

def parse_long_name_params(long_name: str, product: str) -> dict:
    """
    Haal payoffparameters en issuer uit de lange VL-productnaam.

    Opmaak van de lange naam: "{GN|VGN} {ISSUER} {PRODUCT_CODE} {params} {region} {years}"
      bijv. "VGN VLK TP 100-90-60 EU 26-31"  →  issuer="VLK", param2=100, param3=90, param4=60

    Indeling per product (param-sleutels komen overeen met PRODUCT_PAYOFF_FIELDS):
      Trigger Plus Note  → param2=trigger/aflossing, param3=coupon_barrier, param4=redemption
      Index Garantie Note         → param1=protection, param2=participation
      Index Garantie Note Capped  → param1=protection, param2=participation, param3=cap
      Memory Coupon      → param3=coupon_barrier, param4=redemption
      Fixed Rate Note    → param1=premie (bijv. 2.85 uit "2.85%")
    """
    result: dict = {"issuer": "", "param1": "", "param2": "", "param3": "", "param4": ""}
    if not long_name or not product:
        return result

    ln = long_name.strip('"').strip()

    # Issuer: de lange naam begint met GN of VGN, gevolgd door de issuer-afkorting
    # bijv. "GN VLK IGN ..." of "VGN VLK TP ..."
    issuer_m = re.match(r"^V?GN\s+(\S+)", ln, re.IGNORECASE)
    if issuer_m:
        abbrev = issuer_m.group(1).upper()
        result["issuer"] = ISSUER_CODE_MAP.get(abbrev, "")

    NUM = r"(\d+(?:\.\d+)?)"   # vangt een getal met optionele decimalen

    if product == "Trigger Plus Note":
        # "100-90-60": trigger=100, coupon_barrier=90, redemption=60
        m = re.search(rf"{NUM}-{NUM}-{NUM}", ln)
        if m:
            result["param2"] = m.group(1)
            result["param3"] = m.group(2)
            result["param4"] = m.group(3)

    elif product == "Index Garantie Note":
        # "95-100p": protection=95, participation=100
        m = re.search(rf"{NUM}-{NUM}p?", ln, re.IGNORECASE)
        if m:
            result["param1"] = m.group(1)
            result["param2"] = m.group(2)

    elif product == "Index Garantie Note Capped":
        # "95-100-120": protection=95, participation=100, cap=120
        m = re.search(rf"{NUM}-{NUM}-{NUM}", ln)
        if m:
            result["param1"] = m.group(1)
            result["param2"] = m.group(2)
            result["param3"] = m.group(3)

    elif product == "Memory Coupon":
        # "50-50": coupon_barrier=50, redemption=50
        m = re.search(rf"{NUM}-{NUM}", ln)
        if m:
            result["param3"] = m.group(1)
            result["param4"] = m.group(2)

    elif product == "Fixed Rate Note":
        # "2.85%": premie=2.85
        m = re.search(rf"{NUM}%", ln)
        if m:
            result["param1"] = m.group(1)

    return result


# ── Parser voor comments ──────────────────────────────────────────────────────

# Adviseurregels: "{optionele valuta} {bedrag(en)} @{price}% {optionele kwalificatie} voor {name}"
# Ondersteunt:
#   EUR 300k @100% voor Jan vd Ven
#   EUR 650k @99.75 reoffer voor stijn
#   300k @99.75 voor Kurt                      ← zonder valutaprefix
#   EUR 100K+ EUR 250K + EUR 200K @ 99.8 voor Monique  ← gecombineerde bedragen
#   EUR 500k +150k @100% voor Remco            ← extra bedrag
#   EUR 360k ( 250 + 110) @99.75 voor Jean-Paul ← opmerking tussen haakjes
#   EUR 295 @ 100% done voor Rodney Maes       ← kwalificatie 'done'
_ADVISER_RE = re.compile(
    # Optionele valuta vooraan
    r"(?:(?:EUR|USD|GBP|CHF|JPY)\s+)?"
    # Hoofdbedrag + optionele aanvullingen (+150k, + EUR 250K, …)
    r"([\d,.]+\s*(?:k|mio)?"
    r"(?:\s*\+\s*(?:(?:EUR|USD|GBP|CHF|JPY)\s+)?[\d,.]+\s*(?:k|mio)?)*)"
    # Optionele toelichting tussen haakjes, bijv. (250 + 110)
    r"(?:\s*\([^)]*\))?"
    # Prijs: @ NNN [%]
    r"\s*@\s*([\d.]+)\s*%?"
    # Optionele kwalificatie: reoffer / re-offer / done (vastgelegd voor Belgische flow)
    r"(\s+(?:reoffer|re-offer|done))?"
    # Ontvanger
    r"\s+voor\s+([^\n]+?)(?=\s*\n|$|hedged?|btb\b|own\s+book)",
    re.IGNORECASE,
)

# Hedgepartij: "BtB met PARTY @RATE%" of "BtB with PARTY @RATE%"
_HEDGE_RE = re.compile(
    r"btb\s+(?:met|with)\s+([\w][\w\s&.']{0,30}?)\s*@\s*([\d.]+)\s*%?",
    re.IGNORECASE,
)

# Gehedged nominaal: "Hedged EUR NNN BtB" of "Hedged BTB with …"
_HEDGE_AMT_RE = re.compile(
    r"hedged?\s+(?:(?:EUR|USD|GBP|CHF|JPY)\s+)?([\d,.]+\s*(?:k|mio)?)\s+btb",
    re.IGNORECASE,
)


def parse_comments(comments: str) -> dict:
    """
    Haal uit vrije comments-tekst:
      - currency (eerste gevonden valutacode)
      - hedge_party + upfront-tarief
      - btb_amount
      - advisers: lijst met {"name", "amount", "price"} per handelsregel
      - total_sold: geformatteerde som van alle adviseurbedragen (bijv. "1.95 mio")
    """
    result: dict = {
        "currency": "", "hedge_party": "", "upfront": "", "btb_amount": "",
        "advisers": [], "total_sold": "",
    }
    if not comments:
        return result

    comments = comments.strip('"').strip()

    # Valuta: eerste valutacode in de tekst
    cur_m = re.search(r"\b(EUR|USD|GBP|CHF|JPY)\b", comments, re.IGNORECASE)
    if cur_m:
        result["currency"] = cur_m.group(1).upper()

    # Hedgepartij + upfront
    hedge_m = _HEDGE_RE.search(comments)
    if hedge_m:
        result["hedge_party"] = hedge_m.group(1).strip()
        result["upfront"]     = hedge_m.group(2).strip()

    # BtB-nominaal
    amt_m = _HEDGE_AMT_RE.search(comments)
    if amt_m:
        result["btb_amount"] = amt_m.group(1).strip()

    # Adviseurregels
    total_eur = 0.0
    for m in _ADVISER_RE.finditer(comments):
        name = m.group(4).strip().rstrip(",;")
        if not name or re.search(r"btb|hedged?", name, re.IGNORECASE):
            continue
        amount_expr = m.group(1).strip()
        price       = m.group(2).strip()
        qualifier   = (m.group(3) or "").strip().lower()
        is_reoffer  = qualifier in ("reoffer", "re-offer")
        eur = _sum_amount_expr(amount_expr)
        total_eur += eur
        result["advisers"].append({
            "name":      resolve_adviser_name(name),
            "amount":    _format_amount(eur) if eur else amount_expr,
            "price":     price,
            "reoffer":   is_reoffer,
        })

    if total_eur > 0:
        result["total_sold"] = _format_amount(total_eur)

    # Detecteer derdepartijcliënt op basis van trefwoorden in comments
    lower = comments.lower()
    if "alpha capital" in lower:
        result["client_override"] = "Alpha Capital Asset Management B.V."
    elif "oakk" in lower or "cfs" in lower:
        result["client_override"] = "OAKK Capital (CFS)"

    return result


# ── Datumhelper ───────────────────────────────────────────────────────────────

def _fmt_date(val) -> str:
    """
    Formatteer een date/datetime als DD/MM/YYYY voor de GUI.
    Excel bewaart SharePoint-UTC-datums; Amsterdam (CEST) = UTC+2,
    dus 22:00 UTC is lokaal middernacht — voeg 2 h toe voor de juiste datum.
    """
    if val is None:
        return ""
    if isinstance(val, datetime):
        if val.hour >= 20:
            val = val + timedelta(hours=2)
        return val.strftime("%d/%m/%Y")
    s = str(val).strip('"').strip()
    iso = re.match(r"^(\d{4})-(\d{2})-(\d{2})", s)
    if iso:
        return f"{iso.group(3)}/{iso.group(2)}/{iso.group(1)}"
    return s


def _issuer_fallback(raw: dict) -> str:
    """Detecteer de issuer wanneer de lange VL-naam ontbreekt.

    Als de deal een DIP/SNIP-nummer of een Series heeft, is het een VLK-uitgifte.
    """
    dip_snip = str(raw.get("DIP/SNIP") or "").strip('"').strip()
    series = str(raw.get("Series") or "").strip('"').strip()
    if dip_snip or series:
        return "VLK"
    return ""


# ── Hoofdparser voor deals ────────────────────────────────────────────────────

def parse_deal(raw: dict) -> dict:
    """Zet een ruwe SharePoint-rij om naar een GUI-klaar dict."""
    title     = str(raw.get("Title")                 or "").strip('"')
    comments  = str(raw.get("Comments")              or "").strip('"')
    long_name = str(raw.get("Van Lanschot Lange_x")  or "").strip('"')
    hedge_raw = str(raw.get("Hedgeparty")             or "").strip('"')

    tp = parse_title(title)
    cp = parse_comments(comments)
    pp = parse_long_name_params(long_name, tp["product"])

    # VLK-code vereist: SharePoint bewaart dit als Python-bool of als string "True"/"False"
    vlk_raw = raw.get("VanLanschot Code?")
    if isinstance(vlk_raw, bool):
        vlk_code = vlk_raw
    else:
        vlk_code = str(vlk_raw).strip('"').strip().lower() != "false"

    return {
        # ── weergave ─────────────────────────────────────────────────────
        "title":          title,
        "status":         str(raw.get("Status")   or "").strip('"'),
        # ── identiteit ───────────────────────────────────────────────────
        "series":         str(raw.get("Series")   or "").strip(),
        "isin":           str(raw.get("ISIN")      or "").strip('"'),
        "prospectus_type":str(raw.get("DIP/SNIP")  or "").strip('"'),
        "vlk_code":       vlk_code,
        "vl_code":        str(raw.get("VL CODE")   or "").strip('"').strip(),
        # ── datums (DD/MM/YYYY voor de GUI) ──────────────────────────────
        "issue_date":     _fmt_date(raw.get("Issue Date")),
        "trade_date":     _fmt_date(raw.get("DateofResolutions")),
        # ── product ──────────────────────────────────────────────────────
        "product":        tp["product"],
        "maturity":       tp["maturity"],
        "underlyings":    tp["underlyings"],
        # ── issuer + payoffparameters uit lange naam ─────────────────────
        "issuer":         pp["issuer"]
                          or _issuer_fallback(raw),
        "param1":         pp["param1"],
        "param2":         pp["param2"],
        "param3":         pp["param3"],
        "param4":         pp["param4"],
        # ── cliënt ───────────────────────────────────────────────────────
        "client":         cp.get("client_override")
                          or JURISDICTION_TO_CLIENT.get(
                              str(raw.get("Jurisdiction") or "").strip('"').lower(), ""
                          ),
        # ── valuta (uit comments-prefix) ─────────────────────────────────
        "currency":       cp["currency"],
        # ── hedge (kolom Hedgeparty indien aanwezig, anders uit comments) ─
        "hedge_party":    hedge_raw or cp["hedge_party"],
        "upfront":        cp["upfront"],
        "btb_amount":     cp["btb_amount"],
        # ── ruwe comments (getoond in dialoogvoorbeeld) ──────────────────
        "comments":       comments,
        # ── geparste adviseurs (voor Documentatie-mail) ──────────────────
        "advisers":       cp["advisers"],
        # ── total_sold (som van adviseurbedragen, voor PC Mail) ──────────
        "total_sold":     cp["total_sold"],
    }
