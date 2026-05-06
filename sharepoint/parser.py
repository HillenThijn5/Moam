# sharepoint/parser.py
"""
Parses raw SharePoint rows into GUI-ready field dicts.

Title format:
  "YYYYMMDD - {maturity}Y {product_code} {underlying1} {underlying2?}"

Comments format:
  "EUR {size} @{price}% voor {client}\\nHedged EUR {amount} BtB met {party} @{upfront}%"

Van Lanschot Lange_x (long product name) format:
  Trigger/AC:  "VGN VLK TP 100-90-60 EU 26-31"  → trigger=100, coupon_barrier=90, redemption=60
  IGN:         "GN VLK IGN 95-100p EU 26-31"     → protection=95, participation=100
  IGN Capped:  "GN VLK IGNC 95-100-120 EU 26-31" → protection=95, participation=100, cap=120
  MCN:         "VGN VLK MCN 50-50 EU 26-31"      → coupon_barrier=50, redemption=50
  FRN:         "GN VLK FRN 2.85% 26-29"          → premie=2.85
"""
from __future__ import annotations

import re
from datetime import datetime, timedelta

# Maps SharePoint title product codes → PC Mail product names
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

# Maps issuer abbreviations in the long name → full ISSUERS list values
ISSUER_CODE_MAP: dict[str, str] = {
    "VLK":    "VLK",
    "BNP":    "BNP Paribas Issuance",
    "BNPP":   "BNP Paribas Issuance",
    "UBS":    "UBS AG",
    "SG":     "Société Générale",
    "SOCGEN": "Société Générale",
}

# Currency codes that must never be treated as underlyings
KNOWN_CURRENCIES: set[str] = {
    "EUR", "USD", "GBP", "CHF", "JPY", "SEK", "NOK", "DKK", "AUD", "CAD",
}


# ── Amount helpers ────────────────────────────────────────────────────────────

def _parse_amount(s: str) -> float:
    """Convert '300k', '1mio', '1.5 mio', '295' → float in EUR."""
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
        # No suffix: treat as thousands for typical deal sizes (< 10 000)
        return n * 1_000 if n < 10_000 else n


def _sum_amount_expr(expr: str) -> float:
    """
    Sum all numeric amounts in expressions like '100K+ EUR 250K + EUR 200K'.
    Strips embedded currency symbols first.
    """
    clean = re.sub(r"\b(?:EUR|USD|GBP|CHF|JPY|SEK|NOK|DKK)\b", "", expr, flags=re.IGNORECASE)
    parts = re.findall(r"[\d,.]+\s*(?:k|mio)?", clean, re.IGNORECASE)
    return sum(_parse_amount(p) for p in parts)


def _format_amount(eur: float) -> str:
    """Format 1 950 000 → '1.95 mio', 300 000 → '300k'."""
    if eur <= 0:
        return ""
    if eur >= 1_000_000:
        val = eur / 1_000_000
        formatted = f"{val:.2f}".rstrip("0").rstrip(".")
        return f"{formatted} mio"
    val = eur / 1_000
    formatted = f"{val:.0f}"
    return f"{formatted}k"


# ── Title parser ──────────────────────────────────────────────────────────────

def parse_title(title: str) -> dict:
    """
    Parse e.g. '20260504 - 7Y AC SX5E SPX USD' into:
      {"maturity": "7Y", "product": "Trigger Plus Note", "underlyings": ["SX5E", "SPX"]}
    Currency codes (USD, EUR, …) are filtered out of underlyings.
    """
    result: dict = {"maturity": "", "product": "", "underlyings": []}
    title = title.strip('"').strip()

    m = re.match(r"^(\d{8})\s*-\s*(.+)$", title)
    if not m:
        return result

    rest = m.group(2).strip()

    # Maturity: e.g. "5Y", "7Y"
    mat_m = re.match(r"^(\d+Y)\s+(.+)$", rest, re.IGNORECASE)
    if mat_m:
        result["maturity"] = mat_m.group(1).upper()
        rest = mat_m.group(2).strip()

    # Product code — try longest multi-word match first (e.g. "Fixed Rate Note")
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
    # Filter out currency codes so "SX5E USD" → ["SX5E"]
    result["underlyings"] = [
        t for t in tokens[underlyings_start:]
        if t.upper() not in KNOWN_CURRENCIES
    ]
    return result


# ── Long-name param parser ────────────────────────────────────────────────────

def parse_long_name_params(long_name: str, product: str) -> dict:
    """
    Extract payoff parameters and issuer from the VL long product name.

    Long name format: "{GN|VGN} {ISSUER} {PRODUCT_CODE} {params} {region} {years}"
      e.g. "VGN VLK TP 100-90-60 EU 26-31"  →  issuer="VLK", param2=100, param3=90, param4=60

    Mapping per product (param keys match PRODUCT_PAYOFF_FIELDS):
      Trigger Plus Note  → param2=trigger/aflossing, param3=coupon_barrier, param4=redemption
      Index Garantie Note         → param1=protection, param2=participation
      Index Garantie Note Capped  → param1=protection, param2=participation, param3=cap
      Memory Coupon      → param3=coupon_barrier, param4=redemption
      Fixed Rate Note    → param1=premie (e.g. 2.85 from "2.85%")
    """
    result: dict = {"issuer": "", "param1": "", "param2": "", "param3": "", "param4": ""}
    if not long_name or not product:
        return result

    ln = long_name.strip('"').strip()

    # Issuer: long name starts with GN or VGN, then issuer abbreviation
    # e.g. "GN VLK IGN ..." or "VGN VLK TP ..."
    issuer_m = re.match(r"^V?GN\s+(\S+)", ln, re.IGNORECASE)
    if issuer_m:
        abbrev = issuer_m.group(1).upper()
        result["issuer"] = ISSUER_CODE_MAP.get(abbrev, "")

    NUM = r"(\d+(?:\.\d+)?)"   # captures a number with optional decimal

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


# ── Comments parser ───────────────────────────────────────────────────────────

# Adviser lines: "{optional currency} {amount(s)} @{price}% {optional qualifier} voor {name}"
# Handles:
#   EUR 300k @100% voor Jan vd Ven
#   EUR 650k @99.75 reoffer voor stijn
#   300k @99.75 voor Kurt                      ← no currency prefix
#   EUR 100K+ EUR 250K + EUR 200K @ 99.8 voor Monique  ← combined amounts
#   EUR 500k +150k @100% voor Remco            ← add-on amount
#   EUR 360k ( 250 + 110) @99.75 voor Jean-Paul ← bracketed note
#   EUR 295 @ 100% done voor Rodney Maes       ← 'done' qualifier
_ADVISER_RE = re.compile(
    # Optional leading currency
    r"(?:(?:EUR|USD|GBP|CHF|JPY)\s+)?"
    # Primary amount + optional add-ons (+150k, + EUR 250K, …)
    r"([\d,.]+\s*(?:k|mio)?"
    r"(?:\s*\+\s*(?:(?:EUR|USD|GBP|CHF|JPY)\s+)?[\d,.]+\s*(?:k|mio)?)*)"
    # Optional bracketed clarification e.g. (250 + 110)
    r"(?:\s*\([^)]*\))?"
    # Price: @ NNN [%]
    r"\s*@\s*([\d.]+)\s*%?"
    # Optional qualifier: reoffer / re-offer / done
    r"(?:\s+(?:reoffer|re-offer|done))?"
    # Recipient
    r"\s+voor\s+([^\n]+?)(?=\s*\n|$|hedged?|btb\b|own\s+book)",
    re.IGNORECASE,
)

# Hedge party: "BtB met PARTY @RATE%"  or  "BtB with PARTY @RATE%"
_HEDGE_RE = re.compile(
    r"btb\s+(?:met|with)\s+([\w][\w\s&.']{0,30}?)\s*@\s*([\d.]+)\s*%?",
    re.IGNORECASE,
)

# Hedged notional: "Hedged EUR NNN BtB"  or  "Hedged BTB with …"
_HEDGE_AMT_RE = re.compile(
    r"hedged?\s+(?:(?:EUR|USD|GBP|CHF|JPY)\s+)?([\d,.]+\s*(?:k|mio)?)\s+btb",
    re.IGNORECASE,
)


def parse_comments(comments: str) -> dict:
    """
    Extract from free-text comments:
      - currency (first currency code seen)
      - hedge party + upfront rate
      - btb_amount
      - advisers: list of {"name", "amount", "price"} per trade line
      - total_sold: formatted sum of all adviser amounts (e.g. "1.95 mio")
    """
    result: dict = {
        "currency": "", "hedge_party": "", "upfront": "", "btb_amount": "",
        "advisers": [], "total_sold": "",
    }
    if not comments:
        return result

    comments = comments.strip('"').strip()

    # Currency: first currency code in the text
    cur_m = re.search(r"\b(EUR|USD|GBP|CHF|JPY)\b", comments, re.IGNORECASE)
    if cur_m:
        result["currency"] = cur_m.group(1).upper()

    # Hedge party + upfront
    hedge_m = _HEDGE_RE.search(comments)
    if hedge_m:
        result["hedge_party"] = hedge_m.group(1).strip()
        result["upfront"]     = hedge_m.group(2).strip()

    # BtB notional
    amt_m = _HEDGE_AMT_RE.search(comments)
    if amt_m:
        result["btb_amount"] = amt_m.group(1).strip()

    # Adviser entries
    total_eur = 0.0
    for m in _ADVISER_RE.finditer(comments):
        name = m.group(3).strip().rstrip(",;")
        if not name or re.search(r"btb|hedged?", name, re.IGNORECASE):
            continue
        amount_expr = m.group(1).strip()
        price       = m.group(2).strip()
        eur = _sum_amount_expr(amount_expr)
        total_eur += eur
        result["advisers"].append({
            "name":   name,
            "amount": _format_amount(eur) if eur else amount_expr,
            "price":  price,
        })

    if total_eur > 0:
        result["total_sold"] = _format_amount(total_eur)

    return result


# ── Date helper ───────────────────────────────────────────────────────────────

def _fmt_date(val) -> str:
    """
    Format a date/datetime as DD/MM/YYYY for the GUI.
    Excel stores SharePoint UTC dates; Amsterdam (CEST) = UTC+2,
    so 22:00 UTC is midnight local — add 2 h to get the correct date.
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


# ── Main deal parser ──────────────────────────────────────────────────────────

def parse_deal(raw: dict) -> dict:
    """Convert a raw SharePoint row into a GUI-ready dict."""
    title     = str(raw.get("Title")                 or "").strip('"')
    comments  = str(raw.get("Comments")              or "").strip('"')
    long_name = str(raw.get("Van Lanschot Lange_x")  or "").strip('"')
    hedge_raw = str(raw.get("Hedgeparty")             or "").strip('"')

    tp = parse_title(title)
    cp = parse_comments(comments)
    pp = parse_long_name_params(long_name, tp["product"])

    # VLK code required: SharePoint stores as Python bool or string "True"/"False"
    vlk_raw = raw.get("VanLanschot Code?")
    if isinstance(vlk_raw, bool):
        vlk_code = vlk_raw
    else:
        vlk_code = str(vlk_raw).strip('"').strip().lower() != "false"

    return {
        # ── display ──────────────────────────────────────────────────────
        "title":          title,
        "status":         str(raw.get("Status")   or "").strip('"'),
        # ── identity ─────────────────────────────────────────────────────
        "series":         str(raw.get("Series")   or "").strip(),
        "isin":           str(raw.get("ISIN")      or "").strip('"'),
        "prospectus_type":str(raw.get("DIP/SNIP")  or "").strip('"'),
        "vlk_code":       vlk_code,
        "vl_code":        str(raw.get("VL CODE")   or "").strip('"').strip(),
        # ── dates (DD/MM/YYYY for the GUI) ────────────────────────────────
        "issue_date":     _fmt_date(raw.get("Issue Date")),
        "trade_date":     _fmt_date(raw.get("DateofResolutions")),
        # ── product ───────────────────────────────────────────────────────
        "product":        tp["product"],
        "maturity":       tp["maturity"],
        "underlyings":    tp["underlyings"],
        # ── issuer + payoff params from long name ─────────────────────────
        "issuer":         pp["issuer"],
        "param1":         pp["param1"],
        "param2":         pp["param2"],
        "param3":         pp["param3"],
        "param4":         pp["param4"],
        # ── client ────────────────────────────────────────────────────────
        "client":         JURISDICTION_TO_CLIENT.get(
                              str(raw.get("Jurisdiction") or "").strip('"').lower(), ""
                          ),
        # ── currency (from comments prefix) ───────────────────────────────
        "currency":       cp["currency"],
        # ── hedge (Hedgeparty col if present, else parse from comments) ───
        "hedge_party":    hedge_raw or cp["hedge_party"],
        "upfront":        cp["upfront"],
        "btb_amount":     cp["btb_amount"],
        # ── raw comments (shown in dialog preview) ────────────────────────
        "comments":       comments,
        # ── parsed advisers (for Documentatie mail) ───────────────────────
        "advisers":       cp["advisers"],
        # ── total sold (sum of adviser amounts, for PC Mail) ──────────────
        "total_sold":     cp["total_sold"],
    }
