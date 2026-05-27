# statics/data.py
"""
Gecentraliseerde statische data voor alle mailgeneratoren.
Dit is de enige bron voor alle dropdownwaarden, URL's en constanten.
"""
import sys
from pathlib import Path
from statics.loader import (
    load_adviser_map, load_underlyings, load_product_urls, load_prospectus_urls,
)


def _exe_root() -> Path:
    """Projectroot: map van de exe als die draait, anders de bronroot van MoamProject."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent
    return Path(__file__).resolve().parents[1]


# ============================================================================
# PADEN
# ============================================================================

# Lokale projectkopie  gebruikt voor lezen; synchroon gehouden door de refreshlogica
SHAREPOINT_SUMMARY_PATH = str(
    _exe_root() / "sharepoint" / "sharepointsummary.xlsx"
)

# Met OneDrive gesynchroniseerde bron (automatisch bijgewerkt vanuit SharePoint).
# Refresh kopieert dit → SHAREPOINT_SUMMARY_PATH wanneer het bestaat.
SHAREPOINT_ONEDRIVE_PATH = str(
    Path.home() / "OneDrive - Van Lanschot Kempen" / "sharepointsummary.xlsx"
)

# ============================================================================
# PRODUCT- EN BEDRIJFSDATA
# ============================================================================

PRODUCT_TYPES = {
    pt: {
        "label":       pt if pt != "Memory Coupon" else "Memory Coupon Note",
        "brochure_url": load_product_urls().get(pt, {}).get("brochure_url", ""),
        "video_url":    load_product_urls().get(pt, {}).get("video_url"),
        "video_label":  load_product_urls().get(pt, {}).get("video_label"),
    }
    for pt in [
        "Trigger Plus Note",
        "Memory Coupon",
        "Index Garantie Note",
        "Index Garantie Note Capped",
        "Fixed Rate Note",
    ]
}
HEDGEPARTY = [
    "UBS",
    "ING Bank",
    "BNP",
    "Santander",
    "Morgan Stanley",
    "BBVA",
    "Société Générale",
    "Danske",
    "Goldman Sachs",
    "UniCredit",
    "Citi",
    "BofA",
    "J.P. Morgan",
    "KBC",
    "Leonteq",
]

PRODUCTS_NO_UNDERLYING = {"Fixed Rate Note"}


PRODUCT_PAYOFF_FIELDS = {
    #  key → (field_key, label) — volgorde is belangrijk voor de weergave
    #  param1–param4 zijn generieke velden; elk producttype geeft ze een eigen label
    "Trigger Plus Note": [
        ("param1", "Coupon (%)"),
        ("param2", "Aflossing (%)"),
        ("param3", "Coupon Barrier (%)"),
        ("param4", "Redemption Barrier (%)"),
    ],
    "Memory Coupon": [
        ("param1", "Coupon (%)"),
        ("param3", "Coupon Barrier (%)"),
        ("param4", "Redemption Barrier (%)"),
    ],
    "Index Garantie Note": [
        ("param1", "Protection (%)"),
        ("param2", "Participation (%)"),
    ],
    "Index Garantie Note Capped": [
        ("param1", "Protection (%)"),
        ("param2", "Participation (%)"),
        ("param3", "Cap (%)"),
        # Asianing (tail + obs) wordt afgehandeld door een knop in de GUI
    ],
    "Fixed Rate Note": [
        ("param1", "Premie (%)"),
    ],
}

# ============================================================================
# GEDEELDE DROPDOWNS
# ============================================================================

MATURITIES = ["1Y", "2Y", "3Y", "4Y", "5Y", "6Y", "7Y"]

CURRENCIES = [
    "EUR",
    "USD",
    "GBP",
]

ISSUERS = [
    "VLK",
    "BNP Paribas Issuance",
    "Société Générale",
]

CLIENTS = ["VL NL", "VL Belgium", "VL Switzerland", "ING Bank", "Alpha Capital Asset Management B.V.", "OAKK Capital (CFS)"]

# ============================================================================
# SPECIFIEK VOOR DOCUMENTATIE-MAIL
# ============================================================================
_ud = load_underlyings()
UNDERLYINGS      = _ud["list"]
ADVISER_HELPERS: dict[str, str] = load_adviser_map()

DOCUMENTATIE_MAIL_RECIPIENTS = {
    "to": [],
    "fixed_cc": [
        "Robbert van Riel",
        "Bob Stroeken",
        "StructuredInvestments",
    ],
}

# ============================================================================
# SPECIFIEK VOOR MARKETINGMAIL
# ============================================================================

MARKETING_PRODUCT_TYPES = [
    "TRIGGER",
    "MEMORY_COUPON",
    "INDEX_GARANTIE",
    "INDEX_GARANTIE_CAPPED",
]

# Leesbare namen voor marketingproducttypes (gebruikt in titels en Excel)
MARKETING_PRODUCT_TYPE_NAMES = {
    "TRIGGER":               "Trigger Plus Note",
    "INDEX_GARANTIE":        "Index Garantie Note",
    "INDEX_GARANTIE_CAPPED": "Index Garantie Note Capped",
    "MEMORY_COUPON":         "Memory Coupon Note",
}

# GUI-parameterregels voor het tabblad Marketing Mail (stuurt het dynamische formulier aan).
# param1–param4 zijn generieke velden; de labels hieronder verduidelijken de betekenis per producttype.
# param1 = coupon/premie/bescherming, param2 = trigger/participatie,
# param3 = couponbarrière/cap, param4 = bescherming/aflossingsbarrière
MARKETING_PARAMETER_CONFIG = {
    "TRIGGER": {
        "param1": {"label": "Premie %",              "required": True},
        "param2": {"label": "Aflossingsbarrierre %", "required": True},
        "param3": {"label": "Coupon Barrier %",      "required": True},
        "param4": {"label": "Protection %",          "required": True},
        "show_param3": True,
        "asianing": False,
    },
    "MEMORY_COUPON": {
        "param1": {"label": "Premie %",              "required": True},
        "param2": {"label": "Trigger %", "fixed": "n.v.t.", "required": False},
        "param3": {"label": "Coupon Barrier %",      "required": True},
        "param4": {"label": "Protection %",          "required": True},
        "show_param3": True,
        "asianing": False,
    },
    "INDEX_GARANTIE": {
        "param1": {"label": "Protection %",          "required": True},
        "param2": {"label": "Participation %",       "required": True},
        "show_param3": False,
        "asianing": True,
    },
    "INDEX_GARANTIE_CAPPED": {
        "param1": {"label": "Protection %",          "required": True},
        "param2": {"label": "Participation %",       "required": True},
        "param3": {"label": "Cap %",                 "required": True},
        "show_param3": True,
        "asianing": True,
    },
    "AUTOCALL": {
        "param1": {"label": "Premie %",              "required": True},
        "param2": {"label": "Trigger %",             "required": True},
        "param3": {"label": "Coupon Barrier %",      "required": True},
        "param4": {"label": "Protection %",          "required": True},
        "show_param3": True,
        "asianing": False,
    },
}

MARKETING_ISSUERS = [
    "Van Lanschot Kempen N.V. (Fitch: A- / S&P: BBB+)",
    "BNP Paribas Issuance B.V. (Moody's: Aa3 / Fitch: AA- / S&P: A+)",
    "UBS AG (Moody's: Aa3 / S&P: A+ / Fitch: AA-)",
    "SG Issuer, gewaarborgd door Société Générale (Moody's: A1 / S&P: A / Fitch: A)",
]

MARKETING_MAIL_SIGNATURE = "Van Lanschot Kempen Structured Products"


# ============================================================================
# INCREASE / DECREASE-MAIL
# ============================================================================

ID_MAIL_DB_CONN_STR = (
    "DRIVER={ODBC Driver 17 for SQL Server};"
    "SERVER=sql_structuredproducts.fvlprod.fvl;"
    "DATABASE=StructuredProducts2010;"
    "Trusted_Connection=yes;"
    "Encrypt=yes;"
    "TrustServerCertificate=yes;"
)

ID_MAIL_TO = (
    "mtnprogramme@vanlanschot.com; "
    "financialaccounting@vanlanschot.com; "
    "treasuryadm@vanlanschot.com"
)

ID_MAIL_CC = (
    "structuredinvestments@kempen.com; "
    "Riskmanagement@vanlanschotkempen.com; "
    "SAS@kempen.nl; "
    "ReconOs@kempen.nl"
)


# ============================================================================
# API-ENDPOINTS EN EXTERNE URL'S
# ============================================================================

PRIIP_HUB_URL = "https://api.priiphub.com/hub/priip/document"

# ============================================================================
# E-MAILONTVANGERS
# ============================================================================

# Gedeelde CC gebruikt door Marketing Mail en PC Mail
STRUCTURED_INVESTMENTS_EMAIL = "structuredinvestments@vanlanschotkempen.com"

# Marketing Mail — TO-lijst (alle adviseurs/distributie)
MARKETING_MAIL_TO_LIST = [
    "f.b.vaniersel@vanlanschotkempen.com",
    "M.A.J.vandenBroek@vanlanschotkempen.com",
    "K.Seegers@vanlanschotkempen.com",
    "J.Olthof@vanlanschotkempen.com",
    "f.jongepier@vanlanschotkempen.com",
    "P.Arts@vanlanschotkempen.com",
    "e.tjittes@vanlanschotkempen.com",
    "T.H.Sie@vanlanschotkempen.com",
    "m.eeken@vanlanschotkempen.com",
    "A.Bakx@vanlanschotkempen.com",
    "E.D.vanBasten@vanlanschotkempen.com",
    "Y.Doets@vanlanschotkempen.com",
    "beleggingsadvies@vanlanschotkempen.com",
    "j.verhulst@vanlanschotkempen.com",
    "S.Heijdra@vanlanschotkempen.com",
    "h.gruijters@vanlanschotkempen.com",
    "D.vandenBerg@vanlanschotkempen.com",
    "k.droogleeverfortuyn@vanlanschotkempen.com",
    "c.weststrate@vanlanschotkempen.com",
    "a.j.m.vandijk@vanlanschotkempen.com",
    "R.Folkertsma@vanlanschotkempen.com",
    "Huib.Meulenbelt@kempen.com",
    "Tjerk.Bosse@kempen.com",
    "e.vanbraam@vanlanschotkempen.com",
    "J.vanHofwegen@vanlanschotkempen.com",
    "Jaap.Roggeveen@vanlanschotkempen.com",
    "marc.boom@vanlanschotkempen.com",
    "jan.vandeven@vanlanschotkempen.com",
    "m.lof@vanlanschotkempen.com",
    "A.A.vandenNieuwenhuizen@vanlanschotkempen.com",
    "J.vanderHeijden@vanlanschotkempen.com",
    "M.vanderFeltz@vanlanschotkempen.com",
    "B.C.H.Smits@vanlanschotkempen.com",
    "W.Bouhuijs@vanlanschotkempen.com",
    "etienne.vanderwerf@vanlanschotkempen.com",
    "j.vandooremalen@vanlanschotkempen.com",
    "J.Klerks@vanlanschotkempen.com",
    "gerben.debree@vanlanschotkempen.com",
    "t.buisman@vanlanschotkempen.com",
    "beleggingsadviesnltotaal@vanlanschotkempen.com",
    "J.P.vanWijngaarden@vanlanschotkempen.com",
    "b.stroeken@vanlanschotkempen.com",
    "r.vanriel@vanlanschotkempen.com",
    "e.devries@vanlanschotkempen.com",
    "e.lamberty@vanlanschotkempen.com",
]

# PC Mail — TO-lijst (interne ops-/risk-/settlementteams; gescheiden met puntkomma's voor Outlook)
PC_MAIL_TO_RECIPIENTS = (
    "sas@vanlanschotkempen.com; reconos@vanlanschotkempen.com; riskmanagement@vanlanschotkempen.com; "
    "i.sucur@vanlanschotkempen.com; e.lamberty@vanlanschotkempen.com; m.bloem@vanlanschotkempen.com; "
    "f.christensentsu@vanlanschotkempen.com; r.vanbetteraij@vanlanschotkempen.com; s.molloy@vanlanschotkempen.com"
)

# ============================================================================
# UNDERLYINGNAMEN EN ALIASES
# ============================================================================

UNDERLYING_FULL_NAMES = _ud["full_names"]
UNDERLYING_ALIASES    = _ud["aliases"]

# ============================================================================
# PC MAIL-DOCUMENTLINKS (opgebouwd uit tabblad ProductURLs in static_sheet.xlsx)
# ============================================================================
_pu = load_product_urls()

PC_MAIL_BROCHURE_URLS: dict[str, tuple[str, str]] = {
    pt.upper(): (d["brochure_label"], d["brochure_url"])
    for pt, d in _pu.items()
    if d.get("brochure_url")
}

PC_MAIL_VIDEO_URLS: dict[str, tuple[str, str]] = {
    pt.upper(): (d["video_label_pc"], d["video_url"])
    for pt, d in _pu.items()
    if d.get("video_url") and d.get("video_label_pc")
}

PROSPECTUS_URLS: dict[str, tuple[str, str]] = load_prospectus_urls()

