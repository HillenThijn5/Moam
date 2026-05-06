# statics/data.py
"""
Centralized static data for all mail generators.
This is the single source of truth for all dropdown values, URLs, and constants.
"""
import sys
from pathlib import Path
from statics.loader import (
    load_adviser_map, load_underlyings, load_product_urls, load_prospectus_urls,
)


def _exe_root() -> Path:
    """Project root: exe's folder when frozen, otherwise MoamProject source root."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent
    return Path(__file__).resolve().parents[1]


# ============================================================================
# PATHS
# ============================================================================

# Local project copy — used for reading; kept in sync by the refresh logic
SHAREPOINT_SUMMARY_PATH = str(
    _exe_root() / "sharepoint" / "sharepointsummary.xlsx"
)

# OneDrive-synced source (auto-updated from SharePoint).
# Refresh copies this → SHAREPOINT_SUMMARY_PATH when it exists.
SHAREPOINT_ONEDRIVE_PATH = str(
    Path.home() / "OneDrive - Van Lanschot Kempen" / "sharepointsummary.xlsx"
)

# ============================================================================
# PRODUCT & BUSINESS DATA
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
    #  key → (field_key, label)  — order matters for display
    #  param1–param4 are generic slots; each product type assigns its own label
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
        # Asianing (tail + obs) handled by button in the GUI
    ],
    "Fixed Rate Note": [
        ("param1", "Premie (%)"),
    ],
}

# ============================================================================
# SHARED DROPDOWNS
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

CLIENTS = ["VL NL", "VL Belgium", "VL Switzerland", "ING Bank", "Alpha Capital"]

# ============================================================================
# DOCUMENTATIE MAIL SPECIFIC
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
# MARKETING MAIL SPECIFIC
# ============================================================================

MARKETING_PRODUCT_TYPES = [
    "TRIGGER",
    "MEMORY_COUPON",
    "INDEX_GARANTIE",
    "INDEX_GARANTIE_CAPPED",
]

# Human-readable names for marketing product types (used in titles and Excel)
MARKETING_PRODUCT_TYPE_NAMES = {
    "TRIGGER":               "Trigger Plus Note",
    "INDEX_GARANTIE":        "Index Garantie Note",
    "INDEX_GARANTIE_CAPPED": "Index Garantie Note Capped",
    "MEMORY_COUPON":         "Memory Coupon Note",
}

# GUI parameter rows for Marketing Mail tab (drives dynamic form).
# param1–param4 are generic slots; labels below clarify the meaning per product type.
# param1 = coupon/premie/protection, param2 = trigger/participation,
# param3 = coupon barrier/cap, param4 = protection/redemption barrier
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
# INCREASE / DECREASE MAIL
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
# API ENDPOINTS & EXTERNAL URLS
# ============================================================================

PRIIP_HUB_URL = "https://api.priiphub.com/hub/priip/document"

# ============================================================================
# EMAIL RECIPIENTS
# ============================================================================

# Shared CC used by Marketing Mail and PC Mail
STRUCTURED_INVESTMENTS_EMAIL = "structuredinvestments@vanlanschotkempen.com"

# Marketing Mail — TO list (all advisers/distribution)
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

# PC Mail — TO list (internal ops/risk/settlement teams; semicolon-separated for Outlook)
PC_MAIL_TO_RECIPIENTS = (
    "sas@vanlanschotkempen.com; reconos@vanlanschotkempen.com; riskmanagement@vanlanschotkempen.com; "
    "i.sucur@vanlanschotkempen.com; e.lamberty@vanlanschotkempen.com; m.bloem@vanlanschotkempen.com; "
    "f.christensentsu@vanlanschotkempen.com; r.vanbetteraij@vanlanschotkempen.com; s.molloy@vanlanschotkempen.com"
)

# ============================================================================
# UNDERLYING NAMES & ALIASES
# ============================================================================

UNDERLYING_FULL_NAMES = _ud["full_names"]
UNDERLYING_ALIASES    = _ud["aliases"]

# ============================================================================
# PC MAIL DOCUMENT LINKS  (built from ProductURLs sheet in static_sheet.xlsx)
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

