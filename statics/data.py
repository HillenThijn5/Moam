# statics/data.py
"""
Centralized static data for all mail generators.
This is the single source of truth for all dropdown values, URLs, and constants.
"""
from pathlib import Path
from statics.loader import load_adviser_map

# ============================================================================
# PATHS
# ============================================================================

# Local project copy — used for reading; kept in sync by the refresh logic
SHAREPOINT_SUMMARY_PATH = str(
    Path(__file__).resolve().parent.parent / "sharepoint" / "sharepointsummary.xlsx"
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
    "Trigger Plus Note": {
        "label": "Trigger Plus Note",
        "brochure_url": "https://markets.vanlanschotkempen.com/file/GetFile/uploadedfiles/Documents/FvL%2025328%20BRO%20SP%20Trigger%20Notes.pdf",
        "video_url": "https://vimeo.com/765679154",
        "video_label": "Hoe werkt een Trigger note?",
    },
    "Memory Coupon": {
        "label": "Memory Coupon Note",
        "brochure_url": "https://markets.vanlanschotkempen.com/file/GetFile/uploadedfiles/Documents/FvL%2025329%20BRO%20Memory%20Coupon%20Notes.pdf",
        "video_url": None,
        "video_label": None,
    },
    "Index Garantie Note": {
        "label": "Index Garantie Note",
        "brochure_url": "https://markets.vanlanschotkempen.com/file/GetFile/uploadedfiles/Documents/FvL%2025326%20BRO%20Index%20Garantie%20Notes.pdf",
        "video_url": "https://vimeo.com/765678480",
        "video_label": "Hoe werkt een Garantie note?",
    },
    "Index Garantie Note Capped": {
        "label": "Index Garantie Note Capped",
        "brochure_url": "https://markets.vanlanschotkempen.com/file/GetFile/uploadedfiles/Documents/FvL%2025326%20BRO%20Index%20Garantie%20Notes.pdf",
        "video_url": "https://vimeo.com/765678480",
        "video_label": "Hoe werkt een Garantie note?",
    },
    "Fixed Rate Note": {
        "label": "Fixed Rate Note",
        "brochure_url": "",
        "video_url": None,
        "video_label": None,
    },
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
UNDERLYINGS = [
    "SD3E",
    "SX5E",
    "SEESGSEP",
    "VWO US",
    "AEX",
    "SPX",
    "SHELL NA",
    "SCXP",
    "EPRA",
    "SXEP",
    "SXKP",
    "EEM UP",
    "EEM US",
    "SDGP",
    "SOXX US",
    "SX7P",
    "SOLEEE",
    "KOSPI2",
    "RFESGDP",
    "SXXP",
    "SX5EESG",
    "SX6P",
    "SIMSCI",
    "RTY",
    "NKY",
    "HSCEI",
    "SMI",
    "HSI",
    "UKX",
    "TAMSCI",
    "SELRE",
    "EWZ UP",
    "SUBE",
    "USSW10-USSW2",
    "EIISDA30",
    "EIISDA10",
]
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

# Full display names used in Excel (marketing mail)
UNDERLYING_FULL_NAMES = {
    "SX5E":  "EURO STOXX 50 Index",
    "SPX":   "S&P 500 Index",
    "NDX":   "NASDAQ-100 Index",
    "AEX":   "AEX Index",
    "SX7E":  "EURO STOXX Banks Index",
    "SD3E":  "EURO STOXX Select Dividend 30 Index",
    "NKY":   "Nikkei 225 Index",
    "EEM UP":"iShares MSCI Emerging Markets ETF",
    "SDGP":  "STOXX Global Select Dividend 100 Index EUR",
    "SX7P":  "STOXX Europe 600 Banks Index",
    "SXEP":  "STOXX Europe 600 Oil &amp; Gas Index",
    "SXKP":  "STOXX Europe 600 Telecommunications Index",
    "SX6P":  "STOXX Europe 600 Utilities Index",
}

# Short business-safe aliases used in product titles and doc-mail subjects
UNDERLYING_ALIASES = {
    "SX5E":                       "Eurozone",
    "SPX":                        "VS",
    "NDX":                        "US Technology",
    "EURO STOXX BANKS":           "Eurozone Banks",
    "EURO STOXX SELECT DIVIDEND": "Eurozone Dividend",
    "AEX":                        "",
    "SX7E":                       "Eurozone Banks",
    "SD3E":                       "Eurozone Dividend",
    "NKY":                        "JPN",
    "EEM UP":                     "EM",
    "SDGP":                       "Global Dividend",
    "SX7P":                       "",
    "SXEP":                       "",
    "SXKP":                       "",
    "SX6P":                       "",
}

# ============================================================================
# PC MAIL DOCUMENT LINKS
# ============================================================================

# Brochure links shown in the Word template (keyed by product name, uppercase)
PC_MAIL_BROCHURE_URLS: dict[str, tuple[str, str]] = {
    "TRIGGER PLUS NOTE": (
        "General Brochure Trigger Notes",
        "https://markets.vanlanschotkempen.com/uploadedfiles/Documents/FvL%2025328%20BRO%20SP%20Trigger%20Notes.pdf",
    ),
    "MEMORY COUPON": (
        "General Brochure Memory Coupon",
        "https://markets.vanlanschotkempen.com/uploadedfiles/Documents/FvL%2025329%20BRO%20Memory%20Coupon%20Notes.pdf",
    ),
    "INDEX GARANTIE NOTE": (
        "General Brochure Index Garantie Notes",
        "https://markets.vanlanschotkempen.com/uploadedfiles/Documents/Brochure_Index_Garantie_Notes_Jan23.pdf",
    ),
    "INDEX GARANTIE NOTE CAPPED": (
        "General Brochure Index Garantie Notes",
        "https://markets.vanlanschotkempen.com/uploadedfiles/Documents/Brochure_Index_Garantie_Notes_Jan23.pdf",
    ),
}

# Product video links shown in the Word template (keyed by product name, uppercase).
# URL is the same as in PRODUCT_TYPES; label is the PC Mail link text.
PC_MAIL_VIDEO_URLS: dict[str, tuple[str, str]] = {
    key.upper(): (info["video_label_pc"], info["video_url"])
    for key, info in {
        "Trigger Plus Note": {
            "video_label_pc": "Trigger Notes – Product Video",
            "video_url": PRODUCT_TYPES["Trigger Plus Note"]["video_url"],
        },
        "Index Garantie Note": {
            "video_label_pc": "Index Garantie Notes – Product Video",
            "video_url": PRODUCT_TYPES["Index Garantie Note"]["video_url"],
        },
        "Index Garantie Note Capped": {
            "video_label_pc": "Index Garantie Notes – Product Video",
            "video_url": PRODUCT_TYPES["Index Garantie Note Capped"]["video_url"],
        },
    }.items()
    if info["video_url"]
}

# Prospectus URLs keyed by code (DIP / SNIP)
PROSPECTUS_URLS: dict[str, tuple[str, str]] = {
    "DIP": (
        "DIP",
        "https://www.vanlanschotkempen.com/-/media/files/documents/corporate/"
        "investor-relations-en/debt-investors/library/dip/2025/prospectus/"
        "securities-note---15-may-2025.ashx",
    ),
    "SNIP": (
        "SNIP",
        "https://www.vanlanschotkempen.com/-/media/files/documents/corporate/"
        "investor-relations-en/debt-investors/library/snip/2025/prospectus/"
        "securities-note---30-may-2025.ashx",
    ),
}

