# MarketingMail/config.py
from pathlib import Path
from statics.data import (
    MARKETING_MAIL_TO_LIST as TO_LIST,
    STRUCTURED_INVESTMENTS_EMAIL as CC_LIST,
    UNDERLYING_FULL_NAMES,
    MARKETING_MAIL_SIGNATURE as SIGNATURE,
    MARKETING_PARAMETER_CONFIG as PARAMETER_CONFIG,
)

BASE_DIR = Path(__file__).resolve().parent
EXCEL_PATH = BASE_DIR / "templates" / "excel(marketingmail).xlsx"


# ====== EXCEL CELL MAPPING ======
# Maps logical field names to Excel cell addresses for each of the 4 product blocks.
# param1–param4 match the generic parameter slots used throughout the codebase.
PRODUCT_BLOCKS = [
    {
        "name": "B20",
        "issuer": "G22",
        "isin_vl": "G24",
        "ccy_nom": "G26",
        "maturity": "G28",
        "underlying": "G30",
        "startwaarde": "K30",
        "param1": "G32",
        "param2": "G34",
        "param3": "G36",
        "param4": "G38",
        "lbl_param1": "B32",
        "lbl_param2": "B34",
        "lbl_param3": "B36",
        "lbl_param4": "B38",
    },
    {
        "name": "B42",
        "issuer": "G44",
        "isin_vl": "G46",
        "ccy_nom": "G48",
        "maturity": "G50",
        "underlying": "G52",
        "startwaarde": "K52",
        "param1": "G54",
        "param2": "G56",
        "param3": "G58",
        "param4": "G60",
        "lbl_param1": "B54",
        "lbl_param2": "B56",
        "lbl_param3": "B58",
        "lbl_param4": "B60",
    },
    {
        "name": "B65",
        "issuer": "G67",
        "isin_vl": "G69",
        "ccy_nom": "G71",
        "maturity": "G73",
        "underlying": "G75",
        "startwaarde": "K75",
        "param1": "G77",
        "param2": "G79",
        "param3": "G81",
        "param4": "G83",
        "lbl_param1": "B77",
        "lbl_param2": "B79",
        "lbl_param3": "B81",
        "lbl_param4": "B83",
    },
    {
        "name": "B88",
        "issuer": "G90",
        "isin_vl": "G92",
        "ccy_nom": "G94",
        "maturity": "G96",
        "underlying": "G98",
        "startwaarde": "K98",
        "param1": "G100",
        "param2": "G102",
        "param3": "G104",
        "param4": "G106",
        "lbl_param1": "B100",
        "lbl_param2": "B102",
        "lbl_param3": "B104",
        "lbl_param4": "B106",
    },
]

# ====== PRODUCT TYPE LABELS & VALUE OVERRIDES ======
# Maps product_type → Dutch Excel label per param slot, plus optional forced values.
# value_param* overrides the value read from the product dict (use for fixed/n.v.t. cells).
PRODUCT_TYPE_CONFIG = {
    "TRIGGER": {
        "lbl_param1": "Premie:",
        "lbl_param2": "Aflossingsbarriere:",
        "lbl_param3": "Couponbarriere:",
        "lbl_param4": "Bescherming:",
        "default_param2": "100%",   # Aflossingsbarriere defaults to 100% for Trigger products
    },
    "MEMORY_COUPON": {
        "lbl_param1": "Premie:",
        "lbl_param2": "Aflossingsbarriere:",
        "lbl_param3": "Couponbarriere:",
        "lbl_param4": "Bescherming:",
        "value_param2": "n.v.t.",   # Memory Coupon has no aflossingsbarriere
    },
    "INDEX_GARANTIE": {
        "lbl_param1": "Protection:",
        "lbl_param2": "Participatiegraad:",
        "lbl_param3": "Max redemp:",
        "lbl_param4": "Middeling:",
        "value_param3": "n.v.t.",   # No cap for non-capped version
    },
    "INDEX_GARANTIE_CAPPED": {
        "lbl_param1": "Protection:",
        "lbl_param2": "Participatiegraad:",
        "lbl_param3": "Cap:",
        "lbl_param4": "Middeling:",
    },
    "AUTOCALL": {
        "lbl_param1": "Premie:",
        "lbl_param2": "Aflossingsbarriere:",
        "lbl_param3": "Couponbarriere:",
        "lbl_param4": "Bescherming:",
    },
}

# ====== STATIC TEXT TEMPLATES ======
TITLE_CELL = "B9"
INTRO_CELL = "B13"

# ====== OUTLOOK INLINE IMAGE ======
PR_ATTACH_CONTENT_ID = "http://schemas.microsoft.com/mapi/proptag/0x3712001F"
PR_ATTACH_CONTENT_LOCATION = "http://schemas.microsoft.com/mapi/proptag/0x3713001F"


def get_intro_text(denominations: list) -> str:
    """
    Dynamic intro text that includes all denominations.

    Args:
        denominations: List of denomination strings (e.g., ['EUR 100,000', 'USD 150,000'])

    Returns:
        Full intro text with denominations joined by /
    """
    denom_str = " / ".join(denominations)
    return (
        "Van Lanschot Kempen heeft de volgende maatwerk producten gelanceerd. "
        "Vanaf heden kan zolang er Notes beschikbaar zijn vanaf "
        f"{denom_str} mee gedaan worden in deze producten tegen de actuele laatprijs. "
        "Neem voor de actuele laatprijs en/of het opgeven van een order contact op met "
        "Van Lanschot Kempen Structured Products. "
        "Hieronder vindt u beknopt de belangrijkste kenmerken van de nieuwe maatwerk producten. "
        "De volledige documentatie verschijnt binnenkort op www.markets.vanlanschotkempen.com."
    )