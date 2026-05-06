# MarketingMail/product_title.py
from datetime import datetime
from statics.data import UNDERLYING_FULL_NAMES, UNDERLYING_ALIASES, MARKETING_PRODUCT_TYPE_NAMES

# Maps full issuer string (as stored in the model) → short code for product titles
_ISSUER_SHORT: dict[str, str] = {
    "Van Lanschot Kempen N.V.": "VLK",
    "BNP Paribas Issuance B.V.": "BNP",
    "UBS AG": "UBS",
    "SG Issuer": "SG",
}

# Backward-compatible alias used by other modules
PRODUCT_TYPE_NAMES = MARKETING_PRODUCT_TYPE_NAMES


def _shorten_issuer(issuer: str) -> str:
    """Return the short code for a full issuer string, e.g. 'VLK', 'BNP', 'UBS', 'SG'."""
    for full, short in _ISSUER_SHORT.items():
        if issuer.startswith(full):
            return short
    return issuer  # fallback: use as-is



def get_underlying_full_name(underlying: str) -> str:
    if not underlying:
        return ""

    parts = [u.strip() for u in underlying.split("/")]

    names = [
        UNDERLYING_FULL_NAMES.get(u, u)
        for u in parts
    ]

    return " / ".join(names)



def get_underlying_alias(underlying: str) -> str:
    if not underlying:
        return ""

    parts = [u.strip() for u in underlying.split("/")]

    aliases = [
        UNDERLYING_ALIASES.get(u, u)
        for u in parts
    ]

    return " / ".join(aliases)



def extract_year_short(date_str: str) -> str:
    """'17 Apr 2026' → '26'"""
    year = datetime.strptime(date_str, "%d %b %Y").year
    return str(year)[-2:]


def maturity_to_years(maturity: str) -> int:
    """'5 jaar' or '5Y' → 5"""
    token = maturity.split()[0]       # '5' from '5 jaar', '5Y' from '5Y'
    return int(token.rstrip("Yy"))


def get_maturity_years_range(issue_date: str, maturity: str) -> str:
    """
    Generate maturity range like '26-31'.

    Args:
        issue_date: e.g. '17 Apr 2026'
        maturity: e.g. '5 jaar'

    Returns:
        '26-31'
    """
    start_year = extract_year_short(issue_date)
    maturity_years = maturity_to_years(maturity)
    end_year = str((int(start_year) + maturity_years) % 100).zfill(2)
    return f"{start_year}-{end_year}"


def generate_product_title(product: dict) -> str:
    """
    Generate product title with ALIAS (not full name) and maturity.
    'VLK Trigger Plus Note Eurozone 26-31 USD' (uses alias, not full name)
    """
    issuer = _shorten_issuer(product.get("issuer", ""))

    product_type = PRODUCT_TYPE_NAMES.get(
        product.get("product_type"),
        product.get("product_type", "").title()
    )

    # Use ALIAS for product title, not full name
    underlying_alias = get_underlying_alias(product.get("underlying", ""))

    maturity_range = get_maturity_years_range(
        product.get("issue_date", "01 Jan 2026"),
        product.get("maturity", "5 jaar")
    )

    # Add USD suffix if applicable
    currency_suffix = f" {product.get('currency')}" if product.get("currency") == "USD" else ""

    return f"{issuer} {product_type} {underlying_alias} {maturity_range}{currency_suffix}"