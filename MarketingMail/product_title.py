# MarketingMail/product_title.py
from datetime import datetime
from statics.data import UNDERLYING_FULL_NAMES, UNDERLYING_ALIASES, MARKETING_PRODUCT_TYPE_NAMES

# Koppelt de volledige issuer-string (zoals in het model opgeslagen) aan een korte code voor producttitels
_ISSUER_SHORT: dict[str, str] = {
    "Van Lanschot Kempen N.V.": "VLK",
    "BNP Paribas Issuance B.V.": "BNP",
    "UBS AG": "UBS",
    "SG Issuer": "SG",
}

# Alias voor oudere code die andere modules gebruiken
PRODUCT_TYPE_NAMES = MARKETING_PRODUCT_TYPE_NAMES


def _shorten_issuer(issuer: str) -> str:
    """Geef de korte code terug voor een volledige issuer-string, bijv. 'VLK', 'BNP', 'UBS', 'SG'."""
    for full, short in _ISSUER_SHORT.items():
        if issuer.startswith(full):
            return short
    return issuer  # terugvaloptie: gebruik zoals hij is



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
    """'5 jaar' of '5Y' → 5"""
    token = maturity.split()[0]       # '5' uit '5 jaar', '5Y' uit '5Y'
    return int(token.rstrip("Yy"))


def get_maturity_years_range(issue_date: str, maturity: str) -> str:
    """
    Maak een looptijdrange zoals '26-31'.

    Parameters:
        issue_date: bijv. '17 Apr 2026'
        maturity: bijv. '5 jaar'

    Geeft terug:
        '26-31'
    """
    start_year = extract_year_short(issue_date)
    maturity_years = maturity_to_years(maturity)
    end_year = str((int(start_year) + maturity_years) % 100).zfill(2)
    return f"{start_year}-{end_year}"


def generate_product_title(product: dict) -> str:
    """
    Maak een producttitel met ALIAS (niet de volledige naam) en looptijd.
    'VLK Trigger Plus Note Eurozone 26-31 USD' (gebruikt alias, niet de volledige naam)
    """
    issuer = _shorten_issuer(product.get("issuer", ""))

    product_type_key = product.get("product_type", "")
    # Als de coupon barrier 100% is voor een Trigger Plus Note, gebruik dan "Trigger Note"
    barrier = product.get("param3", "").replace("%", "").strip()
    if product_type_key == "TRIGGER" and barrier == "100":
        product_type = "Trigger Note"
    else:
        product_type = PRODUCT_TYPE_NAMES.get(
            product_type_key,
            product_type_key.title()
        )

    # Gebruik ALIAS voor de producttitel, niet de volledige naam
    underlying_alias = get_underlying_alias(product.get("underlying", ""))

    maturity_range = get_maturity_years_range(
        product.get("issue_date", "01 Jan 2026"),
        product.get("maturity", "5 jaar")
    )

    # Voeg een USD-suffix toe als dat van toepassing is
    currency_suffix = f" {product.get('currency')}" if product.get("currency") == "USD" else ""

    return f"{issuer} {product_type} {underlying_alias} {maturity_range}{currency_suffix}"