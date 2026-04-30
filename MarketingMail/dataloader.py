# MarketingMail/dataloader.py
from typing import List
from MarketingMail.models import MarketingProduct
from MarketingMail.config import get_intro_text


def build_denomination_list(products: List[MarketingProduct]) -> list:
    """
    Extract unique denominations from products.
    E.g., [EUR 100,000, USD 150,000]
    """
    denoms = set()
    for p in products:
        denom_str = f"{p.currency} {p.nominal}"
        denoms.add(denom_str)
    return sorted(list(denoms))


def load_marketing_data(
        title: str,
        products: List[MarketingProduct]
) -> dict:
    """
    Main data loader: applies defaults, generates intro text with all denominations.

    Args:
        title: Excel cell B9 value
        products: List of MarketingProduct

    Returns:
        dict with keys: title, intro_text, products (all as dicts)
    """
    # Convert all products to dicts (properties auto-derive)
    normalized_products = [p.to_dict() for p in products]

    # Build denominations from actual products
    denominations = build_denomination_list(products)
    intro_text = get_intro_text(denominations)

    return {
        "title": title,
        "intro_text": intro_text,
        "products": normalized_products,
    }