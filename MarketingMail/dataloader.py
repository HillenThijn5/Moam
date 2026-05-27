# MarketingMail/dataloader.py
from typing import List
from MarketingMail.models import MarketingProduct
from MarketingMail.config import get_intro_text


def build_denomination_list(products: List[MarketingProduct]) -> list:
    """
    Haal unieke denominaties uit producten.
    Bijv. [EUR 100,000, USD 150,000]
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
    Hoofdloader voor data: past defaults toe en maakt introtekst met alle denominaties.

    Parameters:
        title: Waarde voor Excel-cel B9
        products: Lijst met MarketingProduct

    Geeft terug:
        dict met de sleutels: title, intro_text, products (alles als dicts)
    """
    # Zet alle producten om naar dicts (properties worden automatisch afgeleid)
    normalized_products = [p.to_dict() for p in products]

    # Bouw de denominaties op uit de echte producten
    denominations = build_denomination_list(products)
    intro_text = get_intro_text(denominations)

    return {
        "title": title,
        "intro_text": intro_text,
        "products": normalized_products,
    }