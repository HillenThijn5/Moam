"""
Logica voor target market.

Bepaalt de ESG-score, de B24-celwaarde en de Excel-sheetnaam
voor het document Target Market & Distribution Strategy.
"""
from statics.loader import load_esg_map


def compute_esg_score(underlyings) -> int | None:
    """
    Geeft de minimale ESG-score terug van de opgegeven underlyings.
    Accepteert zowel Underlying-objecten (met .ticker) als gewone tickerstrings.
    Geeft None terug als er geen scores zijn gevonden.
    """
    esg_map = load_esg_map()

    scores = []
    for u in underlyings:
        ticker = u.ticker if hasattr(u, "ticker") else u
        score = esg_map.get(ticker)
        if score is not None:
            scores.append(score)

    return min(scores) if scores else None


def build_target_market_b24(esg_score: int | None) -> str:
    """
    Koppelt de ESG-score aan de B24-celwaarde in het target-market-Excel-sjabloon.
    Score > 2 → "Yes - positive target market" (groen). Anders → "Neutral" (oranje).
    """
    if esg_score is not None and esg_score > 2:
        return "Yes - positive target market"
    return "Neutral"


def pick_target_market_sheet(product) -> str:
    """
    Geeft de Excel-sheetnaam terug die voor het opgegeven product moet worden gebruikt.
    Accepteert zowel een PCMailProduct-object als een gewone productnaamstring.
    """
    name = (
        product.product if hasattr(product, "product") else product
    ).strip().upper()

    if name in ("TRIGGER PLUS NOTE", "MEMORY COUPON"):
        return "1260"

    if name in ("INDEX GARANTIE NOTE", "INDEX GARANTIE NOTE CAPPED", "FIXED RATE NOTE"):
        return "1100_1120_above"

    return "1260"  # standaard terugval
