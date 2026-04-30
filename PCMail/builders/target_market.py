"""
Target market logic.

Determines the ESG score, B24 cell value, and Excel sheet name
for the Target Market & Distribution Strategy document.
"""
from statics.loader import load_esg_map


def compute_esg_score(underlyings) -> int | None:
    """
    Returns the highest ESG score among the given underlyings.
    Accepts either Underlying objects (with .ticker) or plain ticker strings.
    Returns None if no scores are found.
    """
    esg_map = load_esg_map()

    scores = []
    for u in underlyings:
        ticker = u.ticker if hasattr(u, "ticker") else u
        score = esg_map.get(ticker)
        if score is not None:
            scores.append(score)

    return max(scores) if scores else None


def build_target_market_b24(esg_score: int | None) -> str:
    """
    Maps ESG score to the B24 cell value in the target market Excel template.
    Score > 2 → positive target market (green). Otherwise → neutral (orange).
    """
    if esg_score is not None and esg_score > 2:
        return "Yes - positive target market"
    return "Neutral"


def pick_target_market_sheet(product) -> str:
    """
    Returns the Excel sheet name to use for the given product.
    Accepts either a PCMailProduct object or a plain product name string.
    """
    name = (
        product.product if hasattr(product, "product") else product
    ).strip().upper()

    if name in ("TRIGGER PLUS NOTE", "MEMORY COUPON"):
        return "1260"

    if name in ("INDEX GARANTIE NOTE", "INDEX GARANTIE NOTE CAPPED", "FIXED RATE NOTE"):
        return "1100_1120_above"

    return "1260"  # default fallback
