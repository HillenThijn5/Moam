
"""
Productspecifieke tekstbouwers.

Elke functie zet producttype + numerieke invoer om naar een leesbare string
voor gebruik in de context van het Word-sjabloon.
"""
from statics.loader import load_parp_dates


def build_payoff(
    product: str,
    coupon_protection: float,
    barrier_cap: float,
    redemption_barrier: float,
    participation: float,
    underlyings: list[str],
    tail: str = "",
    obs: str = "",
    coupon_frequency: str = "Annual",
) -> str:
    """Bouwt de payoff-beschrijvende zin voor het Word-document."""
    ow = "/".join(u for u in underlyings if u)
    p = product.strip()

    freq = (coupon_frequency or "Annual").strip().lower()
    freq_label = {"semi-annual": "semi-annual", "quarterly": "quarterly"}.get(freq, "annual")

    if p == "Trigger Plus Note":
        return (
            f"{participation} autocall barrier, {barrier_cap}% coupon barrier and "
            f"{redemption_barrier}% European style redemption barrier linked to "
            f"{ow} with {coupon_protection}% {freq_label} memory coupon"
        )

    if p == "Memory Coupon":
        return (
            f"{barrier_cap}% coupon barrier and "
            f"{redemption_barrier}% European style redemption barrier linked to "
            f"{ow} with {coupon_protection}% {freq_label} memory coupon"
        )

    if p == "Index Garantie Note":
        base = (
            f"{participation}% participation linked to {ow} "
            f"with {coupon_protection}% minimum redemption"
        )
        asianing_parts = []
        if tail:
            asianing_parts.append(f"tail {tail}m")
        if obs:
            asianing_parts.append(f"{obs}obs")
        if asianing_parts:
            base += ", with " + ", ".join(asianing_parts)
        return base

    if p == "Index Garantie Note Capped":
        base = (
            f"{participation}% participation with {coupon_protection}% minimum redemption "
            f"and {barrier_cap}% maximum redemption linked to {ow}"
        )
        asianing_parts = []
        if tail:
            asianing_parts.append(f"tail {tail}m")
        if obs:
            asianing_parts.append(f"{obs}obs")
        if asianing_parts:
            base += ", with " + ", ".join(asianing_parts)
        return base

    if p == "Fixed Rate Note":
        coupon_desc = {"semi-annual": "semi-annual", "quarterly": "quarterly"}.get(freq, "per annum")
        return f"100% redemption and {coupon_protection}% fixed coupon {coupon_desc}"

    return ""


def build_fee(struct_fee: float, dist_fee: float) -> str:
    """Formatteert de fee-regel. Laat de distributievergoeding weg als die nul is."""
    if dist_fee == 0:
        return f"Structuring fee: {struct_fee}% — No distribution fee."
    return f"Structuring fee: {struct_fee}% — Distribution fee: {dist_fee}%."


def build_issuer_series(issuer: str, prospectus: str, series: str) -> str:
    """Geeft de referentiestring voor uitgevende instelling en series in de Word-context terug."""
    if issuer.strip().upper() == "VLK":
        return f"VLK – under {prospectus} series {series}"
    return f"Issuer {issuer}"


def build_denomination(manual_denom: str, product: str, currency: str = "EUR") -> str:
    """Geeft de coupurestring terug. Gebruikt een handmatige waarde als die is opgegeven, anders standaardwaarden op basis van product en valuta."""
    if manual_denom and manual_denom.strip():
        return manual_denom.strip()

    p       = product.strip()
    non_eur = currency.strip().upper() in ("USD", "GBP")

    if p in ("Trigger Plus Note", "Memory Coupon", "Fixed Rate Note"):
        return "150k + 1k" if non_eur else "100k + 1k"

    if p in ("Index Garantie Note", "Index Garantie Note Capped"):
        return "100k + 1k" if non_eur else "50k + 1k"

    return ""


def build_parp(product: str) -> str:
    """Geeft de PARP-datum voor het product terug, gelezen uit het PARP-tabblad in static_sheet.xlsx."""
    return load_parp_dates().get(product.strip(), "")


def build_eusipa(product: str) -> str:
    """Geeft het EUSIPA-categorie- en subcategorielabel voor het product terug."""
    p = product.strip()

    if p in ("Trigger Plus Note", "Memory Coupon"):
        return (
            "EUSIPA Category 12: Yield Enhancement\n"
            "Sub-Category: 1260 – Express Notes"
        )

    if p == "Index Garantie Note":
        return (
            "EUSIPA Category 11: Capital protection\n"
            "Sub-Category: 1100 – Uncapped Capital Protected Notes"
        )

    if p == "Index Garantie Note Capped":
        return (
            "EUSIPA Category 11: Capital Protection\n"
            "Sub-Category: 1120 – Capped Capital Protection"
        )

    if p == "Fixed Rate Note":
        return (
            "Fixed Income Debt\n"
            "Sub-Category: Senior Preferred Notes"
        )

    return ""


def build_mifid(product: str) -> str:
    """Geeft de MiFID-classificatie voor productcomplexiteit terug."""
    if product.strip() == "Fixed Rate Note":
        return "Non-Complex"
    return "Complex"


def build_kid(product: str, client: str = "") -> str:
    """
    Geeft de KID-bladwijzertekst terug.
    VL Belgium en Fixed Rate Note krijgen een scenarioanalyse-disclaimer in plaats van 'Yes'.
    """
    if (
        client.strip().lower() == "vl belgium"
        or product.strip() == "Fixed Rate Note"
    ):
        return (
            "NB: KID will be created for scenario analysis purposes only. "
            "KID will not be made available to end investors or published on website."
        )
    return "Yes"


def build_compliance(issuer: str) -> str:
    """Geeft de tekst voor Legal/compliance-acties terug op basis van de uitgevende instelling."""
    if issuer.strip().upper() == "VLK":
        return (
            "SP desk will draft Final Terms. "
            "The FT for this product are pre-approved by Legal."
        )
    return f"Final Terms, Summary and KID will be created by {issuer}"


def auto_prospectus_code(product: str) -> str:
    """Geeft 'DIP' terug voor Fixed Rate Notes en 'SNIP' voor alle andere producten."""
    return "DIP" if (product or "").strip().upper() == "FIXED RATE NOTE" else "SNIP"
