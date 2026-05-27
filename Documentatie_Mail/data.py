# Documentatie_Mail/data.py
"""
Bouw de datastructuur voor de Documentatie Mail op.
Zet ruwe invoer om naar het formaat dat de email builder nodig heeft.
"""

from datetime import date


def build_email_data(
        issuer: str,
        currency: str,
        product_type: str,
        maturity: str,
        isin: str,
        trades: list,
        underlyings: list = None,
        to: list = None,
        cc: list = None,
        vlk_code: str = None,
        sender_name: str = None,
) -> dict:
    """
    Bouw de volledige email-data-dict voor de Documentatie Mail.

    Parameters:
        issuer: Naam van de issuer
        currency: Valutacode (bijv. EUR)
        product_type: Naam van het producttype
        maturity: Looptijdstring (bijv. "3Y")
        isin: ISIN-code
        trades: Lijst met trade-dicts met de sleutels: adviser, amount, [price]
        underlyings: Lijst met underlying-tickers
        to: Lijst met namen voor TO-ontvangers
        cc: Lijst met namen voor CC-ontvangers
        vlk_code: VL-code die in de mail moet komen
        sender_name: Naam van de afzender

    Geeft terug:
        dict: Volledige email-data klaar voor email_builder
    """
    return {
        "sender_name": sender_name,
        "product": product_type,
        "isin": isin,
        "maturity": maturity,
        "currency": currency,
        "issuer": issuer,
        "underlyings": underlyings or [],
        "trades": trades or [],
        "to": to or [],
        "cc": cc or [],
        "vlk_code": vlk_code,
        "today": date.today(),
    }