# Documentatie_Mail/data.py
"""
Build email data structure for Documentatie Mail.
Converts raw inputs into the format needed by the email builder.
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
    Build the complete email data dictionary for Documentatie Mail.

    Args:
        issuer: Issuer name
        currency: Currency code (e.g. EUR)
        product_type: Product type name
        maturity: Maturity string (e.g. "3Y")
        isin: ISIN code
        trades: List of trade dicts with keys: adviser, amount, [price]
        underlyings: List of underlying tickers
        to: List of TO recipient names
        cc: List of CC recipient names
        vlk_code: VL code to display in email
        sender_name: Name of sender

    Returns:
        dict: Complete email data ready for email_builder
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