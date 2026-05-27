"""
Klant-specifieke tekstbouwers.

Koppelt klantnaam → jurisdictie, distributeurcode en volledige klantomschrijving
voor gebruik in de context van het Word-sjabloon.
"""


def build_client_text(client: str) -> str:
    """
    Breidt 'VL Belgium' uit met de MiFID-disclaimer voor retailklanten.
    Alle andere klanten worden ongewijzigd teruggegeven.
    """
    if client == "VL Belgium":
        return (
            "VL Belgium – end client is a retail client under MiFID "
            "and is a consumer under Belgian law"
        )
    return client or ""


def build_jurisdiction(client: str) -> str:
    """Geeft de jurisdictiestring terug op basis van de klant."""
    if not client or not client.strip():
        return ""

    c = client.strip().lower()

    if c == "vl belgium":
        return "Belgium"
    if c == "vl switzerland":
        return "Switzerland"

    return "The Netherlands"


def build_distributor(client: str) -> str:
    """
    Zet de klantnaam om naar een distributeurcode.
    Alle VL*-klanten distribueren via VLK; anderen zijn hun eigen distributeur.
    """
    client = (client or "").strip()
    return "VLK" if client.startswith("VL") else client
