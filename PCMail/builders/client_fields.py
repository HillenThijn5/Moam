"""
Client-specific text builders.

Maps client name → jurisdiction, distributor code, and full client description
for use in the Word template context.
"""


def build_client_text(client: str) -> str:
    """
    Expands 'VL Belgium' to include the MiFID retail client disclaimer.
    All other clients are returned as-is.
    """
    if client == "VL Belgium":
        return (
            "VL Belgium – end client is a retail client under MiFID "
            "and is a consumer under Belgian law"
        )
    return client or ""


def build_jurisdiction(client: str) -> str:
    """Returns the jurisdiction string based on the client."""
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
    Maps client name to distributor code.
    All VL* clients distribute through VLK; others are their own distributor.
    """
    client = (client or "").strip()
    return "VLK" if client.startswith("VL") else client
