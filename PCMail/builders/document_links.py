"""
Bouwers voor marketing- en documentlinks.

Bepaalt welke brochures, video's en prospectuslinks per product/klant gelden.
Deze worden in het Word-sjabloon gebruikt om hyperlinks voorwaardelijk weer te geven.

Alle URL's zijn gedefinieerd in statics.data (PC_MAIL_BROCHURE_URLS, PC_MAIL_VIDEO_URLS, PROSPECTUS_URLS).
"""
from statics.data import PC_MAIL_BROCHURE_URLS, PC_MAIL_VIDEO_URLS, PROSPECTUS_URLS


def marketing_logic(client: str, product: str) -> tuple[str, str]:
    """
    Geeft (marketing_flag, other_docs_flag) terug — elk "Yes" of "No".

    - VL NL krijgt zowel marketingmateriaal als andere documenten (video).
    - Alle andere klanten krijgen geen van beide.
    - Memory Coupon krijgt nooit other_docs (er bestaat geen productvideo).
    """
    client_lower = client.strip().lower()
    product_upper = product.strip().upper()

    marketing = "Yes" if client_lower == "vl nl" else "No"
    other_docs = "Yes" if client_lower == "vl nl" else "No"

    if product_upper == "MEMORY COUPON":
        other_docs = "No"

    return marketing, other_docs


def brochure_logic(product: str) -> tuple[str, str]:
    """Geeft (brochure_label, brochure_url) terug voor het product, of ("", "") als er geen bestaat."""
    return PC_MAIL_BROCHURE_URLS.get(product.strip().upper(), ("", ""))


def video_logic(product: str) -> tuple[str, str]:
    """Geeft (video_label, video_url) terug voor het product, of ("", "") als er geen bestaat."""
    return PC_MAIL_VIDEO_URLS.get(product.strip().upper(), ("", ""))


def prospectus_logic(code: str) -> tuple[str, str | None]:
    """
    Geeft (label, url) terug voor de opgegeven prospectuscode.
    Geeft (code, None) terug als de code niet wordt herkend (wordt als platte tekst weergegeven in het Word-document).
    """
    code_upper = (code or "").strip().upper()
    if code_upper in PROSPECTUS_URLS:
        return PROSPECTUS_URLS[code_upper]
    return code, None
