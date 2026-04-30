"""
Marketing and document link builders.

Controls which brochures, videos and prospectus links apply per product/client.
These are used in the Word template to conditionally render hyperlinks.

All URLs are defined in statics.data (PC_MAIL_BROCHURE_URLS, PC_MAIL_VIDEO_URLS, PROSPECTUS_URLS).
"""
from statics.data import PC_MAIL_BROCHURE_URLS, PC_MAIL_VIDEO_URLS, PROSPECTUS_URLS


def marketing_logic(client: str, product: str) -> tuple[str, str]:
    """
    Returns (marketing_flag, other_docs_flag) — each "Yes" or "No".

    - VL NL gets both marketing materials and other docs (video).
    - All other clients get neither.
    - Memory Coupon never gets other_docs (no product video exists).
    """
    client_lower = client.strip().lower()
    product_upper = product.strip().upper()

    marketing = "Yes" if client_lower == "vl nl" else "No"
    other_docs = "Yes" if client_lower == "vl nl" else "No"

    if product_upper == "MEMORY COUPON":
        other_docs = "No"

    return marketing, other_docs


def brochure_logic(product: str) -> tuple[str, str]:
    """Returns (brochure_label, brochure_url) for the product, or ("", "") if none exists."""
    return PC_MAIL_BROCHURE_URLS.get(product.strip().upper(), ("", ""))


def video_logic(product: str) -> tuple[str, str]:
    """Returns (video_label, video_url) for the product, or ("", "") if none exists."""
    return PC_MAIL_VIDEO_URLS.get(product.strip().upper(), ("", ""))


def prospectus_logic(code: str) -> tuple[str, str | None]:
    """
    Returns (label, url) for the given prospectus code.
    Returns (code, None) if the code is unrecognised (renders as plain text in the Word doc).
    """
    code_upper = (code or "").strip().upper()
    if code_upper in PROSPECTUS_URLS:
        return PROSPECTUS_URLS[code_upper]
    return code, None
