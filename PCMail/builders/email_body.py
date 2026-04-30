"""
Email body builders.

Builds the SAS action item list, the Word todo section,
and the full HTML body of the PC notification email.
"""
from PCMail.builders.product_fields import build_compliance, build_denomination


def build_sas_action_items(product) -> list[str]:
    """
    Returns the ordered list of SAS action items for the product.
    Used in both the Word todo section and the email body.
    """
    items = []

    if product.issuer.upper() == "VLK":
        denom = build_denomination(
            getattr(product, "denomination", ""), product.product,
            getattr(product, "currency", "EUR"),
        )
        items.append(
            f"Please create ISIN & Common code "
            f"(TEFRA D/CGN/{product.currency} {denom} denom). No listing"
        )
        items.append("Please set up Note in Bloomberg and Telekurs")

    if product.vlk_code_required:
        items.append("Please set up VL code.")

    return items


def build_sas_word_text(product) -> str:
    """Formats SAS items as a bullet list for the Word template 'todo' bookmark."""
    items = build_sas_action_items(product)

    if not items:
        return ""

    return "• SAS:\n• " + "\n• ".join(items)


def build_todo_html(product) -> str:
    """Renders the <To do> HTML block for the email body."""
    sas_items = build_sas_action_items(product)
    sas_list_html = "".join(f"<li>{item}</li>" for item in sas_items)

    if product.issuer.upper() != "VLK":
        return f"""
        <p><u><b>To do:</b></u></p>
        <p><span style='background-color:yellow'><i>SAS:</i></span></p>
        <ul>
            {sas_list_html}
        </ul>
        <p><i>NB: the Notes will be setup in BBG and Telekurs by {product.issuer} (market = Kempen)</i></p>
        """

    return f"""
    <p><u><b>To do:</b></u></p>
    <p><i>Legal:</i> {build_compliance(product.issuer)}</p>
    <p><span style='background-color:yellow'><i>SAS:</i></span></p>
    <ul>
        {sas_list_html}
    </ul>
    """


def build_email_body_html(todo_html: str, sender: str = "") -> str:
    """Renders the full HTML body of the PC notification email."""
    return f"""
    <html><body style='font-family:Segoe UI; font-size:11pt'>
    <p>Hi,</p>
    <p>
    We would like to notify you of a new private placement deal.<br>
    Please find Termsheet, Individual Product Notification checklist and Target Market/Distribution strategy attached.
    </p>
    {todo_html}
    <p>Regards,<br>{sender}</p>
    </body></html>
    """
