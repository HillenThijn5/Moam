# PCMail/main.py

from PCMail.builders.email_subject import build_email_subject
from PCMail.builders.email_body import build_todo_html, build_email_body_html
from PCMail.builders.word_context import build_word_context
from PCMail.builders.target_market import (
    build_target_market_b24,
    pick_target_market_sheet,
    compute_esg_score,
)
from PCMail.generators.word_generator import render_word
from PCMail.generators.excel_generator import render_target_market
from PCMail.mail.outlook_mailer import send_mail
from PCMail.models.inputdefinition import PCMailProduct, Underlying
from statics.loader import load_benchmark_map
from statics.outlook_user import get_sender_first_name


def run_pc_mail(product: PCMailProduct) -> None:
    """
    Main entrypoint called by the GUI when the user sends the PC mail.
    Receives a fully populated PCMailProduct, renders Word + Excel attachments,
    and opens the draft Outlook email.
    """
    # Enrich underlyings with benchmark data from the static Excel sheet
    benchmark_map = load_benchmark_map()
    for u in product.underlyings:
        bm = benchmark_map.get(u.ticker, {})
        if isinstance(bm, dict):
            u.primary_benchmark = bm.get("primary", u.primary_benchmark)
            u.fallback_benchmark = bm.get("fallback", u.fallback_benchmark)

    # Determine ESG classification for the target market document
    esg_score = compute_esg_score([u.ticker for u in product.underlyings])
    if product.product.strip().upper() == "FIXED RATE NOTE":
        esg_score = 3  # Fixed Rate Note is always score 3 (positive target market)
    b24_value = build_target_market_b24(esg_score)
    sheet_name = pick_target_market_sheet(product)

    # Build Word template context and render both attachments
    todo_html = build_todo_html(product)
    ctx = build_word_context(product, todo_html)
    word_path = render_word(ctx, product.series)
    excel_path = render_target_market(b24_value=b24_value, series=product.series, sheet_name=sheet_name)

    # Compose and open the draft Outlook email
    subject = build_email_subject(product)
    body = build_email_body_html(todo_html=todo_html, sender=get_sender_first_name())
    send_mail(subject=subject, html_body=body, attachments=[word_path, excel_path])
