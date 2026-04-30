# PCMail/builders/logic_builder.py
#
# This file has been split into focused modules. See:
#   product_fields.py  — payoff, fee, kid, eusipa, mifid, parp, denomination, compliance
#   client_fields.py   — build_client_text, build_jurisdiction, build_distributor
#   document_links.py  — marketing_logic, brochure/video/prospectus links
#   target_market.py   — compute_esg_score, build_target_market_b24, pick_target_market_sheet
#   email_body.py      — build_sas_action_items, build_todo_html, build_email_body_html
#   word_context.py    — build_word_context, format_hedge_line
#
# Backward-compatible re-exports (remove once all callers are updated):

from PCMail.builders.word_context import build_word_context, format_hedge_line
from PCMail.builders.target_market import (
    build_target_market_b24,
    pick_target_market_sheet,
    compute_esg_score,
)
from PCMail.builders.email_body import (
    build_email_body_html as build_body_html,
    build_sas_action_items as build_sas_items,
    build_sas_word_text as build_sas_text,
)
from PCMail.builders.product_fields import (
    build_payoff,
    build_fee,
    build_eusipa,
    build_mifid,
    build_kid,
    build_compliance,
    build_parp,
    build_denomination,
    build_issuer_series,
    auto_prospectus_code,
)
from PCMail.builders.client_fields import (
    build_client_text,
    build_jurisdiction,
    build_distributor as distr,
)
from PCMail.builders.document_links import (
    marketing_logic,
    brochure_logic,
    video_logic,
    prospectus_logic,
)

