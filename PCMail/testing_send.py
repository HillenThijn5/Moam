# PCMail/testing_send.py
# Ontwikkelvoorbeeld — alleen gebruikt wanneer main.py direct wordt gestart voor handmatig testen.
# Niet gebruikt in productie (de GUI bouwt PCMailProduct op uit gebruikersinvoer).

from PCMail.models.inputdefinition import PCMailProduct, Underlying
from PCMail.builders.dates import build_trade_and_issue_dates


trade_date, issue_date = build_trade_and_issue_dates()

p = PCMailProduct(
    product="Trigger Plus Note",
    series="1351",
    issuer="VLK",
    currency="EUR",
    client="VL NL",

    issue_size=0,
    sold=0,
    maturity="",
    hedged="False",

    coupon_protection=0,
    participation=0,
    barrier_cap=0,
    redemption_barrier=0,

    trade_date=trade_date,
    issue_date=issue_date,
    struct_fee=0.0,
    dist_fee=0.0,

    vlk_code_required=True,
    attach_target_market=True,
    fill_word=True,
    underlyings=[Underlying(ticker="SX5E")],
)