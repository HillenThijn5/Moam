from dataclasses import dataclass
from typing import List, Optional


@dataclass
class Underlying:
    ticker: str
    primary_benchmark: str = ""
    fallback_benchmark: str = ""


@dataclass
class PCMailProduct:
    # -------------------------
    # Core identity / commercial
    # -------------------------
    product: str
    series: str
    issuer: str
    currency: str
    client: str


    issue_size: str
    sold: str
    maturity: str
    hedged: str


    # -------------------------
    # Payoff inputs (Excel cells)
    # Each field is reused across products with different semantic meaning:
    #   coupon_protection  = coupon % (TPN/MC/FRN)  or  protection/min. redemption % (IGN/IGNC)
    #   participation      = autocall barrier % (TPN)  or  participation % (IGN/IGNC)
    #   barrier_cap        = coupon barrier % (TPN/MC)  or  cap/max. redemption % (IGNC)
    #   redemption_barrier = European redemption barrier % (TPN/MC)
    # -------------------------
    coupon_protection: float
    participation: float
    barrier_cap: float
    redemption_barrier: float



    # -------------------------
    # Regulatory / meta inputs (DIP / SNIP)
    # -------------------------
    struct_fee: float
    dist_fee: float

    vlk_code_required: bool
    attach_target_market: bool
    fill_word: bool

    underlyings: List[Underlying]

    # -------------------------
    # Derived fields — computed automatically, not required from the GUI
    # -------------------------
    parp: Optional[str] = ""
    jurisdiction: Optional[str] = ""
    issuer_series: Optional[str] = ""
    eusipa: Optional[str] = ""
    compliance: Optional[str] = ""
    mifid: Optional[str] = ""
    kid: Optional[str] = ""
    trade_date: Optional[str] = ""
    issue_date: Optional[str] = ""
    distributor: Optional[str] = ""
    denomination: Optional[str] = ""
    prospectus: Optional[str] = ""
    tail: Optional[str] = ""   # Asianing tail in months (e.g. "12" → "12m")
    obs: Optional[str] = ""    # Asianing observations (e.g. "13" → "13obs")