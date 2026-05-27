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
    # Kernidentiteit / commercieel
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
    # Payoff-invoer (Excel-cellen)
    # Elk veld wordt hergebruikt voor producten met een andere semantische betekenis:
    #   coupon_protection  = coupon % (TPN/MC/FRN) of bescherming/min. aflossing % (IGN/IGNC)
    #   participation      = autocallbarrière % (TPN) of participatie % (IGN/IGNC)
    #   barrier_cap        = couponbarrière % (TPN/MC) of cap/max. aflossing % (IGNC)
    #   redemption_barrier = Europese aflossingsbarrière % (TPN/MC)
    # -------------------------
    coupon_protection: float
    participation: float
    barrier_cap: float
    redemption_barrier: float



    # -------------------------
    # Regelgevende / meta-invoer (DIP / SNIP)
    # -------------------------
    struct_fee: float
    dist_fee: float

    vlk_code_required: bool
    attach_target_market: bool
    fill_word: bool

    underlyings: List[Underlying]

    # -------------------------
    # Afgeleide velden — automatisch berekend, niet vereist vanuit de GUI
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
    tail: Optional[str] = ""   # Asianing-staart in maanden (bijv. "12" → "12m")
    obs: Optional[str] = ""    # Asianing-observaties (bijv. "13" → "13obs")
    coupon_frequency: Optional[str] = "Annual"  # Annual / Semi-Annual / Quarterly