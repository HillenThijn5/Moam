# MarketingMail/models.py
from dataclasses import dataclass, field
from typing import Optional
from datetime import date

from PCMail.builders.dates import add_business_days


@dataclass
class MarketingProduct:
    """
    Invoerdefinitie van een product voor de Marketing Mail.

    Verplicht:  product_type, currency, underlying
    Afgeleid:   nominal, product_name, isin, vl_code, issue_date, startwaarde
    Payoff:     coupon_protection, participation, barrier_cap, redemption_barrier
    Asianing:   tail (maanden), obs (observaties) — gebruikt voor IGN/IGNC-middelingstekst
    """
    product_type: str   # TRIGGER, MEMORY_COUPON, INDEX_GARANTIE, INDEX_GARANTIE_CAPPED, AUTOCALL
    currency: str       # EUR, USD
    underlying: str     # bijv. "SX5E" of "SX5E / SPX"
    issuer: str = "Van Lanschot Kempen N.V. (Fitch: A- / S&P: BBB+)"
    maturity: str = "5 jaar"

    # Payoff-velden — zelfde naamgeving als PCMailProduct
    coupon_protection: Optional[str] = None  # Premie % (TRIGGER/MC) of Protection % (IGN)
    participation: Optional[str] = None      # Aflossingsbarrierre % (TRIGGER) of Participatiegraad % (IGN)
    barrier_cap: Optional[str] = None        # Coupon Barrier % (TRIGGER/MC) of Cap % (IGNC)
    redemption_barrier: Optional[str] = None # Protection % (TRIGGER/MC); voor IGN/IGNC → asianing-tekst

    # Asianing (middeling) — alleen voor INDEX_GARANTIE en INDEX_GARANTIE_CAPPED
    tail: Optional[str] = None  # Tail in maanden, bijv. "12" → wordt getoond als "12m"
    obs: Optional[str] = None   # Aantal observaties, bijv. "13" → wordt getoond als "13obs"

    # Datums — standaard automatisch; kan vanuit de GUI overschreven worden
    start_date: str = field(
        default_factory=lambda: date.today().strftime("%d %b %Y")
    )
    issue_date: str = field(
        default_factory=lambda: add_business_days(date.today(), 5).strftime("%d %b %Y")
    )

    @property
    def nominal(self) -> str:
        """Leid nominal af uit currency + product_type."""
        non_eur = self.currency in ("USD", "GBP")
        if self.product_type in ("TRIGGER", "MEMORY_COUPON", "AUTOCALL"):
            return "150,000" if non_eur else "100,000"
        elif self.product_type in ("INDEX_GARANTIE", "INDEX_GARANTIE_CAPPED"):
            return "100,000" if non_eur else "50,000"
        return "TBD"

    @property
    def product_name(self) -> str:
        """Leid de productnaam af uit type, underlying, currency en looptijd."""
        from MarketingMail.product_title import PRODUCT_TYPE_NAMES, get_underlying_alias, get_maturity_years_range, _shorten_issuer

        issuer_short = _shorten_issuer(self.issuer)
        product_type_name = PRODUCT_TYPE_NAMES.get(self.product_type, self.product_type.title())
        underlying_alias = get_underlying_alias(self.underlying)
        maturity_range = get_maturity_years_range(self.issue_date, self.maturity)
        currency_suffix = f" {self.currency}" if self.currency == "USD" else ""

        return f"{issuer_short} {product_type_name} {underlying_alias} {maturity_range}{currency_suffix}"

    @property
    def isin(self) -> str:
        return "TBD"

    @property
    def vl_code(self) -> str:
        return "TBD"

    @property
    def startwaarde(self) -> str:
        """De startwaarde-tekst gebruikt de start-/tradedatum (vandaag), niet de issue date."""
        return f"Startwaarde: TBD ({self.start_date})"

    def _build_asianing_text(self) -> str:
        """Zet tail en obs om naar de middelingstekst voor IGN/IGNC-producten."""
        parts = []
        if self.tail:
            parts.append(f"{self.tail}m")
        if self.obs:
            parts.append(f"{self.obs}obs")
        return ", ".join(parts)

    def to_dict(self) -> dict:
        """Zet om naar een dict voor de Excel-injectieflow."""
        # Voor IGN/IGNC toont de 4e rij (lbl_redemption_barrier / Middeling) de asianing-tekst.
        # Voor TRIGGER/MC toont die de echte redemption barrier-waarde.
        if self.product_type in ("INDEX_GARANTIE", "INDEX_GARANTIE_CAPPED"):
            fourth_value = self._build_asianing_text() or "n.v.t."
        else:
            fourth_value = self.redemption_barrier or ""

        return {
            "product_type":       self.product_type,
            "product_name":       self.product_name,
            "issuer":             self.issuer,
            "currency":           self.currency,
            "nominal":            self.nominal,
            "maturity":           self.maturity,
            "underlying":         self.underlying,
            "isin":               self.isin,
            "vl_code":            self.vl_code,
            "issue_date":         self.issue_date,
            "startwaarde":        self.startwaarde,
            "param1": self.coupon_protection or "",
            "param2": self.participation or "",
            "param3": self.barrier_cap or "",
            "param4": fourth_value,
        }
