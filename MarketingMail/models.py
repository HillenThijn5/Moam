# MarketingMail/models.py
from dataclasses import dataclass, field
from typing import Optional
from datetime import date

from PCMail.builders.dates import add_business_days


@dataclass
class MarketingProduct:
    """
    Input product definition for the Marketing Mail.

    Required:   product_type, currency, underlying
    Derived:    nominal, product_name, isin, vl_code, issue_date, startwaarde
    Payoff:     coupon_protection, participation, barrier_cap, redemption_barrier
    Asianing:   tail (months), obs (observations) — used for IGN/IGNC middeling text
    """
    product_type: str   # TRIGGER, MEMORY_COUPON, INDEX_GARANTIE, INDEX_GARANTIE_CAPPED, AUTOCALL
    currency: str       # EUR, USD
    underlying: str     # e.g. "SX5E" or "SX5E / SPX"
    issuer: str = "Van Lanschot Kempen N.V. (Fitch: A- / S&P: BBB+)"
    maturity: str = "5 jaar"

    # Payoff fields — same naming as PCMailProduct
    coupon_protection: Optional[str] = None  # Premie % (TRIGGER/MC) or Protection % (IGN)
    participation: Optional[str] = None      # Aflossingsbarrierre % (TRIGGER) or Participatiegraad % (IGN)
    barrier_cap: Optional[str] = None        # Coupon Barrier % (TRIGGER/MC) or Cap % (IGNC)
    redemption_barrier: Optional[str] = None # Protection % (TRIGGER/MC); for IGN/IGNC → asianing text

    # Asianing (middeling) — only for INDEX_GARANTIE and INDEX_GARANTIE_CAPPED
    tail: Optional[str] = None  # Tail in months, e.g. "12" → displayed as "12m"
    obs: Optional[str] = None   # Number of observations, e.g. "13" → displayed as "13obs"

    # Dates — default to auto; can be overridden from GUI
    start_date: str = field(
        default_factory=lambda: date.today().strftime("%d %b %Y")
    )
    issue_date: str = field(
        default_factory=lambda: add_business_days(date.today(), 5).strftime("%d %b %Y")
    )

    @property
    def nominal(self) -> str:
        """Derive nominal from currency + product_type."""
        non_eur = self.currency in ("USD", "GBP")
        if self.product_type in ("TRIGGER", "MEMORY_COUPON", "AUTOCALL"):
            return "150,000" if non_eur else "100,000"
        elif self.product_type in ("INDEX_GARANTIE", "INDEX_GARANTIE_CAPPED"):
            return "100,000" if non_eur else "50,000"
        return "TBD"

    @property
    def product_name(self) -> str:
        """Derive product name from type, underlying, currency, with maturity."""
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
        """Startwaarde text uses the start/trade date (today), not issue date."""
        return f"Startwaarde: TBD ({self.start_date})"

    def _build_asianing_text(self) -> str:
        """Formats tail and obs into the middeling text for IGN/IGNC products."""
        parts = []
        if self.tail:
            parts.append(f"{self.tail}m")
        if self.obs:
            parts.append(f"{self.obs}obs")
        return ", ".join(parts)

    def to_dict(self) -> dict:
        """Convert to dict for the Excel injection pipeline."""
        # For IGN/IGNC the 4th row (lbl_redemption_barrier / Middeling) shows asianing text.
        # For TRIGGER/MC it shows the actual redemption barrier value.
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
