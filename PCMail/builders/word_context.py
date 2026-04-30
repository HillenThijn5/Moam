"""
Word template context builder.

Orchestrates all field builders into a single dict that is passed
to the Word template renderer (render_word).

Tuple values ("text", "url") or ("text", "url", "trailing text")
are converted to clickable hyperlinks by the Word generator.
"""
from PCMail.builders.product_fields import (
    build_payoff,
    build_fee,
    build_issuer_series,
    build_parp,
    build_eusipa,
    build_mifid,
    build_kid,
    build_compliance,
    auto_prospectus_code,
)
from PCMail.builders.client_fields import (
    build_client_text,
    build_jurisdiction,
    build_distributor,
)
from PCMail.builders.document_links import (
    marketing_logic,
    brochure_logic,
    video_logic,
    prospectus_logic,
)
from PCMail.builders.email_body import build_sas_word_text


def build_word_context(product, todo_html: str) -> dict:
    """
    Builds the complete bookmark → value mapping for the Word template.
    Called once per mail generation with the fully populated PCMailProduct.
    """
    p = product
    marketing, other_docs = marketing_logic(p.client, p.product)
    primary_bench, fallback_bench = _format_benchmarks(p)

    ctx = {
        # Core identity
        "client":        build_client_text(p.client),
        "ccy":           p.currency,
        "issue":         p.issue_size,
        "sold":          p.sold,
        "hedged":        p.hedged,
        "maturity":      p.maturity,
        "trade_date":    p.trade_date,
        "issue_date":    p.issue_date,

        # Client / regulatory
        "juri":          build_jurisdiction(p.client),
        "distr":         build_distributor(p.client),

        # Product-specific fields
        "parp":          build_parp(p.product),
        "payoff":        build_payoff(
                             product=p.product,
                             coupon_protection=p.coupon_protection,
                             barrier_cap=p.barrier_cap,
                             redemption_barrier=p.redemption_barrier,
                             participation=p.participation,
                             underlyings=[u.ticker for u in p.underlyings],
                             tail=getattr(p, "tail", ""),
                             obs=getattr(p, "obs", ""),
                         ),
        "fee":           build_fee(p.struct_fee, p.dist_fee),
        "issuer_series": build_issuer_series(
                             p.issuer, getattr(p, "prospectus", ""), p.series
                         ),
        "subcat":        build_eusipa(p.product),
        "mifd":          build_mifid(p.product),
        "kid":           build_kid(p.product, p.client),
        "compliance":    build_compliance(p.issuer),

        # Marketing / documents
        "marketing":     marketing,
        "other_docs":    other_docs,

        # Benchmarks
        "benchmark":          primary_bench,
        "fallback_benchmark": fallback_bench,

        # SAS todo block (Word bookmark)
        "todo":          build_sas_word_text(p),
    }

    inlegvel_suffix = (
        " and Product leaflet (inlegvel)"
        if marketing == "Yes" and p.client.strip().lower() == "vl nl"
        else ""
    )
    # Keep separate inlegvel bookmark empty when we append it after brochure link.
    ctx["inlegvel"] = ""

    # Prospectus hyperlink
    pros_code = auto_prospectus_code(p.product)
    pros_text, pros_url = prospectus_logic(pros_code)
    ctx["prospectus"] = (pros_text, pros_url)

    # Brochure hyperlink (only for VL NL)
    if marketing == "Yes":
        text, url = brochure_logic(p.product)
        if text and url:
            ctx["brochure"] = (text, url, inlegvel_suffix)
        else:
            ctx["brochure"] = ""
    else:
        ctx["brochure"] = ""

    # Product video hyperlink (only for VL NL, not Memory Coupon)
    if other_docs == "Yes":
        text, url = video_logic(p.product)
        ctx["video"] = (text, url) if text and url else ""
    else:
        ctx["video"] = ""

    return ctx


def _format_benchmarks(product) -> tuple[str, str]:
    """Joins primary and fallback benchmark names from all underlyings."""
    primary = ", ".join(
        u.primary_benchmark for u in product.underlyings
        if getattr(u, "primary_benchmark", None)
    )
    fallback = ", ".join(
        u.fallback_benchmark for u in product.underlyings
        if getattr(u, "fallback_benchmark", None)
    )
    return primary or "Not Applicable", fallback or "Not Applicable"


def format_hedge_line(book_type: str, hedge_party: str, upfront: str, amount: str) -> str:
    """Formats the hedge description for the Word context."""
    if book_type == "Own Book":
        return "own book"
    return f"{amount} BTB with {hedge_party} @{upfront}%"
