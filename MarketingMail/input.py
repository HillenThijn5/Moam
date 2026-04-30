# MarketingMail/input.py
"""
Single source of input data for marketing mail generation.
Modify this file to change what gets sent.
"""

from MarketingMail.models import MarketingProduct
from MarketingMail.mail_service import create_and_send_marketing_mail

# ===== YOUR INPUT DATA =====
TITLE = "Nieuwe Maatwerk Notes"

PRODUCTS = [
    # TRIGGER products
    MarketingProduct(
        product_type="TRIGGER",
        currency="EUR",
        underlying="SX5E",
        coupon_protection="5.5%",
        participation="70%",
        barrier_cap="10%",
        redemption_barrier="70%",
    ),
    MarketingProduct(
        product_type="TRIGGER",
        currency="USD",
        underlying="SPX",
        coupon_protection="5.5%",
        participation="70%",
        barrier_cap="10%",
        redemption_barrier="70%",
    ),

    # MEMORY COUPON products (participation value auto-set to n.v.t.)
    MarketingProduct(
        product_type="MEMORY_COUPON",
        currency="EUR",
        underlying="SX5E",
        coupon_protection="5.5%",
        barrier_cap="10%",
        redemption_barrier="70%",
    ),
    MarketingProduct(
        product_type="MEMORY_COUPON",
        currency="USD",
        underlying="SPX",
        coupon_protection="5.5%",
        barrier_cap="10%",
        redemption_barrier="70%",
    ),

    # INDEX GARANTIE products (barrier_cap auto-set to n.v.t., uses asianing)
    MarketingProduct(
        product_type="INDEX_GARANTIE",
        currency="EUR",
        underlying="SX5E",
        coupon_protection="100%",
        participation="100%",
        tail="12",
        obs="13",
    ),
    MarketingProduct(
        product_type="INDEX_GARANTIE",
        currency="USD",
        underlying="SPX",
        coupon_protection="100%",
        participation="100%",
    ),

    # INDEX GARANTIE CAPPED (barrier_cap has a value, uses asianing)
    MarketingProduct(
        product_type="INDEX_GARANTIE_CAPPED",
        currency="EUR",
        underlying="SX5E",
        coupon_protection="100%",
        participation="100%",
        barrier_cap="140%",
        tail="12",
        obs="13",
    ),
]

if __name__ == "__main__":
    # Generate and send mail with first 2 products
    mail = create_and_send_marketing_mail(
        title=TITLE,
        products=PRODUCTS,
        choice=2,
    )
    print(f"Mail created: {mail.Subject}")
