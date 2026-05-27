# MarketingMail/input.py
"""
Enige bron voor invoerdata van de marketingmail.
Pas dit bestand aan om te wijzigen wat er verstuurd wordt.
"""

from MarketingMail.models import MarketingProduct
from MarketingMail.mail_service import create_and_send_marketing_mail

# ===== JOUW INVOERDATA =====
TITLE = "Nieuwe Maatwerk Notes"

PRODUCTS = [
    # TRIGGER-producten
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

    # MEMORY COUPON-producten (participation wordt automatisch op n.v.t. gezet)
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

    # INDEX GARANTIE-producten (barrier_cap wordt automatisch n.v.t., gebruikt asianing)
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

    # INDEX GARANTIE CAPPED (barrier_cap heeft een waarde, gebruikt asianing)
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
    # Maak en verstuur een mail met de eerste 2 producten
    mail = create_and_send_marketing_mail(
        title=TITLE,
        products=PRODUCTS,
        choice=2,
    )
    print(f"Mail created: {mail.Subject}")
