# MarketingMail/injection.py
from MarketingMail.config import PRODUCT_BLOCKS, PRODUCT_TYPE_CONFIG
from MarketingMail.product_title import generate_product_title, get_underlying_full_name


def _as_pct(value: str) -> str:
    """Append % to a bare numeric value; leave already-formatted or text values unchanged."""
    if not value:
        return value
    v = value.strip()
    if v.endswith("%"):
        return v
    try:
        float(v)
        return v + "%"
    except ValueError:
        return v  # e.g. "n.v.t.", "12m, 13obs" — leave as-is


def get_labels_and_values(product_dict: dict) -> dict:
    """
    Returns Excel label text and cell values for a product, applying any
    product-type overrides from PRODUCT_TYPE_CONFIG.

    Returns a dict with keys:
        lbl_param1, lbl_param2, lbl_param3, lbl_param4
        value_param1, value_param2, value_param3, value_param4
    """
    product_type = product_dict.get("product_type")
    config = PRODUCT_TYPE_CONFIG.get(product_type)
    if config is None:
        raise ValueError(
            f"Unknown product_type '{product_type}'. "
            f"Valid types: {', '.join(PRODUCT_TYPE_CONFIG.keys())}"
        )

    return {
        # Excel label text (Dutch)
        "lbl_param1": config.get("lbl_param1", "Premie:"),
        "lbl_param2": config.get("lbl_param2", "Aflossingsbarriere:"),
        "lbl_param3": config.get("lbl_param3", "Couponbarriere:"),
        "lbl_param4": config.get("lbl_param4", "Bescherming:"),

        # Values — forced config override > user-entered value > config default > empty
        "value_param1": _as_pct(product_dict.get("param1", "")),
        "value_param2": config.get("value_param2") or _as_pct(product_dict.get("param2", "")) or config.get("default_param2", ""),
        "value_param3": config.get("value_param3") or _as_pct(product_dict.get("param3", "")),
        "value_param4": _as_pct(product_dict.get("param4", "")),
    }


def build_cell_updates(data: dict, choice: int) -> dict:
    """
    Builds a {cell_address: value} dict ready for Excel injection.

    Args:
        data:   dict from load_marketing_data() with 'title', 'intro_text', 'products'
        choice: 1–4 (how many products to include)
    """
    if not 1 <= choice <= 4:
        raise ValueError("choice must be between 1 and 4")

    if len(data["products"]) < choice:
        raise ValueError(
            f"choice={choice} but only {len(data['products'])} products provided"
        )

    updates = {
        "B9":  data["title"],
        "B12": data["intro_text"],
    }

    for i, product in enumerate(data["products"][:choice]):
        block = PRODUCT_BLOCKS[i]
        lv = get_labels_and_values(product)
        product_title = generate_product_title(product)

        updates.update({
            # Title + metadata
            block["name"]:       product_title,
            block["issuer"]:     product.get("issuer", ""),
            block["isin_vl"]:    f'{product.get("isin", "")} / {product.get("vl_code", "")} / {product.get("issue_date", "")}',
            block["ccy_nom"]:    f'{product.get("currency", "")} {product.get("nominal", "")}',
            block["maturity"]:   product.get("maturity", ""),
            block["underlying"]: get_underlying_full_name(product.get("underlying", "")),
            block["startwaarde"]: product.get("startwaarde", ""),

            # Excel labels
            block["lbl_param1"]: lv["lbl_param1"],
            block["lbl_param2"]: lv["lbl_param2"],
            block["lbl_param3"]: lv["lbl_param3"],
            block["lbl_param4"]: lv["lbl_param4"],

            # Excel values
            block["param1"]: lv["value_param1"],
            block["param2"]: lv["value_param2"],
            block["param3"]: lv["value_param3"],
            block["param4"]: lv["value_param4"],
        })

    return updates
