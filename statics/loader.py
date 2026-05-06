# static_info/input.py
#haalt esg scores en benchmarks per onderliggende op


import sys
from pathlib import Path
from openpyxl import load_workbook


def _exe_root() -> Path:
    """Project root: exe's folder when frozen, otherwise MoamProject source root."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent
    return Path(__file__).resolve().parents[1]


_STATIC_PATH = _exe_root() / "statics" / "static_sheet.xlsx"

# Module-level caches — populated on first call, reused forever after
_benchmark_cache: dict | None = None
_esg_cache: dict | None = None
_adviser_cache: dict | None = None


def load_benchmark_map() -> dict[str, dict[str, str]]:
    """
    Reads sheet 'Benchmarks' from STATIC_DATA_PATH.

    Expected headers (row 1):
      ticker | primary_benchmark | fallback_benchmark
    """
    global _benchmark_cache
    if _benchmark_cache is not None:
        return _benchmark_cache
    wb = load_workbook(str(_STATIC_PATH), data_only=True)
    ws = wb["Benchmarks"]

    headers = [str(c.value).strip() if c.value is not None else "" for c in ws[1]]

    # find required columns by header name
    def col_idx(name: str) -> int:
        try:
            return headers.index(name)
        except ValueError:
            return -1

    i_ticker = col_idx("BENCHMARK ID")
    i_primary = col_idx("BENCHMARK NAME")
    i_fallback = col_idx("ALTERNATIVE BENCHMARK NAME")

    out: dict[str, dict[str, str]] = {}

    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or not any(row):
            continue

        ticker = (row[i_ticker] if i_ticker >= 0 else None) or ""
        ticker = str(ticker).strip()
        if not ticker:
            continue

        primary = (row[i_primary] if i_primary >= 0 else "") if i_primary >= 0 else ""
        primary = str(primary).strip() if primary is not None else ""
        if not primary:
            # No benchmark name → skip (e.g. EEM UP is a fund with no benchmark)
            continue

        fallback = (row[i_fallback] if i_fallback >= 0 else "") if i_fallback >= 0 else ""

        out[ticker] = {
            "primary": primary,
            "fallback": str(fallback).strip() if fallback is not None else "",
        }

    wb.close()
    _benchmark_cache = out
    return out

def load_esg_map() -> dict[str, int]:
    global _esg_cache
    if _esg_cache is not None:
        return _esg_cache
    wb = load_workbook(str(_STATIC_PATH), data_only=True)
    ws = wb["ESG"]

    headers = [str(c.value).strip().upper() if c.value else "" for c in ws[1]]

    def col(name: str):
        try:
            return headers.index(name)
        except ValueError:
            return None

    i_ticker = col("TICKER")
    i_score = col("SCORE")

    out = {}

    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or not any(row):
            continue

        ticker = str(row[i_ticker]).strip() if i_ticker is not None else ""
        if not ticker:
            continue

        try:
            score = int(row[i_score]) if i_score is not None else None
        except (TypeError, ValueError):
            score = None

        if score is not None:
            out[ticker] = score

    wb.close()
    _esg_cache = out
    return out


def load_adviser_map() -> dict[str, str]:
    """
    Reads sheet 'Advisers' from static_sheet.xlsx.

    Expected columns (no header row required, but skips blanks):
      A: adviser name   B: helper (CC) name
    """
    global _adviser_cache
    if _adviser_cache is not None:
        return _adviser_cache
    wb = load_workbook(str(_STATIC_PATH), data_only=True)
    ws = wb["advisers"]

    out: dict[str, str] = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or not row[0]:
            continue
        adviser = str(row[0]).strip()
        helper  = str(row[1]).strip() if len(row) > 1 and row[1] else ""
        if adviser:
            out[adviser] = helper

    wb.close()
    _adviser_cache = out
    return out


# ── helpers ──────────────────────────────────────────────────────────────────

def _col(headers: list[str], name: str) -> int:
    try:
        return headers.index(name.upper())
    except ValueError:
        return -1


def _str(val) -> str:
    return str(val).strip() if val is not None else ""


def _or_none(val) -> str | None:
    s = _str(val)
    return s if s else None


# ── Underlyings ───────────────────────────────────────────────────────────────

_underlyings_cache: dict | None = None


def load_underlyings() -> dict:
    """
    Reads sheet 'Benchmarks' for underlying metadata.
    Uses columns: BENCHMARK ID (ticker) | FULL_NAME | ALIAS

    All rows with a BENCHMARK ID are treated as valid underlyings,
    regardless of whether they have a benchmark (e.g. EEM UP is a fund
    with no BENCHMARK NAME but is still a valid underlying).

    Returns {
        'list':       [ticker, ...],          # ordered; drives GUI dropdowns
        'full_names': {ticker: full_name},    # marketing mail Excel display
        'aliases':    {ticker: alias},        # subject / title short names
    }
    """
    global _underlyings_cache
    if _underlyings_cache is not None:
        return _underlyings_cache

    wb = load_workbook(str(_STATIC_PATH), data_only=True)
    try:
        ws = wb["Benchmarks"]
    except KeyError:
        wb.close()
        _underlyings_cache = {"list": [], "full_names": {}, "aliases": {}}
        return _underlyings_cache

    headers = [_str(c.value).upper() for c in ws[1]]
    i_tick  = _col(headers, "BENCHMARK ID")
    i_full  = _col(headers, "FULL_NAME")
    i_alias = _col(headers, "ALIAS")

    tickers, full_names, aliases = [], {}, {}

    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row:
            continue
        ticker = _str(row[i_tick]) if i_tick >= 0 else ""
        if not ticker:
            continue
        tickers.append(ticker)
        if i_full >= 0 and row[i_full]:
            full_names[ticker] = _str(row[i_full])
        if i_alias >= 0 and row[i_alias]:
            aliases[ticker] = _str(row[i_alias])

    wb.close()
    _underlyings_cache = {"list": tickers, "full_names": full_names, "aliases": aliases}
    return _underlyings_cache


# ── Product URLs ──────────────────────────────────────────────────────────────

_product_urls_cache: dict | None = None


def load_product_urls() -> dict:
    """
    Reads sheet 'ProductURLs'.
    Headers (row 1):
      PRODUCT_TYPE | BROCHURE_LABEL | BROCHURE_URL | VIDEO_URL
                   | VIDEO_LABEL_MARKETING | VIDEO_LABEL_PC

    Returns {
        product_type: {
            'brochure_label':       str,
            'brochure_url':         str,
            'video_url':            str | None,
            'video_label':          str | None,   # marketing (Dutch)
            'video_label_pc':       str | None,   # PC mail (English)
        }, ...
    }
    """
    global _product_urls_cache
    if _product_urls_cache is not None:
        return _product_urls_cache

    wb = load_workbook(str(_STATIC_PATH), data_only=True)
    try:
        ws = wb["ProductURLs"]
    except KeyError:
        wb.close()
        _product_urls_cache = {}
        return _product_urls_cache

    headers = [_str(c.value).upper() for c in ws[1]]
    i_type  = _col(headers, "PRODUCT_TYPE")
    i_bl    = _col(headers, "BROCHURE_LABEL")
    i_bu    = _col(headers, "BROCHURE_URL")
    i_vu    = _col(headers, "VIDEO_URL")
    i_vlm   = _col(headers, "VIDEO_LABEL_MARKETING")
    i_vlp   = _col(headers, "VIDEO_LABEL_PC")

    out: dict = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row:
            continue
        pt = _str(row[i_type]) if i_type >= 0 else ""
        if not pt:
            continue
        out[pt] = {
            "brochure_label": _str(row[i_bl])  if i_bl  >= 0 else "",
            "brochure_url":   _str(row[i_bu])  if i_bu  >= 0 else "",
            "video_url":      _or_none(row[i_vu]  if i_vu  >= 0 else None),
            "video_label":    _or_none(row[i_vlm] if i_vlm >= 0 else None),
            "video_label_pc": _or_none(row[i_vlp] if i_vlp >= 0 else None),
        }

    wb.close()
    _product_urls_cache = out
    return out


# ── PARP dates ────────────────────────────────────────────────────────────────

_parp_cache: dict | None = None


def load_parp_dates() -> dict[str, str]:
    """
    Reads sheet 'PARP'.  Headers (row 1): PRODUCT_TYPE | PARP_DATE
    Store dates as plain text, e.g. '24 September 2025'.
    Returns {product_type: date_string}.
    """
    global _parp_cache
    if _parp_cache is not None:
        return _parp_cache

    wb = load_workbook(str(_STATIC_PATH), data_only=True)
    try:
        ws = wb["PARP"]
    except KeyError:
        wb.close()
        _parp_cache = {}
        return _parp_cache

    headers = [_str(c.value).upper() for c in ws[1]]
    i_type = _col(headers, "PRODUCT_TYPE")
    i_date = _col(headers, "PARP_DATE")

    out: dict[str, str] = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row:
            continue
        pt = _str(row[i_type]) if i_type >= 0 else ""
        if not pt:
            continue
        val = row[i_date] if i_date >= 0 else None
        if val is not None:
            from datetime import date as _date, datetime as _dt
            if isinstance(val, (_date, _dt)):
                # Format as "24 September 2025" (no leading zero)
                out[pt] = val.strftime("%d %B %Y").lstrip("0")
            else:
                out[pt] = _str(val)

    wb.close()
    _parp_cache = out
    return out


# ── Prospectus URLs ───────────────────────────────────────────────────────────

_prospectus_cache: dict | None = None


def load_prospectus_urls() -> dict[str, tuple[str, str]]:
    """
    Reads sheet 'ProspectusURLs'.  Headers (row 1): CODE | LABEL | URL
    Returns {code: (label, url)}.
    """
    global _prospectus_cache
    if _prospectus_cache is not None:
        return _prospectus_cache

    wb = load_workbook(str(_STATIC_PATH), data_only=True)
    try:
        ws = wb["ProspectusURLs"]
    except KeyError:
        wb.close()
        _prospectus_cache = {}
        return _prospectus_cache

    headers = [_str(c.value).upper() for c in ws[1]]
    i_code  = _col(headers, "CODE")
    i_label = _col(headers, "LABEL")
    i_url   = _col(headers, "URL")

    out: dict[str, tuple[str, str]] = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row:
            continue
        code = _str(row[i_code]) if i_code >= 0 else ""
        if not code:
            continue
        label = _str(row[i_label]) if i_label >= 0 else code
        url   = _str(row[i_url])   if i_url   >= 0 else ""
        out[code] = (label or code, url)

    wb.close()
    _prospectus_cache = out
    return out
