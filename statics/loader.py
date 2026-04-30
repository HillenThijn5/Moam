# static_info/input.py
#haalt esg scores en benchmarks per onderliggende op


from pathlib import Path
from openpyxl import load_workbook

_STATIC_PATH = Path(__file__).resolve().parent / "static_sheet.xlsx"

# Module-level caches — populated on first call, reused forever after
_benchmark_cache: dict | None = None
_esg_cache: dict | None = None
_adviser_cache: dict | None = None


def load_benchmark_map() -> dict[str, dict[str, str]]:
    global _benchmark_cache
    if _benchmark_cache is not None:
        return _benchmark_cache
    """
    Reads sheet 'Benchmarks' from STATIC_DATA_PATH.

    Expected headers (row 1):
      ticker | primary_benchmark | fallback_benchmark
    """
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
        fallback = (row[i_fallback] if i_fallback >= 0 else "") if i_fallback >= 0 else ""

        out[ticker] = {
            "primary": str(primary).strip() if primary is not None else "",
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
    global _adviser_cache
    if _adviser_cache is not None:
        return _adviser_cache
    """
    Reads sheet 'Advisers' from static_sheet.xlsx.

    Expected columns (no header row required, but skips blanks):
      A: adviser name   B: helper (CC) name
    """
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
