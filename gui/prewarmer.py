# gui/prewarmer.py
"""
Pre-imports heavy modules and caches static data in the background.
Called once at app startup so the first mail generation feels instant.
"""
from __future__ import annotations
import threading

_done = threading.Event()


def start() -> None:
    """Kick off pre-warming in a daemon thread. Returns immediately."""
    t = threading.Thread(target=_run, daemon=True, name="prewarm")
    t.start()


def wait(timeout: float = 60.0) -> None:
    """Block until pre-warming is done (or timeout expires)."""
    _done.wait(timeout=timeout)


def _run() -> None:
    try:
        # ── 1. docxtpl / python-docx / lxml ──────────────────────────────
        import docxtpl          # noqa: F401
        import docx             # noqa: F401

        # ── 2. pywin32 / COM ──────────────────────────────────────────────
        import win32com.client  # noqa: F401

        # ── 3. Static Excel data (results are cached inside loader.py) ────
        from statics.loader import load_benchmark_map, load_esg_map
        load_benchmark_map()
        load_esg_map()

        # ── 4. Word template — load once so docxtpl has it in OS cache ────
        try:
            from PCMail.config.paths import WORD_TEMPLATE_PATH
            import docxtpl as _dt
            _dt.DocxTemplate(str(WORD_TEMPLATE_PATH))  # warms file + parser
        except Exception:
            pass

        # ── 5. Marketing Mail — warm Excel template into OS file cache ────
        try:
            from MarketingMail.config import EXCEL_PATH
            EXCEL_PATH.read_bytes()  # pull template bytes into OS page cache
        except Exception:
            pass

    except Exception:
        pass  # pre-warming is best-effort; errors are silent
    finally:
        _done.set()
