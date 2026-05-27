# gui/prewarmer.py
"""
Importeer zware modules alvast en cache statische data op de achtergrond.
Wordt één keer bij app-start aangeroepen zodat de eerste mail meteen snel voelt.
"""
from __future__ import annotations
import threading
from pathlib import Path

_done = threading.Event()


def start() -> None:
    """Start het voorverwarmen in een daemonthread. Geeft meteen terug."""
    t = threading.Thread(target=_run, daemon=True, name="prewarm")
    t.start()


def wait(timeout: float = 60.0) -> None:
    """Wacht tot het voorverwarmen klaar is (of tot de timeout verloopt)."""
    _done.wait(timeout=timeout)


def _run() -> None:
    try:
        # ── 1. docxtpl / python-docx / lxml ──────────────────────────────
        import docxtpl          # noqa: F401
        import docx             # noqa: F401

        # ── 2. pywin32 / COM ──────────────────────────────────────────────
        import win32com.client  # noqa: F401

        # ── 3. Statische Excel-data (resultaten worden gecachet in loader.py) ──
        from statics.loader import load_benchmark_map, load_esg_map, load_adviser_map
        load_benchmark_map()
        load_esg_map()
        load_adviser_map()

        # ── 4. Word-sjabloon — laad één keer zodat docxtpl hem in de OS-cache heeft ──
        try:
            from PCMail.config.paths import WORD_TEMPLATE_PATH
            import docxtpl as _dt
            _dt.DocxTemplate(str(WORD_TEMPLATE_PATH))  # warmt bestand + parser op
        except Exception:
            pass

        # ── 5. Marketing Mail — warm Excel-sjabloon op in de OS-bestandscache ──
        try:
            from MarketingMail.config import EXCEL_PATH
            EXCEL_PATH.read_bytes()  # trek sjabloonbytes in de OS-page-cache
        except Exception:
            pass

        # ── 6. PCMail Target-Market Excel-sjabloon ───────────────────────
        try:
            from PCMail.config.paths import EXCEL_TM_TEMPLATE_PATH
            EXCEL_TM_TEMPLATE_PATH.read_bytes()
        except Exception:
            pass

        # ── 7. Increase/Decrease Mail Excel-sjabloon ─────────────────────
        try:
            from increase_decrease_mail.mail_service import TEMPLATE_PATH as ID_TEMPLATE
            ID_TEMPLATE.read_bytes()
        except Exception:
            pass

        # ── 8. SharePoint-samenvattings-Excel ────────────────────────────
        try:
            from statics.data import SHAREPOINT_SUMMARY_PATH
            Path(SHAREPOINT_SUMMARY_PATH).read_bytes()
        except Exception:
            pass

    except Exception:
        pass  # voorverwarmen is best-effort; fouten blijven stil
    finally:
        _done.set()
