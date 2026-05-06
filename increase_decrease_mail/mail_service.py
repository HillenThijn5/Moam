# increase_decrease_mail/mail_service.py
"""
Injects deal values into the Excel template, copies range B1:C22,
and pastes it into an Outlook draft — same pattern as Marketing Mail.

Note: Excel must stay open until after sel.Paste() because quitting Excel
flushes the Windows clipboard, which would cause a "Clipboard is empty" error.
"""
from __future__ import annotations
import sys
from pathlib import Path

import win32com.client as win32

from statics.data import ID_MAIL_TO, ID_MAIL_CC
from MarketingMail.excel_handler import ExcelHandler


def _exe_root() -> Path:
    """Project root: exe's folder when frozen, otherwise MoamProject source root."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent
    return Path(__file__).resolve().parents[1]


TEMPLATE_PATH = _exe_root() / "increase_decrease_mail" / "increasedecrease(mailtemplate).xlsx"
SHEET_NAME    = "IncreaseDecrease"
COPY_RANGE    = "B1:C22"


def send_increase_decrease_mail(
    name: str,
    isin: str,
    size: float,
    transfer_price: float | None,
    increase: bool,
) -> None:
    """
    Full workflow:
    1. Open a temp copy of the Excel template.
    2. Inject: C9=Name, C11=Size, C12=Pays/Receives, C13=Amount, C19=Size.
    3. Copy B1:C22 as a picture to the clipboard.
    4. Open Outlook draft, paste the picture into the mail body.
    """
    if not name or not name.strip():
        raise ValueError("Product name is required")
    if not isin or not isin.strip():
        raise ValueError("ISIN is required")
    if size <= 0:
        raise ValueError("Size must be positive")
    pays_recv  = "Pays" if increase else "Receives"
    direction  = "Opbouwen" if increase else "Afbouwen"
    c7_label   = "Increase" if increase else "Decrease"
    amount     = (abs(transfer_price) / 100 * size) if transfer_price is not None else ""

    cell_updates = {
        "C7":  c7_label,
        "C9":  name,
        "C10": isin,
        "C11": size,
        "C12": pays_recv,
        "C13": amount,
        "C19": size,
    }

    subject   = f"SP Control: New MTN - {direction} {isin}"

    # Excel must stay alive until after the paste — quitting Excel flushes the clipboard.
    excel = ExcelHandler(TEMPLATE_PATH)
    excel.__enter__()
    try:
        # ── Steps 1-2: inject values ──────────────────────────────────────
        excel.inject_cells(cell_updates, sheet_name=SHEET_NAME)

        # ── Step 3: copy range to clipboard ──────────────────────────────
        excel.copy_range(COPY_RANGE, sheet_name=SHEET_NAME)

        # ── Step 4: create Outlook draft and paste ────────────────────────
        outlook = win32.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)
        mail.To         = ID_MAIL_TO
        mail.CC         = ID_MAIL_CC
        mail.Subject    = subject
        mail.BodyFormat = 2  # olFormatHTML
        mail.Display()

        editor = mail.GetInspector.WordEditor
        sel    = editor.Application.Selection
        sel.Paste()
    finally:
        # Close Excel only after the paste is done
        excel.__exit__(None, None, None)

