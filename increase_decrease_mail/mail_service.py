# increase_decrease_mail/mail_service.py
"""
Zet dealwaarden in het Excel-template, leest de celwaarden uit
en maakt een Outlook-concept met de inhoud als gewone HTML-tekst (zonder afbeelding/randen).
"""
from __future__ import annotations
import os
import re
import sys
from pathlib import Path

import win32com.client as win32

from statics.data import ID_MAIL_TO, ID_MAIL_CC
from MarketingMail.excel_handler import ExcelHandler


def _get_signature_html() -> str:
    """Lees de standaard-Outlookhandtekening van het bestandssysteem."""
    sig_dir = Path(os.environ.get("APPDATA", "")) / "Microsoft" / "Signatures"
    if not sig_dir.exists():
        return ""
    htm_files = list(sig_dir.glob("*.htm"))
    if not htm_files:
        return ""
    sig_file = htm_files[0]
    try:
        content = sig_file.read_text(encoding="utf-8")
        body_match = re.search(r'<body[^>]*>(.*)</body>', content, re.DOTALL | re.IGNORECASE)
        if body_match:
            return body_match.group(1).strip()
        return content.strip()
    except Exception:
        return ""


def _exe_root() -> Path:
    """Projectmap: de map van de exe bij een gebundelde build, anders de bronmap van MoamProject."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent
    return Path(__file__).resolve().parents[1]


TEMPLATE_PATH = _exe_root() / "increase_decrease_mail" / "increasedecrease(mailtemplate).xlsx"
SHEET_NAME    = "IncreaseDecrease"
COPY_RANGE    = "B1:C22"


def _read_range_as_text(excel: ExcelHandler, address: str, sheet_name: str) -> list[list[str]]:
    """Lees celwaarden uit een bereik en geef ze terug als lijst met rijen (lijst met strings)."""
    ws = excel.wb.Worksheets(sheet_name) if sheet_name else excel.wb.Worksheets(1)
    rng = ws.Range(address)
    values = rng.Value  # Geeft een tuple van tuples terug
    if values is None:
        return []
    rows = []
    for row in values:
        rows.append([str(cell) if cell is not None else "" for cell in row])
    return rows


def _read_range_with_style(excel: ExcelHandler, address: str, sheet_name: str) -> list[list[dict]]:
    """
    Lees celwaarden EN opmaak uit een Excel-bereik via COM.
    Geeft een lijst met rijen terug; elke rij is een lijst met dicts:
      { "value": str, "bold": bool, "italic": bool, "size": float,
        "color": str (hex), "font_name": str }
    """
    ws = excel.wb.Worksheets(sheet_name) if sheet_name else excel.wb.Worksheets(1)
    rng = ws.Range(address)
    rows = []
    for r in range(1, rng.Rows.Count + 1):
        row_data = []
        for c in range(1, rng.Columns.Count + 1):
            cell = rng.Cells(r, c)
            val = cell.Value
            val_str = str(val) if val is not None else ""
            font = cell.Font
            # Lees de kleur uit (geeft een long int terug, zet om naar hex-RGB)
            color_long = font.Color
            if color_long is not None and color_long != 0:
                # Excel slaat dit op als BGR, COM geeft een float terug
                color_int = int(color_long)
                b = (color_int >> 16) & 0xFF
                g = (color_int >> 8) & 0xFF
                r_val = color_int & 0xFF
                color_hex = f"#{r_val:02x}{g:02x}{b:02x}"
            else:
                color_hex = ""
            row_data.append({
                "value": val_str,
                "bold": bool(font.Bold),
                "italic": bool(font.Italic),
                "size": font.Size or 11,
                "font_name": font.Name or "Segoe UI",
                "color": color_hex,
            })
        rows.append(row_data)
    return rows


def _fmt_value(val: str) -> str:
    """Formatteer een celwaarde voor weergave — herstel floats en gebruik Nederlandse getalnotatie."""
    val = val.strip()
    if not val:
        return ""
    # Los float-precisieproblemen op (bijv. "11.899999999999999" → "11.90")
    try:
        f = float(val)
        # Controle op gehele getallen (bijv. 20000.0 → "20000")
        if f == int(f) and '.' in val and not any(c.isalpha() for c in val):
            # Formatteer als geheel getal met Nederlandse duizendtallen
            return f"{int(f):,}".replace(",", ".")
        # Float met nette precisie
        if '.' in val and len(val.split('.')[1]) > 4:
            # Rond af op 2 decimalen, in Nederlandse notatie
            formatted = f"{f:,.2f}"
            # Zet om naar Nederlands: 27,661.20 → 27.661,20
            formatted = formatted.replace(",", "X").replace(".", ",").replace("X", ".")
            return formatted
    except (ValueError, OverflowError):
        pass
    return val


def _rows_to_html(rows: list[list[dict]]) -> str:
    """
    Zet opgemaakte rijen om naar HTML en behoud de Excel-opmaak:
    - Vet, cursief, letterkleur en lettergrootte komen allemaal uit het spreadsheet.
    - Geen tabelranden (strakke look zoals het origineel).
    - Lege rijen worden witruimte.
    """
    html_rows = []
    for row in rows:
        # Kijk of de hele rij leeg is
        if all(not cell["value"].strip() for cell in row):
            html_rows.append('<tr><td colspan="2">&nbsp;</td></tr>')
            continue

        cells_html = []
        for i, cell in enumerate(row):
            val = _fmt_value(cell["value"])
            if not val:
                cells_html.append('<td style="padding-right:30px;">&nbsp;</td>' if i == 0 else '<td>&nbsp;</td>')
                continue

            # Bouw inline styles op vanuit de Excel-opmaak
            styles = []
            if cell["font_name"]:
                styles.append(f"font-family:{cell['font_name']}")
            if cell["size"] and cell["size"] != 11:
                styles.append(f"font-size:{cell['size']}pt")
            if cell["color"]:
                styles.append(f"color:{cell['color']}")
            if i == 0:
                styles.append("padding-right:30px")

            style_str = f' style="{"; ".join(styles)}"' if styles else ""

            # Zet het in bold/italic-tags zoals in Excel
            inner = val
            if cell["bold"]:
                inner = f"<b>{inner}</b>"
            if cell["italic"]:
                inner = f"<i>{inner}</i>"

            cells_html.append(f"<td{style_str}>{inner}</td>")

        html_rows.append(f'<tr>{"".join(cells_html)}</tr>')

    return (
        '<table style="border:none; border-collapse:collapse; font-family:Segoe UI; font-size:11pt;">\n'
        + "\n".join(html_rows)
        + "\n</table>"
    )


def send_increase_decrease_mail(
    name: str,
    isin: str,
    size: float,
    transfer_price: float | None,
    increase: bool,
) -> None:
    """
    Volledige flow:
    1. Open een tijdelijke kopie van het Excel-template.
    2. Vul de waarden in de cellen.
    3. Lees B1:C22 uit als tekstwaarden.
    4. Maak een Outlook-concept met gewone HTML-tekst (zonder afbeelding of celranden).
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

    # Bereken het bedrag met nette afronding om float-precisieproblemen te vermijden
    if transfer_price is not None:
        amount_raw = abs(transfer_price) / 100 * size
        amount = round(amount_raw, 2)
        # Formatteer als Nederlandse valuta: EUR 27.661,20
        amount_str = f"EUR {amount:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    else:
        amount_str = ""

    # Formatteer size als geheel getal (zonder decimalen)
    size_int = int(size)

    cell_updates = {
        "C7":  c7_label,
        "C9":  name,
        "C10": isin,
        "C11": size_int,
        "C12": pays_recv,
        "C13": amount_str,
        "C19": size_int,
    }

    subject   = f"SP Control: New MTN - {direction} {isin}"

    # ── Stappen 1-3: waarden invullen en als tekst uitlezen ─────────────
    with ExcelHandler(TEMPLATE_PATH) as excel:
        excel.inject_cells(cell_updates, sheet_name=SHEET_NAME)
        rows = _read_range_with_style(excel, COPY_RANGE, sheet_name=SHEET_NAME)

    # ── Stap 4: Outlook-concept maken met gewone HTML-tekst ────────────
    body_html = _rows_to_html(rows)

    outlook = win32.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)
    mail.To         = ID_MAIL_TO
    mail.CC         = ID_MAIL_CC
    mail.Subject    = subject
    mail.BodyFormat = 2  # olFormatHTML

    sig_html = _get_signature_html()
    mail.HTMLBody = f"""<html><body style="font-family:Segoe UI; font-size:11pt;">
{body_html}
<br>
{sig_html}
</body></html>"""

    mail.Display()

