# MarketingMail/mail_service.py
import io
import os
import re
import tempfile
import time
from pathlib import Path

import win32clipboard
import win32com.client as win32
from datetime import date, timedelta, datetime
from PIL import Image

from MarketingMail.config import EXCEL_PATH
from statics.data import MARKETING_MAIL_TO_LIST, STRUCTURED_INVESTMENTS_EMAIL, MARKETING_MAIL_SIGNATURE
from MarketingMail.excel_handler import ExcelHandler
from MarketingMail.injection import build_cell_updates
from MarketingMail.dataloader import load_marketing_data
from statics.outlook_user import get_sender_first_name


CF_DIB = 8


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


def _crop_border(img_path: Path, pixels: int = 2) -> None:
    """Snijd boven en links een paar pixels weg om Excel-randjes te verwijderen."""
    try:
        img = Image.open(str(img_path))
        w, h = img.size
        # Alleen croppen als de afbeelding groot genoeg is en het niet opvalt
        if w > pixels * 4 and h > pixels * 4:
            cropped = img.crop((pixels, pixels, w, h))
            cropped.save(str(img_path), format="PNG")
    except Exception:
        pass


def _clipboard_to_png_file() -> Path:
    """Lees de huidige afbeelding van het klembord, sla die op als tijdelijke PNG en geef het pad terug."""
    from PIL import ImageGrab
    import time as _time

    tmp = Path(tempfile.gettempdir()) / "moam_marketing_mail.png"

    # Probeer een paar keer opnieuw — Excel is soms traag met klembordacties
    for attempt in range(5):
        _time.sleep(0.5)
        try:
            img = ImageGrab.grabclipboard()
            if img is not None:
                img.save(str(tmp), format="PNG")
                # Kijk of het bestand niet leeg of corrupt is
                if tmp.stat().st_size > 500:
                    return tmp
        except Exception:
            pass

    # Terugvaloptie: probeer handmatig CF_DIB uit te lezen
    try:
        win32clipboard.OpenClipboard()
        try:
            if win32clipboard.IsClipboardFormatAvailable(CF_DIB):
                data = win32clipboard.GetClipboardData(CF_DIB)
                if len(data) > 100:
                    # Lees BITMAPINFOHEADER uit om de juiste offset te krijgen
                    import struct
                    header_size = struct.unpack_from('<I', data, 0)[0]
                    bit_count = struct.unpack_from('<H', data, 14)[0]
                    # Grootte van de kleurentabel
                    colors_used = struct.unpack_from('<I', data, 32)[0]
                    if bit_count <= 8:
                        palette_size = (colors_used or (1 << bit_count)) * 4
                    else:
                        palette_size = colors_used * 4
                    pixel_offset = 14 + header_size + palette_size
                    file_size = 14 + len(data)
                    bmp_header = b'BM' + struct.pack('<I', file_size) + b'\x00\x00\x00\x00' + struct.pack('<I', pixel_offset)
                    img = Image.open(io.BytesIO(bmp_header + data))
                    img.save(str(tmp), format="PNG")
                    if tmp.stat().st_size > 500:
                        return tmp
        finally:
            win32clipboard.CloseClipboard()
    except Exception:
        pass

    raise RuntimeError(
        "No valid image data found on clipboard after Excel copy. "
        "The hidden Excel instance may not be putting images on the clipboard correctly. "
        "Try closing all Excel instances and retry."
    )


_DUTCH_MONTHS = ["jan", "feb", "mrt", "apr", "mei", "jun",
                 "jul", "aug", "sep", "okt", "nov", "dec"]
_DUTCH_DAYS   = {0: "maandag", 1: "dinsdag", 2: "woensdag",
                 3: "donderdag", 4: "vrijdag"}


def _fmt_nl(d: date) -> str:
    """Formatteer een datum als korte Nederlandse notatie: '17 apr 26'."""
    return f"{d.day} {_DUTCH_MONTHS[d.month - 1]} {str(d.year)[2:]}"


def _build_closing_text(start_date_str: str) -> str:
    try:
        start = datetime.strptime(start_date_str, "%d %b %Y").date()
    except ValueError:
        start = date.today()

    next_bd = start + timedelta(days=1)
    while next_bd.weekday() >= 5:
        next_bd += timedelta(days=1)

    if next_bd == start + timedelta(days=1):
        vanaf = f"vanaf morgen ({_fmt_nl(next_bd)})"
    else:
        vanaf = f"vanaf {_DUTCH_DAYS.get(next_bd.weekday(), '')} ({_fmt_nl(next_bd)})"

    return (
        f"De startwaarde van de onderliggende waarde worden bepaald obv de slotkoers van vandaag ({_fmt_nl(start)}).\r\n"
        f"Aanhaken is vandaag nog mogelijk tegen de uitgifteprijs van 100%, {vanaf} tegen de actuele waarde."
    )


def create_and_send_marketing_mail(
        title: str,
        products,
        choice: int,
        to: str = None,
        cc: str = None,
):
    """
    Volledige flow: data laden → Excel vullen → als afbeelding kopiëren → Outlook-mail maken.

    Parameters:
        title:    Onderwerp van de mail / titel voor Excel B9
        products: Lijst met MarketingProduct-objecten
        choice:   1–4 (aantal productblokken om mee te nemen)
        to:       Overschrijft TO_LIST (optioneel)
        cc:       Overschrijft CC_LIST (optioneel)
    """
    if not 1 <= choice <= 4:
        raise ValueError("choice must be between 1 and 4")

    if not products:
        raise ValueError("At least one product is required")

    if len(products) < choice:
        raise ValueError(
            f"choice={choice} but only {len(products)} products provided"
        )

    to = to or MARKETING_MAIL_TO_LIST
    cc = cc or STRUCTURED_INVESTMENTS_EMAIL

    # Outlook verwacht een string met puntkomma's, geen lijst
    if isinstance(to, list):
        to = "; ".join(to)
    if isinstance(cc, list):
        cc = "; ".join(cc)

    # ===== STAP 1: Data laden en voorbereiden =====
    data = load_marketing_data(title, products)
    updates = build_cell_updates(data, choice)

    # ===== STAP 2: In Excel zetten en afbeelding kopiëren =====
    with ExcelHandler(EXCEL_PATH) as excel:
        excel.inject_cells(updates, sheet_name="Maatwerk Notes Nieuw")
        excel.copy_named_range_as_bitmap(f"only{choice}")

    # ===== STAP 3: Mailinhoud opbouwen =====
    start_date_str = products[0].start_date if products else date.today().strftime("%d %b %Y")
    closing = _build_closing_text(start_date_str)
    body_text = closing

    # ===== STAP 4: Klembordafbeelding opslaan en Outlook-mail maken =====
    # Beperkingen (Outlook 16.0.19929+ security-update):
    # - PropertyAccessor is geblokkeerd → CID handmatig zetten kan niet
    # - WordEditor is geblokkeerd → plakken vanaf het klembord kan niet
    # - data:-URI's, file://-URI's en gewone paden tonen allemaal niets
    # Aanpak: voeg de afbeelding toe als bijlage + verwijs ernaar via cid:bestandsnaam.
    # Outlook koppelt de bijlage automatisch aan de cid als de bestandsnaam matcht.
    # Als dat niet inline rendert, staat de afbeelding in elk geval als bijlage erbij.
    png_path = _clipboard_to_png_file()
    _crop_border(png_path)  # Haal randartefacten linksboven weg
    img_filename = png_path.name  # "moam_marketing_mail.png"

    safe_text = body_text.replace("\r\n", "<br>").replace("\n", "<br>")

    outlook = win32.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)
    mail.To = to
    if cc:
        mail.CC = cc
    mail.Subject = f"Nieuwe Maatwerk Structured Note: {', '.join([p.product_name for p in products[:choice]])}"
    mail.BodyFormat = 2  # olFormatHTML

    # Voeg de bijlage toe VOORDAT je HTMLBody zet
    mail.Attachments.Add(str(png_path))

    # Lees de handtekening van de gebruiker
    sig_html = _get_signature_html()

    # Zet HTMLBody en verwijs met de bestandsnaam als CID naar de bijlage
    mail.HTMLBody = f"""<html><body>
<img src="cid:{img_filename}"><br><br>
<p>{safe_text}</p>
<br>
{sig_html}
</body></html>"""

    mail.Display()
    return mail
