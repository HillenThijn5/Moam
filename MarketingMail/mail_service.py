# MarketingMail/mail_service.py
import win32com.client as win32
from datetime import date, timedelta, datetime

from MarketingMail.config import EXCEL_PATH
from statics.data import MARKETING_MAIL_TO_LIST, STRUCTURED_INVESTMENTS_EMAIL, MARKETING_MAIL_SIGNATURE
from MarketingMail.excel_handler import ExcelHandler
from MarketingMail.injection import build_cell_updates
from MarketingMail.dataloader import load_marketing_data
from statics.outlook_user import get_sender_first_name


_DUTCH_MONTHS = ["jan", "feb", "mrt", "apr", "mei", "jun",
                 "jul", "aug", "sep", "okt", "nov", "dec"]
_DUTCH_DAYS   = {0: "maandag", 1: "dinsdag", 2: "woensdag",
                 3: "donderdag", 4: "vrijdag"}


def _fmt_nl(d: date) -> str:
    """Format a date as Dutch short: '17 apr 26'."""
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
    Complete workflow: load data → inject Excel → copy as picture → create Outlook mail.

    Args:
        title:    Email subject / Excel B9 title
        products: List of MarketingProduct objects
        choice:   1–4 (number of product blocks to include)
        to:       Override TO_LIST (optional)
        cc:       Override CC_LIST (optional)
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

    # Outlook expects a semicolon-separated string, not a list
    if isinstance(to, list):
        to = "; ".join(to)
    if isinstance(cc, list):
        cc = "; ".join(cc)

    # ===== STEP 1: Load and prepare data =====
    data = load_marketing_data(title, products)
    updates = build_cell_updates(data, choice)

    # ===== STEP 2: Inject into Excel & copy image =====
    with ExcelHandler(EXCEL_PATH) as excel:
        excel.inject_cells(updates, sheet_name="Maatwerk Notes Nieuw")
        excel.copy_named_range_as_picture(f"only{choice}")

    # ===== STEP 3: Build mail body =====
    start_date_str = products[0].start_date if products else date.today().strftime("%d %b %Y")
    closing = _build_closing_text(start_date_str)
    body_text = f"{closing}\r\n\r\nGroet,\r\n{get_sender_first_name() or MARKETING_MAIL_SIGNATURE}"

    # ===== STEP 4: Create Outlook mail =====
    outlook = win32.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)
    mail.To = to
    if cc:
        mail.CC = cc

    mail.Subject = f"Nieuwe Maatwerk Structured Note: {', '.join([p.product_name for p in products[:choice]])}"
    mail.BodyFormat = 2
    mail.Display()

    # ===== STEP 5: Paste image into mail body =====
    editor = mail.GetInspector.WordEditor
    sel = editor.Application.Selection

    sel.Paste()
    sel.Collapse(0)  # wdCollapseEnd — move cursor past pasted image to text position
    sel.TypeParagraph()
    sel.TypeParagraph()
    sel.TypeParagraph()
    sel.TypeText(body_text)

    return mail
