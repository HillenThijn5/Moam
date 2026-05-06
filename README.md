# MoamProject — Mail Generator

Internal tool for generating structured product mails at Van Lanschot Kempen.

---

## Setup

```bash
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

## Running

```bash
python -m gui.app
```

## Building the EXE

```powershell
powershell -ExecutionPolicy Bypass -File build_exe.ps1
```

Output: `dist\MoamProject.exe` → copy to the shared drive for deployment.

---

## Mail Generator Workflows

### Marketing Mail

Sends product launch announcements to the advisory network with an Excel-based visual.

**Workflow:**
1. Receives product list (1–4 products) with payoff parameters from the GUI
2. Derives product names, denominations, and intro text automatically
3. Injects product data into an Excel template (cell-level mapping per product block)
4. Copies a named range from the Excel template as a picture to the clipboard
5. Creates an Outlook draft with the picture pasted inline + closing text
6. Sends to the full marketing distribution list

**Key data flow:** GUI input → `MarketingProduct` model → `load_marketing_data` → `build_cell_updates` → Excel injection → clipboard image → Outlook

---

### PC Mail (Product Committee Notification)

Notifies internal operations teams of new private placements with Word + Excel attachments.

**Workflow:**
1. Receives full product specification from the GUI (payoff, fees, client, underlyings)
2. Enriches underlyings with benchmark data from `static_sheet.xlsx`
3. Computes ESG score to determine target market classification (positive/neutral)
4. Builds Word template context (payoff text, fee text, compliance text, hyperlinks)
5. Renders the Word document (Individual Product Notification) from a `.docx` template
6. Renders the Target Market Excel (sets B24 cell value + color-coded fill)
7. Composes the email body with SAS action items (to-do list)
8. Opens Outlook draft with both documents attached

**Key data flow:** GUI input → `PCMailProduct` model → builders (payoff, fee, subject, links) → Word render + Excel render → Outlook with attachments

---

### Documentatie Mail

Sends documentation packages (brochure, Final Terms, EID link) to advisers after a client purchase.

**Workflow:**
1. Receives trade details (adviser, amount, price) and product info from the GUI
2. Builds the email subject with issuer, product, underlyings, and maturity range
3. Generates the HTML body with trade list, VL code, product brochure link, and EID link
4. Derives CC recipients from adviser → helper mapping (from `static_sheet.xlsx`)
5. Opens Outlook draft with signature preserved, resolves all recipient names

**Key data flow:** GUI input → `build_email_data` → `build_subject` + `build_body` → Outlook with HTML body + signature

---

### Increase / Decrease Mail

Notifies treasury and ops teams when an MTN position is being increased or decreased.

**Workflow:**
1. Queries the `StructuredProducts2010` database for current MTN positions and latest transfer price
2. User selects a position and direction (increase/decrease) in the GUI
3. Calculates the transfer amount (`transfer_price / 100 × size`)
4. Injects deal values into an Excel template (`IncreaseDecrease` sheet)
5. Copies the filled range (B1:C22) to the clipboard
6. Pastes the table into an Outlook draft (preserves Excel formatting)

**Key data flow:** SQL query → position selection → Excel injection → clipboard copy → Outlook paste

---

### SharePoint Integration

Loads live deal data from the SharePoint "New Notes Summary" for pre-filling the GUI.

**Workflow:**
1. On startup, copies the OneDrive-synced `sharepointsummary.xlsx` to the local project
2. If no OneDrive copy exists, opens Excel via COM, runs `RefreshAll`, and saves
3. Parses the SharePoint title format to extract product type, maturity, and underlyings
4. Parses the comments field to extract trade amounts, prices, and hedge details
5. Presents parsed deals in the GUI for one-click population of any mail tab

---

## Static Data

All frequently-changing configuration lives in **`statics/static_sheet.xlsx`**, which sits next to the EXE when deployed. The app loads and caches every sheet once at startup — restart the app after any edit.

### Benchmarks sheet

Drives: underlying dropdowns in every tab, benchmark fields in the PC Mail Word document, full names in Marketing Mail, short aliases in Documentatie Mail subjects.

| Column | What to put there |
|--------|-------------------|
| `BENCHMARK ID` | Bloomberg/internal ticker, e.g. `SX5E`. Row order controls dropdown order. |
| `BENCHMARK NAME` | Primary benchmark name, e.g. `EURO STOXX 50 Index`. **Leave blank for funds/ETFs** (e.g. EEM UP) — the app skips blank rows when building the benchmark map. |
| `ALTERNATIVE BENCHMARK NAME` | Fallback benchmark shown when the primary is not available. Leave blank if none. |
| `FULL_NAME` | Full display name used in the Marketing Mail Excel output, e.g. `S&P 500 Index`. Type `&` literally. |
| `ALIAS` | Short name used in email subjects and product titles, e.g. `VS`, `JPN`. Leave blank to use the ticker as-is. |

**To add a new underlying:** add a new row. To remove one: delete the row.

---

### ProductURLs sheet

Drives: brochure and video hyperlinks in the PC Mail Word document; brochure link in Documentatie Mail; video link in Marketing Mail.

| Column | What to put there |
|--------|-------------------|
| `PRODUCT_TYPE` | Exact product name as shown in the app, e.g. `Trigger Plus Note` |
| `BROCHURE_LABEL` | Link text in the Word doc, e.g. `General Brochure Trigger Notes` |
| `BROCHURE_URL` | Full URL to the brochure PDF |
| `VIDEO_URL` | Full Vimeo/YouTube URL. **Leave blank** if there is no video for this product. |
| `VIDEO_LABEL_MARKETING` | Dutch link text for Marketing Mail, e.g. `Hoe werkt een Trigger note?` Leave blank if no video. |
| `VIDEO_LABEL_PC` | English link text for PC Mail Word doc, e.g. `Trigger Notes – Product Video`. Leave blank if no video. |

**When a new brochure is published:** update `BROCHURE_URL` (and optionally `BROCHURE_LABEL`) for the relevant row(s).
**When a new video is published:** update `VIDEO_URL` and both label columns.

---

### PARP sheet

Drives: the PARP date field in the PC Mail (Individual Product Notification) Word document.

| Column | What to put there |
|--------|-------------------|
| `PRODUCT_TYPE` | Exact product name, e.g. `Trigger Plus Note` |
| `PARP_DATE` | Approval date as plain text, e.g. `24 September 2025`. Excel date values also work. |

**Update this sheet every time a new SNIP or DIP prospectus is approved.**

---

### ProspectusURLs sheet

Drives: the prospectus hyperlink in the PC Mail Word document.

| Column | What to put there |
|--------|-------------------|
| `CODE` | `DIP` or `SNIP` (uppercase) |
| `LABEL` | Link text shown in the document, usually the same as `CODE` |
| `URL` | Full URL to the Securities Note on vanlanschotkempen.com |

**Update the URL each time a new Securities Note is published.**

---

### ESG sheet

Drives: target market classification (positive/neutral) in the PC Mail Excel attachment.

| Column | What to put there |
|--------|-------------------|
| `TICKER` | Same ticker as in the Benchmarks sheet |
| `SCORE` | Integer ESG score |

---

### advisers sheet

Drives: the CC field in Documentatie Mail — automatically adds the adviser's helper when a known adviser is selected.

| Column | What to put there |
|--------|-------------------|
| A (no header needed) | Adviser full name |
| B | Helper full name to add as CC |

---

### What is still hardcoded in `statics/data.py`

The following data requires a code change (edit `statics/data.py` and rebuild the EXE) because it reflects structural app behaviour rather than regularly-changing business data:

| Constant | What it contains | When you'd change it |
|----------|-----------------|----------------------|
| `HEDGEPARTY` | List of hedge counterparties in the PC Mail dropdown | New counterparty added or removed |
| `ISSUERS` | Issuer dropdown (PC Mail) | New issuer onboarded |
| `CLIENTS` | Client dropdown (PC Mail) | New client relationship added |
| `CURRENCIES` | Currency dropdown | New currency supported |
| `MATURITIES` | Maturity dropdown values | Range of maturities changes |
| `MARKETING_ISSUERS` | Issuer strings including credit ratings (Marketing Mail) | Rating changes or new issuer |
| `MARKETING_MAIL_TO_LIST` | Full marketing distribution list | Adviser added/removed from distribution |
| `PC_MAIL_TO_RECIPIENTS` | Internal ops/risk/settlement TO addresses (PC Mail) | Team email addresses change |
| `ID_MAIL_TO` / `ID_MAIL_CC` | Increase/decrease mail recipients | Team email addresses change |
| `DOCUMENTATIE_MAIL_RECIPIENTS` | Fixed CC for Documentatie Mail | Fixed CC recipients change |
| `PRODUCT_PAYOFF_FIELDS` | GUI form field labels per product type | New product type or field label changes |
| `MARKETING_PARAMETER_CONFIG` | GUI form structure for Marketing Mail tab | New product type or form layout changes |
| `ID_MAIL_DB_CONN_STR` | SQL Server connection string | Database server or database name changes |
| `PRIIP_HUB_URL` | PRIIP Hub API endpoint | API endpoint changes |

---

## Project Structure

```
MoamProject/
├── gui/              → PyQt GUI (tabs, dialogs, widgets)
├── MarketingMail/    → Marketing mail generator
├── PCMail/           → PC notification mail generator
├── Documentatie_Mail/→ Documentation mail generator
├── increase_decrease_mail/ → Increase/decrease mail generator
├── sharepoint/       → SharePoint reader & parser
├── statics/          → Shared static data, lookups, and Excel loader
└── assets/           → Icons and images
```
