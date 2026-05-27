# PCMail/config/paths.py

from pathlib import Path
import sys
import tempfile


def project_root() -> Path:
    """
    Geeft de hoofdmap van het project (MoamProject) terug,
    zowel in ontwikkeling als in een PyInstaller onefile-exe.
    Externe databestanden (xlsx/docx) staan naast de EXE, niet erin.
    """
    if getattr(sys, "frozen", False):
        # Onefile-EXE: databestanden staan naast het uitvoerbare bestand
        return Path(sys.executable).parent
    else:
        # paths.py → config → PCMail → MoamProject
        return Path(__file__).resolve().parents[2]


def resource_path(relative_path: str) -> Path:
    """
    Bouw een absoluut pad naar een meegeleverd bestand.
    `relative_path` moet relatief zijn aan de projectroot.
    """
    return project_root() / relative_path


# ------------------------------------------------------------------
# PCMail-bronnen
# ------------------------------------------------------------------

WORD_TEMPLATE_PATH = resource_path(
    "PCMail/template/Individual Product Notification - template.docx"
)

EXCEL_TM_TEMPLATE_PATH = resource_path(
    "PCMail/template/pcmailtargetmarkets.xlsx"
)

STATIC_DATA_PATH = resource_path(
    "statics/static_sheet.xlsx"
)

# ------------------------------------------------------------------
# Marketingmail-bronnen
# ------------------------------------------------------------------

MARKETING_EXCEL_TEMPLATE_PATH = resource_path(
    "MarketingMail/templates/excel(marketingmail).xlsx"
)

# ------------------------------------------------------------------
# Tijdelijke uitvoermap (altijd schrijfbaar, buiten de EXE)
# ------------------------------------------------------------------

TEMP_DIR = Path(tempfile.gettempdir()) / "pc_mail_generator"
TEMP_DIR.mkdir(parents=True, exist_ok=True)
