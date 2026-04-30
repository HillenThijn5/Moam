# PCMail/config/paths.py

from pathlib import Path
import sys
import tempfile


def project_root() -> Path:
    """
    Returns the project root folder (MoamProject),
    both in dev and in PyInstaller onefile EXE.
    """
    if hasattr(sys, "_MEIPASS"):
        # PyInstaller onefile temp dir
        return Path(sys._MEIPASS)
    else:
        # paths.py → config → PCMail → MoamProject
        return Path(__file__).resolve().parents[2]


def resource_path(relative_path: str) -> Path:
    """
    Build absolute path to a bundled resource.
    `relative_path` must be relative to project root.
    """
    return project_root() / relative_path


# ------------------------------------------------------------------
# PCMail resources
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
# Marketing mail resources
# ------------------------------------------------------------------

MARKETING_EXCEL_TEMPLATE_PATH = resource_path(
    "MarketingMail/templates/excel(marketingmail).xlsx"
)

# ------------------------------------------------------------------
# Temp output folder (always writable, outside EXE)
# ------------------------------------------------------------------

TEMP_DIR = Path(tempfile.gettempdir()) / "pc_mail_generator"
TEMP_DIR.mkdir(parents=True, exist_ok=True)
