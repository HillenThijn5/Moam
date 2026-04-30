from pathlib import Path
import shutil
import tempfile
from uuid import uuid4

try:
    import win32com.client as win32
except ImportError:
    win32 = None


# Excel constants (late-bound safe)
XL_NORMAL  = -4143
XL_SCREEN  = 1
XL_PICTURE = -4147


class ExcelHandler:
    def __init__(self, excel_path: Path):
        if not win32:
            raise RuntimeError("pywin32 not installed")

        self.excel_path = Path(excel_path)

        # Isolated Excel instance — never touches the user's open workbooks
        self.xl = win32.DispatchEx("Excel.Application")
        self.xl.Visible = False
        self.xl.DisplayAlerts = False
        self.xl.ScreenUpdating = False
        self.xl.Interactive = False   # prevents any focus/taskbar flash

        self.wb = None
        self._working_copy = None

    def __enter__(self):
        src = self.excel_path
        tmp_dir = Path(tempfile.gettempdir())

        self._working_copy = tmp_dir / f"{src.stem}_{uuid4().hex}.xlsx"
        shutil.copy2(src, self._working_copy)

        self.wb = self.xl.Workbooks.Open(
            str(self._working_copy),
            UpdateLinks=0,
            ReadOnly=False,
            AddToMru=False,
        )

        self.xl.CutCopyMode = False
        return self

    def __exit__(self, exc_type, exc, tb):
        try:
            if self.wb:
                self.wb.Close(SaveChanges=False)
        finally:
            self.wb = None

        try:
            if self.xl:
                self.xl.Interactive = True
                self.xl.Quit()
        finally:
            self.xl = None

    # -------------------------------------------------
    # Helpers
    # -------------------------------------------------

    def inject_cells(self, updates: dict, sheet_name: str | None = None):
        if not self.wb:
            raise RuntimeError("Workbook not opened")

        ws = self.wb.Worksheets(sheet_name) if sheet_name else self.wb.Worksheets(1)

        for addr, val in updates.items():
            ws.Range(addr).Value = val

        return ws

    def get_sheet(self, name: str):
        for ws in self.wb.Worksheets:
            if ws.Name.strip().lower() == name.strip().lower():
                return ws
        raise RuntimeError(f"Worksheet '{name}' not found")

    def copy_range_as_picture(self, address: str, sheet_name: str | None = None):
        """Copy an arbitrary cell-address range (e.g. 'B1:C22') as a picture to the clipboard."""
        if not self.wb:
            raise RuntimeError("Workbook not opened")

        ws = self.wb.Worksheets(sheet_name) if sheet_name else self.wb.Worksheets(1)
        ws.Activate()

        rng = ws.Range(address)
        rng.CopyPicture(
            Appearance=XL_SCREEN,
            Format=XL_PICTURE,
        )

    def copy_range(self, address: str, sheet_name: str | None = None):
        """Copy an arbitrary cell-address range (e.g. 'B1:C22') normally to the clipboard (pastes as table)."""
        if not self.wb:
            raise RuntimeError("Workbook not opened")

        ws = self.wb.Worksheets(sheet_name) if sheet_name else self.wb.Worksheets(1)
        ws.Activate()
        ws.Range(address).Copy()

    def copy_named_range_as_picture(self, range_name: str):
        if not self.wb:
            raise RuntimeError("Workbook not opened")

        try:
            rng = self.wb.Names(range_name).RefersToRange
        except Exception:
            raise ValueError(f"Named range '{range_name}' does not exist")

        if rng is None or rng.Cells.Count == 0:
            raise ValueError(f"Named range '{range_name}' resolves to empty range")

        sheet = rng.Worksheet
        sheet.Activate()

        self.xl.Visible = False
        self.xl.WindowState = XL_NORMAL
        self.xl.ActiveWindow.Zoom = 100

        rng.CopyPicture(
            Appearance=XL_SCREEN,
            Format=XL_PICTURE,
        )