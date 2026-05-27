from pathlib import Path
import atexit
import shutil
import tempfile
from uuid import uuid4

try:
    import win32com.client as win32
    import win32gui
    import win32process
except ImportError:
    win32 = None
    win32gui = None
    win32process = None


# Excel-constanten (veilig bij late binding)
XL_NORMAL   = -4143
XL_SCREEN   = 1       # snelle opname van wat op het scherm zou staan (veilig — onze instantie is verborgen)
XL_PICTURE  = -4147
SW_HIDE     = 0

# ── Blijvende verborgen Excel-instantie ─────────────────────────────
# Excel.exe starten is de grootste bottleneck (~3-5 s). Door één
# verborgen instantie open te houden betalen we die kosten maar één keer per appsessie.
_shared_xl = None


def _force_hide(xl) -> None:
    """Verberg geforceerd ALLE vensters van ons verborgen Excel-proces."""
    try:
        pid = xl.ProcessID
    except Exception:
        return

    def _hide_cb(hwnd, _):
        try:
            _, wpid = win32process.GetWindowThreadProcessId(hwnd)
            if wpid == pid:
                win32gui.ShowWindow(hwnd, SW_HIDE)
        except Exception:
            pass
        return True  # ga door met enumereren

    try:
        win32gui.EnumWindows(_hide_cb, None)
    except Exception:
        pass


def _get_shared_xl():
    """Geef een verborgen Excel Application terug en maak die pas aan als het nodig is."""
    global _shared_xl
    if _shared_xl is not None:
        try:
            _shared_xl.Visible  # noqa: B018  — controle; gooit een exceptie als het proces weg is
            return _shared_xl
        except Exception:
            _shared_xl = None

    xl = win32.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    xl.ScreenUpdating = False
    xl.Interactive = False
    _force_hide(xl)
    _shared_xl = xl
    return xl


def _shutdown_excel():
    """Sluit de gedeelde Excel af zodra het Python-proces stopt."""
    global _shared_xl
    if _shared_xl is not None:
        try:
            _shared_xl.Interactive = True
            _shared_xl.Quit()
        except Exception:
            pass
        _shared_xl = None


atexit.register(_shutdown_excel)


class ExcelHandler:
    def __init__(self, excel_path: Path):
        if not win32:
            raise RuntimeError("pywin32 not installed")

        self.excel_path = Path(excel_path)
        self.xl = _get_shared_xl()

        self.wb = None
        self._working_copy = None

    def __enter__(self):
        src = self.excel_path
        tmp_dir = Path(tempfile.gettempdir())

        self._working_copy = tmp_dir / f"{src.stem}_{uuid4().hex}.xlsx"
        shutil.copy2(src, self._working_copy)

        # Onthoud in welk venster de gebruiker zit, zodat we dat straks kunnen herstellen.
        self._prev_fg = None
        try:
            self._prev_fg = win32gui.GetForegroundWindow()
        except Exception:
            pass

        _force_hide(self.xl)
        self.wb = self.xl.Workbooks.Open(
            str(self._working_copy),
            UpdateLinks=0,
            ReadOnly=False,
            AddToMru=False,
        )
        _force_hide(self.xl)

        self.xl.CutCopyMode = False
        return self

    def __exit__(self, exc_type, exc, tb):
        try:
            if self.wb:
                self.wb.Close(SaveChanges=False)
        finally:
            self.wb = None
            _force_hide(self.xl)
            # Zet het venster terug waarin de gebruiker zat voordat we begonnen.
            try:
                if self._prev_fg and win32gui.IsWindow(self._prev_fg):
                    win32gui.SetForegroundWindow(self._prev_fg)
            except Exception:
                pass

    # -------------------------------------------------
    # Hulpfuncties
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
        """Kopieer een willekeurig celbereik (bijv. 'B1:C22') als afbeelding naar het klembord."""
        if not self.wb:
            raise RuntimeError("Workbook not opened")

        ws = self.wb.Worksheets(sheet_name) if sheet_name else self.wb.Worksheets(1)
        ws.Activate()
        _force_hide(self.xl)

        # Verberg rasterlijnen zodat ze niet in de afbeelding terechtkomen
        try:
            self.xl.ActiveWindow.DisplayGridlines = False
        except Exception:
            pass

        # ScreenUpdating moet AAN staan zodat CopyPicture echte pixels kan pakken
        self.xl.ScreenUpdating = True
        try:
            rng = ws.Range(address)
            rng.CopyPicture(
                Appearance=XL_SCREEN,
                Format=XL_PICTURE,
            )
        finally:
            self.xl.ScreenUpdating = False
            _force_hide(self.xl)

    def copy_range_as_bitmap(self, address: str, sheet_name: str | None = None):
        """Kopieer een willekeurig celbereik als bitmap naar het klembord (CF_DIB-formaat)."""
        if not self.wb:
            raise RuntimeError("Workbook not opened")

        ws = self.wb.Worksheets(sheet_name) if sheet_name else self.wb.Worksheets(1)
        ws.Activate()
        _force_hide(self.xl)

        # Verberg rasterlijnen zodat ze niet in de afbeelding terechtkomen
        try:
            self.xl.ActiveWindow.DisplayGridlines = False
        except Exception:
            pass

        # ScreenUpdating moet AAN staan zodat CopyPicture echte pixels kan pakken
        self.xl.ScreenUpdating = True
        try:
            rng = ws.Range(address)
            rng.CopyPicture(
                Appearance=XL_SCREEN,
                Format=2,  # xlBitmap — zet CF_DIB op het klembord
            )
        finally:
            self.xl.ScreenUpdating = False
            _force_hide(self.xl)

    def copy_range(self, address: str, sheet_name: str | None = None):
        """Kopieer een willekeurig celbereik (bijv. 'B1:C22') normaal naar het klembord (plakt als tabel)."""
        if not self.wb:
            raise RuntimeError("Workbook not opened")

        ws = self.wb.Worksheets(sheet_name) if sheet_name else self.wb.Worksheets(1)
        ws.Activate()
        _force_hide(self.xl)
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
        _force_hide(self.xl)

        self.xl.ScreenUpdating = True
        try:
            self.xl.ActiveWindow.Zoom = 100
            rng.CopyPicture(
                Appearance=XL_SCREEN,
                Format=XL_PICTURE,
            )
        finally:
            self.xl.ScreenUpdating = False
            _force_hide(self.xl)

    def copy_named_range_as_bitmap(self, range_name: str):
        """Kopieer een benoemde range als bitmap naar het klembord (CF_DIB-formaat)."""
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
        _force_hide(self.xl)

        # Verberg rasterlijnen om randartefacten aan de zijkanten te voorkomen
        try:
            self.xl.ActiveWindow.DisplayGridlines = False
        except Exception:
            pass

        self.xl.ScreenUpdating = True
        try:
            self.xl.ActiveWindow.Zoom = 100
            rng.CopyPicture(
                Appearance=XL_SCREEN,
                Format=2,  # xlBitmap — zet CF_DIB op het klembord
            )
        finally:
            self.xl.ScreenUpdating = False
            _force_hide(self.xl)


