# sharepoint/reader.py
"""
Reads (and optionally refreshes) the SharePoint "New Notes Summary" Excel file.
The file is cached locally via OneDrive sync, so we copy it to a temp path
to avoid permission errors when the file is also open in Excel.

Refresh strategy: call start_background_refresh() once at app startup.
The picker dialog just reads; it never triggers another refresh.
"""
from __future__ import annotations

import shutil
import tempfile
import threading
import time
from pathlib import Path

import openpyxl

SHEET_NAME = "New Notes Summary"

# Module-level state so the app can check / wait on the startup refresh
_refresh_thread: threading.Thread | None = None
_refresh_done   = threading.Event()
_refresh_error: str = ""


def start_background_refresh(excel_path: str | Path, source_path: str | Path | None = None) -> None:
    """
    Kick off a background refresh.  Safe to call from the main thread; returns immediately.
    source_path: if given (e.g. OneDrive copy), is used as the fast-copy source.
    """
    global _refresh_thread, _refresh_error
    _refresh_done.clear()
    _refresh_error = ""

    def _run():
        global _refresh_error
        try:
            refresh_excel(excel_path, source_path=source_path)
        except Exception as exc:
            _refresh_error = str(exc)
        finally:
            _refresh_done.set()

    _refresh_thread = threading.Thread(target=_run, daemon=True, name="sp-refresh")
    _refresh_thread.start()


def _get_excel_proc_handle(xl) -> object | None:
    """
    Return a Win32 process handle for the Excel instance, or None on failure.
    Uses xl.ProcessID (COM property) — reliable even for windowless instances.
    Falls back to the Hwnd approach for older Excel versions.
    """
    try:
        import win32api
        import win32con
        # xl.ProcessID is available on Excel 2007+ and works for hidden instances
        try:
            pid = xl.ProcessID
        except Exception:
            import win32process
            hwnd = xl.Hwnd
            if not hwnd:
                return None
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
        return win32api.OpenProcess(
            win32con.SYNCHRONIZE | win32con.PROCESS_TERMINATE, False, pid
        )
    except Exception:
        return None


def _kill_excel_process(proc_handle) -> None:
    """
    Wait up to 15 s for the Excel process to exit cleanly after Quit().
    If it is still alive, terminate it forcefully and close the handle.
    """
    if proc_handle is None:
        time.sleep(5)   # fallback: blind wait (hidden Excel, no handle)
        return
    try:
        import win32api
        import win32event
        import win32con
        result = win32event.WaitForSingleObject(proc_handle, 15_000)
        if result != win32con.WAIT_OBJECT_0:   # did not exit in time
            win32api.TerminateProcess(proc_handle, 0)
    except Exception:
        pass
    finally:
        try:
            import win32api as _w32
            _w32.CloseHandle(proc_handle)
        except Exception:
            pass


def refresh_excel(excel_path: str | Path, source_path: str | Path | None = None) -> None:
    """
    Update the local SharePoint summary file.

    Strategy (in priority order):
      1. If *source_path* exists (e.g. an OneDrive sync copy), just copy it → *excel_path*.
         This is instant and never touches Excel at all.
      2. Otherwise open *excel_path* in a fresh hidden Excel instance, run RefreshAll,
         save, then quit.  The instance is always separate from any open Excel the user has.

    Raises on failure so the caller can report the error.
    """
    dst = Path(excel_path).resolve()

    # ── Strategy 1: copy from OneDrive ───────────────────────────────────
    if source_path:
        src = Path(source_path).resolve()
        if src.exists():
            shutil.copy2(src, dst)
            return

    # ── Strategy 2: COM refresh ───────────────────────────────────────────
    if not dst.exists():
        raise FileNotFoundError(f"SharePoint file not found: {dst}")

    try:
        import pythoncom
        import win32com.client as win32
    except ImportError:
        raise RuntimeError(
            "pywin32 is not installed — cannot refresh via Excel COM.\n"
            "Install it with: pip install pywin32"
        )

    tmp = Path(tempfile.gettempdir()) / f"_sp_refresh_{dst.name}"
    shutil.copy2(dst, tmp)
    tmp_path = str(tmp)

    pythoncom.CoInitialize()
    try:
        # Prefer a running Excel instance (fast — no startup cost).
        # Fall back to a new hidden instance if none is open.
        try:
            xl = win32.GetActiveObject("Excel.Application")
            we_started_excel = False
            proc_handle = None
        except Exception:
            xl = win32.DispatchEx("Excel.Application")
            we_started_excel = True
            proc_handle = _get_excel_proc_handle(xl)
            xl.Visible = False   # hide only the instance WE created

        # Freeze all visual updates so the user sees nothing
        prev_screen  = xl.ScreenUpdating
        prev_alerts  = xl.DisplayAlerts
        prev_events  = xl.EnableEvents
        xl.ScreenUpdating = False
        xl.DisplayAlerts  = False
        xl.EnableEvents   = False
        xl.Interactive    = False
        try:
            wb = xl.Workbooks.Open(tmp_path, UpdateLinks=0, ReadOnly=False)
            wb.RefreshAll()
            xl.CalculateUntilAsyncQueriesDone()
            wb.Save()
            wb.Close(SaveChanges=False)
        finally:
            # Always restore the shared instance's settings
            xl.ScreenUpdating = prev_screen
            xl.DisplayAlerts  = prev_alerts
            xl.EnableEvents   = prev_events
            xl.Interactive    = True
            if we_started_excel:
                xl.Quit()
                _kill_excel_process(proc_handle)

        # Copy refreshed temp back to the project file
        shutil.copy2(tmp, dst)
    finally:
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass
        try:
            tmp.unlink(missing_ok=True)
        except Exception:
            pass


def read_deals(excel_path: str | Path) -> list[dict]:
    """
    Read all rows from the SharePoint summary sheet.
    Returns a list of raw dicts keyed by column header.
    """
    src = Path(excel_path)
    if not src.exists():
        raise FileNotFoundError(f"SharePoint summary not found: {src}")

    # Copy to temp to avoid OneDrive / Excel file-lock
    tmp = Path(tempfile.gettempdir()) / f"sp_summary_{src.stem}.xlsx"
    shutil.copy2(src, tmp)

    try:
        wb = openpyxl.load_workbook(tmp, data_only=True, read_only=True)
        try:
            ws = wb[SHEET_NAME]
        except KeyError:
            raise ValueError(
                f"Sheet '{SHEET_NAME}' not found. Available: {wb.sheetnames}"
            )
        rows = list(ws.iter_rows(values_only=True))
        wb.close()
    finally:
        try:
            tmp.unlink(missing_ok=True)
        except Exception:
            pass

    if not rows:
        return []

    # Strip surrounding quotes that Power Query sometimes adds to headers
    headers = [
        str(h).strip('"').strip() if h is not None else "" for h in rows[0]
    ]

    return [
        {headers[i]: (row[i] if i < len(row) else None) for i in range(len(headers))}
        for row in rows[1:]
    ]
