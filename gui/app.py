import sys
import tkinter as tk
from pathlib import Path
from tkinter import ttk

from gui.tabs.documentatie_tab import DocumentatieMailTab
from gui.tabs.marketing_tab import MarketingMailTab
from gui.tabs.pcmail_tab import PCMailTab
from gui.tabs.increase_decrease_tab import IncreaseDecreaseTab


def _icon_path() -> Path:
    """Bepaal het pad naar assets/app.ico voor normale en frozen (PyInstaller)-uitvoeringen."""
    base = Path(sys._MEIPASS) if getattr(sys, "frozen", False) else Path(__file__).resolve().parent.parent
    return base / "assets" / "app.ico"


def _set_taskbar_icon() -> None:
    """Laat Windows het EXE-icoon gebruiken voor de taakbalkknop (niet het Python-icoon)."""
    try:
        import ctypes
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("VLK.MoamProject.1")
    except Exception:
        pass


def _apply_icon(window: tk.Wm) -> None:
    """Zet het programma-icoon op elk Tk- of Toplevel-venster — probeert eerst iconphoto en valt anders terug op iconbitmap."""
    ico = _icon_path()
    if not ico.exists():
        return
    try:
        from PIL import Image, ImageTk
        img = Image.open(str(ico))
        photo = ImageTk.PhotoImage(img)
        window.iconphoto(True, photo)
        window._icon_ref = photo  # voorkom dat garbage collection hem opruimt
    except Exception:
        try:
            window.iconbitmap(str(ico))
        except Exception:
            pass


class MailGeneratorApp(tk.Tk):
    """
    Hoofdvenster van de app.
    Bij het opstarten verbergt die zichzelf, toont een Toplevel-splash terwijl SharePoint
    Excel op de achtergrond ververst, bouwt daarna de volledige UI en laat die zien.
    Tijdens de hele levensduur van de app wordt maar ÉÉN tk.Tk()-exemplaar aangemaakt.
    """

    def __init__(self):
        super().__init__()
        _set_taskbar_icon()
        _apply_icon(self)
        self.withdraw()
        self.products: list[dict] = []
        self.max_products = 41
        # Start het voorverwarmen van zware imports meteen op de achtergrond
        from gui.prewarmer import start as prewarm_start
        prewarm_start()
        self._show_splash()

    # ── Opstartscherm ─────────────────────────────────────────────────────
    def _show_splash(self):
        self._splash = tk.Toplevel(self)
        self._splash.title("Mail Generator")
        self._splash.resizable(False, False)
        self._splash.protocol("WM_DELETE_WINDOW", lambda: None)
        _apply_icon(self._splash)

        frame = ttk.Frame(self._splash, padding=30)
        frame.pack()
        ttk.Label(frame, text="📊  Mail Generator",
                  font=("Segoe UI", 14, "bold")).pack(pady=(0, 10))

        ttk.Label(frame, text="Load latest SharePoint data?",
                  font=("Segoe UI", 10)).pack(pady=(0, 12))

        btn_frame = ttk.Frame(frame)
        btn_frame.pack()
        ttk.Button(btn_frame, text="✅  Yes, refresh data",
                   command=self._start_refresh, width=22).pack(side="left", padx=(0, 8))
        ttk.Button(btn_frame, text="⏩  Skip, open now",
                   command=self._skip_refresh, width=22).pack(side="left")

        self._bar = ttk.Progressbar(frame, mode="indeterminate", length=280)
        self._bar.pack(pady=(14, 0))
        self._status_var = tk.StringVar(value=" ")
        ttk.Label(frame, textvariable=self._status_var,
                  foreground="#555", font=("Segoe UI", 8)).pack(pady=(4, 0))

        # Start de gedeelde verborgen Excel alvast terwijl de gebruiker het opstartscherm leest.
        # Draait op de hoofddraad (COM-safe), zodat de eerste mail meteen klaarstaat.
        self.after(50, self._warm_excel)

        # Centreer het opstartscherm op het scherm
        self._splash.update_idletasks()
        w = self._splash.winfo_reqwidth()
        h = self._splash.winfo_reqheight()
        sw = self._splash.winfo_screenwidth()
        sh = self._splash.winfo_screenheight()
        self._splash.geometry(f"+{(sw - w) // 2}+{(sh - h) // 2}")

    # ── Excel alvast opwarmen ────────────────────────────────────────────
    def _warm_excel(self):
        """Start de blijvende verborgen Excel-instantie terwijl het opstartscherm zichtbaar is."""
        try:
            from MarketingMail.excel_handler import _get_shared_xl
            _get_shared_xl()
        except Exception:
            pass

    # ── Verversen ────────────────────────────────────────────────────────
    def _start_refresh(self):
        """Gebruiker koos om te verversen — start op de achtergrond, toon voortgang en wacht in het opstartscherm."""
        self._status_var.set("Refreshing SharePoint data, please wait…")
        self._bar.start(12)
        try:
            from statics.data import SHAREPOINT_SUMMARY_PATH, SHAREPOINT_ONEDRIVE_PATH
            from sharepoint.reader import start_background_refresh
            start_background_refresh(SHAREPOINT_SUMMARY_PATH,
                                     source_path=SHAREPOINT_ONEDRIVE_PATH)
        except Exception as exc:
            self._status_var.set(f"Refresh error: {exc}")
            self._bar.stop()
            return
        self.after(500, self._poll)

    def _skip_refresh(self):
        """Gebruiker koos om over te slaan — markeer het verversen als klaar en open meteen."""
        from sharepoint.reader import _refresh_done
        _refresh_done.set()
        self._open_main()

    def _poll(self):
        if getattr(self, "_ui_built", False):
            return
        from sharepoint.reader import _refresh_done, _refresh_error
        if _refresh_done.is_set():
            if _refresh_error:
                self._status_var.set(f"⚠️ {_refresh_error} — opening with cached data")
                self._bar.stop()
                self.after(2000, self._open_main)
            else:
                self._open_main()
        else:
            self.after(500, self._poll)

    def _open_main(self):
        if getattr(self, "_ui_built", False):
            return
        self._ui_built = True
        self._bar.stop()
        self._splash.destroy()
        self._build_ui()
        self.deiconify()

    # ── Hoofd-UI ──────────────────────────────────────────────────────────
    def _build_ui(self):
        self.title("Mail Generator - Unified Interface")
        self.geometry("700x800")
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        _apply_icon(self)

        # ── Globale werkbalk ──────────────────────────────────────────────
        toolbar = ttk.Frame(self, relief="flat")
        toolbar.pack(fill="x", padx=8, pady=(6, 0))

        ttk.Button(
            toolbar, text="🔄  Refresh SharePoint",
            command=self._manual_refresh,
        ).pack(side="left")

        ttk.Button(
            toolbar, text="🐛  Debug",
            command=self._show_debug,
        ).pack(side="left", padx=(8, 0))

        self._refresh_status = tk.StringVar(value="")
        ttk.Label(
            toolbar, textvariable=self._refresh_status,
            foreground="#555", font=("Segoe UI", 8),
        ).pack(side="left", padx=(10, 0))

        ttk.Separator(self, orient="horizontal").pack(fill="x", padx=8, pady=(4, 0))

        notebook = ttk.Notebook(self)
        notebook.pack(fill="both", expand=True, padx=5, pady=5)

        doc_tab       = DocumentatieMailTab(notebook)
        marketing_tab = MarketingMailTab(notebook)
        pc_tab        = PCMailTab(notebook)
        id_tab        = IncreaseDecreaseTab(notebook)

        notebook.add(doc_tab.frame,       text="Documentatie Mail")
        notebook.add(marketing_tab.frame, text="Marketing Mail")
        notebook.add(pc_tab.frame,        text="PC Mail")
        notebook.add(id_tab.frame,        text="Increase / Decrease")

    # ── SharePoint handmatig verversen ──────────────────────────────────
    def _manual_refresh(self):
        """Getriggerd door de werkbalkknop — ververs op de achtergrond en toon de status."""
        self._refresh_status.set("⏳ Refreshing…")
        self.update_idletasks()
        try:
            from statics.data import SHAREPOINT_SUMMARY_PATH, SHAREPOINT_ONEDRIVE_PATH
            from sharepoint.reader import start_background_refresh, _refresh_done
            _refresh_done.clear()
            start_background_refresh(SHAREPOINT_SUMMARY_PATH,
                                     source_path=SHAREPOINT_ONEDRIVE_PATH)
        except Exception as exc:
            self._refresh_status.set(f"❌ {exc}")
            return
        self._poll_refresh()

    def _poll_refresh(self):
        from sharepoint.reader import _refresh_done, _refresh_error
        if _refresh_done.is_set():
            if _refresh_error:
                self._refresh_status.set(f"❌ {_refresh_error}")
            else:
                self._refresh_status.set("✅ SharePoint data refreshed")
                self.after(4000, lambda: self._refresh_status.set(""))
        else:
            self.after(500, self._poll_refresh)

    @staticmethod
    def _on_close():
        import sys
        sys.exit(0)

    # ── Debug dialoog ─────────────────────────────────────────────────────
    def _show_debug(self):
        """Draai volledige diagnostiek en toon die in een scrollbaar dialoogvenster."""
        from statics.mail_debug import full_diagnostics
        import threading

        win = tk.Toplevel(self)
        win.title("Mail Generator - Debug Report")
        win.geometry("750x600")
        _apply_icon(win)

        text = tk.Text(win, wrap="word", font=("Consolas", 9))
        vsb = ttk.Scrollbar(win, orient="vertical", command=text.yview)
        text.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        text.pack(fill="both", expand=True)

        text.insert("end", "Running diagnostics...\n")
        text.config(state="disabled")

        def _run():
            report = full_diagnostics()
            def _update():
                text.config(state="normal")
                text.delete("1.0", "end")
                text.insert("end", report)
                text.config(state="disabled")
            win.after(0, _update)

        threading.Thread(target=_run, daemon=True).start()


def main():
    import sys
    app = MailGeneratorApp()
    app.mainloop()
    sys.exit(0)   # zorg voor een nette afsluiting, ook als mainloop terugkomt zonder destroy


if __name__ == "__main__":
    main()