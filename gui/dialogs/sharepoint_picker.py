# gui/dialogs/sharepoint_picker.py
"""
Modal dialog that loads deals from the SharePoint summary Excel and
lets the user select one to auto-fill the PC Mail form.
"""
from __future__ import annotations

import threading
import tkinter as tk
from tkinter import ttk, messagebox


class SharePointPickerDialog:
    """
    Opens a Toplevel, loads deals from the SharePoint Excel in a background
    thread, shows them in a Treeview, and returns the selected parsed deal
    in ``self.result`` after the window closes.
    """

    def __init__(self, parent, excel_path: str):
        self.result: dict | None = None
        self._excel_path = excel_path
        self._deals: list[dict] = []
        self._parsed: dict | None = None

        self.top = tk.Toplevel(parent)
        self.top.title("Load from SharePoint — New Notes Summary")
        self.top.geometry("960x520")
        self.top.resizable(True, True)
        self.top.transient(parent)
        self.top.grab_set()

        self._build_ui()
        self._fetch()

        self.top.wait_window()

    # ── UI ────────────────────────────────────────────────────────────────
    def _build_ui(self):
        # Status
        self._status_var = tk.StringVar(value="Loading deals…")
        ttk.Label(self.top, textvariable=self._status_var,
                  foreground="#555").pack(anchor="w", padx=12, pady=(8, 2))

        # Treeview
        tree_frame = ttk.Frame(self.top)
        tree_frame.pack(fill="both", expand=True, padx=12, pady=4)

        cols = ("title", "status", "series", "issue_date")
        self._tree = ttk.Treeview(
            tree_frame, columns=cols, show="headings",
            selectmode="browse", height=14,
        )
        col_cfg = {
            "title":      ("Title",      450),
            "status":     ("Status",     190),
            "series":     ("Series",      65),
            "issue_date": ("Issue Date",  95),
        }
        for col, (heading, width) in col_cfg.items():
            self._tree.heading(col, text=heading)
            self._tree.column(col, width=width, anchor="w")

        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self._tree.yview)
        self._tree.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        self._tree.pack(fill="both", expand=True)
        self._tree.bind("<<TreeviewSelect>>", self._on_select)
        self._tree.bind("<Double-1>", lambda _: self._on_load())

        # Preview
        preview_frame = ttk.LabelFrame(self.top, text="Parsed fields preview")
        preview_frame.pack(fill="x", padx=12, pady=4)
        self._preview_var = tk.StringVar(value="—")
        ttk.Label(
            preview_frame, textvariable=self._preview_var,
            wraplength=920, justify="left",
        ).pack(padx=8, pady=5)

        # Buttons
        btn_frame = ttk.Frame(self.top)
        btn_frame.pack(fill="x", padx=12, pady=8)
        ttk.Button(btn_frame, text="Load into form",
                   command=self._on_load).pack(side="left", padx=4)
        ttk.Button(btn_frame, text="Cancel",
                   command=self.top.destroy).pack(side="left", padx=4)

    # ── Data fetching ─────────────────────────────────────────────────────
    def _fetch(self):
        def _run():
            try:
                import os
                from sharepoint.reader import read_deals
                from sharepoint.parser import parse_deal

                self.top.after(0, lambda: self._status_var.set("Reading deals…"))
                rows   = read_deals(self._excel_path)
                parsed = [parse_deal(r) for r in rows]

                # Show file modification time so the user can verify freshness
                try:
                    mtime = os.path.getmtime(self._excel_path)
                    from datetime import datetime
                    ts = datetime.fromtimestamp(mtime).strftime("%d/%m/%Y %H:%M")
                except Exception:
                    ts = "?"

                self.top.after(0, lambda: self._populate(parsed, ts))
            except Exception as exc:
                msg = str(exc)
                self.top.after(0, lambda: self._on_error(msg))

        threading.Thread(target=_run, daemon=True).start()

    def _populate(self, deals: list[dict], file_ts: str = ""):
        self._deals = deals
        for d in deals:
            self._tree.insert("", "end", values=(
                d["title"],
                d["status"],
                d["series"],
                d["issue_date"],
            ))
        count = len(deals)
        ts_part = f"  •  file updated {file_ts}" if file_ts else ""
        self._status_var.set(
            f"{count} deal(s) loaded — double-click or select + Load{ts_part}"
        )

    # ── Selection ─────────────────────────────────────────────────────────
    def _on_select(self, _=None):
        sel = self._tree.selection()
        if not sel:
            return
        idx  = self._tree.index(sel[0])
        d    = self._deals[idx]
        self._parsed = d

        params = " | ".join(
            f"p{i+1}={d[f'param{i+1}']}"
            for i in range(4) if d.get(f"param{i+1}")
        ) or "—"
        uls  = ", ".join(d["underlyings"]) or "—"
        hp   = d["hedge_party"] or "—"
        up   = (f"@{d['upfront']}%" if d["upfront"] else "—")
        self._preview_var.set(
            f"Product: {d['product'] or '?'}  |  Maturity: {d['maturity'] or '?'}  |  "
            f"Series: {d['series']}  |  ISIN: {d['isin'] or '—'}  |  "
            f"Client: {d['client'] or '?'}  |  Currency: {d['currency'] or '?'}  |  "
            f"Underlyings: {uls}  |  Trade: {d['trade_date']}  |  Issue: {d['issue_date']}\n"
            f"Params: {params}  |  Hedge: {hp}  {up}  BtB: {d['btb_amount'] or '—'}"
        )

    def _on_load(self):
        if not self._parsed:
            messagebox.showwarning("No selection", "Please select a deal first.",
                                   parent=self.top)
            return
        self.result = self._parsed
        self.top.destroy()

    def _on_error(self, msg: str):
        self._status_var.set("Failed to load — see error")
        messagebox.showerror(
            "SharePoint Error",
            f"Could not read the SharePoint Excel file:\n\n{msg}",
            parent=self.top,
        )
