# gui/tabs/documentatie_tab.py
"""
Documentatie Mail tab - completely independent.
Builds email_data directly without CommonMailInputs middleman.
"""

import tkinter as tk
from tkinter import ttk, messagebox
from datetime import date

from statics.data import (
    ISSUERS, CURRENCIES, PRODUCT_TYPES, MATURITIES,
    UNDERLYINGS, ADVISER_HELPERS, SHAREPOINT_SUMMARY_PATH,
)
from gui.widgets import DynamicListManager, UnderlyingSearchEntry

class DocumentatieMailTab:
    """Documentatie Mail tab - completely independent"""

    def __init__(self, parent):
        self.frame = ttk.Frame(parent)

        # ── SharePoint toolbar ────────────────────────────────────────────
        toolbar = ttk.Frame(self.frame)
        toolbar.pack(fill="x", padx=10, pady=(8, 0))
        ttk.Button(
            toolbar, text="📋 Load from SharePoint",
            command=self._open_sharepoint_picker,
        ).pack(side="left", padx=(0, 8))
        ttk.Label(
            toolbar,
            text="Auto-fill from the SharePoint deal list",
            foreground="#888", font=("Segoe UI", 8),
        ).pack(side="left")

        # ── Product Information ──────────────────────────────────────────
        info_frame = ttk.LabelFrame(self.frame, text="Product Information", width=400)
        info_frame.pack(fill="x", padx=10, pady=10)
        info_frame.pack_propagate(False)

        self.vars = {}
        row = 0
        for label, key, values in [
            ("Issuer",       "issuer",       ISSUERS),
            ("Currency",     "currency",     CURRENCIES),
            ("Product Type", "product_type", list(PRODUCT_TYPES.keys())),
            ("Maturity",     "maturity",     MATURITIES),
        ]:
            ttk.Label(info_frame, text=label + ":").grid(
                row=row, column=0, sticky="w", padx=5, pady=4
            )
            var = tk.StringVar()
            ttk.Combobox(
                info_frame, textvariable=var, values=values,
                state="readonly", width=30
            ).grid(row=row, column=1, padx=5, pady=4)
            self.vars[key] = var
            row += 1

        # ── Underlyings (same pattern as PCMailTab) ──────────────────────
        ttk.Label(info_frame, text="Underlying(s):").grid(
            row=row, column=0, sticky="w", padx=5, pady=4
        )

        self.selected_underlyings: list = []

        ul_pick_row = ttk.Frame(info_frame)
        ul_pick_row.grid(row=row, column=1, padx=5, pady=4, sticky="w")

        self._ul_search = UnderlyingSearchEntry(
            ul_pick_row,
            all_tickers=UNDERLYINGS,
            on_select=self._add_underlying,
        )
        self._ul_search.pack(side="left", padx=(0, 2))

        ttk.Button(ul_pick_row, text="✕", width=2,
                   command=self._remove_last_underlying).pack(side="left", padx=2)

        self._ul_display_var = tk.StringVar(value="—")
        ttk.Label(ul_pick_row, textvariable=self._ul_display_var,
                  foreground="#444", width=22).pack(side="left", padx=(6, 0))

        row += 1

        ttk.Label(info_frame,
                  text="Type ticker → select from list or press Enter.  ✕ removes last.",
                  foreground="#888", font=("Segoe UI", 8)).grid(
            row=row, column=1, sticky="w", padx=5
        )
        row += 1

        # ── ISIN ─────────────────────────────────────────────────────────
        ttk.Label(info_frame, text="ISIN:").grid(
            row=row, column=0, sticky="w", padx=5, pady=4
        )
        self.vars["isin"] = tk.StringVar()
        ttk.Entry(info_frame, textvariable=self.vars["isin"], width=32).grid(
            row=row, column=1, padx=5, pady=4
        )
        row += 1

        # ── VLK code ──────────────────────────────────────────────────────
        ttk.Label(info_frame, text="VLK code:").grid(
            row=row, column=0, sticky="w", padx=5, pady=4
        )
        self.vars["vlk_code"] = tk.StringVar()
        ttk.Entry(info_frame, textvariable=self.vars["vlk_code"], width=32).grid(
            row=row, column=1, padx=5, pady=4
        )

        # ── Advisers & Amounts ───────────────────────────────────────────
        self.adviser_manager = DynamicListManager(
            self.frame,
            "Advisers & Amounts",
            columns=("Adviser Name", "Amount", "Price %"),
            search_options={"Adviser Name": list(ADVISER_HELPERS.keys())},
            defaults={"Price %": "100"},
        )
        self.adviser_manager.pack(fill="both", expand=False, padx=10, pady=10)

        # ── Buttons ──────────────────────────────────────────────────────
        button_frame = ttk.Frame(self.frame)
        button_frame.pack(fill="x", padx=10, pady=10)
        ttk.Button(button_frame, text="Send", command=self._send).pack(
            side="left", padx=5
        )

    # ── Underlying helpers (identical pattern to PCMailTab) ──────────────
    def _add_underlying(self, ticker: str):
        ticker = ticker.strip().upper()
        if ticker and ticker not in self.selected_underlyings:
            self.selected_underlyings.append(ticker)
            self._refresh_ul_display()

    def _remove_last_underlying(self):
        if self.selected_underlyings:
            self.selected_underlyings.pop()
            self._refresh_ul_display()

    def _refresh_ul_display(self):
        if not self.selected_underlyings:
            self._ul_display_var.set("—")
        else:
            n      = len(self.selected_underlyings)
            joined = " / ".join(self.selected_underlyings)
            if len(joined) > 20:
                joined = joined[:18] + "…"
            self._ul_display_var.set(f"{joined}  ({n})")

    # ── SharePoint loader ─────────────────────────────────────────────────
    def _open_sharepoint_picker(self):
        from gui.dialogs.sharepoint_picker import SharePointPickerDialog
        dlg = SharePointPickerDialog(self.frame.winfo_toplevel(),
                                     str(SHAREPOINT_SUMMARY_PATH))
        if dlg.result:
            self._load_from_sharepoint(dlg.result)

    def _load_from_sharepoint(self, deal: dict):
        """Populate form fields from a parsed SharePoint deal."""
        if deal.get("issuer") and deal["issuer"] in ISSUERS:
            self.vars["issuer"].set(deal["issuer"])
        if deal.get("product") and deal["product"] in list(PRODUCT_TYPES.keys()):
            self.vars["product_type"].set(deal["product"])
        if deal.get("currency") and deal["currency"] in CURRENCIES:
            self.vars["currency"].set(deal["currency"])
        if deal.get("maturity") and deal["maturity"] in MATURITIES:
            self.vars["maturity"].set(deal["maturity"])
        if deal.get("isin"):
            self.vars["isin"].set(deal["isin"])
        if deal.get("vl_code"):
            self.vars["vlk_code"].set(deal["vl_code"])

        # Underlyings
        self.selected_underlyings.clear()
        for ul in deal.get("underlyings", []):
            ul = ul.strip().upper()
            if ul:
                self.selected_underlyings.append(ul)
        self._refresh_ul_display()

        # Advisers from parsed comments
        advisers = deal.get("advisers", [])
        if advisers:
            self.adviser_manager.clear()
            for adv in advisers:
                self.adviser_manager.add_item({
                    "Adviser Name": adv["name"],
                    "Amount":       adv["amount"],
                    "Price %":      adv["price"],
                })

    # ── Example loader ───────────────────────────────────────────────────
    # ── Send ─────────────────────────────────────────────────────────────
    def _send(self):
        from Documentatie_Mail.service import send_documentatie_mail

        if not self.vars["isin"].get():
            messagebox.showerror("Error", "ISIN is required")
            return
        if not self.adviser_manager.get_items():
            messagebox.showerror("Error", "Add at least one adviser")
            return

        trades = [
            {
                "adviser": a["Adviser Name"],
                "amount":  a["Amount"],
                "price":   a.get("Price %", "").strip() or "100",
            }
            for a in self.adviser_manager.get_items()
        ]

        email_data = {
            "product":     self.vars["product_type"].get(),
            "isin":        self.vars["isin"].get(),
            "maturity":    self.vars["maturity"].get(),
            "currency":    self.vars["currency"].get(),
            "issuer":      self.vars["issuer"].get(),
            "underlyings": list(self.selected_underlyings),   # ← list, matches email_builder
            "trades":      trades,
            "to": [t["adviser"] for t in trades],  # ← advisers are the recipients
            "cc": [],
            "vlk_code":    self.vars["vlk_code"].get(),
            "today":       date.today(),
        }

        try:
            send_documentatie_mail(email_data)
            messagebox.showinfo("Sent", "Email opened in Outlook ✅")
        except Exception as e:
            messagebox.showerror("Error", str(e))
            raise