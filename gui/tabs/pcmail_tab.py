import tkinter as tk
from tkinter import ttk, messagebox
from datetime import date, timedelta
from statics.data import (
    ISSUERS, CURRENCIES, CLIENTS, PRODUCT_TYPES, MATURITIES,
    HEDGEPARTY, UNDERLYINGS, PRODUCT_PAYOFF_FIELDS, PRODUCTS_NO_UNDERLYING,
    SHAREPOINT_SUMMARY_PATH,
)
from gui.widgets import UnderlyingSearchEntry


def _add_business_days(start: date, n: int) -> date:
    d, added = start, 0
    while added < n:
        d += timedelta(days=1)
        if d.weekday() < 5:
            added += 1
    return d


# ---------------------------------------------------------------------------
# PCMailTab
# ---------------------------------------------------------------------------

class PCMailTab:
    """PC Mail tab – complete, GUI-friendly."""

    def __init__(self, parent):
        self.frame = ttk.Frame(parent)

        self._asianing_tail_var = tk.StringVar()
        self._asianing_obs_var  = tk.StringVar()

        canvas   = tk.Canvas(self.frame, highlightthickness=0)
        scrollbar = ttk.Scrollbar(self.frame, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)

        inner  = ttk.Frame(canvas)
        win_id = canvas.create_window((0, 0), window=inner, anchor="nw")

        canvas.bind("<Configure>",
                    lambda e: canvas.itemconfig(win_id, width=e.width))
        inner.bind("<Configure>",
                   lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind_all("<MouseWheel>",
                        lambda e: canvas.yview_scroll(int(-1 * (e.delta / 120)), "units"))

        self._build_ui(inner)

    # -----------------------------------------------------------------------
    def _build_ui(self, parent):

        # ── SharePoint toolbar ────────────────────────────────────────────
        sp_frame = ttk.Frame(parent)
        sp_frame.pack(fill="x", padx=10, pady=(8, 2))
        ttk.Button(
            sp_frame, text="📋  Load from SharePoint",
            command=self._open_sharepoint_picker,
        ).pack(side="left")
        ttk.Label(
            sp_frame,
            text="Auto-fills product, series, dates, underlyings, client and hedge info",
            foreground="#888", font=("Segoe UI", 8),
        ).pack(side="left", padx=10)

        # ── Section 1 : Product Identity ──────────────────────────────────
        id_frame = ttk.LabelFrame(parent, text="Product Identity")
        id_frame.pack(fill="x", padx=10, pady=(4, 4))

        self.vars = {}

        # Row 0: Issuer | Currency
        ttk.Label(id_frame, text="Issuer:").grid(row=0, column=0, sticky="w", padx=8, pady=3)
        v_issuer = tk.StringVar()
        ttk.Combobox(id_frame, textvariable=v_issuer, values=ISSUERS,
                     state="readonly", width=16).grid(row=0, column=1, padx=8, pady=3, sticky="w")
        self.vars["issuer"] = v_issuer

        ttk.Label(id_frame, text="Currency:").grid(row=0, column=2, sticky="w", padx=(16, 8), pady=3)
        v_currency = tk.StringVar()
        ttk.Combobox(id_frame, textvariable=v_currency, values=CURRENCIES,
                     state="readonly", width=16).grid(row=0, column=3, padx=8, pady=3, sticky="w")
        self.vars["currency"] = v_currency

        # Row 1: Client | Product
        ttk.Label(id_frame, text="Client:").grid(row=1, column=0, sticky="w", padx=8, pady=3)
        v_client = tk.StringVar()
        _client_cb = ttk.Combobox(id_frame, textvariable=v_client, values=CLIENTS,
                                   state="readonly", width=16)
        _client_cb.grid(row=1, column=1, padx=8, pady=3, sticky="w")
        _client_cb.bind("<<ComboboxSelected>>", self._on_client_change)
        self.vars["client"] = v_client

        ttk.Label(id_frame, text="Product:").grid(row=1, column=2, sticky="w", padx=(16, 8), pady=3)
        v_product = tk.StringVar()
        ttk.Combobox(id_frame, textvariable=v_product,
                     values=list(PRODUCT_TYPES.keys()),
                     state="readonly", width=24).grid(row=1, column=3, padx=8, pady=3, sticky="w")
        self.vars["product"] = v_product

        # Row 2: Series | Maturity
        ttk.Label(id_frame, text="Series:").grid(row=2, column=0, sticky="w", padx=8, pady=3)
        v_series = tk.StringVar()
        ttk.Entry(id_frame, textvariable=v_series, width=18).grid(
            row=2, column=1, padx=8, pady=3, sticky="w")
        self.vars["series"] = v_series

        ttk.Label(id_frame, text="Maturity:").grid(row=2, column=2, sticky="w", padx=(16, 8), pady=3)
        v_maturity = tk.StringVar()
        self._maturity_combo = ttk.Combobox(
            id_frame, textvariable=v_maturity,
            values=MATURITIES,
            state="readonly", width=24)
        self._maturity_combo.grid(row=2, column=3, padx=8, pady=3, sticky="w")
        self._maturity_combo.bind("<<ComboboxSelected>>", self._on_maturity_change)
        self.vars["maturity"] = v_maturity

        # ── Section 2 : Commercial ─────────────────────────────────────────
        com_frame = ttk.LabelFrame(parent, text="Commercial")
        com_frame.pack(fill="x", padx=10, pady=4)

        # Row 0: Issue Size | Sold
        ttk.Label(com_frame, text="Issue Size:").grid(row=0, column=0, sticky="w", padx=8, pady=3)
        v_issue = tk.StringVar()
        ttk.Entry(com_frame, textvariable=v_issue, width=18).grid(
            row=0, column=1, padx=8, pady=3, sticky="w")
        self.vars["issue_size"] = v_issue

        ttk.Label(com_frame, text="Sold:").grid(row=0, column=2, sticky="w", padx=(16, 8), pady=3)
        v_sold = tk.StringVar()
        ttk.Entry(com_frame, textvariable=v_sold, width=18).grid(
            row=0, column=3, padx=8, pady=3, sticky="w")
        self.vars["sold"] = v_sold

        # Row 1: hint text for Issue Size / Sold format
        ttk.Label(com_frame, text="(in k or mio — e.g. 500k, 1.25 mio)",
                  foreground="#888", font=("Segoe UI", 8)).grid(
            row=1, column=1, columnspan=3, sticky="w", padx=8, pady=(0, 2))

        # Row 2: Struct Fee | Dist Fee
        ttk.Label(com_frame, text="Struct Fee (%):").grid(row=2, column=0, sticky="w", padx=8, pady=3)
        v_struct = tk.StringVar(value="1.5")
        ttk.Entry(com_frame, textvariable=v_struct, width=10).grid(
            row=2, column=1, padx=8, pady=3, sticky="w")
        self.vars["struct_fee"] = v_struct

        ttk.Label(com_frame, text="Dist Fee (%):").grid(row=2, column=2, sticky="w", padx=(16, 8), pady=3)
        v_dist = tk.StringVar(value="0.0")
        ttk.Entry(com_frame, textvariable=v_dist, width=10).grid(
            row=2, column=3, padx=8, pady=3, sticky="w")
        self.vars["dist_fee"] = v_dist

        # ── Section 3 : Options (VLK code only) ───────────────────────────
        opt_frame = ttk.LabelFrame(parent, text="Options")
        opt_frame.pack(fill="x", padx=10, pady=4)

        self._vlk_code_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            opt_frame,
            text="VLK Code Required",
            variable=self._vlk_code_var,
        ).grid(row=0, column=0, sticky="w", padx=8, pady=4)

        # ── Section 4 : Hedging ───────────────────────────────────────────
        hedge_frame = ttk.LabelFrame(parent, text="Hedging")
        hedge_frame.pack(fill="x", padx=10, pady=4)

        self._hedge_type = tk.StringVar(value="Own Book")

        radio_row = ttk.Frame(hedge_frame)
        radio_row.grid(row=0, column=0, columnspan=4, sticky="w", padx=8, pady=4)

        ttk.Radiobutton(radio_row, text="Own Book", variable=self._hedge_type,
                        value="Own Book", command=self._on_hedge_change).pack(side="left", padx=(0, 12))
        ttk.Radiobutton(radio_row, text="BTB", variable=self._hedge_type,
                        value="BTB", command=self._on_hedge_change).pack(side="left")

        self._btb_frame = ttk.Frame(hedge_frame)
        self._btb_frame.grid(row=1, column=0, columnspan=4, sticky="w", padx=8, pady=(0, 6))

        self._btb_vars = {}

        # Hedge Party — search widget backed by HEDGEPARTY list
        ttk.Label(self._btb_frame, text="Hedge Party:").grid(
            row=0, column=0, sticky="w", padx=(0, 6), pady=2)
        self._btb_hedge_search = UnderlyingSearchEntry(
            self._btb_frame,
            all_tickers=HEDGEPARTY,
            on_select=lambda v: self._btb_hedge_search.set(v),
            auto_clear=False,
        )
        self._btb_hedge_search.grid(row=0, column=1, sticky="w", pady=2)

        # Upfront and Amount — plain entries
        for i, (lbl, key) in enumerate([
            ("Upfront", "upfront"),
            ("Amount",  "amount"),
        ], start=1):
            ttk.Label(self._btb_frame, text=lbl + ":").grid(
                row=i, column=0, sticky="w", padx=(0, 6), pady=2)
            v = tk.StringVar()
            ttk.Entry(self._btb_frame, textvariable=v, width=18).grid(
                row=i, column=1, sticky="w", pady=2)
            self._btb_vars[key] = v

        self._btb_frame.grid_remove()

        # ── Section 5 : Payoff Parameters ─────────────────────────────────
        self._payoff_frame = ttk.LabelFrame(parent, text="Payoff Parameters")
        self._payoff_frame.pack(fill="x", padx=10, pady=4)

        for key in ("param1", "param2", "param3", "param4"):
            self.vars[key] = tk.StringVar(value="")

        self._payoff_widgets = []
        self.vars["product"].trace_add("write", self._on_product_change)

        # ── Section 6 : Underlyings ────────────────────────────────────────
        self._ul_frame = ttk.LabelFrame(parent, text="Underlyings")
        self._ul_frame.pack(fill="x", padx=10, pady=4)
        ul_frame = self._ul_frame

        self.selected_underlyings: list[str] = []

        pick_row = ttk.Frame(ul_frame)
        pick_row.pack(fill="x", padx=8, pady=4)

        ttk.Label(pick_row, text="Search:").pack(side="left")

        self._ul_search = UnderlyingSearchEntry(
            pick_row,
            all_tickers=UNDERLYINGS,
            on_select=self._add_underlying,
        )
        self._ul_search.pack(side="left", padx=(4, 2))

        ttk.Button(pick_row, text="✕", width=2,
                   command=self._remove_last_underlying).pack(side="left", padx=2)

        self._ul_display_var = tk.StringVar(value="—")
        ttk.Label(pick_row, textvariable=self._ul_display_var,
                  foreground="#444", width=30).pack(side="left", padx=(8, 0))

        ttk.Label(ul_frame,
                  text="Type ticker → select from list or press Enter for best match.  "
                       "✕ removes the last one.",
                  foreground="#888", font=("Segoe UI", 8)).pack(
            anchor="w", padx=8, pady=(0, 4))

        # ── Section 7 : Dates ─────────────────────────────────────────────
        date_frame = ttk.LabelFrame(parent, text="Dates")
        date_frame.pack(fill="x", padx=10, pady=4)

        trade = date.today()
        issue = _add_business_days(trade, 5)

        # Trade Date row
        ttk.Label(date_frame, text="Trade Date:").grid(
            row=0, column=0, sticky="w", padx=8, pady=3)
        self._trade_override = tk.BooleanVar(value=False)
        ttk.Checkbutton(date_frame, text="Override",
                        variable=self._trade_override,
                        command=self._toggle_trade_date).grid(
            row=0, column=1, sticky="w", padx=8)
        self._trade_date_var = tk.StringVar(value=trade.strftime("%d/%m/%Y"))
        self._trade_date_entry = ttk.Entry(
            date_frame, textvariable=self._trade_date_var,
            width=14, state="disabled")
        self._trade_date_entry.grid(row=0, column=2, sticky="w", padx=4, pady=3)
        ttk.Label(date_frame, text="(auto = today)",
                  foreground="#888").grid(row=0, column=3, sticky="w", padx=4)

        ttk.Label(date_frame, text="Issue Date:").grid(
            row=1, column=0, sticky="w", padx=8, pady=3)

        self._issue_override = tk.BooleanVar(value=False)
        ttk.Checkbutton(date_frame, text="Override",
                        variable=self._issue_override,
                        command=self._toggle_issue_date).grid(
            row=1, column=1, sticky="w", padx=8)

        self._issue_date_var = tk.StringVar(value=issue.strftime("%d/%m/%Y"))
        self._issue_date_entry = ttk.Entry(
            date_frame, textvariable=self._issue_date_var,
            width=14, state="disabled")
        self._issue_date_entry.grid(row=1, column=2, sticky="w", padx=4, pady=3)

        ttk.Label(date_frame, text="(auto = trade + 5 business days)",
                  foreground="#888").grid(row=1, column=3, sticky="w", padx=4)

        # ── Section 8 : Denomination override ─────────────────────────────
        den_frame = ttk.LabelFrame(parent, text="Denomination")
        den_frame.pack(fill="x", padx=10, pady=4)

        self._den_override = tk.BooleanVar(value=False)
        ttk.Checkbutton(den_frame, text="Override denomination",
                        variable=self._den_override,
                        command=self._toggle_denom).grid(
            row=0, column=0, sticky="w", padx=8, pady=3)

        self._denom_var = tk.StringVar()
        self._denom_entry = ttk.Entry(
            den_frame, textvariable=self._denom_var, width=16, state="disabled")
        self._denom_entry.grid(row=0, column=1, padx=8, pady=3, sticky="w")
        ttk.Label(den_frame, text='e.g. "100k + 1k"',
                  foreground="#888").grid(row=0, column=2, sticky="w", padx=4)

        # ── Buttons ───────────────────────────────────────────────────────
        btn_frame = ttk.Frame(parent)
        btn_frame.pack(fill="x", padx=10, pady=10)

        ttk.Button(btn_frame, text="Clear All",
                   command=self._clear_all).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="Send PC Mail",
                   command=self._send).pack(side="left", padx=5)

        # init payoff visibility
        self._on_product_change()

    # ── Hedge helpers ─────────────────────────────────────────────────────
    def _on_hedge_change(self):
        if self._hedge_type.get() == "BTB":
            self._btb_frame.grid()
        else:
            self._btb_frame.grid_remove()

    # ── Client / maturity auto-fill ───────────────────────────────────────
    def _compute_auto_struct_fee(self) -> str:
        """Return auto-calculated structuring fee as a string, or '' if unknown."""
        if self.vars["client"].get().strip() == "ING Bank":
            return "0.5"
        maturity = self.vars["maturity"].get().strip()
        try:
            years = int(maturity.rstrip("Yy"))
            fee = min(round(years * 0.3, 1), 1.5)
            return str(fee)
        except (ValueError, AttributeError):
            return ""

    def _on_maturity_change(self, event=None):
        fee = self._compute_auto_struct_fee()
        if fee:
            self.vars["struct_fee"].set(fee)

    def _on_client_change(self, event=None):
        client = self.vars["client"].get().strip()
        if client == "ING Bank":
            self.vars["struct_fee"].set("0.5")
            self.vars["dist_fee"].set("0.0")
            self._vlk_code_var.set(False)
            self._hedge_type.set("BTB")
            self._on_hedge_change()
            self._btb_hedge_search.set("ING")
            self._btb_vars["upfront"].set("0.5")
            self._btb_vars["amount"].set("")
        else:
            # Re-apply maturity-based fee when switching away from ING
            self._on_maturity_change()

    def _build_hedged_string(self) -> str:
        if self._hedge_type.get() == "Own Book":
            return "Own Book"
        party   = self._btb_hedge_search.get().strip()
        upfront = self._btb_vars["upfront"].get().strip()
        amount  = self._btb_vars["amount"].get().strip()
        return f"{amount} BTB with {party} @{upfront}%"

    # ── Underlying helpers ────────────────────────────────────────────────
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
            if len(joined) > 28:
                joined = joined[:26] + "…"
            self._ul_display_var.set(f"{joined}  ({n})")

    # ── Product-change ────────────────────────────────────────────────────
    def _on_product_change(self, *_):
        for w in self._payoff_widgets:
            w.destroy()
        self._payoff_widgets.clear()

        product = self.vars["product"].get()
        fields  = PRODUCT_PAYOFF_FIELDS.get(product, [])

        for i, (key, label) in enumerate(fields):
            lw = ttk.Label(self._payoff_frame, text=label + ":")
            lw.grid(row=i, column=0, sticky="w", padx=8, pady=3)
            ew = ttk.Entry(self._payoff_frame, textvariable=self.vars[key], width=10)
            ew.grid(row=i, column=1, padx=8, pady=3, sticky="w")
            self._payoff_widgets.extend([lw, ew])

        if product in ("Index Garantie Note", "Index Garantie Note Capped"):
            self._add_asianing_row(len(fields))

        # Autofill aflossingsbarrière = 100 for Trigger products
        if product == "Trigger Plus Note" and not self.vars["param2"].get().strip():
            self.vars["param2"].set("100")

        # Fixed Rate Note: dist fee = 0.25, struct fee = empty
        if product == "Fixed Rate Note":
            self.vars["struct_fee"].set("")
            self.vars["dist_fee"].set("0.25")
        elif not self.vars["struct_fee"].get().strip():
            # Restore maturity-based default when switching away from FRN
            fee = self._compute_auto_struct_fee()
            if fee:
                self.vars["struct_fee"].set(fee)

        # Show/hide underlying section (Fixed Rate Note has no underlying)
        if product in PRODUCTS_NO_UNDERLYING:
            self._ul_frame.pack_forget()
        elif not self._ul_frame.winfo_ismapped():
            self._ul_frame.pack(fill="x", padx=10, pady=4, after=self._payoff_frame)

    def _add_asianing_row(self, row: int):
        """Asianing button that expands to tail + obs entries, with a cancel/clear option."""
        container = ttk.Frame(self._payoff_frame)
        container.grid(row=row, column=0, columnspan=2, sticky="w", padx=4, pady=2)
        self._payoff_widgets.append(container)

        def _show_button():
            for child in container.winfo_children():
                child.destroy()
            ttk.Label(container, text="Asianing:").grid(row=0, column=0, sticky="w", padx=4)
            ttk.Button(container, text="▶ Set", command=_show_entries).grid(row=0, column=1, padx=4)

        def _clear_and_collapse():
            self._asianing_tail_var.set("")
            self._asianing_obs_var.set("")
            _show_button()

        def _show_entries():
            for child in container.winfo_children():
                child.destroy()
            ttk.Label(container, text="Tail (months):").grid(row=0, column=0, sticky="w", padx=4, pady=2)
            ttk.Entry(container, textvariable=self._asianing_tail_var, width=8).grid(row=0, column=1, padx=4)
            ttk.Label(container, text="Obs:").grid(row=1, column=0, sticky="w", padx=4, pady=2)
            ttk.Entry(container, textvariable=self._asianing_obs_var, width=8).grid(row=1, column=1, padx=4)
            ttk.Button(container, text="✕ Clear asianing",
                       command=_clear_and_collapse).grid(row=2, column=0, columnspan=2,
                                                         sticky="w", padx=4, pady=2)

        # Auto-expand if values already set (e.g. after switching product type back)
        if self._asianing_tail_var.get() or self._asianing_obs_var.get():
            _show_entries()
        else:
            _show_button()

    # ── Date / denom toggles ─────────────────────────────────────────────
    def _toggle_trade_date(self):
        self._trade_date_entry.configure(
            state="normal" if self._trade_override.get() else "disabled")
        if not self._trade_override.get():
            self._trade_date_var.set(date.today().strftime("%d/%m/%Y"))

    def _toggle_issue_date(self):
        self._issue_date_entry.configure(
            state="normal" if self._issue_override.get() else "disabled")

    def _toggle_denom(self):
        self._denom_entry.configure(
            state="normal" if self._den_override.get() else "disabled")

    # ── Clear All ─────────────────────────────────────────────────────────
    def _clear_all(self):
        # Comboboxes / entries in self.vars
        for key, v in self.vars.items():
            v.set("")

        # Fees back to defaults
        self.vars["struct_fee"].set("1.5")
        self.vars["dist_fee"].set("0.0")

        # Options
        self._vlk_code_var.set(True)

        # Hedge → Own Book
        self._hedge_type.set("Own Book")
        self._on_hedge_change()
        self._btb_hedge_search.clear()
        for v in self._btb_vars.values():
            v.set("")

        # Underlyings
        self.selected_underlyings.clear()
        self._ul_search.clear()
        self._refresh_ul_display()

        # Trade date override
        self._trade_override.set(False)
        self._toggle_trade_date()

        # Issue date override
        self._issue_override.set(False)
        self._toggle_issue_date()
        trade = date.today()
        issue = _add_business_days(trade, 5)
        self._issue_date_var.set(issue.strftime("%d/%m/%Y"))

        # Asianing
        self._asianing_tail_var.set("")
        self._asianing_obs_var.set("")

        # Denomination override
        self._den_override.set(False)
        self._toggle_denom()
        self._denom_var.set("")

    # ── Build PCMailProduct ───────────────────────────────────────────────
    def _build_product(self):
        from PCMail.models.inputdefinition import PCMailProduct, Underlying

        trade_dt = date.today()
        if self._trade_override.get():
            try:
                from datetime import datetime as _dt
                trade_dt = _dt.strptime(self._trade_date_var.get().strip(), "%d/%m/%Y").date()
            except ValueError:
                pass
        trade_date_str = trade_dt.strftime("%d/%m/%Y")

        auto_issue = _add_business_days(trade_dt, 5)
        issue_date_str = (
            self._issue_date_var.get().strip()
            if self._issue_override.get()
            else auto_issue.strftime("%d/%m/%Y")
        )

        def _f(key: str) -> float:
            try:
                return float(self.vars[key].get())
            except ValueError:
                return 0.0

        return PCMailProduct(
            product=self.vars["product"].get(),
            series=self.vars["series"].get().strip(),
            issuer=self.vars["issuer"].get(),
            currency=self.vars["currency"].get(),
            client=self.vars["client"].get(),

            issue_size=self.vars["issue_size"].get().strip(),
            sold=self.vars["sold"].get().strip(),
            maturity=self.vars["maturity"].get().strip(),
            hedged=self._build_hedged_string(),

            coupon_protection=_f("param1"),
            participation=_f("param2"),
            barrier_cap=_f("param3"),
            redemption_barrier=_f("param4"),

            trade_date=trade_date_str,
            issue_date=issue_date_str,

            struct_fee=_f("struct_fee"),
            dist_fee=_f("dist_fee"),

            vlk_code_required=self._vlk_code_var.get(),
            attach_target_market=True,
            fill_word=True,

            underlyings=[Underlying(ticker=t) for t in self.selected_underlyings],
            denomination=self._denom_var.get().strip() if self._den_override.get() else "",
            tail=self._asianing_tail_var.get().strip(),
            obs=self._asianing_obs_var.get().strip(),
        )

    # ── SharePoint loader ─────────────────────────────────────────────────
    def _open_sharepoint_picker(self):
        from gui.dialogs.sharepoint_picker import SharePointPickerDialog
        dlg = SharePointPickerDialog(self.frame, SHAREPOINT_SUMMARY_PATH)
        if dlg.result:
            self._load_from_sharepoint(dlg.result)

    def _load_from_sharepoint(self, deal: dict):
        """Populate the form from a parsed SharePoint deal dict."""
        # Clear first so nothing stale remains
        self._clear_all()

        # ── Product identity ──────────────────────────────────────────────
        if deal.get("issuer") and deal["issuer"] in ISSUERS:
            self.vars["issuer"].set(deal["issuer"])
        if deal.get("product") and deal["product"] in list(PRODUCT_TYPES.keys()):
            self.vars["product"].set(deal["product"])
        if deal.get("series"):
            self.vars["series"].set(deal["series"])
        if deal.get("maturity") and deal["maturity"] in MATURITIES:
            self.vars["maturity"].set(deal["maturity"])
        if deal.get("client") and deal["client"] in CLIENTS:
            self.vars["client"].set(deal["client"])
        if deal.get("currency") and deal["currency"] in CURRENCIES:
            self.vars["currency"].set(deal["currency"])

        # ── Payoff params from long name ──────────────────────────────────
        for key in ("param1", "param2", "param3", "param4"):
            val = deal.get(key, "")
            if val:
                self.vars[key].set(val)

        # ── Underlyings ───────────────────────────────────────────────────
        for ul in deal.get("underlyings", []):
            self._add_underlying(ul)

        # ── Dates ─────────────────────────────────────────────────────────
        if deal.get("trade_date"):
            self._trade_override.set(True)
            self._toggle_trade_date()
            self._trade_date_var.set(deal["trade_date"])

        if deal.get("issue_date"):
            self._issue_override.set(True)
            self._toggle_issue_date()
            self._issue_date_var.set(deal["issue_date"])

        # ── VLK code ──────────────────────────────────────────────────────
        if "vlk_code" in deal:
            self._vlk_code_var.set(deal["vlk_code"])

        # ── Sold (total from comments adviser lines) ──────────────────────
        if deal.get("total_sold"):
            self.vars["sold"].set(deal["total_sold"])

        # ── Hedge info ────────────────────────────────────────────────────
        hedge = deal.get("hedge_party", "").strip()
        if hedge:
            self._hedge_type.set("BTB")
            self._on_hedge_change()
            self._btb_hedge_search.set(hedge)
            if deal.get("upfront"):
                self._btb_vars["upfront"].set(deal["upfront"])
            if deal.get("btb_amount"):
                self._btb_vars["amount"].set(deal["btb_amount"])

        # ── Re-render payoff widgets + fee defaults ────────────────────────
        self._on_product_change()
        self._on_maturity_change()

        # Restore params after _on_product_change cleared them
        for key in ("param1", "param2", "param3", "param4"):
            val = deal.get(key, "")
            if val:
                self.vars[key].set(val)

    # ── Actions ───────────────────────────────────────────────────────────
    def _send(self):
        try:
            if not self.vars["product"].get():
                messagebox.showerror("Validation", "Please select a Product.")
                return
            if not self.vars["series"].get().strip():
                messagebox.showerror("Validation", "Series is required.")
                return
            if (not self.selected_underlyings
                    and self.vars["product"].get() not in PRODUCTS_NO_UNDERLYING):
                messagebox.showerror("Validation",
                                     "Please add at least one Underlying.")
                return

            p = self._build_product()

            from PCMail.main import run_pc_mail
            run_pc_mail(p)

            messagebox.showinfo("Success", "✅ PC Mail sent!")

        except Exception as e:
            messagebox.showerror("Error", str(e))
            raise