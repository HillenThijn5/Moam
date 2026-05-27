import tkinter as tk
from tkinter import ttk, messagebox
from datetime import date, timedelta, datetime
from statics.data import (
    CURRENCIES, MATURITIES, MARKETING_PRODUCT_TYPES, MARKETING_ISSUERS,
    UNDERLYINGS, SHAREPOINT_SUMMARY_PATH,
)
from MarketingMail.models import MarketingProduct
from MarketingMail.mail_service import create_and_send_marketing_mail
from statics.data import MARKETING_PARAMETER_CONFIG as PARAMETER_CONFIG
from gui.widgets import UnderlyingSearchEntry


def _add_business_days(start: date, n: int) -> date:
    d, added = start, 0
    while added < n:
        d += timedelta(days=1)
        if d.weekday() < 5:
            added += 1
    return d


class MarketingMailTab:
    """Marketing Mail-tab — één product op het scherm, buffer tot 4 producten."""

    def __init__(self, parent):
        self.frame = ttk.Frame(parent)

        # Buffer voor tot 4 producten
        self.products: list[dict] = []
        self.max_products = 4

        # Asianing-variabelen blijven behouden bij producttypewijzigingen
        self._asianing_tail_var = tk.StringVar(value="")
        self._asianing_obs_var = tk.StringVar(value="")

        # ── SharePoint-werkbalk ──────────────────────────────────────────
        toolbar = ttk.Frame(self.frame)
        toolbar.pack(fill="x", padx=10, pady=(8, 0))
        ttk.Button(
            toolbar, text="📋 Load from SharePoint",
            command=self._open_sharepoint_picker,
        ).pack(side="left", padx=(0, 8))
        ttk.Label(
            toolbar,
            text="Auto-fill current product form from the SharePoint deal list",
            foreground="#888", font=("Segoe UI", 8),
        ).pack(side="left")

        # Bovenkant in twee kolommen: links = instellingen, rechts = productparameters
        top_frame = ttk.Frame(self.frame)
        top_frame.pack(fill="x", padx=10, pady=10)
        top_frame.columnconfigure(0, weight=1)
        top_frame.columnconfigure(1, weight=1)

        info_frame = ttk.LabelFrame(top_frame, text="Marketing Mail Settings")
        info_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 10))

        self.params_frame = ttk.LabelFrame(top_frame, text="Parameters")
        self.params_frame.grid(row=0, column=1, sticky="nsew")
        self.params_frame.columnconfigure(0, weight=0)
        self.params_frame.columnconfigure(1, weight=1)

        self.vars = {}
        row = 0

        # Status van het mandje met onderliggende waarden
        self.selected_underlyings: list[str] = []

        # Producttype
        ttk.Label(info_frame, text="Product Type:").grid(row=row, column=0, sticky="w", padx=5, pady=4)
        default_pt = MARKETING_PRODUCT_TYPES[0] if MARKETING_PRODUCT_TYPES else "TRIGGER"
        self.vars["product_type"] = tk.StringVar(value=default_pt)
        product_values = list(MARKETING_PRODUCT_TYPES) if MARKETING_PRODUCT_TYPES else list(PARAMETER_CONFIG.keys())
        product_combo = ttk.Combobox(
            info_frame, textvariable=self.vars["product_type"],
            values=product_values, state="readonly", width=30,
        )
        product_combo.grid(row=row, column=1, padx=5, pady=4)
        product_combo.bind("<<ComboboxSelected>>", self._on_product_type_changed)
        row += 1

        # Emittent
        ttk.Label(info_frame, text="Issuer:").grid(row=row, column=0, sticky="w", padx=5, pady=4)
        self.vars["issuer"] = tk.StringVar(value=MARKETING_ISSUERS[0])
        issuer_combo = ttk.Combobox(
            info_frame, textvariable=self.vars["issuer"],
            values=MARKETING_ISSUERS, state="readonly", width=30,
        )
        issuer_combo.grid(row=row, column=1, padx=5, pady=4)
        row += 1

        # Underlying — UnderlyingSearchEntry (zelfde patroon als doc mail / PC mail)
        ttk.Label(info_frame, text="Underlying:").grid(row=row, column=0, sticky="w", padx=5, pady=4)
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

        # Valuta
        ttk.Label(info_frame, text="Currency:").grid(row=row, column=0, sticky="w", padx=5, pady=4)
        self.vars["currency"] = tk.StringVar()
        ttk.Combobox(info_frame, textvariable=self.vars["currency"],
                     values=CURRENCIES, state="readonly", width=30).grid(row=row, column=1, padx=5, pady=4)
        row += 1

        # Looptijd
        ttk.Label(info_frame, text="Maturity:").grid(row=row, column=0, sticky="w", padx=5, pady=4)
        self.vars["maturity"] = tk.StringVar(value="5Y")
        ttk.Combobox(info_frame, textvariable=self.vars["maturity"],
                     values=MATURITIES, state="readonly", width=30).grid(row=row, column=1, padx=5, pady=4)
        row += 1

        # Status van de parametervelden
        self.pv_vars: dict[str, tk.StringVar] = {
            key: tk.StringVar(value="") for key in ["param1", "param2", "param3", "param4"]
        }
        self._payoff_widgets: list = []
        self._prev_fixed: set[str] = set()

        # Parametervelden rechts
        self._update_params_widgets()

        # Voorbeeld van de productbuffer
        preview_frame = ttk.LabelFrame(self.frame, text="Buffered Products (max 4)")
        preview_frame.pack(fill="both", expand=False, padx=10, pady=10)
        self.products_listbox = tk.Listbox(preview_frame, height=5)
        self.products_listbox.pack(fill="x", padx=5, pady=5)
        preview_btns = ttk.Frame(preview_frame)
        preview_btns.pack(fill="x", padx=5, pady=5)
        ttk.Button(preview_btns, text="Remove Selected", command=self._remove_selected_product).pack(side="left", padx=5)
        ttk.Button(preview_btns, text="Clear All", command=self._clear_all_products).pack(side="left", padx=5)

        # ── Datums ───────────────────────────────────────────────────────────
        _trade = date.today()
        _issue = _add_business_days(_trade, 5)

        dates_frame = ttk.LabelFrame(self.frame, text="Dates")
        dates_frame.pack(fill="x", padx=10, pady=(0, 4))

        # Strike Date-rij
        ttk.Label(dates_frame, text="Strike Date:").grid(
            row=0, column=0, sticky="w", padx=8, pady=3)
        self._trade_override = tk.BooleanVar(value=False)
        ttk.Checkbutton(dates_frame, text="Override",
                        variable=self._trade_override,
                        command=self._toggle_trade_date).grid(
            row=0, column=1, sticky="w", padx=8)
        self._trade_date_var = tk.StringVar(value=_trade.strftime("%d/%m/%Y"))
        self._trade_date_entry = ttk.Entry(
            dates_frame, textvariable=self._trade_date_var,
            width=14, state="disabled")
        self._trade_date_entry.grid(row=0, column=2, sticky="w", padx=4, pady=3)
        self._trade_date_var.trace_add("write", self._on_trade_date_changed)
        ttk.Label(dates_frame, text="(auto = today)",
                  foreground="#888").grid(row=0, column=3, sticky="w", padx=4)

        # Issue Date-rij
        ttk.Label(dates_frame, text="Issue Date:").grid(
            row=1, column=0, sticky="w", padx=8, pady=3)
        self._issue_override = tk.BooleanVar(value=False)
        ttk.Checkbutton(dates_frame, text="Override",
                        variable=self._issue_override,
                        command=self._toggle_issue_date).grid(
            row=1, column=1, sticky="w", padx=8)
        self._issue_date_var = tk.StringVar(value=_issue.strftime("%d/%m/%Y"))
        self._issue_date_entry = ttk.Entry(
            dates_frame, textvariable=self._issue_date_var,
            width=14, state="disabled")
        self._issue_date_entry.grid(row=1, column=2, sticky="w", padx=4, pady=3)
        ttk.Label(dates_frame, text="(auto = trade + 5 business days)",
                  foreground="#888").grid(row=1, column=3, sticky="w", padx=4)

        # Actieknoppen
        button_frame = ttk.Frame(self.frame)
        button_frame.pack(fill="x", padx=10, pady=10)
        ttk.Button(button_frame, text="+ Product", command=self.on_add_product).pack(side="left", padx=5)
        ttk.Button(button_frame, text="Send", command=self._send).pack(side="left", padx=5)
        self.count_label = ttk.Label(button_frame, text="0/4")
        self.count_label.pack(side="right", padx=5)

        self._refresh_products_preview()

    # ──────────────────────────── mandje met onderliggende waarden ───────────

    def _add_underlying(self, ticker: str):
        ticker = ticker.strip().upper()
        if not ticker or ticker in self.selected_underlyings:
            return
        # Exacte match in bekende onderliggende waarden
        if ticker in UNDERLYINGS:
            self.selected_underlyings.append(ticker)
            self._refresh_ul_display()
            return
        # Benadering: vind onderliggende waarden die beginnen met de verwerkte ticker
        matches = [u for u in UNDERLYINGS if u.upper().startswith(ticker)]
        if len(matches) == 1:
            resolved = matches[0]
            if resolved not in self.selected_underlyings:
                self.selected_underlyings.append(resolved)
                self._refresh_ul_display()
            return
        # Terugval: voeg toe zoals het is (gebruiker kan handmatig corrigeren)
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
            n = len(self.selected_underlyings)
            joined = " / ".join(self.selected_underlyings)
            if len(joined) > 20:
                joined = joined[:18] + "…"
            self._ul_display_var.set(f"{joined}  ({n})")

    # ──────────────────────────── parametervelden ────────────────────────────

    def _update_params_widgets(self):
        """Gooi alle payoff-velden weg en bouw ze opnieuw op voor het huidige producttype."""
        for w in self._payoff_widgets:
            w.destroy()
        self._payoff_widgets.clear()

        product_type = self.vars["product_type"].get().strip()
        cfg = PARAMETER_CONFIG.get(product_type, PARAMETER_CONFIG.get("TRIGGER", {}))
        use_asianing = cfg.get("asianing", False)
        show_param3 = cfg.get("show_param3", True)

        keys_to_show = ["param1", "param2"] + (["param3"] if show_param3 else []) + ([] if use_asianing else ["param4"])
        row = 0
        for key in keys_to_show:
            spec = cfg.get(key, {"label": key, "required": False})
            fixed = spec.get("fixed")

            # Als deze sleutel de vorige keer vast stond maar nu bewerkbaar is, wis dan de oude vaste waarde
            if key in self._prev_fixed and fixed is None:
                self.pv_vars[key].set("")

            if fixed is not None:
                self.pv_vars[key].set(str(fixed))

            lbl = ttk.Label(self.params_frame, text=spec["label"] + ":")
            lbl.grid(row=row, column=0, sticky="w", padx=6, pady=4)
            ent = ttk.Entry(
                self.params_frame, textvariable=self.pv_vars[key], width=18,
                state="disabled" if fixed is not None else "normal",
            )
            ent.grid(row=row, column=1, sticky="w", padx=6, pady=4)
            self._payoff_widgets.extend([lbl, ent])
            row += 1

        if use_asianing:
            self._add_asianing_row(row)

        # Onthoud welke sleutels in deze opbouw vast stonden
        self._prev_fixed = {
            key for key in ["param1", "param2", "param3", "param4"]
            if cfg.get(key, {}).get("fixed") is not None
        }

        # Vul aflossingsbarrière automatisch in = 100 voor TRIGGER producten
        if product_type == "TRIGGER" and not self.pv_vars["param2"].get().strip():
            self.pv_vars["param2"].set("100")

    def _add_asianing_row(self, row: int):
        """Zet de Asianing-knop/velden op de opgegeven grid-rij (zelfde patroon als PC Mail)."""
        container = ttk.Frame(self.params_frame)
        container.grid(row=row, column=0, columnspan=2, sticky="w", padx=4, pady=2)
        self._payoff_widgets.append(container)

        def _show_button():
            for child in container.winfo_children():
                child.destroy()
            ttk.Label(container, text="Middeling:").grid(row=0, column=0, sticky="w", padx=4)
            ttk.Button(container, text="▶ Set", command=_show_entries).grid(row=0, column=1, padx=4)

        def _clear_and_collapse():
            self._asianing_tail_var.set("")
            self._asianing_obs_var.set("")
            _show_button()

        def _show_entries():
            for child in container.winfo_children():
                child.destroy()
            ttk.Label(container, text="Tail (m):").grid(row=0, column=0, sticky="w", padx=4, pady=2)
            ttk.Entry(container, textvariable=self._asianing_tail_var, width=8).grid(row=0, column=1, padx=4)
            ttk.Label(container, text="Obs:").grid(row=1, column=0, sticky="w", padx=4, pady=2)
            ttk.Entry(container, textvariable=self._asianing_obs_var, width=8).grid(row=1, column=1, padx=4)
            ttk.Button(container, text="✕ Clear", command=_clear_and_collapse).grid(
                row=2, column=0, columnspan=2, sticky="w", padx=4, pady=2)

        if self._asianing_tail_var.get() or self._asianing_obs_var.get():
            _show_entries()
        else:
            _show_button()

    def _get_params_values(self) -> dict:
        """Geef gevalideerde parameterwaarden terug, gemapt op de veldnamen van MarketingProduct."""
        product_type = self.vars["product_type"].get().strip()
        cfg = PARAMETER_CONFIG.get(product_type, PARAMETER_CONFIG.get("TRIGGER", {}))
        use_asianing = cfg.get("asianing", False)
        show_param3 = cfg.get("show_param3", True)

        values = {}

        for key in ["param1", "param2"]:
            spec = cfg.get(key, {})
            val = self.pv_vars[key].get().strip()
            if spec.get("required") and not val:
                messagebox.showerror("Missing fields", f"Fill in: {spec.get('label', key)}")
                raise ValueError(f"Missing field: {key}")
            if val and val not in ("n.v.t.", "nvt"):
                values[key] = val

        if show_param3:
            spec = cfg.get("param3", {})
            val = self.pv_vars["param3"].get().strip()
            if spec.get("required") and not val:
                messagebox.showerror("Missing fields", f"Fill in: {spec.get('label', 'param3')}")
                raise ValueError("Missing field: param3")
            if val and val not in ("n.v.t.", "nvt"):
                values["param3"] = val

        if use_asianing:
            values["tail"] = self._asianing_tail_var.get().strip()
            values["obs"] = self._asianing_obs_var.get().strip()
        else:
            spec = cfg.get("param4", {})
            val = self.pv_vars["param4"].get().strip()
            if spec.get("required") and not val:
                messagebox.showerror("Missing fields", f"Fill in: {spec.get('label', 'Protection %')}")
                raise ValueError("Missing field: param4")
            if val:
                values["param4"] = val

        return values

    def _clear_params_values(self):
        for v in self.pv_vars.values():
            v.set("")
        self._asianing_tail_var.set("")
        self._asianing_obs_var.set("")

    def _params_are_empty(self) -> bool:
        all_empty = all(not v.get().strip() for v in self.pv_vars.values())
        no_asianing = not self._asianing_tail_var.get() and not self._asianing_obs_var.get()
        return all_empty and no_asianing

    def _on_product_type_changed(self, event=None):
        self._update_params_widgets()

    # ──────────────────────────── SharePoint-inlader ─────────────────────────

    # Mapt SharePoint-productnaam → productsleutel voor Marketing Mail
    _SP_PRODUCT_MAP: dict[str, str] = {
        "Trigger Plus Note":          "TRIGGER",
        "Memory Coupon":              "MEMORY_COUPON",
        "Index Garantie Note":        "INDEX_GARANTIE",
        "Index Garantie Note Capped": "INDEX_GARANTIE_CAPPED",
    }

    # Zet de korte SharePoint-issuercode om naar de volledige MARKETING_ISSUERS-string
    # Sleutels komen overeen met wat sharepoint/parser.py ISSUER_CODE_MAP teruggeeft
    _SP_ISSUER_MAP: dict[str, str] = {
        "VLK":                 MARKETING_ISSUERS[0],
        "BNP Paribas Issuance": MARKETING_ISSUERS[1],
        "UBS AG":              MARKETING_ISSUERS[2],
        "Société Générale":    MARKETING_ISSUERS[3],
    }

    def _open_sharepoint_picker(self):
        from gui.sharepointgui.sharepoint_picker import SharePointPickerDialog
        dlg = SharePointPickerDialog(
            self.frame.winfo_toplevel(),
            str(SHAREPOINT_SUMMARY_PATH),
        )
        if dlg.result:
            self._load_from_sharepoint(dlg.result)

    def _load_from_sharepoint(self, deal: dict):
        """Vul het huidige productformulier automatisch in vanuit een verwerkte SharePoint-deal."""
        # Producttype
        product_type = self._SP_PRODUCT_MAP.get(deal.get("product", ""))
        if product_type and product_type in MARKETING_PRODUCT_TYPES:
            self.vars["product_type"].set(product_type)
            self._prev_fixed = set()
            self._update_params_widgets()

        # Emittent
        issuer = self._SP_ISSUER_MAP.get(deal.get("issuer", ""))
        if issuer:
            self.vars["issuer"].set(issuer)

        # Valuta
        if deal.get("currency") and deal["currency"] in CURRENCIES:
            self.vars["currency"].set(deal["currency"])

        # Looptijd ("5Y", "7Y" — al in MATURITIES-formaat)
        if deal.get("maturity") and deal["maturity"] in MATURITIES:
            self.vars["maturity"].set(deal["maturity"])

        # Onderliggende waarden
        self.selected_underlyings.clear()
        self._ul_search.clear()
        for ul in deal.get("underlyings", []):
            ul = ul.strip().upper()
            if ul:
                self.selected_underlyings.append(ul)
        self._refresh_ul_display()

        # Payoff-parameters — vul in wat SP heeft verwerkt; de gebruiker kan de rest aanpassen
        for key in ["param1", "param2", "param3", "param4"]:
            val = deal.get(key, "")
            if val:
                self.pv_vars[key].set(val)

        # Datums — zet handmatige velden als SP ze heeft
        if deal.get("trade_date"):
            self._trade_override.set(True)
            self._toggle_trade_date()
            self._trade_date_var.set(deal["trade_date"])
        if deal.get("issue_date"):
            self._issue_override.set(True)
            self._toggle_issue_date()
            self._issue_date_var.set(deal["issue_date"])

    # ──────────────────────────── datumhelpers ───────────────────────────────

    def _toggle_trade_date(self):
        state = "normal" if self._trade_override.get() else "disabled"
        self._trade_date_entry.configure(state=state)
        if not self._trade_override.get():
            # Zet terug naar vandaag en bereken de Issue Date opnieuw
            self._trade_date_var.set(date.today().strftime("%d/%m/%Y"))

    def _toggle_issue_date(self):
        state = "normal" if self._issue_override.get() else "disabled"
        self._issue_date_entry.configure(state=state)
        if not self._issue_override.get():
            self._recompute_issue_date()

    def _on_trade_date_changed(self, *_):
        if not self._issue_override.get():
            self._recompute_issue_date()

    def _recompute_issue_date(self):
        try:
            trade = datetime.strptime(self._trade_date_var.get(), "%d/%m/%Y").date()
            issue = _add_business_days(trade, 5)
            self._issue_date_var.set(issue.strftime("%d/%m/%Y"))
        except ValueError:
            pass  # ongeldige datum tijdens typen — negeer

    def _get_trade_date_model_fmt(self) -> str:
        """Geef de Trade Date terug in het modelformaat '%d %b %Y' (bijv. '17 Apr 2026')."""
        try:
            return datetime.strptime(self._trade_date_var.get(), "%d/%m/%Y").strftime("%d %b %Y")
        except ValueError:
            return date.today().strftime("%d %b %Y")

    def _get_issue_date_model_fmt(self) -> str:
        """Geef de Issue Date terug in het modelformaat '%d %b %Y' (bijv. '17 Apr 2026')."""
        try:
            return datetime.strptime(self._issue_date_var.get(), "%d/%m/%Y").strftime("%d %b %Y")
        except ValueError:
            return _add_business_days(date.today(), 5).strftime("%d %b %Y")

    # ──────────────────────────── productbuffer ──────────────────────────────

    def _form_is_empty(self) -> bool:
        # Onderliggende waarden + valuta zijn de enige bewuste signalen; parameters kunnen automatisch ingevuld zijn
        return (not self.selected_underlyings
                and not self.vars["currency"].get().strip())

    def _collect_current_product(self, allow_empty: bool = False) -> dict | None:
        if allow_empty and self._form_is_empty():
            return None

        product_type = self.vars["product_type"].get().strip()
        currency = self.vars["currency"].get().strip()
        underlyings = list(self.selected_underlyings)

        if not underlyings:
            messagebox.showerror("Missing fields", "Please select at least 1 underlying.")
            return None
        missing = [name for name, val in [("Product Type", product_type), ("Currency", currency)] if not val]
        if missing:
            messagebox.showerror("Missing fields", "Please fill: " + ", ".join(missing))
            return None

        params = self._get_params_values()

        return {
            "product_type": product_type,
            "currency":     currency,
            "underlyings":  underlyings,
            "maturity":     self.vars["maturity"].get() or "5Y",
            "issuer":       self.vars["issuer"].get() or MARKETING_ISSUERS[0],
            "start_date":   self._get_trade_date_model_fmt(),
            "issue_date":   self._get_issue_date_model_fmt(),
            **params,
        }

    def _clear_product_form(self):
        self.vars["product_type"].set("TRIGGER")
        self._prev_fixed = set()
        self.vars["currency"].set("")
        self.vars["maturity"].set("5Y")
        self.vars["issuer"].set(MARKETING_ISSUERS[0])
        self.selected_underlyings.clear()
        self._ul_search.clear()
        self._refresh_ul_display()
        self._clear_params_values()
        self._update_params_widgets()
        # Zet datums terug op automatisch
        self._trade_override.set(False)
        self._toggle_trade_date()
        self._issue_override.set(False)
        self._toggle_issue_date()

    def _refresh_products_preview(self):
        self.products_listbox.delete(0, "end")
        for i, p in enumerate(self.products, start=1):
            u = " / ".join(p.get("underlyings", []))
            extras = []
            for key, label in [("param1", "p1"), ("param2", "p2"), ("param3", "p3"), ("param4", "p4"),
                                ("tail", "tail"), ("obs", "obs")]:
                if p.get(key):
                    extras.append(f"{label}={p[key]}")
            if p.get("issue_date"):
                extras.append(f"issue={p['issue_date']}")
            extra_txt = (" | " + ", ".join(extras)) if extras else ""
            self.products_listbox.insert("end", f"{i}) {p['product_type']} | {u} | {p['currency']} | {p.get('maturity', '5Y')}{extra_txt}")
        self.count_label.config(text=f"{len(self.products)}/{self.max_products}")

    def on_add_product(self):
        if len(self.products) >= self.max_products:
            messagebox.showwarning("Max products", f"Max {self.max_products} products allowed.")
            return
        product = self._collect_current_product()
        if product is None:
            return
        self.products.append(product)
        self._clear_product_form()
        self._refresh_products_preview()

    def _remove_selected_product(self):
        sel = self.products_listbox.curselection()
        if not sel:
            return
        del self.products[sel[0]]
        self._refresh_products_preview()

    def _clear_all_products(self):
        self.products.clear()
        self._refresh_products_preview()

    # ──────────────────────────── voorbeeld + versturen ──────────────────────

    def _send(self):
        try:
            products_to_send = list(self.products)
            current = self._collect_current_product(allow_empty=True)
            if current is not None:
                if len(products_to_send) >= self.max_products:
                    messagebox.showwarning("Max products", f"Max {self.max_products} products allowed.")
                    return
                products_to_send.append(current)

            if not products_to_send:
                messagebox.showerror("Error", "No products to send. Add a product or fill the form.")
                return

            marketing_products = []
            for p in products_to_send:
                underlying_str = " / ".join(p.get("underlyings", []))
                mat_gui = p.get("maturity", "5Y")
                mat_model = mat_gui.replace("Y", " jaar").replace("y", " jaar")
                marketing_products.append(
                    MarketingProduct(
                        product_type=p["product_type"],
                        currency=p["currency"],
                        underlying=underlying_str,
                        maturity=mat_model,
                        issuer=p.get("issuer", MARKETING_ISSUERS[0]),
                        coupon_protection=p.get("param1"),
                        participation=p.get("param2"),
                        barrier_cap=p.get("param3"),
                        redemption_barrier=p.get("param4"),
                        tail=p.get("tail"),
                        obs=p.get("obs"),
                        start_date=p["start_date"],
                        issue_date=p["issue_date"],
                    )
                )

            mail = create_and_send_marketing_mail(
                title="",
                products=marketing_products,
                choice=len(marketing_products),
            )

            self.products.clear()
            self._clear_product_form()
            self._refresh_products_preview()
            messagebox.showinfo("Success", f"✅ Marketing Mail created: {getattr(mail, 'Subject', '(no subject)')}")

        except ValueError:
            pass  # al via messagebox getoond
        except Exception as e:
            from statics.mail_debug import full_diagnostics
            err_msg = str(e)
            answer = messagebox.askyesno(
                "Error",
                f"{err_msg}\n\nWould you like to see the full debug report?"
            )
            if answer:
                report = full_diagnostics()
                self._show_debug_report(report, err_msg)

    def _show_debug_report(self, report: str, error: str):
        """Toon een debugrapport in een scrollbaar venster."""
        import tkinter as tk
        win = tk.Toplevel(self.frame.winfo_toplevel())
        win.title(f"Debug Report - {error[:50]}")
        win.geometry("750x600")
        text = tk.Text(win, wrap="word", font=("Consolas", 9))
        vsb = ttk.Scrollbar(win, orient="vertical", command=text.yview)
        text.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        text.pack(fill="both", expand=True)
        text.insert("end", report)
        text.config(state="disabled")