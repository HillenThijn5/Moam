# gui/tabs/increase_decrease_tab.py
import tkinter as tk
from tkinter import ttk, messagebox
import threading


class IncreaseDecreaseTab:
    """
    Increase / Decrease Mail-tabblad.

    Werkwijze:
    1. Kies bovenaan Increase of Decrease.
    2. Klik op 'Fetch Positions' — haalt short- (Increase) of long-posities (Decrease) op uit de database.
    3. Selecteer een rij in de treeview.
    4. Vul de nominale Size in.
    5. Klik op 'Send' om een concept in Outlook te openen.
    """

    def __init__(self, parent):
        self.frame = ttk.Frame(parent)
        self._positions: list[dict] = []
        self._selected: dict | None = None

        self._build_ui()

    # ──────────────────────────────────────────── UI constructie ─────────────
    def _build_ui(self):
        # ── Richting + ophaalsectie (BOVENKANT) ───────────────────────────────
        fetch_frame = ttk.LabelFrame(self.frame, text="Direction & Database")
        fetch_frame.pack(fill="x", padx=10, pady=(10, 4))

        # Richting-radioknoppen — dit zet de gebruiker als eerste
        ttk.Label(fetch_frame, text="Direction:").grid(
            row=0, column=0, sticky="w", padx=8, pady=6)
        self._direction_var = tk.StringVar(value="Increase")
        dir_row = ttk.Frame(fetch_frame)
        dir_row.grid(row=0, column=1, sticky="w", padx=8, pady=6)
        ttk.Radiobutton(dir_row, text="Increase (Opbouwen)  — short positions",
                        variable=self._direction_var, value="Increase",
                        command=self._on_direction_change).pack(side="left", padx=(0, 16))
        ttk.Radiobutton(dir_row, text="Decrease (Afbouwen)  — long positions",
                        variable=self._direction_var, value="Decrease",
                        command=self._on_direction_change).pack(side="left")

        # Ophalen-knop + status op de rij eronder
        ttk.Button(fetch_frame, text="Fetch Positions",
                   command=self._fetch_positions).grid(
            row=1, column=0, sticky="w", padx=8, pady=(0, 6))
        self._status_var = tk.StringVar(value="Not loaded")
        ttk.Label(fetch_frame, textvariable=self._status_var,
                  foreground="#555").grid(row=1, column=1, sticky="w", padx=8)

        # ── Posities-treeview ────────────────────────────────────────────────
        tree_frame = ttk.LabelFrame(self.frame, text="Positions (select one)")
        tree_frame.pack(fill="both", expand=True, padx=10, pady=4)

        cols = ("Name", "ISIN", "Position", "TransferPrice")
        self._tree = ttk.Treeview(tree_frame, columns=cols, show="headings",
                                  selectmode="browse", height=10)

        col_widths = {"Name": 280, "ISIN": 130, "Position": 90, "TransferPrice": 110}
        for col in cols:
            self._tree.heading(col, text=col)
            self._tree.column(col, width=col_widths[col], anchor="w")

        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self._tree.yview)
        self._tree.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        self._tree.pack(fill="both", expand=True, padx=4, pady=4)
        self._tree.bind("<<TreeviewSelect>>", self._on_select)

        # ── Dealdetails ───────────────────────────────────────────────────────
        deal_frame = ttk.LabelFrame(self.frame, text="Deal Details")
        deal_frame.pack(fill="x", padx=10, pady=4)

        # Pays / Receives (afgeleid van de richting, alleen-lezen)
        ttk.Label(deal_frame, text="Pays / Receives:").grid(
            row=0, column=0, sticky="w", padx=8, pady=4)
        self._pays_var = tk.StringVar(value="Pays")
        ttk.Label(deal_frame, textvariable=self._pays_var,
                  foreground="#003366", font=("Segoe UI", 10, "bold")).grid(
            row=0, column=1, sticky="w", padx=8, pady=4)

        # Grootte
        ttk.Label(deal_frame, text="Size (nominal):").grid(
            row=1, column=0, sticky="w", padx=8, pady=4)
        size_row = ttk.Frame(deal_frame)
        size_row.grid(row=1, column=1, sticky="w", padx=8, pady=4)
        self._size_var = tk.StringVar()
        self._size_entry = ttk.Entry(size_row, textvariable=self._size_var, width=18)
        self._size_entry.pack(side="left")
        self._size_entry.bind("<FocusOut>", self._format_size)
        ttk.Label(size_row, text="  (e.g. 500.000 or 1.250.000)",
                  foreground="#888", font=("Segoe UI", 8)).pack(side="left")

        # Info over het gekozen product
        ttk.Label(deal_frame, text="Selected product:").grid(
            row=2, column=0, sticky="w", padx=8, pady=4)
        self._selected_var = tk.StringVar(value="— none —")
        ttk.Entry(deal_frame, textvariable=self._selected_var,
                  state="readonly", width=55).grid(
            row=2, column=1, sticky="w", padx=8, pady=4)

        # ── Knoppen ───────────────────────────────────────────────────────────
        btn_frame = ttk.Frame(self.frame)
        btn_frame.pack(fill="x", padx=10, pady=10)
        ttk.Button(btn_frame, text="Send Mail", command=self._send).pack(side="left", padx=5)

    # ──────────────────────────────────────────── hulpfuncties ─────────────────
    def _format_size(self, _=None):
        raw = self._size_var.get().replace(".", "").replace(",", "").strip()
        if not raw:
            return
        try:
            self._size_var.set(f"{int(float(raw)):,}".replace(",", "."))
        except ValueError:
            pass

    def _on_direction_change(self):
        """Richting veranderd — werk Pays/Receives bij en reset de tabel."""
        self._pays_var.set("Pays" if self._direction_var.get() == "Increase" else "Receives")
        # Wis posities zodat duidelijk is dat ze voor de nieuwe richting opnieuw opgehaald moeten worden
        self._tree.delete(*self._tree.get_children())
        self._positions = []
        self._selected = None
        self._selected_var.set("— none —")
        self._status_var.set("Direction changed — click Fetch Positions")

    def _on_select(self, _=None):
        sel = self._tree.selection()
        if not sel:
            self._selected = None
            self._selected_var.set("— none —")
            return
        row = self._tree.item(sel[0], "values")
        name, isin, position, tp = row
        self._selected = {
            "Name":          name,
            "ISIN":          isin,
            "Position":      position,
            "TransferPrice": float(tp) if tp not in ("", "None", None) else None,
        }
        self._selected_var.set(f"{name}  |  {isin}")
        # Vul Size automatisch in met de absolute positiewaarde (weergave is al in Nederlandse stijl geformatteerd)
        try:
            abs_pos = abs(int(str(position).replace(".", "").replace(",", "")))
            self._size_var.set(f"{abs_pos:,}".replace(",", "."))
        except (ValueError, TypeError):
            pass

    # ──────────────────────────────────────────── DB-ophalen ──────────────────
    def _fetch_positions(self):
        direction = self._direction_var.get().lower()   # "increase" of "decrease"
        label = "short" if direction == "increase" else "long"
        self._status_var.set(f"Fetching {label} positions…")
        self._tree.delete(*self._tree.get_children())
        self._selected = None
        self._selected_var.set("— none —")

        def _run():
            try:
                from increase_decrease_mail.db import fetch_positions
                rows = fetch_positions(direction=direction)
                self.frame.after(0, lambda: self._populate_tree(rows))
            except Exception as exc:
                msg = str(exc)
                self.frame.after(0, lambda: self._on_fetch_error(msg))

        threading.Thread(target=_run, daemon=True).start()

    def _populate_tree(self, rows: list[dict]):
        self._positions = rows
        # Sorteer op absolute positie: kleinste eerst bij Increase, grootste eerst bij Decrease
        reverse = (self._direction_var.get() == "Decrease")
        rows_sorted = sorted(
            rows,
            key=lambda r: abs(float(r["Position"])) if r.get("Position") not in (None, "") else 0,
            reverse=reverse,
        )
        for r in rows_sorted:
            tp = r.get("TransferPrice")
            tp_str = f"{abs(float(tp)):.4f}" if tp is not None else ""
            raw_pos = r.get("Position")
            try:
                pos_str = f"{int(float(raw_pos)):,}".replace(",", ".")
            except (ValueError, TypeError):
                pos_str = str(raw_pos)
            self._tree.insert("", "end", values=(
                r.get("Name", ""),
                r.get("ISIN", ""),
                pos_str,
                tp_str,
            ))
        self._status_var.set(f"{len(rows)} position(s) loaded")

    def _on_fetch_error(self, msg: str):
        self._status_var.set("Error — see dialog")
        messagebox.showerror("DB Error", f"Could not fetch positions:\n\n{msg}")

    # ──────────────────────────────────────────── versturen ───────────────────
    def _send(self):
        if not self._selected:
            messagebox.showerror("Validation", "Please select a position from the table.")
            return

        size_raw = self._size_var.get().strip().replace(".", "").replace(",", "")
        try:
            size = float(size_raw)
            if size <= 0:
                raise ValueError
        except ValueError:
            messagebox.showerror("Validation", "Please enter a valid positive nominal Size.")
            return

        try:
            from increase_decrease_mail.mail_service import send_increase_decrease_mail
            send_increase_decrease_mail(
                name=self._selected["Name"],
                isin=self._selected["ISIN"],
                size=size,
                transfer_price=self._selected.get("TransferPrice"),
                increase=(self._direction_var.get() == "Increase"),
            )
            messagebox.showinfo("Success", "✅ Mail draft opened in Outlook.")
        except Exception as exc:
            from statics.mail_debug import full_diagnostics
            err_msg = str(exc)
            answer = messagebox.askyesno(
                "Error",
                f"{err_msg}\n\nWould you like to see the full debug report?"
            )
            if answer:
                report = full_diagnostics()
                self._show_debug_report(report, err_msg)

    def _show_debug_report(self, report: str, error: str):
        """Toon debug report in een scrollbaar venster."""
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

