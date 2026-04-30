# gui/tabs/increase_decrease_tab.py
import tkinter as tk
from tkinter import ttk, messagebox
import threading


class IncreaseDecreaseTab:
    """
    Increase / Decrease Mail tab.

    Workflow:
    1. Click 'Fetch Positions' — queries the DB for current short MTN positions.
    2. Select a row in the treeview.
    3. Enter the nominal Size and choose Increase or Decrease.
    4. 'Amount' is auto-calculated as TransferPrice / 100 * Size.
    5. Click 'Send' to open a draft in Outlook.
    """

    def __init__(self, parent):
        self.frame = ttk.Frame(parent)
        self._positions: list[dict] = []   # raw rows from DB
        self._selected: dict | None = None

        self._build_ui()

    # ──────────────────────────────────────────── UI construction ─────────────
    def _build_ui(self):
        # ── Fetch section ─────────────────────────────────────────────────────
        fetch_frame = ttk.LabelFrame(self.frame, text="Database")
        fetch_frame.pack(fill="x", padx=10, pady=(10, 4))

        ttk.Button(fetch_frame, text="Fetch Positions",
                   command=self._fetch_positions).pack(side="left", padx=8, pady=6)

        self._status_var = tk.StringVar(value="Not loaded")
        ttk.Label(fetch_frame, textvariable=self._status_var,
                  foreground="#555").pack(side="left", padx=8)

        # ── Positions treeview ────────────────────────────────────────────────
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

        # ── Deal details ──────────────────────────────────────────────────────
        deal_frame = ttk.LabelFrame(self.frame, text="Deal Details")
        deal_frame.pack(fill="x", padx=10, pady=4)

        # Direction
        ttk.Label(deal_frame, text="Direction:").grid(
            row=0, column=0, sticky="w", padx=8, pady=4)
        self._direction_var = tk.StringVar(value="Increase")
        dir_row = ttk.Frame(deal_frame)
        dir_row.grid(row=0, column=1, sticky="w", padx=8, pady=4)
        ttk.Radiobutton(dir_row, text="Increase (Opbouwen)",
                        variable=self._direction_var, value="Increase",
                        command=self._on_direction_change).pack(side="left", padx=(0, 16))
        ttk.Radiobutton(dir_row, text="Decrease (Afbouwen)",
                        variable=self._direction_var, value="Decrease",
                        command=self._on_direction_change).pack(side="left")

        # Pays / Receives (auto, read-only display)
        ttk.Label(deal_frame, text="Pays / Receives:").grid(
            row=1, column=0, sticky="w", padx=8, pady=4)
        self._pays_var = tk.StringVar(value="Pays")
        ttk.Label(deal_frame, textvariable=self._pays_var,
                  foreground="#003366", font=("Segoe UI", 10, "bold")).grid(
            row=1, column=1, sticky="w", padx=8, pady=4)

        # Size
        ttk.Label(deal_frame, text="Size (nominal):").grid(
            row=2, column=0, sticky="w", padx=8, pady=4)
        size_row = ttk.Frame(deal_frame)
        size_row.grid(row=2, column=1, sticky="w", padx=8, pady=4)
        self._size_var = tk.StringVar()
        self._size_var.trace_add("write", self._on_size_change)
        self._size_entry = ttk.Entry(size_row, textvariable=self._size_var, width=18)
        self._size_entry.pack(side="left")
        self._size_entry.bind("<FocusOut>", self._format_size)
        ttk.Label(size_row, text="  (e.g. 500.000 or 1.250.000)",
                  foreground="#888", font=("Segoe UI", 8)).pack(side="left")

        # Selected product info
        ttk.Label(deal_frame, text="Selected product:").grid(
            row=3, column=0, sticky="w", padx=8, pady=4)
        self._selected_var = tk.StringVar(value="— none —")
        sel_entry = ttk.Entry(deal_frame, textvariable=self._selected_var,
                              state="readonly", width=55)
        sel_entry.grid(row=3, column=1, sticky="w", padx=8, pady=4)

        # ── Buttons ───────────────────────────────────────────────────────────
        btn_frame = ttk.Frame(self.frame)
        btn_frame.pack(fill="x", padx=10, pady=10)
        ttk.Button(btn_frame, text="Send Mail", command=self._send).pack(side="left", padx=5)

    # ──────────────────────────────────────────── helpers ─────────────────────
    def _format_size(self, _=None):
        """On focus-out, reformat the size value with dot-separated thousands."""
        raw = self._size_var.get().replace(".", "").replace(",", "").strip()
        if not raw:
            return
        try:
            value = int(float(raw))
            # Format with dots as thousands separator (Dutch style: 1.250.000)
            self._size_var.set(f"{value:,}".replace(",", "."))
        except ValueError:
            pass

    def _on_direction_change(self):
        self._pays_var.set("Pays" if self._direction_var.get() == "Increase" else "Receives")

    def _on_size_change(self, *_):
        pass  # size field kept for mail injection; no live display needed

    def _on_select(self, _=None):
        sel = self._tree.selection()
        if not sel:
            self._selected = None
            self._selected_var.set("— none —")
            return
        iid = sel[0]
        row = self._tree.item(iid, "values")
        name, isin, position, tp = row
        self._selected = {
            "Name":          name,
            "ISIN":          isin,
            "Position":      position,
            "TransferPrice": float(tp) if tp not in ("", "None", None) else None,
        }
        self._selected_var.set(f"{name}  |  {isin}")

        # ──────────────────────────────────────────── DB fetch ────────────────────
    def _fetch_positions(self):
        self._status_var.set("Fetching…")
        self._tree.delete(*self._tree.get_children())
        self._selected = None
        self._selected_var.set("— none —")

        def _run():
            try:
                from increase_decrease_mail.db import fetch_positions
                rows = fetch_positions()
                self.frame.after(0, lambda: self._populate_tree(rows))
            except Exception as exc:
                msg = str(exc)
                self.frame.after(0, lambda: self._on_fetch_error(msg))

        threading.Thread(target=_run, daemon=True).start()

    def _populate_tree(self, rows: list[dict]):
        self._positions = rows
        for r in rows:
            tp = r.get("TransferPrice")
            tp_str = f"{abs(float(tp)):.4f}" if tp is not None else ""
            self._tree.insert("", "end", values=(
                r.get("Name", ""),
                r.get("ISIN", ""),
                r.get("Position", ""),
                tp_str,
            ))
        self._status_var.set(f"{len(rows)} position(s) loaded")

    def _on_fetch_error(self, msg: str):
        self._status_var.set("Error — see dialog")
        messagebox.showerror("DB Error", f"Could not fetch positions:\n\n{msg}")

    # ──────────────────────────────────────────── send ────────────────────────
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
            messagebox.showerror("Error", str(exc))
            raise
