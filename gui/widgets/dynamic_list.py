from tkinter import ttk, messagebox
import tkinter as tk
from typing import List, Dict, Tuple
from gui.widgets.underlying_search import UnderlyingSearchEntry


class DynamicListManager:
    """Manages add/remove items with optional dropdown or search-entry selection"""

    def __init__(self, parent, label_text: str, columns: Tuple[str, ...] = ("Item",),
                 dropdown_options: Dict[str, List[str]] = None,
                 search_options: Dict[str, List[str]] = None,
                 defaults: Dict[str, str] = None):
        """
        Args:
            parent: parent widget
            label_text: label for the list
            columns: tuple of column names
            dropdown_options: col → list of options, rendered as readonly Combobox
            search_options:   col → list of options, rendered as UnderlyingSearchEntry
                              (type + Enter for best match, same as underlying picker)
            defaults: col → default value used when field is left blank
        """
        self.frame = ttk.LabelFrame(parent, text=label_text)
        self.columns = columns
        self.dropdown_options = dropdown_options or {}
        self.search_options   = search_options   or {}
        self.defaults = defaults or {}
        self.items: List[Dict[str, str]] = []

        # Input row
        input_frame = ttk.Frame(self.frame)
        input_frame.pack(fill="x", padx=5, pady=5)

        self.inputs = {}
        for col in columns:
            ttk.Label(input_frame, text=f"{col}:").pack(side="left", padx=2)

            if col in self.search_options:
                # Search entry: auto_clear=False keeps the selected value in the field
                search_entry = UnderlyingSearchEntry(
                    input_frame,
                    all_tickers=self.search_options[col],
                    on_select=lambda _: None,   # var is shared — no extra callback needed
                    auto_clear=False,
                )
                search_entry.pack(side="left", padx=2)
                # Share the search widget's StringVar so _add_item reads from it directly
                self.inputs[col] = search_entry._var

            elif col in self.dropdown_options:
                var = tk.StringVar()
                ttk.Combobox(
                    input_frame, textvariable=var,
                    values=self.dropdown_options[col],
                    state="readonly", width=12
                ).pack(side="left", padx=2)
                self.inputs[col] = var

            else:
                var = tk.StringVar()
                if col in self.defaults:
                    var.set(self.defaults[col])
                ttk.Entry(input_frame, textvariable=var, width=12).pack(side="left", padx=2)
                self.inputs[col] = var

        ttk.Button(input_frame, text="+ Add", command=self._add_item).pack(side="left", padx=5)

        # Listbox with scrollbar
        list_frame = ttk.Frame(self.frame)
        list_frame.pack(fill="both", expand=True, padx=5, pady=5)

        scrollbar = ttk.Scrollbar(list_frame)
        scrollbar.pack(side="right", fill="y")

        self.listbox = tk.Listbox(list_frame, yscrollcommand=scrollbar.set, height=6)
        self.listbox.pack(side="left", fill="both", expand=True)
        scrollbar.config(command=self.listbox.yview)

        ttk.Button(self.frame, text="- Delete Selected", command=self._delete_item).pack(padx=5, pady=5)

    def _add_item(self):
        """Add item to list. Blank fields fall back to their default value."""
        values = {}
        for col in self.columns:
            raw = self.inputs[col].get().strip()
            # Use default if blank and a default exists, otherwise keep raw
            values[col] = raw or self.defaults.get(col, raw)

        # Only require non-empty for columns that have no default
        required_cols = [c for c in self.columns if c not in self.defaults]
        if all(values[c] for c in required_cols):
            self.items.append(values)
            display_text = " | ".join(values[c] for c in self.columns)
            self.listbox.insert("end", display_text)
            # Reset inputs: restore defaults for defaulted columns, clear others
            for col, var in self.inputs.items():
                var.set(self.defaults.get(col, ""))
        else:
            missing = [c for c in required_cols if not values[c]]
            messagebox.showwarning("Incomplete", f"Please fill in: {', '.join(missing)}")

    def _delete_item(self):
        """Delete selected item"""
        selection = self.listbox.curselection()
        if selection:
            idx = selection[0]
            self.listbox.delete(idx)
            del self.items[idx]

    def add_item(self, values: Dict[str, str]):
        """Programmatically add an item without going through the input widgets."""
        row = {col: values.get(col, self.defaults.get(col, "")) for col in self.columns}
        self.items.append(row)
        display_text = " | ".join(row[c] for c in self.columns)
        self.listbox.insert("end", display_text)

    def get_items(self) -> List[Dict[str, str]]:
        """Return list of items"""
        return self.items

    def clear(self):
        """Clear stored items and UI listbox."""
        self.items.clear()
        self.listbox.delete(0, "end")
        for col, var in self.inputs.items():
            var.set(self.defaults.get(col, ""))   # restore defaults on clear

    def pack(self, *args, **kwargs):
        self.frame.pack(*args, **kwargs)