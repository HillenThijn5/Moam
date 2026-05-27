from tkinter import ttk, messagebox
import tkinter as tk
from typing import List, Dict, Tuple
from gui.widgets.underlying_search import UnderlyingSearchEntry


class DynamicListManager:
    """Beheert het toevoegen en verwijderen van items met optionele keuzelijst- of zoekselectie."""

    def __init__(self, parent, label_text: str, columns: Tuple[str, ...] = ("Item",),
                 dropdown_options: Dict[str, List[str]] = None,
                 search_options: Dict[str, List[str]] = None,
                 defaults: Dict[str, str] = None):
        """
        Args:
            parent: bovenliggende widget
            label_text: label voor de lijst
            columns: tuple van kolomnamen
            dropdown_options: kolom → lijst van opties, weergegeven als readonly Combobox
            search_options:   kolom → lijst van opties, weergegeven als UnderlyingSearchEntry
                              (typen + Enter voor de beste match, hetzelfde als bij de underlying-picker)
            defaults: kolom → standaardwaarde die gebruikt wordt als een veld leeg blijft
        """
        self.frame = ttk.LabelFrame(parent, text=label_text)
        self.columns = columns
        self.dropdown_options = dropdown_options or {}
        self.search_options   = search_options   or {}
        self.defaults = defaults or {}
        self.items: List[Dict[str, str]] = []

        # Input rij
        input_frame = ttk.Frame(self.frame)
        input_frame.pack(fill="x", padx=5, pady=5)

        self.inputs = {}
        for col in columns:
            ttk.Label(input_frame, text=f"{col}:").pack(side="left", padx=2)

            if col in self.search_options:
                # Zoekinvoerveld: auto_clear=False behoudt de geselecteerde waarde in het veld
                search_entry = UnderlyingSearchEntry(
                    input_frame,
                    all_tickers=self.search_options[col],
                    on_select=lambda _: None,   # var is gedeeld — geen extra callback nodig
                    auto_clear=False,
                )
                search_entry.pack(side="left", padx=2)
                # Deel de StringVar van de zoekwidget zodat _add_item er direct van leest
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

        # Listbox met scrollbar
        list_frame = ttk.Frame(self.frame)
        list_frame.pack(fill="both", expand=True, padx=5, pady=5)

        scrollbar = ttk.Scrollbar(list_frame)
        scrollbar.pack(side="right", fill="y")

        self.listbox = tk.Listbox(list_frame, yscrollcommand=scrollbar.set, height=6)
        self.listbox.pack(side="left", fill="both", expand=True)
        scrollbar.config(command=self.listbox.yview)

        ttk.Button(self.frame, text="- Delete Selected", command=self._delete_item).pack(padx=5, pady=5)

    def _add_item(self):
        """Voeg een item toe aan de lijst. Lege velden vallen terug op hun standaardwaarde."""
        values = {}
        for col in self.columns:
            raw = self.inputs[col].get().strip()
            # Gebruik de standaardwaarde als het veld leeg is en er een bestaat, anders behoud de invoer
            values[col] = raw or self.defaults.get(col, raw)

        # Vereis alleen een niet-lege waarde voor kolommen zonder standaardwaarde
        required_cols = [c for c in self.columns if c not in self.defaults]
        if all(values[c] for c in required_cols):
            self.items.append(values)
            display_text = " | ".join(values[c] for c in self.columns)
            self.listbox.insert("end", display_text)
            # Reset invoervelden: herstel standaardwaarden voor die kolommen, wis de rest
            for col, var in self.inputs.items():
                var.set(self.defaults.get(col, ""))
        else:
            missing = [c for c in required_cols if not values[c]]
            messagebox.showwarning("Incomplete", f"Please fill in: {', '.join(missing)}")

    def _delete_item(self):
        """Verwijder geselecteerd item"""
        selection = self.listbox.curselection()
        if selection:
            idx = selection[0]
            self.listbox.delete(idx)
            del self.items[idx]

    def add_item(self, values: Dict[str, str]):
        """Voeg programmatisch een item toe zonder via de invoerwidgets te gaan."""
        row = {col: values.get(col, self.defaults.get(col, "")) for col in self.columns}
        self.items.append(row)
        display_text = " | ".join(row[c] for c in self.columns)
        self.listbox.insert("end", display_text)

    def get_items(self) -> List[Dict[str, str]]:
        """Geef lijst van items terug"""
        return self.items

    def clear(self):
        """Wis de opgeslagen items en het UI-lijstvak."""
        self.items.clear()
        self.listbox.delete(0, "end")
        for col, var in self.inputs.items():
            var.set(self.defaults.get(col, ""))   # herstel standaardwaarden bij wissen

    def pack(self, *args, **kwargs):
        self.frame.pack(*args, **kwargs)