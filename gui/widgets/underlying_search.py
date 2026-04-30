# gui/widgets/underlying_search.py
import tkinter as tk
from tkinter import ttk


class UnderlyingSearchEntry(ttk.Frame):
    """
    Search Entry with filtered Listbox popup.
    Clears itself after a selection; calls on_select(ticker).
    """

    def __init__(self, parent, all_tickers: list, on_select, auto_clear: bool = True, **kw):
        super().__init__(parent, **kw)
        self._all = all_tickers
        self._on_select = on_select
        self._auto_clear = auto_clear
        self._popup = None

        self._var = tk.StringVar()
        self._var.trace_add("write", self._on_type)

        self._entry = ttk.Entry(self, textvariable=self._var, width=18)
        self._entry.pack(side="left")
        self._entry.bind("<Return>",   self._on_enter)
        self._entry.bind("<Down>",     self._focus_popup)
        self._entry.bind("<Escape>",   lambda _: self._close_popup())
        self._entry.bind("<FocusOut>", self._on_focus_out)

    def get(self) -> str:
        return self._var.get().strip()

    def set(self, value: str):
        self._setting = True
        self._var.set(value)
        self._setting = False
        self._close_popup()

    def clear(self):
        self._var.set("")
        self._close_popup()

    def _matches(self, text: str) -> list:
        t = text.strip().upper()
        if not t:
            return []
        starts   = [x for x in self._all if x.upper().startswith(t)]
        contains = [x for x in self._all if t in x.upper() and x not in starts]
        return (starts + contains)[:12]

    def _best_match(self, text: str):
        m = self._matches(text)
        return m[0] if m else None

    def _on_type(self, *_):
        if getattr(self, "_setting", False):
            return
        matches = self._matches(self._var.get())
        if matches:
            self._show_popup(matches)
        else:
            self._close_popup()

    def _show_popup(self, items: list):
        if self._popup:
            self._popup.destroy()

        x = self._entry.winfo_rootx()
        y = self._entry.winfo_rooty() + self._entry.winfo_height()

        popup = tk.Toplevel(self)
        popup.wm_overrideredirect(True)
        popup.geometry(f"+{x}+{y}")
        popup.resizable(False, False)

        lb = tk.Listbox(popup, height=min(8, len(items)),
                        width=20, selectmode="browse",
                        font=("Segoe UI", 9))
        lb.pack()
        for item in items:
            lb.insert("end", item)
        lb.selection_set(0)

        lb.bind("<ButtonRelease-1>", lambda _: self._pick(lb))
        lb.bind("<Return>",          lambda _: self._pick(lb))
        lb.bind("<Escape>",          lambda _: self._close_popup())
        lb.bind("<FocusOut>",        self._on_focus_out)

        self._popup = popup
        self._lb    = lb

    def _close_popup(self):
        if self._popup:
            self._popup.destroy()
            self._popup = None

    def _pick(self, lb):
        sel = lb.curselection()
        if sel:
            self._select(lb.get(sel[0]))

    def _on_enter(self, _=None):
        if self._popup:
            self._pick(self._lb)
        else:
            best = self._best_match(self._var.get())
            if best:
                self._select(best)

    def _focus_popup(self, _=None):
        if self._popup:
            self._lb.focus_set()

    def _on_focus_out(self, _=None):
        self.after(100, self._check_focus)

    def _check_focus(self):
        try:
            focused = self.focus_get()
        except KeyError:
            # Combobox 'popdown' widget can briefly not exist in widget tree;
            # ignore and leave popup open — next event will close it if needed
            return
        if self._popup and focused not in (self._entry, getattr(self, "_lb", None)):
            self._close_popup()

    def _select(self, ticker: str):
        self._close_popup()
        if self._auto_clear:
            self._var.set("")
        else:
            self._var.set(ticker)
        self._on_select(ticker)
