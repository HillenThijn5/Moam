# gui/widgets/overrides.py
from tkinter import ttk
import tkinter as tk
from datetime import datetime, timedelta


class DateOverrideWidget:
    def __init__(
        self,
        parent,
        label: str,
        default_days_offset: int = 0,
        default_value: str | None = None,
        return_default_when_disabled: bool = False,
    ):
        self.frame = ttk.Frame(parent)

        self.enable_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            self.frame,
            text=f"Override {label}",
            variable=self.enable_var,
            command=self._update_state,
        ).pack(side="left", padx=5)

        ttk.Label(self.frame, text="Date:").pack(side="left", padx=2)

        if default_value is None:
            default_value = (datetime.today() + timedelta(days=default_days_offset)).strftime("%Y-%m-%d")

        self._default_value = default_value
        self._return_default_when_disabled = return_default_when_disabled

        self.date_var = tk.StringVar(value=self._default_value)
        ttk.Entry(self.frame, textvariable=self.date_var, width=12).pack(side="left", padx=2)

        self._update_state()

    def _update_state(self):
        state = "normal" if self.enable_var.get() else "disabled"
        for child in self.frame.winfo_children():
            if isinstance(child, ttk.Entry):
                child.config(state=state)

    def set_default(self, value: str):
        """Stel de standaardwaarde in en werk de getoonde waarde bij als de handmatige instelling uit staat."""
        self._default_value = value
        if not self.enable_var.get():
            self.date_var.set(value)

    def reset(self):
        """Schakel de handmatige instelling uit en zet terug naar de standaardwaarde."""
        self.enable_var.set(False)
        self.date_var.set(self._default_value)
        self._update_state()

    def get_value(self) -> str:
        """Geef de handmatig ingestelde datum terug als die actief is, anders de standaardwaarde (optioneel) of leeg."""
        if self.enable_var.get():
            return self.date_var.get()
        return self._default_value if self._return_default_when_disabled else ""

    def pack(self, *args, **kwargs):
        self.frame.pack(*args, **kwargs)
    """Optionele handmatige datuminstelling met selectievakje en standaardterugval."""



# ============================================================================
# HULPFUNCTIE: DenominationOverrideWidget
# ============================================================================
# ============================================================================
# HULPFUNCTIE: DenominationOverrideWidget
# ============================================================================

class DenominationOverrideWidget:
    """Optionele handmatige denominatie-instelling met selectievakje en standaardterugval."""

    def __init__(
        self,
        parent,
        label: str = "Denomination",
        default_value: str = "",
        return_default_when_disabled: bool = False,
    ):
        self.frame = ttk.Frame(parent)

        self.enable_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            self.frame,
            text=f"Override {label}",
            variable=self.enable_var,
            command=self._update_state,
        ).pack(side="left", padx=5)

        ttk.Label(self.frame, text="Value:").pack(side="left", padx=2)

        self._default_value = default_value
        self._return_default_when_disabled = return_default_when_disabled

        self.value_var = tk.StringVar(value=self._default_value)
        ttk.Entry(self.frame, textvariable=self.value_var, width=15).pack(side="left", padx=2)

        self._update_state()

    def _update_state(self):
        state = "normal" if self.enable_var.get() else "disabled"
        for child in self.frame.winfo_children():
            if isinstance(child, ttk.Entry):
                child.config(state=state)

    def set_default(self, value: str):
        """Stel de standaardwaarde in en werk de getoonde waarde bij als de handmatige instelling niet actief is."""
        self._default_value = value
        if not self.enable_var.get():
            self.value_var.set(value)

    def reset(self):
        """Schakel de handmatige instelling uit en zet terug naar de standaardwaarde."""
        self.enable_var.set(False)
        self.value_var.set(self._default_value)
        self._update_state()

    def get_value(self) -> str:
        """Geef de handmatig ingestelde waarde terug als die actief is, anders de standaardwaarde (optioneel) of leeg."""
        if self.enable_var.get():
            return self.value_var.get()
        return self._default_value if self._return_default_when_disabled else ""

    def pack(self, *args, **kwargs):
        self.frame.pack(*args, **kwargs)
# ============================================================================