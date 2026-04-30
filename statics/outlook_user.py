# statics/outlook_user.py
import win32com.client as win32

_cached_first_name: str | None = None


def get_sender_first_name() -> str:
    """
    Returns the first name of the currently logged-in Outlook user.
    E.g. 'Thijn Villain' → 'Thijn'. Cached after first call.
    Falls back to empty string if Outlook is unavailable.
    """
    global _cached_first_name
    if _cached_first_name is not None:
        return _cached_first_name
    try:
        outlook = win32.Dispatch("Outlook.Application")
        display_name = outlook.Session.CurrentUser.Name
        _cached_first_name = display_name.split()[0] if display_name else ""
    except Exception:
        _cached_first_name = ""
    return _cached_first_name
