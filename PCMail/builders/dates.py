"""
Hulpfuncties voor datumcalculatie van handels- en uitgiftedatums.
"""
from datetime import date, timedelta


def add_business_days(start: date, days: int) -> date:
    """Verhoogt `start` met het opgegeven aantal werkdagen (maandag–vrijdag)."""
    current = start
    added = 0
    while added < days:
        current += timedelta(days=1)
        if current.weekday() < 5:
            added += 1
    return current


def build_trade_and_issue_dates() -> tuple[date, date]:
    """Geeft (trade_date=today, issue_date=trade_date + 5 werkdagen) terug."""
    trade_date = date.today()
    issue_date = add_business_days(trade_date, 5)
    return trade_date, issue_date
