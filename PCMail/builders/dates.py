"""
Date calculation utilities for trade and issue dates.
"""
from datetime import date, timedelta


def add_business_days(start: date, days: int) -> date:
    """Advances `start` by the given number of business days (Monday–Friday)."""
    current = start
    added = 0
    while added < days:
        current += timedelta(days=1)
        if current.weekday() < 5:
            added += 1
    return current


def build_trade_and_issue_dates() -> tuple[date, date]:
    """Returns (trade_date=today, issue_date=trade_date + 5 business days)."""
    trade_date = date.today()
    issue_date = add_business_days(trade_date, 5)
    return trade_date, issue_date
