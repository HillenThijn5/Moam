# increase_decrease_mail/db.py
"""
Vraagt de StructuredProducts2010-database uit voor huidige MTN-posities
(Format=13, Book=2) met hun laatste TransferPrice.

Geef direction="increase" mee voor shortposities (Position < 0).
Geef direction="decrease" mee voor longposities (Position > 0).
"""
import pyodbc
from statics.data import ID_MAIL_DB_CONN_STR

_QUERY_TEMPLATE = """
WITH Pos AS (
    SELECT
        II.[Name],
        II.ISIN,
        II.InstrumentIdentifier,
        SUM(P.Position) AS Position
    FROM InstrumentInfo II
    LEFT JOIN InstrumentPositions P
        ON II.InstrumentId = P.InstrumentId
       AND P.Book = 2
    WHERE II.[Format] = 13
    GROUP BY
        II.[Name],
        II.ISIN,
        II.InstrumentIdentifier
)
SELECT
    Pos.[Name],
    Pos.ISIN,
    Pos.Position,
    TP1.TransferPrice
FROM Pos
OUTER APPLY (
    SELECT TOP (1)
        TP.TransferPrice
    FROM TransferPricing TP
    WHERE Pos.InstrumentIdentifier = (TP.InstrumentIdentifier - 1)
    ORDER BY TP.EntryDate DESC
) TP1
WHERE Pos.Position {sign} 0;
"""


def fetch_positions(direction: str = "increase") -> list[dict]:
    """
    Geeft een lijst met dicts terug met de sleutels: Name, ISIN, Position, TransferPrice.

    direction="increase"  → shortposities (Position < 0)
    direction="decrease"  → longposities (Position > 0)

    Gooit pyodbc.Error bij een fout in de connectie of query.
    """
    if direction not in ("increase", "decrease"):
        raise ValueError(f"direction must be 'increase' or 'decrease', got '{direction}'")

    sign = "<" if direction == "increase" else ">"
    query = _QUERY_TEMPLATE.format(sign=sign)

    try:
        conn = pyodbc.connect(ID_MAIL_DB_CONN_STR, timeout=10)
    except pyodbc.Error as e:
        raise ConnectionError(
            f"Could not connect to StructuredProducts2010 database: {e}"
        ) from e

    try:
        cursor = conn.cursor()
        cursor.execute(query)
        columns = [col[0] for col in cursor.description]
        rows = [dict(zip(columns, row)) for row in cursor.fetchall()]
        return rows
    finally:
        conn.close()

