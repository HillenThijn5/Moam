# increase_decrease_mail/db.py
"""
Queries the StructuredProducts2010 database for current MTN positions
(Format=13, Book=2) with their latest TransferPrice.

Pass direction="increase" to get short positions (Position < 0).
Pass direction="decrease" to get long positions  (Position > 0).
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
    Returns a list of dicts with keys: Name, ISIN, Position, TransferPrice.

    direction="increase"  → short positions (Position < 0)
    direction="decrease"  → long positions  (Position > 0)

    Raises pyodbc.Error on connection / query failure.
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

