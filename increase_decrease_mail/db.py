# increase_decrease_mail/db.py
"""
Queries the StructuredProducts2010 database for current short positions
(Format=13, Book=2) with their latest TransferPrice.
"""
import pyodbc
from statics.data import ID_MAIL_DB_CONN_STR

_QUERY = """
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
WHERE Pos.Position < 0;
"""


def fetch_positions() -> list[dict]:
    """
    Returns a list of dicts with keys: Name, ISIN, Position, TransferPrice.
    Raises pyodbc.Error on connection / query failure.
    """
    conn = pyodbc.connect(ID_MAIL_DB_CONN_STR, timeout=10)
    try:
        cursor = conn.cursor()
        cursor.execute(_QUERY)
        columns = [col[0] for col in cursor.description]
        rows = [dict(zip(columns, row)) for row in cursor.fetchall()]
        return rows
    finally:
        conn.close()
