this is my query : WITH Pos AS (
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


this is my connection: import pyodbc

CONN_STR = (
    "DRIVER={ODBC Driver 17 for SQL Server};"
    "SERVER=sql_structuredproducts.fvlprod.fvl;"
    "DATABASE=StructuredProducts2010;"
    "Trusted_Connection=yes;"
    "Encrypt=yes;"
    "TrustServerCertificate=yes;"
)

column names are: Name, ISIN, Position, TransferPrice.









this was my original vba which i used. : Sub Mail_Sheet_Outlook_Body()

    Dim rng As Range
    Dim OutApp As Object
    Dim OutMail As Object
    Dim signature As String
    Dim bodyHtml As String

    With Application
        .EnableEvents = False
        .ScreenUpdating = False
    End With

    Set rng = Sheets("IncreaseDecrease").Range("B1:C22")

    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)

    On Error Resume Next
    With OutMail
        .To = "mtnprogramme@vanlanschot.com; financialaccounting@vanlanschot.com; treasuryadm@vanlanschot.com"
        .CC = "structuredinvestments@kempen.com; Riskmanagement@vanlanschotkempen.com; SAS@kempen.nl; ReconOs@kempen.nl"
        .BCC = ""
        .Subject = Sheets("IncreaseDecrease").Range("E1") & " " & Sheets("IncreaseDecrease").Range("C10")

        'First create and display email so Outlook loads signature
        .Display

        'Read the signature now that Outlook has inserted it
        signature = .htmlBody

        'Your custom HTML content above the signature
        bodyHtml = RangetoHTML(rng)

        'Insert custom content above signature
        .htmlBody = bodyHtml & signature
    End With
    On Error GoTo 0

    With Application
        .EnableEvents = True
        .ScreenUpdating = True
    End With

    Set OutMail = Nothing
    Set OutApp = Nothing
End Sub

Function RangetoHTML(rng As Range)
    Dim fso As Object
    Dim ts As Object
    Dim tempFile As String
    Dim tempWB As Workbook

    tempFile = Environ$("temp") & "\" & Format(Now, "dd-mm-yy h-mm-ss") & ".htm"

    'Copy the range and create a new workbook to past the data in
    rng.Copy
    Set tempWB = Workbooks.Add(1)
    With tempWB.Sheets(1)
        .Cells(1).PasteSpecial Paste:=8
        .Cells(1).PasteSpecial xlPasteValues, , False, False
        .Cells(1).PasteSpecial xlPasteFormats, , False, False
        .Cells(1).Select
        Application.CutCopyMode = False
        On Error Resume Next
        .DrawingObjects.Visible = True
        .DrawingObjects.Delete
        On Error GoTo 0
    End With

    'Publish the sheet to a htm file
    With tempWB.PublishObjects.Add( _
         SourceType:=xlSourceRange, _
         Filename:=tempFile, _
         Sheet:=tempWB.Sheets(1).Name, _
         Source:=tempWB.Sheets(1).UsedRange.Address, _
         HtmlType:=xlHtmlStatic)
        .Publish (True)
    End With

    'Read all data from the htm file into RangetoHTML
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.GetFile(tempFile).OpenAsTextStream(1, -2)
    RangetoHTML = ts.ReadAll
    ts.Close
    RangetoHTML = Replace(RangetoHTML, "align=center x:publishsource=", _
                          "align=left x:publishsource=")

    'Close TempWB
    tempWB.Close SaveChanges:=False

    'Delete the htm file we used in this function
    Kill tempFile

    Set ts = Nothing
    Set fso = Nothing
    Set tempWB = Nothing
End Function


