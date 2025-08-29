Attribute VB_Name = "data_extraction_v3"
Sub ExtractSmartMappedData()
    Dim wsFrom As Worksheet, wsTo As Worksheet
    Dim wb As Workbook
    Dim lastRow As Long, i As Long, headerRow As Range
    Dim colMap As Object
    Dim fromSheetName As String
    Dim headers As Variant, headerName As Variant
    Dim colIndex As Long

    Set wb = ThisWorkbook
    Set colMap = CreateObject("Scripting.Dictionary")

    ' Step 1: Ask for the sheet name
    fromSheetName = InputBox("Enter the sheet name to extract data from:")

    On Error Resume Next
    Set wsFrom = wb.Sheets(fromSheetName)
    On Error GoTo 0

    If wsFrom Is Nothing Then
        MsgBox "? Sheet '" & fromSheetName & "' not found.", vbCritical
        Exit Sub
    End If

    ' Step 2: Create output sheet
    Set wsTo = wb.Sheets.Add(Before:=wb.Sheets(1))
    wsTo.Name = "Smart_Extract_" & Format(Now, "hhmmss")

    ' Step 3: Identify header row (assume headers in row 1)
    Set headerRow = wsFrom.Rows(1)
    headers = Array("Zone", "Article", "Description", "Model", "QTY", "PP", "RSPV", "GV", "Net RSPV")

    For Each headerName In headers
        For colIndex = 1 To headerRow.Cells(1, wsFrom.Columns.Count).End(xlToLeft).Column
            If Trim(UCase(headerRow.Cells(1, colIndex).Value)) = UCase(headerName) Then
                colMap(headerName) = colIndex
                Exit For
            End If
        Next colIndex
        If Not colMap.exists(headerName) Then
            colMap(headerName) = -1 ' Missing column
        End If
    Next

    ' Step 4: Output headers
    With wsTo
        .Cells(1, 1).Value = "Store"
        .Cells(1, 2).Value = "Null"
        .Cells(1, 3).Value = "Customer Article"
        .Cells(1, 4).Value = "Item Description"
        .Cells(1, 5).Value = "Model"
        .Cells(1, 6).Value = "First Name (Brand)"
        .Cells(1, 7).Value = "Sales Qty"
        .Cells(1, 8).Value = "PP"
        .Cells(1, 9).Value = "SP"
        .Cells(1, 10).Value = "GV"
        .Cells(1, 11).Value = "Net SP"
    End With

    ' Step 5: Loop through rows and extract
    lastRow = wsFrom.Cells(wsFrom.Rows.Count, 1).End(xlUp).Row

    For i = 2 To lastRow
        With wsTo
            .Cells(i, 1).Value = GetCell(wsFrom, i, colMap("Zone"))       ' Store
            .Cells(i, 2).Value = ""                                       ' Null
            .Cells(i, 3).Value = GetCell(wsFrom, i, colMap("Article"))       ' Customer Article
            .Cells(i, 4).Value = GetCell(wsFrom, i, colMap("Description"))     ' Item Description
            .Cells(i, 5).Value = GetCell(wsFrom, i, colMap("Model"))     ' Model (from SUPPLR)
            .Cells(i, 6).Formula = "=IFERROR(LEFT(D" & i & ",FIND("" "",D" & i & ")-1),D" & i & ")" ' First Word of Item Description
            .Cells(i, 7).Value = GetCell(wsFrom, i, colMap("QTY"))        ' Sales Qty
            .Cells(i, 8).Value = ""         ' PP
            .Cells(i, 9).Value = GetCell(wsFrom, i, colMap("RSPV"))         ' SP
            .Cells(i, 10).Value = GetCell(wsFrom, i, colMap("GV"))        ' GV (optional)
            .Cells(i, 11).Value = GetCell(wsFrom, i, colMap("Net RSPV"))    ' Net SP (optional)
        End With
    Next i

    MsgBox "? Smart data extraction complete to sheet: " & wsTo.Name, vbInformation
End Sub

' Helper function to safely get cell value or blank
Function GetCell(ws As Worksheet, rowNum As Long, colNum As Long) As Variant
    If colNum = -1 Then
        GetCell = ""
    Else
        GetCell = ws.Cells(rowNum, colNum).Value
    End If
End Function

