Attribute VB_Name = "data_extraction_v3"
Sub ExtractAndFormatFromImageColumns()
    Dim wsFrom As Worksheet
    Dim wsTo As Worksheet
    Dim wb As Workbook
    Dim fromSheetName As String
    Dim lastRow As Long
    Dim i As Long

    Set wb = ThisWorkbook

    ' Step 1: Ask for sheet name
    fromSheetName = InputBox("Enter the sheet name to extract data from:")

    On Error Resume Next
    Set wsFrom = wb.Sheets(fromSheetName)
    On Error GoTo 0

    If wsFrom Is Nothing Then
        MsgBox "? Sheet '" & fromSheetName & "' does not exist.", vbCritical
        Exit Sub
    End If

    ' Step 2: Add new sheet at the start
    Set wsTo = wb.Sheets.Add(Before:=wb.Sheets(1))
    wsTo.Name = "Formatted_" & Format(Now, "hhmmss")

    ' Step 3: Add headers in new sheet
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

    ' Step 4: Determine last row based on ITEM column (Column E)
    lastRow = wsFrom.Cells(wsFrom.Rows.Count, 5).End(xlUp).Row

    ' Step 5: Copy data row-by-row
    For i = 2 To lastRow
        wsTo.Cells(i, 1).Value = wsFrom.Cells(i, 12).Value         ' Zone ? Store
        wsTo.Cells(i, 2).Value = ""                                ' Null
        wsTo.Cells(i, 3).Value = wsFrom.Cells(i, 5).Value          ' ITEM ? Customer Article
        wsTo.Cells(i, 4).Value = wsFrom.Cells(i, 6).Value          ' ITEMDSC ? Item Description
        wsTo.Cells(i, 5).Value = wsFrom.Cells(i, 7).Value          ' BRAND ? Model
        wsTo.Cells(i, 6).Formula = "=IFERROR(LEFT(D" & i & ",FIND("" "",D" & i & ")-1),D" & i & ")" ' First Word of ITMDSC
        wsTo.Cells(i, 7).Value = wsFrom.Cells(i, 13).Value         ' QTY ? Sales Qty
        wsTo.Cells(i, 8).Value = wsFrom.Cells(i, 8).Value          ' PP
        wsTo.Cells(i, 9).Value = wsFrom.Cells(i, 9).Value          ' SP
        wsTo.Cells(i, 10).Value = wsFrom.Cells(i, 10).Value        ' GV
        wsTo.Cells(i, 11).Value = wsFrom.Cells(i, 11).Value        ' Net SP
    Next i

    MsgBox "? Data extracted and formatted to sheet: " & wsTo.Name, vbInformation
End Sub

