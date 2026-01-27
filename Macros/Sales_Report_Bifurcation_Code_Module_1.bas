'Open Code

Option Explicit

'Assign Macro to click (Include Function)

Public Sub IsoscelesTriangle2_Click()
    Milestone_CreateCopies_Rename_CopyPaste 'function
End Sub

Public Sub Milestone_CreateCopies_Rename_CopyPaste()

    Dim newNames As Variant
    newNames = Array("Samer", "Prinu", "Ramy", "Amir", "Johny", "Michel", "Rabih") 'sheet names array (use if need more sheets)

    Dim wb As Workbook
    Set wb = ThisWorkbook

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    On Error GoTo CleanFail

    Dim wsSrc As Worksheet
    Set wsSrc = wb.Worksheets(1)

    Dim lastRow As Long, lastCol As Long
    lastRow = wsSrc.Cells(wsSrc.Rows.Count, "B").End(xlUp).Row

    lastCol = wsSrc.Cells(1, wsSrc.Columns.Count).End(xlToLeft).Column
    If lastCol < 2 Then lastCol = 2

    Dim rngToCopy As Range
    Set rngToCopy = wsSrc.Range(wsSrc.Cells(1, 2), wsSrc.Cells(lastRow, lastCol)) 'B1:last

    Dim i As Long
    For i = LBound(newNames) To UBound(newNames)

        Dim wsNew As Worksheet

        wsSrc.Copy After:=wb.Worksheets(wb.Worksheets.Count)
        Set wsNew = wb.Worksheets(wb.Worksheets.Count)

        wsNew.Name = newNames(i)

        'Remove shapes copied from source
        
        Dim shp As Shape
        For Each shp In wsNew.Shapes
            shp.Delete
        Next shp

        wsNew.Cells.Clear

        rngToCopy.Copy
        With wsNew.Range("A1") 'copy the range to A1 address in new sheets
            .PasteSpecial Paste:=xlPasteValues
            .PasteSpecial Paste:=xlPasteFormats
        End With
        Application.CutCopyMode = False

    Next i

    'Now apply filters in sequence and create _F sheets
    ApplyFiltersAndCreateFilteredSheets wb
    DeleteBaseSheets wb

CleanExit:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Exit Sub

CleanFail:
    Application.CutCopyMode = False
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "Macro stopped due to error: Check original data " & Err.Description, vbExclamation, "Milestone 1+Filters"
End Sub

Private Sub ApplyFiltersAndCreateFilteredSheets(ByVal wb As Workbook)
'creating new sheet - filter rule (align with the array of sheets in previous section)

    '1) Samer -> Samer_F
    FilterAndCopyToNewSheet wb, "Samer", "Samer_F", _
        Array( _
            Array("Sales Loc", Array("UAE")), _
            Array("Country", Array("UAE")), _
            Array("Team", Array("Johny Nevil", "Prinu Raju", "Ramy Hegazy")) _
        )

    '2) Prinu -> Prinu_F
    FilterAndCopyToNewSheet wb, "Prinu", "Prinu_F", _
        Array( _
            Array("Sales Loc", Array("UAE")), _
            Array("Country", Array("UAE")), _
            Array("Team", Array("Prinu Raju")) _
        )

    '3) Ramy -> Ramy_F
    FilterAndCopyToNewSheet wb, "Ramy", "Ramy_F", _
        Array( _
            Array("Team", Array("Johny Nevil", "Ramy Hegazy")), _
            Array("Section", Array("HHH")) _
        )

    '4) Amir -> Amir_F
    FilterAndCopyToNewSheet wb, "Amir", "Amir_F", _
        Array( _
            Array("Team", Array("Amir Hossein Khaksar")) _
        )

    '5) Johny -> Johny_F
    FilterAndCopyToNewSheet wb, "Johny", "Johny_F", _
        Array( _
            Array("Sales Loc", Array("UAE")), _
            Array("Country", Array("UAE")), _
            Array("Team", Array("Johny Nevil")), _
            Array("Section", Array("DPH", "LHH")) _
        )

    '6) Michel -> Michel_F
    FilterAndCopyToNewSheet wb, "Michel", "Michel_F", _
        Array( _
            Array("Sales Loc", Array("PRIME")) _
        )

    '7) Rabih -> Rabih_F
    FilterAndCopyToNewSheet wb, "Rabih", "Rabih_F", _
        Array( _
            Array("Sales Loc", Array("OMAN")) _
        )

End Sub

Private Sub FilterAndCopyToNewSheet(ByVal wb As Workbook, _
                                   ByVal sourceSheetName As String, _
                                   ByVal outputSheetName As String, _
                                   ByVal filters As Variant)

    Dim ws As Worksheet
    Set ws = wb.Worksheets(sourceSheetName)

    'Defining used range (A1 to last used row/column)
    
    Dim lastRow As Long, lastCol As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    If lastRow < 1 Or lastCol < 1 Then Exit Sub

    Dim rng As Range
    Set rng = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))

    'to reset filters
    
    If ws.AutoFilterMode Then ws.AutoFilterMode = False

    'Apply filters sequentially
    Dim i As Long
    For i = LBound(filters) To UBound(filters)

        Dim headerName As String
        headerName = CStr(filters(i)(0))

        Dim criteriaArr As Variant
        criteriaArr = filters(i)(1) 'Array of criteria values

        Dim colIndex As Long
        colIndex = GetHeaderColumnIndex(ws, headerName, lastCol)

        If colIndex = 0 Then
            MsgBox "Header not found on sheet '" & sourceSheetName & "': " & headerName, vbExclamation
            GoTo CleanupFilters
        End If

        'Apply filter
        
        If UBound(criteriaArr) = 0 Then
            rng.AutoFilter Field:=colIndex, Criteria1:=criteriaArr(0)
        Else
            rng.AutoFilter Field:=colIndex, Criteria1:=criteriaArr, Operator:=xlFilterValues
        End If
    Next i

    'Create output sheet
    
    Dim wsOut As Worksheet
    Set wsOut = CreateOrReplaceSheet(wb, outputSheetName)

    'Copy visible only (filtered) range to output sheet
    Dim vis As Range
    On Error Resume Next
    Set vis = rng.SpecialCells(xlCellTypeVisible)
    On Error GoTo 0

    wsOut.Cells.Clear

    If Not vis Is Nothing Then
        vis.Copy
        wsOut.Range("A1").PasteSpecial Paste:=xlPasteValues
        wsOut.Range("A1").PasteSpecial Paste:=xlPasteFormats
        Application.CutCopyMode = False
    Else
        'no visible rows (unlikely, but safe handling)
        wsOut.Range("A1").Value = "No data after filters for: " & sourceSheetName
    End If

CleanupFilters:
    'remove filters from source sheet
    If ws.AutoFilterMode Then ws.AutoFilterMode = False

End Sub

Private Function GetHeaderColumnIndex(ByVal ws As Worksheet, ByVal headerName As String, ByVal lastCol As Long) As Long
    'find header in row 1, match exact text
    Dim c As Long
    For c = 1 To lastCol
        If Trim$(CStr(ws.Cells(1, c).Value)) = headerName Then
            GetHeaderColumnIndex = c
            Exit Function
        End If
    Next c
    GetHeaderColumnIndex = 0
End Function

Private Function CreateOrReplaceSheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0

    If Not ws Is Nothing Then
        Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = True
    End If

    Set CreateOrReplaceSheet = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    CreateOrReplaceSheet.Name = sheetName
End Function

Private Sub DeleteBaseSheets(ByVal wb As Workbook)

    Dim sheetsToDelete As Variant
    sheetsToDelete = Array("Samer", "Prinu", "Ramy", "Amir", "Johny", "Michel", "Rabih") 'update this too when needed. maintain as the original initial array.

    Dim i As Long
    Dim ws As Worksheet

    Application.DisplayAlerts = False

    For i = LBound(sheetsToDelete) To UBound(sheetsToDelete)
        On Error Resume Next
        Set ws = wb.Worksheets(sheetsToDelete(i))
        On Error GoTo 0

        If Not ws Is Nothing Then
            ws.Delete
        End If

        Set ws = Nothing
    Next i

    Application.DisplayAlerts = True

End Sub

'end of code
