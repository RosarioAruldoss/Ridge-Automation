Attribute VB_Name = "Module1"
Option Explicit

'========================
' SETTINGS
'========================
Private Const MAIN_SHEET As String = "MAIN_DATA"
Private Const LOOKUP_SHEET As String = "LOOKUP"

Private Const MAIN_HEADER_ROW As Long = 1
Private Const MAIN_FIRST_DATA_ROW As Long = 2
Private Const MAIN_FIRST_COL As Long = 2 ' B

'========================
' ENTRY POINT
'========================
Public Sub Build_Reports_From_MainData()
    Dim wsMain As Worksheet, wsL As Worksheet

    On Error Resume Next
    Set wsMain = ThisWorkbook.Worksheets(MAIN_SHEET)
    Set wsL = ThisWorkbook.Worksheets(LOOKUP_SHEET)
    On Error GoTo 0

    If wsMain Is Nothing Then
        MsgBox "Sheet not found: " & MAIN_SHEET, vbCritical
        Exit Sub
    End If
    If wsL Is Nothing Then
        MsgBox "Sheet not found: " & LOOKUP_SHEET, vbCritical
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    On Error GoTo CleanFail

    ' 1) Validate headers before doing anything
    If Not ValidateHeadersMatch(wsMain, wsL) Then GoTo CleanExit

    ' 2) Build all output sheets using filter rules
    BuildAllOutputSheets wsMain, wsL

    ' 3) Apply No_GP column deletions after sheets are created
    ApplyNoGPDeletions wsL

    MsgBox "Done. All sheets generated from MAIN_DATA using LOOKUP rules.", vbInformation

CleanExit:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub

CleanFail:
    MsgBox "Macro stopped: " & Err.Description, vbCritical
    Resume CleanExit
End Sub

'=========================================================
' 1) HEADER VALIDATION
' LOOKUP contains vertical list under header "List_of_Headers"
' Must match MAIN_DATA headers (B1 onward) EXACTLY (case-insensitive)
'=========================================================
Private Function ValidateHeadersMatch(ByVal wsMain As Worksheet, ByVal wsL As Worksheet) As Boolean
    Dim listHdrCell As Range
    Set listHdrCell = FindCellExact(wsL, "List_of_Headers")

    If listHdrCell Is Nothing Then
        MsgBox "LOOKUP is missing header: List_of_Headers", vbCritical
        ValidateHeadersMatch = False
        Exit Function
    End If

    Dim mainLastCol As Long
    mainLastCol = wsMain.Cells(MAIN_HEADER_ROW, wsMain.Columns.Count).End(xlToLeft).Column
    If mainLastCol < MAIN_FIRST_COL Then
        MsgBox "MAIN_DATA: No headers found starting at B1.", vbCritical
        ValidateHeadersMatch = False
        Exit Function
    End If

    Dim i As Long
    Dim lookupRow As Long: lookupRow = listHdrCell.Row + 1
    Dim mainCol As Long: mainCol = MAIN_FIRST_COL

    Do While Trim$(CStr(wsL.Cells(lookupRow, listHdrCell.Column).Value)) <> ""
        If mainCol > mainLastCol Then
            MsgBox "Header mismatch: LOOKUP has more headers than MAIN_DATA.", vbCritical
            ValidateHeadersMatch = False
            Exit Function
        End If

        Dim lHdr As String, mHdr As String
        lHdr = NormalizeHeader(CStr(wsL.Cells(lookupRow, listHdrCell.Column).Value))
        mHdr = NormalizeHeader(CStr(wsMain.Cells(MAIN_HEADER_ROW, mainCol).Value))

        If lHdr <> mHdr Then
            MsgBox "Header mismatch at position " & (mainCol - MAIN_FIRST_COL + 1) & vbCrLf & _
                   "LOOKUP: " & wsL.Cells(lookupRow, listHdrCell.Column).Value & vbCrLf & _
                   "MAIN_DATA: " & wsMain.Cells(MAIN_HEADER_ROW, mainCol).Value, vbCritical
            ValidateHeadersMatch = False
            Exit Function
        End If

        lookupRow = lookupRow + 1
        mainCol = mainCol + 1
    Loop

    ' Also ensure MAIN_DATA doesn't have extra headers beyond the LOOKUP list
    If mainCol <= mainLastCol Then
        MsgBox "Header mismatch: MAIN_DATA has extra headers not listed in LOOKUP List_of_Headers.", vbCritical
        ValidateHeadersMatch = False
        Exit Function
    End If

    ValidateHeadersMatch = True
End Function

Private Function NormalizeHeader(ByVal s As String) As String
    NormalizeHeader = LCase$(Trim$(Replace(s, ChrW(160), " ")))
End Function

'=========================================================
' 2) FILTER-RULE ENGINE
' Finds table with headers: OutputSheet, SourceBaseSheet, Header, Mode, Values
' Builds each OutputSheet by applying its rules to MAIN_DATA
'=========================================================
Private Sub BuildAllOutputSheets(ByVal wsMain As Worksheet, ByVal wsL As Worksheet)
    Dim ruleHdrCell As Range
    Set ruleHdrCell = FindCellExact(wsL, "OutputSheet")

    If ruleHdrCell Is Nothing Then
        MsgBox "LOOKUP is missing the filter rules header: OutputSheet", vbCritical
        Exit Sub
    End If

    Dim hdrRow As Long: hdrRow = ruleHdrCell.Row

    Dim colOut As Long, colHeader As Long, colMode As Long, colValues As Long
    colOut = ruleHdrCell.Column
    colHeader = FindHeaderInRow(wsL, hdrRow, "Header")
    colMode = FindHeaderInRow(wsL, hdrRow, "Mode")
    colValues = FindHeaderInRow(wsL, hdrRow, "Values")

    If colHeader = 0 Or colMode = 0 Or colValues = 0 Then
        MsgBox "LOOKUP filter table must contain: OutputSheet, Header, Mode, Values", vbCritical
        Exit Sub
    End If

    ' Build list of distinct OutputSheets
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim r As Long
    r = hdrRow + 1
    Do While Trim$(CStr(wsL.Cells(r, colOut).Value)) <> ""
        Dim outName As String
        outName = Trim$(CStr(wsL.Cells(r, colOut).Value))
        If Not dict.Exists(outName) Then dict.Add outName, True
        r = r + 1
    Loop

    ' For each output sheet, apply rules and create sheet
    Dim key As Variant
    For Each key In dict.Keys
        CreateOutputSheetFromRules wsMain, wsL, hdrRow, CStr(key), colOut, colHeader, colMode, colValues
    Next key
End Sub

Private Sub CreateOutputSheetFromRules( _
    ByVal wsMain As Worksheet, _
    ByVal wsL As Worksheet, _
    ByVal rulesHeaderRow As Long, _
    ByVal outputSheetName As String, _
    ByVal colOut As Long, _
    ByVal colHeader As Long, _
    ByVal colMode As Long, _
    ByVal colValues As Long)

    Dim wsOut As Worksheet
    Set wsOut = CreateOrClearSheet(outputSheetName)

    ' Determine main data range (B1:lastRow/lastCol)
    Dim lastRow As Long, lastCol As Long
    lastCol = wsMain.Cells(MAIN_HEADER_ROW, wsMain.Columns.Count).End(xlToLeft).Column
    lastRow = wsMain.Cells(wsMain.Rows.Count, MAIN_FIRST_COL).End(xlUp).Row

    If lastRow < MAIN_FIRST_DATA_ROW Then
        wsOut.Range("A1").Value = "No data in MAIN_DATA."
        Exit Sub
    End If

    Dim rngAll As Range
    Set rngAll = wsMain.Range(wsMain.Cells(MAIN_HEADER_ROW, MAIN_FIRST_COL), wsMain.Cells(lastRow, lastCol))

    ' Reset filters
    If wsMain.AutoFilterMode Then wsMain.AutoFilterMode = False

    ' Collect EXCLUDE rules separately (so we can handle multi-exclude safely)
    Dim excludeRules As Collection
    Set excludeRules = New Collection

    ' Apply INCLUDE rules via AutoFilter
    Dim r As Long
    r = rulesHeaderRow + 1
    Do While Trim$(CStr(wsL.Cells(r, colOut).Value)) <> ""
        If NormalizeHeader(CStr(wsL.Cells(r, colOut).Value)) = NormalizeHeader(outputSheetName) Then
            Dim hdrName As String, mode As String, valuesText As String
            hdrName = Trim$(CStr(wsL.Cells(r, colHeader).Value))
            mode = UCase$(Trim$(CStr(wsL.Cells(r, colMode).Value)))
            valuesText = Trim$(CStr(wsL.Cells(r, colValues).Value))

            If hdrName <> "" And mode <> "" Then
                Dim fieldIndex As Long
                fieldIndex = GetFieldIndexInRange(wsMain, rngAll, hdrName)

                If fieldIndex = 0 Then
                    MsgBox "Header not found in MAIN_DATA: " & hdrName & vbCrLf & _
                           "While building: " & outputSheetName, vbCritical
                    wsMain.AutoFilterMode = False
                    Exit Sub
                End If

                Dim arrVals As Variant
                arrVals = SplitAndTrim(valuesText)

                If mode = "INCLUDE" Then
                    ApplyIncludeFilter rngAll, fieldIndex, arrVals
                ElseIf mode = "EXCLUDE" Then
                    ' Apply simple exclude with AutoFilter only if single value; otherwise do row-removal after copy
                    If IsArray(arrVals) And UBound(arrVals) >= 0 Then
                        If UBound(arrVals) = 0 Then
                            rngAll.AutoFilter Field:=fieldIndex, Criteria1:="<>" & CStr(arrVals(0))
                        Else
                            ' store for post-copy removal
                            excludeRules.Add Array(hdrName, arrVals)
                        End If
                    End If
                Else
                    MsgBox "Invalid Mode in LOOKUP (use INCLUDE or EXCLUDE): " & mode, vbCritical
                    wsMain.AutoFilterMode = False
                    Exit Sub
                End If
            End If
        End If
        r = r + 1
    Loop

    ' Copy headers to output (as horizontal headers)
    Dim outColCount As Long
    outColCount = rngAll.Columns.Count
    wsOut.Range("A1").Resize(1, outColCount).Value = rngAll.Rows(1).Value

    ' Copy visible data (no clipboard)
    Dim dataBody As Range, visibleData As Range, area As Range
    Dim destRow As Long: destRow = 2

    Set dataBody = rngAll.Offset(1, 0).Resize(rngAll.Rows.Count - 1, rngAll.Columns.Count)

    On Error Resume Next
    Set visibleData = dataBody.SpecialCells(xlCellTypeVisible)
    On Error GoTo 0

    If Not visibleData Is Nothing Then
        For Each area In visibleData.Areas
            wsOut.Cells(destRow, 1).Resize(area.Rows.Count, area.Columns.Count).Value = area.Value
            destRow = destRow + area.Rows.Count
        Next area
    End If

    ' Clear filters for next sheet build
    If wsMain.AutoFilterMode Then wsMain.AutoFilterMode = False

    ' Apply post-copy EXCLUDE rules (multi-value excludes)
    If excludeRules.Count > 0 And wsOut.Cells(wsOut.Rows.Count, 1).End(xlUp).Row >= 2 Then
        ApplyExcludesOnOutput wsOut, excludeRules
    End If

    wsOut.Cells.EntireColumn.AutoFit
End Sub

Private Sub ApplyIncludeFilter(ByVal rngAll As Range, ByVal fieldIndex As Long, ByVal arrVals As Variant)
    If Not IsArray(arrVals) Then Exit Sub

    If UBound(arrVals) = 0 Then
        rngAll.AutoFilter Field:=fieldIndex, Criteria1:=CStr(arrVals(0))
    Else
        rngAll.AutoFilter Field:=fieldIndex, Criteria1:=arrVals, Operator:=xlFilterValues
    End If
End Sub

'=========================================================
' 3) No_GP deletions
' Table headers: No_GP, cols_delete
' Multiple rows per sheet allowed
' Deletes ALL matching header columns (including duplicate headers)
'=========================================================
Private Sub ApplyNoGPDeletions(ByVal wsL As Worksheet)
    Dim ngpHdrCell As Range
    Set ngpHdrCell = FindCellExact(wsL, "No_GP")

    If ngpHdrCell Is Nothing Then
        ' No No_GP table is not a hard error. Just skip.
        Exit Sub
    End If

    Dim hdrRow As Long: hdrRow = ngpHdrCell.Row
    Dim colSheet As Long: colSheet = ngpHdrCell.Column
    Dim colDel As Long: colDel = FindHeaderInRow(wsL, hdrRow, "cols_delete")

    If colDel = 0 Then
        MsgBox "LOOKUP No_GP table is missing header: cols_delete", vbCritical
        Exit Sub
    End If

    Dim r As Long
    r = hdrRow + 1

    Do While Trim$(CStr(wsL.Cells(r, colSheet).Value)) <> ""
        Dim sheetName As String, delList As String
        sheetName = Trim$(CStr(wsL.Cells(r, colSheet).Value))
        delList = Trim$(CStr(wsL.Cells(r, colDel).Value))

        If sheetName <> "" And delList <> "" Then
            Dim arrHeaders As Variant
            arrHeaders = SplitAndTrim(delList) ' splits by comma and trims items

            Dim i As Long
            If IsArray(arrHeaders) Then
                For i = LBound(arrHeaders) To UBound(arrHeaders)
                    If Trim$(CStr(arrHeaders(i))) <> "" Then
                        DeleteAllColumnsByHeader sheetName, CStr(arrHeaders(i))
                    End If
                Next i
            End If
        End If

        r = r + 1
    Loop
End Sub

Private Sub DeleteAllColumnsByHeader(ByVal sheetName As String, ByVal headerName As String)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then Exit Sub

    Dim lastCol As Long
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    If lastCol < 1 Then Exit Sub

    Dim target As String
    target = NormalizeHeader(headerName)

    Dim c As Long
    For c = lastCol To 1 Step -1
        If NormalizeHeader(CStr(ws.Cells(1, c).Value)) = target Then
            ws.Columns(c).Delete
        End If
    Next c
End Sub

'=========================================================
' HELPERS
'=========================================================
Private Function CreateOrClearSheet(ByVal name As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(name)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.name = name
    Else
        ws.Cells.Clear
    End If

    Set CreateOrClearSheet = ws
End Function

Private Function FindCellExact(ByVal ws As Worksheet, ByVal textToFind As String) As Range
    Dim f As Range
    Set f = ws.Cells.Find(What:=textToFind, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
    Set FindCellExact = f
End Function

Private Function FindHeaderInRow(ByVal ws As Worksheet, ByVal headerRow As Long, ByVal headerName As String) As Long
    Dim lastCol As Long, c As Long
    lastCol = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column

    For c = 1 To lastCol
        If NormalizeHeader(CStr(ws.Cells(headerRow, c).Value)) = NormalizeHeader(headerName) Then
            FindHeaderInRow = c
            Exit Function
        End If
    Next c

    FindHeaderInRow = 0
End Function

' Return Field index relative to rngAll (AutoFilter Field is 1..n in rngAll)
Private Function GetFieldIndexInRange(ByVal wsMain As Worksheet, ByVal rngAll As Range, ByVal headerName As String) As Long
    Dim c As Long
    For c = 1 To rngAll.Columns.Count
        If NormalizeHeader(CStr(rngAll.Cells(1, c).Value)) = NormalizeHeader(headerName) Then
            GetFieldIndexInRange = c
            Exit Function
        End If
    Next c
    GetFieldIndexInRange = 0
End Function

Private Function SplitAndTrim(ByVal s As String) As Variant
    Dim t As String
    t = Trim$(s)

    If t = "" Then
        SplitAndTrim = Array()
        Exit Function
    End If

    Dim parts() As String
    parts = Split(t, ",")

    Dim i As Long
    For i = LBound(parts) To UBound(parts)
        parts(i) = Trim$(parts(i))
    Next i

    SplitAndTrim = parts
End Function

' Apply multi-value excludes after data is copied (safe and predictable)
Private Sub ApplyExcludesOnOutput(ByVal wsOut As Worksheet, ByVal excludeRules As Collection)
    Dim lastRow As Long, lastCol As Long
    lastRow = wsOut.Cells(wsOut.Rows.Count, 1).End(xlUp).Row
    lastCol = wsOut.Cells(1, wsOut.Columns.Count).End(xlToLeft).Column
    If lastRow < 2 Or lastCol < 1 Then Exit Sub

    Dim rule As Variant
    For Each rule In excludeRules
        Dim hdrName As String
        Dim arrVals As Variant
        hdrName = CStr(rule(0))
        arrVals = rule(1)

        Dim colIndex As Long
        colIndex = FindHeaderIndexOnSheet(wsOut, hdrName)
        If colIndex = 0 Then GoTo NextRule

        Dim r As Long
        For r = lastRow To 2 Step -1
            Dim v As String
            v = Trim$(CStr(wsOut.Cells(r, colIndex).Value))
            If ValueInArray(v, arrVals) Then
                wsOut.Rows(r).Delete
            End If
        Next r

        ' Recalculate lastRow because rows moved
        lastRow = wsOut.Cells(wsOut.Rows.Count, 1).End(xlUp).Row

NextRule:
    Next rule
End Sub

Private Function FindHeaderIndexOnSheet(ByVal ws As Worksheet, ByVal headerName As String) As Long
    Dim lastCol As Long, c As Long
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    For c = 1 To lastCol
        If NormalizeHeader(CStr(ws.Cells(1, c).Value)) = NormalizeHeader(headerName) Then
            FindHeaderIndexOnSheet = c
            Exit Function
        End If
    Next c
    FindHeaderIndexOnSheet = 0
End Function

Private Function ValueInArray(ByVal v As String, ByVal arr As Variant) As Boolean
    If Not IsArray(arr) Then Exit Function

    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        If NormalizeHeader(v) = NormalizeHeader(CStr(arr(i))) Then
            ValueInArray = True
            Exit Function
        End If
    Next i
    ValueInArray = False
End Function

Sub Graphic4_Click()

End Sub
