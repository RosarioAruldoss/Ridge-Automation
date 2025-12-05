Option Explicit

' Adjust this if needed
Const SRC_HEADER_ROW As Long = 1        ' header in original manual report
Const SRC_DATA_FIRST_ROW As Long = 3    ' data in original manual report starts at row 3

Const MASTER_NAME As String = "Master"

Sub Build_All_Reports_From_Master()
    Dim wsSrc As Worksheet
    Dim wsMaster As Worksheet
    Dim ok As Boolean

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    On Error GoTo CleanFail

    Set wsSrc = ActiveSheet   ' manual report sheet is the active one

    ' Step 1: build filtered Master sheet based on POD date
    Set wsMaster = CreateFilteredMaster(wsSrc)
    If wsMaster Is Nothing Then GoTo CleanExit

    ' Check key columns for blanks / N/A / UNASSIGNED
    ok = ValidateKeyColumns(wsMaster, Array("Sales Loc", "Country", "Team", "Section", "Group"))
    If Not ok Then
        Application.DisplayAlerts = False
        wsMaster.Delete
        Application.DisplayAlerts = True
        GoTo CleanExit
    End If

    ' Now generate the 7 reports from Master

    ' Common list of columns to remove
    Dim colsToRemove As Variant
    colsToRemove = Array("Costed", "Unit Cost", "GP", "GP %", "Workweek", _
                         "Total Item Cost", "GP Value", "GP %")

    ' Step 2: Samer  (Sales Loc: UAE, Country: UAE, Team: all except Amir...)
    CreateReportFromMaster wsMaster, "Samer", _
        salesLoc:="UAE", country:="UAE", _
        teamInclude:=Empty, teamExclude:="Amir Hossien Khaksar", _
        groupFilter:=Empty, sectionInclude:=Empty, sectionExclude:=Empty, _
        removeCols:=colsToRemove

    ' Step 3: Prinu  (Sales Loc: UAE, Country: UAE, Team: Prinu)
    CreateReportFromMaster wsMaster, "Prinu", _
        salesLoc:="UAE", country:="UAE", _
        teamInclude:="Prinu", teamExclude:=Empty, _
        groupFilter:=Empty, sectionInclude:=Empty, sectionExclude:=Empty, _
        removeCols:=colsToRemove

    ' Step 4: Ramy  (Team: Ramy OR Heba, Section: HHH)
    CreateReportFromMaster wsMaster, "Ramy", _
        salesLoc:=Empty, country:=Empty, _
        teamInclude:=Array("Ramy", "Heba"), teamExclude:=Empty, _
        groupFilter:=Empty, sectionInclude:="HHH", sectionExclude:=Empty, _
        removeCols:=colsToRemove

    ' Step 5: Amir  (Team: Amir Hossien Khaksar)
    CreateReportFromMaster wsMaster, "Amir", _
        salesLoc:=Empty, country:=Empty, _
        teamInclude:="Amir Hossien Khaksar", teamExclude:=Empty, _
        groupFilter:=Empty, sectionInclude:=Empty, sectionExclude:=Empty, _
        removeCols:=colsToRemove

    ' Step 6: John  (Sales Loc: UAE, Country: UAE, Group: Online, Section: all except HHH)
    CreateReportFromMaster wsMaster, "John", _
        salesLoc:="UAE", country:="UAE", _
        teamInclude:=Empty, teamExclude:=Empty, _
        groupFilter:="Online", sectionInclude:=Empty, sectionExclude:="HHH", _
        removeCols:=colsToRemove

    ' Step 7: Michel (Sales Loc: Prime) – keep all columns
    CreateReportFromMaster wsMaster, "Michel", _
        salesLoc:="Prime", country:=Empty, _
        teamInclude:=Empty, teamExclude:=Empty, _
        groupFilter:=Empty, sectionInclude:=Empty, sectionExclude:=Empty, _
        removeCols:=Empty

    ' Step 8: Rabih  (Sales Loc: Oman)
    CreateReportFromMaster wsMaster, "Rabih", _
        salesLoc:="Oman", country:=Empty, _
        teamInclude:=Empty, teamExclude:=Empty, _
        groupFilter:=Empty, sectionInclude:=Empty, sectionExclude:=Empty, _
        removeCols:=colsToRemove

    MsgBox "All reports generated successfully.", vbInformation

CleanExit:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub

CleanFail:
    MsgBox "Error: " & Err.Description, vbCritical
    Resume CleanExit
End Sub

' ===========================================================
' Step 1 – Build Master sheet from POD date filter
' ===========================================================
Private Function CreateFilteredMaster(wsSrc As Worksheet) As Worksheet
    Dim podCol As Long, lastRow As Long, lastCol As Long
    Dim r As Long
    Dim sampleCell As Range
    Dim prompt As String
    Dim sFrom As String, sTo As String, sBlanks As String
    Dim dFrom As Date, dTo As Date
    Dim includeBlanks As Boolean
    Dim wsOut As Worksheet
    Dim destRow As Long

    ' Find POD On column in header row
    lastCol = wsSrc.Cells(SRC_HEADER_ROW, wsSrc.Columns.Count).End(xlToLeft).Column
    podCol = 0
    For r = 1 To lastCol
        If LCase$(Trim$(wsSrc.Cells(SRC_HEADER_ROW, r).Value)) = LCase$("POD On") Then
            podCol = r
            Exit For
        End If
    Next r

    If podCol = 0 Then
        MsgBox "'POD On' column not found in header row.", vbCritical
        Exit Function
    End If

    ' Get a sample date format from first nonblank POD cell
    lastRow = wsSrc.Cells(wsSrc.Rows.Count, 1).End(xlUp).Row
    For r = SRC_DATA_FIRST_ROW To lastRow
        If Not IsEmpty(wsSrc.Cells(r, podCol).Value) Then
            Set sampleCell = wsSrc.Cells(r, podCol)
            Exit For
        End If
    Next r

    If sampleCell Is Nothing Then
        MsgBox "No POD dates found to read format from.", vbCritical
        Exit Function
    End If

    prompt = "Enter From POD date in format like this example: " & sampleCell.Text

    sFrom = InputBox(prompt, "POD Date Filter - From")
    If sFrom = "" Then Exit Function

    sTo = InputBox("Enter To POD date in same format. Example: " & sampleCell.Text, _
                   "POD Date Filter - To")
    If sTo = "" Then Exit Function

    On Error GoTo InvalidDate
    dFrom = CDate(sFrom)
    dTo = CDate(sTo)
    On Error GoTo 0

    If dTo < dFrom Then
        MsgBox "To date is earlier than From date.", vbExclamation
        Exit Function
    End If

    sBlanks = InputBox("Include rows where 'POD On' is blank? (Y/N)", _
                       "Include Blanks", "N")
    sBlanks = UCase$(Trim$(sBlanks))
    includeBlanks = (sBlanks = "Y" Or sBlanks = "YES")

    ' Create or replace Master sheet
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets(MASTER_NAME).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    Set wsOut = ThisWorkbook.Worksheets.Add( _
                    After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    wsOut.Name = MASTER_NAME

    ' Copy header row from source to Master row 1
    wsSrc.Rows(SRC_HEADER_ROW).Copy wsOut.Rows(1)

    ' Copy rows that match date criteria
    destRow = 2
    For r = SRC_DATA_FIRST_ROW To lastRow
        Dim v As Variant
        Dim keepRow As Boolean
        v = wsSrc.Cells(r, podCol).Value

        If IsDate(v) Then
            keepRow = (CDate(v) >= dFrom And CDate(v) <= dTo)
        ElseIf includeBlanks And Trim$(CStr(v)) = "" Then
            keepRow = True
        Else
            keepRow = False
        End If

        If keepRow Then
            wsSrc.Range(wsSrc.Cells(r, 1), wsSrc.Cells(r, lastCol)).Copy _
                Destination:=wsOut.Cells(destRow, 1)
            destRow = destRow + 1
        End If
    Next r

    wsOut.Cells.EntireColumn.AutoFit

    Set CreateFilteredMaster = wsOut
    Exit Function

InvalidDate:
    MsgBox "Invalid date entered. Please follow the example shown.", vbCritical
End Function

' ===========================================================
' Validate key lookup columns in Master sheet
' ===========================================================
Private Function ValidateKeyColumns(ws As Worksheet, keyCols As Variant) As Boolean
    Dim lastRow As Long, lastCol As Long
    Dim badCols As String
    Dim i As Long, c As Long, r As Long
    Dim val As String

    ValidateKeyColumns = True

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    For i = LBound(keyCols) To UBound(keyCols)
        c = FindHeaderColumn(ws, keyCols(i))
        If c = 0 Then
            badCols = badCols & vbCrLf & keyCols(i) & " (header not found)"
            ValidateKeyColumns = False
        Else
            For r = 2 To lastRow
                val = UCase$(Trim$(CStr(ws.Cells(r, c).Value)))
                If val = "" Or val = "N/A" Or val = "NA" Or val = "UNASSIGNED" Then
                    badCols = badCols & vbCrLf & keyCols(i)
                    ValidateKeyColumns = False
                    Exit For
                End If
            Next r
        End If
    Next i

    If Not ValidateKeyColumns Then
        MsgBox "The following key columns contain blanks or N/A / UNASSIGNED:" & _
               vbCrLf & badCols & vbCrLf & vbCrLf & _
               "Fix the data and run the macro again.", vbCritical
    End If
End Function

' ===========================================================
' Generic report creator from Master
' ===========================================================
Private Sub CreateReportFromMaster( _
    wsMaster As Worksheet, _
    ByVal reportName As String, _
    ByVal salesLoc As Variant, _
    ByVal country As Variant, _
    ByVal teamInclude As Variant, _
    ByVal teamExclude As Variant, _
    ByVal groupFilter As Variant, _
    ByVal sectionInclude As Variant, _
    ByVal sectionExclude As Variant, _
    ByVal removeCols As Variant)

    Dim rng As Range
    Dim lastRow As Long, lastCol As Long
    Dim colSalesLoc As Long, colCountry As Long
    Dim colTeam As Long, colGroup As Long, colSection As Long
    Dim wsOut As Worksheet
    Dim dataBody As Range, visibleData As Range

    lastRow = wsMaster.Cells(wsMaster.Rows.Count, 1).End(xlUp).Row
    lastCol = wsMaster.Cells(1, wsMaster.Columns.Count).End(xlToLeft).Column

    Set rng = wsMaster.Range(wsMaster.Cells(1, 1), wsMaster.Cells(lastRow, lastCol))

    ' Turn off old filter
    If wsMaster.AutoFilterMode Then wsMaster.AutoFilterMode = False

    colSalesLoc = FindHeaderColumn(wsMaster, "Sales Loc")
    colCountry = FindHeaderColumn(wsMaster, "Country")
    colTeam = FindHeaderColumn(wsMaster, "Team")
    colGroup = FindHeaderColumn(wsMaster, "Group")
    colSection = FindHeaderColumn(wsMaster, "Section")

    ' Apply filters only where criteria provided

    If Not IsEmpty(salesLoc) And colSalesLoc > 0 Then
        rng.AutoFilter Field:=colSalesLoc, Criteria1:=salesLoc
    End If

    If Not IsEmpty(country) And colCountry > 0 Then
        rng.AutoFilter Field:=colCountry, Criteria1:=country
    End If

    If Not IsEmpty(teamInclude) And colTeam > 0 Then
        If IsArray(teamInclude) Then
            ' only handle 2 values here (Ramy & Heba)
            rng.AutoFilter Field:=colTeam, _
                Criteria1:=teamInclude(LBound(teamInclude)), _
                Operator:=xlOr, _
                Criteria2:=teamInclude(UBound(teamInclude))
        Else
            rng.AutoFilter Field:=colTeam, Criteria1:=teamInclude
        End If
    End If

    If Not IsEmpty(teamExclude) And colTeam > 0 Then
        rng.AutoFilter Field:=colTeam, Criteria1:="<>" & CStr(teamExclude)
    End If

    If Not IsEmpty(groupFilter) And colGroup > 0 Then
        rng.AutoFilter Field:=colGroup, Criteria1:=groupFilter
    End If

    If Not IsEmpty(sectionInclude) And colSection > 0 Then
        rng.AutoFilter Field:=colSection, Criteria1:=sectionInclude
    End If

    If Not IsEmpty(sectionExclude) And colSection > 0 Then
        rng.AutoFilter Field:=colSection, Criteria1:="<>" & CStr(sectionExclude)
    End If

    ' Create or replace report sheet
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets(reportName).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    Set wsOut = ThisWorkbook.Worksheets.Add( _
                    After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    wsOut.Name = reportName

    ' Copy header row
    wsMaster.Rows(1).Copy wsOut.Rows(1)

    ' Copy visible data rows from Master (rows 2 down)
    Set dataBody = wsMaster.Range(wsMaster.Cells(2, 1), wsMaster.Cells(lastRow, lastCol))

    On Error Resume Next
    Set visibleData = dataBody.SpecialCells(xlCellTypeVisible)
    On Error GoTo 0

    If Not visibleData Is Nothing Then
        visibleData.Copy wsOut.Cells(2, 1)
    End If

    ' Remove unwanted columns by header name
    If Not IsEmpty(removeCols) Then
        RemoveColumnsByHeader wsOut, removeCols
    End If

    wsOut.Cells.EntireColumn.AutoFit

    ' Clear filter on master for the next report
    If wsMaster.AutoFilterMode Then wsMaster.AutoFilterMode = False
End Sub

' ===========================================================
' Helper: find header column index by name (row 1)
' ===========================================================
Private Function FindHeaderColumn(ws As Worksheet, headerName As String) As Long
    Dim lastCol As Long, c As Long
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    For c = 1 To lastCol
        If LCase$(Trim$(ws.Cells(1, c).Value)) = LCase$(headerName) Then
            FindHeaderColumn = c
            Exit Function
        End If
    Next c
    FindHeaderColumn = 0
End Function

' ===========================================================
' Helper: delete columns from a sheet by header name list
' ===========================================================
Private Sub RemoveColumnsByHeader(ws As Worksheet, headers As Variant)
    Dim i As Long
    Dim colIndex As Long

    For i = LBound(headers) To UBound(headers)
        colIndex = FindHeaderColumn(ws, headers(i))
        If colIndex > 0 Then
            ws.Columns(colIndex).Delete
        End If
    Next i
End Sub
