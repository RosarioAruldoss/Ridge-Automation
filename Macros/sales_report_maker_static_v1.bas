'This note explains how I want this static macros to work.

'I will click on the Blank workbook (new excel sheet) from the sales report - UAE OMAN FILE, I will manually copy paste the current month data from 
'the source file to this new workbook under the sheet name - sheet1.

'There should be a sheet with the macros enabled which will do the following tasks once I click on a button.

'NOTE : THE FIRST ROW IS THE HEADER ROW. 'step 1: REMOVE THE BLANK ROWS FROM THE DATA IN SHEET1
'THE HEADER FIELD TITLES ARE AS FOLLOWS:
'Acct Num/Cust Name/Tax Reg No/Sales Channel Code/Group Name/Cust Class/Territory/Inv Num/Trx Date/Order/Ordered Date/
'Expiration Date/Batch Code/	Source/	Salesrep/	Salesrep/	LPO Num/Cur/E.Rate/Comments/Pod On/Term Due Date/Term Desc/
'Return Ref/Rtn Reason/	Line#	/Brand Code/	Brand Desc	/Item Code	/Item Desc	/Prom Item/Prom Desc	/
'Quantity Ordered/Uom/Inv Qty/Unit SP/Disc Ups/VRate/VAT Amt/Tax Recovery/F Amt/L Amt/Costed/Unit Cost/GP/GP %/
' 'Promo/nonPromo'/Delivery No/Delivery/Target/Workweek/Customer Name/Area/
'	Item Description/'Promo/nonPromo'	/Amount/	Country	/Supplier Item Code/	Section/	Section1	/Section Parent	
'/Total Item Cost/GP Value/GP %/ Group	/Merchandiser	/Sales Exec	/Sales Sup	/Team/SalesLoc/Sales Type/'Offline/Online'
'/Day/POD Month/POD Year/POD Qtr/SubBrand/Item Category/Item Family


'STEP 2: CREATE 7 SHEETS WITH THE NAMES - SAMER, PRINU, RAMY, AMIR, JOHNNY, MICHEL, Rabih

'STEP 4: FROM THE SHEET 1, WE HAVE TO APPLY SOME FILTERS AND DELETE SOME COLUMNS AND PASTE THE DATA IN THE RESPECTIVE SHEETS AS
'MENTIONED IN THE STEP 2.

'STEP 5: THE FILTERS AND DEDUCTION FOR SHEET - 'SAMER' IS AS FOLLOWS:
    'FILTER: HEADER : 
        'SalesLoc'  VALUE : 'UAE'
        'Country'    VALUE : 'UAE'
        'Team'       VALUE : Exclude 'Amir Hossein Khaksar'
    'DELETE COLUMNS with the headers - (Costed)(Unit Cost)(GP)(GP %)(Workweek)(Total Item Cost)(GP Value)(GP %)

    'copy paste the filtered data in the sheet named - 'SAMER'

'STEP 6: THE FILTERS AND DEDUCTION FOR SHEET - 'PRINU' IS AS FOLLOWS:
    'FILTER: HEADER : 
        'SalesLoc'  VALUE : 'UAE'
        'Country'    VALUE : 'UAE'
        'Team'       VALUE : 'Prinu Raju'
    'DELETE COLUMNS with the headers - (Costed)(Unit Cost)(GP)(GP %)(Workweek)(Total Item Cost)(GP Value)(GP %)

    'copy paste the filtered data in the sheet named - 'PRINU'

'STEP 7: THE FILTERS AND DEDUCTION FOR SHEET - 'RAMY' IS AS FOLLOWS:
    'FILTER: HEADER : 
        'Team'       VALUE : 'Ramy Hegazy' & 'Heba Serhal'
        'Section'    VALUE : 'HHH'
    'DELETE COLUMNS with the headers - (Costed)(Unit Cost)(GP)(GP %)(Workweek)(Total Item Cost)(GP Value)(GP %)

    'copy paste the filtered data in the sheet named - 'RAMY'

'STEP 8: THE FILTERS AND DEDUCTION FOR SHEET - 'AMIR' IS AS FOLLOWS:
    'FILTER: HEADER : 
        'Team'       VALUE : 'Amir Hossein Khaksar'
    'DELETE COLUMNS with the headers - (Costed)(Unit Cost)(GP)(GP %)(Workweek)(Total Item Cost)(GP Value)(GP %)

    'copy paste the filtered data in the sheet named - 'AMIR'

'STEP 9: THE FILTERS AND DEDUCTION FOR SHEET - 'JOHNNY' IS AS FOLLOWS:
    'FILTER: HEADER : 
        'SalesLoc'  VALUE : 'UAE'
        'Country'    VALUE : 'UAE'
        'Group'      VALUE : 'Online'
        'Section'    VALUE : Excluding 'HHH'
    'DELETE COLUMNS with the headers - (Costed)(Unit Cost)(GP)(GP %)(Workweek)(Total Item Cost)(GP Value)(GP %)

    'copy paste the filtered data in the sheet named - 'JOHNNY'

'STEP 10: THE FILTERS AND DEDUCTION FOR SHEET - 'MICHEL' IS AS FOLLOWS:
    'FILTER: HEADER : 
        'SalesLoc'  VALUE : 'PRIME'

    'copy paste the filtered data in the sheet named - 'MICHEL'

'STEP 11: THE FILTERS AND DEDUCTION FOR SHEET - 'RABIH' IS AS FOLLOWS:
    'FILTER: HEADER : 
        'SalesLoc'  VALUE : 'OMAN'

    'DELETE COLUMNS with the headers - (Costed)(Unit Cost)(GP)(GP %)(Workweek)(Total Item Cost)(GP Value)(GP %)
    
    'copy paste the filtered data in the sheet named - 'RABIH'

'Now the process is complete and it can be terminated.

Option Explicit

' STATIC MACRO should work in one-click:
' 1) Sheet1: remove blank rows (Row 1 is header, data starts Row 2)
' 2) Create/refresh 7 sheets: SAMER, PRINU, RAMY, AMIR, JOHNNY, MICHEL, RABIH
' 3) From Sheet1 apply filters, copy result to each sheet
' 4) For SAMER/PRINU/RAMY/AMIR/JOHNNY/RABIH delete specific columns (by header)

Public Sub Build_7_Reports_From_Sheet1()
    Dim wsSrc As Worksheet
    Set wsSrc = ThisWorkbook.Worksheets("Sheet1")

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    On Error GoTo CleanFail

    ' Step 1: remove blank rows (based on Column A empties)
    RemoveBlankRows wsSrc

    ' Step 2: create/refresh sheets
    EnsureReportSheet "SAMER"
    EnsureReportSheet "PRINU"
    EnsureReportSheet "RAMY"
    EnsureReportSheet "AMIR"
    EnsureReportSheet "JOHNNY"
    EnsureReportSheet "MICHEL"
    EnsureReportSheet "RABIH"

    ' Columns to remove in most reports
    Dim colsToRemove As Variant
    colsToRemove = Array("Costed", "Unit Cost", "GP", "GP %", "Workweek", _
                         "Total Item Cost", "GP Value", "GP %")

    ' Step 5: SAMER
    CreateReport wsSrc, "SAMER", _
        salesLoc:="UAE", country:="UAE", _
        teamInclude:=Empty, teamExclude:="Amir Hossein Khaksar", _
        groupFilter:=Empty, sectionInclude:=Empty, sectionExclude:=Empty, _
        removeCols:=colsToRemove

    ' Step 6: PRINU
    CreateReport wsSrc, "PRINU", _
        salesLoc:="UAE", country:="UAE", _
        teamInclude:="Prinu Raju", teamExclude:=Empty, _
        groupFilter:=Empty, sectionInclude:=Empty, sectionExclude:=Empty, _
        removeCols:=colsToRemove

    ' Step 7: RAMY
    CreateReport wsSrc, "RAMY", _
        salesLoc:=Empty, country:=Empty, _
        teamInclude:=Array("Ramy Hegazy", "Heba Serhal"), teamExclude:=Empty, _
        groupFilter:=Empty, sectionInclude:="HHH", sectionExclude:=Empty, _
        removeCols:=colsToRemove

    ' Step 8: AMIR
    CreateReport wsSrc, "AMIR", _
        salesLoc:=Empty, country:=Empty, _
        teamInclude:="Amir Hossein Khaksar", teamExclude:=Empty, _
        groupFilter:=Empty, sectionInclude:=Empty, sectionExclude:=Empty, _
        removeCols:=colsToRemove

    ' Step 9: JOHNNY
    CreateReport wsSrc, "JOHNNY", _
        salesLoc:="UAE", country:="UAE", _
        teamInclude:=Empty, teamExclude:=Empty, _
        groupFilter:="Online", sectionInclude:=Empty, sectionExclude:="HHH", _
        removeCols:=colsToRemove

    ' Step 10: MICHEL (no column deletion)
    CreateReport wsSrc, "MICHEL", _
        salesLoc:="PRIME", country:=Empty, _
        teamInclude:=Empty, teamExclude:=Empty, _
        groupFilter:=Empty, sectionInclude:=Empty, sectionExclude:=Empty, _
        removeCols:=Empty

    ' Step 11: RABIH (no column deletion per your latest instruction)
    CreateReport wsSrc, "RABIH", _
        salesLoc:="OMAN", country:=Empty, _
        teamInclude:=Empty, teamExclude:=Empty, _
        groupFilter:=Empty, sectionInclude:=Empty, sectionExclude:=Empty, _
        removeCols:=Empty

    ' Clear filter on source
    If wsSrc.AutoFilterMode Then wsSrc.AutoFilterMode = False

    MsgBox "Process completed. All 7 reports generated.", vbInformation

CleanExit:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub

CleanFail:
    MsgBox "Error: " & Err.Description, vbCritical
    Resume CleanExit
End Sub

'===========================================================
' Step 1: Remove blank rows (entire row blank based on Col A)
'===========================================================
Private Sub RemoveBlankRows(ByVal ws As Worksheet)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then Exit Sub

    Dim r As Long
    For r = lastRow To 2 Step -1
        If Trim$(CStr(ws.Cells(r, 1).Value)) = "" Then
            ws.Rows(r).Delete
        End If
    Next r
End Sub

'===========================================================
' Step 2: Ensure report sheet exists and is cleared
'===========================================================
Private Sub EnsureReportSheet(ByVal sheetName As String)
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = sheetName
    Else
        ws.Cells.Clear
    End If
End Sub

'===========================================================
' Core: Create a report sheet by applying filters on Sheet1
'===========================================================
Private Sub CreateReport( _
    ByVal wsSrc As Worksheet, _
    ByVal reportName As String, _
    ByVal salesLoc As Variant, _
    ByVal country As Variant, _
    ByVal teamInclude As Variant, _
    ByVal teamExclude As Variant, _
    ByVal groupFilter As Variant, _
    ByVal sectionInclude As Variant, _
    ByVal sectionExclude As Variant, _
    ByVal removeCols As Variant)

    Dim wsOut As Worksheet
    Set wsOut = ThisWorkbook.Worksheets(reportName)
    wsOut.Cells.Clear

    Dim lastRow As Long, lastCol As Long
    lastRow = wsSrc.Cells(wsSrc.Rows.Count, 1).End(xlUp).Row
    lastCol = wsSrc.Cells(1, wsSrc.Columns.Count).End(xlToLeft).Column
    If lastRow < 2 Or lastCol < 1 Then Exit Sub

    Dim rng As Range
    Set rng = wsSrc.Range(wsSrc.Cells(1, 1), wsSrc.Cells(lastRow, lastCol))

    ' Clear any existing filters
    If wsSrc.AutoFilterMode Then wsSrc.AutoFilterMode = False

    ' Find needed columns (header row = 1)
    Dim colSalesLoc As Long, colCountry As Long, colTeam As Long, colGroup As Long, colSection As Long
    colSalesLoc = FindHeaderColumn(wsSrc, "Sales Loc")
    colCountry = FindHeaderColumn(wsSrc, "Country")
    colTeam = FindHeaderColumn(wsSrc, "Team")
    colGroup = FindHeaderColumn(wsSrc, "Group")
    colSection = FindHeaderColumn(wsSrc, "Section")

    ' Apply filters only when criteria provided
    If Not IsEmpty(salesLoc) And colSalesLoc > 0 Then
        rng.AutoFilter Field:=colSalesLoc, Criteria1:=salesLoc
    End If

    If Not IsEmpty(country) And colCountry > 0 Then
        rng.AutoFilter Field:=colCountry, Criteria1:=country
    End If

    If Not IsEmpty(teamInclude) And colTeam > 0 Then
        If IsArray(teamInclude) Then
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

    ' Copy header
    wsOut.Rows(1).Resize(1, lastCol).Value = wsSrc.Rows(1).Resize(1, lastCol).Value

    ' Copy visible rows (data)
    Dim dataBody As Range, visibleData As Range, area As Range
    Dim destRow As Long
    destRow = 2

    Set dataBody = wsSrc.Range(wsSrc.Cells(2, 1), wsSrc.Cells(lastRow, lastCol))

    On Error Resume Next
    Set visibleData = dataBody.SpecialCells(xlCellTypeVisible)
    On Error GoTo 0

    If Not visibleData Is Nothing Then
        For Each area In visibleData.Areas
            wsOut.Cells(destRow, 1).Resize(area.Rows.Count, area.Columns.Count).Value = area.Value
            destRow = destRow + area.Rows.Count
        Next area
    End If

    ' Remove columns if requested
    If Not IsEmpty(removeCols) Then
        RemoveColumnsByHeader wsOut, removeCols
    End If

    wsOut.Cells.EntireColumn.AutoFit

    ' Clear filter for next report
    If wsSrc.AutoFilterMode Then wsSrc.AutoFilterMode = False
End Sub

'===========================================================
' Helper: find header column by exact header name in Row 1
'===========================================================
Private Function FindHeaderColumn(ByVal ws As Worksheet, ByVal headerName As String) As Long
    Dim lastCol As Long, c As Long
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    For c = 1 To lastCol
        If LCase$(Trim$(CStr(ws.Cells(1, c).Value))) = LCase$(headerName) Then
            FindHeaderColumn = c
            Exit Function
        End If
    Next c

    FindHeaderColumn = 0
End Function

'===========================================================
' Helper: delete columns by header name list (Row 1)
' Deletes from right-to-left to avoid index shift issues
'===========================================================
Private Sub RemoveColumnsByHeader(ByVal ws As Worksheet, ByVal headers As Variant)
    Dim i As Long, colIndex As Long
    Dim toDelete As Collection
    Set toDelete = New Collection

    ' Collect all matching columns
    For i = LBound(headers) To UBound(headers)
        colIndex = FindHeaderColumn(ws, CStr(headers(i)))
        If colIndex > 0 Then toDelete.Add colIndex
    Next i

    ' Delete in descending order (right-to-left)
    Dim j As Long, k As Long
    Dim arr() As Long
    If toDelete.Count = 0 Then Exit Sub

    ReDim arr(1 To toDelete.Count)
    For j = 1 To toDelete.Count
        arr(j) = CLng(toDelete(j))
    Next j

    For j = LBound(arr) To UBound(arr) - 1
        For k = j + 1 To UBound(arr)
            If arr(j) < arr(k) Then
                Dim tmp As Long
                tmp = arr(j): arr(j) = arr(k): arr(k) = tmp
            End If
        Next k
    Next j

    For j = LBound(arr) To UBound(arr)
        ws.Columns(arr(j)).Delete
    Next j
End Sub
