Sub ExtractListingStatusToNewWorkbook()

    Dim wsSource As Worksheet
    Dim wbNew As Workbook
    Dim wsNew As Worksheet
    Dim listingStatus As String
    Dim tabName As String
    Dim lastRow As Long, lastCol As Long
    Dim rngData As Range, rngVisible As Range
    Dim shp As Shape, shpCopy As Shape
    Dim exportPath As String, originalPath As String, originalName As String
    Dim filteredFileName As String
    Dim cell As Range

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' Set source sheet
    Set wsSource = ActiveSheet

    ' Get original file path
    originalPath = ThisWorkbook.Path
    originalName = ThisWorkbook.Name
    originalName = Replace(originalName, ".xlsm", "")

    ' Prompt for filter keyword
    listingStatus = InputBox("Enter the keyword to filter Listing Status (e.g., 'new', 'delisted'):", "Filter by Listing Status")
    If Trim(listingStatus) = "" Then
        MsgBox "❌ No input given. Operation cancelled.", vbExclamation
        Exit Sub
    End If

    ' Prompt for tab name
    tabName = InputBox("Enter the name for the output sheet/tab:", "Output Tab Name")
    If Trim(tabName) = "" Then
        MsgBox "❌ No sheet name provided. Operation cancelled.", vbExclamation
        Exit Sub
    End If

    ' Find the last row and column
    lastRow = wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).Row
    lastCol = wsSource.Cells(2, wsSource.Columns.Count).End(xlToLeft).Column

    ' Define the full range from row 2 headers
    Set rngData = wsSource.Range(wsSource.Cells(2, 1), wsSource.Cells(lastRow, lastCol))

    ' Apply filter on column 13 with partial match
    If wsSource.AutoFilterMode Then wsSource.AutoFilterMode = False
    rngData.AutoFilter Field:=13, Criteria1:="*" & listingStatus & "*"

    ' Get visible data
    On Error Resume Next
    Set rngVisible = rngData.SpecialCells(xlCellTypeVisible)
    On Error GoTo 0

    If rngVisible Is Nothing Then
        MsgBox "❌ No rows found with Listing Status containing '" & listingStatus & "'.", vbExclamation
        wsSource.AutoFilterMode = False
        Exit Sub
    End If

    ' Create a new workbook
    Set wbNew = Workbooks.Add
    Set wsNew = wbNew.Sheets(1)
    wsNew.Name = tabName

    ' Copy header row 1
    wsSource.Rows(1).Copy Destination:=wsNew.Rows(1)

    ' Copy visible filtered data from row 2 onwards
    rngVisible.Copy Destination:=wsNew.Cells(2, 1)

    ' Copy visible shapes/images
    For Each shp In wsSource.Shapes
        If Not Intersect(shp.TopLeftCell, rngVisible) Is Nothing Then
            Set shpCopy = shp.Duplicate
            shpCopy.Cut
            wsNew.Paste
            With wsNew.Shapes(wsNew.Shapes.Count)
                .Top = wsNew.Cells(shp.TopLeftCell.Row, shp.TopLeftCell.Column).Top
                .Left = wsNew.Cells(shp.TopLeftCell.Row, shp.TopLeftCell.Column).Left
            End With
        End If
    Next shp

    ' Save the new workbook in the same directory
    filteredFileName = listingStatus & "_" & originalName & ".xlsx"
    exportPath = originalPath & "\" & filteredFileName
    wbNew.SaveAs Filename:=exportPath, FileFormat:=xlOpenXMLWorkbook

    ' Cleanup
    wsSource.AutoFilterMode = False
    Application.CutCopyMode = False
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    MsgBox "✅ Filtered data saved as '" & filteredFileName & "' in:" & vbNewLine & exportPath, vbInformation

End Sub
