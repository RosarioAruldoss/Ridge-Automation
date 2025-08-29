Attribute VB_Name = "NewListing"
Sub ExtractByListingStatus()

    Dim wsSource As Worksheet
    Dim wsOutput As Worksheet
    Dim listingStatus As String
    Dim lastRow As Long
    Dim rngData As Range, rngVisible As Range
    Dim shp As Shape, shpCopy As Shape
    Dim destSheetName As String

    ' Set the source sheet
    Set wsSource = ActiveSheet

    ' Ask for the Listing Status value
    listingStatus = InputBox("Enter the Listing Status to filter by (e.g., new, active, delisted):", "Filter by Listing Status")
    If Trim(listingStatus) = "" Then
        MsgBox "? No input given. Operation cancelled.", vbExclamation
        Exit Sub
    End If

    ' Find the last used row
    lastRow = wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).Row

    ' Define the full data range (assuming up to column M)
    Set rngData = wsSource.Range("A1:M" & lastRow)

    ' Remove any existing filters and apply a new one
    If wsSource.AutoFilterMode Then wsSource.AutoFilterMode = False
    rngData.AutoFilter Field:=12, Criteria1:=listingStatus

    ' Check if any data is visible after filtering
    On Error Resume Next
    Set rngVisible = rngData.SpecialCells(xlCellTypeVisible)
    On Error GoTo 0

    If rngVisible Is Nothing Then
        MsgBox "? No rows found with Listing Status = '" & listingStatus & "'.", vbExclamation
        wsSource.AutoFilterMode = False
        Exit Sub
    End If

    ' Prepare the destination sheet
    destSheetName = listingStatus & "_status"
    On Error Resume Next
    Application.DisplayAlerts = False
    Worksheets(destSheetName).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    Set wsOutput = Worksheets.Add(After:=wsSource)
    wsOutput.Name = destSheetName

    ' Copy the filtered data (including formats)
    rngVisible.Copy
    wsOutput.Range("A1").PasteSpecial xlPasteAll

    ' Copy relevant shapes (images) based on visible rows
    For Each shp In wsSource.Shapes
        If Not Intersect(shp.TopLeftCell, rngVisible) Is Nothing Then
            Set shpCopy = shp.Duplicate
            shpCopy.Cut
            wsOutput.Paste
            With wsOutput.Shapes(wsOutput.Shapes.Count)
                .Top = wsOutput.Cells(shp.TopLeftCell.Row - rngData.Row + 1, shp.TopLeftCell.Column).Top
                .Left = wsOutput.Cells(shp.TopLeftCell.Row - rngData.Row + 1, shp.TopLeftCell.Column).Left
            End With
        End If
    Next shp

    ' Clean up
    wsSource.AutoFilterMode = False
    Application.CutCopyMode = False

    MsgBox "? Rows with Listing Status = '" & listingStatus & "' extracted to sheet: '" & destSheetName & "'", vbInformation

End Sub

