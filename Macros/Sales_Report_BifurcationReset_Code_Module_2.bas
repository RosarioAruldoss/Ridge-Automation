'open code

Option Explicit

Public Sub Graphic9_Click()
    ResetWorkbook_Click
End Sub

Public Sub ResetWorkbook_Click()

    Dim wb As Workbook
    Set wb = ThisWorkbook

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    On Error GoTo CleanFail

    '1) Delete all sheets except Sheet 1
    Dim wsKeep As Worksheet
    Set wsKeep = wb.Worksheets(1)

    Dim i As Long
    For i = wb.Worksheets.Count To 1 Step -1
        If wb.Worksheets(i).Name <> wsKeep.Name Then
            wb.Worksheets(i).Delete
        End If
    Next i

    '2) Clear data from row 2 onwards in Sheet 1 (keep header row 1)
    Dim lastRow As Long, lastCol As Long
    With wsKeep
        lastRow = .Cells(.Rows.Count, "B").End(xlUp).Row
        If lastRow < 2 Then GoTo CleanExit

        lastCol = .Cells(1, .Columns.Count).End(xlToLeft).Column
        If lastCol < 1 Then lastCol = 1

        .Range(.Cells(2, 1), .Cells(lastRow, lastCol)).ClearContents
    End With

CleanExit:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Exit Sub

CleanFail:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "Reset macro stopped due to error: " & Err.Description, vbExclamation, "Reset Workbook"
End Sub

'end of code
