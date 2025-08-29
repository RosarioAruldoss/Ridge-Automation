Attribute VB_Name = "Module1"
Sub DeleteRowsWithBlankInColumnA()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long

    ' Set the active sheet
    Set ws = ActiveSheet

    ' Find the last used row in Column A
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' Loop from bottom to top to avoid skipping rows when deleting
    For i = lastRow To 1 Step -1
        If Trim(ws.Cells(i, "A").Value) = "" Then
            ws.Rows(i).Delete
        End If
    Next i

    MsgBox "All rows with blank cells in Column A have been deleted.", vbInformation
End Sub

