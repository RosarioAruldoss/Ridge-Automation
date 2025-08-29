Attribute VB_Name = "WideToLongV2"
Sub TransformWideToLongData()

    Dim wsSource As Worksheet, wsDest As Worksheet
    Dim lastRow As Long, currentCol As Long
    Dim destRow As Long, i As Long
    Dim zoneName As String
    Dim rngQTY As Range, rngSellout As Range

    ' Set worksheets
    Set wsSource = ThisWorkbook.Sheets(1)
    On Error Resume Next
    Set wsDest = ThisWorkbook.Sheets("LongFormat")
    If wsDest Is Nothing Then
        Set wsDest = ThisWorkbook.Sheets.Add(After:=wsSource)
        wsDest.Name = "LongFormat"
    Else
        wsDest.Cells.Clear
    End If
    On Error GoTo 0

    ' Write headers
    wsDest.Range("A1:C1").Value = Array("Zone", "QTY", "Sellout")
    destRow = 2
    currentCol = 1 ' Start from column A

    ' Loop through pairs of columns
    Do While wsSource.Cells(2, currentCol).Value <> ""

        zoneName = wsSource.Cells(1, currentCol).Value

        ' Detect last row in this QTY column
        lastRow = wsSource.Cells(wsSource.Rows.Count, currentCol).End(xlUp).Row
        If lastRow < 2 Then Exit Do ' No usable data

        ' Set QTY and Sellout ranges
        Set rngQTY = wsSource.Range(wsSource.Cells(2, currentCol), wsSource.Cells(lastRow, currentCol))
        Set rngSellout = wsSource.Range(wsSource.Cells(2, currentCol + 1), wsSource.Cells(lastRow, currentCol + 1))

        ' Transfer data
        For i = 1 To rngQTY.Rows.Count
            wsDest.Cells(destRow, 1).Value = zoneName
            wsDest.Cells(destRow, 2).Value = rngQTY.Cells(i, 1).Value
            wsDest.Cells(destRow, 3).Value = rngSellout.Cells(i, 1).Value
            destRow = destRow + 1
        Next i

        currentCol = currentCol + 2 ' Move to next pair

    Loop

    MsgBox "? Wide data dynamically transformed to long format in 'LongFormat' sheet.", vbInformation

End Sub

