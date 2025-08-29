Attribute VB_Name = "DPH_LHH_File_Creator"
Sub SplitDataByCategory()

    Dim ws As Worksheet
    Dim wb As Workbook
    Dim dphWb As Workbook, lhhWb As Workbook
    Dim lastRow As Long, headerRows As Range
    Dim cell As Range
    Dim savePath As String
    
    Set wb = ThisWorkbook
    Set ws = wb.Sheets(1) ' or use sheet name like: Set ws = wb.Sheets("Sheet1")
    
    ' Find the last used row
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' Get file save path
    savePath = wb.Path
    
    ' Copy DPH
    Set dphWb = Workbooks.Add
    ws.Rows("1:2").Copy Destination:=dphWb.Sheets(1).Rows("1:2")
    
    For Each cell In ws.Range("D3:D" & lastRow)
        If LCase(cell.Value) = "dph" Then
            cell.EntireRow.Copy Destination:=dphWb.Sheets(1).Cells(dphWb.Sheets(1).Cells(Rows.Count, 1).End(xlUp).Row + 1, 1)
        End If
    Next cell
    
    dphWb.SaveAs Filename:=savePath & "\DPH.xlsx"
    dphWb.Close SaveChanges:=False
    
    ' Copy LHH
    Set lhhWb = Workbooks.Add
    ws.Rows("1:2").Copy Destination:=lhhWb.Sheets(1).Rows("1:2")
    
    For Each cell In ws.Range("D3:D" & lastRow)
        If LCase(cell.Value) = "lhh" Then
            cell.EntireRow.Copy Destination:=lhhWb.Sheets(1).Cells(lhhWb.Sheets(1).Cells(Rows.Count, 1).End(xlUp).Row + 1, 1)
        End If
    Next cell
    
    lhhWb.SaveAs Filename:=savePath & "\LHH.xlsx"
    lhhWb.Close SaveChanges:=False

    MsgBox "Files DPH.xlsx and LHH.xlsx created successfully!", vbInformation

End Sub

