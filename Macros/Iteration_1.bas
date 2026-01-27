Attribute VB_Name = "Module1"
Option Explicit

Public Sub IsoscelesTriangle2_Click()
    Milestone_CreateCopies_Rename_CopyPaste
End Sub

Public Sub Milestone_CreateCopies_Rename_CopyPaste()

    Dim newNames As Variant
    newNames = Array("Samer", "Prinu", "Ramy", "Amir", "Johny", "Michel", "Rabih")

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
        
        'to remove shape
        Dim shp As Shape
    For Each shp In wsNew.Shapes
        shp.Delete
    Next shp

        wsNew.Cells.Clear

        rngToCopy.Copy
        With wsNew.Range("A1")
            .PasteSpecial Paste:=xlPasteValues
            .PasteSpecial Paste:=xlPasteFormats
        End With
        Application.CutCopyMode = False

    Next i

CleanExit:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Exit Sub

CleanFail:
    Application.CutCopyMode = False
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "Macro stopped due to error: " & Err.Description, vbExclamation, "Milestone 1"
End Sub

