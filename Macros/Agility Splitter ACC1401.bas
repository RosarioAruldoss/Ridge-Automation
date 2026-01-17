Attribute VB_Name = "Module1"

Sub SplitSalesData()

    ' turning off the screen updates for improved performance of the macro

    Application.ScreenUpdating = False 'screen updating off
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    On Error GoTo ErrorHandler ' If there's any error, it will jump to the ErrorHandler section

    Dim ws As Worksheet
    Dim LastRow As Long
    Dim DataRange As Range
    Dim Copy1 As Range, Copy2 As Range, Copy3 As Range, Copy4 As Range
    Dim i As Long
    Dim pasteCols As Variant
    pasteCols = Array(35, 39, 41, 42, 45) ' AI, AM, AO, AP, AS

    ' Set the worksheet - CHANGE "Sheet1" TO YOUR ACTUAL SHEET NAME
    Set ws = ThisWorkbook.Worksheets("Sheet1")
    ws.Activate
    ws.Cells(1, 1).Select ' Select a safe starting cell

    ' Find the last row with data in Column A (assuming headers are in row 1)
    LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    ' Exit if only header exists
    If LastRow < 2 Then
        MsgBox "No data found to process.", vbExclamation
        GoTo SafeExit
    End If
    
    ' Define the original data range (from row 2 to last row.)
    Set DataRange = ws.Range("A2", ws.Cells(LastRow, "CB"))
    
    ' Copy the data 4 times, inserting two blank rows after each copy
    DataRange.Copy
    DataRange.Offset(DataRange.Rows.Count + 1, 0).Insert Shift:=xlDown
    Set Copy1 = DataRange.Offset(DataRange.Rows.Count + 1, 0).Resize(DataRange.Rows.Count, DataRange.Columns.Count)
    
    DataRange.Copy
    Copy1.Offset(Copy1.Rows.Count + 2, 0).Insert Shift:=xlDown
    Set Copy2 = Copy1.Offset(Copy1.Rows.Count + 2, 0).Resize(DataRange.Rows.Count, DataRange.Columns.Count)
    
    DataRange.Copy
    Copy2.Offset(Copy2.Rows.Count + 2, 0).Insert Shift:=xlDown
    Set Copy3 = Copy2.Offset(Copy2.Rows.Count + 2, 0).Resize(DataRange.Rows.Count, DataRange.Columns.Count)
    
    DataRange.Copy
    Copy3.Offset(Copy3.Rows.Count + 2, 0).Insert Shift:=xlDown
    Set Copy4 = Copy3.Offset(Copy3.Rows.Count + 2, 0).Resize(DataRange.Rows.Count, DataRange.Columns.Count)
    
    ' 1. Assign names to Column P (16) for each copy.
    Copy1.Columns(16).Value = "Mohamed Saleh"
    Copy2.Columns(16).Value = "Siraje Zeidan"
    Copy3.Columns(16).Value = "Anees Babu Kozhisseri"
    Copy4.Columns(16).Value = "Mahammed Rafik"
    
    ' 2. Copy Column P from each copy to BN, BO, BP, BQ
    Copy1.Columns(16).Copy Destination:=Copy1.Columns(66)
    Copy2.Columns(16).Copy Destination:=Copy2.Columns(66)
    Copy3.Columns(16).Copy Destination:=Copy3.Columns(66)
    Copy4.Columns(16).Copy Destination:=Copy4.Columns(66)
    Copy1.Columns(16).Copy Destination:=Copy1.Columns(67)
    Copy2.Columns(16).Copy Destination:=Copy2.Columns(67)
    Copy3.Columns(16).Copy Destination:=Copy3.Columns(67)
    Copy4.Columns(16).Copy Destination:=Copy4.Columns(67)
    Copy1.Columns(16).Copy Destination:=Copy1.Columns(68)
    Copy2.Columns(16).Copy Destination:=Copy2.Columns(68)
    Copy3.Columns(16).Copy Destination:=Copy3.Columns(68)
    Copy4.Columns(16).Copy Destination:=Copy4.Columns(68)

       
    ' 3. Change "Mahammed Rafik" & "Anees Babu Kozhisseri" in Column BP to "Prinu Raju"
    Copy4.Columns(69).Copy Destination:=Copy3.Columns(68)
    Copy4.Columns(69).Copy Destination:=Copy4.Columns(68)

    
    ' 4. Apply different formulas to Column AG for each copy based on the original AG value
' The Formula property with relative references will auto-adjust for each row when applied to the entire range.
With Copy1.Columns(33)
    .Formula = "=AG2*0.55" ' This relative reference will change to AG3, AG4, etc., for each row
End With
With Copy2.Columns(33)
    .Formula = "=AG2*0.35"
End With
With Copy3.Columns(33)
    .Formula = "=AG2*0.05"
End With
With Copy4.Columns(33)
    .Formula = "=AG2*0.05"
End With
    
    ' 5. Copy the formula from Column AG to other columns (AI,AM,AO,AP,AS) using the same R1C1 logic
For i = LBound(pasteCols) To UBound(pasteCols)
    ' For Copy1: Apply the SAME R1C1 formula from AG to the new column
    Copy1.Columns(pasteCols(i)).FormulaR1C1 = Copy1.Columns(33).FormulaR1C1
    ' Repeat for Copy2
    Copy2.Columns(pasteCols(i)).FormulaR1C1 = Copy2.Columns(33).FormulaR1C1
    ' Repeat for Copy3
    Copy3.Columns(pasteCols(i)).FormulaR1C1 = Copy3.Columns(33).FormulaR1C1
    ' Repeat for Copy4
    Copy4.Columns(pasteCols(i)).FormulaR1C1 = Copy4.Columns(33).FormulaR1C1
Next i

' Clear the clipboard after the paste operations to avoid any future paste errors
Application.CutCopyMode = False
    
    ' 6a. For Column BD, set the formula to be the same as Column AG (using R1C1 to ensure dynamic behavior)
    Copy1.Columns(56).FormulaR1C1 = Copy1.Columns(33).FormulaR1C1
    Copy2.Columns(56).FormulaR1C1 = Copy2.Columns(33).FormulaR1C1
    Copy3.Columns(56).FormulaR1C1 = Copy3.Columns(33).FormulaR1C1
    Copy4.Columns(56).FormulaR1C1 = Copy4.Columns(33).FormulaR1C1
    
    ' 6b. Copy BC to BJ and BK by setting the same dynamic formula
    ' For BJ (Column 62)
    Copy1.Columns(62).FormulaR1C1 = Copy1.Columns(56).FormulaR1C1
    Copy2.Columns(62).FormulaR1C1 = Copy2.Columns(56).FormulaR1C1
    Copy3.Columns(62).FormulaR1C1 = Copy3.Columns(56).FormulaR1C1
    Copy4.Columns(62).FormulaR1C1 = Copy4.Columns(56).FormulaR1C1

    ' For BK (Column 63)
    Copy1.Columns(63).FormulaR1C1 = Copy1.Columns(56).FormulaR1C1
    Copy2.Columns(63).FormulaR1C1 = Copy2.Columns(56).FormulaR1C1
    Copy3.Columns(63).FormulaR1C1 = Copy3.Columns(56).FormulaR1C1
    Copy4.Columns(63).FormulaR1C1 = Copy4.Columns(56).FormulaR1C1
    
    ' 7. Convert each copy to values INDIVIDUALLY to avoid the "multiple selections" error
    Copy1.Copy
    Copy1.PasteSpecial Paste:=xlPasteValues
    Copy2.Copy
    Copy2.PasteSpecial Paste:=xlPasteValues
    Copy3.Copy
    Copy3.PasteSpecial Paste:=xlPasteValues
    Copy4.Copy
    Copy4.PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    
    ' 8. Remove the blank lines between the copies
    ' We must find the new positions after all the pasting is done
    Dim FirstBlankRow As Long
    FirstBlankRow = Copy1.Row + Copy1.Rows.Count
    ws.Rows(FirstBlankRow & ":" & FirstBlankRow + 1).Delete Shift:=xlUp
    
    FirstBlankRow = Copy2.Row + Copy2.Rows.Count ' Position has shifted after first delete
    ws.Rows(FirstBlankRow & ":" & FirstBlankRow + 1).Delete Shift:=xlUp
    
    FirstBlankRow = Copy3.Row + Copy3.Rows.Count ' Position has shifted again
    ws.Rows(FirstBlankRow & ":" & FirstBlankRow + 1).Delete Shift:=xlUp
    
    MsgBox "Data splitting and formatting completed successfully!", vbInformation

SafeExit:
    ' Restore application settings
    Application.CutCopyMode = False
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description & vbNewLine & "Please ensure your data is formatted correctly and try again.", vbCritical
    Resume SafeExit
End Sub
