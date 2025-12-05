Attribute VB_Name = "Module1"
Option Explicit

' ====== PUBLIC ENTRY POINT ======
Sub CreateCategorySheets_LocknLock()
    Dim wsSrc As Worksheet
    Set wsSrc = ActiveSheet   ' or: Set wsSrc = ThisWorkbook.Worksheets("Master")
    
    CategorySheets_Build wsSrc, True, 2
    MsgBox "Category sheets created successfully.", vbInformation
End Sub


' ====== CORE PROCEDURE ======
' wsSrc         : source worksheet containing the full table
' keepImages    : True to replicate pictures for matching rows
' imageColIndex : column index where images are anchored (2 = column B as requested)
Private Sub CategorySheets_Build(wsSrc As Worksheet, keepImages As Boolean, imageColIndex As Long)
    Dim appCalc As XlCalculation
    Dim wb As Workbook
    Dim headerRow As Long, lastRow As Long, lastCol As Long
    Dim colCategory As Long, r As Long
    Dim catDict As Object ' Scripting.Dictionary (late binding)
    Dim rngHeader As Range
    Dim wsCat As Worksheet, cat As Variant
    Dim rowMap As Object ' sourceRow -> targetRow map for image placement
    Dim ok As Boolean
    
    Set wb = wsSrc.Parent
    
    ' --- performance guardrails ---
    On Error Resume Next
    appCalc = Application.Calculation
    On Error GoTo 0
    
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
        .DisplayAlerts = False
        .Calculation = xlCalculationManual
    End With
    
    On Error GoTo CleanFail
    
    ' --- detect bounds with FIXED HEADER ROW AT 9 ---
    DetectTableBoundsFixedHeader wsSrc, headerRow, lastRow, lastCol
    If headerRow = 0 Or lastRow <= headerRow Or lastCol = 0 Then
        Err.Raise vbObjectError + 510, , "Could not detect a valid table (fixed header row 9 + data)."
    End If
    
    ' find "Category" column in the fixed header row
    colCategory = FindHeaderColumn(wsSrc, headerRow, lastCol, "Category")
    If colCategory = 0 Then
        Err.Raise vbObjectError + 511, , "Header 'Category' not found in row 9."
    End If
    
    Set rngHeader = wsSrc.Range(wsSrc.Cells(headerRow, 1), wsSrc.Cells(headerRow, lastCol))
    
    ' --- build unique category list ---
    Set catDict = CreateObject("Scripting.Dictionary")
    catDict.CompareMode = 1 ' TextCompare
    
    For r = headerRow + 1 To lastRow
        Dim v As String
        v = Trim(CStr(wsSrc.Cells(r, colCategory).Value))
        If Len(v) > 0 Then
            If Not catDict.Exists(v) Then catDict.Add v, v
        End If
    Next r
    
    ' --- create/prepare each category sheet & copy rows ---
    For Each cat In catDict.Keys
        Dim safeName As String
        safeName = SanitizeSheetName(CStr(cat))
        
        Set wsCat = EnsureSheet(wb, safeName)
        wsCat.Cells.Clear
        
        ' copy header
        rngHeader.Copy wsCat.Cells(1, 1)
        wsCat.Rows(1).Font.Bold = True
        
        ' copy matching rows values + formats, build row map for pictures
        Set rowMap = CreateObject("Scripting.Dictionary")
        rowMap.CompareMode = 0
        
        CopyRowsForCategory wsSrc, wsCat, headerRow, lastRow, lastCol, colCategory, CStr(cat), rowMap
        
        ' replicate pictures from column 2 using row map
        If keepImages Then
            CopyImagesByRow wsSrc, wsCat, rowMap, imageColIndex
        End If
        
        ' housekeeping: autofit, freeze header
        With wsCat
            .Rows(1).AutoFilter
            .Cells.EntireColumn.AutoFit
            On Error Resume Next
            .Parent.Windows(1).FreezePanes = False
            .Range("A2").Select
            .Parent.Windows(1).FreezePanes = True
            On Error GoTo 0
        End With
    Next cat
    
    ok = True

CleanExit:
    With Application
        On Error Resume Next
        .Calculation = appCalc
        .ScreenUpdating = True
        .EnableEvents = True
        .DisplayAlerts = True
        On Error GoTo 0
    End With
    If ok Then Exit Sub
    Exit Sub

CleanFail:
    Dim msg As String
    msg = "Macro aborted: " & Err.Description & " (Err #" & Err.Number & ")"
    MsgBox msg, vbExclamation, "Category Sheet Builder"
    GoTo CleanExit
End Sub


' ====== HELPERS ======

' Fixed header row at 9. Derive last row/col from UsedRange with safeguards.
Private Sub DetectTableBoundsFixedHeader(ws As Worksheet, ByRef headerRow As Long, ByRef lastRow As Long, ByRef lastCol As Long)
    Dim ur As Range
    headerRow = 9
    Set ur = ws.UsedRange
    
    ' last data row
    lastRow = ur.Row + ur.Rows.Count - 1
    If lastRow < headerRow + 1 Then
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    End If
    
    ' last column based on header row content
    lastCol = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column
    If lastCol < 1 Then
        lastCol = ur.Column + ur.Columns.Count - 1
    End If
End Sub

' Find a header named targetName in the given header row (case-insensitive)
Private Function FindHeaderColumn(ws As Worksheet, headerRow As Long, lastCol As Long, targetName As String) As Long
    Dim c As Long, txt As String
    FindHeaderColumn = 0
    For c = 1 To lastCol
        txt = Trim$(CStr(ws.Cells(headerRow, c).Value))
        If Len(txt) > 0 Then
            If StrComp(txt, targetName, vbTextCompare) = 0 Then
                FindHeaderColumn = c
                Exit Function
            End If
        End If
    Next c
End Function

' Ensure a sheet exists with given name (re-use if present)
Private Function EnsureSheet(wb As Workbook, sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = sheetName
    End If
    Set EnsureSheet = ws
End Function

' Make sheet-safe name (<=31 chars & invalid chars removed)
Private Function SanitizeSheetName(ByVal s As String) As String
    Dim badChars As Variant, ch As Variant
    badChars = Array(":", "\", "/", "?", "*", "[", "]")
    For Each ch In badChars
        s = Replace$(s, ch, " ")
    Next ch
    s = Trim$(s)
    If Len(s) = 0 Then s = "Category"
    If Len(s) > 31 Then s = Left$(s, 31)
    SanitizeSheetName = s
End Function

' Copy rows where Category equals the given value.
' Builds rowMap: sourceRow -> targetRow (for picture placement).
Private Sub CopyRowsForCategory(wsSrc As Worksheet, wsDst As Worksheet, _
                                headerRow As Long, lastRow As Long, lastCol As Long, _
                                colCategory As Long, categoryValue As String, _
                                ByRef rowMap As Object)
    Dim r As Long, dstRow As Long, rngSrc As Range, rngDst As Range
    dstRow = 2 ' after header
    
    For r = headerRow + 1 To lastRow
        If StrComp(Trim$(CStr(wsSrc.Cells(r, colCategory).Value)), categoryValue, vbTextCompare) = 0 Then
            Set rngSrc = wsSrc.Range(wsSrc.Cells(r, 1), wsSrc.Cells(r, lastCol))
            Set rngDst = wsDst.Range(wsDst.Cells(dstRow, 1), wsDst.Cells(dstRow, lastCol))
            ' Copy values + formats
            rngDst.Value = rngSrc.Value
            rngSrc.Copy
            rngDst.PasteSpecial xlPasteFormats
            Application.CutCopyMode = False
            rowMap(CStr(r)) = dstRow
            dstRow = dstRow + 1
        End If
    Next r
End Sub

' Replicate images from a specific source column (imageColIndex) for mapped rows.
' Places each picture inside the corresponding cell and sets Placement to Move&Size.
Private Sub CopyImagesByRow(wsSrc As Worksheet, wsDst As Worksheet, rowMap As Object, imageColIndex As Long)
    Dim shp As Shape, srcRow As Long, dstRow As Long
    Dim tgtCell As Range
    Dim newShp As Shape
    Dim pad As Single, maxW As Single, maxH As Single
    
    pad = 2 ' small padding
    
    For Each shp In wsSrc.Shapes
        If shp.Type = msoPicture Then
            On Error Resume Next
            srcRow = shp.TopLeftCell.Row
            If Err.Number <> 0 Then
                Err.Clear
                On Error GoTo 0
            Else
                On Error GoTo 0
                If shp.TopLeftCell.Column = imageColIndex Then
                    If rowMap.Exists(CStr(srcRow)) Then
                        dstRow = CLng(rowMap(CStr(srcRow)))
                        Set tgtCell = wsDst.Cells(dstRow, imageColIndex)
                        
                        shp.Copy
                        wsDst.Paste
                        Set newShp = wsDst.Shapes(wsDst.Shapes.Count)
                        
                        newShp.Left = tgtCell.Left + pad
                        newShp.Top = tgtCell.Top + pad
                        newShp.LockAspectRatio = msoTrue
                        newShp.Placement = xlMoveAndSize
                        
                        maxW = tgtCell.Width - 2 * pad
                        maxH = tgtCell.Height - 2 * pad
                        If newShp.Width > maxW Then newShp.Width = maxW
                        If newShp.Height > maxH Then newShp.Height = maxH
                    End If
                End If
            End If
        End If
    Next shp
End Sub


