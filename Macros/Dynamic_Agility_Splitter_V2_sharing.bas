Attribute VB_Name = "Module1"
Option Explicit

Sub Dynamic_Account_Splitting()

    Dim wsL As Worksheet, wsM As Worksheet, wsO As Worksheet
    Dim lastRowM As Long, lastColM As Long, outRow As Long
    Dim headerList As Collection
    Dim hdrCols As Object
    Dim allocDict As Object
    Dim valueSplitCols As Collection
    Dim merchCols As Collection, execCols As Collection, supCols As Collection
    Dim dataArr As Variant, outArr() As Variant
    Dim r As Long, c As Long, i As Long
    Dim j As Variant
    Dim acctNum As String, key As String
    Dim allocRows As Collection
    Dim arrRow As Variant
    Dim allocPct As Double
    Dim acctColRel As Long
    Dim maxSplits As Long
    
    Const HDR_ROW As Long = 1
    Const DATA_START_ROW As Long = 2
    Const START_COL As Long = 2   ' Col B
    
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    Set wsL = ThisWorkbook.Worksheets("LOOKUP")
    Set wsM = ThisWorkbook.Worksheets("MAINDATA")
    
    
    ' 1. VALIDATE MAINDATA HEADERS VS LOOKUP
   
    Set headerList = GetHeaderList(wsL)
    If headerList.Count = 0 Then
        MsgBox "LOOKUP sheet header list is empty.", vbCritical
        GoTo SafeExit
    End If
    
    lastColM = wsM.Cells(HDR_ROW, wsM.Columns.Count).End(xlToLeft).Column
    
    If (lastColM - START_COL + 1) <> headerList.Count Then
        MsgBox "Headers are not matching." & vbCrLf & _
               "MAINDATA header count = " & (lastColM - START_COL + 1) & vbCrLf & _
               "LOOKUP header count = " & headerList.Count, vbCritical
        GoTo SafeExit
    End If
    
    For c = START_COL To lastColM
        If NormalizeText(CStr(wsM.Cells(HDR_ROW, c).Value)) <> NormalizeText(CStr(headerList(c - START_COL + 1))) Then
            MsgBox "Headers are not matching." & vbCrLf & _
                   "Mismatch at MAINDATA column " & c & vbCrLf & _
                   "MAINDATA = " & wsM.Cells(HDR_ROW, c).Value & vbCrLf & _
                   "LOOKUP = " & headerList(c - START_COL + 1), vbCritical
            GoTo SafeExit
        End If
    Next c
    
    
    ' 2. BUILD HEADER MAP FOR MAINDATA
    ' (supports duplicate headers)
    
    Set hdrCols = BuildHeaderCollections(wsM, HDR_ROW, START_COL, lastColM)
    
    If Not hdrCols.Exists(NormalizeText("Acct Num")) Then
        MsgBox "'Acct Num' header not found in MAINDATA.", vbCritical
        GoTo SafeExit
    End If
    
    acctColRel = CLng(hdrCols(NormalizeText("Acct Num")).Item(1))
    
    
    ' 3. READ LOOKUP CONFIG
    
    Set valueSplitCols = ResolveColumns_AllMatches(hdrCols, SplitCSV(GetNeighborValue(wsL, "Value_Split")))
    
    ' For text fields:
    ' Merch / Exec / Sup labels are on the right side
    ' Their mapped output headers are in the cell to the left
    Set merchCols = ResolveColumns_TextMapping(hdrCols, SplitCSV(GetLeftNeighborValue(wsL, "Merch")))
    Set execCols = ResolveColumns_TextMapping(hdrCols, SplitCSV(GetLeftNeighborValue(wsL, "Exec")))
    Set supCols = ResolveColumns_TextMapping(hdrCols, SplitCSV(GetLeftNeighborValue(wsL, "Sup")))
    
    
    ' 4. BUILD ALLOCATION DICTIONARY
    ' key = Acct Num only - t
    
    Set allocDict = BuildAllocationDictionary(wsL)
    If allocDict.Count = 0 Then
        MsgBox "No allocation data found in LOOKUP sheet.", vbCritical
        GoTo SafeExit
    End If
    
    maxSplits = GetMaxSplitCount(allocDict)
    If maxSplits = 0 Then
        MsgBox "No active allocation rows found in LOOKUP sheet.", vbCritical
        GoTo SafeExit
    End If
    
    
    ' 5. LOAD MAINDATA
    
    lastRowM = wsM.Cells(wsM.Rows.Count, START_COL).End(xlUp).Row
    If lastRowM < DATA_START_ROW Then
        MsgBox "No data found in MAINDATA.", vbExclamation
        GoTo SafeExit
    End If
    
    dataArr = wsM.Range(wsM.Cells(HDR_ROW, START_COL), wsM.Cells(lastRowM, lastColM)).Value
    
    
    ' 6. PREP OUTPUT SHEET
    
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets("SPLIT_RESULT").Delete
    Application.DisplayAlerts = True
    On Error GoTo ErrorHandler
    
    Set wsO = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    wsO.Name = "SPLIT_RESULT"
    
    ReDim outArr(1 To ((lastRowM - DATA_START_ROW + 1) * maxSplits) + 1, 1 To UBound(dataArr, 2))
    
    ' Copy headers
    For c = 1 To UBound(dataArr, 2)
        outArr(1, c) = dataArr(1, c)
    Next c
    
    outRow = 2
    
    
    ' 7. PROCESS EACH MAINDATA ROW
    
    For r = 2 To UBound(dataArr, 1)
        
        acctNum = Trim(CStr(dataArr(r, acctColRel)))
        key = NormalizeText(acctNum)
        
        If allocDict.Exists(key) Then
            
            Set allocRows = allocDict(key)
            
            For i = 1 To allocRows.Count
                
                arrRow = allocRows(i)
                allocPct = CDbl(arrRow(4))
                
                ' Step 1: clone original row n no of times
                
                For c = 1 To UBound(dataArr, 2)
                    outArr(outRow, c) = dataArr(r, c)
                Next c
                
                ' Step 2: split numeric columns from the reference list
                
                For Each j In valueSplitCols
                    If IsNumericEx(dataArr(r, CLng(j))) Then
                        outArr(outRow, CLng(j)) = CDbl(CleanNumber(dataArr(r, CLng(j)))) * allocPct
                    Else
                        outArr(outRow, CLng(j)) = dataArr(r, CLng(j))
                    End If
                Next j
                
                ' Step 3: override text columns from the Merch List
                ' Merch bucket
                For Each j In merchCols
                    outArr(outRow, CLng(j)) = arrRow(2)   ' Merch Name
                Next j
                
                ' Exec bucket
                For Each j In execCols
                    outArr(outRow, CLng(j)) = arrRow(1)   ' Exec Name
                Next j
                
                ' Sup bucket
                For Each j In supCols
                    outArr(outRow, CLng(j)) = arrRow(3)   ' Sup Name
                Next j
                
                outRow = outRow + 1
            Next i
            
        Else
            ' No matching allocation found, keep original row once for original unchanged list
            For c = 1 To UBound(dataArr, 2)
                outArr(outRow, c) = dataArr(r, c)
            Next c
            outRow = outRow + 1
        End If
    Next r
    
    
    ' 8. WRITE OUTPUT
    
    wsO.Range(wsO.Cells(1, 1), wsO.Cells(outRow - 1, UBound(dataArr, 2))).Value = outArr
    wsO.Rows(1).Font.Bold = True
    wsO.Cells.EntireColumn.AutoFit
    
    MsgBox "Dynamic account splitting  for Agility completed successfully.", vbInformation

SafeExit:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Exit Sub

ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
    Resume SafeExit

End Sub


' GET HEADER VALIDATION LIST FROM LOOKUP COLUMN A STARTING A2

Private Function GetHeaderList(ws As Worksheet) As Collection
    Dim col As New Collection
    Dim r As Long, lastRow As Long
    Dim txt As String
    
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    For r = 2 To lastRow
        txt = Trim(CStr(ws.Cells(r, 1).Value))
        If txt <> "" Then col.Add txt
    Next r
    
    Set GetHeaderList = col
End Function


' BUILD HEADER COLLECTIONS
' key = normalized header
' value = Collection of relative columns to consider

Private Function BuildHeaderCollections(ws As Worksheet, headerRow As Long, startCol As Long, lastCol As Long) As Object
    Dim d As Object
    Dim c As Long, k As String
    Dim col As Collection
    
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    
    For c = startCol To lastCol
        k = NormalizeText(CStr(ws.Cells(headerRow, c).Value))
        
        If Not d.Exists(k) Then
            Set col = New Collection
            d.Add k, col
        End If
        
        d(k).Add c - startCol + 1
    Next c
    
    Set BuildHeaderCollections = d
End Function


' GET VALUE FROM IMMEDIATE RIGHT OF LABEL

Private Function GetNeighborValue(ws As Worksheet, labelText As String) As String
    Dim f As Range
    
    Set f = ws.Cells.Find(What:=labelText, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
    If f Is Nothing Then
        Err.Raise vbObjectError + 1001, , "Label '" & labelText & "' not found in LOOKUP sheet."
    End If
    
    GetNeighborValue = Trim(CStr(f.Offset(0, 1).Value))
End Function


' GET VALUE FROM IMMEDIATE LEFT OF LABEL
' (Merch / Exec / Sup mapping in your implemented layout)

Private Function GetLeftNeighborValue(ws As Worksheet, labelText As String) As String
    Dim f As Range
    
    Set f = ws.Cells.Find(What:=labelText, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
    If f Is Nothing Then
        Err.Raise vbObjectError + 1002, , "Label '" & labelText & "' not found in LOOKUP sheet."
    End If
    
    GetLeftNeighborValue = Trim(CStr(f.Offset(0, -1).Value))
End Function


' SPLIT CSV STRING INTO ARRAY (from cg)

Private Function SplitCSV(ByVal txt As String) As Variant
    Dim arr As Variant
    Dim i As Long
    
    arr = Split(txt, ",")
    For i = LBound(arr) To UBound(arr)
        arr(i) = Trim(CStr(arr(i)))
    Next i
    
    SplitCSV = arr
End Function


' RESOLVE ALL MATCHING COLUMNS FOR VALUE SPLIT

Private Function ResolveColumns_AllMatches(hdrCols As Object, arr As Variant) As Collection
    Dim result As New Collection
    Dim i As Long, nm As String
    Dim pos As Variant
    
    For i = LBound(arr) To UBound(arr)
        nm = NormalizeText(CStr(arr(i)))
        If nm <> "" Then
            If hdrCols.Exists(nm) Then
                For Each pos In hdrCols(nm)
                    result.Add CLng(pos)
                Next pos
            End If
        End If
    Next i
    
    Set ResolveColumns_AllMatches = result
End Function


' RESOLVE TEXT COLUMNS
' Special handling:
' For duplicate "Salesrep", update only the LAST occurrence
' because first Salesrep is often numeric code and second is name

Private Function ResolveColumns_TextMapping(hdrCols As Object, arr As Variant) As Collection
    Dim result As New Collection
    Dim i As Long, nm As String
    Dim cnt As Long
    
    For i = LBound(arr) To UBound(arr)
        nm = NormalizeText(CStr(arr(i)))
        If nm <> "" Then
            If hdrCols.Exists(nm) Then
                cnt = hdrCols(nm).Count
                
                If nm = NormalizeText("Salesrep") Then
                    result.Add CLng(hdrCols(nm).Item(cnt))   ' last Salesrep only
                Else
                    result.Add CLng(hdrCols(nm).Item(1))     ' first occurrence
                End If
            End If
        End If
    Next i
    
    Set ResolveColumns_TextMapping = result
End Function


' BUILD ALLOCATION DICTIONARY
' key = Acct Num only
' value = Collection of arrays:
'   (0)=Split Seq
'   (1)=Exec Name
'   (2)=Merch Name
'   (3)=Sup Name
'   (4)=AllocDecimal

Private Function BuildAllocationDictionary(ws As Worksheet) As Object

    Dim d As Object
    Dim acctCol As Long, seqCol As Long, execCol As Long
    Dim merchCol As Long, supCol As Long, allocCol As Long, activeCol As Long
    Dim lastCol As Long, lastRow As Long
    Dim c As Long, r As Long
    Dim hdr As String
    Dim acctNum As String, execNm As String, merchNm As String, supNm As String
    Dim allocText As String, activeVal As String, key As String
    Dim seqNo As Long, allocPct As Double
    Dim col As Collection
    
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    For c = 1 To lastCol
        hdr = NormalizeText(CStr(ws.Cells(1, c).Value))
        Select Case hdr
            Case NormalizeText("Acct Num"), NormalizeText("Acct No"), NormalizeText("Acct_No")
                acctCol = c
            Case NormalizeText("Split Seq")
                seqCol = c
            Case NormalizeText("Exec Name")
                execCol = c
            Case NormalizeText("Merch Name")
                merchCol = c
            Case NormalizeText("Sup Name")
                supCol = c
            Case NormalizeText("Alloc %"), NormalizeText("Allocation%")
                allocCol = c
            Case NormalizeText("Active")
                activeCol = c
        End Select
    Next c
    
    If acctCol = 0 Or seqCol = 0 Or execCol = 0 Or merchCol = 0 Or supCol = 0 Or allocCol = 0 Or activeCol = 0 Then
        Err.Raise vbObjectError + 2001, , _
            "Allocation table headers not found properly in LOOKUP row 1." & vbCrLf & _
            "Required: Acct Num, Split Seq, Exec Name, Merch Name, Sup Name, Alloc %, Active"
    End If
    
    lastRow = ws.Cells(ws.Rows.Count, acctCol).End(xlUp).Row
    
    For r = 2 To lastRow
        
        acctNum = Trim(CStr(ws.Cells(r, acctCol).Value))
        activeVal = UCase(Trim(CStr(ws.Cells(r, activeCol).Value)))
        
        If acctNum <> "" And activeVal = "Y" Then
            
            seqNo = CLng(Val(ws.Cells(r, seqCol).Value))
            execNm = Trim(CStr(ws.Cells(r, execCol).Value))
            merchNm = Trim(CStr(ws.Cells(r, merchCol).Value))
            supNm = Trim(CStr(ws.Cells(r, supCol).Value))
            allocText = Trim(CStr(ws.Cells(r, allocCol).Value))
            allocPct = PercentToDecimal(allocText)
            key = NormalizeText(acctNum)
            
            If Not d.Exists(key) Then
                Set col = New Collection
                d.Add key, col
            End If
            
            d(key).Add Array(seqNo, execNm, merchNm, supNm, allocPct)
        End If
    Next r
    
    ' Sort each account's split rows by Split Seq
    Dim dictKey As Variant
    For Each dictKey In d.Keys
        Set d(dictKey) = SortAllocationCollection(d(dictKey))
    Next dictKey
    
    Set BuildAllocationDictionary = d
End Function


' SORT ALLOCATION COLLECTION BY SPLIT SEQ

Private Function SortAllocationCollection(colIn As Collection) As Collection
    Dim arr() As Variant
    Dim i As Long, j As Long
    Dim temp As Variant
    Dim result As New Collection
    
    ReDim arr(1 To colIn.Count)
    
    For i = 1 To colIn.Count
        arr(i) = colIn(i)
    Next i
    
    For i = 1 To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If CLng(arr(i)(0)) > CLng(arr(j)(0)) Then
                temp = arr(i)
                arr(i) = arr(j)
                arr(j) = temp
            End If
        Next j
    Next i
    
    For i = 1 To UBound(arr)
        result.Add arr(i)
    Next i
    
    Set SortAllocationCollection = result
End Function


' GET MAX SPLIT COUNT ACROSS ACCOUNTS

Private Function GetMaxSplitCount(d As Object) As Long
    Dim k As Variant
    Dim mx As Long
    
    mx = 0
    For Each k In d.Keys
        If d(k).Count > mx Then mx = d(k).Count
    Next k
    
    GetMaxSplitCount = mx
End Function


' NORMALIZE TEXT (scope - Change to the original format)

Private Function NormalizeText(ByVal txt As String) As String
    txt = Replace(txt, vbCr, "")
    txt = Replace(txt, vbLf, "")
    txt = Trim(txt)
    NormalizeText = LCase(txt)
End Function


' CONVERT PERCENT TO DECIMAL (scope, check if could be removed after format implementation)
' supports 55 or 55.00 or 55%

Private Function PercentToDecimal(ByVal txt As String) As Double
    Dim v As Double
    
    txt = Replace(txt, "%", "")
    txt = Replace(txt, ",", "")
    txt = Trim(txt)
    
    If txt = "" Then
        PercentToDecimal = 0
    Else
        v = CDbl(txt)
        If v > 1 Then
            PercentToDecimal = v / 100
        Else
            PercentToDecimal = v
        End If
    End If
End Function


' CLEAN NUMERIC TEXT

Private Function CleanNumber(ByVal v As Variant) As Double
    Dim txt As String
    txt = Trim(CStr(v))
    txt = Replace(txt, ",", "")
    If txt = "" Then
        CleanNumber = 0
    Else
        CleanNumber = CDbl(txt)
    End If
End Function


' TEST NUMERIC AFTER CLEANING

Private Function IsNumericEx(ByVal v As Variant) As Boolean
    Dim txt As String
    txt = Trim(CStr(v))
    txt = Replace(txt, ",", "")
    IsNumericEx = IsNumeric(txt)
End Function

