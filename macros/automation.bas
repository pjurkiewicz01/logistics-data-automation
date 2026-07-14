Attribute VB_Name = "Module1"
Option Explicit

'==== SETTINGS ====
Private Const DATA_SHEET As String = "Data"
Private Const OUT_SHEET  As String = "Collections"
Private Const OUT_HEADER_ROW As Long = 6
Private Const SERVICE_NEEDLE As String = "collect"
Private Const SCAN_DATA_TOP As Long = 10

'================ ENTRY =================
Public Sub RunCollections_Stable()
    Dim stage As String
    On Error GoTo Fail

    Dim wsD As Worksheet, wsO As Worksheet
    Set wsD = ThisWorkbook.Worksheets(DATA_SHEET)
    Set wsO = ThisWorkbook.Worksheets(OUT_SHEET)

    '---- Output headers on row 6 (exact) ----
    stage = "Locating output headers"
    Dim oRt As Long, oDp As Long, oNm As Long
    oRt = FindColExactOnRow(wsO, OUT_HEADER_ROW, "Route")
    oDp = FindColExactOnRow(wsO, OUT_HEADER_ROW, "Drop")
    oNm = FindColExactOnRow(wsO, OUT_HEADER_ROW, "Name")
    If oRt * oDp * oNm = 0 Then
        MsgBox "Collections row 6 must have headers 'Route', 'Drop', 'Name'.", vbExclamation
        Exit Sub
    End If
    Dim firstOut As Long
    firstOut = Application.Max(8, OUT_HEADER_ROW + HeaderDepth(wsO, OUT_HEADER_ROW, Array(oRt, oDp, oNm)))
' Optional extra columns on row 6
Dim oSuc As Long, oPrn As Long, oChk As Long
oSuc = FindColExactOnRow(wsO, OUT_HEADER_ROW, "Successful")
oPrn = FindColExactOnRow(wsO, OUT_HEADER_ROW, "Printed")
oChk = FindColExactOnRow(wsO, OUT_HEADER_ROW, "Checked in")

' Rightmost column we should format (covers E,F,G when present)
Dim lastOutCol As Long
lastOutCol = oNm
If oSuc > 0 Then lastOutCol = Application.Max(lastOutCol, oSuc)
If oPrn > 0 Then lastOutCol = Application.Max(lastOutCol, oPrn)
If oChk > 0 Then lastOutCol = Application.Max(lastOutCol, oChk)
    '---- If Data is completely empty -> clear output and exit ----
    stage = "Checking Data emptiness"
    If WorksheetFunction.CountA(wsD.Cells) = 0 Then
        ClearDownColumn_MergeSafe wsO, oRt, firstOut
        ClearDownColumn_MergeSafe wsO, oDp, firstOut
        ClearDownColumn_MergeSafe wsO, oNm, firstOut
        MsgBox "No data in 'Data'. Collections cleared under headers.", vbInformation
        Exit Sub
    End If

    '---- Data headers (exact) ----
    stage = "Locating Data headers"
    Dim rHdr As Long, cSvc As Long, cRt As Long, cDp As Long, cNm As Long
    rHdr = FindDataHeaders(wsD, SCAN_DATA_TOP, cSvc, cRt, cDp, cNm)
    If rHdr = 0 Then
        ClearDownColumn_MergeSafe wsO, oRt, firstOut
        ClearDownColumn_MergeSafe wsO, oDp, firstOut
        ClearDownColumn_MergeSafe wsO, oNm, firstOut
        SafeClearAll wsD
        MsgBox "Couldn't find Data headers: Service / Third Party Route Number / Drop Number / Surname. Cleared both areas.", vbInformation
        Exit Sub
    End If

    '---- Last data row (safe now: cSvc > 0) ----
    stage = "Finding last data row"
    Dim lastRowD As Long
    lastRowD = wsD.Cells(wsD.Rows.Count, cSvc).End(xlUp).Row
    If lastRowD <= rHdr Then
        ClearDownColumn_MergeSafe wsO, oRt, firstOut
        ClearDownColumn_MergeSafe wsO, oDp, firstOut
        ClearDownColumn_MergeSafe wsO, oNm, firstOut
        SafeClearAll wsD
        MsgBox "Data sheet has headers but no rows. Cleared Collections and Data.", vbInformation
        Exit Sub
    End If

    '---- Count matches ----
    stage = "Counting matches"
    Dim r As Long, n As Long
    For r = rHdr + 1 To lastRowD
        If InStr(1, LCase$(Trim$(wsD.Cells(r, cSvc).Text)), SERVICE_NEEDLE, vbTextCompare) > 0 Then n = n + 1
    Next r

    ' Always clear output first
    stage = "Clearing output"
    ClearDownColumn_MergeSafe wsO, oRt, firstOut
    ClearDownColumn_MergeSafe wsO, oDp, firstOut
    ClearDownColumn_MergeSafe wsO, oNm, firstOut

    If n = 0 Then
        SafeClearAll wsD
        MsgBox "No rows where Service contains 'collect'. Collections cleared; Data wiped.", vbInformation
        Exit Sub
    End If

    '---- Build arrays ----
    stage = "Building arrays"
    Dim arrRt() As Variant, arrDp() As Variant, arrNm() As Variant
    ReDim arrRt(1 To n): ReDim arrDp(1 To n): ReDim arrNm(1 To n)

    Dim i As Long: i = 0
    For r = rHdr + 1 To lastRowD
        If InStr(1, LCase$(Trim$(wsD.Cells(r, cSvc).Text)), SERVICE_NEEDLE, vbTextCompare) > 0 Then
            i = i + 1
            arrRt(i) = NeatText(wsD.Cells(r, cRt).Value)
            arrDp(i) = NeatText(wsD.Cells(r, cDp).Value)
            Dim nm As String: nm = Trim$(wsD.Cells(r, cNm).Text)
            If nm = "" Then nm = "(no surname)"
            arrNm(i) = nm
        End If
    Next r

    '---- Sort: Route then Drop ----
    stage = "Sorting arrays"
    Sort2 arrRt, arrDp, arrNm

    '---- Write under row 6 ----
    stage = "Writing output"
    Dim prevRt As String
prevRt = vbNullString

For i = 1 To n
    Dim rr As Long
    rr = firstOut + i - 1

    ' Only write the Route when it changes
    If CStr(arrRt(i)) <> CStr(prevRt) Then
        PutCell wsO, rr, oRt, arrRt(i)
        prevRt = CStr(arrRt(i))
    Else
        PutCell wsO, rr, oRt, ""
    End If

    PutCell wsO, rr, oDp, arrDp(i)
    PutCell wsO, rr, oNm, arrNm(i)
Next i

' ---- Centre all filled cells in the output area ----
Dim lastOutRow As Long
lastOutRow = firstOut + n - 1
' Centre text across Route..Name..(Successful/Printed/Checked in if present)
With wsO.Range(wsO.Cells(firstOut, oRt), wsO.Cells(lastOutRow, lastOutCol))
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
End With

' Make every produced row the same height as the template row (firstOut)
Dim tplHt As Double
tplHt = wsO.Rows(firstOut).RowHeight
For r = firstOut To lastOutRow
    wsO.Rows(r).RowHeight = tplHt
Next r

' Clone the template row’s formatting across the output area (borders, fill, etc.)
Dim fmtSrc As Range, fmtDst As Range
Set fmtSrc = wsO.Range(wsO.Cells(firstOut, oRt), wsO.Cells(firstOut, lastOutCol))
Set fmtDst = wsO.Range(wsO.Cells(firstOut, oRt), wsO.Cells(lastOutRow, lastOutCol))
fmtSrc.Copy
fmtDst.PasteSpecial Paste:=xlPasteFormats
Application.CutCopyMode = False
' --- Extend the box formatting to all filled rows ---
Set fmtSrc = wsO.Range(wsO.Cells(firstOut, oRt), wsO.Cells(firstOut, oNm))          ' one “boxed” template row
Set fmtDst = wsO.Range(wsO.Cells(firstOut, oRt), wsO.Cells(lastOutRow, oNm))        ' entire output area
With wsO.Range(wsO.Cells(firstOut, oRt), wsO.Cells(lastOutRow, oNm))
    ' inside grid
    .Borders(xlInsideVertical).LineStyle = xlContinuous
    .Borders(xlInsideHorizontal).LineStyle = xlContinuous
    ' outside frame
    .Borders(xlEdgeLeft).LineStyle = xlContinuous
    .Borders(xlEdgeTop).LineStyle = xlContinuous
    .Borders(xlEdgeRight).LineStyle = xlContinuous
    .Borders(xlEdgeBottom).LineStyle = xlContinuous
End With

fmtSrc.Copy
fmtDst.PasteSpecial Paste:=xlPasteFormats
Application.CutCopyMode = False

With wsO.Range(wsO.Cells(firstOut, oRt), wsO.Cells(lastOutRow, oNm))
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
End With

    '---- Success: wipe Data ----
    stage = "Clearing Data"
    SafeClearAll wsD

    MsgBox "Collections updated: " & n & " rows. Data sheet cleared.", vbInformation
    Exit Sub

Fail:
    MsgBox "Error at stage: " & stage & vbCrLf & Err.Number & " - " & Err.Description, vbCritical
End Sub

'============ BUTTON SETUP (run once to create the button) =============
Public Sub SetupControlButton()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Control")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = "Control"
    End If

    On Error Resume Next
    ws.Shapes("btnUpdateCollections").Delete
    On Error GoTo 0

    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, 30, 30, 240, 44)
    With shp
        .Name = "btnUpdateCollections"
        .TextFrame2.TextRange.Characters.Text = "Update Collections"
        .TextFrame2.TextRange.Font.Size = 14
        .TextFrame2.TextRange.Font.Bold = msoTrue
        .OnAction = "'" & ThisWorkbook.Name & "'!RunCollections_Stable"
    End With

    ws.Range("A6").Value = "Paste manifest into the 'Data' sheet, then click the button."
    ws.Columns("A:B").AutoFit
    MsgBox "Control sheet/button ready.", vbInformation
End Sub

'================== HEADER FINDERS (exact) ==================
Private Function FindDataHeaders(ws As Worksheet, scanTop As Long, _
                                 ByRef cSvc As Long, ByRef cRt As Long, _
                                 ByRef cDp As Long, ByRef cNm As Long) As Long
    Dim r As Long
    For r = 1 To Application.Min(scanTop, ws.Rows.Count)
        cSvc = FindColExactOnRow(ws, r, "Service")
        cRt = FindColExactOnRow(ws, r, "Third Party Route Number")
        cDp = FindColExactOnRow(ws, r, "Drop Number")
        cNm = FindColExactOnRow(ws, r, "Surname")
        If cSvc > 0 And cRt > 0 And cDp > 0 And cNm > 0 Then
            FindDataHeaders = r
            Exit Function
        End If
    Next r
    cSvc = 0: cRt = 0: cDp = 0: cNm = 0
End Function

Private Function FindColExactOnRow(ws As Worksheet, ByVal rowIndex As Long, headerText As String) As Long
    If rowIndex < 1 Then Exit Function
    Dim lastC As Long: lastC = ws.Cells(rowIndex, ws.Columns.Count).End(xlToLeft).Column
    Dim c As Long, txt As String
    For c = 1 To lastC
        txt = Trim$(ws.Cells(rowIndex, c).Text)
        If StrComp(txt, headerText, vbTextCompare) = 0 Then
            If ws.Cells(rowIndex, c).MergeCells Then
                FindColExactOnRow = ws.Cells(rowIndex, c).MergeArea.Column
            Else
                FindColExactOnRow = c
            End If
            Exit Function
        End If
    Next c
End Function

Private Function HeaderDepth(ws As Worksheet, hdrRow As Long, cols As Variant) As Long
    Dim mx As Long
    Dim k As Variant
    mx = 1

    For Each k In cols
        If ws.Cells(hdrRow, k).MergeCells Then
            mx = Application.Max(mx, ws.Cells(hdrRow, k).MergeArea.Rows.Count)
        End If
    Next k

    HeaderDepth = mx
End Function

'================== WRITE / CLEAR ==================
Private Sub PutCell(ws As Worksheet, ByVal r As Long, ByVal c As Long, ByVal v As Variant)
    On Error Resume Next
    With ws.Cells(r, c)
        If .MergeCells Then .MergeArea.UnMerge
        .Value = v
    End With
    On Error GoTo 0
End Sub

Private Sub ClearDownColumn_MergeSafe(ws As Worksheet, ByVal colIndex As Long, ByVal startRow As Long)
    On Error Resume Next
    Dim lastUsed As Long: lastUsed = ws.Cells(ws.Rows.Count, colIndex).End(xlUp).Row
    Dim rr As Long
    If lastUsed < startRow Then Exit Sub
    For rr = startRow To lastUsed
        With ws.Cells(rr, colIndex)
            If .MergeCells Then .MergeArea.UnMerge
            .ClearContents
        End With
    Next rr
    On Error GoTo 0
End Sub

Private Sub SafeClearAll(ws As Worksheet)
    On Error Resume Next
    ws.Cells.ClearContents
    On Error GoTo 0
End Sub

'================== UTIL: formatting & sort ==================
Private Function NeatText(v As Variant) As String
    If IsError(v) Or IsEmpty(v) Then NeatText = "": Exit Function
    If IsNumeric(v) Then
        If CDbl(v) = CLng(v) Then NeatText = CStr(CLng(v)) Else NeatText = CStr(CDbl(v))
    Else
        NeatText = Trim$(CStr(v))
    End If
End Function

Private Sub Sort2(ByRef a() As Variant, ByRef b() As Variant, ByRef c() As Variant)
    Dim n As Long: n = UBound(a)
    If n <= 1 Then Exit Sub
    Dim idx() As Long, i As Long: ReDim idx(1 To n)
    For i = 1 To n: idx(i) = i: Next i

    Dim ar() As Double, br() As Double, ai() As Boolean, bi() As Boolean
    ReDim ar(1 To n): ReDim br(1 To n): ReDim ai(1 To n): ReDim bi(1 To n)
    For i = 1 To n
        ai(i) = TryNum(a(i), ar(i))
        bi(i) = TryNum(b(i), br(i))
    Next i

    Quick idx, 1, n, a, b, ar, br, ai, bi

    Dim A2() As Variant, B2() As Variant, C2() As Variant
    ReDim A2(1 To n): ReDim B2(1 To n): ReDim C2(1 To n)
    For i = 1 To n
        A2(i) = a(idx(i)): B2(i) = b(idx(i)): C2(i) = c(idx(i))
    Next i
    a = A2: b = B2: c = C2
End Sub

Private Sub Quick(ByRef idx() As Long, ByVal lo As Long, ByVal hi As Long, _
                  ByRef a() As Variant, ByRef b() As Variant, _
                  ByRef ar() As Double, ByRef br() As Double, _
                  ByRef ai() As Boolean, ByRef bi() As Boolean)
    Dim i As Long, j As Long, p As Long, t As Long
    i = lo: j = hi: p = idx((lo + hi) \ 2)
    Do While i <= j
        Do While Cmp(idx(i), p, a, b, ar, br, ai, bi) < 0: i = i + 1: Loop
        Do While Cmp(idx(j), p, a, b, ar, br, ai, bi) > 0: j = j - 1: Loop
        If i <= j Then t = idx(i): idx(i) = idx(j): idx(j) = t: i = i + 1: j = j - 1
    Loop
    If lo < j Then Quick idx, lo, j, a, b, ar, br, ai, bi
    If i < hi Then Quick idx, i, hi, a, b, ar, br, ai, bi
End Sub

Private Function Cmp(ByVal ia As Long, ByVal ib As Long, _
                     ByRef a() As Variant, ByRef b() As Variant, _
                     ByRef ar() As Double, ByRef br() As Double, _
                     ByRef ai() As Boolean, ByRef bi() As Boolean) As Long

    If ai(ia) And ai(ib) Then
        If ar(ia) < ar(ib) Then
            Cmp = -1
        ElseIf ar(ia) > ar(ib) Then
            Cmp = 1
        Else
            Cmp = 0
        End If
    Else
        If CStr(a(ia)) < CStr(a(ib)) Then
            Cmp = -1
        ElseIf CStr(a(ia)) > CStr(a(ib)) Then
            Cmp = 1
        Else
            Cmp = 0
        End If
    End If

    If Cmp = 0 Then
        If bi(ia) And bi(ib) Then
            If br(ia) < br(ib) Then
                Cmp = -1
            ElseIf br(ia) > br(ib) Then
                Cmp = 1
            Else
                Cmp = 0
            End If
        Else
            If CStr(b(ia)) < CStr(b(ib)) Then
                Cmp = -1
            ElseIf CStr(b(ia)) > CStr(b(ib)) Then
                Cmp = 1
            Else
                Cmp = 0
            End If
        End If
    End If
End Function

Private Function TryNum(ByVal v As Variant, ByRef outN As Double) As Boolean
    On Error GoTo noN
    If IsNumeric(v) Then outN = CDbl(v): TryNum = True: Exit Function
    outN = CDbl(Val(CStr(v))): TryNum = True: Exit Function
noN:
    TryNum = False
End Function


