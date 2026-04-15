Attribute VB_Name = "mod_Prefill"
'===================== mod_Prefill =====================
Option Explicit

' ---------- Entry point (centralized) ----------
Public Sub Prefill_From_ItemNo()

    Dim wsForm As Worksheet, wsDB As Worksheet
    Dim itemNo As String
    Dim dbData As Variant, headers As Object
    Dim rowsForItem As Collection
    Dim headerRowIx As Long

    On Error GoTo fail

    Set wsForm = GetFormSheet()
    If wsForm Is Nothing Then
        MsgBox "Form sheet '" & FORM_SHEET_NAME & "' not found.", vbCritical
        Exit Sub
    End If

    itemNo = Trim$(GetItemNoFromForm())
    If Len(itemNo) = 0 Then
        Clear_Form_AutoFields wsForm
        Exit Sub
    End If

    Set wsDB = FindDBSheet()
    If wsDB Is Nothing Then
        MsgBox "Sheet '" & DB_SHEET_NAME & "' not found.", vbExclamation
        Exit Sub
    End If

    dbData = GetDBArray(wsDB)
    If IsEmpty(dbData) Then
        MsgBox "DB sheet is empty.", vbExclamation
        Exit Sub
    End If

    Set headers = BuildHeaderIndex(dbData)
    If headers Is Nothing Then
        MsgBox "Could not read DB headers.", vbExclamation
        Exit Sub
    End If
    If Not headers.Exists(COL_ITEMNO) Then
        MsgBox "DB is missing '" & COL_ITEMNO & "' column.", vbExclamation
        Exit Sub
    End If

    Set rowsForItem = FindRowsForItem(dbData, headers, itemNo)
    If rowsForItem Is Nothing Or rowsForItem.count = 0 Then
        ' Item not found -> open YOUR existing UserForm (no extra controls assumed)
        frmAddItem.txtItemNo.Text = itemNo     ' your form has txtItemNo
        frmAddItem.Show                         ' user clicks Save -> UserForm will append rows to DB and call FillFormFromDB itself
        Exit Sub                                ' do not double-run; the UF calls FillFormFromDB on success
    End If

    headerRowIx = PickHeaderRowIndex(dbData, headers, rowsForItem) ' prefers "Core 1" when present

    ' 1) Header technical fields (+ optional Cust. Item No / Part code)
    Fill_Header_TechFields wsForm, dbData, headers, headerRowIx

    ' 2) Core grid (Core 1/2/3)
    Fill_Core_Grid wsForm, dbData, headers, rowsForItem
    Exit Sub

fail:
    MsgBox "Prefill error: " & Err.Description, vbExclamation

End Sub

' ---------- Header writers ----------
Public Sub Fill_Header_TechFields(wsForm As Worksheet, dbData As Variant, headers As Object, headerRowIx As Long)
    ' Map: DB column -> (Form label, NamedTarget [optional])
    Dim map As Object: Set map = CreateObject("Scripting.Dictionary")
    map.CompareMode = vbTextCompare
    map("CT Type") = Array("CT TYPE", "HDR_CT_TYPE")
    map("RATIO :-") = Array("RATIO :-", "HDR_RATIO_HEADLINE")
    map("RATED VOLTAGE") = Array("RATED VOLTAGE", "HDR_RATED_VOLTAGE")
    map("STC") = Array("STC", "HDR_STC")
    map("I.L.") = Array("I.L.", "HDR_IL")
    map("FREQ.") = Array("FREQ.", "HDR_FREQ")
    map("REF. STD.") = Array("REF. STD.", "HDR_REF_STD")
    If ALLOW_PREFILL_CUST_PART Then map("Cust. Item No / Part code") = Array("Cust. Item No / Part code", "HDR_CUST_PART")

    Dim k As Variant, arr As Variant
    For Each k In map.Keys
        arr = map(k)
        SafeWriteHeaderField wsForm, dbData, headers, headerRowIx, _
            dbColName:=CStr(k), _
            formLabel:=CStr(arr(0)), _
            namedTarget:=CStr(arr(1))
    Next k
End Sub

Private Sub SafeWriteHeaderField(wsForm As Worksheet, dbData As Variant, headers As Object, rowIx As Long, _
                                 ByVal dbColName As String, ByVal formLabel As String, ByVal namedTarget As String)
    Dim valText As String: valText = GetField(dbData, headers, rowIx, dbColName)
    Dim tgt As Range
    Set tgt = TryGetNamedRange(namedTarget, ThisWorkbook)
    If Not tgt Is Nothing Then
        SafeWriteCell tgt, valText
    Else
        Dim lbl As Range, rightCell As Range
        Set lbl = FindLabelCellRight(wsForm, formLabel)
        If Not lbl Is Nothing Then
            Set rightCell = lbl.Offset(0, 1)
            SafeWriteCell rightCell, valText
        End If
    End If
End Sub

' ---------- Core grid fill ----------
Public Sub Fill_Core_Grid(wsForm As Worksheet, dbData As Variant, headers As Object, rowsForItem As Collection)
    Dim particularsCol As Long, core1Col As Long, core2Col As Long, core3Col As Long
    Dim headerRow As Long, firstDataRow As Long, lastDataRow As Long
    If Not LocateGrid_NoUDT(wsForm, particularsCol, core1Col, core2Col, core3Col, headerRow, firstDataRow, lastDataRow) Then
        MsgBox "Could not locate the 'PARTICULARS' grid (need 'PARTICULARS','Core 1','Core 2','Core 3').", vbExclamation
        Exit Sub
    End If

    ' Figure which DB rows are Core 1/2/3
    Dim coreRowIx As Object: Set coreRowIx = CreateObject("Scripting.Dictionary")
    coreRowIx.CompareMode = vbTextCompare
    coreRowIx("Core 1") = 0
    coreRowIx("Core 2") = 0
    coreRowIx("Core 3") = 0

    Dim i As Long, r As Long, coreName As String
    For i = 1 To rowsForItem.count
        r = rowsForItem(i)
        coreName = GetField(dbData, headers, r, "Core")
        If Len(coreName) > 0 Then If coreRowIx.Exists(coreName) Then coreRowIx(coreName) = r
    Next i

    ' Build the map of Form labels -> DB headers (matches your DB header list)
    Dim map As Object: Set map = BuildCoreMap()

    Dim label As Variant, dbCol As String, rowIx As Long, valText As String
    Dim labelRow As Long
    Dim tgt As Range

    For Each label In map.Keys
        dbCol = map(label)
        labelRow = FindRowByLabel(wsForm, particularsCol, firstDataRow, CStr(label))
        If labelRow > 0 Then
            ' Core 1
            rowIx = coreRowIx("Core 1")
            valText = IIf(rowIx > 0, GetField(dbData, headers, rowIx, dbCol), "")
            Set tgt = wsForm.Cells(labelRow, core1Col): SafeWriteCell tgt, valText
            ' Core 2
            rowIx = coreRowIx("Core 2")
            valText = IIf(rowIx > 0, GetField(dbData, headers, rowIx, dbCol), "")
            Set tgt = wsForm.Cells(labelRow, core2Col): SafeWriteCell tgt, valText
            ' Core 3
            rowIx = coreRowIx("Core 3")
            valText = IIf(rowIx > 0, GetField(dbData, headers, rowIx, dbCol), "")
            Set tgt = wsForm.Cells(labelRow, core3Col): SafeWriteCell tgt, valText
        End If
    Next label
End Sub

' ---------- Clear (headers + grid) ----------
Public Sub Clear_Form_AutoFields(wsForm As Worksheet)
    ' Clear core grid (merged-safe)
    Dim particularsCol As Long, core1Col As Long, core2Col As Long, core3Col As Long
    Dim headerRow As Long, firstDataRow As Long, lastDataRow As Long

    If LocateGrid_NoUDT(wsForm, particularsCol, core1Col, core2Col, core3Col, headerRow, firstDataRow, lastDataRow) Then
        If lastDataRow >= firstDataRow Then SafeClearRect wsForm, firstDataRow, core1Col, lastDataRow, core3Col
    End If

    ' Clear header fields we prefill (same mapping used in Fill_Header_TechFields)
    Dim map As Object: Set map = CreateObject("Scripting.Dictionary")
    map.CompareMode = vbTextCompare
    map("CT Type") = Array("CT TYPE", "HDR_CT_TYPE")
    map("RATIO :-") = Array("RATIO :-", "HDR_RATIO_HEADLINE")
    map("RATED VOLTAGE") = Array("RATED VOLTAGE", "HDR_RATED_VOLTAGE")
    map("STC") = Array("STC", "HDR_STC")
    map("I.L.") = Array("I.L.", "HDR_IL")
    map("FREQ.") = Array("FREQ.", "HDR_FREQ")
    map("REF. STD.") = Array("REF. STD.", "HDR_REF_STD")
    If ALLOW_PREFILL_CUST_PART Then map("Cust. Item No / Part code") = Array("Cust. Item No / Part code", "HDR_CUST_PART")

    Dim k As Variant, arr As Variant
    For Each k In map.Keys
        arr = map(k)
        SafeClearHeaderField wsForm, formLabel:=CStr(arr(0)), namedTarget:=CStr(arr(1))
    Next k
End Sub

Private Sub SafeClearHeaderField(wsForm As Worksheet, ByVal formLabel As String, ByVal namedTarget As String)
    Dim tgt As Range, lbl As Range, rightCell As Range
    Set tgt = TryGetNamedRange(namedTarget, ThisWorkbook)
    If Not tgt Is Nothing Then
        SafeClearCell tgt
    Else
        Set lbl = FindLabelCellRight(wsForm, formLabel)
        If Not lbl Is Nothing Then
            Set rightCell = lbl.Offset(0, 1)
            SafeClearCell rightCell
        End If
    End If
End Sub

' ---------- Grid location + tolerant label matching ----------
Public Function LocateGrid_NoUDT(ws As Worksheet, _
    ByRef particularsCol As Long, _
    ByRef core1Col As Long, _
    ByRef core2Col As Long, _
    ByRef core3Col As Long, _
    ByRef headerRow As Long, _
    ByRef firstDataRow As Long, _
    ByRef lastDataRow As Long) As Boolean

    Dim c As Range
    LocateGrid_NoUDT = False
    Set c = FindCell(ws, "PARTICULARS")
    If c Is Nothing Then Exit Function

    particularsCol = c.Column
    headerRow = c.Row
    core1Col = FindInRow(ws, headerRow, "Core 1")
    core2Col = FindInRow(ws, headerRow, "Core 2")
    core3Col = FindInRow(ws, headerRow, "Core 3")
    If core1Col = 0 Or core2Col = 0 Or core3Col = 0 Then Exit Function

    firstDataRow = headerRow + 1
    lastDataRow = FindLastLabelRow(ws, particularsCol, firstDataRow)
    If lastDataRow = 0 Then lastDataRow = ws.Cells(ws.rows.count, particularsCol).End(xlUp).Row
    LocateGrid_NoUDT = True
End Function

Public Function FindLastLabelRow(ws As Worksheet, col As Long, startRow As Long) As Long
    Dim r As Long, blankStreak As Long, lastSeen As Long
    FindLastLabelRow = 0
    r = startRow: blankStreak = 0: lastSeen = 0
    Do While r < ws.rows.count
        If Len(Trim$(CStr(ws.Cells(r, col).Value))) = 0 Then
            blankStreak = blankStreak + 1
            If blankStreak >= 10 Then Exit Do
        Else
            lastSeen = r
            blankStreak = 0
        End If
        r = r + 1
    Loop
    FindLastLabelRow = lastSeen
End Function

Public Function FindRowByLabel(ws As Worksheet, labelCol As Long, startRow As Long, labelText As String) As Long
    Dim r As Long, maxR As Long, txt As String
    maxR = ws.Cells(ws.rows.count, labelCol).End(xlUp).Row
    For r = startRow To maxR
        txt = CStr(ws.Cells(r, labelCol).Value)
        If Norm(txt) = Norm(labelText) Then
            FindRowByLabel = r
            Exit Function
        End If
    Next r
    FindRowByLabel = 0
End Function

Public Function FindCell(ws As Worksheet, labelText As String) As Range
    Dim rng As Range
    For Each rng In ws.UsedRange
        If Norm(CStr(rng.Value)) = Norm(labelText) Then
            Set FindCell = rng
            Exit Function
        End If
    Next rng
    Set FindCell = Nothing
End Function

Public Function FindInRow(ws As Worksheet, rowIx As Long, labelText As String) As Long
    Dim c As Long, lastCol As Long
    lastCol = ws.Cells(rowIx, ws.Columns.count).End(xlToLeft).Column
    For c = 1 To lastCol
        If Norm(CStr(ws.Cells(rowIx, c).Value)) = Norm(labelText) Then
            FindInRow = c
            Exit Function
        End If
    Next c
    FindInRow = 0
End Function

Public Function FindLabelCellRight(ws As Worksheet, labelText As String) As Range
    Dim rng As Range
    For Each rng In ws.UsedRange
        If Norm(CStr(rng.Value)) = Norm(labelText) Then
            Set FindLabelCellRight = rng
            Exit Function
        End If
    Next rng
    Set FindLabelCellRight = Nothing
End Function

Public Function Norm(ByVal s As String) As String
    Dim t As String
    t = UCase$(s)
    t = Replace(t, " ", "")
    t = Replace(t, ":", "")
    t = Replace(t, "-", "")
    t = Replace(t, "’", "")
    t = Replace(t, "'", "")
    t = Replace(t, ".", "")
    t = Replace(t, "(", "")
    t = Replace(t, ")", "")
    t = Replace(t, "/", "")
    t = Replace(t, "@VK/2", "ATVK/2")
    Norm = t
End Function

' ---------- DB field access ----------
Public Function FindRowsForItem(dbData As Variant, headers As Object, itemNo As String) As Collection
    Dim rows As New Collection
    Dim r As Long, lastR As Long, colItem As Long, val As String
    colItem = headers(COL_ITEMNO)
    lastR = UBound(dbData, 1)
    For r = 2 To lastR
        val = CStr(dbData(r, colItem))
        If StrComp(val, itemNo, vbTextCompare) = 0 Then rows.Add r
    Next r
    Set FindRowsForItem = rows
End Function

Public Function PickHeaderRowIndex(dbData As Variant, headers As Object, rows As Collection) As Long
    Dim i As Long, r As Long, colCore As Long
    If rows.count = 0 Then Exit Function
    If headers.Exists("Core") Then
        colCore = headers("Core")
        For i = 1 To rows.count
            r = rows(i)
            If StrComp(CStr(dbData(r, colCore)), "Core 1", vbTextCompare) = 0 Then
                PickHeaderRowIndex = r
                Exit Function
            End If
        Next i
    End If
    PickHeaderRowIndex = rows(1)
End Function

Public Function GetField(dbData As Variant, headers As Object, rowIx As Long, colName As String) As String
    If rowIx <= 0 Then Exit Function
    If Not headers Is Nothing Then
        If headers.Exists(colName) Then
            GetField = CStr(dbData(rowIx, headers(colName)))
            Exit Function
        End If
        ' Fallback: Form label "Core Dimensions" -> DB "Bare Core Dimensions"
        If UCase$(colName) = "CORE DIMENSIONS" And headers.Exists("Bare Core Dimensions") Then
            GetField = CStr(dbData(rowIx, headers("Bare Core Dimensions")))
            Exit Function
        End If
    End If
    GetField = ""
End Function

' ---------- Core map (Form labels -> DB headers) ----------
Public Function BuildCoreMap() As Object
    Dim m As Object: Set m = CreateObject("Scripting.Dictionary")
    m.CompareMode = vbTextCompare

    ' These keys are the Form's "PARTICULARS" labels; values are EXACT DB headers you sent.
    m("RATIO") = "RATIO"
    m("Burden (VA)") = "Burden (VA)"
    m("Accuracy Class") = "Accuracy Class"
    m("ISF") = "ISF"
    m("Min. Knee pt. volt.") = "Min. Knee pt. volt."
    m("Max. Rct @ 75'c") = "Max. Rct @ 75'c"
    m("Max. Exc. C/n. :- @VK/2") = "Max. Exc. C/n. :- @VK/2"

    ' Form shows "Core Dimensions" label; DB header is "Bare Core Dimensions"
    ' GetField() already maps "Core Dimensions" -> "Bare Core Dimensions" as fallback.
    m("Core Dimensions") = "Core Dimensions"

    m("Core Material") = "Core Material"
    m("Core weight (Kg)") = "Core weight (Kg)"
    m("Sec. Total Turns") = "Sec. Total Turns"
    m("Sec. Ter. Marking") = "Sec. Ter. Marking"
    m("Sec. Conductor (S1-S2)") = "Sec. Conductor (S1-S2)"
    m("Sec. Turns (S1-S2)") = "Sec. Turns (S1-S2)"
    m("Sec. Conductor (S2-S3)") = "Sec. Conductor (S2-S3)"
    m("Sec. Turns (S2-S3)") = "Sec. Turns (S2-S3)"
    m("Sec. Conductor (S3-S4)") = "Sec. Conductor (S3-S4)"
    m("Sec. Turns (S3-S4)") = "Sec. Turns (S3-S4)"
    m("Sec. Conductor (S4-S5)") = "Sec. Conductor (S4-S5)"
    m("Sec. Turns (S4-S5)") = "Sec. Turns (S4-S5)"
    m("Sec. Copper weight (kg)") = "Sec. Copper weight (kg)"
    m("Finished Core Dim. (mm)") = "Finished Core Dim. (mm)"
    m("Sec Connection") = "Sec Connection"
    m("Wire Length") = "Wire Length"
    m("Wire Colour") = "Wire Colour"
    m("CT final dim") = "CT final dim"
    m("GA Drg") = "GA Drg"
    m("INS CLASS") = "INS CLASS"
    m("PRI Turns") = "PRI Turns"
m("PRI Copper") = "PRI Copper"
m("Former") = "Former"
m("PRI Length") = "PRI Length"
m("PRI Weight") = "PRI Weight"
m("Sec. Terminal") = "Sec. Terminal"


    Set BuildCoreMap = m
End Function
'=============================================================

