Attribute VB_Name = "mod_TransactionLoader"
Option Explicit

Public Sub LoadTransaction(ByVal headerID As Long)
UI_Begin
UnlockFormSheet

CurrentHeaderID = headerID
ActiveSheet.Shapes("btnEdit").Visible = msoTrue
IsEditMode = False

Dim wsH As Worksheet
Dim wsL As Worksheet
Dim loH As ListObject
Dim loL As ListObject

Dim arrH
Dim arrL

Dim i As Long
Dim rowIndex As Long

Set wsH = ThisWorkbook.Worksheets("Transactions_Header")
Set loH = wsH.ListObjects("T_Header")

Set wsL = ThisWorkbook.Worksheets("Transactions_Lines")
Set loL = wsL.ListObjects("T_Lines")

If loH.ListRows.count = 0 Then Exit Sub

arrH = loH.DataBodyRange.Value

For i = 1 To UBound(arrH, 1)

    If arrH(i, loH.ListColumns("HeaderID").Index) = headerID Then
        rowIndex = i
        Exit For
    End If

Next i

If rowIndex = 0 Then
    MsgBox "Transaction not found.", vbExclamation
    Exit Sub
End If


'================ HEADER LOAD =================

WriteNamed "HDR_PO_ITEM_NO", arrH(rowIndex, loH.ListColumns("PO ITEM NO.").Index)
WriteNamed "HDR_CUS_OD_NO", arrH(rowIndex, loH.ListColumns("CUS. ORDER. NO.").Index)
WriteNamed "Form_ItemNo", arrH(rowIndex, loH.ListColumns("Item No").Index)
WriteNamed "HDR_CT_TYPE", arrH(rowIndex, loH.ListColumns("CT Type").Index)
WriteNamed "HDR_CUST_PART", arrH(rowIndex, loH.ListColumns("Cust. Item No / Part code").Index)
WriteNamed "HDR_RATIO_HEADLINE", arrH(rowIndex, loH.ListColumns("RATIO :-").Index)
WriteNamed "HDR_RATED_VOLTAGE", arrH(rowIndex, loH.ListColumns("RATED VOLTAGE").Index)
WriteNamed "HDR_STC", arrH(rowIndex, loH.ListColumns("STC").Index)
WriteNamed "HDR_IL", arrH(rowIndex, loH.ListColumns("I.L.").Index)
WriteNamed "HDR_FREQ", arrH(rowIndex, loH.ListColumns("FREQ.").Index)
WriteNamed "HDR_REF_STD", arrH(rowIndex, loH.ListColumns("REF. STD.").Index)
WriteNamed "HDR_TI_NO", arrH(rowIndex, loH.ListColumns("TI_No").Index)
WriteNamed "HDR_TI_DATE", arrH(rowIndex, loH.ListColumns("TI_Date").Index)
WriteNamed "HDR_CUSTOMER_NAME", arrH(rowIndex, loH.ListColumns("Customer").Index)
WriteNamed "HDR_CUS_ORDER_DATE", arrH(rowIndex, loH.ListColumns("CUS_ORDER_DATE").Index)
WriteNamed "HDR_WO_NO", arrH(rowIndex, loH.ListColumns("WO_No").Index)
WriteNamed "HDR_QTY", arrH(rowIndex, loH.ListColumns("QTY").Index)
WriteNamed "HDR_SR_NO", arrH(rowIndex, loH.ListColumns("Sr_No").Index)


'================ CLEAR GRID FIRST =================

ClearCoreGrid
ResetCoreGrid

'================ LOAD LINES =================

If loL.ListRows.count = 0 Then GoTo Finish

arrL = loL.DataBodyRange.Value

For i = 1 To UBound(arrL, 1)

    If arrL(i, loL.ListColumns("HeaderID").Index) = headerID Then
        FillCoreFromLine arrL, loL, i
    End If

Next i


Finish:
UI_End
LockFormSheet

End Sub

Private Sub WriteNamed(nm As String, val As Variant)


Dim rng As Range

On Error Resume Next
Set rng = ThisWorkbook.Names(nm).RefersToRange
On Error GoTo 0

If rng Is Nothing Then Exit Sub

If rng.MergeCells Then
    rng.MergeArea.Cells(1, 1).Value = val
Else
    rng.Value = val
End If


End Sub

Private Sub FillCoreFromLine(arr, lo As ListObject, rowIndex As Long)

Dim coreName As String
Dim c As ListColumn
Dim label As String

coreName = arr(rowIndex, lo.ListColumns("Core").Index)

For Each c In lo.ListColumns

    label = c.name

    ' Skip system columns
    If label <> "LineID" _
    And label <> "HeaderID" _
    And label <> "Item No" _
    And label <> "Core" Then

        ' ===== FIX COLUMN NAME MISMATCH =====
        If label = "Bare Core Dimensions" Then
            label = "Core Dimensions"
        End If
        ' ====================================

        WriteCoreGrid label, coreName, arr(rowIndex, c.Index)

    End If

Next c

End Sub


Private Function FindCoreColumn(ws As Worksheet, coreName As String) As Long

Dim c As Range

' Search only first 20 rows where core headers exist
For Each c In ws.rows("1:38").Cells

    If Trim(c.Value) = coreName Then
        FindCoreColumn = c.Column
        Exit Function
    End If

Next c

FindCoreColumn = 0

End Function

Private Sub WriteGrid(ws As Worksheet, coreCol As Long, labelText As String, val As Variant)

Dim r As Range

For Each r In ws.UsedRange

    If Trim(r.Value) = labelText Then

        ' Write value in correct core column
        If ws.Cells(r.Row, coreCol).MergeCells Then
            ws.Cells(r.Row, coreCol).MergeArea.Cells(1, 1).Value = val
        Else
            ws.Cells(r.Row, coreCol).Value = val
        End If

    End If

Next r

End Sub

Private Sub ClearCoreGrid()


Dim ws As Worksheet
Dim r As Range
Dim c As Long

Set ws = ThisWorkbook.Worksheets("Form")

For Each r In ws.UsedRange

    If Trim(r.Value) = "RATIO" _
    Or Trim(r.Value) = "Burden (VA)" _
    Or Trim(r.Value) = "Accuracy Class" _
    Or Trim(r.Value) = "ISF" Then

        For c = 1 To 10

            If r.Offset(0, c).MergeCells Then
                r.Offset(0, c).MergeArea.ClearContents
            Else
                r.Offset(0, c).ClearContents
            End If

        Next c

    End If

Next r


End Sub

Private Sub ResetCoreGrid()


Dim ws As Worksheet
Dim r As Range

Set ws = ThisWorkbook.Worksheets("Form")

For Each r In ws.UsedRange

    If Trim(r.Value) = "RATIO" _
    Or Trim(r.Value) = "Burden (VA)" _
    Or Trim(r.Value) = "Accuracy Class" _
    Or Trim(r.Value) = "ISF" Then

        If r.Offset(0, 1).MergeCells Then
            r.Offset(0, 1).MergeArea.ClearContents
        Else
            r.Offset(0, 1).ClearContents
        End If

        If r.Offset(0, 2).MergeCells Then
            r.Offset(0, 2).MergeArea.ClearContents
        Else
            r.Offset(0, 2).ClearContents
        End If

        If r.Offset(0, 3).MergeCells Then
            r.Offset(0, 3).MergeArea.ClearContents
        Else
            r.Offset(0, 3).ClearContents
        End If

    End If

Next r


End Sub



