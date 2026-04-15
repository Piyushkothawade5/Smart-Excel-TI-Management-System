Attribute VB_Name = "mod_Navigation"
Option Explicit

Public TodayHeaders() As Long
Public CurrentIndex As Long

Private Function ArrayHasData() As Boolean


On Error GoTo NoArray

If UBound(TodayHeaders) >= 1 Then
    ArrayHasData = True
End If

Exit Function


NoArray:
ArrayHasData = False

End Function

Public Sub BuildTodayHeaderList()


Dim ws As Worksheet
Dim lo As ListObject
Dim arr
Dim i As Long
Dim savedCol As Long
Dim idCol As Long
Dim count As Long

Erase TodayHeaders

Set ws = ThisWorkbook.Worksheets("Transactions_Header")
Set lo = ws.ListObjects("T_Header")

If lo.ListRows.count = 0 Then
    MsgBox "No records found in T_Header.", vbInformation
    Exit Sub
End If

savedCol = lo.ListColumns("SavedAt").Index
idCol = lo.ListColumns("HeaderID").Index

arr = lo.DataBodyRange.Value

count = 0

For i = 1 To UBound(arr, 1)

    If DateValue(arr(i, savedCol)) = Date Then
    
        count = count + 1
        
        ReDim Preserve TodayHeaders(1 To count)
        
        TodayHeaders(count) = arr(i, idCol)
    
    End If
    
Next i


If count = 0 Then

    MsgBox "No transactions saved today.", vbInformation
    Exit Sub
    
End If


CurrentIndex = count


End Sub

Public Sub Nav_Last()

BuildTodayHeaderList

If Not ArrayHasData Then Exit Sub

CurrentIndex = UBound(TodayHeaders)

LoadTransaction TodayHeaders(CurrentIndex)


End Sub

Public Sub Nav_First()

Application.EnableEvents = False
BuildTodayHeaderList

If Not ArrayHasData Then Exit Sub

CurrentIndex = 1

LoadTransaction TodayHeaders(CurrentIndex)

End Sub
Public Sub Nav_Next()

If Not BuildTodayHeaderListIfNeeded Then Exit Sub

'If blank form
If CurrentIndex = 0 Then
    CurrentIndex = UBound(TodayHeaders)
ElseIf CurrentIndex < UBound(TodayHeaders) Then
    CurrentIndex = CurrentIndex + 1
Else
    MsgBox "No next record.", vbInformation
    Exit Sub
End If

LoadTransaction TodayHeaders(CurrentIndex)

End Sub

Public Sub Nav_Prev()

If Not BuildTodayHeaderListIfNeeded Then Exit Sub

'If blank form
If CurrentIndex = 0 Then
    CurrentIndex = UBound(TodayHeaders)
ElseIf CurrentIndex > LBound(TodayHeaders) Then
    CurrentIndex = CurrentIndex - 1
Else
    MsgBox "No previous record.", vbInformation
    Exit Sub
End If

LoadTransaction TodayHeaders(CurrentIndex)

End Sub

Private Function BuildTodayHeaderListIfNeeded() As Boolean

Dim ws As Worksheet
Dim lo As ListObject
Dim arr
Dim i As Long
Dim todayDate As Date

todayDate = Date

Set ws = ThisWorkbook.Worksheets("Transactions_Header")
Set lo = ws.ListObjects("T_Header")

If lo.ListRows.count = 0 Then Exit Function

arr = lo.DataBodyRange.Value

ReDim TodayHeaders(1 To lo.ListRows.count)

Dim count As Long
count = 0

For i = 1 To UBound(arr, 1)

    If Int(arr(i, lo.ListColumns("SavedAt").Index)) = todayDate Then
    
        count = count + 1
        TodayHeaders(count) = arr(i, lo.ListColumns("HeaderID").Index)
        
    End If

Next i

If count = 0 Then Exit Function

ReDim Preserve TodayHeaders(1 To count)

BuildTodayHeaderListIfNeeded = True

End Function
