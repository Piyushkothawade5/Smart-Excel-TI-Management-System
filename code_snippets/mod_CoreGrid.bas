Attribute VB_Name = "mod_CoreGrid"

Public Function GetCoreColumns(ws As Worksheet) As Variant

Dim c As Range
Dim col1 As Long, col2 As Long, col3 As Long

For Each c In ws.UsedRange


If Trim(c.Value) = "Core 1" Then col1 = c.Column
If Trim(c.Value) = "Core 2" Then col2 = c.Column
If Trim(c.Value) = "Core 3" Then col3 = c.Column


Next c

GetCoreColumns = Array(col1, col2, col3)

End Function

Public Function ReadCoreGrid() As Collection

Dim ws As Worksheet
Dim lblRange As Range
Dim r As Range
Dim result As New Collection
Dim dict As Object
Dim cores

Set ws = ThisWorkbook.Worksheets("Form")
Set lblRange = ThisWorkbook.Names("CoreLabels").RefersToRange

cores = GetCoreColumns(ws)

For Each r In lblRange


If Trim(r.Value) <> "" Then

    Set dict = CreateObject("Scripting.Dictionary")
    
    dict("Label") = Trim(r.Value)
    dict("Core1") = ws.Cells(r.Row, cores(0)).Value
    dict("Core2") = ws.Cells(r.Row, cores(1)).Value
    dict("Core3") = ws.Cells(r.Row, cores(2)).Value
    
    result.Add dict
    
End If


Next r

Set ReadCoreGrid = result

End Function

Public Sub SaveCoreLines(headerID As Long)

Dim lo As ListObject
Dim ws As Worksheet
Dim rows As Collection
Dim r As Object
Dim lr As ListRow
Dim nextLineID As Long
Dim colIndex As Long

Set ws = ThisWorkbook.Worksheets("Transactions_Lines")
Set lo = ws.ListObjects("T_Lines")

Set rows = ReadCoreGrid()

nextLineID = nextID(lo, "LineID")

Dim coreNames(1 To 3) As String
coreNames(1) = "Core 1"
coreNames(2) = "Core 2"
coreNames(3) = "Core 3"

Dim coreVal
Dim coreNumber As Integer
Dim dbHeader As String

For coreNumber = 1 To 3
Dim hasData As Boolean
hasData = False

For Each r In rows
    If coreNumber = 1 And r("Core1") <> "" Then hasData = True
    If coreNumber = 2 And r("Core2") <> "" Then hasData = True
    If coreNumber = 3 And r("Core3") <> "" Then hasData = True
Next r

If Not hasData Then GoTo NextCore
    Set lr = lo.ListRows.Add

    lr.Range(lo.ListColumns("LineID").Index) = nextLineID
    lr.Range(lo.ListColumns("HeaderID").Index) = headerID
    lr.Range(lo.ListColumns("Item No").Index) = Range("Form_ItemNo").Value
    lr.Range(lo.ListColumns("Core").Index) = coreNames(coreNumber)

    For Each r In rows

        Dim col As ListColumn
        Dim found As Boolean

        found = False
        dbHeader = r("Label")

        ' FIX: Map form label to DB column
        If dbHeader = "Core Dimensions" Then
            dbHeader = "Bare Core Dimensions"
        End If

        For Each col In lo.ListColumns
            If col.name = dbHeader Then
                colIndex = col.Index
                found = True
                Exit For
            End If
        Next col

        If found Then

            If coreNumber = 1 Then coreVal = r("Core1")
            If coreNumber = 2 Then coreVal = r("Core2")
            If coreNumber = 3 Then coreVal = r("Core3")

            lr.Range(colIndex) = coreVal

        End If

    Next r

    nextLineID = nextLineID + 1
NextCore:
Next coreNumber

End Sub





Public Sub WriteCoreGrid(label As String, coreName As String, val As Variant)

Dim ws As Worksheet
Dim lblRange As Range
Dim r As Range
Dim coreCols
Dim colIndex As Long

Set ws = ThisWorkbook.Worksheets("Form")
Set lblRange = ThisWorkbook.Names("CoreLabels").RefersToRange

coreCols = GetCoreColumns(ws)

If coreName = "Core 1" Then colIndex = coreCols(0)
If coreName = "Core 2" Then colIndex = coreCols(1)
If coreName = "Core 3" Then colIndex = coreCols(2)

For Each r In lblRange

If Trim(r.Value) = label Then

    ws.Cells(r.Row, colIndex).Value = val
    Exit Sub
    
End If

Next r

End Sub

Public Function nextID(lo As ListObject, colName As String) As Long

Dim maxID As Long
Dim c As Range

maxID = 0

If lo.ListRows.count = 0 Then
nextID = 1
Exit Function
End If

For Each c In lo.ListColumns(colName).DataBodyRange


If IsNumeric(c.Value) Then
    If c.Value > maxID Then maxID = c.Value
End If


Next c

nextID = maxID + 1

End Function



