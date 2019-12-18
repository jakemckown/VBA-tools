VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Requires:    MRegex.bas
'              MRuntime.bas
'              MTable.bas

Option Explicit

Private pName As String
Private pTable As Range
Private pIndex As Range

Public Property Get Address() As String
    Let Address = pTable.Address
End Property

Public Sub AddRow(ByVal InsertBelow As Range, _
                  ByVal LineStyle As tLineStyleEnum, _
                  ByVal BorderWeight As tBorderWeightEnum)
    If IsEntry(InsertBelow) Then
        Call MRuntime.Save
        Call MRuntime.Fast
        On Error GoTo ErrorHandling
        Set InsertBelow = Intersect(InsertBelow.EntireRow, pTable)
        Call InsertBelow.EntireRow.Copy
        Call InsertBelow.EntireRow.Offset(1).Insert(xlDown)
        Let Application.CutCopyMode = False
        With InsertBelow.Borders(xlEdgeBottom)
            Let .LineStyle = LineStyle
            If Not BorderWeight = tBorderWeightNone Then Let .Weight = BorderWeight
        End With
ErrorHandling:
        Call MRuntime.Restore
    Else
        Call DisplayCellsOutsideTableMessage
        Exit Sub
    End If
    If Not pIndex Is Nothing Then Call RefreshIndex
End Sub

Public Sub Bind(ByVal Name As Name)
    Dim Matches As Collection
    Set Matches = MRegex.GetMatches(Name.Name, "Table_(\w+)")
    Let pName = Matches(1)(3)
    Set pTable = Name.RefersToRange
    Set pIndex = Nothing
    On Error GoTo ErrorHandling
    Dim WS As Worksheet
    Set WS = pTable.Worksheet
    Set pIndex = WS.Names("TableIndex_" & pName).RefersToRange
ErrorHandling:
End Sub

Public Property Get ColumnCount() As Long
    Let ColumnCount = pTable.Columns.Count
End Property

Public Sub DeleteRows(ByVal CellsInTable As Range)
    If IsEntry(CellsInTable) Then
        Dim Message As String
        If IncludesAllRows(CellsInTable) Then
            Let Message = "Cannot delete the selected row(s). At least one row must remain."
            Call MsgBox(Message, vbOKOnly + vbInformation, "Cannot delete row(s)")
            Exit Sub
        End If
        Dim Response As Long
        Let Message = "Delete selected row(s)? This cannot be undone."
        Let Response = MsgBox(Message, vbYesNo + vbQuestion, "Delete selected row(s)?")
        Select Case Response
            Case vbYes
                Call MRuntime.Save
                Call MRuntime.Fast
                On Error GoTo ErrorHandling
                Dim i As Long
                For i = CellsInTable.Rows.Count To 1 Step -1
                    Call CellsInTable.Rows(i).EntireRow.Delete(xlUp)
                Next i
                With pTable
                    Let .Rows(.Rows.Count - 1).Borders(xlEdgeBottom).Weight = xlThin
                End With
                If ActiveCell.Locked Then Call ActiveCell.Offset(-1).Select
                If ActiveCell.Locked Then Call ActiveCell.Offset(1).Select
ErrorHandling:
                Call MRuntime.Restore
            Case vbNo
                Exit Sub
        End Select
    Else
        Call DisplayCellsOutsideTableMessage
        Exit Sub
    End If
    If Not pIndex Is Nothing Then Call RefreshIndex
End Sub

Public Property Get Name() As Long
    Let Name = pName
End Property

Public Property Get RowCount() As Long
    Let Name = pTable.Rows.Count - 1
End Property

Private Sub DisplayCellsOutsideTableMessage()
    Dim Message As String
    Let Message = vbNullString
    Let Message = Message & "One or more cells in the current selection are not "
    Let Message = Message & "within the table. Please select only cells within the "
    Let Message = Message & "table and try again."
    Call MsgBox(Message, vbOKOnly + vbInformation, "One or more cells outside table")
End Sub

Private Function IncludesAllRows(ByVal CellsInTable As Range) As Boolean
    Let IncludesAllRows = True
    On Error GoTo ErrorHandling
    If Not Intersect(CellsInTable.Rows(1).Offset(-1), pTable) Is Nothing _
    Or Not Intersect(CellsInTable.Rows(CellsInTable.Rows.Count).Offset(2), pTable) Is Nothing Then
        Let IncludesAllRows = False
    End If
ErrorHandling:
End Function

Private Function IsEntry(ByVal Cells As Range) As Boolean
    Let IsEntry = False
    On Error GoTo ErrorHandling
    Dim Cell As Range
    For Each Cell In Cells
        If Intersect(Cell, pTable) Is Nothing Then Exit Function
    Next Cell
    Let IsEntry = True
ErrorHandling:
End Function

Private Sub RefreshIndex()
    On Error Resume Next
    Call MRuntime.Save
    Call MRuntime.Fast
    Dim i As Integer
    For i = 1 To pIndex.Rows.Count - 1
        Let pIndex.Cells(i, 1).Value = CStr(i)
    Next i
    Call MRuntime.Restore
End Sub
