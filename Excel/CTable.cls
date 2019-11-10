VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pName As String
Private pTable As Range

Private Const BUTTON_ADD_ROW_TEXT As String = "+"
Private Const BUTTON_DELETE_ROW_TEXT As String = "–"

Public Property Get Address() As String
    Let Address = pTable.Address
End Property

Public Sub AddRow(ByRef InsertBelow As Range)
    If IsEntry(InsertBelow) Then
        Call MRuntime.Save
        Call MRuntime.Fast
        On Error GoTo ErrorHandling
        Set InsertBelow = Intersect(InsertBelow.EntireRow, pTable)
        Call InsertBelow.EntireRow.Copy
        Call InsertBelow.EntireRow.Offset(1).Insert(xlDown)
        Let Application.CutCopyMode = False
        With InsertBelow.Borders(xlEdgeBottom)
            Let .LineStyle = xlContinuous
            Let .Weight = xlHairline
        End With
ErrorHandling:
        Call MRuntime.Restore
    Else
        Dim Message As String
        Let Message = vbNullString
        Let Message = Message & "To add a row to this table, click on the row you wish "
        Let Message = Message & "to insert the new row below, then click the "
        Let Message = Message & BUTTON_ADD_ROW_TEXT
        Let Message = Message & " button."
        Call MsgBox(Message, vbOKOnly + vbInformation, "How to add a row")
        Exit Sub
    End If
    If HasIndex Then Call RefreshIndex
End Sub

Public Sub Bind(ByRef Name As Name)
    Dim Matches As Collection
    Set Matches = MRegex.GetMatches(Name.Name, "Table_(\w+)")
    Let pName = Matches(1)(3)
    Set pTable = Name.RefersToRange
End Sub

Public Property Get ColumnCount() As Long
    Let ColumnCount = pTable.Columns.Count
End Property

Public Sub DeleteRows(ByRef CellsInTable As Range)
    Dim Message As String
    If IsEntry(CellsInTable) Then
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
        Let Message = vbNullString
        Let Message = Message & "To delete rows, click on the row(s) you wish "
        Let Message = Message & "to delete, then click the "
        Let Message = Message & BUTTON_DELETE_ROW_TEXT
        Let Message = Message & " button."
        Call MsgBox(Message, vbOKOnly + vbInformation, "How to delete rows")
        Exit Sub
    End If
    If HasIndex Then Call RefreshIndex
End Sub

Public Property Get HasIndex() As Boolean
    Let HasIndex = False
End Property

Public Property Get Name() As Long
    Let Name = pName
End Property

Public Property Get RowCount() As Long
    Let Name = pTable.Rows.Count - 1
End Property

Private Function IncludesAllRows(ByRef CellsInTable As Range) As Boolean
    Let IncludesAllRows = True
    On Error GoTo ErrorHandling
    If Not Intersect(CellsInTable.Rows(1).Offset(-1), pTable) Is Nothing _
    Or Not Intersect(CellsInTable.Rows(CellsInTable.Rows.Count).Offset(2), pTable) Is Nothing Then
        Let IncludesAllRows = False
    End If
ErrorHandling:
End Function

Private Function IsEntry(ByRef Cells As Range) As Boolean
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
'    On Error Resume Next
'    Call Runtime.SaveSettings
'    Call Runtime.SetToFast(EnableEvents:=False)
'    Dim Index As Range
'    Set Index = WS.Range("TableIndex_" & TableName)
'    Dim i As Integer
'    For i = 1 To Index.Rows.Count - 1
'        Let Index.Cells(i, 1).Value = i
'    Next i
'    Call Runtime.RestoreSettings
End Sub
