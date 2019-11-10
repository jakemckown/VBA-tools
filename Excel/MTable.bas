Attribute VB_Name = "MTable"
Option Explicit

Public Sub AddRow()
    If ActiveWorkbook.Name <> ThisWorkbook.Name Then Exit Sub
    If Not MDeveloper.Mode Then
        Dim Table As CTable
        Set Table = GetTable(Selection)
        If Table Is Nothing Then Exit Sub
        Call Table.AddRow(InsertBelow:=Selection)
    End If
End Sub

Public Sub DeleteRows()
    If ActiveWorkbook.Name <> ThisWorkbook.Name Then Exit Sub
    If Not MDeveloper.Mode Then
        Dim Table As CTable
        Set Table = GetTable(Selection)
        If Table Is Nothing Then Exit Sub
        Call Table.DeleteRows(CellsInTable:=Selection)
    End If
End Sub

Private Function GetTable(ByRef CurrentCell As Range) As CTable
    Dim Table As CTable
    Set Table = New CTable
    Dim Name As Name
    For Each Name In CurrentCell.Worksheet.Names
        If IsTable(Name) Then
            If Not Intersect(Name.RefersToRange, CurrentCell) Is Nothing Then
                Call Table.Bind(Name)
                Set GetTable = Table
                Exit Function
            End If
        End If
    Next Name
    Set GetTable = Nothing
End Function

Private Function IsTable(ByRef Name As Name) As Boolean
    Let IsTable = False
    On Error GoTo ErrorHandling
    If MRegex.Test(Name.Name, "Table_") Then Let IsTable = True
ErrorHandling:
End Function
