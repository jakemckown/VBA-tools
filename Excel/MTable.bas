Attribute VB_Name = "MTable"
' Requires:    MDeveloper.bas
'              MRegex.bas
'              CTable.cls

Option Explicit

Private Const BUTTON_ADD_ROW_TEXT As String = "+"
Private Const BUTTON_DELETE_ROW_TEXT As String = "–"

Public Enum tBorderWeightEnum
    tBorderWeightNone = 0
    tHairline = 1
    tMedium = -4138
    tThick = 4
    tThin = 2
End Enum

Public Enum tLineStyleEnum
    tContinuous = 1
    tDash = -4115
    tDashDot = 4
    tDashDotDot = 5
    tDot = -4118
    tDouble = -4119
    tLineStyleNone = -4142
    tSlantDashDot = 13
End Enum

Public Sub AddRow()
    Call AddRow_ContinuousHairlineBorder
End Sub

Public Sub AddRow_SpecifyBorder(ByVal LineStyle As tLineStyleEnum, _
                                ByVal BorderWeight As tBorderWeightEnum)
    If ActiveWorkbook.Name <> ThisWorkbook.Name Then Exit Sub
    If Not MDeveloper.Mode Then
        Dim Table As CTable
        Set Table = GetTable(Selection)
        If Table Is Nothing Then
            Dim Message As String
            Let Message = vbNullString
            Let Message = Message & "To add a row to the table, click on a cell in the row "
            Let Message = Message & "you wish to insert the new row below, then click the "
            Let Message = Message & BUTTON_ADD_ROW_TEXT & " button."
            Call MsgBox(Message, vbOKOnly + vbInformation, "How to add a row")
            Exit Sub
        End If
        Call Table.AddRow(InsertBelow:=Selection, LineStyle:=LineStyle, BorderWeight:=BorderWeight)
    End If
End Sub

Public Sub AddRow_ContinuousHairlineBorder()
    Call AddRow_SpecifyBorder(LineStyle:=tContinuous, BorderWeight:=tHairline)
End Sub

Public Sub AddRow_ContinuousThinBorder()
    Call AddRow_SpecifyBorder(LineStyle:=tContinuous, BorderWeight:=tThin)
End Sub

Public Sub AddRow_NoBorder()
    Call AddRow_SpecifyBorder(LineStyle:=tLineStyleNone, BorderWeight:=tBorderWeightNone)
End Sub

Public Sub DeleteRows()
    If ActiveWorkbook.Name <> ThisWorkbook.Name Then Exit Sub
    If Not MDeveloper.Mode Then
        Dim Table As CTable
        Set Table = GetTable(Selection)
        If Table Is Nothing Then
            Dim Message As String
            Let Message = vbNullString
            Let Message = Message & "To delete rows from the table, click on cells in "
            Let Message = Message & "one or more row in the table, then click the "
            Let Message = Message & BUTTON_DELETE_ROW_TEXT & " button."
            Call MsgBox(Message, vbOKOnly + vbInformation, "How to delete rows")
            Exit Sub
        End If
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
