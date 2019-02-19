Attribute VB_Name = "MCheckbox"
Option Explicit

Private Const CHECKBOX_FONT_FILE_PATH As String = "C:\Windows\Fonts\webdings.ttf"
Private Const CHECKBOX_FONT_NAME As String = "Webdings"
Public Const CHECKBOX_OFF As String = "c"
Public Const CHECKBOX_ON As String = "g"

Public Sub Click(ByVal Box As Range, _
        Optional ByVal SetValue As Variant, _
        Optional ByVal RowOffset As Long = 0, _
        Optional ByVal ColumnOffset As Long = -1)
    If Not IsMissing(SetValue) Then
        Let Box.Value = SetValue
    Else
        If IsOn(Box) Then
            Let Box.Value = CHECKBOX_OFF
        Else
            Let Box.Value = CHECKBOX_ON
        End If
    End If
    If ActiveCell.Address = Box.Address Then Call Box.Offset(RowOffset, ColumnOffset).Select
End Sub

Public Function IsCheckbox(ByVal Target As Range) As Boolean
    Let IsCheckbox = False
    On Error GoTo ErrorHandling
    If VBA.VarType(Target.Value) = vbString _
    And Target.Font.Name = CHECKBOX_FONT_NAME _
    And (Target.Value = CHECKBOX_ON Or Target.Value = CHECKBOX_OFF) Then Let IsCheckbox = True
ErrorHandling:
End Function

Public Function IsOn(ByVal Box As Range) As Boolean
    Let IsOn = False
    On Error GoTo ErrorHandling
    Select Case Box.Value
        Case CHECKBOX_ON
            Let IsOn = True
        Case CHECKBOX_OFF
            Let IsOn = False
    End Select
ErrorHandling:
End Function
