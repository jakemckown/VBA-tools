Attribute VB_Name = "MFormula"
' Requires:    MRegex.bas
'              MRuntime.bas

Option Explicit

'Private Const FORMULA_EXPRESSION_PATTERN As String = "^???$"
Private Const FORMULA_FUNCTION_PATTERN As String = "^(\w+)\((.*)\)$"

Private FormattedFormula As String

Public Function FormatFormula(ByVal Formula As String, _
                     Optional ByVal IndentSize As Long = 4) As String
    Let FormatFormula = Formula
    On Error GoTo ErrorHandling
    If VBA.Left$(Formula, 1) <> "=" Then Exit Function
    Let FormattedFormula = "="
    If VBA.Mid$(Formula, 2, 1) <> Chr$(10) Then Let FormattedFormula = FormattedFormula & Chr$(10)
    Dim Level As Long
    Let Level = 0
    Let FormattedFormula = FormattedFormula & FormatFormulaValue(Trim$(Mid$(Formula, 2)), IndentSize, Level)
    Let FormatFormula = FormattedFormula
ErrorHandling:
    Let FormattedFormula = vbNullString
End Function

Private Function FormatFormulaExpression(ByVal FormulaExpression As String, _
                                         ByVal IndentSize As Long, _
                                         ByVal Level As Long) As String
    Let FormatFormulaExpression = FormulaExpression
    On Error GoTo ErrorHandling
    Dim Indent As String
    Let Indent = VBA.Space$(IndentSize * Level)
    Dim FormattedExpression As String
    Let FormattedExpression = Indent & FormulaExpression
    Let FormatFormulaExpression = FormattedExpression
ErrorHandling:
End Function

Private Function FormatFormulaFunction(ByVal FormulaFunction As String, _
                                       ByVal IndentSize As Long, _
                                       ByVal Level As Long) As String
    Let FormatFormulaFunction = FormulaFunction
    On Error GoTo ErrorHandling
    Dim Matches As Collection
    Set Matches = MRegex.GetMatches(SourceString:=FormulaFunction, Pattern:=FORMULA_FUNCTION_PATTERN)
    Dim FunctionName As String
    Let FunctionName = Matches(1)(3)
    Dim FunctionArguments As Collection
    Set FunctionArguments = GetFunctionArguments(Matches(1)(4))
    Dim Indent As String
    Let Indent = VBA.Space$(IndentSize * Level)
    Dim FormattedFunction As String
    Let FormattedFunction = Indent & FunctionName & "("
    If FunctionArguments.Count > 0 Then
        Let FormattedFunction = FormattedFunction & Chr$(10)
        Dim i As Long
        For i = 1 To FunctionArguments.Count
            Let FormattedFunction = FormattedFunction & FormatFormulaValue(FunctionArguments(i), IndentSize, Level + 1)
            If i < FunctionArguments.Count Then Let FormattedFunction = FormattedFunction & ","
            Let FormattedFunction = FormattedFunction & Chr$(10)
        Next i
        Let FormattedFunction = FormattedFunction & Indent
    End If
    Let FormattedFunction = FormattedFunction & ")"
    Let FormatFormulaFunction = FormattedFunction
ErrorHandling:
End Function

Private Function FormatFormulaValue(ByVal FormulaValue As String, _
                                    ByVal IndentSize As Long, _
                                    ByVal Level As Long) As String
    Let FormatFormulaValue = FormulaValue
    On Error GoTo ErrorHandling
    If IsFormulaFunction(FormulaValue) Then
        Let FormatFormulaValue = FormatFormulaFunction(FormulaValue, IndentSize, Level)
    ElseIf IsFormulaExpression(FormulaValue) Then
        Let FormatFormulaValue = FormatFormulaExpression(FormulaValue, IndentSize, Level)
    Else
    End If
ErrorHandling:
End Function

Public Sub FormatSelectedFormula()
Attribute FormatSelectedFormula.VB_ProcData.VB_Invoke_Func = "F\n14"
    Dim Cell As Range
    Set Cell = Selection
    Call MRuntime.Save
    Call MRuntime.Fast
    Let Cell.Formula = FormatFormula(Cell.Formula)
    Call MRuntime.Restore
End Sub

Private Function GetFunctionArguments(ByVal ArgumentsString As String) As Collection
    Dim FunctionArguments As Collection
    Set FunctionArguments = New Collection
    On Error GoTo ErrorHandling
    If ArgumentsString <> vbNullString Then
        Dim Level As Long
        Let Level = 0
        Dim Start As Long
        Let Start = 1
        Dim i As Long
        For i = 1 To Len(ArgumentsString)
            Select Case VBA.Mid$(ArgumentsString, i, 1)
                Case "("
                    Let Level = Level + 1
                    
                Case ")"
                    Let Level = Level - 1
                Case ","
                    If Level = 0 Then
                        Call FunctionArguments.Add(VBA.Trim$(VBA.Mid$(ArgumentsString, Start, i - Start)))
                        Let Start = i + 1
                    End If
                Case Else
            End Select
            If i = Len(ArgumentsString) Then Call FunctionArguments.Add(VBA.Trim$(VBA.Mid$(ArgumentsString, Start)))
        Next i
    End If
ErrorHandling:
    Set GetFunctionArguments = FunctionArguments
End Function

Private Function IsFormulaExpression(ByVal FormulaValue As String) As Boolean
    Let IsFormulaExpression = True
    On Error GoTo ErrorHandling
'    Let IsFormulaExpression = MRegex.Test(FormulaValue, FORMULA_EXPRESSION_PATTERN)
ErrorHandling:
End Function

Private Function IsFormulaFunction(ByVal FormulaValue As String) As Boolean
    Let IsFormulaFunction = False
    On Error GoTo ErrorHandling
    Let IsFormulaFunction = MRegex.Test(FormulaValue, FORMULA_FUNCTION_PATTERN)
ErrorHandling:
End Function
