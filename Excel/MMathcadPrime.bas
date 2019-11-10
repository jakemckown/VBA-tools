Attribute VB_Name = "MMathcadPrime"
' https://www.ptcusercommunity.com/thread/60547

Option Explicit

Private Const FILE_PATH As String = "C:\Users\mckownj\OneDrive - kochind.com\Documents\R&D\Mathcad\Calc.mcdx"

Public Sub GetResults()
    Dim Mathcad As Ptc_MathcadPrime_Automation.Application
    Set Mathcad = New Ptc_MathcadPrime_Automation.Application
    Dim Worksheet As Ptc_MathcadPrime_Automation.Worksheet
    Set Worksheet = Mathcad.Open(FILE_PATH)
    Dim Outputs As Ptc_MathcadPrime_Automation.Outputs
    Set Outputs = Worksheet.Outputs
    Dim i As Long, Alias As String, StringValue As String, OutputResult As Ptc_MathcadPrime_Automation.OutputResult
    For i = 1 To Outputs.Count
        Let Alias = Outputs.GetAliasByIndex(i - 1)
        Let StringValue = Worksheet.OutputGetStringValue(Alias)
    Next i
    If StringValue <> "" Then
        Debug.Print Alias & " = " & StringValue
    Else
        Set OutputResult = Worksheet.OutputGetRealValue(Alias)
        Debug.Print Alias & " = " & OutputResult.RealResult & " " & OutputResult.Units
    End If
    Call Mathcad.CloseAll(SaveOption_spDiscardChanges)
    Call Mathcad.Quit(SaveOption_spDiscardChanges)
End Sub
