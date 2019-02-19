Attribute VB_Name = "MRegex"
Option Explicit

Public Function GetMatches(ByVal SourceString As String, _
                           ByVal Pattern As String, _
                  Optional ByVal FindAll As Boolean = True) As Collection
    'Returns a Collection with the following structure:
    '                   1                       2      3              4              ...  SubMatches.Count + 2
    '              [
    '            1   [  PositionInSourceString  Match  SubMatches(1)  SubMatches(2)  ...  SubMatches(SubMatches.Count)  ]
    '            2   [  PositionInSourceString  Match  SubMatches(1)  SubMatches(2)  ...  SubMatches(SubMatches.Count)  ]
    '            .
    '            .
    '            .
    'Matches.Count   [  PositionInSourceString  Match  SubMatches(1)  SubMatches(2)  ...  SubMatches(SubMatches.Count)  ]
    '              ]
    Dim RegExp As Object
    Set RegExp = CreateObject("VBScript.RegExp")
    Let RegExp.Pattern = Pattern
    Let RegExp.Global = FindAll
    Dim Matches As Object
    Set Matches = RegExp.Execute(SourceString)
    Dim MatchCollection As Collection
    Set MatchCollection = New Collection
    Dim i As Long, j As Long, SubMatches As Object, SubMatchCollection As Collection
    For i = 1 To Matches.Count
        Set SubMatchCollection = New Collection
        With SubMatchCollection
            Call .Add(Matches(i - 1).FirstIndex + 1)
            Call .Add(Matches(i - 1).Value)
            Set SubMatches = Matches.Item(i - 1).SubMatches
            For j = 1 To SubMatches.Count
                Call .Add(SubMatches(j - 1))
            Next j
        End With
        Call MatchCollection.Add(SubMatchCollection)
    Next i
    Set GetMatches = MatchCollection
End Function

Public Function GetSubMatch(ByVal SourceString As String, _
                            ByVal Pattern As String, _
                   Optional ByVal Match As Long = 1, _
                   Optional ByVal Index As Long = 1) As String
    Let GetSubMatch = vbNullString
    On Error GoTo ErrorHandling
    Dim Matches As Collection
    Set Matches = GetMatches(SourceString, Pattern)
    Let GetSubMatch = Matches(Match)(Index + 2)
ErrorHandling:
End Function

Public Function Render(ByVal SourceString As String) As String
    Let Render = SourceString
    On Error GoTo ErrorHandling
    Dim RenderedString As String
    Dim Character As String
    Let RenderedString = vbNullString
    Dim i As Long
    For i = 1 To Len(SourceString)
        Let Character = Mid$(String:=SourceString, Start:=i, Length:=1)
        If RequiresEscape(Character) Then
            Let RenderedString = RenderedString & "\" & Character
        Else
            Let RenderedString = RenderedString & Character
        End If
    Next i
    Let Render = RenderedString
ErrorHandling:
End Function

Private Function RequiresEscape(ByVal Character As String) As Boolean
    Let RequiresEscape = False
    On Error GoTo ErrorHandling
    If Len(Character) > 1 Then Exit Function
    Select Case Character
        Case "\", "^", "$", ".", "|", "?", "*", "+", "(", ")", "["
            Let RequiresEscape = True
        Case Else
    End Select
ErrorHandling:
End Function

Public Function Test(ByVal SourceString As String, _
                     ByVal Pattern As String) As Boolean
    Dim RegExp As Object
    Set RegExp = CreateObject("VBScript.RegExp")
    Let RegExp.Pattern = Pattern
    Let Test = RegExp.Test(SourceString)
End Function
