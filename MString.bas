Attribute VB_Name = "MString"
Option Explicit

Public Function Capitalize(ByVal Word As String) As String
    Let Capitalize = Word
    Let Word = Trim$(Word)
    Dim FirstChar As String
    Let FirstChar = Left$(Word, 1)
    If Asc(FirstChar) < 97 Or Asc(FirstChar) > 122 Then Exit Function
    Let FirstChar = Chr$(Asc(FirstChar) - 32)
    Let Capitalize = FirstChar & Mid$(Word, 2)
End Function

Public Function TitleCase(ByVal Text As String, _
                 Optional ByVal CapitalizeAll As Boolean = False) As String
    Let Text = LowerCase(Trim$(Text))
    Dim Words() As String
    Let Words = Split(Text)
    Dim NewString As String
    Let NewString = vbNullString
    Dim i As Long, Word As String, NewWord As String
    For i = 0 To UBound(Words)
        Let Word = Words(i)
        If Word <> vbNullString Then
            Select Case Word
                Case "a", "an", "and", "at", "but", "by", "for", "from", "in", "of", "or", "the", "to", "with"
                    If CapitalizeAll Or i = UBound(Words) Then
                        Let NewWord = Capitalize(Word)
                    Else
                        Let NewWord = Word
                    End If
                Case "lp", "llc"
                    Let NewWord = UpperCase(Word)
                Case Else
                    Let NewWord = Capitalize(Word)
            End Select
            Let NewString = NewString & NewWord
            If i <> UBound(Words) Then Let NewString = NewString & " "
        End If
    Next i
    Let TitleCase = NewString
End Function

Public Function LowerCase(ByVal Text As String) As String
    Let LowerCase = LCase$(Text)
End Function

Public Function UpperCase(ByVal Text As String) As String
    Let UpperCase = UCase$(Text)
End Function
