Attribute VB_Name = "MString"
Option Explicit

Public Const ASCII_LETTERS As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz"
Public Const ASCII_LOWERCASE As String = "abcdefghijklmnopqrstuvwxyz"
Public Const ASCII_UPPERCASE As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
Public Const DECIMAL_DIGITS As String = "0123456789"
Public Const HEXADECIMAL_DIGITS As String = "0123456789abcdefABCDEF"
Public Const OCTAL_DIGITS As String = "01234567"

Public Function Capitalize(ByVal Word As String) As String
    Let Capitalize = Word
    Let Word = Trim$(Word)
    Dim FirstChar As String
    Let FirstChar = Left$(Word, 1)
    If Asc(FirstChar) < 97 Or Asc(FirstChar) > 122 Then Exit Function
    Let FirstChar = Chr$(Asc(FirstChar) - 32)
    Let Capitalize = FirstChar & Mid$(Word, 2)
End Function

Public Function Concatenate(ByRef Strings As Variant, _
                   Optional ByVal Delimiter As String = vbNullString) As String
    Dim Concatenated As String
    Let Concatenated = vbNullString
    On Error GoTo ErrorHandling
    Dim i As Long
    Select Case VBA.VarType(Strings)
        Case Is >= VBA.vbArray
            For i = 0 To UBound(Strings)
                Let Concatenated = Concatenated & CStr(Strings(i))
                If i < UBound(Strings) Then Let Concatenated = Concatenated & Delimiter
            Next i
        Case VBA.vbString
            Let Concatenated = Strings
        Case VBA.vbObject
            If VBA.TypeName(Strings) = "Collection" Then
                For i = 1 To Strings.Count
                    Let Concatenated = Concatenated & CStr(Strings(i))
                    If i < Strings.Count Then Let Concatenated = Concatenated & Delimiter
                Next i
            End If
        Case Else
    End Select
ErrorHandling:
    Let Concatenate = Concatenated
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
