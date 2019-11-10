Attribute VB_Name = "MFileIO"
Option Explicit

' The JSON code below is modified from https://github.com/VBA-tools/VBA-JSON

Public Enum fioIOModeEnum
    fioForAppending = 8
    fioForReading = 1
    fioForWriting = 2
End Enum

Public Enum fioTristateEnum
    fioTristateFalse = 0
    fioTristateMixed = -2
    fioTristateTrue = -1
    fioTristateUseDefault = -2
End Enum

Public Function ConvertToJSON(ByVal JSONValue As Variant, _
                     Optional ByVal Whitespace As Variant, _
                     Optional ByVal CurrentIndentation As Long = 0) As String
    Dim Buffer As String
    Let Buffer = vbNullString
    Dim BufferPosition As Long
    Let BufferPosition = 0
    Dim BufferLength As Long
    Let BufferLength = 0
    Dim IsFirstItem As Boolean
    Let IsFirstItem = True
    Dim PrettyPrint As Boolean
    Let PrettyPrint = Not IsMissing(Whitespace)
    Select Case VBA.VarType(JSONValue)
        Case VBA.vbNull
            Let ConvertToJSON = "null"
        Case VBA.vbString
            Let ConvertToJSON = """" & jsonEncode(JSONValue) & """"
        Case VBA.vbBoolean
            If JSONValue Then
                Let ConvertToJSON = "true"
            Else
                Let ConvertToJSON = "false"
            End If
        Case VBA.vbObject
            Dim Indentation As String
            If PrettyPrint Then
                If VBA.VarType(Whitespace) = VBA.vbString Then
                    Let Indentation = VBA.String$(CurrentIndentation + 1, Whitespace)
                Else
                    Let Indentation = VBA.Space$((CurrentIndentation + 1) * Whitespace)
                End If
            End If
            Dim Converted As String
            If VBA.TypeName(JSONValue) = "Dictionary" Then
                Call jsonBufferAppend(Buffer, "{", BufferPosition, BufferLength)
                Dim Key As Variant, SkipItem As Boolean
                For Each Key In JSONValue.Keys
                    Let Converted = ConvertToJSON(JSONValue(Key), Whitespace, CurrentIndentation + 1)
                    If Converted = vbNullString Then
                        Let SkipItem = jsonIsUndefined(JSONValue(Key))
                    Else
                        Let SkipItem = False
                    End If
                    If Not SkipItem Then
                        If IsFirstItem Then
                            Let IsFirstItem = False
                        Else
                            Call jsonBufferAppend(Buffer, ",", BufferPosition, BufferLength)
                        End If
                        If PrettyPrint Then
                            Let Converted = vbNewLine & Indentation & """" & Key & """: " & Converted
                        Else
                            Let Converted = """" & Key & """:" & Converted
                        End If
                        Call jsonBufferAppend(Buffer, Converted, BufferPosition, BufferLength)
                    End If
                Next Key
                If PrettyPrint Then
                    Call jsonBufferAppend(Buffer, vbNewLine, BufferPosition, BufferLength)
                    If VBA.VarType(Whitespace) = VBA.vbString Then
                        Let Indentation = VBA.String$(CurrentIndentation, Whitespace)
                    Else
                        Let Indentation = VBA.Space$(CurrentIndentation * Whitespace)
                    End If
                End If
                Call jsonBufferAppend(Buffer, Indentation & "}", BufferPosition, BufferLength)
            ElseIf VBA.TypeName(JSONValue) = "Collection" Then
                Call jsonBufferAppend(Buffer, "[", BufferPosition, BufferLength)
                Dim Value As Variant
                For Each Value In JSONValue
                    If IsFirstItem Then
                        Let IsFirstItem = False
                    Else
                        Call jsonBufferAppend(Buffer, ",", BufferPosition, BufferLength)
                    End If
                    Let Converted = ConvertToJSON(Value, Whitespace, CurrentIndentation + 1)
                    If Converted = vbNullString And jsonIsUndefined(Value) Then Let Converted = "null"
                    If PrettyPrint Then Let Converted = vbNewLine & Indentation & Converted
                    Call jsonBufferAppend(Buffer, Converted, BufferPosition, BufferLength)
                Next Value
                If PrettyPrint Then
                    Call jsonBufferAppend(Buffer, vbNewLine, BufferPosition, BufferLength)
                    If VBA.VarType(Whitespace) = VBA.vbString Then
                        Let Indentation = VBA.String$(CurrentIndentation, Whitespace)
                    Else
                        Let Indentation = VBA.Space$(CurrentIndentation * Whitespace)
                    End If
                End If
                Call jsonBufferAppend(Buffer, Indentation & "]", BufferPosition, BufferLength)
            End If
            Let ConvertToJSON = jsonBufferToString(Buffer, BufferPosition)
        Case VBA.vbInteger, VBA.vbLong, VBA.vbSingle, VBA.vbDouble, VBA.vbCurrency, VBA.vbDecimal
            Let ConvertToJSON = VBA.Replace(JSONValue, ",", ".")
        Case Else
            On Error Resume Next
            ConvertToJSON = JSONValue
            On Error GoTo 0
    End Select
End Function

Public Function ConvertToXML(ByVal XMLDocument As Object) As String
    Let ConvertToXML = XMLDocument.XML
End Function

Public Function OpenFile(ByVal FilePath As String, _
                Optional ByVal IOMode As fioIOModeEnum = fioForReading, _
                Optional ByVal Unicode As Boolean = False) As Object
    Dim FileSystem As Object
    Set FileSystem = CreateObject("Scripting.FileSystemObject")
    Dim TextStream As Object
    If FileSystem.FileExists(FilePath) Then
        Dim File As Object
        Set File = FileSystem.GetFile(FilePath)
        Dim Format As fioTristateEnum
        If Unicode Then
            Let Format = fioTristateTrue
        Else
            Let Format = fioTristateFalse
        End If
        Set TextStream = File.OpenAsTextStream(IOMode:=IOMode, Format:=Format)
    Else
        Set TextStream = FileSystem.CreateTextFile(Filename:=FilePath, Overwrite:=True, Unicode:=Unicode)
    End If
    Set OpenFile = TextStream
End Function

Public Function ParentPath(ByVal Path As String, _
                  Optional ByVal Level As Long = 1) As String
    If Level < 1 Then
        Let ParentPath = Path
        Exit Function
    End If
    On Error GoTo ErrorHandling
    Dim i As Long
    For i = 1 To Level
        Let Path = Left$(Path, InStrRev(Path, "\") - 1)
    Next i
ErrorHandling:
    Let ParentPath = Path
End Function

Public Function ParseJSON(ByVal JSON As String) As Object
    Dim Index As Long
    Let Index = 1
    Let JSON = VBA.Replace(VBA.Replace(VBA.Replace(JSON, VBA.vbCr, vbNullString), VBA.vbLf, vbNullString), VBA.vbTab, vbNullString)
    Call jsonSkipSpaces(JSON, Index)
    Select Case VBA.Mid$(JSON, Index, 1)
        Case "{"
            Set ParseJSON = jsonParseObject(JSON, Index)
        Case "["
            Set ParseJSON = jsonParseArray(JSON, Index)
        Case Else
            Call Err.Raise(Number:=10001, Source:="JSONConverter", Description:=jsonParseErrorMessage(JSON, Index, "Expecting: { or ["))
    End Select
End Function

Public Function ParseJSONFromFile(ByVal FilePath As String) As Object
    Dim File As Object
    Set File = OpenFile(FilePath:=FilePath, IOMode:=fioForReading)
    If File Is Nothing Then
        Set ParseJSONFromFile = Nothing
        Exit Function
    End If
    On Error GoTo ErrorHandling
    Dim JSON As String
    Let JSON = File.ReadAll
    Set ParseJSONFromFile = ParseJSON(JSON)
ErrorHandling:
    Call File.Close
End Function

Public Function ParseXML(ByVal XML As String) As Object
    Dim Document As Object
    Set Document = CreateObject("MSXML2.DOMDocument")
    Call Document.LoadXML(XML)
    Set ParseXML = Document
End Function

Public Function ParseXMLFromFile(ByVal FilePath As String) As Object
    Dim Document As Object
    Set Document = CreateObject("MSXML2.DOMDocument")
    Call Document.Load(FilePath)
    Set ParseXMLFromFile = Document
End Function

Public Function ReadAllFromFile(ByVal FilePath As String) As String
    Dim File As Object
    Set File = OpenFile(FilePath)
    On Error GoTo ErrorHandling
    Let ReadAllFromFile = File.ReadAll
ErrorHandling:
    Call File.Close
End Function

Sub Test()
    Dim FilePath As String
    Let FilePath = "K:\Calcs MFG07 PERM\R&D�WIP\Reports\Project Monitor\Queries\JobInfo.sql"
    Debug.Print ReadAllFromFile(FilePath)
End Sub

Public Sub SaveJSONValueToFile(ByRef JSONValue As Object, _
                               ByVal FilePath As String, _
                      Optional ByVal Whitespace As Variant)
    Dim File As Object
    Set File = OpenFile(FilePath:=FilePath, IOMode:=fioForWriting)
    If File Is Nothing Then Exit Sub
    On Error GoTo ErrorHandling
    If IsMissing(Whitespace) Then
        Call File.Write(ConvertToJSON(JSONValue:=JSONValue))
    Else
        Call File.Write(ConvertToJSON(JSONValue:=JSONValue, Whitespace:=Whitespace))
    End If
ErrorHandling:
    Call File.Close
End Sub

Public Sub SaveXMLDocumentToFile(ByRef XMLDocument As Object, _
                                 ByVal FilePath As String)
    Call XMLDocument.Save(FilePath)
End Sub

Private Sub jsonBufferAppend(ByRef Buffer As String, _
                             ByRef Append As Variant, _
                             ByRef BufferPosition As Long, _
                             ByRef BufferLength As Long)
    Dim AppendLength As Long
    Let AppendLength = VBA.Len(Append)
    Dim LengthPlusPosition As Long
    Let LengthPlusPosition = AppendLength + BufferPosition
    If LengthPlusPosition > BufferLength Then
        Dim AddedLength As Long
        Let AddedLength = IIf(AppendLength > BufferLength, AppendLength, BufferLength)
        Let Buffer = Buffer & VBA.Space$(AddedLength)
        Let BufferLength = BufferLength + AddedLength
    End If
    Mid$(Buffer, BufferPosition + 1, AppendLength) = CStr(Append)
    Let BufferPosition = BufferPosition + AppendLength
End Sub

Private Function jsonBufferToString(ByRef Buffer As String, _
                                    ByVal BufferPosition As Long) As String
    If BufferPosition > 0 Then Let jsonBufferToString = VBA.Left$(Buffer, BufferPosition)
End Function

Private Function jsonEncode(ByVal JSON As Variant) As String
    Dim Buffer As String
    Let Buffer = vbNullString
    Dim BufferPosition As Long
    Let BufferPosition = 0
    Dim BufferLength As Long
    Let BufferLength = 0
    Dim Index As Long, Char As String, AscCode As Long
    For Index = 1 To VBA.Len(JSON)
        Let Char = VBA.Mid$(JSON, Index, 1)
        Let AscCode = VBA.AscW(Char)
        If AscCode < 0 Then Let AscCode = AscCode + 65536
        Select Case AscCode
            Case 34
                Let Char = "\"""
            Case 92
                Let Char = "\\"
            Case 47
                Let Char = "\/"
            Case 8
                Let Char = "\b"
            Case 12
                Let Char = "\f"
            Case 10
                Let Char = "\n"
            Case 13
                Let Char = "\r"
            Case 9
                Let Char = "\t"
            Case 0 To 31, 127 To 65535
                Let Char = "\u" & VBA.Right$("0000" & VBA.Hex$(AscCode), 4)
        End Select
        Call jsonBufferAppend(Buffer, Char, BufferPosition, BufferLength)
    Next Index
    Let jsonEncode = jsonBufferToString(Buffer, BufferPosition)
End Function

Private Function jsonIsUndefined(ByVal JSONValue As Variant) As Boolean
    Select Case VBA.VarType(JSONValue)
        Case VBA.vbEmpty
            Let jsonIsUndefined = True
        Case VBA.vbObject
            Select Case VBA.TypeName(JSONValue)
                Case "Empty", "Nothing"
                    Let jsonIsUndefined = True
            End Select
    End Select
End Function

Private Function jsonParseArray(ByRef JSON As String, _
                                ByRef Index As Long) As Collection
    Set jsonParseArray = New Collection
    Call jsonSkipSpaces(JSON, Index)
    If VBA.Mid$(JSON, Index, 1) <> "[" Then
        Call Err.Raise(Number:=10001, Source:="JSONConverter", Description:=jsonParseErrorMessage(JSON, Index, "Expecting: ["))
    Else
        Let Index = Index + 1
        Do
            Call jsonSkipSpaces(JSON, Index)
            If VBA.Mid$(JSON, Index, 1) = "]" Then
                Let Index = Index + 1
                Exit Function
            ElseIf VBA.Mid$(JSON, Index, 1) = "," Then
                Let Index = Index + 1
                Call jsonSkipSpaces(JSON, Index)
            End If
            Call jsonParseArray.Add(jsonParseValue(JSON, Index))
        Loop
    End If
End Function

Private Function jsonParseErrorMessage(ByRef JSON As String, _
                                       ByRef Index As Long, _
                                       ByRef ErrorMessage As String)
    Dim StartIndex As Long
    Let StartIndex = Index - 10
    If StartIndex <= 0 Then Let StartIndex = 1
    Dim StopIndex As Long
    Let StopIndex = Index + 10
    If StopIndex > VBA.Len(JSON) Then Let StopIndex = VBA.Len(JSON)
    Dim Message As String
    Let Message = vbNullString
    Let Message = Message & "Error parsing JSON:" & VBA.vbNewLine
    Let Message = Message & VBA.Mid$(JSON, StartIndex, StopIndex - StartIndex + 1) & VBA.vbNewLine
    Let Message = Message & VBA.Space$(Index - StartIndex) & "^" & VBA.vbNewLine
    Let Message = Message & ErrorMessage
    Let jsonParseErrorMessage = Message
End Function

Private Function jsonParseKey(ByRef JSON As String, _
                              ByRef Index As Long) As String
    If VBA.Mid$(JSON, Index, 1) = """" Then
        Let jsonParseKey = jsonParseString(JSON, Index)
    Else
        Call Err.Raise(Number:=10001, Source:="JSONConverter", Description:=jsonParseErrorMessage(JSON, Index, "Expecting: """))
    End If
    Call jsonSkipSpaces(JSON, Index)
    If VBA.Mid$(JSON, Index, 1) <> ":" Then
        Call Err.Raise(Number:=10001, Source:="JSONConverter", Description:=jsonParseErrorMessage(JSON, Index, "Expecting: :"))
    Else
        Let Index = Index + 1
    End If
End Function

Private Function jsonParseObject(ByRef JSON As String, _
                                 ByRef Index As Long) As Object
    Dim Key As String
    Dim NextChar As String
    Set jsonParseObject = CreateObject("Scripting.Dictionary")
    Call jsonSkipSpaces(JSON, Index)
    If VBA.Mid$(JSON, Index, 1) <> "{" Then
        Call Err.Raise(Number:=10001, Source:="JSONConverter", Description:=jsonParseErrorMessage(JSON, Index, "Expecting: {"))
    Else
        Let Index = Index + 1
        Do
            Call jsonSkipSpaces(JSON, Index)
            If VBA.Mid$(JSON, Index, 1) = "}" Then
                Let Index = Index + 1
                Exit Function
            ElseIf VBA.Mid$(JSON, Index, 1) = "," Then
                Let Index = Index + 1
                Call jsonSkipSpaces(JSON, Index)
            End If
            Let Key = jsonParseKey(JSON, Index)
            Let NextChar = jsonPeek(JSON, Index)
            If NextChar = "[" Or NextChar = "{" Then
                Set jsonParseObject.Item(Key) = jsonParseValue(JSON, Index)
            Else
                Let jsonParseObject.Item(Key) = jsonParseValue(JSON, Index)
            End If
        Loop
    End If
End Function

Private Function jsonParseNumber(ByRef JSON As String, _
                                 ByRef Index As Long) As Double
    Dim Char As String
    Dim Value As String
    Call jsonSkipSpaces(JSON, Index)
    Do While Index > 0 And Index <= Len(JSON)
        Let Char = VBA.Mid$(JSON, Index, 1)
        If VBA.InStr("+-0123456789.eE", Char) Then
            Let Value = Value & Char
            Let Index = Index + 1
        Else
            Let jsonParseNumber = VBA.Val(Value)
            Exit Function
        End If
    Loop
End Function

Private Function jsonParseString(ByRef JSON As String, _
                                 ByRef Index As Long) As String
    Dim Quote As String
    Dim Char As String
    Dim Code As String
    Dim Buffer As String
    Dim BufferPosition As Long
    Dim BufferLength As Long
    Call jsonSkipSpaces(JSON, Index)
    Let Quote = VBA.Mid$(JSON, Index, 1)
    Let Index = Index + 1
    Do While Index > 0 And Index <= Len(JSON)
        Let Char = VBA.Mid$(JSON, Index, 1)
        Select Case Char
            Case "\"
                Let Index = Index + 1
                Let Char = VBA.Mid$(JSON, Index, 1)
                Select Case Char
                    Case """", "\", "/", "'"
                        Call jsonBufferAppend(Buffer, Char, BufferPosition, BufferLength)
                        Let Index = Index + 1
                    Case "b"
                        Call jsonBufferAppend(Buffer, vbBack, BufferPosition, BufferLength)
                        Let Index = Index + 1
                    Case "f"
                        Call jsonBufferAppend(Buffer, vbFormFeed, BufferPosition, BufferLength)
                        Let Index = Index + 1
                    Case "n"
                        Call jsonBufferAppend(Buffer, vbCrLf, BufferPosition, BufferLength)
                        Let Index = Index + 1
                    Case "r"
                        Call jsonBufferAppend(Buffer, vbCr, BufferPosition, BufferLength)
                        Let Index = Index + 1
                    Case "t"
                        Call jsonBufferAppend(Buffer, vbTab, BufferPosition, BufferLength)
                        Let Index = Index + 1
                    Case "u"
                        Let Index = Index + 1
                        Let Code = VBA.Mid$(JSON, Index, 4)
                        Call jsonBufferAppend(Buffer, VBA.ChrW$(VBA.Val("&h" + Code)), BufferPosition, BufferLength)
                        Let Index = Index + 4
                End Select
            Case Quote
                Let jsonParseString = jsonBufferToString(Buffer, BufferPosition)
                Let Index = Index + 1
                Exit Function
            Case Else
                Call jsonBufferAppend(Buffer, Char, BufferPosition, BufferLength)
                Let Index = Index + 1
        End Select
    Loop
End Function

Private Function jsonParseValue(ByRef JSON As String, _
                                ByRef Index As Long) As Variant
    Call jsonSkipSpaces(JSON, Index)
    Select Case VBA.Mid$(JSON, Index, 1)
        Case "{"
            Set jsonParseValue = jsonParseObject(JSON, Index)
        Case "["
            Set jsonParseValue = jsonParseArray(JSON, Index)
        Case """"
            Let jsonParseValue = jsonParseString(JSON, Index)
        Case Else
            If VBA.Mid$(JSON, Index, 4) = "true" Then
                Let jsonParseValue = True
                Let Index = Index + 4
            ElseIf VBA.Mid$(JSON, Index, 5) = "false" Then
                Let jsonParseValue = False
                Let Index = Index + 5
            ElseIf VBA.Mid$(JSON, Index, 4) = "null" Then
                Let jsonParseValue = Null
                Let Index = Index + 4
            ElseIf VBA.InStr("+-0123456789", VBA.Mid$(JSON, Index, 1)) Then
                Let jsonParseValue = jsonParseNumber(JSON, Index)
            Else
                Call Err.Raise(Number:=10001, Source:="JSONConverter", Description:=jsonParseErrorMessage(JSON, Index, "Expecting: String, Number, null, true, false, {, or ["))
            End If
    End Select
End Function

Private Function jsonPeek(ByRef JSON As String, _
                          ByVal Index As Long, _
                 Optional ByRef NumberOfCharacters As Long = 1) As String
    Call jsonSkipSpaces(JSON, Index)
    Let jsonPeek = VBA.Mid$(JSON, Index, NumberOfCharacters)
End Function

Private Sub jsonSkipSpaces(ByRef JSON As String, _
                           ByRef Index As Long)
    Do While Index > 0 And Index <= VBA.Len(JSON) And VBA.Mid$(JSON, Index, 1) = " "
        Let Index = Index + 1
    Loop
End Sub
