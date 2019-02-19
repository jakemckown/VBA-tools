Attribute VB_Name = "MSecurity"
Option Explicit

Public Sub UnprotectWorksheet()
    Dim i1 As Integer, i2 As Integer, i3 As Integer, i4 As Integer, i5 As Integer, i6 As Integer, i7 As Integer, i8 As Integer, i9 As Integer, i10 As Integer, i11 As Integer, i12 As Integer
    Dim Password As String
    On Error Resume Next
    For i1 = 65 To 66
        For i2 = 65 To 66
            For i3 = 65 To 66
                For i4 = 65 To 66
                    For i5 = 65 To 66
                        For i6 = 65 To 66
                            For i7 = 65 To 66
                                For i8 = 65 To 66
                                    For i9 = 65 To 66
                                        For i10 = 65 To 66
                                            For i11 = 65 To 66
                                                For i12 = 32 To 126
                                                    Let Password = Chr(i1) & Chr(i2) & Chr(i3) & Chr(i4) & Chr(i5) & Chr(i6) & Chr(i7) & Chr(i8) & Chr(i9) & Chr(i10) & Chr(i11) & Chr(i12)
                                                    Call ActiveSheet.Unprotect(Password)
                                                    If ActiveSheet.ProtectContents = False Then
                                                        Call MsgBox("Password """ & Password & """ was successful.")
                                                        Exit Sub
                                                    End If
                                                Next
                                            Next
                                        Next
                                    Next
                                Next
                            Next
                        Next
                    Next
                Next
            Next
        Next
    Next
End Sub
