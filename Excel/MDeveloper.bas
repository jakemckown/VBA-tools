Attribute VB_Name = "MDeveloper"
' Requires:    FDeveloper.frm
'              FDeveloper.frx
'              MRuntime.bas

Option Explicit

Private DEVELOPER_MODE As Boolean

Private Const DISPLAY_GRIDLINES_ENABLED As Boolean = True

Public Function Mode() As Boolean
    Let Mode = DEVELOPER_MODE
End Function

Public Sub SetMode(Optional ByVal Value As Boolean = False)
    Let DEVELOPER_MODE = Value
    Select Case DEVELOPER_MODE
        Case True
            Call UnprotectAllWorksheets
            If DISPLAY_GRIDLINES_ENABLED Then Call DisplayGridlines
        Case False
            If DISPLAY_GRIDLINES_ENABLED Then Call DisplayGridlines(False)
            Call ProtectAllWorksheets
            Call MRuntime.Normal
    End Select
End Sub

Public Sub ShowForm()
Attribute ShowForm.VB_ProcData.VB_Invoke_Func = "D\n14"
    Call FDeveloper.Show
End Sub

Private Sub DisplayGridlines(Optional ByVal Value As Boolean = True)
    Call MRuntime.Save
    Call MRuntime.Fast
    Dim CurrentSheet As Worksheet
    Set CurrentSheet = ActiveSheet
    On Error GoTo ErrorHandling
    Dim Worksheet As Worksheet
    For Each Worksheet In ThisWorkbook.Worksheets
        Call Worksheet.Activate
        Let ActiveWindow.DisplayGridlines = Value
    Next Worksheet
ErrorHandling:
    Call CurrentSheet.Activate
    Call MRuntime.Restore
End Sub

Private Sub ProtectAllWorksheets()
    Dim Worksheet As Worksheet
    For Each Worksheet In ThisWorkbook.Worksheets
        Call Worksheet.Protect(UserInterfaceOnly:=True)
    Next Worksheet
End Sub

Private Sub UnprotectAllWorksheets()
    Dim Worksheet As Worksheet
    For Each Worksheet In ThisWorkbook.Worksheets
        Call Worksheet.Unprotect
    Next Worksheet
End Sub
