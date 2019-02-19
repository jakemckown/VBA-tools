VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FDeveloper 
   Caption         =   "Developer"
   ClientHeight    =   1692
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   1932
   OleObjectBlob   =   "FDeveloper.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FDeveloper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Requires:    FDeveloper.frx
'              MDeveloper.bas
'              MRuntime.bas

Option Explicit

Private Sub CheckBox_DeveloperMode_Click()
    Call MDeveloper.SetMode(CheckBox_DeveloperMode.Value)
End Sub

Private Sub CommandButton_Exit_Click()
    Call Unload(Me)
End Sub

Private Sub UserForm_Initialize()
    With Application
        Let Me.Left = .Left + (.Width - Me.Width) / 2
        Let Me.Top = .Top + (.Height - Me.Height) / 2
    End With
    Call MRuntime.Save
    Call MRuntime.Fast
    With CheckBox_DeveloperMode
        Let .Value = MDeveloper.Mode
        Call .SetFocus
    End With
    Call MRuntime.Normal
End Sub
