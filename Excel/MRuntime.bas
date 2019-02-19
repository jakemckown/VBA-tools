Attribute VB_Name = "MRuntime"
#If VBA7 Then
    Public Declare PtrSafe Sub Delay Lib "kernel32" Alias "Sleep" (ByVal Milliseconds As LongPtr)
#Else
    Public Declare Sub Delay Lib "kernel32" Alias "Sleep" (ByVal Milliseconds As Long)
#End If

Option Explicit

Private pCalculation As XlCalculation
Private pDisplayAlerts As Boolean
Private pEnableEvents As Boolean
Private pScreenUpdating As Boolean
Private pStartTime As Date
Private pStopTime As Date

Public Enum rtDisplayTypeEnum
    rtCurrentTime
    rtElapsedTime
End Enum

Public Sub Display(ByVal DisplayType As rtDisplayTypeEnum, _
          Optional ByVal Debugging As Boolean = False)
    Select Case DisplayType
        Case rtCurrentTime
            If Debugging Then
                Debug.Print "Current time: " & Now
            Else
                Call MsgBox( _
                    Prompt:="Current time: " & Now, _
                    Buttons:=vbOKOnly + vbInformation, _
                    Title:="Current Time" _
                )
            End If
        Case rtElapsedTime
            If Debugging Then
                Debug.Print "Elapsed time: " & DateDiff("s", pStartTime, pStopTime) & " s"
            Else
                Call MsgBox( _
                    Prompt:="Elapsed time: " & DateDiff("s", pStartTime, pStopTime) & " s", _
                    Buttons:=vbOKOnly + vbInformation, _
                    Title:="Elapsed Time" _
                )
            End If
        Case Else
    End Select
End Sub

Public Sub Fast(Optional ByVal DisplayAlerts As Boolean = False, _
                Optional ByVal EnableEvents As Boolean = False)
    With Application
        Let .Calculation = xlCalculationManual
        Let .DisplayAlerts = DisplayAlerts
        Let .EnableEvents = EnableEvents
        Let .ScreenUpdating = False
    End With
End Sub

Public Sub Normal(Optional ByVal DisplayAlerts As Boolean = True, _
                  Optional ByVal EnableEvents As Boolean = True)
    With Application
        Let .Calculation = xlCalculationAutomatic
        Let .DisplayAlerts = DisplayAlerts
        Let .EnableEvents = EnableEvents
        Let .ScreenUpdating = True
    End With
End Sub

Public Sub Reset()
    Call Normal
End Sub

Public Sub Restore()
    With Application
        Let .Calculation = pCalculation
        Let .DisplayAlerts = pDisplayAlerts
        Let .EnableEvents = pEnableEvents
        Let .ScreenUpdating = pScreenUpdating
    End With
End Sub

Public Sub Save()
    With Application
        Let pCalculation = .Calculation
        Let pDisplayAlerts = .DisplayAlerts
        Let pEnableEvents = .EnableEvents
        Let pScreenUpdating = .ScreenUpdating
    End With
End Sub

Public Function StartClock() As Date
    Let pStartTime = Now
    Let StartClock = pStartTime
End Function

Public Function StopClock() As Date
    Let pStopTime = Now
    Let StopClock = pStopTime
End Function
