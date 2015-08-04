Attribute VB_Name = "PGP"
Option Explicit
Option Compare Text
Option Base 1
DefLng A-Z

Public Sub PGPRun(Cmd As String, FileName1 As String, Optional FileName2 As String = "")
    Dim PID As Variant, PGPCmd As String
    If Len(FileName2) = 0 Then
        FileName2 = FileName1
    End If
    PGPCmd = Cmd
    PGPCmd = StrTran(PGPCmd, "%1", FileName1)
    PGPCmd = StrTran(PGPCmd, "%2", FileName2)
    PID = Shell(PGPCmd, vbNormalFocus)
    If PID = 0 Then
        MsgBox BPrintF("Команда не выполнима!\n~%s~", PGPCmd), vbCritical, AppTitle
        Exit Sub
    End If
    'DoEvents
    'On Error GoTo WaitShell
    'AppActivate PID, True
    'MsgBox "Команда выполнена", vbInformation, AppTitle
    'Exit Sub

'WaitShell:
    'DoEvents
    'If FileExists(FileName2) Then
    '    Resume Next
    'Else
    '    Resume
    'End If
End Sub
