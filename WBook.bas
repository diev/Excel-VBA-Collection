Attribute VB_Name = "WBook"
Option Explicit
Option Compare Text
Option Base 1
DefLng A-Z

Public Sub WorkbookBeforeClose(Cancel As Boolean)
    On Error Resume Next
    'If Not StrToBool(App.Options("DontConfirmExit")) Then
    '    If Not ActiveWorkbook.Saved Then _
    '        Cancel = MsgBox("Закрыть программу?", vbQuestion + vbOKCancel, App.Title) = vbCancel
    'End If
    'If Not Cancel Then AutoClose
    If App.CloseAllowed Then
        Cancel = False
        AutoClose
    Else
        Cancel = True
        SavePassShow
    End If
End Sub

Public Sub WorkbookOpen()
    AutoOpen
End Sub

Public Sub WorkbookSheetActivate(ByVal Sh As Object)
    Dim s As String
    On Error Resume Next
    If Sh.Name = User.ID Then 'Simple reactivate...
        'Exit Sub
    ElseIf Sh.Name = User.DemoID Then
        User.Demo = True
    ElseIf Len(Sh.Name) = 3 Then
        User.ID = Sh.Name
        LogonShow
        'Load Logon
        'With Logon
        '    .txtID = User.ID
        '    .txtPass.SetFocus
        '    .Show
        'End With
    End If
    'ActiveWindow.Caption = Sh.Name
End Sub

Public Sub WorksheetBeforeRightClick(ByVal Target As Range, Cancel As Boolean)
    '
End Sub

