Attribute VB_Name = "Main"
Option Explicit
Option Compare Text
Option Base 1
DefLng A-Z

Public Const OptionsSection = "Options"

Public App As New CApp

Public Sub AutoOpen()
    'Dim s As String
    'On Error Resume Next
    's = App.Path
    'App.DefaultOptions("Send") = s & "SMail\Mailbox\Bank\S\"
    'App.DefaultOptions("Recv") = s & "SMail\Mailbox\Bank\R\"
    'InitMenuBars
    'User.ID = ActiveSheet.Name
    'Crypto.Password = vbNullString
    'DoEvents
    'CheckRecv
End Sub

Public Sub AutoClose()
    'On Error Resume Next
    'CloseMenuBars
    'ActiveWindow.Caption = Empty
    'Workbooks(App.BookName).Saved = True 'no more prompts!
    'With Application
    '    .DisplayAlerts = False 'no more prompts!
        'DoEvents
        '.Quit
    'End With
    'Workbooks(App.BookName).Close
    'DoEvents
    'With Application
    '    .DisplayAlerts = True
    '    .Caption = .Application
    '    InfoBox "%d books, active %s", Workbooks.Count, ActiveWorkbook.Name
    '    If Workbooks.Count <= 1 Then
    '        .Quit
    '    End If
    'End With
End Sub
