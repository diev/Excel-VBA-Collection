Attribute VB_Name = "Main"
'ВНИМАНИЕ:

'Исходные тексты системы Банк-Клиент ЗАО "Сити Инвест Банк"
'предоставляются клиентам в открытом виде "как есть" на момент
'поставки для настройки и адаптации на компьютере Клиента.

'Также они могут быть интересны студентам и возможным работодателям.
'Любое другое их использование НЕ РАЗРЕШАЕТСЯ в соответствии
'с Законом об охране авторских прав Российской Федерации.
'Вам оказано большое доверие для удобства Вашей работы!

'Автор программы: Дмитрий Евдокимов, 1995-1999-2005
'Адрес в сети Интернет: http://diev.narod.ru/
'E-mail: diev@mail.ru, ICQ: 7372116

'Любые контакты и предложения очень приветствуются!
'Спасибо!

Option Explicit
Option Compare Text
Option Base 1
DefLng A-Z

Public Const OptionsSection = "Options"
'Public Const FilesSection = "Files"
'Public Const PGPSection = "PGP"
'Public Const SMailSection = "SMail"

Public App As New CApp
Public BnkSeek2 As New CBnkSeek2
Public User As New CUser
Public Payment As New CPayment
Public Crypto As New CCrypto
Public SMail As New CSMail

Public Sub AutoOpen()
    'Dim s As String
    On Error Resume Next
    's = App.Path
    'App.DefaultOptions("Send") = s & "SMail\Mailbox\Bank\S\"
    'App.DefaultOptions("Recv") = s & "SMail\Mailbox\Bank\R\"
    'InitMenuBars
    User.ID = ActiveSheet.Name
    Crypto.Password = vbNullString
    DoEvents
    CheckRecv
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

