Attribute VB_Name = "Main"
'ВНИМАНИЕ:

'Исходные тексты системы Банк-Клиент ЗАО "Сити Инвест Банк"
'предоставляются клиентам в открытом виде "как есть" на момент
'поставки для настройки и адаптации на компьютере Клиента.

'Также они могут быть интересны студентам и возможным работодателям.
'Любое другое их использование НЕ РАЗРЕШАЕТСЯ в соответствии
'с Законом об охране авторских прав Российской Федерации.
'Вам оказано большое доверие для удобства Вашей работы!

'Автор программы: Дмитрий Евдокимов, 1995-1999-2001
'Адрес в сети Интернет: http://members.xoom.com/diev/
'E-mail: diev@mail.ru, ICQ: 7372116

'Любые контакты и предложения очень приветствуются!
'Спасибо!

Option Explicit
Option Compare Text
Option Base 1
DefLng A-Z

Public Const BClient2Section = "BClient2"
Public Const BnkSeek2Section = "BnkSeek2"
Public Const PGPSection = "PGP"
Public Const SMailSection = "SMail"

Public App As New CApp
Public BnkSeek2 As New CBnkSeek2
Public User As New CUser
Public Payment As New CPayment
Public PGP As New CPGP
Public SMail As New CSMail

Public Sub AutoOpen()
    'MsgBox because no App here at start!
    Dim s As String
    On Error Resume Next
    
    Application.StatusBar = "Начальная загрузка и проверка условий..."
    If StrToBool(App.Options("DontAutoOpen")) Then
        MsgBox "Выполнение прекращено!", vbCritical, "Блокировка запуска"
        Exit Sub
    End If
    
    If Not IsDir(App.Options("WorkPath")) Then
        MsgBox Bsprintf("Это не Ваша рабочая книга!\n\nДоступ к файлу %s\nразрешен только из %s", App.BookFile, s), vbCritical, "Блокировка запуска"
        Exit Sub
    End If
    
    Application.StatusBar = "Загрузка модулей программы - ждите завершения..."
    InitMenuBars
    Application.StatusBar = "Демонстрационный режим"
    User.Demo = True
    
    PGP.ResetPasswords
    
    s = "Добро пожаловать в демонстрационный режим программы!\n" & _
        "Здесь Вы можете посмотреть программу без какого-либо риска.\n\n" & _
        "Или переключитесь на Ваш рабочий лист, если Вы готовы полноценно\n" & _
        "работать с Вашим счетом и смотреть выписки по нему.\n\n" & _
        "Если листа еще нет, \'Добавьте клиента\' через наше меню \'Банк-Клиент\'\n" & _
        "и НИКОГДА не редактируйте на листах вручную - только через это меню!"
        
    If Workbooks.Count > 1 Then
        s = s & "\n\nВНИМАНИЕ: Обнаружены еще какие-то открытые книги Excel!\n" & _
            "Это может привести к ошибкам в выполнении нашей программы."
    End If
    
    s = s & "\n\nВерсия программы: " & App.Version
    
    InfoHelpBox s, 1
    Application.StatusBar = False
End Sub

Public Sub AutoClose()
    Dim s As String
    On Error Resume Next
    'User.Demo = True
    CloseMenuBars
    s = "В ЦЕЛЯХ БЕЗОПАСНОСТИ ЗАВЕРШИТЕ РАБОТУ С MICROSOFT EXCEL!"
    ActiveWindow.Caption = Empty
    Workbooks(App.BookName).Saved = True 'no more prompts!
    With Application
        .DisplayAlerts = False 'no more prompts!
        .Caption = s
        .StatusBar = s
        DoEvents
        .Quit
    End With
    'Application.Caption = Application.Application
    'Application.StatusBar = False
End Sub

