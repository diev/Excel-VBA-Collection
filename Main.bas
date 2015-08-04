Attribute VB_Name = "Main"
'ВНИМАНИЕ:

'Исходные тексты системы Банк-Клиент ЗАО "Сити Инвест Банк"
'предоставляются клиентам в открытом виде "как есть" на момент
'поставки для настройки и адаптации на компьютере Клиента.

'Также они могут быть интересны студентам и возможным работодателям.
'Любое другое их использование НЕ РАЗРЕШАЕТСЯ в соответствии
'с Законом об охране авторских прав Российской Федерации.
'Вам оказано большое доверие для удобства Вашей работы!

'Автор программы: Дмитрий Евдокимов, 1995-1999-2000
'Адрес в сети Интернет: http://members.xoom.com/diev/
'E-mail: diev@mail.ru, ICQ: 7372116

'Любые контакты и предложения очень приветствуются!
'Спасибо!

Option Explicit
Option Compare Text
Option Base 1
DefLng A-Z

Public Const BnkSeek2Section = "BnkSeek2"
Public Const PGPSection = "PGP"
Public Const SMailSection = "SMail"

Public App As CApp
Public BnkSeek2 As CBnkSeek2
Public User As CUser
Public Payment As CPayment
Public PGP As CPGP
Public SMail As CSMail

Public Sub AutoOpen()
    Dim s As String
    On Error Resume Next
    s = ChangeFileExt(ActiveWorkbook.FullName, "INI")
    If StrToBool(ReadIniFile(s, "Options", _
        "DontAutoOpen")) Then
        MsgBox "Выполнение прекращено!", vbCritical, "Блокировка запуска"
        Exit Sub
    End If
    Application.StatusBar = "Инициализация программы..."
    Set App = New CApp
    Set BnkSeek2 = New CBnkSeek2
    Set User = New CUser
    Set Payment = New CPayment
    Set PGP = New CPGP
    Set SMail = New CSMail
    InitMenuBars
    Application.StatusBar = False
End Sub

Public Sub AutoClose()
    Dim s As String
    On Error Resume Next
    'User.Demo = True
    CloseMenuBars
    s = "В ЦЕЛЯХ БЕЗОПАСНОСТИ ЗАВЕРШИТЕ РАБОТУ С MICROSOFT EXCEL!"
    'ActiveWindow.Caption = Empty
    ActiveWorkbook.Saved = True 'no more prompts!
    With Application
        .DisplayAlerts = False 'no more prompts!
        .Caption = s
        .StatusBar = s
        DoEvents
        .Quit
        DoEvents
        'how are we still here?!
        MsgBox s, vbCritical, App.Title
    End With
End Sub

Public Sub LogonShow()
    On Error Resume Next
    Application.StatusBar = "Вход в систему Банк-Клиент..."
    AutoCheck
    Logon.Show
    Application.StatusBar = False
End Sub

Public Sub PlatEnterShow()
    On Error Resume Next
    Application.StatusBar = "Создание платежного поручения..."
    AutoCheck
    PlatEnter.Show
    Application.StatusBar = False
End Sub

Public Sub NewUserShow()
    On Error Resume Next
    Application.StatusBar = "Добавление нового плательщика..."
    AutoCheck
    Load NewName
    With NewName
        .Mode = "New"
        .Show
    End With
    Application.StatusBar = False
End Sub

Public Sub EditUserShow()
    On Error Resume Next
    Application.StatusBar = "Изменение реквизитов плательщика..."
    AutoCheck
    Load NewName
    With NewName
        .Mode = "Edit"
        .Show
    End With
    Application.StatusBar = False
End Sub

Public Sub BICShow()
    On Error Resume Next
    Application.StatusBar = "Работа со Справочником банков РФ"
    AutoCheck
    Load NewName
    With NewName
        .Mode = "BIC"
        .Show
    End With
    Application.StatusBar = False
End Sub

Public Sub LSShow()
    On Error Resume Next
    Application.StatusBar = "Работа со Справочником банков РФ"
    AutoCheck
    Load NewName
    With NewName
        .Mode = "LS"
        .Show
    End With
    Application.StatusBar = False
End Sub

Public Sub PayUserShow()
    On Error Resume Next
    Application.StatusBar = "Ввод реквизитов получателя платежа..."
    AutoCheck
    Load NewName
    With NewName
        .Mode = "Pay"
        .Show
    End With
    Application.StatusBar = False
End Sub

Public Sub UserPrivateShow()
    On Error Resume Next
    Application.StatusBar = "Ключи и частные настройки клиента..."
    AutoCheck
    UserPrivate.Show
    Application.StatusBar = False
End Sub

Public Sub MailBoxShow()
    On Error Resume Next
    Application.StatusBar = "Работа с почтовыми ящиками программы SMail..."
    AutoCheck
    MailBox.Show
    Application.StatusBar = False
End Sub

Public Sub SavePassShow()
    Application.StatusBar = "Сохранение, смена пароля и выход..."
    AutoCheck
    SaveAsPass.Show
    Application.StatusBar = False
    If App.CloseAllowed Then AutoClose
End Sub

Public Sub ExportPlat()
    On Error Resume Next
    Application.StatusBar = "Шифрование PGP и отправка документов в ящик SMail..."
    AutoCheck
    Payment.EachSelected "ExportPlat", "Отправить все?"
    Application.StatusBar = False
End Sub

Public Sub ExportList()
    On Error Resume Next
    Application.StatusBar = "Выгрузка на диск (экспорт)..."
    AutoCheck
    Payment.ExportToFile
    Application.StatusBar = False
End Sub

Public Sub ImportList()
    On Error Resume Next
    Application.StatusBar = "Загрузка с диска (импорт)..."
    AutoCheck
    Payment.ImportFromFile
    Application.StatusBar = False
End Sub

Public Sub DelRows()
    On Error Resume Next
    Application.StatusBar = "Удаление строк..."
    AutoCheck
    Payment.EachSelected "Delete", "Безвозвратно удалить?"
    Application.StatusBar = False
    If Payment.MoneyLastSelected > 0 Then AddDeletedAmount
End Sub

Public Sub DelUser()
    On Error Resume Next
    Application.StatusBar = "Удаление плательщика..."
    AutoCheck
    With User
        If .Demo Then
            MsgBox "Демонстрационного клиента удалять нельзя!", vbExclamation, App.Title
        ElseIf MsgBox(BPrintF("Действительно удалить из системы\nклиента \'%s\' - %s\nи его ключи шифрования?", _
            .ID, .Name), vbExclamation + vbYesNo, App.Title) = vbYes Then
            .Delete
        End If
    End With
    Application.StatusBar = False
End Sub

Public Sub AutoCheck()
    On Error Resume Next
    If Val(User.BIC) = 0 Then
        MsgBox "Извините, программа перезапустится!", vbExclamation, App.Title
        Restart
    End If
End Sub

Public Sub AutoRestart(Optional s As String = vbNullString)
    On Error Resume Next
    MsgBox BPrintF("Извините, программа перезапустится!\n%s", s), vbExclamation, App.Title
    AutoOpen
End Sub

Public Sub Restart()
    On Error Resume Next
    With App
        If MsgBox("Поискать новые компоненты?", vbQuestion + vbYesNo, .Title) = vbYes Then
            .Setting(BnkSeek2Section, "LocateCanceled") = 0
            .Setting(PGPSection, "LocateCanceled") = 0
            .Setting(SMailSection, "LocateCanceled") = 0
            BnkSeek2.File = vbNullString
            PGP.File = vbNullString
            SMail.File = vbNullString
        End If
    End With
    AutoOpen
End Sub

Public Sub PreviewPlat()
    On Error Resume Next
    Application.StatusBar = "Просмотр и печать поручений..."
    AutoCheck
    Payment.EachSelected "Preview", "Показать все?"
    Application.StatusBar = False
End Sub

Public Sub Info()
    On Error Resume Next
    Application.StatusBar = "Информация о программе..."
    AutoCheck
    App.Info
    Application.StatusBar = False
End Sub

Public Sub FindText()
    On Error Resume Next
    Application.StatusBar = "Поиск текста..."
    AutoCheck
    Payment.FindText
    Application.StatusBar = False
End Sub

'Public Sub FindNext()
'    On Error Resume Next
'    Application.StatusBar = "Поиск текста позже..."
'    AutoCheck
'    Payment.FindNext
'    Application.StatusBar = False
'End Sub
'
'Public Sub FindPrev()
'    On Error Resume Next
'    Application.StatusBar = "Поиск текста раньше..."
'    AutoCheck
'    Payment.FindPrev
'    Application.StatusBar = False
'End Sub

Public Sub SortByDocNo()
    On Error Resume Next
    Application.StatusBar = "Сортировка строк по номеру..."
    AutoCheck
    Payment.SortBy 3
    Application.StatusBar = False
End Sub

Public Sub SortByDocDate()
    On Error Resume Next
    Application.StatusBar = "Сортировка строк по дате..."
    AutoCheck
    Payment.SortBy 4
    Application.StatusBar = False
End Sub

Public Sub SortBySum()
    On Error Resume Next
    Application.StatusBar = "Сортировка строк по сумме..."
    AutoCheck
    Payment.SortBy 5
    Application.StatusBar = False
End Sub

Public Sub SortByName()
    On Error Resume Next
    Application.StatusBar = "Сортировка строк по получателю..."
    AutoCheck
    Payment.SortBy 6
    Application.StatusBar = False
End Sub

Public Sub SortByDetails()
    On Error Resume Next
    Application.StatusBar = "Сортировка строк по назначению..."
    AutoCheck
    Payment.SortBy 11
    Application.StatusBar = False
End Sub

Public Sub AmountChange()
    Dim s As String, c As Currency
    On Error Resume Next
    Application.StatusBar = "Изменение текущего остатка средств..."
    AutoCheck
    c = User.Amount
    s = RSumStr(c, vbCrLf)
    If c < 0 Then s = "минус! " & s
    s = BPrintF("Сейчас на счету %f\n(%s)\n\nВведите новый остаток:", c, s)
    'If c < 0 Then s = s & BPrintF("\nТЕКУЩИЙ МЕНЬШЕ НУЛЯ!")
    s = s & BPrintF("\n\n(Используйте \'+\' и \'-\', чтобы добавить к остатку\nили отнять)")
    s = InputBox(s, App.Title, PlatFormat(c))
    If Len(s) > 0 Then
        c = RVal(s)
        If Left(s, 1) = "+" Then
            c = User.Amount + c
        ElseIf c < 0 Then
            c = User.Amount + c 'negative value!
        End If
        s = RSumStr(c, vbCrLf)
        If c < 0 Then s = "минус! " & s
        If MsgBox(BPrintF("Поставить текущий остаток %f?\n(%s)", c, s), vbQuestion + vbYesNo, App.Title) = vbYes Then
            User.Amount = c
        End If
    End If
    Application.StatusBar = False
End Sub

Public Sub AddDeletedAmount()
    Dim s As String, c As Currency
    On Error Resume Next
    Application.StatusBar = "Возврат средств после удаления..."
    AutoCheck
    c = User.Amount
    s = RSumStr(c, vbCrLf)
    If c < 0 Then s = "минус! " & s
    s = BPrintF("Сейчас на счету %f\n(%s)\n\nДобавить удаленную сумму %f?", c, s, _
        Payment.MoneyLastSelected)
    'If c < 0 Then s = s & BPrintF("\nТЕКУЩИЙ МЕНЬШЕ НУЛЯ!")
    s = InputBox(s, App.Title, BPrintF("+%F", Payment.MoneyLastSelected))
    If Len(s) > 0 Then
        c = RVal(s)
        If Left(s, 1) = "+" Then
            c = User.Amount + c
        ElseIf c < 0 Then
            c = User.Amount + c 'negative value!
        End If
        s = RSumStr(c, vbCrLf)
        If c < 0 Then s = "минус! " & s
        If MsgBox(BPrintF("Поставить текущий остаток %f?\n(%s)", c, s), vbQuestion + vbYesNo, App.Title) = vbYes Then
            User.Amount = c
        End If
    End If
    Application.StatusBar = False
End Sub


