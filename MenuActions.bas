Attribute VB_Name = "MenuActions"
Option Explicit
Option Compare Text
Option Base 1
DefLng A-Z

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
    Application.StatusBar = "Работа со Справочником БИК"
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
    Application.StatusBar = "Работа со Справочником БИК"
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
            InfoBox "Демонстрационного клиента удалять нельзя!"
        ElseIf YesNoBox("Действительно удалить из системы\nклиента %s - %s\nи его ключи шифрования?", _
            .ID, .Name) Then
            .Delete
        End If
    End With
    Application.StatusBar = False
End Sub

Public Sub AutoCheck()
    On Error Resume Next
    If ActiveSheet.Name <> User.ID Then
        Workbooks(App.BookName).Activate
    ElseIf Val(User.BIC) = 0 Then
        WarnBox "Извините, программа перезапустится!"
        Restart
    End If
End Sub

Public Sub AutoRestart(Optional s As String = vbNullString)
    On Error Resume Next
    WarnBox "Извините, программа перезапустится!\n%s", s
    AutoOpen
End Sub

Public Sub Restart()
    On Error Resume Next
    With App
        If YesNoBox("Поискать новые компоненты после перезапуска?") Then
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
    s = Bsprintf("Сейчас на счету %f\n(%s)\n\nВведите новый остаток:", c, s)
    'If c < 0 Then s = s & BSPrintF("\nТЕКУЩИЙ МЕНЬШЕ НУЛЯ!")
    s = s & Bsprintf("\n\n(Используйте \'+\' и \'-\', чтобы добавить к остатку\nили отнять)")
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
        If YesNoBox("Поставить текущий остаток %f?\n(%s)", c, s) Then
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
    s = Bsprintf("Сейчас на счету %f\n(%s)\n\nДобавить удаленную сумму %f?", c, s, _
        Payment.MoneyLastSelected)
    'If c < 0 Then s = s & BSPrintF("\nТЕКУЩИЙ МЕНЬШЕ НУЛЯ!")
    s = InputBox(s, App.Title, Bsprintf("+%F", Payment.MoneyLastSelected))
    If Len(s) > 0 Then
        c = RVal(s)
        If Left(s, 1) = "+" Then
            c = User.Amount + c
        ElseIf c < 0 Then
            c = User.Amount + c 'negative value!
        End If
        s = RSumStr(c, vbCrLf)
        If c < 0 Then s = "минус! " & s
        If YesNoBox("Поставить текущий остаток %f?\n(%s)", c, s) Then
            User.Amount = c
        End If
    End If
    Application.StatusBar = False
End Sub

Public Sub UpdateReceived(File As String)
    If OkCancelBox("ПРОЧТИТЕ ВНИМАТЕЛЬНО ПОРЯДОК ОБНОВЛЕНИЯ:\n\n" & _
        "1. После нажатия кнопки \'OK\' программа Excel сохранит\n" & _
        "этот рабочий файл %s и закроется.\n\n" & _
        "2. Будет запущен самораспаковывающийся файл\n" & _
        "автоматического обновления %s\n" & _
        "для рабочей директории %s.\n\n" & _
        "3. Только убедившись, что Excel закрылся, нажмите кнопку \'OK\'\n" & _
        "в запущенном обновлении. По его завершении можно будет\n" & _
        "открыть снова и продолжить работу с программой Банк-Клиент.\n\n" & _
        "4. После обновления файл следует удалить из принятого.\n\n" & _
        "Закрыть Excel и запустить обновление прямо сейчас?", _
        ActiveWorkbook.FullName, File, App.Path) Then
        
        ActiveWorkbook.Save
        App.CloseAllowed = True
        Shell File & " /auto " & App.Path, vbNormalNoFocus
        AutoClose
    End If
End Sub


Public Sub PrintVeksel()
    Dim OldDate As Date, s As String
    On Error Resume Next
    OldDate = Date
    If YesNoBox("Печать векселя за " & Format(Now - 1, "dddd dd.MM.yyyy HH:mm")) Then
        Date = Date - 1
    Else
        s = Format(Date - 1, "dd.MM.yy")
        s = InputBox("Введите дату для векселя:", App.Title, s)
        If Len(s) = 0 Then Exit Sub
        Date = CDate(s)
    End If
    DoEvents
    
    With Payment
        .ReadRow
        s = .FileName & "." & User.ID
        s = RightPathName(GetWinTempDir, s)
        If .SavePlat(s) Then
            DoEvents
            Time = Time + 0.001
            DoEvents
            ShellWait App.Path & "prnveksl.exe " & s
            DoEvents
        Else
            StopBox "Файл для печати НЕ создан!"
        End If
    End With
    Date = OldDate
    Time = Time - 0.001
    DoEvents
End Sub

