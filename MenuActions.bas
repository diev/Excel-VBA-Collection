Attribute VB_Name = "MenuActions"
Option Explicit
Option Compare Text
Option Base 1
DefLng A-Z

Public Sub PlatEnterShow()
    On Error Resume Next
    AutoCheck
    PlatEnter.Show
End Sub

Public Sub NewUserShow()
    On Error Resume Next
    AutoCheck
    Load NewName
    With NewName
        .Mode = "New"
        .Show
    End With
End Sub

Public Sub EditUserShow()
    On Error Resume Next
    AutoCheck
    Load NewName
    With NewName
        .Mode = "Edit"
        .Show
    End With
End Sub

Public Sub PayUserShow()
    On Error Resume Next
    AutoCheck
    Load NewName
    With NewName
        .Mode = "Pay"
        .Show
    End With
End Sub

Public Sub MailBoxShow()
    On Error Resume Next
    AutoCheck
    If Not CheckRecv Then
        MailBox.Show
    End If
End Sub

Public Sub ExportPlat()
    On Error Resume Next
    AutoCheck
    Payment.EachSelected "ExportPlat", "Отправить все?"
End Sub

Public Sub ExportList()
    On Error Resume Next
    AutoCheck
    Payment.ExportToFile
End Sub

Public Sub ImportList()
    On Error Resume Next
    AutoCheck
    Payment.ImportFromFile
End Sub

Public Sub DelRows()
    On Error Resume Next
    AutoCheck
    Payment.EachSelected "Delete", "Безвозвратно удалить?"
End Sub

Public Sub AutoCheck()
    On Error Resume Next
    'If ActiveSheet.Name <> User.ID Then
    '    Workbooks(App.BookName).Activate
    'Else
    If Val(User.BIC) = 0 Then
        WarnBox "Извините, программа перезапустится!"
        Restart
    End If
End Sub

Public Sub AutoRestart(Optional s As String = vbNullString)
    On Error Resume Next
    WarnBox "Извините, программа перезапустится!\n%s", s
    Restart
End Sub

Public Sub Restart()
    CloseMenuBars
    AutoOpen
    InitMenuBars
    Crypto.Password = vbNullString
End Sub

Public Sub PreviewPlat()
    On Error Resume Next
    AutoCheck
    Payment.EachSelected "Preview", "Показать все?"
End Sub

Public Sub Info()
    On Error Resume Next
    AutoCheck
    App.Info
End Sub

Public Sub FindText()
    On Error Resume Next
    AutoCheck
    Payment.FindText
End Sub

Public Sub UpdateReceived(File As String)
    Dim File2 As String
    File2 = FilePath(File) & "OK-" & FileNameExt(File)
    If IsFile(File2) Then
        Kill File2
        If IsFile(File2) Then
            StopBox "Удалите предыдущий файл обновления\n%s", File2
            Exit Sub
        End If
    End If
    
    If YesNoBox("ПРОЧТИТЕ ВНИМАТЕЛЬНО ПОРЯДОК ОБНОВЛЕНИЯ:\n\n" & _
        "1. После нажатия этой кнопки 'Да' программа закроется.\n" & _
        "2. ЗАКРОЙТЕ EXCEL, ЕСЛИ ОН НЕ ЗАКРОЕТСЯ САМ!!!\n" & _
        "3. Найдите запущенное обновление и нажмите в нем 'OK'.\n" & _
        "4. Дождитесь окончания его автоматической работы\n" & _
        "(оно очень быстро само помелькает и исчезнет).\n" & _
        "5. Запустите программу Банк-Клиент снова\n" & _
        "(и продолжайте дальше работу как обычно).\n" & _
        "6. Удалите выполненное обновление '%s' из принятого.\n\n" & _
        "Приступить?", File2) Then
        
        App.Options("Version") = App.Version
        App.Options("Updated") = vbNullString
        
        ActiveWorkbook.Save
        Name File As File2
        Shell File2 & " /auto " & App.Path, vbNormalNoFocus
        ActiveWorkbook.Close
        Application.Quit
    End If
End Sub

Public Sub CryptoPassword()
    Crypto.ChangePassword
End Sub

Public Sub ExcelPassword()
    Dim s As String
    On Error Resume Next
    With ActiveWorkbook
        s = InputBox("Введите новый пароль Excel для" & vbCrLf & .FullName, _
            App.TITLE)
        If Len(s) > 0 Then
            Kill .FullName
            DoEvents
            .SaveAs .FullName, Password:=s, CreateBackup:=True
        ElseIf .HasPassword Then
            If YesNoBox("Снять прежний пароль?") Then
                Kill .FullName
                DoEvents
                .SaveAs .FullName, Password:=vbNullString, CreateBackup:=True
            End If
        End If
    End With
End Sub

Public Sub PrintVeksel()
    Dim OldDate As Date, s As String
    On Error Resume Next
    OldDate = Date
    If YesNoBox("Печать векселя за " & Format(Now - 1, "dddd dd.MM.yyyy HH:mm")) Then
        Date = Date - 1
    Else
        s = Format(Date - 1, "dd.MM.yy")
        s = InputBox("Введите дату для векселя:", App.TITLE, s)
        If Len(s) = 0 Then Exit Sub
        Date = CDate(s)
    End If
    DoEvents
    
    With Payment
        .ReadRow
        s = .FileName & "." & User.ID4
        s = RightPathName(GetWinTempDir, s)
        OutputFile s, CDos(.AsPlatINI)
        If IsFile(s) Then
            DoEvents
            Time = Time + 0.001
            DoEvents
            ShellWait App.Path & "prnveksl.exe " & s & " /1", vbMinimizedNoFocus
            DoEvents
        Else
            StopBox "Файл для печати НЕ создан!"
        End If
    End With
    Date = OldDate
    Time = Time - 0.001
    DoEvents
End Sub

Public Sub EditID()
    User.Edit
End Sub

Public Sub SAdm()
    Dim File As String
    On Error Resume Next
    File = App.Path & "SMail\SAdm.exe"
    If IsFile(File) Then
        ChDir App.Path & "SMail"  '/////////////////////////////////sadm's bug?
        Shell QFile(File) & " " & QFile(App.Path & "SMail\SMail.cfg"), vbNormalFocus
    Else
        StopBox "У Вас нет программы\n%s", File
    End If
End Sub

Public Sub SSetup()
    Dim File As String
    On Error Resume Next
    File = App.Path & "SMail\SSetup.exe"
    If IsFile(File) Then
        Shell QFile(File) & " " & QFile(App.Path & "SMail\SMail.ctl"), vbNormalFocus
    Else
        StopBox "У Вас нет программы\n%s", File
    End If
End Sub

Public Sub SMailLog()
    Dim File As String
    On Error Resume Next
    File = App.Path & "SMail\SMail.log"
    If IsFile(File) Then
        Shell "notepad.exe " & QFile(File), vbNormalFocus
    Else
        WarnBox "У Вас нет файла\n%s", File
    End If
End Sub

Public Sub SMailCtl()
    Dim File As String
    On Error Resume Next
    File = App.Path & "SMail\SMail.ctl"
    If IsFile(File) Then
        Shell "notepad.exe " & QFile(File), vbNormalFocus
    Else
        WarnBox "У Вас нет файла\n%s", File
    End If
End Sub

Public Sub OpenFolder()
    On Error Resume Next
    Shell "explorer.exe /e,/root," & QFile(App.Path), vbNormalFocus
End Sub

Public Sub OpenFolderR()
    On Error Resume Next
    Shell "explorer.exe /e,/root," & QFile(SMail.Recv), vbNormalFocus
End Sub

Public Sub AskVypRemart()
    On Error Resume Next
    OutputFile SMail.Send & "ask." & User.ID4, CStr(Now)
    SMail.Dial
End Sub

Public Sub AskBnkSeek()
    On Error Resume Next
    OutputFile SMail.Send & "bnkseek2.ask", CStr(Now)
    SMail.Dial
End Sub

Public Sub AskBClient()
    On Error Resume Next
    OutputFile SMail.Send & "bclient2.ask", CStr(Now)
    SMail.Dial
End Sub

Public Sub ShowHelp()
    Dim File As String
    On Error Resume Next
    File = Dir(App.Path & "*.chm")
    If File = vbNullString Then
        WarnBox "Файлов помощи в %s не найдено!", App.Path
    Else
        Do While File <> vbNullString
            Shell "hh " & QFile(App.Path & File), vbNormalFocus
            File = Dir
        Loop
    End If
End Sub

Public Sub SendFiles()
    Dim Files As Variant, File As String, i As Long
    On Error Resume Next
    Files = App.Path
    If BrowseForFiles(Files) Then
        For i = LBound(Files) To UBound(Files)
            File = CStr(Files(i))
            FileCopy File, SMail.Send & FileNameExt(File)
        Next
    End If
End Sub

Public Sub SendNote()
    Dim File As String
    On Error Resume Next
    File = SMail.Send & "Note" & Format(Now, "yymmdd-hhmm") & ".txt"
    Shell "notepad.exe " & QFile(File), vbNormalFocus
End Sub

Public Function CheckRecv()
    Dim File As String, File1 As String, File2 As String
    Dim s As String, s4 As String, s20 As String, sh As Worksheet
    Dim nOk As Long, nErr As Long, nTest As Long, sOk As String, sErr As String, sTest As String, sRep As String
    On Error Resume Next
    
    CheckRecv = True
    
    If Len(App.Options("Updated")) = 0 Then
        s = App.Options("Version")
        sOk = App.Version
        If Len(s) = 0 Then 'New Install
            InfoBox "Версия программы - %s", sOk
            App.Options("Version") = sOk
            App.Options("Updated") = sOk
        ElseIf s = sOk Then
            WarnBox "Было запущено обновление программы,\nно его, кажется, не произошло!\n" & _
                "Попросите Вашего технического специалиста\nпозвонить в Банк.\n\n" & _
                "Сейчас у Вас версия - %s", sOk
            App.Options("Updated") = sOk
        Else
            InfoBox "Версия программы\nобновлена до %s", sOk
            App.Options("Version") = sOk
            App.Options("Updated") = sOk
        End If
    End If
    
    SMail.Valid
    
    'Antivirus?
    File = SMail.Recv & "!antivir.txt"
    s = "Ваш антивирус блокирует работу нашей программы!"
    OutputFile File, CDos(s)
    If IsFile(File) Then
        Kill File
    Else
        WarnBox s
    End If
    
    'ATTN!
    File = Dir(SMail.Recv & "!*.txt")
    Do While File <> vbNullString
        UrgentMessage SMail.Recv & File
        File = Dir
    Loop
    
    'BnkSeek2 Update
    File = SMail.Recv & "BnkSeek2.exe"
    If IsFile(File) Then
        InfoBox "Получено обновление справочника банков.\nПрограмма установит его и перезапустится."
        ShellWait File & " /auto " & App.Path, vbNormalNoFocus
        DoEvents
        Kill File
        DoEvents
        Restart
        DoEvents
        With BnkSeek2
            .LoadFile
            InfoBox "Установлен справочник БИК от %s, банков: %d", _
                DtoC(.Updated), .RecCount
        End With
    End If
    
    'Done Updates
    With App
        s = .Options("Version")
        If s = .Version Then
            If Len(.Options("Updated")) = 0 Then
                File1 = .Path & ThisWorkbook.Name
                File2 = ThisWorkbook.FullName 'check for same UNC!
                If FileDateTime(File1) = FileDateTime(File2) Then
                    .Options("Updated") = Now
                Else
                    WarnBox "Обновление программы не происходит,\n" & _
                        "т.к. Excel использует первый файл вместо второго:\n\n%s%s", _
                        .FileInfo(File2), .FileInfo(File1)
                End If
            End If
        Else
            .Options("Version") = .Version
            .Options("Updated") = Now
            InfoBox "Программа успешно обновлена до %s", .Version
            File = Dir(SMail.Recv & "OK-*.exe")
            Do While File <> vbNullString
                'If YesNoBox("Обновление '%s' проведено\nУдалить файл?", Mid(File, 4)) Then
                    Kill SMail.Recv & File
                'End If
                File = Dir
            Loop
        End If
    End With
    
    'New Updates
    File = Dir(SMail.Recv & "*.exe")
    Do While File <> vbNullString
        If LCase(Left(File, 3)) <> "ok-" Then
            If YesNoBox("Получено новое обновление %s\n(файл от %s)\n\nВыполнить его сейчас?", _
                File, FileDateTime(SMail.Recv & File)) Then
                UpdateReceived SMail.Recv & File
            End If
        End If
        File = Dir
    Loop
    
    '.ID
    File = Dir(SMail.Recv & "*.id")
    Do While File <> vbNullString
        s20 = FileNameOnly(File)
        If Len(s20) < 20 Then
            s = ReadIniFile(SMail.Recv & File, s20, "BIC", App.DefBIC)
            s20 = ReadIniFile(SMail.Recv & File, s20, "LS", s20)
            s20 = SetLSKey(s, s20)
        End If
        s4 = User.ID4(s20)
        s = App.DefLS(s4)
        If YesNoBox("Получены файлы клиента '%s'\nпо счету %s\n\nДобавить его сейчас?", s4, s20) Then
            'ID
            For Each sh In Worksheets
                If App.DefLS(sh.Name) = s Then
                    If YesNoBox("Такой клиент '%s' уже есть!\nУдалить его из очереди на добавление?", s) Then
                        Kill SMail.Recv & File
                        Kill SMail.Recv & s4 & ".plt"
                        Kill SMail.Recv & s20 & ".plt"
                        GoTo NextID
                    'ElseIf YesNoBox("Лист '%s' будет создан заново!", s) Then
                    '    Sheets(s).Delete
                    'Else
                    '    WarnBox "Хорошо, пропускаем мимо."
                    '    GoTo NextID
                    End If
                    Exit For
                End If
            Next
            'For Each sh In Worksheets
            '    If sh.Name = User.DemoID & " (2)" Then
            '        WarnBox "Необходимо удалить лист '000 (2)'"
            '        Sheets(User.DemoID & " (2)").Delete
            '    End If
            'Next
            File1 = SMail.Recv & File
            File2 = App.Path & File
            If IsFile(File2) Then Kill File2
            Name File1 As File2
            Kill File1
            'Sheets(User.DemoID).Copy before:=Sheets(2)
            'If s <> s20 Then '20
            '    Sheets(User.DemoID & " (2)").Name = s20
                User.ID = s20
            'Else '3-4
            '    Sheets(User.DemoID & " (2)").Name = s4
            '    User.ID = s4
            'End If
            
            'PLT
            Payment.ImportFrom1File SMail.Recv & s4 & ".plt"
            Payment.ImportFrom1File SMail.Recv & s20 & ".plt"
        End If
NextID:
        File = Dir
    Loop
    
    '.PGP
    File = Dir(SMail.Recv & "*.pgp")
    If File <> vbNullString Then
        'If LCase(File) <> "remart.pgp" Then
            Select Case _
                YesNoCancelBox("Получены новые ключи \'%s\' -\nрекомендуется перенести их на дискету.\n\n" & _
                "Да - перенести их на дискету в \'%s\'\n" & _
                "Нет - перенести к программе в \'%s\'\n" & _
                "Отмена - переспросить позже.", File, "A:\Keys\", App.Path & "Keys\")
                Case vbyes:
                    App.Options("KeysPath") = "A:\Keys\"
                    WarnBox "Убедитесь, что дискета вставлена\nи готова к записи!"
                    ForceDirectories App.Options("KeysPath")
                    Do While File <> vbNullString
                        File1 = SMail.Recv & File
                        File2 = App.Options("KeysPath") & File
                        nOk = 1
                        Do While IsFile(File2 & Bsprintf(".%d.old", nOk))
                            nOk = nOk + 1
                        Loop
                        If IsFile(File2) Then Name File2 As File2 & Bsprintf(".%d.old", nOk)
                        If IsFile(File2) Then Kill File2
                        If IsFile(File2) Then
                            WarnBox "Дискета защищена от записи!\nПеренос не выполнен."
                        Else
                            Name File1 As File2
                            If IsFile(File2) Then
                                InfoBox "Файл ключа \'%s\' успешно перенесен", File
                                Kill File1
                            Else
                                WarnBox "Файл ключа \'%s\' НЕ перенесен!", File
                            End If
                        End If
                        File = Dir
                    Loop
                Case vbNo:
                    App.Options("KeysPath") = App.Path & "Keys\"
                    ForceDirectories App.Options("KeysPath")
                    Do While File <> vbNullString
                        File1 = SMail.Recv & File
                        File2 = App.Options("KeysPath") & File
                        nOk = 1
                        Do While IsFile(File2 & Bsprintf(".%d.old", nOk))
                            nOk = nOk + 1
                        Loop
                        If IsFile(File2) Then Name File2 As File2 & Bsprintf(".%d.old", nOk)
                        If IsFile(File2) Then Kill File2
                        If IsFile(File2) Then
                            WarnBox "Директория защищена от записи!\nПеренос не выполнен."
                        Else
                            Name File1 As File2
                            If IsFile(File2) Then
                                InfoBox "Файл ключа \'%s\' успешно перенесен", File
                                Kill File1
                            Else
                                WarnBox "Файл ключа \'%s\' НЕ перенесен!", File
                            End If
                        End If
                        File = Dir
                    Loop
                Case Else
                    'next time
            End Select
        'End If
    End If
    
    'O, E, T...
    With SMail
        DoEvents
        File = Dir(.Recv & "O???????.*")
        nOk = 0: sOk = vbNullString
        Do While Len(File) > 0
            If User.IsID4(FileExt(File)) Then
                'Application.StatusBar = Bsprintf("Файл %s ok", File)
                nOk = nOk + 1: sOk = sOk & Bsprintf("  %d. %s\n", nOk, File)
                File1 = .Recv & File
                File2 = .Archive & File
                If Payment.MarkByFile(File1, "Ok", 2) Then 'Green
                    If IsFile(File2) Then Kill File2
                    Name File1 As File2
                    If Not IsFile(File2) Then
                        If YesNoBox("Файл '%s'\nне удалось переместить в архив\nв '%s'\n\n" & _
                            "Удалить его отсюда?", _
                            File1, .Archive) Then
                            Kill File1
                        End If
                    End If
                End If
            End If
            File = Dir
        Loop
                
        DoEvents
        File = Dir(.Recv & "EYP??*.*") 'удаление ошибочного возврата выписок
        Do While Len(File) > 0
            File1 = .Recv & File
            Kill File1
            File = Dir
        Loop
        
        DoEvents
        File = Dir(.Recv & "ESK.*") 'удаление ошибочного запроса выписок
        Do While Len(File) > 0
            File1 = .Recv & File
            Kill File1
            File = Dir
        Loop
        
        File = Dir(.Recv & "E???????.*")
        nErr = 0: sErr = vbNullString
        Do While Len(File) > 0
            If User.IsID4(FileExt(File)) Then
                'Application.StatusBar = Bsprintf("Файл %s с ошибкой!", File)
                nErr = nErr + 1: sErr = sErr & Bsprintf("  %d. %s\n", nErr, File)
                File1 = .Recv & File
                File2 = .Archive & File
                If Payment.MarkByFile(File1, "Err", 12) Then 'Light Red
                    If IsFile(File2) Then Kill File2
                    Name File1 As File2
                    If Not IsFile(File2) Then
                        If YesNoBox("Файл '%s'\nне удалось переместить в архив\nв '%s'\n\n" & _
                            "Удалить его отсюда?", _
                            File1, .Archive) Then
                            Kill File1
                        End If
                    End If
                End If
            End If
            File = Dir
        Loop
        
        DoEvents
        File = Dir(.Recv & "T???????.*")
        nTest = 0: sTest = vbNullString
        Do While Len(File) > 0
            If User.IsID4(FileExt(File)) Then
                'Application.StatusBar = Bsprintf("Файл %s тестовый!", File)
                nTest = nTest + 1: sTest = sTest & Bsprintf("  %d. %s\n", nTest, File)
                File1 = .Recv & File
                File2 = .Archive & File
                If Payment.MarkByFile(File1, "Test", 12) Then 'Light Red
                    If IsFile(File2) Then Kill File2
                    Name File1 As File2
                    If Not IsFile(File2) Then
                        If YesNoBox("Файл '%s'\nне удалось переместить в архив\nв '%s'\n\n" & _
                            "Удалить его отсюда?", _
                            File1, .Archive) Then
                            Kill File1
                        End If
                    End If
                End If
            End If
            File = Dir
        Loop
        'Application.StatusBar = False
        
        DoEvents
        s = vbNullString
        If nOk > 0 Then s = s & Bsprintf("Принято %d:\n%s\n", nOk, sOk)
        If nErr > 0 Then s = s & Bsprintf("Отказано %d:\n%s\n", nErr, sErr)
        If nTest > 0 Then s = s & Bsprintf("Протестировано %d:\n%s\n", nTest, sTest)
        If Len(s) > 0 Then InfoBox s
    End With
    
    If ActiveWorkbook.Worksheets.Count <= 3 Then
        If YesNoBox("У Вас нет Вашего рабочего листа!\nДобавить его сейчас?") Then
            NewUserShow
        End If
    End If
    
    CheckRecv = False
End Function

Public Sub UrgentMessage(File As String)
    On Error Resume Next
    If Not YesNoBox("ВАЖНОЕ СООБЩЕНИЕ ИЗ БАНКА:\n\n%s\n\nНапоминать еще?", _
        CWin(InputFile(File))) Then
        Kill File
    End If
End Sub

Public Sub PGPPassword()
    Crypto.ChangePassword
End Sub

Public Sub ImportNewKeys()
    Dim Floppy As String, File As String, File2 As String
    Dim nOk As Long
    
    On Error Resume Next
    Floppy = "A:\Keys\"
    File = Dir(Floppy & "*.pgp")
    If File <> vbNullString Then
        'OK
    Else
        File = Floppy & "Pubr*.pgp"
        If App.LocateFile(File) Then
            Floppy = FilePath(File)
        Else
            InfoBox "Отказ от выполнения импорта."
            Exit Sub
        End If
    End If
    App.Options("KeysPath") = App.Path & "Keys\"
    ForceDirectories App.Options("KeysPath")

    File = Dir(Floppy & "*.pgp")
    Do While File <> vbNullString
        File2 = App.Options("KeysPath") & File
        nOk = 1
        Do While IsFile(File2 & Bsprintf(".%d.old", nOk))
            nOk = nOk + 1
        Loop
        If IsFile(File2) Then Name File2 As File2 & Bsprintf(".%d.old", nOk)
        If IsFile(File2) Then Kill File2
        If IsFile(File2) Then
            WarnBox "Директория защищена от записи!\nИмпорт не выполнен."
        Else
            FileCopy Floppy & File, File2
            If IsFile(File2) Then
                InfoBox "Файл ключа \'%s\' успешно импортирован", File
            Else
                WarnBox "Файл ключа \'%s\' НЕ импортирован!", File
            End If
        End If
        File = Dir
    Loop
End Sub
