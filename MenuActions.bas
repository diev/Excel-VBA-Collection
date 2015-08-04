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
    Payment.EachSelected "ExportPlat", "��������� ���?"
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
    Payment.EachSelected "Delete", "������������ �������?"
End Sub

Public Sub AutoCheck()
    On Error Resume Next
    'If ActiveSheet.Name <> User.ID Then
    '    Workbooks(App.BookName).Activate
    'Else
    If Val(User.BIC) = 0 Then
        WarnBox "��������, ��������� ��������������!"
        Restart
    End If
End Sub

Public Sub AutoRestart(Optional s As String = vbNullString)
    On Error Resume Next
    WarnBox "��������, ��������� ��������������!\n%s", s
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
    Payment.EachSelected "Preview", "�������� ���?"
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
            StopBox "������� ���������� ���� ����������\n%s", File2
            Exit Sub
        End If
    End If
    
    If YesNoBox("�������� ����������� ������� ����������:\n\n" & _
        "1. ����� ������� ���� ������ '��' ��������� ���������.\n" & _
        "2. �������� EXCEL, ���� �� �� ��������� ���!!!\n" & _
        "3. ������� ���������� ���������� � ������� � ��� 'OK'.\n" & _
        "4. ��������� ��������� ��� �������������� ������\n" & _
        "(��� ����� ������ ���� ���������� � ��������).\n" & _
        "5. ��������� ��������� ����-������ �����\n" & _
        "(� ����������� ������ ������ ��� ������).\n" & _
        "6. ������� ����������� ���������� '%s' �� ���������.\n\n" & _
        "����������?", File2) Then
        
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
        s = InputBox("������� ����� ������ Excel ���" & vbCrLf & .FullName, _
            App.TITLE)
        If Len(s) > 0 Then
            Kill .FullName
            DoEvents
            .SaveAs .FullName, Password:=s, CreateBackup:=True
        ElseIf .HasPassword Then
            If YesNoBox("����� ������� ������?") Then
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
    If YesNoBox("������ ������� �� " & Format(Now - 1, "dddd dd.MM.yyyy HH:mm")) Then
        Date = Date - 1
    Else
        s = Format(Date - 1, "dd.MM.yy")
        s = InputBox("������� ���� ��� �������:", App.TITLE, s)
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
            StopBox "���� ��� ������ �� ������!"
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
        StopBox "� ��� ��� ���������\n%s", File
    End If
End Sub

Public Sub SSetup()
    Dim File As String
    On Error Resume Next
    File = App.Path & "SMail\SSetup.exe"
    If IsFile(File) Then
        Shell QFile(File) & " " & QFile(App.Path & "SMail\SMail.ctl"), vbNormalFocus
    Else
        StopBox "� ��� ��� ���������\n%s", File
    End If
End Sub

Public Sub SMailLog()
    Dim File As String
    On Error Resume Next
    File = App.Path & "SMail\SMail.log"
    If IsFile(File) Then
        Shell "notepad.exe " & QFile(File), vbNormalFocus
    Else
        WarnBox "� ��� ��� �����\n%s", File
    End If
End Sub

Public Sub SMailCtl()
    Dim File As String
    On Error Resume Next
    File = App.Path & "SMail\SMail.ctl"
    If IsFile(File) Then
        Shell "notepad.exe " & QFile(File), vbNormalFocus
    Else
        WarnBox "� ��� ��� �����\n%s", File
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
        WarnBox "������ ������ � %s �� �������!", App.Path
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
            InfoBox "������ ��������� - %s", sOk
            App.Options("Version") = sOk
            App.Options("Updated") = sOk
        ElseIf s = sOk Then
            WarnBox "���� �������� ���������� ���������,\n�� ���, �������, �� ���������!\n" & _
                "��������� ������ ������������ �����������\n��������� � ����.\n\n" & _
                "������ � ��� ������ - %s", sOk
            App.Options("Updated") = sOk
        Else
            InfoBox "������ ���������\n��������� �� %s", sOk
            App.Options("Version") = sOk
            App.Options("Updated") = sOk
        End If
    End If
    
    SMail.Valid
    
    'Antivirus?
    File = SMail.Recv & "!antivir.txt"
    s = "��� ��������� ��������� ������ ����� ���������!"
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
        InfoBox "�������� ���������� ����������� ������.\n��������� ��������� ��� � ��������������."
        ShellWait File & " /auto " & App.Path, vbNormalNoFocus
        DoEvents
        Kill File
        DoEvents
        Restart
        DoEvents
        With BnkSeek2
            .LoadFile
            InfoBox "���������� ���������� ��� �� %s, ������: %d", _
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
                    WarnBox "���������� ��������� �� ����������,\n" & _
                        "�.�. Excel ���������� ������ ���� ������ �������:\n\n%s%s", _
                        .FileInfo(File2), .FileInfo(File1)
                End If
            End If
        Else
            .Options("Version") = .Version
            .Options("Updated") = Now
            InfoBox "��������� ������� ��������� �� %s", .Version
            File = Dir(SMail.Recv & "OK-*.exe")
            Do While File <> vbNullString
                'If YesNoBox("���������� '%s' ���������\n������� ����?", Mid(File, 4)) Then
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
            If YesNoBox("�������� ����� ���������� %s\n(���� �� %s)\n\n��������� ��� ������?", _
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
        If YesNoBox("�������� ����� ������� '%s'\n�� ����� %s\n\n�������� ��� ������?", s4, s20) Then
            'ID
            For Each sh In Worksheets
                If App.DefLS(sh.Name) = s Then
                    If YesNoBox("����� ������ '%s' ��� ����!\n������� ��� �� ������� �� ����������?", s) Then
                        Kill SMail.Recv & File
                        Kill SMail.Recv & s4 & ".plt"
                        Kill SMail.Recv & s20 & ".plt"
                        GoTo NextID
                    'ElseIf YesNoBox("���� '%s' ����� ������ ������!", s) Then
                    '    Sheets(s).Delete
                    'Else
                    '    WarnBox "������, ���������� ����."
                    '    GoTo NextID
                    End If
                    Exit For
                End If
            Next
            'For Each sh In Worksheets
            '    If sh.Name = User.DemoID & " (2)" Then
            '        WarnBox "���������� ������� ���� '000 (2)'"
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
                YesNoCancelBox("�������� ����� ����� \'%s\' -\n������������� ��������� �� �� �������.\n\n" & _
                "�� - ��������� �� �� ������� � \'%s\'\n" & _
                "��� - ��������� � ��������� � \'%s\'\n" & _
                "������ - ������������ �����.", File, "A:\Keys\", App.Path & "Keys\")
                Case vbyes:
                    App.Options("KeysPath") = "A:\Keys\"
                    WarnBox "���������, ��� ������� ���������\n� ������ � ������!"
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
                            WarnBox "������� �������� �� ������!\n������� �� ��������."
                        Else
                            Name File1 As File2
                            If IsFile(File2) Then
                                InfoBox "���� ����� \'%s\' ������� ���������", File
                                Kill File1
                            Else
                                WarnBox "���� ����� \'%s\' �� ���������!", File
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
                            WarnBox "���������� �������� �� ������!\n������� �� ��������."
                        Else
                            Name File1 As File2
                            If IsFile(File2) Then
                                InfoBox "���� ����� \'%s\' ������� ���������", File
                                Kill File1
                            Else
                                WarnBox "���� ����� \'%s\' �� ���������!", File
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
                'Application.StatusBar = Bsprintf("���� %s ok", File)
                nOk = nOk + 1: sOk = sOk & Bsprintf("  %d. %s\n", nOk, File)
                File1 = .Recv & File
                File2 = .Archive & File
                If Payment.MarkByFile(File1, "Ok", 2) Then 'Green
                    If IsFile(File2) Then Kill File2
                    Name File1 As File2
                    If Not IsFile(File2) Then
                        If YesNoBox("���� '%s'\n�� ������� ����������� � �����\n� '%s'\n\n" & _
                            "������� ��� ������?", _
                            File1, .Archive) Then
                            Kill File1
                        End If
                    End If
                End If
            End If
            File = Dir
        Loop
                
        DoEvents
        File = Dir(.Recv & "EYP??*.*") '�������� ���������� �������� �������
        Do While Len(File) > 0
            File1 = .Recv & File
            Kill File1
            File = Dir
        Loop
        
        DoEvents
        File = Dir(.Recv & "ESK.*") '�������� ���������� ������� �������
        Do While Len(File) > 0
            File1 = .Recv & File
            Kill File1
            File = Dir
        Loop
        
        File = Dir(.Recv & "E???????.*")
        nErr = 0: sErr = vbNullString
        Do While Len(File) > 0
            If User.IsID4(FileExt(File)) Then
                'Application.StatusBar = Bsprintf("���� %s � �������!", File)
                nErr = nErr + 1: sErr = sErr & Bsprintf("  %d. %s\n", nErr, File)
                File1 = .Recv & File
                File2 = .Archive & File
                If Payment.MarkByFile(File1, "Err", 12) Then 'Light Red
                    If IsFile(File2) Then Kill File2
                    Name File1 As File2
                    If Not IsFile(File2) Then
                        If YesNoBox("���� '%s'\n�� ������� ����������� � �����\n� '%s'\n\n" & _
                            "������� ��� ������?", _
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
                'Application.StatusBar = Bsprintf("���� %s ��������!", File)
                nTest = nTest + 1: sTest = sTest & Bsprintf("  %d. %s\n", nTest, File)
                File1 = .Recv & File
                File2 = .Archive & File
                If Payment.MarkByFile(File1, "Test", 12) Then 'Light Red
                    If IsFile(File2) Then Kill File2
                    Name File1 As File2
                    If Not IsFile(File2) Then
                        If YesNoBox("���� '%s'\n�� ������� ����������� � �����\n� '%s'\n\n" & _
                            "������� ��� ������?", _
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
        If nOk > 0 Then s = s & Bsprintf("������� %d:\n%s\n", nOk, sOk)
        If nErr > 0 Then s = s & Bsprintf("�������� %d:\n%s\n", nErr, sErr)
        If nTest > 0 Then s = s & Bsprintf("�������������� %d:\n%s\n", nTest, sTest)
        If Len(s) > 0 Then InfoBox s
    End With
    
    If ActiveWorkbook.Worksheets.Count <= 3 Then
        If YesNoBox("� ��� ��� ������ �������� �����!\n�������� ��� ������?") Then
            NewUserShow
        End If
    End If
    
    CheckRecv = False
End Function

Public Sub UrgentMessage(File As String)
    On Error Resume Next
    If Not YesNoBox("������ ��������� �� �����:\n\n%s\n\n���������� ���?", _
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
            InfoBox "����� �� ���������� �������."
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
            WarnBox "���������� �������� �� ������!\n������ �� ��������."
        Else
            FileCopy Floppy & File, File2
            If IsFile(File2) Then
                InfoBox "���� ����� \'%s\' ������� ������������", File
            Else
                WarnBox "���� ����� \'%s\' �� ������������!", File
            End If
        End If
        File = Dir
    Loop
End Sub
