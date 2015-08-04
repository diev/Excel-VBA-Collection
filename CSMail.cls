VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSMail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Option Explicit
Option Compare Text
Option Base 1
DefLng A-Z

'Const SMailSection = "SMail"

Dim mFile As String

Private Sub Class_Initialize()
    mFile = App.Setting(SMailSection, "File")
    If Not IsFile(mFile) Then Me.File = PathDirectories("SMail;\SMail;\SMail+", "smail.exe")
    If App.NoSetting(SMailSection, "Send") Then WriteINI
End Sub

Public Property Get File() As String
    File = mFile
End Property

Public Property Get Path() As String
    Path = FilePath(mFile)
End Property

Public Property Let File(ByVal vNewValue As String)
    mFile = vNewValue
    If Not IsFile(mFile) Then LocateFile
End Property

Public Sub Dial()
    '"SysPath" required in smail.ctl or command line!
    Dim s As String, d As Boolean
    With App
        'd = StrToBool(.Setting(SMailSection, "Debug"))
        If DialOrInet = 2 Then
            s = .Setting(SMailSection, "Inet")
            If Len(s) = 0 Then
                'for just updated versions
                s = .Setting(SMailSection, "Dial") & " Channel1:NET"
                .Setting(SMailSection, "Inet") = s
            End If
            'ShellDialog mFile & " " & s, vbNormalFocus, d
            Shell mFile & " " & s, vbNormalFocus
        Else
            s = .Setting(SMailSection, "Dial")
            'ShellDialog mFile & " " & s, vbNormalFocus, d
            Shell mFile & " " & s, vbNormalFocus
        End If
    End With
    DoEvents
End Sub

Public Property Get Valid() As Boolean
    Valid = IsFile(mFile)
End Property

Public Function LocateFile() As Boolean
    LocateFile = False
    If StrToBool(App.Setting(SMailSection, "LocateCanceled")) Then Exit Function
    mFile = App.Path & "SMail\SMail.exe"
    If BrowseForFile(mFile, "Программы (*.exe;*.bat),*.exe;*.bat", _
        Bsprintf("программу SMail (обычно %s)", mFile), True) Then
        'If LCase(FileNameExt(mFile)) <> "smail.exe" Then
        '    WarnBox "Указанный файл %s\nвозможно, не является программой SMail!", mFile
        'End If
        App.Setting(SMailSection, "File") = FullFile(mFile)
    ElseIf YesNoBox("Отключить SMail и больше не искать?") Then
        App.Setting(SMailSection, "LocateCanceled") = 1
    End If
    LocateFile = IsFile(mFile)
End Function

Public Property Get None() As String
    None = "Программа SMail не установлена!"
End Property

Public Property Get DialOrInet() As Long
    DialOrInet = Val(App.Setting(SMailSection, "DialOrInet"))
    If DialOrInet <> 2 Then DialOrInet = 1
End Property

Public Property Let DialOrInet(ByVal vNewValue As Long)
    If vNewValue <> 2 Then vNewValue = 1
    App.Setting(SMailSection, "DialOrInet") = vNewValue
End Property

Public Property Get Send() As String
    Send = RightSlash(App.Setting(SMailSection, "Send"))
    ForceDirectories Send
End Property

Public Property Let Send(ByVal vNewValue As String)
    App.Setting(SMailSection, "Send") = vNewValue
End Property

Public Property Get Recv() As String
    Recv = RightSlash(App.Setting(SMailSection, "Recv"))
    ForceDirectories Recv
End Property

Public Property Let Recv(ByVal vNewValue As String)
    App.Setting(SMailSection, "Recv") = vNewValue
End Property

Public Property Get Archive() As String
    Archive = RightSlash(App.Setting(SMailSection, "Archive"))
    ForceDirectories Archive
End Property

Public Property Let Archive(ByVal vNewValue As String)
    App.Setting(SMailSection, "Archive") = vNewValue
End Property

Public Sub LastRep()
    Dim s As String, File As String, File1 As String, File2 As String
    Dim nOk As Long, nErr As Long, nTest As Long, sOk As String, sErr As String, sTest As String, sRep As String
    On Error Resume Next
    
    With Me
        File = Dir(.Recv & "O???????." & User.ID)
        nOk = 0: sOk = vbNullString
        Do While Len(File) > 0
            Application.StatusBar = Bsprintf("Файл %s ok", File)
            nOk = nOk + 1: sOk = sOk & Bsprintf("  %d. %s\n", nOk, File)
            If Not StrToBool(App.Options("DontMarkArchive")) Then _
                Payment.MarkByFile FileNameOnly(File), "Ok", 2 'Green
            File1 = .Recv & File
            File2 = .Archive & File
            If IsFile(File2) Then Kill File2
            Name File1 As File2
            File = Dir
        Loop
                
        File = Dir(.Recv & "E???????." & User.ID)
        nErr = 0: sErr = vbNullString
        Do While Len(File) > 0
            Application.StatusBar = Bsprintf("Файл %s с ошибкой!", File)
            nErr = nErr + 1: sErr = sErr & Bsprintf("  %d. %s\n", nErr, File)
            If App.Options("DontMarkArchive") <> "1" Then _
                Payment.MarkByFile FileNameOnly(File), "Err", 12  'Light Red
            File1 = .Recv & File
            File2 = .Archive & File
            If IsFile(File2) Then Kill File2
            Name File1 As File2
            File = Dir
        Loop
        
        File = Dir(.Recv & "T???????." & User.ID)
        nTest = 0: sTest = vbNullString
        Do While Len(File) > 0
            Application.StatusBar = Bsprintf("Файл %s тестовый!", File)
            nTest = nTest + 1: sTest = sTest & Bsprintf("  %d. %s\n", nTest, File)
            If App.Options("DontMarkArchive") <> "1" Then _
                Payment.MarkByFile FileNameOnly(File), "Test", 12  'Light Red
            File1 = .Recv & File
            File2 = .Archive & File
            If IsFile(File2) Then Kill File2
            Name File1 As File2
            File = Dir
        Loop
        Application.StatusBar = False
        
        s = vbNullString
        If nOk > 0 Then s = s & Bsprintf("Принято %d:\n%s\n", nOk, sOk)
        If nErr > 0 Then s = s & Bsprintf("Отказано %d:\n%s\n", nErr, sErr)
        If nTest > 0 Then s = s & Bsprintf("Протестировано %d:\n%s\n", nTest, sTest)
        If Len(s) > 0 Then s = s & Bsprintf("Эти файлы перенесены в архив.\n\n")
        
        'File1 = .Recv & "Rep" & User.ID & ".txt"
        'If IsFile(File1) Then s = s & BSPrintF("Ответ Банка:\n%s", CWin(InputFile(File1)))
        sRep = vbNullString
        File = Dir(.Recv & "Rep???.txt")
        Do While Len(File) > 0
            sRep = sRep & Bsprintf("%s\n\n", CWin(InputFile(.Recv & File)))
            File = Dir
        Loop
        If Len(sRep) > 0 Then s = s & Bsprintf("Ответ Банка:\n\n%s", sRep)
    End With
    
    If Len(s) = 0 Then s = "Ничего нового нет из Банка"
    InfoBox s
End Sub

Public Sub WriteINI()
    Dim s As String
    On Error Resume Next
    With App
        .Setting(SMailSection, "Debug") = 0
        .Setting(SMailSection, "File") = mFile
        s = QFile(ChangeFileExt(mFile, "CTL"))
        .Setting(SMailSection, "DialOrInet") = 1
        If Me.Valid Then
            .Setting(SMailSection, "Dial") = Bsprintf("@%s Poll:Bank SP", s)
            .Setting(SMailSection, "Inet") = Bsprintf("@%s Poll:Bank SP Channel1:NET", s)
            .Setting(SMailSection, "Recv") = Me.Path & "Mailbox\Bank\R\"
            .Setting(SMailSection, "Send") = Me.Path & "Mailbox\Bank\S\"
        Else
            .Setting(SMailSection, "Dial") = "@SMail.ctl Poll:Bank SP"
            .Setting(SMailSection, "Inet") = "@SMail.ctl Poll:Bank SP Channel1:NET"
            .Setting(SMailSection, "Recv") = App.Path & "SMail\Mailbox\Bank\R\"
            .Setting(SMailSection, "Send") = App.Path & "SMail\Mailbox\Bank\S\"
        End If
        .Setting(SMailSection, "Archive") = App.Path & "Archive\"
        If Not IsFile(App.Path & "dial.bat") Then
            s = Bsprintf("@echo off\n%s %s\nexit\n", QFile(mFile), .Setting(SMailSection, "Dial"))
            OutputFile App.Path & "dial.bat", s
        End If
        If Not IsFile(App.Path & "inet.bat") Then
            s = Bsprintf("@echo off\n%s %s\nexit\n", QFile(mFile), .Setting(SMailSection, "Inet"))
            OutputFile App.Path & "inet.bat", s
        End If
    End With
End Sub

