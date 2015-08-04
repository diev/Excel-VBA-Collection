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
    If Not IsFile(mFile) Then Me.File = PathDirectories("SMail;\SMail", "smail.exe")
    If App.Setting(SMailSection, "Send") = vbNullString Then WriteINI
End Sub

Private Sub Class_Terminate()

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
    ShellDialog mFile & " " & App.Setting(SMailSection, "Dial"), _
        vbNormalFocus, StrToBool(App.Setting(SMailSection, "Debug"))
    DoEvents
End Sub

Public Property Get Valid() As Boolean
    Valid = IsFile(mFile)
End Property

Public Function LocateFile()
    LocateFile = False
    If StrToBool(App.Setting(SMailSection, "LocateCanceled")) Then Exit Function
    mFile = App.Path & "SMail"
    If BrowseForFile(mFile, "Программы (*.exe;*.bat),*.exe;*.bat", "программу SMail", True) Then
        App.Setting(SMailSection, "File") = FullFile(mFile)
    ElseIf MsgBox("Отключить SMail и больше не искать?", vbQuestion + vbYesNo, App.Title) = vbYes Then
        App.Setting(SMailSection, "LocateCanceled") = 1
    End If
    LocateFile = IsFile(mFile)
End Function

Public Property Get None() As String
    None = "Программа SMail не установлена!"
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
    Dim nOk As Long, nErr As Long, sOk As String, sErr As String
    On Error Resume Next
    
    With Me
        File = Dir(.Recv & "O???????." & User.ID)
        nOk = 0: sOk = vbNullString
        Do While Len(File) > 0
            Application.StatusBar = BPrintF("Файл \'%s\' ok", File)
            nOk = nOk + 1: sOk = BPrintF("%d. %s\n", nOk, File)
            If App.Options("DontMarkArchive") <> "1" Then _
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
            Application.StatusBar = BPrintF("Файл \'%s\' с ошибкой!", File)
            nErr = nErr + 1: sErr = BPrintF("  %d. %s\n", nErr, File)
            If App.Options("DontMarkArchive") <> "1" Then _
                Payment.MarkByFile FileNameOnly(File), "Err", 12  'Light Red
            File1 = .Recv & File
            File2 = .Archive & File
            If IsFile(File2) Then Kill File2
            Name File1 As File2
            File = Dir
        Loop
        Application.StatusBar = False
        
        s = vbNullString
        If nOk > 0 Then s = s & BPrintF("Принято %d:\n%s\n", nOk, sOk)
        If nErr > 0 Then s = s & BPrintF("Отказано %d:\n%s\n", nErr, sErr)
        
        File1 = .Recv & "Rep" & User.ID & ".txt"
        If IsFile(File1) Then s = s & BPrintF("Ответ Банка:\n%s", CWin(InputFile(File1)))
    End With
    
    If Len(s) = 0 Then s = "Ничего нового нет из Банка"
    MsgBox s, vbInformation, App.Title
End Sub

Public Sub WriteINI()
    Dim s As String
    With App
        .Setting(SMailSection, "Debug") = "0"
        If Me.Valid Then
            .Setting(SMailSection, "Dial") = BPrintF("@%s Poll:Bank,Bank SP", ChangeFileExt(mFile, "CTL"))
            .Setting(SMailSection, "Recv") = Me.Path & "Mailbox\Bank\R\"
            .Setting(SMailSection, "Send") = Me.Path & "Mailbox\Bank\S\"
        Else
            .Setting(SMailSection, "Dial") = "@SMail.ctl Poll:Bank,Bank SP"
            .Setting(SMailSection, "Recv") = App.Path & "SMail\Mailbox\Bank\R\"
            .Setting(SMailSection, "Send") = App.Path & "SMail\Mailbox\Bank\S\"
        End If
        .Setting(SMailSection, "Archive") = App.Path & "Archive\"
        s = BPrintF("@echo off\n%s %s\nexit\n", mFile, .Setting(SMailSection, "Dial"))
        OutputFile App.Path & "dial.bat", s
    End With
End Sub

