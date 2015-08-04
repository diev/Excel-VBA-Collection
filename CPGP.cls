VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CPGP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Option Base 1
DefLng A-Z

'Const PGPSection = "PGP"
'Const DefPGPPass = "000"

Dim mFile As String
Dim mLastID As String
Dim mPass As String

Private Sub Class_Initialize()
    mFile = App.Setting(PGPSection, "File")
    If Not IsFile(mFile) Then Me.File = PathDirectories("PGP;\PGP", "pgp.exe")
    If App.Setting(PGPSection, "Language") = vbNullString Then WriteINI
    mLastID = vbNullString
    mPass = vbNullString 'DefPGPPass
End Sub

Private Sub Class_Terminate()

End Sub

Public Sub ChangePassword()
    Dim s As String
    s = mFile & " " & App.Setting(PGPSection, "Password")
    'If User.ID <> mLastID Then
        mLastID = User.ID
        WriteCFG
    'End If
    ShellDialog s, vbNormalFocus, StrToBool(App.Setting(PGPSection, "Debug"))
    DoEvents
End Sub

Public Property Get Password() As String
    Password = mPass
End Property

Public Property Let Password(ByVal vNewValue As String)
    mPass = vNewValue
    'If InStr(mPass, " ") Then mPass = """" & mPass & """"
End Property

Public Function Encode(File1 As String, File2 As String) As Boolean
    Dim s As String
    On Error Resume Next
    If OverwriteFile(File2) Then
        s = mFile & " " & App.Setting(PGPSection, "Encode")
        s = StrArgs(s, File1, File2)
        If Len(mPass) > 0 Then s = s & BPrintF(" -z\'%s\' +batchmode +force", mPass)
        Kill File2 'possible error
        If User.ID <> mLastID Then
            mLastID = User.ID
            WriteCFG
        End If
        ShellDialog s, vbNormalFocus, StrToBool(App.Setting(PGPSection, "Debug"))
        DoEvents
        Encode = IsFile(File2)
    Else
        Encode = False
    End If
End Function

Public Function Decode(File1 As String, File2 As String) As Boolean
    Dim s As String
    On Error Resume Next
    'If OverwriteFile(File2) Then
        s = mFile & " " & App.Setting(PGPSection, "Decode")
        s = StrArgs(s, File1, File2)
        If Len(mPass) > 0 Then s = s & BPrintF(" -z\'%s\' +batchmode +force", mPass)
        Kill File2 'possible error
        If User.ID <> mLastID Then
            mLastID = User.ID
            WriteCFG
        End If
        ShellDialog s, vbNormalFocus, StrToBool(App.Setting(PGPSection, "Debug"))
        DoEvents
        Decode = IsFile(File2)
    'Else
    '    Decode = False
    'End If
End Function

Public Function DecodeEx(File1 As String) As String
    Dim File2 As String
    On Error Resume Next
    File2 = GetWinTempFile
    If Decode(File1, File2) Then
        DecodeEx = File2
    Else
        Kill File2 'possible error
        DecodeEx = vbNullString
    End If
End Function

Public Property Get File() As String
    File = mFile
End Property

Public Property Let File(ByVal vNewValue As String)
    mFile = vNewValue
    If Not IsFile(mFile) Then LocateFile
End Property

Public Sub WriteINI()
    App.Setting(PGPSection, "File") = vbNullString
    App.Setting(PGPSection, "Debug") = 0
    App.Setting(PGPSection, "Encode") = "-seaw %1 bank -o %2"
    App.Setting(PGPSection, "Decode") = "%1 -o %2"
    App.Setting(PGPSection, "Password") = "-ke"
    App.Setting(PGPSection, "LocateCanceled") = 0
    'PGP.INI
    App.Setting(PGPSection, "RandSeed") = Me.Path & "randseed.bin"
    App.Setting(PGPSection, "Language") = "ru"
    App.Setting(PGPSection, "Charset") = "cp866"
    App.Setting(PGPSection, "Verbose") = 0
    App.Setting(PGPSection, "ShowPass") = 1
    App.Setting(PGPSection, "EncryptToSelf") = 1
End Sub

Public Property Get Valid() As Boolean
    Valid = IsFile(mFile)
End Property

Public Function LocateFile() As Boolean
    LocateFile = False
    If StrToBool(App.Setting(PGPSection, "LocateCanceled")) Then Exit Function
    mFile = App.Path & "PGP"
    If BrowseForFile(mFile, "Программы (*.exe;*.bat),*.exe;*.bat", "программу PGP", True) Then
        App.Setting(PGPSection, "File") = FullFile(mFile)
    ElseIf MsgBox("Отключить PGP и больше не искать?", vbQuestion + vbYesNo, App.Title) = vbYes Then
        App.Setting(PGPSection, "LocateCanceled") = 1
    End If
    LocateFile = IsFile(mFile)
End Function

Public Property Get None() As String
    None = "Программа PGP не установлена!"
End Property

Public Sub Reset()
    If App.Setting(PGPSection, "File") = vbNullString Then WriteINI
    Me.File = App.Setting(PGPSection, "File")
End Sub

Public Property Get MyName() As String
    MyName = User.ID
End Property

Public Property Let MyName(ByVal vNewValue As String)

End Property

Public Property Get Comment() As String
    Comment = "Bank-Client " & User.ID
End Property

Public Property Let Comment(ByVal vNewValue As String)

End Property

Public Property Get PubRing() As String
    With User
        PubRing = .KeysPath & "PUBR" & .ID & ".PGP"
        If Not IsFile(PubRing) Then
            If .LocateKeys Then PubRing = .KeysPath & "PUBR" & .ID & ".PGP"
        End If
    End With
End Property

Public Property Let PubRing(ByVal vNewValue As String)
    User.KeysPath = FilePath(vNewValue)
End Property

Public Property Get SecRing() As String
    With User
        SecRing = .KeysPath & "SECR" & .ID & ".PGP"
        If Not IsFile(SecRing) Then
            If .LocateKeys Then SecRing = .KeysPath & "SECR" & .ID & ".PGP"
        End If
    End With
End Property

Public Property Let SecRing(ByVal vNewValue As String)
    User.KeysPath = FilePath(vNewValue)
End Property

Public Property Get RandSeed() As String
    RandSeed = App.Setting(PGPSection, "RandSeed")
End Property

Public Property Let RandSeed(ByVal vNewValue As String)
    App.Setting(PGPSection, "RandSeed") = vNewValue
End Property

Public Property Get Language() As String
    Language = App.Setting(PGPSection, "Language")
End Property

Public Property Let Language(ByVal vNewValue As String)
    App.Setting(PGPSection, "Language") = vNewValue
End Property

Public Property Get Charset() As String
    Charset = App.Setting(PGPSection, "Charset")
End Property

Public Property Let Charset(ByVal vNewValue As String)
    App.Setting(PGPSection, "Charset") = vNewValue
End Property

Public Property Get Verbose() As Long
    Verbose = App.Setting(PGPSection, "Verbose")
End Property

Public Property Let Verbose(ByVal vNewValue As Long)
    App.Setting(PGPSection, "Verbose") = vNewValue
End Property

Public Property Get ShowPass() As Boolean
    ShowPass = StrToBool(App.Setting(PGPSection, "ShowPass"))
End Property

Public Property Let ShowPass(ByVal vNewValue As Boolean)
    App.Setting(PGPSection, "ShowPass") = BoolToStr(vNewValue)
End Property

Public Property Get EncryptToSelf() As Boolean
    EncryptToSelf = StrToBool(App.Setting(PGPSection, "EncryptToSelf"))
End Property

Public Property Let EncryptToSelf(ByVal vNewValue As Boolean)
    App.Setting(PGPSection, "EncryptToSelf") = BoolToStr(vNewValue)
End Property

Public Property Get CFG() As String
    CFG = ChangeFileExt(mFile, "INI")
End Property

Public Property Get Path() As String
    Path = FilePath(mFile)
End Property

Public Function WriteCFG() As Boolean
    Dim s As String
    With Me
        s = BPrintF("# This file is rewritten with each action!\n# AutoGenerated on %s\n", Now)
        s = s & BPrintF("MyName=%s\n", .MyName)
        s = s & BPrintF("Comment=%s\n", .Comment)
        s = s & BPrintF("PubRing=%s\n", .PubRing)
        s = s & BPrintF("SecRing=%s\n", .SecRing)
        s = s & BPrintF("RandSeed=%s\n", .RandSeed)
        s = s & BPrintF("Language=%s\n", .Language)
        s = s & BPrintF("Charset=%s\n", .Charset)
        s = s & BPrintF("Verbose=%d\n", .Verbose)
        s = s & BPrintF("ShowPass=%s\n", IIf(.ShowPass, "on", "off"))
        s = s & BPrintF("EncryptToSelf=%s\n", IIf(.EncryptToSelf, "on", "off"))
        OutputFile .CFG, s
    End With
End Function
