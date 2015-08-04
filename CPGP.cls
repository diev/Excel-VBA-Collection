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

'PGP exit codes if no "+batchmode" when 0/1 only
Const PGP_EXIT_OK = 0
Const PGP_INVALID_FILE_ERROR = 1
Const PGP_FILE_NOT_FOUND_ERROR = 2
Const PGP_UNKNOWN_FILE_ERROR = 3
Const PGP_NO_BATCH = 4
Const PGP_BAD_ARG_ERROR = 5
Const PGP_INTERRUPT = 6
Const PGP_OUT_OF_MEM = 7

'PGP keyrings errors, base from 10
Const PGP_KEYGEN_ERROR = 10
Const PGP_NONEXIST_KEY_ERROR = 11
Const PGP_KEYRING_ADD_ERROR = 12
Const PGP_KEYRING_EXTRACT_ERROR = 13
Const PGP_KEYRING_EDIT_ERROR = 14
Const PGP_KEYRING_VIEW_ERROR = 15
Const PGP_KEYRING_REMOVE_ERROR = 16
Const PGP_KEYRING_CHECK_ERROR = 17
Const PGP_KEY_SIGNATURE_ERROR = 18
Const PGP_KEYSIG_REMOVE_ERROR = 19
                             
'PGP encode errors, base from 20
Const PGP_SIGNATURE_ERROR = 20
Const PGP_RSA_ENCR_ERROR = 21
Const PGP_ENCR_ERROR = 22
Const PGP_COMPRESS_ERROR = 23
                             
'PGP decode errors, base from 30
Const PGP_SIGNATURE_CHECK_ERROR = 30
Const PGP_RSA_DECR_ERROR = 31
Const PGP_DECR_ERROR = 32
Const PGP_DECOMPRESS_ERROR = 33


Dim mFile As String
Dim mPGPid As String
'Dim mPass As String
Dim mPassCollection As New Collection

Private Sub Class_Initialize()
    mFile = App.Setting(PGPSection, "File")
    If Not IsFile(mFile) Then Me.File = PathDirectories("PGP;\PGP", "pgp.exe")
    If App.NoSetting(PGPSection, "Language") Then WriteINI
    mPGPid = vbNullString
    'mPass = vbNullString 'DefPGPPass
    ResetPasswords
End Sub

Public Sub ChangePassword()
    Dim s As String
    s = mFile & " " & App.Setting(PGPSection, "Password")
    RunPGP s
    ResetPasswords
End Sub

Public Property Get Password() As String
    'Password = mPass
    Dim s As String
    On Error Resume Next
    s = mPassCollection.Item(mPGPid)
    If s = vbNullString Then
        s = InputBox("Введите пароль секретного ключа PGP для " & mPGPid & ":", _
            App.Title)
        If Len(s) > 0 Then
            mPassCollection.Add s, Key:=mPGPid
        End If
    End If
    Password = s
End Property

Public Property Let Password(ByVal vNewValue As String)
    'mPass = vNewValue
    On Error Resume Next
    mPassCollection.Remove mPGPid
    If Len(vNewValue) > 0 Then
        mPassCollection.Add CDos(vNewValue), Key:=mPGPid
    End If
End Property

Public Sub ResetPasswords()
    Dim v As Variant
    On Error Resume Next
    For Each v In mPassCollection
        v.Remove
    Next
End Sub

Public Property Get ID() As String
    ID = mPGPid
End Property

Public Property Let ID(ByVal vNewValue As String)
    If mPGPid <> vNewValue Then
        mPGPid = vNewValue
        'mPass = vbNullString
    End If
End Property

Public Function Encode(File1 As String, File2 As String) As Boolean
    Dim s As String
    On Error Resume Next
    If OverwriteFile(File2) Then
        s = QFile(mFile) & " " & App.Setting(PGPSection, "Encode")
        s = StrArgs(s, QFile(File1), QFile(File2))
        If Len(Password) > 0 Then s = s & Bsprintf(" -z%s +force", QFile(Password)) '+batchmode
        Kill File2 'possible error
        RunPGP s
        Encode = IsFile(File2)
    Else
        Encode = False
    End If
End Function

Public Function Decode(File1 As String, File2 As String) As Boolean
    Dim s As String
    On Error Resume Next
    'If OverwriteFile(File2) Then
        s = QFile(mFile) & " " & App.Setting(PGPSection, "Decode")
        s = StrArgs(s, QFile(File1), QFile(File2))
        If Len(Password) > 0 Then s = s & Bsprintf(" -z%s +force", QFile(Password)) '+batchmode
        Kill File2 'possible error
        RunPGP s
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
    With App
        .Setting(PGPSection, "Debug") = 0
        .Setting(PGPSection, "File") = mFile
        .Setting(PGPSection, "Encode") = "-seaw %1 bank -o %2"
        .Setting(PGPSection, "Decode") = "%1 -o %2"
        .Setting(PGPSection, "Password") = "-ke"
        .Setting(PGPSection, "LocateCanceled") = 0
        'PGP.INI
        .Setting(PGPSection, "RandSeed") = Me.Path & "randseed.bin"
        .Setting(PGPSection, "Language") = "ru"
        .Setting(PGPSection, "Charset") = "cp866"
        .Setting(PGPSection, "Verbose") = 0
        .Setting(PGPSection, "ShowPass") = 1
        .Setting(PGPSection, "EncryptToSelf") = 1
    End With
End Sub

Public Property Get Valid() As Boolean
    Valid = IsFile(mFile)
End Property

Public Function LocateFile() As Boolean
    LocateFile = False
    If StrToBool(App.Setting(PGPSection, "LocateCanceled")) Then Exit Function
    mFile = App.Path & "PGP\PGP.exe"
    If BrowseForFile(mFile, "Программы (*.exe;*.bat),*.exe;*.bat", _
        Bsprintf("программу PGP (обычно %s)", mFile), True) Then
        'If LCase(FileNameExt(mFile)) <> "pgp.exe" Then
        '    WarnBox "Указанный файл %s\nвозможно, не является программой PGP!", mFile
        'End If
        App.Setting(PGPSection, "File") = FullFile(mFile)
    ElseIf YesNoBox("Отключить PGP и больше не искать?") Then
        App.Setting(PGPSection, "LocateCanceled") = 1
    End If
    LocateFile = IsFile(mFile)
End Function

Public Property Get None() As String
    None = "Программа PGP не установлена!"
End Property

Public Sub Reset()
    If App.NoSetting(PGPSection, "File") Then WriteINI
    Me.File = App.Setting(PGPSection, "File")
End Sub

Public Property Get MyName() As String
    MyName = User.ID
End Property

Public Property Let MyName(ByVal vNewValue As String)
    '
End Property

Public Property Get Comment() As String
    Comment = "Bank-Client " & User.ID
End Property

Public Property Let Comment(ByVal vNewValue As String)
    '
End Property

'Public Property Get PubRing() As String
'    With User
'        PubRing = .KeysPath & "PUBR" & .ID & ".PGP"
'        If Not IsFile(PubRing) Then
'            If .LocateKeys Then PubRing = .KeysPath & "PUBR" & .ID & ".PGP"
'        End If
'    End With
'End Property
'
'Public Property Let PubRing(ByVal vNewValue As String)
'    User.KeysPath = FilePath(vNewValue)
'End Property
'
'Public Property Get SecRing() As String
'    With User
'        SecRing = .KeysPath & "SECR" & .ID & ".PGP"
'        If Not IsFile(SecRing) Then
'            If .LocateKeys Then SecRing = .KeysPath & "SECR" & .ID & ".PGP"
'        End If
'    End With
'End Property
'
'Public Property Let SecRing(ByVal vNewValue As String)
'    User.KeysPath = FilePath(vNewValue)
'End Property

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
        s = Bsprintf("# This file is rewritten with each action!\n# AutoGenerated on %s\n", Now)
        s = s & Bsprintf("MyName=%s\n", .ID)
        s = s & Bsprintf("Comment=%s\n", .Comment)
        s = s & Bsprintf("PubRing=%s\n", User.PubRing(.ID))
        s = s & Bsprintf("SecRing=%s\n", User.SecRing(.ID))
        s = s & Bsprintf("RandSeed=%s\n", .RandSeed)
        s = s & Bsprintf("Language=%s\n", .Language)
        s = s & Bsprintf("Charset=%s\n", .Charset)
        s = s & Bsprintf("Verbose=%d\n", .Verbose)
        s = s & Bsprintf("ShowPass=%s\n", IIf(.ShowPass, "on", "off"))
        s = s & Bsprintf("EncryptToSelf=%s\n", IIf(.EncryptToSelf, "on", "off"))
        OutputFile .CFG, s
    End With
End Function

Private Sub RunPGP(RunCmd As String)
    Dim retValue As Long, s As String
    WriteCFG
    retValue = ShellDialog(RunCmd, vbNormalFocus, StrToBool(App.Setting(PGPSection, "Debug")))
    DoEvents
    Exit Sub
    If retValue <> PGP_EXIT_OK Then
        Select Case retValue
            Case PGP_EXIT_OK
                s = "программа PGP завершилась нормально"
            Case PGP_INVALID_FILE_ERROR
                s = "неправильный формат файла"
            Case PGP_FILE_NOT_FOUND_ERROR
                s = "указанный файл не найден"
            Case PGP_UNKNOWN_FILE_ERROR
                s = "неизвестная файловая ошибка"
            Case PGP_NO_BATCH
                s = "операция не для пакетной обработки"
            Case PGP_BAD_ARG_ERROR
                s = "введен неправильный аргумент"
            Case PGP_INTERRUPT
                s = "выполнение прервано пользователем"
            Case PGP_OUT_OF_MEM
                s = "программе PGP не хватает DOS-памяти"
            
            Case PGP_KEYGEN_ERROR
                s = "ошибка генерации ключей"
            Case PGP_NONEXIST_KEY_ERROR
                s = "указанный ключ не существует"
            Case PGP_KEYRING_ADD_ERROR
                s = "ошибка добавления ключей"
            Case PGP_KEYRING_EXTRACT_ERROR
                s = "ошибка извлечения ключей"
            Case PGP_KEYRING_EDIT_ERROR
                s = "ошибка редактирования ключей"
            Case PGP_KEYRING_VIEW_ERROR
                s = "ошибка просмотра ключей"
            Case PGP_KEYRING_REMOVE_ERROR
                s = "ошибка удаления ключей"
            Case PGP_KEYRING_CHECK_ERROR
                s = "ошибка проверки ключей"
            Case PGP_KEY_SIGNATURE_ERROR
                s = "ошибка подписывания ключа или нет права подписи"
            Case PGP_KEYSIG_REMOVE_ERROR
                s = "ошибка снятия подписи с ключа"
                                         
            Case PGP_SIGNATURE_ERROR
                s = "ошибка подписывания или нет права подписи"
            Case PGP_RSA_ENCR_ERROR
                s = "ошибка шифрования RSA"
            Case PGP_ENCR_ERROR
                s = "ошибка шифрования"
            Case PGP_COMPRESS_ERROR
                s = "ошибка сжатия файла"
                                         
            Case PGP_SIGNATURE_CHECK_ERROR
                s = "ошибка проверки подписи на подлинность"
            Case PGP_RSA_DECR_ERROR
                s = "ошибка расшифровывания RSA"
            Case PGP_DECR_ERROR
                s = "ошибка расшифровывания"
            Case PGP_DECOMPRESS_ERROR
                s = "ошибка распаковки сжатого файла"
            
            Case Else
                s = "нет информации о сути этой ошибки"
        End Select
    
        'WarnBox "Ошибка %d (%s)\nпри выполнении команды:\n\n%s", retValue, s, RunCmd
    End If
End Sub
