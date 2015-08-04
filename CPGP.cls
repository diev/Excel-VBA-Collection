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

Dim mExtId As String
'Dim mPass As String
Dim mPassCollection As New Collection

Private Sub Class_Initialize()
    mExtId = User.DemoID 'vbNullString
    'mPass = vbNullString 'DefPGPPass
    ResetPasswords
End Sub

Public Property Get File() As String
    File = App.Path & "PGP\Pgp.exe"
End Property

Public Sub ChangePassword()
    If YesNoBox("ПРОЧТИТЕ ВНИМАТЕЛЬНО ПОРЯДОК СМЕНЫ ПАРОЛЯ PGP\n" & _
        "и приготовьтесь ответить на следующие вопросы:\n\n" & _
        "1. Идентификатор пользователя этого ключа: \'%s\'.\n" & _
        "2. Ввести старый пароль...\n" & _
        "3. Новый идентификатор пользователя - \'N\'.\n" & _
        "4. Сменить ваш пароль - \'Y\'.\n" & _
        "5. Ввести новый пароль...\n\n" & _
        "Приступить?", User.ID4) Then
            mExtId = User.ID4
            RunPGP "-ke"
            ResetPasswords
    End If
End Sub

Public Property Get Password() As String
    'Password = mPass
    Dim s As String
    On Error Resume Next
    If Len(mExtId) = 0 Then mExtId = User.ID4
    s = mPassCollection.Item(mExtId)
    If s = vbNullString Then
        s = InputBox("Введите пароль секретного ключа PGP для " & mExtId & ":", _
            App.TITLE)
        If Len(s) > 0 Then
            mPassCollection.Add s, Key:=mExtId
        End If
    End If
    Password = s
End Property

Public Property Let Password(ByVal vNewValue As String)
    'mPass = vNewValue
    On Error Resume Next
    If Len(vNewValue) = 0 Then ResetPasswords
    If Len(mExtId) = 0 Then mExtId = User.ID4
    mPassCollection.Remove mExtId
    If Len(vNewValue) > 0 Then
        mPassCollection.Add CDos(vNewValue), Key:=mExtId
    End If
End Property

Public Sub ResetPasswords()
    Dim v As Variant
    On Error Resume Next
    For Each v In mPassCollection
        v.Remove
    Next
End Sub

Public Function Encrypt(File1 As String, File2 As String) As Boolean
    Dim s As String
    On Error Resume Next
    Encrypt = False
    If OverwriteFile(File2) Then
        mExtId = FileExt(File2)
        s = StrArgs("-seaw %1 bank -o %2", QFile(File1), QFile(File2))
        If Len(Password) > 0 Then
            s = s & Bsprintf(" -z%s +force", QFile(Password)) '+batchmode
            Kill File2 'possible error
            RunPGP s
            Encrypt = IsFile(File2)
        End If
    End If
End Function

Public Function Decrypt(File1 As String, File2 As String) As Boolean
    Dim s As String
    On Error Resume Next
    Decrypt = False
    'If OverwriteFile(File2) Then
        mExtId = FileExt(File1)
        s = StrArgs("%1 -o %2", QFile(File1), QFile(File2))
        If Len(Password) > 0 Then
            s = s & Bsprintf(" -z%s +force", QFile(Password)) '+batchmode
            Kill File2 'possible error
            RunPGP s
            Decrypt = IsFile(File2)
        End If
    'End If
End Function

Public Function DecryptEx(File1 As String) As String
    Dim File2 As String
    On Error Resume Next
    File2 = GetWinTempFile
    If Decrypt(File1, File2) Then
        DecryptEx = File2
    Else
        Kill File2 'possible error
        DecryptEx = vbNullString
    End If
End Function

Public Property Get Comment() As String
    Comment = "Bank-Client " & User.ID & " - " & Now
End Property

Public Property Get RandSeed() As String
    RandSeed = App.Path & "PGP\RandSeed.bin"
End Property

Public Property Get Verbose() As String
    Dim s As String
    s = App.Options("Verbose")
    Verbose = IIf(Len(s) = 0, "0", s)
End Property

Public Property Get KeysPath() As String
    Dim KeyFile As String
    If Len(mExtId) = 0 Then mExtId = User.ID4
    KeyFile = "Pubr" & mExtId & ".pgp"
    KeysPath = App.Path & "Keys\"
    If Not IsFile(KeysPath & KeyFile) Then
        KeysPath = App.Path & "PGP\"
        If Not IsFile(KeysPath & KeyFile) Then
            KeysPath = App.Options("KeysPath")
            If Not IsFile(KeysPath & KeyFile) Then
                If App.LocateFile(KeyFile) Then
                    KeysPath = FilePath(KeyFile)
                    App.Options("KeysPath") = KeysPath
                End If
            End If
        End If
    End If
End Property

Private Function CommandLineCFG() As String
    Dim Path As String, File As String, s As String
    CommandLineCFG = vbNullString
    With Me
        s = Bsprintf(" -u %s ", mExtId)
        s = s & Bsprintf("+Comment=%s ", QFile(.Comment))
        
        Path = Me.KeysPath
        File = Path & "Pubr" & mExtId & ".pgp"
        If Not IsFile(File) Then
            Exit Function
        End If
        s = s & Bsprintf("+PubRing=%s ", QFile(File))
        
        File = Path & "Secr" & mExtId & ".pgp"
        If Not IsFile(File) Then
            Exit Function
        End If
        s = s & Bsprintf("+SecRing=%s ", QFile(File))
        
        'File = FilePath(.File) & "RandSeed.bin"
        File = App.Path & "PGP\RandSeed.bin"
        If Not IsFile(File) Then
            Exit Function
        End If
        s = s & Bsprintf("+RandSeed=%s ", QFile(File))
        
        s = s & Bsprintf("+Verbose=%s ", .Verbose)
        CommandLineCFG = s & "+ShowPass=on +EncryptToSelf=on"
    End With
End Function

Private Sub RunPGP(RunCmd As String)
    Dim retValue As Long, s As String, File As String
    On Error Resume Next
    File = Me.File
    If IsFile(File) Then
        s = CommandLineCFG()
        If Len(s) = 0 Then Exit Sub
        s = QFile(File) & " " & RunCmd & s
        ShellDialog s, vbNormalFocus, App.BoolOptions("Debug")
    Else
        StopBox "У Вас нет программы\n%s", File
    End If
End Sub

Public Function Info() As String
    Dim s As String, File As String
    File = Me.KeysPath
    s = "Ключи PGP: "
    If UCase(Left(File, 2)) = "A:" Then
        s = s & "на дискете - в " & QFile(File)
    ElseIf Left(File, Len(App.Path)) = App.Path Then
        s = s & "вместе с программой - в " & QFile(Mid(File, Len(App.Path) + 1))
    Else
        s = s & "в другом месте - в " & QFile(File)
    End If
    s = s & Bsprintf("\n(")
    File = Dir(Me.KeysPath & "Secr*.pgp")
    If File <> vbNullString Then
        File = Mid(FileNameOnly(File), 5, 4)
        If User.IsID4(File) Then
            s = s & File
        End If
        File = Dir
    End If
    Do While File <> vbNullString
        File = Mid(FileNameOnly(File), 5, 4)
        If User.IsID4(File) Then
            s = s & ", " & File
        End If
        File = Dir
    Loop
    Info = s & ")" & vbLf
End Function
