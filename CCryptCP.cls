VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CCryptCP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Option Base 1
DefLng A-Z

Public Property Get File() As String
    File = App.Path & "CryptCP.exe"
End Property

Public Function Encrypt(File1 As String, File2 As String) As Boolean
    Dim s As String, File3 As String
    On Error Resume Next
    Encrypt = False
    If Not IsFile(Me.File) Then Exit Function
    File3 = GetWinTempFile
    If OverwriteFile(File2) Then
        'sign
        s = StrArgs("-sign -dn CN=%1 %2 %3", User.LS, QFile(File1), QFile(File3))
        Kill File3 'possible error
        RunCryptCP s
        If IsFile(File3) Then
            'encrypt
            s = StrArgs("-encr -dn CN=%1 %2 %3", User.LS, QFile(File3), QFile(File2))
            Kill File2 'possible error
            RunCryptCP s
            Kill File3
            '
            Encrypt = IsFile(File2)
        Else
            WarnBox "Криптопровайдер %s не работает!\nБудет использована система PGP.", File
        End If
    End If
End Function

Public Function Decrypt(File1 As String, File2 As String) As Boolean
    Dim s As String, ExtId As String, File3 As String
    On Error Resume Next
    Decrypt = False
    'If OverwriteFile(File2) Then
        ExtId = FileExt(File1) ''''''''''''''''''''''''''''''''''''
        File3 = GetWinTempFile
        'decrypt
        s = StrArgs("-decr -dn CN=%1 %2 %3", ExtId, QFile(File1), QFile(File3))
        Kill File3 'possible error
        RunCryptCP s
        If IsFile(File3) Then
            'verify
            s = StrArgs("-verify -dn CN=%1 %2 %3", User.KS, QFile(File3), QFile(File2))
            Kill File2 'possible error
            RunCryptCP s
            Kill File3
            '
            Decrypt = IsFile(File2)
        Else
            'decrypt fail, try verify only
            s = StrArgs("-verify -dn CN=%1 %2 %3", User.KS, QFile(File1), QFile(File2))
            Kill File2 'possible error
            RunCryptCP s
            '
            Decrypt = IsFile(File2)
        End If
    'End If
End Function

Private Sub RunCryptCP(RunCmd As String)
    Dim retValue As Long, s As String, File As String
    On Error Resume Next
    File = Me.File
    If IsFile(File) Then
        s = QFile(File) & " " & RunCmd
        retValue = ShellDialog(s, vbNormalFocus, App.BoolOptions("Debug"))
    Else
        StopBox "У Вас нет программы\n%s", File
    End If
End Sub

Public Function Info() As String
    Info = Bsprintf("Шифрование и ЭЦП обеспечиваются CryptoPro\n%s", App.FileInfo(Me.File))
End Function

