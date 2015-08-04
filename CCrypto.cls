VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CCrypto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Option Base 1
DefLng A-Z

Const FormatPGP = "-----BEGIN PGP MESSAGE-----"

Dim PGP As New CPGP
Dim CryptCP As New CCryptCP

Public Sub ChangePassword()
    PGP.ChangePassword
End Sub

Public Property Let Password(ByVal vNewValue As String)
    PGP.Password = vNewValue
End Property

Public Function Encrypt(File1 As String, File2 As String) As Boolean
    On Error Resume Next
    'CryptCP commonly not used
    'If CryptCP.Encrypt(File1, File2) Then
    '    Encrypt = True
    'Else
        Encrypt = PGP.Encrypt(File1, File2)
    'End If
    Kill File1 'possible error!
End Function

Public Function Decrypt(File1 As String, File2 As String) As Boolean
    Dim s As String
    s = InputFile(File1)
    If Left(s, Len(FormatPGP)) = FormatPGP Then
        Decrypt = PGP.Decrypt(File1, File2)
    'Else
    'CryptCP commonly not used
        'Decrypt = CryptCP.Decrypt(File1, File2)
    End If
End Function

Public Function Info() As String
    'CryptCP commonly not used
    'If IsFile(CryptCP.File) Then
    '    Info = CryptCP.Info
    'Else
        Info = PGP.Info
    'End If
End Function
