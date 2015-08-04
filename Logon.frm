VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Logon 
   Caption         =   "Вход в систему Банк-Клиент"
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5505
   HelpContextID   =   410000
   OleObjectBlob   =   "Logon.frx":0000
   StartUpPosition =   1  'CenterOwner
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
End
Attribute VB_Name = "Logon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit
Option Compare Text
Option Base 1
DefLng A-Z

Private Sub chkChar_Click()
    txtPass.PasswordChar = IIf(chkChar, vbNullString, "*")
    PGP.ShowPass = chkChar
End Sub

Private Sub cmdCancel_Click()
    User.Demo = True
    Unload Me
End Sub

Private Sub cmdOk_Click()
    Hide
    If Val(txtID) = 0 Then txtID = User.DemoID
    User.ID = txtID
    PGP.ID = txtID
    If txtPass.TextLength > 0 Then
        PGP.Password = CDos(txtPass)
    Else
        PGP.Password = vbNullString
    End If
    Unload Me
End Sub

Private Sub cmdPGPPass_Click()
    PGP.ChangePassword
End Sub

Private Sub txtID_Change()
    cmdPGPPass.Enabled = fraPGP.Enabled And txtID = User.ID
End Sub

Private Sub UserForm_Initialize()
    On Error Resume Next
    With App
        With txtID
            .Text = User.ID
            .SelStart = 0
            .SelLength = .TextLength
        End With
        If PGP.Valid Then
            'txtPass = PGP.Password
            chkChar = PGP.ShowPass
            cmdPGPPass.Enabled = txtID = User.ID
        Else
            fraPGP.Caption = PGP.None
            fraPGP.Enabled = False
            lblPGP.Enabled = False
            txtPass.Enabled = False
            txtPass.BackStyle = fmBackStyleTransparent
            cmdPGPPass.Enabled = False
            chkChar.Enabled = False
        End If
    End With
End Sub
