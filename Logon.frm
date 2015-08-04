VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Logon 
   Caption         =   "Вход в систему Банк-Клиент"
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5505
   OleObjectBlob   =   "Logon.frx":0000
   StartUpPosition =   1  'CenterOwner
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
    If chkChar Then
        txtPass.PasswordChar = vbNullString
        App.Setting("PGP", "ShowPass") = "1"
    Else
        txtPass.PasswordChar = "*"
        App.Setting("PGP", "ShowPass") = "0"
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    Hide
    If Val(txtID) = 0 Then txtID = User.DemoID
    User.ID = txtID
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
    cmdOk.Enabled = txtID.TextLength = 0 Or txtID.TextLength = 3
    cmdPGPPass.Enabled = fraPGP.Enabled And txtID = User.ID
End Sub

Private Sub UserForm_Initialize()
    With App
        txtID = User.ID
        If PGP.Valid Then
            'txtPass = PGP.Password
            chkChar = StrToBool(.Setting("PGP", "ShowPass"))
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
