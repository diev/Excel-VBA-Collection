VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserPrivate 
   Caption         =   "Ключи плательщика"
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5505
   OleObjectBlob   =   "UserPrivate.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserPrivate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Explicit
Option Compare Text
Option Base 1
DefLng A-Z

Const None = "(нет)"

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    With txtKeys
        If .Enabled And Not IsDir(.Text) Then
            MsgBox "Директория ключей PGP не найдена!", vbExclamation, App.Title
            .SetFocus
            Exit Sub
        End If
    End With
    Hide
    With User
        .Sign(1) = txtSign1
        .Sign(2) = txtSign2
        .Tel = txtTel
        .No = txtNo
        If txtKeys <> None Then .KeysPath = txtKeys
        .ImportList = txtImport
        .ExportList = txtExport
    End With
    Unload Me
End Sub

Private Sub cmdkeys_Click()
    User.LocateKeys
    txtKeys = IIf(IsDir(User.KeysPath), User.KeysPath, None)
End Sub

Private Sub spnNo_Change()
    On Error Resume Next
    spnNo = 0
End Sub

Private Sub spnNo_SpinDown()
    If txtNo.Value > 1 Then txtNo = txtNo.Value - 1
End Sub

Private Sub spnNo_SpinUp()
    txtNo = txtNo.Value + 1
End Sub

Private Sub txtkeys_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    cmdkeys_Click
End Sub

Private Sub UserForm_Initialize()
    With User
        Caption = Caption & BPrintF(" \'%s\'", .ID)
        spnNo.Min = CLng(.NoMin)
        spnNo.Max = CLng(.NoMax)
        
        txtSign1 = .Sign(1)
        txtSign2 = .Sign(2)
        txtTel = .Tel
        txtNo = .No
        txtImport = .ImportList
        txtExport = .ExportList
        
        If PGP.Valid Then
            txtKeys = .KeysPath
            If IsDir(.KeysPath) Then
                'cmdKeys.Enabled = False
            Else
                txtKeys = None
            End If
        Else
            txtKeys = None
            lblKeys.Enabled = False
            txtKeys.Enabled = False
            cmdKeys.Enabled = False
        End If
    End With
End Sub
