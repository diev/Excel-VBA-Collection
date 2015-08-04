VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserPrivate 
   Caption         =   "Ключи PGP и прочее плательщика"
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5505
   HelpContextID   =   460000
   OleObjectBlob   =   "UserPrivate.frx":0000
   StartUpPosition =   1  'CenterOwner
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
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
    Hide
    With User
        .Sign(1) = txtSign1
        .Sign(2) = txtSign2
        .Tel = txtTel
        .No = txtNo
        .NoMin = txtNoMin
        .NoMax = txtNoMax
        If txtKeys <> None Then .KeysPath = txtKeys
        .ImportList = txtImport
        .ExportList = txtExport
    End With
    Unload Me
End Sub

Private Sub cmdKeys_Click()
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

Private Sub txtKeys_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    cmdKeys_Click
End Sub

Private Sub UserForm_Initialize()
    On Error Resume Next
    With User
        Caption = Caption & Bsprintf(" %s", .ID)
        spnNo.Min = CLng(.NoMin)
        spnNo.Max = CLng(.NoMax)
        
        txtSign1 = .Sign(1)
        txtSign2 = .Sign(2)
        txtTel = .Tel
        txtNo = .No
        txtNoMin = .NoMin
        txtNoMax = .NoMax
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
