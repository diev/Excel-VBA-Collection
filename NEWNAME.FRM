VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} NewName 
   Caption         =   "Реквизиты получателя"
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5505
   HelpContextID   =   430000
   OleObjectBlob   =   "NewName.frx":0000
   StartUpPosition =   1  'CenterOwner
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
End
Attribute VB_Name = "NewName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit
Option Compare Text
Option Base 1
DefLng A-Z

Public Mode As String 'New,Edit,Pay,BIC,LS

Private Sub cmdBank_Click()
    Dim s As String
    AutoCheck
    With BnkSeek2
        If .Valid Then
            DoEvents 'for update
            .MsgInfo txtBIC
        Else
            WarnBox .None
        End If
    End With
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdImportID_Click()
    With User
        User.LocateID App.Path & Right(.LS, 4) & ".ID" '"A:\"
        txtBIC = .BIC
        txtLS = .LS
        txtINN = .INN
        cboName = .Name
    End With
End Sub

Private Sub cmdOk_Click()
    Dim i As Long, s As String
    On Error Resume Next
    If txtBIC.TextLength <> 9 Then
        WarnBox "Длина БИК Банка не 9 символов!"
        txtBIC.SetFocus
        Exit Sub
    End If
    If txtLS.TextLength <> 20 Then
        WarnBox "Длина Л/С не 20 символов!"
        txtLS.SetFocus
        Exit Sub
    End If
    If Right(txtLS, 7) = "0000000" And Not User.Demo Then
        If YesNoBox("Л/С подозрительно кончается на нули!\n\nИсправить?") Then
            txtLS.SetFocus
            Exit Sub
        End If
    End If
    If txtINN.TextLength < 10 Then
        If YesNoBox("Длина ИНН меньше 10 символов!\n\nИсправить?") Then
            txtINN.SetFocus
            Exit Sub
        End If
    End If
    If cboName.TextLength = 0 Then
        WarnBox "Не введено название организации!"
        cboName.SetFocus
        Exit Sub
    End If
    Hide
    If Mode = "Pay" Then
        With Payment
            .Name = cboName
            .INN = txtINN
            .BIC = txtBIC
            .LS = txtLS
        End With
    ElseIf Mode = "New" Then
        With User
            .ID = txtLS
            .Name = cboName
            .INN = txtINN
            .LS = txtLS
            .BIC = txtBIC
            UserPrivateShow
        End With
    ElseIf Mode = "Edit" Then
        With User
            .INN = txtINN
            .Name = cboName
            If Not cboName.MatchFound Then
                .NameN(cboName.ListCount + 1) = cboName
            End If
            .LS = txtLS
            .BIC = txtBIC
            .ResetCaption
            'InfoBox "Данные исправлены"
        End With
        UserPrivateShow
    End If
    Unload Me
End Sub

Private Sub optBank_Click()
    On Error Resume Next
    DoEvents 'for update
    txtLS = SetLSKey(txtBIC, txtLS, optRKC.Value)
End Sub

Private Sub optRKC_Click()
    On Error Resume Next
    DoEvents 'for update
    txtLS = SetLSKey(txtBIC, txtLS, optRKC.Value)
End Sub

Private Sub txtBIC_Change()
    With txtBIC
        lblLenBIC = Bsprintf("%d/%d", .TextLength, .MaxLength)
    End With
End Sub

Private Sub txtBIC_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    On Error Resume Next
    txtBIC = txtBIC & Right(App.DefBIC, 9 - txtBIC.TextLength)
    If Val(Right(txtBIC, 3)) < 3 Then
        If Not optRKC Then
            WarnBox "Похоже, что счет в РКЦ - автопереключение!"
            optRKC = True
        End If
    Else
        If Not optBank Then
            WarnBox "Похоже, что счет в банке - автопереключение!"
            optBank = True
        End If
    End If
End Sub

Private Sub txtINN_Change()
    With txtINN
        lblLenINN = Bsprintf("%d/%d", .TextLength, .MaxLength)
    End With
End Sub

Private Sub txtLS_Change()
    With txtLS
        lblLenLS = Bsprintf("%d/%d", .TextLength, .MaxLength)
    End With
End Sub

Private Sub txtLS_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Dim i As Long
    On Error Resume Next
    txtBIC = txtBIC & Right(App.DefBIC, 9 - txtBIC.TextLength)
    i = InStr(txtLS, "-")
    If i = 0 Then
        txtLS = txtLS & Right(App.DefLS, 20 - txtLS.TextLength)
    Else
        txtLS = Left(txtLS, i - 1) & _
            Mid(App.DefLS, i, 21 - txtLS.TextLength) & _
            Right(txtLS, txtLS.TextLength - i)
    End If
    txtLS = SetLSKey(txtBIC, txtLS, optRKC.Value)
End Sub

Private Sub cboName_Change()
    With cboName
        lblLenName = Bsprintf("%d/%d", .TextLength, .MaxLength)
    End With
End Sub

Private Sub UserForm_Activate()
    Dim i As Long, s As String
    On Error Resume Next
    cboName.ShowDropButtonWhen = fmShowDropButtonWhenNever
    Select Case Mode
        Case "Pay":
            With Payment
                If .Valid Then
                    txtBIC = .BIC
                    txtLS = .LS
                End If
                txtINN = .INN
                cboName = .Name
            End With
        Case "BIC":
            Caption = "Справочник БИК"
            lblLS.Enabled = False
            txtLS.Enabled = False
            txtLS.BackStyle = fmBackStyleTransparent
            lblLenLS.Enabled = False
            lblINN.Enabled = False
            txtINN.Enabled = False
            txtINN.BackStyle = fmBackStyleTransparent
            lblLenINN.Enabled = False
            lblName.Enabled = False
            cboName.Enabled = False
            cboName.BackStyle = fmBackStyleTransparent
            lblLenName.Enabled = False
            cmdOk.Enabled = False
        Case "LS":
            Caption = "Справочник БИК и проверка реквизитов"
            lblINN.Enabled = False
            txtINN.Enabled = False
            lblLenINN.Enabled = False
            txtINN.BackStyle = fmBackStyleTransparent
            lblName.Enabled = False
            cboName.Enabled = False
            cboName.BackStyle = fmBackStyleTransparent
            lblLenName.Enabled = False
            cmdOk.Enabled = False
        Case "New"
            Caption = "Реквизиты нового плательщика"
            With User
                txtBIC = .BIC
                txtLS = .LS
                txtINN = .INN
                cboName = .Name
            End With
            cmdOk.Caption = "Дальше..."
            cmdImportID.Visible = True
            'cmdImportID_Click
        Case "Edit"
            Caption = "Исправление реквизитов плательщика"
            lblBIC.Enabled = False
            txtBIC.Enabled = False
            txtBIC.BackStyle = fmBackStyleTransparent
            lblLenBIC.Enabled = False
            lblLS.Enabled = False
            txtLS.Enabled = False
            txtLS.BackStyle = fmBackStyleTransparent
            lblLenLS.Enabled = False
            fraKey.Enabled = False
            optBank.Enabled = False
            optRKC.Enabled = False
            With User
                txtBIC = .BIC
                txtLS = .LS
                txtINN = .INN
                i = 0
                cboName.Clear
                cboName.ShowDropButtonWhen = fmShowDropButtonWhenAlways
                Do
                    i = i + 1
                    s = .NameN(i)
                    If Len(s) = 0 Then Exit Do
                    cboName.AddItem s
                Loop
                cboName = .Name
            End With
            cmdOk.Caption = "Дальше..."
        Case Else
            'Alert Undefined Mode!
    End Select
End Sub

Private Sub UserForm_Initialize()
    Mode = "Edit"
    txtBIC.ControlTipText = App.DefBIC
    txtLS.ControlTipText = App.DefLS
End Sub
