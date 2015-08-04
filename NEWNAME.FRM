VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} NewName 
   Caption         =   "Реквизиты получателя"
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5505
   OleObjectBlob   =   "NewName.frx":0000
   StartUpPosition =   1  'CenterOwner
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
            s = AutoFillBIC(txtBIC)
            txtBIC = s
            DoEvents 'Some problems with updating TextBox before next...
            .MsgInfo s
        Else
            MsgBox .None, vbExclamation, App.Title
        End If
    End With
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    On Error Resume Next
    If txtBIC.TextLength <> 9 Then
        MsgBox "Длина БИК Банка не 9 символов!", vbExclamation, App.Title
        txtBIC.SetFocus
        Exit Sub
    End If
    If txtLS.TextLength <> 20 Then
        MsgBox "Длина Л/С не 20 символов!", vbExclamation, App.Title
        txtLS.SetFocus
        Exit Sub
    End If
    If Right(txtLS, 3) = "000" And Not User.Demo Then
        If MsgBox(BPrintF("Л/С подозрительно кончается на нули!\n\nИсправить?"), _
            vbQuestion + vbYesNo, App.Title) = vbYes Then
            txtLS.SetFocus
            Exit Sub
        End If
    End If
    If txtINN.TextLength < 10 Then
        If MsgBox(BPrintF("Длина ИНН меньше 10 символов!\n\nИсправить?"), _
            vbQuestion + vbYesNo, App.Title) = vbYes Then
            txtINN.SetFocus
            Exit Sub
        End If
    End If
    If txtName.TextLength = 0 Then
        MsgBox "Не введено название организации!", vbExclamation, App.Title
        txtName.SetFocus
        Exit Sub
    End If
    Hide
    If Mode = "Pay" Then
        With Payment
            .Name = txtName
            .INN = txtINN
            .BIC = txtBIC
            .LS = txtLS
        End With
    ElseIf Mode = "New" Then
        With User
            .ID = Right(txtLS, 3)
            If .ID = Right(txtLS, 3) Then
                .Name = txtName
                .INN = txtINN
                .LS = txtLS
                .BIC = txtBIC
                 UserPrivateShow
            End If
        End With
    ElseIf Mode = "Edit" Then
        With User
            .INN = txtINN
            .Name = txtName
            .LS = txtLS
            .BIC = txtBIC
            'MsgBox "Данные исправлены", vbInformation, App.Title
        End With
        UserPrivateShow
    Else 'Error?
        MsgBox "А что, собственно, Вы здесь набрали?", vbCritical, App.Title
    End If
    Unload Me
End Sub

Private Sub optBank_Click()
    Dim s As String
    If txtLS.TextLength > 0 Then
        s = AutoFillBIC(txtBIC)
        txtBIC = s
        txtLS = AutoFillLS(s, txtLS, optRKC.Value)
        DoEvents
    End If
End Sub

Private Sub optRKC_Click()
    Dim s As String
    If txtLS.TextLength > 0 Then
        s = AutoFillBIC(txtBIC)
        txtBIC = s
        txtLS = AutoFillLS(s, txtLS, optRKC.Value)
        DoEvents
    End If
End Sub

Private Sub txtBIC_Change()
    With txtBIC
        lblLenBIC = BPrintF("%d/%d", .TextLength, .MaxLength)
    End With
End Sub

Private Sub txtBIC_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    txtBIC = AutoFillBIC(txtBIC)
    DoEvents
    If Val(Right(txtBIC, 3)) < 3 Then
        If Not optRKC Then
            MsgBox "Похоже, что счет в РКЦ - автопереключение!", vbExclamation, App.Title
            optRKC = True
        End If
    Else
        If Not optBank Then
            MsgBox "Похоже, что счет в банке - автопереключение!", vbExclamation, App.Title
            optBank = True
        End If
    End If
End Sub

Private Sub txtINN_Change()
    With txtINN
        lblLenINN = BPrintF("%d/%d", .TextLength, .MaxLength)
    End With
End Sub

Private Sub txtLS_Change()
    With txtLS
        lblLenLS = BPrintF("%d/%d", .TextLength, .MaxLength)
    End With
End Sub

Private Sub txtLS_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Dim s As String
    s = AutoFillBIC(txtBIC)
    txtBIC = s
    txtLS = AutoFillLS(s, txtLS, optRKC.Value)
    DoEvents
End Sub

Private Sub txtName_Change()
    With txtName
        lblLenName = BPrintF("%d/%d", .TextLength, .MaxLength)
    End With
End Sub

Private Sub UserForm_Activate()
    Select Case Mode
        Case "Pay":
            With Payment
                If .Valid Then
                    txtBIC = .BIC
                    txtLS = .LS
                End If
                txtINN = .INN
                txtName = .Name
            End With
        Case "BIC":
            Caption = "Справочник банков РФ"
            lblLS.Enabled = False
            txtLS.Enabled = False
            txtLS.BackStyle = fmBackStyleTransparent
            lblLenLS.Enabled = False
            lblINN.Enabled = False
            txtINN.Enabled = False
            txtINN.BackStyle = fmBackStyleTransparent
            lblLenINN.Enabled = False
            lblName.Enabled = False
            txtName.Enabled = False
            txtName.BackStyle = fmBackStyleTransparent
            lblLenName.Enabled = False
            cmdOk.Enabled = False
        Case "LS":
            Caption = "Справочник банков РФ и проверка реквизитов"
            lblINN.Enabled = False
            txtINN.Enabled = False
            lblLenINN.Enabled = False
            txtINN.BackStyle = fmBackStyleTransparent
            lblName.Enabled = False
            txtName.Enabled = False
            txtName.BackStyle = fmBackStyleTransparent
            lblLenName.Enabled = False
            cmdOk.Enabled = False
        Case "New"
            Caption = "Реквизиты нового плательщика (исправьте отличия)"
            With User
                txtBIC = .BIC
                txtLS = .LS
                txtINN = .INN
                txtName = .Name
            End With
            cmdOk.Caption = "Дальше..."
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
                txtName = .Name
            End With
            cmdOk.Caption = "Дальше..."
        Case Else
            'Alert Undefined Mode!
    End Select
End Sub

Private Sub UserForm_Initialize()
    'If Not BnkSeek2.Valid Then
    '    lblBIC = lblBIC & " (" & BnkSeek2.None & ")"
    '    cmdBank.Enabled = False
    'End If
    Mode = "Edit"
    txtBIC.ControlTipText = App.DefBIC
    txtLS.ControlTipText = App.DefLS
End Sub
