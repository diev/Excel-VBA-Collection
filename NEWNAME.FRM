VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} NewName 
   Caption         =   "Ввод новых реквизитов"
   ClientHeight    =   4695
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5415
   OleObjectBlob   =   "NewName.frx":0000
   StartUpPosition =   2  'CenterScreen
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

Public AddNewOk As Boolean

Private Sub cboBIC_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    cboBIC = AutoFillBIC(cboBIC)
End Sub

Private Sub cmdBank_Click()
    Dim s As String, k
    With txtLS
        If .TextLength <> 20 Then
            MsgBox "Длина лицевого счета не 20 знаков!", vbExclamation, AppTitle
            Exit Sub
        End If
        k = InStr(1, .Text, "К")
        If k = 9 Then
            s = .Text
            Mid(s, k, 1) = CStr(CalcBankKey(cboBIC, txtLS))
            .Text = s
            MsgBox BPrintF("Ключ счета откорректирован:\nПроверьте лицевой счет!"), vbExclamation, AppTitle
            .SetFocus
            Exit Sub
        ElseIf k > 0 Then
            MsgBox BPrintF("Ключ счета не на месте:\nПроверьте лицевой счет!"), vbCritical, AppTitle
            .SetFocus
            Exit Sub
        End If
    End With
    If cboBIC.TextLength <> 9 Then
        MsgBox "Длина БИК не 9 знаков!", vbExclamation, AppTitle
        Exit Sub
    End If
    If Not BankKeyValid(cboBIC, txtLS) Then
        s = "Такие реквизиты не соответствуют ключу счета в КО"
        If RkcKeyValid(cboBIC, txtLS) Then
            s = BPrintF("%s,\nно соответствуют ключу счета в РКЦ!", s)
        Else
            s = BPrintF("%s\nи не соответствуют ключу счета в РКЦ!", s)
            MsgBox s, vbCritical, AppTitle
            lblBank.Caption = "Указанные реквизиты не соответствуют друг другу!"
            txtLS.SetFocus
            Exit Sub
        End If
    End If
    cmdBank.Enabled = False
    With Bank
        .BIC = cboBIC
        lblBank.Caption = .FullInfo
    End With
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdNick_Click()
    Dim s As String
    If txtNick.TextLength = 0 Then
        MsgBox "Не указано краткое имя", vbExclamation, AppTitle
        txtNick.SetFocus
        Exit Sub
    End If
    If txtINN.TextLength <> 10 Then
        MsgBox "ИНН должен содержать 10 знаков", vbExclamation, AppTitle
        txtINN.SetFocus
        Exit Sub
    End If
    s = Payment.GetNickUnique(txtNick)
    If s <> txtNick Then
        txtNick = s
        MsgBox BPrintF("Предлагается другой вариант:\n~%s~\nИсправьте, если не удовлетворяет", s), vbExclamation, AppTitle
        txtNick.SetFocus
    End If
    cmdNick.Enabled = False
End Sub

Private Sub cmdOk_Click()
    Dim arr As Variant, s As String
    If cmdNick.Enabled Then
        MultiPage.Value = 0
        MsgBox "Нет проверки уникальности", vbExclamation, AppTitle
        cmdNick.SetFocus
        Exit Sub
    End If
    If cmdBank.Enabled Then
        MultiPage.Value = 1
        MsgBox "Нет проверки Банка", vbExclamation, AppTitle
        cmdBank.SetFocus
        Exit Sub
    End If
    Payment.AddPayee txtNick, cboName, txtINN, txtLS, cboBIC, cboCustom1, cboCustom2
    AddNewOk = True
    Unload Me
End Sub

Private Sub cboBIC_Change()
    lblLenBIC = CStr(cboBIC.TextLength)
    cmdBank.Enabled = True
End Sub

Private Sub txtINN_Change()
    lblLenINN = CStr(txtINN.TextLength)
End Sub

Private Sub txtLS_Change()
    lblLenLS = CStr(txtLS.TextLength)
    cmdBank.Enabled = True
End Sub

Private Sub txtLS_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    txtLS = AutoFillLS(txtLS)
End Sub

Private Sub txtNick_Change()
    cmdNick.Enabled = True
End Sub

Private Sub txtNick_Enter()
    If txtNick.TextLength = 0 Then
        txtNick.Text = NickByName(cboName.Text)
    End If
End Sub

Private Sub UserForm_Initialize()
    Dim arr As Variant
    MultiPage.Value = 0
    cboName.List = Payment.ListNames10
    cboCustom1.List = Payment.CustomList(1)
    cboCustom1.List = Payment.CustomList(2)
    cboBIC.List = Payment.ListBIC10
    lblBnkSeek.Caption = FileVerInfo(Range("BnkSeek").Text, , True)
End Sub
