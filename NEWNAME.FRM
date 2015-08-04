VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} NewName 
   Caption         =   "Реквизиты получателя"
   ClientHeight    =   3600
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6000
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

Public Mode As String 'New,Edit,Pay

Private Sub cmdBank_Click()
    AutoCheck
    BnkSeek2.MsgInfo txtBIC
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    Dim i As Long, s As String
    On Error Resume Next
    
    If User.Demo And Mode <> "New" Then GoTo SkipCheck
    
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
    s = SetLSKey(txtBIC, txtLS)
    If s <> txtLS Then
        If YesNoBox("БИК %s не соответвует Л/С:\n%s (введено)\n%s (должно быть)\n\nИсправить?", txtBIC, txtLS, s) Then
            txtLS = s
            txtLS.SetFocus
            Exit Sub
        End If
    End If
    If Not INNKeyValid(txtINN) Then
        If YesNoBox("ИНН неправильный!\n\nИсправить?") Then
            txtINN.SetFocus
            Exit Sub
        End If
    End If
    If txtKPP.TextLength <> 9 Then
        If YesNoBox("Длина КПП не 9 символов!\n\nИсправить?") Then
            txtKPP.SetFocus
            Exit Sub
        End If
    End If
    If Left(txtKPP.Text, 4) <> Left(txtINN.Text, 4) Then
        If YesNoBox("Коды налоговой инспекции в ИНН и КПП не совпадают!\n\nИсправить?") Then
            txtINN.SetFocus
            Exit Sub
        End If
    End If
    If txtName.TextLength = 0 Then
        WarnBox "Не введено название организации!"
        txtName.SetFocus
        Exit Sub
    End If
    s = txtName.Text
    i = InStr(s, """")
    If i = 0 Then
        'skip
    ElseIf i = 1 Then
        If YesNoBox("Название не начинается с формы собственности (ООО, ЗАО и т.п.)!\n\nИсправить?") Then
            txtName.SetFocus
            Exit Sub
        End If
    ElseIf Mid(s, i - 1, 1) <> " " Then
        If YesNoBox("Перед кавычкой принято ставить пробел.\n\nИсправить?") Then
            txtName.SetFocus
            Exit Sub
        End If
    End If
    
SkipCheck:
    Hide
    If Mode = "Pay" Then
        With Payment
            .Name = txtName
            .INN = txtINN
            .KPP = txtKPP
            .BIC = txtBIC
            .LS = txtLS
        End With
    ElseIf Mode = "New" Then
        With User
            .Add txtBIC, txtLS, txtKPP, txtINN, txtName
            .ID = txtLS
        End With
    ElseIf Mode = "Edit" Then
        With User
            .INN = txtINN
            .KPP = txtKPP
            .Name = txtName
            .LS = txtLS
            .BIC = txtBIC
            .ResetCaption
        End With
    End If
    Unload Me
End Sub

Private Sub txtBIC_Change()
    With txtBIC
        lblLenBIC = Bsprintf("%d/%d", .TextLength, .MaxLength)
    End With
End Sub

Private Sub txtBIC_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    'AutoFill
    'On Error Resume Next
    'txtBIC = txtBIC & Right(App.DefBIC, 9 - txtBIC.TextLength)
End Sub

Private Sub txtINN_Change()
    With txtINN
        lblLenINN = Bsprintf("%d/%d", .TextLength, .MaxLength)
    End With
End Sub

Private Sub txtKPP_Change()
    With txtKPP
        lblLenKPP = Bsprintf("%d/%d", .TextLength, .MaxLength)
    End With
End Sub

Private Sub txtLS_Change()
    With txtLS
        lblLenLS = Bsprintf("%d/%d", .TextLength, .MaxLength)
    End With
End Sub

Private Sub txtLS_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    'AutoFill
    'Dim i As Long
    'On Error Resume Next
    'txtBIC = txtBIC & Right(App.DefBIC, 9 - txtBIC.TextLength)
    'i = InStr(txtLS, "-")
    'If i = 0 Then
    '    txtLS = txtLS & Right(App.DefLS, 20 - txtLS.TextLength)
    'Else
    '    txtLS = Left(txtLS, i - 1) & _
    '        Mid(App.DefLS, i, 21 - txtLS.TextLength) & _
    '        Right(txtLS, txtLS.TextLength - i)
    'End If
    'txtLS = SetLSKey(txtBIC, txtLS)
End Sub

Private Sub txtName_Change()
    With txtName
        lblLenName = Bsprintf("%d/%d", .TextLength, .MaxLength)
    End With
End Sub

Private Sub UserForm_Activate()
    Dim i As Long, s As String
    On Error Resume Next
    Select Case Mode
        Case "Pay":
            With Payment
                txtBIC = .BIC
                txtLS = .LS
                txtINN = .INN
                txtName = .Name
                txtKPP = .KPP 'после Name! (for conversion)
            End With
        Case "New"
            Caption = "Реквизиты нового плательщика"
            With User
                txtBIC = .BIC
                txtLS = .LS
                txtINN = .INN
                txtName = .Name
                txtKPP = .KPP
            End With
        Case "Edit"
            Caption = "Исправление реквизитов плательщика"
            With User
                txtBIC = .BIC
                txtLS = .LS
                txtINN = .INN
                txtName = .Name
                txtKPP = .KPP
            End With
    End Select
End Sub

Private Sub UserForm_Initialize()
    Mode = "Edit"
    txtBIC.ControlTipText = App.DefBIC
    txtLS.ControlTipText = App.DefLS
End Sub
