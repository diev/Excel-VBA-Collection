VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PlatEnter 
   Caption         =   "Нет получателя!"
   ClientHeight    =   3585
   ClientLeft      =   30
   ClientTop       =   270
   ClientWidth     =   5505
   OleObjectBlob   =   "PlatEnter.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "PlatEnter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Option Compare Text
Option Base 1
DefLng A-Z

Private mLoading As Boolean

Private Sub CalcTax()
    Dim n As Variant, t As String, x As Variant
    n = RVal(cboTax.Text)
    If n = 0 Then
        n = "нет"
        t = "НДС не облагается."
    Else
        If InStr(1, cboTax, "%") = 0 Then
            cboTax.Text = cboTax.Text & "%"
        End If
        n = PlatFormat(Sum2Tax(RVal(txtSum), n))
        t = BPrintF("В том числе НДС %s: %s.", cboTax.Text, n)
    End If
    cmdTaxAdd.Caption = n
    cmdTaxAdd.ControlTipText = t
End Sub

Private Sub CalcTaxAdd()
    'If Len(txtDetails) = 0 Then
    '    txtDetails = cmdTaxAdd.ControlTipText
    'ElseIf Right(txtDetails, 1) = " " Then
    '    txtDetails = txtDetails & cmdTaxAdd.ControlTipText
    'Else
    '    txtDetails = txtDetails & " " & cmdTaxAdd.ControlTipText
    'End If
    txtDetails.SelText = cmdTaxAdd.ControlTipText
End Sub

Private Sub cboTax_Change()
    CalcTax
End Sub

Private Sub cmdPayee_Click()
    PayUserShow
    With Payment
        If Len(.Name) > 0 Then
            Caption = "Получатель: " & .Name
            cmdPayee.ControlTipText = .Name
            cmdPayee.Font.Bold = False
        End If
    End With
End Sub

Private Sub cmdTaxAdd_Click()
    CalcTaxAdd
End Sub

Private Sub sbrRows_Change()
    With sbrRows
        .ControlTipText = BPrintF("Строка %d", .Value)
    End With
    If mLoading Then Exit Sub
    With Payment
        Application.GoTo Worksheets(User.ID).Range("$A$1").Cells(sbrRows.Value), False
        .ReadRow sbrRows.Value
        If Len(.Name) = 0 Then
            Caption = "Нет получателя!"
            cmdPayee.Font.Bold = True
            lblNo = "Номер:"
            lblDate = "Дата:"
            txtSum = vbNullString
            txtDetails = vbNullString
        Else
            Caption = "Получатель: " & .Name
            cmdPayee.Font.Bold = False
            cmdPayee.ControlTipText = .Name
            lblNo = BPrintF("Номер %d:", .DocNo)
            lblDate = BPrintF("Дата %n:", .DocDate)
            cboQueue = CStr(.Queue)
            txtSum = PlatFormat(.Sum)
            txtDetails = .Details
        End If
    End With
End Sub

Private Sub spnDate_SpinDown()
    On Error Resume Next
    txtDate = PlatDate(DateAdd("d", -1, RDate(txtDate.Text)))
End Sub

Private Sub spnDate_SpinUp()
    On Error Resume Next
    txtDate = PlatDate(DateAdd("d", 1, RDate(txtDate.Text)))
End Sub

Private Sub txtDetails_Change()
    With txtDetails
        lblLenDetails = BPrintF("%d/%d", .TextLength, .MaxLength)
        'If InStr(1, .Text, "  ") > 0 Then
        '    MsgBox "Не надо вводить лишние пробелы!", vbExclamation, App.Title
        '    .SetFocus
        'End If
    End With
End Sub

Private Sub txtSum_Change()
    With txtSum
        .ControlTipText = RSumStr(RVal(.Text))
        CalcTax
    End With
End Sub

Private Sub cmdCancel_Click()
    Hide
    'ReadArchive 'restore back
    Unload Me
End Sub

Private Sub cmdOk_Click()
    Dim s As String
    If Val(txtNo) = 0 Then
        MsgBox "Не введен номер поручения!", vbExclamation, App.Title
        txtNo.SetFocus
        Exit Sub
    End If
    If Val(txtNo) > User.NoMax Then
        MsgBox "Номер поручения превышает допустимый предел!", vbExclamation, App.Title
        txtNo.SetFocus
        Exit Sub
    End If
    If Val(txtNo) < User.NoMin Then
        MsgBox "Номер поручения ниже допустимого предела!", vbExclamation, App.Title
        txtNo.SetFocus
        Exit Sub
    End If
    If RVal(txtSum) = 0 Then
        MsgBox "Не введена сумма платежа!", vbExclamation, App.Title
        txtSum.SetFocus
        Exit Sub
    End If
    s = txtDetails
    If InStr(1, txtDetails, "^") > 0 Then
        MsgBox BPrintF("Нельзя вводить символ \'^\'!"), vbExclamation, App.Title
        txtDetails.SetFocus
        Exit Sub
    End If
    With txtDetails
        If InStr(1, .Text, "  ") > 0 Then
            MsgBox "Не надо вводить лишние пробелы!", vbExclamation, App.Title
            .Text = StrSpaces1(s)
        End If
        If txtDetails.TextLength = 0 Then
            MsgBox "Не введено назначение платежа!", vbExclamation, App.Title
            txtDetails.SetFocus
            Exit Sub
        End If
    End With
    With Payment
        If Len(.Name) = 0 Then
            MsgBox "Не введен получатель платежа!", vbExclamation, App.Title
            cmdPayee_Click
            Exit Sub
        End If
        If Len(.BIC) = 0 Then
            MsgBox "Не введен банк получателя платежа!", vbExclamation, App.Title
            cmdPayee_Click
            Exit Sub
        End If
        If Len(.LS) = 0 Then
            MsgBox "Не введен л/с получателя платежа!", vbExclamation, App.Title
            cmdPayee_Click
            Exit Sub
        End If
        If Len(.INN) = 0 Then
            MsgBox "Не введен ИНН получателя платежа!", vbExclamation, App.Title
            cmdPayee_Click
            Exit Sub
        End If
        Hide
        .Mark = "?"
        .DocNo = Val(txtNo)
        .DocDate = RDate(txtDate)
        .Queue = cboQueue.Value
        .Sum = RVal(txtSum)
        .Details = txtDetails
        User.No = .DocNo + 1
        User.AmountMinus .Sum
        .FileName = .FileName 'autogeneration
        .WriteRow -1
        Application.GoTo Worksheets(User.ID).Range("$A$1").Cells(.Row), False
    End With
    Unload Me
End Sub

Private Sub spnDate_Change()
    spnDate = 0
End Sub

Private Sub txtDetails_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    With txtDetails
        .Text = StrSpaces1(.Text)
    End With
End Sub

Private Sub txtSum_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    txtSum = PlatFormat(RVal(txtSum))
End Sub

Private Sub UserForm_Initialize()
    Dim n As Long
    mLoading = True
    With Payment
        .ReadRow 0
        If Len(.Name) = 0 Then
            cmdPayee.Font.Bold = True
        Else
            Caption = "Получатель: " & .Name
            cmdPayee.ControlTipText = .Name
            lblNo = BPrintF("Номер %d:", .DocNo)
            lblDate = BPrintF("Дата %n:", .DocDate)
            cboQueue = CStr(.Queue)
            txtSum = PlatFormat(.Sum)
            txtDetails = .Details
        End If
        n = .RowsCount + 1
        sbrRows.Max = n
        If .Row > n Then .Row = n
        sbrRows.Value = .Row
    End With
    txtDate = PlatDate(Date)
    txtDate.ControlTipText = "Сегодня " & PlatDate(Date)
    txtNo = CStr(User.No)
    With cboTax
        .AddItem "нет"
        .AddItem "10%"
        .AddItem "20%"
    End With
    With cboQueue
        .AddItem "1"
        .AddItem "2"
        .AddItem "3"
        .AddItem "4"
        .AddItem "5"
        .AddItem "6"
    End With
    mLoading = False
End Sub

