VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PlatEnter 
   Caption         =   "��� ����������!"
   ClientHeight    =   3600
   ClientLeft      =   30
   ClientTop       =   270
   ClientWidth     =   6000
   HelpContextID   =   440000
   OleObjectBlob   =   "PlatEnter.frx":0000
   StartUpPosition =   1  'CenterOwner
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
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
    Dim n As Variant, T As String, x As Variant
    n = RVal(cboTax.Text)
    If n = 0 Then
        n = "���"
        T = "��� �� ����������."
    Else
        If InStr(1, cboTax, "%") = 0 Then
            cboTax.Text = cboTax.Text & "%"
        End If
        n = PlatFormat(Sum2Tax(RVal(txtSum), n))
        T = Bsprintf("� ��� ����� ��� %s: %s.", cboTax.Text, n)
    End If
    cmdTaxAdd.Caption = n
    cmdTaxAdd.ControlTipText = T
End Sub

Private Sub CalcTaxAdd()
    Dim i As Long
    i = InStr(txtDetails, "!")
    If i > 0 Then
        txtDetails = Left(txtDetails, i) & cmdTaxAdd.ControlTipText
    Else
        txtDetails = txtDetails & "!" & cmdTaxAdd.ControlTipText
    End If
End Sub

Private Sub cboTax_Change()
    CalcTax
End Sub

Private Sub cmdPayee_Click()
    PayUserShow
    With Payment
        If Len(.Name) > 0 Then
            Caption = "����������: " & .Name
            cmdPayee.ControlTipText = .Name
            cmdPayee.Font.Bold = False
        End If
    End With
End Sub

Private Sub cmdSS_Click()
    PlatNAL.Show
    cmdSS.Caption = Payment.SS
End Sub

Private Sub cmdTaxAdd_Click()
    CalcTaxAdd
End Sub

Private Sub sbrRows_Change()
    With sbrRows
        .ControlTipText = Bsprintf("������ %d", .Value)
    End With
    If mLoading Then Exit Sub
    With Payment
        Application.GoTo Worksheets(User.ID).Range("$A$1").Cells(sbrRows.Value), False
        .ReadRow sbrRows.Value
        If Len(.Name) = 0 Then
            Caption = "��� ����������!"
            cmdPayee.Font.Bold = True
            lblNo = "�����:"
            lblDate = "����:"
            txtSum = vbNullString
            txtDetails = vbNullString
        Else
            Caption = "����������: " & .Name
            cmdPayee.Font.Bold = False
            cmdPayee.ControlTipText = .Name
            lblNo = Bsprintf("����� %d:", .DocNo)
            lblDate = Bsprintf("���� %n:", .DocDate)
            cboQueue = CStr(.Queue)
            txtSum = PlatFormat(.Sum)
            txtDetails = .Details
        End If
    End With
End Sub

Private Sub spnDate_Change()
    On Error Resume Next
    txtDate = PlatDate(DateAdd("d", spnDate.Value, RDate(txtDate.Text)))
    spnDate.Value = 0
End Sub

Private Sub spnNo_Change()
    Dim n As Long
    On Error Resume Next
    n = txtNo.Value + spnNo.Value
    If n > 999 Then
        n = 1
    ElseIf n < 1 Then
        n = 999
    End If
    txtNo = CStr(n)
    spnNo.Value = 0
End Sub

Private Sub txtDetails_Change()
    With txtDetails
        lblLenDetails = Bsprintf("%d/%d", .TextLength, .MaxLength)
    End With
End Sub

Private Sub txtSum_Change()
    With txtSum
        .ControlTipText = RSumStr(RVal(.Text))
        CalcTax
    End With
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    Dim s As String
    
    If User.Demo Then GoTo SkipCheck
    
    If Val(txtNo) = 0 Then
        WarnBox "�� ������ ����� ���������!"
        txtNo.SetFocus
        Exit Sub
    End If
    If Val(txtNo) > User.NoMax Then
        WarnBox "����� ��������� ��������� ���������� ������!"
        txtNo.SetFocus
        Exit Sub
    End If
    If Val(txtNo) < User.NoMin Then
        WarnBox "����� ��������� ���� ����������� �������!"
        txtNo.SetFocus
        Exit Sub
    End If
    If RVal(txtSum) = 0 Then
        WarnBox "�� ������� ����� �������!"
        txtSum.SetFocus
        Exit Sub
    End If
    s = Left(Payment.LS, 5)
    If InStr(1, "30122|30123|30230|30231|40807|40813|40814|40815|40818|40819|40820", _
        s) > 0 Then
        If Left(txtDetails, 3) <> "{VO" Then
            WarnBox "�� ������ ������� ������ {VO\n��� ��������� � ������������!\n\n��������� � ����� ��������� ��������."
            txtDetails.SetFocus
            Exit Sub
        End If
    End If
    s = txtDetails
    If InStr(1, txtDetails, "^") > 0 Then
        WarnBox "������ ������� ������ \'^\'!"
        txtDetails.SetFocus
        Exit Sub
    End If
    With txtDetails
        If InStr(1, .Text, "  ") > 0 Then
            WarnBox "�� ���� ������� ������ �������!"
            .Text = StrSpaces1(s)
        End If
        If txtDetails.TextLength = 0 Then
            WarnBox "�� ������� ���������� �������!"
            txtDetails.SetFocus
            Exit Sub
        End If
    End With
    With Payment
        If Len(.Name) = 0 Then
            WarnBox "�� ������ ���������� �������!"
            cmdPayee_Click
            Exit Sub
        End If
        If Len(.BIC) = 0 Then
            WarnBox "�� ������ ���� ���������� �������!"
            cmdPayee_Click
            Exit Sub
        End If
        If Len(.LS) = 0 Then
            WarnBox "�� ������ �/� ���������� �������!"
            cmdPayee_Click
            Exit Sub
        End If
    End With
        
SkipCheck:
    With Payment
        Hide
        .Mark = "?"
        .DocNo = Val(txtNo)
        .DocDate = RDate(txtDate)
        .Queue = cboQueue.Value
        .Sum = RVal(txtSum)
        .Details = txtDetails
        
        s = .ValidationError
        If Len(s) > 0 Then
            WarnBox "��������������!\n��������� �������� �� �������������\n�������� �������� ��������:\n\n%s!", s
        End If
        
        User.No = .DocNo + 1
        '.FileName = .FileName 'autogeneration
        .WriteRow -1
        Application.GoTo Worksheets(User.ID).Range("$A$1").Cells(.Row), False
    End With
    Unload Me
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
    On Error Resume Next
    mLoading = True
    With Payment
        .ReadRow 0
        If Len(.Name) = 0 Then
            cmdPayee.Font.Bold = True
        Else
            Caption = "����������: " & .Name
            cmdPayee.ControlTipText = .Name
            lblNo = Bsprintf("����� %d:", .DocNo)
            lblDate = Bsprintf("���� %n:", .DocDate)
            cboQueue = CStr(.Queue)
            txtSum = PlatFormat(.Sum)
            txtDetails = .Details
            cmdSS.Caption = .SS
        End If
        n = .RowsCount + 1
        If n < 2 Then n = 2
        sbrRows.Max = n
        If .Row > n Then .Row = n
        sbrRows.Value = .Row
    End With
    txtDate = PlatDate(Date)
    txtDate.ControlTipText = "������� " & PlatDate(Date)
    txtNo = CStr(User.No)
    With cboTax
        .AddItem "���"
        .AddItem "18%"
        .AddItem "20%"
        .AddItem "10%"
    End With
    With cboQueue
        .AddItem "6"
        .AddItem "5"
        .AddItem "4"
        .AddItem "3"
        .AddItem "2"
        .AddItem "1"
    End With
    mLoading = False
End Sub

