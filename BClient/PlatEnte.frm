VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PlatEnter 
   Caption         =   "Ввод платежного поручения"
   ClientHeight    =   5535
   ClientLeft      =   30
   ClientTop       =   270
   ClientWidth     =   6735
   OleObjectBlob   =   "PlatEnter.frx":0000
   StartUpPosition =   2  'CenterScreen
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

Const multiPayees = 0
Const multiArchive = 1
Const multiPayment = 2
Const multiDesc = 3

Private Sub cboCustom1_Change()
    If cboCustom1 = EmptyInList Then cboCustom1 = ""
    With lstPayees
        .List = Payment.List(txtName, cboCustom1, cboCustom2)
        lblPayees = .ListCount - 1
    End With
End Sub

Private Sub cboCustom2_Change()
    If cboCustom2 = EmptyInList Then cboCustom2 = ""
    With lstPayees
        .List = Payment.List(txtName, cboCustom1, cboCustom2)
        lblPayees = .ListCount - 1
    End With
End Sub

Private Sub cboSum_Change()
    Dim n As Boolean, s As Currency
    If cboSum = EmptyInList Then cboSum = ""
    n = cboSum.TextLength > 0
    lblBack.Enabled = n
    lblClear.Enabled = n
    With cboSum
        s = RVal(.Text)
        lblSumText.Caption = Format(s, "#,##0.00") & vbCrLf & RSumStr(s, vbCrLf)
    End With
End Sub

Private Sub AddNew()
    Hide
    Load NewName
    With NewName
        .AddNewOk = False
        .Show
        If .AddNewOk Then ResetLists
    End With
    Unload NewName
    Show
End Sub

Private Sub chkThisPayee_Change()
    ArchiveChange
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    Dim arr As Variant
    Dim PayeeNick As String
    With lstPayees
        If .ListIndex < 0 Then
            MultiPage.Value = multiPayees
            MsgBox "Не выбран получатель", vbExclamation, AppTitle
            .SetFocus
            Exit Sub
        Else
            PayeeNick = .List(.ListIndex)
        End If
    End With
    If RVal(cboSum) = 0 Then
        MsgBox "Не введена сумма платежа", vbExclamation, AppTitle
        MultiPage.Value = multiPayment
        cboSum.SetFocus
        Exit Sub
    End If
    With txtDescription
        .Text = strspaces1(.Text)
        If .TextLength = 0 Then
            MsgBox "Не введено назначение платежа", vbExclamation, AppTitle
            MultiPage.Value = multiDesc
            .SetFocus
            Exit Sub
        End If
    End With
    Hide
    Archive.Add PayeeNick, spnNo, txtDate, spnAction, spnQueue, cboSum, lstNDS.Text, txtDescription
    Unload Me
End Sub

Private Sub cmdToday_Click()
    txtDate = CStr(Date)
End Sub

Private Sub lbl0_Click()
    NumPadKey "0"
End Sub

Private Sub lbl00_Click()
    NumPadKey "0"
    NumPadKey "0"
End Sub

Private Sub lbl000_Click()
    NumPadKey "0"
    NumPadKey "0"
    NumPadKey "0"
End Sub

Private Sub lbl1_Click()
    NumPadKey "1"
End Sub

Private Sub lbl2_Click()
    NumPadKey "2"
End Sub

Private Sub lbl3_Click()
    NumPadKey "3"
End Sub

Private Sub lbl4_Click()
    NumPadKey "4"
End Sub

Private Sub lbl5_Click()
    NumPadKey "5"
End Sub

Private Sub lbl6_Click()
    NumPadKey "6"
End Sub

Private Sub lbl7_Click()
    NumPadKey "7"
End Sub

Private Sub lbl8_Click()
    NumPadKey "8"
End Sub

Private Sub lbl9_Click()
    NumPadKey "9"
End Sub

Private Sub lblBack_Click()
    NumPadKey "B"
End Sub

Private Sub lblClear_Click()
    NumPadKey "C"
End Sub

Private Sub lblPoint_Click()
    NumPadKey "."
End Sub

Private Sub lstArchive_Change()
    With lstArchive
        If .ListIndex > -1 Then
            txtDescription.Text = .List(.ListIndex)
        End If
    End With
End Sub

Private Sub lstArchive_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    MultiPage.Value = MultiPage.Value + 1
End Sub

Private Sub lstInserts_Change()
    TextInsert lstInserts.Text
End Sub

Private Sub lstNDS_Click()
    Dim s As String, v1 As Currency, v2 As Double
    v1 = RVal(cboSum.Text)
    If v1 = 0 Then
        MsgBox "Не указана сумма для вычисления НДС", vbExclamation, AppTitle
        Exit Sub
    End If
    v2 = CDbl(RVal(lstNDS.Text))
    If v2 = 0 Then
        TextInsert "НДС не облагается."
        Exit Sub
    End If
    s = IIf(chkIncludingNDS, "Включая НДС ", "") & lstNDS.Text
    v1 = Sum2Tax(v1, v2)
    If v1 > 0 Then
        TextInsert s & ": " & PlatFormat(v1)
    End If
End Sub

Private Sub lstPayees_Change()
    Payment.Nick = lstPayees.Text
End Sub

Private Sub lstPayees_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    With lstPayees
        If .List(.ListIndex) = AddNewInList Then
            AddNew
        Else
            MultiPage.Value = MultiPage.Value + 1
        End If
    End With
End Sub

Private Sub MultiPage_Change()
    Dim OldValue
    With MultiPage
        OldValue = Val(.Tag)
        'Select Case OldValue
        'End Select
        Select Case .Value
            Case 1
                chkThisPayee.Enabled = Len(lstPayees.Text) > 0
                ArchiveChange
            Case Else
        End Select
        .Tag = CStr(.Value)
    End With
End Sub

Private Sub spnAction_Change()
    txtDate_Change
End Sub

Private Sub spnDate_Change()
    With txtDate
        If IsDate(.Text) Then
            .Text = DateAdd("d", spnDate.Value, CDate(.Text))
        End If
    End With
    spnDate.Value = 0
End Sub

Private Sub spnNo_Change()
    txtNo = spnNo
End Sub

Private Sub spnQueue_Change()
    txtQueue = spnQueue
End Sub

Private Sub txtArchive_Change()
    ArchiveChange
End Sub

Private Sub txtDate_Change()
    Dim DocDate As Date
    With txtDate
        If IsDate(.Text) Then
            DocDate = CDate(.Text)
            cmdToday.Enabled = DocDate <> Date
            txtAction.Text = BPrintF("%y (+%d)", _
                DateAdd("d", spnAction.Value, DocDate), spnAction.Value)
        Else
            cmdToday.Enabled = True
        End If
    End With
End Sub

Private Sub txtDescription_Change()
    lblDesc = txtDescription.TextLength
End Sub

Private Sub txtDescription_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    With txtDescription
        .Text = StrTran(.Text, vbCrLf, " ")
        .Text = StrTran(.Text, "  ", " ")
    End With
End Sub

Private Sub txtName_Change()
    With lstPayees
        .List = Payment.List(txtName, cboCustom1, cboCustom2)
        lblPayees = .ListCount - 1
    End With
End Sub

Private Sub txtNo_Change()
    Dim v
    v = Val(txtNo)
    With spnNo
        If v < .Min Or v > .Max Then
            v = .Min
            txtNo = v
        End If
        .Value = v
    End With
End Sub

Private Sub txtQueue_Change()
    Dim v
    v = Val(txtQueue)
    With spnQueue
        If v < .Min Or v > .Max Then
            v = .Min
            txtQueue = v
        End If
        .Value = v
    End With
End Sub

Private Sub UserForm_Initialize()
    Dim Inserts As New CInserts
    MultiPage.Value = multiPayees
    ResetLists
    lstInserts.List = Inserts.List(1)
    lstNDS.List = Inserts.List(2)
    cmdToday_Click
End Sub

Private Sub TextInsert(s As String)
    Dim i
    With txtDescription
        i = .SelStart
        If i > 0 Then
            If Mid(.Text, i, 1) <> " " Then s = " " & s
        End If
        i = i + .SelLength
        If i < .TextLength Then
            If Mid(.Text, i + 1, 1) <> " " Then s = s & " "
        End If
        .SelText = s
        .SetFocus
    End With
End Sub

Private Sub ResetLists()
    With lstPayees
        .List = Payment.List(txtName, cboCustom1, cboCustom2)
        lblPayees = .ListCount - 1
    End With
    cboCustom1.List = Payment.CustomList(1)
    cboCustom2.List = Payment.CustomList(2)
    With lstArchive
        .List = Archive.List(lstPayees.Text)
        .RemoveItem 0
        lblArchive = .ListCount
    End With
    cboSum.List = Archive.ListSum10
End Sub

Private Sub NumPadKey(Key As String)
    Dim i, p
    With cboSum
        i = .SelStart
        p = InStr(1, .Text, "-")
        Select Case Key
            Case "C"
                .Text = ""
            Case "B"
                If i > 0 Then
                    .SelStart = i - 1
                    .SelLength = 1
                    .SelText = ""
                End If
            Case "."
                If p = 0 Then
                    If .TextLength = 0 Then
                        .SelText = "0-"
                    Else
                        .SelText = "-"
                    End If
                End If
            Case Else
                If p = 0 Then
                    .SelText = Key
                ElseIf i < p + 2 Then
                    .SelText = Key
                End If
        End Select
        .SetFocus
    End With
End Sub

Private Sub ArchiveChange()
    With lstArchive
        If chkThisPayee Then
            .List = Archive.List(lstPayees.Text, txtArchive)
        Else
            .List = Archive.List("", txtArchive)
        End If
        .RemoveItem 0
        lblArchive = .ListCount
    End With
End Sub
