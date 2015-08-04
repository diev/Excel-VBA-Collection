VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MailBox 
   Caption         =   "Отправка и прием"
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5505
   OleObjectBlob   =   "MailBox.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MailBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Explicit
Option Compare Text
Option Base 1
DefLng A-Z

Private DefCaption1 As String
Private DefCaption2 As String
Private DefCaption3 As String

Public Mode As String 'R/S/A

Private Sub chkAll_Click()
    RefreshBoxes
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDial_Click()
    With SMail
        If .Valid Then
            Hide
            .Dial
            Unload Me
        Else
            MsgBox .None, vbExclamation, App.Title
        End If
    End With
End Sub

Private Sub cmdFiles_Click()
    Select Case Mode
        Case "R": MailIn
        Case "S": MailOut
        Case "A": MailArch
    End Select
    RefreshBoxes
End Sub

Private Sub cmdKill_Click()
    Dim File As String
    On Error Resume Next
    With lstFiles
        Select Case Mode
            Case "R": File = SMail.Recv & GetParenthesed(.Text)
            Case "S": File = SMail.Send & GetParenthesed(.Text)
            Case "A": File = SMail.Archive & GetParenthesed(.Text)
        End Select
        If MsgBox("Удалить файл?", vbYesNo + vbQuestion, App.Title) = vbYes Then
            Kill File
            If IsFile(File) Then
                MsgBox "Файл удалить не удалось!", vbExclamation, App.Title
            'Else
            '    .RemoveItem .ListIndex
            End If
            RefreshBoxes
            cmdKill.Enabled = .ListCount > 0
            cmdOpen.Enabled = .ListCount > 0
        End If
    End With
End Sub

Private Sub cmdOpen_Click()
    Dim File As String
    Hide
    Select Case Mode
        Case "R": File = SMail.Recv & GetParenthesed(lstFiles)
        Case "S": File = SMail.Send & GetParenthesed(lstFiles)
        Case "A": File = SMail.Archive & GetParenthesed(lstFiles)
    End Select
    If FileExt(File) = User.ID Then
        File = PGP.DecodeEx(File)
        If Not IsFile(File) Then
            MsgBox BPrintF("Программа PGP не смогла расшифровать файл!\nВозможно, ошибка использования ключей"), _
                vbExclamation, App.Title
            Exit Sub
        End If
    End If
    MailDump File
End Sub

Private Sub cmdRep_Click()
    SMail.LastRep
End Sub

Private Sub lstFiles_Change()
    With lstFiles
        cmdKill.Enabled = .ListIndex > -1
        cmdOpen.Enabled = .ListIndex > -1
    End With
End Sub

Private Sub lstFiles_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    cmdOpen_Click
End Sub

Private Sub MailBoxes_Change()
    Select Case MailBoxes.Value
        Case 1: Mode = "R"
        Case 0: Mode = "S"
        Case 2: Mode = "A"
    End Select
    RefreshBoxes
End Sub

Private Sub UserForm_Initialize()
    DefCaption1 = MailBoxes.Tabs(0).Caption
    DefCaption2 = MailBoxes.Tabs(1).Caption
    DefCaption3 = MailBoxes.Tabs(2).Caption
    Mode = "S"
    MailBoxes.Tabs(0).ControlTipText = SMail.Send
    MailBoxes.Tabs(1).ControlTipText = SMail.Recv
    MailBoxes.Tabs(2).ControlTipText = SMail.Archive
    'cmdDial.Enabled = SMail.Valid
    RefreshBoxes
End Sub

Private Sub RefreshBoxes()
    Dim File As String, Ext As String
    Ext = IIf(chkAll, "*", User.ID)
    Select Case Mode
        Case "R":
            With lstFiles
                .Clear
                File = Dir(SMail.Recv & "vyp??-??." & Ext)
                Do While File <> vbNullString
                    .AddItem "Выписка руб. за " & VypDate(File)
                    File = Dir
                Loop
                File = Dir(SMail.Recv & "wyp??-??." & Ext)
                Do While File <> vbNullString
                    .AddItem "Выписка вал. за " & VypDate(File)
                    File = Dir
                Loop
                File = Dir(SMail.Recv & "inv??-??." & Ext)
                Do While File <> vbNullString
                    .AddItem "Предв. инф. на " & VypDate(File)
                    File = Dir
                Loop
                File = Dir(SMail.Recv & "*.txt")
                Do While File <> vbNullString
                    .AddItem "Текст. файл (" & LCase(File) & ")"
                    File = Dir
                Loop
                File = Dir(SMail.Recv & "o???????." & Ext)
                Do While File <> vbNullString
                    .AddItem "Принятый док. " & DocName(File)
                    File = Dir
                Loop
                File = Dir(SMail.Recv & "e???????." & Ext)
                Do While File <> vbNullString
                    .AddItem "Ошибочный док. " & DocName(File)
                    File = Dir
                Loop
                File = Dir(SMail.Recv & "remart.pg?")
                Do While File <> vbNullString
                    .AddItem "Курс валют ЦБ на " & Ot36(Right(File, 1)) & " (" & LCase(File) & ")"
                    File = Dir
                Loop
                chkAll.Caption = BPrintF("Показать всех (%d/%d)", _
                    .ListCount, CountFiles(SMail.Recv & "*.*"))
            End With
        Case "S":
            With lstFiles
                .Clear
                File = Dir(SMail.Send & "*." & Ext)
                Do While File <> vbNullString
                    .AddItem "Плат. док. " & DocName(File)
                    File = Dir
                Loop
                chkAll.Caption = BPrintF("Показать всех (%d/%d)", _
                    .ListCount, CountFiles(SMail.Send & "*.*"))
            End With
        Case "A":
            With lstFiles
                .Clear
                File = Dir(SMail.Archive & "o???????." & Ext)
                Do While File <> vbNullString
                    .AddItem "Принятый док. " & DocName(File)
                    File = Dir
                Loop
                File = Dir(SMail.Archive & "e???????." & Ext)
                Do While File <> vbNullString
                    .AddItem "Ошибочный док. " & DocName(File)
                    File = Dir
                Loop
                chkAll.Caption = BPrintF("Показать всех (%d/%d)", _
                    .ListCount, CountFiles(SMail.Archive & "*.*"))
            End With
    End Select
End Sub

Public Function VypDate(File As String) As String
    Dim mm As Long, dd As Long
    'vypMM-DD.kkk
    mm = Val(Mid(File, 4, 2))
    If InStr(File, "-") Then
        dd = Val(Mid(File, 7, 2))
        VypDate = BPrintF("%d.%02d (%s)", dd, mm, LCase(File))
    Else
        VypDate = BPrintF("%d месяц (%s)", mm, LCase(File))
    End If
End Function

Public Function DocName(File As String) As String
    Dim m As Long, dd As Long, nnn As Long
    'aYMDDNNN.kkk
    m = Ot36(Mid(File, 3, 1))
    dd = Val(Mid(File, 4, 2))
    nnn = Val(Mid(File, 6, 3))
    DocName = BPrintF("№%d от %s.%02d (%s)", nnn, dd, m, LCase(File))
End Function

Public Function CountFiles(Mask As String) As Long
    Dim File As String
    File = Dir(Mask)
    CountFiles = 0
    Do While File <> vbNullString
        CountFiles = CountFiles + 1
        File = Dir
    Loop
End Function
