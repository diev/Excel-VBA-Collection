VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MailBox 
   Caption         =   "Отправка и прием"
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5505
   HelpContextID   =   420000
   OleObjectBlob   =   "MailBox.frx":0000
   StartUpPosition =   1  'CenterOwner
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
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

Private StandardHeight As Long

Public Mode As String 'R/S/A

Private Sub chkAll_Click()
    RefreshBoxes
End Sub

Private Sub chkHeight_Click()
    Dim n As Long
    If chkHeight Then
        n = Application.Height - Height
        lstFiles.Height = lstFiles.Height + n
        chkAll.Top = chkAll.Top + n
        chkHeight.Top = chkHeight.Top + n
        cmdDial.Top = cmdDial.Top + n
        cmdCancel.Top = cmdCancel.Top + n
        Height = Application.Height
        Top = 0
    Else
        Height = StandardHeight
        n = Application.Height - Height
        lstFiles.Height = lstFiles.Height - n
        chkAll.Top = chkAll.Top - n
        chkHeight.Top = chkHeight.Top - n
        cmdDial.Top = cmdDial.Top - n
        cmdCancel.Top = cmdCancel.Top - n
        Top = Abs(Application.Height - Height) \ 2
    End If
End Sub

Private Sub chkInet_Click()
    If chkInet Then
        chkInet.ControlTipText = "Соединение по TCP/IP с stp://194.8.179.114:40696"
        SMail.DialOrInet = 2
    Else
        chkInet.ControlTipText = "Соединение по телефонам 3240691 и 3240696"
        SMail.DialOrInet = 1
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDecode_Click()
    Dim File As String, File2 As String
    On Error Resume Next
    Hide
    File2 = GetParenthesed(lstFiles)
    Select Case Mode
        Case "R": File = SMail.Recv & File2
        Case "S": File = SMail.Send & File2
        Case "A": File = SMail.Archive & File2
    End Select
    File2 = FilePath(User.ExportList) & File2
    If BrowseForSave(File2, "Файлы клиента (*." & User.ID & "),*." & User.ID, _
        "файл для создания копии") Then
        User.ExportList = FilePath(File2)
        If IsFile(File2) Then
            If Not YesNoBox("Файл %s уже существует.\n\nПерезаписать?", File2) Then
                Exit Sub
            End If
        End If
        If LCase(FileExt(File)) = User.ID Then
            If Not PGP.Decode(File, File2) Then
                If YesNoBox("Программа PGP не смогла расшифровать файл!\n" & _
                    "Возможно, ошибка использования ключей.\n\nПросто скопировать?") Then
                    FileCopy File, File2
                End If
            End If
        Else
            FileCopy File, File2
        End If
    End If
End Sub

Private Sub cmdDial_Click()
    With SMail
        If .Valid Then
            Hide
            .Dial
            Unload Me
        Else
            WarnBox .None
        End If
    End With
End Sub

Private Sub cmdFiles_Click()
    Hide
    Select Case Mode
        Case "R": MailIn
        Case "S": MailOut
        Case "A": MailArch
    End Select
    'RefreshBoxes
    Unload Me
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
        If YesNoBox("Удалить файл?\n%s", File) Then
            Kill File
            If IsFile(File) Then
                WarnBox "Файл удалить не удалось!"
            'Else
            '    .RemoveItem .ListIndex
            End If
            RefreshBoxes
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
    MailOpenFile File
End Sub

Private Sub cmdRep_Click()
    SMail.LastRep
End Sub

Private Sub lstFiles_Change()
    With lstFiles
        cmdOpen.Enabled = .ListIndex > -1
        cmdKill.Enabled = .ListIndex > -1
        cmdDecode.Enabled = .ListIndex > -1
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
    Dim File As String
    On Error Resume Next
    Caption = Bsprintf("%s клиента %s", Caption, User.ID)
    DefCaption1 = MailBoxes.Tabs(0).Caption
    DefCaption2 = MailBoxes.Tabs(1).Caption
    DefCaption3 = MailBoxes.Tabs(2).Caption
    Mode = "S"
    MailBoxes.Tabs(0).ControlTipText = SMail.Send
    MailBoxes.Tabs(1).ControlTipText = SMail.Recv
    MailBoxes.Tabs(2).ControlTipText = SMail.Archive
    'cmdDial.Enabled = SMail.Valid
    chkInet = SMail.DialOrInet = 2
    RefreshBoxes
    
    File = Dir(SMail.Recv & "!*.txt")
    Do While File <> vbNullString
        UrgentMessage SMail.Recv & File
        File = Dir
    Loop
    StandardHeight = Height
End Sub

Private Sub RefreshBoxes()
    Dim File As String, myID As String, Ext As String
    myID = IIf(chkAll, "*", User.ID)
    cmdRep.Enabled = IsFile(SMail.Recv & "rep" & myID & ".txt")
    Select Case Mode
        Case "R":
            With lstFiles
                .Clear
                
                File = Dir(SMail.Recv & "*.exe")
                Do While File <> vbNullString
                    .AddItem "Обновление " & UpdateName(File) & FileDT(SMail.Recv & File)
                    File = Dir
                Loop
                
                File = Dir(SMail.Recv & "inv." & myID)
                Do While File <> vbNullString
                    .AddItem "Входящие " & VypDate(File) & FileDT(SMail.Recv & File)
                    File = Dir
                Loop
                
                File = Dir(SMail.Recv & "vyp." & myID)
                Do While File <> vbNullString
                    .AddItem "Исходящие " & VypDate(File) & FileDT(SMail.Recv & File)
                    File = Dir
                Loop
                
                'File = Dir(SMail.Recv & "wyp??*." & myID)
                'Do While File <> vbNullString
                '    .AddItem "Выписка вал. " & VypDate(File) & FileDT(SMail.Recv & File)
                '    File = Dir
                'Loop
                
                File = Dir(SMail.Recv & "vyp??-??." & myID)
                Do While File <> vbNullString
                    .AddItem "Выписка " & VypDate(File) & FileDT(SMail.Recv & File)
                    File = Dir
                Loop
                
                File = Dir(SMail.Recv & "vyp??r*." & myID)
                Do While File <> vbNullString
                    .AddItem "Реестр " & VypDate(File) & FileDT(SMail.Recv & File)
                    File = Dir
                Loop
                
                File = Dir(SMail.Recv & "!*.txt")
                Do While File <> vbNullString
                    .AddItem Bsprintf("Извещение (%s)", _
                        LCase(File)) & FileDT(SMail.Recv & File)
                    File = Dir
                Loop
                File = Dir(SMail.Recv & "rep*.txt")
                Do While File <> vbNullString
                    .AddItem Bsprintf("Итоги (%s)", _
                        LCase(File)) & FileDT(SMail.Recv & File)
                    File = Dir
                Loop
                File = Dir(SMail.Recv & "*.txt")
                Do While File <> vbNullString
                    If Left(File, 1) <> "!" And _
                        LCase(Left(File, 3)) <> "rep" _
                            Then .AddItem Bsprintf("Сообщение (%s)", _
                                LCase(File)) & FileDT(SMail.Recv & File)
                    File = Dir
                Loop
                File = Dir(SMail.Recv & "*.doc")
                Do While File <> vbNullString
                    .AddItem Bsprintf("Файл MS Word (%s)", _
                        LCase(File)) & FileDT(SMail.Recv & File)
                    File = Dir
                Loop
                
                File = Dir(SMail.Recv & "o???????." & myID)
                Do While File <> vbNullString
                    .AddItem "Принятый " & DocName(File) & FileDT(SMail.Recv & File)
                    File = Dir
                Loop
                
                File = Dir(SMail.Recv & "e???????." & myID)
                Do While File <> vbNullString
                    .AddItem "Ошибочный " & DocName(File) & FileDT(SMail.Recv & File)
                    File = Dir
                Loop
                
                File = Dir(SMail.Recv & "t???????." & myID)
                Do While File <> vbNullString
                    .AddItem "Тестовый " & DocName(File) & FileDT(SMail.Recv & File)
                    File = Dir
                Loop
                
                File = Dir(SMail.Recv & "remart.pg?")
                Do While File <> vbNullString
                    .AddItem Bsprintf("Курс валют ЦБ (%s)", LCase(File)) & FileDT(SMail.Recv & File)
                    File = Dir
                Loop
                
                chkAll.Caption = Bsprintf("Все (%d/%d)", _
                    .ListCount, CountFiles(SMail.Recv & "*.*"))
            End With
            
        Case "S":
            With lstFiles
                .Clear
                
                File = Dir(SMail.Send & "*." & myID)
                Do While File <> vbNullString
                    .AddItem DocName(File) & FileDT(SMail.Send & File)
                    File = Dir
                Loop
                
                chkAll.Caption = Bsprintf("Все (%d/%d)", _
                    .ListCount, CountFiles(SMail.Send & "*.*"))
            End With
            
        Case "A":
            With lstFiles
                .Clear
                
                File = Dir(SMail.Archive & "o???????." & myID)
                Do While File <> vbNullString
                    .AddItem "Принятый " & DocName(File) & FileDT(SMail.Archive & File)
                    File = Dir
                Loop
                
                File = Dir(SMail.Archive & "e???????." & myID)
                Do While File <> vbNullString
                    .AddItem "Ошибочный " & DocName(File) & FileDT(SMail.Archive & File)
                    File = Dir
                Loop
                
                File = Dir(SMail.Archive & "t???????." & myID)
                Do While File <> vbNullString
                    .AddItem "Тестовый " & DocName(File) & FileDT(SMail.Archive & File)
                    File = Dir
                Loop
                
                chkAll.Caption = Bsprintf("Все (%d/%d)", _
                    .ListCount, CountFiles(SMail.Archive & "*.*"))
            End With
    End Select
    
    With lstFiles
        cmdOpen.Enabled = .ListIndex > -1
        cmdKill.Enabled = .ListIndex > -1
        cmdDecode.Enabled = .ListIndex > -1
    End With
End Sub

Public Function VypDate(File As String) As String
    'Dim mm As Long, dd As Long
    ''vypMM[-DD].kkk
    'mm = Val(Mid(File, 4, 2))
    'If InStr(File, "-") Then
    '    dd = Val(Mid(File, 7, 2))
    '    VypDate = Bsprintf("%d.%02d (%s)", dd, mm, LCase(File))
    'Else
    '    VypDate = Bsprintf("%d месяц (%s)", mm, LCase(File))
    'End If
    VypDate = Bsprintf("(%s)", LCase(File))
End Function

Public Function DocName(File As String) As String
    Dim m As Long, dd As Long, nnn As Long
    'aYMDDNNN.kkk
    nnn = Val(Mid(File, 6, 3))
    If nnn > 0 Then
        m = Ot36(Mid(File, 3, 1))
        dd = Val(Mid(File, 4, 2))
        DocName = Bsprintf("N %d от %s.%02d (%s)", nnn, dd, m, LCase(File))
    Else
        DocName = Bsprintf("(%s)", LCase(File))
    End If
End Function

Public Function UpdateName(File As String) As String
    UpdateName = Bsprintf("(%s)", LCase(File))
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

Public Sub UrgentMessage(File As String)
    On Error Resume Next
    If Not YesNoBox("ВАЖНОЕ СООБЩЕНИЕ ИЗ БАНКА:\n\n%s\n\nНапоминать еще?", _
        CWin(InputFile(File))) Then
        Kill File
    End If
End Sub

Public Function FileDT(File As String) As String
    Dim dt As Date
    dt = FileDateTime(File)
    If DateDiff("d", dt, Date) = 0 Then
        FileDT = Format(dt, " + dd.MM, HH:mm")
    Else
        FileDT = Format(dt, " - dd.MM, HH:mm")
    End If
End Function
