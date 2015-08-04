VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MailBox 
   Caption         =   "Отправка и прием"
   ClientHeight    =   6720
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6000
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

Const IDExt = ".ID"

Private DefCaption1 As String
Private DefCaption2 As String
Private DefCaption3 As String

Private StandardHeight As Long

Public Mode As String 'R/S/A

Private Sub cboMulti_Change()
    lstFiles.MultiSelect = Int(Right(cboMulti, 1))
End Sub

Private Sub chkAll_Click()
    RefreshBoxes
End Sub

Private Sub chkInet_Click()
    If chkInet Then
        chkInet.ControlTipText = "Соединение по TCP/IP"
        SMail.DialOrInet = 2
    Else
        chkInet.ControlTipText = "Соединение по телефону"
        SMail.DialOrInet = 1
    End If
End Sub

Private Sub chkToday_Click()
    RefreshBoxes
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdClean_Click()
    Dim p As String, s As String, n As Long, dt As Date, File As String, k As Long
    On Error Resume Next
    Select Case Mode
        Case "S": p = SMail.Send
        Case "R": p = SMail.Recv
        Case "A": p = SMail.Archive
    End Select
    dt = Now: n = 0
    File = Dir(p & "*.*")
    Do While File <> vbNullString
        n = n + 1
        If dt > FileDateTime(p & File) Then
            dt = FileDateTime(p & File)
        End If
        File = Dir
    Loop
    s = Bsprintf("В директории %s файлов: %d,\nсамый старый - %n (дней: %d).\n\n" & _
        "Укажите число дней, сколько следует оставить\n" & _
        "(например, 60 = 2 месяца, 0 = оставить как есть)", _
        p, n, dt, DateDiff("d", dt, Now))
    s = InputBox(s, "Очистка старого (рекомендуется)", 0)
    If s = vbNullString Then Exit Sub
    n = Int(s)
    If n > 0 Then
        dt = DateAdd("d", -n, Now)
        If Not YesNoBox("Удалить файлы старее %n\n(%d дней назад)?", dt, n) Then Exit Sub
        n = 0: n = 0
        File = Dir(p & "*.*")
        Do While File <> vbNullString
            If FileDateTime(p & File) < dt Then
                Kill p & File
                n = n + 1
            Else
                k = k + 1
            End If
            File = Dir
        Loop
        RefreshBoxes
        InfoBox "Удалено файлов: %d, оставлено: %d.\n", n, k
    End If
End Sub

Private Sub cmdDecode_Click()
    Dim p As String, i As Long
    On Error Resume Next
    Hide
    Select Case Mode
        Case "R": p = SMail.Recv
        Case "S": p = SMail.Send
        Case "A": p = SMail.Archive
    End Select
    With lstFiles
        If .MultiSelect = fmMultiSelectSingle Then
            DecodeFile p, GetParenthesed(.Value)
        Else
            For i = 0 To .ListCount - 1
                If .Selected(i) Then
                    DecodeFile p, GetParenthesed(.List(i))
                End If
            Next
        End If
    End With
End Sub

Private Sub DecodeFile(p As String, File As String)
    Dim File2 As String, Path2 As String
    On Error Resume Next
    If User.IsID4(FileExt(File)) Then
        If InStr(1, File, "v1c", vbTextCompare) = 1 Then
            File2 = FilePath(User.ExportList) & "kl_to_1c.txt"
            If BrowseForSave(File2, "Файлы 1C (*.txt),*.txt", _
                "файл для загрузки в 1C") Then
                Path2 = FilePath(File2)
                If InStr(1, Path2, SMail.Send, vbTextCompare) = 1 Then
                    StopBox "Не следует мусорить в каталоге отправки!"
                    Exit Sub
                End If
                User.ExportList = Path2
                If IsFile(File2) Then
                    If Not YesNoBox("Файл %s уже существует.\n\nПерезаписать?", File2) Then
                        Exit Sub
                    End If
                End If
                If Not Crypto.Decrypt(p & File, File2) Then
                    If YesNoBox("Программа PGP не смогла расшифровать файл!\n" & _
                        "Возможно, ошибка использования ключей.\n\nПросто скопировать?") Then
                        FileCopy p & File, File2
                    End If
                End If
                'ExportTo1CFile File2
            End If
        Else
            File2 = FilePath(User.ExportList) & File & ".txt"
            If BrowseForSave(File2, "Текстовые файлы (*.txt),*.txt", _
                "файл для создания расшифрованной копии") Then
                Path2 = FilePath(File2)
                If InStr(1, Path2, SMail.Send, vbTextCompare) = 1 Then
                    StopBox "Не следует мусорить в каталоге отправки!"
                    Exit Sub
                End If
                User.ExportList = Path2
                If IsFile(File2) Then
                    If Not YesNoBox("Файл %s уже существует.\n\nПерезаписать?", File2) Then
                        Exit Sub
                    End If
                End If
                If Not Crypto.Decrypt(p & File, File2) Then
                    If YesNoBox("Программа PGP не смогла расшифровать файл!\n" & _
                        "Возможно, ошибка использования ключей.\n\nПросто скопировать?") Then
                        FileCopy p & File, File2
                    End If
                End If
            End If
        End If
    Else
        File2 = FilePath(User.ExportList) & File
        If BrowseForSave(File2, , _
            "файл для создания копии") Then
            Path2 = FilePath(File2)
            If InStr(1, Path2, SMail.Send, vbTextCompare) = 1 Then
                StopBox "Не следует мусорить в каталоге отправки!"
                Exit Sub
            End If
            User.ExportList = Path2
            If IsFile(File2) Then
                If Not YesNoBox("Файл %s уже существует.\n\nПерезаписать?", File2) Then
                    Exit Sub
                End If
            End If
            FileCopy p & File, File2
        End If
    End If
End Sub

Private Sub cmdDial_Click()
    Hide
    SMail.Dial
    Unload Me
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

Private Sub cmdFolder_Click()
    Dim p As String
    On Error Resume Next
    Hide
    Select Case Mode
        Case "R": p = SMail.Recv
        Case "S": p = SMail.Send
        Case "A": p = SMail.Archive
    End Select
    Shell QFile(GetWinDir & "explorer.exe") & " " & _
        QFile(p), vbNormalFocus
End Sub

Private Sub cmdKill_Click()
    Dim p As String, i As Long
    If Not YesNoBox("Удалить файл(ы)?") Then
        Exit Sub
    End If
    On Error Resume Next
    'Hide
    Select Case Mode
        Case "R": p = SMail.Recv
        Case "S": p = SMail.Send
        Case "A": p = SMail.Archive
    End Select
    With lstFiles
        If .MultiSelect = fmMultiSelectSingle Then
            Kill p & GetParenthesed(.Value)
        Else
            For i = 0 To .ListCount - 1
                If .Selected(i) Then
                    Kill p & GetParenthesed(.List(i))
                End If
            Next
        End If
    End With
    RefreshBoxes
End Sub

Private Sub cmdOpen_Click()
    Dim p As String, i As Long
    On Error Resume Next
    Hide
    Select Case Mode
        Case "R": p = SMail.Recv
        Case "S": p = SMail.Send
        Case "A": p = SMail.Archive
    End Select
    With lstFiles
        If .MultiSelect = fmMultiSelectSingle Then
            MailOpenFile p & GetParenthesed(.Value)
        Else
            For i = 0 To .ListCount - 1
                If .Selected(i) Then
                    MailOpenFile p & GetParenthesed(.List(i))
                End If
            Next
        End If
    End With
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
    Dim n As Long
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
    chkAll = App.BoolOptions("ViewAll")
    chkToday = App.BoolOptions("ViewToday")
    chkToday.ControlTipText = Bsprintf("Только %n и %n!", DateAdd("d", -1, Now), Now)
    lstFiles.MultiSelect = App.Options("MultiSelect")
    cboMulti.AddItem "Мульти 0"
    cboMulti.AddItem "Мульти 1"
    cboMulti.AddItem "Мульти 2"
    cboMulti.Text = "Мульти " & CStr(lstFiles.MultiSelect)
    StandardHeight = Height
    With Application
        Move Left, .Top + 5, width, .Height - 10
    End With
    n = Height - StandardHeight
    lstFiles.Height = lstFiles.Height + n
    lblFilesLen.Top = lblFilesLen.Top + n
    chkInet.Top = chkInet.Top + n
    cmdDial.Top = cmdDial.Top + n
    cmdCancel.Top = cmdCancel.Top + n
    'CheckRecv
    RefreshBoxes
End Sub

Private Sub RefreshBoxes()
    Dim File As String, myID As String, Ext As String, dt As Date
    myID = IIf(chkAll, "*", User.ID4)
    dt = IIf(chkToday, DateAdd("d", -1, Now), DateAdd("yyyy", -100, Now))
    cmdRep.Enabled = IsFile(SMail.Recv & "rep" & myID & ".txt")
    Select Case Mode
        Case "R":
            With lstFiles
                .Clear
                
                File = Dir(SMail.Recv & "*.exe")
                Do While File <> vbNullString
                    If LCase(Left(File, 3)) <> "ok-" Then
                        .AddItem "Обновление новое " & UpdateName(File) & FileDT(SMail.Recv & File)
                    End If
                    File = Dir
                Loop
                
                File = Dir(SMail.Recv & "ok-*.exe")
                Do While File <> vbNullString
                    .AddItem "Обновление вып. " & UpdateName(File) & FileDT(SMail.Recv & File)
                    File = Dir
                Loop
                
                File = Dir(SMail.Recv & "inv." & myID)
                Do While File <> vbNullString
                    .AddItem "Входящие " & VypDate(File) & FileDT(SMail.Recv & File)
                    File = Dir
                Loop
                
                File = Dir(SMail.Recv & "vyp??-??." & myID)
                Do While File <> vbNullString
                    If FileDateTime(SMail.Recv & File) >= dt Then
                        .AddItem "Выписка " & VypDate(File) & FileDT(SMail.Recv & File)
                    End If
                    File = Dir
                Loop
                
                File = Dir(SMail.Recv & "v1c??-??." & myID)
                Do While File <> vbNullString
                    If FileDateTime(SMail.Recv & File) >= dt Then
                        .AddItem "Выписка 1C " & VypDate(File) & FileDT(SMail.Recv & File)
                    End If
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
                    If FileDateTime(SMail.Recv & File) >= dt Then
                        .AddItem Bsprintf("Итоги (%s)", _
                            LCase(File)) & FileDT(SMail.Recv & File)
                    End If
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
                File = Dir(SMail.Recv & "*.htm")
                Do While File <> vbNullString
                    .AddItem Bsprintf("Сообщение (%s)", _
                        LCase(File)) & FileDT(SMail.Recv & File)
                    File = Dir
                Loop
                File = Dir(SMail.Recv & "*.gif")
                Do While File <> vbNullString
                    .AddItem Bsprintf("Файл с документом (%s)", _
                        LCase(File)) & FileDT(SMail.Recv & File)
                    File = Dir
                Loop
                File = Dir(SMail.Recv & "*.jpg")
                Do While File <> vbNullString
                    .AddItem Bsprintf("Файл с документом (%s)", _
                        LCase(File)) & FileDT(SMail.Recv & File)
                    File = Dir
                Loop
                File = Dir(SMail.Recv & "*.cer")
                Do While File <> vbNullString
                    .AddItem Bsprintf("Сертификат (%s)", _
                        LCase(File)) & FileDT(SMail.Recv & File)
                    File = Dir
                Loop
                File = Dir(SMail.Recv & "*.doc")
                Do While File <> vbNullString
                    .AddItem Bsprintf("Файл MS Word (%s)", _
                        LCase(File)) & FileDT(SMail.Recv & File)
                    File = Dir
                Loop
                File = Dir(SMail.Recv & "*.rtf")
                Do While File <> vbNullString
                    .AddItem Bsprintf("Файл MS Word (%s)", _
                        LCase(File)) & FileDT(SMail.Recv & File)
                    File = Dir
                Loop
                File = Dir(SMail.Recv & "*.xls")
                Do While File <> vbNullString
                    .AddItem Bsprintf("Файл MS Excel (%s)", _
                        LCase(File)) & FileDT(SMail.Recv & File)
                    File = Dir
                Loop
                File = Dir(SMail.Recv & "*.chm")
                Do While File <> vbNullString
                    .AddItem Bsprintf("Файл помощи (%s)", _
                        LCase(File)) & FileDT(SMail.Recv & File)
                    File = Dir
                Loop
                
                File = Dir(SMail.Recv & "o???????." & myID)
                Do While File <> vbNullString
                    If LCase(Left(File, 3)) <> "ok-" Then
                        .AddItem "Принятый " & DocName(File) & FileDT(SMail.Recv & File)
                    End If
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
                
                File = Dir(SMail.Recv & "cur?????.dbf")
                Do While File <> vbNullString
                    If FileDateTime(SMail.Recv & File) >= dt Then
                        .AddItem Bsprintf("Курс валют ЦБ (%s)", LCase(File)) & FileDT(SMail.Recv & File)
                    End If
                    File = Dir
                Loop
                
                lblFilesLen = Bsprintf("%d/%d", _
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
                
                lblFilesLen = Bsprintf("%d/%d", _
                    .ListCount, CountFiles(SMail.Send & "*.*"))
            End With
            
        Case "A":
            With lstFiles
                .Clear
                
                File = Dir(SMail.Archive & "o???????." & myID)
                Do While File <> vbNullString
                    If FileDateTime(SMail.Archive & File) >= dt Then
                        .AddItem "Принятый " & DocName(File) & FileDT(SMail.Archive & File)
                    End If
                    File = Dir
                Loop
                
                File = Dir(SMail.Archive & "e???????." & myID)
                Do While File <> vbNullString
                    If FileDateTime(SMail.Archive & File) >= dt Then
                        .AddItem "Ошибочный " & DocName(File) & FileDT(SMail.Archive & File)
                    End If
                    File = Dir
                Loop
                
                File = Dir(SMail.Archive & "t???????." & myID)
                Do While File <> vbNullString
                    If FileDateTime(SMail.Archive & File) >= dt Then
                        .AddItem "Тестовый " & DocName(File) & FileDT(SMail.Archive & File)
                    End If
                    File = Dir
                Loop
                
                lblFilesLen = Bsprintf("%d/%d", _
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
    Dim M As Long, dd As Long, nnn As Long, kkk As String
    'aYMDDNNN.kkk
    kkk = FileExt(File)
    If User.IsID4(kkk) Then
        nnn = Val(Mid(File, 6, 3))
        If nnn > 0 Then
            M = Ot36(Mid(File, 3, 1))
            dd = Val(Mid(File, 4, 2))
            DocName = Bsprintf("N %d от %s.%02d (%s)", nnn, dd, M, LCase(File))
        End If
    Else
        DocName = Bsprintf("(%s)", LCase(File))
    End If
End Function

Public Function UpdateName(File As String) As String
    UpdateName = Bsprintf("(%s)", LCase(File))
End Function

Public Function FileDT(File As String) As String
    Dim dt As Date
    dt = FileDateTime(File)
    If DateDiff("d", dt, Date) = 0 Then
        FileDT = Format(dt, " + dd.MM.yy HH:mm")
    Else
        FileDT = Format(dt, " - dd.MM.yy HH:mm")
    End If
End Function

Private Sub UserForm_Terminate()
    App.BoolOptions("ViewAll") = chkAll
    App.BoolOptions("ViewToday") = chkToday
    App.Options("MultiSelect") = lstFiles.MultiSelect
End Sub
