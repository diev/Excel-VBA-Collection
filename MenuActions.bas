Attribute VB_Name = "MenuActions"
Option Explicit
Option Compare Text
Option Base 1
DefLng A-Z

Public Sub LogonShow()
    On Error Resume Next
    Application.StatusBar = "���� � ������� ����-������..."
    AutoCheck
    Logon.Show
    Application.StatusBar = False
End Sub

Public Sub PlatEnterShow()
    On Error Resume Next
    Application.StatusBar = "�������� ���������� ���������..."
    AutoCheck
    PlatEnter.Show
    Application.StatusBar = False
End Sub

Public Sub NewUserShow()
    On Error Resume Next
    Application.StatusBar = "���������� ������ �����������..."
    AutoCheck
    Load NewName
    With NewName
        .Mode = "New"
        .Show
    End With
    Application.StatusBar = False
End Sub

Public Sub EditUserShow()
    On Error Resume Next
    Application.StatusBar = "��������� ���������� �����������..."
    AutoCheck
    Load NewName
    With NewName
        .Mode = "Edit"
        .Show
    End With
    Application.StatusBar = False
End Sub

Public Sub BICShow()
    On Error Resume Next
    Application.StatusBar = "������ �� ������������ ���"
    AutoCheck
    Load NewName
    With NewName
        .Mode = "BIC"
        .Show
    End With
    Application.StatusBar = False
End Sub

Public Sub LSShow()
    On Error Resume Next
    Application.StatusBar = "������ �� ������������ ���"
    AutoCheck
    Load NewName
    With NewName
        .Mode = "LS"
        .Show
    End With
    Application.StatusBar = False
End Sub

Public Sub PayUserShow()
    On Error Resume Next
    Application.StatusBar = "���� ���������� ���������� �������..."
    AutoCheck
    Load NewName
    With NewName
        .Mode = "Pay"
        .Show
    End With
    Application.StatusBar = False
End Sub

Public Sub UserPrivateShow()
    On Error Resume Next
    Application.StatusBar = "����� � ������� ��������� �������..."
    AutoCheck
    UserPrivate.Show
    Application.StatusBar = False
End Sub

Public Sub MailBoxShow()
    On Error Resume Next
    Application.StatusBar = "������ � ��������� ������� ��������� SMail..."
    AutoCheck
    MailBox.Show
    Application.StatusBar = False
End Sub

Public Sub SavePassShow()
    Application.StatusBar = "����������, ����� ������ � �����..."
    AutoCheck
    SaveAsPass.Show
    Application.StatusBar = False
    If App.CloseAllowed Then AutoClose
End Sub

Public Sub ExportPlat()
    On Error Resume Next
    Application.StatusBar = "���������� PGP � �������� ���������� � ���� SMail..."
    AutoCheck
    Payment.EachSelected "ExportPlat", "��������� ���?"
    Application.StatusBar = False
End Sub

Public Sub ExportList()
    On Error Resume Next
    Application.StatusBar = "�������� �� ���� (�������)..."
    AutoCheck
    Payment.ExportToFile
    Application.StatusBar = False
End Sub

Public Sub ImportList()
    On Error Resume Next
    Application.StatusBar = "�������� � ����� (������)..."
    AutoCheck
    Payment.ImportFromFile
    Application.StatusBar = False
End Sub

Public Sub DelRows()
    On Error Resume Next
    Application.StatusBar = "�������� �����..."
    AutoCheck
    Payment.EachSelected "Delete", "������������ �������?"
    Application.StatusBar = False
    If Payment.MoneyLastSelected > 0 Then AddDeletedAmount
End Sub

Public Sub DelUser()
    On Error Resume Next
    Application.StatusBar = "�������� �����������..."
    AutoCheck
    With User
        If .Demo Then
            InfoBox "����������������� ������� ������� ������!"
        ElseIf YesNoBox("������������� ������� �� �������\n������� %s - %s\n� ��� ����� ����������?", _
            .ID, .Name) Then
            .Delete
        End If
    End With
    Application.StatusBar = False
End Sub

Public Sub AutoCheck()
    On Error Resume Next
    If ActiveSheet.Name <> User.ID Then
        Workbooks(App.BookName).Activate
    ElseIf Val(User.BIC) = 0 Then
        WarnBox "��������, ��������� ��������������!"
        Restart
    End If
End Sub

Public Sub AutoRestart(Optional s As String = vbNullString)
    On Error Resume Next
    WarnBox "��������, ��������� ��������������!\n%s", s
    AutoOpen
End Sub

Public Sub Restart()
    On Error Resume Next
    With App
        If YesNoBox("�������� ����� ���������� ����� �����������?") Then
            .Setting(BnkSeek2Section, "LocateCanceled") = 0
            .Setting(PGPSection, "LocateCanceled") = 0
            .Setting(SMailSection, "LocateCanceled") = 0
            BnkSeek2.File = vbNullString
            PGP.File = vbNullString
            SMail.File = vbNullString
        End If
    End With
    AutoOpen
End Sub

Public Sub PreviewPlat()
    On Error Resume Next
    Application.StatusBar = "�������� � ������ ���������..."
    AutoCheck
    Payment.EachSelected "Preview", "�������� ���?"
    Application.StatusBar = False
End Sub

Public Sub Info()
    On Error Resume Next
    Application.StatusBar = "���������� � ���������..."
    AutoCheck
    App.Info
    Application.StatusBar = False
End Sub

Public Sub FindText()
    On Error Resume Next
    Application.StatusBar = "����� ������..."
    AutoCheck
    Payment.FindText
    Application.StatusBar = False
End Sub

'Public Sub FindNext()
'    On Error Resume Next
'    Application.StatusBar = "����� ������ �����..."
'    AutoCheck
'    Payment.FindNext
'    Application.StatusBar = False
'End Sub
'
'Public Sub FindPrev()
'    On Error Resume Next
'    Application.StatusBar = "����� ������ ������..."
'    AutoCheck
'    Payment.FindPrev
'    Application.StatusBar = False
'End Sub

Public Sub SortByDocNo()
    On Error Resume Next
    Application.StatusBar = "���������� ����� �� ������..."
    AutoCheck
    Payment.SortBy 3
    Application.StatusBar = False
End Sub

Public Sub SortByDocDate()
    On Error Resume Next
    Application.StatusBar = "���������� ����� �� ����..."
    AutoCheck
    Payment.SortBy 4
    Application.StatusBar = False
End Sub

Public Sub SortBySum()
    On Error Resume Next
    Application.StatusBar = "���������� ����� �� �����..."
    AutoCheck
    Payment.SortBy 5
    Application.StatusBar = False
End Sub

Public Sub SortByName()
    On Error Resume Next
    Application.StatusBar = "���������� ����� �� ����������..."
    AutoCheck
    Payment.SortBy 6
    Application.StatusBar = False
End Sub

Public Sub SortByDetails()
    On Error Resume Next
    Application.StatusBar = "���������� ����� �� ����������..."
    AutoCheck
    Payment.SortBy 11
    Application.StatusBar = False
End Sub

Public Sub AmountChange()
    Dim s As String, c As Currency
    On Error Resume Next
    Application.StatusBar = "��������� �������� ������� �������..."
    AutoCheck
    c = User.Amount
    s = RSumStr(c, vbCrLf)
    If c < 0 Then s = "�����! " & s
    s = Bsprintf("������ �� ����� %f\n(%s)\n\n������� ����� �������:", c, s)
    'If c < 0 Then s = s & BSPrintF("\n������� ������ ����!")
    s = s & Bsprintf("\n\n(����������� \'+\' � \'-\', ����� �������� � �������\n��� ������)")
    s = InputBox(s, App.Title, PlatFormat(c))
    If Len(s) > 0 Then
        c = RVal(s)
        If Left(s, 1) = "+" Then
            c = User.Amount + c
        ElseIf c < 0 Then
            c = User.Amount + c 'negative value!
        End If
        s = RSumStr(c, vbCrLf)
        If c < 0 Then s = "�����! " & s
        If YesNoBox("��������� ������� ������� %f?\n(%s)", c, s) Then
            User.Amount = c
        End If
    End If
    Application.StatusBar = False
End Sub

Public Sub AddDeletedAmount()
    Dim s As String, c As Currency
    On Error Resume Next
    Application.StatusBar = "������� ������� ����� ��������..."
    AutoCheck
    c = User.Amount
    s = RSumStr(c, vbCrLf)
    If c < 0 Then s = "�����! " & s
    s = Bsprintf("������ �� ����� %f\n(%s)\n\n�������� ��������� ����� %f?", c, s, _
        Payment.MoneyLastSelected)
    'If c < 0 Then s = s & BSPrintF("\n������� ������ ����!")
    s = InputBox(s, App.Title, Bsprintf("+%F", Payment.MoneyLastSelected))
    If Len(s) > 0 Then
        c = RVal(s)
        If Left(s, 1) = "+" Then
            c = User.Amount + c
        ElseIf c < 0 Then
            c = User.Amount + c 'negative value!
        End If
        s = RSumStr(c, vbCrLf)
        If c < 0 Then s = "�����! " & s
        If YesNoBox("��������� ������� ������� %f?\n(%s)", c, s) Then
            User.Amount = c
        End If
    End If
    Application.StatusBar = False
End Sub

Public Sub UpdateReceived(File As String)
    If OkCancelBox("�������� ����������� ������� ����������:\n\n" & _
        "1. ����� ������� ������ \'OK\' ��������� Excel ��������\n" & _
        "���� ������� ���� %s � ���������.\n\n" & _
        "2. ����� ������� ��������������������� ����\n" & _
        "��������������� ���������� %s\n" & _
        "��� ������� ���������� %s.\n\n" & _
        "3. ������ ����������, ��� Excel ��������, ������� ������ \'OK\'\n" & _
        "� ���������� ����������. �� ��� ���������� ����� �����\n" & _
        "������� ����� � ���������� ������ � ���������� ����-������.\n\n" & _
        "4. ����� ���������� ���� ������� ������� �� ���������.\n\n" & _
        "������� Excel � ��������� ���������� ����� ������?", _
        ActiveWorkbook.FullName, File, App.Path) Then
        
        ActiveWorkbook.Save
        App.CloseAllowed = True
        Shell File & " /auto " & App.Path, vbNormalNoFocus
        AutoClose
    End If
End Sub


Public Sub PrintVeksel()
    Dim OldDate As Date, s As String
    On Error Resume Next
    OldDate = Date
    If YesNoBox("������ ������� �� " & Format(Now - 1, "dddd dd.MM.yyyy HH:mm")) Then
        Date = Date - 1
    Else
        s = Format(Date - 1, "dd.MM.yy")
        s = InputBox("������� ���� ��� �������:", App.Title, s)
        If Len(s) = 0 Then Exit Sub
        Date = CDate(s)
    End If
    DoEvents
    
    With Payment
        .ReadRow
        s = .FileName & "." & User.ID
        s = RightPathName(GetWinTempDir, s)
        If .SavePlat(s) Then
            DoEvents
            Time = Time + 0.001
            DoEvents
            ShellWait App.Path & "prnveksl.exe " & s
            DoEvents
        Else
            StopBox "���� ��� ������ �� ������!"
        End If
    End With
    Date = OldDate
    Time = Time - 0.001
    DoEvents
End Sub

