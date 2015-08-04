Attribute VB_Name = "Main"
'��������:

'�������� ������ ������� ����-������ ��� "���� ������ ����"
'��������������� �������� � �������� ���� "��� ����" �� ������
'�������� ��� ��������� � ��������� �� ���������� �������.

'����� ��� ����� ���� ��������� ��������� � ��������� �������������.
'����� ������ �� ������������� �� ����������� � ������������
'� ������� �� ������ ��������� ���� ���������� ���������.
'��� ������� ������� ������� ��� �������� ����� ������!

'����� ���������: ������� ���������, 1995-1999-2001
'����� � ���� ��������: http://members.xoom.com/diev/
'E-mail: diev@mail.ru, ICQ: 7372116

'����� �������� � ����������� ����� ��������������!
'�������!

Option Explicit
Option Compare Text
Option Base 1
DefLng A-Z

Public Const BClient2Section = "BClient2"
Public Const BnkSeek2Section = "BnkSeek2"
Public Const PGPSection = "PGP"
Public Const SMailSection = "SMail"

Public App As New CApp
Public BnkSeek2 As New CBnkSeek2
Public User As New CUser
Public Payment As New CPayment
Public PGP As New CPGP
Public SMail As New CSMail

Public Sub AutoOpen()
    'MsgBox because no App here at start!
    Dim s As String
    On Error Resume Next
    
    Application.StatusBar = "��������� �������� � �������� �������..."
    If StrToBool(App.Options("DontAutoOpen")) Then
        MsgBox "���������� ����������!", vbCritical, "���������� �������"
        Exit Sub
    End If
    
    If Not IsDir(App.Options("WorkPath")) Then
        MsgBox Bsprintf("��� �� ���� ������� �����!\n\n������ � ����� %s\n�������� ������ �� %s", App.BookFile, s), vbCritical, "���������� �������"
        Exit Sub
    End If
    
    Application.StatusBar = "�������� ������� ��������� - ����� ����������..."
    InitMenuBars
    Application.StatusBar = "���������������� �����"
    User.Demo = True
    
    PGP.ResetPasswords
    
    s = "����� ���������� � ���������������� ����� ���������!\n" & _
        "����� �� ������ ���������� ��������� ��� ������-���� �����.\n\n" & _
        "��� ������������� �� ��� ������� ����, ���� �� ������ ����������\n" & _
        "�������� � ����� ������ � �������� ������� �� ����.\n\n" & _
        "���� ����� ��� ���, \'�������� �������\' ����� ���� ���� \'����-������\'\n" & _
        "� ������� �� ������������ �� ������ ������� - ������ ����� ��� ����!"
        
    If Workbooks.Count > 1 Then
        s = s & "\n\n��������: ���������� ��� �����-�� �������� ����� Excel!\n" & _
            "��� ����� �������� � ������� � ���������� ����� ���������."
    End If
    
    s = s & "\n\n������ ���������: " & App.Version
    
    InfoHelpBox s, 1
    Application.StatusBar = False
End Sub

Public Sub AutoClose()
    Dim s As String
    On Error Resume Next
    'User.Demo = True
    CloseMenuBars
    s = "� ����� ������������ ��������� ������ � MICROSOFT EXCEL!"
    ActiveWindow.Caption = Empty
    Workbooks(App.BookName).Saved = True 'no more prompts!
    With Application
        .DisplayAlerts = False 'no more prompts!
        .Caption = s
        .StatusBar = s
        DoEvents
        .Quit
    End With
    'Application.Caption = Application.Application
    'Application.StatusBar = False
End Sub

