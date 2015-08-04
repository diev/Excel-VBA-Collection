Attribute VB_Name = "Menu97"
Option Explicit
Option Compare Text
Option Base 1
DefLng A-Z

Const MenuName = "&����-������"
Const BarName = "���� ������ ����"

Const MenuCount = 6

Dim mMenu(1 To MenuCount) As Variant
Dim mBar(1 To MenuCount) As Variant

Private Sub MainMenu()
    AddItem 1, "&�����...", "FindText", 279
    AddItem 1, "&�������...", "PlatEnterShow", 64
    AddItem 1, "&��������� � �����...", "ImportList", 270
    AddItem 1, "-&�������� � ������", "PreviewPlat", 2174
    If IsFile(App.Path & "prnveksl.exe") Then
        AddItem 1, "&������ �������", "PrintVeksel", 2174
    End If
    AddItem 1, "�&������� �� ����...", "ExportList", 271
    AddItem 1, "��&��������� � ��������", "ExportPlat", 277
    AddItem 1, "-&�������� � �����...", "MailBoxShow", 275
    AddMenu 1, "-�&�����", 2
    AddItem 1, "-� �&��������", "Info", 1954

    AddItem 2, "&������� ������", "DelRows", 67
    AddItem 2, "-&��������� �������...", "EditUserShow", 2148
    AddItem 2, "&�������� �������...", "NewUserShow", 2141
    AddItem 2, "&������������� �����", "ImportNewKeys", 277
    AddMenu 2, "-&���������", 3
    AddMenu 2, "&�������", 4
    AddMenu 2, "&������", 5
    AddMenu 2, "&�����", 6
    AddItem 2, "-&���������� ���������", "Restart", 2144
    
    AddItem 3, "&��������� ����� (SAdm)...", "SAdm", 29
    AddItem 3, "�&�������� ������ (SSetup)...", "SSetup", 29
    AddItem 3, "-&��������� ����� (SMail.ctl)...", "SMailCtl", 29
    AddItem 3, "�&������� ����� (SMail.log)...", "SMailLog", 29
    AddItem 3, "-�&����� ��������� ID...", "EditID", 29
    
    AddItem 4, "&��������� ������� � �����", "AskVypRemart", 29
    AddItem 4, Bsprintf("&���������� ������ (������ %n)", BnkSeek2.Updated), "AskBnkSeek", 29
    AddItem 4, Bsprintf("&���������� ��������� (������ %s)", App.Version), "AskBClient", 29
    
    AddItem 5, "&������� ������ Excel...", "ExcelPassword", 29
    AddItem 5, "�&������ ������ PGP...", "PGPPassword", 29
    AddItem 5, "�&������ ������ SMail...", "SAdm", 29
    
    AddItem 6, "-&����� ���������...", "OpenFolder", 29
    AddItem 6, "�&���� ���������...", "OpenFolderR", 29
    AddItem 6, "-&��������� ����(�)...", "SendFiles", 29
    AddItem 6, "&�������� �������...", "SendNote", 29
End Sub

Public Sub InitMenuBars()
    On Error Resume Next
    Application.ScreenUpdating = False
    CommandBars(BarName).Delete
    CommandBars.ActiveMenuBar.Reset 'MenuBars(xlWorksheet).Reset
    Set mMenu(1) = CommandBars.ActiveMenuBar.Controls.Add(Type:=msoControlPopup, before:=1, temporary:=True)
    mMenu(1).Caption = MenuName
    Set mBar(1) = CommandBars.Add(Name:=BarName, Position:=1, temporary:=True)
    MainMenu
    CommandBars(BarName).Visible = True
    Application.ScreenUpdating = True
End Sub

Public Sub CloseMenuBars()
    On Error Resume Next
    Application.ScreenUpdating = False
    CommandBars(BarName).Delete
    CommandBars.ActiveMenuBar.Reset
    Application.ScreenUpdating = True
End Sub

Private Sub AddItem(Level As Long, Caption As String, Optional Macro As String = vbNullString, Optional Icon As Long = 0)
    Dim Sep As Boolean
    On Error Resume Next
    Sep = Left(Caption, 1) = "-"
    If Sep Then
        Caption = Mid(Caption, 2)
    End If
    With mMenu(Level).CommandBar.Controls.Add(Type:=msoControlButton, ID:=1, temporary:=True)
        .BeginGroup = Sep
        .Caption = Caption
        .OnAction = Macro
        .FaceId = Icon
        .Style = msoButtonIconAndCaption
    End With
    If Level = 1 Then
        With mBar(Level).Controls.Add(Type:=msoControlButton, ID:=1, temporary:=True)
            .BeginGroup = Sep
            .TooltipText = Caption
            .OnAction = Macro
            .FaceId = Icon
            .Style = msoButtonIcon
        End With
    Else
        With mBar(Level).Controls.Add(Type:=msoControlButton, ID:=1, temporary:=True)
            .BeginGroup = Sep
            .Caption = Caption
            .OnAction = Macro
            .FaceId = Icon
            .Style = msoButtonIconAndCaption
        End With
    End If
End Sub

Private Sub AddMenu(Level As Long, Caption As String, SubLevel As Long)
    On Error Resume Next
    Dim Sep As Boolean
    On Error Resume Next
    Sep = Left(Caption, 1) = "-"
    If Sep Then
        Caption = Mid(Caption, 2)
    End If
    Set mMenu(SubLevel) = mMenu(Level).CommandBar.Controls.Add(Type:=msoControlPopup, temporary:=True)
    With mMenu(SubLevel)
        .BeginGroup = Sep
        .Caption = Caption
    End With
    Set mBar(SubLevel) = mBar(Level).Controls.Add(Type:=msoControlPopup, temporary:=True)
    With mBar(SubLevel)
        .BeginGroup = Sep
        .TooltipText = Caption
        .Caption = .TooltipText
    End With
End Sub

