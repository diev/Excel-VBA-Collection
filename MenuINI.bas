Attribute VB_Name = "MenuINI"
Option Explicit
Option Compare Text
Option Base 1
DefLng A-Z

Const MenuSection = "Menu"
Const ServiceMenu = "ServiceMenu"
Const SortMenu = "SortMenu"

Public Sub WriteMenuINI()
    Dim i As Long, s As String
    With App
        '�������� ������ ���� � Excel
        .Setting(MenuSection, "Caption") = CDos("&����-������")
        .Setting(MenuSection, "Before") = 1 '0 - �� ���������
        
        '����������� ���� �� ������ ������ ����
        '.Setting(MenuSection, "RClick") = 1
        
        '�������� ������
        .Setting(MenuSection, "Bar") = CDos("���� ������ ����")
        .Setting(MenuSection, "BarAdd") = 1
        .Setting(MenuSection, "BarPosition") = 1 '0=left, 1=top, 2=right, 3=bottom, 4=floating
        .Setting(MenuSection, "BarVisible") = 1
        
        '������������ ������
        '.Setting(MenuSection, "Bar1") = CDos("���� ������ ���� ����")
        '.Setting(MenuSection, "Bar1Menu") = CDos("[����] ����-������")
        '.Setting(MenuSection, "Bar1Add") = 1
        '.Setting(MenuSection, "Bar1Position") = 3 '0=left, 1=top, 2=right, 3=bottom, 4=floating
        '.Setting(MenuSection, "Bar1Visible") = 0
        
        i = 0
        i = i + 1: .Setting(MenuSection, CStr(i)) = CDos("&���� � �������...") & "\LogonShow\59"
        i = i + 1: .Setting(MenuSection, CStr(i)) = CDos("-&�����...") & "\FindText\279"
        i = i + 1: .Setting(MenuSection, CStr(i)) = CDos("&�������...") & "\PlatEnterShow\64"
        i = i + 1: .Setting(MenuSection, CStr(i)) = CDos("&��������� � �����...") & "\ImportList\270"
        i = i + 1: .Setting(MenuSection, CStr(i)) = CDos("-&�������� � ������") & "\PreviewPlat\2174"
        i = i + 1: .Setting(MenuSection, CStr(i)) = CDos("�&������� �� ����...") & "\ExportList\271"
        i = i + 1: .Setting(MenuSection, CStr(i)) = CDos("��&��������� � ��������") & "\ExportPlat\277"
        i = i + 1: .Setting(MenuSection, CStr(i)) = CDos("-&�������� � �����...") & "\MailBoxShow\275"
        i = i + 1: .Setting(MenuSection, CStr(i)) = CDos("-�&�����") & "\\" & ServiceMenu & "\\" & CDos("������")
        i = i + 1: .Setting(MenuSection, CStr(i)) = CDos("� �&��������") & "\Info\1954"
        i = i + 1: .Setting(MenuSection, CStr(i)) = CDos("-�&���� �� ������ ������") & "\SavePassShow\276"
        .Setting(MenuSection, "Count") = i
        
        i = 0
        i = i + 1: .Setting(ServiceMenu, CStr(i)) = CDos("&���������� ���...") & "\LSShow\176"
        i = i + 1: .Setting(ServiceMenu, CStr(i)) = CDos("-�&��������� �����") & "\\" & SortMenu & "\\" & CDos("����������")
        i = i + 1: .Setting(ServiceMenu, CStr(i)) = CDos("&������� ������") & "\DelRows\67"
        i = i + 1: .Setting(ServiceMenu, CStr(i)) = CDos("-&������� �������...") & "\AmountChange\52"
        i = i + 1: .Setting(ServiceMenu, CStr(i)) = CDos("-&��������� �������...") & "\EditUserShow\2148"
        'i = i + 1: .Setting(ServiceMenu,CStr(i)) = CDos("����� �������...") & "\UserPrivateShow\2148"
        i = i + 1: .Setting(ServiceMenu, CStr(i)) = CDos("&�������� �������...") & "\NewUserShow\2141"
        i = i + 1: .Setting(ServiceMenu, CStr(i)) = CDos("��&����� �������") & "\DelUser\2151"
        i = i + 1: .Setting(ServiceMenu, CStr(i)) = CDos("-&���������� ���������") & "\Restart\2144"
        .Setting(ServiceMenu, "Count") = i
    
        i = 0
        i = i + 1: .Setting(SortMenu, CStr(i)) = CDos("&�����") & "\SortByDocNo\29"
        i = i + 1: .Setting(SortMenu, CStr(i)) = CDos("&����") & "\SortByDocDate\29"
        i = i + 1: .Setting(SortMenu, CStr(i)) = CDos("&�����") & "\SortBySum\29"
        i = i + 1: .Setting(SortMenu, CStr(i)) = CDos("&����������") & "\SortByName\29"
        i = i + 1: .Setting(SortMenu, CStr(i)) = CDos("�&���������") & "\SortByDetails\29"
        .Setting(SortMenu, "Count") = i
    End With
End Sub

