VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CMenu97"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Option Base 1
DefLng A-Z

Const MenuSection = "Menu"
Const ServiceMenu = "ServiceMenu"

Dim mMenu As Variant
Dim mRMenu As Variant
Dim mBar As Variant
Dim mBar1 As Variant

Public Sub InitMenuBars()
    On Error Resume Next
    Application.StatusBar = "Загрузка меню..."
    Application.ScreenUpdating = False
    AddMenu
    AddRClick
    AddBar
    AddBar1
    Application.ScreenUpdating = True
    Application.StatusBar = False
End Sub

Public Sub CloseMenuBars()
    On Error Resume Next
    Application.ScreenUpdating = False
    'If mBar <> Nothing Then
        With mBar
            App.Setting(MenuSection, "BarPosition") = .Position
            App.Setting(MenuSection, "BarVisible") = BoolToStr(.Visible)
            .Delete
        End With
    'End If
    'If mBar1 <> Nothing Then
        With mBar1
            App.Setting(MenuSection, "Bar1Position") = .Position
            App.Setting(MenuSection, "Bar1Visible") = booltistr(.Visible)
            .Delete
        End With
    'End If
    CommandBars.ActiveMenuBar.Reset
    CommandBars("cell").Reset
    Application.ScreenUpdating = True
End Sub

Private Sub AddMenu()
    Dim B As Variant
    On Error Resume Next
    'mMenu.Delete
    CommandBars.ActiveMenuBar.Reset 'MenuBars(xlWorksheet).Reset
    With CommandBars.ActiveMenuBar.Controls
        B = App.Setting(MenuSection, "Before")
        If B > 0 Then
            If B > .Count Then 'Last
                Set mMenu = .Add(Type:=msoControlPopup, temporary:=True)
            Else
                Set mMenu = .Add(Type:=msoControlPopup, before:=B, temporary:=True)
            End If
            AddSubMenu mMenu, App.Setting(MenuSection, "Caption") & "\\" & MenuSection
        End If
    End With
End Sub
    
Private Sub AddRClick()
    On Error Resume Next
    'mRMenu.Delete
    CommandBars("cell").Reset 'internal name!
    If StrToBool(App.Setting(MenuSection, "RClick")) Then
        Set mRMenu = CommandBars("cell").Controls.Add(Type:=msoControlPopup, _
            before:=1, temporary:=True)
        AddSubMenu mRMenu, App.Setting(MenuSection, "Caption") & "\\" & MenuSection
        CommandBars("cell").Controls(2).BeginGroup = True
    End If
End Sub
    
Private Sub AddBar()
    Dim A As Variant, i As Long, n As Long, s As String, Item As Variant
    On Error Resume Next
    mBar.Delete
    s = App.Setting(MenuSection, "Bar")
    If StrToBool(App.Setting(MenuSection, "BarAdd")) Then
        Set mBar = CommandBars.Add(Name:=s, Position:=App.Setting(MenuSection, "BarPosition"), _
            temporary:=True)
        n = App.Setting(MenuSection, "Count")
        For i = 1 To n
            s = App.Setting(MenuSection, CStr(i))
            If Len(s) > 0 Then
                If InStr(s, "\\") > 0 Then
                    A = StrToArr(s, "\\")
                    s = CStr(A(3)) & "\\" & CStr(A(2))
                    If Left(A(1), 1) = "~" Then
                        If Mid(A(1), 2, 1) = "-" Then
                            s = "~-" & s
                        Else
                            s = "~" & s
                        End If
                    ElseIf Left(A(1), 1) = "-" Then
                        s = "-" & s
                    End If
                    Set Item = mBar.Controls.Add(Type:=msoControlPopup, _
                        temporary:=True)
                    AddSubMenu Item, s
                Else
                    Set Item = mBar.Controls.Add(Type:=msoControlButton, _
                        ID:=1, temporary:=True)
                    AddBarItem Item, s
                End If
            End If
        Next
        mBar.Visible = StrToBool(App.Setting(MenuSection, "BarVisible"))
    End If
End Sub

Private Sub AddBar1()
    Dim A As Variant, i As Long, n As Long, s As String, Item As Variant
    On Error Resume Next
    mBar1.Delete
    s = App.Setting(MenuSection, "Bar1")
    If StrToBool(App.Setting(MenuSection, "Bar1Add")) Then
        Set mBar1 = CommandBars.Add(Name:=s, Position:=App.Setting(MenuSection, "Bar1Position"), _
            temporary:=True)
        Set Item = mBar1.Controls.Add(Type:=msoControlPopup, temporary:=True)
        AddSubMenu Item, App.Setting(MenuSection, "Bar1Menu") & "\\" & MenuSection
        mBar1.Visible = StrToBool(App.Setting(MenuSection, "Bar1Visible"))
    End If
End Sub

Private Sub AddSubMenu(subMenu As Variant, s As String)
    Dim A As Variant, i As Long, n As Long, ss As String, Sec As String, Item As Variant
    On Error Resume Next
    A = StrToArr(s, "\\")
    With subMenu
        ss = CStr(A(1))
        If Left(ss, 1) = "~" Then
            .Enabled = False
            ss = Mid(ss, 2)
        End If
        If Left(ss, 1) = "-" Then
            .BeginGroup = True
            .Caption = Mid(ss, 2)
        Else
            .Caption = ss
        End If
        .Tag = CStr(A(4))
    End With
    With subMenu.CommandBar.Controls
        Sec = CStr(A(2))
        n = App.Setting(Sec, "Count")
        For i = 1 To n
            ss = App.Setting(Sec, CStr(i))
            If Len(ss) > 0 Then
                If InStr(ss, "\\") > 0 Then
                    Set Item = .Add(Type:=msoControlPopup, temporary:=True)
                    AddSubMenu Item, ss
                Else
                    Set Item = .Add(Type:=msoControlButton, ID:=1, temporary:=True)
                    AddMenuItem Item, ss
                End If
            End If
        Next
    End With
End Sub

Private Sub AddMenuItem(newItem As Variant, s As String)
    Dim A As Variant, ss As String
    On Error Resume Next
    A = StrToArr(s, "\")
    With newItem
        ss = CStr(A(1))
        If Left(ss, 1) = "~" Then
            .Enabled = False
            ss = Mid(ss, 2)
        End If
        If Left(ss, 1) = "-" Then
            .BeginGroup = True
            .Caption = Mid(ss, 2)
        Else
            .Caption = ss
        End If
        .OnAction = CStr(A(2))
        .FaceId = CLng(A(3))
        .Tag = CStr(A(4))
        '.Style = msoButtonCaption
        .Style = msoButtonIconAndCaption
    End With
End Sub

Private Sub AddBarItem(newItem As Variant, s As String)
    Dim A As Variant, ss As String
    On Error Resume Next
    A = StrToArr(s, "\")
    With newItem
        ss = CStr(A(1))
        If Left(ss, 1) = "~" Then
            .Enabled = False
            ss = Mid(ss, 2)
        End If
        If Left(ss, 1) = "-" Then
            .BeginGroup = True
            .TooltipText = Mid(ss, 2)
        Else
            .TooltipText = ss
        End If
        '.Caption = Left(.TooltipText, 3)
        .OnAction = CStr(A(2))
        .FaceId = CLng(A(3))
        .Tag = CStr(A(4))
        .Style = msoButtonIcon
        '.Style = msoButtonIconAndCaption
    End With
End Sub

Public Sub WriteINI()
    Dim i As Long, s As String
    With App
        'Название нашего меню в Excel
        .Setting(MenuSection, "Caption") = "&1. Банк-Клиент"
        .Setting(MenuSection, "Before") = "1" '0 - не вставлять
        
        'Контекстное меню по правой кнопке мыши
        .Setting(MenuSection, "RClick") = "1"
        
        'Линейный тулбар
        .Setting(MenuSection, "Bar") = "Сити Инвест Банк"
        .Setting(MenuSection, "BarAdd") = "1"
        .Setting(MenuSection, "BarPosition") = "1" '0=left, 1=top, 2=right, 3=bottom, 4=floating
        .Setting(MenuSection, "BarVisible") = "1"
        
        'Менюобразный тулбар
        .Setting(MenuSection, "Bar1") = "Сити Инвест Банк Меню"
        .Setting(MenuSection, "Bar1Menu") = "[Пуск] Банк-Клиент"
        .Setting(MenuSection, "Bar1Add") = "1"
        .Setting(MenuSection, "Bar1Position") = "3" '0=left, 1=top, 2=right, 3=bottom, 4=floating
        .Setting(MenuSection, "Bar1Visible") = "0"
        
        i = 0
        i = i + 1: .Setting(MenuSection, CStr(i)) = "&1. Вход в систему...\LogonShow\59\BPS"
        i = i + 1: .Setting(MenuSection, CStr(i)) = "-&2. Создать...\PlatEnterShow\64\BPS"
        i = i + 1: .Setting(MenuSection, CStr(i)) = "&3. Загрузить с диска...\ImportList\270\BPS"
        i = i + 1: .Setting(MenuSection, CStr(i)) = "-&4. Просмотр и печать\PreviewPlat\2174\BPS"
        i = i + 1: .Setting(MenuSection, CStr(i)) = "&5. Записать на диск...\ExportList\271\BPS"
        i = i + 1: .Setting(MenuSection, CStr(i)) = "&6. Зашифровать к отправке\ExportPlat\277\BPS"
        i = i + 1: .Setting(MenuSection, CStr(i)) = "-&7. Отправка и прием...\MailBoxShow\275\BPS"
        i = i + 1: .Setting(MenuSection, CStr(i)) = "-&8. Сервис\\" & ServiceMenu & "\\Сервис\\BPS"
        i = i + 1: .Setting(MenuSection, CStr(i)) = "&9. О программе\Info\1954\BPS"
        i = i + 1: .Setting(MenuSection, CStr(i)) = "-&0. Выход со сменой пароля\SavePassShow\276\BPS"
        .Setting(MenuSection, "Count") = i
        
        i = 0
        i = i + 1: .Setting(ServiceMenu, CStr(i)) = "&1. Справочник банков РФ...\LSShow\176\BPS"
        i = i + 1: .Setting(ServiceMenu, CStr(i)) = "-&2. Удалить строки\DelRows\67\BPS"
        i = i + 1: .Setting(ServiceMenu, CStr(i)) = "-&3. Реквизиты клиента...\EditUserShow\2148\BPS"
        'i = i + 1: .Setting(ServiceMenu,CStr(i)) = "&4. Ключи клиента...\UserPrivateShow\2148\BPS"
        i = i + 1: .Setting(ServiceMenu, CStr(i)) = "-&4. Добавить клиента...\NewUserShow\2141\BPS"
        i = i + 1: .Setting(ServiceMenu, CStr(i)) = "-&0. Перезапуск программы\Restart\2144\BPS"
        .Setting(ServiceMenu, "Count") = i
    End With
End Sub

Public Sub Reset()
    If App.Setting(MenuSection, "Caption") = vbNullString Then WriteINI
    InitMenuBars
End Sub
