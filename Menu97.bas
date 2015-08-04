Attribute VB_Name = "Menu97"
Option Explicit
Option Compare Text
Option Base 1
DefLng A-Z

Const MenuSection = "Menu"

Dim mMenu As Variant
Dim mRMenu As Variant
Dim mBar As Variant
Dim mBar1 As Variant

Public Sub InitMenuBars()
    On Error Resume Next
    Application.StatusBar = "Загрузка меню..."
    If App.Setting(MenuSection, "Caption") = vbNullString Then WriteMenuINI
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
            App.Setting(MenuSection, "Bar1Visible") = BoolToStr(.Visible)
            .Delete
        End With
    'End If
    CommandBars.ActiveMenuBar.Reset
    CommandBars("cell").Reset
    Application.ScreenUpdating = True
End Sub

Private Sub AddMenu()
    Dim b As Variant
    On Error Resume Next
    'mMenu.Delete
    CommandBars.ActiveMenuBar.Reset 'MenuBars(xlWorksheet).Reset
    With CommandBars.ActiveMenuBar.Controls
        b = App.Setting(MenuSection, "Before")
        If b > 0 Then
            If b > .Count Then 'Last
                Set mMenu = .Add(Type:=msoControlPopup, temporary:=True)
            Else
                Set mMenu = .Add(Type:=msoControlPopup, before:=b, temporary:=True)
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
    If StrToBool(App.Setting(MenuSection, "BarAdd")) Then
        Set mBar = CommandBars.Add(Name:=CWin(App.Setting(MenuSection, "Bar")), _
            Position:=App.Setting(MenuSection, "BarPosition"), temporary:=True)
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
    Dim A As Variant, i As Long, n As Long, Item As Variant
    On Error Resume Next
    mBar1.Delete
    If StrToBool(App.Setting(MenuSection, "Bar1Add")) Then
        Set mBar1 = CommandBars.Add(Name:=CWin(App.Setting(MenuSection, "Bar1")), _
            Position:=App.Setting(MenuSection, "Bar1Position"), temporary:=True)
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
            .Caption = CWin(Mid(ss, 2))
        Else
            .Caption = CWin(ss)
        End If
        .Enabled = .Enabled And CheckEnabled(CStr(A(4)))
        '.Tag = CStr(A(4))
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
            .Caption = CWin(Mid(ss, 2))
        Else
            .Caption = CWin(ss)
        End If
        .OnAction = CStr(A(2))
        .FaceId = CLng(A(3))
        .Enabled = .Enabled And CheckEnabled(CStr(A(4)))
        '.Tag = CStr(A(4))
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
            .TooltipText = CWin(Mid(ss, 2))
        Else
            .TooltipText = CWin(ss)
        End If
        '.Caption = Left(.TooltipText, 3)
        .OnAction = CStr(A(2))
        .FaceId = CLng(A(3))
        .Enabled = .Enabled And CheckEnabled(CStr(A(4)))
        '.Tag = CStr(A(4))
        .Style = msoButtonIcon
        '.Style = msoButtonIconAndCaption
    End With
End Sub

Private Function CheckEnabled(Tag As String) As Boolean
    CheckEnabled = True
    If Tag = vbNullString Then Exit Function
    If InStr(Tag, "B") Then CheckEnabled = CheckEnabled And BnkSeek2.Valid
    If InStr(Tag, "P") Then CheckEnabled = CheckEnabled And PGP.Valid
    If InStr(Tag, "S") Then CheckEnabled = CheckEnabled And SMail.Valid
End Function
