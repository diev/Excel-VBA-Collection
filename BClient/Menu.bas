Attribute VB_Name = "Menu"
Option Explicit
Option Compare Text
Option Base 1
DefLng A-Z

Public Sub ResetMenu()
    MenuBars(xlWorksheet).Reset
    Application.StatusBar = False
End Sub

Public Sub SetMenu()
    'Dim v As Variant, i, m As String, s As String, ss As String
    Dim m As String
    With MenuBars(xlWorksheet)
        .Reset
        'v = GetAllSettings(AppName, "Menus")
        'For i = LBound(v, 1) To UBound(v, 1)
        '    If i = LBound(v, 1) Then
        '        m = CStr(v(i, 0))
        '        s = CStr(v(i, 1))
        '        If Len(m) = 0 Then m = "&Банк-Клиент"
        '        If Len(s) = 0 Then s = "6" 'before Сервис/Tools
        '        .Menus.Add m, before:=Val(s)
        '    Else
        '        s = CStr(v(i, 0))
        '        If Left(s, 1) = "-" Then
        '            .Menus(m).MenuItems.Add "-"
        '        Else
        '            ss = CStr(v(i, 1))
        '            .Menus(m).MenuItems.Add s, ss
        '        End If
        '    End If
        'Next
        
        m = "&Банк-Клиент"
        .Menus.Add m, before:=6
        With .Menus(m).MenuItems
            .Add "&Ввод поручения...", "PlatEnterShow"
            .Add "-"
            .Add "&Просмотр", "PreviewPlat"
            .Add "П&ечать", "PrintPlat"
            .Add "О&тправить в Банк", "ExportPlat"
            .Add "-"
            '.Add "&Импорт из файла...", "ImportPlat"
            '.Add "&Настройки...", "UserOptionsShow"
            .Add "-"
            .Add "Почта в &Банк...", "SMailSend"
            .Add "Сеанс &связи SMail", "SMailDial"
            .Add "Почта из Б&анка...", "SMailRecv"
            .Add "-"
            .Add "&Добавить реквизиты...", "NewNameShow"
            .Add "В&зять поручение из архива", "ArchivedPlat"
            .Add "-"
            .Add "Пе&резапуск", "Reset"
            .Add "-"
            '.Add "&Справка", "HelpShow"
            .Add "&О программе", "AboutShow"
        End With
    End With
End Sub

Public Sub PlatEnterShow()
    PlatEnter.Show
End Sub

Public Sub PreviewPlat()
    'Application.Goto BnkCliWorkbook & "!Платежка", True
    'ActiveWindow.SelectedSheets.PrintPreview
    'Worksheets(Payment.PlatSheet).PrintPreview
    Worksheets(PlatSheet).PrintPreview
End Sub

Public Sub PrintPlat()
    'Application.Goto BnkCliWorkbook & "!Платежка", True
    'ActiveWindow.SelectedSheets.PrintOut
    'Worksheets(Payment.PlatSheet).PrintOut
    Worksheets(PlatSheet).PrintOut
End Sub

Public Sub NewNameShow()
    Do
        Load NewName
        With NewName
            .AddNewOk = False
            .Show
        End With
        Unload NewName
    Loop While MsgBox("Добавить еще?", vbYesNo + vbQuestion, AppTitle) = vbYes
End Sub

'Public Sub UserOptionsShow()
'    UserOptions.Show
'End Sub

Public Sub Reset()
    If MsgBox("Вы действительно хотите перезапустить программу?", vbCritical + vbYesNo, AppTitle) = vbYes Then
        AutoOpen
    End If
End Sub

Public Sub AboutShow()
    About.Show
End Sub

Public Sub SMailDial()
    Dim s As String
    s = Range("Dial").Text
    On Error Resume Next
    If Shell(s, vbNormalFocus) = 0 Then
        ErrorBox Err, BPrintF("Ошибка запуска\n~%s~", s), AppTitle
    End If
End Sub

Public Sub ArchivedPlat()
    Dim n As Long
    With Archive
        Worksheets(.Sheet).Activate
        n = ActiveCell.Row
        .FillByRow n
    End With
    Worksheets(PlatSheet).Activate
End Sub

Public Sub Reset()
    If MsgBox("Вы действительно хотите перезапустить программу?", vbCritical + vbYesNo, AppTitle) = vbYes Then
        AutoOpen
    End If
End Sub

Public Sub Info()
    AboutShow
End Sub
