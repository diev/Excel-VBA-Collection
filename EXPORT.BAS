Attribute VB_Name = "Export"
Option Explicit
Option Compare Text
Option Base 1
DefLng A-Z

Public Sub ExportPlat()
    Dim s As String, FileName1 As String, FileName2 As String
    
    'If GetSetting(AppName, "User", "BIC") <> Range("БИК1").Text Then
    '    MsgBox "Этот платеж не обрабатывается нашей системой!", vbCritical, AppTitle
    '    Exit Sub
    'End If
    'If GetSetting(AppName, "User", "LS") <> Range("Счет1").Text Then
    '    MsgBox "Этот платеж не обрабатывается нашей системой!", vbCritical, AppTitle
    '    Exit Sub
    'End If
    
    s = Range("Файл").Text
    FileName1 = RightPathName(GetWinTempDir, s)
    FileName2 = RightPathName(Range("Out").Text, s)
    If FileExists(FileName2) Then
        If MsgBox(BPrintF("Перезаписать файл?\n~%s~", s), vbExclamation + vbYesNo, _
            AppTitle) = vbNo Then
            Exit Sub
        End If
    Else
        If Len(s) = 0 Then
            MsgBox "Нечего отправлять!", vbExclamation, AppTitle
            Exit Sub
        End If
        If MsgBox(BPrintF("Отправить файл?\n~%s~", s), vbQuestion + vbYesNo, _
            AppTitle) = vbNo Then
            Exit Sub
        End If
    End If
    OutputFile FileName1, ExportString
    PGPRun Range("PGPCode").Text, FileName1, FileName2
    
    Range("Отправлено") = "ОТПРАВЛЕНО!"
    Archive.Mark(Left(s, 8)) = Right(s, 3)
End Sub

Public Function ExportString() As String
    Dim s As String, ss As String, arr As Variant
    s = BPrintF("# %s\n# %s\n# --%s--\n", Range("Файл").Text, CDos(AppTitle), Now)
    
    s = s & BPrintF("\n# [Payment]\n")
    s = s & BPrintF("No   = %d\n", Range("Номер").Text)
    s = s & BPrintF("Date = %s\n", Range("Дата").Text)
    
    s = s & BPrintF("\n# [Payer1]\n")
    s = s & BPrintF("Name1  = %s\n", CDos(Range("Название1").Text))
    s = s & BPrintF("INN1   = %s\n", Range("ИНН1").Text)
    s = s & BPrintF("LS1    = %s\n", Range("Счет1").Text)
    s = s & BPrintF("BIC1   = %s\n", Range("БИК1").Text)
    s = s & BPrintF("Bank1  = %s\n", CDos(Range("Банк1").Text))
    s = s & BPrintF("Place1 = %s\n", CDos(Range("Место1").Text))
    s = s & BPrintF("KS1    = %s\n", Range("Корсчет1").Text)
    
    s = s & BPrintF("\n# [Payee2]\n")
    s = s & BPrintF("Name2  = %s\n", CDos(Range("Название2").Text))
    s = s & BPrintF("INN2   = %s\n", Range("ИНН2").Text)
    s = s & BPrintF("LS2    = %s\n", Range("Счет2").Text)
    s = s & BPrintF("BIC2   = %s\n", Range("БИК2").Text)
    s = s & BPrintF("Bank2  = %s\n", CDos(Range("Банк2").Text))
    s = s & BPrintF("Place2 = %s\n", CDos(Range("Место2").Text))
    s = s & BPrintF("KS2    = %s\n", Range("Корсчет2").Text)
    
    s = s & BPrintF("\n# [Sum]\n")
    s = s & BPrintF("Amount = %s\n", Range("Сумма").Text)
    s = s & BPrintF("Text   = %s\n", CDos(Range("СуммаПрописью").Text))
    
    s = s & BPrintF("\n# [Description]\n")
    ss = CDos(Range("Назначение").Text)
    arr = StrToLines(ss, 68, 4)
    s = s & BPrintF("D1 = %s\n", arr(1))
    s = s & BPrintF("D2 = %s\n", arr(2))
    s = s & BPrintF("D3 = %s\n", arr(3))
    s = s & BPrintF("D4 = %s\n", arr(4))
    
    s = s & BPrintF("\n# [Misc]\n")
    s = s & BPrintF("Delivery = %s\n", CDos(Range("ВидПлатежа").Text))
    s = s & BPrintF("Action   = %s\n", Range("Срок").Text)
    s = s & BPrintF("Queue    = %d\n", Range("Очередность").Text)
    
    s = s & BPrintF("\n# [Sys]\n")
    s = s & BPrintF("ID     = %s\n", Range("ID").Text)
    s = s & BPrintF("Phone  = %s\n", CDos(Range("Phone").Text))
    s = s & BPrintF("File   = %s\n", Range("Файл").Text)
    s = s & BPrintF("Date   = %s\n", DtoS(Date))
    s = s & BPrintF("Time   = %s\n", Time)
    s = s & BPrintF("Op     = %s\n", IIf(Range("БИК1").Text = Range("БИК2").Text, "801", "802"))
    s = s & BPrintF("Type   = %s\n", Range("Тип").Text)
    s = s & BPrintF("Parcel = %s\n", Range("Посылка").Text)
    s = s & BPrintF("Oper   = %s\n", Range("Oper").Text)
    s = s & BPrintF("Folder = %s\n", Range("Folder").Text)
    
    ExportString = s & BPrintF("\n# [eof]\n")
End Function

