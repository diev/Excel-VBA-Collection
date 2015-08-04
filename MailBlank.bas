Attribute VB_Name = "MailBlank"
Option Explicit
Option Compare Text
Option Base 1
DefLng A-Z

Const MailSheet = "Почта"

Public Sub MailIn()
    Dim s As String
    Dim Files As Variant, i As Long
    On Error Resume Next
    Do
        Files = SMail.Recv
        s = Bsprintf("Свои файлы (*.%s;*.txt),*.%s;*.txt", User.ID4, User.ID4) & _
            ",Курсы валют (*.dbf),*.dbf" & _
            ",Документы (*.doc и пр.),*.doc;*.rtf;*.xls;*.htm;*.gif;*.jpg" & _
            ",Ключи PGP (*.pgp),*.pgp" & _
            ",Реквизиты (*.id),*.id" & _
            ",Файлы загрузки (*.plt),*.plt" & _
            ",Обновления (*.exe),*.exe"
        If Not BrowseForFiles(Files, s, _
            "файл(ы) для просмотра") Then Exit Do
        For i = LBound(Files) To UBound(Files)
            MailOpenFile CStr(Files(i))
        Next
    Loop
End Sub

Public Sub MailOut()
    Dim File1 As String
    On Error Resume Next
    Do
        File1 = SMail.Send
        If Not BrowseForFile(File1, _
            Bsprintf("Свои файлы (*.%s;*.txt),*.%s;*.txt", User.ID4, User.ID4), _
            "файл для просмотра") Then Exit Do
        MailOpenFile File1
    Loop
End Sub

Public Sub MailArch()
    Dim s As String
    Dim Files As Variant, i As Long
    On Error Resume Next
    Do
        Files = SMail.Archive
        s = Bsprintf("Свои файлы (*.%s;*.txt),*.%s;*.txt", User.ID4, User.ID4)
        If Not BrowseForFiles(Files, s, _
            "файл(ы) для просмотра") Then Exit Do
        For i = LBound(Files) To UBound(Files)
            MailOpenFile CStr(Files(i))
        Next
    Loop
End Sub

Public Sub MailDump(File As String, Optional KillAfter As Boolean = False)
    Dim SS() As String, s As String, i As Long, ws As Worksheet
    On Error GoTo ErrSheet
    Set ws = Worksheets(MailSheet)
    Application.ScreenUpdating = False
    With ws.Columns("A:A")
        .Clear
        .Font.Name = "Courier New"
        .NumberFormat = "@"
    End With
    Application.ScreenUpdating = True
    
    On Error Resume Next
    s = InputFile(File)
    If Len(s) = 0 Then
        WarnBox "Файл пуст или не читается!"
        If IsFile(File) And KillAfter Then Kill File
        Exit Sub
    Else
        StrToLines CWin(s), SS
    End If
    If KillAfter Then Kill File
    
    On Error GoTo ErrSheet
    Set ws = Worksheets(MailSheet)
    Application.ScreenUpdating = False
    With ws.Columns("A:A")
        .Clear
        .Font.Name = "Courier New"
        .NumberFormat = "@"
    End With
    With ws
        For i = 1 To UBound(SS)
            .Cells(i, 1) = SS(i)
        Next
    End With
    Application.ScreenUpdating = True
    If App.BoolOptions("DontPreviewMail") Then
        Application.GoTo Worksheets(MailSheet).Range("$A$1"), True
    Else
        Worksheets(MailSheet).PrintPreview
    End If
    Exit Sub
    
ErrSheet:
    AutoRestart "Лист платежки не найден"
End Sub

Public Sub MailDumpDbf(File As String)
    Workbooks.Open FileName:=File
    With Range("A1")
        .CurrentRegion.Columns.AutoFit
        .EntireRow.Font.Bold = True
        .Select
    End With
End Sub

Public Sub MailClear()
    On Error Resume Next
    Application.ScreenUpdating = False
    With Worksheets(MailSheet).Columns("A:A")
        .Clear
        .Font.Name = "Courier New"
        .NumberFormat = "@"
    End With
    Application.ScreenUpdating = True
End Sub

Public Sub MailOpenFile(File As String)
    Dim s As String, File2 As String
    If User.IsID4(FileExt(File)) Then
        's = InputFile(File)
        'If Len(s) = 0 Then Exit Sub
        'If Left(s, Len(FormatPGP)) = FormatPGP Then
        '    File = PGP.DecryptEx(File)
        '    If Not IsFile(File) Then
        '        WarnBox Bsprintf("Программа PGP не смогла расшифровать файл!\nУ Вас есть ключи '%s'?", s), _
        '            vbExclamation, App.Title
        '        Exit Sub
        '    End If
        File2 = GetWinTempFile
        If Crypto.Decrypt(File, File2) Then
            MailDump File2, True
        Else
            MailDump File
        End If
        If IsFile(File2) Then Kill File2
    'ElseIf LCase(FileNameOnly(File)) = "remart" Then
    '    MailDump File
    Else
        s = LCase(FileExt(File))
        Select Case s
            Case "exe"
                UpdateReceived File
            Case "txt"
                MailDump File
            Case "dbf"
                MailDumpDbf File
            Case "pgp"
                InfoBox "Добавление производится автоматически"
            Case "id"
                InfoBox "Добавление производится автоматически"
            Case "plt"
                InfoBox "Добавление производится автоматически"
            Case "htm", "html", "xml", "gif", "jpg", "cer", "der"
                If OkCancelBox("Передача файла в Microsoft Explorer") Then
                    Shell "explorer.exe " & LFN(File), vbNormalFocus
                End If
            Case "doc", "rtf"
                If OkCancelBox("Передача файла в Microsoft Word") Then
                    Shell "winword.exe " & LFN(File), vbNormalFocus
                End If
            Case "xls"
                If OkCancelBox("Передача файла в Microsoft Excel") Then
                    Shell "excel.exe " & LFN(File), vbNormalFocus
                End If
            Case "chm"
                If OkCancelBox("Передача файла в просмотр") Then
                    Shell "hh.exe " & LFN(File), vbNormalFocus
                End If
            Case Else
                WarnBox "Неизвестный тип файла '%s'\nВозможно, Вам пора запросить обновление!", s
        End Select
    End If
End Sub
