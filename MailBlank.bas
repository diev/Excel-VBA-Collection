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
        s = Bsprintf("Свои файлы (*.%s;*.txt),*.%s;*.txt", User.ID, User.ID) & _
            ",Курсы валют (remart*.*),remart*.*" & _
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
            Bsprintf("Свои файлы (*.%s;*.txt),*.%s;*.txt", User.ID, User.ID), _
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
        s = Bsprintf("Свои файлы (*.%s;*.txt),*.%s;*.txt", User.ID, User.ID)
        If Not BrowseForFiles(Files, s, _
            "файл(ы) для просмотра") Then Exit Do
        For i = LBound(Files) To UBound(Files)
            MailOpenFile CStr(Files(i))
        Next
    Loop
End Sub

Public Sub MailDump(File As String, Optional KillAfter As Boolean = False)
    Dim ss() As String, s As String, i As Long, ws As Worksheet
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
        StrToLines CWin(s), ss
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
        For i = 1 To UBound(ss)
            .Cells(i, 1) = ss(i)
        Next
    End With
    Application.ScreenUpdating = True
    If StrToBool(App.Options("DontPreviewMail")) Then
        Application.GoTo Worksheets(MailSheet).Range("$A$1"), True
    Else
        Worksheets(MailSheet).PrintPreview
    End If
    Exit Sub
    
ErrSheet:
    AutoRestart "Лист платежки не найден"
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
    Dim s As String, ws As Worksheet
    s = LCase(FileExt(File))
    Select Case s
        Case User.ID
            PGP.ID = s
            'PGP.Password = vbNullString
            File = PGP.DecodeEx(File)
            If Not IsFile(File) Then
                WarnBox Bsprintf("Программа PGP не смогла расшифровать файл!\nВозможно, ошибка использования ключей"), _
                    vbExclamation, App.Title
                Exit Sub
            End If
            MailDump File, True
        Case "exe"
            UpdateReceived File
        Case "txt"
            MailDump File
        Case "doc"
            If OkCancelBox("Передача файла в Microsoft Word") Then
                Shell "winword.exe " & File, vbNormalFocus
            End If
        Case User.DemoID
            PGP.ID = s
            'PGP.Password = vbNullString
            File = PGP.DecodeEx(File)
            If Not IsFile(File) Then
                WarnBox Bsprintf("Программа PGP не смогла расшифровать файл!\nВозможно, ошибка использования ключей"), _
                    vbExclamation, App.Title
                Exit Sub
            End If
            MailDump File, True
        Case Else
            For Each ws In Worksheets
                If ws.Name = s Then
                    If IsFile(App.Path & s & ".id") Then
                        PGP.ID = s
                        'PGP.Password = vbNullString
                        File = PGP.DecodeEx(File)
                        If Not IsFile(File) Then
                            WarnBox Bsprintf("Программа PGP не смогла расшифровать файл!\nВозможно, ошибка использования ключей"), _
                                vbExclamation, App.Title
                            Exit Sub
                        End If
                        MailDump File, True
                        Exit Sub
                    End If
                End If
            Next
            MailDump File
    End Select
End Sub
