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
        s = BPrintF("Свои файлы (*.%s;*.txt),*.%s;*.txt", User.ID, User.ID) & _
            ",Курсы валют (remart*.*),remart*.*"
        If Not BrowseForFiles(Files, s, _
            "файл(ы) для просмотра") Then Exit Do
        For i = LBound(Files) To UBound(Files)
            s = CStr(Files(i))
            If FileExt(s) = User.ID Then
                MailDump PGP.DecodeEx(s)
            Else
                MailDump s
            End If
        Next
    Loop
End Sub

Public Sub MailOut()
    Dim File1 As String
    On Error Resume Next
    Do
        File1 = SMail.Send
        If Not BrowseForFile(File1, _
            BPrintF("Свои файлы (*.%s;*.txt),*.%s;*.txt", User.ID, User.ID), _
            "файл для просмотра") Then Exit Do
        If FileExt(File1) = User.ID Then
            MailDump PGP.DecodeEx(File1)
        Else
            MailDump File1
        End If
    Loop
End Sub

Public Sub MailArch()
    Dim s As String
    Dim Files As Variant, i As Long
    On Error Resume Next
    Do
        Files = SMail.Archive
        s = BPrintF("Свои файлы (*.%s;*.txt),*.%s;*.txt", User.ID, User.ID)
        If Not BrowseForFiles(Files, s, _
            "файл(ы) для просмотра") Then Exit Do
        For i = LBound(Files) To UBound(Files)
            s = CStr(Files(i))
            If FileExt(s) = User.ID Then
                MailDump PGP.DecodeEx(s)
            Else
                MailDump s
            End If
        Next
    Loop
End Sub

Public Sub MailDump(File As String)
    Dim ss As Variant, s As String, i As Long, ws As Worksheet
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
        MsgBox "Файл пуст или не читается!", vbExclamation, App.Title
        If IsFile(File) Then Kill File
        Exit Sub
    Else
        ss = StrToArr(CWin(s), vbCrLf)
    End If
    Kill File
    
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
    's = "$A$1:$A$" & Trim(Str(s.Count))
    'With Range(s)
    '    .Font.Name = "Courier New"
    '    '.NumberFormat = "@"
    '    .Interior.ColorIndex = 35
    '    .Interior.Pattern = xlSolid
    'End With
    'ActiveSheet.PageSetup.PrintArea = s
    
    'ws.Columns("A:A").AutoFit
    'With r.Range("A1", r.Cells(s.LineCount, 1)).Interior
    '    .ColorIndex = 2
    '    .Pattern = xlSolid
    'End With
    'r.Cells(1, 1).Select
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


