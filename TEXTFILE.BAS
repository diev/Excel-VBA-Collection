Attribute VB_Name = "TextFile"
Option Explicit
Option Compare Text
Option Base 1
DefLng A-Z

Public Sub ErrorBox(e As ErrObject, Optional Description As Variant, Optional TITLE As Variant)
    If IsMissing(Description) Then Description = e.Description _
        Else Description = BPrintF("%s\n\%s", Description, e.Description)
    If IsMissing(TITLE) Then TITLE = BPrintF("Ошибка %d в %s", e.Number, e.Source)
    MsgBox Description, vbExclamation, TITLE
End Sub

Public Sub DumpFile()
    Dim s As String, Buf As String, p, d, r
    
    Buf = Cells(1, 1).Text
    If Len(Buf) = 0 Then _
        Buf = Application.GetOpenFilename("Файлы TXT (*.txt),*.txt,Все файлы (*.*),*.*", 1, "Открытие файла TXT")
    
    s = InputFile(Buf)
    If Len(s) = 0 Then Exit Sub
    s = CWin(s) & vbLf
    
    Columns("A:A").Clear
    Cells(1, 1) = Buf
    Application.ScreenUpdating = False
    p = 1: r = 2
    Do
        d = InStr(p, s, vbCrLf) 'DOS eoln
        If d = 0 Then d = InStr(p, s, vbLf) 'else UNIX eoln
        If d = 0 Then Exit Do 'no eoln is eof
        If d = p Then Exit Do
        Buf = Mid(s, p, d - p)
        Cells(r, 1) = "'" & Buf
        p = d + 2
        r = r + 1
    Loop
    
    s = "$A$2:$A$" & Trim(CStr(r - 1))
    With Range(s)
        .Font.Name = "Courier New"
        '.NumberFormat = "@"
        .Interior.ColorIndex = 35
        .Interior.Pattern = xlSolid
    End With
    ActiveSheet.PageSetup.PrintArea = s
    
    Columns("A:A").AutoFit
    Application.ScreenUpdating = True
End Sub

Public Function InputFile(FName As String) As String
    Dim FNum As Long
    On Error GoTo ErrHandler
    FNum = FreeFile
    Open FName For Binary Access Read Lock Write As FNum
        InputFile = Input(LOF(FNum), #FNum)
    Close #FNum
    On Error GoTo 0
    Exit Function
ErrHandler:
    Close #FNum
    InputFile = ""
    ErrorBox Err, BPrintF("Чтение файла ~%s~", FName)
End Function

Public Sub OutputFile(FName As String, s As String)
    Dim FNum As Long
    On Error GoTo ErrHandler
    FNum = FreeFile
    Open FName For Output Access Write Lock Write As FNum
    Close #FNum
    FNum = FreeFile
    Open FName For Binary Access Write Lock Write As FNum
        Put #FNum, , s
    Close #FNum
    On Error GoTo 0
    Exit Sub
ErrHandler:
    Close #FNum
    ErrorBox Err, BPrintF("Запись файла ~%s~", FName)
End Sub

Public Sub AppendFile(FName As String, s As String)
    Dim FNum As Long, p As Long
    On Error GoTo ErrHandler
    FNum = FreeFile
    Open FName For Binary Access Write Lock Write As FNum
        p = LOF(FNum)
        If p = 0 Then p = 1
        Put #FNum, p, s
    Close #FNum
    On Error GoTo 0
    Exit Sub
ErrHandler:
    Close #FNum
    ErrorBox Err, BPrintF("Добавление файла ~%s~", FName)
End Sub

Public Sub WipeFile(FName As String)
    Dim FNum As Long
    On Error GoTo ErrHandler
    FNum = FreeFile
    Open FName For Output Access Write Lock Write As FNum
        Print #FNum, String(LOF(FNum) + 4096, "*")
    Close #FNum
    Kill FName
    On Error GoTo 0
    Exit Sub
ErrHandler:
    Close #FNum
    ErrorBox Err, BPrintF("Затирание файла ~%s~", FName)
End Sub
