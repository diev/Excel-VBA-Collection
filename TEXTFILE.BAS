Attribute VB_Name = "TextFile"
Option Explicit
Option Compare Text
Option Base 1
DefLng A-Z

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
    Open FName For Binary Access Read Shared As FNum
        InputFile = Input(LOF(FNum), #FNum)
    Close #FNum
    On Error GoTo 0
    Exit Function
ErrHandler:
    Close #FNum
    InputFile = vbNullString
    WarnBox "Ошибка чтения из файла\n%s", FName
End Function

Public Sub OutputFile(FName As String, s As String)
    Dim FNum As Long
    On Error Resume Next
    'Kill FName
    FNum = FreeFile
    Open FName For Output Access Write Lock Write As FNum
    Close #FNum
    FNum = FreeFile
    On Error GoTo ErrHandler
    Open FName For Binary Access Write Lock Write As FNum
        Put #FNum, , s
    Close #FNum
    On Error GoTo 0
    Exit Sub
ErrHandler:
    Close #FNum
    WarnBox "Ошибка записи в файл\n%s", FName
End Sub

Public Sub AppendFile(FName As String, s As String)
    Dim FNum As Long, p As Long
    On Error GoTo ErrHandler
    FNum = FreeFile
    Open FName For Binary Access Write Shared As FNum
        p = LOF(FNum) + 1
        Put #FNum, p, s
    Close #FNum
    On Error GoTo 0
    Exit Sub
ErrHandler:
    Close #FNum
    WarnBox "Ошибка добавления в файл\n%s", FName
End Sub

Public Sub PatchFile(FName As String, Addr As String, Codes As String)
    Dim FNum As Long, p As Long, i As Long, n As Long
    On Error GoTo ErrHandler
    FNum = FreeFile
    Open FName For Binary Access Read Write Lock Write As FNum
        p = CLng("&H" & Addr) + 1
        For i = 1 To Len(Codes) Step 2
            Put #FNum, p, CByte("&H" & Mid(Codes, i, 2))
            p = p + 1
        Next
    Close #FNum
    On Error GoTo 0
    Exit Sub
ErrHandler:
    Close #FNum
    WarnBox "Ошибка замены в файле\n%s\nпо адресу %s", FName, Addr
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
    WarnBox "Ошибка затирания файла\n%s", FName
End Sub

Public Function OverwriteFile(FName As String) As Boolean
    If IsFile(FName) Then
        OverwriteFile = YesNoBox("Файл %s\nуже существует с %T!\n\nПерезаписать его?", _
            FName, FileDateTime(FName))
    Else
        OverwriteFile = True
    End If
End Function
