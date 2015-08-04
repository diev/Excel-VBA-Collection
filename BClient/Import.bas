Attribute VB_Name = "Import"
Option Explicit
Option Compare Text
Option Base 1
DefLng A-Z

Public Sub SMailSend()
    Dim FileName As String, TempName As String, s As String
    On Error Resume Next
    FileName = FileEnterByMask(RightPathName(Range("Out").Text, "*.*"))
    If FileExists(FileName) Then
        Select Case UCase(FileExt(FileName))
            Case Range("ID").Text
                PGPRun Range("PGPView").Text, FileName
            Case Else
                'If FileLen(FileName) > 32000 Then
                '    If MsgBox("Файл превышает 32K! Продолжать?", vbExclamation + vbOKCancel, AppTitle) = vbOK Then
                        s = InputFile(FileName)
                        TempName = RightPathName(GetWinTempDir, FileNameExt(FileName))
                        s = CWin(s)
                        OutputFile TempName, s
                        s = RightPathName(GetWinDir, "NOTEPAD ") & TempName
                        If Shell(s, vbNormalFocus) = 0 Then
                            ErrorBox Err, BPrintF("Невозможно запустить\n~%s~", s), AppTitle
                        End If
                        DoEvents
                        Kill TempName
                '    End If
                'End If
        End Select
    End If
End Sub

Public Sub SMailRecv()
    Dim FileName As String, TempName As String, s As String
    On Error Resume Next
    SMailAutoMark
    Do While True
        FileName = FileEnterByMask(RightPathName(Range("In").Text, "*.*"))
        If FileName = "" Then
            Exit Sub
        ElseIf FileExists(FileName) Then
            Select Case UCase(FileExt(FileName))
                Case Range("ID").Text, "OK", "ERR"
                    s = FileNameExt(FileName)
                    Archive.Mark(Left(s, 8)) = FileExt(s)
                    'Decode to a disk file
                    'TempName = RightPathName(GetWinTempDir, FileNameExt(FileName))
                    'PGPRun GetSetting(AppName, "PGP", "Decode"), FileName, TempName
                    's = InputFile(TempName)
                    'OutputFile TempName, CWin(s)
                    's = GetSetting(AppName, "Misc", "Notepad")
                    's = StrTran(s, "%1", TempName)
                    'If Shell(s, vbNormalFocus) = 0 Then
                    '    ErrorBox Err, BPrintF("Невозможно запустить\n~%s~", s), AppTitle
                    'End If
                    'DoEvents
                    
                    'View Only
                    PGPRun Range("PGPView").Text, FileName
                Case "EX_"
                    MsgBox BPrintF("Файл обновления:\nНеобходимо выйти и запустить UPDATE.BAT"), vbInformation, AppTitle
                Case Else
                    'If FileLen(FileName) > 32000 Then
                    '    If MsgBox("Файл превышает 32K! Продолжать?", vbExclamation + vbOKCancel, AppTitle) = vbOK Then
                            s = InputFile(FileName)
                            TempName = RightPathName(GetWinTempDir, FileNameExt(FileName))
                            s = CWin(s)
                            OutputFile TempName, s
                            s = RightPathName(GetWinDir, "NOTEPAD ") & TempName
                            If Shell(s, vbNormalFocus) = 0 Then
                                ErrorBox Err, BPrintF("Невозможно запустить\n~%s~", s), AppTitle
                            End If
                            DoEvents
                            Kill TempName
                    '    End If
                    'End If
            End Select
        End If
    Loop
End Sub

Public Sub SMailAutoMark()
    Dim FileName As String, s As String
    On Error Resume Next
    FileName = Dir(RightPathName(Range("In").Text, "*.*"))
    Do While FileName <> ""
        Select Case UCase(FileExt(FileName))
            Case Range("ID").Text, "OK", "ERR"
                s = FileNameExt(FileName)
                Archive.Mark(Left(s, 8)) = FileExt(s)
        End Select
        FileName = Dir
    Loop
End Sub
