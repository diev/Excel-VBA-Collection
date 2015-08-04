Attribute VB_Name = "ExportInfoRep"
Option Explicit
Option Compare Text
Option Base 1
DefLng A-Z

Public Sub SendInfo()
    Dim n As Long, s As String, i As Long
    On Error Resume Next
    n = FreeFile
    Open SMail.Send & "InfoRep.txt" For Output Access Write Shared As #n
        Print #n, "[App]"
        Print #n, "Title=" & App.TITLE
        Print #n, "Now=" & Now
        Print #n, "CP=1251"
        Print #n,
        Print #n, "Path=" & App.Path
        Print #n, "XLA=" & ThisWorkbook.FullName
        Print #n, "Version=" & App.Version
        Print #n,
        With Application
            s = Choose(Val(.Version) - 6, " 95", " 97", " 2000", " XP", " XP+")
            Print #n, "Excel=" & .Version & s
            Print #n, "OS=" & .OperatingSystem
            Print #n, "Printer=" & .ActivePrinter
        End With
        Print #n,
        With BnkSeek2
            Print #n, "BnkSeek=" & .File
            Print #n, "Updated=" & DtoC(.Updated)
        End With
        Print #n,
        Print #n, "XLS=" & ActiveWorkbook.FullName
        For i = 1 To Worksheets.Count
            With Worksheets(i)
                If IsNumeric(.Name) Then
                    Print #n, Bsprintf("Sheet%d=%s, rows: %d", _
                        i, .Name, .Cells(1, 1).CurrentRegion.Rows.Count)
                Else
                    Print #n, Bsprintf("Sheet%d=%s", _
                        i, .Name)
                End If
            End With
        Next
        Print #n,
        s = Dir(App.Path & "*.id")
        Do While s <> vbNullString
            s = InputFile(App.Path & s)
            Print #n, s
            s = Dir
        Loop
        Print #n,
        PrintSubDirs n, App.Path
        Print #n, "; eof"
    Close #n
End Sub

Private Sub PrintSubDirs(n As Long, Path As String)
    Dim s As String, i As Long, T As Long
    Dim Dirs As New Collection, x As Variant
    Path = RightSlash(Path)
    Print #n, "[" & Path & "]"
    s = Dir(Path & "*.*", vbDirectory)
    Do While s <> vbNullString
        If s <> "." And s <> ".." Then
            If (GetAttr(Path & s) And vbDirectory) = vbDirectory Then
                Print #n, DirString(Path & s)
                Dirs.Add s
            End If
        End If
        s = Dir
    Loop
    i = 0: T = 0
    s = Dir(Path & "*.*")
    Do While s <> vbNullString
        i = i + 1
        T = T + FileLen(Path & s)
        Print #n, FileString(Path, s)
        s = Dir
    Loop
    If i > 0 Then
        Print #n, "Total: " & (T \ 1024) & "K, files: " & i
    End If
    Print #n,
    For Each x In Dirs
        PrintSubDirs n, Path & CStr(x)
    Next
End Sub

Private Function DirString(Path As String) As String
    DirString = " " & Path
End Function

Private Function FileString(Path As String, File As String) As String
    FileString = Format(FileDateTime(Path & File), "dd.MM.yyyy HH:mm") & _
        PadL(CStr(FileLen(Path & File)), 10) & "  " & File
End Function

