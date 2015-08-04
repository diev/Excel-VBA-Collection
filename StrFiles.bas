Attribute VB_Name = "StrFiles"
Option Explicit
Option Compare Text
Option Base 1
DefLng A-Z

Const ExtChar = "."
Const PathChar = "\"
Const UNCChar = "\\"
Const SepChar = ";"

Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
'Declare Function GetFullPathName Lib "kernel32" Alias "GetFullPathNameA" (ByVal lpFileName As String, ByVal nBufferLength As Long, ByVal lpBuffer As String, ByVal lpFilePart As String) As Long

'Windows
Public Function GetWinDir() As String
    Dim s As String, n
    n = 255
    s = Space(n)
    n = GetWindowsDirectory(s, n)
    GetWinDir = RightSlash(Left(s, n))
End Function

Public Function GetWinSysDir() As String
    Dim s As String, n
    n = 255
    s = Space(n)
    n = GetSystemDirectory(s, n)
    GetWinSysDir = RightSlash(Left(s, n))
End Function

Public Function GetWinTempDir() As String
    Dim s As String, n
    n = 255
    s = Space(n)
    n = GetTempPath(n, s)
    GetWinTempDir = RightSlash(Left(s, n))
End Function

Public Function GetWinTempFile(Optional Prefix As String = "TMP") As String
    Dim s As String, n
    n = 255: s = Space(n)
    n = GetTempPath(n, s)
    GetTempFileName Left(s, n), Prefix, 0, s
    GetWinTempFile = Left(s, InStr(s, vbNullChar) - 1)
End Function

'Files
Public Function FullFile(File As String) As String
    Dim Disk As String, Path As String, CurPath As String
    If Len(File) = 0 Then
        FullFile = vbNullString
        Exit Function
    End If
    CurPath = ActiveWorkbook.FullName
    CurPath = Left(CurPath, InStrR(CurPath, PathChar) - 1) 'instead CurDir!
    If Mid(File, 2, 1) = ":" Then
        Disk = Left(File, 2)
        Path = Mid(File, 3)
    ElseIf Left(File, 2) = UNCChar Then
        Disk = vbNullString
        Path = File
    Else 'ignore CurDir
        'Disk = Left(CurDir, 2)
        Disk = Left(CurPath, 2)
        Path = File
    End If
    If Left(Path, 2) = ".\" Then Path = Mid(Path, 3)
    Do While Left(Path, 3) = "..\"
        Disk = Left(Disk, InStrR(Disk, PathChar) - 1)
        Path = Mid(Path, 4)
    Loop
    If Left(Path, 1) = PathChar Then
        FullFile = Disk & Path
    Else
        'FullFile = RightPathName(CurDir(Disk), Path)
        FullFile = RightPathName(CurPath, Path)
    End If
End Function

Public Function FilePath(File As String) As String
    Dim n As Long
    File = FullFile(File)
    If IsDir(File) Then
        FilePath = RightSlash(File)
    Else
        n = InStrR(File, PathChar)
        If n > 0 Then
            FilePath = Left(File, n)
        Else
            FilePath = RightSlash(CurDir)
        End If
    End If
End Function

Public Function FileNameExt(File As String) As String
    Dim n As Long
    n = InStrR(File, PathChar)
    If n = Len(File) Then
        FileNameExt = vbNullString
    ElseIf n > 0 Then
        FileNameExt = Mid(File, n + 1)
    Else
        FileNameExt = File
    End If
End Function

Public Function FileNameOnly(File As String) As String
    Dim n1 As Long, n2 As Long
    If Right(File, 1) = PathChar Then File = Left(File, Len(File) - 1)
    n1 = InStrR(File, PathChar)
    n2 = InStrR(File, ExtChar)
    If n1 > n2 Then
        FileNameOnly = Mid(File, n1 + 1)
    ElseIf n2 > 0 Then
        FileNameOnly = Mid(File, n1 + 1, n2 - n1 - 1)
    Else
        FileNameOnly = Mid(FileNameOnly, n1 + 1)
    End If
End Function

Public Function FileExt(File As String) As String
    Dim n1 As Long, n2 As Long
    If Right(File, 1) = PathChar Then File = Left(File, Len(File) - 1)
    n1 = InStrR(File, PathChar)
    n2 = InStrR(File, ExtChar)
    If n1 > n2 Then
        FileExt = vbNullString
    ElseIf n2 > 0 Then
        FileExt = Mid(File, n2 + 1)
    Else
        FileExt = vbNullString
    End If
End Function

Public Function ChangeFileExt(File As String, FileExt As String) As String
    'No error checking!
    If Len(File) = 0 Then
        ChangeFileExt = vbNullString
    Else
        ChangeFileExt = FullFile(Left(File, InStrR(File, ExtChar)) & FileExt)
    End If
End Function

Public Function RightSlash(FilePath As String) As String
    If Len(FilePath) = 0 Then FilePath = CurDir
    If Right(FilePath, 1) = PathChar Then
        RightSlash = FilePath
    Else
        RightSlash = FilePath & PathChar
    End If
End Function

Public Function RightPathName(FilePath As String, File As String) As String
    RightPathName = RightSlash(FilePath) & File
End Function

'Quote long filenames with spaces in
Public Function QFile(File As String) As String
    If InStr(File, " ") > 0 And Left(File, 1) <> """" Then
        QFile = """" & File & """"
    Else
        QFile = File
    End If
End Function

'Create all dirs in the path
Public Sub ForceDirectories(Path As String)
    Dim Arr As Variant, i As Long, s As String
    If IsDir(Path) Then Exit Sub
    On Error Resume Next
    Path = FullFile(Path)
    Arr = StrToArr(Path, PathChar)
    s = Left(Path, InStr(Path, PathChar) - 1)
    For i = 2 To UBound(Arr)
        s = s & PathChar & CStr(Arr(i))
        MkDir s
    Next
End Sub

'Search for the file a-la PATH
Public Function PathDirectories(Path As String, File As String) As String
    Dim Arr As Variant, i As Long, s As String
    On Error Resume Next
    s = ActiveWorkbook.FullName
    s = Left(s, InStrR(s, PathChar) - 1)
    Path = Path & SepChar & s '& SepChar & Environ("PATH")
    Arr = StrToArr(Path, SepChar)
    For i = 1 To UBound(Arr)
        s = FullFile(RightPathName(CStr(Arr(i)), File))
        If IsFile(s) Then
            PathDirectories = FullFile(s)
            Exit Function
        End If
    Next
    PathDirectories = vbNullString
End Function

Public Function CountFiles(Mask As String) As Long
    Dim File As String
    File = Dir(Mask)
    CountFiles = 0
    Do While File <> vbNullString
        CountFiles = CountFiles + 1
        File = Dir
    Loop
End Function
