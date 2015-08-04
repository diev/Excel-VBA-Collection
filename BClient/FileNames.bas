Attribute VB_Name = "FileNames"
Option Explicit
Option Compare Text
Option Base 1
DefLng A-Z

Public Function FileExists(FileName As String) As Boolean
    FileExists = False
    If Len(FileName) = 0 Then Exit Function
    If Len(Dir(FileName)) = 0 Then Exit Function
    FileExists = True
End Function

Public Function FilePath(FileName As String) As String
    FilePath = Left(FileName, InStrR(FileName, "\"))
End Function

Public Function FileNameExt(FileName As String) As String
    FileNameExt = Mid(FileName, InStrR(FileName, "\") + 1)
End Function

Public Function FileExt(FileName As String) As String
    FileExt = Mid(FileName, InStrR(FileName, ".") + 1)
End Function

Public Function ChangeFileExt(FileName As String, FileExt As String) As String
    ChangeFileExt = Left(FileName, InStrR(FileName, ".")) & FileExt
End Function

Public Function RightSlash(FilePath As String) As String
    If Len(FilePath) = 0 Then
        RightSlash = ".\"
    ElseIf Right(FilePath, 1) = "\" Then
        RightSlash = FilePath
    Else
        RightSlash = FilePath & "\"
    End If
End Function

Public Function RightPathName(FilePath As String, FileName As String) As String
    RightPathName = RightSlash(FilePath) & FileName
End Function

Public Function TempFileName(Optional FileStart As String = "~temp") As String
    Dim n, s As String, FileName As String
    n = 1
    s = String(8 - Len(FileStart), "0")
    FileName = RightSlash(GetWinTempDir) & FileStart
    Do
        TempFileName = FileName & Format(n, s) & ".tmp"
        n = n + 1
    Loop While FileExists(TempFileName)
End Function
