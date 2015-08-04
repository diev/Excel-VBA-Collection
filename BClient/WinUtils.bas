Attribute VB_Name = "WinUtils"
Option Explicit
Option Compare Text
Option Base 1
DefLng A-Z

Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Public Function GetWinDir() As String
    Dim s As String, n
    n = 255
    s = Space(n)
    n = GetWindowsDirectory(s, n)
    GetWinDir = Left(s, n)
End Function

Public Function GetWinSysDir() As String
    Dim s As String, n
    n = 255
    s = Space(n)
    n = GetSystemDirectory(s, n)
    GetWinSysDir = Left(s, n)
End Function

Public Function GetWinTempDir() As String
    Dim s As String, n
    n = 255
    s = Space(n)
    n = GetTempPath(n, s)
    GetWinTempDir = Left(s, n)
End Function
