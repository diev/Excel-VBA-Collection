Attribute VB_Name = "IniFile"
Option Explicit
Option Compare Text
Option Base 1
DefLng A-Z

Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpSectionName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpSectionName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpSectionName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpSectionName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileSection Lib "kernel32" Alias "WritePrivateProfileSectionA" (ByVal lpSectionName As String, ByVal lpString As String, ByVal lpFileName As String) As Long

Public Function GetINIInt(FileName As String, Section As String, Key As String, Optional Default As Long = 0) As Long
    GetINIInt = GetPrivateProfileInt(Section, Key, Default, FileName)
End Function

Public Function GetINIString(FileName As String, Section As String, Key As String, Optional Default As String = "") As String
    Dim s As String, n
    s = Space(255)
    n = GetPrivateProfileString(Section, Key, Default, s, 255, FileName)
    GetINIString = Left(s, n)
End Function

Public Sub WriteINIString(FileName As String, Section As String, Key As String, Optional Value As String = "")
    WritePrivateProfileString Section, Key, Value, FileName
End Sub

Public Function GetINISection(FileName As String, Section As String) As Variant
    Dim s As String, n, arr As Variant, Arr2 As Variant, lb, ub, i
    n = FileLen(FileName)
    s = Space(n)
    GetPrivateProfileSection Section, s, n, FileName
    n = InStr(1, s, Chr(0) & Chr(0))
    s = Left(s, n)
    arr = StrToArr(s, Chr(0), False)
    lb = LBound(arr)
    ub = UBound(arr)
    ReDim Arr2(lb To ub, 2) As String
    For i = lb To ub
        s = arr(i)
        n = InStr(2, s, "=")
        Arr2(i, 1) = Trim(Left(s, n - 1))
        Arr2(i, 2) = Trim(Mid(s, n + 1))
    Next
    GetINISection = Arr2
End Function

Public Sub WriteINISection(FileName As String, Section As String, arr As Variant)
    Dim s As String, lb, ub, i
    s = ""
    lb = LBound(arr)
    ub = UBound(arr)
    For i = lb To ub
        s = s & arr(i, 1) & "=" & arr(i, 2) & Chr(0)
    Next
    s = s & Chr(0)
    WritePrivateProfileSection Section, s, FileName
End Sub
