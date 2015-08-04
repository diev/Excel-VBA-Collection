Attribute VB_Name = "IniFile"
Option Explicit
Option Compare Text
Option Base 1
DefLng A-Z

Private Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpSectionName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpSectionName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpSectionName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
'Private Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpSectionName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
'Private Declare Function WritePrivateProfileSection Lib "kernel32" Alias "WritePrivateProfileSectionA" (ByVal lpSectionName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
'Private Declare Function GetPrivateProfileSectionNames Lib "kernel32" Alias "GetPrivateProfileSectionNamesA" (ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
'Private Declare Function GetPrivateProfileStruct Lib "kernel32" Alias "GetPrivateProfileStructA" (ByVal lpSectionName As String, ByVal lpKeyName As Any, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Boolean
'Private Declare Function WritePrivateProfileStruct Lib "kernel32" Alias "WritePrivateProfileStructA" (ByVal lpSectionName As String, ByVal lpKeyName As Any, ByVal lpString As String, ByVal nSize As Long, ByVal lpFileName As String) As Boolean

Public Function NumIniFile(File As String, Section As String, Key As String, Optional Default As Long = 0) As Long
    NumIniFile = GetPrivateProfileInt(Section, Key, Default, File)
End Function

Public Function ReadIniFile(File As String, Section As String, Key As String, Optional Default As String = vbNullString) As String
    Dim Buf As String, i As Long
    Buf = Space(255)
    i = GetPrivateProfileString(Section, Key, Default, Buf, 255, File)
    ReadIniFile = Left(Buf, i)
End Function

Public Function EmptyIniFile(File As String, Section As String, Key As String) As Boolean
    Dim Buf As String, i As Long
    Buf = Space(255)
    i = GetPrivateProfileString(Section, Key, vbNullString, Buf, 255, File)
    EmptyIniFile = i = 0
End Function

Public Sub WriteIniFile(File As String, Section As String, Key As String, Optional vNewValue As String = vbNullString)
    WritePrivateProfileString Section, Key, vNewValue, File
End Sub
