VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CIniFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Option Base 1
DefLng A-Z

Const Comment = ";#/"

Private Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpSectionName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpSectionName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpSectionName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpSectionName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileSection Lib "kernel32" Alias "WritePrivateProfileSectionA" (ByVal lpSectionName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileSectionNames Lib "kernel32" Alias "GetPrivateProfileSectionNamesA" (ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileStruct Lib "kernel32" Alias "GetPrivateProfileStructA" (ByVal lpSectionName As String, ByVal lpKeyName As Any, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Boolean
Private Declare Function WritePrivateProfileStruct Lib "kernel32" Alias "WritePrivateProfileStructA" (ByVal lpSectionName As String, ByVal lpKeyName As Any, ByVal lpString As String, ByVal nSize As Long, ByVal lpFileName As String) As Boolean

Dim mFile As String
Dim mDefaultSection As String
Dim mReadOnly As Boolean

Private Sub Class_Initialize()
    Reset
End Sub

Private Sub Class_Terminate()
    FlushToDisk
End Sub

Public Sub Reset()
    mReadOnly = False
    mFile = ActiveWorkbook.FullName
    Me.File = ChangeFileExt(mFile, "INI")
End Sub

Public Property Get File() As String
    File = mFile
End Property

Public Property Let File(vNewValue As String)
    mFile = vNewValue
    mDefaultSection = vbNullString
End Property

Public Property Get DefaultSection() As String
    DefaultSection = mDefaultSection
End Property

Public Property Let DefaultSection(ByVal vNewValue As String)
    mDefaultSection = vNewValue
End Property

Public Function GetKeyValue(SectionKey As String, Optional Default As Long = 0) As Long
    Dim i As Long, s As String, k As String: i = InStr(1, SectionKey, "\")
    If i = 0 Then
        s = mDefaultSection: k = SectionKey
    Else
        s = Left(SectionKey, i - 1): k = Mid(SectionKey, i + 1)
    End If
    GetKeyValue = GetPrivateProfileInt(s, k, Default, mFile)
End Function

Public Function GetKey(SectionKey As String, Optional Default As String = vbNullString) As String
    Dim buf As String, i As Long, s As String, k As String
    buf = Space(255): i = InStr(1, SectionKey, "\")
    If i = 0 Then
        s = mDefaultSection: k = SectionKey
    Else
        s = Left(SectionKey, i - 1): k = Mid(SectionKey, i + 1)
    End If
    i = GetPrivateProfileString(s, k, CDos(Default), buf, 255, mFile)
    GetKey = CWin(Left(buf, i))
End Function

Public Property Get KeyEmpty(SectionKey As String) As Boolean
    Dim buf As String, i As Long, s As String, k As String
    buf = Space(255): i = InStr(1, SectionKey, "\")
    If i = 0 Then
        s = mDefaultSection: k = SectionKey
    Else
        s = Left(SectionKey, i - 1): k = Mid(SectionKey, i + 1)
    End If
    i = GetPrivateProfileString(s, k, vbNullString, buf, 255, mFile)
    KeyEmpty = i = 0
End Property

Public Property Get KeyValue(SectionKey As String) As Long
    Dim i As Long, s As String, k As String: i = InStr(1, SectionKey, "\")
    If i = 0 Then
        s = mDefaultSection: k = SectionKey
    Else
        s = Left(SectionKey, i - 1): k = Mid(SectionKey, i + 1)
    End If
    KeyValue = GetPrivateProfileInt(s, k, 0, mFile)
End Property

Public Property Get Key(SectionKey As String) As String
    Dim buf As String, i As Long, s As String, k As String
    buf = Space(255): i = InStr(1, SectionKey, "\")
    If i = 0 Then
        s = mDefaultSection: k = SectionKey
    Else
        s = Left(SectionKey, i - 1): k = Mid(SectionKey, i + 1)
    End If
    i = GetPrivateProfileString(s, k, vbNullString, buf, 255, mFile)
    Key = CWin(Left(buf, i))
End Property

Public Property Let Key(SectionKey As String, vNewValue As String)
    Dim i As Long, s As String, k As String
    If mReadOnly Then Exit Property
    i = InStr(1, SectionKey, "\")
    If i = 0 Then
        s = mDefaultSection: k = SectionKey
    Else
        s = Left(SectionKey, i - 1): k = Mid(SectionKey, i + 1)
    End If
    WritePrivateProfileString s, k, CDos(vNewValue), mFile
End Property

Public Property Get Sections() As Variant
    Dim buf As String, n As Long
    n = FileLen(mFile): buf = Space(n)
    GetPrivateProfileSectionNames buf, n, mFile
    n = InStr(1, buf, vbNullChar & vbNullChar): buf = Left(buf, n)
    Sections = StrToArr(CWin(buf), vbNullChar, False)
End Property

Public Property Get List(SectionName As String) As Variant
    Dim buf As String, n As Long
    n = FileLen(mFile): buf = Space(n)
    GetPrivateProfileSection SectionName, buf, n, mFile
    n = InStr(1, buf, vbNullChar & vbNullChar): buf = Left(buf, n)
    List = StrToArr(CWin(buf), vbNullChar, False)
End Property

Public Property Let List(SectionName As String, vNewValue As Variant)
    Dim buf As String
    If mReadOnly Then Exit Property
    buf = ArrToStr(vNewValue, vbNullChar, False) & vbNullChar
    WritePrivateProfileSection SectionName, CDos(buf), mFile
End Property

Public Sub FlushToDisk()
    On Error Resume Next
    If mReadOnly Then Exit Sub
    WritePrivateProfileStruct vbNullChar, vbNullChar, vbNullChar, 0, mFile
End Sub

Public Property Get KeyBoolean(SectionKey As String) As Boolean
    KeyBoolean = StrToBool(Me.Key(SectionKey))
End Property

Public Property Let KeyBoolean(SectionKey As String, NewValue As Boolean)
    If NewValue Then
        Me.Key(SectionKey) = "1"
    Else
        Me.Key(SectionKey) = "0"
    End If
End Property

Public Property Get Text(SectionName As String) As String
    Dim buf As String, n As Long
    n = FileLen(mFile): buf = Space(n)
    GetPrivateProfileSection SectionName, buf, n, mFile
    n = InStr(1, buf, vbNullChar & vbNullChar): buf = Left(buf, n)
    Text = StrSpaces1(StrTran(CWin(buf), vbNullChar, " "))
End Property

Public Property Let Text(SectionName As String, ByVal NewValue As String)
    If mReadOnly Then Exit Property
    If Right(NewValue, 1) <> vbCrLf Then NewValue = NewValue & vbCrLf
    NewValue = StrTran(NewValue & vbCrLf, vbCrLf, vbNullChar)
    WritePrivateProfileSection SectionName, CDos(NewValue), mFile
End Property

Public Property Get ReadOnly() As Boolean
    ReadOnly = mReadOnly
End Property

Public Property Let ReadOnly(ByVal vNewValue As Boolean)
    mReadOnly = vNewValue
End Property
