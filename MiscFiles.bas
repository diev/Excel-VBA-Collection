Attribute VB_Name = "MiscFiles"
Option Explicit
Option Compare Text
Option Base 1
DefLng A-Z

Const DefaultMask = "Все файлы (*.*),*.*"

Public Function BrowseForFile(ByRef File As String, Optional Mask As String = vbNullString, _
    Optional Capt As String = "файл", Optional Force As Boolean = False) As Boolean
    Dim f As Variant, s1 As String, s2 As String, Home As String
    s1 = FileNameExt(File)
    If Not Force Then
        If IsFile(File) Then
            BrowseForFile = True
            Exit Function
        End If
    End If
    On Error Resume Next
    Home = CurDir
    If Len(Mask) = 0 Then
        Mask = DefaultMask
    Else
        Mask = Mask & "," & DefaultMask
    End If
    With Application
        'set by default
        .DefaultFilePath = Application.Path
        ChDrive .DefaultFilePath
        ChDir .DefaultFilePath
    
        .DefaultFilePath = FilePath(File) 'Some problems on some computers
        ChDrive .DefaultFilePath
        ChDir .DefaultFilePath
    End With
    Do
        f = Application.GetOpenFilename(Mask, 1, "Укажите " & Capt)
        If f <> False Then 'don't change this!
            File = CStr(f)
        Else
            BrowseForFile = False
            ChDrive Home
            ChDir Home
            Exit Function
        End If
        s2 = FileNameExt(File)
        
        If Not Force Then Exit Do
        If UCase(s1) = UCase(s2) Then Exit Do
        If YesNoBox("ВНИМАНИЕ! Возможно, Вы указали не тот файл,\n" & _
            "который ждет от Вас программа:\n\n%s\n(вместо ожидаемого %s)\n\n" & _
            "Все равно использовать этот файл?", File, s1) Then Exit Do
    Loop
    BrowseForFile = IsFile(File)
    ChDrive Home
    ChDir Home
End Function

Public Function BrowseForFiles(ByRef Files As Variant, Optional Mask As String = vbNullString, _
    Optional Capt As String = "файл(ы)", Optional FilterIndex As Long = 1) As Boolean
    Dim f As Variant, Home As String
    On Error Resume Next
    Home = CurDir
    If Len(Mask) = 0 Then
        Mask = DefaultMask
    Else
        Mask = Mask & "," & DefaultMask
    End If
    With Application
        .DefaultFilePath = FilePath(CStr(Files)) 'Some problems on some computers
        ChDrive .DefaultFilePath
        ChDir .DefaultFilePath
        f = .GetOpenFilename(Mask, FilterIndex, "Укажите " & Capt, , True)
    End With
    If f <> False Then 'don't change this!
        BrowseForFiles = True
        Files = f
    Else
        BrowseForFiles = False
    End If
    ChDrive Home
    ChDir Home
End Function

Public Function BrowseForSave(ByRef File As String, Optional Mask As String = vbNullString, _
    Optional Capt As String = "файл") As Boolean
    Dim f As Variant, Home As String
    BrowseForSave = False
    On Error Resume Next
    Home = CurDir
    If Len(Mask) = 0 Then
        Mask = DefaultMask
    Else
        Mask = Mask & "," & DefaultMask
    End If
    With Application
        .DefaultFilePath = FilePath(File) 'Some problems on some computers
        ChDrive .DefaultFilePath
        ChDir .DefaultFilePath
        f = .GetSaveAsFilename(File, Mask, 1, "Укажите " & Capt)
    End With
    ChDrive Home
    ChDir Home
    If f = False Then Exit Function
    File = CStr(f)
    BrowseForSave = True
End Function

Public Function IsFile(File As String) As Boolean
    IsFile = False
    On Error GoTo ErrDir
    IsFile = GetAttr(File)
    IsFile = Not IsDir(File)
ErrDir:
End Function

Public Function IsFile1(RunCmd As String) As Boolean
    Dim File As String, Arr As Variant
    IsFile1 = False
    On Error GoTo ErrDir
    Arr = StrToArr(RunCmd)
    File = Arr(1)
    IsFile1 = GetAttr(File)
    IsFile1 = Not IsDir(File)
ErrDir:
End Function

Public Function IsDir(File As String) As Boolean
    IsDir = False
    On Error GoTo ErrDir
    IsDir = GetAttr(File) And vbDirectory
ErrDir:
End Function

