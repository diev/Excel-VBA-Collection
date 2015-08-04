Attribute VB_Name = "MiscFiles"
Option Explicit
Option Compare Text
Option Base 1
DefLng A-Z

Public Function BrowseForFile(ByRef File As String, Mask As String, _
    Optional Capt As String = "файл", Optional Force As Boolean = False) As Boolean
    Dim f As Variant, s1 As String, s2 As String
    s1 = FileNameExt(File)
    If Not Force Then
        If IsFile(File) Then
            BrowseForFile = True
            Exit Function
        End If
    End If
    On Error Resume Next
    With Application
        'set by default
        .DefaultFilePath = App.Path
        ChDrive .DefaultFilePath
        ChDir .DefaultFilePath
    
        .DefaultFilePath = FilePath(File) 'Some problems on some computers
        ChDrive .DefaultFilePath
        ChDir .DefaultFilePath
    End With
    Do
        f = Application.GetOpenFilename(Mask & _
            ",Все файлы (*.*),*.*", 1, _
            "Укажите " & Capt)
        If f <> False Then 'don't change this!
            File = CStr(f)
        Else
            BrowseForFile = False
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
End Function

Public Function BrowseForFiles(ByRef Files As Variant, Mask As String, _
    Optional Capt As String = "файл(ы)") As Boolean
    Dim f As Variant
    On Error Resume Next
    With Application
        .DefaultFilePath = FilePath(CStr(Files)) 'Some problems on some computers
        ChDrive .DefaultFilePath
        ChDir .DefaultFilePath
        f = .GetOpenFilename(Mask & _
            ",Все файлы (*.*),*.*", 1, _
            "Укажите " & Capt, , True)
    End With
    If f <> False Then 'don't change this!
        BrowseForFiles = True
        Files = f
    Else
        BrowseForFiles = False
    End If
End Function

Public Function BrowseForSave(ByRef File As String, Mask As String, _
    Optional Capt As String = "файл") As Boolean
    Dim f As Variant
    BrowseForSave = False
    On Error Resume Next
    With Application
        .DefaultFilePath = FilePath(File) 'Some problems on some computers
        ChDrive .DefaultFilePath
        ChDir .DefaultFilePath
        f = .GetSaveAsFilename(File, Mask & _
            ",Все файлы (*.*),*.*", 1, _
            "Укажите " & Capt)
    End With
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

