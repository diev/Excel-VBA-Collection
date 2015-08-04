Attribute VB_Name = "MiscFiles"
Option Explicit
Option Compare Text
Option Base 1
DefLng A-Z

Public Function BrowseForFile(ByRef File As String, Mask As String, _
    Optional Capt As String = "файл", Optional Force As Boolean = False) As Boolean
    Dim f As Variant
    If Not Force Then
        If IsFile(File) Then
            BrowseForFile = True
            Exit Function
        End If
    End If
    On Error Resume Next
    'ChDrive File
    'If IsDir(File) Then
    '    ChDir File
    'Else
    '    ChDir FilePath(File)
    'End If
    With Application
        .DefaultFilePath = FilePath(File) 'Some problems on some computers
        ChDrive .DefaultFilePath
        ChDir .DefaultFilePath
        f = .GetOpenFilename(Mask & _
            ",Текстовые файлы (*.txt),*.txt,Все файлы (*.*),*.*", 1, _
            "Укажите " & Capt)
    End With
    If f <> False Then 'don't change this!
        File = CStr(f)
    End If
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
            ",Текстовые файлы (*.txt),*.txt,Все файлы (*.*),*.*", 1, _
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
            ",Текстовые файлы (*.txt),*.txt,Все файлы (*.*),*.*", 1, _
            "Укажите " & Capt)
    End With
    If f = False Then Exit Function
    File = CStr(f)
    BrowseForSave = True
End Function

Public Function IsFile(File As String) As Boolean
    IsFile = False
    On Error GoTo ErrDir
    'If Len(File) = 0 Then Exit Function
    'If Len(Dir(File)) = 0 Then Exit Function
    'IsFile = True
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

