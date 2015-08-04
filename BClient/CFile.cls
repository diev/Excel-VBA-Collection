VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Option Base 1
DefLng A-Z

Private mFile As String
Private mFiles As Variant
Private mFilesIndex As Long
Private mDOS As Boolean

Private Sub Class_Initialize()
    mFile = vbNullString
    mDOS = False
End Sub

Private Sub Class_Terminate()

End Sub

Public Function ChangeExt(ByVal vNewValue As String) As String
    ChangeExt = StrFiles.ChangeFileExt(mFile, vNewValue)
End Function

Public Function CopyTo(ByVal vNewValue As String) As Boolean
    On Error Resume Next
    If MiscFiles.IsDir(vNewValue) Then
        vNewValue = RightPathName(vNewValue, StrFiles.FileNameExt)
    End If
    FileCopy mFile, vNewValue
    CopyTo = MiscFiles.IsFile(vNewValue)
End Function

Public Function CopyToBAK() As Boolean
    Dim s As String
    On Error Resume Next
    s = StrFiles.ChangeFileExt(mFile, "BAK")
    FileCopy mFile, s
    CopyToBAK = MiscFiles.IsFile(s)
End Function

Public Function MoveTo(ByVal vNewValue As String) As Boolean
    On Error Resume Next
    If MiscFiles.IsDir(vNewValue) Then
        vNewValue = RightPathName(vNewValue, StrFiles.FileNameExt)
    End If
    FileCopy mFile, vNewValue
    If MiscFiles.IsFile(vNewValue) Then
        MoveTo = Me.Delete
        mFile = vNewValue
    Else
        MoveTo = False
    End If
End Function

Public Function RenameTo(vNewValue As String) As Boolean
    On Error Resume Next
    Name mFile As vNewValue
    If MiscFiles.IsFile(vNewValue) Then
        RenameTo = True
        mFile = vNewValue
    Else
        RenameTo = False
    End If
End Function

Public Property Get Text() As String
    Text = InputFile(mFile)
    If mDOS Then Text = CWin(Text)
End Property

Public Property Let Text(ByVal vNewValue As String)
    If mDOS Then vNewValue = CDos(vNewValue)
    OutputFile mFile, vNewValue
End Property

Public Sub AddText(vNewValue As String)
    If mDOS Then vNewValue = CDos(vNewValue)
    AppendFile mFile, vNewValue
End Sub

Public Property Get File() As String
    File = mFile
End Property

Public Property Let File(ByVal vNewValue As String)
    mFile = vNewValue
End Property

Public Property Get Path() As String
    Path = StrFiles.FilePath(mFile)
End Property

Public Property Get FileName() As String
    FileName = StrFiles.FileNameOnly(mFile)
End Property

Public Property Get FileNameExt() As String
    FileNameExt = StrFiles.FileNameExt(mFile)
End Property

Public Property Get FileExt() As String
    FileExt = StrFiles.FileExt(mFile)
End Property

Public Property Get FileDate() As Variant
    FileDate = VBA.FileDateTime(mFile)
End Property

Public Property Get Length() As Long
    Length = VBA.FileLen(mFile)
End Property

Public Function Exists() As Boolean
    Exists = MiscFiles.IsFile(mFile)
End Function

Public Function IsDir() As Boolean
    IsDir = MiscFiles.IsDir(mFile)
End Function

Public Function SearchPath(Path As String, File As String) As Boolean
    mFile = PathDirectories(Path, File)
    SearchPath = MiscFiles.IsFile(mFile)
End Function

Public Property Get Strings() As Variant
    Dim ss As New CStrings
    ss.Load mFile, mDOS
    Strings = ss.List
End Property

Public Property Let Strings(ByVal vNewValue As Variant)
    Dim ss As New CStrings
    ss.List = vNewValue
    ss.Save mFile, mDOS
End Property

Public Property Get DOS() As Boolean
    DOS = mDOS
End Property

Public Property Let DOS(ByVal vNewValue As Boolean)
    mDOS = vNewValue
End Property

Public Function Delete() As Boolean
    On Error Resume Next
    Kill mFile
    Delete = MiscFiles.IsFile(mFile)
    mFile = vbNullString
End Function

Public Property Get Valid() As Boolean
    Valid = MiscFiles.IsFile(mFile)
End Property

Public Function FindCount(Optional Mask As String = "*.*") As Long
    Dim s As String
    FindCount = 0
    s = Dir(Mask)
    Do While s <> vbNullString
        FindCount = FindCount + 1
        s = Dir
    Loop
End Function

Public Function FindFirst(Optional Mask As String = "*.*") As Boolean
    'mFile = Dir(Mask)
    'FindFirst = Len(mFile) > 0
    Dim s As String, ss As String
    s = Dir(Mask)
    If s <> vbNullString Then
        ss = s
        Do While s <> vbNullString
            s = Dir
            ss = ss & ";" & s
        Loop
        mFiles = StrToArr(ss, ";")
        mFilesIndex = 1
        mFile = mFiles(mFilesIndex)
    Else
        mFile = vbNullString
        mFiles = Nothing
        mFilesIndex = 0
    End If
End Function

Public Function FindNext() As Boolean
    'mFile = Dir
    'FindNext = Len(mFile) > 0
    On Error GoTo FindNone
    mFilesIndex = mFilesIndex + 1
    mFile = mFiles(mFilesIndex)
    FindNext = True
    Exit Function
FindNone:
    mFile = vbNullString
    mFiles = Nothing
    mFilesIndex = 0
    FindNext = False
End Function

Public Function FindFiles(Optional Mask As String = "*.*") As Variant
    FindFiles = mFiles
End Function
