VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CBnkSeek2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Option Base 1
DefLng A-Z

'BnkSeek2.dbf
Private Type BnkSeek2Type
    Deleted As String * 1 'DBF flag
    BIC As String * 9
    BANK As String * 45
    Place As String * 31
    KS As String * 20
    POST As String * 1
End Type
Const ValidRecSize As Long = 107
Const mBICOffset As Long = 1

'Const BnkSeek2Section = "BnkSeek2"

Dim mFile As String
Dim mUpdated As Date
Dim mRecCount As Long
Dim mDataOffset As Long
Dim mRecSize As Long
Dim mRecord As BnkSeek2Type
Dim mFound As Boolean

Private Sub Class_Initialize()
    mFile = App.Setting(BnkSeek2Section, "File")
    If Not IsFile(mFile) Then Me.File = PathDirectories(vbNullString, "BnkSeek2.dbf")
    'Me.BIC = app.defbic
End Sub

Public Property Get BIC() As String
    BIC = mRecord.BIC
End Property

Public Property Let BIC(ByVal NewValue As String)
    If Not Me.Valid Then Reset
    NewValue = PadL(NewValue, 9, "0")
    If (mRecord.BIC <> NewValue) Then mFound = SearchFile(NewValue)
End Property

Public Property Get Name() As String
    Name = CWin(Trim(mRecord.BANK))
End Property

Public Property Get Place() As String
    Place = CWin(Trim(mRecord.Place))
End Property

Public Property Get KS() As String
    KS = mRecord.KS
End Property

Public Property Get NamePost() As String
    NamePost = mRecord.POST
End Property

Public Property Get Delivery() As String
    If mRecord.POST = "C" Then
        Delivery = "Электронно"
    Else
        Delivery = "Почтой"
    End If
End Property

Public Property Get Info() As String
    With Me
        Info = BPrintF("%s, %s\nБИК %s, К/С %s", .Name, .Place, .BIC, .KS)
    End With
End Property

Public Sub MsgInfo(Optional BIC As String)
    If IsMissing(BIC) Then BIC = mRecord.BIC
    Me.BIC = BIC
    If mFound Then
        MsgBox BPrintF("Справочник банков РФ:\n\n") & Info, vbInformation, App.Title
    Else
        MsgBox BPrintF("Банк c БИК %s в файле \'%s\' не найден!", BIC, mFile), vbExclamation, App.Title
    End If
End Sub

Private Function SearchFile(s As String) As Boolean
    Dim nFile As Long, n As Long, p As Long, BIC As String * 9
    SearchFile = False
    If Not IsFile(mFile) Then Reset
    On Error GoTo ErrFile
    nFile = FreeFile
    Open mFile For Binary Access Read Shared As nFile
        'Seek nFile, mDataOffset
        'Do While Not EOF(nFile)
        '    'RecNo = (Seek(nFile) - mDataOffset) \ mRecSize
        '    Get nFile, , mRecord
        '    If mRecord.NEWNUM = s Then
        '        SearchFile = True
        '        Exit Do
        '    End If
        'Loop
        p = mDataOffset + mBICOffset
        Application.Cursor = xlWait
        For n = 1 To mRecCount
            Get nFile, p, BIC
            If BIC = s Then
                Get nFile, p - mBICOffset, mRecord
                SearchFile = True
                Exit For
            End If
            p = p + mRecSize
        Next
        Application.Cursor = xlDefault
        Close nFile
    Exit Function
    
ErrFile:
    Application.Cursor = xlDefault
    ErrorBox Err
    Close nFile
End Function

Public Property Get File() As String
    File = mFile
End Property

Public Property Let File(NewValue As String)
    mFile = NewValue
    If Not IsFile(mFile) Then LocateFile
    If Not LoadFile Then mFile = vbNullString
End Property

Private Function LoadFile() As Boolean
    Dim nFile As Long, buf As String * 32, b() As Byte
    LoadFile = False
    On Error GoTo ErrFile
    nFile = FreeFile
    Open mFile For Binary Access Read Shared As nFile
        Get nFile, , buf
        On Error GoTo 0
    Close nFile
    StrToBytes buf, b
    If b(2) < 90 Then b(2) = b(2) + 100 'Y2K!
    mUpdated = DateSerial(1900 + b(2), b(3), b(4))
    mRecCount = BytesToValue(b, 5, 4)
    mDataOffset = BytesToValue(b, 9, 2) + 1
    mRecSize = BytesToValue(b, 11, 2)
    LoadFile = mRecSize = ValidRecSize
    If Not LoadFile Then
        MsgBox "Справочник банков РФ не версии 2!", vbExclamation, App.Title
    End If
    Exit Function

ErrFile:
    ErrorBox Err, "Ошибка загрузки BnkSeek2"
    Close nFile
End Function

Public Property Get Valid() As Boolean
    Valid = IsFile(mFile)
End Property

Public Property Get Found() As Boolean
    Found = mFound
End Property

Public Property Get Updated() As Variant
    Updated = mUpdated
End Property

Public Function LocateFile() As Boolean
    mFile = App.Path
    If BrowseForFile(mFile, "Справочники (*.dbf),*.dbf", "Справочник банков РФ (BnkSeek2.dbf)", True) Then
        App.Setting(BnkSeek2Section, "File") = FullFile(mFile)
    End If
    LocateFile = IsFile(mFile)
End Function

Public Property Get None() As String
    None = "Справочник банков РФ не установлен!"
End Property

