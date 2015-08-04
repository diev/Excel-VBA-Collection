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
    Bank As String * 45
    Place As String * 31
    KS As String * 20
    POST As String * 1
End Type
Const ValidRecSize As Long = 107
Const mBICOffset As Long = 1

'Const BnkSeek2Section = "BnkSeek2"

Dim mUpdated As Date
Dim mRecCount As Long
Dim mDataOffset As Long
Dim mRecSize As Long
Dim mRecord As BnkSeek2Type
Dim mFound As Boolean

Public Property Get BIC() As String
    If mFound Then
        BIC = mRecord.BIC
    Else
        BIC = vbNullString
    End If
End Property

Public Property Let BIC(ByVal NewValue As String)
    If Len(NewValue) = 8 Then
        NewValue = "0" & NewValue
    End If
    If Len(NewValue) = 9 Then
    'If (mRecord.BIC <> NewValue) Then 'Caching last
        mFound = SearchFile(NewValue)
    Else
        mFound = False
    End If
End Property

Public Property Get Name() As String
    If mFound Then
        Name = CWin(Trim(mRecord.Bank))
    Else
        Name = vbNullString
    End If
End Property

Public Property Get Place() As String
    If mFound Then
        Place = CWin(Trim(mRecord.Place))
    Else
        Place = vbNullString
    End If
End Property

Public Property Get KS() As String
    If mFound Then
        KS = Trim(mRecord.KS)
    Else
        KS = vbNullString
    End If
End Property

Public Property Get BankInfo() As String
    With Me
        .LoadFile
        BankInfo = Bsprintf("%s, %s\nБИК %s, К/С %s", .Name, .Place, .BIC, .KS)
    End With
End Property

Public Sub MsgInfo(Optional BIC As String)
    If IsMissing(BIC) Then BIC = mRecord.BIC
    Me.BIC = BIC
    If mFound Then
        InfoBox "Справочник БИК для системы Банк-Клиент:\n(состояние на %n)\n\n%s", _
            mUpdated, BankInfo
    Else
        If OkCancelBox("Банк c БИК %s в файле\n%s (обновлялся %n) не найден!\n\n" & _
            "Сейчас будет запрошено обновление справочника.", _
            BIC, Me.File, mUpdated) Then
            AskBnkSeek
        End If
    End If
End Sub

Private Function SearchFile(s As String) As Boolean
    Dim nFile As Long, n As Long, p As Long, BIC As String * 9
    SearchFile = False
    On Error GoTo ErrFile
    nFile = FreeFile
    Open Me.File For Binary Access Read Shared As nFile
        ReInitHeader nFile
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
    WarnBox "Ошибка поиска %s в файле %s", s, Me.File
    Close nFile
End Function

Public Property Get File() As String
    File = App.Path & "BnkSeek2.dbf"
End Property

Public Function LoadFile() As Boolean
    Dim nFile As Long
    LoadFile = False
    On Error GoTo ErrFile
    nFile = FreeFile
    Open Me.File For Binary Access Read Shared As nFile
        ReInitHeader nFile
        On Error GoTo 0
    Close nFile
    LoadFile = mRecSize = ValidRecSize
    If Not LoadFile Then
        WarnBox "Справочник в файле %s неизвестной версии!", Me.File
    End If
    Exit Function

ErrFile:
    WarnBox "Ошибка загрузки Справочника из файла %s", Me.File
    Close nFile
End Function

Public Property Get Found() As Boolean
    Found = mFound
End Property

Public Property Get Updated() As Variant
    Me.LoadFile
    Updated = mUpdated
End Property

Public Property Get RecCount() As Long
    RecCount = mRecCount
End Property

Private Sub ReInitHeader(nFile As Long)
    Dim Buf As String * 32, b() As Byte
    Get nFile, 1, Buf
    StrToBytes Buf, b
    If b(2) < 90 Then b(2) = b(2) + 100 'Y2K!
    mUpdated = DateSerial(1900 + b(2), b(3), b(4))
    mRecCount = BytesToValue(b, 5, 4)
    mDataOffset = BytesToValue(b, 9, 2) + 1
    mRecSize = BytesToValue(b, 11, 2)
End Sub

