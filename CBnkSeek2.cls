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
    Name = CWin(Trim(mRecord.Bank))
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
    ElseIf Left(mRecord.BIC, 6) = Left(App.DefBIC, 6) Then  '044030...
        Delivery = vbNullString
    Else
        Delivery = "Почтой"
    End If
End Property

Public Property Get Info() As String
    With Me
        Info = Bsprintf("%s, %s\nБИК %s, К/С %s", .Name, .Place, .BIC, .KS)
    End With
End Property

Public Sub MsgInfo(Optional BIC As String)
    If IsMissing(BIC) Then BIC = mRecord.BIC
    Me.BIC = BIC
    If mFound Then
        InfoBox "Справочник БИК для системы Банк-Клиент:\n(состояние на %n)\n\n%s", _
            mUpdated, Info
    Else
        If OkCancelBox("Банк c БИК %s в файле\n%s (обновлялся %n) не найден!\n\n" & _
            "Сейчас будет запрошено обновление справочника.\n" & _
            "Свяжитесь дважды с банком с небольшим интервалом\n" & _
            "(2-3 минуты) и Вы его получите.", _
            BIC, mFile, mUpdated) Then
            OutputFile SMail.Send & "bnkseek2.ask", _
                Bsprintf("ID:%s %s\n%s %n\n", User.ID, Now, mFile, mUpdated)
        End If
    End If
End Sub

Private Function SearchFile(s As String) As Boolean
    Dim nFile As Long, n As Long, p As Long, BIC As String * 9
    SearchFile = False
    If Not IsFile(mFile) Then Reset
    On Error GoTo ErrFile
    nFile = FreeFile
    Open mFile For Binary Access Read Shared As nFile
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
    WarnBox "Ошибка поиска %s в файле %s", s, mFile
    Close nFile
End Function

Public Property Get File() As String
    File = mFile
End Property

Public Property Let File(NewValue As String)
    mFile = NewValue
    If Not IsFile(mFile) Then LocateFile
    If Not LoadFile Then mFile = vbNullString
    App.Setting("BnkSeek2", "File") = mFile
End Property

Private Function LoadFile() As Boolean
    Dim nFile As Long
    LoadFile = False
    On Error GoTo ErrFile
    nFile = FreeFile
    Open mFile For Binary Access Read Shared As nFile
        ReInitHeader nFile
        On Error GoTo 0
    Close nFile
    LoadFile = mRecSize = ValidRecSize
    If Not LoadFile Then
        WarnBox "Справочник в файле %s неизвестной версии!", mFile
    End If
    Exit Function

ErrFile:
    WarnBox "Ошибка загрузки Справочника из файла %s", mFile
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
    mFile = App.Path & "BnkSeek2.dbf"
    If BrowseForFile(mFile, "Справочники (*.dbf),*.dbf", _
        Bsprintf("Справочник БИК (обычно %s)", mFile), True) Then
        App.Setting(BnkSeek2Section, "File") = mFile
    End If
    LocateFile = IsFile(mFile)
End Function

Public Property Get None() As String
    None = "Справочник БИК не установлен!"
End Property

Private Sub ReInitHeader(nFile As Long)
    Dim buf As String * 32, b() As Byte
    Get nFile, 1, buf
    StrToBytes buf, b
    If b(2) < 90 Then b(2) = b(2) + 100 'Y2K!
    mUpdated = DateSerial(1900 + b(2), b(3), b(4))
    mRecCount = BytesToValue(b, 5, 4)
    mDataOffset = BytesToValue(b, 9, 2) + 1
    mRecSize = BytesToValue(b, 11, 2)
End Sub
