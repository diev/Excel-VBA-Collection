VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Option Explicit
Option Compare Text
Option Base 1
DefLng A-Z

Const UserDemoID = "000"
Const UserIDExt = ".ID"

Dim mUserID As String
Dim mUserIDFile As String
Dim mLocateCanceled As Boolean

Private Sub Class_Initialize()
    mLocateCanceled = False
    If App.NoSetting(UserDemoID, "Name") Then WriteINI
    Me.ID = UserDemoID
End Sub

Private Sub Class_Terminate()
    '
End Sub

Public Property Get ID() As String
    ID = mUserID
End Property

Public Property Let ID(ByVal vNewValue As String)
    Dim s As String, v As Long, File As String
    On Error Resume Next
    s = Right(vNewValue, 7)
    v = Val(s)
    If v = 0 Then
        mUserID = UserDemoID
    Else
        mUserID = CStr(v)
    End If
    File = mUserID & UserIDExt
    mLocateCanceled = False
    If Me.Demo Then
        mUserIDFile = App.Ini
    Else
        mUserIDFile = App.Path & File
        If Not IsFile(mUserIDFile) Then
            File = mUserIDFile
            If BrowseForFile(File, "Файл реквизитов (*.id),*.id", "файл реквизитов") Then
                FileCopy File, mUserIDFile
            End If
        End If
        If EmptyIniFile(mUserIDFile, mUserID, "Name") Then
            If YesNoBox("Клиент %s не зарегистрирован!\nДобавить его?", mUserID) Then
                WriteIniFile mUserIDFile, mUserID, "Name", CDos("Новый клиент")
                NewUserShow
                Exit Property
            Else
                mUserID = UserDemoID
                mUserIDFile = App.Ini
            End If
        Else
            mUserIDFile = App.Path & File
        End If
    End If
    ResetCaption
    SetUserSheet
    'Commented due to an abnormous screen flashing!...
    'Payment.FillBlank
    'MailClear
End Property

Public Property Get Name() As String
    Name = CWin(ReadIniFile(mUserIDFile, mUserID, "Name"))
End Property

Public Property Let Name(ByVal vNewValue As String)
    WriteIniFile mUserIDFile, mUserID, "Name", CDos(vNewValue)
End Property

Public Property Get NameN(n As Long) As String
    NameN = CWin(ReadIniFile(mUserIDFile, mUserID, "Name" & CStr(n)))
End Property

Public Property Let NameN(n As Long, ByVal vNewValue As String)
    WriteIniFile mUserIDFile, mUserID, "Name" & CStr(n), CDos(vNewValue)
End Property

Public Property Get INN() As String
    INN = CWin(ReadIniFile(mUserIDFile, mUserID, "INN"))
End Property

Public Property Let INN(ByVal vNewValue As String)
    WriteIniFile mUserIDFile, mUserID, "INN", CDos(vNewValue)
End Property

Public Property Get LS() As String
    LS = ReadIniFile(mUserIDFile, mUserID, "LS", App.DefLS)
End Property

Public Property Let LS(ByVal vNewValue As String)
    WriteIniFile mUserIDFile, mUserID, "LS", vNewValue
End Property

Public Property Get BIC() As String
    BIC = ReadIniFile(mUserIDFile, mUserID, "BIC", App.DefBIC)
End Property

Public Property Let BIC(ByVal vNewValue As String)
    If ReadIniFile(mUserIDFile, mUserID, "BIC") <> vNewValue Then
        With BnkSeek2
            .BIC = vNewValue
            WriteIniFile mUserIDFile, mUserID, "BIC", .BIC
            WriteIniFile mUserIDFile, mUserID, "Bank", CDos(.Name)
            WriteIniFile mUserIDFile, mUserID, "Place", CDos(.Place)
            WriteIniFile mUserIDFile, mUserID, "KS", .KS
        End With
    End If
End Property

Public Property Get Bank() As String
    Bank = CWin(ReadIniFile(mUserIDFile, mUserID, "Bank"))
    If Bank = vbNullString Then
        With BnkSeek2
            .BIC = Me.BIC
            Bank = .Name
            Me.Bank = .Name
            Me.Place = .Place
            Me.KS = .KS
        End With
    End If
End Property

Public Property Let Bank(ByVal vNewValue As String)
    WriteIniFile mUserIDFile, mUserID, "Bank", CDos(vNewValue)
End Property

Public Property Get Place() As String
    Place = CWin(ReadIniFile(mUserIDFile, mUserID, "Place"))
End Property

Public Property Let Place(ByVal vNewValue As String)
    WriteIniFile mUserIDFile, mUserID, "Place", CDos(vNewValue)
End Property

Public Property Get KS() As String
    KS = ReadIniFile(mUserIDFile, mUserID, "KS")
End Property

Public Property Let KS(ByVal vNewValue As String)
    WriteIniFile mUserIDFile, mUserID, "KS", vNewValue
End Property

Public Property Get No() As Long
    No = NumIniFile(mUserIDFile, mUserID, "No", 1)
End Property

Public Property Let No(ByVal vNewValue As Long)
    If vNewValue > Me.NoMax Then
        WriteIniFile mUserIDFile, mUserID, "No", CStr(Me.NoMin)
    ElseIf vNewValue < Me.NoMin Then
        WriteIniFile mUserIDFile, mUserID, "No", CStr(Me.NoMin)
    Else
        WriteIniFile mUserIDFile, mUserID, "No", CStr(vNewValue)
    End If
End Property

Public Property Get NoMin() As Long
    NoMin = NumIniFile(mUserIDFile, mUserID, "NoMin", 1)
End Property

Public Property Let NoMin(ByVal vNewValue As Long)
    WriteIniFile mUserIDFile, mUserID, "NoMin", CStr(vNewValue)
    If Me.No < vNewValue Then Me.No = vNewValue
End Property

Public Property Get NoMax() As Long
    NoMax = NumIniFile(mUserIDFile, mUserID, "NoMax", 999)
End Property

Public Property Let NoMax(ByVal vNewValue As Long)
    WriteIniFile mUserIDFile, mUserID, "NoMax", CStr(vNewValue)
    If Me.No > vNewValue Then Me.No = vNewValue
End Property

Public Function SecRing(anID As String) As String
    Dim s As String
    If anID = UserDemoID Then
        s = App.Ini
    Else
        s = App.Path & anID & UserIDExt
    End If
    SecRing = RightPathName(ReadIniFile(s, anID, "KeysPath"), "SECR" & anID & ".PGP")
End Function

Public Function PubRing(anID As String) As String
    Dim s As String
    If anID = UserDemoID Then
        s = App.Ini
    Else
        s = App.Path & anID & UserIDExt
    End If
    PubRing = RightPathName(ReadIniFile(s, anID, "KeysPath"), "PUBR" & anID & ".PGP")
End Function

Public Property Get KeysPath() As String
    KeysPath = RightSlash(ReadIniFile(mUserIDFile, mUserID, "KeysPath"))
End Property

Public Property Let KeysPath(ByVal vNewValue As String)
    WriteIniFile mUserIDFile, mUserID, "KeysPath", vNewValue
End Property

Public Property Get ImportList() As String
    ImportList = ReadIniFile(mUserIDFile, mUserID, "ImportList", App.Path)
End Property

Public Property Let ImportList(ByVal vNewValue As String)
    WriteIniFile mUserIDFile, mUserID, "ImportList", vNewValue
End Property

Public Property Get ExportList() As String
    ExportList = ReadIniFile(mUserIDFile, mUserID, "ExportList", App.Path)
End Property

Public Property Let ExportList(ByVal vNewValue As String)
    WriteIniFile mUserIDFile, mUserID, "ExportList", vNewValue
End Property

Public Property Get Sign(Number As Long) As String
    Sign = CWin(ReadIniFile(mUserIDFile, mUserID, "Sign" & CStr(Number)))
End Property

Public Property Let Sign(Number As Long, ByVal vNewValue As String)
    WriteIniFile mUserIDFile, mUserID, "Sign" & CStr(Number), CDos(vNewValue)
End Property

Public Property Get Tel() As String
    Tel = CWin(ReadIniFile(mUserIDFile, mUserID, "Tel"))
End Property

Public Property Let Tel(ByVal vNewValue As String)
    WriteIniFile mUserIDFile, mUserID, "Tel", CDos(vNewValue)
End Property

Public Property Get Demo() As Boolean
    Demo = mUserID = UserDemoID
End Property

Public Property Let Demo(ByVal vNewValue As Boolean)
    If vNewValue Then Me.ID = UserDemoID
End Property

Public Property Get Valid() As Boolean
    Valid = IsFile(mUserIDFile) And Not EmptyIniFile(mUserIDFile, mUserID, "Name")
End Property

Public Property Get DemoID() As String
    DemoID = UserDemoID
End Property

Public Property Get File() As String
    File = mUserIDFile
End Property

Public Property Let File(ByVal vNewValue As String)
    mUserIDFile = vNewValue
End Property

Public Sub WriteINI()
    With App
        .Setting(UserDemoID, "INN") = "7800000000"
        .Setting(UserDemoID, "Name") = CDos("ЗАО ""Демо""")
        .Setting(UserDemoID, "LS") = App.DefLS
        .Setting(UserDemoID, "BIC") = App.DefBIC
        .Setting(UserDemoID, "Sign1") = CDos("Директор")
        .Setting(UserDemoID, "Sign2") = CDos("Бухгалтер")
        .Setting(UserDemoID, "Tel") = "?"
        .Setting(UserDemoID, "No") = 1
        .Setting(UserDemoID, "NoMin") = 1
        .Setting(UserDemoID, "NoMax") = 999
        .Setting(UserDemoID, "Amount") = "0.00"
        .Setting(UserDemoID, "KeysPath") = App.Path & "PGP\"
        .Setting(UserDemoID, "ExportList") = App.Path
        .Setting(UserDemoID, "ImportList") = App.Path
    End With
End Sub

Public Function LocateKeys() As Boolean
    Dim File As String, File1 As String
    LocateKeys = False
    If mLocateCanceled Then Exit Function
    File1 = "SECR" & mUserID & ".PGP"
    File = PathDirectories("Keys;PGP;PGP\Keys", File1)
    'If Not mLocateCanceled Then
        
    If Not IsFile(File) Then
        If YesNoBox("Ключи PGP не найдены!\nПоискать ключи на дискете?") Then
            File = PathDirectories("A:\Keys;A:\PGP;A:\PGP\Keys", File1)
        End If
    End If
    Do
        File = App.Path & File1
        If BrowseForFile(File, "Все ключи PGP (*.pgp),(*.pgp)", _
            Bsprintf("файл Secr%s.pgp из Вашей пары ключей!", mUserID), True) Then
            File1 = mUserID & ".pgp"
            If LCase(Right(File, Len(File1))) = File1 Then
                Me.KeysPath = FilePath(File)
                LocateKeys = True
                Exit Function
            Else
                StopBox "Указан чужой комплект ключей!"
            End If
        Else
            mLocateCanceled = True
            Exit Do
        End If
    Loop
End Function

Public Sub SetUserSheet()
    Dim ws As Worksheet
    On Error Resume Next
    For Each ws In Worksheets
        If ws.Name = mUserID Then
            ws.Activate
            Exit Sub
        End If
    Next
    For Each ws In Worksheets
        If ws.Name = UserDemoID Then
            ws.Copy before:=Sheets(1)
            Sheets(1).Name = mUserID
            'ActiveWindow.Caption = mUserID
            Exit Sub
        End If
    Next
    WarnBox "Какая-то ошибка установки листа!"
End Sub

Public Sub Delete()
    Dim DelUserID As String, DelUserFile As String, DelSecRing As String, DelPubRing As String
    On Error Resume Next
    With Me
        DelUserID = .ID
        DelUserFile = .File
        DelSecRing = .SecRing(.ID)
        DelPubRing = .PubRing(.ID)
        .Demo = True
    End With
    Worksheets(DelUserID).Delete
    Kill DelUserFile
    If IsFile(DelUserFile) Then _
        WarnBox "Файл настроек %s остался!", DelUserFile
    Kill DelSecRing
    If IsFile(DelSecRing) Then _
        WarnBox "Файл ключей %s остался!", DelSecRing
    Kill DelPubRing
    If IsFile(DelPubRing) Then _
        WarnBox "Файл ключей %s остался!", DelPubRing
End Sub

Public Property Get Amount() As Currency
    Amount = RVal(ReadIniFile(mUserIDFile, mUserID, "Amount", "0.00"))
End Property

Public Property Let Amount(ByVal vNewValue As Currency)
    WriteIniFile mUserIDFile, mUserID, "Amount", PlatFormat(vNewValue)
    ResetCaption
End Property

Public Sub AmountMinus(ByVal vMinusValue As Currency)
    With Me
        .Amount = .Amount - vMinusValue
        ResetCaption
    End With
End Sub

Public Sub ResetCaption()
    With Me
        Application.Caption = Bsprintf("Банк-Клиент %s - %s (%f)", .ID, .Name, .Amount)
    End With
End Sub

Public Function LocateID(Optional Path As String = vbNullString) As Boolean
    On Error Resume Next
    LocateID = False
    If Len(Path) = 0 Then Path = App.Path
    If BrowseForFile(Path, "Файлы реквизитов (*.id),*.id", _
        "файл ID с реквизитами, если он есть", True) Then
        FileCopy Path, App.Path & FileNameExt(Path)
        ID = FileNameOnly(Path)
    End If
    LocateID = IsFile(mUserIDFile)
End Function

