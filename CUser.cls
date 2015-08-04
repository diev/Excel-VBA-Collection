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

Dim mUserID As String
Dim mUserIDFile As String
Dim mLocateCanceled As Boolean

Private Sub Class_Initialize()
    mLocateCanceled = False
    If App.Setting(UserDemoID, "Name") = vbNullString Then WriteINI
    Me.ID = UserDemoID
End Sub

Private Sub Class_Terminate()

End Sub

Public Property Get ID() As String
    ID = mUserID
End Property

Public Property Let ID(ByVal vNewValue As String)
    Dim s As String
    On Error Resume Next
    mUserID = vNewValue
    mLocateCanceled = False
    If Me.Demo Then
        mUserIDFile = App.Ini
    Else
        mUserIDFile = App.Path & mUserID & ".ID"
        If Not IsFile(mUserIDFile) Then
            If MsgBox(BPrintF("Файл реквизитов этого клиента не найден!\nПоискать на дискете?"), _
                vbQuestion + vbOKCancel, App.Title) = vbOK Then
                If IsFile("A:\" & mUserID & ".ID") Then
                    FileCopy "A:\" & mUserID & ".ID", mUserIDFile
                Else
                    MsgBox "На дискете нет файла реквизитов", vbExclamation, App.Title
                End If
            End If
        End If
        If ReadIniFile(mUserIDFile, mUserID, "Name") = vbNullString Then
            If MsgBox(BPrintF("Клиент не зарегистрирован!\nДобавить его?"), _
                vbQuestion + vbYesNo, App.Title) = vbYes Then
                mUserIDFile = App.Path & mUserID & ".ID"
                WriteIniFile mUserIDFile, mUserID, "Name" = "Новый клиент"
            Else
                Me.Demo = True
                mUserIDFile = App.Ini
            End If
        Else
            mUserIDFile = App.Path & mUserID & ".ID"
        End If
    End If
    ResetCaption
    SetUserSheet
    Payment.FillBlank
    MailClear
End Property

Public Property Get Name() As String
    Name = CWin(ReadIniFile(mUserIDFile, mUserID, "Name"))
End Property

Public Property Let Name(ByVal vNewValue As String)
    WriteIniFile mUserIDFile, mUserID, "Name", CDos(vNewValue)
End Property

Public Property Get INN() As String
    INN = ReadIniFile(mUserIDFile, mUserID, "INN")
End Property

Public Property Let INN(ByVal vNewValue As String)
    WriteIniFile mUserIDFile, mUserID, "INN", vNewValue
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

Public Property Get BANK() As String
    BANK = CWin(ReadIniFile(mUserIDFile, mUserID, "Bank"))
    If BANK = vbNullString Then
        With BnkSeek2
            .BIC = Me.BIC
            BANK = .Name
            Me.BANK = .Name
            Me.Place = .Place
            Me.KS = .KS
        End With
    End If
End Property

Public Property Let BANK(ByVal vNewValue As String)
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
    No = CLng(ReadIniFile(mUserIDFile, mUserID, "No", "1"))
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
    NoMin = CLng(ReadIniFile(mUserIDFile, mUserID, "NoMin", "1"))
End Property

Public Property Let NoMin(ByVal vNewValue As Long)
    WriteIniFile mUserIDFile, mUserID, "NoMin", CStr(vNewValue)
    If Me.No < vNewValue Then Me.No = vNewValue
End Property

Public Property Get NoMax() As Long
    NoMax = CLng(ReadIniFile(mUserIDFile, mUserID, "NoMax", "999"))
End Property

Public Property Let NoMax(ByVal vNewValue As Long)
    WriteIniFile mUserIDFile, mUserID, "NoMax", CStr(vNewValue)
    If Me.No > vNewValue Then Me.No = vNewValue
End Property

Public Property Get SecRing() As String
    SecRing = Me.KeysPath & "SECR" & mUserID & ".PGP"
    If Not IsFile(SecRing) Then
        If LocateKeys Then
            SecRing = Me.KeysPath & "SECR" & mUserID & ".PGP"
        Else
            SecRing = vbNullString
        End If
    End If
End Property

Public Property Get PubRing() As String
    PubRing = Me.KeysPath & "PUBR" & mUserID & ".PGP"
    If Not IsFile(PubRing) Then
        If LocateKeys Then
            PubRing = Me.KeysPath & "PUBR" & mUserID & ".PGP"
        Else
            PubRing = vbNullString
        End If
    End If
End Property

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
    Valid = IsFile(mFile)
    If Valid Then Valid = ReadIniFile(mUserIDFile, mUserID, "Name") <> vbNullString
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
        .Setting(UserDemoID, "No") = "1"
        .Setting(UserDemoID, "NoMin") = "1"
        .Setting(UserDemoID, "NoMax") = "999"
        .Setting(UserDemoID, "Amount") = "1000000.00"
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
    File = PathDirectories("PGP;Keys;\Keys;\PGP", File1)
    'If Not mLocateCanceled Then
        
    If Not IsFile(File) Then
        If MsgBox(BPrintF("Ключи PGP не найдены!\nПоискать ключи на дискете?"), _
            vbQuestion + vbOKCancel, App.Title) = vbOK Then
            File = PathDirectories("A:\Keys;A:\PGP;A:\", File1)
        End If
    End If
    Do
        File = App.Path
        If BrowseForFile(File, "Ключи PGP (*.pgp),(*.pgp)", _
            BPrintF("файл из Вашей пары ключей (с номером \'%s\'!)", mUserID), True) Then
            If Right(File, 7) = mUserID & ".PGP" Then
                Me.KeysPath = FilePath(File)
                LocateKeys = True
                Exit Function
            Else
                MsgBox "Указан чужой комплект ключей!", vbCritical, App.Title
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
    MsgBox "Какая-то ошибка установки листа!", vbExclamation, App.Title
End Sub

Public Sub Delete()
    Dim DelUserID As String, DelUserFile As String, DelSecRing As String, DelPubRing As String
    On Error Resume Next
    With Me
        DelUserID = .ID
        DelUserFile = .File
        DelSecRing = .SecRing
        DelPubRing = .PubRing
        .Demo = True
    End With
    Worksheets(DelUserID).Delete
    Kill DelUserFile
    If IsFile(DelUserFile) Then _
        MsgBox BPrintF("Файл настроек \'%s\' остался!", DelUserFile), vbExclamation, App.Title
    Kill DelSecRing
    If IsFile(DelSecRing) Then _
        MsgBox BPrintF("Файл ключей \'%s\' остался!", DelSecRing), vbExclamation, App.Title
    Kill DelPubRing
    If IsFile(DelPubRing) Then _
        MsgBox BPrintF("Файл ключей \'%s\' остался!", DelPubRing), vbExclamation, App.Title
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
    Dim s As String
    With Me
        If .Amount > 0 Then
            s = BPrintF("Банк-Клиент \'%s\' - %s (%f)", .ID, .Name, .Amount)
        Else
            s = BPrintF("Банк-Клиент \'%s\' - %s (минус %f)", .ID, .Name, Abs(.Amount))
        End If
    End With
    Application.Caption = s
End Sub
