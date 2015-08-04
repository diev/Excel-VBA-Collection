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

Const NoDocMin = 1
Const NoDocMax = 999

Dim mUserID As String
Dim mUserIDFile As String

Private Sub Class_Initialize()
    Me.ID = UserDemoID
End Sub

Public Property Get ID() As String
    Dim s As String
    s = Me.ID4(mUserID)
    If mUserID = App.DefLS(s) Then
        ID = s
    Else
        ID = mUserID
    End If
End Property

Public Property Let ID(ByVal vNewValue As String)
    Dim s As String, s4 As String, b As Boolean
    Dim File As String, Arr As Variant
    Dim sh As Worksheet
    On Error Resume Next
    If vNewValue = UserDemoID Then
        mUserID = UserDemoID
        mUserIDFile = App.Path & mUserID & UserIDExt
    Else
        s4 = Me.ID4(vNewValue)
        s = App.DefLS(s4)
        If Len(vNewValue) = 20 Then
            mUserID = vNewValue
        ElseIf vNewValue = UserDemoID Then
            mUserID = UserDemoID
        ElseIf vNewValue = s4 Then
            mUserID = s
        Else
            Exit Property
        End If
        File = mUserID & UserIDExt
        mUserIDFile = App.Path & File
        If Not IsFile(mUserIDFile) Then
            If IsFile(App.Path & s4 & UserIDExt) Then
                s = InputFile(App.Path & s4 & UserIDExt)
                Kill mUserIDFile
                s = "[" & mUserID & Mid(s, InStr(2, s, "]"))
                OutputFile mUserIDFile, s
                If IsFile(mUserIDFile) Then
                    Kill App.Path & s4 & UserIDExt
                End If
            End If
        End If
        s = InputFile(mUserIDFile)
        If Len(s) = 0 Then
            With Me
                .BIC = .BIC
                .LS = .LS
                .KPP = .KPP
                .INN = .INN
                .Name = .Name
            End With
        ElseIf NumIniFile(mUserIDFile, mUserID, "CP", 866) <> 1251 Then
            Kill mUserIDFile
            If InStr(1, s, vbCrLf) = 0 Then 'UNIX
                Arr = StrToArr(s, vbLf)
                s = ArrToStr(Arr, vbCrLf)
            End If
            OutputFile mUserIDFile, CWin(s)
            WriteIniFile mUserIDFile, mUserID, "CP", "1251"
        End If
        'If OkCancelBox("Тщательно проверьте Ваши реквизиты!") Then
        '    EditUserShow
        'End If
    End If
    If App.DefLS(s4) = vNewValue Then '3-4
        s = s4
    Else
        s = mUserID
    End If
    b = False
    For Each sh In Worksheets
        If sh.Name = s4 Or sh.Name = mUserID Then
            b = True
            Exit For
        End If
    Next
    If Not b Then
        NewUserSheet s
    End If
    ResetCaption
    'Commented due to an abnormous screen flashing!...
    'Payment.FillBlank
    'MailClear
End Property

Public Property Get Name() As String
    Dim i1 As Integer, i2 As Integer, s As String
    s = ReadIniFile(mUserIDFile, mUserID, "Name", "ООО ""Демонстрация""")
    If Left(s, 3) = "КПП" Then
        i1 = InStr(3, s, " ") + 1
        i2 = InStr(i1, s, " ")
        Me.KPP = Mid(s, i1, i2 - i1)
        Me.Name = Mid(s, i2 + 1)
    End If
    Name = s
End Property

Public Property Let Name(ByVal vNewValue As String)
    WriteIniFile mUserIDFile, mUserID, "Name", vNewValue
End Property

Public Property Get INN() As String
    INN = ReadIniFile(mUserIDFile, mUserID, "INN", "0")
End Property

Public Property Let INN(ByVal vNewValue As String)
    WriteIniFile mUserIDFile, mUserID, "INN", vNewValue
End Property

Public Property Get KPP() As String
    KPP = ReadIniFile(mUserIDFile, mUserID, "KPP", "0")
End Property

Public Property Let KPP(ByVal vNewValue As String)
    WriteIniFile mUserIDFile, mUserID, "KPP", vNewValue
End Property

Public Property Get OKATO() As String
    OKATO = ReadIniFile(mUserIDFile, mUserID, "OKATO", "0")
End Property

Public Property Let OKATO(ByVal vNewValue As String)
    WriteIniFile mUserIDFile, mUserID, "OKATO", vNewValue
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
        End With
    End If
End Property

Public Property Get Bank() As String
    With BnkSeek2
        .BIC = Me.BIC
        Bank = Trim(.Name)
    End With
End Property

Public Property Get Place() As String
    With BnkSeek2
        .BIC = Me.BIC
        Place = Trim(.Place)
    End With
End Property

Public Property Get KS() As String
    With BnkSeek2
        .BIC = Me.BIC
        KS = Trim(.KS)
    End With
End Property

Public Property Get No() As Long
    No = NumIniFile(mUserIDFile, mUserID, "No", 1)
End Property

Public Property Let No(ByVal vNewValue As Long)
    If vNewValue > NoDocMax Then
        WriteIniFile mUserIDFile, mUserID, "No", NoDocMin
    ElseIf vNewValue < NoDocMin Then
        WriteIniFile mUserIDFile, mUserID, "No", NoDocMin
    Else
        WriteIniFile mUserIDFile, mUserID, "No", CStr(vNewValue)
    End If
End Property

Public Property Get NoMin() As Long
    NoMin = NoDocMin
End Property

Public Property Get NoMax() As Long
    NoMax = NoDocMax
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

Public Sub ResetCaption()
    With Me
        Application.Caption = Bsprintf("Банк-Клиент %s (%s)", .ID, .Name)
    End With
End Sub

Public Sub Edit()
    Dim s As String
    s = Bsprintf("notepad.exe %s", mUserIDFile)
    Shell s, vbNormalFocus
End Sub

Public Function ID4(Optional anID As String = vbNullString) As String
    Dim s As String, v As Long
    If anID = vbNullString Then
        anID = mUserID
    End If
    If anID = DemoID Then
        ID4 = DemoID
    'ElseIf Len(anID) = 20 And Left(anID, 2) = "30" Then
    '    ID4 = "k" & Right(anID, 3)
    Else
        v = Val(Right(anID, 4))
        If v = 0 Then
            ID4 = DemoID
        Else
            ID4 = CStr(v)
        End If
    End If
End Function

Public Function IsID4(s As String) As Boolean
    Dim ws As Worksheet
    If Len(s) = 0 Then
        IsID4 = False
    ElseIf s = DemoID Then
        IsID4 = True
    'ElseIf Left(s, 1) = "k" And IsDigital(Right(s, 3)) Then
    '    For Each ws In Workbooks(App.BookName).Worksheets
    '        If Len(ws.Name) = 20 And Right(s, 3) = Mid(s, 2) Then
    '            IsID4 = True
    '            Exit Function
    '        End If
    '    Next
    '    IsID4 = False
    ElseIf IsDigital(s) Then
        For Each ws In Worksheets
            If Right(ws.Name, Len(s)) = s Then
                IsID4 = True
                Exit Function
            End If
        Next
        IsID4 = False
    Else
        IsID4 = False
    End If
End Function

Public Sub Add(BIC As String, LS As String, KPP As String, INN As String, Name As String)
    Dim File As String
    File = App.Path & LS & UserIDExt
    WriteIniFile File, LS, "BIC", BIC
    WriteIniFile File, LS, "LS", LS
    WriteIniFile File, LS, "KPP", KPP
    WriteIniFile File, LS, "INN", INN
    WriteIniFile File, LS, "Name", Name
    WriteIniFile File, LS, "CP", "1251"
End Sub

Public Function ValidationError() As String
    Dim Item As String, L As Long, s As String
    
    ValidationError = vbNullString
    If Me.Demo Then Exit Function
    
    s = Me.Name
    Item = "Плательщик "
    L = Len(s)
    If L = 0 Then
        ValidationError = Item & "не указан"
        Exit Function
    ElseIf L > 160 Then
        ValidationError = Item & "не может превышать 160 символов"
        Exit Function
    ElseIf InStr(s, "  ") > 0 Then
        ValidationError = Item & "не должен содержать лишних пробелов"
        Exit Function
    ElseIf InStr(s, "?") > 0 Then
        ValidationError = Item & "не должен содержать символ ?"
        Exit Function
    ElseIf InStr(s, "^") > 0 Then
        ValidationError = Item & "не должен содержать символ ^"
        Exit Function
    ElseIf Left(s, 1) = """" Then
        ValidationError = Item & "не должен начинаться с кавычки"
        Exit Function
    ElseIf Left(s, 1) = "-" Then
        ValidationError = Item & "не может начинаться с минуса"
        Exit Function
    End If
    
    s = Me.INN
    Item = "ИНН плательщика "
    If Left(s, 1) = "F" Then 'foreign
        s = Mid(s, 2)
    End If
    L = Len(s)
    If s = "0" Then
        'valid
    ElseIf Not IsDigital(s) Then
        ValidationError = Item & "не из цифр"
        Exit Function
    ElseIf L > 1 And L <> 10 And L <> 12 Then
        ValidationError = Item & "не 10 и не 12 цифр"
        Exit Function
    ElseIf L > 1 And Not INNKeyValid(s) Then
        If User.Demo Then
            WarnBox "%s %s неверный,\nно эта ошибка прощается в деморежиме.", Item, s
        Else
            ValidationError = Item & "неверный по ключу"
            Exit Function
        End If
    End If
    
    s = Me.KPP
    Item = "КПП плательщика "
    L = Len(s)
    If s = "0" Then
        'valid
    ElseIf Not IsDigital(s) Then
        ValidationError = Item & "не из цифр"
        Exit Function
    ElseIf L > 1 And L <> 9 Then
        ValidationError = Item & "не 9 цифр"
        Exit Function
    End If
    
    s = Me.LS
    Item = "Счет плательщика "
    L = Len(s)
    If L <> 20 Then
        ValidationError = Item & "не 20 цифр"
        Exit Function
    ElseIf Not IsDigital(s) Then
        ValidationError = Item & "не из цифр"
        Exit Function
    ElseIf Mid(s, 6, 3) <> "810" Then
        ValidationError = Item & "не имеет кода валюты 810"
        Exit Function
    ElseIf Not LSKeyValid(Me.BIC, s) Then
        ValidationError = Item & "не соответствует БИК по ключу"
        Exit Function
    End If
        
    s = Me.BIC
    Item = "БИК банка плательщика "
    If s <> App.DefBIC Then
        ValidationError = Item & "не " & App.DefBIC
        Exit Function
    End If
End Function


