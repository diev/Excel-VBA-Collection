VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CPayment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Option Base 1
DefLng A-Z

Const PaySheet = "Платежка"
Const BaseRange = "$A$1"
Const TERM = "^"

Dim mRow As Long
Dim mError As String

Dim mFileName As String
Dim mMark As String
Dim mDocNo As Long
Dim mDocDate As Date
Dim mSum As Currency
Dim mName As String
Dim mINN As String
Dim mBIC As String
Dim mLS As String
Dim mQueue As Long
Dim mDetails As String

Dim mCountLastSelected As Long
Dim mMoneyLastSelected As Currency

'Номера столбцов на листе Архива
Const ColFileName = 1
Const ColMark = 2
Const ColDocNo = 3
Const ColDocDate = 4
Const ColSum = 5
Const ColName = 6
Const ColINN = 7
Const ColBIC = 8
Const ColLS = 9
Const ColQueue = 10
Const ColDetails = 11

'Диапазон печати бланка на листе
Const CellRange = "$A$1:$AK$37"

'Подгон печатного листа под квадратные ячейки 5мм
Const CellWidth = 1.71
Const CellHeight = 15#

'Ячейки данных на бланке платежки
Const CellNo = "L4"
Const CellDate = "R4"
Const CellDelivery = "Z4"
Const CellAmount = "E6"
Const CellSum = "X9"
Const CellUserINN = "C9"
Const CellUserName = "A10"
Const CellUserLS = "Y12"
Const CellUserBank = "A15"
Const CellUserBIC = "X15"
Const CellUserPlace = "A16"
Const CellUserKS = "X16"
Const CellBank = "A18"
Const CellBIC = "X18"
Const CellPlace = "A19"
Const CellKS = "X19"
Const CellLS = "X21"
Const CellINN = "C21"
Const CellName = "A22"
Const CellAction = "AF24"
Const CellQueue = "AF25"
Const CellDetails = "A28"
Const CellSign1 = "K35"
Const CellSign2 = "K38"

Private Sub Class_Initialize()
    mError = vbNullString
    Me.Row = 2
    FillBlank
End Sub

Private Sub Class_Terminate()
    
End Sub

Public Property Get DocNo() As Variant
    DocNo = mDocNo
End Property

Public Property Let DocNo(ByVal vNewValue As Variant)
    mDocNo = vNewValue
End Property

Public Property Get DocDate() As Variant
    DocDate = mDocDate
End Property

Public Property Let DocDate(ByVal vNewValue As Variant)
    mDocDate = RDate(vNewValue)
End Property

Public Property Get Sum() As Variant
    Sum = mSum
End Property

Public Property Let Sum(ByVal vNewValue As Variant)
    mSum = RVal(vNewValue)
End Property

Public Property Get Name() As String
    Name = mName
End Property

Public Property Let Name(ByVal vNewValue As String)
    mName = vNewValue
End Property

Public Property Get INN() As String
    INN = mINN
End Property

Public Property Let INN(ByVal vNewValue As String)
    mINN = vNewValue
End Property

Public Property Get LS() As String
    LS = mLS
End Property

Public Property Let LS(ByVal vNewValue As String)
    mLS = vNewValue
End Property

Public Property Get BIC() As String
    BIC = mBIC
End Property

Public Property Let BIC(ByVal vNewValue As String)
    mBIC = vNewValue
End Property

Public Property Get BANK() As String
    BnkSeek2.BIC = Me.BIC
    BANK = BnkSeek2.Name
End Property

Public Property Get Place() As String
    BnkSeek2.BIC = Me.BIC
    Place = BnkSeek2.Place
End Property

Public Property Get KS() As String
    BnkSeek2.BIC = Me.BIC
    KS = BnkSeek2.KS
End Property

Public Property Get Delivery() As String
    BnkSeek2.BIC = Me.BIC
    Delivery = BnkSeek2.Delivery
End Property

Public Property Get Queue() As Variant
    Queue = mQueue
End Property

Public Property Let Queue(ByVal vNewValue As Variant)
    mQueue = vNewValue
End Property

Public Property Get Details() As String
    Details = mDetails
End Property

Public Property Let Details(ByVal vNewValue As String)
    mDetails = vNewValue
End Property

Public Property Get FileName() As String
    BnkSeek2.BIC = mBIC
    FileName = BPrintF("%s%X%X%02d%03d", BnkSeek2.NamePost, _
        Year(mDocDate) - 2000, Month(mDocDate), Day(mDocDate), mDocNo)
    If Len(mFileName) = 0 Then mFileName = FileName
End Property

Public Property Let FileName(ByVal vNewValue As String)
    mFileName = vNewValue
End Property

Public Property Get Mark() As String
    Mark = mMark
End Property

Public Property Let Mark(ByVal vNewValue As String)
    mMark = vNewValue
End Property

Public Property Get Valid() As Boolean
    mError = vbNullString
    If mDocNo <= 0 Then
        mError = "Номер документа не указан"
    ElseIf Year(mDocDate) < 2000 Then
        mError = "Дата документа не действительна"
    ElseIf mSum <= 0 Then
        mError = "Сумма документа не указана"
    ElseIf mQueue = 0 Then
        mError = "Очередность платежа не указана"
    ElseIf Len(mDetails) = 0 Then
        mError = "Не введено назначение платежа"
    ElseIf Len(mLS) = 0 Then
        mError = "Не указан л/с получателя"
    ElseIf Len(mBIC) = 0 Then
        mError = "Не указан БИК банка получателя"
    ElseIf Len(Me.BANK) = 0 Then
        mError = "Не найден банк получателя"
    ElseIf Right(mFileName, 3) = "000" Then
        mError = "Сформировано неправильное имя файла"
    ElseIf Right(User.LS, 3) <> User.ID Then
        mError = "Попытка отправить с чужого л/с"
    End If
    Valid = Len(mError) = 0
End Property

Public Sub FillBlank()
    Dim R As Range
    On Error Resume Next
    Application.ScreenUpdating = False
    Set R = Worksheets(PaySheet).Range(BaseRange)
    R.Range(CellUserName) = User.Name
    R.Range(CellUserINN) = User.INN
    R.Range(CellUserLS) = User.LS
    R.Range(CellUserBIC) = User.BIC
    R.Range(CellUserBank) = User.BANK
    R.Range(CellUserPlace) = User.Place
    R.Range(CellUserKS) = User.KS
    R.Range(CellSign1) = User.Sign(1)
    R.Range(CellSign2) = User.Sign(2)
        
    R.Range(CellNo) = vbNullString
    R.Range(CellDate) = PlatDate(Now)
    R.Range(CellDelivery) = vbNullString
    R.Range(CellAmount) = vbNullString
    R.Range(CellSum) = vbNullString
    R.Range(CellName) = vbNullString
    R.Range(CellINN) = vbNullString
    R.Range(CellBIC) = vbNullString
    R.Range(CellLS) = vbNullString
    R.Range(CellBank) = vbNullString
    R.Range(CellPlace) = vbNullString
    R.Range(CellKS) = vbNullString
    R.Range(CellQueue) = vbNullString
    R.Range(CellDetails) = vbNullString
    'R.Range(CellAction) = vbNullString
    Set R = Worksheets(PaySheet).Range(CellRange)
    R.ColumnWidth = CellWidth
    R.RowHeight = CellHeight
    Application.ScreenUpdating = True
End Sub

Public Sub FillData()
    Dim R As Range
    On Error Resume Next
    Application.ScreenUpdating = False
    Set R = Worksheets(PaySheet).Range(BaseRange)
    With User
        R.Range(CellUserName) = .Name
        R.Range(CellUserINN) = .INN
        R.Range(CellUserLS) = .LS
        R.Range(CellUserBIC) = .BIC
        R.Range(CellUserBank) = .BANK
        R.Range(CellUserPlace) = .Place
        R.Range(CellUserKS) = .KS
        R.Range(CellSign1) = .Sign(1)
        R.Range(CellSign2) = .Sign(2)
    End With
    R.Range(CellNo) = mDocNo
    R.Range(CellDate) = PlatDate(mDocDate)
    R.Range(CellAmount) = RSumStr(mSum)
    R.Range(CellSum) = PlatFormat(mSum)
    R.Range(CellName) = mName
    R.Range(CellINN) = mINN
    R.Range(CellBIC) = mBIC
    R.Range(CellLS) = mLS
    R.Range(CellQueue) = mQueue
    R.Range(CellDetails) = mDetails
    With Me
        R.Range(CellDelivery) = .Delivery
        R.Range(CellBank) = .BANK
        R.Range(CellPlace) = .Place
        R.Range(CellKS) = .KS
    End With
    Set R = Worksheets(PaySheet).Range(CellRange)
    R.ColumnWidth = CellWidth
    R.RowHeight = CellHeight
    Application.ScreenUpdating = True
End Sub

Public Function MarkByFile(FileName As String, Mark As String, Optional Color As Long = 0)
    Dim R As Range, s As String
    s = Mid(FileName, 2)
    MarkByFile = False
    On Error Resume Next
    Set R = Worksheets(User.ID).Range(BaseRange)
    mRow = Me.RowsCount
    Do While mRow > 1
        If InStr(2, R.Cells(mRow, ColFileName).Text, s) = 2 Then 'a bit faster!
            With R.Cells(mRow, ColMark)
                .Value = Mark
                If Color > 0 Then .Font.Color = QBColor(Color)
            End With
            MarkByFile = True
            Exit Function
        End If
        mRow = mRow - 1
    Loop
End Function

Public Property Get Row() As Long
    Row = mRow
End Property

Public Property Let Row(ByVal vNewValue As Long)
    On Error Resume Next
    mRow = NormRow(vNewValue)
    ReadRow
End Property

Public Property Get RowsCount() As Long
    Dim R As Range, Rw As Long
    Set R = Worksheets(User.ID).Range(BaseRange)
    Rw = 1
    Do 'looking for the last used row
        Rw = Rw + 1
    Loop Until Len(R.Cells(Rw, ColDocNo).Text) = 0
    Set R = Nothing
    RowsCount = Rw - 1
End Property

Public Sub PrintPreview()
    On Error Resume Next
    FillData
    If StrToBool(App.Options("DontPreviewPlat")) Then
        Application.GoTo Worksheets(PaySheet).Range("$A$1"), True
    Else
        Worksheets(PaySheet).PrintPreview
    End If
End Sub

Public Function AskAllRows(Optional Ask As String = "Обработать все?") As Boolean
    Dim A As Range, R As Range, s As String
    mCountLastSelected = 0
    mMoneyLastSelected = 0
    With Worksheets(User.ID)
        For Each A In Selection.Areas
            For Each R In A.Rows
                If Len(.Cells(R.Row, ColDocNo).Text) > 0 Then
                    If R.Row > 1 Then
                        mCountLastSelected = mCountLastSelected + 1
                        mMoneyLastSelected = mMoneyLastSelected + .Cells(R.Row, ColSum)
                    End If
                End If
            Next R
        Next A
    End With
    If mCountLastSelected = 1 Then
        'Ask = BPrintF("Помечен 1 документ\nна сумму %f", mMoneyLastSelected)
        AskAllRows = True
        Exit Function
    ElseIf mMoneyLastSelected = 0 Then
        Ask = BPrintF("Помечено документов: %d\nНО ОБЩАЯ СУММА НУЛЕВАЯ!\n\n%s", mCountLastSelected, Ask)
    Else
        Ask = BPrintF("Помечено документов: %d\nНа общую сумму: %f\n\n%s", mCountLastSelected, mMoneyLastSelected, Ask)
    End If
    AskAllRows = MsgBox(Ask, vbQuestion + vbYesNo, App.Title) = vbYes
End Function

Public Function EachSelected(Action As String, Optional Ask As String = "Обработать все?") As Boolean
    Dim A As Range, R As Range
    EachSelected = False
    If Not AskAllRows(Ask) Then Exit Function
    If Action = "Delete" Then
        If MsgBox("Безвозвратно удалить?!", vbExclamation + vbYesNo, App.Title) = vbNo Then
            Exit Function
        End If
    End If
    For Each A In Selection.Areas
        If Action = "Delete" Then
            For Each R In A.Rows
                If R.Row = 1 Then
                    MsgBox "Нельзя удалять заголовок таблицы!", vbCritical, App.Title
                    Exit Function
                End If
            Next R
            A.Rows.EntireRow.Delete
        Else
            For Each R In A.Rows
                If R.Row > 1 Then
                    With Me
                        .ReadRow R.Row
                        If .Valid Then
                            Select Case Action
                                Case "Preview":
                                    PrintPreview
                                Case "ExportFile":
                                    ExportToFile
                                Case "ExportPlat":
                                    SavePlatPGP
                                Case Else
                                    '?
                            End Select
                        Else
                            If MsgBox(BPrintF("Документ в строке %d\nне может быть обработан!\n(%s)\n\nПрервать процесс?", _
                                .Row, .ValidationError), vbExclamation + vbYesNo, App.Title) = vbYes Then Exit Function
                        End If
                    End With
                End If
            Next R
        End If
    Next A
    EachSelected = True
End Function

Public Property Get AsString() As String
    AsString = mFileName & TERM & _
        mMark & TERM & _
        mDocNo & TERM & _
        DtoS(mDocDate) & TERM & _
        PlatFormat(mSum) & TERM & _
        mName & TERM & _
        mINN & TERM & _
        mBIC & TERM & _
        mLS & TERM & _
        mQueue & TERM & _
        mDetails & TERM
End Property

Public Property Let AsString(ByVal vNewValue As String)
    Dim Arr As Variant
    If vNewValue = vbNullString Then Exit Property
    Arr = StrToArr(vNewValue, TERM)
    mFileName = Arr(ColFileName)
    mMark = Arr(ColMark)
    mDocNo = Arr(ColDocNo)
    mDocDate = StoD(CStr(Arr(ColDocDate)))
    mSum = RVal(Arr(ColSum))
    mName = Arr(ColName)
    mINN = Arr(ColINN)
    mBIC = Arr(ColBIC)
    mLS = Arr(ColLS)
    mQueue = Arr(ColQueue)
    mDetails = Arr(ColDetails)
End Property

Public Property Get AsPlatINI() As String
    Dim s As String, ss(4) As String
    On Error Resume Next
    s = BPrintF("; %s\n[File]\n", App.Title)
    s = s & BPrintF("File=%s.%s\nNow=%T\nForm=2\nVer=8\nCP=866\nVersion=%s\n", _
        mFileName, User.ID, Now, App.Version)
    
    s = s & BPrintF("\n[Payment]\n")
    s = s & BPrintF("No=%d\n", mDocNo)
    s = s & BPrintF("Date=%N\n", mDocDate)
    s = s & BPrintF("Sum=%F\n", mSum)
    
    s = s & BPrintF("Queue=%d\n", mQueue)
    
    WrapStrToArr mDetails, ss, 64
    s = s & BPrintF("Text1=%s\n", ss(1))
    s = s & BPrintF("Text2=%s\n", ss(2))
    s = s & BPrintF("Text3=%s\n", ss(3))
    s = s & BPrintF("Text4=%s\n", ss(4))
    
    s = s & BPrintF("\n[From]\n")
    s = s & BPrintF("Name=%s\n", User.Name)
    s = s & BPrintF("INN=%s\n", User.INN)
    s = s & BPrintF("LS=%s\n", User.LS)
    s = s & BPrintF("Sign1=%s\n", User.Sign(1))
    s = s & BPrintF("Sign2=%s\n", User.Sign(2))
    s = s & BPrintF("Tel=%s\n", User.Tel)
    
    s = s & BPrintF("\n[To]\n")
    s = s & BPrintF("Name2=%s\n", mName)
    s = s & BPrintF("INN2=%s\n", mINN)
    s = s & BPrintF("LS2=%s\n", mLS)
    s = s & BPrintF("BIC2=%s\n", mBIC)
    
    s = s & BPrintF("Bank2=%s\n", Me.BANK)
    s = s & BPrintF("Place2=%s\n", Me.Place)
    s = s & BPrintF("KS2=%s\n", Me.KS)
    AsPlatINI = s & BPrintF("\n; eof\n")
End Property

Public Property Get ValidationError() As String
    ValidationError = mError
End Property

Public Function ExportToFile(Optional File As String = vbNullString) As Boolean
    Dim A As Range, R As Range, s As String, x As Integer
    On Error Resume Next
    ExportToFile = False
    If Not AskAllRows() Then Exit Function
    If Not IsFile(File) Then File = User.ExportList
    Application.DefaultFilePath = FilePath(File)
    'If Not IsFile(File) Then
        'If IsDir(File) Then '/////////////////////wrong if nested unexisting dirs/
            File = RightPathName(File, User.ID & To36(Day(Now)) & Format(Now, "hhmm") & ".plt")
            If Not BrowseForSave(File, "Файлы выгрузки (*.plt),*.plt", _
                "файл для выгрузки") Then Exit Function
            User.ExportList = FilePath(File)
        'End If
    'End If
    s = vbNullString
    mCountLastSelected = 0
    For Each A In Selection.Areas
        For Each R In A.Rows
            With Me
                .ReadRow R.Row
                If .Valid Then
                    .MarkByRow R.Row, "Disk", 1 'Blue
                    s = s & Me.AsString & vbCrLf
                    mCountLastSelected = mCountLastSelected + 1
                    Application.StatusBar = BPrintF("Подготовлено: %d", mCountLastSelected)
                Else
                    .MarkByRow R.Row, "Bad", 4 'Red qbcolor
                    If MsgBox(BPrintF("Документ в строке %d\nне может быть обработан!\n(%s)\n\nПрервать процесс?", _
                        .Row, .ValidationError), vbExclamation + vbYesNo, App.Title) = vbYes Then
                        Application.StatusBar = False
                        Exit Function
                    End If
                End If
            End With
        Next R
    Next A
    Application.StatusBar = False
    If Len(s) = 0 Then
        MsgBox "Нечего выгружать!", vbExclamation, App.Title
        Exit Function
    Else
        If Not StrToBool(App.Options("DontExportDos")) Then s = CDos(s)
    End If
    If IsFile(File) Then
        x = MsgBox(BPrintF("Файл \'%s\'уже существует!\nЧтобы добавить, нажмите \'Да\'\nПерезаписать - \'Нет\'", File), _
            vbQuestion + vbYesNoCancel, App.Title)
        Select Case x
            Case vbYes:
                AppendFile File, s
            Case vbNo:
                OutputFile File, s
            Case Else
                MsgBox "Выгрузка НЕ выполнена!", vbExclamation, App.Title
                Exit Function
        End Select
    Else
        OutputFile File, s
    End If
    If IsFile(File) Then
        MsgBox BPrintF("Выгрузка выполнена в файл\n\'%s\'\nРазмер файла: %d байт", _
            File, FileLen(File)), vbInformation, App.Title
        ExportToFile = True
    Else
        MsgBox "Ошибка выгрузки!", vbCritical, App.Title
        ExportToFile = False
    End If
End Function

Public Function ImportFromFile(Optional File As String = vbNullString) As Boolean
    Dim Files As Variant, i As Long
    On Error Resume Next
    If IsFile(File) Then
        ImportFromFile = ImportFrom1File(File)
    Else
        Files = User.ImportList
        If BrowseForFiles(Files, "Файлы загрузки (*.plt),*.plt", "файл(ы) для загрузки") Then
            User.ImportList = FilePath(CStr(Files(1)))
            'MultiSelect
            For i = LBound(Files) To UBound(Files)
                ImportFromFile = ImportFrom1File(CStr(Files(i)))
            Next
        Else
            ImportFromFile = False
        End If
    End If
End Function

Public Function ImportFrom1File(File As String) As Boolean
    On Error Resume Next
    Dim s As String, i As Long, ss As Variant, xReply As Long
    s = InputFile(File)
    If Len(s) > 0 Then
        If Not StrToBool(App.Options("DontExportDos")) Then s = CWin(s)
        s = StrTran(s, vbCrLf, vbLf) 'DOS/UNIX ready
        ss = StrToArr(s, vbLf)
        'Application.ScreenUpdating = False
        mCountLastSelected = 0
        mMoneyLastSelected = 0
        mRow = Me.RowsCount
        For i = 1 To UBound(ss)
            s = ss(i)
            If Len(s) > 0 Then
                mCountLastSelected = mCountLastSelected + 1
                Me.AsString = ss(mCountLastSelected)
                mRow = mRow + 1
                WriteRow mRow
                mMoneyLastSelected = mMoneyLastSelected + Me.Sum
                Application.StatusBar = BPrintF("Загружено: %d", mCountLastSelected)
            End If
        Next
        'Application.ScreenUpdating = True
        Application.StatusBar = False
    
        xReply = MsgBox(BPrintF("Завершена загрузка из файла\n\'%s\'\n\nДокументов: %d\nОбщая сумма: %f\n\nУдалить его?", _
            File, mCountLastSelected, mMoneyLastSelected), vbQuestion + vbYesNoCancel, App.Title)
        Select Case xReply
            Case vbYes: Kill File
            Case vbNo
            Case Else
                Exit Function
        End Select
    Else
        xReply = MsgBox(BPrintF("Ошибка загрузки из файла\n\'%s\'\n\nУдалить его?", File), vbQuestion + vbYesNoCancel + vbDefaultButton2, App.Title)
        Select Case xReply
            Case vbYes: Kill File
            Case vbNo:
            Case Else
                Exit Function
        End Select
    End If
End Function

Public Function SavePlat(File As String) As Boolean
    SavePlat = False
    If OverwriteFile(File) Then
        OutputFile File, CDos(Me.AsPlatINI)
        SavePlat = IsFile(File)
    End If
End Function

Public Function LoadPlat(File As String) As Boolean
    Dim Plat As New CIniFile
    On Error Resume Next
    mFileName = FileNameOnly(File)
    With Plat
        .ReadOnly = True
        .DefaultSection = "Payment"
        mDocNo = .KeyValue("No")
        mDocDate = RDate(.KeyValue("Date"))
        mSum = RVal(.Key("Sum"))
        mQueue = .KeyValue("Queue")
        mDetails = Trim(.Key("Text1") & " " & _
            .Key("Text2") & " " & .Key("Text3") & " " & .Key("Text4"))
        
        .DefaultSection = "To"
        mName = .Key("Name2")
        mINN = .Key("INN2")
        mLS = .Key("LS2")
        mBIC = .Key("BIC2")
    End With
    WriteRow
    LoadPlat = Me.Valid
End Function

Public Function SavePlatPGP() As Boolean
    Dim File As String, File1 As String, File2 As String
    On Error Resume Next
    File = Me.FileName & "." & User.ID
    File1 = RightPathName(GetWinTempDir, File)
    File2 = SMail.Send & File
    If SavePlat(File1) Then
        If PGP.Encode(File1, File2) Then
            If Not StrToBool(App.Options("DontMarkArchive")) Then _
                MarkByFile mFileName, "Mail", 1  'Blue
            MsgBox BPrintF("Файл \'%s\' готов", File), vbInformation, App.Title
        Else
            MsgBox "Шифрованный файл НЕ создан!", vbCritical, App.Title
        End If
        Kill File1
    Else
        MsgBox "Файл для шифрования НЕ создан!", vbCritical, App.Title
    End If
    SavePlatPGP = IsFile(File2)
End Function

Public Sub Delete()
    On Error Resume Next
    If mRow > 1 Then Worksheets(User.ID).Range(BaseRange).Rows(mRow).EntireRow.Delete
End Sub

Public Sub Clear()
    mFileName = vbNullString
    mMark = vbNullString
    mDocNo = 0
    mDocDate = Now
    mSum = 0
    mName = vbNullString
    mINN = vbNullString
    mBIC = vbNullString
    mLS = vbNullString
    mQueue = 0
    mDetails = vbNullString
End Sub

Public Sub ReadRow(Optional Row As Long)
    If Not IsMissing(Row) Then mRow = NormRow(Row)
    With Worksheets(User.ID).Range(BaseRange)
        mFileName = .Cells(mRow, ColFileName)
        mMark = .Cells(mRow, ColMark)
        mDocNo = .Cells(mRow, ColDocNo)
        mDocDate = .Cells(mRow, ColDocDate)
        mSum = .Cells(mRow, ColSum)
        mName = .Cells(mRow, ColName)
        mINN = .Cells(mRow, ColINN)
        mBIC = .Cells(mRow, ColBIC)
        mLS = .Cells(mRow, ColLS)
        mQueue = .Cells(mRow, ColQueue)
        mDetails = .Cells(mRow, ColDetails)
    End With
End Sub

Public Sub WriteRow(Optional Row As Long)
    If Not IsMissing(Row) Then mRow = NormRow(Row)
    With Worksheets(User.ID).Range(BaseRange)
        .Cells(mRow, ColFileName) = mFileName
        .Cells(mRow, ColMark) = mMark
        .Cells(mRow, ColDocNo) = mDocNo
        .Cells(mRow, ColDocDate) = mDocDate
        .Cells(mRow, ColSum) = mSum
        .Cells(mRow, ColName) = mName
        .Cells(mRow, ColINN) = mINN
        .Cells(mRow, ColBIC) = mBIC
        .Cells(mRow, ColLS) = mLS
        .Cells(mRow, ColQueue) = mQueue
        .Cells(mRow, ColDetails) = mDetails
    End With
End Sub

Private Function NormRow(ByVal Row As Long) As Long
    Dim NewRow As Long
    NewRow = 2
    With Worksheets(User.ID).Range(BaseRange)
        Do While Len(.Cells(NewRow, ColDocNo).Text) > 0
            NewRow = NewRow + 1
        Loop
    End With
    If Row = 0 Then
        Worksheets(User.ID).Activate
        Row = ActiveCell.Row
    ElseIf Row = -1 Then
        Row = NewRow
    End If
    If Row < 2 Then
        NormRow = 2
    ElseIf Row > NewRow Then
        NormRow = NewRow
    Else
        NormRow = Row
    End If
End Function

Public Sub SortBy(Col As Long)
    On Error Resume Next
    Worksheets(User.ID).Range(BaseRange).Sort _
        Key1:=Cells(1, Col), Order1:=xlAscending, Header:=xlGuess, _
        OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom
End Sub

Public Sub FindText()
    Static s As String
    On Error Resume Next
    s = InputBox("Введите часть строки для поиска:", App.Title, s)
    If s = vbNullString Then Exit Sub
    Worksheets(User.ID).Cells.Find(What:=s, after:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False).Activate
End Sub

'Public Sub FindNext()
'    On Error Resume Next
'    Worksheets(User.ID).Cells.FindNext(after:=ActiveCell).Activate
'End Sub
'
'Public Sub FindPrev()
'    On Error Resume Next
'    Worksheets(User.ID).Cells.FindPrev(after:=ActiveCell).Activate
'End Sub

Public Sub MarkByRow(Row As Long, Mark As String, Optional Color As Long = 0)
    On Error Resume Next
    With Worksheets(User.ID).Range(BaseRange)
        .Cells(Row, ColMark) = Mark
        If Color > 0 Then .Cells(Row, ColMark).Font.Color = QBColor(Color)
    End With
End Sub

Public Property Get CountLastSelected() As Long
    CountLastSelected = mCountLastSelected
End Property

Public Property Get MoneyLastSelected() As Currency
    MoneyLastSelected = mMoneyLastSelected
End Property

