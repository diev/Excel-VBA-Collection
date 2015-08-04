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
Const WRAP = "!"

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
Const CellRange = "$A$1:$AI$38"

'Подгон печатного листа под квадратные ячейки 5мм
Const CellWidth = 1.8
Const CellHeight = 14.5

'Ячейки данных на бланке платежки
Const CellNo = "M4"
Const CellDate = "R4"
Const CellDelivery = "Z4"
Const CellAmount = "E6"
Const CellSum = "X9"
Const CellUserINN = "C9"
Const CellUserName = "A10"
Const CellUserLS = "X12"
Const CellUserBank = "A15"
Const CellUserBIC = "X15"
Const CellUserPlace = "G17"
Const CellUserKS = "X16"
Const CellBank = "A18"
Const CellBIC = "X18"
Const CellPlace = "G20"
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

Public Property Get Bank() As String
    BnkSeek2.BIC = Me.BIC
    Bank = BnkSeek2.Name
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
    'FileName = Bsprintf("%s%X%X%02d%03d", BnkSeek2.NamePost, _
    '    Year(mDocDate) - 2000, Month(mDocDate), Day(mDocDate), mDocNo)
    FileName = Bsprintf("A%X%X%02d%03d", _
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
    ElseIf Year(mDocDate) < 1999 Then
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
    ElseIf Len(Me.Bank) = 0 Then
        mError = "Не найден банк получателя"
    ElseIf Right(mFileName, 3) = "000" Then 'нет номера платежки
        mError = "Сформировано неправильное имя файла"
    ElseIf Right(User.LS, Len(User.ID)) <> User.ID Then
        mError = Bsprintf("Попытка отправить от клиента %s\nсо счета %s", User.ID, User.LS)
    End If
    Valid = Len(mError) = 0
End Property

Public Sub FillBlank()
    Dim r As Range
    On Error Resume Next
    Application.ScreenUpdating = False
    Set r = Worksheets(PaySheet).Range(BaseRange)
    r.Range(CellUserName) = StrTran(User.Name, WRAP, vbLf)
    r.Range(CellUserINN) = User.INN
    r.Range(CellUserLS) = User.LS
    r.Range(CellUserBIC) = User.BIC
    r.Range(CellUserBank) = User.Bank
    r.Range(CellUserPlace) = User.Place
    r.Range(CellUserKS) = User.KS
    r.Range(CellSign1) = User.Sign(1)
    r.Range(CellSign2) = User.Sign(2)
        
    r.Range(CellNo) = vbNullString
    r.Range(CellDate) = PlatDate(Now)
    r.Range(CellDelivery) = vbNullString
    r.Range(CellAmount) = vbNullString
    r.Range(CellSum) = vbNullString
    r.Range(CellName) = vbNullString
    r.Range(CellINN) = vbNullString
    r.Range(CellBIC) = vbNullString
    r.Range(CellLS) = vbNullString
    r.Range(CellBank) = vbNullString
    r.Range(CellPlace) = vbNullString
    r.Range(CellKS) = vbNullString
    r.Range(CellQueue) = vbNullString
    r.Range(CellDetails) = vbNullString
    'R.Range(CellAction) = vbNullString
    Set r = Worksheets(PaySheet).Range(CellRange)
    r.ColumnWidth = CellWidth
    r.RowHeight = CellHeight
    Application.ScreenUpdating = True
End Sub

Public Sub FillData()
    Dim r As Range, i As Long
    On Error Resume Next
    Application.ScreenUpdating = False
    Set r = Worksheets(PaySheet).Range(BaseRange)
    With User
        r.Range(CellUserName) = StrTran(.Name, WRAP, vbLf)
        r.Range(CellUserINN) = .INN
        r.Range(CellUserLS) = .LS
        r.Range(CellUserBIC) = .BIC
        r.Range(CellUserBank) = .Bank
        r.Range(CellUserPlace) = .Place
        r.Range(CellUserKS) = .KS
        r.Range(CellSign1) = .Sign(1)
        r.Range(CellSign2) = .Sign(2)
    End With
    r.Range(CellNo) = mDocNo
    r.Range(CellDate) = PlatDate(mDocDate)
    r.Range(CellAmount) = RSumStr(mSum)
    r.Range(CellSum) = PlatFormat(mSum)
    r.Range(CellName) = StrTran(mName, WRAP, vbLf)
    r.Range(CellINN) = mINN
    r.Range(CellBIC) = mBIC
    r.Range(CellLS) = mLS
    r.Range(CellQueue) = mQueue
    r.Range(CellDetails) = StrTran(mDetails, WRAP, vbLf)
    
    With Me
        'r.Range(CellDelivery) = .Delivery
        r.Range(CellBank) = .Bank
        r.Range(CellPlace) = .Place
        r.Range(CellKS) = .KS
    End With
    Set r = Worksheets(PaySheet).Range(CellRange)
    r.ColumnWidth = CellWidth
    r.RowHeight = CellHeight
    Application.ScreenUpdating = True
End Sub

Public Function MarkByFile(FileName As String, Mark As String, Optional Color As Long = 0) As Boolean
    Dim r As Range, s As String
    s = Mid(FileName, 2)
    MarkByFile = False
    On Error Resume Next
    Set r = Worksheets(User.ID).Range(BaseRange)
    mRow = Me.RowsCount
    Do While mRow > 1
        If InStr(2, r.Cells(mRow, ColFileName).Text, s) = 2 Then 'a bit faster!
            With r.Cells(mRow, ColMark)
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
    Dim r As Range
    On Error Resume Next
    Set r = Worksheets(User.ID).Range(BaseRange)
    RowsCount = 1
    Do While Len(r.Cells(RowsCount, ColDocNo).Text) > 0 'looking for the last used row
        RowsCount = RowsCount + 1
    Loop
    Set r = Nothing
    RowsCount = RowsCount - 1
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
    Dim A As Range, r As Range, s As String
    mCountLastSelected = 0
    mMoneyLastSelected = 0
    With Worksheets(User.ID)
        For Each A In Selection.Areas
            For Each r In A.Rows
                If Len(.Cells(r.Row, ColDocNo).Text) > 0 Then
                    If r.Row > 1 Then
                        mCountLastSelected = mCountLastSelected + 1
                        mMoneyLastSelected = mMoneyLastSelected + .Cells(r.Row, ColSum)
                    End If
                End If
            Next r
        Next A
    End With
    If mCountLastSelected = 1 Then
        'Ask = BSPrintF("Помечен 1 документ\nна сумму %f", mMoneyLastSelected)
        AskAllRows = True
        Exit Function
    ElseIf mMoneyLastSelected = 0 Then
        Ask = Bsprintf("Помечено документов: %d\nНО ОБЩАЯ СУММА НУЛЕВАЯ!\n\n%s", mCountLastSelected, Ask)
    Else
        Ask = Bsprintf("Помечено документов: %d\nНа общую сумму: %f\n\n%s", mCountLastSelected, mMoneyLastSelected, Ask)
    End If
    AskAllRows = OkCancelBox(Ask)
End Function

Public Function EachSelected(Action As String, Optional Ask As String = "Обработать все?") As Boolean
    Dim A As Range, r As Range
    EachSelected = False
    If Not AskAllRows(Ask) Then Exit Function
    If Action = "Delete" Then
        If Not OkCancelBox("Безвозвратно удалить?!") Then
            Exit Function
        End If
    End If
    For Each A In Selection.Areas
        If Action = "Delete" Then
            For Each r In A.Rows
                If r.Row = 1 Then
                    StopBox "Нельзя удалять заголовок таблицы!"
                    Exit Function
                End If
            Next r
            A.Rows.EntireRow.Delete
        Else
            For Each r In A.Rows
                If r.Row > 1 Then
                    With Me
                        .ReadRow r.Row
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
                            If YesNoBox("Документ в строке %d\nне может быть обработан!\n(%s)\n\nПрервать процесс?", _
                                .Row, .ValidationError) Then Exit Function
                        End If
                    End With
                End If
            Next r
        End If
    Next A
    EachSelected = True
End Function

Public Property Get AsString() As String
    AsString = mFileName & TERM & _
        mMark & TERM & _
        mDocNo & TERM & _
        DtoC(mDocDate) & TERM & _
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
    On Error Resume Next
    If vNewValue = vbNullString Then Exit Property
    'If Len(vNewValue) < ColDetails Then Exit Property
    Arr = StrToArr(vNewValue, FindTerm(vNewValue)) 'autoTERM from ;^|
    mFileName = Arr(ColFileName)
    mMark = Arr(ColMark)
    mDocNo = Arr(ColDocNo)
    mDocDate = RDate(Arr(ColDocDate)) 'StoD(CStr(Arr(ColDocDate)))
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
    s = Bsprintf("; %s\n[File]\n", App.Title)
    s = s & Bsprintf("File=%s.%s\nNow=%T\nForm=2\nVer=10\nCP=866\nVersion=%s\nPath=%s\n", _
        mFileName, User.ID, Now, App.Version, App.Path)
    
    s = s & Bsprintf("\n[Payment]\n")
    s = s & Bsprintf("No=%d\n", mDocNo)
    s = s & Bsprintf("Date=%n\n", mDocDate) 'dd.mm.yyyy
    s = s & Bsprintf("Sum=%F\n", mSum) '0-00
    
    s = s & Bsprintf("Queue=%d\n", mQueue)
    
    WrapStrToArr mDetails, ss, 64
    s = s & Bsprintf("Text1=%s\n", ss(1))
    s = s & Bsprintf("Text2=%s\n", ss(2))
    s = s & Bsprintf("Text3=%s\n", ss(3))
    s = s & Bsprintf("Text4=%s\n", ss(4))
    
    s = s & Bsprintf("\n[From]\n")
    s = s & Bsprintf("Name=%s\n", StrSpaces1(User.Name))
    s = s & Bsprintf("INN=%s\n", User.INN)
    s = s & Bsprintf("LS=%s\n", User.LS)
'    s = s & Bsprintf("Sign1=%s\n", User.Sign(1))
'    s = s & Bsprintf("Sign2=%s\n", User.Sign(2))
'    s = s & Bsprintf("Tel=%s\n", User.Tel)
    
    s = s & Bsprintf("\n[To]\n")
    s = s & Bsprintf("Name2=%s\n", StrSpaces1(mName))
    s = s & Bsprintf("INN2=%s\n", mINN)
    s = s & Bsprintf("LS2=%s\n", mLS)
    s = s & Bsprintf("BIC2=%s\n", mBIC)
    
    s = s & Bsprintf("Bank2=%s\n", Me.Bank)
    s = s & Bsprintf("Place2=%s\n", Me.Place)
    s = s & Bsprintf("KS2=%s\n", Me.KS)
    AsPlatINI = s & Bsprintf("\n; eof\n")
End Property

Public Property Get ValidationError() As String
    ValidationError = mError
End Property

Public Function ExportToFile(Optional File As String = vbNullString) As Boolean
    Dim A As Range, r As Range, s As String, x As Integer
    On Error Resume Next
    ExportToFile = False
    If Not AskAllRows() Then Exit Function
    If Not IsFile(File) Then File = User.ExportList
    Application.DefaultFilePath = FilePath(File)
    'If Not IsFile(File) Then
        'If IsDir(File) Then '/////////////////////wrong if nested unexisting dirs/
            's = User.ID & To36(Day(Now)) & Format(Now, "hhmm") & ".plt" '8.3
            s = User.ID & Format(Now, "-dd-hhmm") & ".plt" 'LFN
            File = RightPathName(File, s)
            If Not BrowseForSave(File, "Файлы выгрузки (*.plt),*.plt", _
                "файл для выгрузки") Then Exit Function
            User.ExportList = FilePath(File)
        'End If
    'End If
    s = vbNullString
    mCountLastSelected = 0
    For Each A In Selection.Areas
        For Each r In A.Rows
            With Me
                .ReadRow r.Row
                If .Valid Then
                    .MarkByRow r.Row, "Disk", 1 'Blue
                    s = s & Me.AsString & vbCrLf
                    mCountLastSelected = mCountLastSelected + 1
                    Application.StatusBar = Bsprintf("Подготовлено: %d", mCountLastSelected)
                Else
                    .MarkByRow r.Row, "Bad", 4 'Red qbcolor
                    If YesNoBox("Документ в строке %d\nне может быть обработан!\n(%s)\n\nПрервать процесс?", _
                        .Row, .ValidationError) = vbYes Then
                        Application.StatusBar = False
                        Exit Function
                    End If
                End If
            End With
        Next r
    Next A
    Application.StatusBar = False
    If Len(s) = 0 Then
        WarnBox "Нечего выгружать!"
        Exit Function
    Else
        If Not StrToBool(App.Options("DontExportDos")) Then s = CDos(s)
    End If
    If IsFile(File) Then
        x = MsgBox(Bsprintf("Файл %s уже существует!\nЧтобы добавить, нажмите \'Да\'\nПерезаписать - \'Нет\'", File), _
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
        InfoBox "Выгрузка выполнена в файл\n%s\n\nРазмер файла: %d байт", _
            File, FileLen(File)
        ExportToFile = True
    Else
        StopBox "Ошибка выгрузки!"
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
        If BrowseForFiles(Files, "Файлы загрузки (*.plt),*.plt,Файлы бухгалтерии (*.dbf),*.dbf,Файлы 1C:Windows (*.txt),*.txt", _
            "файл(ы) для загрузки") Then
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
    Dim s As String, i As Long, ss() As String
    On Error Resume Next
    Select Case LCase(FileExt(File))
        Case "plt"
            ImportFrom1File = ImportFrom1PLTFile(File)
        Case "dbf"
            ImportFrom1File = ImportFrom1DBFFile(File)
        Case "txt"
            ImportFrom1File = ImportFrom1CFile(File)
    End Select
End Function

Public Function ImportFrom1PLTFile(File As String) As Boolean
    Dim s As String, i As Long, ss() As String
    Dim xReply As Variant 'VbMsgBoxResult 'comp. for Office97
    On Error Resume Next
    s = InputFile(File)
    If Len(s) > 0 Then
        If Not StrToBool(App.Options("DontExportDos")) Then s = CWin(s)
        StrToLines s, ss
        'Application.ScreenUpdating = False
        mCountLastSelected = 0
        mMoneyLastSelected = 0
        mRow = Me.RowsCount
        For i = 1 To UBound(ss)
            s = ss(i)
            If InStrCount(s, TERM) = ColDetails Then
                mCountLastSelected = mCountLastSelected + 1
                Me.AsString = ss(mCountLastSelected)
                mRow = mRow + 1
                WriteRow mRow
                mMoneyLastSelected = mMoneyLastSelected + Me.Sum
                Application.StatusBar = Bsprintf("Загружено: %d", mCountLastSelected)
            End If
        Next
        'Application.ScreenUpdating = True
        Application.StatusBar = False
    
        xReply = MsgBox(Bsprintf("Завершена загрузка из файла\n%s\n\nДокументов: %d\nОбщая сумма: %f\n\nУдалить его?", _
            File, mCountLastSelected, mMoneyLastSelected), vbQuestion + vbYesNoCancel, App.Title)
        Select Case xReply
            Case vbYes: Kill File
            Case vbNo:
            Case Else
                Exit Function
        End Select
    Else
        xReply = MsgBox(Bsprintf("Ошибка загрузки из файла\n%s\n\nУдалить его?", File), vbQuestion + vbYesNoCancel + vbDefaultButton2, App.Title)
        Select Case xReply
            Case vbYes: Kill File
            Case vbNo:
            Case Else
                Exit Function
        End Select
    End If
End Function

Public Function ImportFrom1DBFFile(File As String) As Boolean
    Dim xReply As Variant 'VbMsgBoxResult 'comp. for Office97
    On Error Resume Next
    
    ReadDbfFile File, Me.RowsCount + 1
    ImportFrom1DBFFile = True

    'xReply = MsgBox(Bsprintf("Завершена загрузка из файла\n%s\n\nДокументов: %d\nОбщая сумма: %f\n\nУдалить его?", _
    '    File, mCountLastSelected, mMoneyLastSelected), vbQuestion + vbYesNoCancel, App.Title)
    xReply = MsgBox(Bsprintf("Завершена загрузка из файла\n%s\n\nУдалить его?", _
        File), vbQuestion + vbYesNoCancel, App.Title)
    Select Case xReply
        Case vbYes: Kill File
        Case vbNo:
        Case Else
            Exit Function
    End Select
End Function

Public Function ImportFrom1CFile(File As String) As Boolean
    Dim s As String, File2 As String, i As Long, v As String ', ss() As String
    Dim xReply As Variant 'VbMsgBoxResult 'comp. for Office97
    On Error Resume Next
    
    Const Format1C = "1CClientBankExchange"
    Const SecBegin = "СекцияДокумент=Платежное поручение"
    Const SecEnd = "КонецДокумента"
    Const Sec1C = "TMP1C"
    
    ImportFrom1CFile = False
    s = InputFile(File)
    'If Len(s) > 0 Then
    '    If Not StrToBool(App.Options("DontExportDos")) Then s = CWin(s)
    '    StrToLines s, ss
    '    'Application.ScreenUpdating = False
    If Left(s, Len(Format1C)) = Format1C Then
        's = "[" & Format1C & "]" & Mid(s, Len(Format1C) + 1)
        mCountLastSelected = 0
        mMoneyLastSelected = 0
        mRow = Me.RowsCount
        Do
            i = InStr(s, SecBegin)
            If i = 0 Then Exit Do
            s = "[" & Sec1C & "]" & Mid(s, i + Len(SecBegin))
            
            File2 = GetWinTempFile(Sec1C)
            'Kill File2
            OutputFile File2, s
            v = ReadIniFile(File2, Sec1C, "ПлательщикСчет")
            If v <> User.LS Then
                If Not YesNoBox("Это не Ваш счет %s!\nПродолжать?", v) Then
                    Kill File2
                    Exit Function
                End If
            End If
        
        '    For i = 1 To UBound(ss)
        '        s = ss(i)
        '        If InStrCount(s, TERM) = ColDetails Then
                    mCountLastSelected = mCountLastSelected + 1
        '            Me.AsString = ss(mCountLastSelected)
                    With Me
                        .DocNo = NumIniFile(File2, Sec1C, "Номер", 1)
                        .DocDate = RDate(ReadIniFile(File2, Sec1C, "Дата", Now))
                        .Sum = RVal(ReadIniFile(File2, Sec1C, "Сумма"))
                        .Name = ReadIniFile(File2, Sec1C, "Получатель1")
                        .INN = ReadIniFile(File2, Sec1C, "ПолучательИНН", "нет")
                        .BIC = ReadIniFile(File2, Sec1C, "ПолучательБИК")
                        .LS = ReadIniFile(File2, Sec1C, "ПолучательРасчСчет")
                        .Queue = NumIniFile(File2, Sec1C, "Очередность", 6)
                        .Details = ReadIniFile(File2, Sec1C, "НазначениеПлатежа")
                        Kill File2
                    End With
                    mRow = mRow + 1
                    WriteRow mRow
                    mMoneyLastSelected = mMoneyLastSelected + Me.Sum
                    Application.StatusBar = Bsprintf("Загружено: %d", mCountLastSelected)
    '        End If
    '    Next
        Loop
    '    'Application.ScreenUpdating = True
    '    Application.StatusBar = False
    '
        xReply = MsgBox(Bsprintf("Завершена загрузка из файла\n%s\n\nДокументов: %d\nОбщая сумма: %f\n\nУдалить его?", _
            File, mCountLastSelected, mMoneyLastSelected), vbQuestion + vbYesNoCancel, App.Title)
        Select Case xReply
            Case vbYes: Kill File
            Case vbNo:
            Case Else
                Exit Function
        End Select
    Else
        xReply = MsgBox(Bsprintf("Ошибка загрузки из файла\n%s\n\nУдалить его?", File), vbQuestion + vbYesNoCancel + vbDefaultButton2, App.Title)
        Select Case xReply
            Case vbYes: Kill File
            Case vbNo:
            Case Else
                Exit Function
        End Select
    End If
    ImportFrom1CFile = True
End Function

Public Function SavePlat(File As String) As Boolean
    SavePlat = False
    If OverwriteFile(File) Then
        OutputFile File, CDos(Me.AsPlatINI)
        SavePlat = IsFile(File)
    End If
End Function

Public Function LoadPlat(File As String) As Boolean
    Const Section1 = "Payment"
    Const Section2 = "To"
    On Error Resume Next
    mFileName = FileNameOnly(File)
    mDocNo = NumIniFile(mFileName, Section1, "No", 1)
    mDocDate = RDate(ReadIniFile(mFileName, Section1, "Date", Now))
    mSum = RVal(ReadIniFile(mFileName, Section1, "Sum", "0-00"))
    mQueue = NumIniFile(mFileName, Section1, "Queue", 6)
    mDetails = Trim(ReadIniFile(mFileName, Section1, "Text1") & " " & _
        ReadIniFile(mFileName, Section1, "Text2") & " " & _
        ReadIniFile(mFileName, Section1, "Text3") & " " & _
        ReadIniFile(mFileName, Section1, "Text4"))
    mName = ReadIniFile(mFileName, Section2, "Name2")
    mINN = ReadIniFile(mFileName, Section2, "INN2")
    mLS = ReadIniFile(mFileName, Section2, "LS2")
    mBIC = ReadIniFile(mFileName, Section2, "BIC2")
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
        PGP.ID = User.ID
        If PGP.Encode(File1, File2) Then
            If Not StrToBool(App.Options("DontMarkArchive")) Then _
                MarkByFile mFileName, "Mail", 1  'Blue
            InfoBox "Файл %s готов", File
        Else
            StopBox "Шифрованный файл НЕ создан!"
        End If
        Kill File1
    Else
        StopBox "Файл для шифрования НЕ создан!"
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
        mBIC = PadL(.Cells(mRow, ColBIC), 9, "0")
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
    Worksheets(User.ID).Cells.Find(What:=s, after:=ActiveCell, _
        LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, _
        SearchDirection:=xlNext, MatchCase:=False).Activate
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

