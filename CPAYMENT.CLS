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
Const NALT = ";"

Dim mRow As Long

Dim mLSFrom As String

Dim mFileName As String
Dim mMark As String
Dim mDocNo As Long
Dim mDocDate As Date
Dim mSum As Currency
Dim mName As String
Dim mINN As String
Dim mKPP As String
Dim mBIC As String
Dim mLS As String
Dim mQueue As Long
Dim mDetails As String
Dim mSS As String
Dim mNAL(1 To 7) As String

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
Const CellRange = "$A$1:$AJ$39"

'Подгон печатного листа под квадратные ячейки 5мм
Const CellWidth = 1.8
Const CellHeight = 14.5

'Ячейки данных на бланке платежки
Const CellNo = "M4"
Const CellDate = "R4"
Const CellDelivery = "Z4"
Const CellSS = "AI3"
Const CellAmount = "E6"
Const CellSum = "X9"
Const CellUserINN = "C9"
Const CellUserKPP = "M9"
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
Const CellKPP = "M21"
Const CellName = "A22"
Const CellAction = "AF24"
Const CellQueue = "AF25"
Const CellNAL1 = "A27"
Const CellNAL2 = "J27"
Const CellNAL3 = "P27"
Const CellNAL4 = "R27"
Const CellNAL5 = "W27"
Const CellNAL6 = "AD27"
Const CellNAL7 = "AI27"
Const CellDetails = "A28"
Const CellSign1 = "K36"
Const CellSign2 = "K39"

Private Sub Class_Initialize()
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
    Dim i1 As Long, i2 As Long
    If Left(vNewValue, 4) = "ИНН " Then '1C
        vNewValue = Trim(Mid(vNewValue, 5))
        i1 = InStr(vNewValue, " ")
        i2 = InStr(vNewValue, "\")
        If i2 = 0 Or i2 > i1 Then 'ИНН 7812345678 ООО ...
            mINN = Left(vNewValue, i1 - 1)
            If IsDigital(mINN) Then
                vNewValue = Trim(Mid(vNewValue, i1 + 1))
            Else
                mINN = vbNullString
            End If
        End If
    End If
    If Left(vNewValue, 1) = "\" Then 'ИНН 7812345678\781234567 ООО ... 1C
        vNewValue = Trim(Mid(vNewValue, 2))
        i1 = InStr(vNewValue, " ")
        mKPP = Left(vNewValue, i1 - 1)
        If IsDigital(mKPP) Then
            vNewValue = Trim(Mid(vNewValue, i1 + 1))
        Else
            mKPP = vbNullString
        End If
    ElseIf Left(vNewValue, 4) = "КПП " Then 'КПП 781234567 ООО ...
        vNewValue = Trim(Mid(vNewValue, 5))
        i1 = InStr(vNewValue, " ")
        If i1 > 0 Then
            mKPP = Left(vNewValue, i1 - 1)
            If IsDigital(mKPP) Then
                vNewValue = Trim(Mid(vNewValue, i1 + 1))
            Else
                mKPP = vbNullString
            End If
        'Else 'ООО"НетПробеловНигде"
        End If
    Else 'ООО ...
        'mINN = vbNullString
        'mKPP = vbNullString
    End If
    mName = vNewValue
End Property

Public Property Get INN() As String
    INN = mINN
End Property

Public Property Let INN(ByVal vNewValue As String)
    mINN = vNewValue
End Property

Public Property Get KPP() As String
    KPP = mKPP
End Property

Public Property Let KPP(ByVal vNewValue As String)
    mKPP = vNewValue
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
    If Len(vNewValue) = 9 Then
        mBIC = vNewValue
    ElseIf Len(vNewValue) = 8 Then
        mBIC = "0" & vNewValue
    Else
        mBIC = vbNullString
    End If
End Property

Public Property Get Bank() As String
    BnkSeek2.BIC = mBIC
    Bank = BnkSeek2.Name
End Property

Public Property Get Place() As String
    BnkSeek2.BIC = mBIC
    Place = BnkSeek2.Place
End Property

Public Property Get KS() As String
    BnkSeek2.BIC = mBIC
    KS = BnkSeek2.KS
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

Public Property Get SS() As String
    SS = mSS
End Property

Public Property Let SS(ByVal vNewValue As String)
    Dim i As Long
    If Len(vNewValue) = 2 Then
        mSS = vNewValue
    Else
        mSS = vbNullString
        For i = LBound(mNAL) To UBound(mNAL)
            mNAL(i) = vbNullString
        Next
    End If
End Property

Public Property Get NAL(i As Long) As String
    If Len(mSS) = 0 Then
        NAL = vbNullString
    Else
        If mNAL(i) = vbNullString Then
            NAL = "0"
        Else
            NAL = mNAL(i)
        End If
    End If
End Property

Public Property Let NAL(i As Long, ByVal vNewValue As String)
    mNAL(i) = vNewValue
End Property

Public Property Get FileName() As String
    'BnkSeek2.BIC = mBIC
    'FileName = Bsprintf("%s%X%X%02d%03d", BnkSeek2.NamePost, _
    '    Year(mDocDate) - 2000, Month(mDocDate), Day(mDocDate), mDocNo)
    FileName = Bsprintf("A%X%X%02d%03d", _
        Year(mDocDate) - 2000, Month(mDocDate), Day(mDocDate), mDocNo)
    If Len(mFileName) = 0 Then mFileName = FileName
End Property

Public Property Get Mark() As String
    Mark = mMark
End Property

Public Property Let Mark(ByVal vNewValue As String)
    mMark = vNewValue
End Property

Public Sub FillBlank()
    Dim r As Range
    On Error Resume Next
    Application.ScreenUpdating = False
    Set r = Worksheets(PaySheet).Range(BaseRange)
    r.Range(CellUserName) = StrTran(User.Name, WRAP, vbLf)
    r.Range(CellUserINN) = User.INN
    r.Range(CellUserKPP) = User.KPP
    r.Range(CellUserLS) = User.LS
    r.Range(CellUserBIC) = User.BIC
    r.Range(CellUserBank) = User.Bank
    r.Range(CellUserPlace) = User.Place
    r.Range(CellUserKS) = User.KS
    'r.Range(CellSign1) = User.Sign(1)
    'r.Range(CellSign2) = User.Sign(2)
        
    r.Range(CellNo) = vbNullString
    r.Range(CellDate) = PlatDate(Now)
    r.Range(CellDelivery) = vbNullString
    r.Range(CellAmount) = vbNullString
    r.Range(CellSum) = vbNullString
    r.Range(CellName) = vbNullString
    r.Range(CellINN) = vbNullString
    r.Range(CellKPP) = vbNullString
    r.Range(CellBIC) = vbNullString
    r.Range(CellLS) = vbNullString
    r.Range(CellBank) = vbNullString
    r.Range(CellPlace) = vbNullString
    r.Range(CellKS) = vbNullString
    r.Range(CellQueue) = vbNullString
    r.Range(CellDetails) = vbNullString
    r.Range(CellSS) = vbNullString
    r.Range(CellNAL1) = vbNullString
    r.Range(CellNAL2) = vbNullString
    r.Range(CellNAL3) = vbNullString
    r.Range(CellNAL4) = vbNullString
    r.Range(CellNAL5) = vbNullString
    r.Range(CellNAL6) = vbNullString
    r.Range(CellNAL7) = vbNullString
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
        r.Range(CellUserKPP) = .KPP
        r.Range(CellUserLS) = .LS
        r.Range(CellUserBIC) = .BIC
        r.Range(CellUserBank) = .Bank
        r.Range(CellUserPlace) = .Place
        r.Range(CellUserKS) = .KS
        'r.Range(CellSign1) = .Sign(1)
        'r.Range(CellSign2) = .Sign(2)
    End With
    With Me
        r.Range(CellNo) = .DocNo
        r.Range(CellDate) = PlatDate(.DocDate)
        r.Range(CellAmount) = RSumStr(.Sum)
        r.Range(CellSum) = PlatFormat(.Sum)
        r.Range(CellName) = StrTran(.Name, WRAP, vbLf)
        r.Range(CellINN) = .INN
        r.Range(CellKPP) = .KPP
        r.Range(CellBIC) = .BIC
        r.Range(CellLS) = .LS
        r.Range(CellQueue) = .Queue
        r.Range(CellDetails) = StrTran(.Details, WRAP, vbLf)
        r.Range(CellSS) = .SS
        r.Range(CellNAL1) = .NAL(1)
        r.Range(CellNAL2) = .NAL(2)
        r.Range(CellNAL3) = .NAL(3)
        r.Range(CellNAL4) = .NAL(4)
        r.Range(CellNAL5) = .NAL(5)
        r.Range(CellNAL6) = .NAL(6)
        r.Range(CellNAL7) = .NAL(7)
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

Public Function MarkByFile(File As String, Mark As String, Optional Color As Long = 0) As Boolean
    Dim r As Range, s As String, File2 As String
    Dim n As Long, ws As Worksheet
    MarkByFile = False
    On Error Resume Next
    File2 = GetWinTempFile
    If Crypto.Decrypt(File, File2) Then
        LoadPlat File2
        Kill File2
    Else
        WarnBox "Файл %s не расшифровать!", FileNameExt(File)
        If Not LoadPlat(File) Then
            WarnBox "Файл %s не прочитать!", FileNameExt(File)
            Exit Function
        End If
    End If
    If Len(mLSFrom) <> 20 Then Exit Function
    For Each ws In ActiveWorkbook.Worksheets
        s = ws.Name
        If IsDigital(s) Then
            If App.DefLS(s) = mLSFrom Then
                Set r = ws.Range(BaseRange)
                n = 1
                Do While Len(r.Cells(n, ColDocNo).Text) > 0 'looking for the last used row
                    n = n + 1
                Loop
                mRow = n - 1
                Do While mRow > 1
                    If r.Cells(mRow, ColDocNo) = mDocNo Then
                        If r.Cells(mRow, ColDocDate) = mDocDate Then
                            If r.Cells(mRow, ColSum) = mSum Then
                                n = mRow
                                Exit Do
                            End If
                        End If
                    End If
                    mRow = mRow - 1
                Loop
                'WriteRow N, s
                'MarkByRow N, Mark, Color
                
                If LCase(r.Cells(n, ColMark).Text) = "ok" Then
                    WarnBox "Файл %s уже был принят!", FileNameExt(File)
                Else
                    r.Cells(n, ColMark) = Mark
                    r.Cells(n, ColMark).Font.Color = QBColor(Color)
                    r.Cells(n, ColDocNo) = mDocNo
                    r.Cells(n, ColDocDate) = mDocDate
                    r.Cells(n, ColSum) = mSum
                    r.Cells(n, ColName) = KppAndName()
                    r.Cells(n, ColINN) = mINN
                    r.Cells(n, ColBIC) = mBIC
                    r.Cells(n, ColLS) = mLS
                    r.Cells(n, ColQueue) = mQueue
                    r.Cells(n, ColDetails) = NalAndDetails()
                End If
                
                MarkByFile = True
                Exit Function
            End If
        End If
    Next
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
    If App.BoolOptions("DontPreviewPlat") Then
        Application.GoTo Worksheets(PaySheet).Range("$A$1"), True
    Else
        Worksheets(PaySheet).PrintPreview
    End If
End Sub

Public Function AskAllRows(Optional Ask As String = "Обработать все?") As Boolean
    Dim a As Range, r As Range, s As String
    mCountLastSelected = 0
    mMoneyLastSelected = 0
    With Worksheets(User.ID)
        For Each a In Selection.Areas
            For Each r In a.Rows
                If Len(.Cells(r.Row, ColDocNo).Text) > 0 Then
                    If r.Row > 1 Then
                        mCountLastSelected = mCountLastSelected + 1
                        mMoneyLastSelected = mMoneyLastSelected + .Cells(r.Row, ColSum)
                    End If
                End If
            Next r
        Next a
    End With
    'If mCountLastSelected = 1 Then
    '    'Ask = BSPrintF("Помечен 1 документ\nна сумму %f", mMoneyLastSelected)
    '    AskAllRows = True
    '    Exit Function
    'Else
    If mMoneyLastSelected = 0 Then
        Ask = Bsprintf("Помечено документов: %d\nНО ОБЩАЯ СУММА НУЛЕВАЯ!\n\n%s", mCountLastSelected, Ask)
    Else
        Ask = Bsprintf("Помечено документов: %d\nНа общую сумму: %f\n\n%s", mCountLastSelected, mMoneyLastSelected, Ask)
    End If
    With User
        AskAllRows = OkCancelBox("Плательщик %s\n%s\nИНН %s / КПП %s\n\n%s\n\n%s", _
            .LS, .Name, .INN, .KPP, Ask, "Перед обработкой файл будет автосохранен!")
    End With
End Function

Public Function EachSelected(Action As String, Optional Ask As String = "Обработать все?") As Boolean
    Dim a As Range, r As Range
    Dim s As String, nOk As Long, nErr As Long
    EachSelected = False
    Select Case Action
        Case "Preview":
            For Each a In Selection.Areas
                For Each r In a.Rows
                    If r.Row > 1 Then
                        Me.ReadRow r.Row
                        PrintPreview
                    End If
                Next r
            Next a
        Case "ExportPlat":
            s = User.ValidationError
            If Len(s) > 0 Then
                StopBox "Исправьте ошибки в Ваших реквизитах плательщика:\n\n%s!", s
                Exit Function
            End If
            If AskAllRows(Ask) Then
                ActiveWorkbook.Save
            Else
                Exit Function
            End If
            nOk = 0
            nErr = 0
            For Each a In Selection.Areas
                For Each r In a.Rows
                    If r.Row > 1 Then
                        With Me
                            .ReadRow r.Row
                            s = LCase(.Mark)
                            If s = "mail" Then
                                s = "Уже зашифрован - отправьте его."
                            ElseIf s = "ok" Then
                                s = "Уже принят Банком."
                            Else
                                s = .ValidationError
                            End If
                            If Len(s) = 0 Then
                                If SaveCryptedPlat Then
                                    nOk = nOk + 1
                                Else
                                    nErr = nErr + 1
                                End If
                            Else
                                nErr = nErr + 1
                                If YesNoBox("Документ N %d на сумму %f\nв строке %d не может быть правильно обработан:\n\n%s!\n\nПрервать процесс?", _
                                    .DocNo, .Sum, .Row, s) Then
                                    WarnBox "Обработка помеченных документов прервана.\nУспешных: %d\nОшибочных: %d", nOk, nErr
                                    Exit Function
                                End If
                            End If
                        End With
                    End If
                Next r
            Next a
            InfoBox "Обработка помеченных документов выполнена.\nУспешных: %d\nОшибочных: %d", nOk, nErr
        Case "Delete":
            If OkCancelBox("Безвозвратно удалить?!") Then
                For Each a In Selection.Areas
                    For Each r In a.Rows
                        If r.Row = 1 Then
                            StopBox "Нельзя удалять заголовок таблицы!"
                            Exit Function
                        End If
                    Next r
                    a.Rows.EntireRow.Delete
                Next a
            Else
                Exit Function
            End If
        Case Else
            '?
    End Select
    EachSelected = True
End Function

Public Property Get AsString() As String
    AsString = mFileName & TERM & _
        mMark & TERM & _
        mDocNo & TERM & _
        DtoC(mDocDate) & TERM & _
        PlatFormat(mSum) & TERM & _
        KppAndName & TERM & _
        mINN & TERM & _
        mBIC & TERM & _
        mLS & TERM & _
        mQueue & TERM & _
        NalAndDetails & TERM
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
    mKPP = vbNullString
    Me.Name = Arr(ColName) 'Auto KPP and Name
    mINN = Arr(ColINN)
    mBIC = Arr(ColBIC)
    mLS = Arr(ColLS)
    mQueue = Arr(ColQueue)
    Me.SS = vbNullString 'PreNulling!!!
    Me.Details = Arr(ColDetails) 'Auto SS & NAL1-7 later
End Property

Public Property Get AsPlatINI() As String
    Dim s As String, SS(4) As String, i As Long
    On Error Resume Next
    s = Bsprintf("; %s\n[File]\n", App.TITLE)
    s = s & Bsprintf("File=%s.%s\n", Me.FileName, User.ID4)
    s = s & Bsprintf("Now=%T\n", Now)
    s = s & Bsprintf("Form=2\nVer=11\nCP=866\nVersion=%s\n", App.Version)
    s = s & Bsprintf("Path=%s\n", App.Path)
    
    s = s & Bsprintf("\n[Payment]\n")
    s = s & Bsprintf("No=%d\n", Me.DocNo)
    s = s & Bsprintf("Date=%n\n", Me.DocDate) 'dd.mm.yyyy
    s = s & Bsprintf("Sum=%F\n", Me.Sum) '0-00
    
    s = s & Bsprintf("Queue=%d\n", Me.Queue)
    
    WrapStrToArr Me.Details, SS, 64
    s = s & Bsprintf("Text1=%s\n", SS(1))
    s = s & Bsprintf("Text2=%s\n", SS(2))
    s = s & Bsprintf("Text3=%s\n", SS(3))
    s = s & Bsprintf("Text4=%s\n", SS(4))
    
    s = s & Bsprintf("\n[From]\n")
    s = s & Bsprintf("Name=%s\n", StrSpaces1(User.Name))
    s = s & Bsprintf("INN=%s\n", User.INN)
    s = s & Bsprintf("KPP=%s\n", User.KPP)
    s = s & Bsprintf("LS=%s\n", User.LS)
    
    s = s & Bsprintf("\n[To]\n")
    s = s & Bsprintf("Name2=%s\n", Me.Name)
    s = s & Bsprintf("INN2=%s\n", Me.INN)
    s = s & Bsprintf("KPP2=%s\n", Me.KPP)
    s = s & Bsprintf("LS2=%s\n", Me.LS)
    s = s & Bsprintf("BIC2=%s\n", Me.BIC)
    
    's = s & Bsprintf("Bank2=%s\n", Me.Bank)
    's = s & Bsprintf("Place2=%s\n", Me.Place)
    's = s & Bsprintf("KS2=%s\n", Me.KS)
    
    If Len(Me.SS) = 2 Then
        s = s & Bsprintf("SS=%s\n", Me.SS)
        For i = LBound(mNAL) To UBound(mNAL)
            s = s & Bsprintf("NAL%d=%s\n", i, Me.NAL(i))
        Next
    End If
    
    AsPlatINI = s & Bsprintf("\n; eof\n")
End Property

Public Property Get AsED101(Optional No As Long = 1) As String
    Dim s As String
    On Error Resume Next
    
'<?xml version="1.0" encoding="WINDOWS-1251"?>
'<PacketEPD xmlns="urn:cbr-ru:ed:v1.0" EDNo="1" EDDate="2003-04-14" EDAuthor="4525545000" EDQuantity="2" Sum="1">
'   <ED101 xmlns="urn:cbr-ru:ed:v1.0" EDNo="7" EDDate="2003-04-14" EDAuthor="4525545000" PaytKind="1" Sum="2400000" TurnoverKind="1" TransKind="01" ChargeOffDate="2003-04-14" Priority="6" ReceiptDate="2003-04-14">
'      <AccDoc AccDocNo="4" AccDocDate="2003-04-14"></AccDoc>
'      <Payer INN="7726274727" PersonalAcc="40702810200203001037">
'          <Name>ООО ТЕХНО ПЛЮС</Name>
'          <Bank BIC="044525545" CorrespAcc="30101810300000000545"></Bank>
'      </Payer>
'      <Payee INN="7726062105" PersonalAcc="40702810010130010079">
'          <Name>ООО ТД ТОРНАДО-ПРОДУКТ</Name>
'          <Bank BIC="044525219" CorrespAcc="30101810500000000219"></Bank>
'      </Payee>
'      <Purpose>ОПЛАТА ПО ДОГОВОРУ 95456 ОТ 15.01.2003 В ТОМ ЧИСЛЕ НДС 4000 РУБ</Purpose>
'   </ED101>
'</PacketEPD>
    
    s = Bsprintf("<ED101 xmlns=\'urn:cbr-ru:ed:v1.0\' ")
    s = s & Bsprintf("EDNo=\'%d\' EDDate=\'%m\' ", No, Now)
    s = s & Bsprintf("EDAuthor=\'%s\' ", User.ID) '"4030702000"
    's = s & Bsprintf("PaytKind=\'1\' ")
    s = s & Bsprintf("Sum=\'%d\' ", Me.Sum * 100)
    's = s & Bsprintf("TurnoverKind=\'1\' ")
    's = s & Bsprintf("TransKind=\'01\' ")
    's = s & Bsprintf("ChargeOffDate=\'%m\' ", Now)
    s = s & Bsprintf("Priority=\'%d\'", Me.Queue)
    's = s & Bsprintf("ReceiptDate=\'%m\'", Now)
    s = s & Bsprintf(">\n")
    
    's = s & Bsprintf("\t<AccDoc AccDocNo=\'%d\' AccDocDate=\'%m\'></AccDoc>\n", Me.DocNo, Me.DocDate)
    s = s & Bsprintf("\t<AccDoc AccDocNo=\'%d\' AccDocDate=\'%m\'/>\n", Me.DocNo, Me.DocDate)
    
    With User
        s = s & Bsprintf("\t<Payer INN=\'%s\' PersonalAcc=\'%s\'", .INN, .LS)
        If Len(.KPP) > 1 Then s = s & Bsprintf(" KPP=\'%s\'", .KPP)
        s = s & Bsprintf(">\n\t\t<Name>%s</Name>\n", StrXMLencode(.Name))
    
        s = s & Bsprintf("\t\t<Bank BIC=\'%s\'", .BIC)
        If Len(.KS) = 20 Then s = s & Bsprintf(" CorrespAcc=\'%s\'", .KS)
        's = s & Bsprintf("></Bank>\n\t</Payer>\n")
        s = s & Bsprintf("/></Payer>\n")
    End With

    With Me
        s = s & Bsprintf("\t<Payee INN=\'%s\' PersonalAcc=\'%s\'", .INN, .LS)
        If Len(.KPP) > 1 Then s = s & Bsprintf(" KPP=\'%s\'", .KPP)
        s = s & Bsprintf(">\n\t\t<Name>%s</Name>\n", StrXMLencode(.Name))
    
        s = s & Bsprintf("\t\t<Bank BIC=\'%s\'", .BIC)
        If Len(.KS) = 20 Then s = s & Bsprintf(" CorrespAcc=\'%s\'", .KS)
        's = s & Bsprintf("></Bank>\n\t</Payee>\n")
        s = s & Bsprintf("/></Payee>\n")

        s = s & Bsprintf("\t<Purpose>%s</Purpose>\n", StrXMLencode(.Details))
        
        If Len(.SS) > 0 Then
            s = s & Bsprintf("\t<DepartmentalInfo DraverStatus=\'%s\' ", .SS)
            s = s & Bsprintf("CBC=\'%s\' OKATO=\'%s\' ", .NAL(1), .NAL(2))
            s = s & Bsprintf("PaytReason=\'%s\' TaxPeriod=\'%s\' ", .NAL(3), .NAL(4))
            s = s & Bsprintf("DocNo=\'%s\' DocDate=\'%s\' TaxPaytKind=\'%s\'/>\n", .NAL(5), .NAL(6), .NAL(7))
            's = s & Bsprintf("</DepartmentalInfo>\n")
        End If
    End With
    
    AsED101 = s & Bsprintf("</ED101>\n")
End Property

Public Function ExportToFile(Optional File As String = vbNullString) As Boolean
    Dim a As Range, r As Range, s As String, x As Integer
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
            If Not BrowseForSave(File, "Файлы выгрузки (*.plt),*.plt,Файлы УФЭБС (*.xml),*.xml", _
                "файл для выгрузки") Then Exit Function
            User.ExportList = FilePath(File)
        'End If
    'End If
    s = vbNullString
    mCountLastSelected = 0
    Select Case LCase(FileExt(File))
        Case "plt":
            For Each a In Selection.Areas
                For Each r In a.Rows
                    With Me
                        .ReadRow r.Row
                        If mDocNo > 0 Then
                            .MarkByRow r.Row, "Disk", 1 'Blue
                            s = s & Me.AsString & vbCrLf
                            mCountLastSelected = mCountLastSelected + 1
                            Application.StatusBar = Bsprintf("Подготовлено: %d", mCountLastSelected)
                        End If
                    End With
                Next r
            Next a
            If Not App.BoolOptions("DontExportDos") Then s = CDos(s)
            Application.StatusBar = False
            If Len(s) = 0 Then
                WarnBox "Нечего выгружать!"
                Exit Function
            End If
            If IsFile(File) Then
                x = MsgBox(Bsprintf("Файл %s уже существует!\nЧтобы добавить, нажмите \'Да\'\nПерезаписать - \'Нет\'", File), _
                    vbQuestion + vbYesNoCancel, App.TITLE)
                Select Case x
                    Case vbyes:
                        AppendFile File, s
                    Case vbNo:
                        OutputFile File, s
                    Case Else
                        MsgBox "Выгрузка НЕ выполнена!", vbExclamation, App.TITLE
                        Exit Function
                End Select
            Else
                OutputFile File, s
            End If

        Case "xml":
            For Each a In Selection.Areas
                For Each r In a.Rows
                    With Me
                        .ReadRow r.Row
                        If mDocNo > 0 Then
                            .MarkByRow r.Row, "Disk", 1 'Blue
                            s = s & Me.AsED101(mCountLastSelected + 1)
                            mCountLastSelected = mCountLastSelected + 1
                            Application.StatusBar = Bsprintf("Подготовлено: %d", mCountLastSelected)
                        End If
                    End With
                Next r
            Next a
            Application.StatusBar = False
            If Len(s) = 0 Then
                WarnBox "Нечего выгружать!"
                Exit Function
            End If
            x = Val(App.Options("EDNo")) + 1
            If x > 999999 Then x = 1
            App.Options("EDNo") = CStr(x)
            s = Bsprintf("<?xml version=\'1.0\' encoding=\'WINDOWS-1251\'?>\n" & _
                "<PacketEPD xmlns=\'urn:cbr-ru:ed:v1.0\' " & _
                "EDNo=\'%d\' EDDate=\'%m\' EDAuthor=\'%s\' " & _
                "EDReceiver=\'4030702000\' EDQuantity=\'%d\' Sum=\'%d\'>\n", _
                x, Now, User.ID, mCountLastSelected, mMoneyLastSelected * 100) & _
                s & Bsprintf("</PacketEPD>\n")
            OutputFile File, s

        Case Else
            WarnBox "Не указан формат вывода!"
            Exit Function
    End Select
    If IsFile(File) Then
        InfoBox "Выгрузка выполнена в файл\n%s\n\nРазмер файла: %d байт", _
            File, FileLen(File)
        ExportToFile = True
    Else
        StopBox "Ошибка выгрузки!"
    End If
End Function

Public Function ImportFromFile(Optional File As String = vbNullString) As Boolean
    Dim Files As Variant, i As Long, s As String
    On Error Resume Next
    If IsFile(File) Then
        ImportFromFile = ImportFrom1File(File)
    Else
        Files = User.ImportList
        s = User.ID4
        i = Val(App.Options("ImportIndex"))
        If i = 0 Then i = 1
        If BrowseForFiles(Files, "Файлы загрузки (*.plt),*.plt,Файлы бухгалтерии (*.dbf),*.dbf,Файлы 1C:Windows (*.txt),*.txt," & _
            Bsprintf("Платежки клиента %s (*.%s),*.%s", s, s, s), _
            "файл(ы) для загрузки", i) Then
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
    If IsFile(File) Then
        Select Case LCase(FileExt(File))
            Case "plt"
                App.Options("ImportIndex") = 1
                ImportFrom1File = ImportFrom1PLTFile(File)
            Case "dbf"
                App.Options("ImportIndex") = 2
                ImportFrom1File = ImportFrom1DBFFile(File)
            Case "txt"
                App.Options("ImportIndex") = 3
                ImportFrom1File = ImportFrom1CFile(File)
            Case User.ID4
                App.Options("ImportIndex") = 4
                ImportFrom1File = ImportFrom1PlatFile(File)
        End Select
    End If
End Function

Public Function ImportFrom1PLTFile(File As String) As Boolean
    Dim s As String, i As Long, ArrS() As String
    Dim xReply As Variant 'VbMsgBoxResult 'comp. for Office97
    On Error Resume Next
    s = InputFile(File)
    If Len(s) > 0 Then
        If Not StrToBool(App.Options("DontExportDos")) Then s = CWin(s)
        StrToLines s, ArrS
        'Application.ScreenUpdating = False
        mCountLastSelected = 0
        mMoneyLastSelected = 0
        mRow = Me.RowsCount
        For i = 1 To UBound(ArrS)
            s = ArrS(i)
            If InStrCount(s, TERM) = ColDetails Then
                mCountLastSelected = mCountLastSelected + 1
                Me.AsString = ArrS(mCountLastSelected)
                mRow = mRow + 1
                WriteRow mRow
                mMoneyLastSelected = mMoneyLastSelected + Me.Sum
                Application.StatusBar = Bsprintf("Загружено: %d", mCountLastSelected)
            End If
        Next
        'Application.ScreenUpdating = True
        Application.StatusBar = False
    
        xReply = MsgBox(Bsprintf("Завершена загрузка из файла\n%s\n\nДокументов: %d\nОбщая сумма: %f\n\nУдалить его?", _
            File, mCountLastSelected, mMoneyLastSelected), vbQuestion + vbYesNoCancel, App.TITLE)
        Select Case xReply
            Case vbyes: Kill File
            Case vbNo:
            Case Else
                Exit Function
        End Select
    Else
        xReply = MsgBox(Bsprintf("Нечего грузить из пустого файла\n%s\n\nУдалить его?", File), vbQuestion + vbYesNoCancel + vbDefaultButton2, App.TITLE)
        Select Case xReply
            Case vbyes: Kill File
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
        File), vbQuestion + vbYesNoCancel, App.TITLE)
    Select Case xReply
        Case vbyes: Kill File
        Case vbNo:
        Case Else
            Exit Function
    End Select
End Function

Public Function ImportFrom1CFile(File As String) As Boolean
    Dim s As String, File2 As String, i As Long, n As Long, Sec As String, v As String
    Dim fi As Long, fo As Long
    Dim xReply As Variant 'VbMsgBoxResult 'comp. for Office97
    On Error Resume Next
    
    'Const Format1C = "1CClientBankExchange"
    Const SecBegin = "СекцияДокумент=Платежное поручение"
    'Const SecEnd = "КонецДокумента"
    
    ImportFrom1CFile = False
    
    fi = FreeFile
    Open File For Input Access Read Shared As #fi
    File2 = GetWinTempFile("1C")
    fo = FreeFile
    Open File2 For Output Access Write Shared As #fo
    n = 0
    Print #fo, "[Header]"
    Do While Not EOF(fi)
        Line Input #fi, s
        If s = SecBegin Then
            n = n + 1
            Print #fo, "[Doc" & n & "]"
        ElseIf InStr(s, "=") > 1 Then
            Print #fo, s
        End If
    Loop
    Close #fi
    Close #fo
    
    If n > 0 Then
        mCountLastSelected = 0
        mMoneyLastSelected = 0
        mRow = Me.RowsCount
        
        For i = 1 To n
            Sec = "Doc" & i
            v = ReadIniFile(File2, Sec, "ПлательщикСчет")
            If v <> User.LS Then
                If Not YesNoBox("Это не Ваш счет %s!\nПродолжать?", v) Then
                    Kill File2
                    Exit Function
                End If
            End If
            
            'Света со Шпалерной
            v = ReadIniFile(File2, Sec, "ПлательщикКПП", "0")
            If v <> User.KPP Then
                If YesNoBox("Из 1C загружаются документы\nс КПП %s вместо вашего %s\n\nЗаменить как в 1C?", _
                    v, User.KPP) Then
                    User.KPP = v
                    WarnBox "Не забудьте отправить эти платежи,\nпока снова не сменили КПП!\nСейчас %s", v
                End If
            End If
        
            mCountLastSelected = mCountLastSelected + 1
            With Me
                .Mark = "1C"
                .DocNo = NumIniFile(File2, Sec, "Номер", 1)
                .DocDate = RDate(ReadIniFile(File2, Sec, "Дата", Now))
                .Sum = RVal(ReadIniFile(File2, Sec, "Сумма"))
                .Name = ReadIniFile(File2, Sec, "Получатель") '"Получатель1"
                .INN = ReadIniFile(File2, Sec, "ПолучательИНН", "0") 'after Name!
                .KPP = ReadIniFile(File2, Sec, "ПолучательКПП", "0") 'after Name!
                .BIC = ReadIniFile(File2, Sec, "ПолучательБИК")
                .LS = ReadIniFile(File2, Sec, "ПолучательРасчСчет")
                If Len(.LS) = 0 Then '1Cv8
                    .LS = ReadIniFile(File2, Sec, "ПолучательСчет")
                End If
                .Queue = NumIniFile(File2, Sec, "Очередность", 6)
                .SS = ReadIniFile(File2, Sec, "СтатусСоставителя")
                If Len(.SS) > 0 Then
                    .NAL(1) = ReadIniFile(File2, Sec, "ПоказательКБК")
                    .NAL(2) = ReadIniFile(File2, Sec, "ОКАТО")
                    .NAL(3) = ReadIniFile(File2, Sec, "ПоказательОснования")
                    .NAL(4) = ReadIniFile(File2, Sec, "ПоказательПериода")
                    .NAL(5) = ReadIniFile(File2, Sec, "ПоказательНомера") '/////////???
                    .NAL(6) = ReadIniFile(File2, Sec, "ПоказательДаты")
                    .NAL(7) = ReadIniFile(File2, Sec, "ПоказательТипа")
                Else
                    .NAL(1) = vbNullString
                    .NAL(2) = vbNullString
                    .NAL(3) = vbNullString
                    .NAL(4) = vbNullString
                    .NAL(5) = vbNullString
                    .NAL(6) = vbNullString
                    .NAL(7) = vbNullString
                End If
                .Details = ReadIniFile(File2, Sec, "НазначениеПлатежа")
            End With
            mRow = mRow + 1
            WriteRow mRow
            mMoneyLastSelected = mMoneyLastSelected + Me.Sum
            Application.StatusBar = Bsprintf("Загружено: %d", mCountLastSelected)
        Next
        Kill File2
        
        xReply = MsgBox(Bsprintf("Завершена загрузка из файла\n%s\n\nДокументов: %d\nОбщая сумма: %f\n\nУдалить его?", _
            File, mCountLastSelected, mMoneyLastSelected), vbQuestion + vbYesNoCancel, App.TITLE)
        Select Case xReply
            Case vbyes: Kill File
            Case vbNo:
            Case Else
                Exit Function
        End Select
    Else
        xReply = MsgBox(Bsprintf("Ошибка загрузки из файла\n%s\n\nУдалить его?", File), vbQuestion + vbYesNoCancel + vbDefaultButton2, App.TITLE)
        Select Case xReply
            Case vbyes: Kill File
            Case vbNo:
            Case Else
                Exit Function
        End Select
    End If
    ImportFrom1CFile = True
End Function

Public Function ImportFrom1PlatFile(File As String) As Boolean
    Dim Mark As String
    On Error Resume Next
    Select Case UCase(Left(FileNameExt(File), 1))
        Case "O"
            Mark = "Ok?"
        Case "E"
            Mark = "Err?"
        Case "T"
            Mark = "Test?"
        Case "A"
            Mark = "Mail?"
        Case Else
            Mark = "?"
    End Select
    MarkByFile File, Mark
End Function

Public Function LoadPlat(File As String) As Boolean
    Dim i As Long
    Const Section1 = "Payment"
    Const Section2 = "From"
    Const Section3 = "To"
    On Error Resume Next
    'mFileName = FileNameOnly(File)
    'mMark = Mark
    mDocNo = NumIniFile(File, Section1, "No", 1)
    mDocDate = RDate(ReadIniFile(File, Section1, "Date", Now))
    mSum = RVal(ReadIniFile(File, Section1, "Sum", "0-00"))
    mQueue = NumIniFile(File, Section1, "Queue", 6)
    mDetails = CWin(Trim(ReadIniFile(File, Section1, "Text1") & " " & _
        ReadIniFile(File, Section1, "Text2") & " " & _
        ReadIniFile(File, Section1, "Text3") & " " & _
        ReadIniFile(File, Section1, "Text4")))
    
    mLSFrom = ReadIniFile(File, Section2, "LS")
    
    mName = CWin(ReadIniFile(File, Section3, "Name2"))
    mINN = ReadIniFile(File, Section3, "INN2")
    mKPP = ReadIniFile(File, Section3, "KPP2", "0")
    mLS = ReadIniFile(File, Section3, "LS2")
    mBIC = ReadIniFile(File, Section3, "BIC2")
    mSS = ReadIniFile(File, Section3, "SS")
    For i = LBound(mNAL) To UBound(mNAL)
        mNAL(i) = CWin(ReadIniFile(File, Section3, Bsprintf("NAL%d", i)))
    Next
    'WriteRow
    LoadPlat = mSum > 0 'Me.Valid
End Function

Public Function SaveCryptedPlat() As Boolean
    Dim File As String, File1 As String, File2 As String
    On Error Resume Next
    File = Me.FileName & "." & User.ID4
    File1 = RightPathName(GetWinTempDir, File)
    File2 = SMail.Send & File
    'If OverwriteFile(File1) Then
        OutputFile File1, CDos(Me.AsPlatINI)
        If Crypto.Encrypt(File1, File2) Then
            If Not App.BoolOptions("DontMarkArchive") Then
                'MarkByFile File2, "Mail", 1    'Blue
                MarkByRow Payment.Row, "Mail", 1     'Blue
            End If
            InfoBox "Файл %s готов к отправке", File
        Else
            StopBox "Шифрованный файл НЕ создан!\nОшибка использования ключей PGP?"
        End If
    'Else
    '    StopBox "К отправке оставлен прежний файл!"
    'End If
    SaveCryptedPlat = IsFile(File2)
End Function

Public Sub Delete()
    On Error Resume Next
    If mRow > 1 Then Worksheets(User.ID).Range(BaseRange).Rows(mRow).EntireRow.Delete
End Sub

Public Sub Clear()
    'mFileName = vbNullString
    mMark = vbNullString
    mDocNo = 0
    mDocDate = Now
    mSum = 0
    mName = vbNullString
    mINN = vbNullString
    mKPP = vbNullString
    mBIC = vbNullString
    mLS = vbNullString
    mQueue = 0
    mDetails = vbNullString
    Me.SS = vbNullString
End Sub

Public Sub ReadRow(Optional Row As Long)
    If Not IsMissing(Row) Then mRow = NormRow(Row)
    With Worksheets(User.ID).Range(BaseRange)
        Me.Clear
        'mFileName = .Cells(mRow, ColFileName)
        mMark = .Cells(mRow, ColMark)
        mDocNo = .Cells(mRow, ColDocNo)
        mDocDate = .Cells(mRow, ColDocDate)
        mSum = .Cells(mRow, ColSum)
        Me.Name = StrSpaces1(.Cells(mRow, ColName))
        mINN = .Cells(mRow, ColINN)
        Me.BIC = .Cells(mRow, ColBIC)
        mLS = .Cells(mRow, ColLS)
        mQueue = .Cells(mRow, ColQueue)
        ParseDetails StrSpaces1(.Cells(mRow, ColDetails))
    End With
End Sub

Public Sub WriteRow(Optional Row As Long, Optional wsName As String = vbNullString)
    If Not IsMissing(Row) Then mRow = NormRow(Row)
    If Len(wsName) = 0 Then wsName = User.ID
    With Worksheets(wsName).Range(BaseRange)
        'If Len(.Cells(mRow, ColFileName).Text) = 0 Then mFileName = FileName()
        '.Cells(mRow, ColFileName) = mFileName
        If Len(mMark) = 0 Then mMark = "?"
        .Cells(mRow, ColMark) = mMark
        .Cells(mRow, ColDocNo) = mDocNo
        .Cells(mRow, ColDocDate) = mDocDate
        .Cells(mRow, ColSum) = mSum
        .Cells(mRow, ColName) = KppAndName()
        .Cells(mRow, ColINN) = mINN
        .Cells(mRow, ColBIC) = mBIC
        .Cells(mRow, ColLS) = mLS
        .Cells(mRow, ColQueue) = mQueue
        .Cells(mRow, ColDetails) = NalAndDetails()
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
    s = InputBox("Введите часть строки для поиска:", App.TITLE, s)
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

Public Function KppAndName() As String
    KppAndName = "КПП " & mKPP & " " & mName
End Function

Public Function NalAndDetails() As String
    Dim i As Long, s As String
    If Len(mSS) = 2 Then
        s = mSS & NALT
        For i = LBound(mNAL) To UBound(mNAL)
            s = s & Me.NAL(i) & NALT
        Next
        NalAndDetails = s & mDetails
    Else
        NalAndDetails = mDetails
    End If
End Function

Private Sub ParseDetails(vNewValue As String)
    Dim Arr As Variant, i As Long
    If InStr("01", Left(vNewValue, 1)) > 0 And Mid(vNewValue, 3, 1) = NALT Then
        Arr = StrToArr(vNewValue, NALT)
        Me.SS = Arr(LBound(Arr))
        mDetails = vbNullString
        For i = LBound(Arr) + 1 To UBound(Arr)
            If i > UBound(mNAL) + 1 Then
                mDetails = mDetails & Arr(i) & NALT
            Else
                mNAL(i - 1) = Arr(i)
            End If
        Next
    ElseIf Mid(vNewValue, 4, 3) = "KNF" Then
        Me.SS = vbNullString
        mDetails = vNewValue
    ElseIf IsDigital(Left(vNewValue, 1)) And InStrCount(vNewValue, NALT) >= 7 Then
        Me.SS = "01"
        Arr = StrToArr(vNewValue, NALT)
        mDetails = vbNullString
        For i = LBound(Arr) To UBound(Arr)
            If i > UBound(mNAL) Then
                mDetails = mDetails & Arr(i) & NALT
            Else
                mNAL(i) = Arr(i)
            End If
        Next
    Else
        Me.SS = vbNullString
        mDetails = vNewValue
    End If
    If Right(mDetails, 1) = NALT Then
        mDetails = Left(mDetails, Len(mDetails) - 1)
    End If
End Sub

Public Function ValidationError() As String
    Dim Item As String, L As Long, s As String
    
    If User.Demo Then GoTo SkipCheck
    
    Item = "Номер документа "
    If mDocNo < 1 Then
        ValidationError = Item & "не указан"
        Exit Function
    ElseIf mDocNo > 999 Then
        If mBIC <> App.DefBIC Then
            ValidationError = Item & "не может превышать 999 для других банков"
            Exit Function
        ElseIf mDocNo > 99999 Then
            ValidationError = Item & "не может превышать 99999 для нашего Банка"
            Exit Function
        End If
    End If
        
    Item = "Дата документа "
    If Year(mDocDate) < 1999 Then
        ValidationError = Item & "не действительна"
        Exit Function
    End If
        
    Item = "Сумма платежа "
    If mSum < 0.01 Then
        ValidationError = Item & "не указана"
        Exit Function
    End If
        
    Item = "Очередность платежа "
    If mQueue < 1 Or mQueue > 6 Then
        ValidationError = Item & "не от 6 до 1"
        Exit Function
    End If
        
    s = User.Name
    Item = "Плательщик платежа "
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
    
    s = mName
    Item = "Получатель платежа "
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
    
    s = mINN
    Item = "ИНН получателя платежа "
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
    
    s = mKPP
    Item = "КПП получателя платежа "
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
    
    s = mDetails
    Item = "Назначение платежа "
    L = Len(s)
    If L = 0 Then
        ValidationError = Item & "не указано"
        Exit Function
    ElseIf L > 210 Then
        ValidationError = Item & "не может превышать 210 символов"
        Exit Function
    ElseIf InStr(s, "  ") > 0 Then
        ValidationError = Item & "не должно содержать лишних пробелов"
        Exit Function
    ElseIf InStr(s, "^") > 0 Then
        ValidationError = Item & "не должно содержать символ ^"
        Exit Function
    ElseIf Left(s, 1) = """" Then
        ValidationError = Item & "не должно начинаться с кавычки"
        Exit Function
    ElseIf Left(s, 1) = "-" Then
        ValidationError = Item & "не может начинаться с минуса"
        Exit Function
    'ElseIf InStr("30122,30123,40813,40814,40815", Left(mLS, 5)) > 0 Then
    '    If Not (IsDigital(Left(s, 3)) And StrComp(Mid(s, 4, 3), "KNF", vbBinaryCompare) = 0) Then
    '        ValidationError = Item & "не имеет слова KNF после трех цифр"
    '        Exit Function
    '    End If
    End If
    
    s = mSS
    Item = "Статус налогоплательщика "
    L = Len(mSS)
    If L = 0 Then
        If mQueue <> 6 Then
            ValidationError = Item & "отсутствует при очередности не 6"
            Exit Function
        ElseIf Left(mLS, 5) = "40101" Then
            ValidationError = Item & "отсутствует при платеже в бюджет"
            Exit Function
        ElseIf Left(mLS, 5) = "40314" Then
            ValidationError = Item & "отсутствует при платеже в таможню"
            Exit Function
        End If
    ElseIf InStr("01,02,03,04,05,06,07,08,09,10,11,12,13,14,15", s) = 0 Then
        ValidationError = Item & "не 01-15"
        Exit Function
    Else
    
        s = mNAL(1)
        Item = "Код бюджетной классификации "
        L = Len(s)
        If s = "0" Then
            'valid
        ElseIf L = 0 Then
            ValidationError = Item & "не указан"
            Exit Function
        ElseIf Left(mLS, 5) = "40101" Then 'Бюджет
            If L <> 20 Then
                ValidationError = Item & "не из 20 цифр"
                Exit Function
            'ElseIf InStr("1000,2000,3000", Mid(s, 14, 3)) = 0 Then
            '    ValidationError = Item & "не 1000,2000,3000 в 14-17"
            '    Exit Function
            End If
        ElseIf L > 20 Then
            ValidationError = Item & "не может превышать 20 цифр"
            Exit Function
        ElseIf Not IsDigital(s) Then
            ValidationError = Item & "не из цифр"
            Exit Function
        End If
            
        s = mNAL(2)
        Item = "Код ОКАТО "
        L = Len(s)
        If s = "0" Then
            'valid
        ElseIf L = 0 Then
            ValidationError = Item & "не указан"
            Exit Function
        ElseIf L > 11 Then
            ValidationError = Item & "не может превышать 11 цифр"
            Exit Function
        ElseIf Not IsDigital(s) Then
            ValidationError = Item & "не из цифр"
            Exit Function
        End If
            
        s = mNAL(3)
        Item = "Основание платежа "
        L = Len(s)
        If s = "0" Then
            'valid
        ElseIf L = 0 Then
            ValidationError = Item & "не указано"
            Exit Function
        ElseIf InStr("ТП,ЗД,БФ,ТР,РС,ОТ,РТ,ВУ,ПР,АП,АР", s) = 0 Then
            ValidationError = Item & "не ТП,ЗД,БФ,ТР,РС,ОТ,РТ,ВУ,ПР,АП,АР"
            Exit Function
        End If
            
        s = mNAL(4)
        Item = "Налоговый период "
        L = Len(s)
        If s = "0" Then
            'valid
        ElseIf L = 0 Then
            ValidationError = Item & "не указан"
            Exit Function
        ElseIf L <> 10 Then
            ValidationError = Item & "не из 10 знаков"
            Exit Function
        ElseIf Mid(s, 3, 1) <> "." And Mid(s, 6, 1) <> "." Then
            ValidationError = Item & "не через точки"
            Exit Function
        ElseIf Not RIsDate(s) Then
            If InStr("Д1,Д2,Д3,МС", Left(s, 2)) > 0 Then
                If InStr("01,02,03,04,05,06,07,08,09,10,11,12", Mid(s, 4, 2)) = 0 Then
                    ValidationError = Item & "не имеет месяц 01-12"
                    Exit Function
                End If
            ElseIf Left(s, 2) = "КВ" Then
                If InStr("01,02,03,04", Mid(s, 4, 2)) = 0 Then
                    ValidationError = Item & "не имеет квартал 01-04"
                    Exit Function
                End If
            ElseIf Left(s, 2) = "ПЛ" Then
                If InStr("01,02", Mid(s, 4, 2)) = 0 Then
                    ValidationError = Item & "не имеет полугодие 01-02"
                    Exit Function
                End If
            ElseIf Left(s, 2) = "ГД" Then
                If Left(s, 6) <> "ГД.00." Then
                    ValidationError = Item & "не имеет ГД.00"
                    Exit Function
                End If
            Else
                ValidationError = Item & "не Д1,Д2,Д3,МС,КВ,ПЛ,ГД и не дата"
                Exit Function
            End If
        End If
        
        s = mNAL(5)
        Item = "Номер налогового документа "
        L = Len(s)
        If s = "0" Then
            'valid
        ElseIf L = 0 Then
            ValidationError = Item & "не указан"
            Exit Function
        ElseIf L > 15 Then
            ValidationError = Item & "не может превышать 15 знаков"
            Exit Function
        End If
            
        s = mNAL(6)
        Item = "Дата налогового документа "
        L = Len(s)
        If s = "0" Then
            'valid
        ElseIf L = 0 Then
            ValidationError = Item & "не указана"
            Exit Function
        ElseIf L <> 10 Then
            ValidationError = Item & "не из 10 знаков"
            Exit Function
        ElseIf Mid(s, 3, 1) <> "." And Mid(s, 6, 1) <> "." Then
            ValidationError = Item & "не через точки"
            Exit Function
        ElseIf Not RIsDate(s) Then
            ValidationError = Item & "не дата"
            Exit Function
        End If
    
        s = mNAL(7)
        Item = "Тип налогового документа "
        L = Len(s)
        If s = "0" Then
            'valid
        ElseIf L = 0 Then
            ValidationError = Item & "не указан"
            Exit Function
        ElseIf InStr("НС,ПЛ,ГП,ВЗ,АВ,ПЕ,ПЦ,СА,АШ,ИШ", s) = 0 Then
            ValidationError = Item & "не НС,ПЛ,ГП,ВЗ,АВ,ПЕ,ПЦ,СА,АШ,ИШ"
            Exit Function
        End If
            
    End If
        
    s = mLS
    Item = "Счет получателя "
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
    ElseIf Not LSKeyValid(mBIC, s) Then
        ValidationError = Item & "не соответствует БИК по ключу"
        Exit Function
    End If
        
    s = mBIC
    Item = "БИК банка получателя "
    If Len(s) <> 9 Then
        ValidationError = Item & "не 9 цифр"
        Exit Function
    ElseIf Not IsDigital(s) Then
        ValidationError = Item & "не из цифр"
        Exit Function
    ElseIf Left(s, 2) <> "04" Then
        ValidationError = Item & "не начинается с 04"
        Exit Function
    ElseIf Len(Me.Bank) = 0 Then
        ValidationError = Item & "не найден в справочнике"
        Exit Function
    End If
    
SkipCheck:
    ValidationError = vbNullString
End Function
