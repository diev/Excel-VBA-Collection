VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CBank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit
Option Compare Text
Option Base 1
DefLng A-Z

Dim mBIC As String * 9
Dim mName As String
Dim mPlace As String
Dim mKS As String * 20
Dim mPostCode As String * 6
Dim mAddress As String
Dim mTel As String
Dim mDateIzm As Date
Dim mDelivery As String * 1
Dim mRecNo As Long

Public Property Let BIC(ByVal vNewValue As String)
    If mBIC <> vNewValue Then
        mBIC = PadL(vNewValue, 9, "0")
        BnkSeekBIC
    End If
End Property

Public Property Get BIC() As String
    BIC = mBIC
End Property

Public Function Name() As String
    Name = mName
End Function

Public Function Place() As String
    Place = mPlace
End Function

Public Function KS() As String
    KS = mKS
End Function

Public Function PostCode() As String
    PostCode = mPostCode
End Function

Public Function Address() As String
    Address = mAddress
End Function

Public Function Tel() As String
    Tel = mTel
End Function

Public Function DateIzm() As String
    DateIzm = mDateIzm
End Function

Public Function Delivery() As String
    Delivery = mDelivery
End Function

Public Function Bankier() As String
    Select Case mDelivery
        Case "A": Bankier = "SM1"
        Case "B": Bankier = "SM2"
        Case Else: Bankier = " "
    End Select
End Function

Public Function PrintDelivery() As String
    Select Case mDelivery
        Case "C": PrintDelivery = "Электронно"
        Case Else: PrintDelivery = "Почтой"
    End Select
End Function

Public Function RecNo() As Long
    RecNo = mRecNo
End Function

Private Sub BnkSeekBIC()
    Dim Rec As RecordType, Buf As String, n
    Rec = SeekDbfFile(Range("BnkSeek").Text, "NEWNUM", mBIC)
    mRecNo = Rec.RecNo
    If mRecNo = 0 Then
        Exit Sub
    End If
     
    mKS = FieldValue(Rec, "KSNP")
     
    If FieldValue(Rec, "P") = "+" Then
        Buf = ""
    Else
        Buf = FieldValue(Rec, "PZN")
        Select Case Buf
            Case "00": Buf = "ГРКЦ "
            Case "10": Buf = "РКЦ "
            Case "20": Buf = "Б "
            Case "21": Buf = "КБ "
            Case "22": Buf = "СБ "
            Case "23": Buf = "АКБ "
            Case "24": Buf = "ЧКБ "
            Case "25": Buf = "КОПБ "
            Case "26": Buf = "АПБ "
            Case "30": Buf = "ФБ "
            Case "31": Buf = "ФКБ "
            Case "32": Buf = "Отд."
            Case "33": Buf = "ФАКБ "
            Case "34": Buf = "ФЧКБ "
            Case "35": Buf = "ФКОБ "
            Case "36": Buf = "Отд."
            Case "40": Buf = "ПУ "
            Case "50": Buf = "ЦХ "
            Case "70": Buf = "КУ "
            Case "71": Buf = "КЛ "
            Case "72": Buf = "ОРЦБ "
            Case "90": Buf = "ЛИКВ "
            Case "98": Buf = "ИСКЛ "
            Case "99": Buf = "ОТЗВ "
            Case Else
                Buf = ""
        End Select
    End If
    mName = Buf & FieldValue(Rec, "NAMEP")
     
    n = Val(FieldValue(Rec, "TNP"))
    mPlace = Choose(n, "Г.", "П.", "С.", "ПГТ ", "СТ-ЦА ", "АУЛ ", "РП ") & _
        FieldValue(Rec, "NNP")
     
    If Left(mBIC, 4) = "0440" Then 'С-Петербург
        Buf = Mid(mBIC, 5, 2)
        mDelivery = IIf(Buf = "30", "A", "B") 'внутригор. / внутрирег.
         
        'Buf = Right(mBIC, 3) 'Sorry, это в счете клиента...
        'If Buf = "790" Or Buf = "791" Then
        '    'Требуется таблица филиалов!
        '    Mid(mKS, 10, 4) = "xxxx"
        'End If
         
    Else 'иногородние
        n = Val(FieldValue(Rec, "UER"))
        mDelivery = IIf(n = 1 Or n = 3, "C", "P") 'межрег. / почтовые
    End If
    
    mPostCode = FieldValue(Rec, "IND")
    mAddress = FieldValue(Rec, "ADR")
    mTel = FieldValue(Rec, "TELEF")
    mTel = StrTran(mTel, ")", ") ")
    mTel = StrTran(mTel, ",", ", ")
    'mTel = StrTran(mTel, "  ", " ")
    mDateIzm = FieldValue(Rec, "DT_IZM")
End Sub

Public Function Info() As String
    Info = BPrintF("%s, %s\nБИК %s, К/С %s", mName, mPlace, mBIC, mKS)
End Function

Public Function FullInfo() As String
    FullInfo = BPrintF("%s\nИнд.: %s, адрес: %s\nтел.: %s, дата изм.: %s", Info, mPostCode, mAddress, mTel, mDateIzm)
End Function

Public Sub FillForm(Position As Long, BIC As String)
    Me.BIC = BIC
    If Position = 1 Then
        Range("БИК1") = mBIC
        Range("Банк1") = mName
        Range("Место1") = mPlace
        Range("Корсчет1") = mKS
    Else
        Range("БИК2") = mBIC
        Range("Банк2") = mName
        Range("Место2") = mPlace
        Range("Корсчет2") = mKS
        Range("ВидПлатежа") = PrintDelivery
        Range("Тип") = Bankier
        Range("Посылка") = mDelivery
    End If
End Sub
