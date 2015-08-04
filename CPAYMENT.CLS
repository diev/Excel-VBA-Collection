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

'Платежка
Const mPaymenrSheet = "Платежка"

'Получатели
Const mPayeesSheet = "Получатели"
'Поля
Const mPayeesNick = 1 'Получатель (кр.название)
Const mPayeesCustom1 = 2 'Признак1
Const mPayeesCustom2 = 3 'Признак2
Const mPayeesName = 4 'Получатель (полное назв.)
Const mPayeesINN = 5 'ИНН
Const mPayeesLS = 6 'Счет
Const mPayeesBIC = 7 'БИК
'Итого
'Const mPayeesSize = 7 'Общее число столбцов

Dim mPayees As New CDataSheet
Dim mPayeeRecNo As Long
Dim mPayeeArr As Variant

Private Sub Class_Initialize()
    mPayees.Sheet = mPayeesSheet
    mPayeeRecNo = 1
    mPayeeArr = mPayees.GetRecord
End Sub

Public Sub AddPayee(Nick As String, Name As String, INN As String, LS As String, BIC As String, _
    Optional Custom1 As String = "", Optional Custom2 As String = "")
    mPayeeArr = mPayees.GetRecord
    mPayeeArr(mPayeesNick) = Nick
    mPayeeArr(mPayeesCustom1) = Custom1
    mPayeeArr(mPayeesCustom2) = Custom2
    mPayeeArr(mPayeesName) = Name
    mPayeeArr(mPayeesINN) = INN
    mPayeeArr(mPayeesLS) = LS
    mPayeeArr(mPayeesBIC) = BIC
    mPayees.InsertRecord mPayeeArr
    mPayeeRecNo = 1
End Sub

Public Sub FillForm()
    Application.ScreenUpdating = False
        Range("ИНН2") = mPayeeArr(mPayeesINN)
        Range("Название2") = mPayeeArr(mPayeesName)
        Range("Счет2") = mPayeeArr(mPayeesLS)
        Bank.FillForm 2, CStr(mPayeeArr(mPayeesBIC))
        Range("Отправлено") = ""
    Application.ScreenUpdating = True
End Sub

Public Sub FillActiveCell()
    With Worksheets(mPayeesSheet)
        Nick(Position) = .Cells(Selection.Row, mPayeesNick).Text
    End With
    FillForm
End Sub

Public Function List(Optional NickNameLS As String = "", _
    Optional Custom1 As String = "", Optional Custom2 As String = "") As Variant
    Dim arr() As String, i, n, l, c1, c2, RecCount
    Dim sRange As String, sNick, sName, sLS
    n = 1
    RecCount = mPayees.RecCount
    sRange = mPayees.DataRangeAddress
    sNick = mPayeesNick
    sName = mPayeesName
    sLS = mPayeesLS
    c1 = Len(Custom1)
    c2 = Len(Custom2)
    If RecCount > 0 Then
        With Range(sRange)
            ReDim arr(RecCount + 1) As String
            l = Len(NickNameLS)
            If l + c1 + c2 = 0 Then
                n = RecCount + 1
                For i = 1 To RecCount
                    arr(i + 1) = .Cells(i, sNick).Text
                Next
            Else
                For i = 1 To RecCount
                    If InStr(1, .Cells(i, sNick).Text & .Cells(i, sName).Text & _
                        .Cells(i, sLS).Text, NickNameLS, vbTextCompare) > 0 Then
                        If c1 + c2 = 0 Then
                            n = n + 1
                            arr(n) = .Cells(i, sNick).Text
                        ElseIf .Cells(i, mPayeesCustom1).Text = Custom1 And .Cells(i, mPayeesCustom2).Text = Custom2 Then
                            n = n + 1
                            arr(n) = .Cells(i, mPayeesNick).Text
                        End If
                    End If
                Next
            End If
        End With
    End If
    ReDim Preserve arr(n) As String
    arr(1) = AddNewInList
    List = arr
End Function

Public Function CustomList(No As Long) As Variant
    Dim arr() As String, i, j, n, s As String, b As Boolean, RecCount, custn
    custn = IIf(No = 1, mPayeesCustom1, mPayeesCustom2)
    n = 1
    RecCount = mPayees.RecCount
    If RecCount > 0 Then
        ReDim arr(RecCount + 1) As String
        With Range(mPayees.DataRangeAddress)
            For i = 1 To RecCount
                s = .Cells(i, custn).Text
                If Len(s) > 0 Then 'take filled only
                    b = False
                    For j = 1 To n - 1 'was it before?
                        If arr(j) = s Then
                            b = True 'found previous!
                            Exit For
                        End If
                    Next
                    If Not b Then 'add new
                        n = n + 1
                        arr(n) = s
                    End If
                End If
            Next
        End With
    End If
    ReDim Preserve arr(n) As String
    arr(1) = ""
    CustomList = arr
End Function

Public Function ListNames10() As Variant
    Dim arr As Variant
    arr = mPayees.GetStrings(mPayeesName, 10)
    If Not IsArray(arr) Then
        ReDim arr(1) As String
        arr(1) = ""
    End If
    ListNames10 = arr
End Function

Public Function ListBIC10() As Variant
    Dim arr As Variant
    arr = mPayees.GetStrings(mPayeesBIC, 10)
    If Not IsArray(arr) Then
        ReDim arr(1) As String
        arr(1) = ""
    End If
    ListBIC10 = arr
End Function

Public Property Get PlatSheet() As String
    Sheet = mPaymenrSheet
End Property

Public Property Get Sheet() As String
    Sheet = mPayees.Sheet
End Property

Public Property Let Sheet(ByVal vNewValue As String)
    mPayees.Sheet = vNewValue
End Property

Public Property Get Nick() As String
    Nick = mPayeeArr(mPayeeNick)
End Property

Public Property Let Nick(ByVal vNewValue As String)
    'If vNewValue <> mPayeeArr(mPayeesNick) Then
        mPayeeRecNo = mPayees.GetNickRecNo(vNewValue)
        mPayeeArr = mPayees.GetRecord(mPayeeRecNo)
    'End If
End Property

Public Function GetNickUnique(Name As String) As String
    GetNickUnique = mPayees.GetNickUnique(Name)
End Function

