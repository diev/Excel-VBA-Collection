VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CArchive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit
Option Compare Text
Option Base 1
DefLng A-Z

'Архив
Const mArchiveSheet = "Архив"
'Поля
Const mArchivePayeeNick = 1 'Получатель (кр.название)
Const mArchiveNo = 2 'Номер
Const mArchiveDate = 3 'Дата
Const mArchiveFile = 4 'Файл
Const mArchiveMark = 5 'Отметка Банка
Const mArchiveAction = 6 'Срок
Const mArchiveQueue = 7 'Очередность
Const mArchiveSum = 8 'Сумма
Const mArchiveNDS = 9 'НДС
Const mArchiveDescription = 10 'Назначение
'Итого
Const mArchiveSize = 10 'Общее число столбцов

Dim mArchive As New CDataSheet
Dim mListCount As Long

Private Sub Class_Initialize()
    mArchive.Sheet = mArchiveSheet
End Sub

Public Sub Add(PayeeNick As String, DocNo As Variant, DocDate As Variant, Action As Variant, Queue As Variant, _
    Sum As String, NDS As String, Description As String)
    Dim arr As Variant, s As New CSum
    arr = mArchive.GetRecord
    arr(mArchivePayeeNick) = PayeeNick
    arr(mArchiveNo) = DocNo
    arr(mArchiveDate) = CDate(DocDate)
    arr(mArchiveAction) = Action
    arr(mArchiveQueue) = Queue
    arr(mArchiveSum) = RVal(Sum)
    arr(mArchiveNDS) = NDS
    arr(mArchiveDescription) = Description
    
    Range("Номер") = arr(mArchiveNo)
    Range("Дата") = arr(mArchiveDate)
    Range("Срок") = DateAdd("d", arr(mArchiveAction), arr(mArchiveDate))
    Range("Очередность") = arr(mArchiveQueue)
    s.Value = arr(mArchiveSum)
    Range("Сумма") = s.PlatText
    Range("СуммаПрописью") = s.AmountText
    Range("Назначение") = arr(mArchiveDescription)
    Payment.FillForm
    MakeSheetVMD
    
    arr(mArchiveFile) = Left(Range("Файл").Text, 8)
    mArchive.InsertRecord arr
End Sub

Public Sub FillByRow(n As Long)
    Dim s As New CSum
    Application.ScreenUpdating = False
    With Worksheets(mArchiveSheet)
        Payment.Nick = .Cells(n, mArchivePayeeNick).Text
        Payment.FillForm
    
        Range("Номер") = .Cells(n, mArchiveNo).Text
        Range("Дата") = .Cells(n, mArchiveDate)
        Range("Срок") = DateAdd("d", .Cells(n, mArchiveAction), .Cells(n, mArchiveDate))
        Range("Очередность") = .Cells(n, mArchiveQueue).Text
        s.Value = .Cells(n, mArchiveSum)
        Range("Сумма") = s.PlatText
        Range("СуммаПрописью") = s.AmountText
        Range("Назначение") = .Cells(n, mArchiveDescription).Text
        
        'Range("Файл") = .Cells(n, mArchiveFile).Text
        'Range("Тип") = .Cells(n, mArchiveFile).Text '///////////////////////////
        'Range("Посылка") = .Cells(n, mArchiveFile).Text '///////////////////////////
        MakeSheetVMD
    End With
    Application.ScreenUpdating = True
End Sub

Public Function List(Optional Nick As String = "", Optional Substr As String = "") As Variant
    Dim arr() As String, i, n, RecCount, s As String, n1, n2
    n = 1
    n1 = Len(Nick)
    n2 = Len(Substr)
    RecCount = mArchive.RecCount
    If RecCount > 0 Then
        ReDim arr(RecCount + 1) As String
        With Range(mArchive.DataRangeAddress)
            If n1 > 0 And n2 > 0 Then
                For i = 1 To RecCount
                    If .Cells(i, mArchivePayeeNick).Text = Nick Then
                        s = .Cells(i, mArchiveDescription).Text
                        If Len(s) > 0 Then
                            If InStr(1, s, Substr, vbTextCompare) > 0 Then
                                n = n + 1
                                arr(n) = s
                            End If
                        End If
                    End If
                Next
            ElseIf n1 = 0 And n2 > 0 Then
                For i = 1 To RecCount
                    s = .Cells(i, mArchiveDescription).Text
                    If Len(s) > 0 Then
                        If InStr(1, s, Substr, vbTextCompare) > 0 Then
                            n = n + 1
                            arr(n) = s
                        End If
                    End If
                Next
            ElseIf n1 > 0 And n2 = 0 Then
                For i = 1 To RecCount
                    If .Cells(i, mArchivePayeeNick).Text = Nick Then
                        s = .Cells(i, mArchiveDescription).Text
                        If Len(s) > 0 Then
                            n = n + 1
                            arr(n) = s
                        End If
                    End If
                Next
            Else
                For i = 1 To RecCount
                    s = .Cells(i, mArchiveDescription).Text
                    If Len(s) > 0 Then
                        n = n + 1
                        arr(n) = s
                    End If
                Next
            End If
        End With
    End If
    ReDim Preserve arr(n) As String
    arr(1) = AddNewInList
    List = arr
End Function

Public Function ListRec(Item As Long, Optional Nick As String = "", Optional Substr As String = "") As Long
    Dim i, n, RecCount
    RecCount = mArchive.RecCount
    If RecCount = 0 Then
        ListRec = 0
    ElseIf Len(Nick) = 0 Then
        With Range(mArchive.DataRangeAddress)
            n = 0
            For i = 1 To RecCount
                If Len(.Cells(i, mArchiveDescription).Text) > 0 Then
                    n = n + 1
                    If n = Item Then
                        ListRec = i
                        Exit Function
                    End If
                End If
            Next
        End With
    Else
        With Range(mArchive.DataRangeAddress)
            n = 0
            For i = 1 To RecCount
                If .Cells(i, mArchivePayeeNick).Text = Nick Then
                    If Len(.Cells(i, mArchiveDescription).Text) > 0 Then
                        n = n + 1
                        If n = Item Then
                            ListRec = i
                            Exit Function
                        End If
                    End If
                End If
            Next
        End With
    End If
End Function

Public Function ListSum10() As Variant
    Dim arr As Variant
    arr = mArchive.GetStrings(mArchiveSum, 10)
    If Not IsArray(arr) Then
        ReDim arr(1) As String
        arr(1) = ""
    End If
    ListSum10 = arr
End Function

Public Property Get Sheet() As String
    Sheet = mArchive.Sheet
End Property

Public Property Let Sheet(ByVal vNewValue As String)
    mArchive.Sheet = vNewValue
End Property

Public Property Get Mark(FileName As String) As String
    Dim i, RecCount
    Mark = ""
    RecCount = mArchive.RecCount
    With Range(mArchive.DataRangeAddress)
        For i = 1 To RecCount
            If .Cells(i, mArchiveFile).Text = FileName Then
                Mark = .Cells(i, mArchiveMark)
                Exit Property
            End If
        Next
    End With
End Property

Public Property Let Mark(FileName As String, ByVal vNewValue As String)
    Dim i, RecCount
    RecCount = mArchive.RecCount
    With Range(mArchive.DataRangeAddress)
        For i = 1 To RecCount
            If .Cells(i, mArchiveFile).Text = FileName Then
                .Cells(i, mArchiveMark) = vNewValue
                Exit Property
            End If
        Next
    End With
End Property
