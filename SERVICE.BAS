Attribute VB_Name = "Service"
Option Explicit
Option Compare Text
Option Base 1
DefLng A-Z

Public Sub IfPaint(Optional Fields As Boolean = False)
    Dim c As Range, cSrc, cDest
    With Range("Бланк")
        cDest = .Cells(1).Interior.ColorIndex 'A1 - get dest. color
        cSrc = .Cells(2).Interior.ColorIndex 'A2 - get back color
        .Cells(1).Interior.ColorIndex = cSrc 'restore sample A1
    End With
    If Fields Then
        cSrc = Range("Номер").Interior.ColorIndex
    End If
    If cDest = cSrc Then
        MsgBox "Выберите отличный цвет ячейки A1 и повторите", vbInformation, AppTitle
        Exit Sub
    End If
    ScreenUpdating = False
    For Each c In Range("Бланк")
        If c.Interior.ColorIndex = cSrc Then
            c.Interior.ColorIndex = cDest
        End If
    Next
    ScreenUpdating = True
End Sub

'Не очень страшная функция - очищает платежку!
Public Sub ClearPlat()
    Range("Номер") = Empty
    Range("Дата") = Empty
    Range("Срок") = Empty
    Range("Очередность") = Empty
    Range("Сумма") = Empty
    Range("СуммаПрописью") = Empty
    Range("Назначение") = Empty
    
    Range("ИНН2") = Empty
    Range("Название2") = Empty
    Range("Счет2") = Empty
    Range("БИК2") = Empty
    Range("Банк2") = Empty
    Range("Место2") = Empty
    Range("Корсчет2") = Empty
    
    Range("Тип") = Empty
    Range("ВидПлатежа") = Empty
    Range("Файл") = Empty
    Range("Посылка") = Empty
    Range("Отправлено") = Empty
End Sub

'Очень страшная функция - уничтожает всю платежку!
Public Sub ClearPlatAll()
    ClearPlat
    Range("ИНН1") = Empty
    Range("Название1") = Empty
    Range("Счет1") = Empty
    Range("БИК1") = Empty
    Range("Банк1") = Empty
    Range("Место1") = Empty
    Range("Корсчет1") = Empty
End Sub
