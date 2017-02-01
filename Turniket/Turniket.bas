'(c) Дмитрий Евдокимов, ред. 01.02.2017

' Исходные данные:
' 1) Этот XLSM-файл с модулем Turniket.bas;
' 2) XLS или CSV-файл с турникетов;
' 3) TXT-файл с парковки;
' 4) Присвоить нужный период с Date1 по Date2 ниже.
'
' Убеждаемся, что есть лист "Отчет" (он будет очищен), а листы "Парковка" и "Турникет" будут удалены и загружены снова
' Меню "Разработчик" - "Макросы" - выбираем единственный макрос TurnOver - "Выполнить" - ждем
'
' На листе "Отчет" кликаем внутри таблицы данных
' Меню "Данные" - "Промежуточный итог" (ФИО, Сумма, Дробь)
' "Сохранить" - отдать в Кадры
'

Option Explicit
Option Compare Text

'Фильтр с Date1 по Date2
Const Date1 As Date = #1/1/2017# 'mm/dd/yyyy
Const Date2 As Date = #1/31/2017# 'mm/dd/yyyy

'Столбцы с турникета
Const TURNIKET As String = "Турникет"
Const ColTName As Long = 2
Const ColTDate As Long = 3
Const ColTTime As Long = 4
Const ColTEvent As Long = 5

'Столбцы с парковки
Const PARKING As String = "Парковка"
Const ColPDate As Long = 1
Const ColPTime As Long = 2
Const ColPObject As Long = 3
Const ColPName As Long = 4

'Столбцы отчета
Const REPORT As String = "Отчет"
Const ColRName As Long = 1
Const ColRDate As Long = 2
Const ColRLogin As Long = 3
Const ColRObjin As Long = 4
Const ColRLogout As Long = 5
Const ColRObjout As Long = 6
Const ColRHours As Long = 7
Const ColRMins As Long = 8
Const ColRTotal As Long = 9

Sub TurnOver()
    Dim SheetFile As Variant
    Dim WB As String
    
    Dim Sheet1 As Worksheet
    Dim Sheet2 As Worksheet
    Dim Row1 As Long
    Dim Row2 As Long
    
    Dim StatusStr As String
    
    Dim SName As String
    Dim SDate As String
    Dim TDate As Date
    Dim S As String
    
    Dim nMins As Long
    Dim i As Long
    
    'Очистка отчета
    Application.DisplayStatusBar = True
    WB = ActiveWorkbook.Name
    Set Sheet2 = ActiveWorkbook.Worksheets(REPORT)
    Sheet2.Cells.Delete
    Row2 = 1
    
    'GoTo TurniketLoaded '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    'Ищем данные с турникета
    Application.StatusBar = "Загрузка данных с турникета..."
    ChDir CurDir
    SheetFile = Application.GetOpenFilename("Excel (*.xls;*.csv), *.xls;*.csv", , "Данные с турникета (файл Excel)")
    If SheetFile = False Then Exit Sub

    For Each Sheet1 In Sheets
        If Sheet1.Name = TURNIKET Then Sheet1.Delete
    Next
    
    Workbooks.Open Filename:=SheetFile
    
    If LCase(Right(SheetFile, 4)) = ".csv" Then
        Columns("A:A").Select
        Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
            TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
            Semicolon:=True, Comma:=False, Space:=False, Other:=False, FieldInfo _
            :=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, 1)), _
            TrailingMinusNumbers:=True
    End If
    
    Sheets(1).Select
    Sheets(1).Copy Before:=Workbooks(WB).Sheets(1)
    Workbooks(WB).Activate
    Sheets(1).Select
    Sheets(1).Name = TURNIKET

TurniketLoaded:
    Set Sheet1 = ActiveWorkbook.Worksheets(TURNIKET)
    
    Application.StatusBar = "Отбор данных с турникета..."
    Sheet2.Activate
    
    Row1 = 2
    StatusStr = ""
    Do While Len(Trim(Sheet1.Cells(Row1, ColTDate).Text)) > 0
        SDate = Trim(Sheet1.Cells(Row1, ColTDate).Text)
        If StatusStr <> SDate Then
            StatusStr = SDate
            Application.StatusBar = "Отбор данных с турникета... " & StatusStr
            Sheet2.Cells(Row2, ColRDate).Select
            Sheet2.Columns("A:D").AutoFit
            DoEvents
        End If
        
        S = Trim(Sheet1.Cells(Row1, ColTEvent).Text)
        If S = "Проход" Then
            TDate = CDate(Sheet1.Cells(Row1, ColTDate))
            If Date1 <= TDate And TDate <= Date2 Then
                SName = Trim(Sheet1.Cells(Row1, ColTName).Text)
                If Len(SName) > 0 Then
                    Sheet2.Cells(Row2, ColRName) = Replace(SName, "  ", " ")
                    Sheet2.Cells(Row2, ColRDate) = TDate
                    Sheet2.Cells(Row2, ColRLogin) = TDate + CDate(Sheet1.Cells(Row1, ColTTime))
                    Sheet2.Cells(Row2, ColRObjin) = S
                    Row2 = Row2 + 1
                End If
            End If
        End If
        Row1 = Row1 + 1
    Loop
    Set Sheet1 = Nothing
    
    'GoTo ParkingLoaded '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    'Ищем данные с парковки
    Application.StatusBar = "Загрузка данных с парковки..."
    SheetFile = Application.GetOpenFilename("Text (*.txt), *.txt", , "Данные с парковки (текстовый файл)")
    If SheetFile = False Then Exit Sub
    
    For Each Sheet1 In Sheets
        If Sheet1.Name = PARKING Then Sheet1.Delete
    Next
    
    Workbooks.OpenText Filename:=SheetFile, Origin:=1251 _
        , StartRow:=1, DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, _
        ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, Comma:=False _
        , Space:=False, Other:=False, FieldInfo:=Array(Array(1, 1), Array(2, 1), _
        Array(3, 1), Array(4, 1)), TrailingMinusNumbers:=True
    Sheets(1).Select
    Sheets(1).Copy Before:=Workbooks(WB).Sheets(1)
    Sheets(1).Select
    Sheets(1).Name = PARKING
    
ParkingLoaded:
    Set Sheet1 = ActiveWorkbook.Worksheets(PARKING)
    Sheet1.Columns("A:D").AutoFit
    
    Application.StatusBar = "Отбор данных с парковки..."
    Sheet2.Activate
    
    Row1 = 3
    StatusStr = ""
    Do While Len(Sheet1.Cells(Row1, ColPDate).Text) = 10 'dd.mm.yyyy (maybe eof)
        SDate = Sheet1.Cells(Row1, ColPDate).Text
        If StatusStr <> SDate Then
            StatusStr = SDate
            Application.StatusBar = "Отбор данных с парковки... " & StatusStr
            Sheet2.Cells(Row2, ColRDate).Select
            DoEvents
        End If
        
        SName = Trim(Sheet1.Cells(Row1, ColPName).Text)
        If Len(SName) > 0 Then
            TDate = CDate(Sheet1.Cells(Row1, ColPDate))
            If Date1 <= TDate And TDate <= Date2 Then
                Sheet2.Cells(Row2, ColRName) = FIO(SName)
                Sheet2.Cells(Row2, ColRDate) = TDate
                Sheet2.Cells(Row2, ColRLogin) = TDate + CDate(Sheet1.Cells(Row1, ColPTime))
                Sheet2.Cells(Row2, ColRObjin) = Sheet1.Cells(Row1, ColPObject)
                Row2 = Row2 + 1
            End If
        End If
        Row1 = Row1 + 1
    Loop
    Set Sheet1 = Nothing
    
    'Сортируем
    Application.StatusBar = "Сортировка по времени... "
    Sheet2.Sort.SortFields.Clear
    Sheet2.Sort.SortFields.Add Key:=Range("C1"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Отчет").Sort
        .SetRange Range("A1:D" & Row2 - 1)
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    'Поиск ухода
    Application.StatusBar = "Поиск времени ухода... "
    Row1 = 1
    
    SName = ""
    Do While Len(Sheet2.Cells(Row1, ColRName).Text) > 0
        SDate = Sheet2.Cells(Row1, ColRDate)
        If SName <> Sheet2.Cells(Row1, ColRName).Text Then
            SName = Sheet2.Cells(Row1, ColRName).Text
            Application.StatusBar = "Поиск времени ухода... " & SDate & " " & SName
            Sheet2.Cells(Row1, ColRName).Select
            DoEvents
        End If
        'If Left(SName, ColRName) <> "-" Then
            Row2 = Row1 + 1
            Do While Sheet2.Cells(Row2, ColRDate).Text = SDate
                If Sheet2.Cells(Row2, ColRName).Text = SName Then
                    Sheet2.Cells(Row1, ColRLogout).FormulaR1C1 = Sheet2.Cells(Row2, ColRLogin)
                    Sheet2.Cells(Row1, ColRObjout) = Sheet2.Cells(Row2, ColRObjin)
                    nMins = DateDiff("n", Sheet2.Cells(Row1, ColRLogin), Sheet2.Cells(Row1, ColRLogout)) - 48 'Обед 48 минут
                    If nMins > 0 Then
                        Sheet2.Cells(Row1, ColRHours).FormulaR1C1 = "=RC[-2]-RC[-4]-48/60/24" ' = nMins
                        Sheet2.Cells(Row1, ColRMins) = nMins
                        Sheet2.Cells(Row1, ColRTotal).FormulaR1C1 = "=RC[-1]/60" ' = nMins \ 60
                    End If
                    'Sheet2.Cells(Row2, ColRName) = "-" & Sheet2.Cells(Row2, 1).Text
                    Sheet2.Rows(Row2).Delete
                Else
                    Row2 = Row2 + 1
                End If
            Loop
        'End If
        Row1 = Row1 + 1
    Loop
    
    'Финальная красота
    Application.StatusBar = "Сортировка по ФИО... "
    Row1 = Row1 - 1
    Sheet2.Sort.SortFields.Clear
    Sheet2.Sort.SortFields.Add Key:=Range( _
        "A1:A" & Row1), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    Sheet2.Sort.SortFields.Add Key:=Range( _
        "B1:B" & Row1), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With Sheet2.Sort
        .SetRange Range("A1:I" & Row1)
        '.Header = xlYes
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    Sheet2.Rows(1).Insert
    Row2 = 1
    Sheet2.Cells(Row2, ColRName) = "ФИО" 'A
    Sheet2.Cells(Row2, ColRDate) = "Дата" 'B
    Sheet2.Cells(Row2, ColRLogin) = "Приход" 'C
    Sheet2.Cells(Row2, ColRObjin) = "Вход" 'D
    Sheet2.Cells(Row2, ColRLogout) = "Уход" 'E
    Sheet2.Cells(Row2, ColRObjout) = "Выход" 'F
    Sheet2.Cells(Row2, ColRHours) = "Часы" 'G
    Sheet2.Cells(Row2, ColRMins) = "Минуты" 'H
    Sheet2.Cells(Row2, ColRTotal) = "Дробь" 'I
    
    Sheet2.Rows(Row2).Font.Bold = True
    Sheet2.Columns(ColRName).NumberFormat = "@"
    Sheet2.Columns(ColRLogin).NumberFormat = "h:mm;@"
    Sheet2.Cells(Row2, ColRObjin).NumberFormat = "@"
    Sheet2.Columns(ColRLogout).NumberFormat = "h:mm;@"
    Sheet2.Cells(Row2, ColRObjout).NumberFormat = "@"
    Sheet2.Columns(ColRHours).NumberFormat = "h:mm;@"
    Sheet2.Columns(ColRMins).NumberFormat = "0"
    Sheet2.Columns(ColRTotal).NumberFormat = "0.00"
    Sheet2.Columns("A:I").EntireColumn.AutoFit
    
    'Sheet2.Cells("A2").Subtotal GroupBy:=1, Function:=xlSum, TotalList:=Array(5), _
        Replace:=True, PageBreaks:=False, SummaryBelowData:=True
    'Sheet2.Outline.ShowLevels RowLevels:=2
    
    Application.StatusBar = False
End Sub

Function FIO(S As String)
    Dim A() As String
    S = Replace(S, "  ", " ")
    A = Split(S)
    If UBound(A) = 2 Then
        FIO = A(0) & " " & Left(A(1), 1) & "." & Left(A(2), 1) & "."
    Else
        'MsgBox ("Ошибка в ФИО с парковки")
        'Stop
        FIO = S
    End If
End Function
