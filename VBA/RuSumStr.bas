Attribute VB_Name = "RuSumStr"
Option Explicit
Option Compare Text
Option Base 1
DefLng A-Z

'Читает всю строку с цифрами и возвращает дробное целое число,
'независимо от наличия пробелов и букв в этой строке,
'но дробная часть отделяется после последней точки,
'запятой, знака равенства или минуса (или указанного списка разделителей).
Public Function RVal(Value As String, Optional Delim As String = ".,=-") As Currency
    Dim i, b() As Byte, s As String, n, p
    RVal = 0
    n = 0
    p = 0
    StrToBytes b, Trim(Value)
    For i = LBound(b) To UBound(b)
        Select Case b(i)
        Case 48 To 57 '"0" To "9"
            RVal = RVal * 10 + (b(i) - 48)
            n = n + 1
        Case Else
            If InStr(Delim, Chr(b(i))) > 0 Then
                p = n
            End If
        End Select
    Next
    If b(LBound(b)) = 45 Then '-
        RVal = -RVal
    End If
    If p > 0 Then
        RVal = RVal * 10 ^ (p - n)
    End If
End Function

'Public Function RVal(Value As String, Optional Delim As String = ".,=-") As Currency
'    Dim i As long, b As String * 1, ss As String: ss = ""
'    For i = 1 To Len(Value)
'        b = Mid$(Value, i, 1)
'        Select Case b
'        Case "0" To "9"
'            ss = ss & b
'        Case Else
'            If InStr(Delim, b) > 0 Then ss = ss & "."
'        End Select
'    Next
'    RVal = Val(ss)
'End Function

'Возвращает сумму прописью из дроби (например, полученной из RVal()
Public Function RSumStr(Rubles As Currency, Optional Delim As String = " ") As String
Attribute RSumStr.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim s As String, n, ss As String: n = 0: ss = ""

    s = Format(Rubles, "000000000000000.00")
    
    n = Val(Mid$(s, 1, 3))
    If n > 0 Then ss = ss & Nums999Str(n, True) & "триллион" & EndStr(n, "", "а", "ов") & Delim
    
    n = Val(Mid$(s, 4, 3))
    If n > 0 Then ss = ss & Nums999Str(n, True) & "миллиард" & EndStr(n, "", "а", "ов") & Delim
        
    n = Val(Mid$(s, 7, 3))
    If n > 0 Then ss = ss & Nums999Str(n, True) & "миллион" & EndStr(n, "", "а", "ов") & Delim
        
    n = Val(Mid$(s, 10, 3))
    If n > 0 Then ss = ss & Nums999Str(n, False) & "тысяч" & EndStr(n, "а", "и", "") & Delim
    
    n = Val(Mid$(s, 13, 3))
    If n > 0 Then ss = ss & Nums999Str(n, True)
    
    If Len(ss) = 0 Then ss = "ноль "
    ss = ss & "рубл" & EndStr(n, "ь ", "я ", "ей ")
            
    n = Val(Right$(s, 2))
    ss = ss & Right$(s, 2) & " копе" & EndStr(n, "йка.", "йки.", "ек.")
    
    RSumStr = UCase$(Left$(ss, 1)) & Mid$(ss, 2)
End Function

'Возвращает число из 3 разрядов прописью в зависимости от рода
Private Function Nums999Str(ByVal n As Long, ByVal male As Boolean) As String
    Dim ss As String: ss = ""
        
    '0..999
    If n > 99 Then
        ss = ss & Choose(n \ 100, "сто", "двести", "триста", "четыреста", _
            "пятьсот", "шестьсот", "семьсот", "восемьсот", "девятьсот") _
            & " "
        n = n Mod 100
    End If
    
    '0..99
    If n > 19 Then
        ss = ss & Choose(n \ 10 - 1, "двадцать", "тридцать", "сорок", _
            "пятьдесят", "шестьдесят", "семьдесят", "восемьдесят", "девяносто") _
            & " "
        n = n Mod 10
    End If
    
    '0..19
    If n > 9 Then
        ss = ss & Choose(n - 9, "десять", "одиннадцать", "двенадцать", "тринадцать", "четырнадцать", _
            "пятнадцать", "шестнадцать", "семнадцать", "восемнадцать", "девятнадцать") _
            & " "
    ElseIf n > 2 Then
        ss = ss & Choose(n - 2, "три", "четыре", "пять", "шесть", "семь", "восемь", "девять") _
            & " "
    ElseIf n = 2 Then
        ss = ss & IIf(male, "два", "две") & " "
    ElseIf n = 1 Then
        ss = ss & IIf(male, "один", "одна") & " "
    'Else n = 0
    End If
    
    Nums999Str = ss
End Function

'Возвращает одну из указанных строчек в зависимости от числа
'Например, "рубл" & EndStr(10, "ь", "я", "ей") вернет "рублей"
Public Function EndStr(ByVal n As Long, s1 As String, s2 As String, s5 As String) As String
    If n > 19 Then n = n Mod 10
    Select Case n
    Case 1
        EndStr = s1
    Case 2 To 4
        EndStr = s2
    Case Else
        EndStr = s5
    End Select
End Function

'Формат со знаком равенства/дефиса вместо десятичной точки
Public Function PlatFormat(Rubles As Currency, Optional Delim As Variant = "-") As String
Attribute PlatFormat.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim ss As String
    ss = Format(Rubles, "0.00") '"#,##0.00"
    Mid$(ss, Len(ss) - 2, 1) = Delim
    PlatFormat = ss
End Function

'Формат Y2K
Public Function PlatDate(DATA As Variant) As String
    PlatDate = Format(CDate(DATA), "dd.mm.yyyy")
End Function

'Сумма НДС из итоговой суммы (x + Y% = z)
Public Function Sum2Tax(TotalRubles As Currency, TaxPercent As Double) As Currency
    Sum2Tax = TotalRubles * TaxPercent / (100 + TaxPercent)
End Function

'Читает дату в любом формате
Public Function RDate(v As Variant) As Date 'instead CDate()
    Dim s As String, n As Long
    On Error Resume Next
    If IsDate(v) Then
        RDate = CDate(v)
    ElseIf Val(v) > 19000101 Then 'from 01.01.1900
        RDate = StoD(CStr(v))
    Else
        s = CStr(v)
        n = InStr(10, s, "г", vbTextCompare)
        If n > 10 Then 'skip Авг.
            s = Left(s, n - 1)
        End If
        RDate = DateValue(s)
    End If
End Function

'Проверяет, дата ли это - в любом формате
Public Function RIsDate(v As Variant) As Boolean 'instead IsDate()
    On Error Resume Next
    RIsDate = RDate(v) > 0
End Function
