Attribute VB_Name = "Russian"
Option Explicit
Option Compare Text
Option Base 1
DefLng A-Z

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

Public Function RIsDate(v As Variant) As Boolean 'instead IsDate()
    On Error Resume Next
    RIsDate = RDate(v) > 0
End Function

'Читает всю строку с цифрами и возвращает дробное целое число,
'независимо от наличия пробелов и букв в этой строке,
'но дробная часть отделяется после последней точки,
'запятой, знака равенства или минуса (или указанного списка разделителей).
Public Function RVal(Value As Variant, Optional Delim As String = ".,=-") As Currency
    Dim i, b() As Byte, s As String, n, p
    RVal = 0: n = 0: p = 0: s = Trim(CStr(Value))
    If Len(s) = 0 Then Exit Function
    StrToBytes s, b
    For i = 1 To Len(s)
        Select Case b(i)
            Case 48 To 57 '"0" To "9"
                RVal = RVal * 10 + b(i) - 48
                n = n + 1
            Case Else
                If InStr(Delim, Chr(b(i))) > 0 Then p = n
        End Select
    Next
    If b(1) = 45 Then RVal = -RVal
    If p > 0 Then RVal = RVal * 10 ^ (p - n)
End Function

'Public Function RVal(Value As String, Optional Delim As String = ".,=-") As Currency
'    Dim i As long, b As String * 1, ss As String: ss = vbNullString
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

