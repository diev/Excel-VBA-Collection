Attribute VB_Name = "StrUtils"
Option Explicit
Option Compare Text
Option Base 1
DefLng A-Z

Public Function DtoS(YYYYMMDD As Variant) As String
    If IsDate(YYYYMMDD) Then
        DtoS = Format(YYYYMMDD, "yyyymmdd")
    Else
        DtoS = Space(8)
        'DtoS = Format(DateSerial(2099, Month(Now), Day(Now)), "yyyymmdd")
        
    End If
End Function

Public Function DtoC(DDMMYY As Variant) As String
    If IsDate(DDMMYY) Then
        DtoC = Format(DDMMYY, "dd.mm.yyyy")
    Else
        'DtoC = "  .  .  "
        'DtoC = DateSerial(2099, Month(Now), Day(Now))
    End If
End Function

Public Function StoD(YYYYMMDD As String) As Date
    StoD = DateValue(Left(YYYYMMDD, 4) & "-" & Mid(YYYYMMDD, 5, 2) & "-" & Right(YYYYMMDD, 2))
End Function

Public Function CTime(HHMM As Variant) As String
    If IsDate(HHMM) Then
        CTime = Format(CDate(HHMM), "hh:mm")
    Else
        CTime = "00:00"
    End If
End Function

Public Function Pad(Text As String, Length As Long, Optional Char As Variant = 32) As String
    Pad = PadR(Text, Length, Char)
End Function

Public Function PadR(Text As String, Length As Long, Optional Char As Variant = 32) As String
    If Length > 0 Then
        PadR = Left(Text & String(Length, Char), Length)
    Else
        PadR = Trim(Text)
    End If
End Function

Public Function PadL(Text As String, Length As Long, Optional Char As Variant = 32) As String
    If Length > 0 Then
        PadL = Right(String(Length, Char) & Text, Length)
    Else
        PadL = Trim(Text)
    End If
End Function

Public Function PadC(Text As String, Length As Long, Optional Char As Variant = 32) As String
    If Length > 0 Then
        PadC = Mid(String(Length, Char) & Text & String(Length, Char), (Length - Len(Text)) \ 2, Length)
    Else
        PadC = Trim(Text)
    End If
End Function 'doesn't work!

Public Function PadLine(Text As String, Length As Long, Optional Delim As String = "|", Optional Char As Variant = 32) As String
    Dim LText As String, CText As String, RText As String, n1, n2
    n1 = InStr(1, Text, Delim)
    n2 = InStrR(Text, Delim)
    If n1 = 0 Then
        PadLine = PadR(Text, Length, Char)
    Else
        LText = Left(Text, n1 - 1)
        CText = Mid(Text, n1 + 1, n2 - n1 - 1)
        RText = Mid(Text, n2 + 1)
        PadLine = LText & PadC(CText, Length - Len(LText) - Len(RText), Char) & RText
    End If
End Function

Public Function PadLine2(LText As String, RText As String, Length As Long, Optional Delim As String = "|", Optional Char As Variant = 32) As String
    PadLine2 = PadR(LText, Length - Len(RText), Char) & RText
End Function

Public Function PadLine3(LText As String, CText As String, RText As String, Length As Long, Optional Delim As String = "|", Optional Char As Variant = 32) As String
    PadLine3 = LText & PadC(CText, Length - Len(LText) - Len(RText), Char) & RText
End Function

Public Function BPrintF(FormatStr As String, ParamArray Args() As Variant) As String
    'like the C function sprintf()
    Dim i, n, C As String, s As String
    BPrintF = FormatStr
    BPrintF = StrTran(BPrintF, "~", """")
    BPrintF = StrTran(BPrintF, "\n", vbCrLf)
    BPrintF = StrTran(BPrintF, "\t", vbTab)
    For i = LBound(Args) To UBound(Args)
        n = InStr(1, BPrintF, "%")
        If n = 0 Then
            Exit For
        End If
        C = Mid(BPrintF, n, 2)
        Select Case C
            Case "%s"
                s = CStr(Args(i))
            Case "%d"
                s = CInt(Args(i))
            Case "%y"
                s = DtoC(Args(i)) 'dd.mm.yy'
            Case "%t"
                s = CTime(Args(i)) 'hh:mm'
            Case "%w"
                s = Format(Args(i), "dd.mm.yy hh:mm")
            Case Else
                s = C
        End Select
        BPrintF = StrTran(BPrintF, C, s, 1)
    Next
End Function

'Join() in Office2000
Public Function ArrToStr(Arr As Variant, Optional Delim As String = " ", Optional KeepEmpty As Boolean = True) As String
    Dim i, s As String
    ArrToStr = ""
    For i = LBound(Arr) To UBound(Arr)
        s = Trim(Arr(i))
        If KeepEmpty Then
            ArrToStr = ArrToStr & Trim(Arr(i)) & Delim
        ElseIf Len(s) > 0 Then
            ArrToStr = ArrToStr & Trim(Arr(i)) & Delim
        End If
    Next
End Function

'Split() in Office2000
Public Function StrToArr(s As String, Optional Delim As String = " ", Optional KeepEmpty As Boolean = True) As Variant
    Dim i, Arr As Variant, n1, n2, n, ss As String
    i = InStrCount(s, Delim)
    If Right(s, Len(Delim)) <> Delim Then i = i + 1
    ReDim Arr(i) As String
    n1 = 1
    n2 = 1
    n = LBound(Arr)
    For i = LBound(Arr) To UBound(Arr)
        n2 = InStr(n1, s, Delim)
        If n2 = 0 Then n2 = Len(s) + 1
        ss = Trim(Mid(s, n1, n2 - n1))
        If KeepEmpty Then
            Arr(i) = ss
            n = n + 1
        ElseIf Len(s) > 0 Then
            Arr(i) = ss
            n = n + 1
        End If
        n1 = n2 + Len(Delim)
    Next
    If n < UBound(Arr) Then
        ReDim Preserve Arr(n)
    End If
    StrToArr = Arr
End Function

'Replace() in Office2000
Public Function StrTran(Where As String, What As String, Dest As String, Optional Count As Long = 0) As String
    Dim i: i = 0
    StrTran = Where
    Do
        i = InStr(i + 1, StrTran, What)
        If i > 0 Then StrTran = Left(StrTran, i - 1) & Dest & Mid(StrTran, i + Len(What))
        Count = Count - 1
    Loop Until i = 0 Or Count = 0
End Function

'Only one space between must survive!
Public Function StrSpaces1(Where As String) As String
    Dim i: i = 1
    StrSpaces1 = Trim(Where)
    StrSpaces1 = StrTran(StrSpaces1, vbTab, " ")
    StrSpaces1 = StrTran(StrSpaces1, vbCrLf, " ")
    StrSpaces1 = StrTran(StrSpaces1, vbCr, " ")
    StrSpaces1 = StrTran(StrSpaces1, vbLf, " ")
    Do
        i = InStr(i, StrSpaces1, Space(2))
        If i > 0 Then StrSpaces1 = Left(StrSpaces1, i) & Mid(StrSpaces1, i + 2)
    Loop Until i = 0
End Function

'InStrRev() in Office2000
Public Function InStrR(Where As String, What As String, _
    Optional after As Long = 1, Optional before As Long = 0) As Long
    Dim i, s As String
    If before = 0 Then
        s = Where
    Else
        s = Left(Where, before - 1)
    End If
    InStrR = InStr(after, s, What)
    If InStrR = 0 Then Exit Function
    Do
        i = InStr(InStrR + 1, s, What)
        If i = 0 Then Exit Function
        InStrR = i
    Loop
End Function

Public Function InStrCount(Where As String, What As String) As Long
    Dim i
    i = 0
    InStrCount = 0
    Do
        i = InStr(i + 1, Where, What)
        If i = 0 Then Exit Function
        InStrCount = InStrCount + 1
    Loop
End Function

Public Function NickByName(Name As String) As String
    'Выделяет последнюю фразу в кавычках, а иначе - последнее слово
    Dim i, n1, n2
    n1 = 1
    n2 = Len(Name)
    For i = n2 To n1 Step -1
        If Mid(Name, i, 1) = """" Then
            n2 = i - 1
            Exit For
        End If
    Next
    For i = n2 To n1 Step -1
        If Mid(Name, i, 1) = """" Then
            n1 = i + 1
            Exit For
        End If
    Next
    
    If n1 = 1 Then
        For i = n2 To n1 Step -1
            If Mid(Name, i, 1) = " " Then
                n1 = i + 1
                Exit For
            End If
        Next
    End If
    NickByName = Mid(Name, n1, n2 - n1 + 1)
End Function

Public Function BoolToStr(b As Boolean) As String
    BoolToStr = IIf(b, "1", "0")
End Function

Public Function StrToBool(s As String) As Boolean
    StrToBool = InStr(1, "*1*T*TRUE*Y*YES*ON*Д*ДА*", _
        "*" & UCase(s) & "*") > 0
End Function

'NEED BE TESTED!
Public Function StrToLines(ByVal s As String, Length As Long, Optional Lines As Long = 0) As Variant
    Dim i, n, after, before, Arr As Variant
    s = StrSpaces1(s)
    after = 1
    Do
        before = after + Length + 1
        i = InStrR(s, " ", after, before)
        If i = 0 Then Exit Do
        Mid(s, i, 1) = vbLf
        after = i
    Loop While after < Len(s)
    If Lines > 0 Then
        For n = InStrCount(s, vbLf) To Lines
            s = s & vbLf
        Next
    End If
    StrToLines = StrToArr(s, vbLf)
End Function
