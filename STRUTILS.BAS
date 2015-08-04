Attribute VB_Name = "StrUtils"
Option Explicit
Option Compare Text
Option Base 1
DefLng A-Z

Public Function DtoS(YYYYMMDD As Variant) As String
    On Error Resume Next
    DtoS = vbNullString
    DtoS = Format(YYYYMMDD, "yyyymmdd")
End Function

Public Function DtoC(DDMMYY As Variant, Optional Century As Boolean = True) As String
    On Error Resume Next
    DtoC = vbNullString
    DtoC = Format(DDMMYY, IIf(Century, "dd.MM.yyyy", "dd.MM.yy"))
End Function

Public Function StoD(YYYYMMDD As String) As Date
    On Error Resume Next
    StoD = DateValue("01.01.1970")
    StoD = DateValue(Left(YYYYMMDD, 4) & "-" & Mid(YYYYMMDD, 5, 2) & "-" & Right(YYYYMMDD, 2))
End Function

Public Function CTime(HHMM As Variant) As String
    On Error Resume Next
    CTime = "00:00"
    CTime = Format(CDate(HHMM), "HH:mm")
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

Public Function StrXMLencode(What As String) As String
    StrXMLencode = StrTran(What, "&", "&amp;")
    StrXMLencode = StrTran(StrXMLencode, """", "&quot;")
    StrXMLencode = StrTran(StrXMLencode, "<", "&lt;")
    StrXMLencode = StrTran(StrXMLencode, ">", "&gt;")
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

Public Function StrToBool(s As String) As Boolean 'Default = False
    StrToBool = InStr(1, "*1*T*TRUE*Y*YES*ON*Д*ДА*", _
        "*" & UCase(s) & "*") > 0
End Function

Public Sub StrToLines(s As String, a() As String)
    Dim start As Long, pos As Long, Row As Long
    Dim Buf As String
    ReDim a(InStrCount(s, vbLf) + 1)
    start = 1: Row = 0
    If InStr(start, s, vbCrLf) > 0 Then 'DOS
        s = s & vbCrLf
        Do
            pos = InStr(start, s, vbCrLf)
            If pos = 0 Then Exit Do 'no eoln is eof
            Buf = Mid(s, start, pos - start)
            start = pos + 2
            Row = Row + 1
            a(Row) = RTrim(Buf)
        Loop
    Else 'UNIX
        s = s & vbLf
        Do
            pos = InStr(start, s, vbLf)
            If pos = 0 Then Exit Do 'no eoln is eof
            Buf = Mid(s, start, pos - start)
            start = pos + 1
            Row = Row + 1
            a(Row) = RTrim(Buf)
        Loop
    End If
    Do While Len(a(Row)) = 0
        Row = Row - 1
    Loop
    ReDim Preserve a(Row)
End Sub

Public Function StrArgs(Where As String, ParamArray Args() As Variant) As String
    Dim i As Long, n As Long
    StrArgs = Where: n = 0
    For i = LBound(Args) To UBound(Args)
        n = n + 1
        StrArgs = StrTran(StrArgs, "%" & CStr(n), CStr(Args(i)))
    Next
End Function

Public Function GetParenthesed(s As String) As String
    Dim n1 As Long, n2 As Long
    GetParenthesed = vbNullString
    On Error GoTo ErrGet
    n1 = InStr(1, s, "(") + 1
    n2 = InStr(n1, s, ")") - 1
    GetParenthesed = Mid(s, n1, n2 - n1 + 1)
ErrGet:
End Function

Public Sub WrapStrToArr(What As String, Arr() As String, Length As Long)
    Dim i As Long, n1 As Long, n2 As Long
    What = What & " "
    n1 = 1
    For i = LBound(Arr) To UBound(Arr)
        n2 = InStrR(What, " ", n1, n1 + Length)
        If n2 = 0 Then
            Arr(i) = Mid(What, n1, Length)
            n1 = n1 + Length
        Else
            Arr(i) = Trim(Mid(What, n1, n2 - n1))
            n1 = n2 + 1
        End If
    Next
End Sub

Public Function FindTerm(s As String) As String
    Dim i As Long, b() As Byte
    StrToBytes s, b
    For i = 1 To Len(s)
        Select Case b(i)
            Case 59, 94, 124 ';^|
                FindTerm = Chr(b(i))
                Exit Function
                
            'Case 9, 32, 46 '.
            'Case 48 To 57 '0..9
            'Case 65 To 90 'A..Z
            'Case 97 To 122 'a..z
            'Case Else
            '    FindTerm = Chr(b(i))
            '    Exit Function
        End Select
    Next
    FindTerm = vbNullString
End Function

Public Function IsDigital(s As String) As Boolean
    'IsDigital = CStr(Val(s)) = s
    Dim i As Long, b() As Byte
    StrToBytes s, b
    For i = 1 To Len(s)
        Select Case b(i)
            Case 48 To 57 '0..9
                '
            Case Else
                IsDigital = False
                Exit Function
        End Select
    Next
    IsDigital = True
End Function

Public Function LFN(s As String) As String
    If InStr(s, " ") > 0 Then
        LFN = """" & s & """"
    Else
        LFN = s
    End If
End Function
