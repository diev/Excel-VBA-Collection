Attribute VB_Name = "Base36"
Option Explicit
Option Compare Text
Option Base 1
DefLng A-Z

Public Function To36(Value As Variant) As String
    Dim n, v
    
    On Error GoTo ErrHandler
    If Value < 36 Then
        If Value < 10 Then '0..9
            To36 = Chr(48 + Value) '"0"+x
        Else 'A..Z
            To36 = Chr(55 + Value) '"A"-10+x
        End If
    Else
        To36 = ""
        n = Value
        Do While n > 0
            v = n Mod 36
            If v < 10 Then v = v + 48 Else v = v + 55
            To36 = Chr(v) & To36
            n = n \ 36
        Loop
    End If
    Exit Function
    
ErrHandler:
    To36 = "-1" 'Overflow!
End Function

Public Function Ot36(Value As String) As Long 'Integer for Asc(), Long for Win32
    Dim i, n, v
    
    On Error GoTo ErrHandler
    n = Len(Value)
    If n = 1 Then '1 digit only
        Ot36 = Asc(Value)
        Select Case Ot36
            Case 48 To 57 '0..9
                Ot36 = Ot36 - 48
            Case 65 To 90 'A..Z
                Ot36 = Ot36 - 55
            Case 97 To 122 'a..z
                Ot36 = Ot36 - 87
            Case Else
                GoTo ErrHandler
        End Select
    Else
        Ot36 = 0
        For i = 1 To n
            v = Asc(Mid(Value, i, 1))
            Select Case v
                Case 48 To 57 '0..9
                    v = v - 48
                Case 65 To 90 'A..Z
                    v = v - 55
                Case 97 To 122 'a..z
                    v = v - 87
                Case Else
                    GoTo ErrHandler
            End Select
            Ot36 = Ot36 + v * (36 ^ (n - i))
        Next
    End If
    Exit Function

ErrHandler:
    Ot36 = -1 'Overflow!
End Function

Public Function DateTo36(Optional Value As Variant) As String
    If IsMissing(Value) Then Value = Now
    
    DateTo36 = To36(Year(Value) - 1990) 'до 2026 года
    DateTo36 = DateTo36 & To36(Month(Value))
    DateTo36 = DateTo36 & To36(Day(Value))
    DateTo36 = DateTo36 & To36(Hour(Value))
    
    'Точность до десятых секунды (от начала текущего часа)
    DateTo36 = DateTo36 & PadL(To36((Timer - Hour(Value) * 3600) * 10), 3, "0")
End Function

Public Function DateOt36(Value As String, Optional vFormat As String = "dd.mm.yyyy hh:mm:ss") As String
    Dim n, yy, mm, dd, h, m, s
    
    n = Len(Value) 'ymdhxxx(36), xxx=s*10
    If n = 7 Then
        yy = Ot36(Left(Value, 1)) + 1990
    Else
        '2000 для учета високосного года, if any
        yy = 2000
    End If
    mm = Ot36(Mid(Value, n - 5, 1))
    dd = Ot36(Mid(Value, n - 4, 1))
    h = Ot36(Mid(Value, n - 3, 1))
    m = Ot36(Right(Value, 3)) \ 10
    s = m Mod 60
    m = m \ 60
    DateOt36 = Format(DateSerial(yy, mm, dd) + CDate(h & ":" & m & ":" & s), vFormat)
End Function


