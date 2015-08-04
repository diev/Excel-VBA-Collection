Attribute VB_Name = "Base36"
Option Explicit
Option Compare Text
Option Base 1
DefLng A-Z

Public Function To36(Value As Variant) As String
    Dim n As Long, v As Long
    
    On Error GoTo ErrHandler
    If Value < 10 Then '0..9
        To36 = Chr(48 + Value) '"0"+x
    ElseIf Value < 36 Then 'A..Z
        To36 = Chr(55 + Value) '"A"-10+x
    Else 'If Value < 2147483648# Then
        To36 = vbNullString
        n = CCur(Value)
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

Public Function Ot36(Value As String) As Long
    Dim i, n, v, b() As Byte
    
    'On Error GoTo ErrHandler
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
        StrToBytes b, Value
        Ot36 = 0
        For i = 1 To n
            v = b(i)
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
