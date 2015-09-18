Attribute VB_Name = "CWinDos"
Option Explicit
Option Base 1
DefLng A-Z

Public Function CWin(Dos As String) As String
    Dim i As Long, x As Byte, C() As Byte
    
    StrToBytes C, Dos
    For i = LBound(C) To UBound(C)
        x = C(i)
        If x > 127 Then
            If x <= 175 Then 'А..Яа..п
                C(i) = x + 64
            ElseIf x >= 224 And x <= 239 Then 'р..я
                C(i) = x + 16
            
            ElseIf x = 240 Then 'Ё
                C(i) = 168
            ElseIf x = 241 Then 'ё
                C(i) = 184
            ElseIf x = 252 Then '№
                c(i) = 185
            
            ElseIf x = 193 Or x = 194 Or x = 196 Then
                C(i) = 45 '-
            ElseIf x = 179 Or x = 180 Or x = 195 Or x = 197 Then
                C(i) = 124 '|
            ElseIf x = 191 Or x = 192 Or x = 217 Or x = 218 Then
                C(i) = 43 '+
            
            End If
        End If
    Next
    CWin = BytesToStr(C)
End Function

Public Function CDos(Win As String) As String
    Dim i As Long, x As Byte, C() As Byte
    
    StrToBytes C, Win
    For i = LBound(C) To UBound(C)
        x = C(i)
        If x > 127 Then
            If x >= 192 And x <= 239 Then 'А..яа..п
                C(i) = x - 64
            ElseIf x >= 240 And x <= 255 Then 'р..я
                C(i) = x - 16
            
            ElseIf x = 168 Then 'Ё240, Е133
                C(i) = 133
            ElseIf x = 184 Then 'ё241, е165
                C(i) = 165
            
            ElseIf x = 185 Then '№->N
                'C(i) = 78 'N
                C(i) = 252 '№
            ElseIf x = 150 Or x = 151 Then '--
                C(i) = 45
            
            End If
        End If
    Next
    CDos = BytesToStr(C)
End Function
