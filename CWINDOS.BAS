Attribute VB_Name = "CWinDos"
Option Explicit
Option Compare Text
Option Base 1
DefLng A-Z

Public Function CWin(DOS As String) As String
    Dim i As Long, n As Long, b() As Byte, x As Byte
    n = Len(DOS)
    If n = 0 Then
        CWin = DOS
        Exit Function
    End If
    StrToBytes DOS, b
    For i = 1 To n
        x = b(i)
        If x > 127 Then
            If x < 176 Then 'Р128..пр..я175
                b(i) = x + 64
            ElseIf x > 223 Then '№224..џ239
                If x < 240 Then
                    b(i) = x + 16
                ElseIf x = 240 Then 'Ј240
                    b(i) = 168
                ElseIf x = 241 Then 'И241
                    b(i) = 184
                End If
            
            ElseIf x = 193 Or x = 194 Or x = 196 Then
                b(i) = 45 '-
            ElseIf x = 179 Or x = 180 Or x = 195 Or x = 197 Then
                b(i) = 124 '|
            ElseIf x = 191 Or x = 192 Or x = 217 Or x = 218 Then
                b(i) = 43 '+
            
            End If
        End If
    Next
    CWin = BytesToStr(b)
End Function

Public Function CDos(Win As String) As String
    Dim i As Long, n As Long, b() As Byte, x As Byte
    n = Len(Win)
    If n = 0 Then
        CDos = Win
        Exit Function
    End If
    StrToBytes Win, b
    For i = 1 To n
        x = b(i)
        If x > 127 Then
            If x > 239 Then '№240..џ255
                b(i) = x - 16
            ElseIf x > 191 Then 'Р191..пр..я239
                b(i) = x - 64
            
            ElseIf x = 168 Then 'Ј240
                b(i) = 133 'Х133
            ElseIf x = 184 Then 'И241
                b(i) = 165 'х165
            
            ElseIf x = 185 Then 'Й->N
                b(i) = 78 'N
            ElseIf x = 150 Or x = 151 Then
                b(i) = 45 '--
            
            End If
        End If
    Next
    CDos = BytesToStr(b)
End Function

