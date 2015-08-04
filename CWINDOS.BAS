Attribute VB_Name = "CWinDos"
Option Explicit
Option Base 1
DefLng A-Z

Public Function CWin(Dos As String) As String
    Dim i As Long, x As Byte, c() As Byte
    
    StrToBytes c, Dos
    For i = LBound(c) To UBound(c)
        x = c(i)
        If x > 127 Then
            If x <= 175 Then 'Р..пр..я
                c(i) = x + 64
            ElseIf x >= 224 And x <= 239 Then '№..џ
                c(i) = x + 16
            
            ElseIf x = 240 Then 'Ј
                c(i) = 168
            ElseIf x = 241 Then 'И
                c(i) = 184
            
            ElseIf x = 193 Or x = 194 Or x = 196 Then
                c(i) = 45 '-
            ElseIf x = 179 Or x = 180 Or x = 195 Or x = 197 Then
                c(i) = 124 '|
            ElseIf x = 191 Or x = 192 Or x = 217 Or x = 218 Then
                c(i) = 43 '+
            
            End If
        End If
    Next
    CWin = BytesToStr(c)
End Function

Public Function CDos(Win As String) As String
    Dim i As Long, x As Byte, c() As Byte
    
    StrToBytes c, Win
    For i = LBound(c) To UBound(c)
        x = c(i)
        If x > 127 Then
            If x >= 192 And x <= 239 Then 'Р..пр..я
                c(i) = x - 64
            ElseIf x >= 240 And x <= 255 Then '№..џ
                c(i) = x - 16
            
            ElseIf x = 168 Then 'Ј240, Х133
                c(i) = 133
            ElseIf x = 184 Then 'И241, х165
                c(i) = 165
            
            ElseIf x = 185 Then 'Й->N
                c(i) = 78
            ElseIf x = 150 Or x = 151 Then '--
                c(i) = 45
            
            End If
        End If
    Next
    CDos = BytesToStr(c)
End Function

'Public Function CWin(Dos As String) As String
'    Dim i As long, x As long
'
'    CWin = Dos
'    For i = 1 To Len(Dos)
'        x = Asc(Mid(Dos, i, 1))
'        If x > 127 Then
'            If x <= 175 Then 'Р..пр..я
'                Mid(CWin, i, 1) = Chr(x + 64)
'            ElseIf x >= 224 And x <= 239 Then '№..џ
'                Mid(CWin, i, 1) = Chr(x + 16)
'            ElseIf x = 240 Then 'Ј
'                Mid(CWin, i, 1) = Chr(168)
'            ElseIf x = 241 Then 'И
'                Mid(CWin, i, 1) = Chr(184)
'            End If
'        End If
'    Next
'End Function

'Public Function CDos(Win As String) As String
'    Dim i As long, x As long
'
'    CDos = Win
'    For i = 1 To Len(Win)
'        x = Asc(Mid(Win, i, 1))
'        If x > 127 Then
'            If x >= 192 And x <= 239 Then 'Р..пр..я
'                Mid(CDos, i, 1) = Chr(x - 64)
'            ElseIf x >= 240 And x <= 255 Then '№..џ
'                Mid(CDos, i, 1) = Chr(x - 16)
'            ElseIf x = 168 Then 'Ј
'                Mid(CDos, i, 1) = Chr(240)
'            ElseIf x = 184 Then 'И
'                Mid(CDos, i, 1) = Chr(241)
'            ElseIf x = 185 Then 'Й
'                Mid(CDos, i, 1) = Chr(78)
'            End If
'        End If
'    Next
'End Function


