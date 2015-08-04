Attribute VB_Name = "CWinDos"
Option Explicit
Option Compare Text
Option Base 1
DefLng A-Z

Public Function CWin(DOS As String) As String
    Dim i As Long, n As Long, b() As Byte
    n = Len(DOS)
    If n = 0 Then
        CWin = DOS
        Exit Function
    End If
    StrToBytes DOS, b
    For i = 1 To n
        Select Case b(i)
            Case 0 To 127
                'US-ASCII
            Case 128 To 175 'Р128..пр..я175
                b(i) = b(i) + 64
            Case 224 To 239 '№224..џ239
                b(i) = b(i) + 16
            Case 240 'Ј240
                b(i) = 168
            Case 241 'И241
                b(i) = 184
            Case 193, 194, 196
                b(i) = 45 '-
            Case 179, 180, 195, 197
                b(i) = 124 '|
            Case 191, 192, 217, 218
                b(i) = 43 '+
            Case Else
                b(i) = 46 '.
        End Select
    Next
    CWin = BytesToStr(b)
End Function

Public Function CDos(Win As String) As String
    Dim i As Long, n As Long, b() As Byte
    n = Len(Win)
    If n = 0 Then
        CDos = Win
        Exit Function
    End If
    StrToBytes Win, b
    For i = 1 To n
        Select Case b(i)
            Case 0 To 127
                'US-ASCII
            Case 240 To 255 '№240..џ255
                b(i) = b(i) - 16
            Case 191 To 239 'Р191..пр..я239
                b(i) = b(i) - 64
            Case 168 'Ј240
                b(i) = 133 'Х133
            Case 184 'И241
                b(i) = 165 'х165
            Case 185 'Й->N
                b(i) = 78 'N
            Case 150, 151
                b(i) = 45 '--
            Case 174, 187, 145 To 148
                b(i) = 34 '"
            Case Else
                b(i) = 46 '.
        End Select
    Next
    CDos = BytesToStr(b)
End Function

