Attribute VB_Name = "StrBytes"
Option Explicit
Option Compare Text
Option Base 1
DefLng A-Z

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
    (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)

Public Function BytesToStr(Bytes() As Byte) As String
    Dim LB As Long, UB As Long, Length As Long
    LB = LBound(Bytes): UB = UBound(Bytes): Length = UB - LB + 1
    BytesToStr = Space(Length)
    CopyMemory ByVal BytesToStr, Bytes(LB), Length
End Function

Public Function BytesToValue(Bytes() As Byte, Optional start As Long = 1, Optional Length As Long = 1) As Long
    Select Case Length
        Case 1: 'byte
            BytesToValue = Bytes(start)
        Case 2 To 4: 'word/dword
            CopyMemory BytesToValue, Bytes(start), Length
    End Select
End Function

Public Sub ValueToBytes(NewValue As Long, Bytes() As Byte, Optional start As Long = 1, Optional Length As Long = 1)
    Select Case Length
        Case 1: 'byte
            Bytes(start) = NewValue
        Case 2 To 4: 'word/dword
            CopyMemory Bytes(start), NewValue, Length
    End Select
End Sub

Public Function TrimNullChar(Bytes() As Byte, Optional start As Long = 1) As String
    Dim Length As Long: Length = UBound(Bytes)
    Dim s As String: s = String(Length, vbNullChar)
    CopyMemory ByVal s, Bytes(start), Length
    TrimNullChar = Left(s, InStr(s & vbNullChar, vbNullChar) - 1)
End Function

Public Sub StrToBytes(AString As String, Bytes() As Byte)
    Dim Length As Long: Length = Len(AString)
    If Length = 0 Then
        'Bytes = vbEmpty
    Else
        ReDim Bytes(1 To Length)
        CopyMemory Bytes(1), ByVal AString, Length
    End If
End Sub
