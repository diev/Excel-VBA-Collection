'Attribute VB_Name = "Bytes"
Option Explicit
Option Base 0 'Binary bytes like the C language!!!
DefLng A-Z

Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)

Sub StrToBytes(ab() As Byte, s As String)
    If IsArrayEmpty(ab) Then If Len(s) > 0 Then ReDim ab(0 To Len(s) - 1) Else ReDim ab(0)
    Dim cab ' As SysInt
    ' Copy to existing array, padding or truncating if needed
    cab = UBound(ab) - LBound(ab) + 1
    If Len(s) < cab Then s = s & String(cab - Len(s), 0)
    CopyMemory ab(LBound(ab)), ByVal s, cab
    
    'If IsArrayEmpty(ab) Then
    '    ' Just assign to empty array
    '    ab = StrConv(s, vbFromUnicode)
    'Else
    '    Dim cab ' As SysInt
    '    ' Copy to existing array, padding or truncating if needed
    '    cab = UBound(ab) - LBound(ab) + 1
    '    If Len(s) < cab Then s = s & String(cab - Len(s), 0)
    '    CopyMemory ab(LBound(ab)), ByVal s, cab
    'End If
End Sub

Function BytesToStr(ab() As Byte) As String
    'BytesToStr = StrConv(ab(), vbUnicode)
    BytesToStr = String(LenBytes(ab), 0)
    CopyMemory ByVal BytesToStr, ab(LBound(ab)), LenBytes(ab)
End Function

' Read string with length in first byte
Function BytesToPStr(ab() As Byte, Optional iOffset As Long = 0) As String
    BytesToPStr = MidBytes(ab, iOffset + 1, ab(iOffset))
End Function

Function BytesToWord(abBuf() As Byte, Optional iOffset As Long = 0) As Integer
    Dim w As Integer
    CopyMemory w, abBuf(iOffset), 2
    BytesToWord = w
End Function

Function BytesToDWord(abBuf() As Byte, Optional iOffset As Long = 0) As Long
    Dim dw As Long
    CopyMemory dw, abBuf(iOffset), 4
    BytesToDWord = dw
End Function

Sub BytesFromWord(w As Long, abBuf() As Byte, Optional iOffset As Long = 0)
    CopyMemory abBuf(iOffset), w, 2
End Sub

Sub BytesFromDWord(dw As Long, abBuf() As Byte, Optional iOffset As Long = 0)
    CopyMemory abBuf(iOffset), dw, 4
End Sub

'' Emulate relevant Basic string functions for arrays of bytes:
''     Len$             LenBytes
''     Mid$ function    MidBytes
''     Mid$ statement   InsBytes sub
''     Left$            LeftBytes
''     Right$           RightBytes

' LenBytes - Emulates Len for array of bytes
Function LenBytes(ab() As Byte) As Long
    LenBytes = UBound(ab) - LBound(ab) + 1
End Function

' MidBytes - emulates Mid$ function for array of bytes
' (Note that MidBytes does not emulate Mid$ exactly--string fields
' in byte arrays are often null-padded, and MidBytes can extract
' non-null portion.)
Function MidBytes(ab() As Byte, iOffset, Optional vLen As Variant, _
                  Optional vToNull As Variant) As String
    Dim s As String, fToNull As Boolean, cab ' As SysInt
    If Not IsMissing(vToNull) Then fToNull = vToNull
    ' Calculate length
    If IsMissing(vLen) Then
        cab = LenBytes(ab) - iOffset
    Else
        cab = vLen
    End If
    ' Assign and return string
    s = String$(cab, 0)
    CopyMemory ByVal s, ab(iOffset), cab
    If fToNull Then
        cab = InStr(s & vbNullChar, vbNullChar)
        MidBytes = Left$(s, cab - 1)
    Else
        MidBytes = s
    End If
End Function

' InsBytes - Emulates Mid$ statement for array of bytes
' (Note that InsBytes does not emulate Mid$ exactly--it inserts
' a null-padded string into a fixed-size field in order to work
' better with common use of byte arrays.)
Sub InsBytes(sIns As String, ab() As Byte, iOffset, _
             Optional vLen As Variant, Optional sPad As Byte = 0)
    Dim cab ' As SysInt
    ' Calculate length
    If IsMissing(vLen) Then
        cab = Len(sIns)
    Else
        cab = vLen
        ' Null-pad insertion string if too short
        If (Len(sIns) < cab) Then
            sIns = sIns & String$(cab - Len(sIns), sPad)
        End If
    End If
    ' Insert string
    CopyMemory ab(iOffset), ByVal sIns, cab
End Sub

' LeftBytes - Emulates Left$ function for array of bytes
Function LeftBytes(ab() As Byte, iLen) As String
    Dim s As String
    s = String$(iLen, 0)
    CopyMemory ByVal s, ab(LBound(ab)), iLen
    LeftBytes = s
End Function

' RightBytes - Emulates Right$ function for array of bytes
Function RightBytes(ab() As Byte, iLen) As String
    Dim s As String
    s = String$(iLen, 0)
    CopyMemory ByVal s, ab(UBound(ab) - iLen + 1), iLen
    RightBytes = s
End Function

' FillBytes - Fills field in array of bytes with given byte
Sub FillBytes(ab() As Byte, Optional b As Byte = 0, Optional iOffset As Long = 0, Optional iLen As Long)
    Dim i ' As SysInt
    If IsArrayEmpty(ab) Then ReDim ab(iLen)
    If IsMissing(iLen) Then iLen = UBound(ab) - iOffset + 1
    For i = iOffset To iOffset + iLen - 1
        ab(i) = b
    Next
End Sub

' InStrBytes is not implemented because a simple version would
' simply be equivalent to InStr(ab(), s). This creates a temporary
' string for ab() on every call. An efficient version that works
' directly on arrays of bytes could be written in C.

'However here is my release
Function InStrBytes(ab() As Byte, b As Byte, Optional iOffset As Long = 0) As Long
    Dim i ' As SysInt
    For i = iOffset To UBound(ab)
        If ab(i) = b Then
            InStrBytes = i
            Exit Function
        End If
    Next
    InStrBytes = -1 'Not found
End Function

Function IsArrayEmpty(ab() As Byte) As Boolean
    Dim v As Variant
    On Error Resume Next
    v = ab(LBound(ab))
    IsArrayEmpty = (Err <> 0)
End Function
