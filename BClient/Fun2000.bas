Attribute VB_Name = "Fun2000"
Option Explicit
Option Compare Text
Option Base 1
DefLng A-Z

'Join() in Office2000
Public Function ArrToStr(Arr As Variant, Optional Delim As String = " ", Optional KeepEmpty As Boolean = True) As String
'    On Error GoTo Join97
'    ArrToStr = VBA.Join(arr, Delim)
'    Exit Function
'
'Join97:
    Dim i, s As String
    ArrToStr = vbNullString
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
Public Function StrToArr(s As String, Optional Delim As String = " ", _
    Optional KeepEmpty As Boolean = True) As Variant
'    On Error GoTo Split97
'    StrToArr = VBA.Split(s, Delim, , vbTextCompare)
'    Exit Function
'
'Split97:
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
            Arr(n) = ss
            n = n + 1
        ElseIf Len(s) > 0 Then
            Arr(n) = ss
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
Public Function StrTran(Where As String, What As String, Dest As String, _
    Optional Count As Long = 0, Optional Start As Long = 1) As String
'    On Error GoTo Replace97
'    StrTran = VBA.Replace(Where, What, Dest, , , vbTextCompare)
'    Exit Function
'
'Replace97:
    Dim i: i = Start - 1
    StrTran = Where
    Do
        i = InStr(i + 1, StrTran, What)
        If i > 0 Then StrTran = Left(StrTran, i - 1) & Dest & Mid(StrTran, i + Len(What))
        Count = Count - 1
    Loop Until i = 0 Or Count = 0
End Function

'InStrRev() in Office2000
Public Function InStrR(Where As String, What As String, _
    Optional after As Long = 1, Optional before As Long = 0) As Long
'    On Error GoTo InStrRev97
'    InStrR = VBA.InStrRev(Where, What, , vbTextCompare)
'    Exit Function
'
'InStrRev97:
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
