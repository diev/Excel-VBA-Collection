Attribute VB_Name = "KeyValue"
Option Explicit
Option Compare Text
Option Base 1
DefLng A-Z

Public Function CalcLSKey(ByVal BIC As String, ByVal LS As String) As Long
    BIC = PadL(BIC, 9, "0"): Mid(LS, 9, 1) = "0"
    If Val(Right(BIC, 3)) < 4 Then 'РКЦ?
        CalcLSKey = CalcLSKeyValue("0" + Mid(BIC, 5, 2) + LS)
    Else 'Кредитная организация
        CalcLSKey = CalcLSKeyValue(Right(BIC, 3) + LS)
    End If
End Function

Public Function LSKeyValid(ByVal BIC As String, ByVal LS As String) As Boolean
    Dim KeyProbe As Long
    KeyProbe = Mid(LS, 9, 1)
    BIC = PadL(BIC, 9, "0"): Mid(LS, 9, 1) = "0"
    If Val(Right(BIC, 3)) < 4 Then 'РКЦ?
        LSKeyValid = CalcLSKeyValue("0" + Mid(BIC, 5, 2) + LS) = KeyProbe
    Else 'Кредитная организация
        LSKeyValid = CalcLSKeyValue(Right(BIC, 3) + LS) = KeyProbe
    End If
End Function

Public Function CalcLSKeyValue(BICLS As String) As Long
    Dim i As Long, nSum As Long, ab() As Byte
    
    StrToBytes BICLS + "0", ab 'Align to 24 (step 3)
    For i = 1 To 24
        ab(i) = ab(i) - 48 'Asc("0")
    Next
    
    nSum = 0
    For i = 1 To 24 Step 3
        nSum = nSum + ab(i) * 7 Mod 10
        nSum = nSum + ab(i + 1) ' * 1 Mod 10
        nSum = nSum + ab(i + 2) * 3 Mod 10 'Aligned 0*x=0, it's no sense
    Next
    
    CalcLSKeyValue = nSum * 3 Mod 10
End Function

Public Function SetLSKey(BIC As String, ByVal LS As String) As String
    Mid(LS, 9, 1) = CStr(CalcLSKey(BIC, LS))
    SetLSKey = LS
End Function

Public Function INNKeyValid(ByVal INN As String) As Boolean
    Dim i As Long, nSum As Long, ab() As Byte
    
    If Left(INN, 1) = "F" Then
        INN = Mid(INN, 2)
    End If
    Select Case Len(INN)
        Case 0
            INNKeyValid = True
        Case 1
            INNKeyValid = INN = "0"
        Case 10
            StrToBytes INN, ab
            For i = 1 To 10
                ab(i) = ab(i) - 48 'Asc("0")
            Next
            nSum = 0
            nSum = nSum + ab(1) * 2
            nSum = nSum + ab(2) * 4
            nSum = nSum + ab(3) * 10
            nSum = nSum + ab(4) * 3
            nSum = nSum + ab(5) * 5
            nSum = nSum + ab(6) * 9
            nSum = nSum + ab(7) * 4
            nSum = nSum + ab(8) * 6
            nSum = nSum + ab(9) * 8
            nSum = nSum Mod 11 Mod 10
            INNKeyValid = ab(10) = nSum
        Case 12
            StrToBytes INN, ab
            For i = 1 To 12
                ab(i) = ab(i) - 48 'Asc("0")
            Next
            nSum = 0
            nSum = nSum + ab(1) * 7
            nSum = nSum + ab(2) * 2
            nSum = nSum + ab(3) * 4
            nSum = nSum + ab(4) * 10
            nSum = nSum + ab(5) * 3
            nSum = nSum + ab(6) * 5
            nSum = nSum + ab(7) * 9
            nSum = nSum + ab(8) * 4
            nSum = nSum + ab(9) * 6
            nSum = nSum + ab(10) * 8
            nSum = nSum Mod 11 Mod 10
            If ab(11) <> nSum Then
                INNKeyValid = False
            Else
                'ab(11) = nSum
                nSum = 0
                nSum = nSum + ab(1) * 3
                nSum = nSum + ab(2) * 7
                nSum = nSum + ab(3) * 2
                nSum = nSum + ab(4) * 4
                nSum = nSum + ab(5) * 10
                nSum = nSum + ab(6) * 3
                nSum = nSum + ab(7) * 5
                nSum = nSum + ab(8) * 9
                nSum = nSum + ab(9) * 4
                nSum = nSum + ab(10) * 6
                nSum = nSum + ab(11) * 8
                nSum = nSum Mod 11 Mod 10
                INNKeyValid = ab(12) = nSum
            End If
        Case Else
            INNKeyValid = False
    End Select
End Function
