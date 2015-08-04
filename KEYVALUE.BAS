Attribute VB_Name = "KeyValue"
Option Explicit
Option Compare Text
Option Base 1
DefLng A-Z

Public Function CalcLSKey(ByVal BIC As String, ByVal LS As String, Optional RKC As Boolean = False)
    BIC = PadL(BIC, 9, "0"): Mid(LS, 9, 1) = "0"
    If RKC Then 'РКЦ
        CalcLSKey = CalcLSKeyValue("0" + Mid(BIC, 5, 2) + LS)
    Else 'Кредитная организация
        CalcLSKey = CalcLSKeyValue(Right(BIC, 3) + LS)
    End If
End Function

Public Function LSKeyValid(ByVal BIC As String, ByVal LS As String, Optional RKC As Boolean = False) As Boolean
    BIC = PadL(BIC, 9, "0"): Mid(LS, 9, 1) = "0"
    If RKC Then 'РКЦ
        LSKeyValid = CalcLSKeyValue("0" + Mid(BIC, 5, 2) + LS) = 0
    Else 'Кредитная организация
        LSKeyValid = CalcLSKeyValue(Right(BIC, 3) + LS) = 0
    End If
End Function

Public Function CalcLSKeyValue(BICLS As String) As Long
    Dim i, nSum, ab() As Byte
    
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

Public Function SetLSKey(BIC As String, ByVal LS As String, Optional RKC As Boolean = False) As String
    Mid(LS, 9, 1) = CStr(CalcLSKey(BIC, LS, RKC))
    SetLSKey = LS
End Function

