Attribute VB_Name = "KeyValue"
Option Explicit
Option Compare Text
Option Base 1
DefLng A-Z

Public Function CalcBankKey(BIC As String, LS As String)
    Mid(LS, 9, 1) = "0"
    CalcBankKey = CalcKeyValue(Right(BIC, 3) + LS) 'Кредитная организация
End Function

Public Function CalcRkcKey(BIC As String, LS As String)
    Mid(LS, 9, 1) = "0"
    CalcRkcKey = CalcKeyValue("0" + Mid(BIC, 5, 2) + LS) 'РКЦ
End Function

Public Function BankKeyValid(BIC As String, LS As String) As Boolean
    BankKeyValid = CalcKeyValue(Right(BIC, 3) + LS) = 0 'Кредитная организация
End Function

Public Function RkcKeyValid(BIC As String, LS As String) As Boolean
    RkcKeyValid = CalcKeyValue("0" + Mid(BIC, 5, 2) + LS) = 0 'РКЦ
End Function

Private Function CalcKeyValue(BICLS As String)
    Dim i, nSum, ab() As Byte
    
    StrToBytes ab, BICLS + "0" 'Align to 24 (step 3)
    For i = 0 To 23
        ab(i) = ab(i) - 48 'Asc("0")
    Next
    
    nSum = 0
    For i = 0 To 23 Step 3
        nSum = nSum + ab(i) * 7 Mod 10
        nSum = nSum + ab(i + 1) * 1 Mod 10
        nSum = nSum + ab(i + 2) * 3 Mod 10 'Aligned 0*x=0, it's no sense
    Next
    
    CalcKeyValue = nSum * 3 Mod 10
End Function

Public Function AutoFillLS(ByVal LS As String, Optional BIC As String) As String
    Const LS20 = "40702810К00000000000"
    Dim i, n
    i = InStr(1, LS, "-")
    If i = 0 Then
        LS = "-" & LS
        i = 1
    End If
    n = 20 - Len(LS) + 1 'chars to insert
    AutoFillLS = Left(LS, i - 1) & _
        Mid(LS20, i, n) & Right(LS, 20 - i - n + 1)
    If Len(BIC) > 0 Then
        i = InStr(1, AutoFillLS, "К")
        If i > 0 Then
            Mid(AutoFillLS, i, 1) = CStr(CalcBankKey(BIC, AutoFillLS))
        End If
    End If
End Function

Public Function AutoFillBIC(BIC As String) As String
    Const BIC9 = "044030000"
    Dim n
    n = 9 - Len(BIC) 'chars to insert
    AutoFillBIC = Left(BIC9, n) & BIC
End Function


