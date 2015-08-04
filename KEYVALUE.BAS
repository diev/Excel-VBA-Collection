Attribute VB_Name = "KeyValue"
Option Explicit
Option Compare Text
Option Base 1
DefLng A-Z

Public Function CalcKey(ByVal BIC As String, ByVal LS As String, Optional RKC As Boolean = False)
    BIC = PadL(BIC, 9, "0"): Mid(LS, 9, 1) = "0"
    If RKC Then 'РКЦ
        CalcKey = CalcKeyValue("0" + Mid(BIC, 5, 2) + LS)
    Else 'Кредитная организация
        CalcKey = CalcKeyValue(Right(BIC, 3) + LS)
    End If
End Function

Public Function KeyValid(ByVal BIC As String, ByVal LS As String, Optional RKC As Boolean = False) As Boolean
    BIC = PadL(BIC, 9, "0"): Mid(LS, 9, 1) = "0"
    If RKC Then 'РКЦ
        KeyValid = CalcKeyValue("0" + Mid(BIC, 5, 2) + LS) = 0
    Else 'Кредитная организация
        KeyValid = CalcKeyValue(Right(BIC, 3) + LS) = 0
    End If
End Function

Private Function CalcKeyValue(BICLS As String) As Long
    Dim i, nSum, ab() As Byte
    
    StrToBytes BICLS + "0", ab 'Align to 24 (step 3)
    For i = 1 To 24
        ab(i) = ab(i) - 48 'Asc("0")
    Next
    
    nSum = 0
    For i = 1 To 24 Step 3
        nSum = nSum + ab(i) * 7 Mod 10
        nSum = nSum + ab(i + 1) * 1 Mod 10
        nSum = nSum + ab(i + 2) * 3 Mod 10 'Aligned 0*x=0, it's no sense
    Next
    
    CalcKeyValue = nSum * 3 Mod 10
End Function

Public Function AutoFillLS(BIC As String, ByVal LS As String, Optional RKC As Boolean = False) As String
    If Len(LS) = 0 Then LS = App.DefLS
    If InStr(1, BIC & LS, "-") Then
        MsgBox BPrintF("Символ \'-\' в этой версии не поддерживается!"), _
            vbExclamation, App.Title
        AutoFillLS = App.DefLS
    Else
        LS = Left(App.DefLS, 20 - Len(LS)) & LS
        Mid(LS, 9, 1) = CStr(CalcKey(AutoFillBIC(BIC), LS, RKC))
        AutoFillLS = LS
    End If
End Function

Public Function AutoFillBIC(BIC As String) As String
    If Len(BIC) = 0 Then
        AutoFillBIC = App.DefBIC
    ElseIf InStr(1, BIC, "-") Then
        MsgBox BPrintF("Символ \'-\' в этой версии не поддерживается!"), _
            vbExclamation, App.Title
        AutoFillBIC = App.DefBIC
    Else
        AutoFillBIC = Left(App.DefBIC, 9 - Len(BIC)) & BIC
    End If
End Function

