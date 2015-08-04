Attribute VB_Name = "Russian"
Option Explicit
Option Compare Text
Option Base 1
DefLng A-Z

Public Function RDate(v As Variant) As Date 'instead CDate()
    Dim s As String
    On Error GoTo ErrDate1
    RDate = CDate(v)
    If IsDate(RDate) Then Exit Function
ErrDate1:
    On Error GoTo ErrDate2
    s = StrTran(CStr(v), ".", "/")
    RDate = CDate(s)
    If IsDate(RDate) Then Exit Function
ErrDate2:
    On Error GoTo ErrDate3
    s = StrTran(CStr(v), "/", ".")
    RDate = CDate(s)
    If IsDate(RDate) Then Exit Function
ErrDate3:
    'rdate=?
    RDate = Now
End Function

Public Function RIsDate(v As Variant) As Boolean 'instead IsDate()
    Dim s As String
    On Error GoTo ErrDate1
    RIsDate = IsDate(v)
    If RIsDate Then Exit Function
ErrDate1:
    On Error GoTo ErrDate2
    s = StrTran(CStr(v), ".", "/")
    RIsDate = IsDate(s)
    If RIsDate Then Exit Function
ErrDate2:
    On Error GoTo ErrDate3
    s = StrTran(CStr(v), "/", ".")
    RIsDate = IsDate(s)
    If RIsDate Then Exit Function
ErrDate3:
    RIsDate = False
End Function


