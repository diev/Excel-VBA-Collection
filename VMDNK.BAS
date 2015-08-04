Attribute VB_Name = "VMDNK"
Option Explicit
Option Compare Text
Option Base 1
DefLng A-Z

Public Sub MakeSheetVMD()
    Range("Файл") = Range("Посылка").Text & _
        GetMMDD(DtoS(Range("Дата"))) & _
        PadL(Range("Номер").Text, 3, "0") & "." & _
        Range("ID").Text  'Right(Range("Счет1").Text, 3)
End Sub

Public Function GetMMDD(YYYYMMDD As String) As String
    GetMMDD = Mid(YYYYMMDD, 4, 1) & to36(Val(Mid(YYYYMMDD, 5, 2))) & Right(YYYYMMDD, 2)
    'GetMMDD = Right(YYYYMMDD, 4)
End Function

