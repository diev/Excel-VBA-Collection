Attribute VB_Name = "Info"
Option Explicit
Option Compare Text
Option Base 1
DefLng A-Z

Public Function BnkUtilsInfo() As String
    With ThisWorkbook
        BnkUtilsInfo = FileVerInfo(.FullName, "Банковские вспомогательные утилиты")
    End With
End Function

Public Function FileVerInfo(FileName As String, Optional FileComment As String = "", Optional FileUpdated As Boolean = False) As String
    Dim n
    If Len(FileComment) > 0 Then
        FileVerInfo = BPrintF("%s - " & FileComment & "\n~%s~ ", FileNameExt(FileName), FileName)
    Else
        FileVerInfo = BPrintF("~%s~ ", FileName)
    End If
    If FileExists(FileName) Then
        n = FileLen(FileName) \ 1024
        If n = 0 Then n = 1
        FileVerInfo = FileVerInfo & BPrintF("(%dk) - %y", n, FileDateTime(FileName))
        If FileUpdated Then
            FileVerInfo = FileVerInfo & BPrintF("\nот последнего обновления прошло %d дней", _
                DateDiff("d", FileDateTime(FileName), Date))
        End If
    Else
        FileVerInfo = FileVerInfo & "не найден!"
    End If
End Function
