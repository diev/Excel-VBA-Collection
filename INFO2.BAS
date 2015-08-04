Attribute VB_Name = "Info"
Option Explicit
Option Compare Text
Option Base 1
DefLng A-Z

Public Function BnkClienInfo() As String
    With ThisWorkbook
        BnkClienInfo = FileVerInfo(.FullName, "Основная программа ~Банк-Клиент~")
    End With
End Function

Public Function Plat999Info() As String
    With ActiveWorkbook
        Plat999Info = FileVerInfo(.FullName, "Данные пользователя ~Банк-Клиент~")
    End With
End Function

Public Function BnkSeekInfo() As String
    BnkSeekInfo = FileVerInfo(Range("BnkSeek").Text, "Справочник Банков", True)
End Function

Public Sub VerInfo()
    Dim VerInfo As String
    VerInfo = BPrintF("Версии установленных файлов:\n\n%s\n\n%s\n\n%s\n\n%s", _
        Plat999Info, BnkClienInfo, BnkUtilsInfo, BnkSeekInfo)
    MsgBox VerInfo, vbInformation, AppTitle
End Sub

Public Sub Demo()
    MsgBox BPrintF("Действие недоступно в этой версии\nЗапросите обновленную версию\n\n%s", BnkClienInfo), vbExclamation, AppTitle
End Sub
