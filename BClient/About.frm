VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} About 
   Caption         =   "О программе ""Банк-Клиент"""
   ClientHeight    =   3810
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4845
   OleObjectBlob   =   "About.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "About"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Option Compare Text
Option Base 1
DefLng A-Z

Private Sub cmdOk_Click()
    Application.StatusBar = False
    Unload Me
End Sub

Private Sub cmdVerInfo_Click()
    VerInfo
End Sub

Private Sub UserForm_Initialize()
    Application.StatusBar = AppTitle
    lblInfo.Caption = _
        BPrintF("Программа %s\n\n" & _
        "Автор: %s, %d\n%s\n\n" & _
        "Проверьте файлы вашей версии\nи запросите обновление!", _
        AppTitle, "Дмитрий Евдокимов", 1999, "http://members.xoom.com/diev")
    lblAddress.Caption = BPrintF("Наш адрес:\n\n%s\n%s\n%s", _
        "Санкт-Петербург, Московский пр-т, 143", _
        "тел. (812) 324-06-90, факс 324-06-95", _
        "e-mail bclient@cib.lek.ru, web http://www.cityinvest.sp.ru")
End Sub
