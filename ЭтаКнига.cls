VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Ёта нига"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
0 Option Explicit
Option Compare Text
Option Base 1
DefLng A-Z

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    CloseBar
End Sub

Private Sub Workbook_Open()
    AddBar
End Sub





