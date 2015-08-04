VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CInserts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit
Option Compare Text
Option Base 1
DefLng A-Z

'Подстановки
Const mInsertsSheet = "Подстановки"
Const mInsertsText = 1 'Подстановки
Const mInsertsNDS = 2 'НДС
'Итого
Const mInsertsSize = 2 'Общее число столбцов

Dim mInserts As New CDataSheet

Private Sub Class_Initialize()
    mInserts.Sheet = mInsertsSheet
End Sub

Public Function List(Column As Long) As Variant
    Dim arr As Variant
    arr = mInserts.GetStrings(Column)
    If Not IsArray(arr) Then
        ReDim arr(1) As String
        arr(1) = ""
    End If
    List = arr
End Function
