VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Option Base 1
DefLng A-Z

Dim mSum As Currency

Public Property Get Value() As Variant
    Value = mSum
End Property

Public Property Let Value(ByVal vNewValue As Variant)
    mSum = RVal(CStr(vNewValue))
End Property

Public Property Get AmountText() As String
    AmountText = RSumStr(mSum)
End Property

Public Property Get PlatText() As String
    PlatText = PlatFormat(mSum)
End Property

Public Property Get Text() As String
    AmountText = Format(mSum, "#,##0.00")
End Property

Public Property Get Tax(Percent As Variant) As Variant
    Tax = Sum2Tax(mSum, CDbl(Percent))
End Property
