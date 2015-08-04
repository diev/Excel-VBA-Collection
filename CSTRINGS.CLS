VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CStrings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Option Base 1
DefLng A-Z

Dim mStrings() As String
Dim mFileName As String
Dim mChanged As Boolean

Private Sub Class_Initialize()

End Sub

Private Sub Class_Terminate()

End Sub

Public Property Get List() As Variant

End Property

Public Property Let List(ByVal vNewValue As Variant)

End Property

Public Sub Save()

End Sub

Public Sub Load()

End Sub

Public Function Count()

End Function

Public Property Get Line() As Variant

End Property

Public Property Let Line(ByVal vNewValue As Variant)

End Property

Public Property Get Comment() As Variant

End Property

Public Property Let Comment(ByVal vNewValue As Variant)

End Property

Public Property Get Value() As Variant

End Property

Public Property Let Value(ByVal vNewValue As Variant)

End Property

Public Property Get Delimiter() As Variant

End Property

Public Property Let Delimiter(ByVal vNewValue As Variant)

End Property

Public Property Get Changed() As Variant

End Property

Public Property Let Changed(ByVal vNewValue As Variant)

End Property

Public Property Get FileName() As Variant

End Property

Public Property Let FileName(ByVal vNewValue As Variant)

End Property

Public Property Get LineFeed() As Variant

End Property

Public Property Let LineFeed(ByVal vNewValue As Variant)

End Property

Public Sub Add()

End Sub

Public Sub Delete()

End Sub
