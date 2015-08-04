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
Dim mFile As String

Private Sub Class_Initialize()
    Reset
End Sub

Public Property Get List() As Variant
    List = mStrings
End Property

Public Property Let List(vNewValue As Variant)
    Dim i As Long, n As Long
    n = UBound(vNewValue) - LBound(vNewValue) + 1
    ReDim mStrings(n)
    For i = 1 To n
        mStrings(i) = vNewValue(i)
    Next
End Property

Public Function Save(Optional File As String, Optional WinToDos As Boolean = True) As Boolean
    Dim s As String
    Save = False
    If IsMissing(File) Then File = mFile
    s = Me.Text
    If WinToDos Then s = CDos(s)
    OutputFile File, s
    Save = IsFile(File)
End Function

Public Function Load(File As String, Optional DosToWin As Boolean = True) As Boolean
    Dim s As String
    Load = False
    If Not IsFile(File) Then Exit Function
    If FileLen(File) = 0 Then Exit Function
    mFile = File
    s = InputFile(File)
    If Len(s) = 0 Then Exit Function
    If DosToWin Then s = CWin(s)
    Me.Text = StrTran(s, vbTab, Space(8))
    Load = True
End Function

Public Property Get LineCount() As Long
    LineCount = UBound(mStrings) - LBound(mStrings) + 1
End Property

Public Property Get Line(No As Long) As String
    Line = vbNullString
    If No < LBound(mStrings) Or UBound(mStrings) < No Then Exit Property
    Line = mStrings(No)
End Property

Public Property Let Line(No As Long, ByVal vNewValue As String)
    If No < LBound(mStrings) Or UBound(mStrings) < No Then Exit Property
    mStrings(No) = vNewValue
End Property

Public Property Get File() As String
    File = mFile
End Property

Public Property Get Text() As String
    Dim i As Long
    Text = vbNullString
    For i = LBound(mStrings) To UBound(mStrings)
        Text = Text & mStrings(i) & vbCrLf
    Next
End Property

Public Property Let Text(ByVal vNewValue As String)
    StrToLines vNewValue, mStrings
End Property

Public Sub Reset()
    mFile = vbNullString
End Sub
