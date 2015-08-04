Attribute VB_Name = "Pads"
Option Explicit
Option Compare Text
Option Base 1
DefLng A-Z

Public Function Pad(Text As String, Length As Long, Optional Char As Variant = 32) As String
    Pad = PadR(Text, Length, Char)
End Function

Public Function PadR(Text As String, Length As Long, Optional Char As Variant = 32) As String
    If Length > 0 Then
        PadR = Left(Text & String(Length, Char), Length)
    Else
        PadR = Trim(Text)
    End If
End Function

Public Function PadL(Text As String, Length As Long, Optional Char As Variant = 32) As String
    If Length > 0 Then
        PadL = Right(String(Length, Char) & Text, Length)
    Else
        PadL = Trim(Text)
    End If
End Function

Public Function PadC(Text As String, Length As Long, Optional Char As Variant = 32) As String
    If Length > 0 Then
        PadC = Mid(String(Length, Char) & Text & String(Length, Char), (Length - Len(Text)) \ 2, Length)
    Else
        PadC = Trim(Text)
    End If
End Function 'doesn't work!

Public Function PadLine(Text As String, Length As Long, Optional Delim As String = "|", Optional Char As Variant = 32) As String
    Dim LText As String, CText As String, RText As String, n1, n2
    n1 = InStr(1, Text, Delim)
    n2 = InStrR(Text, Delim)
    If n1 = 0 Then
        PadLine = PadR(Text, Length, Char)
    Else
        LText = Left(Text, n1 - 1)
        CText = Mid(Text, n1 + 1, n2 - n1 - 1)
        RText = Mid(Text, n2 + 1)
        PadLine = LText & PadC(CText, Length - Len(LText) - Len(RText), Char) & RText
    End If
End Function

Public Function PadLine2(LText As String, RText As String, Length As Long, Optional Delim As String = "|", Optional Char As Variant = 32) As String
    PadLine2 = PadR(LText, Length - Len(RText), Char) & RText
End Function

Public Function PadLine3(LText As String, CText As String, RText As String, Length As Long, Optional Delim As String = "|", Optional Char As Variant = 32) As String
    PadLine3 = LText & PadC(CText, Length - Len(LText) - Len(RText), Char) & RText
End Function

