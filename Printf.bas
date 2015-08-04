Attribute VB_Name = "Printf"
Option Explicit
Option Compare Binary 'BINARY!!!
Option Base 1
DefLng A-Z

Dim simple As Boolean, realign As Boolean
Dim width As String, part As String, before As Boolean, argc As Long

Public Function BPrintF(FormatStr As String, ParamArray Args() As Variant) As String
    'like the C function sprintf()
    'required to investigate user32.wvsprintfA()!
    Dim n1 As Long, n2 As Long, c As String, s As String
    BPrintF = vbNullString
    n1 = 1
    Do
        n2 = InStr(n1, FormatStr, "\")
        If n2 = 0 Then
            BPrintF = BPrintF & Mid(FormatStr, n1)
            Exit Do
        Else
            BPrintF = BPrintF & Mid(FormatStr, n1, n2 - n1)
        End If
        c = Mid(FormatStr, n2 + 1, 1)
        Select Case c
            Case "\":
                BPrintF = BPrintF & "\"
            Case "'":
                BPrintF = BPrintF & """" 'like C \"
            Case "a":
                Beep
            Case "b":
                BPrintF = Left(BPrintF, Len(BPrintF) - 1) '''''''''''''''''need to be last
            Case "n":
                BPrintF = BPrintF & vbCrLf
            'Case "r":
            '    BPrintF = BPrintF & vbLf 'UNIX
            Case "t":
                BPrintF = BPrintF & vbTab
            Case "0":
                BPrintF = BPrintF & vbNullChar
            Case Else
                BPrintF = BPrintF & c
        End Select
        n1 = n2 + 2: n2 = n1
    Loop
    
    argc = LBound(Args)
    n1 = 1
    FormatStr = BPrintF
    BPrintF = vbNullString
    Do
        simple = True
        realign = False
        width = vbNullString
        part = vbNullString
        before = True
        n2 = InStr(n1, FormatStr, "%")
        If n2 = 0 Then
            BPrintF = BPrintF & Mid(FormatStr, n1)
            Exit Do
        Else
            BPrintF = BPrintF & Mid(FormatStr, n1, n2 - n1)
        End If
        Do
            n2 = n2 + 1
            c = Mid(FormatStr, n2, 1)
            Select Case c
                Case "%":
                    BPrintF = BPrintF & "%"
                    Exit Do
                Case "c":
                    BPrintF = BPrintF & Chr(Args(argc))
                    argc = argc + 1
                    Exit Do
                Case "d":
                    BPrintF = BPrintF & DFormat(Args(argc))
                    argc = argc + 1
                    Exit Do
                Case "s":
                    BPrintF = BPrintF & SFormat(Args(argc))
                    argc = argc + 1
                    Exit Do
                    
                'Changed behavior!
                Case "f":
                    BPrintF = BPrintF & FFormat(Args(argc))
                    argc = argc + 1
                    Exit Do
                Case "F":
                    BPrintF = BPrintF & PFormat(Args(argc))
                    argc = argc + 1
                    Exit Do
                Case "x":
                    BPrintF = BPrintF & DFormat(LCase(To36(Args(argc))))
                    argc = argc + 1
                    Exit Do
                Case "X":
                    BPrintF = BPrintF & DFormat(UCase(To36(Args(argc))))
                    argc = argc + 1
                    Exit Do
                    
                'Extra formats!!!
                Case "n":
                    BPrintF = BPrintF & Format(Args(argc), "dd.MM.yyyy")
                    argc = argc + 1
                    Exit Do
                Case "N":
                    BPrintF = BPrintF & DtoS(Args(argc))
                    argc = argc + 1
                    Exit Do
                Case "t":
                    BPrintF = BPrintF & Format(Args(argc), "HH:mm")
                    argc = argc + 1
                    Exit Do
                Case "T":
                    BPrintF = BPrintF & Format(Args(argc), "dd.MM.yyyy HH:mm")
                    argc = argc + 1
                    Exit Do
                
                'Digital preprocessing
                Case "0":
                    simple = False
                    If Len(width) = 0 Then width = "0"
                    If before Then
                        width = width & c
                    Else
                        part = part & c
                    End If
                Case "1" To "9":
                    simple = False
                    If Len(width) = 0 Then width = " "
                    If before Then
                        width = width & c
                    Else
                        part = part & c
                    End If
                Case "-":
                    simple = False
                    realign = True
                Case "*":
                    simple = False
                    If Len(width) = 0 Then width = " "
                    If before Then
                        width = width & CStr(Args(argc))
                    Else
                        part = CStr(Args(argc))
                    End If
                    argc = argc + 1
                Case ".":
                    simple = False
                    before = False
                    If Len(width) = 0 Then width = " "
                
                'Something goes wrong...
                Case Else
                    BPrintF = BPrintF & c
                    Exit Do
            End Select
        Loop
        n1 = n2 + 1
    Loop
End Function

Public Function DFormat(v As Variant) As String
    Dim s As String
    s = CStr(v)
    If Not simple Then
        If realign Then 'align to left
            s = PadR(s, Val(width), Left(width, 1))
            If Len(part) > 0 Then s = Right(s, Val(part))
        Else 'align to right
            s = PadL(s, Val(width), Left(width, 1))
            If Len(part) > 0 Then s = Left(s, Val(part))
        End If
    End If
    DFormat = s
End Function

Public Function FFormat(v As Variant, Optional Delim As String = ".") As String
    Dim s As String
    If simple Then
        s = Format(v, "#,0.00")
        Mid(s, Len(s) - 2, 1) = Delim
    Else
        s = Format(v, "#,0." & String(Val(part), "0"))
        Mid(s, Len(s) - Val(part), 1) = Delim
        If realign Then 'align to left
            s = PadR(s, Val(width), Left(width, 1))
        Else 'align to right
            s = PadL(s, Val(width), Left(width, 1))
        End If
    End If
    FFormat = s
End Function

Public Function PFormat(v As Variant, Optional Delim As String = "-") As String
    Dim s As String
    If simple Then
        s = Format(v, "0.00")
        Mid(s, Len(s) - 2, 1) = Delim
    Else
        s = Format(v, "0." & String(Val(part), "0"))
        Mid(s, Len(s) - Val(part), 1) = Delim
        If realign Then 'align to left
            s = PadR(s, Val(width), Left(width, 1))
        Else 'align to right
            s = PadL(s, Val(width), Left(width, 1))
        End If
    End If
    PFormat = s
End Function

Public Function SFormat(v As Variant) As String
    Dim s As String
    s = CStr(v)
    If Not simple Then
        If realign Then 'align to right
            s = PadL(s, Val(width), Left(width, 1))
            If Len(part) > 0 Then s = Right(s, Val(part))
        Else 'align to left
            s = PadR(s, Val(width), Left(width, 1))
            If Len(part) > 0 Then s = Left(s, Val(part))
        End If
    End If
    SFormat = s
End Function

