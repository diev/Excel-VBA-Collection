Attribute VB_Name = "Export1C"
'This Module is not used because files for 1C are sent from the Bankier itself.

Option Explicit
Option Compare Text
Option Base 1
DefLng A-Z

Private Const T1 = "Электронная выписка из лицевого счета N "
Private Const T3 = "за период с "
Private Const T4 = " по "
Private Const T6 = "Входящее сальдо "
Private Const T7 = "Итого оборотов "

Public Sub test()
    ExportTo1CFile App.Path & "vyp01-14.1802.txt"
End Sub

Public Sub ExportTo1CFile(File As String)
    Dim FileNo As Long, FileNo2 As Long, Buf As String, n As Long, s As String
    Dim Pt As New CPayment, Debet As Boolean
    Dim CIBLS As String, CIBName As String
    Dim StartDate As Date, EndDate As Date
    Dim SIN As Currency, SDT As Currency, SKT As Currency, SOUT As Currency
    
    FileNo = FreeFile
    Open File For Input Access Read As FileNo
    Do While Not EOF(FileNo)
        Line Input #FileNo, Buf
        Buf = CWin(RTrim(Buf))
        If InStr(1, Buf, T1, vbTextCompare) = 1 Then
            CIBLS = Right(Buf, 20)
            
            Line Input #FileNo, Buf
            CIBName = CWin(Trim(Mid(Buf, 13)))
            
            Line Input #FileNo, Buf
            Buf = CWin(RTrim(Buf))
            If InStr(1, Buf, T3, vbTextCompare) = 1 Then
                n = InStr(1, Buf, T4, vbTextCompare)
                StartDate = RDate(Mid(Buf, Len(T3) + 1, n - Len(T3) - 1))
                EndDate = RDate(Mid(Buf, n + Len(T4)))
            Else
                StartDate = RDate(Mid(Buf, 4))
                EndDate = StartDate
            End If
        ElseIf InStr(1, Buf, T6, vbTextCompare) = 21 Then
            SIN = RVal(Mid(Buf, 63, 16))
        ElseIf InStr(1, Buf, T7, vbTextCompare) = 21 Then
            SDT = RVal(Mid(Buf, 46, 16))
            SKT = RVal(Mid(Buf, 63, 16))
        
            Line Input #FileNo, Buf
            SOUT = RVal(Mid(Buf, 63, 16))
        End If
    Loop
    Close FileNo
    
    Kill App.Path & "kl_to_1c.txt"
    FileNo2 = FreeFile
    Open App.Path & "kl_to_1c.txt" For Output Access Write As FileNo2
    Print #FileNo2, "1CClientBankExchange"
    Print #FileNo2, "ВерсияФормата=1.01"
    Print #FileNo2, "Кодировка=Windows"
    Print #FileNo2, "ДатаНачала="; StartDate
    Print #FileNo2, "ДатаКонца="; EndDate
    Print #FileNo2, "РасчСчет="; CIBLS
        
    Print #FileNo2, "СекцияРасчСчет"
    Print #FileNo2, "ДатаНачала="; StartDate
    Print #FileNo2, "ДатаКонца="; EndDate
    Print #FileNo2, "РасчСчет="; CIBLS
    Print #FileNo2, "НачальныйОстаток="; SumFormat(SIN)
    Print #FileNo2, "ВсегоПоступило="; SumFormat(SKT)
    Print #FileNo2, "ВсегоСписано="; SumFormat(SDT)
    Print #FileNo2, "КонечныйОстаток="; SumFormat(SOUT)
    Print #FileNo2, "КонецРасчСчет"
    
    'Print #FileNo2, ""

    FileNo = FreeFile
    Open File For Input Access Read As FileNo
    n = 0
    Do While Not EOF(FileNo)
        Line Input #FileNo, Buf
        If Left(Buf, 1) = "-" Then
            n = n + 1
        ElseIf n = 2 Then
            s = Left(Buf, 4)
            If RVal(s) > 0 Then
                With Pt
                    .DocNo = s
                    .DocDate = Mid(Buf, 6, 8)
                    .BIC = Mid(Buf, 15, 9)
                    .LS = Mid(Buf, 25, 20)
                    Debet = Mid(Buf, 61, 1) <> " "
                    If Debet Then
                        .Sum = RVal(Mid(Buf, 46, 16))
                    Else
                        .Sum = RVal(Mid(Buf, 63, 16))
                    End If
                    .Name = vbNullString ''''''''''''''''''''''''
                    .Details = vbNullString ''''''''''''''''''''''''
                End With
            
            End If
        End If
    Loop

    Print #FileNo2, "КонецФайла"
    Close FileNo
    Close FileNo2
End Sub
