VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   5580
   ClientLeft      =   1050
   ClientTop       =   1380
   ClientWidth     =   9915
   OleObjectBlob   =   "UserForm1.frx":0000
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'UserForm1 + ListBox1 + ListBox2

Option Explicit
Option Base 0 '1
DefLng A-Z

Dim ActiveBox As Integer

Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    With ListBox1
        If .ListIndex > 0 Then
            Cells(ActiveRow, .ListIndex).Select
            Close Me
        End If
    End With
End Sub

Private Sub ListBox1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    CommonKeyPress KeyAscii
End Sub

Private Sub ListBox2_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    With ListBox2
        If .ListIndex > 0 Then
            Cells(ActiveRow, .ListIndex + TU0 - 1).Select
            Close Me
        End If
    End With
End Sub

Private Sub ListBox2_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    CommonKeyPress KeyAscii
    With ListBox2
        If .ListIndex > 0 Then
            Select Case KeyAscii
                Case Asc("0")
                    Cells(ActiveRow, .ListIndex + TU0 - 1).Select
                    Me.Hide
                Case Asc("1")
                    Cells(ActiveRow, .ListIndex + TU1 - 1).Select
                    Me.Hide
                Case Asc("2")
                    Cells(ActiveRow, .ListIndex + TU2 - 1).Select
                    Me.Hide
                Case Asc("3")
                    Cells(ActiveRow, .ListIndex + TU3 - 1).Select
                    Me.Hide
                Case Asc("4")
                    Cells(ActiveRow, .ListIndex + TU4 - 1).Select
                    Me.Hide
            End Select
        End If
    End With
End Sub

Private Sub CommonKeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    'Debug.Print KeyAscii
    Select Case KeyAscii
        Case 45 'Gray-
            If ActiveRow - 1 > THRow Then
                ActiveRow = ActiveRow - 1
                FillByRow
            End If
        Case 32, 43 'Space, Gray+
            If Len(Cells(ActiveRow + 1, PRIM_1).Text) > 0 Then
                ActiveRow = ActiveRow + 1
                FillByRow
            End If
        Case 13, 27 'Enter, Esc
            Me.Hide
    End Select
End Sub

Private Sub UserForm_Initialize()
    Dim c As Integer, i As Integer
    
    For i = 1 To LAST
        TH(i) = ActiveSheet.Cells(THRow, i).Text
    Next i
    ActiveRow = ActiveCell.Row
    
    With Me
        .Caption = "Строка " & ActiveRow
        .Move Application.Left, Application.Top, Application.width, Application.Height
        ListBox1.Move 0, 0, .InsideWidth \ 5, .InsideHeight
        ListBox2.Move ListBox1.width, 0, .InsideWidth - ListBox1.width, .InsideHeight
    End With
    
    With ListBox1
        .ColumnCount = 2
        .ColumnWidths = "2 cm;10 cm"
        .AddItem "Поле:"
        For c = VERSION To PRIZ_SD
            If NotResrvFld1(c) Then
                .AddItem Fld(c)
            End If
        Next c
        .List(0, 1) = "Информация об операции:"
    End With
    
    With ListBox2
        .ColumnCount = 5
        '.ColumnWidths = "2 cm;;;;"
        .AddItem "Поле:"
        For c = TU0 To RESRV_B02
            If NotResrvFld2(c) Then
                .AddItem Fld(c)
            End If
        Next c
    End With
    
    FillByRow
End Sub

Public Sub FillByRow()
    Dim t As Integer, c As Integer, i As Integer
    Dim cw As String
    
    For i = 1 To LAST
        TR(i) = ActiveSheet.Cells(ActiveRow, i).Text
    Next i
    
    With Me
        .Caption = "Строка " & ActiveRow & ": " & TR(PRIM_1)
    End With
    
    With ListBox1
        .ForeColor = QBColor(0)
        i = 1
        For c = VERSION To PRIZ_SD
            If NotResrvFld1(c) Then
                .List(i, 1) = Descr1(c)
                i = i + 1
            End If
        Next c
    
        .ListIndex = 0
    End With
    
    With ListBox2
        .ForeColor = QBColor(0)
        cw = "2 cm"
        
        t = 1: .List(0, t) = "0. Сведения о лице:"
        cw = cw & ";"
        i = 1
        For c = TU0 To RESRV_B02
            If NotResrvFld2(c) Then
                .List(i, t) = Descr2(c)
                i = i + 1
            End If
        Next c
        
        t = 2: .List(0, t) = "1. Пред-ль лица:"
        cw = cw & ";"
        If TR(TU1) = "0" Then
            .List(0, t) = "1."
            cw = cw & "1 cm"
        End If
        i = 1
        For c = TU1 To RESRV12
            If NotResrvFld2(c) Then
                .List(i, t) = Descr2(c)
                i = i + 1
            End If
        Next c
        
        t = 3: .List(0, t) = "2. Пред-ль получателя:"
        cw = cw & ";"
        If TR(TU2) = "0" Then
            .List(0, t) = "2."
            cw = cw & "1 cm"
        End If
        i = 1
        For c = TU2 To RESRV22
            If NotResrvFld2(c) Then
                .List(i, t) = Descr2(c)
                i = i + 1
            End If
        Next c
        
        t = 4: .List(0, t) = "3. Получатель:"
        cw = cw & ";"
        i = 1
        For c = TU3 To RESRV_B32
            If NotResrvFld2(c) Then
                .List(i, t) = Descr2(c)
                i = i + 1
            End If
        Next c
        
'        t = 5: .List(0, t) = "4. Третье лицо:"
'        i = 1
'        For c = TU4 To RESERV612
'            If NotResrvFld2(c) Then
'                .List(i, t) = Descr2(c)
'                i = i + 1
'            End If
'        Next c
'
'
'
'
'        For c = TU0 To BP_0
'            If NotResrvFld2(c) Then
'                .List(i, 1) = Descr2(c) 'from 38 to 86
'                .List(i, 2) = Descr2(c + 49) 'from 87 to 121
'                .List(i, 3) = Descr2(c + 85) 'from 123 to 157
'                .List(i, 4) = Descr2(c + 121) 'from 159 to 207
'                '.List(i, 5) = Descr2(c + 170) 'from 208 to 242
'                i = i + 1
'            End If
'        Next c
'        For c = VP_0 To RESRV_B02
'            If NotResrvFld2(c) Then
'                .List(i, 1) = Descr2(c) 'from 38 to 86
'                .List(i, 4) = Descr2(c + 121) 'from 159 to 207
'                i = i + 1
'            End If
'        Next c
'
'        i = 1
'        For c = TU1 To RESRV12
'            If NotResrvFld2(c) Then
'                .List(i, 2) = Descr2(c)
'                i = i + 1
'            End If
'        Next c
'
'        i = 1
'        For c = TU2 To RESRV22
'            If NotResrvFld2(c) Then
'                .List(i, 3) = Descr2(c)
'                i = i + 1
'            End If
'        Next c
'
'        i = 1
'        For c = TU3 To RESRV_B32
'            If NotResrvFld2(c) Then
'                .List(i, 4) = Descr2(c)
'                i = i + 1
'            End If
'        Next c
'
'        'i = 1
'        'For c = TU4 To RESERV612
'        '    If NotResrvFld2(c) Then
'        '        .List(i, 5) = Descr2(c)
'        '        i = i + 1
'        '    End If
'        'Next c
        
        .ColumnWidths = cw
        .ListIndex = 0
    End With
End Sub

Public Function Fld(c As Integer) As String
    Dim s As String
    s = TH(c)
    Fld = Left(s, InStr(s, " ") - 1)
    If (Right(Fld, 1) = "0") Then Fld = Left(Fld, Len(Fld) - 1)
End Function

Public Function NotResrvFld1(c As Integer) As Boolean
    Select Case c
        Case VERSION, REFER_R2, NUMBF_S, BRANCH, KTU_SS, BIK_SS, NUMBF_SS
            NotResrvFld1 = False
        Case Else
            NotResrvFld1 = True
    End Select
End Function

Public Function NotResrvFld2(c As Integer) As Boolean
    Dim s As String
    s = Left(TH(c), 5)
    NotResrvFld2 = s <> "RESRV"
End Function

Public Function Descr1(c As Integer) As String
    Dim r As String, s As String
    s = TR(c)
    r = s 'default
    ActiveBox = 1
    Select Case c
        Case VERSION
            Select Case s
                Case "2"
                    r = s & " - такая версия"
                Case Else
                    r = Er(s)
            End Select
        Case ACTION
            Select Case s
                Case "1"
                    r = s & " - добавление"
                Case "2"
                    r = s & " - исправление"
                Case "3"
                    r = s & " - замена"
                Case "4"
                    r = s & " - удаление"
                Case Else
                    r = Er(s)
            End Select
        Case REGN
            Select Case s
                Case "3194"
                    r = s & " - рег/н"
                Case Else
                    r = Er(s)
            End Select
        Case ND_KO
            Select Case s
                Case "7831001422"
                    r = s & " - ИНН"
                Case Else
                    r = Er(s)
            End Select
        Case KTU_S
            Select Case s
                Case "40"
                    r = s & " - ОКАТО"
                Case Else
                    r = Er(s)
            End Select
        Case BIK_S
            Select Case s
                Case "044030702"
                    r = s & " - БИК"
                Case Else
                    r = Er(s)
            End Select
        Case NUMBF_S
            If s <> "0" Then r = Er(s)
        Case BRANCH
            If s <> "0" Then r = Er(s)
        Case KTU_SS
            If s <> "0" Then r = Er(s)
        Case BIK_SS
            If s <> "0" Then r = Er(s)
        Case NUMBF_SS
            If s <> "0" Then r = Er(s)
        Case TERROR
            Select Case s
                Case "1"
                    r = s & " - приостановление"
                Case "2"
                    r = s & " - совершение"
                Case "0"
                    r = s & " - иное"
                Case Else
                    r = Er(s)
            End Select
        Case CURREN
            Select Case s
                Case "643"
                    r = s & " - рубли"
                Case "840"
                    r = s & " - доллары"
                Case "978"
                    r = s & " - евро"
                Case Else
                    r = Er(s)
            End Select
        Case DATE_S, DATE_PAY_D
            Select Case s
                Case "01.01.2099"
                    'r = s & " - нет"
                    r = "н/д"
                Case Else
                    If Not IsDate(s) Then r = Er(s)
            End Select
        Case DATA, DATE_P 'дата должна быть!
            Select Case s
                Case "01.01.2099"
                    'r = Er(s)
                    r = Er("н/д")
                Case Else
                    If Not IsDate(s) Then r = Er(s)
            End Select
        Case B_PAYER
            Select Case s
                Case "1"
                    r = s & " - клиент"
                Case "2"
                    r = s & " - банк"
                Case "0"
                    r = s & " - некто"
                Case Else
                    r = Er(s)
            End Select
        Case B_RECIP
            Select Case s
                Case "1"
                    r = s & " - клиент"
                Case "2"
                    r = s & " - банк"
                Case "0"
                    r = s & " - некто"
                Case Else
                    r = Er(s)
            End Select
        Case PART
            Select Case s
                Case "1"
                    r = s & " - от третьего лица"
                Case "2"
                    r = s & " - для третьего лица"
                Case "0"
                    r = s & " - без третьих лиц"
                Case Else
                    r = Er(s)
            End Select
        Case CURREN_CON
            Select Case s
                'Case "643"
                '    r = s & " - продажа рублей"
                Case "840"
                    r = s & " - продажа долларов"
                Case "978"
                    r = s & " - продажа евро"
                Case "0"
                    r = s & " - не конверсия"
                Case Else
                    r = Er(s)
            End Select
        Case PRIZ_SD
            Select Case s
                Case "0"
                    r = s & " - деньги"
                Case "1"
                    r = s & " - имущество"
                Case Else
                    r = Er(s)
            End Select
    End Select
    Descr1 = r
End Function

Public Function Descr2(c As Integer) As String
    Dim r As String, s As String
    s = TR(c)
    r = s 'default
    ActiveBox = 2
    Select Case c
        Case TU0
            Select Case s
                Case "1"
                    r = s & " - юрлицо"
                Case "2"
                    r = s & " - физлицо"
                Case "3"
                    r = s & " - ИП"
                Case "4"
                    If TR(B_PAYER) <> "0" Then
                        r = Er(s)
                    Else
                        r = s & " - не установлено"
                    End If
                Case Else
                    r = Er(s)
            End Select
        Case TU1, TU2
            Select Case s
                'Case "1" 'у нас это ошибка!
                '    r = s & " - юрлицо"
                Case "2"
                    r = s & " - физлицо"
                'Case "3" 'у нас это ошибка!
                '    r = s & " - ИП"
                Case "0"
                    r = s '& " - без представителя"
                Case Else
                    r = Er(s)
            End Select
        Case TU3
            Select Case s
                Case "1"
                    r = s & " - юрлицо"
                Case "2"
                    r = s & " - физлицо"
                Case "3"
                    r = s & " - ИП"
                Case "4"
                    If TR(B_RECIP) <> "0" Then
                        r = Er(s)
                    Else
                        r = s & " - не установлено"
                    End If
                Case Else
                    r = Er(s)
            End Select
        Case TU4
            Select Case s
                Case "1"
                    r = s & " - юрлицо"
                Case "2"
                    r = s & " - физлицо"
                Case "3"
                    r = s & " - ИП"
                Case Else
                    r = Er(s)
            End Select
        Case AMR_S0, ADRESS_S0, AMR_S1, ADRESS_S1, AMR_S2, ADRESS_S2, AMR_S3, ADRESS_S3, AMR_S4, ADRESS_S4
            Select Case s
                Case "00"
                    r = s & " - иностранец"
                Case "0"
                    r = s
                Case Else
                    r = s & " - ОКАТО"
            End Select
        Case ND0
            If s <> "0" Then
                Select Case TR(TU0)
                    Case "1"
                        If Len(s) <> 10 Then r = Er(s)
                    Case "2", "3"
                        If Len(s) <> 12 Then r = Er(s)
                    Case Else
                        r = Er(s)
                End Select
            End If
        Case ND1
            If s <> "0" Then
                Select Case TR(TU1)
                    Case "1"
                        If Len(s) <> 10 Then r = Er(s)
                    Case "2", "3"
                        If Len(s) <> 12 Then r = Er(s)
                    Case Else
                        r = Er(s)
                End Select
            End If
        Case ND2
            If s <> "0" Then
                Select Case TR(TU2)
                    Case "1"
                        If Len(s) <> 10 Then r = Er(s)
                    Case "2", "3"
                        If Len(s) <> 12 Then r = Er(s)
                    Case Else
                        r = Er(s)
                End Select
            End If
        Case ND3
            If s <> "0" Then
                Select Case TR(TU3)
                    Case "1"
                        If Len(s) <> 10 Then r = Er(s)
                    Case "2", "3"
                        If Len(s) <> 12 Then r = Er(s)
                    Case Else
                        r = Er(s)
                End Select
            End If
        Case ND4
            If s <> "0" Then
                Select Case TR(TU4)
                    Case "1"
                        If Len(s) <> 10 Then r = Er(s)
                    Case "2", "3"
                        If Len(s) <> 12 Then r = Er(s)
                    Case Else
                        r = Er(s)
                End Select
            End If
        Case VD11, VD21, VD31, VD41
            If InStr(s, " ") > 0 Then r = Er(s)
        Case VD03, VD06, VD07, MC_02, MC_03, VD13, VD16, VD17, MC_12, MC_13, VD23, VD26, VD27, MC_22, MC_23, VD33, VD36, VD37, MC_32, MC_33, VD43, VD46, VD47, MC_42, MC_43
            Select Case s
                Case "01.01.2099"
                    'r = s & " - нет"
                    r = "н/д"
                Case Else
                    If Not IsDate(s) Then r = Er(s)
            End Select
        Case GR0, GR1, GR2, GR3, GR4 'дата должна быть! (если кто есть вообще)
            Select Case s
                Case "01.01.2099"
                    If c = GR0 And TR(BIK_B0) <> "0" Then
                        'r = s & " - iao"
                        r = "н/д"
                    ElseIf c = GR1 And TR(TU1) = "0" Then
                        'r = s & " - нет"
                        r = "н/д"
                    ElseIf c = GR2 And TR(TU2) = "0" Then
                        'r = s & " - нет"
                        r = "н/д"
                    Else
                        'r = Er(s)
                        r = Er("н/д")
                    End If
                Case Else
                    If Not IsDate(s) Then r = Er(s)
            End Select
        Case VP_0, VP_3
            Select Case s
                Case "1"
                    r = s & " - выгодопр-ль идентифицирован"
                Case "2"
                    r = s & " - выгодопр-ль не идентифицирован"
                Case "0"
                    r = s
                Case Else
                    r = Er(s)
            End Select
        Case RESRV02, RESRV_B02, RESRV12, RESRV22, RESRV32, RESRV_B32, RESRV42, RESERV612 'резерв
            Select Case s
                Case "0"
                    r = s & " - резерв"
                Case Else
                    r = Er(s)
            End Select
        Case ACC_B0, ACC_COR_B0, ACC_B3, ACC_COR_B3
            Select Case s
                Case "0"
                    r = s
                Case Else
                    If Len(s) <> 20 Then
                        r = Er(s)
                    Else
                        r = s
                    End If
            End Select
        Case CARD_B0, CARD_B3
            Select Case s
                Case "1"
                    r = s & " - карта банка"
                Case "2"
                    r = s & " - карта чужого"
                Case "3"
                    r = s
                Case "0"
                    r = s '& " - не карта"
                Case Else
                    r = Er(s)
            End Select
    End Select
    Descr2 = r
End Function

Private Function Er(s As String) As String
    Er = "<!> " & s '& " = ОШИБКА!"
    Select Case ActiveBox
        Case 1
            ListBox1.ForeColor = QBColor(12) 'light red
        Case 2
            ListBox2.ForeColor = QBColor(12) 'light red
    End Select
End Function

Private Sub UserForm_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    CommonKeyPress KeyAscii
End Sub
