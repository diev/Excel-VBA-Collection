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

Const THRow As Integer = 3 'Table Header

Const VERSION As Integer = 1
Const ACTION As Integer = 2
Const NUMB_P As Integer = 3
Const DATE_P As Integer = 4
Const DATE_S As Integer = 5
Const TEL As Integer = 6
Const REFER_R2 As Integer = 7
Const REGN As Integer = 8
Const ND_KO As Integer = 9
Const KTU_S As Integer = 10
Const BIK_S As Integer = 11
Const NUMBF_S As Integer = 12
Const BRANCH As Integer = 13
Const KTU_SS As Integer = 14
Const BIK_SS As Integer = 15
Const NUMBF_SS As Integer = 16
Const TERROR As Integer = 17
Const VO As Integer = 18
Const DOP_V As Integer = 19
Const DATA As Integer = 20
Const SUME As Integer = 21
Const SUM As Integer = 22
Const CURREN As Integer = 23
Const PRIM_1 As Integer = 24
Const PRIM_2 As Integer = 25
Const NUM_PAY_D As Integer = 26
Const DATE_PAY_D As Integer = 27
Const METAL As Integer = 28
Const PRIZ6001 As Integer = 29
Const B_PAYER As Integer = 30
Const B_RECIP As Integer = 31
Const PART As Integer = 32
Const DESCR_1 As Integer = 33
Const DESCR_2 As Integer = 34
Const CURREN_CON As Integer = 35
Const SUM_CON As Integer = 36
Const PRIZ_SD As Integer = 37
Const TU0 As Integer = 38
Const PRU0 As Integer = 39
Const NAMEU0 As Integer = 40
Const KODCR0 As Integer = 41
Const KODCN0 As Integer = 42
Const AMR_S0 As Integer = 43
Const AMR_R0 As Integer = 44
Const AMR_G0 As Integer = 45
Const AMR_U0 As Integer = 46
Const AMR_D0 As Integer = 47
Const AMR_K0 As Integer = 48
Const AMR_O0 As Integer = 49
Const ADRESS_S0 As Integer = 50
Const ADRESS_R0 As Integer = 51
Const ADRESS_G0 As Integer = 52
Const ADRESS_U0 As Integer = 53
Const ADRESS_D0 As Integer = 54
Const ADRESS_K0 As Integer = 55
Const ADRESS_O0 As Integer = 56
Const KD0 As Integer = 57
Const SD0 As Integer = 58
Const RG0 As Integer = 59
Const ND0 As Integer = 60
Const VD01 As Integer = 61
Const VD02 As Integer = 62
Const VD03 As Integer = 63
Const VD04 As Integer = 64
Const VD05 As Integer = 65
Const VD06 As Integer = 66
Const VD07 As Integer = 67
Const MC_01 As Integer = 68
Const MC_02 As Integer = 69
Const MC_03 As Integer = 70
Const GR0 As Integer = 71
Const BP_0 As Integer = 72
Const VP_0 As Integer = 73
Const RESRV02 As Integer = 74
Const ACC_B0 As Integer = 75
Const ACC_COR_B0 As Integer = 76
Const NAME_IS_B0 As Integer = 77
Const BIK_IS_B0 As Integer = 78
Const CARD_B0 As Integer = 79
Const NAME_B0 As Integer = 80
Const KODCN_B0 As Integer = 81
Const BIK_B0 As Integer = 82
Const NAME_R0 As Integer = 83
Const KODCN_R0 As Integer = 84
Const BIK_R0 As Integer = 85
Const RESRV_B02 As Integer = 86
Const TU1 As Integer = 87
Const PRU1 As Integer = 88
Const NAMEU1 As Integer = 89
Const KODCR1 As Integer = 90
Const KODCN1 As Integer = 91
Const AMR_S1 As Integer = 92
Const AMR_R1 As Integer = 93
Const AMR_G1 As Integer = 94
Const AMR_U1 As Integer = 95
Const AMR_D1 As Integer = 96
Const AMR_K1 As Integer = 97
Const AMR_O1 As Integer = 98
Const ADRESS_S1 As Integer = 99
Const ADRESS_R1 As Integer = 100
Const ADRESS_G1 As Integer = 101
Const ADRESS_U1 As Integer = 102
Const ADRESS_D1 As Integer = 103
Const ADRESS_K1 As Integer = 104
Const ADRESS_O1 As Integer = 105
Const KD1 As Integer = 106
Const SD1 As Integer = 107
Const RG1 As Integer = 108
Const ND1 As Integer = 109
Const VD11 As Integer = 110
Const VD12 As Integer = 111
Const VD13 As Integer = 112
Const VD14 As Integer = 113
Const VD15 As Integer = 114
Const VD16 As Integer = 115
Const VD17 As Integer = 116
Const MC_11 As Integer = 117
Const MC_12 As Integer = 118
Const MC_13 As Integer = 119
Const GR1 As Integer = 120
Const BP_1 As Integer = 121
Const RESRV12 As Integer = 122
Const TU2 As Integer = 123
Const PRU2 As Integer = 124
Const NAMEU2 As Integer = 125
Const KODCR2 As Integer = 126
Const KODCN2 As Integer = 127
Const AMR_S2 As Integer = 128
Const AMR_R2 As Integer = 129
Const AMR_G2 As Integer = 130
Const AMR_U2 As Integer = 131
Const AMR_D2 As Integer = 132
Const AMR_K2 As Integer = 133
Const AMR_O2 As Integer = 134
Const ADRESS_S2 As Integer = 135
Const ADRESS_R2 As Integer = 136
Const ADRESS_G2 As Integer = 137
Const ADRESS_U2 As Integer = 138
Const ADRESS_D2 As Integer = 139
Const ADRESS_K2 As Integer = 140
Const ADRESS_O2 As Integer = 141
Const KD2 As Integer = 142
Const SD2 As Integer = 143
Const RG2 As Integer = 144
Const ND2 As Integer = 145
Const VD21 As Integer = 146
Const VD22 As Integer = 147
Const VD23 As Integer = 148
Const VD24 As Integer = 149
Const VD25 As Integer = 150
Const VD26 As Integer = 151
Const VD27 As Integer = 152
Const MC_21 As Integer = 153
Const MC_22 As Integer = 154
Const MC_23 As Integer = 155
Const GR2 As Integer = 156
Const BP_2 As Integer = 157
Const RESRV22 As Integer = 158
Const TU3 As Integer = 159
Const PRU3 As Integer = 160
Const NAMEU3 As Integer = 161
Const KODCR3 As Integer = 162
Const KODCN3 As Integer = 163
Const AMR_S3 As Integer = 164
Const AMR_R3 As Integer = 165
Const AMR_G3 As Integer = 166
Const AMR_U3 As Integer = 167
Const AMR_D3 As Integer = 168
Const AMR_K3 As Integer = 169
Const AMR_O3 As Integer = 170
Const ADRESS_S3 As Integer = 171
Const ADRESS_R3 As Integer = 172
Const ADRESS_G3 As Integer = 173
Const ADRESS_U3 As Integer = 174
Const ADRESS_D3 As Integer = 175
Const ADRESS_K3 As Integer = 176
Const ADRESS_O3 As Integer = 177
Const KD3 As Integer = 178
Const SD3 As Integer = 179
Const RG3 As Integer = 180
Const ND3 As Integer = 181
Const VD31 As Integer = 182
Const VD32 As Integer = 183
Const VD33 As Integer = 184
Const VD34 As Integer = 185
Const VD35 As Integer = 186
Const VD36 As Integer = 187
Const VD37 As Integer = 188
Const MC_31 As Integer = 189
Const MC_32 As Integer = 190
Const MC_33 As Integer = 191
Const GR3 As Integer = 192
Const BP_3 As Integer = 193
Const VP_3 As Integer = 194
Const RESRV32 As Integer = 195
Const ACC_B3 As Integer = 196
Const ACC_COR_B3 As Integer = 197
Const NAME_IS_B3 As Integer = 198
Const BIK_IS_B3 As Integer = 199
Const CARD_B3 As Integer = 200
Const NAME_B3 As Integer = 201
Const KODCN_B3 As Integer = 202
Const BIK_B3 As Integer = 203
Const NAME_R3 As Integer = 204
Const KODCN_R3 As Integer = 205
Const BIK_R3 As Integer = 206
Const RESRV_B32 As Integer = 207
Const TU4 As Integer = 208
Const PRU4 As Integer = 209
Const NAMEU4 As Integer = 210
Const KODCR4 As Integer = 211
Const KODCN4 As Integer = 212
Const AMR_S4 As Integer = 213
Const AMR_R4 As Integer = 214
Const AMR_G4 As Integer = 215
Const AMR_U4 As Integer = 216
Const AMR_D4 As Integer = 217
Const AMR_K4 As Integer = 218
Const AMR_O4 As Integer = 219
Const ADRESS_S4 As Integer = 220
Const ADRESS_R4 As Integer = 221
Const ADRESS_G4 As Integer = 222
Const ADRESS_U4 As Integer = 223
Const ADRESS_D4 As Integer = 224
Const ADRESS_K4 As Integer = 225
Const ADRESS_O4 As Integer = 226
Const KD4 As Integer = 227
Const SD4 As Integer = 228
Const RG4 As Integer = 229
Const ND4 As Integer = 230
Const VD41 As Integer = 231
Const VD42 As Integer = 232
Const VD43 As Integer = 233
Const VD44 As Integer = 234
Const VD45 As Integer = 235
Const VD46 As Integer = 236
Const VD47 As Integer = 237
Const MC_41 As Integer = 238
Const MC_42 As Integer = 239
Const MC_43 As Integer = 240
Const GR4 As Integer = 241
Const BP_4 As Integer = 242
Const RESRV42 As Integer = 243
Const RESERV612 As Integer = 244

Dim TH(1 To 244) As String
Dim TR(1 To 244) As String

Dim ActiveRow As Integer

Private Sub UserForm_Initialize()
    Dim c As Integer, i As Integer
    
    For i = 1 To 244
        TH(i) = ActiveSheet.Cells(THRow, i).Text
    Next i
    ActiveRow = ActiveCell.Row
    
    With UserForm1
        .Caption = "Проверка строки " & ActiveRow
        .Move 20, 40, Application.width - 40, Application.Height - 100
        ListBox1.Move 0, 0, .InsideWidth \ 5, .InsideHeight
        ListBox2.Move ListBox1.width, 0, .InsideWidth - ListBox1.width, .InsideHeight
    End With
    
    With ListBox1
        .ColumnCount = 2
        .ColumnWidths = "3 cm;"
        .AddItem "Поле:"
        For c = 1 To 37
            If NotResrvFld1(c) Then
                .AddItem Fld(c)
            End If
        Next c
        .List(0, 1) = "Информация об операции:"
    End With
    
    With ListBox2
        .ColumnCount = 6
        .ColumnWidths = "3 cm;;;;;"
        .AddItem "Поле:"
        For c = 38 To 86
            If NotResrvFld2(c) Then
                .AddItem Fld(c)
            End If
        Next c
        .List(0, 1) = "0. Сведения о лице:"
        .List(0, 2) = "1. Представитель лица:"
        .List(0, 3) = "2. Представитель получателя:"
        .List(0, 4) = "3. Получатель:"
        .List(0, 5) = "4. Третье лицо:"
    End With
    
    FillByRow
End Sub

Public Sub FillByRow()
    Dim c As Integer, i As Integer
    
    For i = 1 To 244
        TR(i) = ActiveSheet.Cells(ActiveRow, i).Text
    Next i
    
    With UserForm1
        .Caption = "Проверка строки " & ActiveRow
    End With
    
    With ListBox1
        i = 1
        For c = 1 To 37
            If NotResrvFld1(c) Then
                .List(i, 1) = Descr1(c)
                i = i + 1
            End If
        Next c
    
        .ListIndex = 0
    End With
    
    With ListBox2
        i = 1
        For c = TU0 To BP_0
            If NotResrvFld2(c) Then
                .List(i, 1) = Descr2(c) 'from 38 to 86
                .List(i, 2) = Descr2(c + 49) 'from 87 to 121
                .List(i, 3) = Descr2(c + 85) 'from 123 to 157
                .List(i, 4) = Descr2(c + 121) 'from 159 to 207
                .List(i, 5) = Descr2(c + 170) 'from 208 to 242
                i = i + 1
            End If
        Next c
        For c = VP_0 To RESRV_B02
            If NotResrvFld2(c) Then
                .List(i, 1) = Descr2(c) 'from 38 to 86
                .List(i, 4) = Descr2(c + 121) 'from 159 to 207
                i = i + 1
            End If
        Next c
        
        i = 1
        For c = 87 To 121
            If NotResrvFld2(c) Then
                .List(i, 2) = Descr2(c)
                i = i + 1
            End If
        Next c
        
        i = 1
        For c = 123 To 157
            If NotResrvFld2(c) Then
                .List(i, 3) = Descr2(c)
                i = i + 1
            End If
        Next c
        
        i = 1
        For c = 159 To 207
            If NotResrvFld2(c) Then
                .List(i, 4) = Descr2(c)
                i = i + 1
            End If
        Next c
        
        i = 1
        For c = 208 To 242
            If NotResrvFld2(c) Then
                .List(i, 5) = Descr2(c)
                i = i + 1
            End If
        Next c
        
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
                    r = s & " - нет"
                Case Else
                    If Not IsDate(s) Then r = Er(s)
            End Select
        Case DATA, DATE_P 'дата должна быть!
            Select Case s
                Case "01.01.2099"
                    r = Er(s)
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
                    r = s & " - без представителя"
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
                    r = s & " - нет"
                Case Else
                    If Not IsDate(s) Then r = Er(s)
            End Select
        Case GR0, GR1, GR2, GR3, GR4 'дата должна быть!
            Select Case s
                Case "01.01.2099"
                    r = Er(s)
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
                    r = s & " - не карта"
                Case Else
                    r = Er(s)
            End Select
    End Select
    Descr2 = r
End Function

Private Function Er(s As String) As String
    Er = s & " = ОШИБКА!"
End Function
