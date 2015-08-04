Attribute VB_Name = "ChkData"
Option Explicit
Option Base 0 '1
DefLng A-Z

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

Private Row As Integer
Private Col As Integer

Public Sub CheckData()
    Dim db As Range, r As Range, t As Integer, s As String, rule As Integer
    
    Row = ActiveCell.Row
    Set db = ActiveSheet.Range("A4")
    
    If Row > 3 Then
        Select Case YesNoCancelBox("Начать проверку с начала?")
        Case vbYes
            Row = 1
        Case vbNo
            Row = ActiveCell.Row - 3
        Case Else
            Exit Sub
        End Select
    Else
        Row = 1
    End If
    
    Do
        If Len(db.Cells(Row, 1).Text) = 0 Then
            InfoBox "Проверка %d строк окончена", Row - 1
            Exit Do
        End If
        
        Set r = db.Rows(Row)
                
        rule = 1
        t = TERROR
        If r.Columns(t).Text <> "0" Then r.Columns(t) = "0"
        
        rule = 2
        t = DOP_V
        If r.Columns(t).Text <> "0" Then r.Columns(t) = "0"
        
        rule = 3
        If r.Columns(B_PAYER).Text = "1" And r.Columns(TU0).Text = "1" Then 'Отправитель - Клиент Банка
            t = TU1
            If r.Columns(t).Text = "0" Then
                r.Columns(t).Select
                If Not OkCancelBox("Правило %d\nНе должно быть 0", rule) Then Exit Sub
            End If
            t = NAMEU1
            If r.Columns(t).Text = "0" Then
                r.Columns(t).Select
                If Not OkCancelBox("Правило %d\nНе должно быть 0", rule) Then Exit Sub
            End If
            t = KODCR1
            If r.Columns(t).Text = "0" Then
                r.Columns(t).Select
                If Not OkCancelBox("Правило %d\nНе должно быть 0", rule) Then Exit Sub
            End If
            t = KODCN1
            If r.Columns(t).Text = "0" Then
                r.Columns(t).Select
                If Not OkCancelBox("Правило %d\nНе должно быть 0", rule) Then Exit Sub
            End If
            t = AMR_S1
            If r.Columns(t).Text = "0" Then
                r.Columns(t).Select
                If Not OkCancelBox("Правило %d\nНе должно быть 0", rule) Then Exit Sub
            End If
            t = AMR_G1
            If r.Columns(t).Text = "0" Then
                r.Columns(t).Select
                If Not OkCancelBox("Правило %d\nНе должно быть 0", rule) Then Exit Sub
            End If
            t = AMR_U1
            If r.Columns(t).Text = "0" Then
                r.Columns(t).Select
                If Not OkCancelBox("Правило %d\nНе должно быть 0", rule) Then Exit Sub
            End If
            t = AMR_D1
            If r.Columns(t).Text = "0" Then
                r.Columns(t).Select
                If Not OkCancelBox("Правило %d\nНе должно быть 0", rule) Then Exit Sub
            End If
            t = AMR_O1
            If r.Columns(t).Text = "0" Then
                r.Columns(t).Select
                If Not OkCancelBox("Правило %d\nНе должно быть 0", rule) Then Exit Sub
            End If
            t = ADRESS_S1
            If r.Columns(t).Text = "0" Then
                r.Columns(t).Select
                If Not OkCancelBox("Правило %d\nНе должно быть 0", rule) Then Exit Sub
            End If
            t = ADRESS_G1
            If r.Columns(t).Text = "0" Then
                r.Columns(t).Select
                If Not OkCancelBox("Правило %d\nНе должно быть 0", rule) Then Exit Sub
            End If
            t = ADRESS_U1
            If r.Columns(t).Text = "0" Then
                r.Columns(t).Select
                If Not OkCancelBox("Правило %d\nНе должно быть 0", rule) Then Exit Sub
            End If
            t = ADRESS_D1
            If r.Columns(t).Text = "0" Then
                r.Columns(t).Select
                If Not OkCancelBox("Правило %d\nНе должно быть 0", rule) Then Exit Sub
            End If
            t = ADRESS_O1
            If r.Columns(t).Text = "0" Then
                r.Columns(t).Select
                If Not OkCancelBox("Правило %d\nНе должно быть 0", rule) Then Exit Sub
            End If
            t = KD1
            If r.Columns(t).Text = "0" Then
                r.Columns(t).Select
                If Not OkCancelBox("Правило %d\nНе должно быть 0", rule) Then Exit Sub
            End If
            t = SD1
            If r.Columns(t).Text = "0" Then
                r.Columns(t).Select
                If Not OkCancelBox("Правило %d\nНе должно быть 0", rule) Then Exit Sub
            End If
            t = ND1
            If r.Columns(t).Text = "0" Then
                r.Columns(t).Select
                If Not OkCancelBox("Правило %d\nНе должно быть 0", rule) Then Exit Sub
            End If
            t = VD11
            If r.Columns(t).Text = "0" Then
                r.Columns(t).Select
                If Not OkCancelBox("Правило %d\nНе должно быть 0", rule) Then Exit Sub
            End If
            t = VD12
            If r.Columns(t).Text = "0" Then
                r.Columns(t).Select
                If Not OkCancelBox("Правило %d\nНе должно быть 0", rule) Then Exit Sub
            End If
            t = VD13
            If r.Columns(t).Text = "01.01.2099" Then
                r.Columns(t).Select
                If Not OkCancelBox("Правило %d\nНе должно быть 01.01.2099", rule) Then Exit Sub
            End If
            t = GR1
            If r.Columns(t).Text = "01.01.2099" Then
                r.Columns(t).Select
                If Not OkCancelBox("Правило %d\nНе должно быть 01.01.2099", rule) Then Exit Sub
            End If
            t = BP_1
            If r.Columns(t).Text = "0" Then
                r.Columns(t).Select
                If Not OkCancelBox("Правило %d\nНе должно быть 0", rule) Then Exit Sub
            End If
        End If
        
        rule = 4
        If r.Columns(B_PAYER).Text = "1" Then
            t = GR0
            If r.Columns(t).Text = "01.01.2099" Then
                r.Columns(t).Select
                If Not OkCancelBox("Правило %d\nНе должно быть 01.01.2099", rule) Then Exit Sub
            End If
        End If
        
        rule = 5
        If r.Columns(B_RECIP).Text = "1" Then 'Получатель - Клиент Банка
            t = GR3
            If r.Columns(t).Text = "01.01.2099" Then
                r.Columns(t).Select
                If Not OkCancelBox("Правило %d\nНе должно быть 01.01.2099", rule) Then Exit Sub
            End If
        End If
        
        rule = 6
        If r.Columns(B_RECIP).Text = "1" And r.Columns(TU3).Text = "1" Then 'Получатель - Клиент Банка
            t = TU2
            If r.Columns(t).Text = "0" Then
                r.Columns(t).Select
                If Not OkCancelBox("Правило %d\nНе должно быть 0", rule) Then Exit Sub
            End If
            t = NAMEU2
            If r.Columns(t).Text = "0" Then
                r.Columns(t).Select
                If Not OkCancelBox("Правило %d\nНе должно быть 0", rule) Then Exit Sub
            End If
            t = KODCR2
            If r.Columns(t).Text = "0" Then
                r.Columns(t).Select
                If Not OkCancelBox("Правило %d\nНе должно быть 0", rule) Then Exit Sub
            End If
            t = KODCN2
            If r.Columns(t).Text = "0" Then
                r.Columns(t).Select
                If Not OkCancelBox("Правило %d\nНе должно быть 0", rule) Then Exit Sub
            End If
            t = AMR_S2
            If r.Columns(t).Text = "0" Then
                r.Columns(t).Select
                If Not OkCancelBox("Правило %d\nНе должно быть 0", rule) Then Exit Sub
            End If
            t = AMR_G2
            If r.Columns(t).Text = "0" Then
                r.Columns(t).Select
                If Not OkCancelBox("Правило %d\nНе должно быть 0", rule) Then Exit Sub
            End If
            t = AMR_U2
            If r.Columns(t).Text = "0" Then
                r.Columns(t).Select
                If Not OkCancelBox("Правило %d\nНе должно быть 0", rule) Then Exit Sub
            End If
            t = AMR_D2
            If r.Columns(t).Text = "0" Then
                r.Columns(t).Select
                If Not OkCancelBox("Правило %d\nНе должно быть 0", rule) Then Exit Sub
            End If
            t = AMR_O2
            If r.Columns(t).Text = "0" Then
                r.Columns(t).Select
                If Not OkCancelBox("Правило %d\nНе должно быть 0", rule) Then Exit Sub
            End If
            t = ADRESS_S2
            If r.Columns(t).Text = "0" Then
                r.Columns(t).Select
                If Not OkCancelBox("Правило %d\nНе должно быть 0", rule) Then Exit Sub
            End If
            t = ADRESS_G2
            If r.Columns(t).Text = "0" Then
                r.Columns(t).Select
                If Not OkCancelBox("Правило %d\nНе должно быть 0", rule) Then Exit Sub
            End If
            t = ADRESS_U2
            If r.Columns(t).Text = "0" Then
                r.Columns(t).Select
                If Not OkCancelBox("Правило %d\nНе должно быть 0", rule) Then Exit Sub
            End If
            t = ADRESS_D2
            If r.Columns(t).Text = "0" Then
                r.Columns(t).Select
                If Not OkCancelBox("Правило %d\nНе должно быть 0", rule) Then Exit Sub
            End If
            t = ADRESS_O2
            If r.Columns(t).Text = "0" Then
                r.Columns(t).Select
                If Not OkCancelBox("Правило %d\nНе должно быть 0", rule) Then Exit Sub
            End If
            t = KD2
            If r.Columns(t).Text = "0" Then
                r.Columns(t).Select
                If Not OkCancelBox("Правило %d\nНе должно быть 0", rule) Then Exit Sub
            End If
            t = SD2
            If r.Columns(t).Text = "0" Then
                r.Columns(t).Select
                If Not OkCancelBox("Правило %d\nНе должно быть 0", rule) Then Exit Sub
            End If
            t = ND2
            If r.Columns(t).Text = "0" Then
                r.Columns(t).Select
                If Not OkCancelBox("Правило %d\nНе должно быть 0", rule) Then Exit Sub
            End If
            t = VD21
            If r.Columns(t).Text = "0" Then
                r.Columns(t).Select
                If Not OkCancelBox("Правило %d\nНе должно быть 0", rule) Then Exit Sub
            End If
            t = VD22
            If r.Columns(t).Text = "0" Then
                r.Columns(t).Select
                If Not OkCancelBox("Правило %d\nНе должно быть 0", rule) Then Exit Sub
            End If
            t = VD23
            If r.Columns(t).Text = "01.01.2099" Then
                r.Columns(t).Select
                If Not OkCancelBox("Правило %d\nНе должно быть 01.01.2099", rule) Then Exit Sub
            End If
            t = GR2
            If r.Columns(t).Text = "01.01.2099" Then
                r.Columns(t).Select
                If Not OkCancelBox("Правило %d\nНе должно быть 01.01.2099", rule) Then Exit Sub
            End If
            t = BP_2
            If r.Columns(t).Text = "0" Then
                r.Columns(t).Select
                If Not OkCancelBox("Правило %d\nНе должно быть 0", rule) Then Exit Sub
            End If
        End If
        
        rule = 7
        If r.Columns(B_PAYER).Text = "0" Then 'Отправитель - не Клиент Банка
            t = KODCR0
            If r.Columns(t).Text = "0" Then
                r.Columns(t).Select
                If Not OkCancelBox("Правило %d\nНе должно быть 0", rule) Then Exit Sub
            End If
            t = KODCN0
            If r.Columns(t).Text = "0" Then
                r.Columns(t).Select
                If Not OkCancelBox("Правило %d\nНе должно быть 0", rule) Then Exit Sub
            End If
            t = ND0
            If r.Columns(t).Text = "0" Then
                r.Columns(t).Select
                If Not OkCancelBox("Правило %d\nНе должно быть 0", rule) Then Exit Sub
            End If
            t = ACC_COR_B0
            If r.Columns(t).Text = "0" Then
                r.Columns(t).Select
                If Not OkCancelBox("Правило %d\nНе должно быть 0", rule) Then Exit Sub
            End If
            t = NAME_B0
            If r.Columns(t).Text = "0" Then
                r.Columns(t).Select
                If Not OkCancelBox("Правило %d\nНе должно быть 0", rule) Then Exit Sub
            End If
            t = KODCN_B0
            If r.Columns(t).Text = "0" Then
                r.Columns(t).Select
                If Not OkCancelBox("Правило %d\nНе должно быть 0", rule) Then Exit Sub
            End If
            t = BIK_B0
            If r.Columns(t).Text = "0" Then
                r.Columns(t).Select
                If Not OkCancelBox("Правило %d\nНе должно быть 0", rule) Then Exit Sub
            End If
            t = NAME_R0
            If r.Columns(t).Text = "0" Then
                r.Columns(t).Select
                If Not OkCancelBox("Правило %d\nНе должно быть 0", rule) Then Exit Sub
            End If
            t = KODCN_R0
            If r.Columns(t).Text = "0" Then
                r.Columns(t).Select
                If Not OkCancelBox("Правило %d\nНе должно быть 0", rule) Then Exit Sub
            End If
            t = BIK_R0
            If r.Columns(t).Text = "0" Then
                r.Columns(t).Select
                If Not OkCancelBox("Правило %d\nНе должно быть 0", rule) Then Exit Sub
            End If
        End If
        
        rule = 8
        If r.Columns(B_RECIP).Text = "0" Then 'Получатель - не Клиент Банка
            t = KODCR3
            If r.Columns(t).Text = "0" Then
                r.Columns(t).Select
                If Not OkCancelBox("Правило %d\nНе должно быть 0", rule) Then Exit Sub
            End If
            t = KODCN3
            If r.Columns(t).Text = "0" Then
                r.Columns(t).Select
                If Not OkCancelBox("Правило %d\nНе должно быть 0", rule) Then Exit Sub
            End If
            t = ND3
            If r.Columns(t).Text = "0" Then
                r.Columns(t).Select
                If Not OkCancelBox("Правило %d\nНе должно быть 0", rule) Then Exit Sub
            End If
            t = ACC_COR_B3
            If r.Columns(t).Text = "0" Then
                r.Columns(t).Select
                If Not OkCancelBox("Правило %d\nНе должно быть 0", rule) Then Exit Sub
            End If
            t = NAME_B3
            If r.Columns(t).Text = "0" Then
                r.Columns(t).Select
                If Not OkCancelBox("Правило %d\nНе должно быть 0", rule) Then Exit Sub
            End If
            t = KODCN_B3
            If r.Columns(t).Text = "0" Then
                r.Columns(t).Select
                If Not OkCancelBox("Правило %d\nНе должно быть 0", rule) Then Exit Sub
            End If
            t = BIK_B3
            If r.Columns(t).Text = "0" Then
                r.Columns(t).Select
                If Not OkCancelBox("Правило %d\nНе должно быть 0", rule) Then Exit Sub
            End If
            t = NAME_R3
            If r.Columns(t).Text = "0" Then
                r.Columns(t).Select
                If Not OkCancelBox("Правило %d\nНе должно быть 0", rule) Then Exit Sub
            End If
            t = KODCN_R3
            If r.Columns(t).Text = "0" Then
                r.Columns(t).Select
                If Not OkCancelBox("Правило %d\nНе должно быть 0", rule) Then Exit Sub
            End If
            t = BIK_R3
            If r.Columns(t).Text = "0" Then
                r.Columns(t).Select
                If Not OkCancelBox("Правило %d\nНе должно быть 0", rule) Then Exit Sub
            End If
        End If
        
        rule = 9
        If r.Columns(B_PAYER).Text = "1" And r.Columns(B_RECIP).Text = "1" Then
            t = ACC_COR_B0
            If r.Columns(t).Text <> "0" Then
                r.Columns(t).Select
                If Not OkCancelBox("Правило %d\nДолжно быть 0", rule) Then Exit Sub
            End If
            t = ACC_COR_B3
            If r.Columns(t).Text <> "0" Then
                r.Columns(t).Select
                If Not OkCancelBox("Правило %d\nДолжно быть 0", rule) Then Exit Sub
            End If
        End If
            
'        rule = 10
'        If r.Columns(B_PAYER).Text = "1" Then 'Отправитель - Клиент Банка
'            t = ACC_COR_B0
'            If r.Columns(t).Text <> "0" Then
'                r.Columns(t).Select
'                If Not OkCancelBox("Правило %d\nДолжно быть 0", rule) Then Exit Sub
'            End If
'        End If
            
        rule = 10
        If r.Columns(VO).Text = "5003" Or r.Columns(VO).Text = "8001" Then
            t = PRIZ_SD
            If r.Columns(t).Text <> "1" Then r.Columns(t) = "1"
        End If
            
        rule = 11
        If r.Columns(VO).Text = "5003" Or r.Columns(VO).Text = "6001" Or r.Columns(VO).Text = "8001" Then
            t = DATE_S
            'If r.Columns(t).Text = "01.01.2099" Then r.Columns(t) = r.Columns(DATE_P)
            r.Columns(t) = r.Columns(DATE_P)
        End If
            
        rule = 12
        t = NUMB_P
        If Row > 1 Then
            r.Columns(t) = db.Cells(1, NUMB_P) + Row - 1
        End If
                
'                t = AMR_G0
'                If Col = t And r.Columns(t).Text = "СПб" Then
'                    r.Columns(t).Select
'                    s = "г.Санкт-Петербург"
'                    If YesNoBox("Меняем СПб на %s?", s) Then r.Columns(t) = s
'                End If
'
'                t = AMR_K0
'                If Col = t And r.Columns(t).Text = "А" Then
'                    r.Columns(t).Select
'                    s = "лит.А"
'                    If YesNoBox("Меняем А на %s?", s) Then r.Columns(t) = s
'                End If
        
        rule = 13
        t = BIK_B0
        If r.Columns(ACC_COR_B0).Text = "0" And r.Columns(BIK_B0).Text <> "0" Then
            If Not OkCancelBox("Правило %d\nДолжно быть 0?", rule) Then Exit Sub
        ElseIf r.Columns(ACC_COR_B0).Text = App.DefKS And r.Columns(BIK_B0).Text <> "0" Then
            If Not OkCancelBox("Правило %d\nДолжно быть 0", rule) Then Exit Sub
        ElseIf Right(r.Columns(ACC_COR_B0).Text, 3) <> Right(r.Columns(BIK_B0).Text, 3) Then
            If Not OkCancelBox("Правило %d\nБИК и к/с должны совпадать", rule) Then Exit Sub
        End If
        
        rule = 14
        t = BIK_B3
        If r.Columns(ACC_COR_B3).Text = "0" And r.Columns(BIK_B3).Text <> "0" Then
            If Not OkCancelBox("Правило %d\nДолжно быть 0?", rule) Then Exit Sub
        ElseIf r.Columns(ACC_COR_B3).Text = App.DefKS And r.Columns(BIK_B3).Text <> "0" Then
            If Not OkCancelBox("Правило %d\nДолжно быть 0", rule) Then Exit Sub
        ElseIf Right(r.Columns(ACC_COR_B3).Text, 3) <> Right(r.Columns(BIK_B3).Text, 3) Then
            If Not OkCancelBox("Правило %d\nБИК и к/с должны совпадать", rule) Then Exit Sub
        End If
        
        Set r = Nothing
        DoEvents
        Row = Row + 1
    Loop
    
    Set db = Nothing
End Sub

Private Function Field(Name As String) As Boolean
    Dim s As String, n As Integer
    s = Cells(3, Col).Text
    n = InStr(s, " ")
    s = Left(s, n - 1)
    Field = Name = s
End Function

Private Sub Fields()
    Dim s As String, n As Integer
    Row = 3
    Col = 1
    Do While (Len(Cells(Row, Col).Text) > 0)
        s = Cells(Row, Col).Text
        n = InStr(s, " ")
        s = Left(s, n - 1)
        Debug.Print "Const " & s & " As Integer = " & Col
        Col = Col + 1
        If Col > 100 Then Exit Do
    Loop
End Sub
