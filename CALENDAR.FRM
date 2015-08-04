VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Calendar 
   Caption         =   "Календарь"
   ClientHeight    =   4335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4095
   OleObjectBlob   =   "Calendar.frx":0000
   StartUpPosition =   1  'CenterOwner
   Tag             =   "23sdf"
End
Attribute VB_Name = "Calendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Option Base 1
DefLng A-Z

Public DateEntered As Variant

Private Sub cmdCancel_Click()
    DateEntered = Null
    Hide
End Sub

Private Sub cmdOk_Click()
    Hide
End Sub

Private Sub cmdToday_Click()
    Reset
End Sub

Private Sub lstMonth_Change()
    CalcDays
End Sub

Private Sub Label1_Click()
    SetDate Label1.Caption
End Sub

Private Sub Label2_Click()
    SetDate Label2.Caption
End Sub

Private Sub Label3_Click()
    SetDate Label3.Caption
End Sub

Private Sub Label4_Click()
    SetDate Label4.Caption
End Sub

Private Sub Label5_Click()
    SetDate Label5.Caption
End Sub

Private Sub Label6_Click()
    SetDate Label6.Caption
End Sub

Private Sub Label7_Click()
    SetDate Label7.Caption
End Sub

Private Sub Label8_Click()
    SetDate Label8.Caption
End Sub

Private Sub Label9_Click()
    SetDate Label9.Caption
End Sub

Private Sub Label10_Click()
    SetDate Label10.Caption
End Sub

Private Sub Label11_Click()
    SetDate Label11.Caption
End Sub

Private Sub Label12_Click()
    SetDate Label12.Caption
End Sub

Private Sub Label13_Click()
    SetDate Label13.Caption
End Sub

Private Sub Label14_Click()
    SetDate Label14.Caption
End Sub

Private Sub Label15_Click()
    SetDate Label15.Caption
End Sub

Private Sub Label16_Click()
    SetDate Label16.Caption
End Sub

Private Sub Label17_Click()
    SetDate Label17.Caption
End Sub

Private Sub Label18_Click()
    SetDate Label18.Caption
End Sub

Private Sub Label19_Click()
    SetDate Label19.Caption
End Sub

Private Sub Label20_Click()
    SetDate Label20.Caption
End Sub

Private Sub Label21_Click()
    SetDate Label21.Caption
End Sub

Private Sub Label22_Click()
    SetDate Label22.Caption
End Sub

Private Sub Label23_Click()
    SetDate Label23.Caption
End Sub

Private Sub Label24_Click()
    SetDate Label24.Caption
End Sub

Private Sub Label25_Click()
    SetDate Label25.Caption
End Sub

Private Sub Label26_Click()
    SetDate Label26.Caption
End Sub

Private Sub Label27_Click()
    SetDate Label27.Caption
End Sub

Private Sub Label28_Click()
    SetDate Label28.Caption
End Sub

Private Sub Label29_Click()
    SetDate Label29.Caption
End Sub

Private Sub Label30_Click()
    SetDate Label30.Caption
End Sub

Private Sub Label31_Click()
    SetDate Label31.Caption
End Sub

Private Sub Label32_Click()
    SetDate Label32.Caption
End Sub

Private Sub Label33_Click()
    SetDate Label33.Caption
End Sub

Private Sub Label34_Click()
    SetDate Label34.Caption
End Sub

Private Sub Label35_Click()
    SetDate Label35.Caption
End Sub

Private Sub Label36_Click()
    SetDate Label36.Caption
End Sub

Private Sub Label37_Click()
    SetDate Label37.Caption
End Sub

Private Sub Label38_Click()
    SetDate Label38.Caption
End Sub

Private Sub Label39_Click()
    SetDate Label39.Caption
End Sub

Private Sub Label40_Click()
    SetDate Label40.Caption
End Sub

Private Sub Label41_Click()
    SetDate Label41.Caption
End Sub

Private Sub Label42_Click()
    SetDate Label42.Caption
End Sub

Private Sub Label1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    cmdOk_Click
End Sub

Private Sub Label2_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    cmdOk_Click
End Sub

Private Sub Label3_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    cmdOk_Click
End Sub

Private Sub Label4_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    cmdOk_Click
End Sub

Private Sub Label5_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    cmdOk_Click
End Sub

Private Sub Label6_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    cmdOk_Click
End Sub

Private Sub Label7_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    cmdOk_Click
End Sub

Private Sub Label8_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    cmdOk_Click
End Sub

Private Sub Label9_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    cmdOk_Click
End Sub

Private Sub Label10_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    cmdOk_Click
End Sub

Private Sub Label11_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    cmdOk_Click
End Sub

Private Sub Label12_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    cmdOk_Click
End Sub

Private Sub Label13_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    cmdOk_Click
End Sub

Private Sub Label14_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    cmdOk_Click
End Sub

Private Sub Label15_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    cmdOk_Click
End Sub

Private Sub Label16_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    cmdOk_Click
End Sub

Private Sub Label17_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    cmdOk_Click
End Sub

Private Sub Label18_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    cmdOk_Click
End Sub

Private Sub Label19_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    cmdOk_Click
End Sub

Private Sub Label20_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    cmdOk_Click
End Sub

Private Sub Label21_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    cmdOk_Click
End Sub

Private Sub Label22_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    cmdOk_Click
End Sub

Private Sub Label23_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    cmdOk_Click
End Sub

Private Sub Label24_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    cmdOk_Click
End Sub

Private Sub Label25_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    cmdOk_Click
End Sub

Private Sub Label26_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    cmdOk_Click
End Sub

Private Sub Label27_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    cmdOk_Click
End Sub

Private Sub Label28_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    cmdOk_Click
End Sub

Private Sub Label29_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    cmdOk_Click
End Sub

Private Sub Label30_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    cmdOk_Click
End Sub

Private Sub Label31_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    cmdOk_Click
End Sub

Private Sub Label32_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    cmdOk_Click
End Sub

Private Sub Label33_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    cmdOk_Click
End Sub

Private Sub Label34_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    cmdOk_Click
End Sub

Private Sub Label35_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    cmdOk_Click
End Sub

Private Sub Label36_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    cmdOk_Click
End Sub

Private Sub Label37_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    cmdOk_Click
End Sub

Private Sub Label38_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    cmdOk_Click
End Sub

Private Sub Label39_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    cmdOk_Click
End Sub

Private Sub Label40_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    cmdOk_Click
End Sub

Private Sub Label41_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    cmdOk_Click
End Sub

Private Sub Label42_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    cmdOk_Click
End Sub

Private Sub lstMonth_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    cmdOk_Click
End Sub

Private Sub spnDay_Change()
    With spnDay
        DateEntered = DateAdd("d", .Value, DateEntered)
        .Value = 0
    End With
    DisplayDate
End Sub

Private Sub spnWeek_Change()
    With spnWeek
        DateEntered = DateAdd("d", .Value, DateEntered)
        .Value = 0
    End With
    DisplayDate
End Sub

Private Sub spnYear_Change()
    With txtYear
        .Value = .Value + spnYear.Value
    End With
End Sub

Private Sub txtYear_Change()
    spnYear.Value = 0
    CalcDays
End Sub

Private Sub UserForm_Initialize()
    Reset
End Sub

Private Sub CalcDays()
    Dim w, m, m1, d, d1, i, n, arr(42) As Variant
    
    On Error Resume Next
    m = lstMonth.ListIndex + 1
    m1 = CDate("01." & CStr(lstMonth.ListIndex + 1) & "." & txtYear.Text)
    w = Weekday(m1, vbUseSystemDayOfWeek)
    d1 = DateAdd("d", -w, m1)
    
    For i = 1 To 42
        d = DateAdd("d", i, d1)
        If Month(d) = m Then
            arr(i) = CStr(Day(d))
        Else
            arr(i) = ""
        End If
    Next
    
    Label1.Caption = arr(1)
    Label2.Caption = arr(2)
    Label3.Caption = arr(3)
    Label4.Caption = arr(4)
    Label5.Caption = arr(5)
    Label6.Caption = arr(6)
    Label7.Caption = arr(7)
    Label8.Caption = arr(8)
    Label9.Caption = arr(9)
    Label10.Caption = arr(10)
    Label11.Caption = arr(11)
    Label12.Caption = arr(12)
    Label13.Caption = arr(13)
    Label14.Caption = arr(14)
    Label15.Caption = arr(15)
    Label16.Caption = arr(16)
    Label17.Caption = arr(17)
    Label18.Caption = arr(18)
    Label19.Caption = arr(19)
    Label20.Caption = arr(20)
    Label21.Caption = arr(21)
    Label22.Caption = arr(22)
    Label23.Caption = arr(23)
    Label24.Caption = arr(24)
    Label25.Caption = arr(25)
    Label26.Caption = arr(26)
    Label27.Caption = arr(27)
    Label28.Caption = arr(28)
    Label29.Caption = arr(29)
    Label30.Caption = arr(30)
    Label31.Caption = arr(31)
    Label32.Caption = arr(32)
    Label33.Caption = arr(33)
    Label34.Caption = arr(34)
    Label35.Caption = arr(35)
    Label36.Caption = arr(36)
    Label37.Caption = arr(37)
    Label38.Caption = arr(38)
    Label39.Caption = arr(39)
    Label40.Caption = arr(40)
    Label41.Caption = arr(41)
    Label42.Caption = arr(42)
    
    SetDate CStr(Day(DateEntered))
End Sub

Private Sub SetDate(d As String)
    If Len(d) > 0 Then
        DateEntered = CDate(d & "." & CStr(lstMonth.ListIndex + 1) & "." & txtYear.Text)
        DisplayDate
    End If
End Sub

Private Sub DisplayDate()
    lblDate.Caption = Format(DateEntered, "dddd dd mmmm yyyy г.", vbMonday, vbFirstJan1)
    cmdToday.Enabled = DateEntered <> Date
End Sub

Private Sub Reset()
    DateEntered = Date
    With cmdToday
        .ControlTipText = Format(Date, "dddd dd mmmm yyyy", vbMonday, vbFirstJan1) & " г."
        .Enabled = False
    End With
    
    With lstMonth
        .Clear
        .AddItem "01 Январь"
        .AddItem "02 Февраль"
        .AddItem "03 Март"
        .AddItem "04 Апрель"
        .AddItem "05 Май"
        .AddItem "06 Июнь"
        .AddItem "07 Июль"
        .AddItem "08 Август"
        .AddItem "09 Сентябрь"
        .AddItem "10 Октябрь"
        .AddItem "11 Ноябрь"
        .AddItem "12 Декабрь"
        .ListIndex = Month(Date) - 1
    End With
    
    txtYear.Value = Year(Date)
    CalcDays
End Sub
