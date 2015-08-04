VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PlatNAL 
   Caption         =   "������������ ������� � ������"
   ClientHeight    =   3600
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   6000
   OleObjectBlob   =   "PlatNAL.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "PlatNAL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Option Base 1
DefLng A-Z

Const NULNAL = "0"

Private Sub cboSS_Change()
    If Val(Left(cboSS, 2)) = 0 Then
        cboNAL1 = vbNullString
        cboNAL2 = vbNullString
        cboNAL3 = vbNullString
        cboNAL4 = vbNullString
        cboNAL5 = vbNullString
        cboNAL6 = vbNullString
        cboNAL7 = vbNullString
    Else
        If cboNAL1.TextLength = 0 Then cboNAL1 = NULNAL
        If cboNAL2.TextLength = 0 Then cboNAL2 = NULNAL
        If cboNAL3.TextLength = 0 Then cboNAL3 = cboNAL3.List(0)
        If cboNAL4.TextLength = 0 Then cboNAL4 = cboNAL4.List(3)
        If cboNAL5.TextLength = 0 Then cboNAL5 = NULNAL
        If cboNAL6.TextLength = 0 Then cboNAL6 = DtoC(Now)
        If cboNAL7.TextLength = 0 Then cboNAL7 = cboNAL7.List(0)
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    With Payment
        .SS = Trim(Left(cboSS, 2))
        .NAL(1) = cboNAL1
        .NAL(2) = cboNAL2
        .NAL(3) = Left(cboNAL3, 2)
        .NAL(4) = cboNAL4
        .NAL(5) = cboNAL5
        .NAL(6) = cboNAL6
        .NAL(7) = Left(cboNAL7, 2)
    End With
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    Dim s As String, i As Long
    On Error Resume Next
    With cboSS
        .AddItem "01/����������������"
        .AddItem "02/��������� �����"
        .AddItem "03/������� �������"
        .AddItem "04/��������� �����"
        .AddItem "05/������ ���������"
        .AddItem "06/�������� ������������������"
        .AddItem "07/���������� �����"
        .AddItem "08/���������� ���� ����.������"
        .AddItem "09/���.���������������"
        .AddItem "10/������� ��������"
        .AddItem "11/�������, ���.�������"
        .AddItem "12/����� ������.���-��"
        .AddItem "13/���� ���.���� �� �����"
        .AddItem "14/������� ���.�����"
        .AddItem "15/���.���� ��� �����"
        .AddItem "  /��� - ������� ������"
    End With
    With cboNAL1
        s = Payment.NAL(1)
        .AddItem s
        .AddItem "18210000000001000000"
        .AddItem "18210000000002000000"
        .AddItem "18210000000003000000"
        If s <> NULNAL Then
            .AddItem NULNAL
        End If
        .Text = .List(0)
    End With
    With cboNAL2
        s = User.OKATO
        .AddItem s
        If s <> NULNAL Then
            .AddItem NULNAL
        End If
    End With
    With cboNAL3
        .AddItem "��/������� ������"
        .AddItem "��/�������������"
        .AddItem "��/���.����"
        .AddItem "��/����������"
        .AddItem "��/���������"
        .AddItem "��/��������"
        .AddItem "��/����������������"
        .AddItem "��/������� ����������"
        .AddItem "��/���������������"
        .AddItem "��/��� ��������"
        .AddItem "��/�������������� ���."
        .AddItem NULNAL
        .Text = .List(0)
    End With
    s = Right(DtoC(Now), 8)
    With cboNAL4
        .AddItem "�1" & s
        .AddItem "�2" & s
        .AddItem "�3" & s
        .AddItem "��" & Right(DtoC(DateAdd("m", -1, Now)), 8)
        .AddItem "��" & s
        s = Right(s, 5)
        .AddItem "��.01" & s
        .AddItem "��.02" & s
        .AddItem "��.03" & s
        .AddItem "��.04" & s
        .AddItem "��.01" & s
        .AddItem "��.02" & s
        .AddItem "��.00" & s
        .AddItem DtoC(Now)
        .AddItem NULNAL
    End With
    With cboNAL5
        .AddItem "��"
        .AddItem "��"
        .AddItem "��"
        .AddItem "��"
        .AddItem "��"
        .AddItem "��"
        .AddItem "��"
        .AddItem "��"
        .AddItem NULNAL
    End With
    With cboNAL6
        .AddItem NULNAL
    End With
    With cboNAL7
        .AddItem "��/�����,����"
        .AddItem "��/�����"
        .AddItem "��/�������"
        .AddItem "��/�����"
        .AddItem "��/�����,����������"
        .AddItem "��/����"
        .AddItem "��/��������"
        .AddItem "��/�������"
        .AddItem "��/���.�����"
        .AddItem "��/���� �����"
        .AddItem NULNAL
        .Text = .List(0)
    End With
    With Payment
        s = .SS
        For i = 0 To cboSS.ListCount - 1
            If Left(cboSS.List(i), 2) = s Then
                cboSS.Text = cboSS.List(i)
                Exit For
            End If
        Next
        'cboNAL1 = .NAL(1)
        s = .NAL(1)
        For i = 0 To cboNAL1.ListCount - 1
            If cboNAL1.List(i) = s Then
                cboNAL1.Text = cboNAL1.List(i)
                Exit For
            End If
        Next
        cboNAL2 = .NAL(2)
        s = .NAL(3)
        For i = 0 To cboNAL3.ListCount - 1
            If Left(cboNAL3.List(i), 2) = s Then
                cboNAL3.Text = cboNAL3.List(i)
                Exit For
            End If
        Next
        cboNAL4 = .NAL(4)
        cboNAL5 = .NAL(5)
        cboNAL6 = .NAL(6)
        s = .NAL(7)
        For i = 0 To cboNAL7.ListCount - 1
            If Left(cboNAL7.List(i), 2) = s Then
                cboNAL7.Text = cboNAL7.List(i)
                Exit For
            End If
        Next
    End With
End Sub

