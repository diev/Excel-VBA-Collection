Attribute VB_Name = "WBook"
Option Explicit
Option Compare Text
Option Base 1
DefLng A-Z

Public Sub WorkbookActivate()
    On Error Resume Next
    InitMenuBars
    User.ID = ActiveSheet.Name
End Sub

Public Sub WorkbookBeforeClose(Cancel As Boolean)
    On Error Resume Next
    AutoClose
    Cancel = False
End Sub

Public Sub WorkbookDeactivate()
    On Error Resume Next
    Application.Caption = Application & " (����� ������ ����-������!)"
    CloseMenuBars
End Sub

Public Sub WorkbookOpen()
    On Error Resume Next
    AutoOpen
End Sub

Public Sub WorkbookSheetActivate(ByVal sh As Object)
    On Error Resume Next
    User.ID = sh.Name
End Sub

Public Sub NewUserSheet(Optional Name As String = "000")
    'Dim ws As Worksheet
    On Error Resume Next
    'For Each ws In Sheets
    '    If ws.Name = Name Then
    '        WarnBox "���� � ������ %s ��� ����������!", Name
    '        ws.Activate
    '        Exit Sub
    '    End If
    'Next
    Sheets.Add
    ActiveSheet.Name = Name
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "����"
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "�������"
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "�����"
    Range("D1").Select
    ActiveCell.FormulaR1C1 = "����"
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "�����"
    Range("F1").Select
    ActiveCell.FormulaR1C1 = "����������"
    Range("G1").Select
    ActiveCell.FormulaR1C1 = "���"
    Range("H1").Select
    ActiveCell.FormulaR1C1 = "���"
    Range("I1").Select
    ActiveCell.FormulaR1C1 = "����"
    Range("J1").Select
    ActiveCell.FormulaR1C1 = "����"
    Range("K1").Select
    ActiveCell.FormulaR1C1 = "����������"
    Range("A1:K1").Select
    With Selection.Interior
        .ColorIndex = 16
        .Pattern = xlSolid
    End With
    Selection.Font.ColorIndex = 2
    Columns("A:K").Select
    Selection.NumberFormat = "@"
    Range("B:D,J:J").Select
    Selection.HorizontalAlignment = xlCenter
    Columns("D:D").Select
    Selection.NumberFormat = "dd.mm.yyyy"
    Columns("E:E").Select
    Selection.NumberFormat = "#,##0.00"
    Range("E1").Select
    Selection.HorizontalAlignment = xlCenter
    Columns("A:K").EntireColumn.AutoFit
    Columns("D:D").ColumnWidth = 10
    Columns("E:E").ColumnWidth = 10
    Columns("F:F").ColumnWidth = 20
    Columns("G:G").ColumnWidth = 10
    Columns("H:H").ColumnWidth = 10
    Columns("I:I").ColumnWidth = 20
    Columns("K:K").ColumnWidth = 40
    Range("C2").Select
End Sub

