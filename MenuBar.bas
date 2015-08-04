Attribute VB_Name = "MenuBar"
Option Explicit
Option Compare Text
Option Base 1
DefLng A-Z

Const BarName = "161�"
Const BarPosition = 1

Public Sub AddBar()
    Dim i As Long
    On Error Resume Next
    CommandBars(BarName).Delete
    With CommandBars.Add(Name:=BarName, Position:=BarPosition, temporary:=True)
        .Visible = True
        i = 0
        With .Controls
            With .Add(Type:=msoControlButton, ID:=1, temporary:=True)
                i = i + 1
                .Caption = "&" & CStr(i) & ". ���������"
                .TooltipText = "��������� DBF"
                .OnAction = "DumpDBFFile"
                .Style = msoButtonCaption
            End With
            With .Add(Type:=msoControlButton, ID:=1, temporary:=True)
                i = i + 1
                .Caption = "&" & CStr(i) & ". ��������"
                .TooltipText = "�������� ��� ���� DBF"
                .OnAction = "DumpDBFFile2"
                .Style = msoButtonCaption
                .BeginGroup = True
            End With
            With .Add(Type:=msoControlButton, ID:=1, temporary:=True)
                i = i + 1
                .Caption = "&" & CStr(i) & ". ���������"
                .TooltipText = "��������� ������"
                .OnAction = "CheckData"
                .Style = msoButtonCaption
                .BeginGroup = True
            End With
            With .Add(Type:=msoControlButton, ID:=1, temporary:=True)
                i = i + 1
                .Caption = "&" & CStr(i) & ". ���������"
                .TooltipText = "��������� DBF"
                .OnAction = "WriteDBFFile"
                .Style = msoButtonCaption
                .BeginGroup = True
            End With
            With .Add(Type:=msoControlButton, ID:=1, temporary:=True)
                i = i + 1
                .Caption = "&" & CStr(i) & ". �������� � ������"
                .TooltipText = "�������� �� �������� � ������"
                .OnAction = "ExportComita"
                .Style = msoButtonCaption
                .BeginGroup = True
            End With
            With .Add(Type:=msoControlButton, ID:=1, temporary:=True)
                i = i + 1
                .Caption = "&" & CStr(i) & ". ��������� � ��"
                .TooltipText = "�������� �� �������� � ��"
                .OnAction = "ExportSVK"
                .Style = msoButtonCaption
                .BeginGroup = True
            End With
            With .Add(Type:=msoControlButton, ID:=1, temporary:=True)
                i = i + 1
                .Caption = "&" & CStr(i) & ". ������"
                .TooltipText = "����������� ����.�������"
                .OnAction = "PrintDBFDigest"
                .Style = msoButtonCaption
                .BeginGroup = True
            End With
        End With
    End With
End Sub

Public Sub CloseBar()
    On Error Resume Next
    CommandBars(BarName).Delete
End Sub

