Attribute VB_Name = "MenuBar"
Option Explicit
Option Compare Text
Option Base 1
DefLng A-Z

Const BarName = "161п"
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
                .Caption = "&" & CStr(i) & ". Загрузить"
                .TooltipText = "Загрузить DBF"
                .OnAction = "DumpDBFFile"
                .Style = msoButtonCaption
            End With
            With .Add(Type:=msoControlButton, ID:=1, temporary:=True)
                i = i + 1
                .Caption = "&" & CStr(i) & ". Добавить"
                .TooltipText = "Добавить ещё один DBF"
                .OnAction = "DumpDBFFile2"
                .Style = msoButtonCaption
                .BeginGroup = True
            End With
            With .Add(Type:=msoControlButton, ID:=1, temporary:=True)
                i = i + 1
                .Caption = "&" & CStr(i) & ". Просмотр"
                .TooltipText = "Просмотреть данные"
                .OnAction = "ViewForm"
                .Style = msoButtonCaption
                .BeginGroup = True
            End With
            With .Add(Type:=msoControlButton, ID:=1, temporary:=True)
                i = i + 1
                .Caption = "&" & CStr(i) & ". Проверить"
                .TooltipText = "Проверить данные"
                .OnAction = "CheckData"
                .Style = msoButtonCaption
                .BeginGroup = True
            End With
            With .Add(Type:=msoControlButton, ID:=1, temporary:=True)
                i = i + 1
                .Caption = "&" & CStr(i) & ". Сохранить"
                .TooltipText = "Сохранить DBF"
                .OnAction = "WriteDBFFile"
                .Style = msoButtonCaption
                .BeginGroup = True
            End With
            With .Add(Type:=msoControlButton, ID:=1, temporary:=True)
                i = i + 1
                .Caption = "&" & CStr(i) & ". Передать в Комиту"
                .TooltipText = "Передать на проверку в Комиту"
                .OnAction = "ExportComita"
                .Style = msoButtonCaption
                .BeginGroup = True
            End With
            With .Add(Type:=msoControlButton, ID:=1, temporary:=True)
                i = i + 1
                .Caption = "&" & CStr(i) & ". Отправить в ЦБ"
                .TooltipText = "Передать на отправку в ЦБ"
                .OnAction = "ExportSVK"
                .Style = msoButtonCaption
                .BeginGroup = True
            End With
            With .Add(Type:=msoControlButton, ID:=1, temporary:=True)
                i = i + 1
                .Caption = "&" & CStr(i) & ". Печать"
                .TooltipText = "Распечатать сокр.вариант"
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
