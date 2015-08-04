Attribute VB_Name = "Init"
Option Explicit
Option Compare Text
Option Base 1
DefLng A-Z

Public Sub AutoOpen()
    With Application
        .Cursor = xlWait
        .StatusBar = "Ждите, идет загрузка..."
        .Caption = AppTitle
    End With
    
    'Worksheets(Payment.Sheet).Activate
    Worksheets("Платежка").Activate
    
    With ActiveWindow
        '.Caption = GetSetting(AppName, "User", "Nick", NoUser) & " - " & ActiveSheet.Name
        .WindowState = xlMaximized
    End With
    
    With ActiveWorkbook
        SetMenu
    End With
    
    With Application
        .StatusBar = False
        .Cursor = xlDefault
    End With
    
    'If Len(GetSetting(AppName, "User", "Nick", "")) = 0 Then
    '    MsgBox "Пожалуйста, проверьте все настройки!", vbInformation, AppTitle
    '    UserOptionsShow
    'End If
End Sub

Public Sub AutoClose()
    With Application
        .Caption = Empty
        .StatusBar = False
    End With
    
    With ActiveWindow
        .Caption = Empty
    End With
    
    With MenuBars(xlWorksheet)
        .Reset
    End With
End Sub
