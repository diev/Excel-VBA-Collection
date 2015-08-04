Attribute VB_Name = "Workbook"
Option Explicit
Option Compare Text
Option Base 1
DefLng A-Z

Public Function WorkName() As String
    WorkName = ActiveWorkbook.Name
End Function

Public Function WorkFile() As String
    WorkFile = ActiveWorkbook.FullName
End Function

Public Function WorkPath() As String
    WorkPath = FilePath(ActiveWorkbook.FullName)
End Function

Public Sub WorkbookBeforeClose(Cancel As Boolean)
    On Error Resume Next
    Workbooks(WorkName).Activate
    Cancel = MsgBox("«акрыть программу?", vbQuestion + vbOKCancel, AppTitle) = vbCancel
    If Not Cancel Then
        With ActiveWorkbook
            If Not .CreateBackup Then
                MsgBox BPrintF("Ќасто€тельно рекомендуетс€ делать резерную копию\n" & _
                    "дл€ предотвращени€ потери данных"), vbExclamation, AppTitle
            End If
            'If Not .HasPassword Then
            '    MsgBox BPrintF("Ќасто€тельно рекомендуетс€ установить пароль\n" & _
            '        "дл€ предотвращени€ несанкционированного доступа!"), vbExclamation, AppTitle
            'End If
        End With
        AutoClose
    End If
End Sub

Public Sub WorkbookOpen()
    AutoOpen
End Sub

Public Sub WorkbookSheetActivate(ByVal Sh As Object)
    ActiveWindow.Caption = Sh.Name
End Sub

