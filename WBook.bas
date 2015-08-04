Attribute VB_Name = "WBook"
Option Explicit
Option Compare Text
Option Base 1
DefLng A-Z

Public Sub WorkbookBeforeClose(Cancel As Boolean)
    On Error Resume Next
    If App.CloseAllowed Then
        Cancel = False
        AutoClose
    Else
        Cancel = True
        SavePassShow
        If App.CloseAllowed Then
            Cancel = False
            AutoClose
        End If
    End If
End Sub

Public Sub WorkbookOpen()
    AutoOpen
End Sub

Public Sub WorkbookSheetActivate(ByVal Sh As Object)
    Dim s As String
    On Error Resume Next
    If Sh.Name = User.ID Then 'Simple reactivate...
        'Exit Sub
    ElseIf Sh.Name = User.DemoID Then
        User.Demo = True
    ElseIf Val(Sh.Name) > 0 Then
        User.ID = Sh.Name
        'LogonShow
    End If
End Sub
