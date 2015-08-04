VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SaveAsPass 
   Caption         =   "Смена пароля, сохранение и выход"
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5505
   OleObjectBlob   =   "SaveAsPass.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SaveAsPass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Option Explicit
Option Compare Binary 'BINARY for passwords!!!
Option Base 1
DefLng A-Z

Private Sub chkChar_Click()
    EnableBoxes
End Sub

Private Sub cmdCancel_Click()
    ActiveWorkbook.Saved = False
    App.CloseAllowed = False
    Unload Me
End Sub

Private Sub cmdOk_Click()
    On Error Resume Next
    With ActiveWorkbook
        If optNoPass Then
            If .HasPassword Then
                Kill .FullName
                .SaveAs .FullName, Password:=vbNullString, CreateBackup:=True
            Else
                .Save
            End If
        ElseIf optPrevPass Then
            .Save
        ElseIf optNewPass Then
            If txtPass1.TextLength = 0 Or txtPass2.TextLength = 0 Then
                MsgBox "Пароль не введен!", vbExclamation, App.Title
                txtPass1.SetFocus
                Exit Sub
            End If
            If txtPass1 <> txtPass2 Then
                MsgBox "Пароли не совпадают!", vbExclamation, App.Title
                txtPass2.SetFocus
                Exit Sub
            End If
            Kill .FullName
            .SaveAs .FullName, Password:=txtPass1, CreateBackup:=True
        ElseIf optNoSave Then
            .Saved = True
        Else
            '?
        End If
        '.Saved = True
        Hide
        App.CloseAllowed = True
        '.Close
    End With
    Unload Me
    'AutoClose
End Sub

Private Sub optNewPass_Click()
    EnableBoxes
End Sub

Private Sub optNoPass_Click()
    EnableBoxes
End Sub

Private Sub optNoSave_Click()
    EnableBoxes
End Sub

Private Sub optPrevPass_Click()
    EnableBoxes
End Sub

Private Sub UserForm_Initialize()
    'chkChar = StrToBool(App.Setting("PGP", "ShowPass"))
    EnableBoxes
    If ActiveWorkbook.HasPassword Then
        optPrevPass.Enabled = True
        optPrevPass = True
    End If
End Sub

Private Sub EnableBoxes()
    With txtPass1
        If optNewPass Then
            .Enabled = True
            .BackStyle = fmBackStyleOpaque
        Else
            .Enabled = False
            .BackStyle = fmBackStyleTransparent
        End If
    End With
    With txtPass2
        If optNewPass Then
            .Enabled = True
            .BackStyle = fmBackStyleOpaque
        Else
            .Enabled = False
            .BackStyle = fmBackStyleTransparent
        End If
    End With
    chkChar.Enabled = optNewPass
    If chkChar Then
        txtPass1.PasswordChar = vbNullString
        txtPass2.PasswordChar = vbNullString
    Else
        txtPass1.PasswordChar = "*"
        txtPass2.PasswordChar = "*"
    End If
End Sub

