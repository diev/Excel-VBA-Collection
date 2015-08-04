VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FileEnter 
   Caption         =   "Выбор файла"
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7335
   OleObjectBlob   =   "FileEnter.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FileEnter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Option Base 1
DefLng A-Z

Public FileMask As String
Public MaskEdit As Boolean
Private Path As String

Private Sub cmdCancel_Click()
    Hide
End Sub

Private Sub cmdKill_Click()
    Dim s As String
    With lstFiles
        If .ListIndex = -1 Then Exit Sub
        cmdOk.Enabled = False
        cmdKill.Enabled = False
        s = Path & .Text
        If MsgBox(BPrintF("Действительно удалить?\n~%s~", s), vbExclamation + vbYesNo, "Удаление файла") = vbYes Then
            .RemoveItem .ListIndex
            lblFiles.Caption = BPrintF("Найдено:\n%d", .ListCount)
            Kill s
        End If
    End With
End Sub

Private Sub cmdOk_Click()
    With lstFiles
        FileMask = Path & .Text
    End With
    Hide
End Sub

Private Sub lstFiles_Change()
    Dim b As Boolean
    b = lstFiles.ListIndex > -1
    cmdOk.Enabled = b
    cmdKill.Enabled = b
End Sub

Public Sub FillList()
    Dim FileName As String, s As String, i, m
    cmdOk.Enabled = False
    cmdKill.Enabled = False
    With txtMask
        .Text = FileMask
        .Locked = Not MaskEdit
    End With
    Path = FilePath(FileMask)
    FileName = Dir(FileMask)
    With lstFiles
        .Clear
        i = 0: m = 0
        Do While FileName <> ""
            .AddItem FileName
            s = FileLen(Path & FileName) & Format(FileDateTime(Path & FileName), "  dd.mm.yy hh:mm")
            If Len(s) > m Then m = Len(s)
            .List(i, 1) = s
            i = i + 1
            FileName = Dir
        Loop
        lblFiles.Caption = BPrintF("Найдено:\n%d", .ListCount)
        For i = 0 To .ListCount - 1
            .List(i, 1) = PadL(.List(i, 1), m)
        Next
    End With
    FileMask = ""
End Sub

Private Sub lstFiles_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    cmdOk_Click
End Sub

Private Sub spnFiles_Change()
    With Application
        lstFiles.ColumnWidths = spnFiles * 1000 \ 35 & ";" & 6000 \ 35
    End With
End Sub

Private Sub txtMask_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If MaskEdit Then 'Not .Locked
        FileMask = txtMask.Text
        FillList
    End If
End Sub

