VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSMail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Option Base 1
DefLng A-Z

'Const SMailSection = "SMail"

Private Sub Class_Initialize()
    ForceDirectories Me.Send
    ForceDirectories Me.Recv
    ForceDirectories Me.Archive
End Sub

Public Function Valid()
    Dim File As String
    On Error GoTo NoSMail
    If IsFile(App.Path & "SMail\SMail.exe") Then
        Valid = True
        Exit Function
    Else
        If YesNoBox("У Вас нет программы связи с Банком!\nВозможно, Вы не установили файл SML*.EXE с дистрибутивной дискеты.\n\nПоискать сейчас на диске A:?") Then
            File = Dir("a:\sml*.exe")
            If File <> vbNullString Then
                ShellDialog File, vbNormalFocus, App.BoolOptions("Debug")
            End If
        End If
    End If
    Valid = IsFile(App.Path & "SMail\SMail.exe")
    Exit Function
NoSMail:
    StopBox "Нет программы связи!\nИщите Вашу дистрибутивную дискету."
    Valid = False
End Function

Public Sub Dial()
    Dim s As String, n As Long
    On Error Resume Next
    If Valid Then
        With Me
            n = CountFiles(.Send & "*.*")
            s = " @" & QFile(App.Path & "SMail\SMail.ctl") & " Poll:Bank SP"
            If DialOrInet = 2 Then
                s = s & " Channel1:NET IP_Timeout:15"
            End If
            Shell QFile(App.Path & "SMail\SMail.exe") & s, vbNormalFocus
            If n > 0 Then
                DoEvents
                OutputFile QFile(App.Path & "SMail\Flags\Bank.pll"), "Dial again!"
            End If
        End With
    End If
End Sub

Public Property Get DialOrInet() As Long
    DialOrInet = Val(App.Options("DialOrInet"))
    If DialOrInet = 0 Then
        If YesNoBox("Включить связь с Банком через Интернет?") Then
            DialOrInet = 2
        Else
            DialOrInet = 1
        End If
        Me.DialOrInet = DialOrInet
    End If
End Property

Public Property Let DialOrInet(ByVal vNewValue As Long)
    If vNewValue <> 2 Then vNewValue = 1
    App.Options("DialOrInet") = vNewValue
End Property

Public Property Get Send() As String
    Send = App.Path & "SMail\Mailbox\Bank\S\" 'App.Options("Send")
End Property

Public Property Get Recv() As String
    Recv = App.Path & "SMail\Mailbox\Bank\R\" 'App.Options("Recv")
End Property

Public Property Get Archive() As String
    Archive = App.Path & "Archive\"
End Property

Public Sub LastRep()
    Dim File As String, s As String
    With Me
        File = Dir(.Recv & "Rep???*.txt")
        s = vbNullString
        Do While Len(File) > 0
            s = s & Bsprintf("%s\n", CWin(InputFile(.Recv & File)))
            File = Dir
        Loop
    End With
    If Len(s) > 0 Then
        InfoBox s
    End If
End Sub

Public Function FlagDate(Ext As String) As String
    Dim s As String
    s = App.Path & "SMail\Flags\Bank." & Ext
    If IsFile(s) Then
        FlagDate = Bsprintf("%T", FileDateTime(s))
    Else
        FlagDate = "нет"
    End If
End Function

