Attribute VB_Name = "FunShell"
Option Explicit
Option Compare Text
Option Base 1
DefLng A-Z

Private Const INFINITE = -1&

Private Declare Function WaitForSingleObject Lib "kernel32" _
    (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" _
    (ByVal dwAccess As Long, ByVal fInherit As Integer, ByVal hObject As Long) As Long
    
Public Sub ShellWait(RunCmd As String, Optional RunWindow As Long = vbMinimizedNoFocus)
    Dim retVal As Long, pID As Long, pHandle As Long
    On Error GoTo ErrShell
    pID = Shell(RunCmd, RunWindow)
    pHandle = OpenProcess(&H100000, True, pID)
    retVal = WaitForSingleObject(pHandle, INFINITE)
    Exit Sub
ErrShell:
    MsgBox "������ �������!" & vbCrLf & Err.Description, vbCritical, App.Title
End Sub

Public Sub ShellDialog(RunCmd As String, Optional RunWindow As Long = vbNormalFocus, Optional RunDebug As Boolean = False)
    Application.StatusBar = "������ ������� ��������� � �������� �� ����������..."
    If RunDebug Then
        MsgBox BPrintF("������ %d ����. �� ���������� %s\n\n%s", Len(RunCmd), CurDir, RunCmd), vbExclamation, App.Title
    End If
    ShellWait RunCmd, RunWindow
    Application.StatusBar = False
End Sub
