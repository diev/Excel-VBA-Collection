Attribute VB_Name = "FuncShell"
Option Explicit
Option Compare Text
Option Base 1
DefLng A-Z

Private Type STARTUPINFO
    cb As Long
    lpReserved As String
    lpDesktop As String
    lpTitle As String
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Long
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type

Private Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessID As Long
    dwThreadID As Long
End Type

Private Const NORMAL_PRIORITY_CLASS = &H20&
Private Const INFINITE = -1&

Private Declare Function WaitForSingleObject Lib "kernel32" _
    (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long

Private Declare Function CreateProcessA Lib "kernel32" _
    (ByVal lpApplicationName As String, ByVal lpCommandLine As String, _
    ByVal lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, _
    ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, _
    ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As String, _
    lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) _
    As Long
    
Private Declare Function CloseHandle Lib "kernel32" _
    (ByVal hObject As Long) As Long
    
Private Declare Function GetExitCodeProcess Lib "kernel32" _
    (ByVal hProcess As Long, lpExitCode As Long) As Long

Public Function ShellWait(RunCmd As String, Optional RunWindow As Long = vbMinimizedNoFocus) As Long
    Dim proc As PROCESS_INFORMATION
    Dim start As STARTUPINFO
    Dim retVal As Long
    
    'Init
    start.cb = Len(start)
    
    'Start
    retVal = CreateProcessA(vbNullString, RunCmd, 0&, 0&, -1&, _
        NORMAL_PRIORITY_CLASS, 0&, vbNullString, start, proc)
            
    'Wait...
    retVal = WaitForSingleObject(proc.hProcess, INFINITE)
    
    'ExitCode
    GetExitCodeProcess proc.hProcess, retVal
    CloseHandle proc.hThread
    CloseHandle proc.hProcess
    
    'Return
    ShellWait = retVal 'And &HFF&
End Function

Public Function ShellDialog(RunCmd As String, Optional RunWindow As Long = vbNormalFocus, Optional RunDebug As Boolean = False) As Long
    Dim retVal As Long
    Application.StatusBar = "Запуск внешней программы и ожидание ее завершения..."
    If RunDebug Then
        WarnBox "Запуск %d симв. из директории %s\n\n%s", Len(RunCmd), CurDir, RunCmd
        retVal = ShellWait(RunCmd, vbNormalFocus)
        WarnBox "Код завершения: %d", retVal
    Else
        retVal = ShellWait(RunCmd, RunWindow)
    End If
    Application.StatusBar = False
    ShellDialog = retVal
End Function
