Attribute VB_Name = "DirRecursion"
Dim r As Long

Public Sub AllTree(Optional StartFolder As String = "C:\")
    Dim fso As Object
    
    With Range("A1:E1")
        .CurrentRegion.Clear
        .Value = Array("File", "Size", "Date", "Time", "Path")
        .Font.Bold = True
    End With
    
    r = 1
    Set fso = CreateObject("Scripting.FileSystemObject")
    AllTreeFiles fso.GetFolder(StartFolder)
    
    Application.Goto Range("A1")
    Application.StatusBar = False
    MsgBox CStr(r - 1) & " загружено", vbInformation
End Sub

Public Sub AllTreeFiles(fd As Object)
    Dim f1 As Object
    If fd.Path = Environ("windir") Then Exit Sub 'skip C:\WINDOWS
    If r > 65000 Then
        MsgBox "Превышение листа MS Excel", vbExclamation
        Exit Sub
    End If
    With Application
        .StatusBar = CStr(r) & " " & fd.Path
        '.ScreenUpdating = False
        For Each f1 In fd.Files
            r = r + 1
            Cells(r, 1) = f1.Name
            Cells(r, 2) = f1.Size
            Cells(r, 3) = Format(f1.DateLastModified, "dd.mm.yy")
            Cells(r, 4) = Format(f1.DateLastModified, "hh:mm")
            Cells(r, 5) = fd.Path
        Next
        '.ScreenUpdating = True
    End With
    DoEvents
    For Each f1 In fd.SubFolders
        AllTreeFiles f1
    Next
End Sub

Public Sub CleanOnce()
    'Sort before by File, then Size!
    Dim s1 As String, s2 As String, bSame As Boolean
    Dim i As Long, n As Long
    i = 2: n = 0: bSame = False
    Do
        s1 = UCase(Cells(i, 1).Text) & Cells(i, 2).Text
        s2 = UCase(Cells(i + 1, 1).Text) & Cells(i + 1, 2).Text
        If Len(s2) = 0 Then Exit Do
        If s1 = s2 Then
            bSame = True
            i = i + 1
        ElseIf bSame Then
            bSame = False
            Application.StatusBar = CStr(n) & " del / " & CStr(i) & " dbl - " & Cells(i, 1).Text
            i = i + 1
            DoEvents
        Else
            Rows(i).Delete: n = n + 1
        End If
    Loop
    Application.StatusBar = False
    s1 = CStr(n) & " del / " & CStr(i) & " dbl"
    MsgBox s1, vbInformation
End Sub
