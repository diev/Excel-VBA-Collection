Option Explicit

Public Sub Jpg_To_Xlsx()
    Dim JpgFolder As String, JpgFile As String
    Dim XlsFolder As String, XlsFile As String
    
    Dim file As String
    Dim counter As Long: counter = 0
    
    With Application.FileDialog(msoFileDialogFolderPicker)
        .AllowMultiSelect = False
        
        .Title = "Выберите папку с исходными JPG"
        If .Show <> -1 Then Exit Sub
        JpgFolder = .SelectedItems(1) & "\"
        
        .Title = "Выберите папку для готовых XLSX"
        If .Show <> -1 Then Exit Sub
        XlsFolder = .SelectedItems(1) & "\"
    End With
    Application.StatusBar = "Ждите..."
    
    Application.ScreenUpdating = False
    'Application.EnableEvents = False
    Application.DisplayAlerts = False
    With Workbooks.Add
        .RemoveDocumentInformation (xlRDIAll)
    
        file = Dir(JpgFolder & "*.jpg")
        Do While Len(file) > 0
            counter = counter + 1
            'Application.StatusBar = "Делаем " & file
            
            JpgFile = JpgFolder & file
            XlsFile = XlsFolder & Replace(file, ".jpg", ".xlsx", , , vbTextCompare)
        
            With .Worksheets(1).Shapes.AddPicture(JpgFile, False, True, 0, 0, 0, 0)
                .LockAspectRatio = True
                .Select
            End With
            Selection.ShapeRange.ScaleHeight 1, msoTrue, msoScaleFromTopLeft
            Selection.PrintObject = msoFalse
                        
            .SaveAs Filename:=XlsFile, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
            .Worksheets(1).Shapes(1).Delete
            
            file = Dir
            'If counter = 10 Then Exit Do
            If counter Mod 5 = 0 Then
                Application.ScreenUpdating = True
                Application.StatusBar = "Ждите... (обработано " & counter & " файлов)"
                Application.ScreenUpdating = False
            End If
        Loop
        .Close False
    End With
    Application.StatusBar = False
    Application.ScreenUpdating = True
    'Application.EnableEvents = True
    Application.DisplayAlerts = True
End Sub
