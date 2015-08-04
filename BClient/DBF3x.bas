Attribute VB_Name = "DBF3x"
Option Explicit
Option Base 0 '1
DefLng A-Z

Const BufSizeLimit = 8192 '4096 is one FAT32 disk cluster

Const DBF_C As Byte = 67
Const DBF_D As Byte = 68
Const DBF_L As Byte = 76
Const DBF_M As Byte = 77 'Not supported!
Const DBF_N As Byte = 78

Public Type FieldType
    Name As String
    Type As Byte
    Len As Long
    Dec As Long
    dbfFormat As String
    xlsFormat As String
End Type

Public Type HeaderType
    File As String
    UpdateDate As String
    RecCount As Long
    DataOffset As Long
    RecSize As Long
    FieldsCount As Long
    Fields() As FieldType
End Type

Public Type RecordType
    RecNo As Long
    FieldsCount As Long
    Fields() As FieldType
    Values() As String
End Type

Private Sub ParseHeader(Header As HeaderType, sHeader As String)
    Dim Buf() As Byte, i, p
    
    StrToBytes Buf, sHeader
    With Header
        .UpdateDate = Format(DateSerial(Buf(1), Buf(2), Buf(3)), "dd.mm.yy")
        .RecCount = BytesToDWord(Buf, 4)
        .DataOffset = BytesToWord(Buf, 8)
        .RecSize = BytesToWord(Buf, 10)
        .FieldsCount = .DataOffset \ 32 - 1  'minus head
        ReDim .Fields(.FieldsCount)

        p = 32 'Header size
        For i = 1 To .FieldsCount
            'If Buf(p) = 13 Then Exit For 'OD terminator byte
        
            With .Fields(i)
                .Name = MidBytes(Buf, p, 10, 1)
                .Type = Buf(p + 11)
                .Len = Buf(p + 16)
                .Dec = 0
                .xlsFormat = "@"
                Select Case .Type
                    Case DBF_C
                        'Below is the standard but not our real life
                        .Len = BytesToWord(Buf, p + 16) 'comment this if it is going wrong
                        .dbfFormat = "C" & CStr(.Len)
                    Case DBF_D
                        .dbfFormat = "D8"
                        .xlsFormat = "dd.mm.yyyy"
                    Case DBF_L
                        .dbfFormat = "L1"
                    Case DBF_N
                        .Dec = Buf(p + 17) 'sometimes 18 ?!
                        If .Dec = 0 Then
                            .dbfFormat = "N" & CStr(.Len)
                            .xlsFormat = "#,##0"
                        Else
                            .dbfFormat = "N" & CStr(.Len) & "." & CStr(.Dec)
                            .xlsFormat = "#,##0." & String(.Dec, "0")
                        End If
                    Case DBF_M 'Not supported!
                        .dbfFormat = "MEMO"
                    Case Else 'Game Over!'
                        .dbfFormat = "ERROR!"
                End Select
            End With
            p = p + 32
        Next
    End With
End Sub

Private Function CreateHeader(Header As HeaderType) As String
    Dim Buf() As Byte, i, p, l As Long
    
    With Header
        FillBytes Buf, , , BufSizeLimit
                
        Buf(0) = 3 'No MEMO allowed!
        Buf(1) = Val(Format(Date, "yy")) 'Y2K!
        Buf(2) = Month(Date)
        Buf(3) = Day(Date)
        
        p = 32 'size of header
        l = 1 'delete mark
        For i = 1 To .FieldsCount
            With .Fields(i)
                InsBytes .Name, Buf, p
                Buf(p + 11) = .Type
                
                'real life maybe instead of the standard below
                If .Type = DBF_C Then
                    BytesFromWord .Len, Buf, p + 16
                Else
                    Buf(p + 16) = .Len
                    Buf(p + 17) = .Dec
                End If
                
                'Fast N-fields preparation
                If .Dec = 0 Then
                    .xlsFormat = "0"
                Else
                    .xlsFormat = "0." & String(.Dec, "0")
                End If
                l = l + .Len
            End With
            p = p + 32 'size of a field definition
        Next
        BytesFromDWord .RecCount, Buf, 4
        BytesFromWord p + 2, Buf, 8 'DataOffset
        BytesFromWord l, Buf, 10 'RecSize
    End With
    Buf(p) = 13 'OD terminator byte
    CreateHeader = LeftBytes(Buf, p + 2)
End Function

Public Sub DumpDbfFile()
    Dim Header As HeaderType
    Dim db As Range, r, C, p
    Dim nFile As Long, Buffer As String, Buf As String
    
    Application.StatusBar = "Инициализация..."
    Buf = Cells(1, 1).Text
    If Len(Buf) = 0 Or Buf = "False" Then
        ChDrive "C"
        ChDir "C:\Temp"
        Buf = Application.GetOpenFilename("Файлы 161п (*.??1;*.dbf),*.??1;*.dbf,Файлы DBF (*.dbf),*.dbf,Все файлы (*.*),*.*", 1, "Открытие файла DBF по 161п")
    End If
    
    Cells.Clear
    Application.StatusBar = "Открытие файла..."
    
    'On Error GoTo ErrFile
    nFile = FreeFile
    Open Buf For Binary Access Read Lock Write As nFile
    r = LOF(nFile)
    
    If r < 32 Then
        MsgBox "Файл не DBF или пуст", vbExclamation, "Ошибка"
        GoTo ErrFile
    End If
    
    If r > BufSizeLimit Then r = BufSizeLimit
    Buffer = Input(r, nFile)
    ParseHeader Header, Buffer
    On Error GoTo 0
    
    With Header
        .File = Buf
        For C = 1 To .FieldsCount
            With .Fields(C)
                Columns(C).NumberFormat = .xlsFormat
                Select Case .Type
                    Case DBF_C
                        Columns(C).Font.Color = QBColor(0)
                    Case DBF_D
                        Columns(C).Font.Color = QBColor(2)
                    Case DBF_L
                        Columns(C).Font.Color = QBColor(1)
                    Case DBF_N
                        Columns(C).Font.Color = QBColor(5)
                    Case Else
                        Columns(C).Font.Color = QBColor(12)
                End Select
            End With
        Next
        
        Seek nFile, .DataOffset + 1
        Set db = ActiveSheet.Range("A4")
        
        On Error Resume Next 'skip every bad data!
        For r = 1 To .RecCount
            Buffer = Input(.RecSize, nFile)
            p = 2 'skip the delete flag byte
            For C = 1 To .FieldsCount
                With .Fields(C)
                    Buf = Trim(Mid(Buffer, p, .Len))
                    Select Case .Type
                        Case DBF_C
                            If Len(Buf) > 0 Then
                                db.Cells(r, C) = CWin(Buf)
                            Else
                                db.Cells(r, C) = "0"
                            End If
                            If db.Cells(r, C) = "0" Then _
                                db.Cells(r, C).Font.Color = QBColor(12)
                        Case DBF_D
                            db.Cells(r, C) = StoD(Buf)
                 '           Select Case Year(db.Cells(r, c))
                 '               Case 2099
                 '                   db.Cells(r, c).Font.Color = QBColor(12)
                 '               Case Is > 2020
                 '                   db.Cells(r, c).Font.Color = QBColor(13)
                 '           End Select
                        Case DBF_L
                            db.Cells(r, C) = Buf
                        Case DBF_N
                            db.Cells(r, C) = Val(Buf)
                            If db.Cells(r, C) = 0 Then _
                                db.Cells(r, C).Font.Color = QBColor(12)
                        Case Else
                            db.Cells(r, C) = Buf
                    End Select
                    
                    'User Final Prefs
                    db.Cells(r, C).Font.Name = "Calibri"
                    db.Cells(r, C).Font.Size = 11
                    db.Cells(r, C).HorizontalAlignment = xlLeft
                    
                    p = p + .Len
                End With
            Next
            Application.StatusBar = "Загружено " & r & " из " & .RecCount
            If r Mod 10 = 0 Then
                If r = 30 Then Application.ScreenUpdating = False
                DoEvents
            End If
        Next
        On Error GoTo 0
        Close nFile
        
        Application.ScreenUpdating = False
        'For c = 1 To .FieldsCount
        '    Columns(c).AutoFit
        'Next
        
        Application.ScreenUpdating = True
        Cells(1, 1).Font.Color = QBColor(12)
        Cells(1, 1) = .File
        Cells(2, 1).Font.Color = QBColor(0)
        Cells(2, 1) = "Последнее изменение " & .UpdateDate & _
            ", записей " & .RecCount
        Rows(3).Font.Color = QBColor(9)
        For C = 1 To .FieldsCount
            With .Fields(C)
                Cells(3, C) = .Name & " " & .dbfFormat
            End With
        Next
    End With
    Application.StatusBar = False
    db.Cells(1, 1).Select
    Set db = Nothing
    Exit Sub
    
ErrFile:
    ErrorBox Err
    Close nFile
    Application.StatusBar = False
End Sub

Public Sub DumpDbfFile2()
    Dim Header As HeaderType
    Dim db As Range, r, C, p
    Dim nFile As Long, Buffer As String, Buf As String
    
    Application.StatusBar = "Инициализация..."
    Buf = Application.GetOpenFilename("Файлы 161п (*.0??;*.dbf),*.0??;*.dbf,Файлы DBF (*.dbf),*.dbf,Все файлы (*.*),*.*", 1, "Открытие файла DBF по 161п")
    
    Application.StatusBar = "Открытие файла..."
    
    'On Error GoTo ErrFile
    nFile = FreeFile
    Open Buf For Binary Access Read Lock Write As nFile
    r = LOF(nFile)
    
    If r < 32 Then
        MsgBox "Файл не DBF или пуст", vbExclamation, "Ошибка"
        GoTo ErrFile
    End If
    
    If r > BufSizeLimit Then r = BufSizeLimit
    Buffer = Input(r, nFile)
    ParseHeader Header, Buffer
    On Error GoTo 0
    
    With Header
        Seek nFile, .DataOffset + 1
        
        ActiveSheet.Range("A4").Select
        Selection.End(xlDown).Select

        Set db = ActiveSheet.Range("A4").CurrentRegion.End(xlDown).Offset(1, 0)
        
        On Error Resume Next 'skip every bad data!
        For r = 1 To .RecCount
            Buffer = Input(.RecSize, nFile)
            p = 2 'skip the delete flag byte
            For C = 1 To .FieldsCount
                With .Fields(C)
                    Buf = Trim(Mid(Buffer, p, .Len))
                    Select Case .Type
                        Case DBF_C
                            If Len(Buf) > 0 Then
                                db.Cells(r, C) = CWin(Buf)
                            Else
                                db.Cells(r, C) = "0"
                            End If
                            If db.Cells(r, C) = "0" Then _
                                db.Cells(r, C).Font.Color = QBColor(12)
                        Case DBF_D
                            db.Cells(r, C) = StoD(Buf)
                 '           Select Case Year(db.Cells(r, c))
                 '               Case 2099
                 '                   db.Cells(r, c).Font.Color = QBColor(12)
                 '               Case Is > 2020
                 '                   db.Cells(r, c).Font.Color = QBColor(13)
                 '           End Select
                        Case DBF_L
                            db.Cells(r, C) = Buf
                        Case DBF_N
                            db.Cells(r, C) = Val(Buf)
                            If db.Cells(r, C) = 0 Then _
                                db.Cells(r, C).Font.Color = QBColor(12)
                        Case Else
                            db.Cells(r, C) = Buf
                    End Select
                    
                    'User Final Prefs
                    db.Cells(r, C).Font.Name = "Calibri"
                    db.Cells(r, C).Font.Size = 11
                    db.Cells(r, C).HorizontalAlignment = xlLeft
                    
                    p = p + .Len
                End With
            Next
            Application.StatusBar = "Загружено " & r & " из " & .RecCount
            If r Mod 10 = 0 Then
                If r = 30 Then Application.ScreenUpdating = False
                DoEvents
            End If
        Next
        On Error GoTo 0
        Close nFile
        
        'Application.ScreenUpdating = False
        'For c = 1 To .FieldsCount
        '    Columns(c).AutoFit
        'Next
        
        Application.ScreenUpdating = True
        'Cells(1, 1).Font.Color = QBColor(12)
        'Cells(1, 1) = .File
        'Cells(2, 1).Font.Color = QBColor(0)
        'Cells(2, 1) = "Последнее изменение " & .UpdateDate & _
        '    ", записей " & .RecCount
        'Rows(3).Font.Color = QBColor(9)
        'For c = 1 To .FieldsCount
        '    With .Fields(c)
        '        Cells(3, c) = .Name & " " & .dbfFormat
        '    End With
        'Next
    End With
    Application.StatusBar = False
    db.Cells(1, 1).Select
    Set db = Nothing
    Exit Sub
    
ErrFile:
    ErrorBox Err
    Close nFile
    Application.StatusBar = False
End Sub

Public Sub WriteDbfFile() 'dbfgen
    Dim db As Range, Header As HeaderType
    Dim r, C, p, EmptyRow As Boolean
    Dim nFile As Long, Buffer As String, Buf As String
    
    Application.StatusBar = "Инициализация..."
    With Header
        .File = Cells(1, 1)
        If Len(.File) = 0 Or .File = "False" Then
            .File = Application.GetSaveAsFilename("30702000", "Файлы 161п (*.0??),*.0??,Файлы DBF (*.dbf),*.dbf,Все файлы (*.*),*.*", 1, "Сохранение файла DBF по 161п")
        End If
    
        '.FieldsCount = Cells(3, 1).End(xlToRight).Column
        C = 1
        Do
            If Len(Trim(Cells(3, C).Text)) = 0 Then
                .FieldsCount = C - 1
                Exit Do
            End If
            C = C + 1
        Loop
        ReDim .Fields(.FieldsCount)
        
        p = 1 'del flag
        For C = 1 To .FieldsCount
            With .Fields(C)
                Buf = Trim(Cells(3, C).Text)
                r = InStr(Buf, " ")
                .Name = Left(Buf, r - 1)
                .dbfFormat = Trim(Mid(Buf, r + 1))
                .Type = Asc(UCase(.dbfFormat))
                .Len = Fix(Val(Mid(.dbfFormat, 2)))
                If .Type = DBF_N Then
                    .Dec = Val(Mid(.dbfFormat, InStr(2, .dbfFormat, ".") + 1))
                Else
                    .Dec = 0
                End If
                p = p + .Len
            End With
        Next
        .RecSize = p
        
        'Check for correct length every C-cell and counting records
        Set db = ActiveSheet.Range("A4")
        r = 1
        Do
            EmptyRow = True
            For C = 1 To .FieldsCount
                With .Fields(C)
                    p = Len(Trim(db.Cells(r, C).Text))
                    If EmptyRow And p > 0 Then EmptyRow = False
                    If .Type = DBF_C And p > .Len Then
                        db.Cells(r, C).Select
                        MsgBox BPrintF("Размер символьного поля %d\nВведено %d", .Len, p), vbExclamation, "Ошибка"
                        Application.StatusBar = False
                        Exit Sub
                    End If
                End With
            Next
            Application.StatusBar = BPrintF("Проверка строки %d", r)
            DoEvents
            If EmptyRow Then
                .RecCount = r - 1
                Exit Do
            End If
            r = r + 1
        Loop
        
        Cells(1, 1) = .File
        Cells(2, 1) = "Последнее изменение " & Date & _
            ", записей " & .RecCount
        
        Buffer = CreateHeader(Header)
        
        On Error GoTo ErrFile
        'truncate file
        nFile = FreeFile
        Open .File For Output Access Write Lock Read Write As nFile
        Close nFile
    
        nFile = FreeFile
        Open .File For Binary Access Write Lock Read Write As nFile
        Put nFile, 1, Buffer 'whole prepared header
        On Error GoTo 0
        
        For r = 1 To .RecCount
            Buffer = " " 'the delete flag byte
            For C = 1 To .FieldsCount
                With .Fields(C)
                    Buf = Trim(db.Cells(r, C).Text)
                    'If Buf = "" Then
                    '    Buffer = Buffer & Space(.Len)
                    'Else
                        Select Case .Type
                            Case DBF_C
                                If Buf = "" Then Buf = "0"
                                Buffer = Buffer & Pad(CDos(Buf), .Len)
                            Case DBF_D
                                Buffer = Buffer & DtoS(Buf)
                            Case DBF_L
                                Buffer = Buffer & IIf(Buf = "T", "T", "F")
                            Case DBF_N
                                'using fast preparation
                                Buf = PadL(Format(RVal(Buf), .xlsFormat), .Len)
                                If .Dec > 0 And InStr(Buf, ",") > 0 Then
                                    Mid(Buf, InStr(Buf, ","), 1) = "."
                                End If
                                Buffer = Buffer & Buf
                            Case Else
                                Buffer = Buffer & Pad(Buf, .Len)
                        End Select
                    'End If
                End With
            Next
            Put nFile, , Buffer
            
            Application.StatusBar = "Записано " & r & " из " & .RecCount
            If r Mod 10 = 0 Then DoEvents
        Next
        
        Put nFile, , Chr(26) 'eof mark
        Close nFile
        
        Buf = .File
    End With
    Set db = Nothing
    Application.StatusBar = False
    
    Application.StatusBar = "Резервная копия..."
    FileCopy Buf, RepFile(Buf)
    Application.StatusBar = False
    Exit Sub
    
ErrFile:
    ErrorBox Err
    Close nFile
    Application.StatusBar = False
End Sub

Public Function CDbf(Value As Variant, dbfFormat As String) As String
    Dim t As String, l As Long, d As Long
    
    t = UCase(Left(dbfFormat, 1))
    l = Fix(Val(Mid(dbfFormat, 2)))
                    
    Select Case t
        Case "C"
            CDbf = Pad(CDos(CStr(Value)), l)
        'Case "D"
            'CDbf = DtoS(Value)
        Case "L"
            CDbf = IIf(CBool(Value), "T", "F")
        Case "N"
            d = Val(Mid(dbfFormat, InStr(2, dbfFormat, ".") + 1))
            If d = 0 Then
                CDbf = PadL(Format(RVal(CStr(Value)), "0"), l)
            Else
                CDbf = PadL(Format(RVal(CStr(Value)), "0." & String(d, "0")), l)
                Mid(CDbf, Len(CDbf) - d, 1) = "."
            End If
        Case Else
            CDbf = Pad(CStr(Value), l)
    End Select
End Function

Public Function SeekDbfFile(File As String, Field As String, Value As String) As RecordType
    Dim Header As HeaderType
    Dim r, C, p, nField
    Dim nFile As Long, Buffer As String, Buf As String
    
    SeekDbfFile.RecNo = 0
    Application.StatusBar = "Поиск в файле..."
    
    On Error GoTo ErrFile
    nFile = FreeFile
    Open File For Binary Access Read Lock Write As nFile
    r = LOF(nFile)
    If r > BufSizeLimit Then r = BufSizeLimit
    Buffer = Input(r, nFile)
    ParseHeader Header, Buffer
    On Error GoTo 0
    
    With Header
        p = 1
        nField = 0
        ReDim SeekDbfFile.Fields(.FieldsCount)
        Buf = UCase(Field)
        For C = 1 To .FieldsCount
            With SeekDbfFile.Fields(C)
                .Name = Header.Fields(C).Name
                .Type = Header.Fields(C).Type
                .Len = Header.Fields(C).Len
                .Dec = Header.Fields(C).Dec
                .dbfFormat = Header.Fields(C).dbfFormat
                .xlsFormat = Header.Fields(C).xlsFormat
            End With
        Next
        For C = 1 To .FieldsCount
            With SeekDbfFile.Fields(C)
                If .Name = Buf Then
                    nField = C
                    Buf = String(.Len, " ")
                    Exit For
                Else
                    p = p + .Len
                End If
            End With
        Next
        If nField = 0 Then
            Close nFile
            MsgBox "Поле не найдено: " & Field, vbExclamation, "Ошибка"
            Exit Function
        End If
        
        On Error Resume Next 'skip every bad data!
        For r = 1 To .RecCount
            Get nFile, (.DataOffset + 1) + (r - 1) * .RecSize + p, Buf
            If Buf = Value Then
                Seek nFile, (.DataOffset + 1) + (r - 1) * .RecSize
                Buffer = Input(.RecSize, nFile)
                p = 2 'delete flag byte
                With SeekDbfFile
                    .RecNo = r
                    .FieldsCount = Header.FieldsCount
                    ReDim .Values(.FieldsCount)
                    For C = 1 To .FieldsCount
                        .Values(C) = Mid(Buffer, p, .Fields(C).Len)
                        p = p + .Fields(C).Len
                    Next
                End With
                Exit For
                If r Mod 10 = 0 Then DoEvents
            End If
        Next
        On Error GoTo 0
        Close nFile
    End With
    Application.StatusBar = False
    Exit Function
    
ErrFile:
    ErrorBox Err
    Close nFile
    Application.StatusBar = False
End Function

Public Function FieldValue(Rec As RecordType, Name As String) As Variant
    Dim C, Buf As String
    
    Buf = UCase(Name)
    With Rec
        For C = 1 To .FieldsCount
            If .Fields(C).Name = Buf Then
                Buf = Trim(.Values(C))
                Select Case .Fields(C).Type
                    Case DBF_C
                        FieldValue = CWin(Buf)
                    Case DBF_D
                        If Len(Buf) > 0 Then FieldValue = StoD(Buf)
                    Case DBF_L
                        FieldValue = CBool(Buf)
                    Case DBF_N
                        If Val(Buf) <> 0 Then _
                            FieldValue = Val(Buf)
                    Case Else
                        FieldValue = Buf
                End Select
            End If
        Next
    End With
End Function

Public Function RepFile(Buf As String) As String
    Dim y As String, m As String, d As String
    Dim dt As String, p As String, C As Long
    
    dt = DtoS(Now)
    y = Left(dt, 4)
    m = Mid(dt, 5, 2)
    d = Right(dt, 2)
    
    'p = FilePath(Buf) & "Rep\" & y & "\" & m & "\" & d
    p = "G:\OD\FORMS\F161p\Rep\" & y & "\" & m & "\" & d
    ForceDirectories p
    p = RightPathName(p, FileNameExt(Buf))
    
    dt = p
    C = 1
    Do While IsFile(dt)
        C = C + 1
        dt = p & "." & C
    Loop
    RepFile = dt
End Function
