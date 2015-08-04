VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CApp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Option Base 1
DefLng A-Z

Const AppTitle = "Банк-Клиент ЗАО ""Сити Инвест Банк"""
Const AppVersion = "06.09.2006"

Const CIBLS = "40702810100000000000"
Const CIBBIC = "044030702"

Public Property Get Path() As String
    Path = FilePath(ActiveWorkbook.FullName)
End Property

Public Property Get TITLE() As String
    TITLE = AppTitle & " вер. " & AppVersion
End Property

Public Property Get Version() As String
    Version = AppVersion
End Property

Public Property Get DefBIC() As String
    DefBIC = CIBBIC
End Property

Public Property Get DefLS(Optional ID As String = vbNullString) As String
    Dim LS As String
    If ID = vbNullString Then
        DefLS = CIBLS
    Else
        LS = Left(CIBLS, 20 - Len(ID)) & ID
        LS = SetLSKey(CIBBIC, LS)
        DefLS = LS
    End If
End Property

Public Sub Info()
    Dim s As String, File As String, ID4 As String
    On Error Resume Next
    'If User.Demo Then
    '    s = Bsprintf("ДЕМОНСТРАЦИОННАЯ ВЕРСИЯ!\n")
    'Else
    '    With User
    '        s = Bsprintf("Номер клиента: %s\n%s\nЛ/с: %s\n", .ID, .Name, .LS)
    '    End With
    'End If
    '
    'fmt = "\nПрограмма Банк-Клиент: %s\nДАТА ВЕРСИИ: %s\n"
    's = s & Bsprintf(fmt, ThisWorkbook.FullName, App.Version)
    '
    'fmt = "\nРабочая книга Excel: %s\n"
    's = s & Bsprintf(fmt, ActiveWorkbook.FullName)
    '
    's = s & BnkSeek2.Info
    's = s & Crypto.Info
    's = s & SMail.Info
    '
    'fmt = "\nЗагрузка: %s\nВыгрузка: %s"
    's = s & Bsprintf(fmt, User.ImportList, User.ExportList)
    '
    'If OkCancelBox("%s\n\nОтчет о конфигурации будет выслан в Банк.\nЕго подготовка займет некоторое время.", s) Then
    '    SendInfo
    'End If
    
    s = vbNullString
    With User
        s = s & Bsprintf("Номер клиента: %s\n\n(%s)\n(счет %s)\n\n", _
            .ID4, .Name, .LS)
    End With
    
    With App
        s = s & Bsprintf("ВЕРСИЯ ПРОГРАММЫ: %s\nОбновлена: %s\nДиректория установки: %s\n\n", _
            .Version, .Options("Updated"), QFile(.Path))
        s = s & Bsprintf("%s%s\n", _
            FileInfo(ActiveWorkbook.FullName), _
            FileInfo(ThisWorkbook.FullName))
    End With
    
    s = s & Bsprintf("%s\n", Crypto.Info)
    
    With BnkSeek2
        File = .File
        If IsFile(File) Then
            .LoadFile
            s = s & Bsprintf("Справочник БИК от %s (банков: %d)\n%s\n", _
                DtoC(.Updated), .RecCount, FileInfo(File))
        Else
            s = s & Bsprintf("Справочник БИК не установлен!\n")
        End If
    End With
    
    With SMail
        File = App.Path & "SMail\Flags\Bank.fin"
        If IsFile(File) Then
            s = s & Bsprintf("Последний сеанс связи: %s (%s)\n", _
                FileDateTime(App.Path & "SMail\Flags\Bank.fin"), _
                IIf(.DialOrInet = 2, "через Интернет", "по модему"))
            s = s & Bsprintf("Файлов в принятом: %d (последний %s)\n\n", _
                CountFiles(.Recv & "*.*"), FileDateTime(App.Path & "SMail\Flags\Bank.rcv"))
        Else
            s = s & Bsprintf("Сеансов связи не было или SMail не установлен!\n\n")
        End If
    End With
    
    If User.Demo Then
        s = Bsprintf("ДЕМОНСТРАЦИОННАЯ ВЕРСИЯ\n\n%s", s)
    Else
        s = Bsprintf("ПРИ ЗВОНКЕ В БАНК НАЗОВИТЕ СЛЕДУЮЩЕЕ:\n\n%s", s)
    End If
    File = Dir(App.Path & "*.chm")
    If File <> vbNullString Then
        s = s & Bsprintf("Файлы помощи:\n")
        Do While File <> vbNullString
            s = s & Bsprintf("%s\n", FileInfo(App.Path & File))
            File = Dir
        Loop
        If MsgBox(s & "Показать помощь?", vbInformation + vbOKCancel, App.TITLE) = vbOK Then
            ShowHelp
        End If
    Else
        's = s & Bsprintf("Файлов помощи не найдено!")
        s = s & Bsprintf("Помощь на сайте www.cibank.ru в разделе Банк-Клиент")
        InfoBox "%s", s
    End If
    
End Sub

Public Function FileInfo(File As String) As String
    Dim s As String
    If FilePath(File) = App.Path Then
        s = FileNameExt(File)
    Else
        s = File
    End If
    FileInfo = Bsprintf("(файл: %s, дата: %s, размер: %s)\n", _
        QFile(s), FileDateTime(File), FileLen(File))
End Function

Public Property Get Options(Key As String) As String
    Options = GetSetting(ActiveWorkbook.Name, OptionsSection, Key, vbNullString)
End Property

Public Property Let Options(Key As String, ByVal vNewValue As String)
    SaveSetting ActiveWorkbook.Name, OptionsSection, Key, CStr(vNewValue)
End Property

Public Property Let DefaultOptions(Key As String, ByVal vNewValue As String)
    If GetSetting(ActiveWorkbook.Name, OptionsSection, Key, vbNullString) = vbNullString Then
        SaveSetting ActiveWorkbook.Name, OptionsSection, Key, vNewValue
    End If
End Property

Public Property Get BoolOptions(Key As String) As Boolean
    BoolOptions = GetSetting(ActiveWorkbook.Name, OptionsSection, Key, "0") = "1"
End Property

Public Property Let BoolOptions(Key As String, ByVal vNewValue As Boolean)
    SaveSetting ActiveWorkbook.Name, OptionsSection, Key, IIf(vNewValue, "1", "0")
End Property

Public Property Get Setting(Section As String, Key As String) As Variant
    Setting = GetSetting(ActiveWorkbook.Name, Section, Key)
End Property

Public Property Get NoSetting(Section As String, Key As String) As Boolean
    NoSetting = GetSetting(ActiveWorkbook.Name, Section, Key, vbNullString) = vbNullString
End Property

Public Property Let Setting(Section As String, Key As String, ByVal vNewValue As Variant)
    SaveSetting ActiveWorkbook.Name, Section, Key, CStr(vNewValue)
End Property

'Public Function FileSetting(Key As String, ByRef File As String) As Boolean
'    FileSetting = False
'    File = GetSetting(ActiveWorkbook.Name, FilesSection, Key, File)
'    If LocateFile(File) Then
'        SaveSetting ActiveWorkbook.Name, FilesSection, Key, File
'        FileSetting = True
'    End If
'End Function

Public Function LocateFile(ByRef File As String) As Boolean
    Dim f As Variant, s As String
    On Error Resume Next
    LocateFile = False
    With Application
        'set by default
        .DefaultFilePath = App.Path
        ChDrive .DefaultFilePath
        ChDir .DefaultFilePath
    
        .DefaultFilePath = FilePath(File) 'Some problems on some computers
        ChDrive .DefaultFilePath
        ChDir .DefaultFilePath
    End With
    WarnBox "Файл %s не найден!\nТребуется указать его местонахождение", File
    s = FileExt(File)
    s = Bsprintf("Файлы (*.%s),*.%s,Все файлы (*.*),*.*", s, s)
    f = Application.GetOpenFilename(s, 1, "Укажите " & File)
    If f <> False Then 'don't change this!
        File = CStr(f)
        LocateFile = True
    Else
        File = vbNullString
    End If
End Function

