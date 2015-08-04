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
Const AppVersion = "2000-10-16"

Const CIBLS = "40702810100000000000"
Const CIBBIC = "044030702"

Dim mAppPath As String
Dim mAppIni As String
Dim mCloseAllowed As Boolean

Public Property Get File() As String
    File = ThisWorkbook.FullName
End Property

Public Property Get UserData() As String
    UserData = ActiveWorkbook.FullName
End Property

Public Property Get Path() As String
    Path = mAppPath
End Property

Public Property Get Ini() As String
    Ini = mAppIni
End Property

Public Property Get Title() As String
    Title = AppTitle
End Property

Public Property Get Version() As String
    Version = AppVersion
End Property

Public Property Get DefBIC() As String
    DefBIC = CIBBIC
End Property

Public Property Get DefLS() As String
    DefLS = CIBLS
End Property

Public Sub Info()
    Dim s As String, fmt As String
    s = BPrintF("PLT2000a - Программа ввода платежных поручений\n" & _
        "с интегрированной шифрацией и цифровой подписью PGP\nи почтовым обменом по системе SMail\n\n")
    If User.Demo Then
        s = s & BPrintF("ДЕМОНСТРАЦИОННАЯ ВЕРСИЯ!\n\n")
    Else
        s = s & BPrintF("Номер клиента: %s\n%s\n\n", User.ID, User.Name)
    End If
    
    fmt = Chr(169) & "1995-2000 Дмитрий Евдокимов\n" & _
        "Экспертная оценка: Игорь Соляр\n" & _
        "Тестирование: Ирина Равжаева\nВерсия: %s\n\n"
    s = s & BPrintF(fmt, App.Version)

    With Application
        fmt = "Microsoft Excel версии %s (%d)\n%s\n\n"
        s = s & BPrintF(fmt, .Version, .Build, .OperatingSystem)
    End With

    fmt = "Данные пользователя: \'%s\'\n" & _
        "Программа Банк-Клиент: \'%s\'\n\n"
    s = s & BPrintF(fmt, ActiveWorkbook.FullName, ThisWorkbook.FullName)
    
    If BnkSeek2.Valid Then
        fmt = "Справочник банков РФ версии 2: \'%s\'\n(последнее обновление: %s)\n\n"
        s = s & BPrintF(fmt, BnkSeek2.File, DtoSC(BnkSeek2.Updated))
    Else
        s = s & BPrintF(BnkSeek2.None & "\n")
    End If
    
    If PGP.Valid Then
        fmt = "Программа PGP: \'%s\'\n"
        s = s & BPrintF(fmt, PGP.File)
        
        'fmt = "Секр. ключи: \'%s\'\n" & _
        '    "Публ. ключи: \'%s\'\n"
        's = s & BPrintF(fmt, User.SecRing, User.PubRing)
    Else
        s = s & BPrintF(PGP.None & "\n")
    End If
    
    If SMail.Valid Then
        fmt = "Программа SMail: \'%s\'\n"
        s = s & BPrintF(fmt, SMail.File)
    Else
        s = s & BPrintF(SMail.None & "\n")
    End If
    
    fmt = "Прием: \'%s\'\nОтпр.: \'%s\'\n\n"
    s = s & BPrintF(fmt, SMail.Recv, SMail.Send)
    
    'fmt = "\nЗагрузка с диска: \'%s\'\nЗапись на диск: \'%s\'\n"
    's = s & BPrintF(fmt, User.ImportList, User.ExportList)
    
    MsgBox s, vbInformation, App.Title
End Sub

Private Sub Class_Initialize()
    With ActiveWorkbook
        mAppPath = Left(.FullName, InStrR(.FullName, "\"))
        mAppIni = mAppPath & Left(.Name, InStr(.Name, ".")) & "INI"
    End With
End Sub

Public Property Get Options(Key As String) As Variant
    Options = ReadIniFile(mAppIni, "Options", Key)
End Property

Public Property Let Options(Key As String, ByVal vNewValue As Variant)
    'Select Case VarType(vNewValue)
    '    Case vbBoolean:
    '        WriteIniFile mAppIni, "Options", Key, BoolToStr(vNewValue)
    '    Case vbString:
    '        WriteIniFile mAppIni, "Options", Key, vNewValue
    '    Case Else
            WriteIniFile mAppIni, "Options", Key, CStr(vNewValue)
    'End Select
End Property

Public Property Get Setting(Section As String, Key As String) As Variant
    Setting = ReadIniFile(mAppIni, Section, Key)
End Property

Public Property Let Setting(Section As String, Key As String, ByVal vNewValue As Variant)
    'Select Case VarType(vNewValue)
    '    Case vbBoolean:
    '        WriteIniFile mAppIni, Section, Key, BoolToStr(vNewValue)
    '    Case vbString:
    '        WriteIniFile mAppIni, Section, Key, vNewValue
    '    Case Else
            WriteIniFile mAppIni, Section, Key, CStr(vNewValue)
    'End Select
End Property

Public Property Get CloseAllowed() As Boolean
    CloseAllowed = mCloseAllowed
End Property

Public Property Let CloseAllowed(ByVal vNewValue As Boolean)
    mCloseAllowed = vNewValue
End Property
