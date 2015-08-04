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
Const AppVersion = "14.03.2002"

Const CIBLS = "40702810100000000000"
Const CIBBIC = "044030702"

Dim mBookName As String
Dim mAppPath As String
Dim mAppIni As String
Dim mAppHelp As String
Dim mCloseAllowed As Boolean

Public Property Get File() As String
    File = ThisWorkbook.FullName
End Property

Public Property Get BookFile() As String
    BookFile = Workbooks(mBookName).FullName
End Property

Public Property Get BookName() As String
    BookName = mBookName
End Property

Public Property Get Path() As String
    Path = mAppPath
End Property

Public Property Get Ini() As String
    Ini = mAppIni
End Property

Public Property Get HelpFile() As String
    HelpFile = mAppHelp
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
    If User.Demo Then
        s = Bsprintf("ДЕМОНСТРАЦИОННАЯ ВЕРСИЯ!\n")
    Else
        s = Bsprintf("Номер клиента: %s\n%s\n", User.ID, User.Name)
    End If
    
    fmt = "\nРабочая книга Excel: %s\n"
    s = s & Bsprintf(fmt, Workbooks(mBookName).FullName)
    
    If PGP.Valid Then
        fmt = "Ключи PGP клиента: %s\n"
        s = s & Bsprintf(fmt, User.KeysPath)
    End If
    
    fmt = "\nПрограмма Банк-Клиент: %s\nДата выпуска: %s\n"
    s = s & Bsprintf(fmt, ThisWorkbook.FullName, App.Version)
    
    With BnkSeek2
        If .Valid Then
            .File = .File 'ReInitHeader for .Updated
            fmt = "\nСправочник БИК: %s\nПоследнее обновление: %s\n"
            s = s & Bsprintf(fmt, .File, DtoC(.Updated))
        Else
            s = s & Bsprintf("\n%s\n", .None)
        End If
    End With
    
    With PGP
        If .Valid Then
            fmt = "\nПрограмма PGP: %s\n"
            s = s & Bsprintf(fmt, .File)
        Else
            s = s & Bsprintf("\n%s\n", .None)
        End If
    End With
    
    With SMail
        If .Valid Then
            fmt = "Программа SMail: %s\n"
            s = s & Bsprintf(fmt, .File)
        Else
            s = s & Bsprintf("%s\n", .None)
        End If
    End With
    
    fmt = "\nПрием: %s\nОтправка: %s\n"
    s = s & Bsprintf(fmt, SMail.Recv, SMail.Send)
    
    fmt = "\nЗагрузка: %s\nВыгрузка: %s"
    s = s & Bsprintf(fmt, User.ImportList, User.ExportList)
    
    InfoHelpBox "%s", s, 1
    
    SendInfo
End Sub

Private Sub Class_Initialize()
    With ActiveWorkbook
        mBookName = .Name
        mAppPath = FilePath(.FullName)
        mAppIni = ChangeFileExt(.FullName, "ini")
        If Options("WorkPath") = vbNullString Then
            Options("WorkPath") = mAppPath
        End If
    End With
    With ThisWorkbook
        mAppHelp = ChangeFileExt(.FullName, "hlp")
        'Application.Help = mAppHelp
    End With
End Sub

Public Property Get Options(Key As String) As String
    Options = ReadIniFile(mAppIni, BClient2Section, Key)
End Property

Public Property Let Options(Key As String, ByVal vNewValue As String)
    WriteIniFile mAppIni, BClient2Section, Key, CStr(vNewValue)
End Property

Public Property Get Setting(Section As String, Key As String) As Variant
    Setting = ReadIniFile(mAppIni, Section, Key)
End Property

Public Property Get NoSetting(Section As String, Key As String) As Boolean
    NoSetting = EmptyIniFile(mAppIni, Section, Key)
End Property

Public Property Let Setting(Section As String, Key As String, ByVal vNewValue As Variant)
    WriteIniFile mAppIni, Section, Key, CStr(vNewValue)
End Property

Public Property Get CloseAllowed() As Boolean
    CloseAllowed = mCloseAllowed
End Property

Public Property Let CloseAllowed(ByVal vNewValue As Boolean)
    mCloseAllowed = vNewValue
End Property

