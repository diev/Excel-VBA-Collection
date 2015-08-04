Attribute VB_Name = "MenuINI"
Option Explicit
Option Compare Text
Option Base 1
DefLng A-Z

Const MenuSection = "Menu"
Const ServiceMenu = "ServiceMenu"
Const SortMenu = "SortMenu"

Public Sub WriteMenuINI()
    Dim i As Long, s As String
    With App
        'Название нашего меню в Excel
        .Setting(MenuSection, "Caption") = CDos("&Банк-Клиент")
        .Setting(MenuSection, "Before") = 1 '0 - не вставлять
        
        'Контекстное меню по правой кнопке мыши
        '.Setting(MenuSection, "RClick") = 1
        
        'Линейный тулбар
        .Setting(MenuSection, "Bar") = CDos("Сити Инвест Банк")
        .Setting(MenuSection, "BarAdd") = 1
        .Setting(MenuSection, "BarPosition") = 1 '0=left, 1=top, 2=right, 3=bottom, 4=floating
        .Setting(MenuSection, "BarVisible") = 1
        
        'Менюобразный тулбар
        '.Setting(MenuSection, "Bar1") = CDos("Сити Инвест Банк Меню")
        '.Setting(MenuSection, "Bar1Menu") = CDos("[Пуск] Банк-Клиент")
        '.Setting(MenuSection, "Bar1Add") = 1
        '.Setting(MenuSection, "Bar1Position") = 3 '0=left, 1=top, 2=right, 3=bottom, 4=floating
        '.Setting(MenuSection, "Bar1Visible") = 0
        
        i = 0
        i = i + 1: .Setting(MenuSection, CStr(i)) = CDos("&Вход в систему...") & "\LogonShow\59"
        i = i + 1: .Setting(MenuSection, CStr(i)) = CDos("-&Найти...") & "\FindText\279"
        i = i + 1: .Setting(MenuSection, CStr(i)) = CDos("&Создать...") & "\PlatEnterShow\64"
        i = i + 1: .Setting(MenuSection, CStr(i)) = CDos("&Загрузить с диска...") & "\ImportList\270"
        i = i + 1: .Setting(MenuSection, CStr(i)) = CDos("-&Просмотр и печать") & "\PreviewPlat\2174"
        i = i + 1: .Setting(MenuSection, CStr(i)) = CDos("З&аписать на диск...") & "\ExportList\271"
        i = i + 1: .Setting(MenuSection, CStr(i)) = CDos("За&шифровать к отправке") & "\ExportPlat\277"
        i = i + 1: .Setting(MenuSection, CStr(i)) = CDos("-&Отправка и прием...") & "\MailBoxShow\275"
        i = i + 1: .Setting(MenuSection, CStr(i)) = CDos("-С&ервис") & "\\" & ServiceMenu & "\\" & CDos("Сервис")
        i = i + 1: .Setting(MenuSection, CStr(i)) = CDos("О п&рограмме") & "\Info\1954"
        i = i + 1: .Setting(MenuSection, CStr(i)) = CDos("-В&ыход со сменой пароля") & "\SavePassShow\276"
        .Setting(MenuSection, "Count") = i
        
        i = 0
        i = i + 1: .Setting(ServiceMenu, CStr(i)) = CDos("&Справочник БИК...") & "\LSShow\176"
        i = i + 1: .Setting(ServiceMenu, CStr(i)) = CDos("-С&ортировка строк") & "\\" & SortMenu & "\\" & CDos("Сортировка")
        i = i + 1: .Setting(ServiceMenu, CStr(i)) = CDos("&Удалить строки") & "\DelRows\67"
        i = i + 1: .Setting(ServiceMenu, CStr(i)) = CDos("-&Текущий остаток...") & "\AmountChange\52"
        i = i + 1: .Setting(ServiceMenu, CStr(i)) = CDos("-&Реквизиты клиента...") & "\EditUserShow\2148"
        'i = i + 1: .Setting(ServiceMenu,CStr(i)) = CDos("Ключи клиента...") & "\UserPrivateShow\2148"
        i = i + 1: .Setting(ServiceMenu, CStr(i)) = CDos("&Добавить клиента...") & "\NewUserShow\2141"
        i = i + 1: .Setting(ServiceMenu, CStr(i)) = CDos("Уд&алить клиента") & "\DelUser\2151"
        i = i + 1: .Setting(ServiceMenu, CStr(i)) = CDos("-&Перезапуск программы") & "\Restart\2144"
        .Setting(ServiceMenu, "Count") = i
    
        i = 0
        i = i + 1: .Setting(SortMenu, CStr(i)) = CDos("&Номер") & "\SortByDocNo\29"
        i = i + 1: .Setting(SortMenu, CStr(i)) = CDos("&Дата") & "\SortByDocDate\29"
        i = i + 1: .Setting(SortMenu, CStr(i)) = CDos("&Сумма") & "\SortBySum\29"
        i = i + 1: .Setting(SortMenu, CStr(i)) = CDos("&Получатель") & "\SortByName\29"
        i = i + 1: .Setting(SortMenu, CStr(i)) = CDos("Н&азначение") & "\SortByDetails\29"
        .Setting(SortMenu, "Count") = i
    End With
End Sub

