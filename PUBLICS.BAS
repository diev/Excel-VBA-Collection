Attribute VB_Name = "Publics"
Option Explicit
Option Compare Text
Option Base 1
DefLng A-Z

Public Const AppName = "Plt2000" 'Registry folder
Public Const AppTitle = "Банк-Клиент ЗАО ""Сити Инвест Банк"""
Public Const AppUser = "Банк-Клиент "
Public Const NoUser = "НЕТ РЕГИСТРАЦИИ!"

'Add New.../Empty in lists
Public Const AddNewInList = "< Добавить... >"
Public Const NoneInList = "< нет >"
Public Const EmptyInList = ""

Public Payment As New CPayment
Public Archive As New CArchive
Public Bank As New CBank
