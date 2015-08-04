Attribute VB_Name = "BnkForms"
Option Explicit
Option Base 1
DefLng A-Z

Public Function FileEnterByMask(Optional FileMask As String = "*.*", Optional MaskEdit As Boolean = False) As String
    Load FileEnter
    With FileEnter
        .FileMask = FileMask
        .MaskEdit = MaskEdit
        .FillList
        .Show
        FileEnterByMask = .FileMask
    End With
    Unload FileEnter
End Function

Public Function DateByCalendar(Optional StartDate As Variant) As Variant
    Load Calendar
    With Calendar
        If IsMissing(StartDate) Then
            StartDate = Date
        End If
        .DateEntered = StartDate
        .Show
        DateByCalendar = .DateEntered
    End With
    Unload Calendar
End Function

Public Sub CalendarShow()
    Calendar.Show
    Unload Calendar
End Sub

