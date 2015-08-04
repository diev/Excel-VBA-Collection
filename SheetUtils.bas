Attribute VB_Name = "SheetUtils"
Option Explicit
Option Compare Text
Option Base 1
DefLng A-Z

Public Sub PrintDBFDigest()
Attribute PrintDBFDigest.VB_Description = "Макрос записан 20.06.2006 (User)"
Attribute PrintDBFDigest.VB_ProcData.VB_Invoke_Func = " \n14"
    Application.ScreenUpdating = False
    Columns("A:B").Select
    Selection.Delete Shift:=xlToLeft
    Columns("B:P").Select
    Selection.Delete Shift:=xlToLeft
    Columns("C:C").Select
    Selection.Delete Shift:=xlToLeft
    Columns("G:Q").Select
    Selection.Delete Shift:=xlToLeft
    Columns("H:BD").Select
    Selection.Delete Shift:=xlToLeft
    Columns("I:AO").Select
    Selection.Delete Shift:=xlToLeft
    Columns("F:F").Select
    Selection.Delete Shift:=xlToLeft
    
    Application.ScreenUpdating = True
    Columns("A:H").Select
    Columns("A:H").EntireColumn.AutoFit
    Range("A1").Select
End Sub
