Attribute VB_Name = "Export"
Option Explicit
Option Compare Text
Option Base 1
DefLng A-Z

Public Sub ExportComita()
    Dim fi As String, fo As String
    fi = Cells(1, 1)
    fo = "G:\OD\FORMS\ARM FM\import_oes\"
    FileCopy fi, fo & FileNameExt(fi)
    On Error Resume Next
    Shell "G:\OD\FORMS\ARM FM\COURIER.BAT"
End Sub

Public Sub ExportSVK()
    Dim fi As String, fo As String
    fi = Cells(1, 1)
    fo = "G:\OD\FORMS\F161p\out\"
    FileCopy fi, fo & FileNameExt(fi)
End Sub
