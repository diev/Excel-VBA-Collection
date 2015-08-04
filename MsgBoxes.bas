Attribute VB_Name = "MsgBoxes"
Option Explicit
Option Compare Text
Option Base 1
DefLng A-Z

Public Sub InfoBox(FormatStr As String, ParamArray Args() As Variant)
    MsgBox Bvsprintf(FormatStr, CVar(Args)), vbInformation, App.TITLE
End Sub

Public Sub WarnBox(FormatStr As String, ParamArray Args() As Variant)
    MsgBox Bvsprintf(FormatStr, CVar(Args)), vbExclamation, App.TITLE
End Sub

Public Sub StopBox(FormatStr As String, ParamArray Args() As Variant)
    MsgBox Bvsprintf(FormatStr, CVar(Args)), vbCritical, App.TITLE
End Sub

Public Function OkCancelBox(FormatStr As String, ParamArray Args() As Variant) As Boolean
    OkCancelBox = MsgBox(Bvsprintf(FormatStr, CVar(Args)), vbExclamation + vbOKCancel, _
        App.TITLE) = vbOK
End Function

Public Function YesNoBox(FormatStr As String, ParamArray Args() As Variant) As Boolean
    YesNoBox = MsgBox(Bvsprintf(FormatStr, CVar(Args)), vbQuestion + vbYesNo, _
        App.TITLE) = vbyes
End Function

Public Function YesNoCancelBox(FormatStr As String, ParamArray Args() As Variant) As Long
    YesNoCancelBox = MsgBox(Bvsprintf(FormatStr, CVar(Args)), vbQuestion + vbYesNoCancel, _
        App.TITLE)
End Function



