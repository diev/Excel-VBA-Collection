Attribute VB_Name = "MsgBoxes"
Option Explicit
Option Compare Text
Option Base 1
DefLng A-Z

Public Sub InfoBox(FormatStr As String, ParamArray Args() As Variant)
    MsgBox Bvsprintf(FormatStr, CVar(Args)), vbInformation, App.Title
End Sub

Public Sub InfoHelpBox(FormatStr As String, ParamArray Args() As Variant)
    MsgBox Bvsprintf(FormatStr, CVar(Args)), vbInformation + vbMsgBoxHelpButton, _
        App.Title, App.HelpFile, Args(UBound(Args))
End Sub

Public Sub WarnBox(FormatStr As String, ParamArray Args() As Variant)
    MsgBox Bvsprintf(FormatStr, CVar(Args)), vbExclamation, App.Title
End Sub

Public Sub WarnHelpBox(FormatStr As String, ParamArray Args() As Variant)
    MsgBox Bvsprintf(FormatStr, CVar(Args)), vbExclamation + vbMsgBoxHelpButton, _
        App.Title, App.HelpFile, Args(UBound(Args))
End Sub

Public Sub StopBox(FormatStr As String, ParamArray Args() As Variant)
    MsgBox Bvsprintf(FormatStr, CVar(Args)), vbCritical, App.Title
End Sub

Public Sub StopHelpBox(FormatStr As String, ParamArray Args() As Variant)
    MsgBox Bvsprintf(FormatStr, CVar(Args)), vbCritical + vbMsgBoxHelpButton, _
        App.Title, App.HelpFile, Args(UBound(Args))
End Sub

Public Function OkCancelBox(FormatStr As String, ParamArray Args() As Variant) As Boolean
    OkCancelBox = MsgBox(Bvsprintf(FormatStr, CVar(Args)), vbExclamation + vbOKCancel, _
        App.Title) = vbOK
End Function

Public Function OkCancelHelpBox(FormatStr As String, ParamArray Args() As Variant) As Boolean
    OkCancelHelpBox = MsgBox(Bvsprintf(FormatStr, CVar(Args)), vbExclamation + vbOKCancel + _
        vbMsgBoxHelpButton, App.Title, App.HelpFile, Args(UBound(Args))) = vbOK
End Function

Public Function YesNoBox(FormatStr As String, ParamArray Args() As Variant) As Boolean
    YesNoBox = MsgBox(Bvsprintf(FormatStr, CVar(Args)), vbQuestion + vbYesNo, _
        App.Title) = vbYes
End Function

Public Function YesNoHelpBox(FormatStr As String, ParamArray Args() As Variant) As Boolean
    YesNoHelpBox = MsgBox(Bvsprintf(FormatStr, CVar(Args)), vbQuestion + vbYesNo + _
        vbMsgBoxHelpButton, App.Title, App.HelpFile, Args(UBound(Args))) = vbYes
End Function



