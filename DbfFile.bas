Attribute VB_Name = "DbfFile"
Option Explicit
Option Compare Text
Option Base 1
DefLng A-Z

Const DBF_C As Byte = 67
Const DBF_S As Byte = 167 'No conversion Dos->Win
Const DBF_D As Byte = 68
Const DBF_N As Byte = 78

'Public Function StoD(YYYYMMDD As String) As Date
'    StoD = DateSerial(Val(Left(YYYYMMDD, 4)), _
'        Val(mID(YYYYMMDD, 5, 2)), Val(Right(YYYYMMDD, 2)))
'End Function
'
'Public Function DtoS(YYYYMMDD As Date) As String
'    DtoS = Format(YYYYMMDD, "yyyyMMdd")
'End Function
'
'Public Function CtoD(DDMMYYYY As String) As Date
'    CtoD = RDate(DDMMYYYY)
'End Function
'
'Public Function DtoC(DDMMYYYY As Date) As String
'    DtoC = Format(DDMMYYYY, "dd.MM.yyyy")
'End Function

Public Sub ReadDbfFile(File As String, Optional StartRow As Long = 2)
    Dim db As Range, r, c, p
    Dim nFile As Long, Buffer As String, BufField As String, Bytes() As Byte
    
    Dim RecCount As Long
    Dim DataOffset As Long
    Dim RecSize As Long
    Dim FieldsCount As Long
    Dim FieldType() As Byte
    Dim FieldLen() As Long
    
    On Error GoTo ErrFile
    nFile = FreeFile
    Open File For Binary Access Read Lock Write As nFile
    'Application.Cursor = xlWait
    'Application.ScreenUpdating = False
    r = LOF(nFile)
    If r > 4096 Then r = 4096 'one FAT32 disk cluster
    Buffer = Input(r, nFile)
    StrToBytes Buffer, Bytes
    RecCount = BytesToValue(Bytes, 5, 4)
    DataOffset = BytesToValue(Bytes, 9, 2)
    RecSize = BytesToValue(Bytes, 11, 2)
    FieldsCount = DataOffset \ 32 - 1  'minus head
    ReDim FieldType(FieldsCount)
    ReDim FieldLen(FieldsCount)

    p = 33 'Header size + 1st
    For c = 1 To FieldsCount
        FieldType(c) = Bytes(p + 11)
        FieldLen(c) = Bytes(p + 16)
        p = p + 32
    Next
    
    'No conversion Dos->Win (selected manually for speed)
    FieldType(1) = DBF_S
    FieldType(2) = DBF_S
    FieldType(7) = DBF_S
    FieldType(8) = DBF_S
    FieldType(9) = DBF_S
    
    Seek nFile, DataOffset + 1
    Set db = Worksheets(User.ID).Range("A" & CStr(StartRow))
    
    On Error Resume Next 'skip every bad data!
    For r = 1 To RecCount
        Buffer = Input(RecSize, nFile)
        p = 2 'skip the delete flag byte
        For c = 1 To FieldsCount
            BufField = Trim(Mid(Buffer, p, FieldLen(c)))
            Select Case FieldType(c)
                Case DBF_C
                    db.Cells(r, c) = CWin(BufField)
                Case DBF_S
                    db.Cells(r, c) = BufField
                Case DBF_D
                    If Len(BufField) > 0 Then db.Cells(r, c) = StoD(BufField)
                Case DBF_N
                    db.Cells(r, c) = Val(BufField)
            End Select
            p = p + FieldLen(c)
        Next
        If r Mod 100 = 0 Then
            DoEvents
            'Application.ScreenUpdating = False
        End If
    Next
ErrFile:
    'Application.ScreenUpdating = True
    Application.Cursor = xlDefault
    Close nFile
    Set db = Nothing
End Sub

