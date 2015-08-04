VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CDataSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit
Option Compare Text
Option Base 1
DefLng A-Z

Dim mDataSheet As String
Dim mHead As String
Dim mData As String
Dim mRecCount As Long
Dim mColCount As Long
Dim mLastRec As Long

Public Property Get RecCount() As Long
    RecCount = mRecCount
End Property

Public Property Get ColCount() As Long
    ColCount = mColCount
End Property

Public Property Get Sheet() As String
    Sheet = mDataSheet
End Property

Public Property Let Sheet(ByVal vNewValue As String)
    Reset vNewValue
End Property

Public Sub Reset(Optional DataSheet As String = "")
    If Len(DataSheet) > 0 Then mDataSheet = DataSheet
    mHead = mDataSheet & "!$A$1"
    mData = mDataSheet & "!$A$2"
    mRecCount = 0
    With Range(mData)
        Do While Len(.Offset(mRecCount).Text) > 0
            mRecCount = mRecCount + 1
        Loop
    End With
    mColCount = 0
    With Range(mHead)
        Do While Len(.Offset(, mColCount).Text) > 0
            mColCount = mColCount + 1
        Loop
    End With
    mLastRec = 0
End Sub

Public Property Get CurrentRangeAddress() As String
    If mRecCount > 0 Then
        CurrentRangeAddress = mDataSheet & "!" & _
            Range(mHead).CurrentRegion.Offset(1).Resize(mRecCount, mColCount).Address
    Else
        CurrentRangeAddress = ""
    End If
End Property

Public Property Get DataRangeAddress() As String
    DataRangeAddress = mDataSheet & "!" & Range(mData).Address
End Property

Public Property Get HeadRangeAddress() As String
    HeadRangeAddress = mDataSheet & "!" & Range(mHead).Address
End Property

Public Function GetRecord(Optional RecNo As Long = 0) As Variant
    Dim arr() As Variant, i
    ReDim arr(mColCount)
    If RecNo > 0 And RecNo <= mRecCount Then
        With Range(mData)
            For i = 1 To mColCount
                arr(i) = .Cells(RecNo, i)
            Next
        End With
    End If
    GetRecord = arr
End Function

Public Function GetField(RecNo As Long, FieldNo As Long) As Variant
    If RecNo > 0 And RecNo <= mRecCount Then
        With Range(mData)
            GetField = .Cells(RecNo, FieldNo)
        End With
    Else
        GetField = vbNull
    End If
End Function

Public Function GetFieldText(RecNo As Long, FieldNo As Long) As String
    If RecNo > 0 And RecNo <= mRecCount Then
        With Range(mData)
            GetFieldText = .Cells(RecNo, FieldNo).Text
        End With
    Else
        GetFieldText = ""
    End If
End Function

Public Function GetNickRecNo(Nick As String) As Long
    With Range(mData)
        For mLastRec = 1 To mRecCount
            If .Cells(mLastRec).Text = Nick Then
                GetNickRecNo = mLastRec
                Exit Function
            End If
        Next
    End With
    mLastRec = 0
    GetNickRecNo = 0
End Function

Public Function GetNickRecord(Nick As String) As Variant
    Dim arr() As Variant, i
    ReDim arr(mColCount)
    With Range(mData)
        For mLastRec = 1 To mRecCount
            If .Cells(mLastRec).Text = Nick Then
                For i = 1 To mColCount
                    arr(i) = .Cells(mLastRec, i)
                Next
                GetNickRecord = arr
                Exit Function
            End If
        Next
    End With
    mLastRec = 0
    GetNickRecord = Null
End Function

Public Function GetNextNickRecord(Nick As String) As Variant
    Dim arr() As Variant, i
    If mLastRec > 0 Then
        ReDim arr(mColCount)
        With Range(mData)
            For mLastRec = mLastRec + 1 To mRecCount
                If .Cells(mLastRec) = Nick Then
                    For i = 1 To mColCount
                        arr(i) = .Cells(mLastRec, i)
                    Next
                    GetNextNickRecord = arr
                    Exit Function
                End If
            Next
        End With
    End If
    mLastRec = 0
    GetNextNickRecord = Null
End Function

Public Sub PutRecord(RecNo As Long, ByVal vNewValue As Variant)
    Dim i
    With Range(mData)
        ScreenUpdating = False
            For i = 1 To mColCount
                .Cells(RecNo, i) = vNewValue(i)
            Next
        ScreenUpdating = False
    End With
End Sub

Public Sub AddRecord(ByVal vNewValue As Variant)
    Dim i
    mRecCount = mRecCount + 1
    With Range(mData)
        ScreenUpdating = False
            For i = 1 To mColCount
                .Cells(mRecCount, i) = vNewValue(i)
            Next
        ScreenUpdating = True
    End With
End Sub

Public Sub InsertRecord(ByVal vNewValue As Variant)
    Dim i
    mRecCount = mRecCount + 1
    With Worksheets(mDataSheet)
        Application.ScreenUpdating = False
            .Rows(2).Insert '1 for Headers, 2... for Data
            .Rows(3).Copy .Rows(2) 'copy empty row to clear header formats if any
            .Rows(2).ClearContents
            For i = 1 To mColCount
                .Cells(2, i) = vNewValue(i)
            Next
        Application.ScreenUpdating = True
    End With
End Sub

Public Function IsNickUnique(Nick As String) As Boolean
    Dim i
    IsNickUnique = False
    With Range(mData)
        For i = 1 To mRecCount
            If .Offset(i).Text = Nick Then Exit Function
        Next
    End With
    IsNickUnique = True
End Function

Public Function GetNickUnique(Nick As String) As String
    Dim i
    GetNickUnique = Nick
    i = 1
    Do Until IsNickUnique(GetNickUnique)
        i = i + 1
        GetNickUnique = Nick & " (" & CStr(i) & ")" 'Nick (2)
    Loop
End Function

Public Function ColName(Column As Long) As String
    ColName = Range(mHead).Offset(, Column - 1).Text
End Function

Public Function GetStrings(Column As Long, Optional TopCount As Long = 0) As Variant
    Dim arr() As String, i, j, n, s As String, b As Boolean
    If mRecCount = 0 Then
        GetStrings = Null
        Exit Function
    End If
    ReDim arr(mRecCount)
    n = 1
    With Range(mData)
        For i = 1 To mRecCount
            s = .Cells(i, Column).Text
            If Len(s) > 0 Then 'take filled only
                b = False
                For j = 1 To n 'was it before?
                    If arr(j) = s Then
                        b = True 'found previous!
                        Exit For
                    End If
                Next
                'If j > n Then 'too but this maybe wrong in other VBA
                If Not b Then 'add new
                    arr(n) = s
                    n = n + 1
                    If n = TopCount Then Exit For
                End If
            End If
        Next
    End With
    If n > 1 Then
        ReDim Preserve arr(n - 1)
        GetStrings = arr
    Else
        GetStrings = Null
    End If
End Function

Public Function GetList(ParamArray Column() As Variant) As Variant
    Dim arr As Variant, i, j, lb, ub
    If mRecCount = 0 Then
        GetList = Null
        Exit Function
    End If
    lb = LBound(Column)
    ub = UBound(Column)
    i = ub - lb + 1
    With Range(mData)
        If i = 1 Then 'separate 1D for speed
            ReDim arr(1 To mRecCount) As String
            j = Column(lb)
            For i = 1 To mRecCount
                arr(i) = .Cells(i, j).Text
            Next
        Else
            ReDim arr(1 To mRecCount, i) As String
            For i = 1 To mRecCount
                For j = lb To ub
                    arr(i, j + 1) = .Cells(i, Column(j)).Text
                Next
            Next
        End If
    End With
    GetList = arr
End Function

Public Function GetListFiltered(ParamArray Column() As Variant) As Variant
    Dim arr As Variant, farr As Variant, i, j, k, lb, ub, flb, fub, b As Boolean, n, x
    If mRecCount = 0 Then
        GetListFiltered = Null
        Exit Function
    End If
    lb = LBound(Column)
    ub = UBound(Column) - 1
    farr = Column(ub + 1) 'last param as filter array(k1, "v1", k2, "v2", ...)
    flb = LBound(farr)
    fub = UBound(farr)
    x = ub - lb + 1
    n = 0
    With Range(mData)
        If x = 1 Then 'separate 1D for speed
            ReDim arr(1 To mRecCount) As String
            j = Column(lb)
            For i = 1 To mRecCount
                b = True
                For k = flb To fub Step 2
                    If InStr(1, .Cells(i, farr(k)).Text, farr(k + 1)) = 0 Then
                        b = False
                        Exit For
                    End If
                Next
                If b Then
                    n = n + 1
                    arr(n) = .Cells(i, j).Text
                    DoEvents
                End If
            Next
            If n = 0 Then
                arr = Null
            Else
                ReDim Preserve arr(1 To n) As String
            End If
        Else
            ReDim arr(1 To mRecCount, x) As String
            For i = 1 To mRecCount
                b = True
                For k = flb To fub Step 2
                    If InStr(1, .Cells(i, farr(k)).Text, farr(k + 1)) = 0 Then
                        b = False
                        Exit For
                    End If
                Next
                If b Then
                    n = n + 1
                    For j = lb To ub
                        arr(n, j + 1) = .Cells(i, Column(j)).Text
                    Next
                    DoEvents
                End If
            Next
            If n = 0 Then
                arr = Null
            Else
                ReDim Preserve arr(1 To n, x) As String
            End If
        End If
    End With
    GetListFiltered = arr
End Function

Public Sub Zap()
    Dim i
    With Worksheets(mDataSheet)
        Application.ScreenUpdating = False
            For i = 1 To mRecCount
                .Rows(2).Delete
            Next
        Application.ScreenUpdating = True
    End With
    Reset
End Sub

Public Sub Repair()
    Dim i, j, n, b As Boolean
    With Range(mData)
        i = 1
        n = 0
        b = True
        Do While b
            If Len(.Cells(i).Text) = 0 Then
                b = False
                For j = 2 To mColCount
                    If Len(.Cells(i, j).Text) > 0 Then
                        b = True
                        n = n + 1
                        .Cells(i, 1).Text = Format(n, "00000000")
                        Exit For
                    End If
                Next
            End If
        Loop
    End With
    Reset
End Sub

Public Property Get RecNo() As Long
    RecNo = mLastRec
End Property

Public Property Let RecNo(ByVal vNewValue As Long)
    If 0 <= vNewValue And vNewValue <= mRecCount Then
        mLastRec = RecNo
    End If
End Property
