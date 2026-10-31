Attribute VB_Name = "modTools"
Option Explicit

Public Sub loadFlxRecord(rs As ADODB.Recordset, flx As Object, Optional lbl As Object = Nothing)
    
    With rs
        '.Open
        If .RecordCount = 0 Then
            flx.Clear
            flx.Rows = 2
            Exit Sub
        End If
    Set flx.DataSource = rs
    If Not (lbl Is Nothing) Then lbl.Caption = lbl.Caption & ": " & .RecordCount
    '.Close
    End With

End Sub

Public Sub addToTextBox(rs As ADODB.Recordset, txt1 As Object)
    
    With rs
        If .RecordCount <> 0 Then
            .MoveFirst
                txt1 = .Fields(0)
        Else
            txt1 = ""
        End If
    End With
End Sub

Public Sub displayCRviewer(rs As ADODB.Recordset, crViewer As Object, _
    crDsr As Object, Optional pValue As String)
    crDsr.DiscardSavedData
    crDsr.Database.SetDataSource rs
    crViewer.ReportSource = crDsr
    crViewer.ViewReport
 
End Sub

Public Sub displayCboObject(ByRef objcbo As Object, rs As ADODB.Recordset, ByRef arrStr() As String, _
    fld1 As String, fld2 As String)
                            
    On Error GoTo LocalErr
    Dim arrIndex As Integer
    objcbo.Clear
    If rs.RecordCount <= 0 Then
        Exit Sub
    End If
    
    ReDim arrStr(rs.RecordCount, 1) As String
    arrIndex = 0
    rs.MoveFirst
    Do While Not rs.EOF
        arrStr(arrIndex, 0) = rs(fld1)
        arrStr(arrIndex, 1) = rs(fld2)
           
        Call objcbo.AddItem(rs(fld2), arrIndex)
        arrIndex = arrIndex + 1
        rs.MoveNext
    Loop
    
Exit Sub
LocalErr:
    MsgBox Err.Description, vbCritical, "Error ---> displayCboObject()"
End Sub

Public Sub addToListObjectCBO(ByRef objcbo As Object, rs As ADODB.Recordset, ByRef arrStr() As String, _
                            fld1 As String, fld2 As String)
                            
    On Error GoTo LocalErr
    Dim arrIndex As Integer
    objcbo.Clear
    If rs.RecordCount <= 0 Then
        Exit Sub
    End If
    
    ReDim arrStr(rs.RecordCount, 1) As String
    arrIndex = 0
    rs.MoveFirst
    Do While Not rs.EOF
        arrStr(arrIndex, 0) = rs(fld1)
        arrStr(arrIndex, 1) = rs(fld2)
           
        Call objcbo.AddItem(rs(fld2), arrIndex)
        arrIndex = arrIndex + 1
        rs.MoveNext
    Loop
Exit Sub
LocalErr:
    MsgBox Err.Description, vbCritical, "Error ---> displayCboObject()"
End Sub

Public Sub addCBOListBranch(ByRef objcbo As Object, rs As ADODB.Recordset)
On Error GoTo LocalErr

LocalErr:
    MsgBox Err.Description, vbCritical, "Error --->"
End Sub


' Added By SMC
' 2016-08-04
Public Sub addToListObject(ByRef objWList As Object, rs As ADODB.Recordset, Optional dontClear As Boolean = False _
    , Optional searchString As String = "", Optional RemoveEnterKey As Boolean = False, Optional FieldIndex As Integer = 0)
On Error GoTo er
    Dim Index As Integer
    Dim Tmp As String
    
    If Not dontClear Then
        objWList.Clear
        Index = 0
    Else
        Index = objWList.ListCount
    End If

    If rs.RecordCount <= 0 Then Exit Sub
    
    objWList.Tag = 0
    Index = objWList.ListCount
    
    rs.MoveFirst
    While Not rs.EOF
        Tmp = rs.Fields(FieldIndex).Value
        
        If Trim(searchString) <> "" Then
            If Trim(searchString) = Trim(Tmp) Then
                objWList.Tag = Index
            End If
        End If
        
        If RemoveEnterKey Then
            Tmp = strRemove13And10Char(Tmp)
        End If
        
        objWList.AddItem Tmp
        rs.MoveNext
        Index = Index + 1
    Wend
    If objWList.ListCount > 0 Then objWList.ListIndex = CInt(objWList.Tag)
Exit Sub
er:
    DoEvents
    MsgBox Err.Description, vbCritical, "Error ->addToListObject()"
End Sub

' Added By SMC
' 2016-08-04
Function strRemove13And10Char(ByVal str As String) As String
    Dim cntSpace As Long
    Dim counter As Long
    Dim Tmp As String
    
    str = Trim(str)
    strRemove13And10Char = ""
    cntSpace = 0
    For counter = 1 To Len(str)
        Tmp = Mid(str, counter, 1)
        
        If Asc(Tmp) = 10 Or Asc(Tmp) = 13 Then Tmp = " "
        
        strRemove13And10Char = strRemove13And10Char & Tmp
    Next counter
End Function


Public Sub flxFind(objFlx As Object, objText As Object, Caption As String, flxTTLCol As Integer, ByRef flxRowNo As Integer, ByRef flxColNo As Integer)
On Error GoTo er

Dim fRowIndex As Integer, fColIndex As Integer

If objText.Text = "" Then Exit Sub
With objFlx
    For fRowIndex = flxRowNo To .Rows - 1
        For fColIndex = flxColNo To flxTTLCol - 1
            If InStr(UCase(.TextMatrix(fRowIndex, fColIndex)), UCase(objText.Text)) > 0 Then
                .HighLight = flexHighlightNever
                .Refresh
                DoEvents
                .Col = 0
                .Row = fRowIndex
                .ColSel = flxTTLCol - 1
                .RowSel = fRowIndex
                .TopRow = fRowIndex
                flxRowNo = IIf(fColIndex = flxTTLCol - 1, fRowIndex + 1, fRowIndex)
                flxColNo = IIf(fColIndex = flxTTLCol - 1, 1, fColIndex + 1)
                'flxActRow = fRowIndex
                DoEvents
                .HighLight = flexHighlightAlways
                Exit Sub
            End If
        Next fColIndex
        flxColNo = 1
    Next fRowIndex
End With
MsgBox "REACH END OF ROW! ", vbInformation + vbOKOnly, Caption
flxRowNo = 1
flxColNo = 1
Exit Sub
er:
    MsgBox Err.Description, vbCritical, "Error->flxFind()"
End Sub

Public Function getFlxColsTTL(objFlx As Object, FlxCols As Integer) As Double
On Error GoTo er
    Dim i  As Integer
    Dim tmpSTR As String
    getFlxColsTTL = 0
    For i = 1 To objFlx.Rows - 1
        tmpSTR = objFlx.TextMatrix(i, FlxCols)
        tmpSTR = Trim(Replace(tmpSTR, ",", ""))
        If Trim(tmpSTR) <> "" Then
            getFlxColsTTL = getFlxColsTTL + Val(tmpSTR)
        End If
    Next i
Exit Function
er:
    MsgBox Err.Description, vbCritical, "Error->getFlxColsTTL()"
End Function

Public Function getFlxRowIndex(objFlx As Object, rowStart As Integer, search As String, colIndex As Integer, Optional EqualMatched As Boolean = True) As Integer
On Error GoTo er
    Dim RowIndex As Integer
    getFlxRowIndex = -1
    search = Trim(search)
    If rowStart <= 0 Then rowStart = 1
    For RowIndex = rowStart To objFlx.Rows - 1
        If EqualMatched Then
            If search = Trim(objFlx.TextMatrix(RowIndex, colIndex)) Then
                getFlxRowIndex = RowIndex
                Exit Function
            End If
        Else
            If InStr(Trim(objFlx.TextMatrix(RowIndex, colIndex)), search) Then
                getFlxRowIndex = RowIndex
                Exit Function
            End If
        End If
    Next
Exit Function
er:
    DoEvents
    MsgBox Err.Description, vbCritical, "Error-getFlxRowIndex()"
End Function

Public Function getListIndex(ByVal objWLst As Object, ByVal Value As String) As Integer
    On Error Resume Next
    Value = Trim(Value)
    getListIndex = -1
    Dim i As Integer
    For i = 0 To objWLst.ListCount - 1
        If Trim(objWLst.List(i)) = Value Then
            getListIndex = i
            Exit Function
        End If
    Next
End Function

Sub addToListObjectWithArray(ByRef objWList As Object, rs As ADODB.Recordset, ByRef arrStr() As String, fldDesc As String, maxCols As Integer)
On Error GoTo er
    Dim arrIndex As Integer
    Dim i As Integer
    
    objWList.Clear
    If rs.RecordCount <= 0 Then
        Exit Sub
    End If
 
    maxCols = maxCols - 1
    
    ReDim arrStr(rs.RecordCount, maxCols) As String
    
    arrIndex = 0
    rs.MoveFirst
    Do While Not rs.EOF
        For i = 0 To maxCols
            arrStr(arrIndex, i) = rs.Fields(i)
        Next i
        Call objWList.AddItem(rs(fldDesc), arrIndex)
        arrIndex = arrIndex + 1
        rs.MoveNext
    Loop
    
    If rs.State <> 0 Then rs.Close
    
Exit Sub
er:
    DoEvents
    MsgBox Err.Description, vbCritical, "Error ->addToListObjectWithArray()"
End Sub

Sub addToListObjectUsingArray(ByRef objWList As Object, arrStr() As String)
On Error GoTo er
    Dim arrIndex As Integer
    Dim i As Integer
    
    objWList.Clear
    
    For i = 0 To UBound(arrStr)
        Call objWList.AddItem(arrStr(i))
    Next i
    
Exit Sub
er:
    DoEvents
    MsgBox Err.Description, vbCritical, "Error ->addToListObjectUsingArray()"
End Sub

Sub addToListObjectUsingArrayWithCols(ByRef objWList As Object, ByRef arrStr() As String, arrSrc() As String, MaxCol As Integer, IndexDesc As Integer)
On Error GoTo er
    Dim TTLArrCount As Integer
    Dim arrIndex As Integer
    Dim RunningArrIndex As Integer
    Dim jCol As Integer
    
    objWList.Clear
    TTLArrCount = UBound(arrSrc)
    ReDim arrStr(TTLArrCount, MaxCol) As String
    RunningArrIndex = 0
    For arrIndex = 0 To TTLArrCount
        If Trim(arrSrc(arrIndex, IndexDesc)) <> "" Or arrIndex = 0 Then
            For jCol = 0 To MaxCol
                arrStr(RunningArrIndex, jCol) = arrSrc(arrIndex, jCol)
            Next jCol
            
            Call objWList.AddItem(arrSrc(arrIndex, IndexDesc), RunningArrIndex)
            RunningArrIndex = RunningArrIndex + 1
        End If
    Next arrIndex
Exit Sub
er:
    DoEvents
    MsgBox Err.Description, vbCritical, "Error ->addToListObjectUsingArrayWithCols()"
End Sub

Sub addToListObjectWithArrayGivenField(ByRef objWList As Object, rs As ADODB.Recordset, ByRef arrStr() As String, fldCode As String, fldDesc As String)
On Error GoTo er
    Dim arrIndex As Integer
    objWList.Clear
    If rs.RecordCount <= 0 Then
        Exit Sub
    End If
     
    ReDim arrStr(rs.RecordCount, 1) As String
    
    arrIndex = 0
    rs.MoveFirst
    Do While Not rs.EOF
        arrStr(arrIndex, 0) = rs(fldCode)
        arrStr(arrIndex, 1) = rs(fldDesc)
        
        Call objWList.AddItem(rs(fldDesc), arrIndex)
        arrIndex = arrIndex + 1
        rs.MoveNext
    Loop
    
    If rs.State <> 0 Then rs.Close
Exit Sub
er:
    DoEvents
    MsgBox Err.Description, vbCritical, "Error ->addToListObjectWithArrayGivenField()"
End Sub



Public Sub changeListIndex(objWList As Object, ByVal Value As String)
On Error Resume Next
    Dim i As Integer
    For i = 0 To objWList.ListCount - 1
        If UCase(Trim(objWList.List(i))) = UCase(Trim(Value)) Then
            objWList.ListIndex = i
            Exit Sub
        End If
    Next
End Sub

Public Sub fillArray(arrStr() As String, objRs As ADODB.Recordset)
On Error GoTo er
Dim iRow As Long
Dim jCol As Long
Dim ttlRow As Long

If objRs.State = 0 Then Exit Sub
If objRs.RecordCount <= 0 Then Exit Sub
iRow = 0
ttlRow = objRs.RecordCount - 1

ReDim arrStr(ttlRow, objRs.Fields.Count - 1) As String
objRs.MoveFirst
Do While Not objRs.EOF
    For jCol = 0 To objRs.Fields.Count - 1
        arrStr(iRow, jCol) = CStr(objRs.Fields(jCol).Value)
    Next jCol
    
    iRow = iRow + 1
    objRs.MoveNext
Loop

Exit Sub
er:
    DoEvents
    MsgBox Err.Description, vbCritical, "Error->fillArray()"
End Sub

Sub fillListObjectDiscountDesc(ByRef objWList As Object, arrSrc() As String, ByRef arrStr() As String, Active As Boolean)
On Error GoTo er
    Dim ttlData As Long
    Dim strTmp As String
    Dim varValue  As Variant
    Dim objRaw As New Collection
    Dim objRawData(2) As String
    
    Dim arrIndex As Long
    
    Dim iRow As Long
    Dim ttlRow As Long
    
    objWList.Clear
 
    ttlRow = UBound(arrSrc())
    ttlData = 0
    
    'Array Legend
    '0-> DISCOUNT_TITLEID
    '1-> DESCRIPTION
    '2-> ACTIVE
    For iRow = 0 To ttlRow
        strTmp = Trim(arrSrc(iRow, 2))
        If Active = True Then
            If Trim(strTmp) = "1" Then
                objRawData(1) = arrSrc(iRow, 0) ' Key
                objRawData(2) = arrSrc(iRow, 1) 'Data
                
                Call objRaw.Add(objRawData, arrSrc(iRow, 0))
                
                ttlData = ttlData + 1
            End If
        Else
            objRawData(1) = arrSrc(iRow, 0)  ' Key
            objRawData(2) = arrSrc(iRow, 1) 'Data
            
            Call objRaw.Add(objRawData, arrSrc(iRow, 0))
            
            ttlData = ttlData + 1
        End If
        
    Next iRow
    
  
    ReDim arrStr(ttlData) As String
    
    arrIndex = 0
    Call objWList.AddItem("", arrIndex)
    arrStr(arrIndex) = ""
 
    
    For Each varValue In objRaw
        arrIndex = arrIndex + 1
        arrStr(arrIndex) = varValue(1)

        Call objWList.AddItem(varValue(2), arrIndex)
    Next
Exit Sub
er:
    DoEvents
    MsgBox Err.Description, vbCritical, "Error ->fillListObjectDiscountType()"
End Sub

Public Function GetLastDayOfTheMonth(pMonthNo As Integer, pYEAR As Integer) As Integer
    GetLastDayOfTheMonth = -1
    Select Case pMonthNo
        Case 1, 3, 5, 7, 8, 10, 12
            GetLastDayOfTheMonth = 31
        Case 2
            If (pYEAR Mod 4) = 0 Then
                GetLastDayOfTheMonth = 29
            Else
                GetLastDayOfTheMonth = 28
            End If
        Case Else
            GetLastDayOfTheMonth = 30
    End Select
End Function
    

