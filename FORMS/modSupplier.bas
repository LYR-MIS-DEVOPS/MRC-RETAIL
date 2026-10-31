Attribute VB_Name = "modSupplier"
Option Explicit
Private Qry As String

Public Sub sp_supplier(objRs As ADODB.Recordset)
    On Error GoTo LocalErr
        
        Qry = "CALL sp_mrcretails.SP_SUPPLIER('" & SysBranchID & "');"
        If objRs.State <> 0 Then objRs.Close
        Set objRs = LYRcon.Execute(Qry)
        
Exit Sub
LocalErr:
    MsgBox Err.Description, vbCritical, "Error!"
End Sub

Public Sub modDefaultSupplierMainWarehouse(rs As ADODB.Recordset, txt1 As Object, txt2 As Object)
 
    With rs
        If .RecordCount = 1 Then
            .MoveFirst
            txt1.Text = .Fields(0)
            txt2.Text = .Fields(1)
        End If
    End With

End Sub


'Public Sub loadFlxRecord(rs As ADODB.Recordset, flx As Object, Optional lbl As Object = Nothing)
'
'    With rs
'        '.Open
'        If .RecordCount = 0 Then
'            flx.Clear
'            flx.Rows = 2
'            Exit Sub
'        End If
'    Set flx.DataSource = rs
'    If Not (lbl Is Nothing) Then lbl.Caption = lbl.Caption & ": " & .RecordCount
'    '.Close
'    End With
'
'End Sub
