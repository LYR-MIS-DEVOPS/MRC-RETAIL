Attribute VB_Name = "modBilling"
Option Explicit
Private Qry As String
Private rs As New ADODB.Recordset

Public Sub sp_billing_new_list_bill_dtls(pBMAINID As String, objRs As ADODB.Recordset)
On Error GoTo er
    'sp_mrcretails.sp_billing_new_list_bill_dtls(
    '  pBRANCHID    Int
    '  ,pBMAINID    Int
    ')
    Qry = "Call sp_mrcretails.sp_billing_new_list_bill_dtls('" & SysBranchID & "','" & pBMAINID & "');"
    Set objRs = New ADODB.Recordset
    If objRs.State <> 0 Then objRs.Close
    objRs.Open Qry, LYRcon, adOpenStatic
Exit Sub
er:
    DoEvents
    MsgBox Err.Description, vbCritical, "Error->sp_billing_new_list_bill_dtls()"
End Sub

Public Sub sp_billing_new_list_year(objRs As ADODB.Recordset)
On Error GoTo er
    'sp_mrcretails.sp_billing_new_list_year(
    '   pBRANCHID    Int
    ')
    Qry = "Call sp_mrcretails.sp_billing_new_list_year('" & SysBranchID & "');"
    Set objRs = New ADODB.Recordset
    If objRs.State <> 0 Then objRs.Close
    objRs.Open Qry, LYRcon, adOpenStatic
Exit Sub
er:
    DoEvents
    MsgBox Err.Description, vbCritical, "Error->sp_billing_new_list_year()"
End Sub

Public Sub sp_billing_new_list_summary(pYEAR As String, objRs As ADODB.Recordset)
On Error GoTo er
    'sp_mrcretails.sp_billing_new_list_summary(
    '  pBRANCHID    Int
    '  ,pYEAR       Int
    ')
    Qry = "Call sp_mrcretails.sp_billing_new_list_summary('" & SysBranchID & "','" & pYEAR & "');"
    Set objRs = New ADODB.Recordset
    If objRs.State <> 0 Then objRs.Close
    objRs.Open Qry, LYRcon, adOpenStatic
Exit Sub
er:
    DoEvents
    MsgBox Err.Description, vbCritical, "Error->sp_billing_new_list_summary()"
End Sub

Public Sub sp_billing_new_process(pCUSTOMERID As String, pTRANSTYPE As String, pBILLMonth As String, pBILLYear As String _
    , pBILLINGSCHEME As String, pBILLINGSCHEDULE As String, pCUTOFF As String _
    , objRs As ADODB.Recordset)
On Error GoTo er
    'sp_mrcretails.sp_billing_new_process(
    '  pBRANCHID          Int
    '  ,pBRANCHCODE       VARCHAR(4)
    '  ,pCUSTOMERID       Int
    '  ,pTRANSTYPE        Varchar(50)
    '  ,pBILLMonth        Int
    '  ,pBILLYear         Int
    '  ,pBILLINGSCHEME    Varchar(60) -- M - > BI-MONTHLY, BM - > BI-MONTHLY, WK - > WEEKLY
    '  ,pBILLINGSCHEDULE  Varchar(60) -- ME - > MONTH END, 15 - > 15TH DAY, WK - > WEEKLY
    '  ,pCUTOFF           Date -- Date Cut-Off
    '  ,pENCODER          Int
    ')
    Dim Ldate As Date
    If IsDate(pCUTOFF) Then Ldate = CDate(pCUTOFF) Else Ldate = Now()
    
    pCUTOFF = Format(Ldate, "YYYY-m-dd")
    
    Qry = "Call sp_mrcretails.sp_billing_new_process('" & SysBranchID & "','" & sysBranchCode & "','" & pCUSTOMERID & "'" & _
                        ",'" & pTRANSTYPE & "','" & pBILLMonth & "','" & pBILLYear & "','" & pBILLINGSCHEME & "','" & pBILLINGSCHEDULE & "','" & pCUTOFF & "','" & UserID & "');"
    'InputBox Qry, Qry, Qry
    Set objRs = New ADODB.Recordset
    If objRs.State <> 0 Then objRs.Close
    objRs.Open Qry, LYRcon, adOpenStatic
Exit Sub
er:
    DoEvents
    MsgBox Err.Description, vbCritical, "Error->sp_billing_new_process()"
End Sub

Public Sub sp_billing_new_list_period(objRs As ADODB.Recordset)
On Error GoTo er
    'sp_mrcretails.sp_billing_new_list_period()
    Qry = "Call sp_mrcretails.sp_billing_new_list_period();"
    Set objRs = New ADODB.Recordset
    If objRs.State <> 0 Then objRs.Close
    objRs.Open Qry, LYRcon, adOpenStatic
Exit Sub
er:
    DoEvents
    MsgBox Err.Description, vbCritical, "Error->sp_billing_new_list_period()"
End Sub

Public Sub sp_billing_new_preview(pBILLINGNO As String, objRs As ADODB.Recordset)
On Error GoTo er
    'sp_mrcretails.sp_billing_new_preview(
    '    pBRANCHID     Int
    '    ,pBILLINGNO   Varchar(20)
    '    ,pENCODER     Int
    ')
    Qry = "Call sp_mrcretails.sp_billing_new_preview('" & SysBranchID & "','" & pBILLINGNO & "','" & UserID & "');"
    Set objRs = New ADODB.Recordset
    If objRs.State <> 0 Then objRs.Close
    objRs.Open Qry, LYRcon, adOpenStatic
Exit Sub
er:
    DoEvents
    MsgBox Err.Description, vbCritical, "Error->sp_billing_new_preview()"
End Sub
