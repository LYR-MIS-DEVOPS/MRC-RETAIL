Attribute VB_Name = "modAllocation"
Option Explicit
Private Qry As String
Private rs As New ADODB.Recordset
Public LoadOnProcess As Boolean

Public Function getMySQLDateFormat(PDate As String) As String
On Error Resume Next
    Dim Ldate As Date
    If IsDate(PDate) Then
        Ldate = CDate(PDate)
    Else
        Ldate = Now()
    End If
    
    getMySQLDateFormat = Format(Ldate, "YYYY-m-dd")
End Function


Public Sub sp_checks_update(pCHECKID As String, pCUSTOMERID As String, pCUSTOMER As String, pADDRESS As String, pPRNO As String, _
            pPRDATE As String, pCHKACCTNAME As String, pBANK As String, pCHECKNO As String, pAMOUNT As Double, _
            pSTATUS As String, pPAIDTO As String, pCHECKDATE As String, pCHECKTYPE As String, _
            objRs As ADODB.Recordset)
On Error GoTo er
    'sp_mrcretails.sp_checks_update(
    '  pBRANCHID       Int
    '  ,pBRANCHCODE    VARCHAR(4)
    '  ,pCUSTOMERID    INT
    '  ,pCHECKID       INT
    '  ,pCUSTOMER      Varchar(50)
    '  ,pADDRESS       Varchar(50)
    '  ,pPRNO          Varchar(10)
    '  ,pPRDATE        date
    '  ,pCHKACCTNAME   Varchar(50)
    '  ,pBANK          Varchar(70)
    '  ,pCHECKNO       Varchar(20)
    '  ,pAMOUNT        double
    '  ,pSTATUS        Varchar(250)
    '  ,pPAIDTO        Varchar(50)
    '  ,pCHECKDATE     date
    '  ,pCHECKTYPE     Varchar(20)
    '  ,pUPDATEDBY     Int
    ')
    Qry = "Call sp_mrcretails.sp_checks_update('" & SysBranchID & "','" & sysBranchCode & "','" & pCUSTOMERID & "','" & pCHECKID & "','" & pCUSTOMER & "','" & pADDRESS & "','" & pPRNO & "'" _
                & ",'" & pPRDATE & "','" & pCHKACCTNAME & "','" & pBANK & "','" & pCHECKNO & "','" & pAMOUNT & "','" & pSTATUS & "','" & pPAIDTO & "','" & pCHECKDATE & "'" _
                & ",'" & pCHECKTYPE & "'," & UserID & " );"
    'InputBox Qry, Qry, Qry
    Set objRs = New ADODB.Recordset
    If objRs.State <> 0 Then objRs.Close
    objRs.Open Qry, LYRcon, adOpenStatic
Exit Sub
er:
    DoEvents
    MsgBox Err.Description, vbCritical, "Error->sp_checks_create()"
End Sub
 

Public Sub sp_checks_create(pCUSTOMERID As String, pCUSTOMER As String, pADDRESS As String, pPRNO As String, _
            pPRDATE As String, pCHKACCTNAME As String, pBANK As String, pCHECKNO As String, pAMOUNT As Double, _
            pSTATUS As String, pPAIDTO As String, pCHECKDATE As String, pCHECKTYPE As String, pDate4BrkDwn As String, pDate4Dpsit As String, _
            objRs As ADODB.Recordset)
On Error GoTo er
    'sp_mrcretails.sp_checks_create(
    '  pBRANCHID       Int
    '  ,pBRANCHCODE    VARCHAR(4)
    '  ,pCUSTOMERID    INT
    '  ,pCUSTOMER      Varchar(50)
    '  ,pADDRESS       varchar(50)
    '  ,pPRNO          varchar(10)
    '  ,pPRDATE        date
    '  ,pCHKACCTNAME   varchar(50)
    '  ,pBANK          varchar(70)
    '  ,pCHECKNO       varchar(20)
    '  ,pAMOUNT        double
    '  ,pSTATUS        varchar(250)
    '  ,pPAIDTO        Varchar(50)
    '  ,pCHECKDATE     date
    '  ,pCHECKTYPE     varchar(20)
    '  ,pDate4BrkDwn   Varchar(20)
    '  ,pDate4Dpsit    Varchar(20)
    '  ,pUPDATEDBY     Int
    ')
    
    Qry = "Call sp_mrcretails.sp_checks_create('" & SysBranchID & "','" & sysBranchCode & "','" & pCUSTOMERID & "','" & pCUSTOMER & "','" & pADDRESS & "','" & pPRNO & "'" _
                & ",'" & pPRDATE & "','" & pCHKACCTNAME & "','" & pBANK & "','" & pCHECKNO & "','" & pAMOUNT & "','" & pSTATUS & "','" & pPAIDTO & "','" & pCHECKDATE & "'" _
                & ",'" & pCHECKTYPE & "','" & pDate4BrkDwn & "','" & pDate4Dpsit & "'," & UserID & " );"
    Set objRs = New ADODB.Recordset
    If objRs.State <> 0 Then objRs.Close
    objRs.Open Qry, LYRcon, adOpenStatic
Exit Sub
er:
    DoEvents
    MsgBox Err.Description, vbCritical, "Error->sp_checks_create()"
End Sub

Public Sub sp_check_allocation_list(pCHKSPRID As String, objRs As ADODB.Recordset)
On Error GoTo er
    'sp_mrcretails.sp_check_allocation_list(
    '  pBRANCHID       Int
    '  ,pBRANCHCODE    VARCHAR(4)
    '  ,pCHKSPRID      Int
    ')
    Qry = "Call sp_mrcretails.sp_check_allocation_list('" & SysBranchID & "','" & sysBranchCode & "','" & pCHKSPRID & "');"
    Set objRs = New ADODB.Recordset
    'InputBox Qry, Qry, Qry
    If objRs.State <> 0 Then objRs.Close
    objRs.Open Qry, LYRcon, adOpenStatic
Exit Sub
er:
    DoEvents
    MsgBox Err.Description, vbCritical, "Error->sp_check_allocation_list()"
End Sub

Public Sub sp_allocation_summary(pDateFrom As String, pDateTo As String, objRs As ADODB.Recordset)
On Error GoTo er
    'sp_mrcretails.sp_allocation_summary(
    '  pBRANCHID     Int
    '  ,pBRANCHCODE  VARCHAR(4)
    '  ,pDATEFROM    Varchar(60)
    '  ,pDATETO      Varchar(60)
    '  ,pEMPID       Int
    ')
    If Trim(pDateFrom) <> "" Then pDateFrom = getMySQLDateFormat(pDateFrom)
    
    If Trim(pDateTo) <> "" Then pDateTo = getMySQLDateFormat(pDateTo)
    
    If pDateFrom = pDateTo Then
        pDateTo = ""
    End If
    
    Qry = "Call sp_mrcretails.sp_allocation_summary('" & SysBranchID & "','" & sysBranchCode & "','" & pDateFrom & "','" & pDateTo & "','-1');"
    Set objRs = New ADODB.Recordset
    'InputBox Qry, Qry, Qry
    If objRs.State <> 0 Then objRs.Close
    objRs.Open Qry, LYRcon, adOpenStatic
Exit Sub
er:
    DoEvents
    MsgBox Err.Description, vbCritical, "Error->sp_allocation_summary()"
End Sub

Public Sub sp_allocation_auto(pBRANCHID As Integer, pCUSTOMERID As Integer, pBILLINGNO As String, _
            pALLOCAMOUNT As String, pALLOCATIONID As Integer, pENCODER As Integer, objRs As ADODB.Recordset)
On Error GoTo er
    'sp_mrcretails.sp_allocation_auto(
    '  pBRANCHID        Int
    '  ,pBRANCHCODE     VARCHAR(4)
    '  ,pCUSTOMERID     Int
    '  ,pBILLINGNO      Varchar(20)
    '  ,pALLOCAMOUNT    DECIMAL(24,2)
    '  ,pALLOCATIONID   Int
    '  ,pENCODER        Int
    ')
    Qry = "Call sp_mrcretails.sp_allocation_auto(" & pBRANCHID & ",'" & sysBranchCode & "'," & pCUSTOMERID & ",'" & _
            pBILLINGNO & "','" & pALLOCAMOUNT & "','" & pALLOCATIONID & "'," & pENCODER & ");"
    'InputBox Qry, Qry, Qry
    Set objRs = New ADODB.Recordset
    If objRs.State <> 0 Then objRs.Close
    objRs.Open Qry, LYRcon, adOpenStatic
    
Exit Sub
er:
    DoEvents
    MsgBox Err.Description, vbCritical, "Error->sp_allocation_auto()"
End Sub

Public Sub sp_allocation_list(pBRANCHID As Integer, pALLOCATIONID As String, objRs As ADODB.Recordset)
On Error GoTo er
    'PROCEDURE sp_mrcretails.sp_allocation_list(
    '  pBRANCHID           Int
    '  ,pBRANCHCODE        VARCHAR(4)
    '  ,pALLOCATIONID      Int
    ')
    
    Qry = "Call sp_mrcretails.sp_allocation_list( " & pBRANCHID & ",'" & sysBranchCode & "'," & pALLOCATIONID & ");"
    Set objRs = New ADODB.Recordset
    If objRs.State <> 0 Then objRs.Close
    objRs.Open Qry, LYRcon, adOpenStatic
    
Exit Sub
er:
    DoEvents
    MsgBox Err.Description, vbCritical, "Error->sp_allocation_list()"
End Sub

Public Sub sp_allocation_printform(pALLOCATIONID As Integer, objRs As ADODB.Recordset)
On Error GoTo er
    'sp_mrcretails.sp_allocation_printform(
    '  pBRANCHID        Int
    '  ,pBRANCHCODE     Varchar(4)
    '  ,pALLOCATIONID   Int
    '  ,pENCODER        Int
    ')
    Qry = "Call sp_mrcretails.sp_allocation_printform( " & SysBranchID & ",'" & sysBranchCode & "'," & pALLOCATIONID & "," & UserID & ");"
    Set objRs = New ADODB.Recordset
    If objRs.State <> 0 Then objRs.Close
    objRs.Open Qry, LYRcon, adOpenStatic
    
Exit Sub
er:
    DoEvents
    MsgBox Err.Description, vbCritical, "Error->sp_allocation_printform()"
End Sub

Public Sub sp_allocation_printform_cash_attachment(pALLOCATIONID As Integer, objRs As ADODB.Recordset)
On Error GoTo er
    'PROCEDURE sp_mrcretails.sp_allocation_printform_cash_attachment(
    '  pBRANCHID        Int
    '  ,pBRANCHCODE     Varchar(4)
    '  ,pALLOCATIONID   Int
    '  ,pENCODER        Int
    ')
    Qry = "Call sp_mrcretails.sp_allocation_printform_cash_attachment( " & SysBranchID & ",'" & sysBranchCode & "'," & pALLOCATIONID & "," & UserID & ");"
    Set objRs = New ADODB.Recordset
    If objRs.State <> 0 Then objRs.Close
    objRs.Open Qry, LYRcon, adOpenStatic
Exit Sub
er:
    DoEvents
    MsgBox Err.Description, vbCritical, "Error->sp_allocation_printform_cash_attachment()"
End Sub

Public Sub sp_allocation_printform_check_attachment(pALLOCATIONID As Integer, objRs As ADODB.Recordset)
On Error GoTo er
    ' sp_mrcretails.sp_allocation_printform_check_attachment(
    '   pBRANCHID       Int
    '  ,pBRANCHCODE     Varchar(4)
    '  ,pALLOCATIONID   Int
    '  ,pENCODER        Int
    ')
    Qry = "Call sp_mrcretails.sp_allocation_printform_check_attachment( " & SysBranchID & ",'" & sysBranchCode & "'," & pALLOCATIONID & "," & UserID & ");"
    Set objRs = New ADODB.Recordset
    If objRs.State <> 0 Then objRs.Close
    objRs.Open Qry, LYRcon, adOpenStatic
Exit Sub
er:
    DoEvents
    MsgBox Err.Description, vbCritical, "Error->sp_allocation_printform_check_attachment()"
End Sub

Public Sub sp_allocation_create(pBRANCHID As Integer, pACCOUNT_TYPE As String, pPAYTYPE As String _
           , pCUSTOMERID As Integer, pCHKSPRID As Integer, pCHECKNO As String, pTAXWHELDPERCENT As String _
           , pTAXWHELDAMOUNT As String, pENCODER As Integer, ByRef pALLOCATIONID As String _
           , ByRef pDATECREATED As String, pALLOCATIONSTATUS As String, ByRef pLevel As String _
           , ByRef pMessage As String)
           
On Error GoTo er
    'Call sp_mrcretails.sp_allocation_create(
    '  12               -- pBRANCHID          Int
    '  ,'MST'           -- pBRANCHCODE        VARCHAR(4)
    '  ,'JOB ORDER'     -- pACCOUNT_TYPE      Varchar(20)
    '  ,'CHECK PAYMENT' -- pPAYTYPE           Varchar(100) -- CA -> CASH PAYMENT, CH -> CHECK PAYMENT, AKLDGMT -> ACKNOWLEDGEMENT
    '  ,67              -- pCUSTOMERID        Int
    '  ,12479           -- pCHKSPRID          Int
    '  ,'1001455'       -- pCHECKNO           Varchar(20)
    '  ,1               -- pTAXWHELDPERCENT   Decimal(6,2)
    '  ,3.15            -- pTAXWHELDAMOUNT    Decimal(24,2)
    '  ,160             -- pENCODER           Int
    ')
    
    Qry = "Call sp_mrcretails.sp_allocation_create(" & pBRANCHID & ",'" & sysBranchCode & "','" & pACCOUNT_TYPE & _
                "','" & pPAYTYPE & "'," & pCUSTOMERID & "," & pCHKSPRID & ",'" & pCHECKNO & "','" & _
                pTAXWHELDPERCENT & "','" & pTAXWHELDAMOUNT & "'," & pENCODER & ");"
                
    'InputBox Qry, Qry, Qry
    Set rs = New ADODB.Recordset
    If rs.State <> 0 Then rs.Close
    rs.Open Qry, LYRcon, adOpenStatic
    
    If rs.State = 0 Then
        pLevel = "Error"
        pMessage = "Missing Output parameters!"
        
        pALLOCATIONID = ""
        pDATECREATED = "<AUTO>"
        pALLOCATIONSTATUS = ""
    ElseIf rs.RecordCount > 0 Then
        pLevel = rs!level
        pMessage = rs!message
        pALLOCATIONSTATUS = rs!ALLOCATIONSTATUS
        pALLOCATIONID = rs!AllocationID
        pDATECREATED = rs!DateCreated
    End If
    If rs.State <> 0 Then rs.Close
    
Exit Sub
er:
    DoEvents
    MsgBox Err.Description, vbCritical, "Error->sp_allocation_create()"
End Sub

Public Sub sp_allocation_update(pBRANCHID As Integer, pALLOCATIONID As Integer, _
                                pACCOUNT_TYPE As String, pPAYTYPE As String, pCUSTOMERID As Integer, pCHKSPRID As Integer, _
                                pCHECKNO As String, pTAXWHELDPERCENT As String, pTAXWHELDAMOUNT As String, pENCODER As Integer, _
                                ByRef pLevel As String, ByRef pMessage As String)
   
On Error GoTo er
    
    'PROCEDURE sp_mrcretails.sp_allocation_update(
    '  pBRANCHID           Int
    '  ,pBRANCHCODE        VARCHAR(4)
    '  ,pALLOCATIONID      Int
    '  ,pACCOUNT_TYPE      Varchar(20)
    '  ,pPAYTYPE           Varchar(100) -- CA -> CASH PAYMENT, CH -> CHECK PAYMENT, AKLDGMT -> ACKNOWLEDGEMENT
    '  ,pCUSTOMERID        Int
    '  ,pCHKSPRID          Int
    '  ,pCHECKNO           Varchar(20)
    '  ,pTAXWHELDPERCENT   Decimal(6,2)
    '  ,pTAXWHELDAMOUNT    Decimal(24,2)
    '  ,pENCODER           Int
    ')
    
    Qry = "Call sp_mrcretails.sp_allocation_update('" & pBRANCHID & "','" & sysBranchCode & "'," & pALLOCATIONID & ",'" & pACCOUNT_TYPE & _
                                                   "','" & pPAYTYPE & "'," & pCUSTOMERID & "," & pCHKSPRID & ",'" & pCHECKNO & _
                                                   "','" & pTAXWHELDPERCENT & "','" & pTAXWHELDAMOUNT & "'," & pENCODER & " );"
    
    'InputBox Qry, Qry, Qry
    Set rs = New ADODB.Recordset
    If rs.State <> 0 Then rs.Close
    rs.Open Qry, LYRcon, adOpenStatic
    If rs.State = 0 Then
        pLevel = "Error"
        pMessage = "Missing Output parameters!"
    Else
        pLevel = rs!level
        pMessage = rs!message
     End If
    If rs.State <> 0 Then rs.Close
Exit Sub
er:
DoEvents
    MsgBox Err.Description, vbCritical, "Error->sp_allocation_update()"
End Sub

Public Sub sp_onaccterms_list_ledger(pBRANCHID As Integer, pCUSTOMERID As Integer, pREFTYPE As String, _
            pREFNO As String, objRs As ADODB.Recordset)
On Error GoTo er
    'PROCEDURE sp_mrcretails.sp_onaccterms_list_ledger(
    '  pBRANCHID      Int
    '  ,pBRANCHCODE   VARCHAR(4)
    '  ,pCUSTOMERID   Int
    '  ,pREFTYPE      Varchar(20)
    '  ,pREFNO        VARCHAR(20)
    ')
    Qry = "Call sp_mrcretails.sp_onaccterms_list_ledger(" & pBRANCHID & ",'" & sysBranchCode & "'," & pCUSTOMERID & ",'" & pREFTYPE & "','" & pREFNO & "');"
    'InputBox Qry, Qry, Qry
    
    Set objRs = New ADODB.Recordset
    If objRs.State <> 0 Then objRs.Close
    objRs.Open Qry, LYRcon, adOpenStatic
    
Exit Sub
er:
    DoEvents
    MsgBox Err.Description, vbCritical, "Error->sp_onaccterms_list_ledger()"
End Sub

Public Sub sp_onaccterms_list_allocation(pBRANCHID As Integer, pCUSTOMERID As Integer, pACCOUNT_TYPE As String, _
        objRs As ADODB.Recordset)
On Error GoTo er
    'PROCEDURE sp_mrcretails.sp_onaccterms_list_allocation(
    '  pBRANCHID        Int
    '  ,pBRANCHCODE     VARCHAR(4)
    '  ,pCUSTOMERID     Int
    ')
    Qry = "Call sp_mrcretails.sp_onaccterms_list_allocation(" & pBRANCHID & ",'" & sysBranchCode & "'," & pCUSTOMERID & ",'" & pACCOUNT_TYPE & "');"
    Set objRs = New ADODB.Recordset
    If objRs.State <> 0 Then objRs.Close
    objRs.Open Qry, LYRcon, adOpenStatic
Exit Sub
er:
    DoEvents
    MsgBox Err.Description, vbCritical, "Error->sp_onaccterms_list_allocation()"
End Sub

Public Sub sp_allocation_list_curbilling(pBRANCHID As Integer, pCUSTOMERID As Integer, objRs As ADODB.Recordset)
On Error GoTo er
    'PROCEDURE sp_mrcretails.sp_allocation_list_curbilling(
    '  pBRANCHID        Int
    '  ,pBRANCHCODE     VARCHAR(4)
    '  ,pCUSTOMERID     Int
    ')
    Qry = "Call sp_mrcretails.sp_allocation_list_curbilling(" & pBRANCHID & ",'" & sysBranchCode & "'," & pCUSTOMERID & ");"
    Set objRs = New ADODB.Recordset
    If objRs.State <> 0 Then objRs.Close
    objRs.Open Qry, LYRcon, adOpenStatic
    
Exit Sub
er:
    DoEvents
    MsgBox Err.Description, vbCritical, "Error->sp_allocation_list_curbilling()"
End Sub

Public Sub sp_checks_list(pBRANCHID As Integer, pCUSTOMERID As Integer, objRs As ADODB.Recordset)
On Error GoTo er
    'PROCEDURE sp_mrcretails.sp_checks_list(
    '  pBRANCHID       Int
    '  ,pBRANCHCODE    VARCHAR(4)
    '  ,pCUSTOMERID    Int
    ')
    
    Qry = "Call sp_mrcretails.sp_checks_list( " & pBRANCHID & ",'" & sysBranchCode & "'," & pCUSTOMERID & ");"
    Set objRs = New ADODB.Recordset
    If objRs.State <> 0 Then objRs.Close
    objRs.Open Qry, LYRcon, adOpenStatic
    
Exit Sub
er:
    DoEvents
    MsgBox Err.Description, vbCritical, "Error->sp_checks_list()"
End Sub

Public Sub sp_customer_list(pBRANCHID As Integer, objRs As ADODB.Recordset)
On Error GoTo er
    'PROCEDURE sp_mrcretails.sp_customer_list(
    '  pBRANCHID           Int
    '  ,pBRANCHCODE        VARCHAR(4)
    ')
    
    Qry = "Call sp_mrcretails.sp_customer_list(" & pBRANCHID & ",'" & sysBranchCode & "' );"
    Set objRs = New ADODB.Recordset
    If objRs.State <> 0 Then objRs.Close
    objRs.Open Qry, LYRcon, adOpenStatic
    
Exit Sub
er:
    DoEvents
    MsgBox Err.Description, vbCritical, "Error->sp_customer_list()"
End Sub

Public Sub sp_cash_list(pBRANCHID As Integer, pCUSTOMERID As Integer, pALLOCATIONID As Integer, objRs As ADODB.Recordset)
On Error GoTo er
    'PROCEDURE sp_mrcretails.sp_cash_list(
    '  pBRANCHID           Int
    '  ,pBRANCHCODE        VARCHAR(4)
    '  ,pCUSTOMERID        Int
    '  ,pALLOCATIONID    Int
    ')
    Qry = "Call sp_mrcretails.sp_cash_list( " & pBRANCHID & ",'" & sysBranchCode & "'," & pCUSTOMERID & ",'" & pALLOCATIONID & "');"
    Set objRs = New ADODB.Recordset
    If objRs.State <> 0 Then objRs.Close
    objRs.Open Qry, LYRcon, adOpenStatic
Exit Sub
er:
    DoEvents
    MsgBox Err.Description, vbCritical, "Error->sp_cash_list()"
End Sub

Public Sub sp_customer_list_active_allocation(pCUSTOMERID As String, objRs As ADODB.Recordset)
On Error GoTo er
    'sp_mrcretails.sp_customer_list_active_allocation(
    '  pBRANCHID int
    '  ,pBRANCHCODE varchar(4)
    '  ,pCUSTOMERID  Int
    ')
    Qry = "Call sp_mrcretails.sp_customer_list_active_allocation('" & SysBranchID & "','" & sysBranchCode & "','" & pCUSTOMERID & "');"
    Set objRs = New ADODB.Recordset
    If objRs.State <> 0 Then objRs.Close
    objRs.Open Qry, LYRcon, adOpenStatic
Exit Sub
er:
    DoEvents
    MsgBox Err.Description, vbCritical, "Error->sp_customer_list_active_allocation()"
End Sub

Public Sub sp_payments_list_allocation(pALLOCATIONID As String, objRs As ADODB.Recordset)
On Error GoTo er
    'sp_mrcretails.sp_payments_list_allocation(
    '   pBRANCHID      Int
    '  ,pBRANCHCODE    Varchar(4)
    '  ,pALLOCATIONID  Varchar(20)
    ')
    Qry = "Call sp_mrcretails.sp_payments_list_allocation('" & SysBranchID & "','" & sysBranchCode & "','" & pALLOCATIONID & "');"
    Set objRs = New ADODB.Recordset
    If objRs.State <> 0 Then objRs.Close
    objRs.Open Qry, LYRcon, adOpenStatic
Exit Sub
er:
    DoEvents
    MsgBox Err.Description, vbCritical, "Error->sp_payments_list_allocation()"
End Sub

Public Sub sp_payments_create(pALLOCATIONID As String, pCUSTOMERID As String, pORNO As String, pCRNO As String _
        , pREFDATE As String, pTRANSPTYPE As String, pACCT_CODE _
        , objRs As ADODB.Recordset)
On Error GoTo er
    'sp_mrcretails.sp_payments_create(
    '  pBRANCHID        Int
    '  ,pBRANCHCODE     Varchar(4)
    '  ,pALLOCATIONID   Varchar(20)
    '  ,pCUSTOMERID     Varchar(14)
    '  ,pORNO           Varchar(20)
    '  ,pCRNO           Varchar(20)
    '  ,pREFDATE        Date
    '  ,pTRANSPTYPE     Varchar(200) -- CREDIT CARD PAYMENT,INTER-BRANCH PAYMENT,ON-LINE PAYMENT
    '  ,pENCODER        Int
    ')
    Qry = "Call sp_mrcretails.sp_payments_create('" & SysBranchID & "','" & sysBranchCode & "','" & pALLOCATIONID & "'," & pCUSTOMERID & ",'" & pORNO & "','" & pCRNO & "'" & _
          " ,'" & pREFDATE & "','" & pTRANSPTYPE & "'," & UserID & ",'" & pACCT_CODE & "');"
    
    'InputBox Qry, Qry, Qry
    Set objRs = New ADODB.Recordset
    If objRs.State <> 0 Then objRs.Close
    objRs.Open Qry, LYRcon, adOpenStatic
    
Exit Sub
er:
    DoEvents
    MsgBox Err.Description, vbCritical, "Error->sp_payments_create()"
End Sub

Public Sub sp_payments_cancel(pPAYMENTID As String, pREMARKS As String, objRs As ADODB.Recordset)
On Error GoTo er
    'sp_mrcretails.sp_payments_cancel(
    '  pBRANCHID     Int
    '  ,pBRANCHCODE  Varchar(4)
    '  ,pPAYMENTID   Int
    '  ,pREMARKS     Varchar(500)
    '  ,pEMPID       Int
    ')
    Qry = "Call sp_mrcretails.sp_payments_cancel('" & SysBranchID & "','" & sysBranchCode & "','" & pPAYMENTID & "','" & pREMARKS & "'," & UserID & ");"
    
    'InputBox Qry, Qry, Qry
    Set objRs = New ADODB.Recordset
    If objRs.State <> 0 Then objRs.Close
    objRs.Open Qry, LYRcon, adOpenStatic
Exit Sub
er:
    DoEvents
    MsgBox Err.Description, vbCritical, "Error->sp_payments_cancel()"
End Sub




