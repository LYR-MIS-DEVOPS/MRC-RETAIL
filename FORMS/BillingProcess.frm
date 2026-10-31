VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form BillingProcess 
   Caption         =   "Billing"
   ClientHeight    =   9150
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   14445
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9150
   ScaleWidth      =   14445
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab tabBill 
      Height          =   7815
      Left            =   120
      TabIndex        =   18
      Tag             =   "XY"
      Top             =   1320
      Width           =   14175
      _ExtentX        =   25003
      _ExtentY        =   13785
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "SUMMARY"
      TabPicture(0)   =   "BillingProcess.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "flxBill"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "PRINT PREVIEW"
      TabPicture(1)   =   "BillingProcess.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "CRViewer1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin CRVIEWERLibCtl.CRViewer CRViewer1 
         Height          =   7335
         Left            =   60
         TabIndex        =   20
         Tag             =   "XY"
         Top             =   360
         Width           =   14055
         DisplayGroupTree=   -1  'True
         DisplayToolbar  =   -1  'True
         EnableGroupTree =   -1  'True
         EnableNavigationControls=   -1  'True
         EnableStopButton=   -1  'True
         EnablePrintButton=   -1  'True
         EnableZoomControl=   -1  'True
         EnableCloseButton=   -1  'True
         EnableProgressControl=   -1  'True
         EnableSearchControl=   -1  'True
         EnableRefreshButton=   -1  'True
         EnableDrillDown =   -1  'True
         EnableAnimationControl=   -1  'True
         EnableSelectExpertButton=   0   'False
         EnableToolbar   =   -1  'True
         DisplayBorder   =   -1  'True
         DisplayTabs     =   -1  'True
         DisplayBackgroundEdge=   -1  'True
         SelectionFormula=   ""
         EnablePopupMenu =   -1  'True
         EnableExportButton=   0   'False
         EnableSearchExpertButton=   0   'False
         EnableHelpButton=   0   'False
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxBill 
         Height          =   7335
         Left            =   -74940
         TabIndex        =   19
         Tag             =   "XYF"
         Top             =   360
         Width           =   14055
         _ExtentX        =   24791
         _ExtentY        =   12938
         _Version        =   393216
         RowHeightMin    =   285
         BackColorBkg    =   16777215
         SelectionMode   =   1
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.TextBox txtDueDate 
      Height          =   285
      Left            =   4560
      TabIndex        =   17
      Top             =   120
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txtSQL 
      Height          =   285
      Left            =   8520
      TabIndex        =   16
      Top             =   120
      Visible         =   0   'False
      Width           =   5775
   End
   Begin MSComctlLib.ProgressBar Pbar 
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Tag             =   "XYF"
      Top             =   9150
      Visible         =   0   'False
      Width           =   13575
      _ExtentX        =   23945
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   6000
      Top             =   0
   End
   Begin VB.TextBox txtYr 
      Height          =   285
      Left            =   7800
      TabIndex        =   11
      Top             =   120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtMo 
      Height          =   285
      Left            =   6960
      TabIndex        =   10
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   6480
      Top             =   0
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Tag             =   "XY"
      Top             =   720
      Width           =   14175
      Begin VB.ComboBox cboBilling 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1680
         TabIndex        =   9
         Tag             =   "XYF"
         Top             =   172
         Width           =   3135
      End
      Begin VB.TextBox txtFind 
         Height          =   285
         Left            =   12120
         TabIndex        =   5
         Tag             =   "XYF"
         Top             =   195
         Width           =   1935
      End
      Begin lvButton.lvButtons_H cmdFind 
         Height          =   375
         Left            =   10800
         TabIndex        =   6
         Tag             =   "XY"
         Top             =   150
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         Caption         =   "Find Next"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         ImgAlign        =   1
         Image           =   "BillingProcess.frx":0038
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdNxtBill 
         Height          =   375
         Left            =   4920
         TabIndex        =   15
         Tag             =   "XY"
         Top             =   150
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         Caption         =   "&New Billing"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         ImgAlign        =   1
         Image           =   "BillingProcess.frx":05D2
         cBack           =   -2147483633
      End
      Begin VB.Label Label1 
         Caption         =   "Billing Period:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   210
         Width           =   1575
      End
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   3
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   661
      Caption         =   "E&xit Report"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      ImgAlign        =   1
      Image           =   "BillingProcess.frx":0B6C
      cBack           =   -2147483633
   End
   Begin MSFlexGridLib.MSFlexGrid flxDtls 
      Height          =   1095
      Left            =   120
      TabIndex        =   13
      Top             =   9600
      Visible         =   0   'False
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   1931
      _Version        =   393216
      Rows            =   1
      Cols            =   6
      FixedCols       =   0
      FormatString    =   "CustomerID | RefNo | RefDate  | BillType  | AmtDM  | AmtCM  "
   End
   Begin VB.Label lblTotal 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   13680
      TabIndex        =   14
      Tag             =   "XY"
      Top             =   9150
      Width           =   615
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   360
      Top             =   8640
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   14685
      m_MDI_Height    =   9720
   End
   Begin VB.Label lblHeading 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "MONTHLY BILLING"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   2
      Left            =   9720
      TabIndex        =   2
      Top             =   450
      Width           =   4500
   End
   Begin VB.Label lblHeading 
      BackStyle       =   0  'Transparent
      Caption         =   "MOTO>R REPUBLIC CORP."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   450
      Width           =   2615
   End
   Begin VB.Label lblHeading 
      BackStyle       =   0  'Transparent
      Caption         =   "MOTO>R REPUBLIC CORP."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   1
      Left            =   150
      TabIndex        =   1
      Top             =   475
      Width           =   2615
   End
   Begin VB.Label lblHeading 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "MONTHLY BILLING"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   3
      Left            =   9735
      TabIndex        =   3
      Top             =   480
      Width           =   4500
   End
   Begin VB.Menu mnuINV 
      Caption         =   "Customer Billing"
      Visible         =   0   'False
      Begin VB.Menu mnuProcess 
         Caption         =   "Process Billing"
      End
      Begin VB.Menu mnuPreview 
         Caption         =   "Print Preview"
      End
   End
End
Attribute VB_Name = "BillingProcess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsCustBill As New ADODB.Recordset
Dim rsPrn As New ADODB.Recordset
Dim ActRow As Integer, fRow As Integer, fCol As Integer

Private Sub cboBilling_Click()
    Dim MoYr() As String
    
    If cboBilling.ListIndex = -1 Or cboBilling.Text = "" Then Exit Sub
    MoYr = Split(cboBilling.Text, "-")
    txtYr.Text = Trim(MoYr(1))
    txtMo.Text = Month(DateValue(UCase$(MoYr(0)) & "1, " & txtYr.Text))
    txtDueDate.Text = DateAdd("m", 1, DateSerial(txtYr.Text, txtMo.Text, 15))
    Call DisplayBilling
End Sub

Private Sub cmdFind_Click()
    Dim i As Integer, j As Integer
        
    If txtFind.Text = "" Then Exit Sub
    With flxBill
        For i = fRow To .Rows - 1
            For j = fCol To 8
                If InStr(.TextMatrix(i, j), txtFind.Text) > 0 Then
                    .Col = 2: .Row = i
                    .ColSel = 8:  .RowSel = i
                    .TopRow = i
                    fRow = IIf(j = 8, i + 1, i)
                    fCol = IIf(j = 8, 2, j + 1)
                    Exit Sub
                End If
            Next j
            fCol = 2
        Next i
    End With
    MsgBox "No match found!", vbInformation + vbOKOnly, "Find"
    fRow = 1
    fCol = 2
End Sub

Private Sub cmdNxtbill_Click()
    Dim PrvMo As Integer, PrvYr As Integer, Ans As Integer
    Dim Qry As String, MyAcc As String
    Dim CutOffDate As Date
    
    'Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE (((USERACCESS.EMPID)=" & UserID & ") AND ((USERACCESS.MENUINDX)=65));"
    'mySQL
    Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE USERACCESS.EMPID=" & UserID & " AND USERACCESS.MENUINDX=65;"
    MyAcc = GetFieldValue(Qry)
    If MyAcc <> "FULL ACCESS" Then
        MsgBox "You don't have full access to this transaction." & vbCrLf & "Please contact your System Administrator", vbExclamation + vbOKOnly, "Access Denied"
        Exit Sub
    End If

    'Validate if Next Period wla pa nag lapse
    Qry = "SELECT BILLING.BILLYR, BILLING.BILLMO From BILLING GROUP BY BILLING.BILLYR, BILLING.BILLMO ORDER BY BILLING.BILLYR DESC , BILLING.BILLMO DESC;"
    PrvYr = Val(GetFieldValue(Qry))
    Qry = "SELECT BILLING.BILLMO, BILLING.BILLYR From BILLING GROUP BY BILLING.BILLYR, BILLING.BILLMO ORDER BY BILLING.BILLYR DESC , BILLING.BILLMO DESC;"
    PrvMo = Val(GetFieldValue(Qry))
    If PrvYr = 0 Then
        Timer2.Enabled = True
        Exit Sub
    End If
    CutOffDate = DateAdd("m", 2, DateSerial(PrvYr, PrvMo, 1))
    CutOffDate = DateAdd("d", -1, CutOffDate)
    If Date < CutOffDate Then
        MsgBox "Billing Cut-Off will due on " & CutOffDate & vbCrLf & "Billing can be process after the cutoff", vbInformation + vbOKOnly, "Billing"
        Exit Sub
    End If
    txtMo.Text = Month(CutOffDate)
    txtYr.Text = Year(CutOffDate)
    txtDueDate.Text = DateAdd("m", 1, DateSerial(txtYr.Text, txtMo.Text, 15))
    Ans = MsgBox("Do you want to Process Billing for the period " & MonthName(txtMo.Text) & " - " & txtYr.Text & "?", vbQuestion + vbYesNo, "Process Billing")
    If Ans = vbNo Then Exit Sub
    'Display all SP in flxBill
    cboBilling.AddItem MonthName(txtMo.Text) & " - " & txtYr.Text
    cboBilling.Text = MonthName(txtMo.Text) & " - " & txtYr.Text
    'Qry = "SELECT CUSTOMER.CUSTOMERID, CUSTOMER.CUSTOMERNAME From CUSTOMER WHERE (((CUSTOMER.CTYPE)<>'INDIVIDUAL')) ORDER BY CUSTOMER.CUSTOMERNAME;"
    'mySQL
    Qry = "SELECT CUSTOMER.CUSTOMERID, CUSTOMER.CUSTOMERNAME From CUSTOMER WHERE CUSTOMER.CTYPE<>'INDIVIDUAL' ORDER BY CUSTOMER.CUSTOMERNAME;"
    If rsCustBill.State = adStateOpen Then rsCustBill.Close
    rsCustBill.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    If rsCustBill.RecordCount = 0 Then
        flxBill.Clear
        flxBill.Rows = 2
        flxBill.Cols = 10
        Call RXResizer1_ResizeComplete
        Exit Sub
    End If
    Set flxBill.DataSource = rsCustBill
    ActRow = 1
    Call RXResizer1_ResizeComplete
    Call ProcessBilling
                           
End Sub

Private Sub cmdTrans_Click(Index As Integer)
    Unload Me
    'MotoRep.tbMain.Visible = True
    'MotoRep.imgBG.Visible = True
            MotoRep.toolBarButton.Visible = True
End Sub

Private Sub flxBill_EnterCell()
    ActRow = flxBill.Row
End Sub

Private Sub flxBill_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
    If Button = vbRightButton Then
        Call PopupMenu(mnuINV)
    End If
End Sub

Private Sub Form_Load()
    Dim rsBill As New ADODB.Recordset
    Dim Qry As String
    
    'Qry = "SELECT BILLING.BILLMO, BILLING.BILLYR, Count(BILLING.CUSTOMERID) AS CountOfCUSTOMERID From BILLING " & _
          "GROUP BY BILLING.BILLMO, BILLING.BILLYR, BILLING.BILLMO ORDER BY BILLING.BILLYR, BILLING.BILLMO;"
    'mySQL
    Qry = "SELECT BILLING.BILLMO, BILLING.BILLYR, Count(BILLING.CUSTOMERID) AS CountOfCUSTOMERID From BILLING " & _
          "GROUP BY BILLING.BILLMO, BILLING.BILLYR, BILLING.BILLMO ORDER BY BILLING.BILLYR, BILLING.BILLMO;"
    With rsBill
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                cboBilling.AddItem MonthName(!BILLMO) & " - " & !BILLYR
                .MoveNext
            Loop
        End If
        .Close
    End With
    tabBill.Tab = 0
    If cboBilling.ListCount > 0 Then
        Timer1.Enabled = True
    Else
        Timer2.Enabled = True
    End If
    MotoRep.sBar.Panels(6).Text = "Process Customer Billing"
    'MotoRep.tbMain.Visible = False
    MotoRep.imgBG.Visible = False
    MotoRep.toolBarButton.Visible = False
    ActRow = 1
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    MotoRep.sBar.Panels(6).Text = ""
    If rsCustBill.State = adStateOpen Then rsCustBill.Close
End Sub

Private Sub mnuPreview_Click()
    Dim crSOA As New rptSOA
    Dim Qry As String, Dfrom As String, Dto As String
    
    
    Dfrom = DateSerial(txtYr.Text, txtMo.Text, 1)
    Dto = DateAdd("d", -1, DateAdd("m", 1, Dfrom))
    'Display Print Preview
    Qry = "SELECT CUSTOMER.CUSTOMERNAME, CUSTOMER.ADDRESS, BILLING.DUEDATE, BILLING.PREVBAL, BILLING.CURRENTBILL, BILLING.PAYMENTS, BILLING.DMCM, BILLING.BALANCE, BILLINGDTLS.BILLTYPE, BILLINGDTLS.REFDATE, BILLINGDTLS.REFNO, BILLINGDTLS.AMOUNTDM, BILLINGDTLS.AMOUNTCM, CUSTOMER.CONTACTPERSON, CUSTOMER.CTYPE " & _
          "FROM CUSTOMER INNER JOIN (BILLINGDTLS INNER JOIN BILLING ON BILLINGDTLS.BILLINGID = BILLING.BILLINGID) ON CUSTOMER.CUSTOMERID = BILLING.CUSTOMERID Where (((BILLING.BILLMO) = " & txtMo.Text & ") And ((BILLING.BILLYR) = " & txtYr.Text & ") And ((BILLING.CUSTOMERID) = " & flxBill.TextMatrix(ActRow, 1) & ")) ORDER BY BILLINGDTLS.REFDATE, BILLINGDTLS.REFNO;"
    
    If rsPrn.State = adStateOpen Then rsPrn.Close
    rsPrn.Open Qry, LYRcon, adOpenDynamic, adLockOptimistic
    crSOA.DiscardSavedData
    crSOA.Database.SetDataSource rsPrn
    crSOA.ParameterFields(1).ClearCurrentValueAndRange
    crSOA.ParameterFields(1).AddCurrentValue Dfrom
    crSOA.ParameterFields(2).ClearCurrentValueAndRange
    crSOA.ParameterFields(2).AddCurrentValue Dto
    crSOA.ParameterFields(3).ClearCurrentValueAndRange
    crSOA.ParameterFields(3).AddCurrentValue GetSettings("CASHIER")
    crSOA.ParameterFields(4).ClearCurrentValueAndRange
    crSOA.ParameterFields(4).AddCurrentValue GetSettings("MANAGER")
    crSOA.ParameterFields(5).ClearCurrentValueAndRange
    crSOA.ParameterFields(5).AddCurrentValue GetSettings("OIC")
    crSOA.ParameterFields(6).ClearCurrentValueAndRange
    crSOA.ParameterFields(6).AddCurrentValue SysBranch
    crSOA.ParameterFields(7).ClearCurrentValueAndRange
    crSOA.ParameterFields(7).AddCurrentValue SysAddress
    crSOA.ParameterFields(8).ClearCurrentValueAndRange
    crSOA.ParameterFields(8).AddCurrentValue SysTel
    
    CRViewer1.ReportSource = crSOA
    CRViewer1.ViewReport
    CRViewer1.DisplayGroupTree = False
    CRViewer1.EnableExportButton = True
    tabBill.Tab = 1
End Sub

Private Sub mnuProcess_Click()
    Dim rsAccts As New ADODB.Recordset
    Dim dStart As Date, dEnd As Date, Ans As Integer, BillID As Integer
    Dim Qry As String, rRet As String, MyType As String, MyAcc As String
    Dim j As Integer, TtlBill As Currency, PrvBal As Currency, TtlPay As Currency
    Dim SvcAmt As Currency, AdjAmt As Currency, AmtRet As Currency, InvAmt As Currency, TtlDmCm As Currency
    
    'Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE (((USERACCESS.EMPID)=" & UserID & ") AND ((USERACCESS.MENUINDX)=65));"
    'mySQL
    Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE USERACCESS.EMPID=" & UserID & " AND USERACCESS.MENUINDX=65;"
    
    MyAcc = GetFieldValue(Qry)
    If MyAcc <> "FULL ACCESS" Then
        MsgBox "You don't have full access to this transaction." & vbCrLf & "Please contact your System Administrator", vbExclamation + vbOKOnly, "Access Denied"
        Exit Sub
    End If
    
    dStart = DateSerial(txtYr.Text, Val(txtMo.Text), 1)
    dEnd = DateAdd("d", -1, DateAdd("m", 1, dStart))

    'Validate Process if Billing already exist
    If Not IsDate(txtDueDate.Text) Then
        MsgBox "Invalid Due Date!", vbExclamation + vbOKOnly, "Validation Error"
        Exit Sub
    End If
    'Qry = "SELECT BILLING.BILLINGID From BILLING WHERE (((BILLING.CUSTOMERID)=" & flxBill.TextMatrix(ActRow, 1) & ") AND ((BILLING.BILLMO)=" & txtMo.Text & ") AND ((BILLING.BILLYR)=" & txtYr.Text & "));"
    'mySQL
    Qry = "SELECT BILLING.BILLINGID From BILLING WHERE (((BILLING.CUSTOMERID)=" & flxBill.TextMatrix(ActRow, 1) & ") AND ((BILLING.BILLMO)=" & txtMo.Text & ") AND ((BILLING.BILLYR)=" & txtYr.Text & "));"
  
    BillID = Val(GetFieldValue(Qry))
    If BillID > 0 Then
        Ans = MsgBox("Warning!!! Billing Already exist!" & vbCrLf & "Do you want to Re-Process the Billing?", vbQuestion + vbYesNo, "Update Billing")
        If Ans = vbNo Then Exit Sub
    End If
    'Lock objects
    Frame1.Enabled = False
    cmdTrans(3).Enabled = False
    flxBill.Enabled = False
    
    'Get On Account for Spare Parts
    TtlBill = 0
    TtlDmCm = 0
    TtlPay = 0
    flxDtls.Rows = 1                ' AND ((ONACCTERMS.STATUS)='OPEN') has been removed to qry 9/26/08
    'Qry = "SELECT ONACCTERMS.CUSTOMERID, ONACCTERMS.SPINVNO, ONACCTERMS.TRANSDATE, ONACCTERMS.BALANCE, ONACCTERMS.ACCTAMOUNT From ONACCTERMS " & _
          "WHERE (((ONACCTERMS.CUSTOMERID)=" & flxBill.TextMatrix(ActRow, 1) & ") AND (Not (ONACCTERMS.SPINVNO) Is Null) AND ((ONACCTERMS.TRANSDATE)>=#" & dStart & "# And (ONACCTERMS.TRANSDATE)<=#" & dEnd & "#) AND ((ONACCTERMS.STATUS)<>'CANCELLED')) ORDER BY ONACCTERMS.TRANSDATE;"
    'mySQL
    Qry = "SELECT ONACCTERMS.CUSTOMERID, ONACCTERMS.SPINVNO, ONACCTERMS.TRANSDATE, ONACCTERMS.BALANCE, ONACCTERMS.ACCTAMOUNT From ONACCTERMS " & _
          "WHERE ONACCTERMS.CUSTOMERID=" & flxBill.TextMatrix(ActRow, 1) & " AND Not (ONACCTERMS.SPINVNO) Is Null AND ONACCTERMS.TRANSDATE>='" & Format(dStart, "yyyy-mm-dd") & "' And ONACCTERMS.TRANSDATE<='" & Format(dEnd, "yyyy-mm-dd") & "' AND ONACCTERMS.STATUS<>'CANCELLED' ORDER BY ONACCTERMS.TRANSDATE;"

    With rsAccts
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                'rRet = "SELECT Sum(SRETURNSP.TOTALAMT) AS SumOfTOTALAMT From SRETURNSP " & _
                       "WHERE (((SRETURNSP.INVOICE)='" & !SPINVNO & "') AND ((SRETURNSP.REMARKS) Not Like 'CANCELLED'));"
                'mySQL
                rRet = "SELECT Sum(SRETURNSP.TOTALAMT) AS SumOfTOTALAMT From SRETURNSP " & _
                       "WHERE SRETURNSP.INVOICE='" & !SPINVNO & "' AND (SRETURNSP.REMARKS) Not Like 'CANCELLED';"
                AmtRet = Val(GetFieldValue(rRet))
                AmtAdj = GetDMCMamt(!SPINVNO, "INVOICE")
                InvAmt = !AcctAmount    '!Balance + AmtRet + AmtAdj
                flxDtls.AddItem !CUSTOMERID & vbTab & "INV. NO. " & !SPINVNO & vbTab & !TRANSDATE & vbTab & "SPARE PARTS" & vbTab & InvAmt
                TtlBill = TtlBill + InvAmt
                .MoveNext
            Loop
        End If
        .Close
    End With
    
    'Display Sales Return
    'Qry = "SELECT SPINVOICE.CUSTOMERID, SRETURNSP.RETNO, SRETURNSP.RETDATE, SRETURNSP.TOTALAMT FROM SRETURNSP INNER JOIN SPINVOICE ON SRETURNSP.INVOICE = SPINVOICE.SPINVNO " & _
          "WHERE (((SPINVOICE.CUSTOMERID)=" & flxBill.TextMatrix(ActRow, 1) & ") AND ((SRETURNSP.RETDATE)>=#" & dStart & "# And (SRETURNSP.RETDATE)<=#" & dEnd & "#) AND ((SRETURNSP.REMARKS)<>'CANCELLED'));"
    Qry = "SELECT SPINVOICE.CUSTOMERID, SRETURNSP.RETNO, SRETURNSP.RETDATE, SRETURNSP.TOTALAMT FROM SRETURNSP INNER JOIN SPINVOICE ON SRETURNSP.INVOICE = SPINVOICE.SPINVNO " & _
          "WHERE SPINVOICE.CUSTOMERID=" & flxBill.TextMatrix(ActRow, 1) & " AND SRETURNSP.RETDATE>='" & Format(dStart, "yyyy-mm-dd") & "' And SRETURNSP.RETDATE<='" & Format(dEnd, "yyyy-mm-dd") & "' AND SRETURNSP.REMARKS<>'CANCELLED';"
    With rsAccts
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                flxDtls.AddItem !CUSTOMERID & vbTab & "RET. NO. " & !RetNo & vbTab & !RetDate & vbTab & "SPARE PARTS" & vbTab & vbTab & !TotalAmt
                TtlDmCm = TtlDmCm + !TotalAmt
                .MoveNext
            Loop
        End If
        .Close
    End With
    
    'Display Payment for SpareParts only
    'Qry = "SELECT LEDGER.CUSTOMERID, LEDGER.REFNO, LEDGER.REFDATE, Sum(LEDGER.CREDIT) AS SumOfCREDIT From Ledger " & _
          "WHERE (((LEDGER.REFTYPE)='Payment') AND ((LEDGER.JOBID) Is Null Or (LEDGER.JOBID)=0)) GROUP BY LEDGER.CUSTOMERID, LEDGER.REFNO, LEDGER.REFDATE " & _
          "HAVING (((LEDGER.CUSTOMERID)=" & flxBill.TextMatrix(ActRow, 1) & ") AND ((LEDGER.REFDATE)>=#" & dStart & "# And (LEDGER.REFDATE)<=#" & dEnd & "#));"
    
    'mySQL
    'Qry = "SELECT LEDGER.CUSTOMERID, LEDGER.REFNO, LEDGER.REFDATE, Sum(LEDGER.CREDIT) AS SumOfCREDIT From Ledger WHERE LEDGER.REFTYPE = 'Payment' AND (LEDGER.JOBID Is Null Or LEDGER.JOBID=0) " & _
            "GROUP BY LEDGER.CUSTOMERID, LEDGER.REFNO, LEDGER.REFDATE HAVING LEDGER.CUSTOMERID=" & flxBill.TextMatrix(ActRow, 1) & " AND LEDGER.REFDATE >= '" & Format(dStart, "yyyy-mm-dd") & "' " & _
            "And LEDGER.REFDATE<='" & Format(dEnd, "yyyy-mm-dd") & "';"
    
    Qry = "SELECT " & _
            "Ledger.CUSTOMERID," & _
            "Ledger.REFNO," & _
            "Ledger.REFDATE," & _
            "Sum(Ledger.CREDIT) As SumOfCREDIT " & _
          "From Ledger " & _
          "Where " & _
            "LEDGER.REFTYPE = 'Payment' AND " & _
            "LEDGER.JOBID IS NULL OR " & _
            "Ledger.JOBID = 0 " & _
          " Group By Ledger.CUSTOMERID," & _
            "Ledger.REFNO," & _
            "Ledger.REFDATE " & _
          "Having " & _
            "Ledger.CUSTOMERID = " & flxBill.TextMatrix(ActRow, 1) & " AND " & _
            "Ledger.REFDATE >= '" & Format(dStart, "yyyy-mm-dd") & "' AND " & _
            "Ledger.REFDATE <= '" & Format(dEnd, "yyyy-mm-dd") & "';"
    
    With rsAccts
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                flxDtls.AddItem !CUSTOMERID & vbTab & "O.R. NO. " & !RefNo & vbTab & !REFDATE & vbTab & "SPARE PARTS" & vbTab & vbTab & !SumOfCREDIT
                TtlPay = TtlPay + !SumOfCREDIT
                .MoveNext
            Loop
        End If
        .Close
    End With
    
    'Display Payment for Job Order only
    'Qry = "SELECT LEDGER.CUSTOMERID, LEDGER.REFNO, LEDGER.REFDATE, Sum(LEDGER.CREDIT) AS SumOfCREDIT From Ledger " & _
          "WHERE (((LEDGER.REFTYPE)='Payment') AND ((LEDGER.JOBID)>0)) GROUP BY LEDGER.CUSTOMERID, LEDGER.REFNO, LEDGER.REFDATE " & _
          "HAVING (((LEDGER.CUSTOMERID)=" & flxBill.TextMatrix(ActRow, 1) & ") AND ((LEDGER.REFDATE)>=#" & dStart & "# And (LEDGER.REFDATE)<=#" & dEnd & "#));"
    'mySQL error having statement
    'Qry = "SELECT LEDGER.CUSTOMERID, LEDGER.REFNO, LEDGER.REFDATE, Sum(LEDGER.CREDIT) AS SumOfCREDIT From Ledger " & _
          "WHERE LEDGER.REFTYPE='Payment' AND LEDGER.JOBID>0 GROUP BY LEDGER.CUSTOMERID, LEDGER.REFNO, LEDGER.REFDATE " & _
          "HAVING LEDGER.CUSTOMERID=" & flxBill.TextMatrix(ActRow, 1) & " AND LEDGER.REFDATE>='" & Format(dStart, "yyyy-mm-dd") & "' And LEDGER.REFDATE<='" & Format(dEnd, "yyyy-mm-dd") & "';"
    Qry = "SELECT " & _
            "Ledger.CUSTOMERID," & _
            "Ledger.REFNO," & _
            "Ledger.REFDATE," & _
            "Sum(Ledger.CREDIT) As SumOfCREDIT " & _
          "From Ledger " & _
          "Where " & _
            "LEDGER.REFTYPE = 'Payment' AND " & _
            "Ledger.JOBID > 0 " & _
          " Group By Ledger.CUSTOMERID," & _
            "Ledger.REFNO," & _
            "Ledger.REFDATE " & _
          "Having " & _
            "Ledger.CUSTOMERID = " & flxBill.TextMatrix(ActRow, 1) & " AND " & _
            "Ledger.REFDATE >= '" & Format(dStart, "yyyy-mm-dd") & "' AND " & _
            "Ledger.REFDATE <= '" & Format(dEnd, "yyyy-mm-dd") & "';"

    With rsAccts
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                flxDtls.AddItem !CUSTOMERID & vbTab & "O.R. NO. " & !RefNo & vbTab & !REFDATE & vbTab & "SERVICE" & vbTab & vbTab & !SumOfCREDIT
                TtlPay = TtlPay + !SumOfCREDIT
                .MoveNext
            Loop
        End If
        .Close
    End With
    
    'Qry to Display On Account for Service
    'Qry = "SELECT ONACCTERMS.CUSTOMERID, ONACCTERMS.JOBORDERNO, ONACCTERMS.TRANSDATE, ONACCTERMS.BALANCE From ONACCTERMS " & _
          "WHERE (((ONACCTERMS.CUSTOMERID)=" & flxBill.TextMatrix(ActRow, 1) & ") AND (Not (ONACCTERMS.JOBORDERNO) Is Null) AND ((ONACCTERMS.TRANSDATE)>=#" & dStart & "# And (ONACCTERMS.TRANSDATE)<=#" & dEnd & "#) AND ((ONACCTERMS.BALANCE)>=1) AND ((ONACCTERMS.STATUS)='OPEN') AND ((ONACCTERMS.SPINVNO) Is Null)) ORDER BY ONACCTERMS.TRANSDATE;"
    'mySQL
    Qry = "SELECT ONACCTERMS.CUSTOMERID, ONACCTERMS.JOBORDERNO, ONACCTERMS.TRANSDATE, ONACCTERMS.BALANCE From ONACCTERMS " & _
          "WHERE ONACCTERMS.CUSTOMERID=" & flxBill.TextMatrix(ActRow, 1) & " AND Not (ONACCTERMS.JOBORDERNO) Is Null AND ONACCTERMS.TRANSDATE>='" & Format(dStart, "yyyy-mm-dd") & "' And ONACCTERMS.TRANSDATE<='" & Format(dEnd, "yyyy-mm-dd") & "' AND ONACCTERMS.BALANCE>=1 AND ONACCTERMS.STATUS='OPEN' AND (ONACCTERMS.SPINVNO) Is Null ORDER BY ONACCTERMS.TRANSDATE;"
    With rsAccts
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                SvcAmt = GetDMCMamt(!JOBORDERNO, "JOB ORDER")
                flxDtls.AddItem !CUSTOMERID & vbTab & "JOB. NO. " & !JOBORDERNO & vbTab & !TRANSDATE & vbTab & "SERVICE" & vbTab & !Balance + SvcAmt
                TtlBill = TtlBill + !Balance + SvcAmt
                .MoveNext
            Loop
        End If
        .Close
    End With
    
    'Display DMCM for Service AND SpareParts
    'Qry = "SELECT ADJDMCM.CUSTOMERID, ADJDMCM.ADJNO, ADJDMCM.ADJDATE, ADJDMCM.DEBIT, ADJDMCM.CREDIT, ADJDMCM.TRANSTYPE From ADJDMCM " & _
          "WHERE (((ADJDMCM.CUSTOMERID)=" & flxBill.TextMatrix(ActRow, 1) & ") AND ((ADJDMCM.ADJDATE)>=#" & dStart & "# And (ADJDMCM.ADJDATE)<=#" & dEnd & "#) AND ((ADJDMCM.REASON)<>'CANCELLED')) ORDER BY ADJDMCM.ADJDATE;"
    'mySQL
    Qry = "SELECT ADJDMCM.CUSTOMERID, ADJDMCM.ADJNO, ADJDMCM.ADJDATE, ADJDMCM.DEBIT, ADJDMCM.CREDIT, ADJDMCM.TRANSTYPE From ADJDMCM " & _
          "WHERE ADJDMCM.CUSTOMERID=" & flxBill.TextMatrix(ActRow, 1) & " AND ADJDMCM.ADJDATE>='" & Format(dStart, "yyyy-mm-dd") & "' And ADJDMCM.ADJDATE<='" & Format(dEnd, "yyyy-mm-dd") & "' AND ADJDMCM.REASON<>'CANCELLED' ORDER BY ADJDMCM.ADJDATE;"
    With rsAccts
        .Open Qry, LYRcon, adOpenKeyset, adLockBatchOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                flxDtls.AddItem !CUSTOMERID & vbTab & "DMCM NO. " & !ADJNO & vbTab & !ADJDATE & vbTab & IIf(!TRANSTYPE = "JOB ORDER", "SERVICE", "SPARE PARTS") & vbTab & !Debit & vbTab & !Credit
                TtlDmCm = TtlDmCm - !Debit + !Credit
                .MoveNext
            Loop
        End If
        .Close
    End With
    
    'Other Charges from Disbursement
    'Qry = "SELECT Disbursement.CUSTOMERID, Disbursement.CVNUMBER, Disbursement.CVDATE, Disbursement.AMOUNT From Disbursement " & _
          "Where (((Disbursement.CUSTOMERID) = " & flxBill.TextMatrix(ActRow, 1) & ") And ((Disbursement.CVDate) >= #" & dStart & "# And (Disbursement.CVDate) <= #" & dEnd & "#) And ((Disbursement.AMOUNT) > 0)) ORDER BY Disbursement.CVNUMBER;"
    Qry = "SELECT Disbursement.CUSTOMERID, Disbursement.CVNUMBER, Disbursement.CVDATE, Disbursement.AMOUNT From Disbursement " & _
          "Where Disbursement.CUSTOMERID = " & flxBill.TextMatrix(ActRow, 1) & " And Disbursement.CVDate >= '" & Format(dStart, "yyyy-mm-dd") & "' And Disbursement.CVDate <='" & Format(dEnd, "yyyy-mm-dd") & "' And Disbursement.AMOUNT > 0 ORDER BY Disbursement.CVNUMBER;"
    With rsAccts
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                flxDtls.AddItem !CUSTOMERID & vbTab & "PCV. NO. " & Format(!CVNUMBER, "0000") & vbTab & !CVDate & vbTab & "OTHERS" & vbTab & !Amount
                TtlBill = TtlBill + !Amount
                .MoveNext
            Loop
        End If
        .Close
    End With
        
    'Payment for Disbursement
    'Qry = "SELECT PAYMENTS.CUSTOMERID, PAYMENTS.ORNO, PAYMENTS.PAYDATE, PAYMENTS.OTHERS, PAYMENTS.REMARKS From PAYMENTS " & _
          "WHERE (((PAYMENTS.CUSTOMERID)=" & flxBill.TextMatrix(ActRow, 1) & ") AND ((PAYMENTS.PAYDATE)>=#" & dStart & "# And (PAYMENTS.PAYDATE)<=#" & dEnd & "#) AND ((PAYMENTS.OTHERS)>0) AND ((PAYMENTS.REMARKS) Like 'Other Accts%'));"
    'mySQL
    Qry = "SELECT PAYMENTS.CUSTOMERID, PAYMENTS.ORNO, PAYMENTS.PAYDATE, PAYMENTS.OTHERS, PAYMENTS.REMARKS From PAYMENTS " & _
          "WHERE PAYMENTS.CUSTOMERID=" & flxBill.TextMatrix(ActRow, 1) & " AND PAYMENTS.PAYDATE>='" & Format(dStart, "yyyy-mm-dd") & "' And PAYMENTS.PAYDATE<='" & Format(dEnd, "yyyy-mm-dd") & "' AND PAYMENTS.OTHERS>0 AND (PAYMENTS.REMARKS) Like 'Other Accts%';"
    
    With rsAccts
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                flxDtls.AddItem !CUSTOMERID & vbTab & "O.R. NO. " & !ORNO & vbTab & !PayDate & vbTab & "OTHERS" & vbTab & vbTab & !Others
                TtlPay = TtlPay + !Others
                .MoveNext
            Loop
        End If
        .Close
    End With
    
    'get Previous Balance , Payments and DMCM
    PrvBal = GetPreviousBalance(flxBill.TextMatrix(ActRow, 1))
    TotalBal = (TtlBill + PrvBal) - (TtlPay + TtlDmCm)
    If BillID > 0 Then
        'LYRcon.Execute "DELETE BILLING.* From BILLING WHERE (((BILLING.CUSTOMERID)=" & flxBill.TextMatrix(ActRow, 1) & ") AND ((BILLING.BILLMO)=" & txtMo.Text & ") AND ((BILLING.BILLYR)=" & txtYr.Text & "));"
        'mySQL
        LYRcon.Execute "DELETE From BILLING WHERE (((BILLING.CUSTOMERID)=" & flxBill.TextMatrix(ActRow, 1) & ") AND ((BILLING.BILLMO)=" & txtMo.Text & ") AND ((BILLING.BILLYR)=" & txtYr.Text & "));"
        'LYRcon.Execute "DELETE BILLINGDTLS.* From BILLINGDTLS WHERE (((BILLINGDTLS.BILLINGID)=" & BillID & ") AND ((BILLINGDTLS.CUSTOMERID)=" & flxBill.TextMatrix(ActRow, 1) & "));"
        LYRcon.Execute "DELETE From BILLINGDTLS WHERE BILLINGDTLS.BILLINGID=" & BillID & " AND BILLINGDTLS.CUSTOMERID=" & flxBill.TextMatrix(ActRow, 1) & ";"
    End If
    If TotalBal <> 0 Then
        'Save Billing
        'LYRcon.Execute "INSERT INTO BILLING ( CUSTOMERID, DUEDATE, BILLMO, BILLYR, CURRENTBILL, PREVBAL, PAYMENTS, DMCM, BALANCE ) " & _
                       "VALUES (" & flxBill.TextMatrix(ActRow, 1) & ", #" & txtDueDate.Text & "#, " & txtMo.Text & ", " & txtYr.Text & ", " & TtlBill & ", " & PrvBal & ", " & TtlPay & ", " & TtlDmCm & ", " & TotalBal & ");"
        LYRcon.Execute "INSERT INTO BILLING(CUSTOMERID, DUEDATE, BILLMO, BILLYR, CURRENTBILL, PREVBAL, PAYMENTS, DMCM, BALANCE ) " & _
                       "VALUES(" & flxBill.TextMatrix(ActRow, 1) & ", '" & Format(txtDueDate.Text, "yyyy-mm-dd") & "', " & txtMo.Text & ", " & txtYr.Text & ", " & TtlBill & ", " & PrvBal & ", " & TtlPay & ", " & TtlDmCm & ", " & TotalBal & ");"
        'Qry = "SELECT BILLING.BILLINGID From BILLING WHERE (((BILLING.CUSTOMERID)=" & flxBill.TextMatrix(ActRow, 1) & ") AND ((BILLING.BILLMO)=" & txtMo.Text & ") AND ((BILLING.BILLYR)=" & txtYr.Text & "));"
        'mySQl
        Qry = "SELECT BILLING.BILLINGID From BILLING WHERE (((BILLING.CUSTOMERID)=" & flxBill.TextMatrix(ActRow, 1) & ") AND ((BILLING.BILLMO)=" & txtMo.Text & ") AND ((BILLING.BILLYR)=" & txtYr.Text & "));"
        BillID = Val(GetFieldValue(Qry))
        'Save Billing Details
        For j = 1 To flxDtls.Rows - 1
            LYRcon.Execute "INSERT INTO BILLINGDTLS ( BILLINGID, CUSTOMERID, REFNO, REFDATE, BILLTYPE, AMOUNTDM, AMOUNTCM ) " & _
                           "VALUES (" & BillID & ", " & flxBill.TextMatrix(ActRow, 1) & ", '" & flxDtls.TextMatrix(j, 1) & "', '" & Format(flxDtls.TextMatrix(j, 2), "yyyy-mm-dd") & "', '" & flxDtls.TextMatrix(j, 3) & "', " & Val(flxDtls.TextMatrix(j, 4)) & ", " & Val(flxDtls.TextMatrix(j, 5)) & ");"
        Next j
    End If
    'With flxBill      1                      2          3            4        5        6          7        8       9     10
    '    .FormatString = "   |  |< NAME OF CUSTOMER  |^DUE DATE|>PREV. BAL.|>CURRENT|>PAYMENTS |>DMCM|>AMOUNT DUE|^MONTH |^YEAR"
    flxBill.TextMatrix(ActRow, 4) = Format(PrvBal, "#,##0.00")
    flxBill.TextMatrix(ActRow, 5) = Format(TtlBill, "#,##0.00")
    flxBill.TextMatrix(ActRow, 6) = Format(TtlPay, "#,##0.00")
    flxBill.TextMatrix(ActRow, 7) = Format(TtlDmCm, "#,##0.00")
    flxBill.TextMatrix(ActRow, 8) = Format(TotalBal, "#,##0.00")
    MsgBox "Finish Processing Billing for " & MonthName(txtMo.Text) & " " & txtYr.Text & "!", vbInformation + vbOKOnly, "Update Complete"
    'Lock objects
    Frame1.Enabled = True
    cmdTrans(3).Enabled = True
    flxBill.Enabled = True
    
End Sub

Private Sub RXResizer1_ResizeComplete()
    Dim i As Integer, j As Integer
    
    lblHeading(2).Left = Frame1.Left + (Frame1.Width - lblHeading(2).Width) - 30
    lblHeading(3).Left = Frame1.Left + (Frame1.Width - lblHeading(3).Width)
    
    With flxBill             '   1       2          3               4          5         6       7          8        9     10
        .FormatString = "   |  |< NAME OF CUSTOMER  |^DUE DATE|>PREV. BAL.|>CURRENT|>PAYMENTS |>DMCM|>AMOUNT DUE|^MONTH |^YEAR"
        .ColWidth(0) = 300
        .ColWidth(1) = 0
        For i = 2 To 10
            .ColWidth(i) = (.Width - 300) * 0.12
        Next i
        .ColWidth(2) = (.Width - 300) * 0.25
        .ZOrder
        For i = 1 To .Rows - 1
            .TextMatrix(i, 3) = Format(.TextMatrix(i, 3), "mm/dd/yyyy")
            For j = 4 To 8
                .TextMatrix(i, j) = Format(.TextMatrix(i, j), "#,##0.00")
            Next j
        Next i
    End With
    
    txtFind.Top = cmdFind.Top
End Sub

Private Sub tabBill_Click(PreviousTab As Integer)
    If PreviousTab = 0 Then
        cmdFind.Enabled = False
        cmdNxtBill.Enabled = False
    Else
        cmdFind.Enabled = True
        cmdNxtBill.Enabled = True
    End If
End Sub


Private Sub Timer1_Timer()
    cboBilling.ListIndex = cboBilling.ListCount - 1
    'Call cbobilling_Click
    Timer1.Enabled = False
End Sub

Private Sub Timer2_Timer()
    Dim Ans As Integer, StartDate As Date
    Dim Qry As String
    
    'Display all SP in flxBill
    'Qry = "SELECT CUSTOMER.CUSTOMERID, CUSTOMER.CUSTOMERNAME From CUSTOMER WHERE (((CUSTOMER.CTYPE)<>'INDIVIDUAL')) ORDER BY CUSTOMER.CUSTOMERNAME;"
    'mySQL
    Qry = "SELECT CUSTOMER.CUSTOMERID, CUSTOMER.CUSTOMERNAME From CUSTOMER WHERE CUSTOMER.CTYPE<>'INDIVIDUAL' ORDER BY CUSTOMER.CUSTOMERNAME;"
    If rsCustBill.State = adStateOpen Then rsCustBill.Close
    rsCustBill.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    If rsCustBill.RecordCount = 0 Then
        flxBill.Clear
        flxBill.Rows = 2
        flxBill.Cols = 10
        Call RXResizer1_ResizeComplete
        Exit Sub
    End If
    Set flxBill.DataSource = rsCustBill
    ActRow = 1
    Call RXResizer1_ResizeComplete
    
    'Identify 1st month-year cutoff
    'Qry = "SELECT SPINVOICE.INVDATE FROM SPINVOICE INNER JOIN CUSTOMER ON SPINVOICE.CUSTOMERID = CUSTOMER.CUSTOMERID " & _
          "WHERE (((SPINVOICE.INVTYPE)='CREDIT') AND ((CUSTOMER.CTYPE)<>'INDIVIDUAL')) ORDER BY SPINVOICE.INVDATE;"
    Qry = "SELECT SPINVOICE.INVDATE FROM SPINVOICE INNER JOIN CUSTOMER ON SPINVOICE.CUSTOMERID = CUSTOMER.CUSTOMERID " & _
          "WHERE SPINVOICE.INVTYPE='CREDIT' AND CUSTOMER.CTYPE<>'INDIVIDUAL' ORDER BY SPINVOICE.INVDATE;"

    StartDate = GetFieldValue(Qry)
    txtMo.Text = Month(StartDate)
    txtYr.Text = Year(StartDate)
    
    Ans = MsgBox("Do you want to Process Billing for the Month of " & MonthName(txtMo.Text) & " - " & txtYr.Text, vbQuestion + vbYesNo, MonthName(txtMo.Text) & " Cuf-off")
    If Ans = vbNo Then
        Timer2.Enabled = False
        Exit Sub
    End If
    cboBilling.AddItem MonthName(txtMo.Text) & " - " & txtYr.Text
    cboBilling.Text = MonthName(txtMo.Text) & " - " & txtYr.Text
    txtDueDate.Text = DateAdd("m", 1, DateSerial(txtYr.Text, txtMo.Text, 15))
    Call ProcessBilling
    Timer2.Enabled = False
End Sub

Private Sub txtFind_Change()
    fRow = 1
    fCol = 2
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then Call cmdFind_Click
End Sub

Private Sub ProcessBilling()
    Dim rsAccts As New ADODB.Recordset
    Dim dStart As Date, dEnd As Date, BillID As Long
    Dim Qry As String, Qinv As String, TotalBal As Currency
    Dim i As Integer, j As Integer, TtlBill As Currency, PrvBal As Currency, TtlPay As Currency
    Dim MyType As String, rRet As String, AmtRet As Currency, InvAmt As Currency, TtlDmCm As Currency
    Dim SvcAmt As Currency, AdjAmt As Currency, QCridt As String, SumOfCredt As Long
    
    'Lock objects
    Frame1.Enabled = False
    cmdTrans(3).Enabled = False
    flxBill.Enabled = False
    
    dStart = DateSerial(txtYr.Text, Val(txtMo.Text), 1)
    dEnd = DateAdd("d", -1, DateAdd("m", 1, dStart))
        
    Pbar.Value = 0
    Pbar.Max = flxBill.Rows - 1
    Pbar.Visible = True
    
    For i = 1 To flxBill.Rows - 1
        'Start Progress bar
        'Make Sure All On-Account Invoice and JO are accounted
            
        'Get On Account for Spare Parts
        TtlBill = 0
        TtlDmCm = 0
        TtlPay = 0
        flxDtls.Rows = 1
        'Qry = "SELECT ONACCTERMS.CUSTOMERID, ONACCTERMS.SPINVNO, ONACCTERMS.TRANSDATE, ONACCTERMS.BALANCE, ONACCTERMS.ACCTAMOUNT From ONACCTERMS " & _
              "WHERE (((ONACCTERMS.CUSTOMERID)=" & flxBill.TextMatrix(i, 1) & ") AND (Not (ONACCTERMS.SPINVNO) Is Null) AND ((ONACCTERMS.TRANSDATE)>=#" & dStart & "# And (ONACCTERMS.TRANSDATE)<=#" & dEnd & "#) AND ((ONACCTERMS.STATUS)<>'CANCELLED')) ORDER BY ONACCTERMS.TRANSDATE;"
        'mySQL
        Qry = "SELECT ONACCTERMS.CUSTOMERID, ONACCTERMS.SPINVNO, ONACCTERMS.TRANSDATE, ONACCTERMS.BALANCE, ONACCTERMS.ACCTAMOUNT From ONACCTERMS " & _
              "WHERE ONACCTERMS.CUSTOMERID=" & flxBill.TextMatrix(i, 1) & " AND Not (ONACCTERMS.SPINVNO) Is Null AND ONACCTERMS.TRANSDATE>='" & Format(dStart, "yyyy-mm-dd") & "' And ONACCTERMS.TRANSDATE<='" & Format(dEnd, "yyyy-mm-dd") & "' AND ONACCTERMS.STATUS<>'CANCELLED' ORDER BY ONACCTERMS.TRANSDATE;"
        
        With rsAccts
            .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
            If .RecordCount > 0 Then
                .MoveFirst
                Do While Not .EOF
                    'rRet = "SELECT Sum(SRETURNSP.TOTALAMT) AS SumOfTOTALAMT From SRETURNSP " & _
                           "WHERE (((SRETURNSP.INVOICE)='" & !SPINVNO & "') AND ((SRETURNSP.REMARKS) Not Like 'CANCELLED'));"
                    'mySQL
                    rRet = "SELECT Sum(SRETURNSP.TOTALAMT) AS SumOfTOTALAMT From SRETURNSP " & _
                           "WHERE SRETURNSP.INVOICE='" & !SPINVNO & "' AND (SRETURNSP.REMARKS) Not Like 'CANCELLED';"
                    AmtRet = Val(GetFieldValue(rRet))
                    AmtAdj = GetDMCMamt(!SPINVNO, "INVOICE")
                    InvAmt = !AcctAmount    '!Balance + AmtRet + AmtAdj
                    flxDtls.AddItem !CUSTOMERID & vbTab & "INV. NO. " & !SPINVNO & vbTab & !TRANSDATE & vbTab & "SPARE PARTS" & vbTab & InvAmt
                    TtlBill = TtlBill + InvAmt
                    .MoveNext
                Loop
            End If
            .Close
        End With
        
        'Display Sales Return
        'Qry = "SELECT SPINVOICE.CUSTOMERID, SRETURNSP.RETNO, SRETURNSP.RETDATE, SRETURNSP.TOTALAMT FROM SRETURNSP INNER JOIN SPINVOICE ON SRETURNSP.INVOICE = SPINVOICE.SPINVNO " & _
              "WHERE (((SPINVOICE.CUSTOMERID)=" & flxBill.TextMatrix(i, 1) & ") AND ((SRETURNSP.RETDATE)>=#" & dStart & "# And (SRETURNSP.RETDATE)<=#" & dEnd & "#) AND ((SRETURNSP.REMARKS)<>'CANCELLED'));"
        'mySQL
        Qry = "SELECT SPINVOICE.CUSTOMERID, SRETURNSP.RETNO, SRETURNSP.RETDATE, SRETURNSP.TOTALAMT FROM SRETURNSP INNER JOIN SPINVOICE ON SRETURNSP.INVOICE = SPINVOICE.SPINVNO " & _
              "WHERE SPINVOICE.CUSTOMERID=" & flxBill.TextMatrix(i, 1) & " AND SRETURNSP.RETDATE>='" & Format(dStart, "yyyy-mm-dd") & "' And SRETURNSP.RETDATE<='" & Format(dEnd, "yyyy-mm-dd") & "' AND SRETURNSP.REMARKS<>'CANCELLED';"
        With rsAccts
            .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
            If .RecordCount > 0 Then
                .MoveFirst
                Do While Not .EOF
                    flxDtls.AddItem !CUSTOMERID & vbTab & "RET. NO. " & !RetNo & vbTab & !RetDate & vbTab & "SPARE PARTS" & vbTab & vbTab & !TotalAmt
                    TtlDmCm = TtlDmCm + !TotalAmt
                    .MoveNext
                Loop
            End If
            .Close
        End With
        
        'Display Payment for SpareParts only
        'Qry = "SELECT LEDGER.CUSTOMERID, LEDGER.REFNO, LEDGER.REFDATE, Sum(LEDGER.CREDIT) AS SumOfCREDIT From Ledger " & _
        '  "WHERE (((LEDGER.REFTYPE)='Payment') AND ((LEDGER.JOBID) Is Null Or (LEDGER.JOBID)=0)) GROUP BY LEDGER.CUSTOMERID, LEDGER.REFNO, LEDGER.REFDATE " & _
        '  "HAVING (((LEDGER.CUSTOMERID)=" & flxBill.TextMatrix(ActRow, 1) & ") AND ((LEDGER.REFDATE)>='" & Format(dStart, "yyyy-mm-dd") & "' And (LEDGER.REFDATE)<='" & Format(dEnd, "yyyy-mm-dd") & "'));"
        
        'mySQL error having statement
        'Qry = "SELECT LEDGER.CUSTOMERID, LEDGER.REFNO, LEDGER.REFDATE, Sum(LEDGER.CREDIT) AS SumOfCREDIT From Ledger WHERE LEDGER.REFTYPE='Payment' AND ((LEDGER.JOBID) Is Null Or LEDGER.JOBID=0) " & _
            "GROUP BY LEDGER.CUSTOMERID, LEDGER.REFNO, LEDGER.REFDATE HAVING LEDGER.CUSTOMERID=" & flxBill.TextMatrix(ActRow, 1) & " AND " & _
            "LEDGER.REFDATE>='" & Format(dStart, "yyyy-mm-dd") & "' And LEDGER.REFDATE<='" & Format(dEnd, "yyyy-mm-dd") & "';"
        'New query coded as indented form
        
        Qry = "SELECT " & _
            "Ledger.CUSTOMERID," & _
            "Ledger.REFNO," & _
            "Ledger.REFDATE," & _
            "Sum(Ledger.CREDIT) As SumOfCREDIT " & _
          "From Ledger " & _
          "Where " & _
            "LEDGER.REFTYPE = 'Payment' AND " & _
            "LEDGER.JOBID IS NULL OR " & _
            "Ledger.JOBID = 0 " & _
          " Group By Ledger.CUSTOMERID," & _
            "Ledger.REFNO," & _
            "Ledger.REFDATE " & _
          "Having " & _
            "Ledger.CUSTOMERID = " & flxBill.TextMatrix(ActRow, 1) & " AND " & _
            "Ledger.REFDATE >= '" & Format(dStart, "yyyy-mm-dd") & "' AND " & _
            "Ledger.REFDATE <= '" & Format(dEnd, "yyyy-mm-dd") & "';"

          With rsAccts
            .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
            If .RecordCount > 0 Then
                .MoveFirst
                Do While Not .EOF
                    flxDtls.AddItem !CUSTOMERID & vbTab & "O.R. NO. " & !RefNo & vbTab & !REFDATE & vbTab & "SPARE PARTS" & vbTab & vbTab & !SumOfCREDIT
                    TtlPay = TtlPay + !SumOfCREDIT
                    .MoveNext
                Loop
            End If
            .Close
        End With
        
        'Display Payment for Job Order only
        'Qry = "SELECT LEDGER.CUSTOMERID, LEDGER.REFNO, LEDGER.REFDATE, Sum(LEDGER.CREDIT) AS SumOfCREDIT From Ledger " & _
              "WHERE (((LEDGER.REFTYPE)='Payment') AND ((LEDGER.JOBID)>0)) GROUP BY LEDGER.CUSTOMERID, LEDGER.REFNO, LEDGER.REFDATE " & _
              "HAVING (((LEDGER.CUSTOMERID)=" & flxBill.TextMatrix(i, 1) & ") AND ((LEDGER.REFDATE)>=#" & dStart & "# And (LEDGER.REFDATE)<=#" & dEnd & "#));"
        'mysql
        Qry = "SELECT " & _
                "Ledger.CUSTOMERID," & _
                "Ledger.REFNO," & _
                "Ledger.REFDATE," & _
                "Sum(Ledger.CREDIT) As SumOfCREDIT " & _
              "From " & _
                "Ledger " & _
              "Where " & _
                "LEDGER.REFTYPE = 'Payment' AND " & _
                "LEDGER.JOBID > 0 " & _
              "Group By " & _
                "Ledger.CUSTOMERID," & _
                "Ledger.REFNO," & _
                "Ledger.REFDATE " & _
              "Having " & _
                "Ledger.CUSTOMERID = " & flxBill.TextMatrix(i, 1) & " AND " & _
                "Ledger.REFDATE >= '" & Format(dStart, "yyyy-mm-dd") & "' AND " & _
                "Ledger.REFDATE <= '" & Format(dEnd, "yyyy-mm-dd") & "';"
                
        With rsAccts
            .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
            If .RecordCount > 0 Then
                .MoveFirst
                Do While Not .EOF
                    flxDtls.AddItem !CUSTOMERID & vbTab & "O.R. NO. " & !RefNo & vbTab & !REFDATE & vbTab & "SERVICE" & vbTab & vbTab & !SumOfCREDIT
                    TtlPay = TtlPay + !SumOfCREDIT
                    .MoveNext
                Loop
            End If
            .Close
        End With
        
                        
        'Qry to Display On Account for Service
        'Qry = "SELECT ONACCTERMS.CUSTOMERID, ONACCTERMS.JOBORDERNO, ONACCTERMS.TRANSDATE, ONACCTERMS.BALANCE From ONACCTERMS " & _
              "WHERE (((ONACCTERMS.CUSTOMERID)=" & flxBill.TextMatrix(i, 1) & ") AND (Not (ONACCTERMS.JOBORDERNO) Is Null) AND ((ONACCTERMS.TRANSDATE)>=#" & dStart & "# And (ONACCTERMS.TRANSDATE)<=#" & dEnd & "#) AND ((ONACCTERMS.BALANCE)>=1) AND ((ONACCTERMS.STATUS)='OPEN') AND ((ONACCTERMS.SPINVNO) Is Null)) ORDER BY ONACCTERMS.TRANSDATE;"
        'mysql
        Qry = "SELECT ONACCTERMS.CUSTOMERID, ONACCTERMS.JOBORDERNO, ONACCTERMS.TRANSDATE, ONACCTERMS.BALANCE From ONACCTERMS " & _
              "WHERE (((ONACCTERMS.CUSTOMERID)=" & flxBill.TextMatrix(i, 1) & ") AND (Not (ONACCTERMS.JOBORDERNO) Is Null) AND ((ONACCTERMS.TRANSDATE)>='" & Format(dStart, "yyyy-mm-dd") & "' And (ONACCTERMS.TRANSDATE)<='" & Format(dEnd, "yyyy-mm-dd") & "') AND ((ONACCTERMS.BALANCE)>=1) AND ((ONACCTERMS.STATUS)='OPEN') AND ((ONACCTERMS.SPINVNO) Is Null)) ORDER BY ONACCTERMS.TRANSDATE;"
              
        With rsAccts
            .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
            If .RecordCount > 0 Then
                .MoveFirst
                Do While Not .EOF
                    SvcAmt = GetDMCMamt(!JOBORDERNO, "JOB ORDER")
                    flxDtls.AddItem !CUSTOMERID & vbTab & "JOB. NO. " & !JOBORDERNO & vbTab & !TRANSDATE & vbTab & "SERVICE" & vbTab & !Balance + SvcAmt
                    TtlBill = TtlBill + !Balance + SvcAmt
                    .MoveNext
                Loop
            End If
            .Close
        End With
        
        'Display DMCM for Service AND SpareParts
        'Qry = "SELECT ADJDMCM.CUSTOMERID, ADJDMCM.REFNO, ADJDMCM.ADJDATE, ADJDMCM.DEBIT, ADJDMCM.CREDIT, ADJDMCM.TRANSTYPE From ADJDMCM " & _
              "WHERE (((ADJDMCM.CUSTOMERID)=" & flxBill.TextMatrix(i, 1) & ") AND ((ADJDMCM.ADJDATE)>=#" & dStart & "# And (ADJDMCM.ADJDATE)<=#" & dEnd & "#) AND ((ADJDMCM.REASON)<>'CANCELLED')) ORDER BY ADJDMCM.ADJDATE;"
        'mysql
        Qry = "SELECT ADJDMCM.CUSTOMERID, ADJDMCM.REFNO, ADJDMCM.ADJDATE, ADJDMCM.DEBIT, ADJDMCM.CREDIT, ADJDMCM.TRANSTYPE From ADJDMCM " & _
              "WHERE (((ADJDMCM.CUSTOMERID)=" & flxBill.TextMatrix(i, 1) & ") AND ((ADJDMCM.ADJDATE)>='" & Format(dStart, "yyyy-mm-dd") & "' And (ADJDMCM.ADJDATE)<='" & Format(dEnd, "yyyy-mm-dd") & "') AND ((ADJDMCM.REASON)<>'CANCELLED')) ORDER BY ADJDMCM.ADJDATE;"
     
        With rsAccts
            .Open Qry, LYRcon, adOpenKeyset, adLockBatchOptimistic
            If .RecordCount > 0 Then
                .MoveFirst
                Do While Not .EOF
                    flxDtls.AddItem !CUSTOMERID & vbTab & "DMCM NO. " & !RefNo & vbTab & !ADJDATE & vbTab & IIf(!TRANSTYPE = "JOB ORDER", "SERVICE", "SPARE PARTS") & vbTab & !Debit & vbTab & !Credit
                    TtlDmCm = TtlDmCm - !Debit + !Credit
                    .MoveNext
                Loop
            End If
            .Close
        End With
        
        'Other Charges from Disbursement
        'Qry = "SELECT Disbursement.CUSTOMERID, Disbursement.CVNUMBER, Disbursement.CVDATE, Disbursement.AMOUNT From Disbursement " & _
              "Where (((Disbursement.CUSTOMERID) = " & flxBill.TextMatrix(i, 1) & ") And ((Disbursement.CVDate) >= #" & dStart & "# And (Disbursement.CVDate) <= #" & dEnd & "#) And ((Disbursement.AMOUNT) > 0)) ORDER BY Disbursement.CVNUMBER;"
        'Mysql
        Qry = "SELECT Disbursement.CUSTOMERID, Disbursement.CVNUMBER, Disbursement.CVDATE, Disbursement.AMOUNT From Disbursement " & _
              "Where (((Disbursement.CUSTOMERID) = " & flxBill.TextMatrix(i, 1) & ") And ((Disbursement.CVDate) >= '" & Format(dStart, "yyyy-mm-dd") & "' And (Disbursement.CVDate) <= '" & Format(dEnd, "yyyy-mm-dd") & "') And ((Disbursement.AMOUNT) > 0)) ORDER BY Disbursement.CVNUMBER;"
              
        With rsAccts
            .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
            If .RecordCount > 0 Then
                .MoveFirst
                Do While Not .EOF
                    flxDtls.AddItem !CUSTOMERID & vbTab & "PCV. NO. " & Format(!CVNUMBER, "0000") & vbTab & !CVDate & vbTab & "OTHERS" & vbTab & !Amount
                    TtlBill = TtlBill + !Amount
                    .MoveNext
                Loop
            End If
            .Close
        End With
        
        'Payment for Disbursement
        'Qry = "SELECT PAYMENTS.CUSTOMERID, PAYMENTS.ORNO, PAYMENTS.PAYDATE, PAYMENTS.OTHERS, PAYMENTS.REMARKS From PAYMENTS " & _
              "WHERE (((PAYMENTS.CUSTOMERID)=" & flxBill.TextMatrix(i, 1) & ") AND ((PAYMENTS.PAYDATE)>=#" & dStart & "# And (PAYMENTS.PAYDATE)<=#" & dEnd & "#) AND ((PAYMENTS.OTHERS)>0) AND ((PAYMENTS.REMARKS) Like 'Other Accts%'));"
        'mysql
        Qry = "SELECT PAYMENTS.CUSTOMERID, PAYMENTS.ORNO, PAYMENTS.PAYDATE, PAYMENTS.OTHERS, PAYMENTS.REMARKS From PAYMENTS " & _
              "WHERE (((PAYMENTS.CUSTOMERID)=" & flxBill.TextMatrix(i, 1) & ") AND ((PAYMENTS.PAYDATE)>='" & Format(dStart, "yyyy-mm-dd") & "' And (PAYMENTS.PAYDATE)<='" & Format(dEnd, "yyyy-mm-dd") & "') AND ((PAYMENTS.OTHERS)>0) AND ((PAYMENTS.REMARKS) Like 'Other Accts%'));"
        
        With rsAccts
            .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
            If .RecordCount > 0 Then
                .MoveFirst
                Do While Not .EOF
                    flxDtls.AddItem !CUSTOMERID & vbTab & "O.R. NO. " & !ORNO & vbTab & !PayDate & vbTab & "OTHERS" & vbTab & vbTab & !Others
                    TtlPay = TtlPay + !Others
                    .MoveNext
                Loop
            End If
            .Close
        End With
        
        'get Previous Balance , Payments and DMCM
        PrvBal = GetPreviousBalance(flxBill.TextMatrix(i, 1))
        TotalBal = (TtlBill + PrvBal) - (TtlPay + TtlDmCm)
        If TotalBal <> 0 Then
            'Save Billing
            'LYRcon.Execute "INSERT INTO BILLING ( CUSTOMERID, DUEDATE, BILLMO, BILLYR, CURRENTBILL, PREVBAL, PAYMENTS, DMCM, BALANCE ) " & _
                           "VALUES (" & flxBill.TextMatrix(i, 1) & ", #" & txtDueDate.Text & "#, " & txtMo.Text & ", " & txtYr.Text & ", " & TtlBill & ", " & PrvBal & ", " & TtlPay & ", " & TtlDmCm & ", " & TotalBal & ");"
            'mySQL
            LYRcon.Execute "INSERT INTO BILLING(CUSTOMERID, DUEDATE, BILLMO, BILLYR, CURRENTBILL, PREVBAL, PAYMENTS, DMCM, BALANCE ) " & _
                           "VALUES(" & flxBill.TextMatrix(i, 1) & ", '" & Format(txtDueDate.Text, "yyyy-mm-dd") & "', " & txtMo.Text & ", " & txtYr.Text & ", " & TtlBill & ", " & PrvBal & ", " & TtlPay & ", " & TtlDmCm & ", " & TotalBal & ");"
            
            'Qry = "SELECT BILLING.BILLINGID From BILLING WHERE (((BILLING.CUSTOMERID)=" & flxBill.TextMatrix(i, 1) & ") AND ((BILLING.BILLMO)=" & txtMo.Text & ") AND ((BILLING.BILLYR)=" & txtYr.Text & "));"
            'mysql
            Qry = "SELECT BILLING.BILLINGID From BILLING WHERE (((BILLING.CUSTOMERID)=" & flxBill.TextMatrix(i, 1) & ") AND ((BILLING.BILLMO)=" & txtMo.Text & ") AND ((BILLING.BILLYR)=" & txtYr.Text & "));"
           
            BillID = Val(GetFieldValue(Qry))
                       
            'Save Billing Details
            For j = 1 To flxDtls.Rows - 1
                'LYRcon.Execute "INSERT INTO BILLINGDTLS ( BILLINGID, CUSTOMERID, REFNO, REFDATE, BILLTYPE, AMOUNTDM, AMOUNTCM ) " & _
                               "VALUES (" & BillID & ", " & flxBill.TextMatrix(i, 1) & ", '" & flxDtls.TextMatrix(j, 1) & "', #" & flxDtls.TextMatrix(j, 2) & "#, '" & flxDtls.TextMatrix(j, 3) & "', " & Val(flxDtls.TextMatrix(j, 4)) & ", " & Val(flxDtls.TextMatrix(j, 5)) & ");"
                'mysql
                LYRcon.Execute "INSERT INTO BILLINGDTLS(BILLINGID, CUSTOMERID, REFNO, REFDATE, BILLTYPE, AMOUNTDM, AMOUNTCM) " & _
                               "VALUES(" & BillID & ", " & flxBill.TextMatrix(i, 1) & ", '" & flxDtls.TextMatrix(j, 1) & "', '" & Format(flxDtls.TextMatrix(j, 2), "yyyy-mm-dd") & "', '" & flxDtls.TextMatrix(j, 3) & "', " & Val(flxDtls.TextMatrix(j, 4)) & ", " & Val(flxDtls.TextMatrix(j, 5)) & ");"
            Next j
        End If
        Pbar.Value = Pbar.Value + 1
        DoEvents
        pPct = (Pbar.Value / Pbar.Max)
        lblTotal.Caption = Format$(pPct, "##0.0%")
        DoEvents
    Next i
    DoEvents
    
    MsgBox "Finish Processing Billing for " & MonthName(txtMo.Text) & " " & txtYr.Text & "!", vbInformation + vbOKOnly, "Update Complete"
    Pbar.Visible = False
    lblTotal.Caption = ""
    Frame1.Enabled = True
    cmdTrans(3).Enabled = True
    flxBill.Enabled = True
    Call DisplayBilling
End Sub

Function GetDMCMamt(MyRef As String, RefType As String) As Currency
    Dim rsDmCm As New ADODB.Recordset
    Dim Qry As String
    
    'Qry = "SELECT ADJDMCM.REFNO, ADJDMCM.TRANSTYPE, ADJDMCM.DEBIT, ADJDMCM.CREDIT, ADJDMCM.REASON From ADJDMCM " & _
          "WHERE (((ADJDMCM.REFNO)='" & MyRef & "') AND ((ADJDMCM.TRANSTYPE)='" & RefType & "') AND ((ADJDMCM.REASON)<>'CANCELLED'));"
    'mySQL
    Qry = "SELECT ADJDMCM.REFNO, ADJDMCM.TRANSTYPE, ADJDMCM.DEBIT, ADJDMCM.CREDIT, ADJDMCM.REASON From ADJDMCM " & _
          "WHERE (((ADJDMCM.REFNO)='" & MyRef & "') AND ((ADJDMCM.TRANSTYPE)='" & RefType & "') AND ((ADJDMCM.REASON)<>'CANCELLED'));"
    
    With rsDmCm
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            If !Credit > 0 Then
                GetDMCMamt = !Credit
            Else
                GetDMCMamt = !Debit * -1
            End If
        Else
            GetDMCMamt = 0
        End If
        .Close
    End With
End Function

Function GetPreviousBalance(MyID As Integer) As Currency
    Dim rsPrev As New ADODB.Recordset
    Dim Qry As String, BillDate As Date, tmpBal As Currency
    
    'Qry = "SELECT BILLING.* From BILLING Where (((BILLING.CUSTOMERID) = " & MyId & ")) ORDER BY BILLING.BILLYR DESC , BILLING.BILLMO DESC;"
    'mySQL
    Qry = "SELECT BILLING.* From BILLING Where BILLING.CUSTOMERID = " & MyID & " ORDER BY BILLING.BILLYR DESC , BILLING.BILLMO DESC;"
   
    tmpBal = 0
    With rsPrev
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                BillDate = DateSerial(!BILLYR, !BILLMO, 1)
                If BillDate < DateSerial(txtYr.Text, txtMo.Text, 1) Then
                    tmpBal = IIf(IsNull(!Balance), 0, !Balance)
                    Exit Do
                End If
                .MoveNext
            Loop
        Else
            tmpBal = 0
        End If
        .Close
    End With
    GetPreviousBalance = tmpBal
End Function
Private Sub DisplayBilling()
    Dim crSOA As New rptSOA
    Dim Qry As String, Dfrom As String, Dto As String
    
    'Qry = "SELECT BILLING.CUSTOMERID, CUSTOMER.CUSTOMERNAME, BILLING.DUEDATE, BILLING.PREVBAL, BILLING.CURRENTBILL, BILLING.PAYMENTS, BILLING.DMCM, BILLING.BALANCE, BILLING.BILLMO, BILLING.BILLYR " & _
          "FROM CUSTOMER INNER JOIN BILLING ON CUSTOMER.CUSTOMERID = BILLING.CUSTOMERID Where (((BILLING.BILLMO) = " & txtMo.Text & ") And ((BILLING.BILLYR) = " & txtYr.Text & ")) ORDER BY CUSTOMER.CUSTOMERNAME;"
    'mySQL
    Qry = "SELECT BILLING.CUSTOMERID, CUSTOMER.CUSTOMERNAME, BILLING.DUEDATE, BILLING.PREVBAL, BILLING.CURRENTBILL, BILLING.PAYMENTS, BILLING.DMCM, BILLING.BALANCE, BILLING.BILLMO, BILLING.BILLYR " & _
          "FROM CUSTOMER INNER JOIN BILLING ON CUSTOMER.CUSTOMERID = BILLING.CUSTOMERID Where (((BILLING.BILLMO) = " & txtMo.Text & ") And ((BILLING.BILLYR) = " & txtYr.Text & ")) ORDER BY CUSTOMER.CUSTOMERNAME;"

    If rsCustBill.State = adStateOpen Then rsCustBill.Close
    rsCustBill.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    If rsCustBill.RecordCount = 0 Then
        flxBill.Clear
        flxBill.Rows = 2
        flxBill.Cols = 8
        Call RXResizer1_ResizeComplete
        Exit Sub
    End If
    Set flxBill.DataSource = rsCustBill
    ActRow = 1
    Call RXResizer1_ResizeComplete
    
    'Display Print Preview
    Dfrom = DateSerial(txtYr.Text, txtMo.Text, 1)
    Dto = DateAdd("d", -1, DateAdd("m", 1, Dfrom))
    'old    Qry = "SELECT CUSTOMER.CUSTOMERNAME, CUSTOMER.ADDRESS, BILLING.DUEDATE, BILLING.PREVBAL, BILLING.CURRENTBILL, BILLING.PAYMENTS, BILLING.DMCM, BILLING.BALANCE, BILLINGDTLS.BILLTYPE, BILLINGDTLS.REFDATE, BILLINGDTLS.REFNO, BILLINGDTLS.AMOUNTDM, BILLINGDTLS.AMOUNTCM, CUSTOMER.CONTACTPERSON, CUSTOMER.CTYPE " & _
          "FROM BILLINGDTLS INNER JOIN (BILLING INNER JOIN CUSTOMER ON BILLING.CUSTOMERID = CUSTOMER.CUSTOMERID) ON BILLINGDTLS.CUSTOMERID = CUSTOMER.CUSTOMERID Where (((BILLING.DUEDATE) = #" & txtDueDate.Text & "#) And ((BILLINGDTLS.REFDATE) >= #" & dFrom & "# And (BILLINGDTLS.REFDATE) <= #" & dTo & "#)) ORDER BY BILLINGDTLS.REFDATE, BILLINGDTLS.REFNO;"
    
    'noted if we remove ( ) query can't run error syntax
    Qry = "SELECT CUSTOMER.CUSTOMERNAME, CUSTOMER.ADDRESS, BILLING.DUEDATE, BILLING.PREVBAL, BILLING.CURRENTBILL, BILLING.PAYMENTS, BILLING.DMCM, BILLING.BALANCE, BILLINGDTLS.BILLTYPE, BILLINGDTLS.REFDATE, BILLINGDTLS.REFNO, BILLINGDTLS.AMOUNTDM, BILLINGDTLS.AMOUNTCM, CUSTOMER.CONTACTPERSON, CUSTOMER.CTYPE " & _
          "FROM CUSTOMER INNER JOIN (BILLINGDTLS INNER JOIN BILLING ON BILLINGDTLS.BILLINGID = BILLING.BILLINGID) ON CUSTOMER.CUSTOMERID = BILLING.CUSTOMERID Where (((BILLING.BILLMO) = " & txtMo.Text & ") And ((BILLING.BILLYR) = " & txtYr.Text & ")) ORDER BY BILLINGDTLS.REFDATE, BILLINGDTLS.REFNO;"
    
    If rsPrn.State = adStateOpen Then rsPrn.Close
    rsPrn.Open Qry, LYRcon, adOpenDynamic, adLockOptimistic
    'crSOA.DiscardSavedData
    crSOA.Database.SetDataSource rsPrn
    crSOA.ParameterFields(1).ClearCurrentValueAndRange
    crSOA.ParameterFields(1).AddCurrentValue Dfrom
    crSOA.ParameterFields(2).ClearCurrentValueAndRange
    crSOA.ParameterFields(2).AddCurrentValue Dto
    crSOA.ParameterFields(3).ClearCurrentValueAndRange
    crSOA.ParameterFields(3).AddCurrentValue GetSettings("CASHIER")
    crSOA.ParameterFields(4).ClearCurrentValueAndRange
    crSOA.ParameterFields(4).AddCurrentValue GetSettings("MANAGER")
    crSOA.ParameterFields(5).ClearCurrentValueAndRange
    crSOA.ParameterFields(5).AddCurrentValue GetSettings("OIC")
    CRViewer1.ReportSource = crSOA
    CRViewer1.ViewReport
    CRViewer1.DisplayGroupTree = False
    CRViewer1.EnableExportButton = True
    
End Sub

