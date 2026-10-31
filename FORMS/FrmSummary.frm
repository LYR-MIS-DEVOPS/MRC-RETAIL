VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmSummary 
   Caption         =   "SUMMARY"
   ClientHeight    =   8715
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   14640
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8715
   ScaleWidth      =   14640
   WindowState     =   2  'Maximized
   Begin VB.CommandButton statUpdate 
      Caption         =   "Status Update"
      Height          =   495
      Left            =   9360
      TabIndex        =   29
      Tag             =   "XY"
      Top             =   840
      Visible         =   0   'False
      Width           =   1695
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   2
      Left            =   1800
      TabIndex        =   9
      Top             =   0
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   661
      Caption         =   "&Print Summary"
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
      Image           =   "FrmSummary.frx":0000
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   1
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   661
      Caption         =   "Show &Form"
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
      Image           =   "FrmSummary.frx":059A
      cBack           =   -2147483633
   End
   Begin VB.TextBox txtStatus 
      Height          =   375
      Left            =   15120
      TabIndex        =   26
      Top             =   2880
      Visible         =   0   'False
      Width           =   615
   End
   Begin MSComDlg.CommonDialog cdlgVan 
      Left            =   15480
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ProgressBar Pbar 
      Height          =   135
      Left            =   120
      TabIndex        =   25
      Tag             =   "xyf"
      Top             =   8520
      Visible         =   0   'False
      Width           =   14415
      _ExtentX        =   25426
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.ComboBox cboRebuilt 
      Height          =   315
      Left            =   11880
      TabIndex        =   23
      Tag             =   "xyf"
      Top             =   0
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.TextBox txtFSCamt 
      Height          =   285
      Left            =   3000
      TabIndex        =   22
      Top             =   8760
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxSum 
      Height          =   7095
      Left            =   120
      TabIndex        =   13
      Tag             =   "XYF"
      Top             =   1440
      Width           =   14415
      _ExtentX        =   25426
      _ExtentY        =   12515
      _Version        =   393216
      RowHeightMin    =   285
      SelectionMode   =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Tag             =   "XY"
      Top             =   720
      Width           =   14415
      Begin VB.TextBox txtFind 
         Height          =   285
         Left            =   12360
         TabIndex        =   5
         Tag             =   "XYF"
         Top             =   195
         Width           =   1935
      End
      Begin lvButton.lvButtons_H cmdFind 
         Height          =   375
         Left            =   11040
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
         Image           =   "FrmSummary.frx":0B34
         cBack           =   -2147483633
      End
      Begin MSComCtl2.DTPicker dFrom 
         Height          =   315
         Left            =   1440
         TabIndex        =   7
         Top             =   180
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         Format          =   130154497
         CurrentDate     =   39322
      End
      Begin lvButton.lvButtons_H cmdRefresh 
         Height          =   375
         Left            =   5400
         TabIndex        =   14
         ToolTipText     =   "Execute"
         Top             =   150
         Visible         =   0   'False
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
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
         Image           =   "FrmSummary.frx":10CE
         cBack           =   -2147483633
      End
      Begin MSComCtl2.DTPicker dTo 
         Height          =   315
         Left            =   3720
         TabIndex        =   15
         Top             =   180
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         Format          =   130154497
         CurrentDate     =   39322
      End
      Begin lvButton.lvButtons_H cmdPrev 
         Height          =   375
         Left            =   5880
         TabIndex        =   16
         ToolTipText     =   "Execute"
         Top             =   150
         Visible         =   0   'False
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
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
         Image           =   "FrmSummary.frx":1668
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdNext 
         Height          =   375
         Left            =   3240
         TabIndex        =   12
         ToolTipText     =   "Execute"
         Top             =   150
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
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
         Image           =   "FrmSummary.frx":1C02
         cBack           =   -2147483633
      End
      Begin VB.Label lblTO 
         Alignment       =   2  'Center
         Caption         =   "TO"
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
         Left            =   3200
         TabIndex        =   28
         Top             =   210
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Date Covered:"
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
         TabIndex        =   11
         Top             =   210
         Width           =   1335
      End
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   3
      Left            =   3600
      TabIndex        =   10
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
      Image           =   "FrmSummary.frx":219C
      cBack           =   -2147483633
   End
   Begin MSFlexGridLib.MSFlexGrid flxInvCancel 
      Height          =   735
      Left            =   120
      TabIndex        =   17
      Top             =   9120
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   1296
      _Version        =   393216
      FixedCols       =   0
      FormatString    =   "REFERENCE NO.|^REF. TYPE"
   End
   Begin MSFlexGridLib.MSFlexGrid flxCancel 
      Height          =   1695
      Left            =   960
      TabIndex        =   18
      Top             =   9120
      Visible         =   0   'False
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   2990
      _Version        =   393216
      Rows            =   1
      Cols            =   6
      FixedCols       =   0
      FormatString    =   "Part Number|<Adj. Type|< IN      |< OUT   |<Source | SPID"
   End
   Begin MSFlexGridLib.MSFlexGrid flxCancelPuRet 
      Height          =   1695
      Left            =   5640
      TabIndex        =   19
      Top             =   9120
      Visible         =   0   'False
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   2990
      _Version        =   393216
      Rows            =   1
      Cols            =   7
      FixedCols       =   0
      FormatString    =   "Qty|<Unit|<Part Number|<Description |<Price | Amount|<SPID"
   End
   Begin MSFlexGridLib.MSFlexGrid flxCancelSalRet 
      Height          =   1695
      Left            =   3000
      TabIndex        =   20
      Top             =   9120
      Visible         =   0   'False
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   2990
      _Version        =   393216
      Rows            =   1
      Cols            =   7
      FixedCols       =   0
      FormatString    =   "Qty|<Unit|<Part Number|<Description |<Price | Amount|<SPID"
   End
   Begin MSFlexGridLib.MSFlexGrid FlxPrev 
      Height          =   1095
      Left            =   9000
      TabIndex        =   21
      Top             =   9240
      Visible         =   0   'False
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   1931
      _Version        =   393216
      Rows            =   1
      FixedCols       =   0
      FormatString    =   "SPID|<PREVIOUS LOCATION "
   End
   Begin lvButton.lvButtons_H cmdExportTRA 
      Height          =   375
      Left            =   5400
      TabIndex        =   24
      Top             =   0
      Visible         =   0   'False
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   661
      Caption         =   "Import TRA"
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
      Image           =   "FrmSummary.frx":2736
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdSaveExcel 
      Height          =   375
      Left            =   1800
      TabIndex        =   27
      Top             =   0
      Visible         =   0   'False
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   661
      Caption         =   "&Save to Excel"
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
      Image           =   "FrmSummary.frx":3148
      cBack           =   -2147483633
   End
   Begin MSComDlg.CommonDialog dlgDEMO 
      Left            =   11400
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   360
      Top             =   8640
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   14880
      m_MDI_Height    =   9285
   End
   Begin VB.Label lblHeading 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Form SUMMARY"
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
      Left            =   9960
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
      Top             =   480
      Width           =   2610
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
      Caption         =   "Form SUMMARY"
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
      Left            =   9975
      TabIndex        =   3
      Top             =   480
      Width           =   4500
   End
   Begin VB.Menu mnuSum 
      Caption         =   "Summary"
      Visible         =   0   'False
      Begin VB.Menu mnuCancel 
         Caption         =   "Cancel Job Order"
      End
   End
   Begin VB.Menu mnuPOmonitor 
      Caption         =   "PO MONITOR"
      Visible         =   0   'False
      Begin VB.Menu mnuMonitorPO 
         Caption         =   "Monitor Purchase Order"
      End
   End
   Begin VB.Menu mnuDepositValidation 
      Caption         =   "DEPOSIT SLIP"
      Visible         =   0   'False
      Begin VB.Menu mnuDeptSlipValid 
         Caption         =   "Deposit Slip Validation"
      End
   End
End
Attribute VB_Name = "FrmSummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsSummary As New ADODB.Recordset
Dim ActRow As Integer, fRow As Integer, fCol As Integer, TCol As Integer, AccID As Integer, ColPrn As Integer
Dim Qry As String, ShowDateRange As Boolean
Public mySrc As String
Dim myDCode As String

Public Sub RefreshRecord(dCode As Integer)
    Dim i As Integer
    myDCode = Val(dCode)
    
    If mySrc = "Job Order" Then
        If dCode = 1 Then
            '***Access
            'Qry = "SELECT SRVCJOBORDER.DATEPROMISED, SRVCJOBORDER.JOBORDERNO, SRVCJOBORDER.CUSTOMER, SRVCUSTOMER.ADDRESS, SRVCJOBORDER.PAYTYPE, SRVCJOBORDER.BALANCE, SRVCJOBORDER.ORNO, SRVCUSTOMER.MODEL, SRVCJOBORDER.KMREADING, SRVCJOBORDER.JOTYPE, SRVCJOBORDER.MECHANIC, SRVCJOBORDER.PRINTCOPY, SRVCJOBORDER.REMARKS, SRVCJOBORDER.JOBORDERID " & _
                  "FROM SRVCJOBORDER INNER JOIN SRVCUSTOMER ON SRVCJOBORDER.ACCOUNTNO = SRVCUSTOMER.ACCOUNTNO Where (((SRVCJOBORDER.DATEPROMISED) = #" & dFrom.Value & "#)) ORDER BY SRVCJOBORDER.DATEPROMISED, SRVCJOBORDER.JOBORDERNO;"
            'mySQL
            Qry = "SELECT SRVCJOBORDER.DATEPROMISED, SRVCJOBORDER.JOBORDERNO, SRVCJOBORDER.CUSTOMER, SRVCUSTOMER.ADDRESS, SRVCJOBORDER.PAYTYPE, SRVCJOBORDER.BALANCE, SRVCJOBORDER.ORNO, SRVCUSTOMER.MODEL, SRVCJOBORDER.KMREADING, SRVCJOBORDER.JOTYPE, SRVCJOBORDER.MECHANIC, SRVCJOBORDER.PRINTCOPY, SRVCJOBORDER.REMARKS, SRVCJOBORDER.JOBORDERID " & _
                  "FROM SRVCJOBORDER INNER JOIN SRVCUSTOMER ON SRVCJOBORDER.ACCOUNTNO = SRVCUSTOMER.ACCOUNTNO Where SRVCJOBORDER.DATEPROMISED = '" & Format(Dfrom.Value, "yyyy-mm-dd") & "' ORDER BY SRVCJOBORDER.DATEPROMISED, SRVCJOBORDER.JOBORDERNO;"
            
        Else
            '***Access
            'Qry = "SELECT SRVCJOBORDER.DATEPROMISED, SRVCJOBORDER.JOBORDERNO, SRVCJOBORDER.CUSTOMER, SRVCUSTOMER.ADDRESS, SRVCJOBORDER.PAYTYPE, SRVCJOBORDER.BALANCE, SRVCJOBORDER.ORNO, SRVCUSTOMER.MODEL, SRVCJOBORDER.KMREADING, SRVCJOBORDER.JOTYPE, SRVCJOBORDER.MECHANIC, SRVCJOBORDER.PRINTCOPY, SRVCJOBORDER.REMARKS, SRVCJOBORDER.JOBORDERID " & _
                  "FROM SRVCJOBORDER INNER JOIN SRVCUSTOMER ON SRVCJOBORDER.ACCOUNTNO = SRVCUSTOMER.ACCOUNTNO Where (((SRVCJOBORDER.DATEPROMISED) >= #" & dFrom.Value & "# And (SRVCJOBORDER.DATEPROMISED) <= #" & dTo.Value & "#)) ORDER BY SRVCJOBORDER.DATEPROMISED, SRVCJOBORDER.JOBORDERNO;"
            'mySQL
            Qry = "SELECT SRVCJOBORDER.DATEPROMISED, SRVCJOBORDER.JOBORDERNO, SRVCJOBORDER.CUSTOMER, SRVCUSTOMER.ADDRESS, SRVCJOBORDER.PAYTYPE, SRVCJOBORDER.BALANCE, SRVCJOBORDER.ORNO, SRVCUSTOMER.MODEL, SRVCJOBORDER.KMREADING, SRVCJOBORDER.JOTYPE, SRVCJOBORDER.MECHANIC, SRVCJOBORDER.PRINTCOPY, SRVCJOBORDER.REMARKS, SRVCJOBORDER.JOBORDERID " & _
                  "FROM SRVCJOBORDER INNER JOIN SRVCUSTOMER ON SRVCJOBORDER.ACCOUNTNO = SRVCUSTOMER.ACCOUNTNO Where SRVCJOBORDER.DATEPROMISED >= '" & Format(Dfrom.Value, "yyyy-mm-dd") & "' And SRVCJOBORDER.DATEPROMISED <= '" & Format(Dto.Value, "yyyy-mm-dd") & "' ORDER BY SRVCJOBORDER.DATEPROMISED, SRVCJOBORDER.JOBORDERNO;"
        
        End If
        
    ElseIf mySrc = "SP Incoming" Or mySrc = "SP Incoming Supplier" Then
        If dCode = 1 Then
            Qry = "SELECT DATE_FORMAT(SPRECEIVE.DATEREC,'%c/%d/%Y'), SPRECEIVE.PACKINGNO, SPRECEIVE.SUPPLIER, SPRECEIVE.SHIPPER, DATE_FORMAT(SPRECEIVE.PACKDATE,'%c/%d/%Y'), SPRECEIVE.REMARKS, SPRECEIVE.CHECKEDBY, SPRECEIVE.PRINTCOPY, SPRECEIVE.SPRRNO " & _
                  "From SPRECEIVE Where SPRECEIVE.DATEREC = '" & Format(Dfrom.Value, "yyyy-mm-dd") & "' ORDER BY SPRECEIVE.DATEREC, SPRECEIVE.PACKINGNO DESC;"
        Else
            Qry = "SELECT DATE_FORMAT(SPRECEIVE.DATEREC,'%c/%d/%Y'), SPRECEIVE.PACKINGNO, SPRECEIVE.SUPPLIER, SPRECEIVE.SHIPPER, DATE_FORMAT(SPRECEIVE.PACKDATE,'%c/%d/%Y'), SPRECEIVE.REMARKS, SPRECEIVE.CHECKEDBY, SPRECEIVE.PRINTCOPY, SPRECEIVE.SPRRNO " & _
                  "From SPRECEIVE Where SPRECEIVE.DATEREC >= '" & Format(Dfrom.Value, "yyyy-mm-dd") & "' And SPRECEIVE.DATEREC <= '" & Format(Dto.Value, "yyyy-mm-dd") & "' ORDER BY SPRECEIVE.DATEREC, SPRECEIVE.PACKINGNO DESC;"
        End If
        
    ElseIf mySrc = "Clerical Adjustment" Then
        If dCode = 1 Then
            '***Access
            'Qry = "SELECT ADJCLERICAL.ADJDATE, ADJCLERICAL.ADJNO, ADJCLERICAL.ITEMADJUSTED, ADJCLERICAL.FIELDADJUSTED, ADJCLERICAL.REFNO, ADJCLERICAL.OLDVALUE, ADJCLERICAL.NEWVALUE, ADJCLERICAL.REASON, ADJCLERICAL.REMARKS, ADJCLERICAL.ENCODER, ADJCLERICAL.PRINTCOPY " & _
                  "From ADJCLERICAL Where (((ADJCLERICAL.ADJDATE) = #" & dFrom.Value & "#)) ORDER BY ADJCLERICAL.ADJDATE, ADJCLERICAL.ADJNO DESC;"
            'mySQL
            Qry = "SELECT ADJCLERICAL.ADJDATE, ADJCLERICAL.ADJNO, ADJCLERICAL.ITEMADJUSTED, ADJCLERICAL.FIELDADJUSTED, ADJCLERICAL.REFNO, ADJCLERICAL.OLDVALUE, ADJCLERICAL.NEWVALUE, ADJCLERICAL.REASON, ADJCLERICAL.REMARKS, ADJCLERICAL.ENCODER, ADJCLERICAL.PRINTCOPY " & _
                  "From ADJCLERICAL Where ADJCLERICAL.ADJDATE = '" & Format(Dfrom.Value, "yyyy-mm-dd") & "' ORDER BY ADJCLERICAL.ADJDATE, ADJCLERICAL.ADJNO DESC;"
            
        Else
            '***Access
            'Qry = "SELECT ADJCLERICAL.ADJDATE, ADJCLERICAL.ADJNO, ADJCLERICAL.ITEMADJUSTED, ADJCLERICAL.FIELDADJUSTED, ADJCLERICAL.REFNO, ADJCLERICAL.OLDVALUE, ADJCLERICAL.NEWVALUE, ADJCLERICAL.REASON, ADJCLERICAL.REMARKS, ADJCLERICAL.ENCODER, ADJCLERICAL.PRINTCOPY " & _
                  "From ADJCLERICAL Where (((ADJCLERICAL.ADJDATE) >= #" & dFrom.Value & "# And (ADJCLERICAL.ADJDATE) <= #" & dTo.Value & "#)) ORDER BY ADJCLERICAL.ADJDATE, ADJCLERICAL.ADJNO DESC;"
            'mySQL
            Qry = "SELECT ADJCLERICAL.ADJDATE, ADJCLERICAL.ADJNO, ADJCLERICAL.ITEMADJUSTED, ADJCLERICAL.FIELDADJUSTED, ADJCLERICAL.REFNO, ADJCLERICAL.OLDVALUE, ADJCLERICAL.NEWVALUE, ADJCLERICAL.REASON, ADJCLERICAL.REMARKS, ADJCLERICAL.ENCODER, ADJCLERICAL.PRINTCOPY " & _
                  "From ADJCLERICAL Where ADJCLERICAL.ADJDATE >= '" & Format(Dfrom.Value, "yyyy-mm-dd") & "' And ADJCLERICAL.ADJDATE <= '" & Format(Dto.Value, "yyyy-mm-dd") & "' ORDER BY ADJCLERICAL.ADJDATE, ADJCLERICAL.ADJNO DESC;"
        
        End If
    ElseIf mySrc = "DMCM [A/R]" Then
        If dCode = 1 Then
            '***Access
            'Qry = "SELECT ADJDMCM.ADJDATE, ADJDMCM.ADJNO, ADJDMCM.CUSTOMER, ADJDMCM.TRANSTYPE, ADJDMCM.REFNO, ADJDMCM.DEBIT, ADJDMCM.CREDIT, ADJDMCM.BALANCE, ADJDMCM.STATUS, ADJDMCM.REASON, ADJDMCM.PRINTCOPY From ADJDMCM " & _
                  "Where (((ADJDMCM.ADJDATE) = #" & dFrom.Value & "#)) ORDER BY ADJDMCM.ADJDATE DESC , ADJDMCM.ADJNO DESC;"
            'mySQL
            Qry = "SELECT ADJDMCM.ADJDATE, ADJDMCM.ADJNO, ADJDMCM.CUSTOMER, ADJDMCM.TRANSTYPE, ADJDMCM.REFNO, ADJDMCM.DEBIT, ADJDMCM.CREDIT, ADJDMCM.BALANCE, ADJDMCM.STATUS, ADJDMCM.REASON, ADJDMCM.PRINTCOPY From ADJDMCM " & _
                  "Where ADJDMCM.ADJDATE = '" & Format(Dfrom.Value, "yyyy-mm-dd") & "' ORDER BY ADJDMCM.ADJDATE DESC , ADJDMCM.ADJNO DESC;"
            'Qry = "SELECT " & _
                    "ADJDMCM.ADJDATE, " & _
                    "ADJDMCM.ADJNO, " & _
                    "ADJDMCM.CUSTOMER, " & _
                    "ADJDMCM.TRANSTYPE, " & _
                    "ADJDMCM.REFNO, " & _
                    "spinvoice.INVDATE, " & _
                    "IF(ADJDMCM.DEBIT=0,' ',ADJDMCM.DEBIT), " & _
                    "IF(ADJDMCM.CREDIT=0,' ',ADJDMCM.CREDIT), " & _
                    "ADJDMCM.BALANCE, " & _
                    "ADJDMCM.`STATUS`, " & _
                    "ADJDMCM.REASON, " & _
                    "ADJDMCM.PRINTCOPY " & _
                "From " & _
                    "spinvoice " & _
                    "INNER JOIN ADJDMCM ON (spinvoice.SPINVNO = ADJDMCM.REFNO) " & _
                "Where " & _
                    "ADJDMCM.ADJDATE = '" & Format(dFrom.Value, "yyyy-mm-dd") & "' " & _
                "Order By " & _
                    "ADJDMCM.ADJDATE DESC, " & _
                    "ADJDMCM.ADJNO Desc"

        Else
            '***Access
            'Qry = "SELECT ADJDMCM.ADJDATE, ADJDMCM.ADJNO, ADJDMCM.CUSTOMER, ADJDMCM.TRANSTYPE, ADJDMCM.REFNO, ADJDMCM.DEBIT, ADJDMCM.CREDIT, ADJDMCM.BALANCE, ADJDMCM.STATUS, ADJDMCM.REASON, ADJDMCM.PRINTCOPY From ADJDMCM " & _
                  "Where (((ADJDMCM.ADJDATE) >= #" & dFrom.Value & "# And (ADJDMCM.ADJDATE) <= #" & dTo.Value & "#)) ORDER BY ADJDMCM.ADJDATE DESC , ADJDMCM.ADJNO DESC;"
            'mySQL
            Qry = "SELECT ADJDMCM.ADJDATE, ADJDMCM.ADJNO, ADJDMCM.CUSTOMER, ADJDMCM.TRANSTYPE, ADJDMCM.REFNO, ADJDMCM.DEBIT, ADJDMCM.CREDIT, ADJDMCM.BALANCE, ADJDMCM.STATUS, ADJDMCM.REASON, ADJDMCM.PRINTCOPY From ADJDMCM " & _
                  "Where ADJDMCM.ADJDATE >= '" & Format(Dfrom.Value, "yyyy-mm-dd") & "' And ADJDMCM.ADJDATE <= '" & Format(Dto.Value, "yyyy-mm-dd") & "' ORDER BY ADJDMCM.ADJDATE DESC , ADJDMCM.ADJNO DESC;"
            'mysql
            'Qry = "SELECT " & _
                    "ADJDMCM.ADJDATE, " & _
                    "ADJDMCM.ADJNO, " & _
                    "ADJDMCM.CUSTOMER, " & _
                    "ADJDMCM.TRANSTYPE, " & _
                    "ADJDMCM.REFNO, " & _
                    "spinvoice.INVDATE, " & _
                    "IF(ADJDMCM.DEBIT=0,' ',ADJDMCM.DEBIT), " & _
                    "IF(ADJDMCM.CREDIT=0,' ',ADJDMCM.CREDIT), " & _
                    "ADJDMCM.BALANCE, " & _
                    "ADJDMCM.`STATUS`, " & _
                    "ADJDMCM.REASON, " & _
                    "ADJDMCM.PRINTCOPY " & _
                "From " & _
                    "spinvoice " & _
                    "INNER JOIN ADJDMCM ON (spinvoice.SPINVNO = ADJDMCM.REFNO) " & _
                "Where " & _
                    "ADJDMCM.ADJDATE >= '" & Format(dFrom.Value, "yyyy-mm-dd") & "' AND " & _
                    "ADJDMCM.ADJDATE <= '" & Format(dTo.Value, "yyyy-mm-dd") & "' " & _
                "Order By " & _
                    "ADJDMCM.ADJDATE DESC, " & _
                    "ADJDMCM.ADJNO Desc"
            
        End If
    
    ElseIf mySrc = "Inventory Adjustment" Then
        If dCode = 1 Then
            '***Access
            'Qry = "SELECT INVNTRYADJ.ADJDATE, INVNTRYADJ.ADJNO, INVNTRYADJ.DOCREF, INVNTRYADJ.REASON, INVNTRYADJ.REMARKS, INVNTRYADJ.PREPBY, INVNTRYADJ.APRVBY, INVNTRYADJ.PRINTCOPY " & _
                  "From InvntryADJ Where (((InvntryADJ.ADJDATE) = #" & dFrom.Value & "#)) ORDER BY INVNTRYADJ.ADJDATE, INVNTRYADJ.ADJNO;"
            'mySQL
            Qry = "SELECT INVNTRYADJ.ADJDATE, INVNTRYADJ.ADJNO, INVNTRYADJ.DOCREF, INVNTRYADJ.REASON, INVNTRYADJ.REMARKS, INVNTRYADJ.PREPBY, INVNTRYADJ.APRVBY, INVNTRYADJ.PRINTCOPY " & _
                  "From InvntryADJ Where InvntryADJ.ADJDATE = '" & Format(Dfrom.Value, "yyy-mm-dd") & "' ORDER BY INVNTRYADJ.ADJDATE, INVNTRYADJ.ADJNO;"
            
        Else
            '***Access
            'Qry = "SELECT INVNTRYADJ.ADJDATE, INVNTRYADJ.ADJNO, INVNTRYADJ.DOCREF, INVNTRYADJ.REASON, INVNTRYADJ.REMARKS, INVNTRYADJ.PREPBY, INVNTRYADJ.APRVBY, INVNTRYADJ.PRINTCOPY " & _
                  "From InvntryADJ Where (((InvntryADJ.ADJDATE) >= #" & dFrom.Value & "# And (InvntryADJ.ADJDATE) <= #" & dTo.Value & "#)) ORDER BY INVNTRYADJ.ADJDATE, INVNTRYADJ.ADJNO;"
            'mySQL
            Qry = "SELECT INVNTRYADJ.ADJDATE, INVNTRYADJ.ADJNO, INVNTRYADJ.DOCREF, INVNTRYADJ.REASON, INVNTRYADJ.REMARKS, INVNTRYADJ.PREPBY, INVNTRYADJ.APRVBY, INVNTRYADJ.PRINTCOPY " & _
                  "From InvntryADJ Where InvntryADJ.ADJDATE >= '" & Format(Dfrom.Value, "yyyy-mm-dd") & "' And InvntryADJ.ADJDATE <= '" & Format(Dto.Value, "yyyy-mm-dd") & "' ORDER BY INVNTRYADJ.ADJDATE, INVNTRYADJ.ADJNO;"
            
            
        End If
    ElseIf mySrc = "Customer P.O." Then
        If dCode = 1 Then
            '***Access
            'Qry = "SELECT CUSTPO.CPODATE, CUSTPO.CPONO, CUSTPO.CUSTOMER, IIf([APPROVED]=True,'APPROVED','PENDING') AS STATUS, CUSTPO.INVOICENO, CUSTPO.REMARKS, CUSTPO.PRINTCOPY, CUSTPO.CPOID " & _
                  "From CUSTPO Where (((CUSTPO.CPODATE) = #" & dFrom.Value & "#)) ORDER BY CUSTPO.CPODATE DESC , CUSTPO.CPONO DESC;"
            'mySQL
            Qry = "SELECT CUSTPO.CPODATE, CUSTPO.CPONO, CUSTPO.CUSTOMER, If(CUSTPO.APPROVED=1,'APPROVED','PENDING') AS STATUS, CUSTPO.INVOICENO, CUSTPO.REMARKS, CUSTPO.PRINTCOPY, CUSTPO.CPOID " & _
                  "From CUSTPO Where CUSTPO.CPODATE = '" & Format(Dfrom.Value, "yyyy-mm-dd") & "' ORDER BY CUSTPO.CPODATE DESC , CUSTPO.CPONO DESC;"
        
        Else
            '***Access
            'Qry = "SELECT CUSTPO.CPODATE, CUSTPO.CPONO, CUSTPO.CUSTOMER, IIf([APPROVED]=True,'APPROVED','PENDING') AS STATUS, CUSTPO.INVOICENO, CUSTPO.REMARKS, CUSTPO.PRINTCOPY, CUSTPO.CPOID " & _
                  "From CUSTPO Where (((CUSTPO.CPODATE) >= #" & dFrom.Value & "# And (CUSTPO.CPODATE) <= #" & dTo.Value & "#)) ORDER BY CUSTPO.CPODATE DESC , CUSTPO.CPONO DESC;"
            'mySQL
            Qry = "SELECT CUSTPO.CPODATE, CUSTPO.CPONO, CUSTPO.CUSTOMER, If(CUSTPO.APPROVED=1,'APPROVED','PENDING') AS STATUS, CUSTPO.INVOICENO, CUSTPO.REMARKS, CUSTPO.PRINTCOPY, CUSTPO.CPOID " & _
                  "From CUSTPO Where CUSTPO.CPODATE >= '" & Format(Dfrom.Value, "yyyy-mm-dd") & "' And CUSTPO.CPODATE <= '" & Format(Dto.Value, "yyyy-mm-dd") & "' ORDER BY CUSTPO.CPODATE DESC , CUSTPO.CPONO DESC;"
            
        End If
    ElseIf mySrc = "Top Selling Spare Parts" Then
        Qry = "SELECT TOPSALESUMM.DATECREATED, TOPSALESUMM.SALESPERIOD, TOPSALESUMM.ITEMTYPE, TOPSALESUMM.TOPVALUE, TOPSALESUMM.NOTE, TOPSALESUMM.TOPSALESID " & _
              "From TOPSALESUMM ORDER BY TOPSALESUMM.DATECREATED DESC;"
      
    ElseIf mySrc = "Purchase Return" Then
        If dCode = 1 Then
            '***Access
            'Qry = "SELECT PURCHASERET.DATERET, PURCHASERET.PURTRNO, PURCHASERET.PACKINGNO, SPRECEIVE.DATEREC, SPRECEIVE.SUPPLIER, PURCHASERET.AMOUNT, PURCHASERET.REMARKS, PURCHASERET.PRINTCOPY, PURCHASERET.PURTRNID " & _
                  "FROM PURCHASERET INNER JOIN SPRECEIVE ON PURCHASERET.PACKINGNO = SPRECEIVE.PACKINGNO WHERE (((PURCHASERET.DATERET)=#" & dFrom.Value & "#)) ORDER BY PURCHASERET.PURTRNO DESC;"
            'mySQL
            Qry = "SELECT PURCHASERET.DATERET, PURCHASERET.PURTRNO, PURCHASERET.PACKINGNO, SPRECEIVE.DATEREC, SPRECEIVE.SUPPLIER, PURCHASERET.AMOUNT, PURCHASERET.REMARKS, PURCHASERET.PRINTCOPY, PURCHASERET.PURTRNID " & _
                  "FROM PURCHASERET INNER JOIN SPRECEIVE ON PURCHASERET.PACKINGNO = SPRECEIVE.PACKINGNO WHERE PURCHASERET.DATERET='" & Format(Dfrom.Value, "yyyy-mm-dd") & "' ORDER BY PURCHASERET.PURTRNO DESC;"
     
        Else
            '***Access
            'Qry = "SELECT PURCHASERET.DATERET, PURCHASERET.PURTRNO, PURCHASERET.PACKINGNO, SPRECEIVE.DATEREC, SPRECEIVE.SUPPLIER, PURCHASERET.AMOUNT, PURCHASERET.REMARKS, PURCHASERET.PRINTCOPY, PURCHASERET.PURTRNID " & _
                  "FROM PURCHASERET INNER JOIN SPRECEIVE ON PURCHASERET.PACKINGNO = SPRECEIVE.PACKINGNO Where (((PURCHASERET.DATERET) >= #" & dFrom.Value & "# And (PURCHASERET.DATERET) <= #" & dTo.Value & "#)) ORDER BY PURCHASERET.PURTRNO DESC;"
            Qry = "SELECT PURCHASERET.DATERET, PURCHASERET.PURTRNO, PURCHASERET.PACKINGNO, SPRECEIVE.DATEREC, SPRECEIVE.SUPPLIER, PURCHASERET.AMOUNT, PURCHASERET.REMARKS, PURCHASERET.PRINTCOPY, PURCHASERET.PURTRNID " & _
                  "FROM PURCHASERET INNER JOIN SPRECEIVE ON PURCHASERET.PACKINGNO = SPRECEIVE.PACKINGNO Where PURCHASERET.DATERET >= '" & Format(Dfrom.Value, "yyyy-mm-dd") & "' And PURCHASERET.DATERET <= '" & Format(Dto.Value, "yyyy-mm-dd") & "' ORDER BY PURCHASERET.PURTRNO DESC;"
        
        End If
    ElseIf mySrc = "Requisition" Then
        If dCode = 1 Then
            '***Access
            'Qry = "SELECT REQUISITION.REQDATE, REQUISITION.REQNO, REQUISITION.CUSTOMER, REQUISITION.AMOUNT, REQUISITION.JOBNO, REQUISITION.INVOICENO, REQUISITION.ORNO, REQUISITION.REMARKS, REQUISITION.PRINTCOPY, REQUISITION.REQID " & _
                  "From Requisition WHERE (((REQUISITION.REQDATE)=#" & dFrom.Value & "#));"
            'mySQL
            Qry = "SELECT REQUISITION.REQDATE, REQUISITION.REQNO, REQUISITION.CUSTOMER, REQUISITION.AMOUNT, REQUISITION.JOBNO, REQUISITION.INVOICENO, REQUISITION.ORNO, REQUISITION.REMARKS, REQUISITION.PRINTCOPY, REQUISITION.REQID " & _
                  "From Requisition WHERE REQUISITION.REQDATE='" & Format(Dfrom.Value, "yyyy-mm-dd") & "';"
            
        Else
            '***Access
            'Qry = "SELECT REQUISITION.REQDATE, REQUISITION.REQNO, REQUISITION.CUSTOMER, REQUISITION.AMOUNT, REQUISITION.JOBNO, REQUISITION.INVOICENO, REQUISITION.ORNO, REQUISITION.REMARKS, REQUISITION.PRINTCOPY, REQUISITION.REQID " & _
                  "From Requisition WHERE (((REQUISITION.REQDATE)>=#" & dFrom.Value & "# And (REQUISITION.REQDATE)<=#" & dTo.Value & "#));"
        
            'mySQL
            Qry = "SELECT REQUISITION.REQDATE, REQUISITION.REQNO, REQUISITION.CUSTOMER, REQUISITION.AMOUNT, REQUISITION.JOBNO, REQUISITION.INVOICENO, REQUISITION.ORNO, REQUISITION.REMARKS, REQUISITION.PRINTCOPY, REQUISITION.REQID " & _
                  "From Requisition WHERE REQUISITION.REQDATE>='" & Format(Dfrom.Value, "yyyy-mm-dd") & "' And REQUISITION.REQDATE<='" & Format(Dto.Value, "yyyy-mm-dd") & "';"
        End If
    ElseIf mySrc = "Location Adjustment" Then
        If dCode = 1 Then
            '***Access
            'Qry = "SELECT TRANSFERLOC.DATETRANS, TRANSFERLOC.TRANSNO, TRANSFERLOC.REFERENCE, TRANSFERLOC.REASON, TRANSFERLOC.PREPBY, TRANSFERLOC.APRVBY, TRANSFERLOC.PRINTCOPY, TRANSFERLOC.TRANSLOCID " & _
                  "From TRANSFERLOC Where (((TRANSFERLOC.DATETRANS) = #" & dFrom.Value & "#)) ORDER BY TRANSFERLOC.DATETRANS DESC , TRANSFERLOC.TRANSNO DESC;"
            'mySQL
            Qry = "SELECT TRANSFERLOC.DATETRANS, TRANSFERLOC.TRANSNO, TRANSFERLOC.REFERENCE, TRANSFERLOC.REASON, TRANSFERLOC.PREPBY, TRANSFERLOC.APRVBY, TRANSFERLOC.PRINTCOPY, TRANSFERLOC.TRANSLOCID " & _
                  "From TRANSFERLOC Where TRANSFERLOC.DATETRANS = '" & Format(Dfrom.Value, "yyyy-mm-dd") & "' ORDER BY TRANSFERLOC.DATETRANS DESC , TRANSFERLOC.TRANSNO DESC;"
        Else
            '***Access
            'Qry = "SELECT TRANSFERLOC.DATETRANS, TRANSFERLOC.TRANSNO, TRANSFERLOC.REFERENCE, TRANSFERLOC.REASON, TRANSFERLOC.PREPBY, TRANSFERLOC.APRVBY, TRANSFERLOC.PRINTCOPY, TRANSFERLOC.TRANSLOCID " & _
                  "From TRANSFERLOC Where (((TRANSFERLOC.DATETRANS) >= #" & dFrom.Value & "# And (TRANSFERLOC.DATETRANS) <= #" & dTo.Value & "#)) ORDER BY TRANSFERLOC.DATETRANS DESC , TRANSFERLOC.TRANSNO DESC;"
            'mySQL
            Qry = "SELECT TRANSFERLOC.DATETRANS, TRANSFERLOC.TRANSNO, TRANSFERLOC.REFERENCE, TRANSFERLOC.REASON, TRANSFERLOC.PREPBY, TRANSFERLOC.APRVBY, TRANSFERLOC.PRINTCOPY, TRANSFERLOC.TRANSLOCID " & _
                  "From TRANSFERLOC Where TRANSFERLOC.DATETRANS >= '" & Format(Dfrom.Value, "yyyy-mm-dd") & "' And TRANSFERLOC.DATETRANS <= '" & Format(Dto.Value, "yyyy-mm-dd") & "' ORDER BY TRANSFERLOC.DATETRANS DESC , TRANSFERLOC.TRANSNO DESC;"
            
        End If
    ElseIf mySrc = "Received Spare Parts" Or mySrc = "Received Spare Parts DVOWH" Then
        If dCode = 1 Then
            '***Access
            'Qry = "SELECT SPRECEIVE.DATEREC, SPRECEIVE.PACKINGNO, SPRECEIVE.SUPPLIER, SPRECEIVE.SHIPPER, SPRECEIVE.WAYBILLNO, SPRECEIVE.PACKDATE, SPRECEIVE.REMARKS, SPRECEIVE.CHECKEDBY, SPRECEIVE.PRINTCOPY, SPRECEIVE.SPRRNO " & _
                  "From SPRECEIVE Where (((SPRECEIVE.DATEREC) = #" & dFrom.Value & "#)) ORDER BY SPRECEIVE.DATEREC, SPRECEIVE.PACKINGNO DESC;"
            'mySQl
            'Qry = "SELECT SPRECEIVE.DATEREC, SPRECEIVE.PACKINGNO, SPRECEIVE.SUPPLIER, SPRECEIVE.SHIPPER, SPRECEIVE.POREF, SPRECEIVE.PACKDATE, SPRECEIVE.REMARKS, SPRECEIVE.CHECKEDBY, SPRECEIVE.PRINTCOPY, SPRECEIVE.SPRRNO, SPRECEIVE.AUTOLOAD  " & _
            '      "From SPRECEIVE Where SPRECEIVE.DATEREC = '" & Format(dFrom.Value, "yyyy-mm-dd") & "' ORDER BY SPRECEIVE.DATEREC, SPRECEIVE.PACKINGNO DESC;"

            Qry = "SELECT SPRECEIVE.DATEREC, SPRECEIVE.PACKINGNO, SPRECEIVE.SUPPLIER, SPRECEIVE.SHIPPER, SPRECEIVE.POREF, SPRECEIVE.PACKDATE, SPRECEIVE.REMARKS, " & _
                          "COUNT(itr.SPAREPARTID) as TOTALPN, SUM(itr.RECEIVE) as TOTALQTY, IFNULL(sprec.DAMAGE,0) as DAMAGE, " & _
                          "IFNULL(sprec.LACKING,0) as LACKING, IFNULL(sprec.EXCESS,0) as EXCESS, IFNULL(sprec.WRONGPN,0) as WRONGPN, " & _
                          "SPRECEIVE.CHECKEDBY , SPRECEIVE.PRINTCOPY, SPRECEIVE.SPRRNO, SPRECEIVE.AutoLoad " & _
                  "From SPRECEIVE " & _
                  "LEFT JOIN SPITR as itr ON itr.REFERENCE = spreceive.PACKINGNO " & _
                  "LEFT JOIN (Select SUM(DAMAGE) as DAMAGE, SUM(LACKING) as LACKING, SUM(EXCESS) as EXCESS, SUM(WRONGPN) as WRONGPN, SPAREPARTID, PACKINGNO FROM sprecdiscrepancy GROUP BY PACKINGNO) as sprec ON sprec.packingNO = itr.REFERENCE " & _
                  "Where SPRECEIVE.DATEREC = '" & Format(Dfrom.Value, "yyyy-mm-dd") & "' " & _
                  "GROUP BY itr.REFERENCE ORDER BY SPRECEIVE.DATEREC, SPRECEIVE.PACKINGNO DESC "
        Else
            '***Access
            'Qry = "SELECT SPRECEIVE.DATEREC, SPRECEIVE.PACKINGNO, SPRECEIVE.SUPPLIER, SPRECEIVE.SHIPPER, SPRECEIVE.WAYBILLNO, SPRECEIVE.PACKDATE, SPRECEIVE.REMARKS, SPRECEIVE.CHECKEDBY, SPRECEIVE.PRINTCOPY, SPRECEIVE.SPRRNO " & _
                  "From SPRECEIVE Where (((SPRECEIVE.DATEREC) >= #" & dFrom.Value & "# And (SPRECEIVE.DATEREC) <= #" & dTo.Value & "#)) ORDER BY SPRECEIVE.DATEREC, SPRECEIVE.PACKINGNO DESC;"
            'mySQL
            'Qry = "SELECT SPRECEIVE.DATEREC, SPRECEIVE.PACKINGNO, SPRECEIVE.SUPPLIER, SPRECEIVE.SHIPPER, SPRECEIVE.POREF, SPRECEIVE.PACKDATE, SPRECEIVE.REMARKS, SPRECEIVE.CHECKEDBY, SPRECEIVE.PRINTCOPY, SPRECEIVE.SPRRNO, SPRECEIVE.AUTOLOAD " & _
            '      "From SPRECEIVE Where SPRECEIVE.DATEREC >= '" & Format(dFrom.Value, "yyyy-mm-dd") & "' And SPRECEIVE.DATEREC <= '" & Format(dTo.Value, "yyyy-mm-dd") & "' ORDER BY SPRECEIVE.DATEREC, SPRECEIVE.PACKINGNO DESC;"
            Qry = "SELECT SPRECEIVE.DATEREC, SPRECEIVE.PACKINGNO, SPRECEIVE.SUPPLIER, SPRECEIVE.SHIPPER, SPRECEIVE.POREF, SPRECEIVE.PACKDATE, SPRECEIVE.REMARKS, " & _
                          "COUNT(itr.SPAREPARTID) as TOTALPN, SUM(itr.RECEIVE) as TOTALQTY, IFNULL(sprec.DAMAGE,0) as DAMAGE, " & _
                          "IFNULL(sprec.LACKING,0) as LACKING, IFNULL(sprec.EXCESS,0) as EXCESS, IFNULL(sprec.WRONGPN,0) as WRONGPN, " & _
                          "SPRECEIVE.CHECKEDBY , SPRECEIVE.PRINTCOPY, SPRECEIVE.SPRRNO, SPRECEIVE.AutoLoad " & _
                  "From SPRECEIVE " & _
                  "LEFT JOIN SPITR as itr ON itr.REFERENCE = spreceive.PACKINGNO " & _
                  "LEFT JOIN (Select SUM(DAMAGE) as DAMAGE, SUM(LACKING) as LACKING, SUM(EXCESS) as EXCESS, SUM(WRONGPN) as WRONGPN, SPAREPARTID, PACKINGNO FROM sprecdiscrepancy GROUP BY PACKINGNO) as sprec ON sprec.packingNO = itr.REFERENCE " & _
                  "Where SPRECEIVE.DATEREC >= '" & Format(Dfrom.Value, "yyyy-mm-dd") & "' And SPRECEIVE.DATEREC <= '" & Format(Dto.Value, "yyyy-mm-dd") & "' " & _
                  "GROUP BY itr.REFERENCE ORDER BY SPRECEIVE.DATEREC, SPRECEIVE.PACKINGNO DESC "
                  
        End If
    ElseIf mySrc = "Cashier Cash Flow" Then
        If dCode = 1 Then
            Qry = "SELECT DATE_FORMAT(cashflow.CFDATE,'%c/%d/%Y'), cashflow.PREPAREDBY, cashflow.APPROVEDBY, cashflow.REMARKS, cashflow.CASHFLOWID " & _
                  "From cashflow Where cashflow.CFDATE = '" & Format(Dfrom.Value, "YYYY-MM-DD") & "' Order By cashflow.CFDATE Desc;"
        Else
            Qry = "SELECT DATE_FORMAT(cashflow.CFDATE,'%c/%d/%Y'), cashflow.PREPAREDBY, cashflow.APPROVEDBY, cashflow.REMARKS, cashflow.CASHFLOWID " & _
                  "From cashflow Where cashflow.CFDATE >= '" & Format(Dfrom.Value, "YYYY-MM-DD") & "' and cashflow.CFDATE <= '" & Format(Dto.Value, "YYYY-MM-DD") & "' Order By cashflow.CFDATE Desc;"
        End If
    
    ElseIf mySrc = "Manager's Cash Flow" Then
        If dCode = 1 Then
            Qry = "SELECT DATE_FORMAT(cashflowmgr.CFDATE,'%c/%d/%Y'), cashflowmgr.PREPAREDBY, cashflowmgr.APPROVEDBY, cashflowmgr.REMARKS, cashflowmgr.MGRCFID, valid.TtlDeposit " & _
                  "From cashflowmgr Left Join (Select MGRCFID, COUNT(MGRCFID) as TtlDeposit From depositslipinfo GROUP BY MGRCFID) as valid ON valid.MGRCFID = cashflowmgr.MGRCFID Where cashflowmgr.CFDATE = '" & Format(Dfrom.Value, "YYYY-MM-DD") & "' Order By cashflowmgr.CFDATE Desc;"
        Else
            Qry = "SELECT DATE_FORMAT(cashflowmgr.CFDATE,'%c/%d/%Y'), cashflowmgr.PREPAREDBY, cashflowmgr.APPROVEDBY, cashflowmgr.REMARKS, cashflowmgr.MGRCFID, valid.TtlDeposit " & _
                  "From cashflowmgr Left Join (Select MGRCFID, COUNT(MGRCFID) as TtlDeposit From depositslipinfo GROUP BY MGRCFID) as valid ON valid.MGRCFID = cashflowmgr.MGRCFID Where cashflowmgr.CFDATE >= '" & Format(Dfrom.Value, "YYYY-MM-DD") & "' and cashflowmgr.CFDATE <= '" & Format(Dto.Value, "YYYY-MM-DD") & "' Order By cashflowmgr.CFDATE Desc;"
        End If
    
    ElseIf mySrc = "Sales Order" Then
        If dCode = 1 Then
            Qry = "SELECT DATE_FORMAT(spinvoice.INVDATE,'%c/%d/%Y'), spinvoice.SPINVNO, spinvoice.INVTYPE, spinvoice.CHARGETO, Format(spinvoice.AMOUNT,2), spinvoice.REMARKS, spsalesorder.ORDERNO, spsalesorder.FRONTLINER " & _
                "From spsalesorder INNER JOIN spinvoice ON (spsalesorder.SPSORDERID = spinvoice.SPSORDERID) " & _
                "Where spinvoice.INVDATE = '" & Format(Dfrom.Value, "yyyy-mm-dd") & "' AND spinvoice.REMARKS <> 'CANCELLED' Order By spsalesorder.FRONTLINER, spinvoice.INVDATE, spinvoice.SPINVNO"
            
        Else
            Qry = "SELECT DATE_FORMAT(spinvoice.INVDATE,'%c/%d/%Y'), spinvoice.SPINVNO, spinvoice.INVTYPE, spinvoice.CHARGETO, Format(spinvoice.AMOUNT,2), spinvoice.REMARKS, spsalesorder.ORDERNO, spsalesorder.FRONTLINER " & _
                "From spsalesorder INNER JOIN spinvoice ON (spsalesorder.SPSORDERID = spinvoice.SPSORDERID) " & _
                "Where spinvoice.INVDATE >= '" & Format(Dfrom.Value, "yyyy-mm-dd") & "' AND spinvoice.INVDATE <= '" & Format(Dto.Value, "yyyy-mm-dd") & "' AND spinvoice.REMARKS <> 'CANCELLED' Order By spsalesorder.FRONTLINER, spinvoice.INVDATE, spinvoice.SPINVNO;"
        End If
    ElseIf mySrc = "Mechanic Income" Then
        If dCode = 1 Then
            '***Access
            'Qry = "SELECT SRVCJOBORDER.MECHANIC, SRVCJOBORDER.JOBORDERNO, SRVCJOBORDER.DATEPROMISED, SRVCJOBORDER.JOTYPE, SRVCJOBORDER.TOTALCOST From SRVCJOBORDER " & _
                  "Where (((SRVCJOBORDER.DATEPROMISED) = #" & dFrom.Value & "#) And ((SRVCJOBORDER.PRINTCOPY) > 0)) ORDER BY SRVCJOBORDER.MECHANIC, SRVCJOBORDER.DATEPROMISED;"
            'mySQl
            Qry = "SELECT SRVCJOBORDER.MECHANIC, SRVCJOBORDER.JOBORDERNO, SRVCJOBORDER.DATEPROMISED, SRVCJOBORDER.JOTYPE, SRVCJOBORDER.TOTALCOST From SRVCJOBORDER " & _
                  "Where SRVCJOBORDER.DATEPROMISED = '" & Format(Dfrom.Value, "yyyy-mm-dd") & "' And SRVCJOBORDER.PRINTCOPY > 0 And SRVCJOBORDER.REMARKS <> 'CANCELLED' ORDER BY SRVCJOBORDER.MECHANIC, SRVCJOBORDER.DATEPROMISED;"
            
        Else
            '***Access
            'Qry = "SELECT SRVCJOBORDER.MECHANIC, SRVCJOBORDER.JOBORDERNO, SRVCJOBORDER.DATEPROMISED, SRVCJOBORDER.JOTYPE, SRVCJOBORDER.TOTALCOST From SRVCJOBORDER " & _
                  "Where (((SRVCJOBORDER.DATEPROMISED) >= #" & dFrom.Value & "# And (SRVCJOBORDER.DATEPROMISED) <= #" & dTo.Value & "#) And ((SRVCJOBORDER.PRINTCOPY) > 0)) ORDER BY SRVCJOBORDER.MECHANIC, SRVCJOBORDER.DATEPROMISED;"
            'mySQL
            Qry = "SELECT SRVCJOBORDER.MECHANIC, SRVCJOBORDER.JOBORDERNO, SRVCJOBORDER.DATEPROMISED, SRVCJOBORDER.JOTYPE, SRVCJOBORDER.TOTALCOST From SRVCJOBORDER " & _
                  "Where SRVCJOBORDER.DATEPROMISED >= '" & Format(Dfrom.Value, "yyyy-mm-dd") & "' And SRVCJOBORDER.DATEPROMISED <= '" & Format(Dto.Value, "yyyy-mm-dd") & "' And SRVCJOBORDER.PRINTCOPY > 0 And SRVCJOBORDER.REMARKS <> 'CANCELLED' ORDER BY SRVCJOBORDER.MECHANIC, SRVCJOBORDER.DATEPROMISED;"
            
        End If
    ElseIf mySrc = "Cash Voucher" Then
        If dCode = 1 Then
            '***Access
            'Qry = "SELECT CASHVOUCHER.CVDATE, CASHVOUCHER.VOUCHERNO, CASHVOUCHER.PAIDTO, CASHVOUCHER.POSITION, CASHVOUCHER.PARTICULAR, CASHVOUCHER.AMOUNT, CASHVOUCHER.REMARKS, CASHVOUCHER.PRINTCOPY " & _
                  "From CashVoucher Where (((CashVoucher.CVDate) = #" & dFrom & "#)) ORDER BY CASHVOUCHER.CVDATE DESC , CASHVOUCHER.VOUCHERNO;"
            'mySQL
            Qry = "SELECT CASHVOUCHER.CVDATE, CASHVOUCHER.VOUCHERNO, CASHVOUCHER.PAIDTO, CASHVOUCHER.POSITION, CASHVOUCHER.PARTICULAR, CASHVOUCHER.AMOUNT, CASHVOUCHER.REMARKS, CASHVOUCHER.PRINTCOPY " & _
                  "From CashVoucher Where (((CashVoucher.CVDate) = '" & Format(Dfrom, "yyyy-mm-dd") & "')) ORDER BY CASHVOUCHER.CVDATE DESC , CASHVOUCHER.VOUCHERNO;"
            
        Else
            '***Access
            'Qry = "SELECT CASHVOUCHER.CVDATE, CASHVOUCHER.VOUCHERNO, CASHVOUCHER.PAIDTO, CASHVOUCHER.POSITION, CASHVOUCHER.PARTICULAR, CASHVOUCHER.AMOUNT, CASHVOUCHER.REMARKS, CASHVOUCHER.PRINTCOPY From CASHVOUCHER " & _
                  "Where (((CASHVOUCHER.CVDate) >= #" & dFrom & "# And (CASHVOUCHER.CVDate) <= #" & dTo & "#)) ORDER BY CASHVOUCHER.CVDATE DESC, CASHVOUCHER.VOUCHERNO;"
            'mySQL
            Qry = "SELECT CASHVOUCHER.CVDATE, CASHVOUCHER.VOUCHERNO, CASHVOUCHER.PAIDTO, CASHVOUCHER.POSITION, CASHVOUCHER.PARTICULAR, CASHVOUCHER.AMOUNT, CASHVOUCHER.REMARKS, CASHVOUCHER.PRINTCOPY From CASHVOUCHER " & _
                  "Where (((CASHVOUCHER.CVDate) >= '" & Format(Dfrom, "yyyy-mm-dd") & "' And (CASHVOUCHER.CVDate) <= '" & Format(Dto, "yyyy-mm-dd") & "')) ORDER BY CASHVOUCHER.CVDATE DESC, CASHVOUCHER.VOUCHERNO;"
            
        End If
    
    ElseIf mySrc = "Advances for Liquidation" Then
        If dCode = 1 Then
            Qry = "SELECT date_format(advliquidation.AFLDATE,'%c/%d/%Y'), advliquidation.AFLNO, format(advliquidation.AMTREC,2), format(advliquidation.AMTRETURN,2), format(advliquidation.AFLAMOUNT,2), advliquidation.PREPAREDBY, advliquidation.REMARKS, advliquidation.PRINTCOPY, advliquidation.AFLID " & _
                  "From advliquidation Where advliquidation.AFLDATE = '" & Format(Dfrom.Value, "YYYY-MM-DD") & "' Order By advliquidation.AFLNO;"
        Else
            Qry = "SELECT date_format(advliquidation.AFLDATE,'%c/%d/%Y'), advliquidation.AFLNO, format(advliquidation.AMTREC,2), format(advliquidation.AMTRETURN,2), format(advliquidation.AFLAMOUNT,2), advliquidation.PREPAREDBY, advliquidation.REMARKS, advliquidation.PRINTCOPY, advliquidation.AFLID " & _
                  "From advliquidation Where advliquidation.AFLDATE >= '" & Format(Dfrom.Value, "YYYY-MM-DD") & "' and advliquidation.AFLDATE <= '" & Format(Dto.Value, "YYYY-MM-DD") & "' Order By advliquidation.AFLNO;"
        End If
        
    ElseIf mySrc = "Cash Return" Then
        If dCode = 1 Then
            Qry = "SELECT date_format(cashreturn.CRDATE,'%c/%d/%Y'), cashreturn.CASHRETNO, format(cashreturn.CRAMOUNT,2), cashreturn.REMARKS, cashreturn.PREPAREDBY, cashreturn.PRINTCOPY, cashreturn.CRID " & _
                  "From cashreturn Where cashreturn.CRDATE = '" & Format(Dfrom.Value, "YYYY-MM-DD") & "';"
        Else
            Qry = "SELECT date_format(cashreturn.CRDATE,'%c/%d/%Y'), cashreturn.CASHRETNO, format(cashreturn.CRAMOUNT,2), cashreturn.REMARKS, cashreturn.PREPAREDBY, cashreturn.PRINTCOPY, cashreturn.CRID " & _
                  "From cashreturn Where cashreturn.CRDATE >= '" & Format(Dfrom.Value, "YYYY-MM-DD") & "' and cashreturn.CRDATE <= '" & Format(Dto.Value, "YYYY-MM-DD") & "';"
        End If
    ElseIf mySrc = "Sales Return" Then
        If dCode = 1 Then
            '***Access
            'Qry = "SELECT SRETURNSP.RETDATE, SRETURNSP.RETNO, SRETURNSP.INVOICE, SPINVOICE.CHARGETO, SRETURNSP.REASON, SRETURNSP.TOTALAMT, SRETURNSP.REMARKS, SRETURNSP.PRINTCOPY, SRETURNSP.RETURNID " & _
                  "FROM SRETURNSP INNER JOIN SPINVOICE ON SRETURNSP.INVOICE = SPINVOICE.SPINVNO Where (((SRETURNSP.RETDATE) = #" & dFrom.Value & "#)) ORDER BY SRETURNSP.RETDATE, SRETURNSP.RETNO DESC;"
            'mySQL
            'Qry = "SELECT SRETURNSP.RETDATE, SRETURNSP.RETNO, SRETURNSP.INVOICE, SPINVOICE.CHARGETO, SRETURNSP.REASON, SRETURNSP.TOTALAMT, SRETURNSP.REMARKS, SRETURNSP.PRINTCOPY, SRETURNSP.RETURNID " & _
                  "FROM SRETURNSP INNER JOIN SPINVOICE ON SRETURNSP.INVOICE = SPINVOICE.SPINVNO Where SRETURNSP.RETDATE = '" & Format(dFrom.Value, "yyyy-mm-dd") & "' ORDER BY SRETURNSP.RETDATE, SRETURNSP.RETNO DESC;"
            
            'modify Report 06/08/2010
             
             Qry = "SELECT " & _
                        "SRETURNSP.RETDATE, " & _
                        "SRETURNSP.RETNO, " & _
                        "SRETURNSP.INVOICE, " & _
                        "SPINVOICE.INVDATE, " & _
                        "SPINVOICE.CHARGETO, " & _
                        "Cast(IF(MONTH(SRETURNSP.RETDATE) = MONTH(SPINVOICE.INVDATE), SRETURNSP.TOTALAMT, '') as char) AS THMOS, " & _
                        "Cast(IF(MONTH(SRETURNSP.RETDATE) = MONTH(SPINVOICE.INVDATE),'',SRETURNSP.TOTALAMT) as char) AS PRMOS, " & _
                        "SRETURNSP.TOTALAMT, " & _
                        "SRETURNSP.REMARKS, " & _
                        "SRETURNSP.PRINTCOPY, " & _
                        "SRETURNSP.RETURNID " & _
                    "From " & _
                        "SRETURNSP " & _
                        "INNER JOIN SPINVOICE ON (SRETURNSP.INVOICE = SPINVOICE.SPINVNO) " & _
                    "Where " & _
                        "SRETURNSP.RETDATE = '" & Format(Dfrom.Value, "yyyy-mm-dd") & "' " & _
                        "And Case When SRETURNSP.INVTYPE = 0 Then spinvoice.INVTYPE = 'CREDIT' Else spinvoice.INVTYPE = 'CASH' End " & _
                    "Order By " & _
                        "SRETURNSP.RETDATE, " & _
                        "SRETURNSP.RETNO Desc"
        Else
            '***Access
            'Qry = "SELECT SRETURNSP.RETDATE, SRETURNSP.RETNO, SRETURNSP.INVOICE, SPINVOICE.CHARGETO, SRETURNSP.REASON, SRETURNSP.TOTALAMT, SRETURNSP.REMARKS, SRETURNSP.PRINTCOPY, SRETURNSP.RETURNID " & _
                  "FROM SRETURNSP INNER JOIN SPINVOICE ON SRETURNSP.INVOICE = SPINVOICE.SPINVNO Where (((SRETURNSP.RETDATE) >= #" & dFrom.Value & "# And (SRETURNSP.RETDATE) <= #" & dTo.Value & "#)) ORDER BY SRETURNSP.RETDATE, SRETURNSP.RETNO DESC;"
            'mySQL
            'Qry = "SELECT SRETURNSP.RETDATE, SRETURNSP.RETNO, SRETURNSP.INVOICE, SPINVOICE.CHARGETO, SRETURNSP.REASON, SRETURNSP.TOTALAMT, SRETURNSP.REMARKS, SRETURNSP.PRINTCOPY, SRETURNSP.RETURNID " & _
                  "FROM SRETURNSP INNER JOIN SPINVOICE ON SRETURNSP.INVOICE = SPINVOICE.SPINVNO Where SRETURNSP.RETDATE >= '" & Format(dFrom.Value, "yyyy-mm-dd") & "' And SRETURNSP.RETDATE <= '" & Format(dTo.Value, "yyyy-mm-dd") & "' ORDER BY SRETURNSP.RE
            
            'modify Report 06/08/2010
             
             Qry = "SELECT " & _
                        "SRETURNSP.RETDATE, " & _
                        "SRETURNSP.RETNO, " & _
                        "SRETURNSP.INVOICE, " & _
                        "SPINVOICE.INVDATE, " & _
                        "SPINVOICE.CHARGETO, " & _
                        "Cast(IF(MONTH(SRETURNSP.RETDATE) = MONTH(SPINVOICE.INVDATE), SRETURNSP.TOTALAMT, '') as char) AS THMOS, " & _
                        "Cast(IF(MONTH(SRETURNSP.RETDATE) = MONTH(SPINVOICE.INVDATE),'',SRETURNSP.TOTALAMT) as char) AS PRMOS, " & _
                        "SRETURNSP.TOTALAMT, " & _
                        "SRETURNSP.REMARKS, " & _
                        "SRETURNSP.PRINTCOPY, " & _
                        "SRETURNSP.RETURNID " & _
                    "From " & _
                        "SRETURNSP " & _
                        "INNER JOIN SPINVOICE ON (SRETURNSP.INVOICE = SPINVOICE.SPINVNO) " & _
                    "Where " & _
                        "SRETURNSP.RETDATE >= '" & Format(Dfrom.Value, "yyyy-mm-dd") & "' AND " & _
                        "SRETURNSP.RETDATE <= '" & Format(Dto.Value, "yyyy-mm-dd") & "' " & _
                        "And Case When SRETURNSP.INVTYPE = 0 Then spinvoice.INVTYPE = 'CREDIT' Else spinvoice.INVTYPE = 'CASH' End " & _
                    "Order By " & _
                        "SRETURNSP.RETDATE, " & _
                        "SRETURNSP.RETNO Desc"
        
        
        End If
    ElseIf mySrc = "Branch Purchase Order" Then
    
        If dCode = 1 Then
             Qry = "SELECT spordertaking.SQSERVE, spordertaking.BRANCHNAME, ' ', spordertaking.DATESERVE, spordertaking.DATECONFIRM, '', spordertaking.CONFIRMNO, " & _
                "spordertaking.TOTALAMOUNT, spordertaking.BRANCHID, spordertaking.PRINTPICK, spordertaking.FINALQT " & _
                "From spordertaking Where spordertaking.DATESERVE >= '" & Format(Dfrom.Value, "yyyy-mm-dd") & "' Order By spordertaking.ORDERIDSQ;"
        Else
             Qry = "SELECT spordertaking.SQSERVE, spordertaking.BRANCHNAME, ' ', spordertaking.DATESERVE, spordertaking.DATECONFIRM,' ',spordertaking.CONFIRMNO, " & _
               "spordertaking.TOTALAMOUNT, spordertaking.BRANCHID, spordertaking.PRINTPICK, spordertaking.FINALQT " & _
               "From spordertaking Where spordertaking.DATESERVE >= '" & Format(Dfrom.Value, "yyyy-mm-dd") & "' AND spordertaking.DATESERVE <= '" & Format(Dto.Value, "yyyy-mm-dd") & "' Order by spordertaking.ORDERIDSQ;"

        End If
           
    ElseIf mySrc = "Trust Receipt Agreement" Then
        If dCode = 1 Then
            Qry = "SELECT date_format(spinvoice.INVDATE,'%c/%d/%Y'), spinvoice.TRANSNO, spinvoice.CHARGETO, spinvoice.ADDRESS, spinvoice.INVTYPE, spinvoice.AMOUNT, spinvoice.CTYPE, spinvoice.REMARKS, spinvoice.FRONTLINER, spinvoice.SPSORDERID, spinvoice.PRINTCOPY " & _
                  "From spinvoice Where spinvoice.TRANSNO LIKE 'TRA#%' AND spinvoice.INVDATE = '" & Format(Dfrom.Value, "yyyy-mm-dd") & "' Order By spinvoice.InvDate Desc"
        Else
            Qry = "SELECT date_format(spinvoice.INVDATE,'%c/%d/%Y'), spinvoice.TRANSNO, spinvoice.CHARGETO, spinvoice.ADDRESS, spinvoice.INVTYPE, spinvoice.AMOUNT, spinvoice.CTYPE, spinvoice.REMARKS, spinvoice.FRONTLINER, spinvoice.SPSORDERID, spinvoice.PRINTCOPY " & _
                  "From spinvoice Where spinvoice.TRANSNO LIKE 'TRA#%' AND spinvoice.INVDATE >= '" & Format(Dfrom.Value, "yyyy-mm-dd") & "' AND spinvoice.INVDATE <= '" & Format(Dto.Value, "yyyy-mm-dd") & "' Order By spinvoice.InvDate Desc"
        
        End If
    ElseIf mySrc = "Branch Order" Then
        statUpdate.Visible = True
        If dCode = 1 Then
            'Qry = "SELECT SPORDER.ORDERDATE, SPORDER.REFERENCE, SUPPLIER.SUPPLIERNAME, SPORDER.ORDERTYPE, SPORDER.CONFIRMNO, SPORDER.SUBMITTEDBY, SPORDER.APPROVEDBY, SPORDER.PRINTCOPY, SPORDER.QTYORDER, SPORDER.QTYREC, SPORDER.BACKORDER, SPORDER.BOLASTUPDATE " & _
            '      "FROM SPORDER INNER JOIN SUPPLIER ON SPORDER.SUPPLIERID = SUPPLIER.SUPPLIERID Where SPORDER.ORDERDATE = '" & Format(dFrom.Value, "YYYY-MM-DD") & "' ORDER BY SPORDER.ORDERDATE DESC;"
            Qry = "SELECT date_format(sporder.ORDERDATE, '%c-%d-%Y'), sporder.REFERENCE, sporder.`STATUS`, sporder.CONFIRMNO, sporder.QTYPN, sporder.QTYRECPN, sporder.BACKORDERPN, sporder.QTYORDER, sporder.QTYREC, sporder.BACKORDER, DATEDIFF(sporder.LEADTIME, sporder.ORDERDATE) AS LEADTIME, sporder.PICKEDBY, sporder.PRINTCOPY " & _
                  "From sporder Where SPORDER.ORDERDATE = '" & Format(Dfrom.Value, "YYYY-MM-DD") & "' ORDER BY SPORDER.ORDERDATE DESC;"
        
        Else
            '  QUERY SA PAG DISPLAY SA SUMMARY
            Qry = "SELECT date_format(sporder.ORDERDATE, '%c-%d-%Y'), sporder.REFERENCE, sporder.`STATUS`, sporder.CONFIRMNO, sporder.QTYPN, sporder.QTYRECPN, sporder.BACKORDERPN, sporder.QTYORDER, sporder.QTYREC, sporder.BACKORDER, DATEDIFF(sporder.LEADTIME, sporder.ORDERDATE) AS LEADTIME, sporder.PICKEDBY, sporder.PRINTCOPY " & _
                  "From sporder Where SPORDER.ORDERDATE >= '" & Format(Dfrom.Value, "YYYY-MM-DD") & "' and SPORDER.ORDERDATE <= '" & Format(Dto.Value, "YYYY-MM-DD") & "' ORDER BY SPORDER.ORDERDATE DESC;"
                
        End If
'**************************** CODE UPDATED: 08/04/2017 **********************************
    ElseIf mySrc = "SETTLEMENT DISCREPANCY" Then
        If dCode = 1 Then
            Qry = "call sp_partsmatrix.sp_SDiscrepancy_Summary('" & sysBranchCode & "', " & dCode & ", '" & Format(Dfrom.Value, "YYYY-MM-DD") & "', '" & Format(Dto.Value, "YYYY-MM-DD") & "')"
        Else
            Qry = "call sp_partsmatrix.sp_SDiscrepancy_Summary('" & sysBranchCode & "', " & dCode & ", '" & Format(Dfrom.Value, "YYYY-MM-DD") & "', '" & Format(Dto.Value, "YYYY-MM-DD") & "')"
        End If
    ElseIf mySrc = "SDiscrepancy Adjustment" Then
        If dCode = 1 Then
            Qry = "SELECT INVNTRYADJ.ADJDATE, INVNTRYADJ.ADJNO, INVNTRYADJ.DOCREF, INVNTRYADJ.REASON, INVNTRYADJ.REMARKS, INVNTRYADJ.PREPBY, INVNTRYADJ.APRVBY, INVNTRYADJ.PRINTCOPY " & _
                  "From InvntryADJ Where InvntryADJ.ADJDATE = '" & Format(Dfrom.Value, "yyyy-mm-dd") & "' ORDER BY INVNTRYADJ.ADJDATE, INVNTRYADJ.ADJNO;"
        Else
            Qry = "SELECT INVNTRYADJ.ADJDATE, INVNTRYADJ.ADJNO, INVNTRYADJ.DOCREF, INVNTRYADJ.REASON, INVNTRYADJ.REMARKS, INVNTRYADJ.PREPBY, INVNTRYADJ.APRVBY, INVNTRYADJ.PRINTCOPY " & _
                  "From InvntryADJ Where InvntryADJ.ADJDATE >= '" & Format(Dfrom.Value, "yyyy-mm-dd") & "' And InvntryADJ.ADJDATE <= '" & Format(Dto.Value, "yyyy-mm-dd") & "' ORDER BY INVNTRYADJ.ADJDATE, INVNTRYADJ.ADJNO;"
        End If
    ElseIf mySrc = "Daily Cash Deposit Report" Then
        Qry = "call sp_mrcretails.sp_cashdeposit_summary('" & sysBranchCode & "','" & Format(Dfrom.Value, "yyyy-mm-dd") & "','" & Format(Dto.Value, "yyyy-mm-dd") & "','" & "SUMMARY" & "')"
'************************ END CODE UPDATED: 08/04/2017 **********************************
    ElseIf mySrc = "Detailed Inventory Outgoing" Then
        Qry = "call sp_mrcretails.sp_acctg_DIO('" & sysBranchCode & "','" & Format(Dfrom.Value, "yyyy-mm-dd") & "','" & Format(Dto.Value, "yyyy-mm-dd") & "')"
    ElseIf mySrc = "Detailed Inventory Incoming" Then
        Qry = "call sp_mrcretails.sp_acctg_DII('" & sysBranchCode & "','" & Format(Dfrom.Value, "yyyy-mm-dd") & "','" & Format(Dto.Value, "yyyy-mm-dd") & "')"
    ElseIf mySrc = "Sales Book" Then
        Qry = "call sp_mrcretails.sp_acctg_SalesBook('" & sysBranchCode & "','" & Format(Dfrom.Value, "yyyy-mm-dd") & "','" & Format(Dto.Value, "yyyy-mm-dd") & "')"
    ElseIf mySrc = "Sales Return Report" Then
        Qry = "call sp_mrcretails.sp_acctg_SalesReturn('" & sysBranchCode & "','" & Format(Dfrom.Value, "yyyy-mm-dd") & "','" & Format(Dto.Value, "yyyy-mm-dd") & "')"
    ElseIf mySrc = "Service Receipts Book" Then
        Qry = "call sp_mrcretails.sp_acctg_JobOrder('" & sysBranchCode & "','" & Format(Dfrom.Value, "yyyy-mm-dd") & "','" & Format(Dto.Value, "yyyy-mm-dd") & "')"
'************************ END CODE UPDATED: 07/10/2019 **********************************
    ElseIf mySrc = "ADJUSTMENT REQUEST" Then
        Qry = "call sp_mrcretails.sp_other_adjustment_summary('" & sysBranchCode & "'," & Val(dCode) & ",'" & Format(Dfrom.Value, "yyyy-mm-dd") & "','" & Format(Dto.Value, "yyyy-mm-dd") & "')"
    ElseIf mySrc = "CANCELLATION MODULE" Then
        Qry = "call sp_mrcretails.sp_cancellation_module_summary('" & sysBranchCode & "'," & Val(dCode) & ",'" & Format(Dfrom.Value, "yyyy-mm-dd") & "','" & Format(Dto.Value, "yyyy-mm-dd") & "')"
    End If
    If rsSummary.State = adStateOpen Then rsSummary.Close
    rsSummary.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    If rsSummary.RecordCount = 0 Then
        flxSum.Clear
        flxSum.Rows = 2
        flxSum.Cols = TCol
        Call RXResizer1_ResizeComplete
        Exit Sub
    End If

    If mySrc = "Detailed Inventory Incoming" Then
        With rsSummary
            .MoveFirst
            Call RXResizer1_ResizeComplete
            flxSum.Rows = .RecordCount + 1
            If .RecordCount > 0 Then
                For i = 1 To flxSum.Rows - 1
                    flxSum.TextMatrix(i, 1) = !TRANSDATE
                    flxSum.TextMatrix(i, 2) = !REFERENCE
                    flxSum.TextMatrix(i, 3) = !PACKDATE
                    flxSum.TextMatrix(i, 4) = !PartNumber
                    flxSum.TextMatrix(i, 5) = !Description
                    flxSum.TextMatrix(i, 6) = !STATUTORY
                    flxSum.TextMatrix(i, 7) = !Receive
                    flxSum.TextMatrix(i, 8) = !Cost
                    flxSum.TextMatrix(i, 9) = !MDSEINVTY
                    flxSum.TextMatrix(i, 10) = !OTHERINVTY
                    flxSum.TextMatrix(i, 11) = !INPUTVAT
                    flxSum.TextMatrix(i, 12) = !TtlCost
                    flxSum.TextMatrix(i, 13) = !TTLCOSTOTHER
                    flxSum.TextMatrix(i, 14) = !RECTYPE
                    flxSum.TextMatrix(i, 15) = !RECEIVEDFROM
                    flxSum.TextMatrix(i, 16) = !SUPPLIERNAME
                    .MoveNext
                Next i
            End If
        End With
    ElseIf mySrc = "Daily Cash Deposit Report" Then
        With rsSummary
            .MoveFirst
            Call RXResizer1_ResizeComplete
            flxSum.Rows = .RecordCount + 1
            If .RecordCount > 0 Then
                For i = 1 To flxSum.Rows - 1
                'DATE, SPCASHWALKIN, SPCASHWHSALE, SERVICE, OTHER_INCOME, TTLSALES, DOWNPAYMENT, SPCOLL, SRVCOLL,
                'OTHERCOLL, TTLCOLL, RECFUNDS, CASHRETURN, CASHDISBURSEMENT, CREDITCARD, ONLINEPAYMENT,
                'INTERBRANCHPAYMENT, TOTALLESS, NETCOLLECTION, CASH, CHECK, TOTAL, DEPOSITEDAMOUNT, CASHOUTAMOUNT,
                'EXCESSSHORTAGE, VARIANCE, DATEDEPOSIT, INTERVALS, DEPOSITREF, REMARKS
                    flxSum.TextMatrix(i, 1) = !Date
                    flxSum.TextMatrix(i, 2) = Format(!SPCASHWALKIN, "###,###,###.00")
                    flxSum.TextMatrix(i, 3) = Format(!SPCASHWHSALE, "###,###,###.00")
                    flxSum.TextMatrix(i, 4) = Format(!Service, "###,###,###.00")
                    flxSum.TextMatrix(i, 5) = Format(!OTHER_INCOME, "###,###,###.00")
                    flxSum.TextMatrix(i, 6) = Format(!TTLSALES, "###,###,###.00")
                    flxSum.TextMatrix(i, 7) = Format(!DownPayment, "###,###,###.00")
                    flxSum.TextMatrix(i, 8) = Format(!SPCOLL, "###,###,###.00")
                    flxSum.TextMatrix(i, 9) = Format(!SRVCOLL, "###,###,###.00")
                    flxSum.TextMatrix(i, 10) = Format(!OTHERCOLL, "###,###,###.00")
                    flxSum.TextMatrix(i, 11) = Format(!TTLCOLL, "###,###,###.00")
                    flxSum.TextMatrix(i, 12) = Format(!RECFUNDS, "###,###,###.00")
                    flxSum.TextMatrix(i, 13) = Format(!CashReturn, "###,###,###.00")
                    flxSum.TextMatrix(i, 14) = Format(!SalesReturn, "###,###,###.00")
                    flxSum.TextMatrix(i, 15) = Format(!CASHDISBURSEMENT, "###,###,###.00")
                    flxSum.TextMatrix(i, 16) = Format(!CREDITCARD, "###,###,###.00")
                    flxSum.TextMatrix(i, 17) = Format(!ONLINEPAYMENT, "###,###,###.00")
                    flxSum.TextMatrix(i, 18) = Format(!INTERBRANCHPAYMENT, "###,###,###.00")
                    flxSum.TextMatrix(i, 19) = Format(!TOTALLESS, "###,###,###.00")
                    flxSum.TextMatrix(i, 20) = Format(!NETCOLLECTION, "###,###,###.00")
                    flxSum.TextMatrix(i, 21) = Format(!Cash, "###,###,###.00")
                    flxSum.TextMatrix(i, 22) = Format(!Check, "###,###,###.00")
                    flxSum.TextMatrix(i, 23) = Format(!Total, "###,###,###.00")
                    flxSum.TextMatrix(i, 24) = Format(!DEPOSITEDAMOUNT, "###,###,###.00")
                    flxSum.TextMatrix(i, 25) = Format(!CASHOUTAMOUNT, "###,###,###.00")
                    flxSum.TextMatrix(i, 26) = Format(!EXCESSSHORTAGE, "###,###,###.00")
                    flxSum.TextMatrix(i, 27) = Format(!VARIANCE, "###,###,###.00")
                    flxSum.TextMatrix(i, 28) = Format(!DATEDEPOSIT, "MM/DD/YYYY")
                    flxSum.TextMatrix(i, 29) = !INTERVALS
                    flxSum.TextMatrix(i, 30) = !DEPOSITREF
                    flxSum.TextMatrix(i, 31) = !Remarks
                    
                    If flxSum.TextMatrix(i, 1) = "TOTAL" Then 'TOTAL
                        flxSum.Col = 1: flxSum.Row = i: flxSum.ColSel = 30: flxSum.RowSel = i: flxSum.FillStyle = flexFillRepeat: flxSum.CellBackColor = RGB(211, 211, 211)
                        flxSum.Col = 2
                    Else
                        flxSum.Col = 1: flxSum.Row = i: flxSum.ColSel = 30: flxSum.RowSel = i: flxSum.FillStyle = flexFillRepeat: flxSum.CellBackColor = RGB(250, 250, 250)
                        flxSum.Col = 2
                    End If
                    .MoveNext
                Next i
            End If
        End With
    ElseIf mySrc = "ADJUSTMENT REQUEST" Then
        With rsSummary
            .MoveFirst
            Call RXResizer1_ResizeComplete
            flxSum.Rows = .RecordCount + 1
            If .RecordCount > 0 Then
                For i = 1 To flxSum.Rows - 1
                'oa.req_date, oa.req_status, oa.ticketno, ifnull(oa.controlno,'') as controlno, ifnull(date_format(oa.date_final,'%m/%d/%Y'),'') as date_final,
                'oa.subj_request, emp.EMPLOYEENAME as REQUESTBY, ifnull(oa.aud_apprvd_by,'') as aud_apprvd_by, ifnull(oa.mis_incharge,'') as mis_incharge,
                'ifnull(oa.script_file,'') as script_file
                    flxSum.TextMatrix(i, 1) = !req_date
                    flxSum.TextMatrix(i, 2) = !req_status
                    flxSum.TextMatrix(i, 3) = !ticketno
                    flxSum.TextMatrix(i, 4) = !controlno
                    flxSum.TextMatrix(i, 5) = !ADJNO
                    flxSum.TextMatrix(i, 6) = !date_final
                    flxSum.TextMatrix(i, 7) = !subj_request
                    flxSum.TextMatrix(i, 8) = !REQUESTBY
                    flxSum.TextMatrix(i, 9) = !aud_apprvd_by
                    flxSum.TextMatrix(i, 10) = mis_incharge
                    flxSum.TextMatrix(i, 11) = script_file
                    .MoveNext
                Next i
            End If
        End With
    ElseIf mySrc <> "Daily Cash Deposit Report" Then
        Set flxSum.DataSource = rsSummary
    End If
    DoEvents
    If mySrc = "Branch Purchase Order" Then Call DisplayRef
    
    ActRow = 1
    flxSum.Visible = True
    Call RXResizer1_ResizeComplete
End Sub
Private Sub DisplayRef()
    Dim qryRef As String, r As Integer, Fil As Long, ConRef As String
    Dim rsDRef As New ADODB.Recordset
    
    For Fil = 1 To flxSum.Rows - 1
        ConRef = ""
        qryRef = "SELECT spordrconfirm.REFERENCE From spordrconfirm Where spordrconfirm.SQSERVE =" & flxSum.TextMatrix(Fil, 1) & ";"
        If rsDRef.State = adStateOpen Then rsDRef.Close
        rsDRef.Open qryRef, LYRcon, adOpenKeyset, adLockOptimistic
        If rsDRef.RecordCount > 0 Then
            rsDRef.MoveFirst
                 For r = 1 To rsDRef.RecordCount
                    ConRef = ConRef & rsDRef!REFERENCE & IIf((rsDRef.RecordCount = 1), "", ",")
                    rsDRef.MoveNext
                Next r
                flxSum.TextMatrix(Fil, 3) = IIf((rsDRef.RecordCount = 1), ConRef, Left$(ConRef, Len(ConRef) - 1))
        End If
    Next Fil
    
 End Sub

Private Sub cboRebuilt_Click()

    If mySrc = "Job Order" Then
        If cmdNext.Tag = "1" Then
            If cboRebuilt.Text = "View All" Then
                Qry = "SELECT SRVCJOBORDER.DATEPROMISED, SRVCJOBORDER.JOBORDERNO, SRVCJOBORDER.CUSTOMER, SRVCUSTOMER.ADDRESS, SRVCJOBORDER.PAYTYPE, SRVCJOBORDER.BALANCE, SRVCJOBORDER.ORNO, SRVCUSTOMER.MODEL, SRVCJOBORDER.KMREADING, SRVCJOBORDER.JOTYPE, SRVCJOBORDER.MECHANIC, SRVCJOBORDER.PRINTCOPY, SRVCJOBORDER.REMARKS, SRVCJOBORDER.JOBORDERID " & _
                      "FROM SRVCJOBORDER INNER JOIN SRVCUSTOMER ON SRVCJOBORDER.ACCOUNTNO = SRVCUSTOMER.ACCOUNTNO Where SRVCJOBORDER.DATEPROMISED = '" & Format(Dfrom.Value, "yyyy-mm-dd") & "' ORDER BY SRVCJOBORDER.DATEPROMISED, SRVCJOBORDER.JOBORDERNO;"
            ElseIf cboRebuilt.Text = "View Machine Shop Only" Then
                Qry = "SELECT SRVCJOBORDER.DATEPROMISED, SRVCJOBORDER.JOBORDERNO, SRVCJOBORDER.CUSTOMER, SRVCUSTOMER.ADDRESS, SRVCJOBORDER.PAYTYPE, SRVCJOBORDER.BALANCE, SRVCJOBORDER.ORNO, SRVCUSTOMER.MODEL, SRVCJOBORDER.KMREADING, SRVCJOBORDER.JOTYPE, SRVCJOBORDER.MECHANIC, SRVCJOBORDER.PRINTCOPY, SRVCJOBORDER.REMARKS, SRVCJOBORDER.JOBORDERID " & _
                      "FROM SRVCJOBORDER INNER JOIN SRVCUSTOMER ON SRVCJOBORDER.ACCOUNTNO = SRVCUSTOMER.ACCOUNTNO Where SRVCJOBORDER.DATEPROMISED = '" & Format(Dfrom.Value, "yyyy-mm-dd") & "' And SRVCJOBORDER.JOTYPE = 'MACHINE SHOP' ORDER BY SRVCJOBORDER.DATEPROMISED, SRVCJOBORDER.JOBORDERNO;"
            ElseIf cboRebuilt.Text = "Exclude Machine Shop" Then
                Qry = "SELECT SRVCJOBORDER.DATEPROMISED, SRVCJOBORDER.JOBORDERNO, SRVCJOBORDER.CUSTOMER, SRVCUSTOMER.ADDRESS, SRVCJOBORDER.PAYTYPE, SRVCJOBORDER.BALANCE, SRVCJOBORDER.ORNO, SRVCUSTOMER.MODEL, SRVCJOBORDER.KMREADING, SRVCJOBORDER.JOTYPE, SRVCJOBORDER.MECHANIC, SRVCJOBORDER.PRINTCOPY, SRVCJOBORDER.REMARKS, SRVCJOBORDER.JOBORDERID " & _
                      "FROM SRVCJOBORDER INNER JOIN SRVCUSTOMER ON SRVCJOBORDER.ACCOUNTNO = SRVCUSTOMER.ACCOUNTNO Where SRVCJOBORDER.DATEPROMISED = '" & Format(Dfrom.Value, "yyyy-mm-dd") & "' And SRVCJOBORDER.JOTYPE <> 'MACHINE SHOP' ORDER BY SRVCJOBORDER.DATEPROMISED, SRVCJOBORDER.JOBORDERNO;"
            End If
        Else
            If cboRebuilt.Text = "View All" Then
                Qry = "SELECT SRVCJOBORDER.DATEPROMISED, SRVCJOBORDER.JOBORDERNO, SRVCJOBORDER.CUSTOMER, SRVCUSTOMER.ADDRESS, SRVCJOBORDER.PAYTYPE, SRVCJOBORDER.BALANCE, SRVCJOBORDER.ORNO, SRVCUSTOMER.MODEL, SRVCJOBORDER.KMREADING, SRVCJOBORDER.JOTYPE, SRVCJOBORDER.MECHANIC, SRVCJOBORDER.PRINTCOPY, SRVCJOBORDER.REMARKS, SRVCJOBORDER.JOBORDERID " & _
                      "FROM SRVCJOBORDER INNER JOIN SRVCUSTOMER ON SRVCJOBORDER.ACCOUNTNO = SRVCUSTOMER.ACCOUNTNO Where SRVCJOBORDER.DATEPROMISED >= '" & Format(Dfrom.Value, "yyyy-mm-dd") & "' And SRVCJOBORDER.DATEPROMISED <= '" & Format(Dto.Value, "yyyy-mm-dd") & "' ORDER BY SRVCJOBORDER.DATEPROMISED, SRVCJOBORDER.JOBORDERNO;"
            ElseIf cboRebuilt.Text = "View Machine Shop Only" Then
                Qry = "SELECT SRVCJOBORDER.DATEPROMISED, SRVCJOBORDER.JOBORDERNO, SRVCJOBORDER.CUSTOMER, SRVCUSTOMER.ADDRESS, SRVCJOBORDER.PAYTYPE, SRVCJOBORDER.BALANCE, SRVCJOBORDER.ORNO, SRVCUSTOMER.MODEL, SRVCJOBORDER.KMREADING, SRVCJOBORDER.JOTYPE, SRVCJOBORDER.MECHANIC, SRVCJOBORDER.PRINTCOPY, SRVCJOBORDER.REMARKS, SRVCJOBORDER.JOBORDERID " & _
                      "FROM SRVCJOBORDER INNER JOIN SRVCUSTOMER ON SRVCJOBORDER.ACCOUNTNO = SRVCUSTOMER.ACCOUNTNO Where SRVCJOBORDER.DATEPROMISED >= '" & Format(Dfrom.Value, "yyyy-mm-dd") & "' And SRVCJOBORDER.DATEPROMISED <= '" & Format(Dto.Value, "yyyy-mm-dd") & "' And SRVCJOBORDER.JOTYPE = 'MACHINE SHOP' ORDER BY SRVCJOBORDER.DATEPROMISED, SRVCJOBORDER.JOBORDERNO;"
            ElseIf cboRebuilt.Text = "Exclude Machine Shop" Then
                Qry = "SELECT SRVCJOBORDER.DATEPROMISED, SRVCJOBORDER.JOBORDERNO, SRVCJOBORDER.CUSTOMER, SRVCUSTOMER.ADDRESS, SRVCJOBORDER.PAYTYPE, SRVCJOBORDER.BALANCE, SRVCJOBORDER.ORNO, SRVCUSTOMER.MODEL, SRVCJOBORDER.KMREADING, SRVCJOBORDER.JOTYPE, SRVCJOBORDER.MECHANIC, SRVCJOBORDER.PRINTCOPY, SRVCJOBORDER.REMARKS, SRVCJOBORDER.JOBORDERID " & _
                      "FROM SRVCJOBORDER INNER JOIN SRVCUSTOMER ON SRVCJOBORDER.ACCOUNTNO = SRVCUSTOMER.ACCOUNTNO Where SRVCJOBORDER.DATEPROMISED >= '" & Format(Dfrom.Value, "yyyy-mm-dd") & "' And SRVCJOBORDER.DATEPROMISED <= '" & Format(Dto.Value, "yyyy-mm-dd") & "' And SRVCJOBORDER.JOTYPE <> 'MACHINE SHOP' ORDER BY SRVCJOBORDER.DATEPROMISED, SRVCJOBORDER.JOBORDERNO;"
            End If
            
        End If

        If rsSummary.State = adStateOpen Then rsSummary.Close
        rsSummary.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If rsSummary.RecordCount = 0 Then
            flxSum.Clear
            flxSum.Rows = 2
            flxSum.Cols = TCol
            Call RXResizer1_ResizeComplete
            Exit Sub
        End If
        Set flxSum.DataSource = rsSummary
        DoEvents
        
        ActRow = 1
        flxSum.Visible = True
        Call RXResizer1_ResizeComplete
    End If
End Sub

Private Sub cmdExportTRA_Click()
    Dim rsTRA As New ADODB.Recordset
    Dim i As Integer, Ans As Integer
    Dim myFile As String, Qry As String, MyExt As String
    
    TRAImport.Show vbModal
    

End Sub

Private Sub cmdFind_Click()
    Dim i As Integer, j As Integer
        
    If txtFind.Text = "" Then Exit Sub
    With flxSum
        For i = fRow To .Rows - 1
            For j = fCol To TCol - 1
                If InStr(.TextMatrix(i, j), txtFind.Text) > 0 Then
                    .Col = 1: .Row = i
                    .ColSel = TCol - 1: .RowSel = i
                    .TopRow = i
                    fRow = IIf(j = TCol - 1, i + 1, i)
                    fCol = IIf(j = TCol - 1, 1, j + 1)
                    ActRow = i
                    Exit Sub
                End If
            Next j
            fCol = 1
        Next i
    End With
    MsgBox "No match found!", vbInformation + vbOKOnly, "Find"
    fRow = 1
    fCol = 1
End Sub

Private Sub cmdNext_Click()
    cmdNext.Tag = "2"
    cmdNext.Visible = False
    cmdPrev.Left = cmdNext.Left
    cmdPrev.Top = cmdNext.Top
    cmdPrev.Visible = True
    Dto.Visible = True
    cmdRefresh.Visible = True
    Dto.Value = Dfrom.Value
End Sub

Private Sub cmdPrev_Click()
    cmdNext.Tag = "1"
    cmdNext.Visible = True
    cmdPrev.Visible = False
    Dto.Visible = False
    cmdRefresh.Visible = False
    Call RefreshRecord(1)
End Sub

Private Sub cmdRefresh_Click()
    Call RefreshRecord(2)
End Sub

Private Sub cmdSaveExcel_Click()
On Error GoTo er
    Dim XLapp As Excel.Application
    Dim wbkMVR As Excel.Workbook
    Dim wksLineA1 As Excel.Worksheet
    Dim sSaveAsFileName
    Dim i As Long
    Dim j As Long, x As Long
    
    
    With dlgDEMO
        .CancelError = True
        .Flags = cdlOFNHideReadOnly
        .filename = ""
        .Filter = "Microsoft Excel (*.xls)|*.xls|All Files(*.*)|*.*"
        .ShowOpen
        If dlgDEMO.filename = "" Then Exit Sub
        sSaveAsFileName = .filename
    End With
    
    Screen.MousePointer = vbHourglass
    Me.Enabled = False
    Set XLapp = CreateObject("Excel.Application")
    XLapp.Visible = True
    XLapp.Workbooks.Add
    Set wbkMVR = XLapp.ActiveWorkbook
    Set wksLineA1 = wbkMVR.Sheets(1)
    
    If mySrc = "Daily Cash Deposit Report" Then
        wksLineA1.Name = "Sales"
        wksLineA1.Cells(1, 1).Value = "DATE"
        wksLineA1.Cells(1, 2).Value = "SPAREPARTS CASH (WALK-IN)"
        wksLineA1.Cells(1, 3).Value = "SPAREPARTS CASH (WHOLESALE)"
        wksLineA1.Cells(1, 4).Value = "SERVICE"
        wksLineA1.Cells(1, 5).Value = "OTHER INCOME"
        wksLineA1.Cells(1, 6).Value = "TOTAL CASH SALES"
        wksLineA1.Cells(1, 7).Value = "DOWNPAYMENT"
        wksLineA1.Cells(1, 8).Value = "A/R COLLECTION - SPAREPARTS"
        wksLineA1.Cells(1, 9).Value = "A/R COLLECTION - SERVICE"
        wksLineA1.Cells(1, 10).Value = "OTHER COLLECTION"
        wksLineA1.Cells(1, 11).Value = "TOTAL COLLECTION"
        wksLineA1.Cells(1, 12).Value = "RECEIVED/BORROWED FUNDS"
        wksLineA1.Cells(1, 13).Value = "CASH RETURN"
        wksLineA1.Cells(1, 14).Value = "CASH DISBURSEMENT"
        wksLineA1.Cells(1, 15).Value = "CREDIT CARD"
        wksLineA1.Cells(1, 16).Value = "ONLINE PAYMENT"
        wksLineA1.Cells(1, 17).Value = "INTER-BRANCH PAYMENT"
        wksLineA1.Cells(1, 18).Value = "TOTAL LESS"
        wksLineA1.Cells(1, 19).Value = "NET COLLECTION"
        wksLineA1.Cells(1, 20).Value = "CASH"
        wksLineA1.Cells(1, 21).Value = "CHECK"
        wksLineA1.Cells(1, 22).Value = "TOTAL"
        wksLineA1.Cells(1, 23).Value = "DEPOSITED AMOUNT"
        wksLineA1.Cells(1, 24).Value = "CASH-OUT AMOUNT"
        wksLineA1.Cells(1, 25).Value = "EXCESS/SHORTAGE"
        wksLineA1.Cells(1, 26).Value = "VARIANCE"
        wksLineA1.Cells(1, 27).Value = "DATE DEPOSIT"
        wksLineA1.Cells(1, 28).Value = "INTERVALS"
        wksLineA1.Cells(1, 28).Value = "DEPOSIT REF."
        wksLineA1.Cells(1, 29).Value = "REMARKS"
        With flxSum
            For i = 1 To .Rows - 1
                For j = 1 To 30
                    wksLineA1.Cells(i + 1, j).Value = .TextMatrix(i, j)
                Next j
            DoEvents
            Next i
        End With
        wksLineA1.Range("A1:AD" & (i - 1)).Columns.AutoFit
        wksLineA1.Range("A1:AD1").Font.Bold = True
        XLapp.DisplayAlerts = False
    ElseIf mySrc = "Detailed Inventory Outgoing" Then
        wksLineA1.Name = "DIO"
        wksLineA1.Cells(1, 1).Value = "COMPANY NAME: " & SysBranch
        wksLineA1.Cells(2, 1).Value = "ADDRESS: " & SysAddress
        wksLineA1.Cells(4, 1).Value = "DETAILED INVENTORY OUTGOING"
        wksLineA1.Cells(5, 1).Value = "DATE COVERED: " & Format(Dfrom.Value, "MMMM DD, YYYY") & " TO " & Format(Dto.Value, "MMMM DD, YYYY")
        wksLineA1.Cells(6, 1).Value = "DATE"
        wksLineA1.Cells(6, 2).Value = "REF. NO"
        wksLineA1.Cells(6, 3).Value = "PART NUMBER"
        wksLineA1.Cells(6, 4).Value = "PART DESCRIPTIONS"
        wksLineA1.Cells(6, 5).Value = "STATUTORY REPORTING"
        wksLineA1.Cells(6, 6).Value = "QTY"
        wksLineA1.Cells(6, 7).Value = "COST w/o VAT"
        wksLineA1.Cells(6, 8).Value = "MDSE. INVTY."
        wksLineA1.Cells(6, 9).Value = "VAT INPUT"
        wksLineA1.Cells(6, 10).Value = "TOTAL COST w/ VAT"
        wksLineA1.Cells(6, 11).Value = "OUTGOING TYPE"
        wksLineA1.Cells(6, 12).Value = "DELIVER TO CUSTOMER"
        wksLineA1.Cells(6, 13).Value = "SUPPLIER'S NAME"
        With flxSum
            x = 6
            For i = 1 To .Rows - 1
                For j = 1 To 13
                    wksLineA1.Cells(x + 1, j).Value = .TextMatrix(i, j)
                Next j
            DoEvents
                x = x + 1
            Next i
        End With
        wksLineA1.Range("A6:M" & (x - 1)).Columns.AutoFit
        wksLineA1.Range("A1").Font.Bold = True
        wksLineA1.Range("A4").Font.Bold = True
        wksLineA1.Range("A6:M6").Font.Bold = True
        XLapp.DisplayAlerts = False
    ElseIf mySrc = "Detailed Inventory Incoming" Then
        wksLineA1.Name = "DII"
        wksLineA1.Cells(1, 1).Value = "COMPANY NAME: " & SysBranch
        wksLineA1.Cells(2, 1).Value = "ADDRESS: " & SysAddress
        wksLineA1.Cells(4, 1).Value = "DETAILED INVENTORY INCOMING"
        wksLineA1.Cells(5, 1).Value = "DATE COVERED: " & Format(Dfrom.Value, "MMMM DD, YYYY") & " TO " & Format(Dto.Value, "MMMM DD, YYYY")
        wksLineA1.Cells(6, 1).Value = "DATE RECEIVED"
        wksLineA1.Cells(6, 2).Value = "REF. NO"
        wksLineA1.Cells(6, 3).Value = "PACKING DATE"
        wksLineA1.Cells(6, 4).Value = "PART NUMBER"
        wksLineA1.Cells(6, 5).Value = "PART DESCRIPTIONS"
        wksLineA1.Cells(6, 6).Value = "STATUTORY REPORTING"
        wksLineA1.Cells(6, 7).Value = "QTY"
        wksLineA1.Cells(6, 8).Value = "COST w/o VAT"
        wksLineA1.Cells(6, 9).Value = "MDSE. INVTY."
        wksLineA1.Cells(6, 10).Value = "OTHER INVTY."
        wksLineA1.Cells(6, 11).Value = "VAT INPUT"
        wksLineA1.Cells(6, 12).Value = "TOTAL COST w/ VAT"
        wksLineA1.Cells(6, 13).Value = "TOTAL COST w/ VAT OTHER"
        wksLineA1.Cells(6, 14).Value = "RECEIVED TYPE"
        wksLineA1.Cells(6, 15).Value = "RECEIVED FROM"
        wksLineA1.Cells(6, 16).Value = "SUPPLIER'S NAME"
        With flxSum
            x = 6
            For i = 1 To .Rows - 1
                For j = 1 To 16
                    wksLineA1.Cells(x + 1, j).Value = .TextMatrix(i, j)
                Next j
            DoEvents
                x = x + 1
            Next i
        End With
        wksLineA1.Range("A6:P" & (x - 1)).Columns.AutoFit
        wksLineA1.Range("A1").Font.Bold = True
        wksLineA1.Range("A4").Font.Bold = True
        wksLineA1.Range("A6:P6").Font.Bold = True
        XLapp.DisplayAlerts = False
    ElseIf mySrc = "Sales Book" Then
        wksLineA1.Name = "Sales Book"
        wksLineA1.Cells(1, 1).Value = "COMPANY NAME: " & SysBranch
        wksLineA1.Cells(2, 1).Value = "ADDRESS: " & SysAddress
        wksLineA1.Cells(4, 1).Value = "SALES BOOK"
        wksLineA1.Cells(5, 1).Value = "DATE COVERED: " & Format(Dfrom.Value, "MMMM DD, YYYY") & " TO " & Format(Dto.Value, "MMMM DD, YYYY")
        wksLineA1.Cells(6, 1).Value = "DATE"
        wksLineA1.Cells(6, 2).Value = "CUSTOMER'S NAME"
        wksLineA1.Cells(6, 3).Value = "INVOICE NUMBER"
        wksLineA1.Cells(6, 4).Value = "CLASSIFICATION"
        wksLineA1.Cells(6, 5).Value = "SUPPLIER'S NAME"
        wksLineA1.Cells(6, 6).Value = "STATUTORY REPORTING"
        wksLineA1.Cells(6, 7).Value = "TRANSTYPE"
        wksLineA1.Cells(6, 8).Value = "PART NUMBER"
        wksLineA1.Cells(6, 9).Value = "PART DESCRIPTION"
        wksLineA1.Cells(6, 10).Value = "QTY"
        wksLineA1.Cells(6, 11).Value = "UNIT PRICE"
        wksLineA1.Cells(6, 12).Value = "UNIT COST"
        wksLineA1.Cells(6, 13).Value = "TOTAL COST"
        wksLineA1.Cells(6, 14).Value = "TOTAL SALES"
        wksLineA1.Cells(6, 15).Value = "CUSTOMER TYPE"
        wksLineA1.Cells(6, 16).Value = "FRONTLINER"
        wksLineA1.Cells(6, 17).Value = "ADDRESS"
        With flxSum
            x = 6
            For i = 1 To .Rows - 1
                For j = 1 To 17
                    wksLineA1.Cells(x + 1, j).Value = .TextMatrix(i, j)
                Next j
            DoEvents
                x = x + 1
            Next i
        End With
        wksLineA1.Range("A6:Q" & (x - 1)).Columns.AutoFit
        wksLineA1.Range("A1").Font.Bold = True
        wksLineA1.Range("A4").Font.Bold = True
        wksLineA1.Range("A6:Q6").Font.Bold = True
        XLapp.DisplayAlerts = False
    ElseIf mySrc = "Sales Return Report" Then
        wksLineA1.Name = "Sales Return"
        wksLineA1.Cells(1, 1).Value = "COMPANY NAME: " & SysBranch
        wksLineA1.Cells(2, 1).Value = "ADDRESS: " & SysAddress
        wksLineA1.Cells(4, 1).Value = "SALES RETURN"
        wksLineA1.Cells(5, 1).Value = "DATE COVERED: " & Format(Dfrom.Value, "MMMM DD, YYYY") & " TO " & Format(Dto.Value, "MMMM DD, YYYY")
        wksLineA1.Cells(6, 1).Value = "DATE"
        wksLineA1.Cells(6, 2).Value = "SALE RET. NO."
        wksLineA1.Cells(6, 3).Value = "INVOICE NUMBER"
        wksLineA1.Cells(6, 4).Value = "INVOICE DATE"
        wksLineA1.Cells(6, 5).Value = "INVOICE TYPE"
        wksLineA1.Cells(6, 6).Value = "STATUTORY REPORTING"
        wksLineA1.Cells(6, 7).Value = "SUPPLIER'S NAME"
        wksLineA1.Cells(6, 8).Value = "CUSTOMER NAME"
        wksLineA1.Cells(6, 9).Value = "PART NUMBER"
        wksLineA1.Cells(6, 10).Value = "PART DESCRIPTION"
        wksLineA1.Cells(6, 11).Value = "QTY"
        wksLineA1.Cells(6, 12).Value = "UNIT PRICE"
        wksLineA1.Cells(6, 13).Value = "TOTAL PRICE"
        wksLineA1.Cells(6, 14).Value = "UNIT COST"
        wksLineA1.Cells(6, 15).Value = "TOTAL COST"
        wksLineA1.Cells(6, 16).Value = "AR - TRADE"
        wksLineA1.Cells(6, 17).Value = "COST OF SALES"
        wksLineA1.Cells(6, 18).Value = "MDSE. INVTY"
        wksLineA1.Cells(6, 19).Value = "SALES RETURN AND ALLOW."
        wksLineA1.Cells(6, 20).Value = "INPUT TAX"
        wksLineA1.Cells(6, 21).Value = "REMARKS"
        wksLineA1.Cells(6, 22).Value = "CUSTOMER TYPE"
        wksLineA1.Cells(6, 23).Value = "FRONTLINER"
        With flxSum
            x = 6
            For i = 1 To .Rows - 1
                For j = 1 To 23
                    wksLineA1.Cells(x + 1, j).Value = .TextMatrix(i, j)
                Next j
            DoEvents
                x = x + 1
            Next i
        End With
        wksLineA1.Range("A6:W" & (x - 1)).Columns.AutoFit
        wksLineA1.Range("A1").Font.Bold = True
        wksLineA1.Range("A4").Font.Bold = True
        wksLineA1.Range("A6:W6").Font.Bold = True
        XLapp.DisplayAlerts = False
    ElseIf mySrc = "Service Receipts Book" Then
        wksLineA1.Name = "SRB"
        wksLineA1.Cells(1, 1).Value = "COMPANY NAME: " & SysBranch
        wksLineA1.Cells(2, 1).Value = "ADDRESS: " & SysAddress
        wksLineA1.Cells(4, 1).Value = "SERVICE RECEIPTS BOOK"
        wksLineA1.Cells(5, 1).Value = "DATE COVERED: " & Format(Dfrom.Value, "MMMM DD, YYYY") & " TO " & Format(Dto.Value, "MMMM DD, YYYY")
        wksLineA1.Cells(6, 1).Value = "DATE"
        wksLineA1.Cells(6, 2).Value = "CUSTOMER'S NAME"
        wksLineA1.Cells(6, 3).Value = "JO NO."
        wksLineA1.Cells(6, 4).Value = "OR NO."
        wksLineA1.Cells(6, 5).Value = "JO TYPE"
        wksLineA1.Cells(6, 6).Value = "JO AMOUNT"
        wksLineA1.Cells(6, 7).Value = "ACCOUNT RECEIVABLE"
        wksLineA1.Cells(6, 8).Value = "SERVICE INCOME"
        wksLineA1.Cells(6, 9).Value = "VAT OUTPUT"
        wksLineA1.Cells(6, 10).Value = "MECHANIC"
        wksLineA1.Cells(6, 11).Value = "REMARKS"
        wksLineA1.Cells(6, 12).Value = "MC MODEL"
        wksLineA1.Cells(6, 13).Value = "KM RDG."
        wksLineA1.Cells(6, 14).Value = "SERVICE TYPE"
        With flxSum
            x = 6
            For i = 1 To .Rows - 1
                For j = 1 To 14
                    wksLineA1.Cells(x + 1, j).Value = .TextMatrix(i, j)
                Next j
            DoEvents
                x = x + 1
            Next i
        End With
        wksLineA1.Range("A6:O" & (x - 1)).Columns.AutoFit
        wksLineA1.Range("A1").Font.Bold = True
        wksLineA1.Range("A4").Font.Bold = True
        wksLineA1.Range("A6:O6").Font.Bold = True
        XLapp.DisplayAlerts = False
    End If
    
    wbkMVR.SaveAs filename:=sSaveAsFileName, FileFormat:=xlNormal, Password:="", WriteResPassword:="", ReadOnlyRecommended:=False, CreateBackup:=False
    
    wbkMVR.Close SaveChanges:=False
    XLapp.Quit
    
    Screen.MousePointer = vbDefault
    Me.Enabled = True
    
    Set XLapp = Nothing
    Set wbkMVR = Nothing
    Set wksLineA1 = Nothing
    MsgBox "Finish Saving to Excel!", vbInformation + vbOKOnly, Me.Caption
    
    Exit Sub
er:
    Screen.MousePointer = vbDefault
    Me.Enabled = True
    MsgBox Err.Description, vbCritical, "Error!"
End Sub

Private Sub cmdTrans_Click(Index As Integer)
    Dim dCov As String, pQRY As String, JOid As String
    statUpdate.Visible = False
    Select Case Index
        Case 1  'Show Form
            With flxSum
                If .Rows = 2 And flxSum.TextMatrix(1, 1) = "" Then
                    MsgBox "There is no record to view!", vbInformation + vbOKOnly, "Confirm"
                    Exit Sub
                End If
                If mySrc = "Job Order" Then
                    Call JobOrder.DisplayJobOrder(flxSum.TextMatrix(ActRow, 14))
                    Unload Me
                ElseIf mySrc = "Manager's Cash Flow" Then
                    'If DateValue(.TextMatrix(ActRow, 1)) >= DateValue("05/10/2010") Then '5/10/2010-tagum
                    '    Call CashFlowMgr.ShowRecord(Val(.TextMatrix(ActRow, 5)))
                    '    Unload Me
                    'Else
                    '    CashFlowMgrOld.CFdate.Value = DateValue(.TextMatrix(ActRow, 1))
                    '    Call CashFlowMgrOld.ViewCashFlow
                    '    Unload Me
                    If SysBranch = "MOTO-R REPUBLIC CORP." Then
                        'If DateValue(.TextMatrix(ActRow, 1)) >= DateValue("05/10/2010") Then
                            Call CashFlowMgr.ShowRecord(Val(.TextMatrix(ActRow, 5)))
                            Unload Me
                        'Else
                        '    CashFlowMgrOld.CFdate.Value = DateValue(.TextMatrix(ActRow, 1))
                        '    Call CashFlowMgrOld.ViewCashFlow
                        'End If
                    ElseIf SysBranch = "LYR MOTORPARTS CORP. - ULAS" Then
                        'If DateValue(.TextMatrix(ActRow, 1)) >= DateValue("06/10/2010") Then
                            Call CashFlowMgr.ShowRecord(Val(.TextMatrix(ActRow, 5)))
                            Unload Me
                        'Else
                        '    CashFlowMgrOld.CFdate.Value = DateValue(.TextMatrix(ActRow, 1))
                        '    Call CashFlowMgrOld.ViewCashFlow
                        'End If
                    Else
                        Call CashFlowMgr.ShowRecord(Val(.TextMatrix(ActRow, 5)))
                        Unload Me
                    End If
                ElseIf mySrc = "Cashier Cash Flow" Then
                    'All Branches of Moto-r Davao
                    '5/10/2010-TAGUM
                    '6/10/2010-ULAS
                    '4/01/2011 Depot
                    '"7/25/2011" - VAN - LYR SHOP
                    If SysBranch = "MOTO-R REPUBLIC CORP." Then
                        'If DateValue(.TextMatrix(ActRow, 1)) >= DateValue("05/10/2010") Then
                            Call CashFlowCashier.ShowRecord(Val(.TextMatrix(ActRow, 5)))
                            Unload Me
                        'Else
                        '    CashFlowCashierOld.CFdate.Value = DateValue(.TextMatrix(ActRow, 1))
                        '    Call CashFlowCashierOld.ViewCashFlow
                        '    Unload Me
                        'End If
                    ElseIf SysBranch = "LYR MOTORPARTS CORP. - ULAS" Then
                        'If DateValue(.TextMatrix(ActRow, 1)) >= DateValue("06/10/2010") Then
                            Call CashFlowCashier.ShowRecord(Val(.TextMatrix(ActRow, 5)))
                            Unload Me
                        'Else
                        '    CashFlowCashierOld.CFdate.Value = DateValue(.TextMatrix(ActRow, 1))
                        '    Call CashFlowCashierOld.ViewCashFlow
                        '    Unload Me
                        'End If
                    Else
                        'Depot, Van, LYR SHOP
                        Call CashFlowCashier.ShowRecord(Val(.TextMatrix(ActRow, 5)))
                        Unload Me
                    End If
                    
                ElseIf mySrc = "SP Incoming Supplier" And .TextMatrix(ActRow, 6) <> "(ONLINE ENTRY)" Then
                    Call SPincomingSupplier.ShowRecord(.TextMatrix(ActRow, 9))
                    Unload Me
                ElseIf mySrc = "SP Incoming" Then
                    Call SPincoming.ShowRecord(.TextMatrix(ActRow, 9))
                    Unload Me
                ElseIf mySrc = "Clerical Adjustment" Then
                    Call ClericalADJ.DisplayClrADJ(flxSum.TextMatrix(ActRow, 2))
                    ClericalADJ.Show vbModal
                ElseIf mySrc = "DMCM [A/R]" Then
                    Call DMCMadj.ShowRecord(flxSum.TextMatrix(ActRow, 2))
                    Unload Me
                ElseIf mySrc = "Inventory Adjustment" Then
                    Call InvntryADJ.DisplayInvAdj(flxSum.TextMatrix(ActRow, 2))
                    Unload Me
                ElseIf mySrc = "Customer P.O." Then
                    Call POCust.ShowRecord(Val(flxSum.TextMatrix(ActRow, 8)))
                    Unload Me
                ElseIf mySrc = "Top Selling Spare Parts" Then
                    SPTopSales.Show
                    Call SPTopSales.ShowRecord(flxSum.TextMatrix(ActRow, 6))
                    Unload Me
                ElseIf mySrc = "Purchase Return" Then
                    Call PurchaseReturn.DisplayPurchaseReturn(flxSum.TextMatrix(ActRow, 9))
                    Unload Me
                ElseIf mySrc = "Requisition" Then
                    Call Requisition.DisplayRequisition(flxSum.TextMatrix(ActRow, 2))
                    Requisition.Show vbModal
                ElseIf mySrc = "Sales Return" Then
                    'Debug
                    Call SalesReturn.DisplaySalesReturn(flxSum.TextMatrix(ActRow, 11))
                    Unload Me
                ElseIf mySrc = "Location Adjustment" Then
                    Call SPLocAdj.DisplayLocAdj(flxSum.TextMatrix(ActRow, 8))
                    Unload Me
                ElseIf mySrc = "Received Spare Parts" Or mySrc = "Received Spare Parts DVOWH" Then
                    'If InStr(flxSum.TextMatrix(ActRow, 3), "WAREHOUSE") Then
                    If flxSum.TextMatrix(ActRow, 17) = 0 Then
                        If mySrc = "Received Spare Parts DVOWH" Then
                            Call SPincomingWhAgdao.ShowRecord(flxSum.TextMatrix(ActRow, 2))
                        Else
                            Call SPincoming.ShowRecord(flxSum.TextMatrix(ActRow, 2))
                        End If
                    Else
                        Call SPbranchIncoming.ShowRecord(flxSum.TextMatrix(ActRow, 2))
                    End If
                    Unload Me
                'ElseIf mySrc = "Sales Order" Then
                    'Call SPSorder.DisplaySalesOrder(flxSum.TextMatrix(ActRow, 7))
                    'Unload Me
                    
                ElseIf mySrc = "Mechanic Income" Then
                    '***Access
                    'Qry = "SELECT SRVCJOBORDER.JOBORDERID From SRVCJOBORDER WHERE (((SRVCJOBORDER.JOBORDERNO)='" & flxSVC.TextMatrix(ActRow, 2) & "'));"
                    'mySQL
                    Qry = "SELECT SRVCJOBORDER.JOBORDERID From SRVCJOBORDER WHERE SRVCJOBORDER.JOBORDERNO='" & flxSum.TextMatrix(ActRow, 2) & "';"
                    JOid = GetFieldValue(Qry)
                    Call JobOrder.DisplayJobOrder(JOid)
                    Unload Me
                ElseIf mySrc = "Cash Voucher" Then
                    Call CashVoucher.ShowRecord(flxSum.TextMatrix(ActRow, 2))
                    CashVoucher.Show vbModal
                ElseIf mySrc = "Advances for Liquidation" Then
                    ExpAdvLiq.SrcSumm = True
                    ExpAdvLiq.txtAFLid.Text = .TextMatrix(ActRow, 9)
                    Call ExpAdvLiq.ShowRecord
                    ExpAdvLiq.Show vbModal
                ElseIf mySrc = "Cash Return" Then
                    ExpCashRet.SrcSumm = True
                    ExpCashRet.txtCRid.Text = .TextMatrix(ActRow, 7)
                    Call ExpCashRet.ShowRecord
                    ExpCashRet.Show vbModal
                ElseIf mySrc = "Branch Purchase Order" Then
                    Call BranchPO.ShowRecordOT(.TextMatrix(ActRow, 1))
                    Unload Me
                ElseIf mySrc = "Trust Receipt Agreement" Then
                    Call TRA.ShowRecord(Val(.TextMatrix(ActRow, 10)))
                    Unload Me
                ElseIf mySrc = "Branch Order" Then
                    If Left(.TextMatrix(ActRow, 2), 1) = "R" Then
                        Call BranchReplPO.DisplaySPOrder(.TextMatrix(ActRow, 2))
                    Else
                        Call SPordering.DisplaySPOrder(.TextMatrix(ActRow, 2))
                    End If
                    Unload Me
'************************ CODE UPDATED: 08/04/17 ******************************
                ElseIf mySrc = "SETTLEMENT DISCREPANCY" Then
                    Call frmSDiscrepancy.ShowRecord(.TextMatrix(ActRow, 9), .TextMatrix(ActRow, 12))
                    Unload Me
                ElseIf mySrc = "SDiscrepancy Adjustment" Then
                    Call SDiscInvtryAdj.DisplayInvAdj(flxSum.TextMatrix(ActRow, 2))
                    Unload Me
'******************** END CODE UPDATED: 08/04/17 ******************************
                ElseIf mySrc = "ADJUSTMENT REQUEST" Then
                    Call frmOtherAdjustment.ShowRecord(flxSum.TextMatrix(ActRow, 3))
                    Unload Me
                ElseIf mySrc = "CANCELLATION MODULE" Then
                    Call frmCancellation.ShowRecord(flxSum.TextMatrix(ActRow, 2))
                    frmCancellation.Show vbModal
                End If
            End With
            
        Case 2  'Print
            If mySrc = "Job Order" Then
                dCov = "Date Covered: " & Format(Dfrom.Value, "Mmm. dd, yyyy") & IIf(Val(cmdNext.Tag) = 2, " to " & Format(Dto.Value, "Mmm. dd, yyyy"), "")
                If cboRebuilt.ListIndex <> -1 Then
                    If cmdNext.Tag = "1" Then
                        If cboRebuilt.Text = "View All" Then
                            Qry = "SELECT SRVCJOBORDER.DATEPROMISED, SRVCJOBORDER.JOBORDERNO, SRVCJOBORDER.CUSTOMER, SRVCJOBORDER.PAYTYPE, SRVCJOBORDER.TOTALCOST, SRVCJOBORDER.ORNO, SRVCUSTOMER.MODEL, SRVCJOBORDER.KMREADING, SRVCJOBORDER.JOTYPE, SRVCJOBORDER.MECHANIC, SRVCJOBORDER.REMARKS, SRVCJOBORDER.DISCOUNT " & _
                                  "FROM SRVCJOBORDER INNER JOIN SRVCUSTOMER ON SRVCJOBORDER.ACCOUNTNO = SRVCUSTOMER.ACCOUNTNO Where SRVCJOBORDER.DATEPROMISED = '" & Format(Dfrom.Value, "yyyy-mm-dd") & "' And SRVCJOBORDER.PRINTCOPY > 0 ORDER BY SRVCJOBORDER.DATEPROMISED, SRVCJOBORDER.JOBORDERNO;"
                        ElseIf cboRebuilt.Text = "View Machine Shop Only" Then
                            Qry = "SELECT SRVCJOBORDER.DATEPROMISED, SRVCJOBORDER.JOBORDERNO, SRVCJOBORDER.CUSTOMER, SRVCJOBORDER.PAYTYPE, SRVCJOBORDER.TOTALCOST, SRVCJOBORDER.ORNO, SRVCUSTOMER.MODEL, SRVCJOBORDER.KMREADING, SRVCJOBORDER.JOTYPE, SRVCJOBORDER.MECHANIC, SRVCJOBORDER.REMARKS, SRVCJOBORDER.DISCOUNT " & _
                                  "FROM SRVCJOBORDER INNER JOIN SRVCUSTOMER ON SRVCJOBORDER.ACCOUNTNO = SRVCUSTOMER.ACCOUNTNO Where SRVCJOBORDER.DATEPROMISED = '" & Format(Dfrom.Value, "yyyy-mm-dd") & "' And SRVCJOBORDER.JOTYPE = 'MACHINE SHOP' And SRVCJOBORDER.PRINTCOPY > 0 ORDER BY SRVCJOBORDER.DATEPROMISED, SRVCJOBORDER.JOBORDERNO;"
                        ElseIf cboRebuilt.Text = "Exclude Machine Shop" Then
                            Qry = "SELECT SRVCJOBORDER.DATEPROMISED, SRVCJOBORDER.JOBORDERNO, SRVCJOBORDER.CUSTOMER, SRVCJOBORDER.PAYTYPE, SRVCJOBORDER.TOTALCOST, SRVCJOBORDER.ORNO, SRVCUSTOMER.MODEL, SRVCJOBORDER.KMREADING, SRVCJOBORDER.JOTYPE, SRVCJOBORDER.MECHANIC, SRVCJOBORDER.REMARKS, SRVCJOBORDER.DISCOUNT " & _
                                  "FROM SRVCJOBORDER INNER JOIN SRVCUSTOMER ON SRVCJOBORDER.ACCOUNTNO = SRVCUSTOMER.ACCOUNTNO Where SRVCJOBORDER.DATEPROMISED = '" & Format(Dfrom.Value, "yyyy-mm-dd") & "' And SRVCJOBORDER.JOTYPE <> 'MACHINE SHOP' And SRVCJOBORDER.PRINTCOPY > 0 ORDER BY SRVCJOBORDER.DATEPROMISED, SRVCJOBORDER.JOBORDERNO;"
                        End If
                    Else
                        If cboRebuilt.Text = "View All" Then
                            Qry = "SELECT SRVCJOBORDER.DATEPROMISED, SRVCJOBORDER.JOBORDERNO, SRVCJOBORDER.CUSTOMER, SRVCJOBORDER.PAYTYPE, SRVCJOBORDER.TOTALCOST, SRVCJOBORDER.ORNO, SRVCUSTOMER.MODEL, SRVCJOBORDER.KMREADING, SRVCJOBORDER.JOTYPE, SRVCJOBORDER.MECHANIC, SRVCJOBORDER.REMARKS, SRVCJOBORDER.DISCOUNT " & _
                                  "FROM SRVCJOBORDER INNER JOIN SRVCUSTOMER ON SRVCJOBORDER.ACCOUNTNO = SRVCUSTOMER.ACCOUNTNO Where SRVCJOBORDER.DATEPROMISED >= '" & Format(Dfrom.Value, "yyyy-mm-dd") & "' And SRVCJOBORDER.DATEPROMISED <= '" & Format(Dto.Value, "yyyy-mm-dd") & "' And SRVCJOBORDER.PRINTCOPY > 0 ORDER BY SRVCJOBORDER.DATEPROMISED, SRVCJOBORDER.JOBORDERNO;"
                        ElseIf cboRebuilt.Text = "View Machine Shop Only" Then
                            Qry = "SELECT SRVCJOBORDER.DATEPROMISED, SRVCJOBORDER.JOBORDERNO, SRVCJOBORDER.CUSTOMER, SRVCJOBORDER.PAYTYPE, SRVCJOBORDER.TOTALCOST, SRVCJOBORDER.ORNO, SRVCUSTOMER.MODEL, SRVCJOBORDER.KMREADING, SRVCJOBORDER.JOTYPE, SRVCJOBORDER.MECHANIC, SRVCJOBORDER.REMARKS, SRVCJOBORDER.DISCOUNT " & _
                                  "FROM SRVCJOBORDER INNER JOIN SRVCUSTOMER ON SRVCJOBORDER.ACCOUNTNO = SRVCUSTOMER.ACCOUNTNO Where SRVCJOBORDER.DATEPROMISED >= '" & Format(Dfrom.Value, "yyyy-mm-dd") & "' And SRVCJOBORDER.DATEPROMISED <= '" & Format(Dto.Value, "yyyy-mm-dd") & "' And SRVCJOBORDER.JOTYPE = 'MACHINE SHOP' And SRVCJOBORDER.PRINTCOPY > 0 ORDER BY SRVCJOBORDER.DATEPROMISED, SRVCJOBORDER.JOBORDERNO;"
                        ElseIf cboRebuilt.Text = "Exclude Machine Shop" Then
                            Qry = "SELECT SRVCJOBORDER.DATEPROMISED, SRVCJOBORDER.JOBORDERNO, SRVCJOBORDER.CUSTOMER, SRVCJOBORDER.PAYTYPE, SRVCJOBORDER.TOTALCOST, SRVCJOBORDER.ORNO, SRVCUSTOMER.MODEL, SRVCJOBORDER.KMREADING, SRVCJOBORDER.JOTYPE, SRVCJOBORDER.MECHANIC, SRVCJOBORDER.REMARKS, SRVCJOBORDER.DISCOUNT " & _
                                  "FROM SRVCJOBORDER INNER JOIN SRVCUSTOMER ON SRVCJOBORDER.ACCOUNTNO = SRVCUSTOMER.ACCOUNTNO Where SRVCJOBORDER.DATEPROMISED >= '" & Format(Dfrom.Value, "yyyy-mm-dd") & "' And SRVCJOBORDER.DATEPROMISED <= '" & Format(Dto.Value, "yyyy-mm-dd") & "' And SRVCJOBORDER.JOTYPE <> 'MACHINE SHOP' And SRVCJOBORDER.PRINTCOPY > 0 ORDER BY SRVCJOBORDER.DATEPROMISED, SRVCJOBORDER.JOBORDERNO;"
                        End If
                        
                    End If
                Else
                    If Val(cmdNext.Tag) = 1 Then
                        '***Access
                        'Qry = "SELECT SRVCJOBORDER.DATEPROMISED, SRVCJOBORDER.JOBORDERNO, SRVCJOBORDER.CUSTOMER, SRVCJOBORDER.PAYTYPE, SRVCJOBORDER.TOTALCOST, SRVCJOBORDER.ORNO, SRVCUSTOMER.MODEL, SRVCJOBORDER.KMREADING, SRVCJOBORDER.JOTYPE, SRVCJOBORDER.MECHANIC, SRVCJOBORDER.REMARKS, SRVCJOBORDER.DISCOUNT " & _
                              "FROM SRVCJOBORDER INNER JOIN SRVCUSTOMER ON SRVCJOBORDER.ACCOUNTNO = SRVCUSTOMER.ACCOUNTNO Where (((SRVCJOBORDER.DATEPROMISED) = #" & dFrom.Value & "#) And ((SRVCJOBORDER.PRINTCOPY) > 0)) ORDER BY SRVCJOBORDER.DATEPROMISED, SRVCJOBORDER.JOBORDERNO;"
                        'mySQL
                        Qry = "SELECT SRVCJOBORDER.DATEPROMISED, SRVCJOBORDER.JOBORDERNO, SRVCJOBORDER.CUSTOMER, SRVCJOBORDER.PAYTYPE, SRVCJOBORDER.TOTALCOST, SRVCJOBORDER.ORNO, SRVCUSTOMER.MODEL, SRVCJOBORDER.KMREADING, SRVCJOBORDER.JOTYPE, SRVCJOBORDER.MECHANIC, SRVCJOBORDER.REMARKS, SRVCJOBORDER.DISCOUNT " & _
                              "FROM SRVCJOBORDER INNER JOIN SRVCUSTOMER ON SRVCJOBORDER.ACCOUNTNO = SRVCUSTOMER.ACCOUNTNO Where SRVCJOBORDER.DATEPROMISED = '" & Format(Dfrom.Value, "yyyy-mm-dd") & "' And SRVCJOBORDER.PRINTCOPY > 0 ORDER BY SRVCJOBORDER.DATEPROMISED, SRVCJOBORDER.JOBORDERNO;"
                        
                    Else
                        '***Access
                        'Qry = "SELECT SRVCJOBORDER.DATEPROMISED, SRVCJOBORDER.JOBORDERNO, SRVCJOBORDER.CUSTOMER, SRVCJOBORDER.PAYTYPE, SRVCJOBORDER.TOTALCOST, SRVCJOBORDER.ORNO, SRVCUSTOMER.MODEL, SRVCJOBORDER.KMREADING, SRVCJOBORDER.JOTYPE, SRVCJOBORDER.MECHANIC, SRVCJOBORDER.REMARKS, SRVCJOBORDER.DISCOUNT " & _
                              "FROM SRVCJOBORDER INNER JOIN SRVCUSTOMER ON SRVCJOBORDER.ACCOUNTNO = SRVCUSTOMER.ACCOUNTNO Where (((SRVCJOBORDER.DATEPROMISED) >= #" & dFrom.Value & "# And (SRVCJOBORDER.DATEPROMISED) <= #" & dTo.Value & "#) And ((SRVCJOBORDER.PRINTCOPY) > 0)) ORDER BY SRVCJOBORDER.DATEPROMISED, SRVCJOBORDER.JOBORDERNO;"
                        'mySQl
                        Qry = "SELECT SRVCJOBORDER.DATEPROMISED, SRVCJOBORDER.JOBORDERNO, SRVCJOBORDER.CUSTOMER, SRVCJOBORDER.PAYTYPE, SRVCJOBORDER.TOTALCOST, SRVCJOBORDER.ORNO, SRVCUSTOMER.MODEL, SRVCJOBORDER.KMREADING, SRVCJOBORDER.JOTYPE, SRVCJOBORDER.MECHANIC, SRVCJOBORDER.REMARKS, SRVCJOBORDER.DISCOUNT " & _
                              "FROM SRVCJOBORDER INNER JOIN SRVCUSTOMER ON SRVCJOBORDER.ACCOUNTNO = SRVCUSTOMER.ACCOUNTNO Where SRVCJOBORDER.DATEPROMISED >= '" & Format(Dfrom.Value, "yyyy-mm-dd") & "' And SRVCJOBORDER.DATEPROMISED <= '" & Format(Dto.Value, "yyyy-mm-dd") & "' And SRVCJOBORDER.PRINTCOPY > 0 ORDER BY SRVCJOBORDER.DATEPROMISED, SRVCJOBORDER.JOBORDERNO;"
                        
                    End If
                End If
                PRNmc.QryPrn = Qry
                PRNmc.ParamValue1 = dCov
                PRNmc.ParamValue2 = "Prepared By: " & GetUserName
                PRNmc.MySource = "JOSUMMARY"
                PRNmc.Show
            
            ElseIf mySrc = "SP Incoming" Or mySrc = "SP Incoming" Then
                If Val(cmdNext.Tag) = 2 Then
                    dCov = "Date Covered: " & Dfrom.Value & " to " & Dto.Value
                    '***Access
                    'Qry = "SELECT SPRECEIVE.DATEREC, SPRECEIVE.PACKINGNO, SPRECEIVE.SUPPLIER, SPRECEIVE.SHIPPER, SPRECEIVE.WAYBILLNO, SPRECEIVE.PACKDATE, SPRECEIVE.REMARKS, SPRECEIVE.CHECKEDBY, SPRECEIVE.PRINTCOPY, SPRECEIVE.SPRRNO " & _
                          "From SPRECEIVE Where (((SPRECEIVE.DATEREC) >= #" & dFrom.Value & "# And (SPRECEIVE.DATEREC) <= #" & dTo.Value & "#)) ORDER BY SPRECEIVE.DATEREC, SPRECEIVE.PACKINGNO DESC;"
                    'mySQl
                    Qry = "SELECT SPRECEIVE.DATEREC, SPRECEIVE.PACKINGNO, SPRECEIVE.SUPPLIER, SPRECEIVE.SHIPPER, SPRECEIVE.WAYBILLNO, SPRECEIVE.PACKDATE, SPRECEIVE.REMARKS, SPRECEIVE.CHECKEDBY, SPRECEIVE.PRINTCOPY, SPRECEIVE.SPRRNO " & _
                          "From SPRECEIVE Where SPRECEIVE.DATEREC >= '" & Format(Dfrom.Value, "yyyy-mm-dd") & "' And SPRECEIVE.DATEREC <= '" & Format(Dto.Value, "yyyy-mm-dd") & "' ORDER BY SPRECEIVE.DATEREC, SPRECEIVE.PACKINGNO DESC;"
                    'Stop here
                Else
                    '***Access
                    'Qry = "SELECT SPRECEIVE.DATEREC, SPRECEIVE.PACKINGNO, SPRECEIVE.SUPPLIER, SPRECEIVE.SHIPPER, SPRECEIVE.WAYBILLNO, SPRECEIVE.PACKDATE, SPRECEIVE.REMARKS, SPRECEIVE.CHECKEDBY, SPRECEIVE.PRINTCOPY, SPRECEIVE.SPRRNO " & _
                          "From SPRECEIVE Where (((SPRECEIVE.DATEREC) = #" & dFrom.Value & "#)) ORDER BY SPRECEIVE.DATEREC, SPRECEIVE.PACKINGNO DESC;"
                    'mySQL
                    Qry = "SELECT SPRECEIVE.DATEREC, SPRECEIVE.PACKINGNO, SPRECEIVE.SUPPLIER, SPRECEIVE.SHIPPER, SPRECEIVE.WAYBILLNO, SPRECEIVE.PACKDATE, SPRECEIVE.REMARKS, SPRECEIVE.CHECKEDBY, SPRECEIVE.PRINTCOPY, SPRECEIVE.SPRRNO " & _
                          "From SPRECEIVE Where SPRECEIVE.DATEREC = '" & Format(Dfrom.Value, "yyyy-mm-dd") & "' ORDER BY SPRECEIVE.DATEREC, SPRECEIVE.PACKINGNO DESC;"
                    
                    dCov = "Date Covered: " & Dfrom.Value
                End If
                PRNmc.MySource = "SPRECEIVESUMM"
                PRNmc.QryPrn = Qry
                PRNmc.ParamValue1 = dCov
                PRNmc.Show
    
            ElseIf mySrc = "Sales Return" Then
                dCov = "Date Covered: " & Format(Dfrom.Value, "Mmm. dd, yyyy") & IIf(Val(cmdNext.Tag) = 2, " to " & Format(Dto.Value, "Mmm. dd, yyyy"), "")
                If Val(cmdNext.Tag) = 1 Then
                    '***Access
                    'Qry = "SELECT SRETURNSP.RETDATE, SRETURNSP.RETNO, SRETURNSP.INVOICE, SPINVOICE.CHARGETO, SRETURNSP.REASON, SRETURNSP.TOTALAMT, SRETURNSP.REMARKS " & _
                          "FROM SRETURNSP INNER JOIN SPINVOICE ON SRETURNSP.INVOICE = SPINVOICE.SPINVNO Where (((SRETURNSP.RETDATE) = #" & dFrom.Value & "#)) ORDER BY SRETURNSP.RETDATE, SRETURNSP.RETNO;"
                    'mySQl
                    'Qry = "SELECT SRETURNSP.RETDATE, SRETURNSP.RETNO, SRETURNSP.INVOICE, SPINVOICE.CHARGETO, SRETURNSP.REASON, SRETURNSP.TOTALAMT, SRETURNSP.REMARKS " & _
                          "FROM SRETURNSP INNER JOIN SPINVOICE ON SRETURNSP.INVOICE = SPINVOICE.SPINVNO Where SRETURNSP.RETDATE = '" & Format(dFrom.Value, "yyyy-mm-dd") & "' ORDER BY SRETURNSP.RETDATE, SRETURNSP.RETNO;"
                    Qry = "SELECT " & _
                        "SRETURNSP.RETDATE, " & _
                        "SRETURNSP.RETNO, " & _
                        "SRETURNSP.INVOICE, " & _
                        "SPINVOICE.INVDATE, " & _
                        "SPINVOICE.CHARGETO, " & _
                        "Cast(IF(MONTH(SRETURNSP.RETDATE) = MONTH(SPINVOICE.INVDATE), SRETURNSP.TOTALAMT, '') as char) AS THMOS, " & _
                        "Cast(IF(MONTH(SRETURNSP.RETDATE) = MONTH(SPINVOICE.INVDATE),'',SRETURNSP.TOTALAMT) as char) AS PRMOS, " & _
                        "SRETURNSP.TOTALAMT, " & _
                        "SRETURNSP.REMARKS, " & _
                        "SRETURNSP.PRINTCOPY, " & _
                        "SRETURNSP.RETURNID " & _
                    "From " & _
                        "SRETURNSP " & _
                        "INNER JOIN SPINVOICE ON (SRETURNSP.INVOICE = SPINVOICE.SPINVNO) " & _
                    "Where " & _
                        "SRETURNSP.RETDATE >= '" & Format(Dfrom.Value, "yyyy-mm-dd") & "' " & _
                    "Order By " & _
                        "SRETURNSP.RETDATE, " & _
                        "SRETURNSP.RETNO Desc"
        
                Else
                    '***Access
                    'Qry = "SELECT SRETURNSP.RETDATE, SRETURNSP.RETNO, SRETURNSP.INVOICE, SPINVOICE.CHARGETO, SRETURNSP.REASON, SRETURNSP.TOTALAMT, SRETURNSP.REMARKS " & _
                          "FROM SRETURNSP INNER JOIN SPINVOICE ON SRETURNSP.INVOICE = SPINVOICE.SPINVNO Where (((SRETURNSP.RETDATE) >= #" & dFrom.Value & "# And (SRETURNSP.RETDATE) <= #" & dTo.Value & "#)) ORDER BY SRETURNSP.RETDATE, SRETURNSP.RETNO;"
                    'mySQl
                    'Qry = "SELECT SRETURNSP.RETDATE, SRETURNSP.RETNO, SRETURNSP.INVOICE, SPINVOICE.CHARGETO, SRETURNSP.REASON, SRETURNSP.TOTALAMT, SRETURNSP.REMARKS " & _
                          "FROM SRETURNSP INNER JOIN SPINVOICE ON SRETURNSP.INVOICE = SPINVOICE.SPINVNO Where SRETURNSP.RETDATE >= '" & Format(dFrom.Value, "yyyy-mm-dd") & "' And SRETURNSP.RETDATE <= '" & Format(dTo.Value, "yyyy-mm-dd") & "' ORDER BY SRET
                    Qry = "SELECT " & _
                        "SRETURNSP.RETDATE, " & _
                        "SRETURNSP.RETNO, " & _
                        "SRETURNSP.INVOICE, " & _
                        "SPINVOICE.INVDATE, " & _
                        "SPINVOICE.CHARGETO, " & _
                        "Cast(IF(MONTH(SRETURNSP.RETDATE) = MONTH(SPINVOICE.INVDATE), SRETURNSP.TOTALAMT, '') as char) AS THMOS, " & _
                        "Cast(IF(MONTH(SRETURNSP.RETDATE) = MONTH(SPINVOICE.INVDATE),'',SRETURNSP.TOTALAMT) as char) AS PRMOS, " & _
                        "SRETURNSP.TOTALAMT, " & _
                        "SRETURNSP.REMARKS, " & _
                        "SRETURNSP.PRINTCOPY, " & _
                        "SRETURNSP.RETURNID " & _
                    "From " & _
                        "SRETURNSP " & _
                        "INNER JOIN SPINVOICE ON (SRETURNSP.INVOICE = SPINVOICE.SPINVNO) " & _
                    "Where " & _
                        "SRETURNSP.RETDATE >= '" & Format(Dfrom.Value, "yyyy-mm-dd") & "' AND " & _
                        "SRETURNSP.RETDATE <= '" & Format(Dto.Value, "yyyy-mm-dd") & "' " & _
                        "And Case When SRETURNSP.INVTYPE = 0 Then spinvoice.INVTYPE = 'CREDIT' Else spinvoice.INVTYPE = 'CASH' End " & _
                    "Order By " & _
                        "SRETURNSP.RETDATE, " & _
                        "SRETURNSP.RETNO Desc"
                End If
                PRNmc.QryPrn = Qry
                PRNmc.ParamValue1 = dCov
                '***Access
                'pQRY = "SELECT SETTINGS.VALUE From SETTINGS " & _
                       "WHERE (((SETTINGS.DESCRIPTION)='EMP DEFAULT VALUE') AND ((SETTINGS.REMARKS)='OIC'));"
                'mySQl
                pQRY = "SELECT SETTINGS.VALUE From SETTINGS " & _
                       "WHERE SETTINGS.DESCRIPTION='EMP DEFAULT VALUE' AND SETTINGS.REMARKS='OIC';"
                
                PRNmc.ParamValue2 = GetFieldValue(pQRY)
                '***Access
                'pQRY = "SELECT SETTINGS.VALUE From SETTINGS " & _
                       "WHERE (((SETTINGS.DESCRIPTION)='EMP DEFAULT VALUE') AND ((SETTINGS.REMARKS)='MANAGER'));"
                'mySQL
                pQRY = "SELECT SETTINGS.VALUE From SETTINGS " & _
                       "WHERE SETTINGS.DESCRIPTION='EMP DEFAULT VALUE' AND SETTINGS.REMARKS='MANAGER';"
                
                PRNmc.ParamValue3 = GetFieldValue(pQRY)
                PRNmc.ParamValue4 = "BRANCH MANAGER"
                PRNmc.ParamValue5 = "AREA OPERATION MANAGER"
                PRNmc.MySource = "SALESRETURNSUMMARY"
                PRNmc.Show
                
                
            ElseIf mySrc = "Received Spare Parts" Or mySrc = "Received Spare Parts DVOWH" Then
                If Val(cmdNext.Tag) = 2 Then
                    dCov = "Date Covered: " & Dfrom.Value & " to " & Dto.Value
                    '***Access
                    'Qry = "SELECT SPRECEIVE.DATEREC, SPRECEIVE.PACKINGNO, SPRECEIVE.SUPPLIER, SPRECEIVE.SHIPPER, SPRECEIVE.WAYBILLNO, SPRECEIVE.PACKDATE, SPRECEIVE.REMARKS, SPRECEIVE.CHECKEDBY, SPRECEIVE.PRINTCOPY, SPRECEIVE.SPRRNO " & _
                          "From SPRECEIVE Where (((SPRECEIVE.DATEREC) >= #" & dFrom.Value & "# And (SPRECEIVE.DATEREC) <= #" & dTo.Value & "#)) ORDER BY SPRECEIVE.DATEREC, SPRECEIVE.PACKINGNO DESC;"
                    'mySQl
                    'Qry = "SELECT SPRECEIVE.DATEREC, SPRECEIVE.PACKINGNO, SPRECEIVE.SUPPLIER, SPRECEIVE.SHIPPER, SPRECEIVE.WAYBILLNO, SPRECEIVE.PACKDATE, SPRECEIVE.REMARKS, SPRECEIVE.CHECKEDBY, SPRECEIVE.PRINTCOPY, SPRECEIVE.SPRRNO " & _
                    '      "From SPRECEIVE Where SPRECEIVE.DATEREC >= '" & Format(dFrom.Value, "yyyy-mm-dd") & "' And SPRECEIVE.DATEREC <= '" & Format(dTo.Value, "yyyy-mm-dd") & "' ORDER BY SPRECEIVE.DATEREC, SPRECEIVE.PACKINGNO DESC;"
                    'Stop here
                    
                    Qry = "SELECT SPRECEIVE.DATEREC, SPRECEIVE.PACKINGNO, SPRECEIVE.SUPPLIER, SPRECEIVE.SHIPPER, SPRECEIVE.POREF, SPRECEIVE.PACKDATE, SPRECEIVE.REMARKS, " & _
                                 "COUNT(itr.SPAREPARTID) as TOTALPN, SUM(itr.RECEIVE) as TOTALQTY, IFNULL(sprec.DAMAGE,0) as DAMAGE, " & _
                                 "IFNULL(sprec.LACKING,0) as LACKING, IFNULL(sprec.EXCESS,0) as EXCESS, IFNULL(sprec.WRONGPN,0) as WRONGPN, " & _
                                 "SPRECEIVE.CHECKEDBY , SPRECEIVE.PRINTCOPY, SPRECEIVE.SPRRNO, SPRECEIVE.AutoLoad " & _
                          "From SPRECEIVE " & _
                          "LEFT JOIN SPITR as itr ON itr.REFERENCE = spreceive.PACKINGNO " & _
                          "LEFT JOIN (Select SUM(DAMAGE) as DAMAGE, SUM(LACKING) as LACKING, SUM(EXCESS) as EXCESS, SUM(WRONGPN) as WRONGPN, SPAREPARTID, PACKINGNO FROM sprecdiscrepancy GROUP BY PACKINGNO) as sprec ON sprec.packingNO = itr.REFERENCE " & _
                          "Where SPRECEIVE.DATEREC >= '" & Format(Dfrom.Value, "yyyy-mm-dd") & "' And SPRECEIVE.DATEREC <= '" & Format(Dto.Value, "yyyy-mm-dd") & "' " & _
                          "GROUP BY itr.REFERENCE ORDER BY SPRECEIVE.DATEREC, SPRECEIVE.PACKINGNO DESC;"
                Else
                    '***Access
                    'Qry = "SELECT SPRECEIVE.DATEREC, SPRECEIVE.PACKINGNO, SPRECEIVE.SUPPLIER, SPRECEIVE.SHIPPER, SPRECEIVE.WAYBILLNO, SPRECEIVE.PACKDATE, SPRECEIVE.REMARKS, SPRECEIVE.CHECKEDBY, SPRECEIVE.PRINTCOPY, SPRECEIVE.SPRRNO " & _
                          "From SPRECEIVE Where (((SPRECEIVE.DATEREC) = #" & dFrom.Value & "#)) ORDER BY SPRECEIVE.DATEREC, SPRECEIVE.PACKINGNO DESC;"
                    'mySQL
                    'Qry = "SELECT SPRECEIVE.DATEREC, SPRECEIVE.PACKINGNO, SPRECEIVE.SUPPLIER, SPRECEIVE.SHIPPER, SPRECEIVE.WAYBILLNO, SPRECEIVE.PACKDATE, SPRECEIVE.REMARKS, SPRECEIVE.CHECKEDBY, SPRECEIVE.PRINTCOPY, SPRECEIVE.SPRRNO " & _
                    '      "From SPRECEIVE Where SPRECEIVE.DATEREC = '" & Format(dFrom.Value, "yyyy-mm-dd") & "' ORDER BY SPRECEIVE.DATEREC, SPRECEIVE.PACKINGNO DESC;"
                    Qry = "SELECT SPRECEIVE.DATEREC, SPRECEIVE.PACKINGNO, SPRECEIVE.SUPPLIER, SPRECEIVE.SHIPPER, SPRECEIVE.POREF, SPRECEIVE.PACKDATE, SPRECEIVE.REMARKS, " & _
                                 "COUNT(itr.SPAREPARTID) as TOTALPN, SUM(itr.RECEIVE) as TOTALQTY, IFNULL(sprec.DAMAGE,0) as DAMAGE, " & _
                                 "IFNULL(sprec.LACKING,0) as LACKING, IFNULL(sprec.EXCESS,0) as EXCESS, IFNULL(sprec.WRONGPN,0) as WRONGPN, " & _
                                 "SPRECEIVE.CHECKEDBY , SPRECEIVE.PRINTCOPY, SPRECEIVE.SPRRNO, SPRECEIVE.AutoLoad " & _
                          "From SPRECEIVE " & _
                          "LEFT JOIN SPITR as itr ON itr.REFERENCE = spreceive.PACKINGNO " & _
                          "LEFT JOIN (Select SUM(DAMAGE) as DAMAGE, SUM(LACKING) as LACKING, SUM(EXCESS) as EXCESS, SUM(WRONGPN) as WRONGPN, SPAREPARTID, PACKINGNO FROM sprecdiscrepancy GROUP BY PACKINGNO) as sprec ON sprec.packingNO = itr.REFERENCE " & _
                          "Where SPRECEIVE.DATEREC = '" & Format(Dfrom.Value, "yyyy-mm-dd") & "' " & _
                          "GROUP BY itr.REFERENCE ORDER BY SPRECEIVE.DATEREC, SPRECEIVE.PACKINGNO DESC;"
                          
                          
                    dCov = "Date Covered: " & Dfrom.Value
                End If
                PRNmc.MySource = "SPRECEIVESUMM"
                PRNmc.QryPrn = Qry
                PRNmc.ParamValue1 = dCov
                PRNmc.Show
            ElseIf mySrc = "Mechanic Income" Then
                If Val(cmdNext.Tag) = 1 Then
                    '***Access
                    'Qry = "SELECT SRVCJOBORDER.MECHANIC, SRVCJOBORDER.JOBORDERNO, SRVCJOBORDER.DATEPROMISED, SRVCJOBORDER.JOTYPE, SRVCJOBORDER.TOTALCOST From SRVCJOBORDER " & _
                          "Where (((SRVCJOBORDER.DATEPROMISED) = #" & dFrom.Value & "#) And ((SRVCJOBORDER.PRINTCOPY) > 0)) ORDER BY SRVCJOBORDER.MECHANIC, SRVCJOBORDER.DATEPROMISED;"
                    'mySQl
                    Qry = "SELECT SRVCJOBORDER.MECHANIC, SRVCJOBORDER.JOBORDERNO, SRVCJOBORDER.DATEPROMISED, SRVCJOBORDER.JOTYPE, SRVCJOBORDER.TOTALCOST From SRVCJOBORDER " & _
                          "Where SRVCJOBORDER.DATEPROMISED = '" & Format(Dfrom.Value, "yyyy-mm-dd") & "' And SRVCJOBORDER.PRINTCOPY > 0 And SRVCJOBORDER.REMARKS <> 'CANCELLED' ORDER BY SRVCJOBORDER.MECHANIC, SRVCJOBORDER.DATEPROMISED;"
                    
                Else
                    '***Access
                    'Qry = "SELECT SRVCJOBORDER.MECHANIC, SRVCJOBORDER.JOBORDERNO, SRVCJOBORDER.DATEPROMISED, SRVCJOBORDER.JOTYPE, SRVCJOBORDER.TOTALCOST From SRVCJOBORDER " & _
                          "Where (((SRVCJOBORDER.DATEPROMISED) >= #" & dFrom.Value & "# And (SRVCJOBORDER.DATEPROMISED) <= #" & dTo.Value & "#) And ((SRVCJOBORDER.PRINTCOPY) > 0)) ORDER BY SRVCJOBORDER.MECHANIC, SRVCJOBORDER.DATEPROMISED;"
                    'mySQL
                    Qry = "SELECT SRVCJOBORDER.MECHANIC, SRVCJOBORDER.JOBORDERNO, SRVCJOBORDER.DATEPROMISED, SRVCJOBORDER.JOTYPE, SRVCJOBORDER.TOTALCOST From SRVCJOBORDER " & _
                          "Where SRVCJOBORDER.DATEPROMISED >= '" & Format(Dfrom.Value, "yyyy-mm-dd") & "' And SRVCJOBORDER.DATEPROMISED <= '" & Format(Dto.Value, "yyyy-mm-dd") & "' And SRVCJOBORDER.PRINTCOPY > 0 AND SRVCJOBORDER.REMARKS <> 'CANCELLED' ORDER BY SRVCJOBORDER.MECHANIC, SRVCJOBORDER.DATEPROMISED;"
                    
                End If
                
                PRNmc.QryPrn = Qry
                PRNmc.ParamValue1 = "Date Covered: " & Format(Dfrom.Value, "Mmm. dd, yyyy") & IIf(Val(cmdNext.Tag) = 1, "", " to " & Format(Dto.Value, "Mmm. dd, yyyy"))
                PRNmc.ParamValue2 = Val(txtFSCamt.Text)
                PRNmc.MySource = "SRVCMECH"
                PRNmc.Show
            ElseIf mySrc = "DMCM [A/R]" Then
                If Val(cmdNext.Tag) = 1 Then
                
'                    Qry = "SELECT " & _
'                        "ADJDMCM.ADJDATE, " & _
'                        "ADJDMCM.ADJNO, " & _
'                        "ADJDMCM.CUSTOMER, " & _
'                        "ADJDMCM.TRANSTYPE, " & _
'                        "ADJDMCM.REFNO, " & _
'                        "spinvoice.INVDATE, " & _
'                        "ADJDMCM.DEBIT, " & _
'                        "ADJDMCM.CREDIT, " & _
'                        "ADJDMCM.BALANCE, " & _
'                        "ADJDMCM.`STATUS`, " & _
'                        "ADJDMCM.REASON, " & _
'                        "ADJDMCM.PRINTCOPY " & _
'                    "From " & _
'                        "spinvoice " & _
'                        "INNER JOIN ADJDMCM ON (spinvoice.SPINVNO = ADJDMCM.REFNO) " & _
'                    "Where " & _
'                        "ADJDMCM.ADJDATE = '" & Format(dFrom.Value, "yyyy-mm-dd") & "' " & _
'                    "Order By " & _
'                        "ADJDMCM.ADJDATE DESC, " & _
'                        "ADJDMCM.ADJNO Desc"

                    Qry = "SELECT " & _
                              "main.ADJDATE, " & _
                              "main.ADJNO, " & _
                              "main.CUSTOMER, " & _
                              "main.TRANSTYPE, " & _
                              "main.REFNO, " & _
                              "CONVERT( (case main.TRANSTYPE when 'invoice' then inv.INVDATE " & _
                              "when 'job order' then  jo.DATERECEIVED else '' END) USING UTF8) as INVDATE, " & _
                              "main.DEBIT, " & _
                              "main.CREDIT, " & _
                              "main.BALANCE, " & _
                              "main.`STATUS`, " & _
                              "main.REASON, " & _
                              "Main.PRINTCOPY " & _
                            "From " & _
                              "ADJDMCM As Main " & _
                            "Left Join " & _
                                "(SELECT INVDATE,SPINVNO  FROM spinvoice WHERE INVDATE = '" & Format(Dfrom.Value, "yyyy-mm-dd") & "') as inv on main.REFNO = inv.SPINVNO " & _
                            "Left Join " & _
                                "(SELECT DATERECEIVED,JOBORDERNO FROM srvcjoborder WHERE DATERECEIVED = '" & Format(Dfrom.Value, "yyyy-mm-dd") & "') as jo on main.REFNO=jo.JOBORDERNO " & _
                            "Where " & _
                              "Main.ADJDATE= '" & Format(Dfrom.Value, "yyyy-mm-dd") & "' " & _
                            "Order By " & _
                              "main.ADJDATE DESC, " & _
                              "Main.ADJNO Desc"

                Else
'                    Qry = "SELECT " & _
'                        "ADJDMCM.ADJDATE, " & _
'                        "ADJDMCM.ADJNO, " & _
'                        "ADJDMCM.CUSTOMER, " & _
'                        "ADJDMCM.TRANSTYPE, " & _
'                        "ADJDMCM.REFNO, " & _
'                        "spinvoice.INVDATE, " & _
'                        "ADJDMCM.DEBIT, " & _
'                        "ADJDMCM.CREDIT, " & _
'                        "ADJDMCM.BALANCE, " & _
'                        "ADJDMCM.`STATUS`, " & _
'                        "ADJDMCM.REASON, " & _
'                        "ADJDMCM.PRINTCOPY " & _
'                    "From " & _
'                        "spinvoice " & _
'                        "INNER JOIN ADJDMCM ON (spinvoice.SPINVNO = ADJDMCM.REFNO) " & _
'                    "Where " & _
'                        "ADJDMCM.ADJDATE >= '" & Format(dFrom.Value, "yyyy-mm-dd") & "' AND " & _
'                        "ADJDMCM.ADJDATE <= '" & Format(dTo.Value, "yyyy-mm-dd") & "' " & _
'                    "Order By " & _
'                        "ADJDMCM.ADJDATE DESC, " & _
'                        "ADJDMCM.ADJNO Desc"

                        Qry = "SELECT " & _
                              "main.ADJDATE, " & _
                              "main.ADJNO, " & _
                              "main.CUSTOMER, " & _
                              "main.TRANSTYPE, " & _
                              "main.REFNO, " & _
                              "case main.TRANSTYPE when 'invoice' then inv.INVDATE " & _
                              "when 'job order' then  jo.DATERECEIVED else '' END as INVDATE, " & _
                              "main.DEBIT, " & _
                              "main.CREDIT, " & _
                              "main.BALANCE, " & _
                              "main.`STATUS`, " & _
                              "main.REASON, " & _
                              "Main.PRINTCOPY " & _
                            "From " & _
                              "ADJDMCM As Main " & _
                            "Left Join " & _
                                "(SELECT INVDATE,SPINVNO  FROM spinvoice WHERE INVDATE BETWEEN '" & Format(Dfrom.Value, "yyyy-mm-dd") & "' and '" & Format(Dto.Value, "yyyy-mm-dd") & "') as inv on main.REFNO = inv.SPINVNO " & _
                            "Left Join " & _
                                "(SELECT DATERECEIVED,JOBORDERNO FROM srvcjoborder WHERE DATERECEIVED BETWEEN '" & Format(Dfrom.Value, "yyyy-mm-dd") & "' and '" & Format(Dto.Value, "yyyy-mm-dd") & "') as jo on main.REFNO=jo.JOBORDERNO " & _
                            "Where " & _
                              "Main.ADJDATE between '" & Format(Dfrom.Value, "yyyy-mm-dd") & "' and '" & Format(Dto.Value, "yyyy-mm-dd") & "' " & _
                            "Order By " & _
                              "main.ADJDATE DESC, " & _
                              "Main.ADJNO Desc"
                End If
                               
                PRNmc.QryPrn = Qry
                PRNmc.ParamValue1 = "Date Covered: " & Format(Dfrom.Value, "Mmm. dd, yyyy") & IIf(Val(cmdNext.Tag) = 1, "", " to " & Format(Dto.Value, "Mmm. dd, yyyy"))
                pQRY = "SELECT " & _
                            "SETTINGS.`Value` " & _
                        "From " & _
                            "SETTINGS " & _
                        "Where " & _
                            "SETTINGS.DESCRIPTION = 'BRANCH O.I.C' AND " & _
                        "SETTINGS.REMARKS = 'OIC'"
                PRNmc.ParamValue2 = GetFieldValue(pQRY)
                PRNmc.ParamValue3 = GetUserName
                pQRY = "SELECT " & _
                                "SETTINGS.`VALUE` " & _
                        "From " & _
                                "SETTINGS " & _
                        "WHERE " & _
                                "SETTINGS.DESCRIPTION='EMP DEFAULT VALUE' AND " & _
                                "SETTINGS.REMARKS='MANAGER';"
                
                PRNmc.ParamValue4 = GetFieldValue(pQRY)
                PRNmc.MySource = "DMCM [A/R]"
                PRNmc.Show

            ElseIf mySrc = "Sales Order" Then
                If Val(cmdNext.Tag) = 1 Then
                    Qry = "SELECT DATE_FORMAT(spinvoice.INVDATE,'%c/%d/%Y'), spinvoice.SPINVNO, spinvoice.INVTYPE, spinvoice.CHARGETO, Format(spinvoice.AMOUNT,2), spinvoice.REMARKS, spsalesorder.ORDERNO, spsalesorder.FRONTLINER " & _
                        "From spsalesorder INNER JOIN spinvoice ON (spsalesorder.SPSORDERID = spinvoice.SPSORDERID) " & _
                        "Where spinvoice.INVDATE = '" & Format(Dfrom.Value, "yyyy-mm-dd") & "' AND spinvoice.REMARKS <> 'CANCELLED' Order By spsalesorder.FRONTLINER, spinvoice.INVDATE, spinvoice.SPINVNO;"
                    
                Else
                    Qry = "SELECT DATE_FORMAT(spinvoice.INVDATE,'%c/%d/%Y'), spinvoice.SPINVNO, spinvoice.INVTYPE, spinvoice.CHARGETO, Format(spinvoice.AMOUNT,2), spinvoice.REMARKS, spsalesorder.ORDERNO, spsalesorder.FRONTLINER " & _
                        "From spsalesorder INNER JOIN spinvoice ON (spsalesorder.SPSORDERID = spinvoice.SPSORDERID) " & _
                        "Where spinvoice.INVDATE >= '" & Format(Dfrom.Value, "yyyy-mm-dd") & "' AND spinvoice.INVDATE <= '" & Format(Dto.Value, "yyyy-mm-dd") & "' AND spinvoice.REMARKS <> 'CANCELLED' Order By spsalesorder.FRONTLINER, spinvoice.INVDATE, spinvoice.SPINVNO"
                End If
                PRNmc.QryPrn = Qry
                PRNmc.ParamValue1 = SysBranch
                PRNmc.ParamValue2 = SysAddress
                PRNmc.ParamValue3 = "Date Covered: " & Format(Dfrom.Value, "Mmm. dd, yyyy") & IIf(Val(cmdNext.Tag) = 1, "", " to " & Format(Dto.Value, "Mmm. dd, yyyy"))
                PRNmc.MySource = "FRONTLINE SUMMARY"
                PRNmc.Show
'**************** CODE UPDATED: 10/11/2017 ********************
            ElseIf mySrc = "Requisition" Then
                Qry = "call sp_mrcretails.sp_requisition_rpt(" & SysBranchID & ",'" & sysBranchCode & "', '" & Format(Dfrom.Value, "YYYY-mm-dd") & "', '" & Format(Dto.Value, "YYYY-mm-dd") & "'," & UserID & ", " & Val(cmdNext.Tag) & ");"
                
                PRNmc.QryPrn = Qry
                PRNmc.ParamValue1 = Format(Dfrom.Value, "Mmm. dd, yyyy") & IIf(Val(cmdNext.Tag) = 1, "", " to " & Format(Dto.Value, "Mmm. dd, yyyy"))
                PRNmc.MySource = "Requisition"
                PRNmc.Show
            ElseIf mySrc = "Daily Cash Deposit Report" Then
                Qry = "call sp_mrcretails.sp_cashdeposit_summary('" & sysBranchCode & "','" & Format(Dfrom.Value, "yyyy-mm-dd") & "','" & Format(Dto.Value, "yyyy-mm-dd") & "','" & "PREVIEW" & "')"
                frmPrinCDSummary.QryPrn = Qry
                frmPrinCDSummary.ParamValue1 = SysBranch
                frmPrinCDSummary.ParamValue2 = SysAddress
                frmPrinCDSummary.ParamValue3 = Format(Dfrom.Value, "MMMM dd, YYYY") & " To " & Format(Dto.Value, "MMMM dd, YYYY")
                frmPrinCDSummary.ParamValue4 = GetFieldValue("Select VALUE From Settings Where DESCRIPTION = 'ACCOUNTING SECTION HEAD'")
                frmPrinCDSummary.ParamValue5 = GetFieldValue("Select VALUE From Settings Where DESCRIPTION = 'ACCOUNTING DIVISION HEAD'")
                
                frmPrinCDSummary.MySource = "Daily Cash Deposit Report"
                frmPrinCDSummary.Show
                
            ElseIf mySrc = "Inventory Adjustment" Then
                Qry = "call sp_mrcretails.sp_inventoryAdj_summary('" & sysBranchCode & "'," & Val(myDCode) & ",'" & Format(Dfrom.Value, "yyyy-mm-dd") & "','" & Format(Dto.Value, "yyyy-mm-dd") & "')"
                PRNmc.QryPrn = Qry
                PRNmc.ParamValue1 = SysBranch
                PRNmc.ParamValue2 = SysAddress
                PRNmc.ParamValue3 = "DATE COVERED: " & Format(Dfrom.Value, "MMMM dd, yyyy") & IIf(Val(myDCode) = 1, "", " TO " & Format(Dto.Value, "MMMM dd, yyyy"))
                PRNmc.ParamValue4 = GetFieldValue("Select CONCAT(FIRSTNAME, ' ', LASTNAME) as EMPLOYEENAME From Employees Where EMPID = " & Val(UserID) & " ")
                
                PRNmc.MySource = "Inventory Adjustment"
                PRNmc.Show
'**************** CODE UPDATED: 10/11/2017 ********************
            ElseIf mySrc = "Detailed Inventory Outgoing" Then
                Qry = "call sp_mrcretails.sp_acctg_DIO('" & sysBranchCode & "','" & Format(Dfrom.Value, "yyyy-mm-dd") & "','" & Format(Dto.Value, "yyyy-mm-dd") & "')"
                frmPrinCDSummary.QryPrn = Qry
                frmPrinCDSummary.ParamValue1 = SysBranch
                frmPrinCDSummary.ParamValue2 = SysAddress
                frmPrinCDSummary.ParamValue3 = "DATE COVERED: " & Format(Dfrom.Value, "MMMM dd, YYYY") & " To " & Format(Dto.Value, "MMMM dd, YYYY")
                frmPrinCDSummary.ParamValue4 = GetFieldValue("Select CONCAT(FIRSTNAME, ' ', LASTNAME) as EMPLOYEENAME From Employees Where EMPID = " & Val(UserID) & " ")
                frmPrinCDSummary.ParamValue5 = GetFieldValue("Select VALUE From Settings Where DESCRIPTION = 'ACCOUNTING SECTION HEAD'")
                frmPrinCDSummary.ParamValue6 = GetFieldValue("Select VALUE From Settings Where DESCRIPTION = 'ACCOUNTING DIVISION HEAD'")
                
                frmPrinCDSummary.MySource = "Detailed Inventory Outgoing"
                frmPrinCDSummary.Show
            ElseIf mySrc = "Detailed Inventory Incoming" Then
                Qry = "call sp_mrcretails.sp_acctg_DII('" & sysBranchCode & "','" & Format(Dfrom.Value, "yyyy-mm-dd") & "','" & Format(Dto.Value, "yyyy-mm-dd") & "')"
                frmPrinCDSummary.QryPrn = Qry
                frmPrinCDSummary.ParamValue1 = SysBranch
                frmPrinCDSummary.ParamValue2 = SysAddress
                frmPrinCDSummary.ParamValue3 = "DATE COVERED: " & Format(Dfrom.Value, "MMMM dd, YYYY") & " To " & Format(Dto.Value, "MMMM dd, YYYY")
                frmPrinCDSummary.ParamValue4 = GetFieldValue("Select CONCAT(FIRSTNAME, ' ', LASTNAME) as EMPLOYEENAME From Employees Where EMPID = " & Val(UserID) & " ")
                frmPrinCDSummary.ParamValue5 = GetFieldValue("Select VALUE From Settings Where DESCRIPTION = 'ACCOUNTING SECTION HEAD'")
                frmPrinCDSummary.ParamValue6 = GetFieldValue("Select VALUE From Settings Where DESCRIPTION = 'ACCOUNTING DIVISION HEAD'")
                
                frmPrinCDSummary.MySource = "Detailed Inventory Incoming"
                frmPrinCDSummary.Show
            ElseIf mySrc = "Detailed Inventory Incoming" Then
                Qry = "call sp_mrcretails.sp_acctg_DII('" & sysBranchCode & "','" & Format(Dfrom.Value, "yyyy-mm-dd") & "','" & Format(Dto.Value, "yyyy-mm-dd") & "')"
                frmPrinCDSummary.QryPrn = Qry
                frmPrinCDSummary.ParamValue1 = SysBranch
                frmPrinCDSummary.ParamValue2 = SysAddress
                frmPrinCDSummary.ParamValue3 = "DATE COVERED: " & Format(Dfrom.Value, "MMMM dd, YYYY") & " To " & Format(Dto.Value, "MMMM dd, YYYY")
                frmPrinCDSummary.ParamValue4 = GetFieldValue("Select CONCAT(FIRSTNAME, ' ', LASTNAME) as EMPLOYEENAME From Employees Where EMPID = " & Val(UserID) & " ")
                frmPrinCDSummary.ParamValue5 = GetFieldValue("Select VALUE From Settings Where DESCRIPTION = 'ACCOUNTING SECTION HEAD'")
                frmPrinCDSummary.ParamValue6 = GetFieldValue("Select VALUE From Settings Where DESCRIPTION = 'ACCOUNTING DIVISION HEAD'")
                
                frmPrinCDSummary.MySource = "Detailed Inventory Incoming"
                frmPrinCDSummary.Show
            ElseIf mySrc = "Sales Book" Then
                Qry = "call sp_mrcretails.sp_acctg_SalesBook('" & sysBranchCode & "','" & Format(Dfrom.Value, "yyyy-mm-dd") & "','" & Format(Dto.Value, "yyyy-mm-dd") & "')"
                frmPrinCDSummary.QryPrn = Qry
                frmPrinCDSummary.ParamValue1 = SysBranch
                frmPrinCDSummary.ParamValue2 = SysAddress
                frmPrinCDSummary.ParamValue3 = "DATE COVERED: " & Format(Dfrom.Value, "MMMM dd, YYYY") & " To " & Format(Dto.Value, "MMMM dd, YYYY")
                frmPrinCDSummary.ParamValue4 = GetFieldValue("Select CONCAT(FIRSTNAME, ' ', LASTNAME) as EMPLOYEENAME From Employees Where EMPID = " & Val(UserID) & " ")
                frmPrinCDSummary.ParamValue5 = GetFieldValue("Select VALUE From Settings Where DESCRIPTION = 'ACCOUNTING SECTION HEAD'")
                frmPrinCDSummary.ParamValue6 = GetFieldValue("Select VALUE From Settings Where DESCRIPTION = 'ACCOUNTING DIVISION HEAD'")
                
                frmPrinCDSummary.MySource = "Sales Book"
                frmPrinCDSummary.Show
            ElseIf mySrc = "Sales Return Report" Then
                Qry = "call sp_mrcretails.sp_acctg_SalesReturn('" & sysBranchCode & "','" & Format(Dfrom.Value, "yyyy-mm-dd") & "','" & Format(Dto.Value, "yyyy-mm-dd") & "')"
                frmPrinCDSummary.QryPrn = Qry
                frmPrinCDSummary.ParamValue1 = SysBranch
                frmPrinCDSummary.ParamValue2 = SysAddress
                frmPrinCDSummary.ParamValue3 = "DATE COVERED: " & Format(Dfrom.Value, "MMMM dd, YYYY") & " To " & Format(Dto.Value, "MMMM dd, YYYY")
                frmPrinCDSummary.ParamValue4 = GetFieldValue("Select CONCAT(FIRSTNAME, ' ', LASTNAME) as EMPLOYEENAME From Employees Where EMPID = " & Val(UserID) & " ")
                frmPrinCDSummary.ParamValue5 = GetFieldValue("Select VALUE From Settings Where DESCRIPTION = 'ACCOUNTING SECTION HEAD'")
                frmPrinCDSummary.ParamValue6 = GetFieldValue("Select VALUE From Settings Where DESCRIPTION = 'ACCOUNTING DIVISION HEAD'")
                
                frmPrinCDSummary.MySource = "Sales Return Report"
                frmPrinCDSummary.Show
            ElseIf mySrc = "Service Receipts Book" Then
                Qry = "call sp_mrcretails.sp_acctg_JobOrder('" & sysBranchCode & "','" & Format(Dfrom.Value, "yyyy-mm-dd") & "','" & Format(Dto.Value, "yyyy-mm-dd") & "')"
                frmPrinCDSummary.QryPrn = Qry
                frmPrinCDSummary.ParamValue1 = SysBranch
                frmPrinCDSummary.ParamValue2 = SysAddress
                frmPrinCDSummary.ParamValue3 = "DATE COVERED: " & Format(Dfrom.Value, "MMMM dd, YYYY") & " To " & Format(Dto.Value, "MMMM dd, YYYY")
                frmPrinCDSummary.ParamValue4 = GetFieldValue("Select CONCAT(FIRSTNAME, ' ', LASTNAME) as EMPLOYEENAME From Employees Where EMPID = " & Val(UserID) & " ")
                frmPrinCDSummary.ParamValue5 = GetFieldValue("Select VALUE From Settings Where DESCRIPTION = 'ACCOUNTING SECTION HEAD'")
                frmPrinCDSummary.ParamValue6 = GetFieldValue("Select VALUE From Settings Where DESCRIPTION = 'ACCOUNTING DIVISION HEAD'")
                
                frmPrinCDSummary.MySource = "Service Receipts Book"
                frmPrinCDSummary.Show
            Else
                MsgBox "Not yet available", vbInformation + vbOKOnly, "Print"
            End If
        Case 3  'Exit
            Unload Me
            'MotoRep.tbMain.Visible = True
            'MotoRep.imgBG.Visible = TruetoolBarButton
'**************** CODE UPDATED: 08/15/2017 *******************
            If mySrc = "SETTLEMENT DISCREPANCY" Then
                MotoRep.toolBarButton.Visible = False
'            ElseIf mySrc = "Branch Order" Then
'                MotoRep.toolBarButton.Visible = False
            Else
                MotoRep.toolBarButton.Visible = True
            End If
'************ END CODE UPDATED: 08/15/2017 *******************
    End Select
        
End Sub

Private Sub dFrom_CloseUp()
    FrDate = Dfrom.Value
    If cmdNext.Tag = "1" Then
        Call RefreshRecord(1)
    End If
End Sub

Private Sub dTo_CloseUp()
    ToDate = Dto.Value
    If Dto.Value < Dfrom.Value Then
        MsgBox "Please choose a date not less than " & Dfrom.Value, vbExclamation + vbOKOnly, "Validation Error"
        Dto.Value = Dfrom.Value
    End If
End Sub

Private Sub flxSum_EnterCell()
    ActRow = flxSum.Row
End Sub
Private Sub flxSum_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If flxSum.TextMatrix(ActRow, 1) = "" Or mySrc = "SP Incoming" Or mySrc = "SP Incoming Supplier" Or mySrc = "Received Spare Parts" Or mySrc = "Received Spare Parts DVOWH" Or mySrc = "Sales Order" Or mySrc = "Mechanic Income" Or mySrc = "Top Selling Spare Parts" Or mySrc = "Cashier Cash Flow" Then Exit Sub
    If Not IsDate(flxSum.TextMatrix(ActRow, 1)) Then
        Exit Sub
    Else
        If mySrc <> "Branch Order" Then
            If mySrc <> "Manager's Cash Flow" Then
                If DateValue(flxSum.TextMatrix(ActRow, 1)) <> Date Then Exit Sub
            End If
        End If
    End If
    If Button = vbRightButton Then
        If mySrc = "Branch Order" Then
            'mnuCancel.Caption = "Monitor Purchase Request"
            Call PopupMenu(mnuPOmonitor)
        ElseIf mySrc = "Manager's Cash Flow" Then
            Call PopupMenu(mnuDepositValidation)
        Else
            mnuCancel.Caption = "Cancel " & mySrc
            Call PopupMenu(mnuSum)
        End If
    End If
End Sub

Private Sub Form_Load()
    Dim Qry As String
    
    lblHeading(0).ToolTipText = "New"
    Dfrom.Value = Date
    cmdNext.Tag = "1"
    ShowDateRange = False
    'e-set ang table based sa Mysource
    If mySrc = "Job Order" Then
        TCol = 14: ColPrn = 12: AccID = 43
        cboRebuilt.AddItem "View All"
        cboRebuilt.AddItem "View Machine Shop Only"
        cboRebuilt.AddItem "Exclude Machine Shop"
        cboRebuilt.Visible = True
        dCode = 1
    ElseIf mySrc = "SP Incoming" Or mySrc = "SP Incoming Supplier" Then
        TCol = 9: ColPrn = 8: AccID = 52
        Call cmdNext_Click
        Dfrom.Value = DateSerial(Year(Date), Month(Date), 1)
        Dto.Value = Date
        ShowDateRange = True
        'cmdChkClosed.Visible = True
        'Call ShowCommands(1)
    
    ElseIf mySrc = "Clerical Adjustment" Then
        TCol = 11: ColPrn = 11: AccID = 56
        Dfrom.Value = DateSerial(Year(Date), Month(Date), 1)
        Dto = Date
        ShowDateRange = True
    ElseIf mySrc = "DMCM [A/R]" Then
        TCol = 11: ColPrn = 11: AccID = 54
        Dfrom.Value = DateSerial(Year(Date), Month(Date), 1)
        Dto = Date
        ShowDateRange = True
    ElseIf mySrc = "Inventory Adjustment" Then
        TCol = 8: ColPrn = 8: AccID = 55
        Dfrom.Value = DateSerial(Year(Date), Month(Date), 1)
        Dto = Date
        ShowDateRange = True
    ElseIf mySrc = "Customer P.O." Then
        TCol = 8: ColPrn = 0: AccID = 40
        Dfrom.Value = DateSerial(Year(Date), Month(Date), 1)
        Dto = Date
        ShowDateRange = True
    ElseIf mySrc = "Top Selling Spare Parts" Then
        TCol = 6: AccID = 97
    ElseIf mySrc = "Purchase Return" Then
        TCol = 9: ColPrn = 0: AccID = 24
        Call cmdNext_Click
        If Day(Date) <= 7 Then
            Dfrom.Value = DateAdd("m", -1, DateSerial(Year(Date), Month(Date), 1))
        Else
            Dfrom.Value = DateSerial(Year(Date), Month(Date), 1)
        End If
        ShowDateRange = True
        
    ElseIf mySrc = "Requisition" Then
        TCol = 9: ColPrn = 0: AccID = 39
    ElseIf mySrc = "Sales Return" Then
        TCol = 9: ColPrn = 0: AccID = 29
    ElseIf mySrc = "Location Adjustment" Then
        TCol = 8: ColPrn = 0: AccID = 35
        Call cmdNext_Click
        If Day(Date) <= 7 Then
            Dfrom.Value = DateAdd("m", -1, DateSerial(Year(Date), Month(Date), 1))
        Else
            Dfrom.Value = DateSerial(Year(Date), Month(Date), 1)
        End If
        ShowDateRange = True
    ElseIf mySrc = "Received Spare Parts" Or mySrc = "Received Spare Parts DVOWH" Then
        TCol = 11: ColPrn = 0: AccID = 23
        Call cmdNext_Click
        If Day(Date) <= 7 Then
            Dfrom.Value = DateAdd("m", -1, DateSerial(Year(Date), Month(Date), 1))
        Else
            Dfrom.Value = DateSerial(Year(Date), Month(Date), 1)
        End If
        ShowDateRange = True
    ElseIf mySrc = "Sales Order" Then
        TCol = 7: ColPrn = 0: AccID = 28
    ElseIf mySrc = "Mechanic Income" Then
        TCol = 8: ColPrn = 0: AccID = 42
        '***Access
        'Qry = "SELECT SETTINGS.VALUE From SETTINGS WHERE (((SETTINGS.DESCRIPTION)='FSC AMOUNT'));"
        'mySQl
        Qry = "SELECT SETTINGS.VALUE From SETTINGS WHERE SETTINGS.DESCRIPTION='FSC AMOUNT';"
        
        txtFSCamt.Text = GetFieldValue(Qry)
    ElseIf mySrc = "Cash Voucher" Then
        TCol = 9: ColPrn = 8: AccID = 58
        Call cmdNext_Click
        Dfrom.Value = DateSerial(Year(Date), Month(Date), 1)
        Dto = Date
        ShowDateRange = True
    ElseIf mySrc = "Cashier Cash Flow" Then
        TCol = 6: ColPrn = 0: AccID = 48
        Call cmdNext_Click
        Dfrom.Value = DateSerial(Year(Date), Month(Date), 1)
        Dto.Value = Date
        ShowDateRange = True
    ElseIf mySrc = "Manager's Cash Flow" Then
        TCol = 7: ColPrn = 0: AccID = 48
        Call cmdNext_Click
        Dfrom.Value = DateSerial(Year(Date), Month(Date), 1)
        Dto.Value = Date
        ShowDateRange = True
        
    ElseIf mySrc = "Advances for Liquidation" Then
        TCol = 9: ColPrn = 0: AccID = 65
        Dfrom = Date
        Dto.Value = Date
        cmdNext.Tag = "1"
        'ShowDateRange = False
    ElseIf mySrc = "Cash Return" Then
        TCol = 7: ColPrn = 0: AccID = 64
        Call cmdNext_Click
        Dfrom.Value = DateSerial(Year(Date), Month(Date), 1)
        Dto.Value = Date
        ShowDateRange = True
    ElseIf mySrc = "Branch Purchase Order" Then
        TCol = 12: ColPrn = 0
        'AccID =0
        Dfrom = Date
        Dto.Value = Date
    ElseIf mySrc = "Trust Receipt Agreement" Then
        TCol = 12: ColPrn = 11: AccID = 64
        Call cmdNext_Click
        Dfrom.Value = DateSerial(Year(Date), Month(Date), 1)
        Dto.Value = Date
        If InStr(SysBranch, "LYR MOTORPARTS CORP. - VAN") > 0 Then
            cmdExportTRA.Visible = True
        End If
        ShowDateRange = True
    ElseIf mySrc = "Branch Order" Then
        TCol = 8: ColPrn = 7: AccID = 53
        Call cmdNext_Click
        Dfrom.Value = DateSerial(Year(Date), Month(Date), 1)
        Dto.Value = Date
        ShowDateRange = True
        'Call ShowCommands(1)
'*********************** CODE UPDATED: 08/07/2017 **********************
    ElseIf mySrc = "SETTLEMENT DISCREPANCY" Then
        TCol = 12: ColPrn = 11
        Call cmdNext_Click
        Dfrom.Value = DateSerial(Year(Date), Month(Date), 1)
        Dto.Value = Date
        ShowDateRange = True
        'Call ShowCommands(1)
    ElseIf mySrc = "SDiscrepancy Adjustment" Then
        TCol = 8: ColPrn = 8: AccID = 55
        Call cmdNext_Click
        Dfrom.Value = DateSerial(Year(Date), Month(Date), 1)
        Dto.Value = Date
        ShowDateRange = True
        'Call ShowCommands(1)
    ElseIf mySrc = "Daily Cash Deposit Report" Then
        TCol = 31: ColPrn = 31
        Call cmdNext_Click
        Dfrom.Value = DateSerial(Year(Date), Month(Date), 1)
        Dto.Value = Date
        cmdSaveExcel.Visible = True
        cmdPrev.Visible = False
        cmdTrans(2).Left = cmdTrans(1).Left
        cmdTrans(1).Visible = False
        lblTO.Visible = True
        ShowDateRange = True
        flxSum.SelectionMode = flexSelectionFree
    ElseIf mySrc = "Detailed Inventory Outgoing" Then
        TCol = 14: ColPrn = 14
        Call cmdNext_Click
        Dfrom.Value = DateSerial(Year(Date), Month(Date), 1)
        Dto.Value = Date
        cmdSaveExcel.Visible = True
        cmdPrev.Visible = False
        cmdTrans(2).Left = cmdTrans(1).Left
        cmdTrans(1).Visible = False
        lblTO.Visible = True
        ShowDateRange = True
    ElseIf mySrc = "Detailed Inventory Incoming" Then
        TCol = 16: ColPrn = 16
        Call cmdNext_Click
        Dfrom.Value = DateSerial(Year(Date), Month(Date), 1)
        Dto.Value = Date
        cmdSaveExcel.Visible = True
        cmdPrev.Visible = False
        cmdTrans(2).Left = cmdTrans(1).Left
        cmdTrans(1).Visible = False
        lblTO.Visible = True
        ShowDateRange = True
    ElseIf mySrc = "Sales Book" Then
        TCol = 17: ColPrn = 17
        Call cmdNext_Click
        Dfrom.Value = DateSerial(Year(Date), Month(Date), 1)
        Dto.Value = Date
        cmdSaveExcel.Visible = True
        cmdPrev.Visible = False
        cmdTrans(2).Left = cmdTrans(1).Left
        cmdTrans(1).Visible = False
        lblTO.Visible = True
        ShowDateRange = True
    ElseIf mySrc = "Sales Return Report" Then
        TCol = 23: ColPrn = 23
        Call cmdNext_Click
        Dfrom.Value = DateSerial(Year(Date), Month(Date), 1)
        Dto.Value = Date
        cmdSaveExcel.Visible = True
        cmdPrev.Visible = False
        cmdTrans(2).Left = cmdTrans(1).Left
        cmdTrans(1).Visible = False
        lblTO.Visible = True
        ShowDateRange = True
    ElseIf mySrc = "Service Receipts Book" Then
        TCol = 14: ColPrn = 14
        Call cmdNext_Click
        Dfrom.Value = DateSerial(Year(Date), Month(Date), 1)
        Dto.Value = Date
        cmdSaveExcel.Visible = True
        cmdPrev.Visible = False
        cmdTrans(2).Left = cmdTrans(1).Left
        cmdTrans(1).Visible = False
        lblTO.Visible = True
        ShowDateRange = True
'******************* END CODE UPDATED: 08/07/2017 **********************
    ElseIf mySrc = "ADJUSTMENT REQUEST" Then
        TCol = 12: ColPrn = 12
        Dfrom.Value = DateSerial(Year(Date), Month(Date), 1)
        Dto = Date
        ShowDateRange = True
    ElseIf mySrc = "CANCELLATION MODULE" Then
        TCol = 12: ColPrn = 12
        Dfrom.Value = DateSerial(Year(Date), Month(Date), 1)
        Dto = Date
        ShowDateRange = True
    End If
    
    If FrDate <> DateValue("1/1/1977") Then
        Dfrom.Value = FrDate
        Dto.Value = ToDate
    Else
        FrDate = Dfrom.Value
        ToDate = Dto.Value
    End If
    'If mySrc = "Purchase Return" Or mySrc = "Location Adjustment" Or mySrc = "Cash Voucher" Or mySrc = "Received Spare Parts" Or mySrc = "Cashier Cash Flow" Or mySrc = "Cash Return" Or mySrc = "Manager's Cash Flow" Or mySrc = "Trust Receipt Agreement" Then
    If ShowDateRange Then
        Call RefreshRecord(2)
    ElseIf mySrc <> "Detailed Inventory Outgoing" Or mySrc <> "Detailed Inventory Incoming" Or mySrc <> "Sales Book" Or mySrc <> "Sales Return Report" Or mySrc <> "Service Receipts Book" Then
        Call RefreshRecord(1)
    End If
    ActRow = 0
    MotoRep.sBar.Panels(6).Text = mySrc & " Summary"
    lblHeading(2).Caption = mySrc & " Summary"
    lblHeading(3).Caption = mySrc & " Summary"
    Me.Caption = mySrc & " Summary"
    'MotoRep.tbMain.Visible = False
    MotoRep.imgBG.Visible = False
    MotoRep.toolBarButton.Visible = False
    ActRow = 1
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    MotoRep.sBar.Panels(6).Text = ""
    If rsSummary.State = adStateOpen Then rsSummary.Close
End Sub

Private Sub mnuCancel_Click()
    Dim Qry As String, MyAcc As String
    
    '***Access
    'Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE (((USERACCESS.EMPID)=" & UserID & ") AND ((USERACCESS.MENUINDX)=" & AccID & "));"
    'mySQl
    Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE USERACCESS.EMPID=" & UserID & " AND USERACCESS.MENUINDX=" & AccID & ";"
    
    MyAcc = GetFieldValue(Qry)
    If MyAcc <> "FULL ACCESS" Then
        MsgBox "You don't have full access to this transaction." & vbCrLf & "Please contact your System Administrator", vbExclamation + vbOKOnly, "Access Denied"
        Exit Sub
    End If
    MsgBox "Not applicable to this Transaction", vbInformation + vbOKOnly, "Validation"
'    If mySrc = "Job Order" Then
'        Call CancelJO
'    ElseIf mySrc = "Clerical Adjustment" Then
'        Call CancelClericalAdj
'    ElseIf mySrc = "DMCM [A/R]" Then
'        Call CancelDMCM
'    ElseIf mySrc = "Inventory Adjustment" Then
'        Call CancelInvAdj
'    ElseIf mySrc = "Customer P.O." Then
'        Call CancelCustPO
'    ElseIf mySrc = "Purchase Return" Then
'        Call CancelPurchaseReturn
'    ElseIf mySrc = "Requisition" Then
'        Call CancelRequisition
'    ElseIf mySrc = "Sales Return" Then
'        Call CancelSalesReturn
'    ElseIf mySrc = "Location Adjustment" Then
'        Call CancelLocAdj
'    ElseIf mySrc = "Cash Voucher" Then
'        Call CancelCashVoucher
'    End If
End Sub

Private Sub CancelCashVoucher()
    Dim Ans As Integer, MyReason As String, Qry As String
    Dim myAcctDesc As String * 50
    
    Ans = MsgBox("Do you want to Cancel Voucher?" & vbCrLf & vbCrLf & _
        "C.V. Date : " & flxSum.TextMatrix(ActRow, 1) & vbCrLf & _
        "Voucher No: " & flxSum.TextMatrix(ActRow, 2) & vbCrLf & _
        "Received By: " & flxSum.TextMatrix(ActRow, 3) & vbCrLf & _
        "Particular : " & flxSum.TextMatrix(ActRow, 5), vbQuestion + vbYesNo, "Cancel")
    If Ans = vbNo Then Exit Sub
    MyReason = InputBox("Please indicate reason for Cancellation of Voucher:", "Reason for Cancel")
    If MyReason = "" Then
        MsgBox "Unable to indicate Reason..." & vbCrLf & "Cancellation of Voucher failed!", vbExclamation + vbOKOnly, "Cancel Failed"
        Exit Sub
    End If
    MyReason = Replace(MyReason, "'", "`")
    'Update Voucher
    '***Access
    'LYRcon.Execute "UPDATE CASHVOUCHER SET CASHVOUCHER.REMARKS = 'CANCELLED' WHERE (((CASHVOUCHER.VOUCHERNO)=" & flxSum.TextMatrix(ActRow, 2) & "));"
    'mySQL
    LYRcon.Execute "UPDATE CASHVOUCHER SET CASHVOUCHER.REMARKS = 'CANCELLED' WHERE CASHVOUCHER.VOUCHERNO=" & flxSum.TextMatrix(ActRow, 2) & ";"
    
    'Add to Disbursement
    myAcctDesc = "CASH VOUCHER [" & flxSum.TextMatrix(ActRow, 5) & "]"
    '***Access
    'LYRcon.Execute "INSERT INTO DISBURSEMENT ( CVDATE, CVNUMBER, DISBURSETYPE, ACCTDESC, PAYEE, PARTICULARS, AMOUNT, STATUS, PRINTCOPY, UPDATEDBY, LASTUPDATE ) " & _
                   "VALUES (#" & flxSum.TextMatrix(ActRow, 1) & "#, " & flxSum.TextMatrix(ActRow, 2) & ", 'Cancelled Vouchers', '" & myAcctDesc & "', '" & flxSum.TextMatrix(ActRow, 3) & "', '" & MyReason & "', 0, 'Final', 1, " & UserID & ", #" & Now & "#);"
    'mySQL
    LYRcon.Execute "INSERT INTO DISBURSEMENT ( CVDATE, CVNUMBER, DISBURSETYPE, ACCTDESC, PAYEE, PARTICULARS, AMOUNT, STATUS, PRINTCOPY, UPDATEDBY, LASTUPDATE ) " & _
                   "VALUES ('" & Format(flxSum.TextMatrix(ActRow, 1), "yyyy-mm-dd") & "', " & flxSum.TextMatrix(ActRow, 2) & ", 'Cancelled Vouchers', '" & myAcctDesc & "', '" & flxSum.TextMatrix(ActRow, 3) & "', '" & MyReason & "', 0, 'Final', 1, " & UserID & ", '" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "');"
    
    'Update Auditrans
    '***Access
    'LYRcon.Execute "INSERT INTO AUDITRANS ( TRANSDATE, TRANSTYPE, TRANSDOC, REFNO, REFDATE, REASON, EMPID ) " & _
                    "VALUES (#" & Date & "#, 'CANCEL', 'CASH VOUCHER', '" & flxSum.TextMatrix(ActRow, 2) & "', #" & flxSum.TextMatrix(ActRow, 1) & "#, '" & Left$(MyReason, 200) & "', " & UserID & ");"
    'mySQL
    LYRcon.Execute "INSERT INTO AUDITRANS ( TRANSDATE, TRANSTYPE, TRANSDOC, REFNO, REFDATE, REASON, EMPID ) " & _
                    "VALUES ('" & Format(Date, "yyyy-mm-dd") & "', 'CANCEL', 'CASH VOUCHER', '" & flxSum.TextMatrix(ActRow, 2) & "', '" & Format(flxSum.TextMatrix(ActRow, 1), "yyyy-mm-dd") & "', '" & Left(MyReason, 200) & "', " & UserID & ");"
    
    flxSum.TextMatrix(ActRow, 6) = "0.00"
    flxSum.TextMatrix(ActRow, 7) = "CANCELLED"
    MsgBox "Voucher Cancelled!", vbInformation + vbOKOnly

End Sub

Private Sub CancelLocAdj()
    Dim Ans As Integer, MyReason As String, Qry As String
    Dim rsPRV As New ADODB.Recordset
    
    
    Ans = MsgBox("Do you want to Cancel Adjustment with Control No. " & flxSum.TextMatrix(ActRow, 2), vbQuestion + vbYesNo, "Cancel")
    If Ans = vbNo Then Exit Sub
    MyReason = InputBox("Please indicate reason for Cancellation of Adjustment:", "Reason for Cancel")
    If MyReason = "" Then
        MsgBox "Unable to indicate Reason..." & vbCrLf & "Cancellation of Job Order failed!", vbExclamation + vbOKOnly, "Cancel Failed"
        Exit Sub
    End If
    'Update SP Location Adjustment
    '***Access
    'LYRcon.Execute "UPDATE TRANSFERLOC SET TRANSFERLOC.REFERENCE = 'CANCELLED' WHERE (((TRANSFERLOC.TRANSLOCID)=" & flxSum.TextMatrix(ActRow, 8) & "));"
    'Qry = "SELECT TRANSLOCDTLS.SPAREPARTID, TRANSLOCDTLS.PREVLOC From TRANSLOCDTLS WHERE (((TRANSLOCDTLS.TRANSLOCID)=" & flxSum.TextMatrix(ActRow, 8) & "));"
    
    'mySQL
    LYRcon.Execute "UPDATE TRANSFERLOC SET TRANSFERLOC.REFERENCE = 'CANCELLED' WHERE TRANSFERLOC.TRANSLOCID=" & flxSum.TextMatrix(ActRow, 8) & ";"
    Qry = "SELECT TRANSLOCDTLS.SPAREPARTID, TRANSLOCDTLS.PREVLOC From TRANSLOCDTLS WHERE TRANSLOCDTLS.TRANSLOCID=" & flxSum.TextMatrix(ActRow, 8) & ";"
    
    With rsPRV
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            FlxPrev.Rows = 1
            .MoveFirst
            Do While Not .EOF
                FlxPrev.AddItem !SparepartID & vbTab & !PrevLoc
                .MoveNext
            Loop
        End If
        .Close
    End With
    With FlxPrev
        For i = 1 To .Rows - 1
            '***Access
            'LYRcon.Execute "UPDATE SPAREPARTS SET SPAREPARTS.LOCATION = '" & .TextMatrix(i, 1) & "' WHERE (((SPAREPARTS.SPAREPARTID)=" & .TextMatrix(i, 0) & "));"
            'mySQL
            LYRcon.Execute "UPDATE SPAREPARTS SET SPAREPARTS.LOCATION = '" & .TextMatrix(i, 1) & "' WHERE SPAREPARTS.SPAREPARTID=" & .TextMatrix(i, 0) & ";"
        Next i
    End With
    'Update Audit Trans
    '***Access
    'LYRcon.Execute "INSERT INTO AUDITRANS ( TRANSDATE, TRANSTYPE, TRANSDOC, REFNO, REFDATE, REASON, EMPID ) " & _
            "VALUES (#" & Date & "#, 'CANCEL', 'SP LOCATION ADJ.', '" & flxSum.TextMatrix(ActRow, 2) & "', #" & flxSum.TextMatrix(ActRow, 1) & "#, '" & MyReason & "', " & UserID & ");"
    'mySQL
    LYRcon.Execute "INSERT INTO AUDITRANS ( TRANSDATE, TRANSTYPE, TRANSDOC, REFNO, REFDATE, REASON, EMPID ) " & _
            "VALUES ('" & Format(Date, "yyyy-mm-dd") & "', 'CANCEL', 'SP LOCATION ADJ.', '" & flxSum.TextMatrix(ActRow, 2) & "', '" & Format(flxSum.TextMatrix(ActRow, 1), "yyyy-mm-dd") & "', '" & MyReason & "', " & UserID & ");"
  
   'Update Grid
    flxSum.TextMatrix(ActRow, 3) = "CANCELLED"
    With flxSum     'Change row color
        .Col = 1: .Row = ActRow: .ColSel = 8: .RowSel = ActRow: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 0, 0)
        .Col = 2
    End With
    MsgBox "Spare Parts Location Adjustment Cancelled!", vbInformation + vbOKOnly

End Sub
Private Sub CancelSalesReturn()
    Dim rsCDtls As New ADODB.Recordset
    Dim rsLedger As New ADODB.Recordset
    Dim Ans As Integer, MyReason As String, Qry As String, MyDetails As String, MyAcc As String
    Dim MyUnit As String, MyDesc As String, MyInv As String, MyProc As String
    Dim i As Integer, CustID As Integer, InvBal As Single, TtlAmtPd As Currency, MyCredit As Currency, MyDebit As Currency
    
    'Display Details in flxCancelSalRet
    '***Access
    'Qry = "SELECT SPITR.RECEIVE, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPITR.PRICE, [RECEIVE]*[PRICE] AS AMOUNT, SPITR.SPAREPARTID " & _
          "FROM SPITR INNER JOIN SPAREPARTS ON SPITR.SPAREPARTID = SPAREPARTS.SPAREPARTID " & _
          "WHERE (((SPITR.REFERENCE)='RET#" & flxSum.TextMatrix(ActRow, 2) & "') AND ((SPITR.TRANSTYPE)='SALES RETURN'));"
    'mySQL
    Qry = "SELECT SPITR.RECEIVE, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPITR.PRICE, SPITR.RECEIVE*SPITR.PRICE AS AMOUNT, SPITR.SPAREPARTID " & _
          "FROM SPITR INNER JOIN SPAREPARTS ON SPITR.SPAREPARTID = SPAREPARTS.SPAREPARTID " & _
          "WHERE SPITR.REFERENCE='RET#" & flxSum.TextMatrix(ActRow, 2) & "' AND SPITR.TRANSTYPE='SALES RETURN';"
    
    With rsCDtls
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        flxCancelSalRet.Rows = 1
        If .RecordCount > 0 Then
            .MoveFirst
            MyDetails = ""
            Do While Not .EOF
                '0-QTY,  1-PARTNO,  2-DESCRIPTION, 3-PRICE, 4-AMOUNT, 5-SPID
                flxCancelSalRet.AddItem !Receive & vbTab & !PartNumber & vbTab & !Description & vbTab & !Price & vbTab & !Amount & vbTab & !SparepartID
                If InStr(UCase$(!Description), "4T OIL") > 0 Then
                    MyUnit = IIf(!Receive > 1, "QRTS", "QRT")
                Else
                    MyUnit = IIf(!Receive > 1, "PCS", "PC")
                End If
                If Len(!Description) < 30 Then
                    MyDesc = !Description & Space$(30 - Len(!Description))
                Else
                    MyDesc = Left$(!Description, 30)
                End If
                
                MyDetails = IIf(MyDetails <> "", MyDetails & vbCrLf, "") & !Receive & vbTab & MyUnit & vbTab & !PartNumber & Space$(20 - Len(!PartNumber)) & vbTab & MyDesc & vbTab & Format(!Price, "#,##0.00") & vbTab & Format(!Amount, "#,##0.00")
                'If MyDetails <> "" Then
                '    MyDetails = MyDetails & vbCrLf & !RECEIVE & vbTab & MyUnit & vbTab & !PartNumber & Space$(20 - Len(!PartNumber)) & vbTab & MyDesc & vbTab & Format(!Price, "#,##0.00") & vbTab & Format(!Amount, "#,##0.00")
                'Else
                '    MyDetails = "" '& !RECEIVE & vbTab & MyUnit & vbTab & !PartNumber & Space$(20 - Len(!PartNumber)) & vbTab & MyDesc & vbTab & Format(!Price, "#,##0.00") & vbTab & Format(!Amount, "#,##0.00")
         
                'End If
                .MoveNext
            Loop
        End If
        .Close
    End With
    Ans = MsgBox("Do you want to Cancel Sales Return?" & vbCrLf & vbCrLf & _
        "Sales Return No.: " & flxSum.TextMatrix(ActRow, 2) & vbTab & vbTab & vbTab & vbTab & "Sales Return Date : " & flxSum.TextMatrix(ActRow, 1) & vbCrLf & vbCrLf & _
        "QTY." & vbTab & "UNIT" & vbTab & "PART NUMBER" & vbTab & vbTab & "DESCRIPTION" & vbTab & vbTab & vbTab & "PRICE" & vbTab & "AMOUNT" & vbCrLf & MyDetails, vbYesNo, "Cancel")
    If Ans = vbNo Then Exit Sub
    MyReason = InputBox("Please indicate reason for Cancellation of Sales Return:", "Reason for Cancel")
    If MyReason = "" Then
        MsgBox "Unable to indicate Reason..." & vbCrLf & "Cancellation of Sales Return failed!", vbExclamation + vbOKOnly, "Cancel Failed"
        Exit Sub
    End If
    'Retrace transaction made
    'Update Sales Return Remarks to CANCELLED
    '***Access
    'LYRcon.Execute "UPDATE SRETURNSP SET SRETURNSP.REMARKS = 'CANCELLED', SRETURNSP.TOTALAMT = 0 WHERE (((SRETURNSP.RETNO)='" & flxSum.TextMatrix(ActRow, 2) & "'));"
    'mySQL
    LYRcon.Execute "UPDATE SRETURNSP SET SRETURNSP.REMARKS = 'CANCELLED', SRETURNSP.TOTALAMT = 0 WHERE SRETURNSP.RETNO='" & flxSum.TextMatrix(ActRow, 2) & "';"
    
    With flxCancelSalRet
        For i = 1 To .Rows - 1
            'Cancel SPTIR
            '***Access
            'LYRcon.Execute "UPDATE SPITR SET SPITR.TRANSTYPE = 'CANCELLED', SPITR.RECEIVE = 0 " & _
                           "WHERE (((SPITR.REFERENCE)='RET#" & flxSum.TextMatrix(ActRow, 2) & "') AND ((SPITR.SPAREPARTID)=" & .TextMatrix(i, 5) & "));"
            'mySQL
            LYRcon.Execute "UPDATE SPITR SET SPITR.TRANSTYPE = 'CANCELLED', SPITR.RECEIVE = 0 " & _
                           "WHERE SPITR.REFERENCE='RET#" & flxSum.TextMatrix(ActRow, 2) & "' AND SPITR.SPAREPARTID=" & .TextMatrix(i, 5) & ";"
            'Update Qty-On Hand
            '***Access
            'LYRcon.Execute "UPDATE SPAREPARTS SET SPAREPARTS.QTYONHAND = [SPAREPARTS].[QTYONHAND]-" & Val(.TextMatrix(i, 0)) & _
                           " WHERE (((SPAREPARTS.SPAREPARTID)=" & .TextMatrix(i, 5) & "));"
            'mySQl
            LYRcon.Execute "UPDATE SPAREPARTS SET SPAREPARTS.QTYONHAND = SPAREPARTS.QTYONHAND-" & Val(.TextMatrix(i, 0)) & _
                           " WHERE SPAREPARTS.SPAREPARTID=" & .TextMatrix(i, 5) & ";"
        
        Next i
    End With
    
    '***Access
    'Qry = "SELECT SPINVOICE.SPINVNO From SPINVOICE WHERE (((SPINVOICE.SPINVNO)='" & flxSum.TextMatrix(ActRow, 3) & "'));"
    'mysql
    Qry = "SELECT SPINVOICE.SPINVNO From SPINVOICE WHERE SPINVOICE.SPINVNO='" & flxSum.TextMatrix(ActRow, 3) & "';"
    
    If CtrlNumExist(Qry) Then   'Update Source Invoice field remarks to Null
        'LYRcon.Execute "UPDATE SPINVOICE SET SPINVOICE.REMARKS = Null WHERE (((SPINVOICE.SPINVNO)='" & flxSum.TextMatrix(ActRow, 3) & "'));"
        'mySQL
        LYRcon.Execute "UPDATE SPINVOICE SET SPINVOICE.REMARKS = Null WHERE SPINVOICE.SPINVNO='" & flxSum.TextMatrix(ActRow, 3) & "';"
        
    End If
    '***Access
    'Qry = "SELECT PAYMENTS.INVOICENO From PAYMENTS WHERE (((PAYMENTS.INVOICENO)='" & flxSum.TextMatrix(ActRow, 3) & "'));"
    'mySQL
    Qry = "SELECT PAYMENTS.INVOICENO From PAYMENTS WHERE PAYMENTS.INVOICENO='" & flxSum.TextMatrix(ActRow, 3) & "';"
    
    If CtrlNumExist(Qry) Then   'If Cash Update Payment's remarks to null
        '***Access
        'LYRcon.Execute "UPDATE PAYMENTS SET PAYMENTS.REMARKS = Null WHERE (((PAYMENTS.INVOICENO)='" & flxSum.TextMatrix(ActRow, 3) & "'));"
        'mySQL
        LYRcon.Execute "UPDATE PAYMENTS SET PAYMENTS.REMARKS = Null WHERE PAYMENTS.INVOICENO='" & flxSum.TextMatrix(ActRow, 3) & "';"
        
    End If
    'Update Ledger and On-Account Tables if Credit
    CustID = GetFieldValue("SELECT SPINVOICE.CUSTOMERID, SPINVOICE.SPINVNO FROM SPINVOICE WHERE (((SPINVOICE.SPINVNO)='" & flxSum.TextMatrix(ActRow, 3) & "'));")
    '***Access
    'Qry = "SELECT LEDGER.* From LEDGER WHERE (((LEDGER.REFNO)='" & flxSum.TextMatrix(ActRow, 2) & "') AND ((LEDGER.CUSTOMERID)=" & CustID & "));"
    
    'mySQL
    Qry = "SELECT LEDGER.* From LEDGER WHERE LEDGER.REFNO='" & flxSum.TextMatrix(ActRow, 2) & "' AND LEDGER.CUSTOMERID=" & CustID & ";"
    
    If CtrlNumExist(Qry) Then
        'Cancel Sales Return in Ledger
        '***Access
        'LYRcon.Execute "UPDATE LEDGER SET LEDGER.REFTYPE = 'CANCELLED', LEDGER.CUSTOMERID = 0 WHERE (((LEDGER.REFNO)='" & flxSum.TextMatrix(ActRow, 2) & "'));"
        'mySQL
        LYRcon.Execute "UPDATE LEDGER SET LEDGER.REFTYPE = 'CANCELLED', LEDGER.CUSTOMERID = 0 WHERE LEDGER.REFNO='" & flxSum.TextMatrix(ActRow, 2) & "';"
        
        'Update OS Bal in Ledger
        '***Access
        'InvBal = Val(GetFieldValue("SELECT SPINVOICE.AMOUNT From SPINVOICE WHERE (((SPINVOICE.SPINVNO)='" & flxSum.TextMatrix(ActRow, 3) & "'));"))
        'mySQL
        InvBal = Val(GetFieldValue("SELECT SPINVOICE.AMOUNT From SPINVOICE WHERE SPINVOICE.SPINVNO='" & flxSum.TextMatrix(ActRow, 3) & "';"))
        
        TtlAmtPd = 0
        '***Access
        'Qry = "SELECT LEDGER.* From LEDGER WHERE (((LEDGER.SPINVNO)='" & flxSum.TextMatrix(ActRow, 3) & "') AND ((LEDGER.CUSTOMERID)=" & CustID & ")) ORDER BY LEDGER.SEQNO;"
        
        'mySQL
        Qry = "SELECT LEDGER.* From LEDGER WHERE LEDGER.SPINVNO='" & flxSum.TextMatrix(ActRow, 3) & "' AND LEDGER.CUSTOMERID=" & CustID & " ORDER BY LEDGER.SEQNO;"
        
        With rsLedger   'Update Ledger
            .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
            If .RecordCount > 0 Then
                .MoveFirst
                Do While Not .EOF
                    MyCredit = IIf(IsNull(!CREDIT), 0, !CREDIT)
                    TtlAmtPd = TtlAmtPd + MyCredit
                    MyDebit = IIf(IsNull(!Debit), 0, !Debit)
                    If MyCredit > 0 Then
                        InvBal = InvBal - MyCredit
                    ElseIf MyDebit > 0 Then
                        InvBal = InvBal + MyDebit
                    End If
                    !OSbal = InvBal
                    .MoveNext
                Loop
            End If
            .Close
        End With
        'Update Onaccount table
        '***Access
        'LYRcon.Execute "UPDATE ONACCTERMS SET ONACCTERMS.TTLAMTPAID = " & TtlAmtPd & ", ONACCTERMS.BALANCE = " & InvBal & ", ONACCTERMS.STATUS = 'OPEN', ONACCTERMS.LASTUPDATE = #" & Now & "# " & _
                       "WHERE (((ONACCTERMS.SPINVNO)='" & flxSum.TextMatrix(ActRow, 3) & "'));"
        'mySQL
        LYRcon.Execute "UPDATE ONACCTERMS SET ONACCTERMS.TTLAMTPAID = " & TtlAmtPd & ", ONACCTERMS.BALANCE = " & InvBal & ", ONACCTERMS.STATUS = 'OPEN', ONACCTERMS.LASTUPDATE = '" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "' " & _
                       "WHERE ONACCTERMS.SPINVNO='" & flxSum.TextMatrix(ActRow, 3) & "';"
        
        
        'Update Customer Outstanding Balance
        '***Access
        'LYRcon.Execute "UPDATE CUSTOMER SET CUSTOMER.OSBAL = CUSTOMER.OSBAL+" & PureNum(flxSum.TextMatrix(ActRow, 6)) & " WHERE (((CUSTOMER.CUSTOMERID)=" & CustID & "));"
        
        'mySQL
        LYRcon.Execute "UPDATE CUSTOMER SET CUSTOMER.OSBAL = CUSTOMER.OSBAL+" & Val(Replace(flxSum.TextMatrix(ActRow, 6), ",", "")) & " WHERE CUSTOMER.CUSTOMERID=" & CustID & ";"
        
    End If
    'Update Audit Trans
    '***Access
    'LYRcon.Execute "INSERT INTO AUDITRANS ( TRANSDATE, TRANSTYPE, TRANSDOC, REFNO, REFDATE, REASON, EMPID ) " & _
            "VALUES (#" & Date & "#, 'CANCEL', 'SALES RETURN', '" & flxSum.TextMatrix(ActRow, 2) & "', #" & flxSum.TextMatrix(ActRow, 1) & "#, '" & MyReason & "', " & UserID & ");"
    'mySQL
    LYRcon.Execute "INSERT INTO AUDITRANS ( TRANSDATE, TRANSTYPE, TRANSDOC, REFNO, REFDATE, REASON, EMPID ) " & _
            "VALUES ('" & Format(Date, "yyyy-mm-dd") & "', 'CANCEL', 'SALES RETURN', '" & flxSum.TextMatrix(ActRow, 2) & "', '" & Format(flxSum.TextMatrix(ActRow, 1), "yyyy-mm-dd") & "', '" & MyReason & "', " & UserID & ");"
    
    'Update Grid
    With flxSum
        .TextMatrix(ActRow, 7) = "CANCELLED"
        .TextMatrix(ActRow, 6) = "0.00"
        .Col = 1: .Row = ActRow: .ColSel = 8: .RowSel = ActRow: .FillStyle = flexFillRepeat: .CellBackColor = RGB(240, 0, 0)
        .Col = 2
    End With
    MsgBox "Sales Return Cancelled!", vbInformation + vbOKOnly, "Cancel"
                    
End Sub

Function GetReqID(MyRN As String) As Long
    Dim rsReq As New ADODB.Recordset
    Dim Qry As String
    
    '***Access
    'Qry = "SELECT REQUISITION.REQID From Requisition WHERE (((REQUISITION.REQNO)='" & MyRN & "'));"
    'mySQL
    Qry = "SELECT REQUISITION.REQID From Requisition WHERE REQUISITION.REQNO='" & MyRN & "';"
    
    With rsReq
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            GetReqID = !ReqID
        Else
            GetReqID = 0
        End If
        .Close
    End With
End Function

Private Sub CancelRequisition()
    Dim rsSPItr As New ADODB.Recordset
    Dim Ans As Integer, i As Integer, MyReason As String, Qry As String
    Dim ReqID As Long
            
    'Validate Requisition if exist in invoice
    ReqID = GetReqID(flxSum.TextMatrix(ActRow, 2))
    '***Access
    'Qry = "SELECT SPINVOICE.* From SPINVOICE WHERE (((SPINVOICE.REQID)=" & ReqID & ") AND ((SPINVOICE.REMARKS)<>'CANCELLED'));"
    
    'mySQL
    Qry = "SELECT SPINVOICE.* From SPINVOICE WHERE SPINVOICE.REQID=" & ReqID & " AND SPINVOICE.REMARKS<>'CANCELLED';"
    
    If CtrlNumExist(Qry) Then
        MsgBox "Unable to cancel requisition because an Invoice was already issued!", vbExclamation + vbOKOnly, "Cancel Failed"
        Exit Sub
    End If
    
    If Val(flxSum.TextMatrix(ActRow, 9)) = 0 Then
        Ans = MsgBox("Requisition No. " & flxSum.TextMatrix(ActRow, 2) & " is not yet printed!" & vbCrLf & "You can still used the requisition to another customer..." & vbCrLf & vbCrLf & _
              "Are you sure you want to cancel requisition?", vbQuestion + vbYesNo, "Cancel")
        If Ans = vbNo Then Exit Sub
    End If
    Ans = MsgBox("Do you want to Cancel Requisition?" & vbCrLf & vbCrLf & _
        "S.I. Date : " & flxSum.TextMatrix(ActRow, 1) & vbCrLf & "Req. No: " & flxSum.TextMatrix(ActRow, 2) & vbCrLf & _
        "Charge to: " & flxSum.TextMatrix(ActRow, 3) & vbCrLf & _
        "Req. Amount: " & flxSum.TextMatrix(ActRow, 4), vbQuestion + vbYesNo, "Cancel")
    If Ans = vbNo Then Exit Sub
    MyReason = InputBox("Please indicate reason for Cancellation of Requisition:", "Reason for Cancel")
    If MyReason = "" Then
        MsgBox "Unable to indicate Reason..." & vbCrLf & "Cancellation of Requisition failed!", vbExclamation + vbOKOnly, "Cancel Failed"
        Exit Sub
    End If
    'Update Requisition
    '***Access
    'LYRcon.Execute "UPDATE REQUISITION SET REQUISITION.REMARKS = 'CANCELLED' WHERE (((REQUISITION.REQNO)='" & flxSum.TextMatrix(ActRow, 2) & "'));"
    'mySQL
    LYRcon.Execute "UPDATE REQUISITION SET REQUISITION.REMARKS = 'CANCELLED' WHERE REQUISITION.REQNO='" & flxSum.TextMatrix(ActRow, 2) & "';"
    
    'Update Grid
    flxSum.TextMatrix(ActRow, 8) = "CANCELLED"
    With flxSum
        .Col = 1: .Row = ActRow: .ColSel = 9: .RowSel = ActRow: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 100, 100)
        .Col = 2
    End With
    MsgBox "Requisition Cancelled!", vbInformation + vbOKOnly, "Requisition Summary"

End Sub


Private Sub CancelPurchaseReturn()
    Dim rsCDtls As New ADODB.Recordset
    Dim rsLedger As New ADODB.Recordset
    Dim Ans As Integer, MyReason As String, Qry As String, MyDetails As String, MyAcc As String
    Dim MyUnit As String, MyDesc As String, MyInv As String, MyProc As String, MySpc As String
    Dim i As Integer, CustID As Integer, InvBal As Single, TtlAmtPd As Currency, MyCredit As Currency, MyDebit As Currency
    
    'Display Details in flxCancelpuret
    '***Access
    'Qry = "SELECT SPITR.OUT, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPITR.PRICE, [OUT]*[PRICE] AS AMOUNT, SPITR.SPAREPARTID " & _
          "FROM SPITR INNER JOIN SPAREPARTS ON SPITR.SPAREPARTID = SPAREPARTS.SPAREPARTID " & _
          "WHERE (((SPITR.REFERENCE)='RET#" & flxSum.TextMatrix(ActRow, 2) & "') AND ((SPITR.TRANSTYPE)='PURCHASE RETURN'));"
    'mySQL
    Qry = "SELECT SPITR.`OUT`, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPITR.PRICE, SPITR.`OUT`*SPITR.PRICE AS AMOUNT, SPITR.SPAREPARTID " & _
          "FROM SPITR INNER JOIN SPAREPARTS ON SPITR.SPAREPARTID = SPAREPARTS.SPAREPARTID " & _
          "WHERE SPITR.REFERENCE='RET#" & flxSum.TextMatrix(ActRow, 2) & "' AND SPITR.TRANSTYPE='PURCHASE RETURN';"
    
    
    With rsCDtls
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        flxCancelPuRet.Rows = 1
        If .RecordCount > 0 Then
            .MoveFirst
            MyDetails = ""
            Do While Not .EOF
                                    '0-QTY,  1-PARTNO,  2-DESCRIPTION, 3-PRICE, 4-AMOUNT, 5-SPID
                flxCancelPuRet.AddItem !OUT & vbTab & !PartNumber & vbTab & !Description & vbTab & !Price & vbTab & !Amount & vbTab & !SparepartID
                If InStr(UCase$(!Description), "4T OIL") > 0 Then
                    MyUnit = IIf(!OUT > 1, "QRTS", "QRT")
                Else
                    MyUnit = IIf(!OUT > 1, "PCS", "PC")
                End If
                If Len(!Description) < 30 Then
                    MyDesc = !Description & Space$(30 - Len(!Description))
                Else
                    MyDesc = Left$(!Description, 30)
                End If
                If Len(!PartNumber) < 20 Then
                    MySpc = Space$(20 - Len(!PartNumber))
                Else
                    MySpc = ""
                End If
                MyDetails = IIf(MyDetails <> "", MyDetails & vbCrLf, "") & !OUT & vbTab & MyUnit & vbTab & !PartNumber & MySpc & vbTab & MyDesc & vbTab & Format(!Price, "#,##0.00") & vbTab & Format(!Amount, "#,##0.00")
                .MoveNext
            Loop
        End If
        .Close
    End With
    Ans = MsgBox("Do you want to Cancel Purchase Return?" & vbCrLf & vbCrLf & _
        "Purchase Return No.: " & flxSum.TextMatrix(ActRow, 2) & vbTab & vbTab & vbTab & vbTab & "Purchase Return Date : " & flxSum.TextMatrix(ActRow, 1) & vbCrLf & vbCrLf & _
        "QTY." & vbTab & "UNIT" & vbTab & "PART NUMBER" & vbTab & vbTab & "DESCRIPTION" & vbTab & vbTab & vbTab & "PRICE" & vbTab & "AMOUNT" & vbCrLf & MyDetails, vbYesNo, "Cancel")
    If Ans = vbNo Then Exit Sub
    MyReason = InputBox("Please indicate reason for Cancellation of Sales Return:", "Reason for Cancel")
    If MyReason = "" Then
        MsgBox "Unable to indicate Reason..." & vbCrLf & "Cancellation of Sales Return failed!", vbExclamation + vbOKOnly, "Cancel Failed"
        Exit Sub
    End If
    'Retrace transaction made
    'Update Sales Return Remarks to CANCELLED
    '***Access
    'LYRcon.Execute "UPDATE PURCHASERET SET PURCHASERET.REMARKS = 'CANCELLED', PURCHASERET.AMOUNT = 0 WHERE (((PURCHASERET.PURTRNID)=" & flxSum.TextMatrix(ActRow, 9) & "));"
    'mySQL
    LYRcon.Execute "UPDATE PURCHASERET SET PURCHASERET.REMARKS = 'CANCELLED', PURCHASERET.AMOUNT = 0 WHERE PURCHASERET.PURTRNID=" & flxSum.TextMatrix(ActRow, 9) & ";"
    
    With flxCancelPuRet
        For i = 1 To .Rows - 1
            'Cancel SPTIR
            '***Access
            'LYRcon.Execute "UPDATE SPITR SET SPITR.TRANSTYPE = 'CANCELLED', SPITR.OUT = 0 " & _
                           "WHERE (((SPITR.REFERENCE)='RET#" & flxSum.TextMatrix(ActRow, 2) & "') AND ((SPITR.SPAREPARTID)=" & .TextMatrix(i, 5) & "));"
            'mySQL
            LYRcon.Execute "UPDATE SPITR SET SPITR.TRANSTYPE = 'CANCELLED', SPITR.`OUT` = 0 " & _
                           "WHERE SPITR.REFERENCE='RET#" & flxSum.TextMatrix(ActRow, 2) & "' AND SPITR.SPAREPARTID=" & .TextMatrix(i, 5) & ";"
            
            'Update Qty-On Hand
            'LYRcon.Execute "UPDATE SPAREPARTS SET SPAREPARTS.QTYONHAND = [SPAREPARTS].[QTYONHAND]+" & Val(.TextMatrix(i, 0)) & _
                           " WHERE (((SPAREPARTS.SPAREPARTID)=" & .TextMatrix(i, 5) & "));"
            'mySQl
            LYRcon.Execute "UPDATE SPAREPARTS SET SPAREPARTS.QTYONHAND = SPAREPARTS.QTYONHAND+" & Val(.TextMatrix(i, 0)) & _
                           " WHERE SPAREPARTS.SPAREPARTID=" & .TextMatrix(i, 5) & ";"
        
        Next i
    End With
    
    'Update Audit Trans
    '***Access
    'LYRcon.Execute "INSERT INTO AUDITRANS ( TRANSDATE, TRANSTYPE, TRANSDOC, REFNO, REFDATE, REASON, EMPID ) " & _
            "VALUES (#" & Date & "#, 'CANCEL', 'PURCHASE RETURN', '" & flxSum.TextMatrix(ActRow, 2) & "', #" & flxSum.TextMatrix(ActRow, 1) & "#, '" & MyReason & "', " & UserID & ");"
    'mySQL
    LYRcon.Execute "INSERT INTO AUDITRANS ( TRANSDATE, TRANSTYPE, TRANSDOC, REFNO, REFDATE, REASON, EMPID ) " & _
            "VALUES ('" & Format(Date, "yyyy-mm-dd") & "', 'CANCEL', 'PURCHASE RETURN', '" & flxSum.TextMatrix(ActRow, 2) & "', '" & Format(flxSum.TextMatrix(ActRow, 1), "yyyy-mm-dd") & "', '" & MyReason & "', " & UserID & ");"
    
    'Update Grid
    With flxSum
        .TextMatrix(ActRow, 7) = "CANCELLED"
        .TextMatrix(ActRow, 6) = "0.00"
        .Col = 1: .Row = ActRow: .ColSel = 9: .RowSel = ActRow: .FillStyle = flexFillRepeat: .CellBackColor = RGB(240, 0, 0)
        .Col = 2
    End With
    MsgBox "Purchase Return Cancelled!", vbInformation + vbOKOnly, "Cancel"
                    
End Sub
Private Sub CancelCustPO()
    Dim MyReason As String, Qry As String
    Dim ReqID As Long, Ans As Integer, i As Integer
            
    'Validate  if existing invoice is cancelled
    If flxSum.TextMatrix(ActRow, 5) <> "" Then
        '***Access
        'Qry = "SELECT SPINVOICE.* From SPINVOICE WHERE (((SPINVOICE.SPINVNO)='" & flxSum.TextMatrix(ActRow, 5) & "') or ((SPINVOICE.REMARKS)<>'CANCELLED'));"
        'mySql
        Qry = "SELECT SPINVOICE.* From SPINVOICE WHERE SPINVOICE.SPINVNO='" & flxSum.TextMatrix(ActRow, 5) & "' or SPINVOICE.REMARKS<>'CANCELLED';"
        
        If CtrlNumExist(Qry) Then
            MsgBox "Unable to cancel P.O. because an Invoice was alread issued!", vbExclamation + vbOKOnly, "Cancel Failed"
            Exit Sub
        End If
    End If
    If Val(flxSum.TextMatrix(ActRow, 7)) = 0 Then
        Ans = MsgBox("P.O. No. " & flxSum.TextMatrix(ActRow, 2) & " is not yet printed!" & vbCrLf & "You can still edit the P.O." & vbCrLf & vbCrLf & _
              "Are you sure you want to cancel P.O.?", vbQuestion + vbYesNo, "Cancel")
        If Ans = vbNo Then Exit Sub
    End If
    Ans = MsgBox("Do you want to Cancel P.O.?" & vbCrLf & vbCrLf & _
        "P.O. Date : " & flxSum.TextMatrix(ActRow, 1) & vbCrLf & "P.O. No: " & flxSum.TextMatrix(ActRow, 2) & vbCrLf & _
        "Charge to: " & flxSum.TextMatrix(ActRow, 3), vbQuestion + vbYesNo, "Cancel")
    If Ans = vbNo Then Exit Sub
    MyReason = InputBox("Please indicate reason for Cancellation of Requisition:", "Reason for Cancel")
    If MyReason = "" Then
        MsgBox "Unable to indicate Reason..." & vbCrLf & "Cancellation of Requisition failed!", vbExclamation + vbOKOnly, "Cancel Failed"
        Exit Sub
    End If
    'Update Requisition
    '***Access
    'LYRcon.Execute "UPDATE CUSTPO SET CUSTPO.REMARKS = 'CANCELLED ' & [CUSTPO].[REMARKS] WHERE (((CUSTPO.CPOID)=" & flxSum.TextMatrix(ActRow, 8) & "));"
    'mySQL
    LYRcon.Execute "UPDATE CUSTPO SET CUSTPO.REMARKS = 'CANCELLED ' & CUSTPO.REMARKS WHERE CUSTPO.CPOID=" & flxSum.TextMatrix(ActRow, 8) & ";"

    'Update Grid
    flxSum.TextMatrix(ActRow, 6) = "CANCELLED"
    With flxSum
        .Col = 1: .Row = ActRow: .ColSel = 8: .RowSel = ActRow: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 150, 150)
        .Col = 2
    End With
    MsgBox "P.O. is now Cancelled!", vbInformation + vbOKOnly, "P.O. Summary"

End Sub
Private Sub CancelInvAdj()
    Dim rsCDtls As New ADODB.Recordset
    Dim Ans As Integer, MyReason As String, Qry As String, MyDetails As String, MyAcc As String
    
    If InStr(UCase$(flxSum.TextMatrix(ActRow, 5)), "CANCELLED") > 0 Then
        MsgBox "Inv. Adjustment No. " & flxSum.TextMatrix(ActRow, 2) & " was already cancelled!", vbExclamation + vbOKOnly, "Cancel"
        Exit Sub
    End If
'    If Val(flxSum.TextMatrix(ActRow, 8)) > 0 Then
'        MsgBox "Inv. Adjustment No. " & flxSum.TextMatrix(ActRow, 2) & " was already printed and final!", vbExclamation + vbOKOnly, "Cancel"
'        Exit Sub
'    End If
    
    'Display Details in flxCancel
    '***Access
    'Qry = "SELECT SPAREPARTS.PARTNUMBER, SPITR.REMARKS, SPITR.RECEIVE, SPITR.OUT, SPITR.SOURCE, SPITR.SPAREPARTID " & _
          "FROM SPITR INNER JOIN SPAREPARTS ON SPITR.SPAREPARTID = SPAREPARTS.SPAREPARTID " & _
          "WHERE (((SPITR.REFERENCE)='ADJ#" & flxSum.TextMatrix(ActRow, 2) & "') AND ((SPITR.TRANSTYPE)='INVENTORY ADJ.'));"
    'mySQL
    Qry = "SELECT SPAREPARTS.PARTNUMBER, SPITR.REMARKS, SPITR.RECEIVE, SPITR.`OUT`, SPITR.SOURCE, SPITR.SPAREPARTID " & _
          "FROM SPITR INNER JOIN SPAREPARTS ON SPITR.SPAREPARTID = SPAREPARTS.SPAREPARTID " & _
          "WHERE SPITR.REFERENCE='ADJ#" & flxSum.TextMatrix(ActRow, 2) & "' AND SPITR.TRANSTYPE='INVENTORY ADJ.';"
    
    
    With rsCDtls
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        flxCancel.Rows = 1
        If .RecordCount > 0 Then
            .MoveFirst
            MyDetails = ""
            Do While Not .EOF
                '0-Part Number,  1-Adj. Type,  2-IN, 3-OUT, 4-Source, 5-SPID
                flxCancel.AddItem !PartNumber & vbTab & IIf(IsNull(!Remarks), "", !Remarks) & vbTab & !Receive & vbTab & !OUT & vbTab & IIf(IsNull(!Source), "", !Source) & vbTab & !SparepartID
                MyDetails = IIf(MyDetails <> "", MyDetails & vbCrLf, "") & !PartNumber & vbTab & flxCancel.TextMatrix(flxCancel.Rows - 1, 1) & Space$(2) & vbTab & IIf(InStr(UCase$(flxCancel.TextMatrix(flxCancel.Rows - 1, 1)), "DEBIT") > 0, !Receive, !OUT) & vbTab & flxCancel.TextMatrix(flxCancel.Rows - 1, 4)
                .MoveNext
            Loop
        End If
        .Close
    End With
    Ans = MsgBox("Do you want to Cancel Adjustment?" & vbCrLf & vbCrLf & _
        "Adjustment No: " & flxSum.TextMatrix(ActRow, 2) & vbTab & "Adjustment Date : " & flxSum.TextMatrix(ActRow, 1) & vbCrLf & vbCrLf & _
        "PART NUMBER" & vbTab & "ADJ. TYPE" & vbTab & "QTY." & vbTab & "SOURCE" & vbCrLf & MyDetails, vbYesNo, "Cancel")
    If Ans = vbNo Then Exit Sub
    MyReason = InputBox("Please indicate reason for Cancellation of Adjustment:", "Reason for Cancel")
    If MyReason = "" Then
        MsgBox "Unable to indicate Reason..." & vbCrLf & "Cancellation of Adjustment failed!", vbExclamation + vbOKOnly, "Cancel Failed"
        Exit Sub
    End If
    'Retrace transaction made
    '***Access
    'LYRcon.Execute "UPDATE INVNTRYADJ SET INVNTRYADJ.REMARKS = 'CANCELLED' WHERE (((INVNTRYADJ.ADJNO)='" & flxSum.TextMatrix(ActRow, 2) & "'));"
    'Qry = "SELECT INVNTRYADJ.APPLYCHARGES, INVNTRYADJ.ADJNO From InvntryADJ WHERE (((INVNTRYADJ.ADJNO)='" & flxSum.TextMatrix(ActRow, 2) & "'));"
    'If GetFieldValue(Qry) = True Then   'Update Apply Charges if true
    '    LYRcon.Execute "UPDATE INVADJCHARGES SET INVADJCHARGES.EMPID = 0 WHERE (((INVADJCHARGES.ADJNO)='" & flxSum.TextMatrix(ActRow, 2) & "'));"
    'End If
    
    'mySQL
    LYRcon.Execute "UPDATE INVNTRYADJ SET INVNTRYADJ.REMARKS = 'CANCELLED' WHERE INVNTRYADJ.ADJNO='" & flxSum.TextMatrix(ActRow, 2) & "';"
    Qry = "SELECT INVNTRYADJ.APPLYCHARGES, INVNTRYADJ.ADJNO From InvntryADJ WHERE INVNTRYADJ.ADJNO='" & flxSum.TextMatrix(ActRow, 2) & "';"
    If GetFieldValue(Qry) = True Then   'Update Apply Charges if true
        LYRcon.Execute "UPDATE INVADJCHARGES SET INVADJCHARGES.EMPID = 0 WHERE INVADJCHARGES.ADJNO='" & flxSum.TextMatrix(ActRow, 2) & "';"
    End If
 
    'Update Details
    With flxCancel
        For i = 1 To .Rows - 1
            'If InStr(UCase$(.TextMatrix(i, 1)), "DEBIT") > 0 Then
            '    LYRcon.Execute "UPDATE SPAREPARTS SET SPAREPARTS.QTYONHAND = [SPAREPARTS].[QTYONHAND]-" & Val(.TextMatrix(i, 2)) & _
            '                   " WHERE (((SPAREPARTS.SPAREPARTID)=" & .TextMatrix(i, 5) & "));"
            '    LYRcon.Execute "UPDATE SPITR SET SPITR.TRANSTYPE = 'CANCELLED', SPITR.RECEIVE = 0 " & _
            '                   "WHERE (((SPITR.REFERENCE)='ADJ#" & flxSum.TextMatrix(ActRow, 2) & "') AND ((SPITR.SPAREPARTID)=" & .TextMatrix(i, 5) & "));"
            'Else
            '    LYRcon.Execute "UPDATE SPAREPARTS SET SPAREPARTS.QTYONHAND = [SPAREPARTS].[QTYONHAND]+" & Val(.TextMatrix(i, 3)) & _
            '                   " WHERE (((SPAREPARTS.SPAREPARTID)=" & .TextMatrix(i, 5) & "));"
            '    LYRcon.Execute "UPDATE SPITR SET SPITR.TRANSTYPE = 'CANCELLED', SPITR.OUT = 0 " & _
            '                   "WHERE (((SPITR.REFERENCE)='ADJ#" & flxSum.TextMatrix(ActRow, 2) & "') AND ((SPITR.SPAREPARTID)=" & .TextMatrix(i, 5) & "));"
            'End If
            'mySQL
            If InStr(UCase$(.TextMatrix(i, 1)), "DEBIT") > 0 Then
                LYRcon.Execute "UPDATE SPAREPARTS SET SPAREPARTS.QTYONHAND = SPAREPARTS.QTYONHAND-" & Val(.TextMatrix(i, 2)) & _
                               " WHERE SPAREPARTS.SPAREPARTID=" & .TextMatrix(i, 5) & ";"
                LYRcon.Execute "UPDATE SPITR SET SPITR.TRANSTYPE = 'CANCELLED', SPITR.RECEIVE = 0 " & _
                               "WHERE SPITR.REFERENCE='ADJ#" & flxSum.TextMatrix(ActRow, 2) & "' AND SPITR.SPAREPARTID=" & .TextMatrix(i, 5) & ";"
            Else
                LYRcon.Execute "UPDATE SPAREPARTS SET SPAREPARTS.QTYONHAND = SPAREPARTS.QTYONHAND+" & Val(.TextMatrix(i, 3)) & _
                               " WHERE SPAREPARTS.SPAREPARTID=" & .TextMatrix(i, 5) & ";"
                LYRcon.Execute "UPDATE SPITR SET SPITR.TRANSTYPE = 'CANCELLED', SPITR.`OUT` = 0 " & _
                               "WHERE SPITR.REFERENCE='ADJ#" & flxSum.TextMatrix(ActRow, 2) & "' AND SPITR.SPAREPARTID=" & .TextMatrix(i, 5) & ";"
            End If
        Next i
    End With
    
    'Update Audit Trans
    '***Access
    'LYRcon.Execute "INSERT INTO AUDITRANS ( TRANSDATE, TRANSTYPE, TRANSDOC, REFNO, REFDATE, REASON, EMPID ) " & _
            "VALUES (#" & Date & "#, 'CANCEL', 'INVENTORY ADJ.', '" & flxSum.TextMatrix(ActRow, 2) & "', #" & flxSum.TextMatrix(ActRow, 1) & "#, '" & MyReason & "', " & UserID & ");"
    'mySQL
    LYRcon.Execute "INSERT INTO AUDITRANS ( TRANSDATE, TRANSTYPE, TRANSDOC, REFNO, REFDATE, REASON, EMPID ) " & _
            "VALUES ('" & Format(Date, "yyyy-mm-dd") & "', 'CANCEL', 'INVENTORY ADJ.', '" & flxSum.TextMatrix(ActRow, 2) & "', '" & Format(flxSum.TextMatrix(ActRow, 1), "yyyy-mm-dd") & "', '" & MyReason & "', " & UserID & ");"
   
   'Update Grid
    flxSum.TextMatrix(ActRow, 5) = "CANCELLED"
    With flxSum     'Change row color
        .Col = 1: .Row = ActRow: .ColSel = 8: .RowSel = ActRow: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 0, 0)
        .Col = 2
    End With
    MsgBox "Adjustment Cancelled!", vbInformation + vbOKOnly

End Sub
    
Private Sub CancelDMCM()
    Dim Ans As Integer, Bal As Currency, AmtAdj As Currency
    Dim Qry As String, MyTrans As String, RefType As String, AdjType As String, MyStat As String, JOrem As String
    Dim MyReason As String, MyAcc As String
    
    If InStr(UCase$(flxSum.TextMatrix(ActRow, 10)), "CANCELLED") > 0 Then
        MsgBox "DMCM No. " & flxSum.TextMatrix(ActRow, 2) & " was already cancelled!", vbExclamation + vbOKOnly, "Cancel"
        Exit Sub
    End If
'    If Val(flxSum.TextMatrix(ActRow, 11)) > 0 Then
'        MsgBox "DMCM No. " & flxSum.TextMatrix(ActRow, 2) & " was already printed and final!", vbExclamation + vbOKOnly, "Cancel"
'        Exit Sub
'    End If
    
    Ans = MsgBox("Do you want to Cancel DMCM?", vbQuestion + vbYesNo, "Cancel")
    If Ans = vbNo Then Exit Sub
    MyReason = InputBox("Please indicate reason for Cancellation of DMCM:", "Reason for Cancel")
    If MyReason = "" Then
        MsgBox "Unable to indicate Reason..." & vbCrLf & "Cancellation of Adjustment failed!", vbExclamation + vbOKOnly, "Cancel Failed"
        Exit Sub
    End If
        
    MyTrans = flxSum.TextMatrix(ActRow, 4)
    If Val(Replace(flxSum.TextMatrix(ActRow, 7), ",", "")) > 0 Then
        AmtAdj = Val(Replace(flxSum.TextMatrix(ActRow, 7), ",", ""))
        AdjType = "CREDIT"
    Else
        AmtAdj = Val(Replace(flxSum.TextMatrix(ActRow, 6), ",", ""))
        AdjType = "DEBIT"
    End If
    Bal = Val(Replace(flxSum.TextMatrix(ActRow, 8), ",", ""))
    'UPDATE AdjDMCM
    '***Access
    'LYRcon.Execute "UPDATE ADJDMCM SET ADJDMCM.REASON = 'CANCELLED', ADJDMCM.CUSTOMERID = 0 WHERE (((ADJDMCM.ADJNO)='" & flxSum.TextMatrix(ActRow, 2) & "'));"
    
    'mySQL
    LYRcon.Execute "UPDATE ADJDMCM SET ADJDMCM.REASON = 'CANCELLED', ADJDMCM.CUSTOMERID = 0 WHERE ADJDMCM.ADJNO='" & flxSum.TextMatrix(ActRow, 2) & "';"
    
    If MyTrans = "INVOICE" Then
        '***Access
        'Qry = "SELECT SPINVOICE.INVTYPE From SPINVOICE WHERE (((SPINVOICE.SPINVNO)='" & flxSum.TextMatrix(ActRow, 4) & "'));"
        
        'mySQL
        Qry = "SELECT SPINVOICE.INVTYPE From SPINVOICE WHERE SPINVOICE.SPINVNO='" & flxSum.TextMatrix(ActRow, 5) & "';"
        
        RefType = GetFieldValue(Qry)
        If RefType = "CREDIT" Then  'INVOICE TYPE = CREDIT
            Call UpdateLedger(flxSum.TextMatrix(ActRow, 2))
            '***Access
            'LYRcon.Execute "UPDATE SPINVOICE SET SPINVOICE.REMARKS = Null WHERE (((SPINVOICE.SPINVNO)='" & flxSum.TextMatrix(ActRow, 5) & "'));"
            
            'mySQL
            LYRcon.Execute "UPDATE SPINVOICE SET SPINVOICE.REMARKS = Null WHERE SPINVOICE.SPINVNO='" & flxSum.TextMatrix(ActRow, 5) & "';"
            
            'update onaccount
            If AdjType = "CREDIT" Then  'DMCM ADJ TYPE = CREDIT
                MyStat = IIf(Bal + AmtAdj > 0, "OPEN", "CLOSED")
                '***Access
                'LYRcon.Execute "UPDATE ONACCTERMS SET ONACCTERMS.BALANCE = " & Bal + AmtAdj & ", ONACCTERMS.STATUS = '" & MyStat & "' " & _
                               "WHERE (((ONACCTERMS.SPINVNO)='" & flxSum.TextMatrix(ActRow, 5) & "'));"
                'mySQL
                LYRcon.Execute "UPDATE ONACCTERMS SET ONACCTERMS.BALANCE = " & Bal + AmtAdj & ", ONACCTERMS.STATUS = '" & MyStat & "' " & _
                               "WHERE ONACCTERMS.SPINVNO='" & flxSum.TextMatrix(ActRow, 5) & "';"
                               
             Else
                MyStat = IIf(Bal - AmtAdj > 0, "OPEN", "CLOSED")
                '***Access
                'LYRcon.Execute "UPDATE ONACCTERMS SET ONACCTERMS.BALANCE = " & Bal - AmtAdj & ", ONACCTERMS.STATUS = '" & MyStat & "' " & _
                               "WHERE (((ONACCTERMS.SPINVNO)='" & flxSum.TextMatrix(ActRow, 5) & "'));"
                'mySQL
                LYRcon.Execute "UPDATE ONACCTERMS SET ONACCTERMS.BALANCE = " & Bal - AmtAdj & ", ONACCTERMS.STATUS = '" & MyStat & "' " & _
                               "WHERE ONACCTERMS.SPINVNO='" & flxSum.TextMatrix(ActRow, 5) & "';"
                
            End If
        ElseIf RefType = "CASH" Then
            If AdjType = "CREDIT" Then
                '***Access
                'LYRcon.Execute "UPDATE SPINVOICE SET SPINVOICE.AMOUNT = " & Bal + AmtAdj & ", SPINVOICE.REMARKS = Null " & _
                               "WHERE (((SPINVOICE.SPINVNO)='" & flxSum.TextMatrix(ActRow, 5) & "'));"
                'mySQl
                LYRcon.Execute "UPDATE SPINVOICE SET SPINVOICE.AMOUNT = " & Bal + AmtAdj & ", SPINVOICE.REMARKS = Null " & _
                               "WHERE SPINVOICE.SPINVNO='" & flxSum.TextMatrix(ActRow, 5) & "';"
            
            ElseIf AdjType = "DEBIT" Then
                '***Access
                'LYRcon.Execute "UPDATE SPINVOICE SET SPINVOICE.AMOUNT = " & Bal - AmtAdj & ", SPINVOICE.REMARKS = Null " & _
                               "WHERE (((SPINVOICE.SPINVNO)='" & flxSum.TextMatrix(ActRow, 5) & "'));"
                'mySQL
                LYRcon.Execute "UPDATE SPINVOICE SET SPINVOICE.AMOUNT = " & Bal - AmtAdj & ", SPINVOICE.REMARKS = Null " & _
                               "WHERE SPINVOICE.SPINVNO='" & flxSum.TextMatrix(ActRow, 5) & "';"
                
            End If
        End If
    ElseIf MyTrans = "JOB ORDER" Then
        '***Access
        'Qry = "SELECT SRVCJOBORDER.PAYTYPE From SRVCJOBORDER WHERE (((SRVCJOBORDER.JOBORDERNO)='" & flxSum.TextMatrix(ActRow, 5) & "'));"
        'mySQL
        Qry = "SELECT SRVCJOBORDER.PAYTYPE From SRVCJOBORDER WHERE SRVCJOBORDER.JOBORDERNO='" & flxSum.TextMatrix(ActRow, 5) & "';"
        
        RefType = GetFieldValue(Qry)
        '***Access
        'Qry = "SELECT SRVCJOBORDER.REMARKS From SRVCJOBORDER WHERE (((SRVCJOBORDER.JOBORDERNO)='" & flxSum.TextMatrix(ActRow, 5) & "'));"
        'mySQL
        Qry = "SELECT SRVCJOBORDER.REMARKS From SRVCJOBORDER WHERE SRVCJOBORDER.JOBORDERNO='" & flxSum.TextMatrix(ActRow, 5) & "';"
        
        JOrem = GetFieldValue(Qry)
        JOrem = Replace(JOrem, "DMCM ADJ# " & flxSum.TextMatrix(ActRow, 2), "")
        If RefType = "ON ACCOUNT" Then  'JOB ORDER PAY TYPE = ON ACCOUNT
            Call UpdateLedger(flxSum.TextMatrix(ActRow, 2))
            '***Access
            'LYRcon.Execute "UPDATE SRVCJOBORDER SET SRVCJOBORDER.REMARKS = '" & JOrem & "' " & _
                       "WHERE (((SRVCJOBORDER.JOBORDERNO)='" & flxSum.TextMatrix(ActRow, 5) & "'));"
            'mySQL
            LYRcon.Execute "UPDATE SRVCJOBORDER SET SRVCJOBORDER.REMARKS = '" & JOrem & "' " & _
                       "WHERE SRVCJOBORDER.JOBORDERNO='" & flxSum.TextMatrix(ActRow, 5) & "';"
            
            'update onaccount
            If AdjType = "CREDIT" Then  'DMCM ADJ TYPE = CREDIT
                MyStat = IIf(Bal + AmtAdj > 0, "OPEN", "CLOSED")
                '***Access
                'LYRcon.Execute "UPDATE ONACCTERMS SET ONACCTERMS.BALANCE = " & Bal + AmtAdj & ", ONACCTERMS.STATUS = '" & MyStat & "' " & _
                               "WHERE (((ONACCTERMS.JOBORDERNO)='" & flxSum.TextMatrix(ActRow, 5) & "') AND ((ONACCTERMS.SPINVNO) Is Null));"
                'mySQL
                LYRcon.Execute "UPDATE ONACCTERMS SET ONACCTERMS.BALANCE = " & Bal + AmtAdj & ", ONACCTERMS.STATUS = '" & MyStat & "' " & _
                               "WHERE ONACCTERMS.JOBORDERNO='" & flxSum.TextMatrix(ActRow, 5) & "' AND (ONACCTERMS.SPINVNO) Is Null;"
                
            Else
                MyStat = IIf(Bal - AmtAdj > 0, "OPEN", "CLOSED")
                '***Access
                'LYRcon.Execute "UPDATE ONACCTERMS SET ONACCTERMS.BALANCE = " & Bal - AmtAdj & ", ONACCTERMS.STATUS = '" & MyStat & "' " & _
                               "WHERE (((ONACCTERMS.JOBORDERNO)='" & flxSum.TextMatrix(ActRow, 5) & "') AND ((ONACCTERMS.SPINVNO) Is Null));"
                'mySQL
                LYRcon.Execute "UPDATE ONACCTERMS SET ONACCTERMS.BALANCE = " & Bal - AmtAdj & ", ONACCTERMS.STATUS = '" & MyStat & "' " & _
                               "WHERE ONACCTERMS.JOBORDERNO='" & flxSum.TextMatrix(ActRow, 5) & "' AND (ONACCTERMS.SPINVNO) Is Null;"
            End If
        ElseIf RefType = "CASH" Then
            If AdjType = "CREDIT" Then
                '***Access
                'LYRcon.Execute "UPDATE SRVCJOBORDER SET SRVCJOBORDER.BALANCE = " & Bal + AmtAdj & ", SRVCJOBORDER.REMARKS = '" & JOrem & "' " & _
                               "WHERE (((SRVCJOBORDER.JOBORDERNO)='" & flxSum.TextMatrix(ActRow, 5) & "'));"
                'mySQL
                LYRcon.Execute "UPDATE SRVCJOBORDER SET SRVCJOBORDER.BALANCE = " & Bal + AmtAdj & ", SRVCJOBORDER.REMARKS = '" & JOrem & "' " & _
                               "WHERE SRVCJOBORDER.JOBORDERNO='" & flxSum.TextMatrix(ActRow, 5) & "';"
                
            ElseIf AdjType = "DEBIT" Then
                '***Access
                'LYRcon.Execute "UPDATE SRVCJOBORDER SET SRVCJOBORDER.BALANCE = " & Bal - AmtAdj & ", SRVCJOBORDER.REMARKS = '" & JOrem & "' " & _
                               "WHERE (((SRVCJOBORDER.JOBORDERNO)='" & flxSum.TextMatrix(ActRow, 5) & "'));"
                'mySQL
                LYRcon.Execute "UPDATE SRVCJOBORDER SET SRVCJOBORDER.BALANCE = " & Bal - AmtAdj & ", SRVCJOBORDER.REMARKS = '" & JOrem & "' " & _
                               "WHERE SRVCJOBORDER.JOBORDERNO='" & flxSum.TextMatrix(ActRow, 5) & "';"
                
            End If
        End If
    End If
   'Update Grid
    flxSum.TextMatrix(ActRow, 10) = "CANCELLED"
    With flxSum     'Change row color
        .Col = 1: .Row = ActRow: .ColSel = 11: .RowSel = ActRow: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 0, 0)
        .Col = 2
    End With
    '***Access
    'LYRcon.Execute "INSERT INTO AUDITRANS ( TRANSDATE, TRANSTYPE, TRANSDOC, REFNO, REFDATE, REASON, EMPID ) " & _
                   "VALUES (#" & Date & "#, 'CANCEL', 'DMCM ADJ.', '" & flxSum.TextMatrix(ActRow, 2) & "', #" & flxSum.TextMatrix(ActRow, 1) & "#, '" & MyReason & "', " & UserID & ");"
    'mySQL
    LYRcon.Execute "INSERT INTO AUDITRANS ( TRANSDATE, TRANSTYPE, TRANSDOC, REFNO, REFDATE, REASON, EMPID ) " & _
                   "VALUES ('" & Format(Date, "yyyy-mm-dd") & "', 'CANCEL', 'DMCM ADJ.', '" & flxSum.TextMatrix(ActRow, 2) & "', '" & Format(flxSum.TextMatrix(ActRow, 1), "yyyy-mm-dd") & "', '" & MyReason & "', " & UserID & ");"
    
    MsgBox "Adjustment Cancelled!", vbInformation + vbOKOnly
End Sub

Private Sub UpdateLedger(MyRefNo As String)
    Dim rsLedger As New ADODB.Recordset
    Dim Qry As String
    Dim i As Integer, CustID As Integer
    Dim MyCredit As Currency, MyDebit As Currency, InvBal As Currency
    
    '***Access
    'Qry = "SELECT LEDGER.* FROM LEDGER WHERE (((LEDGER.REFNO)='" & MyRefNo & "') AND ((LEDGER.REFTYPE)='DMCM'));"
    'myql
    Qry = "SELECT LEDGER.* FROM LEDGER WHERE LEDGER.REFNO='" & MyRefNo & "' AND LEDGER.REFTYPE='DMCM';"
    
    With rsLedger   'Update Ledger
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            flxInvCancel.Rows = 1
            CustID = !CUSTOMERID
            Do While Not .EOF
                MyJOid = IIf(IsNull(!JOBID), 0, !JOBID)
                flxInvCancel.AddItem !SPINVNO & vbTab & IIf(MyJOid > 0, "JOB", "INV")
                .MoveNext
            Loop
        End If
        .Close
    End With
    For i = 1 To flxInvCancel.Rows - 1
        If flxInvCancel.TextMatrix(i, 1) = "INV" Then
            '***Access
            'Qry = "SELECT LEDGER.* FROM LEDGER WHERE (((LEDGER.SPINVNO)='" & flxInvCancel.TextMatrix(i, 0) & "') AND ((LEDGER.CUSTOMERID)=" & CustID & ")) ORDER BY LEDGER.SEQNO;"
            'mySQL
            Qry = "SELECT LEDGER.* FROM LEDGER WHERE LEDGER.SPINVNO='" & flxInvCancel.TextMatrix(i, 0) & "' AND LEDGER.CUSTOMERID=" & CustID & " ORDER BY LEDGER.SEQNO;"
            
            InvBal = GetInvAmt(flxInvCancel.TextMatrix(i, 0))
        Else
            '***Access
            'Qry = "SELECT ONACCTERMS.ACCTAMOUNT FROM ONACCTERMS WHERE (((ONACCTERMS.SPINVNO) Is Null) AND ((ONACCTERMS.JOBORDERNO)='" & flxInvCancel.TextMatrix(i, 0) & "'));"
            'mySQL
             Qry = "SELECT ONACCTERMS.ACCTAMOUNT FROM ONACCTERMS WHERE (ONACCTERMS.SPINVNO) Is Null AND ONACCTERMS.JOBORDERNO='" & flxInvCancel.TextMatrix(i, 0) & "';"
            
            InvBal = GetFieldValue(Qry)
            '***Access
            'Qry = "SELECT LEDGER.* From Ledger WHERE (((LEDGER.SPINVNO)='" & flxInvCancel.TextMatrix(i, 0) & "') AND ((LEDGER.CUSTOMERID)=" & CustID & ") AND ((LEDGER.JOBID)>0)) ORDER BY LEDGER.SEQNO;"
            'mySQL
            Qry = "SELECT LEDGER.* From Ledger WHERE LEDGER.SPINVNO='" & flxInvCancel.TextMatrix(i, 0) & "' AND LEDGER.CUSTOMERID=" & CustID & " AND LEDGER.JOBID>0 ORDER BY LEDGER.SEQNO;"
            
        End If
        With rsLedger   'Update Ledger
            .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
            If .RecordCount > 0 Then
                .MoveFirst
                Do While Not .EOF
                    If !RefNo <> MyRefNo And !RefType <> "DMCM" Then
                        MyCredit = IIf(IsNull(!CREDIT), 0, !CREDIT)
                        TtlAmtPd = TtlAmtPd + MyCredit
                        MyDebit = IIf(IsNull(!Debit), 0, !Debit)
                        If MyCredit > 0 Then
                            InvBal = InvBal - MyCredit
                        ElseIf MyDebit > 0 Then
                            InvBal = InvBal + MyDebit
                        End If
                        !OSbal = InvBal
                    End If
                    .MoveNext
                Loop
            End If
            .Close
        End With
    Next i
    'Update Ledger's RefType to CANCELLED and CustomerID=0 where RefNo=MyOR
    LYRcon.Execute "UPDATE Ledger SET Ledger.REFTYPE = 'CANCELLED', Ledger.CUSTOMERID = 0 WHERE Ledger.REFTYPE='DMCM' AND Ledger.REFNO='" & MyRefNo & "';"
End Sub

Function GetInvAmt(InvNo As String) As Single
    Dim rsInv As New ADODB.Recordset
    Dim Qry As String
    
    '***Access
    'Qry = "SELECT SPINVOICE.AMOUNT From SPINVOICE WHERE (((SPINVOICE.SPINVNO)='" & InvNo & "'));"
    'mySQL
    Qry = "SELECT SPINVOICE.AMOUNT From SPINVOICE WHERE SPINVOICE.SPINVNO='" & InvNo & "';"
    
    With rsInv
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            GetInvAmt = !Amount
        Else
            GetInvAmt = 0
        End If
        .Close
    End With
    
End Function

Private Sub CancelClericalAdj()
    Dim Ans As Integer, MyReason As String, Qry As String, MyRefType As String, MyAcc As String
    Dim ReqID As Long

    If InStr(UCase$(flxSum.TextMatrix(ActRow, 9)), "CANCELLED") > 0 Then
        MsgBox "Clerical Adj. No. " & flxSum.TextMatrix(ActRow, 2) & " was already cancelled!", vbExclamation + vbOKOnly, "Cancel"
        Exit Sub
    End If
    If Val(flxSum.TextMatrix(ActRow, 11)) > 0 Then
        MsgBox "Clerical Adj. No. " & flxSum.TextMatrix(ActRow, 2) & " was already printed and final!", vbExclamation + vbOKOnly, "Cancel"
        Exit Sub
    End If
    
    Ans = MsgBox("Do you want to Cancel Adjustment?" & vbCrLf & vbCrLf & _
        "Adj. Date : " & flxSum.TextMatrix(ActRow, 1) & vbCrLf & "Adjustment No: " & flxSum.TextMatrix(ActRow, 2) & vbCrLf & _
        "Transaction: " & flxSum.TextMatrix(ActRow, 3), vbQuestion + vbYesNo, "Cancel")
    If Ans = vbNo Then Exit Sub
    MyReason = InputBox("Please indicate reason for Cancellation of Adjustment:", "Reason for Cancel")
    If MyReason = "" Then
        MsgBox "Unable to indicate Reason..." & vbCrLf & "Cancellation of Adjustment failed!", vbExclamation + vbOKOnly, "Cancel Failed"
        Exit Sub
    End If
    '...Retrace transactions base on Item Adjusted
    With flxSum
        If .TextMatrix(ActRow, 3) = "PART NUMBER" Then
            '***Access
            'LYRcon.Execute "UPDATE SPAREPARTS SET SPAREPARTS.PARTNUMBER = '" & .TextMatrix(ActRow, 5) & "' " & _
                   "WHERE (((SPAREPARTS.PARTNUMBER)='" & .TextMatrix(ActRow, 6) & "'));"
            'LYRcon.Execute "UPDATE SOURCESP SET SOURCESP.PARTNUMBER = '" & .TextMatrix(ActRow, 5) & "' " & _
                           "WHERE (((SOURCESP.PARTNUMBER)='" & .TextMatrix(ActRow, 6) & "'));"
            'mySQl
            LYRcon.Execute "UPDATE SPAREPARTS SET SPAREPARTS.PARTNUMBER = '" & .TextMatrix(ActRow, 5) & "' " & _
                   "WHERE SPAREPARTS.PARTNUMBER='" & .TextMatrix(ActRow, 6) & "';"
            LYRcon.Execute "UPDATE SOURCESP SET SOURCESP.PARTNUMBER = '" & .TextMatrix(ActRow, 5) & "' " & _
                           "WHERE SOURCESP.PARTNUMBER='" & .TextMatrix(ActRow, 6) & "';"
            
        End If
        If .TextMatrix(ActRow, 3) = "INVOICE" And .TextMatrix(ActRow, 4) = "INVOICE NUMBER" Then
            '***Access
            'Qry = "SELECT SPINVOICE.SPINVNO From SPINVOICE WHERE (((SPINVOICE.SPINVNO)='" & .TextMatrix(ActRow, 6) & "'));"
            'mySQL
            Qry = "SELECT SPINVOICE.SPINVNO From SPINVOICE WHERE SPINVOICE.SPINVNO='" & .TextMatrix(ActRow, 6) & "';"
            
            If CtrlNumExist(Qry) Then
                MsgBox "Unable to Cancel Transaction because Invoice No. " & .TextMatrix(ActRow, 6) & " was already used.", vbExclamation + vbOKOnly, "Cancel Failed"
                Exit Sub
            End If
            'Cancel Invoice Number
            '***Access
            'Qry = "SELECT SPINVOICE.INVTYPE From SPINVOICE WHERE (((SPINVOICE.SPINVNO)='" & .TextMatrix(ActRow, 7) & "'));"
            'mySQL
            Qry = "SELECT SPINVOICE.INVTYPE From SPINVOICE WHERE SPINVOICE.SPINVNO='" & .TextMatrix(ActRow, 7) & "';"
            
            MyRefType = GetFieldValue(Qry)
            eMsg = "Line 170: Updating Invoice Number"
            '***Access
            'LYRcon.Execute "UPDATE SPINVOICE SET SPINVOICE.SPINVNO = '" & .TextMatrix(ActRow, 6) & "', SPINVOICE.TRANSNO = 'SI#-" & .TextMatrix(ActRow, 6) & "' " & _
                           "WHERE (((SPINVOICE.SPINVNO)='" & .TextMatrix(ActRow, 7) & "'));"
            'mySQL
            LYRcon.Execute "UPDATE SPINVOICE SET SPINVOICE.SPINVNO = '" & .TextMatrix(ActRow, 6) & "', SPINVOICE.TRANSNO = 'SI#-" & .TextMatrix(ActRow, 6) & "' " & _
                           "WHERE SPINVOICE.SPINVNO='" & .TextMatrix(ActRow, 7) & "';"
            
            eMsg = "Line 173: Updating Reference of Transaction Register"
            '***Access
            'LYRcon.Execute "UPDATE SPITR SET SPITR.REFERENCE = 'SI#-" & .TextMatrix(ActRow, 6) & "' " & _
                           "WHERE (((SPITR.REFERENCE)='SI#-" & .TextMatrix(ActRow, 7) & "'));"
            'mySQL
            LYRcon.Execute "UPDATE SPITR SET SPITR.REFERENCE = 'SI#-" & .TextMatrix(ActRow, 6) & "' " & _
                           "WHERE SPITR.REFERENCE='SI#-" & .TextMatrix(ActRow, 7) & "';"
            
            If MyRefType = "CASH" Then
                eMsg = "Line 177: Updating Invoice Number on table Payments"
                '***Access
                'LYRcon.Execute "UPDATE PAYMENTS SET PAYMENTS.INVOICENO = '" & .TextMatrix(ActRow, 6) & "' WHERE (((PAYMENTS.INVOICENO)='" & .TextMatrix(ActRow, 7) & "'));"
                'mySQL
                LYRcon.Execute "UPDATE PAYMENTS SET PAYMENTS.INVOICENO = '" & .TextMatrix(ActRow, 6) & "' WHERE PAYMENTS.INVOICENO='" & .TextMatrix(ActRow, 7) & "';"
                
            Else
                eMsg = "Line 180: Updating Invoice Number on table On Account Terms"
                '***Access
                LYRcon.Execute "UPDATE ONACCTERMS SET ONACCTERMS.SPINVNO = '" & .TextMatrix(ActRow, 6) & "' WHERE (((ONACCTERMS.SPINVNO)='" & .TextMatrix(ActRow, 7) & "'));"
                Qry = "SELECT LEDGER.SPINVNO From Ledger WHERE (((LEDGER.SPINVNO)='" & .TextMatrix(ActRow, 7) & "') AND ((LEDGER.JOBID)=0));"
                
                'mySQL
                LYRcon.Execute "UPDATE ONACCTERMS SET ONACCTERMS.SPINVNO = '" & .TextMatrix(ActRow, 6) & "' WHERE ONACCTERMS.SPINVNO='" & .TextMatrix(ActRow, 7) & "';"
                Qry = "SELECT LEDGER.SPINVNO From Ledger WHERE LEDGER.SPINVNO='" & .TextMatrix(ActRow, 7) & "' AND LEDGER.JOBID=0;"
                
                If CtrlNumExist(Qry) Then
                    eMsg = "Line 184: Updating Invoice Number on table Ledger"
                    '***Access
                    'LYRcon.Execute "UPDATE LEDGER SET LEDGER.SPINVNO = '" & .TextMatrix(ActRow, 6) & "' " & _
                            "WHERE (((LEDGER.SPINVNO)='" & .TextMatrix(ActRow, 7) & "') AND ((LEDGER.JOBID)=0));"
                    'mySQL
                    LYRcon.Execute "UPDATE LEDGER SET LEDGER.SPINVNO = '" & .TextMatrix(ActRow, 6) & "' " & _
                            "WHERE LEDGER.SPINVNO='" & .TextMatrix(ActRow, 7) & "' AND LEDGER.JOBID=0;"
                
                End If
            End If
            '***Access
            'Qry = "SELECT SPINVOICE.REQID From SPINVOICE WHERE (((SPINVOICE.SPINVNO)='" & .TextMatrix(ActRow, 6) & "'));"
            
            'mySQl
            Qry = "SELECT SPINVOICE.REQID From SPINVOICE WHERE SPINVOICE.SPINVNO='" & .TextMatrix(ActRow, 6) & "';"
            
            ReqID = Val(GetFieldValue(Qry))
            If ReqID > 0 Then
                eMsg = "Line 199: Updating Invoice Number on table Requisition"
                '***Access
                'LYRcon.Execute "UPDATE REQUISITION SET REQUISITION.INVOICENO = '" & .TextMatrix(ActRow, 6) & "' WHERE (((REQUISITION.REQID)=" & ReqID & "));"
                'mySQL
                LYRcon.Execute "UPDATE REQUISITION SET REQUISITION.INVOICENO = '" & .TextMatrix(ActRow, 6) & "' WHERE REQUISITION.REQID=" & ReqID & ";"
            
            End If
        End If
        If .TextMatrix(ActRow, 3) = "INVOICE" And .TextMatrix(ActRow, 4) = "INVOICE DATE" Then
            '***Access
            'Qry = "SELECT SPINVOICE.SPINVNO From SPINVOICE WHERE (((SPINVOICE.SPINVNO)='" & .TextMatrix(ActRow, 5) & "'));"
            
            'mySQL
            Qry = "SELECT SPINVOICE.SPINVNO From SPINVOICE WHERE SPINVOICE.SPINVNO='" & .TextMatrix(ActRow, 5) & "';"
            
            If Not CtrlNumExist(Qry) Then
                MsgBox "Unable to Cancel Transaction because Invoice No. " & .TextMatrix(ActRow, 6) & " does not exist", vbExclamation + vbOKOnly, "Cancel Failed"
                Exit Sub
            End If
            '***Access
            'Qry = "SELECT SPINVOICE.INVTYPE From SPINVOICE WHERE (((SPINVOICE.SPINVNO)='" & .TextMatrix(ActRow, 5) & "'));"
            'mySQL
            Qry = "SELECT SPINVOICE.INVTYPE From SPINVOICE WHERE SPINVOICE.SPINVNO='" & .TextMatrix(ActRow, 5) & "';"
            
            MyRefType = GetFieldValue(Qry)
            eMsg = "Line 200: Updating Invoice Date"
            '***Access
            'LYRcon.Execute "UPDATE SPINVOICE SET SPINVOICE.INVDATE = #" & .TextMatrix(ActRow, 6) & "# WHERE (((SPINVOICE.SPINVNO)='" & .TextMatrix(ActRow, 5) & "'));"
            'mySQL
            LYRcon.Execute "UPDATE SPINVOICE SET SPINVOICE.INVDATE = '" & Format(.TextMatrix(ActRow, 6), "yyyy-mm-dd") & "' WHERE SPINVOICE.SPINVNO='" & .TextMatrix(ActRow, 5) & "';"
            
            eMsg = "Line 202: Updating Date on Transaction Register"
            '***Access
            'LYRcon.Execute "UPDATE SPITR SET SPITR.TRANSDATE = #" & .TextMatrix(ActRow, 6) & "# WHERE (((SPITR.REFERENCE)='SI#-" & .TextMatrix(ActRow, 5) & "'));"
            'mySQL
            LYRcon.Execute "UPDATE SPITR SET SPITR.TRANSDATE = '" & Format(.TextMatrix(ActRow, 6), "yyyy-mm-dd") & "' WHERE SPITR.REFERENCE='SI#-" & .TextMatrix(ActRow, 5) & "';"
            
            If MyRefType = "CASH" Then
                eMsg = "Line 205: Updating Invoice Date on table Payments"
                '***Access
                'LYRcon.Execute "UPDATE PAYMENTS SET PAYMENTS.PAYDATE= #" & .TextMatrix(ActRow, 6) & "# WHERE (((PAYMENTS.INVOICENO)='" & .TextMatrix(ActRow, 5) & "'));"
                'mySQL
                LYRcon.Execute "UPDATE PAYMENTS SET PAYMENTS.PAYDATE= '" & Format(.TextMatrix(ActRow, 6), "yyyy-mm-dd") & "' WHERE PAYMENTS.INVOICENO='" & .TextMatrix(ActRow, 5) & "';"
                
            Else
                eMsg = "Line 208: Updating Invoice Date on table On Account Terms"
                
                'mySQL
                LYRcon.Execute "UPDATE ONACCTERMS SET ONACCTERMS.TRANSDATE = '" & Format(.TextMatrix(ActRow, 6), "yyyy-mm-dd") & "' WHERE ONACCTERMS.SPINVNO='" & .TextMatrix(ActRow, 5) & "';"
            
            End If
        End If
        '***Access
        'LYRcon.Execute "UPDATE ADJCLERICAL SET ADJCLERICAL.REMARKS = 'CANCELLED' WHERE (((ADJCLERICAL.ADJNO)='" & .TextMatrix(ActRow, 2) & "'));"
        'mySQL
        LYRcon.Execute "UPDATE ADJCLERICAL SET ADJCLERICAL.REMARKS = 'CANCELLED' WHERE ADJCLERICAL.ADJNO='" & .TextMatrix(ActRow, 2) & "';"
        
        'Update Audit Trans
        '***Access
        'LYRcon.Execute "INSERT INTO AUDITRANS ( TRANSDATE, TRANSTYPE, TRANSDOC, REFNO, REFDATE, REASON, EMPID ) " & _
                 "VALUES (#" & Date & "#, 'CANCEL', 'CLERICAL ADJ.', '" & .TextMatrix(ActRow, 2) & "', #" & .TextMatrix(ActRow, 1) & "#, '" & MyReason & "', " & UserID & ");"
        'mySQL
        LYRcon.Execute "INSERT INTO AUDITRANS ( TRANSDATE, TRANSTYPE, TRANSDOC, REFNO, REFDATE, REASON, EMPID ) " & _
                 "VALUES ('" & Format(Date, "yyyy-mm-dd") & "', 'CANCEL', 'CLERICAL ADJ.', '" & .TextMatrix(ActRow, 2) & "', '" & Format(.TextMatrix(ActRow, 1), "yyyy-mm-dd") & "', '" & MyReason & "', " & UserID & ");"
        
        'Update Grid
         .TextMatrix(ActRow, 9) = "CANCELLED"
        'Change row color
        .Col = 1: .Row = ActRow: .ColSel = 11: .RowSel = ActRow: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 0, 0)
        .Col = 2
        MsgBox .TextMatrix(ActRow, 3) & " Adjustment Cancelled!", vbInformation + vbOKOnly, "Cancel"
    End With

End Sub

Private Sub CancelJO()
    Dim Ans As Integer, MyReason As String
    
    If flxSum.TextMatrix(ActRow, 7) <> "" Then
        MsgBox "Job Order already Paid! Cancellation not allowed!", vbExclamation + vbOKOnly, "Validation"
        Exit Sub
    End If
    If InStr(UCase$(flxSum.TextMatrix(ActRow, 13)), "CANCELLED") > 0 Then
        MsgBox "Job Order No. " & flxSum.TextMatrix(ActRow, 2) & " was already cancelled!", vbExclamation + vbOKOnly, "Cancel"
        Exit Sub
    End If
    'If Val(flxSum.TextMatrix(ActRow, 12)) > 0 Then
    '    MsgBox "Job Order No. " & flxSum.TextMatrix(ActRow, 2) & " was already printed and final!", vbExclamation + vbOKOnly, "Cancel"
    '    Exit Sub
    'End If
    
    Ans = MsgBox("Do you want to Cancel Job Order?" & vbCrLf & vbCrLf & _
        "Job Order Date : " & flxSum.TextMatrix(ActRow, 1) & vbCrLf & "Job Order No: " & flxSum.TextMatrix(ActRow, 2) & vbCrLf & _
        "Customer: " & flxSum.TextMatrix(ActRow, 3) & vbCrLf & _
        "JO# Amount: " & flxSum.TextMatrix(ActRow, 6), vbQuestion + vbYesNo, "Cancel")
    If Ans = vbNo Then Exit Sub
    MyReason = InputBox("Please indicate reason for Cancellation of Job Order:", "Reason for Cancel")
    If MyReason = "" Then
        MsgBox "Unable to indicate Reason..." & vbCrLf & "Cancellation of Job Order failed!", vbExclamation + vbOKOnly, "Cancel Failed"
        Exit Sub
    End If
    'Update Job Order
    LYRcon.Execute "UPDATE SRVCJOBORDER SET SRVCJOBORDER.REMARKS = 'CANCELLED', SRVCJOBORDER.TOTALCOST = 0, SRVCJOBORDER.BALANCE = 0 WHERE (((SRVCJOBORDER.JOBORDERID)=" & flxSum.TextMatrix(ActRow, 14) & "));"
    'Update On account
    If flxSum.TextMatrix(ActRow, 5) = "ON ACCOUNT" Then
        LYRcon.Execute "UPDATE ONACCTERMS SET ONACCTERMS.STATUS = 'CANCELLED' WHERE (((ONACCTERMS.JOBORDERID)=" & flxSum.TextMatrix(ActRow, 14) & "));"
        
        'mySQL
        'LYRcon.Execute "UPDATE ONACCTERMS SET ONACCTERMS.STATUS = 'CANCELLED' WHERE ONACCTERMS.JOBORDERID=" & flxSum.TextMatrix(ActRow, 14) & ";"
        
    End If
    'Update Audit Trans
    '***Access
    'LYRcon.Execute "INSERT INTO AUDITRANS ( TRANSDATE, TRANSTYPE, TRANSDOC, REFNO, REFDATE, REASON, EMPID ) " & _
            "VALUES (#" & Date & "#, 'CANCEL', 'JOB ORDER', '" & flxSum.TextMatrix(ActRow, 2) & "', #" & flxSum.TextMatrix(ActRow, 1) & "#, '" & MyReason & "', " & UserID & ");"
    
    'mySQL
    LYRcon.Execute "INSERT INTO AUDITRANS ( TRANSDATE, TRANSTYPE, TRANSDOC, REFNO, REFDATE, REASON, EMPID ) " & _
            "VALUES ('" & Format(Date, "yyyy-mm-dd") & "', 'CANCEL', 'JOB ORDER', '" & flxSum.TextMatrix(ActRow, 2) & "', '" & Format(flxSum.TextMatrix(ActRow, 1), "yyyy-mm-dd") & "', '" & MyReason & "', " & UserID & ");"
    
   'Update Grid
    flxSum.TextMatrix(ActRow, 13) = "CANCELLED"
    flxSum.TextMatrix(ActRow, 6) = "0.00"
    With flxSum     'Change row color
        .Col = 1: .Row = ActRow: .ColSel = 14: .RowSel = ActRow: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 0, 0)
        .Col = 2
    End With
    MsgBox "Job Order Cancelled!", vbInformation + vbOKOnly
    
End Sub

Private Sub mnuMonitorPO_Click()
    Call monitorPurRequest
End Sub
Private Sub mnuDeptSlipValid_click()
    Call CashFlowMgrValidation.ShowRecord(flxSum.TextMatrix(ActRow, 1))
    CashFlowMgrValidation.Show vbModal
End Sub

Public Sub monitorPurRequest()
    Dim rsMonitor As New ADODB.Recordset
    Dim updtStatus As Boolean
    
    On Error GoTo LocalErr
    
    Qry = "SELECT sporderbranch.`STATUS` From sporderbranch " & _
                "Where sporderbranch.REFERENCE = '" & flxSum.TextMatrix(ActRow, 2) & "' AND sporderbranch.ORDERDATE = '" & Format(flxSum.TextMatrix(ActRow, 1), "yyyy-mm-dd") & "';"

    With rsMonitor
        Call ConnectMain
        .Open Qry, MainCon, adOpenKeyset, adLockOptimistic
            If .RecordCount > 0 Then
                .MoveFirst
                txtStatus.Text = IIf(IsNull(!STATUS), "", !STATUS)
                updtStatus = True
            End If
        .Close
    End With
    MainCon.Close
    
    
        If flxSum.TextMatrix(ActRow, 3) = txtStatus.Text Then
            MsgBox "No Update for this transaction", vbCritical + vbOKOnly, "No Update"
        ElseIf txtStatus.Text = "" Then
            MsgBox "No Update for Status", vbCritical + vbOKOnly, "No Update"
        Else
            With flxSum
                LYRcon.Execute "UPDATE sporder SET `STATUS` = '" & txtStatus.Text & "' Where sporder.REFERENCE = '" & .TextMatrix(ActRow, 2) & "' AND sporder.ORDERDATE = '" & Format(.TextMatrix(ActRow, 1), "yyyy-mm-dd") & "';"
                MsgBox "The Status of PO Ref#" + flxSum.TextMatrix(ActRow, 2) + " was updated to " + txtStatus.Text + "!", vbInformation + vbOKOnly, "Status Updated"
            End With
        End If
        Call RefreshRecord(2)
    Exit Sub
LocalErr:
    MsgBox ErrMsg & vbCrLf & Err.Number & " " & Err.Description & vbCrLf & vbCrLf & "Please inform M.I.S. Department", vbCritical + vbOKOnly, "Error"

    
End Sub


Private Sub RXResizer1_ResizeComplete()
    Dim i As Integer, j As Integer
    lblHeading(2).Left = Frame1.Left + (Frame1.Width - lblHeading(2).Width) - 30
    lblHeading(3).Left = Frame1.Left + (Frame1.Width - lblHeading(3).Width)
    
    With flxSum
        If mySrc = "SP Incoming" Or mySrc = "SP Incoming Supplier" Then
            .FormatString = "   |^DATE REC.|<PACKING NO.|<SUPPLIER|<SHIPPER|^PACK DATE|<REMARKS|<CHECKER|<PRINT COPY"
            .ColWidth(0) = 300
            .ColWidth(1) = (.Width - 300) * 0.12
            .ColWidth(2) = (.Width - 300) * 0.12
            .ColWidth(3) = (.Width - 300) * 0.17
            .ColWidth(4) = (.Width - 300) * 0.17
            .ColWidth(5) = (.Width - 300) * 0.12
            .ColWidth(6) = (.Width - 300) * 0.15
            .ColWidth(7) = (.Width - 300) * 0.15
            .ColWidth(8) = (.Width - 300) * 0.15
            .ColWidth(9) = 0
            For i = 1 To .Rows - 1
                If Val(.TextMatrix(i, 8)) = 0 Then 'Not yet printed
                    .Col = 1: .Row = i: .ColSel = 8: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 250, 200)
                    .Col = 2
                End If
                If InStr(.TextMatrix(i, 6), "CANCELLED") > 0 Then 'Cancelled
                    .Col = 1: .Row = i: .ColSel = 9: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 100, 100)
                    .Col = 2
                End If
            Next i

    
        ElseIf mySrc = "Job Order" Then
            
            .FormatString = "   |^ DATE|^JOB ORDER NO.|<CUSTOMER|<ADDRESS|^PAYMENT TYPE |>BALANCE|^O.R. NO.|<MC MODEL|>KM READING|<J.O. TYPE|<MECHANIC|>PRINT COPY|REMARKS"
            .ColWidth(0) = 300
            .ColWidth(1) = (.Width - 300) * 0.1
            .ColWidth(2) = (.Width - 300) * 0.1
            .ColWidth(3) = (.Width - 300) * 0.2
            .ColWidth(4) = (.Width - 300) * 0.2
            .ColWidth(5) = (.Width - 300) * 0.15
            .ColWidth(6) = (.Width - 300) * 0.12
            .ColWidth(7) = (.Width - 300) * 0.12
            .ColWidth(8) = (.Width - 300) * 0.15
            .ColWidth(9) = (.Width - 300) * 0.12
            .ColWidth(10) = (.Width - 300) * 0.2
            .ColWidth(11) = (.Width - 300) * 0.2
            .ColWidth(12) = (.Width - 300) * 0.12
            .ColWidth(13) = (.Width - 300) * 0.15
            .ColWidth(14) = 0
            For i = 1 To .Rows - 1
                .TextMatrix(i, 1) = Format$(.TextMatrix(i, 1), "mm/dd/yyyy")
                .TextMatrix(i, 6) = Format$(.TextMatrix(i, 6), "#,##0.00")
                If Val(.TextMatrix(i, 12)) = 0 Then
                    '.Col = 0: .Row = i: .CellFontName = "Webdings": .CellFontSize = 10: .TextMatrix(i, 0) = lblP.Caption
                    .Col = 1: .Row = i: .ColSel = 13: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 250, 180)
                    .Col = 2
                End If
                If .TextMatrix(i, 13) = "CANCELLED" Then
                    '.Col = 0: .Row = i: .CellFontName = "Webdings": .CellFontSize = 10: .TextMatrix(i, 0) = lblP.Caption
                    .Col = 1: .Row = i: .ColSel = 13: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(255, 0, 0)
                    .Col = 2
                End If
            Next i
        ElseIf mySrc = "Clerical Adjustment" Then
                                '   1       2          3          4       5            6        7        8         9          10
            .FormatString = "   |^ DATE|^ADJ. NO.|<TRANSACTION|<FIELD|<REF. NO.|<OLD VALUE |<NEW VALUE|<REASON |<REMARKS|<ENCODER|<PRINT COPY"
            .ColWidth(0) = 300
            For i = 1 To 10
                .ColWidth(i) = (.Width - 300) * 0.12
            Next i
            .ColWidth(11) = (.Width - 300) * 0.08
            For i = 1 To .Rows - 1
                .TextMatrix(i, 1) = Format$(.TextMatrix(i, 1), "mm/dd/yyyy")
                If Val(.TextMatrix(i, 11)) = 0 Then
                    .Col = 1: .Row = i: .ColSel = 11: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 250, 180)
                    .Col = 2
                End If
                If .TextMatrix(i, 9) = "CANCELLED" Then
                    .Col = 1: .Row = i: .ColSel = 11: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(255, 0, 0)
                    .Col = 2
                End If
            Next i
        ElseIf mySrc = "DMCM [A/R]" Then 'ADJDMCM.ADJDATE, ADJDMCM.ADJNO, ADJDMCM.CUSTOMER, ADJDMCM.TRANSTYPE, ADJDMCM.REFNO, ADJDMCM.DEBIT, ADJDMCM.CREDIT, ADJDMCM.BALANCE, ADJDMCM.STATUS, ADJDMCM.REASON, ADJDMCM.PRINTCOPY
                            '       1       2           3       4         5         6       7        8       9        10      11      12
            .FormatString = "   |^ DATE|^ADJ. NO.|<CUSTOMER|<TRANSACTION|^REF. NO.|>DEBIT |>CREDIT|>BALANCE|<STATUS|<REASON|<PRN COPY"
            .ColWidth(0) = 300
            .ColWidth(1) = (.Width - 300) * 0.1
            .ColWidth(2) = (.Width - 300) * 0.1
            .ColWidth(3) = (.Width - 300) * 0.2
            .ColWidth(4) = (.Width - 300) * 0.1
            .ColWidth(5) = (.Width - 300) * 0.1
            .ColWidth(6) = (.Width - 300) * 0.1
            .ColWidth(7) = (.Width - 300) * 0.12
            .ColWidth(8) = (.Width - 300) * 0.12
            .ColWidth(9) = (.Width - 300) * 0.12
            .ColWidth(10) = (.Width - 300) * 0.2
            .ColWidth(11) = (.Width - 300) * 0.08
            '.ColWidth(12) = (.Width - 300) * 0.12
            
            For i = 1 To .Rows - 1
                .TextMatrix(i, 1) = Format$(.TextMatrix(i, 1), "mm/dd/yyyy")
                '.TextMatrix(i, 6) = Format$(.TextMatrix(i, 6), "mm/dd/yyyy")
                .TextMatrix(i, 7) = Format$(.TextMatrix(i, 7), "#,##0.00")
                .TextMatrix(i, 8) = Format$(.TextMatrix(i, 8), "#,##0.00")
                .TextMatrix(i, 6) = Format$(.TextMatrix(i, 6), "#,##0.00")
                If Val(.TextMatrix(i, 11)) = 0 Then
                    '.Col = 0: .Row = i: .CellFontName = "Webdings": .CellFontSize = 10: .TextMatrix(i, 0) = lblP.Caption
                    .Col = 1: .Row = i: .ColSel = 11: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 250, 180)
                    .Col = 2
                End If
                If .TextMatrix(i, 10) = "CANCELLED" Then
                    '.Col = 0: .Row = i: .CellFontName = "Webdings": .CellFontSize = 10: .TextMatrix(i, 0) = lblP.Caption
                    .Col = 1: .Row = i: .ColSel = 11: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(255, 0, 0)
                    .Col = 2
                End If
            Next i
                
        ElseIf mySrc = "Inventory Adjustment" Then
                             '       1       2           3       4       5           6           7           8
            .FormatString = "   |^ DATE|^ADJ. NO.|<REFERENCE|<REASON|<REMARKS |<PREPARED BY|<APPROVED BY|<PRINT COPY"
            .ColWidth(0) = 300
            .ColWidth(1) = (.Width - 300) * 0.12
            .ColWidth(2) = (.Width - 300) * 0.12
            .ColWidth(3) = (.Width - 300) * 0.12
            .ColWidth(4) = (.Width - 300) * 0.2
            .ColWidth(5) = (.Width - 300) * 0.2
            .ColWidth(6) = (.Width - 300) * 0.2
            .ColWidth(7) = (.Width - 300) * 0.2
            .ColWidth(8) = (.Width - 300) * 0.15
            For i = 1 To .Rows - 1
                .TextMatrix(i, 1) = Format$(.TextMatrix(i, 1), "mm/dd/yyyy")
                If Val(.TextMatrix(i, 8)) = 0 Then
                    '.Col = 0: .Row = i: .CellFontName = "Webdings": .CellFontSize = 10: .TextMatrix(i, 0) = lblP.Caption
                    .Col = 1: .Row = i: .ColSel = 8: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 250, 180)
                    .Col = 2
                End If
                If .TextMatrix(i, 5) = "CANCELLED" Then
                    '.Col = 0: .Row = i: .CellFontName = "Webdings": .CellFontSize = 10: .TextMatrix(i, 0) = lblP.Caption
                    .Col = 1: .Row = i: .ColSel = 8: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(255, 0, 0)
                    .Col = 2
                End If
            Next i
        ElseIf mySrc = "Top Selling Spare Parts" Then
            .FormatString = "   |^DATE CREATED|<SALES PERIOD|<ITEM TYPE|<NO. OF SPAREPARTS|<NOTE "
            .ColWidth(0) = 300
            .ColWidth(1) = .Width * 0.12
            .ColWidth(2) = .Width * 0.21
            .ColWidth(3) = .Width * 0.21
            .ColWidth(4) = .Width * 0.21
            .ColWidth(5) = .Width * 0.21
            .ColWidth(6) = 0
            For i = 1 To .Rows - 1
                .TextMatrix(i, 1) = Format$(.TextMatrix(i, 1), "mm/dd/yyyy")
                If IsNumeric(.TextMatrix(i, 2)) Then    'Identify Sales Period
                    If Val(.TextMatrix(i, 2)) = 0 Then
                        .TextMatrix(i, 2) = "FROM BEGINNING"
                    Else
                        j = Val(.TextMatrix(i, 2))
                        .TextMatrix(i, 2) = UCase$(Format$(DateAdd("M", j * -1, DateValue(.TextMatrix(i, 1))), "MMM. 'YY")) & " TO " & UCase$(Format$(DateAdd("M", -1, DateValue(.TextMatrix(i, 1))), "MMM. 'YY"))
                    End If
                End If
                If IsNumeric(.TextMatrix(i, 4)) Then   'Identify Top Level
                    .TextMatrix(i, 4) = "TOP " & .TextMatrix(i, 4)
                End If
            Next i
        ElseIf mySrc = "Customer P.O." Then
            .FormatString = "   |^P.O. DATE|^P.O. NO.|<CUSTOMER NAME|>STATUS|^INVOICE NO.|^REMARKS|^PRINT "
            .ColWidth(0) = 300
            .ColWidth(1) = .Width * 0.11
            .ColWidth(2) = .Width * 0.11
            .ColWidth(3) = .Width * 0.2
            .ColWidth(4) = .Width * 0.15
            .ColWidth(5) = .Width * 0.11
            .ColWidth(6) = .Width * 0.2
            .ColWidth(7) = .Width * 0.11
            .ColWidth(8) = 0
            For i = 1 To .Rows - 1
                .TextMatrix(i, 1) = Format$(.TextMatrix(i, 1), "mm/dd/yyyy")
                If .TextMatrix(i, 4) = "PENDING" Then
                    .Col = 1: .Row = i: .ColSel = 8: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 250, 200)
                    .Col = 2
                End If
                If Val(.TextMatrix(i, 7)) = 0 And .TextMatrix(i, 4) = "APPROVED" Then 'Not yet printed
                    .Col = 1: .Row = i: .ColSel = 8: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(0, 230, 0)
                    .Col = 2
                End If
                If InStr(.TextMatrix(i, 6), "CANCELLED") > 0 Then 'Cancelled
                    .Col = 1: .Row = i: .ColSel = 8: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 150, 150)
                    .Col = 2
                End If
            Next i
        
        ElseIf mySrc = "Purchase Return" Then
            .FormatString = "   |^RET. DATE|^RETURN NO.|^PACKING NO.|^PACK DATE|<SUPPLIER|>AMOUNT|<REMARKS |^PRINTED"
            .ColWidth(0) = 300
            .ColWidth(1) = (.Width - 300) * 0.1
            .ColWidth(2) = (.Width - 300) * 0.1
            .ColWidth(3) = (.Width - 300) * 0.1
            .ColWidth(4) = (.Width - 300) * 0.1
            .ColWidth(5) = (.Width - 300) * 0.2
            .ColWidth(6) = (.Width - 300) * 0.1
            .ColWidth(7) = (.Width - 300) * 0.2
            .ColWidth(8) = (.Width - 300) * 0.07
            .ColWidth(9) = 0
            For i = 1 To .Rows - 1
                .TextMatrix(i, 1) = Format$(.TextMatrix(i, 1), "mm/dd/yyyy")
                .TextMatrix(i, 4) = Format$(.TextMatrix(i, 4), "mm/dd/yyyy")
                .TextMatrix(i, 6) = Format$(.TextMatrix(i, 6), "#,##0.00")
                If Val(.TextMatrix(i, 8)) = 0 Then  'PRINT COPY
                    .Col = 1: .Row = i: .ColSel = 8: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 250, 180)
                    .Col = 2
                End If
                If .TextMatrix(i, 7) = "CANCELLED" Then 'REMARKS
                    .Col = 1: .Row = i: .ColSel = 8: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(240, 0, 0)
                    .Col = 2
                End If
            Next i
        ElseIf mySrc = "Requisition" Then
            .FormatString = "   |^REQ. DATE|^REQ. NO.|<CUSTOMER NAME|>REQ. AMOUNT|^JOB ORDER NO.|^INVOICE NO.|^O.R. NO.|<REMARKS|^PRN "
            .ColWidth(0) = 300
            .ColWidth(1) = .Width * 0.12
            .ColWidth(2) = .Width * 0.12
            .ColWidth(3) = .Width * 0.2
            .ColWidth(4) = .Width * 0.12
            .ColWidth(5) = .Width * 0.12
            .ColWidth(6) = .Width * 0.12
            .ColWidth(7) = .Width * 0.12
            .ColWidth(8) = .Width * 0.15
            .ColWidth(9) = .Width * 0.1
            For i = 1 To .Rows - 1
                .TextMatrix(i, 1) = Format$(.TextMatrix(i, 1), "mm/dd/yyyy")
                .TextMatrix(i, 4) = Format$(.TextMatrix(i, 4), "###,##0.00")
                If Val(.TextMatrix(i, 9)) = 0 Then 'Not yet printed
                    .Col = 1: .Row = i: .ColSel = 9: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 250, 200)
                    .Col = 2
                End If
                If InStr(.TextMatrix(i, 8), "CANCELLED") > 0 Then 'Cancelled
                    .Col = 1: .Row = i: .ColSel = 9: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 100, 100)
                    .Col = 2
                End If
            Next i
        ElseIf mySrc = "Sales Return" Then
            .FormatString = "   |^DATE RETURNED|^SALES RET.#|^INVOICE#|^INVOICE DATE|<NAME OF CUSTOMER|^SALES RET.THIS MONTH|^SALES RET.-PRIOR MONTH SALES|>TOTAL|<REMARKS |^PRINTED"
            .ColWidth(0) = 300
            .ColWidth(1) = (.Width - 300) * 0.15  'Date return
            .ColWidth(2) = (.Width - 300) * 0.12  'Sales RET no
            .ColWidth(3) = (.Width - 300) * 0.12  'Invoice No
            .ColWidth(4) = (.Width - 300) * 0.12  'Invoice date
            .ColWidth(5) = (.Width - 300) * 0.25  'Name of Customer
            .ColWidth(6) = (.Width - 300) * 0.2  'Sales Return This Month
            .ColWidth(7) = (.Width - 300) * 0.23  'Sales Return Prior Month Sales
            .ColWidth(8) = (.Width - 300) * 0.06  'Total
            .ColWidth(9) = (.Width - 300) * 0.26  'REMARKS
            .ColWidth(10) = (.Width - 300) * 0.08 'Printed
            .ColWidth(11) = 0
            For i = 1 To .Rows - 1
                .TextMatrix(i, 1) = Format$(.TextMatrix(i, 1), "mm/dd/yyyy")
                .TextMatrix(i, 4) = Format$(.TextMatrix(i, 4), "mm/dd/yyyy")
                .TextMatrix(i, 6) = Format$(.TextMatrix(i, 6), "#,##0.00")
                .TextMatrix(i, 7) = Format$(.TextMatrix(i, 7), "#,##0.00")
                .TextMatrix(i, 8) = Format$(.TextMatrix(i, 8), "#,##0.00")
                If Val(.TextMatrix(i, 10)) = 0 Then  'PRINT COPY
                    .Col = 1: .Row = i: .ColSel = 10: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 250, 180)
                    .Col = 2
                End If
                If .TextMatrix(i, 9) = "CANCELLED" Then 'REMARKS
                    .Col = 1: .Row = i: .ColSel = 10: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(240, 0, 0)
                    .Col = 2
                End If
            Next i
        ElseIf mySrc = "Location Adjustment" Then
            .FormatString = "   |^ DATE|^CONTROL NO.|<REFERENCE|<REASON|<PREPARED BY|<APPROVED BY|^PRINT COPY"
            .ColWidth(0) = 300
            .ColWidth(1) = (.Width - 300) * 0.12
            .ColWidth(2) = (.Width - 300) * 0.12
            .ColWidth(3) = (.Width - 300) * 0.2
            .ColWidth(4) = (.Width - 300) * 0.4
            .ColWidth(5) = (.Width - 300) * 0.15
            .ColWidth(6) = (.Width - 300) * 0.15
            .ColWidth(7) = (.Width - 300) * 0.1
            .ColWidth(8) = 0
            For i = 1 To .Rows - 1
                .TextMatrix(i, 2) = Format$(.TextMatrix(i, 2), "0000")
                .TextMatrix(i, 1) = Format$(.TextMatrix(i, 1), "MM/DD/YYYY")
                If Val(.TextMatrix(i, 7)) = 0 Then
                    '.Col = 0: .Row = i: .CellFontName = "Webdings": .CellFontSize = 10: .TextMatrix(i, 0) = lblP.Caption
                    .Col = 1: .Row = i: .ColSel = .Cols - 1: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 250, 180)
                    .Col = 2
                End If
                If .TextMatrix(i, 3) = "CANCELLED" Then
                    .Col = 1: .Row = i: .ColSel = 8: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(255, 0, 0)
                    .Col = 2
                End If
            Next i
        ElseIf mySrc = "Received Spare Parts" Or mySrc = "Received Spare Parts DVOWH" Then
                                      '1          2            3         4        5             6           7       8        9
            .FormatString = "   |^DATE REC.|<PACKING NO.|<SUPPLIER|<SHIPPER|<P.O. REF.#|^PACK DATE|<REMARKS|<TOTALPN|<TOTALQTY|<DAMAGE|<LACKING|<EXCESS|<WRONGPN|<CHECKER|<PRINTCOPY"
            .ColWidth(0) = 300
            .ColWidth(1) = (.Width - 300) * 0.12    'DATE REC
            .ColWidth(2) = (.Width - 300) * 0.12    'PACKING NO
            .ColWidth(3) = (.Width - 300) * 0.17    'SUPPLIER
            .ColWidth(4) = (.Width - 300) * 0.17    'SHIPPER
            .ColWidth(5) = (.Width - 300) * 0.12    'POREF
            .ColWidth(6) = (.Width - 300) * 0.12    'PACKDATE
            .ColWidth(7) = (.Width - 300) * 0.15    'REMARKS
            .ColWidth(8) = (.Width - 300) * 0.1    'TOTAL PN
            .ColWidth(9) = (.Width - 300) * 0.1    'TOTAL QTY
            .ColWidth(10) = (.Width - 300) * 0.1   'DAMAGE
            .ColWidth(11) = (.Width - 300) * 0.1   'LACKING
            .ColWidth(12) = (.Width - 300) * 0.1   'EXCESS
            .ColWidth(13) = (.Width - 300) * 0.1   'WRONGPN
            .ColWidth(14) = (.Width - 300) * 0.17   'CHECKER
            .ColWidth(15) = (.Width - 300) * 0.1   'PRINT COPY
            .ColWidth(16) = 0                       'SPRRNO
            .ColWidth(17) = 0                       'AUTOLOAD
            For i = 1 To .Rows - 1
                .TextMatrix(i, 1) = Format$(.TextMatrix(i, 1), "mm/dd/yyyy")
                .TextMatrix(i, 6) = Format$(.TextMatrix(i, 6), "mm/dd/yyyy")
                If Val(.TextMatrix(i, 15)) = 0 Then 'Not yet printed
                    .Col = 1: .Row = i: .ColSel = 9: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 250, 200)
                    .Col = 2
                End If
                If InStr(.TextMatrix(i, 7), "CANCELLED") > 0 Then 'Cancelled
                    .Col = 1: .Row = i: .ColSel = 10: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 100, 100)
                    .Col = 2
                End If
            Next i
        ElseIf mySrc = "Advances for Liquidation" Then
            .FormatString = "  |^AFL Date |^AFL No. |>Amt. Rec.  |>Amt. Ret. |>AFL Amount |<Prepared By |<Remarks |^Prn Copy|"
            .ColWidth(0) = 300
            .ColWidth(1) = (.Width - 400) * 0.1
            .ColWidth(2) = (.Width - 400) * 0.1
            .ColWidth(3) = (.Width - 400) * 0.1
            .ColWidth(4) = (.Width - 400) * 0.1
            .ColWidth(5) = (.Width - 400) * 0.1
            .ColWidth(6) = (.Width - 400) * 0.2
            .ColWidth(7) = (.Width - 400) * 0.2
            .ColWidth(8) = (.Width - 400) * 0.08
            .ColWidth(9) = 0
            For i = 1 To .Rows - 1
                If InStr(UCase$(.TextMatrix(i, 7)), "CANCELLED") > 0 Then 'Cancelled
                    .Col = 1: .Row = i: .ColSel = TCol: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 100, 100)
                    .Col = 2
                End If
                If Val(.TextMatrix(i, 8)) = 0 Then
                    .Col = 1: .Row = i: .ColSel = TCol: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 250, 200)
                    .Col = 2
                End If
            Next i
            
        ElseIf mySrc = "Sales Order" Then
            .FormatString = "  |^INV. DATE|^INVOICE NO.|<INV. TYPE|<CUSTOMER|> INV. AMOUNT|<REMARKS|^ORDERNO|<FRONTLINER"
            .ColWidth(0) = 300
            .ColWidth(1) = (.Width - 300) * 0.1
            .ColWidth(2) = (.Width - 300) * 0.1
            .ColWidth(3) = (.Width - 300) * 0.1
            .ColWidth(4) = (.Width - 300) * 0.15
            .ColWidth(5) = (.Width - 300) * 0.1
            .ColWidth(6) = (.Width - 300) * 0.15
            .ColWidth(7) = (.Width - 300) * 0.1
            .ColWidth(8) = (.Width - 300) * 0.15
            
        ElseIf mySrc = "Mechanic Income" Then
            .FormatString = "   |<MECHANIC NAME|^J.O. NO.|^J.O. DATE|<JOB ORDER TYPE|>J.O. AMOUNT|>FSC AMOUNT|>TOTAL "
            .ColWidth(0) = 300
            .ColWidth(1) = .Width * 0.2
            .ColWidth(2) = .Width * 0.12
            .ColWidth(3) = .Width * 0.12
            .ColWidth(4) = .Width * 0.17
            .ColWidth(5) = .Width * 0.11
            .ColWidth(6) = .Width * 0.11
            .ColWidth(7) = .Width * 0.11
            For i = 1 To .Rows - 1
                .TextMatrix(i, 3) = Format$(.TextMatrix(i, 3), "mm/dd/yyyy")
                .TextMatrix(i, 5) = Format$(.TextMatrix(i, 5), "#,##0.00")
                If InStr(.TextMatrix(i, 4), "FSC") > 0 Then .TextMatrix(i, 6) = txtFSCamt.Text
                .TextMatrix(i, 7) = Format(Val(Replace(.TextMatrix(i, 5), ",", "")) + Val(Replace(.TextMatrix(i, 6), ",", "")), "#,##0.00")
            Next i
        ElseIf mySrc = "Cashier Cash Flow" Then
            .FormatString = "  |^Cash Flow Date |<Name of Cashier |< Approved By  |<Remarks "
            .ColWidth(0) = 300
            .ColWidth(1) = (.Width - 400) * 0.15
            .ColWidth(2) = (.Width - 400) * 0.25
            .ColWidth(3) = (.Width - 400) * 0.25
            .ColWidth(4) = (.Width - 400) * 0.3
            .ColWidth(5) = 0
        
        ElseIf mySrc = "Manager's Cash Flow" Then
            .FormatString = "  |^Cash Flow Date |<Prepared By |< Approved By  |<Remarks | |^Total Validated"
            .ColWidth(0) = 300
            .ColWidth(1) = (.Width - 400) * 0.15
            .ColWidth(2) = (.Width - 400) * 0.25
            .ColWidth(3) = (.Width - 400) * 0.25
            .ColWidth(4) = (.Width - 400) * 0.25
            .ColWidth(5) = 0
            .ColWidth(6) = (.Width - 400) * 0.08
            For i = 1 To .Rows - 1
                'If Len(.TextMatrix(i, 2)) < 4 Then .TextMatrix(i, 2) = Format(.TextMatrix(i, 2), "0000")
'                If InStr(UCase$(.TextMatrix(i, 6)), "CANCELLED") > 0 Then 'Cancelled
'                    .Col = 1: .Row = i: .ColSel = TCol: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 100, 100)
'                    .Col = 2
'                End If
                If Val(.TextMatrix(i, 6)) = 0 Then
                    '.Col = 0: .Row = i: .CellFontName = "Webdings": .CellFontSize = 10 ': .Text = lblP.Caption
                    .Col = 1: .Row = i: .ColSel = 6: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 250, 200)
                    .Col = 2
                End If
            Next i
        ElseIf mySrc = "Cash Return" Then
            .FormatString = "  |^  CR Date |^Cash Return# |>Amount  |<Remarks |Prepared By |^Prn Copy|"
            .ColWidth(0) = 300
            .ColWidth(1) = (.Width - 400) * 0.12
            .ColWidth(2) = (.Width - 400) * 0.12
            .ColWidth(3) = (.Width - 400) * 0.15
            .ColWidth(4) = (.Width - 400) * 0.25
            .ColWidth(5) = (.Width - 400) * 0.25
            .ColWidth(6) = (.Width - 400) * 0.07
            .ColWidth(7) = 0
            For i = 1 To .Rows - 1
                'If Len(.TextMatrix(i, 2)) < 4 Then .TextMatrix(i, 2) = Format(.TextMatrix(i, 2), "0000")
                If InStr(UCase$(.TextMatrix(i, 4)), "CANCELLED") > 0 Then 'Cancelled
                    .Col = 1: .Row = i: .ColSel = TCol: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 100, 100)
                    .Col = 2
                End If
                If Val(.TextMatrix(i, 6)) = 0 Then
                    '.Col = 0: .Row = i: .CellFontName = "Webdings": .CellFontSize = 10 ': .Text = lblP.Caption
                    .Col = 1: .Row = i: .ColSel = TCol: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 250, 200)
                    .Col = 2
                End If
            Next i

        ElseIf mySrc = "Cash Voucher" Then
            .FormatString = "   |^C.V. DATE|^C.V. NO.|<PAYEE|<DESIGNATION|<PARTICULARS|>C.V. AMOUNT|<REMARKS |^PRN COPY"
            .ColWidth(0) = 300
            .ColWidth(1) = .Width * 0.1
            .ColWidth(2) = .Width * 0.1
            .ColWidth(3) = .Width * 0.15
            .ColWidth(4) = .Width * 0.15
            .ColWidth(5) = .Width * 0.15
            .ColWidth(6) = .Width * 0.1
            .ColWidth(7) = .Width * 0.15
            .ColWidth(8) = .Width * 0.05
            For i = 1 To .Rows - 1
                .TextMatrix(i, 1) = Format$(.TextMatrix(i, 1), "mm/dd/yyyy")
                '.TextMatrix(i, 2) = .TextMatrix(i, 2)
                .TextMatrix(i, 6) = Format$(.TextMatrix(i, 6), "#,##0.00")
                If Val(.TextMatrix(i, 8)) = 0 Then 'Not yet printed
                    .Col = 1: .Row = i: .ColSel = TCol - 1: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 250, 200)
                    .Col = 2
                End If
            Next i
        ElseIf mySrc = "Trust Receipt Agreement" Then
            .FormatString = "  |^INV. DATE|^REF. NO. |<NAME OF CUSTOMER |<ADDRESS |^REF. TYPE|>TRA AMOUNT |^CUSTOMER TYPE|<REMARKS |^AGENT |^SPID "
            .ColWidth(0) = 300
            .ColWidth(1) = (.Width - 300) * 0.08
            .ColWidth(2) = (.Width - 300) * 0.12
            .ColWidth(3) = (.Width - 300) * 0.2
            .ColWidth(4) = (.Width - 300) * 0.2
            .ColWidth(5) = (.Width - 300) * 0.09
            .ColWidth(6) = (.Width - 300) * 0.09
            .ColWidth(7) = (.Width - 300) * 0.09
            .ColWidth(8) = (.Width - 300) * 0.25
            .ColWidth(9) = (.Width - 300) * 0.2
            .ColWidth(10) = 0
            .ColWidth(11) = 0
            
            For i = 1 To .Rows - 1
                '.TextMatrix(i, 1) = Format$(.TextMatrix(i, 1), "mm/dd/yyyy")
                '.TextMatrix(i, 2) = Format$(Val(.TextMatrix(i, 2)), "0000")
                .TextMatrix(i, 6) = Format$(.TextMatrix(i, 6), "#,##0.00")
                If Val(.TextMatrix(i, 11)) = 0 Then 'Not yet printed
                    .Col = 1: .Row = i: .ColSel = TCol - 1: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 250, 200)
                    .Col = 2
                End If
            Next i
        ElseIf mySrc = "Branch Order" Then
            '1-ORDERDATE, 2-REFERENCE, 3-SUPPLIERNAME, 4-ORDERTYPE,5-CONFIRMNO,6-SUBMITTEDBY,7-APPROVEDBY,8-PRINTCOPY,9-QTYORDER, 10-QTYREC,11-BACKORDER,12-BOLASTUPDATE "
            '.FormatString = "   |^ORDER DATE|^RR NO.|<SUPPLIER NAME|^ORDER TYPE|^CONFIRM NO.|<SUBMITTED BY|<APPROVED BY|^PRN COPY"
            '                 0       1         2       3                   4          5         6       7     8            9        10         11       12        13
            .FormatString = "   |^ORDER DATE|^RR NO.|^WAREHOUSE STATUS|^CONFIRM NO.|^TOTAL PN|^REC PN|^BO PN|^TOTAL QTY.|^REC QTY|^BO QTY.|^LEAD TIME|<PICKER|<PRINT COPY"
            .ColWidth(0) = 300
            .ColWidth(1) = (.Width - 300) * 0.08
            .ColWidth(2) = (.Width - 300) * 0.08
            .ColWidth(3) = (.Width - 300) * 0.1
            .ColWidth(4) = (.Width - 300) * 0.08
            .ColWidth(5) = (.Width - 300) * 0.05
            .ColWidth(6) = (.Width - 300) * 0.05
            .ColWidth(7) = (.Width - 300) * 0.05
            .ColWidth(8) = (.Width - 300) * 0.06
            .ColWidth(9) = (.Width - 300) * 0.05
            .ColWidth(10) = (.Width - 300) * 0.05
            .ColWidth(11) = (.Width - 300) * 0.13
            .ColWidth(12) = (.Width - 300) * 0.14
            .ColWidth(13) = (.Width - 300) * 0.06
            For i = 1 To .Rows - 1
                '.TextMatrix(i, 1) = Format$(.TextMatrix(i, 1), "mm/dd/yyyy")
                '.TextMatrix(i, 6) = Format$(.TextMatrix(i, 6), "mm/dd/yyyy")
                '.TextMatrix(i, 9) = Format$(.TextMatrix(i, 9), "mm/dd/yyyy")
                If .TextMatrix(i, 13) = "0" Then
                    '.Col = 0: .Row = i: .CellFontName = "Webdings": .CellFontSize = 10: .TextMatrix(i, 0) = lblP.Caption
                    .Col = 1: .Row = i: .ColSel = 13: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 250, 180)
                    .Col = 2
                End If
                If .TextMatrix(i, 3) = "CANCELLED" Then
                    '.Col = 0: .Row = i: .CellFontName = "Webdings": .CellFontSize = 10: .TextMatrix(i, 0) = lblP.Caption
                    .Col = 1: .Row = i: .ColSel = 13: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(255, 0, 0)
                    .Col = 2
                End If
                If Val(.TextMatrix(i, 11)) <> 0 Then
                   .TextMatrix(i, 11) = IIf(Val(.TextMatrix(i, 11)) >= 2, Val(.TextMatrix(i, 11)) & " Days", Val(.TextMatrix(i, 11)) & " Day")
                End If
            Next i
            

        ElseIf mySrc = "Branch Purchase Order" Then
            '               0      1          2               3              4         5          6         7                8     9  10    11
            .FormatString = "   |^SRV NO. |^BRANCH NAME |^ P.O REFERENCE |^DATE SRV|^DATE CONF.|^STATUS|^TRANSFER#|^TOTAL AMOUNT|^BID|^PICK^FQT"
            .ColWidth(0) = 300
            .ColWidth(1) = (.Width - 300) * 0.07
            .ColWidth(2) = (.Width - 300) * 0.13
            .ColWidth(3) = (.Width - 300) * 0.2
            .ColWidth(4) = (.Width - 300) * 0.09
            .ColWidth(5) = (.Width - 300) * 0.09
            .ColWidth(6) = (.Width - 300) * 0.09
            .ColWidth(7) = (.Width - 300) * 0.16
            .ColWidth(8) = (.Width - 300) * 0.16
            .ColWidth(9) = 0
            .ColWidth(10) = 0
            .ColWidth(11) = 0
            .ColWidth(12) = 0
          
            For i = 1 To .Rows - 1
                 If Val(.TextMatrix(i, 10)) = 0 Then 'NoT Printed for Picking
                      .Col = 1: .Row = i: .ColSel = 11: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 250, 200)
                      .Col = 2
                      .TextMatrix(i, 6) = "PENDING"
                End If
                If Val(.TextMatrix(i, 11)) = 1 Then 'Final for Older Taking
                      .Col = 1: .Row = i: .ColSel = 11: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 100, 300)
                      .Col = 2
                      .TextMatrix(i, 6) = "PENDING"
                End If
                If Val(.TextMatrix(i, 11)) = 1 Then 'Final for Older Taking
                      .Col = 1: .Row = i: .ColSel = 11: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = &H80000005
                      .Col = 2
              
                      .TextMatrix(i, 6) = "FINAL"
                End If
                If Val(.TextMatrix(i, 11)) = 2 Then 'Not Final for Older Taking
                      .Col = 1: .Row = i: .ColSel = 11: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 100, 100)
                      .Col = 2
                      .TextMatrix(i, 6) = "CANCELLED"
                End If
                If .TextMatrix(i, 1) = "" And .Rows = 2 Then
                      .Col = 1: .Row = i: .ColSel = 11: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = &H80000005
                      .Col = 2
                     .TextMatrix(i, 6) = ""
                End If
                .TextMatrix(i, 8) = Format(.TextMatrix(i, 8), "###,###,##0.00")
            Next i
'************************ CODE UPDATED: 08/04/17 ******************************
        ElseIf mySrc = "SETTLEMENT DISCREPANCY" Then
            '               0      1          2           3            4              5               6          7         8         9           10          11
            .FormatString = "   |^REF NO. |^DATE REF |^BRANCHNAME |^TOTAL DAMAGE |^TOTAL LACKING |^TOTAL EXCESS |^RR NO. |^DATE REC |^STATUS |^PREPARED BY |^APPROVED BY"
            .ColWidth(0) = 300
            .ColWidth(1) = (.Width - 300) * 0.1   'REF NO
            .ColWidth(2) = (.Width - 300) * 0.1   'DATE REF
            .ColWidth(3) = (.Width - 300) * 0.2    'BRANCHNAME
            .ColWidth(4) = (.Width - 300) * 0.1   'TOTAL DAMAGE
            .ColWidth(5) = (.Width - 300) * 0.1   'TOTAL LACKING
            .ColWidth(6) = (.Width - 300) * 0.1   'TOTAL EXCESS
            .ColWidth(7) = (.Width - 300) * 0.1   'RR No.
            .ColWidth(8) = (.Width - 300) * 0.1   'DATE REC
            .ColWidth(9) = (.Width - 300) * 0.1   'STATUS
            .ColWidth(10) = (.Width - 300) * 0.2  'PREPARED BY
            .ColWidth(11) = (.Width - 300) * 0.2  'APPROVED BY
            .ColWidth(12) = 0
          
            For i = 1 To .Rows - 1
                 If .TextMatrix(i, 9) = "PENDING" Then 'NoT Printed for TRANSFER
                      .Col = 1: .Row = i: .ColSel = 11: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 250, 200)
                      .Col = 2
                End If
            Next i
        ElseIf mySrc = "SDiscrepancy Adjustment" Then
                             '       1       2           3       4       5           6           7           8
            .FormatString = "   |^ DATE|^ADJ. NO.|<REFERENCE|<REASON|<REMARKS |<PREPARED BY|<APPROVED BY|<PRINT COPY"
            .ColWidth(0) = 300
            .ColWidth(1) = (.Width - 300) * 0.12
            .ColWidth(2) = (.Width - 300) * 0.12
            .ColWidth(3) = (.Width - 300) * 0.12
            .ColWidth(4) = (.Width - 300) * 0.2
            .ColWidth(5) = (.Width - 300) * 0.2
            .ColWidth(6) = (.Width - 300) * 0.2
            .ColWidth(7) = (.Width - 300) * 0.2
            .ColWidth(8) = (.Width - 300) * 0.15
            For i = 1 To .Rows - 1
                .TextMatrix(i, 1) = Format$(.TextMatrix(i, 1), "mm/dd/yyyy")
                If Val(.TextMatrix(i, 8)) = 0 Then
                    '.Col = 0: .Row = i: .CellFontName = "Webdings": .CellFontSize = 10: .TextMatrix(i, 0) = lblP.Caption
                    .Col = 1: .Row = i: .ColSel = 8: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 250, 180)
                    .Col = 2
                End If
                If .TextMatrix(i, 5) = "CANCELLED" Then
                    '.Col = 0: .Row = i: .CellFontName = "Webdings": .CellFontSize = 10: .TextMatrix(i, 0) = lblP.Caption
                    .Col = 1: .Row = i: .ColSel = 8: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(255, 0, 0)
                    .Col = 2
                End If
            Next i
'******************** END CODE UPDATED: 08/04/17 ******************************
        ElseIf mySrc = "Daily Cash Deposit Report" Then
                             '       1       2                          3                   4       5                   6               7            8          9           10                  11                      12                      13              14              15          16                  17                  18             19           20      21      22          23              24                  25                          26              27              28              29              30                  31
            .FormatString = "  |^DATE |>SPAREPARTS (WALK-IN) |>SPAREPARTS (WHOLESALE) |>SERVICE |>OTHER INCOME |>TOTAL CASH SALES |>DOWNPAYMENT |>SPAREPARTS |>SERVICE |>OTHER COLLECTION |>TOTAL COLLECTION |>RECEIVED/BORROWED FUNDS |>CASH RETURN |>SALES RETURN |>CASH DISBURSEMENT |>CREDIT CARD |>ONLINE PAYMENT |>INTER-BRANCH PAYMENT |>TOTAL LESS |>NET COLLECTION |>CASH |>CHECK |>TOTAL |>DEPOSITED AMOUNT |>CASH-OUT AMOUNT |>TTL RECEIPTS VS TTL CF |>TTL CF VS DEPOSIT |^DATE DEPOSITED |^INTERVALS |^DEPOSIT SLIP REF. NO. |<REMARKS"
            .ColWidth(0) = 300
            .ColWidth(1) = (.Width - 300) * 0.1
            .ColWidth(2) = (.Width - 300) * 0.13
            .ColWidth(3) = (.Width - 300) * 0.13
            .ColWidth(4) = (.Width - 300) * 0.12
            .ColWidth(5) = (.Width - 300) * 0.12
            .ColWidth(6) = (.Width - 300) * 0.15
            .ColWidth(7) = (.Width - 300) * 0.12
            .ColWidth(8) = (.Width - 300) * 0.15
            .ColWidth(9) = (.Width - 300) * 0.15
            .ColWidth(10) = (.Width - 300) * 0.12
            .ColWidth(11) = (.Width - 300) * 0.12
            .ColWidth(12) = (.Width - 300) * 0.17
            .ColWidth(13) = (.Width - 300) * 0.12
            .ColWidth(14) = (.Width - 300) * 0.12
            .ColWidth(15) = (.Width - 300) * 0.12
            .ColWidth(16) = (.Width - 300) * 0.12
            .ColWidth(17) = (.Width - 300) * 0.12
            .ColWidth(18) = (.Width - 300) * 0.15
            .ColWidth(19) = (.Width - 300) * 0.12
            .ColWidth(20) = (.Width - 300) * 0.12
            .ColWidth(21) = (.Width - 300) * 0.12
            .ColWidth(22) = (.Width - 300) * 0.12
            .ColWidth(23) = (.Width - 300) * 0.12
            .ColWidth(24) = (.Width - 300) * 0.12
            .ColWidth(25) = (.Width - 300) * 0.12
            .ColWidth(26) = (.Width - 300) * 0.15
            .ColWidth(27) = (.Width - 300) * 0.15
            .ColWidth(28) = (.Width - 300) * 0.17
            .ColWidth(29) = (.Width - 300) * 0.08
            .ColWidth(30) = (.Width - 300) * 0.15
            .ColWidth(31) = (.Width - 300) * 0.35
'******************** END CODE UPDATED: 08/04/17 ******************************
        ElseIf mySrc = "Detailed Inventory Outgoing" Then
            '                     1        2          3              4              5         6         7             8             9            10                11             12           13
            .FormatString = "  |^DATE |^REF NO |<PART NUMBER |<PART DECRIPTION |^STATUTORY |^QTY |>COST w/o VAT |>MDSE INVTY |>VAT INPUT |>TOTAL COST w/VAT |^OUTGOING TYPE |<DELIVERED TO |<SUPPLIER'S NAME"
            .ColWidth(0) = 300
            .ColWidth(1) = (.Width - 300) * 0.1         'DATE
            .ColWidth(2) = (.Width - 300) * 0.1         'REF#
            .ColWidth(3) = (.Width - 300) * 0.2         'PART#
            .ColWidth(4) = (.Width - 300) * 0.25        'DESCRIPTION
            .ColWidth(5) = (.Width - 300) * 0.1         'STATUTORY
            .ColWidth(6) = (.Width - 300) * 0.07        'QTY
            .ColWidth(7) = (.Width - 300) * 0.08        'COST
            .ColWidth(8) = (.Width - 300) * 0.08        'MDS INVTY
            .ColWidth(9) = (.Width - 300) * 0.08        'VAT INPUT
            .ColWidth(10) = (.Width - 300) * 0.11       'TTL COST
            .ColWidth(11) = (.Width - 300) * 0.1        'OUTGOING TYPE
            .ColWidth(12) = (.Width - 300) * 0.22       'DELIVER TO
            .ColWidth(13) = (.Width - 300) * 0.22       'SUPPLIERS NAME
        ElseIf mySrc = "Detailed Inventory Incoming" Then
            '                     1           2          3            4              5              6         7          8            9            10             11             12                    13                     14              15            16
            .FormatString = "  |^DATE REC |^REF NO |^PACK DATE |<PART NUMBER |<PART DECRIPTION |^STATUTORY |^QTY |>COST w/o VAT |>MDSE INVTY |>OTHERS INVTY |>VAT INPUT |>TOTAL COST w/VAT |>TOTAL COST w/VAT OTHERS |^RECEIVED TYPE |<RECEIVED FROM |<SUPPLIER'S NAME"
            .ColWidth(0) = 300
            .ColWidth(1) = (.Width - 300) * 0.1         'DATE
            .ColWidth(2) = (.Width - 300) * 0.1         'REF#
            .ColWidth(3) = (.Width - 300) * 0.1         'PACK DATE
            .ColWidth(4) = (.Width - 300) * 0.2         'PART#
            .ColWidth(5) = (.Width - 300) * 0.25        'DESCRIPTION
            .ColWidth(6) = (.Width - 300) * 0.1         'STATUTORY
            .ColWidth(7) = (.Width - 300) * 0.08        'QTY
            .ColWidth(8) = (.Width - 300) * 0.08        'COST
            .ColWidth(9) = (.Width - 300) * 0.08        'MDS INVTY
            .ColWidth(10) = (.Width - 300) * 0.08       'OTHERS INVTY
            .ColWidth(11) = (.Width - 300) * 0.08       'VAT INPUT
            .ColWidth(12) = (.Width - 300) * 0.1        'TTL COST
            .ColWidth(13) = (.Width - 300) * 0.1        'TTL COST OTHERS
            .ColWidth(14) = (.Width - 300) * 0.1        'RECEIVED TYPE
            .ColWidth(15) = (.Width - 300) * 0.22       'RECEIVED FROM
            .ColWidth(16) = (.Width - 300) * 0.22       'SUPPLIERS NAME
        ElseIf mySrc = "Sales Book" Then
            '                      1            2             3              4                5              6           7            8              9             10       11            12          13            14            15            16          17
            .FormatString = "  |^DATE |<CUSTOMER'S NAME |^INVOICE NO |^CLASSIFICATION |<SUPPLIER'S NAME |^STATUTORY |^TRANSTYPE |<PARTNUMBER |<PART DESCRIPTION |^QTY |>UNIT PRICE |>UNIT COST |>TOTAL COST |>TOTAL SALES |^CUSTOMER TYPE |<FRONTLINER |<ADDRESS"
            .ColWidth(0) = 300
            .ColWidth(1) = (.Width - 300) * 0.1         'DATE
            .ColWidth(2) = (.Width - 300) * 0.2         'CUSTOMER NAME
            .ColWidth(3) = (.Width - 300) * 0.1         'INVOICENO
            .ColWidth(4) = (.Width - 300) * 0.12        'CLASSIFICATION
            .ColWidth(5) = (.Width - 300) * 0.2         'SUPPLIERS NAME
            .ColWidth(6) = (.Width - 300) * 0.1         'STATUTORY
            .ColWidth(7) = (.Width - 300) * 0.1         'TRANSTYPE
            .ColWidth(8) = (.Width - 300) * 0.15        'PARTNUMBER
            .ColWidth(9) = (.Width - 300) * 0.2         'DESCRIPTION
            .ColWidth(10) = (.Width - 300) * 0.08       'QTY
            .ColWidth(11) = (.Width - 300) * 0.1        'UNIT PRICE
            .ColWidth(12) = (.Width - 300) * 0.1        'UNIT COST
            .ColWidth(13) = (.Width - 300) * 0.12       'TTL COST
            .ColWidth(14) = (.Width - 300) * 0.12       'TTL PRICE
            .ColWidth(15) = (.Width - 300) * 0.12       'CUSTOMER TYPE
            .ColWidth(16) = (.Width - 300) * 0.2        'FRONTLINER
            .ColWidth(17) = (.Width - 300) * 0.2        'ADDRESS
        ElseIf mySrc = "Sales Return Report" Then
            '                      1           2            3           4           5           6               7                  8            9             10             11       12           13            14           15         16           17              18                19                 20        21           22            23
            .FormatString = "  |^DATE |^SALES RET NO |^INVOICE NO |^INV. DATE |^INV. TYPE |^STATUTORY |<SUPPLIER'S NAME |<CUSTOMER'S NAME |<PARTNUMBER |<PART DESCRIPTION |^QTY |>UNIT PRICE |>TOTAL PRICE |>UNIT COST |>TOTAL COST |>AR-TRADE |>COST OF SALES |>MDSE. INVTY |>SALES RETURN & ALLOW. |>INPUT TAX |<REMARKS |^CUSTOMER TYPE |<FRONTLINER"
            .ColWidth(0) = 300
            .ColWidth(1) = (.Width - 300) * 0.1         'DATE
            .ColWidth(2) = (.Width - 300) * 0.1         'SALES RET NO
            .ColWidth(3) = (.Width - 300) * 0.1         'INVOICENO
            .ColWidth(4) = (.Width - 300) * 0.1         'INV. DATE
            .ColWidth(5) = (.Width - 300) * 0.1         'INV. TYPE
            .ColWidth(6) = (.Width - 300) * 0.1         'STATUTORY
            .ColWidth(7) = (.Width - 300) * 0.2         'SUPPLIER'S NAME
            .ColWidth(8) = (.Width - 300) * 0.2         'CUSTOMER'S NAME
            .ColWidth(9) = (.Width - 300) * 0.15        'PARTNUMBER
            .ColWidth(10) = (.Width - 300) * 0.2        'PART DESCRIPTION
            .ColWidth(11) = (.Width - 300) * 0.08       'QTY
            .ColWidth(12) = (.Width - 300) * 0.1        'UNIT PRICE
            .ColWidth(13) = (.Width - 300) * 0.11       'TTL PRICE
            .ColWidth(14) = (.Width - 300) * 0.1        'UNIT COST
            .ColWidth(15) = (.Width - 300) * 0.11       'TTL COST
            .ColWidth(16) = (.Width - 300) * 0.1        'AR-TRADE
            .ColWidth(17) = (.Width - 300) * 0.1        'COST OF SALES
            .ColWidth(18) = (.Width - 300) * 0.1        'MDSE INVTY
            .ColWidth(19) = (.Width - 300) * 0.1        'SALES RETURN AND ALLOW
            .ColWidth(20) = (.Width - 300) * 0.1        'INPUT TAX
            .ColWidth(21) = (.Width - 300) * 0.2        'REMARKS
            .ColWidth(22) = (.Width - 300) * 0.12       'CUSTOMER TYPE
            .ColWidth(23) = (.Width - 300) * 0.2        'FRONTLINER
        ElseIf mySrc = "Service Receipts Book" Then
            '                     1        2            3           4         5             6             7             8             9            10        11      12        13           14
            .FormatString = "  |^DATE |^CUSTOMER |^JOB ORDER NO |^O.R NO |^J.O TYPE |>TOTAL AMOUNT |>ACCOUNT REC. |>SRVC INCOME |>VAT OUTPUT |<MECHANIC |^REMARKS |<MODEL |<KMREADING |^SERVICE TYPE"
            .ColWidth(0) = 300
            .ColWidth(1) = (.Width - 300) * 0.1         'DATE
            .ColWidth(2) = (.Width - 300) * 0.2         'CUSTOMER
            .ColWidth(3) = (.Width - 300) * 0.1         'JOB ORDER NO
            .ColWidth(4) = (.Width - 300) * 0.1         'O.R NO
            .ColWidth(5) = (.Width - 300) * 0.11        'J.O TYPE
            .ColWidth(6) = (.Width - 300) * 0.08        'TOTAL AMOUNT
            .ColWidth(7) = (.Width - 300) * 0.1         'ACCOUNT REC.
            .ColWidth(8) = (.Width - 300) * 0.1         'SRVC INCOME
            .ColWidth(9) = (.Width - 300) * 0.1         'VAT OUTPUT
            .ColWidth(10) = (.Width - 300) * 0.2        'MECHANIC
            .ColWidth(11) = (.Width - 300) * 0.2        'REMARKS
            .ColWidth(12) = (.Width - 300) * 0.15       'MODEL
            .ColWidth(13) = (.Width - 300) * 0.1        'KMREADING
            .ColWidth(14) = (.Width - 300) * 0.1        'SERVICE TYPE
'******************** END CODE UPDATED: 07/10/19 ******************************
        ElseIf mySrc = "ADJUSTMENT REQUEST" Then
            '                     1        2         3           4              5               6                 7             8                     9               10
            .FormatString = "  |^DATE |^STATUS |^TICKET NO. |^CONTROL NO. |^ADJ. NO. |^DATE FINAL |<SUBJECT OF REQUEST |<REQUEST BY |<AUDIT APPROVED BY |<MIS IN-INCHARGE |<SCRIPT FILENAME"
            .ColWidth(0) = 300
            .ColWidth(1) = (.Width - 300) * 0.1         'DATE
            .ColWidth(2) = (.Width - 300) * 0.1         'STATUS
            .ColWidth(3) = (.Width - 300) * 0.1         'TICKET NO.
            .ColWidth(4) = (.Width - 300) * 0.1         'CONTROL NO.
            .ColWidth(5) = (.Width - 300) * 0.1         'ADJ. NO.
            .ColWidth(6) = (.Width - 300) * 0.1         'DATE FINAL
            .ColWidth(7) = (.Width - 300) * 0.2         'SUBJECT OF REQUEST
            .ColWidth(8) = (.Width - 300) * 0.15        'REQUEST BY
            .ColWidth(9) = (.Width - 300) * 0.15        'AUDIT APPROVED BY
            .ColWidth(10) = (.Width - 300) * 0.15        'MIS IN-INCHARGE
            .ColWidth(11) = (.Width - 300) * 0.15       'SCRIPT FILENAME
        ElseIf mySrc = "CANCELLATION MODULE" Then
            '                           1             2         3           4                 5          6        7       8      9          10            11
            .FormatString = "  |^DATE CREATED |^CANCEL NO. |^STATUS |<TRANSACTION TYPE |^PROG. VALUE |<REASON |<NAME |>AMOUNT |^DATE |<PREPARED BY |<APPROVED BY"
            .ColWidth(0) = 300
            .ColWidth(1) = (.Width - 300) * 0.1         'DATE CREATED
            .ColWidth(2) = (.Width - 300) * 0.1         'CANCEL NO.
            .ColWidth(3) = (.Width - 300) * 0.1         'STATUS
            .ColWidth(4) = (.Width - 300) * 0.1         'TRANSACTION TYPE
            .ColWidth(5) = (.Width - 300) * 0.1         'PROG. VALUE
            .ColWidth(6) = (.Width - 300) * 0.2         'REASON
            .ColWidth(7) = (.Width - 300) * 0.15        'NAME
            .ColWidth(8) = (.Width - 300) * 0.1         'AMOUNT
            .ColWidth(9) = (.Width - 300) * 0.1         'DATE
            .ColWidth(10) = (.Width - 300) * 0.15       'PREPARED BY
            .ColWidth(11) = (.Width - 300) * 0.15       'APPROVED BY
        End If
    End With
    
    cmdFind.Height = cmdRefresh.Height
    cmdFind.Top = cmdRefresh.Top
    txtFind.Top = cmdFind.Top
    
End Sub

Private Sub statUpdate_Click()
On Error GoTo er
Dim Qry
Qry = "call sp_mrcretails.sp_updateOrderStatus_fromWarehouse('" & sysBranchCode & "')"
If rsSummary.State = adStateOpen Then rsSummary.Close
    rsSummary.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        Call RefreshRecord(2)
        MsgBox "Order status updated successfully.", vbInformation, "Success!"
        Exit Sub
        
er:
    Screen.MousePointer = vbDefault
    Me.Enabled = True
    MsgBox Err.Description, vbCritical, "Error!"
End Sub

Private Sub txtFind_Change()
    fRow = 1
    fCol = 1
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then Call cmdFind_Click
End Sub

