VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAgingSum 
   Caption         =   "Aging Summary"
   ClientHeight    =   7890
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14250
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7890
   ScaleWidth      =   14250
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab tabProcessing 
      Height          =   975
      Left            =   4920
      TabIndex        =   13
      Tag             =   "xyf"
      Top             =   3360
      Visible         =   0   'False
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   1720
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Current Aging"
      TabPicture(0)   =   "frmAgingSum.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblProcess"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Image2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label21"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         Caption         =   "PLEASE WAIT...."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   960
         TabIndex        =   15
         Tag             =   "xyf"
         Top             =   360
         Width           =   3255
      End
      Begin VB.Image Image2 
         Height          =   720
         Left            =   240
         Picture         =   "frmAgingSum.frx":001C
         Stretch         =   -1  'True
         Tag             =   "xyf"
         Top             =   120
         Width           =   720
      End
      Begin VB.Label lblProcess 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Loading Accounts Receivables ..."
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
         Left            =   720
         TabIndex        =   14
         Tag             =   "xyf"
         Top             =   600
         Width           =   3975
      End
   End
   Begin VB.Timer tmrExecForceRun 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   13200
      Top             =   0
   End
   Begin lvButton.lvButtons_H cmdReprocess 
      Height          =   375
      Left            =   1680
      TabIndex        =   12
      Top             =   45
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
      Caption         =   "&Current Aging"
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
      Image           =   "frmAgingSum.frx":0A14
      cBack           =   -2147483633
   End
   Begin MSComDlg.CommonDialog dlgDEMO 
      Left            =   13680
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin lvButton.lvButtons_H cmdSaveExcel 
      Height          =   375
      Left            =   0
      TabIndex        =   10
      Top             =   45
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
      Caption         =   "&Save Excel"
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
      Image           =   "frmAgingSum.frx":0FAE
      cBack           =   -2147483633
   End
   Begin VB.Timer tmrInit 
      Enabled         =   0   'False
      Left            =   11760
      Top             =   0
   End
   Begin lvButton.lvButtons_H cmdExit 
      Height          =   375
      Left            =   3360
      TabIndex        =   1
      Top             =   45
      Width           =   1680
      _ExtentX        =   2963
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
      Image           =   "frmAgingSum.frx":1548
      cBack           =   -2147483633
   End
   Begin VB.Frame Frame1 
      Height          =   7455
      Left            =   120
      TabIndex        =   2
      Tag             =   "xyf"
      Top             =   360
      Width           =   14055
      Begin TabDlg.SSTab tabAgingSum 
         Height          =   5895
         Left            =   120
         TabIndex        =   0
         Tag             =   "xyf"
         Top             =   1440
         Width           =   13815
         _ExtentX        =   24368
         _ExtentY        =   10398
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   10
         TabHeight       =   520
         TabCaption(0)   =   "Grid View"
         TabPicture(0)   =   "frmAgingSum.frx":1AE2
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "flxAgingSum"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "pBar"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).ControlCount=   2
         TabCaption(1)   =   "Print Preview"
         TabPicture(1)   =   "frmAgingSum.frx":1AFE
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "CRViewer1"
         Tab(1).ControlCount=   1
         Begin MSComctlLib.ProgressBar pBar 
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Tag             =   "xyf"
            Top             =   5520
            Visible         =   0   'False
            Width           =   13575
            _ExtentX        =   23945
            _ExtentY        =   450
            _Version        =   393216
            Appearance      =   1
         End
         Begin CRVIEWERLibCtl.CRViewer CRViewer1 
            Height          =   5415
            Left            =   -74880
            TabIndex        =   17
            Tag             =   "xyf"
            Top             =   360
            Width           =   13575
            DisplayGroupTree=   -1  'True
            DisplayToolbar  =   -1  'True
            EnableGroupTree =   0   'False
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
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxAgingSum 
            Height          =   5055
            Left            =   120
            TabIndex        =   18
            Tag             =   "xyf"
            Top             =   360
            Width           =   13575
            _ExtentX        =   23945
            _ExtentY        =   8916
            _Version        =   393216
            FillStyle       =   1
            AllowUserResizing=   1
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
      End
      Begin VB.Frame Frame2 
         Height          =   1215
         Left            =   120
         TabIndex        =   3
         Tag             =   "xy"
         Top             =   120
         Width           =   13815
         Begin VB.ComboBox cboRefType 
            Height          =   315
            ItemData        =   "frmAgingSum.frx":1B1A
            Left            =   4800
            List            =   "frmAgingSum.frx":1B27
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Tag             =   "xyf"
            Top             =   240
            Visible         =   0   'False
            Width           =   2775
         End
         Begin VB.ComboBox cboYear 
            Height          =   315
            Left            =   3840
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Tag             =   "xyf"
            Top             =   720
            Width           =   855
         End
         Begin VB.ComboBox cboMonth 
            Height          =   315
            Left            =   1920
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Tag             =   "xyf"
            Top             =   720
            Width           =   1815
         End
         Begin VB.ComboBox cboAgingType 
            Height          =   315
            ItemData        =   "frmAgingSum.frx":1B44
            Left            =   1920
            List            =   "frmAgingSum.frx":1B46
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Tag             =   "xyf"
            Top             =   240
            Width           =   2775
         End
         Begin lvButton.lvButtons_H cmdRefresh 
            Height          =   375
            Left            =   4800
            TabIndex        =   9
            Tag             =   "xyf"
            Top             =   660
            Width           =   480
            _ExtentX        =   847
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
            ImgAlign        =   4
            Image           =   "frmAgingSum.frx":1B48
            cBack           =   -2147483633
         End
         Begin VB.Label Label2 
            Caption         =   "For the month of:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   360
            TabIndex        =   6
            Tag             =   "xyf"
            Top             =   720
            Width           =   1455
         End
         Begin VB.Label Label1 
            Caption         =   "Aging Type:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   360
            TabIndex        =   5
            Tag             =   "xyf"
            Top             =   240
            Width           =   1575
         End
      End
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   10800
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   14490
      m_MDI_Height    =   8460
   End
   Begin VB.Menu mnuAging 
      Caption         =   "Aging"
      Enabled         =   0   'False
      Visible         =   0   'False
   End
End
Attribute VB_Name = "frmAgingSum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim param_Date As String
Public ParamValue1 As String
Dim ParamValueLookUp As Boolean

Private Sub cboAgingType_Click()
    ParamValueLookUp = True
    If cboAgingType.Text = "Collection Efficiency Details" Then
        cboRefType.Visible = True
    Else
        cboRefType.Visible = False
    End If
End Sub

Private Sub cboMonth_Click()
    ParamValueLookUp = True
End Sub

Private Sub cboRefType_Click()
    ParamValueLookUp = True
End Sub

Private Sub cboYear_Click()
    ParamValueLookUp = True
End Sub

Private Sub cmdExit_Click()
    Unload Me
    'MotoRep.imgBG.Visible = True
    MotoRep.toolBarButton.Visible = True
End Sub

Private Sub cmdRefresh_Click()
On Error GoTo conError
    Dim Qry As String
    Dim FRowLookUp As Integer
    
    If cboAgingType.Text = "" Or cboMonth.Text = "" Or cboYear.Text = "" Then
        MsgBox "Please complete field!", vbCritical, "Error"
        Exit Sub
    End If
    
    param_Date = cboYear.Text & "-" & (cboMonth.ListIndex + 1) & "-1"
    If cboAgingType.Text = "Details" Then
        Qry = "Call sp_mrcretails.sp_rpt_AgingDetails('" & param_Date & "'," & SysBranchID & ")"
        
    ElseIf cboAgingType.Text = "Customer" Then
        Qry = "Call sp_mrcretails.sp_rpt_AgingSummaryPerCustomer('" & param_Date & "'," & SysBranchID & ")"
        
    ElseIf cboAgingType.Text = "Collection Efficiency Summary" Then
        Qry = "Call sp_mrcretails.sp_rpt_AgingCollEfficiencySummary('" & param_Date & "'," & SysBranchID & ")"
        
    ElseIf cboAgingType.Text = "Collection Efficiency Details" Then
        If Trim(cboRefType.Text) = "" Then
            MsgBox "MISSING REF. TYPE PARAMETER!", vbExclamation, "VALIDATION"
            cboRefType.SetFocus
            Exit Sub
        End If
        Qry = "Call sp_mrcretails.sp_rpt_AgingCollEfficiencyDetails('" & param_Date & "','" & Trim(cboRefType.Text) & "'," & SysBranchID & ")"
    End If
    
    Call RXResizer1_ResizeComplete
     
    Screen.MousePointer = vbHourglass
    DoEvents
    Me.Enabled = False
    DoEvents
    
    ParamValueLookUp = False
    
    tabAgingSum.Tab = 0
    DoEvents
    Set rsAnyUse = LYRcon.Execute(Qry)
    With rsAnyUse
        Pbar.Visible = True
        
        FRowLookUp = 1
        If .RecordCount > 0 Then
            .MoveFirst
            Pbar.Value = 0
            Pbar.Max = .RecordCount
                Do While Not .EOF
                    If cboAgingType.Text = "Details" Then
                        '                 0        1          2          3         4          5         6        7         8           9
                        '.FormatString = "   |<Customer Name |<Ref. Type |<Ref. No.|^Ref. Date|>Ref. Amount|^Terms|^Due Date|>Not Yet Due|>Current" & _
                        '"|>Due Next Mo.|>1st Due|>2nd Due|>3rd&Up Due|>Total Overdue|>Payments|>Sales Return |>CM |>DM |>Coll. Nxt Mo. |<Month Of"
                        '      10           11      12        13           14            15         16         17   18      19            20
                        'CUSTOMERNAME, RefNo, RefType, TransDate, AcctAmount, Terms, DueDate, NotYetDue, Current, DueNextMo
                        ', OverDue1Mo, OverDue2Mos, OverDue3MosUp, TotalOverDue, Payments, SalesReturns, CM, DM, CollNextMo, AgingMainID, LastModified, ModifiedBy, MonthEnd, ForTheMonth, BRANCHNAME, BRANCHCODE, BRANCHADDRESS
                        If FRowLookUp = 1 Then
                            flxAgingSum.TextMatrix(1, 1) = !CustomerName
                            flxAgingSum.TextMatrix(1, 2) = !RefType
                            flxAgingSum.TextMatrix(1, 3) = !RefNo
                            flxAgingSum.TextMatrix(1, 4) = !TRANSDATE
                            flxAgingSum.TextMatrix(1, 5) = Format(!AcctAmount, "#,###,##0.00")
                            flxAgingSum.TextMatrix(1, 6) = !Terms
                            flxAgingSum.TextMatrix(1, 7) = !DUEDATE
                            flxAgingSum.TextMatrix(1, 8) = Format(!Notyetdue, "#,###,##0.00")
                            flxAgingSum.TextMatrix(1, 9) = Format(!Current, "#,###,##0.00")
                            flxAgingSum.TextMatrix(1, 10) = Format(!DueNextMo, "#,###,##0.00")
                            flxAgingSum.TextMatrix(1, 11) = Format(!OverDue1Mo, "#,###,##0.00")
                            flxAgingSum.TextMatrix(1, 12) = Format(!OverDue2Mos, "#,###,##0.00")
                            flxAgingSum.TextMatrix(1, 13) = Format(!OverDue3MosUp, "#,###,##0.00")
                            flxAgingSum.TextMatrix(1, 14) = Format(!TotalOverDue, "#,###,##0.00")
                            flxAgingSum.TextMatrix(1, 15) = Format(!Payments, "#,###,##0.00")
                            flxAgingSum.TextMatrix(1, 16) = Format(!SalesReturns, "#,###,##0.00")
                            flxAgingSum.TextMatrix(1, 17) = Format(!CM, "#,###,##0.00")
                            flxAgingSum.TextMatrix(1, 18) = Format(!DM, "#,###,##0.00")
                            flxAgingSum.TextMatrix(1, 19) = Format(!CollNextMo, "#,###,##0.00")
                            flxAgingSum.TextMatrix(1, 20) = !forthemonth
                        Else
                            flxAgingSum.AddItem _
                                    vbTab & !CustomerName & _
                                    vbTab & !RefType & _
                                    vbTab & !RefNo & _
                                    vbTab & !TRANSDATE & _
                                    vbTab & Format(!AcctAmount, "#,###,##0.00") & _
                                    vbTab & !Terms & _
                                    vbTab & !DUEDATE & _
                                    vbTab & Format(!Notyetdue, "#,###,##0.00") & _
                                    vbTab & Format(!Current, "#,###,##0.00") & _
                                    vbTab & Format(!DueNextMo, "#,###,##0.00") & _
                                    vbTab & Format(!OverDue1Mo, "#,###,##0.00") & _
                                    vbTab & Format(!OverDue2Mos, "#,###,##0.00") & _
                                    vbTab & Format(!OverDue3MosUp, "#,###,##0.00") & _
                                    vbTab & Format(!TotalOverDue, "#,###,##0.00") & _
                                    vbTab & Format(!Payments, "#,###,##0.00") & _
                                    vbTab & Format(!SalesReturns, "#,###,##0.00") & _
                                    vbTab & Format(!CM, "#,###,##0.00") & _
                                    vbTab & Format(!DM, "#,###,##0.00") & _
                                    vbTab & Format(!CollNextMo, "#,###,##0.00") & _
                                    vbTab & !forthemonth
                        End If
                    ElseIf cboAgingType.Text = "Customer" Then
                        'AgingMainID, CUSTOMERID, TotInvoiceNo, TotJoNO, TotPCV, AcctAmount, NotYetDue, Current, DueNextMo, OverDue1Mo
                        ',OverDue2Mos, OverDue3MosUp, TotalOverDue, Payments, CollEfficiency, SalesReturns, CM, DM, CollNextMo
                        ',CUSTOMERNAME, CreditLimit, ForTheMonth, BRANCHNAME, BRANCHCODE, BRANCHADDRESS
            
                        '                   0     1           2          3             4            5         6           7
                        '.FormatString = "   |<Customer Name|>Tot. Inv.|>Tot. J.O.|>Tot. PCV |>Credit Limit|>Not Yet Due|>Current" & _
                        '"|>Due Next Mo.|>1st Due|>2nd Due|>3rd&Up Due|>Total Overdue|>Payments|^Coll. Efficiency|>Sales Return|>Coll.Nxt Mo.|<Month Of"
                        '   8              9       10       11          12              13        14                15              16          17
                        If FRowLookUp = 1 Then
                            flxAgingSum.TextMatrix(1, 1) = !CustomerName
                            flxAgingSum.TextMatrix(1, 2) = Format(!TotInvoiceNo, "#,###,##0")
                            flxAgingSum.TextMatrix(1, 3) = Format(!TotJoNO, "#,###,##0")
                            flxAgingSum.TextMatrix(1, 4) = Format(!TotPCV, "#,###,##0")
                            flxAgingSum.TextMatrix(1, 5) = Format(!CreditLimit, "#,###,##0.00")
                            flxAgingSum.TextMatrix(1, 6) = Format(!Notyetdue, "#,###,##0.00")
                            flxAgingSum.TextMatrix(1, 7) = Format(!Current, "#,###,##0.00")
                            flxAgingSum.TextMatrix(1, 8) = Format(!DueNextMo, "#,###,##0.00")
                            flxAgingSum.TextMatrix(1, 9) = Format(!OverDue1Mo, "#,###,##0.00")
                            flxAgingSum.TextMatrix(1, 10) = Format(!OverDue2Mos, "#,###,##0.00")
                            flxAgingSum.TextMatrix(1, 11) = Format(!OverDue3MosUp, "#,###,##0.00")
                            flxAgingSum.TextMatrix(1, 12) = Format(!TotalOverDue, "#,###,##0.00")
                            flxAgingSum.TextMatrix(1, 13) = Format(!Payments, "#,###,##0.00")
                            flxAgingSum.TextMatrix(1, 14) = !CollEfficiency
                            flxAgingSum.TextMatrix(1, 15) = Format(!SalesReturns, "#,###,##0.00")
                            flxAgingSum.TextMatrix(1, 16) = Format(!CollNextMo, "#,###,##0.00")
                            flxAgingSum.TextMatrix(1, 17) = !forthemonth
                        Else
                            flxAgingSum.AddItem _
                                vbTab & !CustomerName & _
                                vbTab & Format(!TotInvoiceNo, "#,###,##0") & _
                                vbTab & Format(!TotJoNO, "#,###,##0") & _
                                vbTab & Format(!TotPCV, "#,###,##0") & _
                                vbTab & Format(!CreditLimit, "#,###,##0.00") & _
                                vbTab & Format(!Notyetdue, "#,###,##0.00") & _
                                vbTab & Format(!Current, "#,###,##0.00") & _
                                vbTab & Format(!DueNextMo, "#,###,##0.00") & _
                                vbTab & Format(!OverDue1Mo, "#,###,##0.00") & _
                                vbTab & Format(!OverDue2Mos, "#,###,##0.00") & _
                                vbTab & Format(!OverDue3MosUp, "#,###,##0.00") & _
                                vbTab & Format(!TotalOverDue, "#,###,##0.00") & _
                                vbTab & Format(!Payments, "#,###,##0.00") & _
                                vbTab & !CollEfficiency & _
                                vbTab & Format(!SalesReturns, "#,###,##0.00") & _
                                vbTab & Format(!CollNextMo, "#,###,##0.00") & _
                                vbTab & !forthemonth
                        End If
 
                    ElseIf cboAgingType.Text = "Collection Efficiency Summary" Then
                        '                 0      1          2                 3          4           5           6           7         8             9
                        '.FormatString = "   |<REF. TYPE |>Coll. Efficiency |^Tot.Cust |^Tot. Inv. |^Tot. J.O. |^Tot. PCV |>Tot.Amnt |>Not Yet Due |>Current" & _
                        '"|>Due Next Mo.|>1st Due|>2nd Due|>3rd&Up Due|>Total Overdue|>Payments|>Sales Return |>CM |>DM |>Coll.Next Mo.|<Month of"
                        '     10           11       12       13          14             15        16           17  18    19             20
                        
                        'RefType, CollEfficiency, TotCustomer, TotInvoiceNo, TotJoNO, TotPCV, AcctAmount, NotYetDue, Current, DueNextMo, OverDue1Mo
                        ',OverDue2Mos, OverDue3MosUp, TotalOverDue, Payments, SalesReturns, CM, DM, CollNextMo, ForTheMonth, BRANCHNAME, BRANCHCODE, BRANCHADDRESS
                        If FRowLookUp = 1 Then
                            flxAgingSum.TextMatrix(1, 1) = !RefType
                            flxAgingSum.TextMatrix(1, 2) = !CollEfficiency
                            flxAgingSum.TextMatrix(1, 3) = Format(!TotCustomer, "#,###,##0")
                            flxAgingSum.TextMatrix(1, 4) = Format(!TotInvoiceNo, "#,###,##0")
                            flxAgingSum.TextMatrix(1, 5) = Format(!TotJoNO, "#,###,##0")
                            flxAgingSum.TextMatrix(1, 6) = Format(!TotPCV, "#,###,##0")
                            flxAgingSum.TextMatrix(1, 7) = Format(!AcctAmount, "#,###,##0.00")
                            flxAgingSum.TextMatrix(1, 8) = Format(!Notyetdue, "#,###,##0.00")
                            flxAgingSum.TextMatrix(1, 9) = Format(!Current, "#,###,##0.00")
                            flxAgingSum.TextMatrix(1, 10) = Format(!DueNextMo, "#,###,##0.00")
                            flxAgingSum.TextMatrix(1, 11) = Format(!OverDue1Mo, "#,###,##0.00")
                            flxAgingSum.TextMatrix(1, 12) = Format(!OverDue2Mos, "#,###,##0.00")
                            flxAgingSum.TextMatrix(1, 13) = Format(!OverDue3MosUp, "#,###,##0.00")
                            flxAgingSum.TextMatrix(1, 14) = Format(!TotalOverDue, "#,###,##0.00")
                            flxAgingSum.TextMatrix(1, 15) = Format(!Payments, "#,###,##0.00")
                            flxAgingSum.TextMatrix(1, 16) = Format(!SalesReturns, "#,###,##0.00")
                            flxAgingSum.TextMatrix(1, 17) = Format(!CM, "#,###,##0.00")
                            flxAgingSum.TextMatrix(1, 18) = Format(!DM, "#,###,##0.00")
                            flxAgingSum.TextMatrix(1, 19) = Format(!CollNextMo, "#,###,##0.00")
                            flxAgingSum.TextMatrix(1, 20) = !forthemonth
                        Else
                            flxAgingSum.AddItem _
                                vbTab & !RefType & _
                                vbTab & !CollEfficiency & _
                                vbTab & Format(!TotCustomer, "#,###,##0") & _
                                vbTab & Format(!TotInvoiceNo, "#,###,##0") & _
                                vbTab & Format(!TotJoNO, "#,###,##0") & _
                                vbTab & Format(!TotPCV, "#,###,##0") & _
                                vbTab & Format(!AcctAmount, "#,###,##0.00") & _
                                vbTab & Format(!Notyetdue, "#,###,##0.00") & _
                                vbTab & Format(!Current, "#,###,##0.00") & _
                                vbTab & Format(!DueNextMo, "#,###,##0.00") & _
                                vbTab & Format(!OverDue1Mo, "#,###,##0.00") & _
                                vbTab & Format(!OverDue2Mos, "#,###,##0.00") & _
                                vbTab & Format(!OverDue3MosUp, "#,###,##0.00") & _
                                vbTab & Format(!TotalOverDue, "#,###,##0.00") & _
                                vbTab & Format(!Payments, "#,###,##0.00") & _
                                vbTab & Format(!SalesReturns, "#,###,##0.00") & _
                                vbTab & Format(!CM, "#,###,##0.00") & _
                                vbTab & Format(!DM, "#,###,##0.00") & _
                                vbTab & Format(!CollNextMo, "#,###,##0.00") & _
                                vbTab & !forthemonth
                        End If
                    
   
                    ElseIf cboAgingType.Text = "Collection Efficiency Details" Then
                        'CUSTOMERNAME, RefNo, RefType, TransDate, AcctAmount, Terms, DueDate, NotYetDue, Current, DueNextMo, OverDue1Mo, OverDue2Mos
                        ',OverDue3MosUp, TotalOverDue, Payments, SalesReturns, CM, DM, CollNextMo, AgingMainID, LastModified
                        ',ModifiedBy, MonthEnd, ForTheMonth, BRANCHNAME, BRANCHCODE, BRANCHADDRESS, RefTypeCaption
                        
                        '                 0         1          2        3         4           5         6        7         8          9
                        '.FormatString = "   |<Customer Name|<Ref. Type|<Ref. No.|^Ref.Date|>Ref.Amount|^Terms|^Due Date|>Not Yet Due|>Current" & _
                        '                "   |>Due Next Mo.|>1st Due|>2nd Due|>3rd&Up Due|>Total Overdue|>Payments|>Sales Return |>CM  |>DM  |>Coll.Nxt Mo.|<Month Of"
                        '                     10            11        12        13          14           15           16          17    18    19            20
                        If FRowLookUp = 1 Then
                            flxAgingSum.TextMatrix(1, 1) = !CustomerName
                            flxAgingSum.TextMatrix(1, 2) = !RefType
                            flxAgingSum.TextMatrix(1, 3) = !RefNo
                            flxAgingSum.TextMatrix(1, 4) = !TRANSDATE
                            flxAgingSum.TextMatrix(1, 5) = Format(!AcctAmount, "#,###,##0.00")
                            flxAgingSum.TextMatrix(1, 6) = !Terms
                            flxAgingSum.TextMatrix(1, 7) = !DUEDATE
                            flxAgingSum.TextMatrix(1, 8) = Format(!Notyetdue, "#,###,##0.00")
                            flxAgingSum.TextMatrix(1, 9) = Format(!Current, "#,###,##0.00")
                            flxAgingSum.TextMatrix(1, 10) = Format(!DueNextMo, "#,###,##0.00")
                            flxAgingSum.TextMatrix(1, 11) = Format(!OverDue1Mo, "#,###,##0.00")
                            flxAgingSum.TextMatrix(1, 12) = Format(!OverDue2Mos, "#,###,##0.00")
                            flxAgingSum.TextMatrix(1, 13) = Format(!OverDue3MosUp, "#,###,##0.00")
                            flxAgingSum.TextMatrix(1, 14) = Format(!TotalOverDue, "#,###,##0.00")
                            flxAgingSum.TextMatrix(1, 15) = Format(!Payments, "#,###,##0.00")
                            flxAgingSum.TextMatrix(1, 16) = Format(!SalesReturns, "#,###,##0.00")
                            flxAgingSum.TextMatrix(1, 17) = Format(!CM, "#,###,##0.00")
                            flxAgingSum.TextMatrix(1, 18) = Format(!DM, "#,###,##0.00")
                            flxAgingSum.TextMatrix(1, 19) = Format(!CollNextMo, "#,###,##0.00")
                            flxAgingSum.TextMatrix(1, 20) = !forthemonth
                        Else
                            flxAgingSum.AddItem _
                                vbTab & !CustomerName & _
                                vbTab & !RefType & _
                                vbTab & !RefNo & _
                                vbTab & !TRANSDATE & _
                                vbTab & Format(!AcctAmount, "#,###,##0.00") & _
                                vbTab & !Terms & _
                                vbTab & !DUEDATE & _
                                vbTab & Format(!Notyetdue, "#,###,##0.00") & _
                                vbTab & Format(!Current, "#,###,##0.00") & _
                                vbTab & Format(!DueNextMo, "#,###,##0.00") & _
                                vbTab & Format(!OverDue1Mo, "#,###,##0.00") & _
                                vbTab & Format(!OverDue2Mos, "#,###,##0.00") & _
                                vbTab & Format(!OverDue3MosUp, "#,###,##0.00") & _
                                vbTab & Format(!TotalOverDue, "#,###,##0.00") & _
                                vbTab & Format(!Payments, "#,###,##0.00") & _
                                vbTab & Format(!SalesReturns, "#,###,##0.00") & _
                                vbTab & Format(!CM, "#,###,##0.00") & _
                                vbTab & Format(!DM, "#,###,##0.00") & _
                                vbTab & Format(!CollNextMo, "#,###,##0.00") & _
                                vbTab & !forthemonth
                        End If
                    End If
                
                .MoveNext
                Pbar.Value = Pbar.Value + 1
                
                FRowLookUp = 2
                DoEvents
            Loop
        Else
            flxAgingSum.Cols = 2
            MsgBox "No record found", vbCritical, "No Record"
            Pbar.Visible = False
            GoTo ExitGate
        End If
        .Close
    End With
    
ExitGate:
    Me.Enabled = True
    Screen.MousePointer = vbDefault
    
    Pbar.Visible = False
Exit Sub
conError:
    Me.Enabled = True
    Screen.MousePointer = vbDefault
    MsgBox Err.Description, vbCritical, "Error"
End Sub

Private Sub cmdReprocess_Click()
    Me.Enabled = False
    tabProcessing.Visible = True
    frmAgingSum.Enabled = False
    tmrExecForceRun.Enabled = True
End Sub

Private Sub cmdSaveExcel_Click()
On Error GoTo er
    Dim XLapp As Excel.Application
    Dim wbkMVR As Excel.Workbook
    Dim wksLineA1 As Excel.Worksheet
    Dim sSaveAsFileName
    Dim i As Integer
    Dim j As Integer
    
    If ParamValueLookUp = True Then
        MsgBox "The parameters value was change, please reload the record.", vbExclamation, "Validation!"
        Exit Sub
    End If
    
    If cboAgingType.Text = "" Or cboMonth.Text = "" Or cboYear.Text = "" Then
        MsgBox "Please complete field", vbCritical, "Error"
        Exit Sub
    End If
    
    If flxAgingSum.TextMatrix(1, 1) = "" Then
        MsgBox "No Data to Extract", vbCritical, "Error"
        Exit Sub
    End If
    
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
    
 
    If cboAgingType.Text = "Details" Then
        '                 0        1          2          3         4          5         6        7         8           9
        '.FormatString = "   |<Customer Name |<Ref. Type |<Ref. No.|^Ref. Date|>Ref. Amount|^Terms|^Due Date|>Not Yet Due|>Current" & _
        '"|>Due Next Mo.|>1st Due|>2nd Due|>3rd&Up Due|>Total Overdue|>Payments|>Sales Return |>CM |>DM |>Coll. Nxt Mo. |<Month Of"
        '      10           11      12        13           14            15         16         17   18      19            20
        wksLineA1.Name = "Aging Details"
        wksLineA1.Cells(1, 1).Value = "Customer Name"
        wksLineA1.Cells(1, 2).Value = "Ref. Type"
        wksLineA1.Cells(1, 3).Value = "Ref. No."
        wksLineA1.Cells(1, 4).Value = "Ref. Date"
        wksLineA1.Cells(1, 5).Value = "Ref. Amount"
        wksLineA1.Cells(1, 6).Value = "Terms"
        wksLineA1.Cells(1, 7).Value = "Due Date"
        wksLineA1.Cells(1, 8).Value = "Not Yet Due"
        wksLineA1.Cells(1, 9).Value = "Current"
        wksLineA1.Cells(1, 10).Value = "Due Next Mo."
        wksLineA1.Cells(1, 11).Value = "1st Due"
        wksLineA1.Cells(1, 12).Value = "2nd Due"
        wksLineA1.Cells(1, 13).Value = "3rd&Up Due"
        wksLineA1.Cells(1, 14).Value = "Total Overdue"
        wksLineA1.Cells(1, 15).Value = "Payments"
        wksLineA1.Cells(1, 16).Value = "Sales Return"
        wksLineA1.Cells(1, 17).Value = "CM"
        wksLineA1.Cells(1, 18).Value = "DM"
        wksLineA1.Cells(1, 19).Value = "Coll. Nxt Mo."
        wksLineA1.Cells(1, 20).Value = "Month Of"
        With flxAgingSum
            'Pbar.Value = 0
            'Pbar.Max = .Rows - 1
            For i = 1 To .Rows - 1
                For j = 1 To 20
                    wksLineA1.Cells(i + 1, j).Value = .TextMatrix(i, j)
                Next j
                'Pbar.Value = Pbar.Value + 1
                DoEvents
            Next i
        End With
        wksLineA1.Range("A1:T" & (i - 1)).Columns.AutoFit
        wksLineA1.Range("A1:T1").Font.Bold = True
        XLapp.DisplayAlerts = False
        
    ElseIf cboAgingType.Text = "Customer" Then
        '                   0     1           2          3             4            5         6           7
        '.FormatString = "   |<Customer Name|>Tot. Inv.|>Tot. J.O.|>Tot. PCV |>Credit Limit|>Not Yet Due|>Current" & _
        ' "|>Due Next Mo.|>1st Due|>2nd Due|>3rd&Up Due|>Total Overdue|>Payments|^Coll. Efficiency|>Sales Return|>Coll. Nxt Mo.|<Month Of"
        '   8              9       10       11          12              13        14                15              16          17
        
            
        wksLineA1.Name = "Aging by Customer"
        wksLineA1.Cells(1, 1).Value = "Customer Name"
        wksLineA1.Cells(1, 2).Value = "Tot. Inv."
        wksLineA1.Cells(1, 3).Value = "Tot. J.O."
        wksLineA1.Cells(1, 4).Value = "Tot. PCV"
        wksLineA1.Cells(1, 5).Value = "Credit Limit"
        wksLineA1.Cells(1, 6).Value = "Not Yet Due"
        wksLineA1.Cells(1, 7).Value = "Current"
        wksLineA1.Cells(1, 8).Value = "Due Next Mo."
        wksLineA1.Cells(1, 9).Value = "1st Due"
        wksLineA1.Cells(1, 10).Value = "2nd Due"
        wksLineA1.Cells(1, 11).Value = "3rd&Up Due"
        wksLineA1.Cells(1, 12).Value = "Total Overdue"
        wksLineA1.Cells(1, 13).Value = "Payments"
        wksLineA1.Cells(1, 14).Value = "Coll. Efficiency"
        wksLineA1.Cells(1, 15).Value = "Sales Return"
        wksLineA1.Cells(1, 16).Value = "Coll. Nxt Mo."
        wksLineA1.Cells(1, 17).Value = "Month Of"
        With flxAgingSum
            For i = 1 To .Rows - 1
                For j = 1 To 17
                    wksLineA1.Cells(i + 1, j).Value = .TextMatrix(i, j)
                Next j
                DoEvents
            Next i
        End With
        wksLineA1.Range("A1:Q" & (i - 1)).Columns.AutoFit
        wksLineA1.Range("A1:Q1").Font.Bold = True
        XLapp.DisplayAlerts = False
        
    ElseIf cboAgingType.Text = "Collection Efficiency Summary" Then
        '                 0      1          2                 3          4           5           6           7         8             9
        '.FormatString = "   |<REF. TYPE |>Coll. Efficiency |^Tot.Cust |^Tot. Inv. |^Tot. J.O. |^Tot. PCV |>Tot.Amnt |>Not Yet Due |>Current" & _
        '"|>Due Next Mo.|>1st Due|>2nd Due|>3rd&Up Due|>Total Overdue|>Payments|>Sales Return |>CM |>DM |>Coll.Next Mo.|<Month of"
        '     10           11       12       13          14             15        16           17  18    19             20
        wksLineA1.Name = "Collection Efficiency Summary"
        wksLineA1.Cells(1, 1).Value = "REF. TYPE"
        wksLineA1.Cells(1, 2).Value = "Coll. Efficiency"
        wksLineA1.Cells(1, 3).Value = "Tot. Cust"
        wksLineA1.Cells(1, 4).Value = "Tot. Inv."
        wksLineA1.Cells(1, 5).Value = "Tot. J.O."
        wksLineA1.Cells(1, 6).Value = "Tot. PCV"
        wksLineA1.Cells(1, 7).Value = "Tot. Amnt"
        wksLineA1.Cells(1, 8).Value = "Not Yet Due"
        wksLineA1.Cells(1, 9).Value = "Current"
        wksLineA1.Cells(1, 10).Value = "Due Next Mo."
        wksLineA1.Cells(1, 11).Value = "1st Due"
        wksLineA1.Cells(1, 12).Value = "2nd Due"
        wksLineA1.Cells(1, 13).Value = "3rd&Up Due"
        wksLineA1.Cells(1, 14).Value = "Total Overdue"
        wksLineA1.Cells(1, 15).Value = "Payments"
        wksLineA1.Cells(1, 16).Value = "Sales Return"
        wksLineA1.Cells(1, 17).Value = "CM"
        wksLineA1.Cells(1, 18).Value = "DM"
        wksLineA1.Cells(1, 19).Value = "Coll.Nxt Mo."
        wksLineA1.Cells(1, 20).Value = "Month Of"
        
        With flxAgingSum
            For i = 1 To .Rows - 1
                For j = 1 To 20
                    wksLineA1.Cells(i + 1, j).Value = .TextMatrix(i, j)
                Next j
                DoEvents
            Next i
        End With
        wksLineA1.Range("A1:T" & (i - 1)).Columns.AutoFit
        wksLineA1.Range("A1:T1").Font.Bold = True
        XLapp.DisplayAlerts = False
        
    ElseIf cboAgingType.Text = "Collection Efficiency Details" Then
        '                 0         1          2        3         4           5         6        7         8          9
        '.FormatString = "   |<Customer Name|<Ref. Type|<Ref. No.|^Ref.Date|>Ref. Amount|^Terms|^Due Date|>Not Yet Due|>Current" & _
        '                "   |>Due Next Mo.|>1st Due|>2nd Due|>3rd&Up Due|>Total Overdue|>Payments|>Sales Return |>CM  |>DM  |>Coll. Nxt Mo.|<Month Of"
        '                     10            11        12        13          14           15           16          17    18    19            20
        wksLineA1.Name = "AgingDetailsWithEfficiency"
        wksLineA1.Cells(1, 1).Value = "Customer Name"
        wksLineA1.Cells(1, 2).Value = "Ref. Type"
        wksLineA1.Cells(1, 3).Value = "Ref. No."
        wksLineA1.Cells(1, 4).Value = "Ref. Date"
        wksLineA1.Cells(1, 5).Value = "Ref. Amount"
        wksLineA1.Cells(1, 6).Value = "Terms"
        wksLineA1.Cells(1, 7).Value = "Due Date"
        wksLineA1.Cells(1, 8).Value = "Not Yet Due"
        wksLineA1.Cells(1, 9).Value = "Current"
        wksLineA1.Cells(1, 10).Value = "Due Next Mo."
        wksLineA1.Cells(1, 11).Value = "1st Due"
        wksLineA1.Cells(1, 12).Value = "2nd Due"
        wksLineA1.Cells(1, 13).Value = "3rd&Up Due"
        wksLineA1.Cells(1, 14).Value = "Total Overdue"
        wksLineA1.Cells(1, 15).Value = "Payments"
        wksLineA1.Cells(1, 16).Value = "Sales Return"
        wksLineA1.Cells(1, 17).Value = "CM"
        wksLineA1.Cells(1, 18).Value = "DM"
        wksLineA1.Cells(1, 19).Value = "Coll.Nxt Mo."
        wksLineA1.Cells(1, 20).Value = "Month Of"
        With flxAgingSum
            For i = 1 To .Rows - 1
                For j = 1 To 20
                    wksLineA1.Cells(i + 1, j).Value = .TextMatrix(i, j)
                Next j
                DoEvents
            Next i
        End With
        wksLineA1.Range("A1:T" & (i - 1)).Columns.AutoFit
        wksLineA1.Range("A1:T1").Font.Bold = True
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


Private Sub Form_Load()
On Error GoTo er
    Dim i As Integer
    
    MotoRep.sBar.Panels(6).Text = "Aging Summary"
    
    'Display Aging Type
    With cboAgingType
        .Clear
        .AddItem "Details"
        .AddItem "Customer"
        .AddItem "Collection Efficiency Summary"
        .AddItem "Collection Efficiency Details"
    End With
    
    'Display Month name
    cboMonth.Clear
    For i = 1 To 12
        cboMonth.AddItem MonthName(i)
    Next i
    
    
    tabAgingSum.Tab = 0
    Call RXResizer1_ResizeComplete
    tmrInit.Enabled = True
    Call tmrInit_Timer
Exit Sub
er:
    MsgBox Err.Description, vbCritical, "Error!"
    Call Unload(Me)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If rsAnyUse.State = adStateOpen Then rsAnyUse.Close
End Sub

Private Sub RXResizer1_ResizeComplete()
    
    With flxAgingSum
        .ClearStructure
        .Clear
        .Rows = 2
        .FixedRows = 1
        .FixedCols = 1
        
        If cboAgingType.Text = "Details" Then
            '                 0        1          2          3         4          5         6        7         8           9
            .FormatString = "   |<Customer Name |<Ref. Type |<Ref. No.|^Ref. Date|>Ref. Amount|^Terms|^Due Date|>Not Yet Due|>Current" & _
            "|>Due Next Mo.|>1st Due|>2nd Due|>3rd&Up Due|>Total Overdue|>Payments|>Sales Return |>CM |>DM |>Coll. Nxt Mo. |<Month Of"
            '      10           11      12        13           14            15         16         17   18      19            20
            'CUSTOMERNAME, RefNo, RefType, TransDate, AcctAmount, Terms, DueDate, NotYetDue, Current, DueNextMo
            ', OverDue1Mo, OverDue2Mos, OverDue3MosUp, TotalOverDue, Payments, SalesReturns, CM, DM, CollNextMo
            ', AgingMainID, LastModified, ModifiedBy, MonthEnd, ForTheMonth, BRANCHNAME, BRANCHCODE, BRANCHADDRESS
            .Cols = 21
            .ColWidth(0) = 300
            .ColWidth(1) = (.Width - 300) * 0.2
            .ColWidth(2) = (.Width - 300) * 0.1
            .ColWidth(3) = (.Width - 300) * 0.08
            .ColWidth(4) = (.Width - 300) * 0.08
            .ColWidth(5) = (.Width - 300) * 0.08
            .ColWidth(6) = (.Width - 300) * 0.08
            .ColWidth(7) = (.Width - 300) * 0.08
            .ColWidth(8) = (.Width - 300) * 0.08
            .ColWidth(9) = (.Width - 300) * 0.08
            .ColWidth(10) = (.Width - 300) * 0.08
            .ColWidth(11) = (.Width - 300) * 0.08
            .ColWidth(12) = (.Width - 300) * 0.08
            .ColWidth(13) = (.Width - 300) * 0.08
            .ColWidth(14) = (.Width - 300) * 0.08
            .ColWidth(15) = (.Width - 300) * 0.08
            .ColWidth(16) = (.Width - 300) * 0.08
            .ColWidth(17) = (.Width - 300) * 0.08
            .ColWidth(18) = (.Width - 300) * 0.08
            .ColWidth(19) = (.Width - 300) * 0.2
            .ColWidth(20) = (.Width - 300) * 0.1
            
        
        ElseIf cboAgingType.Text = "Customer" Then
            'AgingMainID, CUSTOMERID, TotInvoiceNo, TotJoNO, TotPCV, AcctAmount, NotYetDue, Current, DueNextMo, OverDue1Mo
            ',OverDue2Mos, OverDue3MosUp, TotalOverDue, Payments, CollEfficiency, SalesReturns, CM, DM, CollNextMo
            ',CUSTOMERNAME, CreditLimit, ForTheMonth, BRANCHNAME, BRANCHCODE, BRANCHADDRESS

            '                   0     1           2          3             4            5         6           7
            .FormatString = "   |<Customer Name|>Tot. Inv.|>Tot. J.O.|>Tot. PCV |>Credit Limit|>Not Yet Due|>Current" & _
                "|>Due Next Mo.|>1st Due|>2nd Due|>3rd&Up Due|>Total Overdue|>Payments|^Coll. Efficiency|>Sales Return|>Coll. Nxt Mo.|<Month Of"
            '   8              9       10       11          12              13        14                15              16          17
            
            .Cols = 18
            .ColWidth(0) = 300
            .ColWidth(1) = (.Width - 300) * 0.3
            .ColWidth(2) = (.Width - 300) * 0.06
            .ColWidth(3) = (.Width - 300) * 0.06
            .ColWidth(4) = (.Width - 300) * 0.06
            .ColWidth(5) = (.Width - 300) * 0.06
            .ColWidth(6) = (.Width - 300) * 0.08
            .ColWidth(7) = (.Width - 300) * 0.08
            .ColWidth(8) = (.Width - 300) * 0.08
            .ColWidth(9) = (.Width - 300) * 0.08
            .ColWidth(10) = (.Width - 300) * 0.08
            .ColWidth(11) = (.Width - 300) * 0.08
            .ColWidth(12) = (.Width - 300) * 0.08
            .ColWidth(13) = (.Width - 300) * 0.08
            .ColWidth(14) = (.Width - 300) * 0.08
            .ColWidth(15) = (.Width - 300) * 0.08
            .ColWidth(16) = (.Width - 300) * 0.08
            .ColWidth(17) = (.Width - 300) * 0.1

        ElseIf cboAgingType.Text = "Collection Efficiency Summary" Then
            '                 0      1          2                 3          4           5           6           7         8             9
            .FormatString = "   |<REF. TYPE |>Coll. Efficiency |^Tot.Cust |^Tot. Inv. |^Tot. J.O. |^Tot. PCV |>Tot.Amnt |>Not Yet Due |>Current" & _
            "|>Due Next Mo.|>1st Due|>2nd Due|>3rd&Up Due|>Total Overdue|>Payments|>Sales Return |>CM |>DM |>Coll.Next Mo.|<Month of"
            '     10           11       12       13          14             15        16           17  18    19             20
            
            'RefType, CollEfficiency, TotCustomer, TotInvoiceNo, TotJoNO, TotPCV, AcctAmount, NotYetDue, Current, DueNextMo, OverDue1Mo
            ',OverDue2Mos, OverDue3MosUp, TotalOverDue, Payments, SalesReturns, CM, DM, CollNextMo, ForTheMonth, BRANCHNAME, BRANCHCODE, BRANCHADDRESS
            
            .Cols = 21
            .ColWidth(0) = 300
            .ColWidth(1) = (.Width - 300) * 0.14
            .ColWidth(2) = (.Width - 300) * 0.08
            .ColWidth(3) = (.Width - 300) * 0.05
            .ColWidth(4) = (.Width - 300) * 0.05
            .ColWidth(5) = (.Width - 300) * 0.05
            .ColWidth(6) = (.Width - 300) * 0.09
            .ColWidth(7) = (.Width - 300) * 0.09
            .ColWidth(8) = (.Width - 300) * 0.09
            .ColWidth(9) = (.Width - 300) * 0.09
            .ColWidth(10) = (.Width - 300) * 0.1
            .ColWidth(11) = (.Width - 300) * 0.1
            .ColWidth(12) = (.Width - 300) * 0.1
            .ColWidth(13) = (.Width - 300) * 0.1
            .ColWidth(14) = (.Width - 300) * 0.1
            .ColWidth(15) = (.Width - 300) * 0.1
            .ColWidth(16) = (.Width - 300) * 0.1
            .ColWidth(17) = (.Width - 300) * 0.1
            .ColWidth(18) = (.Width - 300) * 0.1
            .ColWidth(19) = (.Width - 300) * 0.1
            .ColWidth(20) = (.Width - 300) * 0.1
            
        ElseIf cboAgingType.Text = "Collection Efficiency Details" Then
            'CUSTOMERNAME, RefNo, RefType, TransDate, AcctAmount, Terms, DueDate, NotYetDue, Current, DueNextMo, OverDue1Mo, OverDue2Mos
            ',OverDue3MosUp, TotalOverDue, Payments, SalesReturns, CM, DM, CollNextMo, AgingMainID, LastModified
            ',ModifiedBy, MonthEnd, ForTheMonth, BRANCHNAME, BRANCHCODE, BRANCHADDRESS, RefTypeCaption
            
            '                 0         1          2        3         4           5         6        7         8          9
            .FormatString = "   |<Customer Name|<Ref. Type|<Ref. No.|^Ref.Date|>Ref. Amount|^Terms|^Due Date|>Not Yet Due|>Current" & _
                            "   |>Due Next Mo.|>1st Due|>2nd Due|>3rd&Up Due|>Total Overdue|>Payments|>Sales Return |>CM  |>DM  |>Coll. Nxt Mo.|<Month Of"
            '                     10            11        12        13          14           15           16          17    18    19            20
            .Cols = 21
            .ColWidth(0) = 300
            .ColWidth(1) = (.Width - 300) * 0.2
            .ColWidth(2) = (.Width - 300) * 0.09
            .ColWidth(3) = (.Width - 300) * 0.08
            .ColWidth(4) = (.Width - 300) * 0.08
            .ColWidth(5) = (.Width - 300) * 0.08
            .ColWidth(6) = (.Width - 300) * 0.08
            .ColWidth(7) = (.Width - 300) * 0.08
            .ColWidth(8) = (.Width - 300) * 0.08
            .ColWidth(9) = (.Width - 300) * 0.08
            .ColWidth(10) = (.Width - 300) * 0.08
            .ColWidth(11) = (.Width - 300) * 0.08
            .ColWidth(12) = (.Width - 300) * 0.08
            .ColWidth(13) = (.Width - 300) * 0.08
            .ColWidth(14) = (.Width - 300) * 0.08
            .ColWidth(15) = (.Width - 300) * 0.08
            .ColWidth(16) = (.Width - 300) * 0.08
            .ColWidth(17) = (.Width - 300) * 0.08
            .ColWidth(18) = (.Width - 300) * 0.08
            .ColWidth(19) = (.Width - 300) * 0.08
            .ColWidth(20) = (.Width - 300) * 0.16
        Else
            Exit Sub
        End If
        
    End With

End Sub

Private Sub tabAgingSum_Click(PreviousTab As Integer)
On Error GoTo errRpt
    Dim Qry As String
    Dim rsPrint As New ADODB.Recordset
    param_Date = cboYear.Text & "-" & (cboMonth.ListIndex + 1) & "-1"
    
    If tabAgingSum.Tab = 1 Then 'Tab Print Preview
        
        If cboAgingType.Text = "" Or cboMonth.Text = "" Or cboYear.Text = "" Then
                MsgBox "Please complete field!", vbCritical, "Error"
                tabAgingSum.Tab = 0
                Exit Sub
        End If
        
        If cboAgingType.Text = "Details" Then
            Dim crAgingDetails As New rptAgingDetails
            Qry = "call sp_mrcretails.sp_rpt_AgingDetails('" & param_Date & "'," & SysBranchID & ")"
            If rsPrint.State <> 0 Then rsPrint.Close
            Set rsPrint = LYRcon.Execute(Qry)
            DoEvents
            crAgingDetails.DiscardSavedData
            crAgingDetails.Database.SetDataSource rsPrint
            DoEvents
            CRViewer1.ReportSource = crAgingDetails
            CRViewer1.EnableExportButton = True
            CRViewer1.ViewReport
                
        ElseIf cboAgingType.Text = "Customer" Then
 
            Dim crAgingCus As New rptAgingCustomerSum
            Qry = "call sp_mrcretails.sp_rpt_AgingSummaryPerCustomer('" & param_Date & "'," & SysBranchID & ")"
            If rsPrint.State <> 0 Then rsPrint.Close
            Set rsPrint = LYRcon.Execute(Qry)
            DoEvents
            crAgingCus.DiscardSavedData
            crAgingCus.Database.Tables(1).SetDataSource rsPrint
            DoEvents
            CRViewer1.ReportSource = crAgingCus
            CRViewer1.ViewReport
            

        ElseIf cboAgingType.Text = "Collection Efficiency Summary" Then
            Dim crColEfficiencySum As New rptAgingColEfficiencySum
            Qry = "Call sp_mrcretails.sp_rpt_AgingCollEfficiencySummary('" & param_Date & "'," & SysBranchID & ")"
            If rsPrint.State <> 0 Then rsPrint.Close
            Set rsPrint = LYRcon.Execute(Qry)
            DoEvents
            crColEfficiencySum.DiscardSavedData
            crColEfficiencySum.Database.Tables(1).SetDataSource rsPrint
            DoEvents
            CRViewer1.ReportSource = crColEfficiencySum
            CRViewer1.ViewReport
            
        ElseIf cboAgingType.Text = "Collection Efficiency Details" Then
            If Trim(cboRefType.Text) = "" Then
                MsgBox "MISSING REF. TYPE PARAMETER!", vbExclamation, "VALIDATION"
                tabAgingSum.Tab = 0
                cboRefType.SetFocus
                Exit Sub
            End If
            
            Dim crColEfficiencyDTLS As New rptAgingColEfficiencyDTLS
            Qry = "Call sp_mrcretails.sp_rpt_AgingCollEfficiencyDetails('" & param_Date & "','" & Trim(cboRefType.Text) & "'," & SysBranchID & ")"
            If rsPrint.State <> 0 Then rsPrint.Close
            Set rsPrint = LYRcon.Execute(Qry)
            DoEvents
            crColEfficiencyDTLS.DiscardSavedData
            crColEfficiencyDTLS.Database.Tables(1).SetDataSource rsPrint
            DoEvents
            CRViewer1.ReportSource = crColEfficiencyDTLS
            CRViewer1.ViewReport
            
        End If
           
    End If
    Exit Sub
errRpt:
    MsgBox Err.Description, vbCritical, "Error->tabAgingSum_Click()"
End Sub

Private Sub tmrInit_Timer()
    Dim Qry As String
    'Display Year
    Qry = "Call sp_mrcretails.sp_GetYear(" & SysBranchID & ")"
    With cboYear
        .Clear
        If rsAnyUse.State <> 0 Then rsAnyUse.Close
        rsAnyUse.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
            If rsAnyUse.RecordCount > 0 Then
                rsAnyUse.MoveFirst
                Do While Not rsAnyUse.EOF
                    cboYear.AddItem rsAnyUse!Yr
                rsAnyUse.MoveNext
                Loop
            End If
        rsAnyUse.Close
    End With
    tmrInit.Enabled = False
End Sub

Private Sub tmrExecForceRun_Timer()
On Error GoTo er
    tmrExecForceRun.Enabled = False
    If rsAnyUse.State <> 0 Then rsAnyUse.Close
    Set rsAnyUse = LYRcon.Execute("Call sp_mrcretails.sp_agingmain(month(curdate()), year(curdate()), 'frun'," & SysBranchID & ")")
    If Not (rsAnyUse.EOF And rsAnyUse.BOF) Then
        If rsAnyUse!level = "Error" Then
            MsgBox rsAnyUse!message, vbCritical, "Error Message"
            tabProcessing.Visible = False
        Else
            MsgBox rsAnyUse!message, vbInformation, "Process Granted"
            tabProcessing.Visible = False
        End If
    End If
    If rsAnyUse.State <> 0 Then rsAnyUse.Close
    tmrInit.Enabled = True
    Call tmrInit_Timer
    frmAgingSum.Enabled = True
    Me.Enabled = True
Exit Sub
er:
    Me.Enabled = True
    frmAgingSum.Enabled = True
    MsgBox Err.Description, vbCritical, "Error->tmrExecForceRun_Timer()"
End Sub
