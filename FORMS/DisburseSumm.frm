VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form DisburseSumm 
   Caption         =   "Petty Cash Voucher Summary"
   ClientHeight    =   8640
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11940
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8640
   ScaleWidth      =   11940
   WindowState     =   2  'Maximized
   Begin MSFlexGridLib.MSFlexGrid flxCV 
      Height          =   615
      Left            =   1680
      TabIndex        =   8
      Top             =   9720
      Visible         =   0   'False
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   1085
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      FormatString    =   "CV No   |<Amount    |<Date        |<Remarks"
   End
   Begin VB.TextBox txtCriteria 
      Height          =   285
      Left            =   5160
      TabIndex        =   19
      Top             =   9240
      Visible         =   0   'False
      Width           =   2055
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxDisb 
      Height          =   7335
      Left            =   120
      TabIndex        =   18
      Tag             =   "XY"
      Top             =   1200
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   12938
      _Version        =   393216
      RowHeightMin    =   280
      SelectionMode   =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   120
      TabIndex        =   9
      Tag             =   "XY"
      Top             =   480
      Width           =   11655
      Begin VB.TextBox txtFind 
         Height          =   285
         Left            =   9600
         TabIndex        =   10
         Tag             =   "XYF"
         Top             =   195
         Width           =   1935
      End
      Begin lvButton.lvButtons_H cmdFind 
         Height          =   375
         Left            =   8280
         TabIndex        =   11
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
         Image           =   "DisburseSumm.frx":0000
         cBack           =   -2147483633
      End
      Begin MSComCtl2.DTPicker dFrom 
         Height          =   315
         Left            =   1440
         TabIndex        =   12
         Top             =   180
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         Format          =   101449729
         CurrentDate     =   39322
      End
      Begin lvButton.lvButtons_H cmdNext 
         Height          =   375
         Left            =   3240
         TabIndex        =   13
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
         Image           =   "DisburseSumm.frx":059A
         cBack           =   -2147483633
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
         Image           =   "DisburseSumm.frx":0B34
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
         Format          =   101449729
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
         Image           =   "DisburseSumm.frx":10CE
         cBack           =   -2147483633
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
         TabIndex        =   17
         Top             =   210
         Width           =   1335
      End
   End
   Begin CRVIEWERLibCtl.CRViewer vwrDisb 
      Height          =   7935
      Left            =   120
      TabIndex        =   4
      Tag             =   "XY"
      Top             =   480
      Width           =   11655
      DisplayGroupTree=   0   'False
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
   Begin VB.Timer tmrDropRev 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   4440
      Top             =   9240
   End
   Begin VB.Timer tmrFloatRev 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   4440
      Top             =   8760
   End
   Begin VB.Timer tmrDrop 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3840
      Top             =   9240
   End
   Begin VB.Timer tmrFloat 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3840
      Top             =   8760
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   2
      Left            =   3360
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
      Caption         =   "&Show Voucher"
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
      Image           =   "DisburseSumm.frx":1668
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   3
      Left            =   5040
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
      Caption         =   "&Filter Records"
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
      Image           =   "DisburseSumm.frx":1C02
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
      Caption         =   "&Grid View"
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
      Image           =   "DisburseSumm.frx":219C
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   1
      Left            =   1680
      TabIndex        =   3
      Top             =   0
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
      Image           =   "DisburseSumm.frx":2736
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   4
      Left            =   6720
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
      Caption         =   "&Print Preview"
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
      Image           =   "DisburseSumm.frx":2CD0
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   5
      Left            =   8400
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
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
      Image           =   "DisburseSumm.frx":326A
      cBack           =   -2147483633
   End
   Begin VB.Label lblRemarks 
      Height          =   255
      Left            =   5160
      TabIndex        =   20
      Top             =   9000
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.Label lblP 
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   9240
      Visible         =   0   'False
      Width           =   255
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   360
      Top             =   8640
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   12180
      m_MDI_Height    =   9210
   End
   Begin VB.Menu mnuPCV 
      Caption         =   "PCV"
      Visible         =   0   'False
      Begin VB.Menu mnuCancel 
         Caption         =   "Cancel PCV"
      End
   End
End
Attribute VB_Name = "DisburseSumm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsRPT As New ADODB.Recordset

Dim BtnIndx As Integer, ActRow As Integer, Qry As String
Dim fRow As Integer, fCol As Integer
Public Sub ShowFilteredRecords()
    Qry = ""
    If Val(cmdNext.Tag) = 1 Then
        '***Access
        'Qry = "SELECT DISBURSEMENT.CVDATE, Format([CVNUMBER],'0000') AS CVNO, DISBURSEMENT.DISBURSETYPE, DISBURSEMENT.ACCTDESC, DISBURSEMENT.PAYEE, DISBURSEMENT.PARTICULARS, DISBURSEMENT.AMOUNT, DISBURSEMENT.STATUS, DISBURSEMENT.VOUCHERID " & _
              "From Disbursement " & txtCriteria.Text & " And ((Disbursement.CVDate) = #" & dFrom.Value & "#)) ORDER BY DISBURSEMENT.CVDATE DESC , DISBURSEMENT.CVNUMBER DESC;"
        'mySQL
        Qry = "SELECT DISBURSEMENT.CVDATE, DISBURSEMENT.CVNUMBER AS CVNO, DISBURSEMENT.DISBURSETYPE, DISBURSEMENT.ACCTDESC, DISBURSEMENT.PAYEE, DISBURSEMENT.PARTICULARS, DISBURSEMENT.AMOUNT, DISBURSEMENT.`STATUS`, DISBURSEMENT.VOUCHERID " & _
            "From Disbursement " & txtCriteria.Text & " And ((Disbursement.CVDate) = '" & Format(dFrom.Value, "yyyy-mm-dd") & "')) ORDER BY DISBURSEMENT.CVDATE DESC , DISBURSEMENT.CVNUMBER DESC;"
    
    Else
        'Qry = "SELECT Disbursement.CVDATE, Format([CVNUMBER],'0000') AS CVNO, Disbursement.DISBURSETYPE, Disbursement.ACCTDESC, Disbursement.PAYEE, Disbursement.PARTICULARS, Disbursement.AMOUNT, Disbursement.STATUS, Disbursement.VOUCHERID " & _
              "From Disbursement " & txtCriteria.Text & " And ((Disbursement.CVDate) >= #" & dFrom.Value & "# And (Disbursement.CVDate) <= #" & dTo.Value & "#)) ORDER BY Disbursement.CVDATE DESC , Disbursement.CVNUMBER DESC;"
        'mySQL
        Qry = "SELECT Disbursement.CVDATE, Disbursement.CVNUMBER AS CVNO, Disbursement.DISBURSETYPE, Disbursement.ACCTDESC, Disbursement.PAYEE, Disbursement.PARTICULARS, Disbursement.AMOUNT, Disbursement.STATUS, Disbursement.VOUCHERID " & _
              "From Disbursement " & txtCriteria.Text & " And ((Disbursement.CVDate) >= '" & Format(dFrom.Value, "yyyy-mm-dd") & "' And (Disbursement.CVDate) <= '" & Format(dTo.Value, "yyyy-mm-dd") & "')) ORDER BY Disbursement.CVDATE DESC , Disbursement.CVNUMBER DESC;"
    
    End If
    If rsRPT.State = adStateOpen Then rsRPT.Close
    rsRPT.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    If rsRPT.RecordCount = 0 Then
        flxDisb.Clear
        flxDisb.Rows = 2
        flxDisb.Cols = 10
        Call RXResizer1_ResizeComplete
        Exit Sub
    End If
    Set flxDisb.DataSource = rsRPT
    ActRow = 1
    flxDisb.Visible = True
    Call RXResizer1_ResizeComplete
    flxCV.Rows = 1
End Sub

Private Sub RefreshRecord(dCode As Integer)
    If dCode = 1 Then
        'Qry = "SELECT DISBURSEMENT.CVDATE, Format([CVNUMBER],'0000') AS CVNO, DISBURSEMENT.DISBURSETYPE, DISBURSEMENT.ACCTDESC, DISBURSEMENT.PAYEE, DISBURSEMENT.PARTICULARS, DISBURSEMENT.AMOUNT, DISBURSEMENT.STATUS, DISBURSEMENT.VOUCHERID From Disbursement " & _
              "Where (((Disbursement.CVDate) = #" & dFrom.Value & "#)) ORDER BY DISBURSEMENT.CVDATE DESC , DISBURSEMENT.CVNUMBER DESC;"
        'mySQL
        Qry = "SELECT DISBURSEMENT.CVDATE, DISBURSEMENT.CVNUMBER AS CVNO, DISBURSEMENT.DISBURSETYPE, DISBURSEMENT.ACCTDESC, DISBURSEMENT.PAYEE, DISBURSEMENT.PARTICULARS, DISBURSEMENT.AMOUNT, DISBURSEMENT.STATUS, DISBURSEMENT.VOUCHERID From Disbursement " & _
              "Where Disbursement.CVDate = '" & Format(dFrom.Value, "yyyy-mm-dd") & "' ORDER BY DISBURSEMENT.CVDATE DESC , DISBURSEMENT.CVNUMBER DESC;"
    Else
        'Qry = "SELECT Disbursement.CVDATE, Format([CVNUMBER],'0000') AS CVNO, Disbursement.DISBURSETYPE, Disbursement.ACCTDESC, Disbursement.PAYEE, Disbursement.PARTICULARS, Disbursement.AMOUNT, Disbursement.STATUS, Disbursement.VOUCHERID From Disbursement " & _
              "Where (((Disbursement.CVDate) >= #" & dFrom.Value & "# And (Disbursement.CVDate) <= #" & dTo.Value & "#)) ORDER BY Disbursement.CVDATE DESC , Disbursement.CVNUMBER DESC;"
        'mySQL
        Qry = "SELECT Disbursement.CVDATE, Disbursement.CVNUMBER AS CVNO, Disbursement.DISBURSETYPE, Disbursement.ACCTDESC, Disbursement.PAYEE, Disbursement.PARTICULARS, Disbursement.AMOUNT, Disbursement.`STATUS`, Disbursement.VOUCHERID From Disbursement " & _
              "Where Disbursement.CVDate >= '" & Format(dFrom.Value, "yyyy-mm-dd") & "' And Disbursement.CVDate <= '" & Format(dTo.Value, "yyyy-mm-dd") & "' ORDER BY Disbursement.CVDATE DESC , Disbursement.CVNUMBER DESC;"
    End If
    If rsRPT.State = adStateOpen Then rsRPT.Close
    rsRPT.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    If rsRPT.RecordCount = 0 Then
        flxDisb.Clear
        flxDisb.Rows = 2
        flxDisb.Cols = 10
        Call RXResizer1_ResizeComplete
        Exit Sub
    End If
    Set flxDisb.DataSource = rsRPT
    ActRow = 1
    flxDisb.Visible = True
    Call RXResizer1_ResizeComplete
    Call LoadCashVoucher
    txtCriteria.Text = ""
    lblRemarks.Caption = ""
End Sub

Private Sub cmdFind_Click()
    Dim i As Integer, j As Integer
        
    If txtFind.Text = "" Then Exit Sub
    With flxDisb
        For i = fRow To .Rows - 1
            For j = fCol To 8
                If InStr(UCase$(.TextMatrix(i, j)), txtFind.Text) > 0 Then
                    .Col = 1: .Row = i
                    .ColSel = 8: .RowSel = i
                    .TopRow = i
                    fRow = IIf(j = 8, i + 1, i)
                    fCol = IIf(j = 8, 1, j + 1)
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
    dTo.Visible = True
    cmdRefresh.Visible = True
    dTo.Value = dFrom.Value
End Sub

Private Sub cmdPrev_Click()
    cmdNext.Tag = "1"
    cmdNext.Visible = True
    cmdPrev.Visible = False
    dTo.Visible = False
    cmdRefresh.Visible = False
    Call RefreshRecord(1)
End Sub

Private Sub cmdRefresh_Click()
    Call RefreshRecord(2)
End Sub


Private Sub cmdTrans_Click(Index As Integer)
    Dim MyQry As String
    
    Select Case Index
        Case 0  'Grid View
            BtnIndx = 1
            tmrFloat.Enabled = True
            Call PrintPreView(False)
        Case 1  'Exit
            Unload Me
            'MotoRep.tbMain.Visible = True
            'MotoRep.imgBG.Visible = True
            MotoRep.toolBarButton.Visible = True
            'Disbursement.Show
        Case 2  'Show Voucher
            With flxDisb
                If ActRow = 0 Or ActRow > .Rows - 1 Then
                    MsgBox "Please select a record to view!", vbInformation + vbOKOnly, "Confirm"
                    Exit Sub
                End If
                If flxDisb.TextMatrix(ActRow, 3) = "Cancelled Vouchers" Then
                    MsgBox "Unable to view Cancelled Vouchers!", vbExclamation + vbOKOnly, "View"
                    Exit Sub
                End If
                DisbursementV2.PCVID = .TextMatrix(ActRow, 9)
                Call DisbursementV2.DisplayVoucher
            End With
            Unload Me
            DisbursementV2.Show vbModal
        Case 3  'Filter
            For i = 0 To 2
                 With SPfilter.cboLookIn(i)
                   .AddItem "PCV NUMBER"
                   .AddItem "DISBURSEMENT TYPE"
                   .AddItem "ACCOUNT DESCRIPTION"
                   .AddItem "NAME OF PAYEE"
                   .AddItem "PARTICULARS"
                   .AddItem "PCV AMOUNT"
                 End With
            Next i
            SPfilter.fldNames = "Disbursement.CVNUMBER*Disbursement.DISBURSETYPE*Disbursement.ACCTDESC*Disbursement.PAYEE*Disbursement.PARTICULARS*Disbursement.amount"
            SPfilter.fldNdx = "0" 'start w/ 0
            SPfilter.fldDates = ""
            SPfilter.fldSupplier = 10
            SPfilter.MySource = "DISBURSEMENT"
            SPfilter.Show vbModal
            
        Case 4  'Print Preview
            BtnIndx = 5
            tmrFloatRev.Enabled = True
            Call PrintPreView(True)
            If txtCriteria.Text = "" Then
                If Val(cmdNext.Tag) = 1 Then
                    'MyQry = "SELECT DISBURSEMENT.CVDATE, Format([CVNUMBER],'0000') AS CVNO, DISBURSEMENT.DISBURSETYPE, DISBURSEMENT.ACCTDESC, DISBURSEMENT.PAYEE, DISBURSEMENT.PARTICULARS, DISBURSEMENT.AMOUNT, Disbursement.DEPARTMENT " & _
                            "From Disbursement WHERE (((DISBURSEMENT.CVDATE)=#" & dFrom.Value & "#) AND ((Disbursement.STATUS)='Final')) ORDER BY Disbursement.DEPARTMENT DESC;"
                    'mySQL
                    MyQry = "SELECT DISBURSEMENT.CVDATE, DISBURSEMENT.CVNUMBER AS CVNO, DISBURSEMENT.DISBURSETYPE, DISBURSEMENT.ACCTDESC, DISBURSEMENT.PAYEE, DISBURSEMENT.PARTICULARS, DISBURSEMENT.AMOUNT, Disbursement.DEPARTMENT " & _
                            "From Disbursement WHERE DISBURSEMENT.CVDATE='" & Format(dFrom.Value, "yyyy-mm-dd") & "' AND Disbursement.`STATUS`='Final' ORDER BY Disbursement.DEPARTMENT DESC;"
                Else
                    'MyQry = "SELECT Disbursement.CVDATE, Format([CVNUMBER],'0000') AS CVNO, Disbursement.DISBURSETYPE, Disbursement.ACCTDESC, Disbursement.PAYEE, Disbursement.PARTICULARS, Disbursement.AMOUNT, Disbursement.DEPARTMENT " & _
                            "From Disbursement WHERE (((Disbursement.CVDATE)>=#" & dFrom.Value & "# And (Disbursement.CVDATE)<=#" & dTo.Value & "#) AND ((Disbursement.STATUS)='Final')) ORDER BY Disbursement.DEPARTMENT DESC;"
                    'mySQL
                    MyQry = "SELECT Disbursement.CVDATE, Disbursement.CVNUMBER AS CVNO, Disbursement.DISBURSETYPE, Disbursement.ACCTDESC, Disbursement.PAYEE, Disbursement.PARTICULARS, Disbursement.AMOUNT, Disbursement.DEPARTMENT " & _
                            "From Disbursement WHERE Disbursement.CVDATE>='" & Format(dFrom.Value, "yyyy-mm-dd") & "' And Disbursement.CVDATE<='" & Format(dTo.Value, "yyyy-mm-dd") & "' AND Disbursement.`STATUS`='Final' ORDER BY Disbursement.DEPARTMENT DESC;"
                End If
            Else
                If Val(cmdNext.Tag) = 1 Then
                    'MyQry = "SELECT DISBURSEMENT.CVDATE, Format([CVNUMBER],'0000') AS CVNO, DISBURSEMENT.DISBURSETYPE, DISBURSEMENT.ACCTDESC, DISBURSEMENT.PAYEE, DISBURSEMENT.PARTICULARS, DISBURSEMENT.AMOUNT, Disbursement.DEPARTMENT " & _
                            "From Disbursement " & txtCriteria.Text & " AND ((Disbursement.CVDate) = #" & dFrom.Value & "#)) ORDER BY DISBURSEMENT.CVDATE DESC , DISBURSEMENT.CVNUMBER DESC;"
                    'mySQL
                    MyQry = "SELECT DISBURSEMENT.CVDATE, DISBURSEMENT.CVNUMBER AS CVNO, DISBURSEMENT.DISBURSETYPE, DISBURSEMENT.ACCTDESC, DISBURSEMENT.PAYEE, DISBURSEMENT.PARTICULARS, DISBURSEMENT.AMOUNT, Disbursement.DEPARTMENT " & _
                            "From Disbursement " & txtCriteria.Text & " AND Disbursement.CVDate = '" & Format(dFrom.Value, "yyyy-mm-dd") & "' ORDER BY DISBURSEMENT.CVDATE DESC , DISBURSEMENT.CVNUMBER DESC;"
                Else
                    'MyQry = "SELECT Disbursement.CVDATE, Format([CVNUMBER],'0000') AS CVNO, Disbursement.DISBURSETYPE, Disbursement.ACCTDESC, Disbursement.PAYEE, Disbursement.PARTICULARS, Disbursement.AMOUNT, Disbursement.DEPARTMENT " & _
                            "From Disbursement " & txtCriteria.Text & " AND ((Disbursement.CVDate) >= #" & dFrom.Value & "# And (Disbursement.CVDate) <= #" & dTo.Value & "#)) ORDER BY Disbursement.CVDATE DESC , Disbursement.CVNUMBER DESC;"
                    'mySQL
                    MyQry = "SELECT Disbursement.CVDATE, Disbursement.CVNUMBER AS CVNO, Disbursement.DISBURSETYPE, Disbursement.ACCTDESC, Disbursement.PAYEE, Disbursement.PARTICULARS, Disbursement.AMOUNT, Disbursement.DEPARTMENT " & _
                            "From Disbursement " & txtCriteria.Text & " AND ((Disbursement.CVDate) >= '" & Format(dFrom.Value, "yyyy-mm-dd") & "' And (Disbursement.CVDate) <= '" & Format(dTo.Value, "yyyy-mm-dd") & "')) ORDER BY Disbursement.CVDATE DESC , Disbursement.CVNUMBER DESC;"
                End If
            End If
            If dFrom.Value <= DateValue("05/10/2010") Then
                Call DisplayOldReport(MyQry)
            Else
                Call DisplayReport(MyQry)
            End If
        Case 5  'Exit
            'MotoRep.tbMain.Visible = True
            'MotoRep.imgBG.Visible = True
            MotoRep.toolBarButton.Visible = True
            Unload Me
            'Disbursement.Show
    End Select
End Sub

Private Sub dFrom_CloseUp()
    If cmdNext.Tag = "1" Then
        Call RefreshRecord(1)
    End If
End Sub

Private Sub dTo_CloseUp()
    If dTo.Value < dFrom.Value Then
        MsgBox "Please choose a date not less than " & dFrom.Value, vbExclamation + vbOKOnly, "Validation Error"
        dTo.Value = dFrom.Value
    End If
End Sub

Private Sub flxDisb_EnterCell()
    ActRow = flxDisb.Row
End Sub

Private Sub flxDisb_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Integer, NoCV As Boolean
    
    NoCV = True
    For i = 1 To flxCV.Rows - 1
        If flxCV.TextMatrix(i, 3) = "EXPENSE" Then NoCV = False
    Next i
    If Not NoCV Then Exit Sub
    If flxDisb.Rows = 2 Then ActRow = 1
    If Button = vbRightButton And flxDisb.TextMatrix(ActRow, 8) = "Final" And flxDisb.TextMatrix(ActRow, 3) <> "Cancelled Vouchers" Then
        Call PopupMenu(mnuPCV)
    End If
End Sub

Private Sub Form_Load()
    Dim i As Integer
    
    dFrom.Value = Date
    cmdNext.Tag = "1"
    Call RefreshRecord(1)
    ActRow = 1
    MotoRep.sBar.Panels(6).Text = "Disbursement Summary"
    BtnIndx = 1
    tmrFloat.Enabled = True
    Call PrintPreView(False)
End Sub

Private Sub LoadCashVoucher()
    Dim rsCV As New ADODB.Recordset
    Dim MyQry As String

    With rsCV
        If Val(cmdNext.Tag) = 1 Then
            'MyQry = "SELECT CASHVOUCHER.VOUCHERNO, CASHVOUCHER.AMOUNT, CASHVOUCHER.CVDATE, CASHVOUCHER.REMARKS From CASHVOUCHER " & _
                    "WHERE (((CASHVOUCHER.CVDATE)=#" & dFrom.Value & "#) AND ((CASHVOUCHER.REMARKS)<>'CANCELLED'));"
            'mySQL
            MyQry = "SELECT CASHVOUCHER.VOUCHERNO, CASHVOUCHER.AMOUNT, CASHVOUCHER.CVDATE, CASHVOUCHER.REMARKS From CASHVOUCHER " & _
                    "WHERE CASHVOUCHER.CVDATE='" & Format(dFrom.Value, "yyyy-mm-dd") & "' AND CASHVOUCHER.REMARKS<>'CANCELLED';"
            
        Else
            'MyQry = "SELECT CASHVOUCHER.VOUCHERNO, CASHVOUCHER.AMOUNT, CASHVOUCHER.CVDATE, CASHVOUCHER.REMARKS From CashVoucher " & _
                    "WHERE (((CASHVOUCHER.CVDATE)>=#" & dFrom.Value & "# And (CASHVOUCHER.CVDATE)<=#" & dTo.Value & "#) AND ((CASHVOUCHER.REMARKS)<>'CANCELLED'));"
            'mySQL
            MyQry = "SELECT CASHVOUCHER.VOUCHERNO, CASHVOUCHER.AMOUNT, CASHVOUCHER.CVDATE, CASHVOUCHER.REMARKS From CashVoucher " & _
                    "WHERE CASHVOUCHER.CVDATE>='" & Format(dFrom.Value, "yyyy-mm-dd") & "' And CASHVOUCHER.CVDATE<='" & Format(dTo.Value, "yyyy-mm-dd") & "' AND CASHVOUCHER.REMARKS<>'CANCELLED';"
        End If
        .Open MyQry, LYRcon, adOpenKeyset, adLockOptimistic
        flxCV.Rows = 1
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                flxCV.AddItem Format(!VOUCHERNO, "0000") & vbTab & !Amount & vbTab & !CVDate & vbTab & !Remarks
                .MoveNext
            Loop
        End If
        .Close
    End With
End Sub

Private Sub DisplayOldReport(MQ As String)
    Dim crDisb As New rptDisbOld
    Dim Disburser As String, Qry As String, MyCVs As String
    Dim i As Integer, CVamt As Currency
    
    Disburser = GetSettings("DISBURSER")
    If rsRPT.State = adStateOpen Then rsRPT.Close
    rsRPT.Open MQ, LYRcon, adOpenKeyset, adLockOptimistic
    crDisb.DiscardSavedData
    crDisb.Database.SetDataSource rsRPT
    crDisb.EnableParameterPrompting = False
    If Val(cmdNext.Tag) = 1 Then
        crDisb.ParameterFields(1).ClearCurrentValueAndRange
        crDisb.ParameterFields(1).AddCurrentValue Format$(dFrom.Value, "Mmmm dd, yyyy")
    Else
        crDisb.ParameterFields(1).ClearCurrentValueAndRange
        crDisb.ParameterFields(1).AddCurrentValue Format$(dFrom.Value, "Mm/dd/yyyy") & " to " & Format$(dTo.Value, "Mm/dd/yyyy")
    End If
    crDisb.ParameterFields(2).ClearCurrentValueAndRange
    crDisb.ParameterFields(2).AddCurrentValue Disburser
    Select Case flxCV.Rows      'Identify number of vouchers included in report
        Case 1      'No voucher available
            crDisb.ParameterFields(3).ClearCurrentValueAndRange
            crDisb.ParameterFields(3).AddCurrentValue IIf(txtCriteria.Text = "", "Cash Voucher No. [No Available]:", "")
            For i = 4 To 10
                crDisb.ParameterFields(i).ClearCurrentValueAndRange
                crDisb.ParameterFields(i).AddCurrentValue IIf(i Mod 2 = 0, 0, "")
            Next i
        Case 2      '1 voucher available
            crDisb.ParameterFields(3).ClearCurrentValueAndRange
            crDisb.ParameterFields(3).AddCurrentValue "Cash Voucher No. [" & flxCV.TextMatrix(1, 0) & "]:"
            crDisb.ParameterFields(4).ClearCurrentValueAndRange
            crDisb.ParameterFields(4).AddCurrentValue Val(flxCV.TextMatrix(1, 1))
            For i = 5 To 10
                crDisb.ParameterFields(i).ClearCurrentValueAndRange
                crDisb.ParameterFields(i).AddCurrentValue IIf(i Mod 2 = 0, 0, "")
            Next i
        Case 3      '2 voucher available
            crDisb.ParameterFields(3).ClearCurrentValueAndRange
            crDisb.ParameterFields(3).AddCurrentValue "Cash Voucher No. [" & flxCV.TextMatrix(1, 0) & "]:"
            crDisb.ParameterFields(4).ClearCurrentValueAndRange
            crDisb.ParameterFields(4).AddCurrentValue Val(flxCV.TextMatrix(1, 1))
            crDisb.ParameterFields(5).ClearCurrentValueAndRange
            crDisb.ParameterFields(5).AddCurrentValue "Cash Voucher No. [" & flxCV.TextMatrix(2, 0) & "]:"
            crDisb.ParameterFields(6).ClearCurrentValueAndRange
            crDisb.ParameterFields(6).AddCurrentValue Val(flxCV.TextMatrix(2, 1))
            crDisb.ParameterFields(7).ClearCurrentValueAndRange
            crDisb.ParameterFields(7).AddCurrentValue "Total :"
            crDisb.ParameterFields(8).ClearCurrentValueAndRange
            crDisb.ParameterFields(8).AddCurrentValue Val(flxCV.TextMatrix(1, 1)) + Val(flxCV.TextMatrix(2, 1))
            crDisb.ParameterFields(9).ClearCurrentValueAndRange
            crDisb.ParameterFields(9).AddCurrentValue ""
            crDisb.ParameterFields(10).ClearCurrentValueAndRange
            crDisb.ParameterFields(10).AddCurrentValue 0
        Case 4      '3 voucher available
            crDisb.ParameterFields(3).ClearCurrentValueAndRange
            crDisb.ParameterFields(3).AddCurrentValue "Cash Voucher No. [" & flxCV.TextMatrix(1, 0) & "]:"
            crDisb.ParameterFields(4).ClearCurrentValueAndRange
            crDisb.ParameterFields(4).AddCurrentValue Val(flxCV.TextMatrix(1, 1))
            crDisb.ParameterFields(5).ClearCurrentValueAndRange
            crDisb.ParameterFields(5).AddCurrentValue "Cash Voucher No. [" & flxCV.TextMatrix(2, 0) & "]:"
            crDisb.ParameterFields(6).ClearCurrentValueAndRange
            crDisb.ParameterFields(6).AddCurrentValue Val(flxCV.TextMatrix(2, 1))
            crDisb.ParameterFields(7).ClearCurrentValueAndRange
            crDisb.ParameterFields(7).AddCurrentValue "Cash Voucher No. [" & flxCV.TextMatrix(3, 0) & "]:"
            crDisb.ParameterFields(8).ClearCurrentValueAndRange
            crDisb.ParameterFields(8).AddCurrentValue Val(flxCV.TextMatrix(3, 1))
            crDisb.ParameterFields(9).ClearCurrentValueAndRange
            crDisb.ParameterFields(9).AddCurrentValue "Total :"
            crDisb.ParameterFields(10).ClearCurrentValueAndRange
            crDisb.ParameterFields(10).AddCurrentValue Val(flxCV.TextMatrix(1, 1)) + Val(flxCV.TextMatrix(2, 1)) + Val(flxCV.TextMatrix(3, 1))
        Case Else
            If Val(cmdNext.Tag) = 1 Then
                Qry = "SELECT Format([CVNUMBER],'0000') AS CVNO From Disbursement Where (((Disbursement.CVDate) = #" & dFrom.Value & "#)) ORDER BY Format([CVNUMBER],'0000')"
            Else
                Qry = "SELECT Format([CVNUMBER],'0000') AS CVNO From Disbursement Where (((Disbursement.CVDate) >= #" & dFrom.Value & "# And (Disbursement.CVDate) <= #" & dTo.Value & "#))ORDER BY Format([CVNUMBER],'0000');"
            End If
            MyCVs = GetSeries(Qry)
            For i = 1 To flxCV.Rows - 1
                CVamt = CVamt + Val(flxCV.TextMatrix(i, 1))
            Next i
            crDisb.ParameterFields(3).ClearCurrentValueAndRange
            crDisb.ParameterFields(3).AddCurrentValue "CV Nos. [" & MyCVs & "]:"
            crDisb.ParameterFields(4).ClearCurrentValueAndRange
            crDisb.ParameterFields(4).AddCurrentValue Val(CVamt)
            For i = 5 To 10
                crDisb.ParameterFields(i).ClearCurrentValueAndRange
                crDisb.ParameterFields(i).AddCurrentValue IIf(i Mod 2 = 0, 0, "")
            Next i
    End Select
    crDisb.ParameterFields(11).ClearCurrentValueAndRange
    crDisb.ParameterFields(11).AddCurrentValue lblRemarks.Caption
    
    vwrDisb.ReportSource = crDisb
    vwrDisb.EnableExportButton = True
    vwrDisb.ViewReport
End Sub



Private Sub DisplayReport(MQ As String)
    Dim crDisb As New rptDisb
    Dim Disburser As String, Qry As String, MyCVs As String, AFLseries As String
    Dim i As Integer, CVamt As Currency, AdvLiq As Currency
    
    Disburser = GetSettings("DISBURSER")
    If rsRPT.State = adStateOpen Then rsRPT.Close
    
    rsRPT.Open MQ, LYRcon, adOpenKeyset, adLockOptimistic
    crDisb.DiscardSavedData
    crDisb.Database.SetDataSource rsRPT
    crDisb.EnableParameterPrompting = False
    If Val(cmdNext.Tag) = 1 Then
        Qry = "SELECT SUM(advliquidation.AFLAMOUNT) As sumAFLamt From advliquidation " & _
              "Where advliquidation.AFLDATE = '" & Format(dFrom.Value, "yyyy-mm-dd") & "' AND  advliquidation.PRINTCOPY > 0 AND advliquidation.REMARKS <> 'CANCELLED';"
        AdvLiq = Val(GetFieldValue(Qry))
        Qry = "SELECT advliquidation.AFLNO From advliquidation " & _
              "Where advliquidation.AFLDATE = '" & Format(dFrom.Value, "yyyy-mm-dd") & "' AND  advliquidation.PRINTCOPY > 0 AND advliquidation.REMARKS <> 'CANCELLED' Order By advliquidation.AFLNO;"
        AFLseries = GetOR(Qry)
        Qry = "SELECT CONCAT('CV Check:', cashflowmgr.PCVCHECK, ',  CV#:', cashflowmgr.EXPENSECV, ',  CV Amount: ', format(cashflowmgr.EXPENSEAMT, 2)) AS CheckVoucher " & _
              "From CashFlowMgr Where cashflowmgr.CFDATE = '" & Format(dFrom.Value, "yyyy-mm-dd") & "'"
        MyCVs = GetFieldValue(Qry)
    
        crDisb.ParameterFields(1).ClearCurrentValueAndRange
        crDisb.ParameterFields(1).AddCurrentValue Format$(dFrom.Value, "Mmmm dd, yyyy")
    Else
        Qry = "SELECT SUM(advliquidation.AFLAMOUNT) As sumAFLamt From advliquidation " & _
              "Where advliquidation.AFLDATE >= '" & Format(dFrom.Value, "yyyy-mm-dd") & "' AND advliquidation.AFLDATE <= '" & Format(dTo.Value, "yyyy-mm-dd") & "' AND  advliquidation.PRINTCOPY > 0 AND advliquidation.REMARKS <> 'CANCELLED';"
        AdvLiq = Val(GetFieldValue(Qry))
        Qry = "SELECT advliquidation.AFLNO From advliquidation " & _
              "Where advliquidation.AFLDATE >= '" & Format(dFrom.Value, "yyyy-mm-dd") & "' AND advliquidation.AFLDATE <= '" & Format(dTo.Value, "yyyy-mm-dd") & "' AND  advliquidation.PRINTCOPY > 0 AND advliquidation.REMARKS <> 'CANCELLED' Order By advliquidation.AFLNO;"
        AFLseries = GetOR(Qry)
        MyCVs = ""
        crDisb.ParameterFields(1).ClearCurrentValueAndRange
        crDisb.ParameterFields(1).AddCurrentValue Format$(dFrom.Value, "Mm/dd/yyyy") & " to " & Format$(dTo.Value, "Mm/dd/yyyy")
    End If
    crDisb.ParameterFields(2).ClearCurrentValueAndRange
    crDisb.ParameterFields(2).AddCurrentValue Disburser
    crDisb.ParameterFields(3).ClearCurrentValueAndRange
    crDisb.ParameterFields(3).AddCurrentValue "Total Adv. for Liquidation [" & AFLseries & "]"
    crDisb.ParameterFields(4).ClearCurrentValueAndRange
    crDisb.ParameterFields(4).AddCurrentValue Val(AdvLiq)
    crDisb.ParameterFields(5).ClearCurrentValueAndRange
    crDisb.ParameterFields(5).AddCurrentValue MyCVs
    Qry = "SELECT branch.ADDRESS From branch Where branch.SYSTEMBRANCH = 1"
    crDisb.ParameterFields(6).ClearCurrentValueAndRange
    crDisb.ParameterFields(6).AddCurrentValue GetFieldValue(Qry)
    crDisb.ParameterFields(7).ClearCurrentValueAndRange
    crDisb.ParameterFields(7).AddCurrentValue lblRemarks.Caption
    crDisb.ParameterFields(8).ClearCurrentValueAndRange
    crDisb.ParameterFields(8).AddCurrentValue SysBranch
    vwrDisb.ReportSource = crDisb
    vwrDisb.EnableExportButton = True
    vwrDisb.ViewReport
End Sub

Function GetSeries(Qry As String) As String
    Dim rsCashOR As New ADODB.Recordset
    Dim MyBig As String, ORseries As String, LastSer As String
    Dim X As Integer
    
    With rsCashOR
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            ORseries = ""
            X = 1
            Do While Not .EOF
               If X = 1 Then
                    MyBig = .Fields(0)
                    ORseries = .Fields(0)
                Else
                    If Val(.Fields(0)) = Val(MyBig) + 1 Then
                        MyBig = .Fields(0)
                        If Right$(ORseries, 1) <> "-" Then
                            ORseries = ORseries & "-"
                            LastSer = .Fields(0)
                        End If
                    Else
                        If Right$(ORseries, 1) = "-" Then
                            ORseries = ORseries & IIf(Left$(LastSer, 3) <> Left$(Format$(MyBig, "0000"), 3), Format(MyBig, "0000"), Right$(MyBig, 2)) & ", " & .Fields(0)
                            MyBig = .Fields(0)
                        Else
                            MyBig = .Fields(0)
                            ORseries = ORseries & ", " & .Fields(0)
                        End If
                    End If
                End If
                X = X + 1
                .MoveNext
            Loop
            If Right$(ORseries, 1) = "-" Then ORseries = ORseries & MyBig
        End If
        .Close
    End With
    GetSeries = ORseries
End Function

Private Sub PrintPreView(Ppvw As Boolean)
    If Ppvw Then
        Frame1.Visible = False
        flxDisb.Visible = False
        vwrDisb.Visible = True
    Else
        vwrDisb.Visible = False
        Frame1.Visible = True
        flxDisb.Visible = True
    End If
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If rsRPT.State = adStateOpen Then rsRPT.Close
    MotoRep.sBar.Panels(6).Text = ""
End Sub

Private Sub mnuCancel_Click()
    Dim Ans As Integer, MyReason As String, MyAcc As String, Qry As String
    
    
    'Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE (((USERACCESS.EMPID)=" & UserID & ") AND ((USERACCESS.MENUINDX)=59));"
    'mySQL
    Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE USERACCESS.EMPID=" & UserID & " AND USERACCESS.MENUINDX=59;"
    MyAcc = GetFieldValue(Qry)
    If MyAcc <> "FULL ACCESS" Then
        MsgBox "You don't have full access to this transaction." & vbCrLf & "Please contact your System Administrator", vbExclamation + vbOKOnly, "Access Denied"
        Exit Sub
    End If
    
    Ans = MsgBox("Do you want to Cancel Voucher?" & vbCrLf & vbCrLf & _
        "PCV Date : " & flxDisb.TextMatrix(ActRow, 1) & vbCrLf & "Voucher No: " & flxDisb.TextMatrix(ActRow, 2) & vbCrLf & _
        "Received By: " & flxDisb.TextMatrix(ActRow, 5) & vbCrLf & _
        "Account Desc.: " & flxDisb.TextMatrix(ActRow, 4), vbQuestion + vbYesNo, "Cancel")
    If Ans = vbNo Then Exit Sub
    MyReason = InputBox("Please indicate reason for Cancellation of PCV:", "Reason for Cancel")
    If MyReason = "" Then
        MsgBox "Unable to indicate Reason..." & vbCrLf & "Cancellation of PCV failed!", vbExclamation + vbOKOnly, "Cancel Failed"
        Exit Sub
    End If
    'If Disburse type = Refund ... revert all updates
    If InStr(flxDisb.TextMatrix(ActRow, 4), "REFUND") > 0 And InStr(flxDisb.TextMatrix(ActRow, 6), "Sales Return") > 0 Then Call CancelRefund
    'Update Disbursement
    'LYRcon.Execute "UPDATE DISBURSEMENT SET DISBURSEMENT.DISBURSETYPE = 'Cancelled Vouchers', DISBURSEMENT.PARTICULARS = '" & MyReason & "', DISBURSEMENT.AMOUNT = 0 " & _
                   "WHERE (((DISBURSEMENT.CVNUMBER)=" & flxDisb.TextMatrix(ActRow, 2) & "));"
    'mySQL
    LYRcon.Execute "UPDATE DISBURSEMENT SET DISBURSEMENT.DISBURSETYPE = 'Cancelled Vouchers', DISBURSEMENT.PARTICULARS = '" & MyReason & "', DISBURSEMENT.AMOUNT = 0 " & _
                   "WHERE DISBURSEMENT.CVNUMBER=" & Val(Replace(flxDisb.TextMatrix(ActRow, 2), ",", "")) & ";"
    'Update Audit Trans
    'LYRcon.Execute "INSERT INTO AUDITRANS ( TRANSDATE, TRANSTYPE, TRANSDOC, REFNO, REFDATE, REASON, EMPID ) " & _
            "VALUES (#" & Date & "#, 'CANCEL', 'PCV', '" & flxDisb.TextMatrix(ActRow, 2) & "', #" & flxDisb.TextMatrix(ActRow, 1) & "#, '" & MyReason & "', " & UserID & ");"
    'mySQL
    LYRcon.Execute "INSERT INTO AUDITRANS ( TRANSDATE, TRANSTYPE, TRANSDOC, REFNO, REFDATE, REASON, EMPID ) " & _
            "VALUES ('" & Format(Date, "yyyy-mm-dd") & "', 'CANCEL', 'PCV', '" & flxDisb.TextMatrix(ActRow, 2) & "', '" & Format(flxDisb.TextMatrix(ActRow, 1), "yyyy-mm-dd") & "', '" & MyReason & "', " & UserID & ");"
    
    'Update Grid
    flxDisb.TextMatrix(ActRow, 3) = "Cancelled Vouchers"
    flxDisb.TextMatrix(ActRow, 6) = MyReason
    flxDisb.TextMatrix(ActRow, 7) = "0.00"
    MsgBox "Voucher Cancelled!", vbInformation + vbOKOnly
End Sub

Private Sub CancelRefund()
    Dim Qry As String, Chrg As String, RefType As String, RefNo As String, MyInv As String
    Dim RefAmt As Currency, Bal As Currency, CustID As Integer
    
    'Qry = "SELECT DISBURSEMENT.CHARGETO From Disbursement WHERE (((DISBURSEMENT.CVNUMBER)=" & flxDisb.TextMatrix(ActRow, 2) & "));"
    'mySQL
    Qry = "SELECT DISBURSEMENT.CHARGETO From Disbursement WHERE DISBURSEMENT.CVNUMBER=" & flxDisb.TextMatrix(ActRow, 2) & ";"
    Chrg = GetFieldValue(Qry)
    If Chrg = "" Then Exit Sub
    If InStr(Chrg, "SR#") > 0 Then
        RefType = "SALES RETURN"
        RefNo = Right$(Chrg, Len(Chrg) - 3)
        'Qry = "SELECT SRETURNSP.INVOICE From SRETURNSP WHERE (((SRETURNSP.RETNO)='" & RefNo & "'));"
        'mysql
        Qry = "SELECT SRETURNSP.INVOICE From SRETURNSP WHERE SRETURNSP.RETNO='" & RefNo & "';"
        
        MyInv = GetFieldValue(Qry)
    ElseIf InStr(Chrg, "OR#") > 0 Then
        RefType = "O.R."
        RefNo = Right$(Chrg, Len(Chrg) - 3)
        'Qry = "SELECT Ledger.SPINVNO From Ledger WHERE (((Ledger.REFTYPE)='Payment') AND ((Ledger.REFNO)='" & RefNo & "'));"
        'mySQL
        Qry = "SELECT Ledger.SPINVNO From Ledger WHERE Ledger.REFTYPE='Payment' AND Ledger.REFNO='" & RefNo & "';"
    
        MyInv = GetFieldValue(Qry)
    End If
    If MyInv = "" Then Exit Sub
    'Qry = "SELECT LEDGER.SPINVNO, LEDGER.REFTYPE, LEDGER.REFNO From Ledger " & _
          "WHERE (((LEDGER.SPINVNO)='" & MyInv & "') AND ((LEDGER.REFTYPE)='Refund') AND ((LEDGER.REFNO)='" & flxDisb.TextMatrix(ActRow, 2) & "'));"
    'mySQL
    Qry = "SELECT LEDGER.SPINVNO, LEDGER.REFTYPE, LEDGER.REFNO From Ledger " & _
          "WHERE LEDGER.SPINVNO='" & MyInv & "' AND LEDGER.REFTYPE='Refund' AND LEDGER.REFNO='" & flxDisb.TextMatrix(ActRow, 2) & "';"
    
    If CtrlNumExist(Qry) Then
        'LYRcon.Execute "UPDATE Ledger SET Ledger.REFTYPE = 'CANCELLED', Ledger.CUSTOMERID = 0 " & _
                         "WHERE (((Ledger.SPINVNO)='" & MyInv & "') AND ((Ledger.REFNO)='" & flxDisb.TextMatrix(ActRow, 2) & "'));"
        'mySQL
        LYRcon.Execute "UPDATE Ledger SET Ledger.REFTYPE = 'CANCELLED', Ledger.CUSTOMERID = 0 " & _
                         "WHERE Ledger.SPINVNO='" & MyInv & "' AND Ledger.REFNO='" & flxDisb.TextMatrix(ActRow, 2) & "';"
    End If
    'Qry = "SELECT ONACCTERMS.BALANCE From ONACCTERMS WHERE (((ONACCTERMS.SPINVNO)='" & MyInv & "'));"
    'mySQL
    Qry = "SELECT ONACCTERMS.BALANCE From ONACCTERMS WHERE ONACCTERMS.SPINVNO='" & MyInv & "';"
    
    If CtrlNumExist(Qry) Then
        RefAmt = Val(Replace(flxDisb.TextMatrix(ActRow, 7), ",", ""))
        Bal = Val(GetFieldValue(Qry))
        'LYRcon.Execute "UPDATE ONACCTERMS SET ONACCTERMS.BALANCE = " & Bal - RefAmt & ", ONACCTERMS.STATUS = '" & IIf(Bal + RefAmt > 0, "OPEN", "CLOSED") & "' WHERE (((ONACCTERMS.SPINVNO)='" & MyInv & "'));"
        'mySQL
        LYRcon.Execute "UPDATE ONACCTERMS SET ONACCTERMS.BALANCE = " & Bal - RefAmt & ", ONACCTERMS.`STATUS` = '" & IIf(Bal + RefAmt > 0, "OPEN", "CLOSED") & "' WHERE ONACCTERMS.SPINVNO='" & MyInv & "';"
        'Qry = "SELECT ONACCTERMS.CUSTOMERID From ONACCTERMS WHERE (((ONACCTERMS.SPINVNO)='" & MyInv & "'));"
        'mySQL
        Qry = "SELECT ONACCTERMS.CUSTOMERID From ONACCTERMS WHERE ONACCTERMS.SPINVNO='" & MyInv & "';"
        CustID = Val(GetFieldValue(Qry))
        'Update Customer Balance
        'Qry = "SELECT Sum(ONACCTERMS.BALANCE) AS SumOfBALANCE From ONACCTERMS " & _
              "WHERE (((ONACCTERMS.CUSTOMERID)=" & CustID & ") AND ((ONACCTERMS.STATUS)='OPEN'));"
        'MYsql
        Qry = "SELECT Sum(round(ONACCTERMS.BALANCE)) AS SumOfBALANCE From ONACCTERMS " & _
              "WHERE ONACCTERMS.CUSTOMERID=" & CustID & " AND ONACCTERMS.`STATUS`='OPEN';"
        
        Bal = Val(GetFieldValue(Qry))
        'LYRcon.Execute "UPDATE CUSTOMER SET CUSTOMER.OSBAL = " & Bal & " WHERE (((CUSTOMER.CUSTOMERID)=" & CustID & "));"
        'mySQL
        LYRcon.Execute "UPDATE CUSTOMER SET CUSTOMER.OSBAL = " & Bal & " WHERE CUSTOMER.CUSTOMERID=" & CustID & ";"
    End If
    'LYRcon.Execute "UPDATE SRETURNSP SET SRETURNSP.REMARKS = '', SRETURNSP.PROREF = '' WHERE (((SRETURNSP.RETNO)='" & RefNo & "'));"
    'mySQL
    LYRcon.Execute "UPDATE SRETURNSP SET SRETURNSP.REMARKS = '', SRETURNSP.PROREF = '' WHERE SRETURNSP.RETNO='" & RefNo & "';"

End Sub

Function GetOR(Qry As String) As String
    Dim rsCashOR As New ADODB.Recordset
    Dim MyBig As String, ORseries As String, LastSer As String
    Dim X As Integer
    
    With rsCashOR
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            ORseries = ""
            X = 1
            Do While Not .EOF
               If X = 1 Then
                    MyBig = .Fields(0)
                    ORseries = .Fields(0)
                Else
                    If Val(.Fields(0)) = Val(MyBig) + 1 Then
                        MyBig = Val(.Fields(0))
                        If Right$(ORseries, 1) <> "-" Then
                            ORseries = ORseries & "-"
                            LastSer = .Fields(0)
                        End If
                    Else
                        If Right$(ORseries, 1) = "-" Then
                            ORseries = ORseries & IIf(Left$(LastSer, 3) <> Left$(Format$(MyBig, "0000"), 3), Format(MyBig, "0000"), Right$(MyBig, 2)) & ", " & .Fields(0)
                            MyBig = .Fields(0)
                        Else
                            MyBig = .Fields(0)
                            ORseries = ORseries & ", " & .Fields(0)
                        End If
                    End If
                End If
                X = X + 1
                .MoveNext
            Loop
            If Right$(ORseries, 1) = "-" Then ORseries = ORseries & MyBig
        End If
        .Close
    End With
    GetOR = ORseries
End Function



Private Sub RXResizer1_ResizeComplete()
    Dim i As Integer
    
    With flxDisb
        .FormatString = "  |^  PCV Date |^PVC Number |< Disbursement Type |<Account Description  |< Name of Payee |<Particulars |>PCV Amount|<Status"
        .ColWidth(0) = 300
        .ColWidth(1) = (.Width - 400) * 0.1
        .ColWidth(2) = (.Width - 400) * 0.08
        .ColWidth(3) = (.Width - 400) * 0.15
        .ColWidth(4) = (.Width - 400) * 0.2
        .ColWidth(5) = (.Width - 400) * 0.15
        .ColWidth(6) = (.Width - 400) * 0.2
        .ColWidth(7) = (.Width - 400) * 0.07
        .ColWidth(8) = (.Width - 400) * 0.05
        .ColWidth(9) = 0
        For i = 1 To .Rows - 1
            .TextMatrix(i, 1) = Format(.TextMatrix(i, 1), "mm/dd/yyyy")
            .TextMatrix(i, 7) = Format(.TextMatrix(i, 7), "#,##0.00")
            If InStr(UCase$(.TextMatrix(i, 3)), "CANCELLED") > 0 Then 'Cancelled
                .Col = 1: .Row = i: .ColSel = 9: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 100, 100)
                .Col = 2
            End If
            If .TextMatrix(i, 8) = "Temp" Then
                .Col = 0: .Row = i: .CellFontName = "Webdings": .CellFontSize = 10: .Text = lblP.Caption
                .Col = 1: .Row = i: .ColSel = 9: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 250, 200)
                .Col = 2
            End If
        Next i
    End With
    cmdFind.Height = cmdRefresh.Height
    cmdFind.Top = cmdRefresh.Top
    txtFind.Top = cmdFind.Top
    
End Sub

Private Sub tmrDrop_Timer()
    cmdTrans(BtnIndx).Top = cmdTrans(BtnIndx).Top + 50
    If cmdTrans(BtnIndx).Top >= 0 Then
        cmdTrans(BtnIndx).Top = 0
        BtnIndx = BtnIndx + 1
        If BtnIndx = 6 Then     'last index + 1
            tmrDrop.Enabled = False
        End If
    End If
End Sub

Private Sub tmrDropRev_Timer()
    cmdTrans(BtnIndx).Top = cmdTrans(BtnIndx).Top + 50
    If cmdTrans(BtnIndx).Top >= 0 Then
        cmdTrans(BtnIndx).Top = 0
        BtnIndx = BtnIndx + 1
        If BtnIndx = 2 Then
            tmrDropRev.Enabled = False
        End If
    End If
End Sub

Private Sub tmrFloat_Timer()
    Dim i As Integer
    
    cmdTrans(BtnIndx).Top = cmdTrans(BtnIndx).Top - 50
    If cmdTrans(BtnIndx).Top <= -375 Then
        cmdTrans(BtnIndx).Visible = False
        BtnIndx = BtnIndx - 1
        If BtnIndx = -1 Then
            tmrFloat.Enabled = False
            BtnIndx = 2 'Index of Button view grid
            For i = 2 To 5
                cmdTrans(i).Top = -375
                cmdTrans(i).Left = IIf(i = 2, 0, cmdTrans(2).Width * (i - 2))
                cmdTrans(i).Visible = True
            Next i
            tmrDrop.Enabled = True
        End If
    End If
End Sub

Private Sub tmrFloatRev_Timer()
    
    cmdTrans(BtnIndx).Top = cmdTrans(BtnIndx).Top - 50
    If cmdTrans(BtnIndx).Top <= -375 Then
        cmdTrans(BtnIndx).Visible = False
        BtnIndx = BtnIndx - 1
        If BtnIndx = 1 Then 'Index of Button Exit
            tmrFloatRev.Enabled = False
            BtnIndx = 0
            cmdTrans(0).Visible = True
            cmdTrans(1).Visible = True
            tmrDropRev.Enabled = True
        End If
    End If
End Sub

Private Sub txtFind_Change()
    fRow = 1
    fCol = 1
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then Call cmdFind_Click
End Sub

