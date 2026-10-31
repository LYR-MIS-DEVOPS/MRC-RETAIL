VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form FrmSummary 
   Caption         =   "Form Summary"
   ClientHeight    =   8550
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11865
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8550
   ScaleWidth      =   11865
   WindowState     =   2  'Maximized
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
      Top             =   1350
      Width           =   11655
      _ExtentX        =   20558
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
      Width           =   11655
      Begin VB.TextBox txtFind 
         Height          =   285
         Left            =   9600
         TabIndex        =   5
         Tag             =   "XYF"
         Top             =   195
         Width           =   1935
      End
      Begin lvButton.lvButtons_H cmdFind 
         Height          =   375
         Left            =   8280
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
         Image           =   "FrmSummary.frx":0000
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
         Format          =   94699521
         CurrentDate     =   39322
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
         Image           =   "FrmSummary.frx":059A
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
         Image           =   "FrmSummary.frx":0B34
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
         Format          =   94699521
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
         Image           =   "FrmSummary.frx":10CE
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
         TabIndex        =   11
         Top             =   210
         Width           =   1335
      End
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
      Image           =   "FrmSummary.frx":1668
      cBack           =   -2147483633
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
      Image           =   "FrmSummary.frx":1C02
      cBack           =   -2147483633
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
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   360
      Top             =   8640
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   12105
      m_MDI_Height    =   9120
      OriginalBorder  =   10
      OriginalCaption =   31
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
      Left            =   7200
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
      Left            =   7215
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
End
Attribute VB_Name = "FrmSummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rssummary As New ADODB.Recordset
Dim ActRow As Integer, fRow As Integer, fCol As Integer, TCol As Integer, AccID As Integer, ColPrn As Integer
Dim Qry As String
Public mySrc As String

Public Sub RefreshRecord(dCode As Integer)
    Dim i As Integer
    
    If mySrc = "Job Order" Then
        If dCode = 1 Then
            '***Access
            'Qry = "SELECT SRVCJOBORDER.DATEPROMISED, SRVCJOBORDER.JOBORDERNO, SRVCJOBORDER.CUSTOMER, SRVCUSTOMER.ADDRESS, SRVCJOBORDER.PAYTYPE, SRVCJOBORDER.BALANCE, SRVCJOBORDER.ORNO, SRVCUSTOMER.MODEL, SRVCJOBORDER.KMREADING, SRVCJOBORDER.JOTYPE, SRVCJOBORDER.MECHANIC, SRVCJOBORDER.PRINTCOPY, SRVCJOBORDER.REMARKS, SRVCJOBORDER.JOBORDERID " & _
                  "FROM SRVCJOBORDER INNER JOIN SRVCUSTOMER ON SRVCJOBORDER.ACCOUNTNO = SRVCUSTOMER.ACCOUNTNO Where (((SRVCJOBORDER.DATEPROMISED) = #" & dFrom.Value & "#)) ORDER BY SRVCJOBORDER.DATEPROMISED, SRVCJOBORDER.JOBORDERNO;"
            'mySQL
            Qry = "SELECT SRVCJOBORDER.DATEPROMISED, SRVCJOBORDER.JOBORDERNO, SRVCJOBORDER.CUSTOMER, SRVCUSTOMER.ADDRESS, SRVCJOBORDER.PAYTYPE, SRVCJOBORDER.BALANCE, SRVCJOBORDER.ORNO, SRVCUSTOMER.MODEL, SRVCJOBORDER.KMREADING, SRVCJOBORDER.JOTYPE, SRVCJOBORDER.MECHANIC, SRVCJOBORDER.PRINTCOPY, SRVCJOBORDER.REMARKS, SRVCJOBORDER.JOBORDERID " & _
                  "FROM SRVCJOBORDER INNER JOIN SRVCUSTOMER ON SRVCJOBORDER.ACCOUNTNO = SRVCUSTOMER.ACCOUNTNO Where SRVCJOBORDER.DATEPROMISED = '" & Format(dFrom.Value, "yyyy-mm-dd") & "' ORDER BY SRVCJOBORDER.DATEPROMISED, SRVCJOBORDER.JOBORDERNO;"
            
        Else
            '***Access
            'Qry = "SELECT SRVCJOBORDER.DATEPROMISED, SRVCJOBORDER.JOBORDERNO, SRVCJOBORDER.CUSTOMER, SRVCUSTOMER.ADDRESS, SRVCJOBORDER.PAYTYPE, SRVCJOBORDER.BALANCE, SRVCJOBORDER.ORNO, SRVCUSTOMER.MODEL, SRVCJOBORDER.KMREADING, SRVCJOBORDER.JOTYPE, SRVCJOBORDER.MECHANIC, SRVCJOBORDER.PRINTCOPY, SRVCJOBORDER.REMARKS, SRVCJOBORDER.JOBORDERID " & _
                  "FROM SRVCJOBORDER INNER JOIN SRVCUSTOMER ON SRVCJOBORDER.ACCOUNTNO = SRVCUSTOMER.ACCOUNTNO Where (((SRVCJOBORDER.DATEPROMISED) >= #" & dFrom.Value & "# And (SRVCJOBORDER.DATEPROMISED) <= #" & dTo.Value & "#)) ORDER BY SRVCJOBORDER.DATEPROMISED, SRVCJOBORDER.JOBORDERNO;"
            'mySQL
            Qry = "SELECT SRVCJOBORDER.DATEPROMISED, SRVCJOBORDER.JOBORDERNO, SRVCJOBORDER.CUSTOMER, SRVCUSTOMER.ADDRESS, SRVCJOBORDER.PAYTYPE, SRVCJOBORDER.BALANCE, SRVCJOBORDER.ORNO, SRVCUSTOMER.MODEL, SRVCJOBORDER.KMREADING, SRVCJOBORDER.JOTYPE, SRVCJOBORDER.MECHANIC, SRVCJOBORDER.PRINTCOPY, SRVCJOBORDER.REMARKS, SRVCJOBORDER.JOBORDERID " & _
                  "FROM SRVCJOBORDER INNER JOIN SRVCUSTOMER ON SRVCJOBORDER.ACCOUNTNO = SRVCUSTOMER.ACCOUNTNO Where SRVCJOBORDER.DATEPROMISED >= '" & Format(dFrom.Value, "yyyy-mm-dd") & "' And SRVCJOBORDER.DATEPROMISED <= '" & Format(dTo.Value, "yyyy-mm-dd") & "' ORDER BY SRVCJOBORDER.DATEPROMISED, SRVCJOBORDER.JOBORDERNO;"
        
        End If
    ElseIf mySrc = "Clerical Adjustment" Then
        If dCode = 1 Then
            '***Access
            'Qry = "SELECT ADJCLERICAL.ADJDATE, ADJCLERICAL.ADJNO, ADJCLERICAL.ITEMADJUSTED, ADJCLERICAL.FIELDADJUSTED, ADJCLERICAL.REFNO, ADJCLERICAL.OLDVALUE, ADJCLERICAL.NEWVALUE, ADJCLERICAL.REASON, ADJCLERICAL.REMARKS, ADJCLERICAL.ENCODER, ADJCLERICAL.PRINTCOPY " & _
                  "From ADJCLERICAL Where (((ADJCLERICAL.ADJDATE) = #" & dFrom.Value & "#)) ORDER BY ADJCLERICAL.ADJDATE, ADJCLERICAL.ADJNO DESC;"
            'mySQL
            Qry = "SELECT ADJCLERICAL.ADJDATE, ADJCLERICAL.ADJNO, ADJCLERICAL.ITEMADJUSTED, ADJCLERICAL.FIELDADJUSTED, ADJCLERICAL.REFNO, ADJCLERICAL.OLDVALUE, ADJCLERICAL.NEWVALUE, ADJCLERICAL.REASON, ADJCLERICAL.REMARKS, ADJCLERICAL.ENCODER, ADJCLERICAL.PRINTCOPY " & _
                  "From ADJCLERICAL Where ADJCLERICAL.ADJDATE = '" & Format(dFrom.Value, "yyyy-mm-dd") & "' ORDER BY ADJCLERICAL.ADJDATE, ADJCLERICAL.ADJNO DESC;"
            
        Else
            '***Access
            'Qry = "SELECT ADJCLERICAL.ADJDATE, ADJCLERICAL.ADJNO, ADJCLERICAL.ITEMADJUSTED, ADJCLERICAL.FIELDADJUSTED, ADJCLERICAL.REFNO, ADJCLERICAL.OLDVALUE, ADJCLERICAL.NEWVALUE, ADJCLERICAL.REASON, ADJCLERICAL.REMARKS, ADJCLERICAL.ENCODER, ADJCLERICAL.PRINTCOPY " & _
                  "From ADJCLERICAL Where (((ADJCLERICAL.ADJDATE) >= #" & dFrom.Value & "# And (ADJCLERICAL.ADJDATE) <= #" & dTo.Value & "#)) ORDER BY ADJCLERICAL.ADJDATE, ADJCLERICAL.ADJNO DESC;"
            'mySQL
            Qry = "SELECT ADJCLERICAL.ADJDATE, ADJCLERICAL.ADJNO, ADJCLERICAL.ITEMADJUSTED, ADJCLERICAL.FIELDADJUSTED, ADJCLERICAL.REFNO, ADJCLERICAL.OLDVALUE, ADJCLERICAL.NEWVALUE, ADJCLERICAL.REASON, ADJCLERICAL.REMARKS, ADJCLERICAL.ENCODER, ADJCLERICAL.PRINTCOPY " & _
                  "From ADJCLERICAL Where ADJCLERICAL.ADJDATE >= '" & Format(dFrom.Value, "yyyy-mm-dd") & "' And ADJCLERICAL.ADJDATE <= '" & Format(dTo.Value, "yyyy-mm-dd") & "' ORDER BY ADJCLERICAL.ADJDATE, ADJCLERICAL.ADJNO DESC;"
        
        End If
    ElseIf mySrc = "DMCM [A/R]" Then
        If dCode = 1 Then
            '***Access
            'Qry = "SELECT ADJDMCM.ADJDATE, ADJDMCM.ADJNO, ADJDMCM.CUSTOMER, ADJDMCM.TRANSTYPE, ADJDMCM.REFNO, ADJDMCM.DEBIT, ADJDMCM.CREDIT, ADJDMCM.BALANCE, ADJDMCM.STATUS, ADJDMCM.REASON, ADJDMCM.PRINTCOPY From ADJDMCM " & _
                  "Where (((ADJDMCM.ADJDATE) = #" & dFrom.Value & "#)) ORDER BY ADJDMCM.ADJDATE DESC , ADJDMCM.ADJNO DESC;"
            'mySQL
            'Qry = "SELECT ADJDMCM.ADJDATE, ADJDMCM.ADJNO, ADJDMCM.CUSTOMER, ADJDMCM.TRANSTYPE, ADJDMCM.REFNO, ADJDMCM.DEBIT, ADJDMCM.CREDIT, ADJDMCM.BALANCE, ADJDMCM.STATUS, ADJDMCM.REASON, ADJDMCM.PRINTCOPY From ADJDMCM " & _
                  "Where ADJDMCM.ADJDATE = '" & Format(dFrom.Value, "yyyy-mm-dd") & "' ORDER BY ADJDMCM.ADJDATE DESC , ADJDMCM.ADJNO DESC;"
            Qry = "SELECT " & _
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
            'Qry = "SELECT ADJDMCM.ADJDATE, ADJDMCM.ADJNO, ADJDMCM.CUSTOMER, ADJDMCM.TRANSTYPE, ADJDMCM.REFNO, ADJDMCM.DEBIT, ADJDMCM.CREDIT, ADJDMCM.BALANCE, ADJDMCM.STATUS, ADJDMCM.REASON, ADJDMCM.PRINTCOPY From ADJDMCM " & _
                  "Where ADJDMCM.ADJDATE >= '" & Format(dFrom.Value, "yyyy-mm-dd") & "' And ADJDMCM.ADJDATE <= '" & Format(dTo.Value, "yyyy-mm-dd") & "' ORDER BY ADJDMCM.ADJDATE DESC , ADJDMCM.ADJNO DESC;"
            'mysql
            Qry = "SELECT " & _
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
                  "From InvntryADJ Where InvntryADJ.ADJDATE = '" & Format(dFrom.Value, "yyy-mm-dd") & "' ORDER BY INVNTRYADJ.ADJDATE, INVNTRYADJ.ADJNO;"
            
        Else
            '***Access
            'Qry = "SELECT INVNTRYADJ.ADJDATE, INVNTRYADJ.ADJNO, INVNTRYADJ.DOCREF, INVNTRYADJ.REASON, INVNTRYADJ.REMARKS, INVNTRYADJ.PREPBY, INVNTRYADJ.APRVBY, INVNTRYADJ.PRINTCOPY " & _
                  "From InvntryADJ Where (((InvntryADJ.ADJDATE) >= #" & dFrom.Value & "# And (InvntryADJ.ADJDATE) <= #" & dTo.Value & "#)) ORDER BY INVNTRYADJ.ADJDATE, INVNTRYADJ.ADJNO;"
            'mySQL
            Qry = "SELECT INVNTRYADJ.ADJDATE, INVNTRYADJ.ADJNO, INVNTRYADJ.DOCREF, INVNTRYADJ.REASON, INVNTRYADJ.REMARKS, INVNTRYADJ.PREPBY, INVNTRYADJ.APRVBY, INVNTRYADJ.PRINTCOPY " & _
                  "From InvntryADJ Where InvntryADJ.ADJDATE >= '" & Format(dFrom.Value, "yyyy-mm-dd") & "' And InvntryADJ.ADJDATE <= '" & Format(dTo.Value, "yyyy-mm-dd") & "' ORDER BY INVNTRYADJ.ADJDATE, INVNTRYADJ.ADJNO;"
            
            
        End If
    ElseIf mySrc = "Customer P.O." Then
        If dCode = 1 Then
            '***Access
            'Qry = "SELECT CUSTPO.CPODATE, CUSTPO.CPONO, CUSTPO.CUSTOMER, IIf([APPROVED]=True,'APPROVED','PENDING') AS STATUS, CUSTPO.INVOICENO, CUSTPO.REMARKS, CUSTPO.PRINTCOPY, CUSTPO.CPOID " & _
                  "From CUSTPO Where (((CUSTPO.CPODATE) = #" & dFrom.Value & "#)) ORDER BY CUSTPO.CPODATE DESC , CUSTPO.CPONO DESC;"
            'mySQL
            Qry = "SELECT CUSTPO.CPODATE, CUSTPO.CPONO, CUSTPO.CUSTOMER, If(CUSTPO.APPROVED=1,'APPROVED','PENDING') AS STATUS, CUSTPO.INVOICENO, CUSTPO.REMARKS, CUSTPO.PRINTCOPY, CUSTPO.CPOID " & _
                  "From CUSTPO Where CUSTPO.CPODATE = '" & Format(dFrom.Value, "yyyy-mm-dd") & "' ORDER BY CUSTPO.CPODATE DESC , CUSTPO.CPONO DESC;"
        
        Else
            '***Access
            'Qry = "SELECT CUSTPO.CPODATE, CUSTPO.CPONO, CUSTPO.CUSTOMER, IIf([APPROVED]=True,'APPROVED','PENDING') AS STATUS, CUSTPO.INVOICENO, CUSTPO.REMARKS, CUSTPO.PRINTCOPY, CUSTPO.CPOID " & _
                  "From CUSTPO Where (((CUSTPO.CPODATE) >= #" & dFrom.Value & "# And (CUSTPO.CPODATE) <= #" & dTo.Value & "#)) ORDER BY CUSTPO.CPODATE DESC , CUSTPO.CPONO DESC;"
            'mySQL
            Qry = "SELECT CUSTPO.CPODATE, CUSTPO.CPONO, CUSTPO.CUSTOMER, If(CUSTPO.APPROVED=1,'APPROVED','PENDING') AS STATUS, CUSTPO.INVOICENO, CUSTPO.REMARKS, CUSTPO.PRINTCOPY, CUSTPO.CPOID " & _
                  "From CUSTPO Where CUSTPO.CPODATE >= '" & Format(dFrom.Value, "yyyy-mm-dd") & "' And CUSTPO.CPODATE <= '" & Format(dTo.Value, "yyyy-mm-dd") & "' ORDER BY CUSTPO.CPODATE DESC , CUSTPO.CPONO DESC;"
            
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
                  "FROM PURCHASERET INNER JOIN SPRECEIVE ON PURCHASERET.PACKINGNO = SPRECEIVE.PACKINGNO WHERE PURCHASERET.DATERET='" & Format(dFrom.Value, "yyyy-mm-dd") & "' ORDER BY PURCHASERET.PURTRNO DESC;"
     
        Else
            '***Access
            'Qry = "SELECT PURCHASERET.DATERET, PURCHASERET.PURTRNO, PURCHASERET.PACKINGNO, SPRECEIVE.DATEREC, SPRECEIVE.SUPPLIER, PURCHASERET.AMOUNT, PURCHASERET.REMARKS, PURCHASERET.PRINTCOPY, PURCHASERET.PURTRNID " & _
                  "FROM PURCHASERET INNER JOIN SPRECEIVE ON PURCHASERET.PACKINGNO = SPRECEIVE.PACKINGNO Where (((PURCHASERET.DATERET) >= #" & dFrom.Value & "# And (PURCHASERET.DATERET) <= #" & dTo.Value & "#)) ORDER BY PURCHASERET.PURTRNO DESC;"
            Qry = "SELECT PURCHASERET.DATERET, PURCHASERET.PURTRNO, PURCHASERET.PACKINGNO, SPRECEIVE.DATEREC, SPRECEIVE.SUPPLIER, PURCHASERET.AMOUNT, PURCHASERET.REMARKS, PURCHASERET.PRINTCOPY, PURCHASERET.PURTRNID " & _
                  "FROM PURCHASERET INNER JOIN SPRECEIVE ON PURCHASERET.PACKINGNO = SPRECEIVE.PACKINGNO Where PURCHASERET.DATERET >= '" & Format(dFrom.Value, "yyyy-mm-dd") & "' And PURCHASERET.DATERET <= '" & Format(dTo.Value, "yyyy-mm-dd") & "' ORDER BY PURCHASERET.PURTRNO DESC;"
        
        End If
    ElseIf mySrc = "Requisition" Then
        If dCode = 1 Then
            '***Access
            'Qry = "SELECT REQUISITION.REQDATE, REQUISITION.REQNO, REQUISITION.CUSTOMER, REQUISITION.AMOUNT, REQUISITION.JOBNO, REQUISITION.INVOICENO, REQUISITION.ORNO, REQUISITION.REMARKS, REQUISITION.PRINTCOPY, REQUISITION.REQID " & _
                  "From Requisition WHERE (((REQUISITION.REQDATE)=#" & dFrom.Value & "#));"
            'mySQL
            Qry = "SELECT REQUISITION.REQDATE, REQUISITION.REQNO, REQUISITION.CUSTOMER, REQUISITION.AMOUNT, REQUISITION.JOBNO, REQUISITION.INVOICENO, REQUISITION.ORNO, REQUISITION.REMARKS, REQUISITION.PRINTCOPY, REQUISITION.REQID " & _
                  "From Requisition WHERE REQUISITION.REQDATE='" & Format(dFrom.Value, "yyyy-mm-dd") & "';"
            
        Else
            '***Access
            'Qry = "SELECT REQUISITION.REQDATE, REQUISITION.REQNO, REQUISITION.CUSTOMER, REQUISITION.AMOUNT, REQUISITION.JOBNO, REQUISITION.INVOICENO, REQUISITION.ORNO, REQUISITION.REMARKS, REQUISITION.PRINTCOPY, REQUISITION.REQID " & _
                  "From Requisition WHERE (((REQUISITION.REQDATE)>=#" & dFrom.Value & "# And (REQUISITION.REQDATE)<=#" & dTo.Value & "#));"
        
            'mySQL
            Qry = "SELECT REQUISITION.REQDATE, REQUISITION.REQNO, REQUISITION.CUSTOMER, REQUISITION.AMOUNT, REQUISITION.JOBNO, REQUISITION.INVOICENO, REQUISITION.ORNO, REQUISITION.REMARKS, REQUISITION.PRINTCOPY, REQUISITION.REQID " & _
                  "From Requisition WHERE REQUISITION.REQDATE>='" & Format(dFrom.Value, "yyyy-mm-dd") & "' And REQUISITION.REQDATE<='" & Format(dTo.Value, "yyyy-mm-dd") & "';"
        End If
    ElseIf mySrc = "Location Adjustment" Then
        If dCode = 1 Then
            '***Access
            'Qry = "SELECT TRANSFERLOC.DATETRANS, TRANSFERLOC.TRANSNO, TRANSFERLOC.REFERENCE, TRANSFERLOC.REASON, TRANSFERLOC.PREPBY, TRANSFERLOC.APRVBY, TRANSFERLOC.PRINTCOPY, TRANSFERLOC.TRANSLOCID " & _
                  "From TRANSFERLOC Where (((TRANSFERLOC.DATETRANS) = #" & dFrom.Value & "#)) ORDER BY TRANSFERLOC.DATETRANS DESC , TRANSFERLOC.TRANSNO DESC;"
            'mySQL
            Qry = "SELECT TRANSFERLOC.DATETRANS, TRANSFERLOC.TRANSNO, TRANSFERLOC.REFERENCE, TRANSFERLOC.REASON, TRANSFERLOC.PREPBY, TRANSFERLOC.APRVBY, TRANSFERLOC.PRINTCOPY, TRANSFERLOC.TRANSLOCID " & _
                  "From TRANSFERLOC Where TRANSFERLOC.DATETRANS = '" & Format(dFrom.Value, "yyyy-mm-dd") & "' ORDER BY TRANSFERLOC.DATETRANS DESC , TRANSFERLOC.TRANSNO DESC;"
        Else
            '***Access
            'Qry = "SELECT TRANSFERLOC.DATETRANS, TRANSFERLOC.TRANSNO, TRANSFERLOC.REFERENCE, TRANSFERLOC.REASON, TRANSFERLOC.PREPBY, TRANSFERLOC.APRVBY, TRANSFERLOC.PRINTCOPY, TRANSFERLOC.TRANSLOCID " & _
                  "From TRANSFERLOC Where (((TRANSFERLOC.DATETRANS) >= #" & dFrom.Value & "# And (TRANSFERLOC.DATETRANS) <= #" & dTo.Value & "#)) ORDER BY TRANSFERLOC.DATETRANS DESC , TRANSFERLOC.TRANSNO DESC;"
            'mySQL
            Qry = "SELECT TRANSFERLOC.DATETRANS, TRANSFERLOC.TRANSNO, TRANSFERLOC.REFERENCE, TRANSFERLOC.REASON, TRANSFERLOC.PREPBY, TRANSFERLOC.APRVBY, TRANSFERLOC.PRINTCOPY, TRANSFERLOC.TRANSLOCID " & _
                  "From TRANSFERLOC Where TRANSFERLOC.DATETRANS >= '" & Format(dFrom.Value, "yyyy-mm-dd") & "' And TRANSFERLOC.DATETRANS <= '" & Format(dTo.Value, "yyyy-mm-dd") & "' ORDER BY TRANSFERLOC.DATETRANS DESC , TRANSFERLOC.TRANSNO DESC;"
            
        End If
    ElseIf mySrc = "Received Spare Parts" Then
        If dCode = 1 Then
            '***Access
            'Qry = "SELECT SPRECEIVE.DATEREC, SPRECEIVE.PACKINGNO, SPRECEIVE.SUPPLIER, SPRECEIVE.SHIPPER, SPRECEIVE.WAYBILLNO, SPRECEIVE.PACKDATE, SPRECEIVE.REMARKS, SPRECEIVE.CHECKEDBY, SPRECEIVE.PRINTCOPY, SPRECEIVE.SPRRNO " & _
                  "From SPRECEIVE Where (((SPRECEIVE.DATEREC) = #" & dFrom.Value & "#)) ORDER BY SPRECEIVE.DATEREC, SPRECEIVE.PACKINGNO DESC;"
            'mySQl
            Qry = "SELECT SPRECEIVE.DATEREC, SPRECEIVE.PACKINGNO, SPRECEIVE.SUPPLIER, SPRECEIVE.SHIPPER, SPRECEIVE.WAYBILLNO, SPRECEIVE.PACKDATE, SPRECEIVE.REMARKS, SPRECEIVE.CHECKEDBY, SPRECEIVE.PRINTCOPY, SPRECEIVE.SPRRNO " & _
                  "From SPRECEIVE Where SPRECEIVE.DATEREC = '" & Format(dFrom.Value, "yyyy-mm-dd") & "' ORDER BY SPRECEIVE.DATEREC, SPRECEIVE.PACKINGNO DESC;"

        Else
            '***Access
            'Qry = "SELECT SPRECEIVE.DATEREC, SPRECEIVE.PACKINGNO, SPRECEIVE.SUPPLIER, SPRECEIVE.SHIPPER, SPRECEIVE.WAYBILLNO, SPRECEIVE.PACKDATE, SPRECEIVE.REMARKS, SPRECEIVE.CHECKEDBY, SPRECEIVE.PRINTCOPY, SPRECEIVE.SPRRNO " & _
                  "From SPRECEIVE Where (((SPRECEIVE.DATEREC) >= #" & dFrom.Value & "# And (SPRECEIVE.DATEREC) <= #" & dTo.Value & "#)) ORDER BY SPRECEIVE.DATEREC, SPRECEIVE.PACKINGNO DESC;"
            'mySQL
            Qry = "SELECT SPRECEIVE.DATEREC, SPRECEIVE.PACKINGNO, SPRECEIVE.SUPPLIER, SPRECEIVE.SHIPPER, SPRECEIVE.WAYBILLNO, SPRECEIVE.PACKDATE, SPRECEIVE.REMARKS, SPRECEIVE.CHECKEDBY, SPRECEIVE.PRINTCOPY, SPRECEIVE.SPRRNO " & _
                  "From SPRECEIVE Where SPRECEIVE.DATEREC >= '" & Format(dFrom.Value, "yyyy-mm-dd") & "' And SPRECEIVE.DATEREC <= '" & Format(dTo.Value, "yyyy-mm-dd") & "' ORDER BY SPRECEIVE.DATEREC, SPRECEIVE.PACKINGNO DESC;"
        
        End If
    ElseIf mySrc = "Cashier Cash Flow" Then
        If dCode = 1 Then
            Qry = "SELECT DATE_FORMAT(cashflow.CFDATE,'%c/%d/%Y'), cashflow.PREPAREDBY, cashflow.APPROVEDBY, cashflow.REMARKS, cashflow.CASHFLOWID " & _
                  "From cashflow Where cashflow.CFDATE = '" & Format(dFrom.Value, "YYYY-MM-DD") & "' Order By cashflow.CFDATE Desc;"
        Else
            Qry = "SELECT DATE_FORMAT(cashflow.CFDATE,'%c/%d/%Y'), cashflow.PREPAREDBY, cashflow.APPROVEDBY, cashflow.REMARKS, cashflow.CASHFLOWID " & _
                  "From cashflow Where cashflow.CFDATE >= '" & Format(dFrom.Value, "YYYY-MM-DD") & "' and cashflow.CFDATE <= '" & Format(dTo.Value, "YYYY-MM-DD") & "' Order By cashflow.CFDATE Desc;"
        End If
    
    ElseIf mySrc = "Manager's Cash Flow" Then
        If dCode = 1 Then
            Qry = "SELECT DATE_FORMAT(cashflowmgr.CFDATE,'%c/%d/%Y'), cashflowmgr.PREPAREDBY, cashflowmgr.APPROVEDBY, cashflowmgr.REMARKS, cashflowmgr.MGRCFID " & _
                  "From cashflowmgr Where cashflowmgr.CFDATE = '" & Format(dFrom.Value, "YYYY-MM-DD") & "' Order By cashflowmgr.CFDATE Desc;"
        Else
            Qry = "SELECT DATE_FORMAT(cashflowmgr.CFDATE,'%c/%d/%Y'), cashflowmgr.PREPAREDBY, cashflowmgr.APPROVEDBY, cashflowmgr.REMARKS, cashflowmgr.MGRCFID " & _
                  "From cashflowmgr Where cashflowmgr.CFDATE >= '" & Format(dFrom.Value, "YYYY-MM-DD") & "' and cashflowmgr.CFDATE <= '" & Format(dTo.Value, "YYYY-MM-DD") & "' Order By cashflowmgr.CFDATE Desc;"
        End If
    
    ElseIf mySrc = "Sales Order" Then
        If dCode = 1 Then
            '***Access
            'Qry = "SELECT SPSALESORDER.ORDERDATE, SPSALESORDER.ORDERNO, SPSALESORDER.FRONTLINER, SPSALESORDER.TOTALPRICE, SPSALESORDER.INVOICENO, SPSALESORDER.REFERENCE, SPSALESORDER.SPSORDERID " & _
                  "From SPSALESORDER Where (((SPSALESORDER.ORDERDATE) = #" & dFrom.Value & "#)) ORDER BY SPSALESORDER.ORDERNO DESC;"
            'mySQL
            Qry = "SELECT SPSALESORDER.ORDERDATE, SPSALESORDER.ORDERNO, SPSALESORDER.FRONTLINER, SPSALESORDER.TOTALPRICE, SPSALESORDER.INVOICENO, SPSALESORDER.REFERENCE, SPSALESORDER.SPSORDERID " & _
                  "From SPSALESORDER Where SPSALESORDER.ORDERDATE = '" & Format(dFrom.Value, "yyyy-mm-dd") & "' ORDER BY SPSALESORDER.ORDERNO DESC;"
        
        Else
            '***Access
            'Qry = "SELECT SPSALESORDER.ORDERDATE, SPSALESORDER.ORDERNO, SPSALESORDER.FRONTLINER, SPSALESORDER.TOTALPRICE, SPSALESORDER.INVOICENO, SPSALESORDER.REFERENCE, SPSALESORDER.SPSORDERID " & _
                  "From SPSALESORDER Where (((SPSALESORDER.ORDERDATE) >= #" & dFrom.Value & "# And (SPSALESORDER.ORDERDATE) <= #" & dTo.Value & "#)) ORDER BY SPSALESORDER.ORDERNO DESC;"
            'mySQL
            Qry = "SELECT SPSALESORDER.ORDERDATE, SPSALESORDER.ORDERNO, SPSALESORDER.FRONTLINER, SPSALESORDER.TOTALPRICE, SPSALESORDER.INVOICENO, SPSALESORDER.REFERENCE, SPSALESORDER.SPSORDERID " & _
                  "From SPSALESORDER Where SPSALESORDER.ORDERDATE >= '" & Format(dFrom.Value, "yyyy-mm-dd") & "' And SPSALESORDER.ORDERDATE <= '" & Format(dTo.Value, "yyyy-mm-dd") & "' ORDER BY SPSALESORDER.ORDERNO DESC;"
            
        End If
    ElseIf mySrc = "Mechanic Income" Then
        If dCode = 1 Then
            '***Access
            'Qry = "SELECT SRVCJOBORDER.MECHANIC, SRVCJOBORDER.JOBORDERNO, SRVCJOBORDER.DATEPROMISED, SRVCJOBORDER.JOTYPE, SRVCJOBORDER.TOTALCOST From SRVCJOBORDER " & _
                  "Where (((SRVCJOBORDER.DATEPROMISED) = #" & dFrom.Value & "#) And ((SRVCJOBORDER.PRINTCOPY) > 0)) ORDER BY SRVCJOBORDER.MECHANIC, SRVCJOBORDER.DATEPROMISED;"
            'mySQl
            Qry = "SELECT SRVCJOBORDER.MECHANIC, SRVCJOBORDER.JOBORDERNO, SRVCJOBORDER.DATEPROMISED, SRVCJOBORDER.JOTYPE, SRVCJOBORDER.TOTALCOST From SRVCJOBORDER " & _
                  "Where SRVCJOBORDER.DATEPROMISED = '" & Format(dFrom.Value, "yyyy-mm-dd") & "' And SRVCJOBORDER.PRINTCOPY > 0 And SRVCJOBORDER.REMARKS <> 'CANCELLED' ORDER BY SRVCJOBORDER.MECHANIC, SRVCJOBORDER.DATEPROMISED;"
            
        Else
            '***Access
            'Qry = "SELECT SRVCJOBORDER.MECHANIC, SRVCJOBORDER.JOBORDERNO, SRVCJOBORDER.DATEPROMISED, SRVCJOBORDER.JOTYPE, SRVCJOBORDER.TOTALCOST From SRVCJOBORDER " & _
                  "Where (((SRVCJOBORDER.DATEPROMISED) >= #" & dFrom.Value & "# And (SRVCJOBORDER.DATEPROMISED) <= #" & dTo.Value & "#) And ((SRVCJOBORDER.PRINTCOPY) > 0)) ORDER BY SRVCJOBORDER.MECHANIC, SRVCJOBORDER.DATEPROMISED;"
            'mySQL
            Qry = "SELECT SRVCJOBORDER.MECHANIC, SRVCJOBORDER.JOBORDERNO, SRVCJOBORDER.DATEPROMISED, SRVCJOBORDER.JOTYPE, SRVCJOBORDER.TOTALCOST From SRVCJOBORDER " & _
                  "Where SRVCJOBORDER.DATEPROMISED >= '" & Format(dFrom.Value, "yyyy-mm-dd") & "' And SRVCJOBORDER.DATEPROMISED <= '" & Format(dTo.Value, "yyyy-mm-dd") & "' And SRVCJOBORDER.PRINTCOPY > 0 And SRVCJOBORDER.REMARKS <> 'CANCELLED' ORDER BY SRVCJOBORDER.MECHANIC, SRVCJOBORDER.DATEPROMISED;"
            
        End If
    ElseIf mySrc = "Cash Voucher" Then
        If dCode = 1 Then
            '***Access
            'Qry = "SELECT CASHVOUCHER.CVDATE, CASHVOUCHER.VOUCHERNO, CASHVOUCHER.PAIDTO, CASHVOUCHER.POSITION, CASHVOUCHER.PARTICULAR, CASHVOUCHER.AMOUNT, CASHVOUCHER.REMARKS, CASHVOUCHER.PRINTCOPY " & _
                  "From CashVoucher Where (((CashVoucher.CVDate) = #" & dFrom & "#)) ORDER BY CASHVOUCHER.CVDATE DESC , CASHVOUCHER.VOUCHERNO;"
            'mySQL
            Qry = "SELECT CASHVOUCHER.CVDATE, CASHVOUCHER.VOUCHERNO, CASHVOUCHER.PAIDTO, CASHVOUCHER.POSITION, CASHVOUCHER.PARTICULAR, CASHVOUCHER.AMOUNT, CASHVOUCHER.REMARKS, CASHVOUCHER.PRINTCOPY " & _
                  "From CashVoucher Where (((CashVoucher.CVDate) = '" & Format(dFrom, "yyyy-mm-dd") & "')) ORDER BY CASHVOUCHER.CVDATE DESC , CASHVOUCHER.VOUCHERNO;"
            
        Else
            '***Access
            'Qry = "SELECT CASHVOUCHER.CVDATE, CASHVOUCHER.VOUCHERNO, CASHVOUCHER.PAIDTO, CASHVOUCHER.POSITION, CASHVOUCHER.PARTICULAR, CASHVOUCHER.AMOUNT, CASHVOUCHER.REMARKS, CASHVOUCHER.PRINTCOPY From CASHVOUCHER " & _
                  "Where (((CASHVOUCHER.CVDate) >= #" & dFrom & "# And (CASHVOUCHER.CVDate) <= #" & dTo & "#)) ORDER BY CASHVOUCHER.CVDATE DESC, CASHVOUCHER.VOUCHERNO;"
            'mySQL
            Qry = "SELECT CASHVOUCHER.CVDATE, CASHVOUCHER.VOUCHERNO, CASHVOUCHER.PAIDTO, CASHVOUCHER.POSITION, CASHVOUCHER.PARTICULAR, CASHVOUCHER.AMOUNT, CASHVOUCHER.REMARKS, CASHVOUCHER.PRINTCOPY From CASHVOUCHER " & _
                  "Where (((CASHVOUCHER.CVDate) >= '" & Format(dFrom, "yyyy-mm-dd") & "' And (CASHVOUCHER.CVDate) <= '" & Format(dTo, "yyyy-mm-dd") & "')) ORDER BY CASHVOUCHER.CVDATE DESC, CASHVOUCHER.VOUCHERNO;"
            
        End If
    
    ElseIf mySrc = "Advances for Liquidation" Then
        If dCode = 1 Then
            Qry = "SELECT date_format(advliquidation.AFLDATE,'%c/%d/%Y'), advliquidation.AFLNO, format(advliquidation.AMTREC,2), format(advliquidation.AMTRETURN,2), format(advliquidation.AFLAMOUNT,2), advliquidation.PREPAREDBY, advliquidation.REMARKS, advliquidation.PRINTCOPY, advliquidation.AFLID " & _
                  "From advliquidation Where advliquidation.AFLDATE = '" & Format(dFrom.Value, "YYYY-MM-DD") & "' Order By advliquidation.AFLNO;"
        Else
            Qry = "SELECT date_format(advliquidation.AFLDATE,'%c/%d/%Y'), advliquidation.AFLNO, format(advliquidation.AMTREC,2), format(advliquidation.AMTRETURN,2), format(advliquidation.AFLAMOUNT,2), advliquidation.PREPAREDBY, advliquidation.REMARKS, advliquidation.PRINTCOPY, advliquidation.AFLID " & _
                  "From advliquidation Where advliquidation.AFLDATE >= '" & Format(dFrom.Value, "YYYY-MM-DD") & "' and advliquidation.AFLDATE <= '" & Format(dTo.Value, "YYYY-MM-DD") & "' Order By advliquidation.AFLNO;"
        End If
        
    ElseIf mySrc = "Cash Return" Then
        If dCode = 1 Then
            Qry = "SELECT date_format(cashreturn.CRDATE,'%c/%d/%Y'), cashreturn.CASHRETNO, format(cashreturn.CRAMOUNT,2), cashreturn.REMARKS, cashreturn.PREPAREDBY, cashreturn.PRINTCOPY, cashreturn.CRID " & _
                  "From cashreturn Where cashreturn.CRDATE = '" & Format(dFrom.Value, "YYYY-MM-DD") & "';"
        Else
            Qry = "SELECT date_format(cashreturn.CRDATE,'%c/%d/%Y'), cashreturn.CASHRETNO, format(cashreturn.CRAMOUNT,2), cashreturn.REMARKS, cashreturn.PREPAREDBY, cashreturn.PRINTCOPY, cashreturn.CRID " & _
                  "From cashreturn Where cashreturn.CRDATE >= '" & Format(dFrom.Value, "YYYY-MM-DD") & "' and cashreturn.CRDATE <= '" & Format(dTo.Value, "YYYY-MM-DD") & "';"
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
                        "IF(MONTH(SRETURNSP.RETDATE) = MONTH(SPINVOICE.INVDATE), SRETURNSP.TOTALAMT, '') AS THMOS, " & _
                        "IF(MONTH(SRETURNSP.RETDATE) = MONTH(SPINVOICE.INVDATE),'',SRETURNSP.TOTALAMT) AS PRMOS, " & _
                        "SRETURNSP.TOTALAMT, " & _
                        "SRETURNSP.REMARKS, " & _
                        "SRETURNSP.PRINTCOPY, " & _
                        "SRETURNSP.RETURNID " & _
                    "From " & _
                        "SRETURNSP " & _
                        "INNER JOIN SPINVOICE ON (SRETURNSP.INVOICE = SPINVOICE.SPINVNO) " & _
                    "Where " & _
                        "SRETURNSP.RETDATE = '" & Format(dFrom.Value, "yyyy-mm-dd") & "' " & _
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
                        "IF(MONTH(SRETURNSP.RETDATE) = MONTH(SPINVOICE.INVDATE), SRETURNSP.TOTALAMT, '') AS THMOS, " & _
                        "IF(MONTH(SRETURNSP.RETDATE) = MONTH(SPINVOICE.INVDATE),'',SRETURNSP.TOTALAMT) AS PRMOS, " & _
                        "SRETURNSP.TOTALAMT, " & _
                        "SRETURNSP.REMARKS, " & _
                        "SRETURNSP.PRINTCOPY, " & _
                        "SRETURNSP.RETURNID " & _
                    "From " & _
                        "SRETURNSP " & _
                        "INNER JOIN SPINVOICE ON (SRETURNSP.INVOICE = SPINVOICE.SPINVNO) " & _
                    "Where " & _
                        "SRETURNSP.RETDATE >= '" & Format(dFrom.Value, "yyyy-mm-dd") & "' AND " & _
                        "SRETURNSP.RETDATE <= '" & Format(dTo.Value, "yyyy-mm-dd") & "' " & _
                    "Order By " & _
                        "SRETURNSP.RETDATE, " & _
                        "SRETURNSP.RETNO Desc"
        
        
        End If
        ElseIf mySrc = "Branch Purchase Order" Then
        
            If dCode = 1 Then
                 Qry = "SELECT spordertaking.SQSERVE, spordertaking.BRANCHNAME, ' ', spordertaking.DATESERVE, spordertaking.DATECONFIRM, '', spordertaking.CONFIRMNO, " & _
                    "spordertaking.TOTALAMOUNT, spordertaking.BRANCHID, spordertaking.PRINTPICK, spordertaking.FINALQT " & _
                    "From spordertaking Where spordertaking.DATESERVE >= '" & Format(dFrom.Value, "yyyy-mm-dd") & "' Order By spordertaking.ORDERIDSQ;"
            Else
                 Qry = "SELECT spordertaking.SQSERVE, spordertaking.BRANCHNAME, ' ', spordertaking.DATESERVE, spordertaking.DATECONFIRM,' ',spordertaking.CONFIRMNO, " & _
                   "spordertaking.TOTALAMOUNT, spordertaking.BRANCHID, spordertaking.PRINTPICK, spordertaking.FINALQT " & _
                   "From spordertaking Where spordertaking.DATESERVE >= '" & Format(dFrom.Value, "yyyy-mm-dd") & "' AND spordertaking.DATESERVE <= '" & Format(dTo.Value, "yyyy-mm-dd") & "' Order by spordertaking.ORDERIDSQ;"

            End If
           
    End If
    If rssummary.State = adStateOpen Then rssummary.Close
    rssummary.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    If rssummary.RecordCount = 0 Then
        flxSum.Clear
        flxSum.Rows = 2
        flxSum.Cols = TCol
        Call RXResizer1_ResizeComplete
        Exit Sub
    End If
    Set flxSum.DataSource = rssummary
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
    Dim dCov As String, pQRY As String, JOid As String
    
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
                ElseIf mySrc = "Received Spare Parts" Then
                    Call SPincoming.ShowRecord(flxSum.TextMatrix(ActRow, 2))
                    Unload Me
                'ElseIf mySrc = "Sales Order" Then
                    'Call SPSorder.DisplaySalesOrder(flxSum.TextMatrix(ActRow, 7))
                    'Unload Me
                ElseIf mySrc = "Cashier Cash Flow" Then
                    'All Branches of Moto-r Davao
                    '5/10/2010-TAGUM
                    '6/10/2010-ULAS
                    '4/01/2011 Depot
                    '"7/25/2011" - VAN - LYR SHOP
                    If SysBranch = "MOTO-R REPUBLIC CORP." Then
                        If DateValue(.TextMatrix(ActRow, 1)) >= DateValue("05/10/2010") Then
                            Call CashFlowCashier.ShowRecord(Val(.TextMatrix(ActRow, 5)))
                            Unload Me
                        Else
                            CashFlowCashierOld.CFdate.Value = DateValue(.TextMatrix(ActRow, 1))
                            Call CashFlowCashierOld.ViewCashFlow
                            Unload Me
                        End If
                    ElseIf SysBranch = "MOTO-R ULAS" Then
                        If DateValue(.TextMatrix(ActRow, 1)) >= DateValue("06/10/2010") Then
                            Call CashFlowCashier.ShowRecord(Val(.TextMatrix(ActRow, 5)))
                            Unload Me
                        Else
                            CashFlowCashierOld.CFdate.Value = DateValue(.TextMatrix(ActRow, 1))
                            Call CashFlowCashierOld.ViewCashFlow
                            Unload Me
                        End If
                    Else
                        'Depot, Van, LYR SHOP
                        Call CashFlowCashier.ShowRecord(Val(.TextMatrix(ActRow, 5)))
                    End If
                    
                ElseIf mySrc = "Manager's Cash Flow" Then
                    'If DateValue(.TextMatrix(ActRow, 1)) >= DateValue("05/10/2010") Then '5/10/2010-tagum
                    '    Call CashFlowMgr.ShowRecord(Val(.TextMatrix(ActRow, 5)))
                    '    Unload Me
                    'Else
                    '    CashFlowMgrOld.CFdate.Value = DateValue(.TextMatrix(ActRow, 1))
                    '    Call CashFlowMgrOld.ViewCashFlow
                    '    Unload Me
                    If SysBranch = "MOTO-R REPUBLIC CORP." Then
                        If DateValue(.TextMatrix(ActRow, 1)) >= DateValue("05/10/2010") Then
                            Call CashFlowMgr.ShowRecord(Val(.TextMatrix(ActRow, 5)))
                            Unload Me
                        Else
                            CashFlowMgrOld.CFdate.Value = DateValue(.TextMatrix(ActRow, 1))
                            Call CashFlowMgrOld.ViewCashFlow
                        End If
                    ElseIf SysBranch = "MOTO-R ULAS" Then
                        If DateValue(.TextMatrix(ActRow, 1)) >= DateValue("06/10/2010") Then
                            Call CashFlowMgr.ShowRecord(Val(.TextMatrix(ActRow, 5)))
                            Unload Me
                        Else
                            CashFlowMgrOld.CFdate.Value = DateValue(.TextMatrix(ActRow, 1))
                            Call CashFlowMgrOld.ViewCashFlow
                        End If
                    Else
                        Call CashFlowMgr.ShowRecord(Val(.TextMatrix(ActRow, 5)))
                        Unload Me
                    End If
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

                End If
            End With
            
        Case 2  'Print
            If mySrc = "Job Order" Then
                dCov = "Date Covered: " & Format(dFrom.Value, "Mmm. dd, yyyy") & IIf(Val(cmdNext.Tag) = 2, " to " & Format(dTo.Value, "Mmm. dd, yyyy"), "")
                If Val(cmdNext.Tag) = 1 Then
                    '***Access
                    'Qry = "SELECT SRVCJOBORDER.DATEPROMISED, SRVCJOBORDER.JOBORDERNO, SRVCJOBORDER.CUSTOMER, SRVCJOBORDER.PAYTYPE, SRVCJOBORDER.TOTALCOST, SRVCJOBORDER.ORNO, SRVCUSTOMER.MODEL, SRVCJOBORDER.KMREADING, SRVCJOBORDER.JOTYPE, SRVCJOBORDER.MECHANIC, SRVCJOBORDER.REMARKS, SRVCJOBORDER.DISCOUNT " & _
                          "FROM SRVCJOBORDER INNER JOIN SRVCUSTOMER ON SRVCJOBORDER.ACCOUNTNO = SRVCUSTOMER.ACCOUNTNO Where (((SRVCJOBORDER.DATEPROMISED) = #" & dFrom.Value & "#) And ((SRVCJOBORDER.PRINTCOPY) > 0)) ORDER BY SRVCJOBORDER.DATEPROMISED, SRVCJOBORDER.JOBORDERNO;"
                    'mySQL
                    Qry = "SELECT SRVCJOBORDER.DATEPROMISED, SRVCJOBORDER.JOBORDERNO, SRVCJOBORDER.CUSTOMER, SRVCJOBORDER.PAYTYPE, SRVCJOBORDER.TOTALCOST, SRVCJOBORDER.ORNO, SRVCUSTOMER.MODEL, SRVCJOBORDER.KMREADING, SRVCJOBORDER.JOTYPE, SRVCJOBORDER.MECHANIC, SRVCJOBORDER.REMARKS, SRVCJOBORDER.DISCOUNT " & _
                          "FROM SRVCJOBORDER INNER JOIN SRVCUSTOMER ON SRVCJOBORDER.ACCOUNTNO = SRVCUSTOMER.ACCOUNTNO Where SRVCJOBORDER.DATEPROMISED = '" & Format(dFrom.Value, "yyyy-mm-dd") & "' And SRVCJOBORDER.PRINTCOPY > 0 ORDER BY SRVCJOBORDER.DATEPROMISED, SRVCJOBORDER.JOBORDERNO;"
                    
                Else
                    '***Access
                    'Qry = "SELECT SRVCJOBORDER.DATEPROMISED, SRVCJOBORDER.JOBORDERNO, SRVCJOBORDER.CUSTOMER, SRVCJOBORDER.PAYTYPE, SRVCJOBORDER.TOTALCOST, SRVCJOBORDER.ORNO, SRVCUSTOMER.MODEL, SRVCJOBORDER.KMREADING, SRVCJOBORDER.JOTYPE, SRVCJOBORDER.MECHANIC, SRVCJOBORDER.REMARKS, SRVCJOBORDER.DISCOUNT " & _
                          "FROM SRVCJOBORDER INNER JOIN SRVCUSTOMER ON SRVCJOBORDER.ACCOUNTNO = SRVCUSTOMER.ACCOUNTNO Where (((SRVCJOBORDER.DATEPROMISED) >= #" & dFrom.Value & "# And (SRVCJOBORDER.DATEPROMISED) <= #" & dTo.Value & "#) And ((SRVCJOBORDER.PRINTCOPY) > 0)) ORDER BY SRVCJOBORDER.DATEPROMISED, SRVCJOBORDER.JOBORDERNO;"
                    'mySQl
                    Qry = "SELECT SRVCJOBORDER.DATEPROMISED, SRVCJOBORDER.JOBORDERNO, SRVCJOBORDER.CUSTOMER, SRVCJOBORDER.PAYTYPE, SRVCJOBORDER.TOTALCOST, SRVCJOBORDER.ORNO, SRVCUSTOMER.MODEL, SRVCJOBORDER.KMREADING, SRVCJOBORDER.JOTYPE, SRVCJOBORDER.MECHANIC, SRVCJOBORDER.REMARKS, SRVCJOBORDER.DISCOUNT " & _
                          "FROM SRVCJOBORDER INNER JOIN SRVCUSTOMER ON SRVCJOBORDER.ACCOUNTNO = SRVCUSTOMER.ACCOUNTNO Where SRVCJOBORDER.DATEPROMISED >= '" & Format(dFrom.Value, "yyyy-mm-dd") & "' And SRVCJOBORDER.DATEPROMISED <= '" & Format(dTo.Value, "yyyy-mm-dd") & "' And SRVCJOBORDER.PRINTCOPY > 0 ORDER BY SRVCJOBORDER.DATEPROMISED, SRVCJOBORDER.JOBORDERNO;"
                    
                End If
                PRNmc.QryPrn = Qry
                PRNmc.ParamValue1 = dCov
                PRNmc.ParamValue2 = "Prepared By: " & GetUserName
                PRNmc.MySource = "JOSUMMARY"
                PRNmc.Show
                
            ElseIf mySrc = "Sales Return" Then
                dCov = "Date Covered: " & Format(dFrom.Value, "Mmm. dd, yyyy") & IIf(Val(cmdNext.Tag) = 2, " to " & Format(dTo.Value, "Mmm. dd, yyyy"), "")
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
                        "IF(MONTH(SRETURNSP.RETDATE) = MONTH(SPINVOICE.INVDATE), SRETURNSP.TOTALAMT, '') AS THMOS, " & _
                        "IF(MONTH(SRETURNSP.RETDATE) = MONTH(SPINVOICE.INVDATE),'',SRETURNSP.TOTALAMT) AS PRMOS, " & _
                        "SRETURNSP.TOTALAMT, " & _
                        "SRETURNSP.REMARKS, " & _
                        "SRETURNSP.PRINTCOPY, " & _
                        "SRETURNSP.RETURNID " & _
                    "From " & _
                        "SRETURNSP " & _
                        "INNER JOIN SPINVOICE ON (SRETURNSP.INVOICE = SPINVOICE.SPINVNO) " & _
                    "Where " & _
                        "SRETURNSP.RETDATE >= '" & Format(dFrom.Value, "yyyy-mm-dd") & "' " & _
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
                        "IF(MONTH(SRETURNSP.RETDATE) = MONTH(SPINVOICE.INVDATE), SRETURNSP.TOTALAMT, '') AS THMOS, " & _
                        "IF(MONTH(SRETURNSP.RETDATE) = MONTH(SPINVOICE.INVDATE),'',SRETURNSP.TOTALAMT) AS PRMOS, " & _
                        "SRETURNSP.TOTALAMT, " & _
                        "SRETURNSP.REMARKS, " & _
                        "SRETURNSP.PRINTCOPY, " & _
                        "SRETURNSP.RETURNID " & _
                    "From " & _
                        "SRETURNSP " & _
                        "INNER JOIN SPINVOICE ON (SRETURNSP.INVOICE = SPINVOICE.SPINVNO) " & _
                    "Where " & _
                        "SRETURNSP.RETDATE >= '" & Format(dFrom.Value, "yyyy-mm-dd") & "' AND " & _
                        "SRETURNSP.RETDATE <= '" & Format(dTo.Value, "yyyy-mm-dd") & "' " & _
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
                
                
            ElseIf mySrc = "Received Spare Parts" Then
                If Val(cmdNext.Tag) = 2 Then
                    dCov = "Date Covered: " & dFrom.Value & " to " & dTo.Value
                    '***Access
                    'Qry = "SELECT SPRECEIVE.DATEREC, SPRECEIVE.PACKINGNO, SPRECEIVE.SUPPLIER, SPRECEIVE.SHIPPER, SPRECEIVE.WAYBILLNO, SPRECEIVE.PACKDATE, SPRECEIVE.REMARKS, SPRECEIVE.CHECKEDBY, SPRECEIVE.PRINTCOPY, SPRECEIVE.SPRRNO " & _
                          "From SPRECEIVE Where (((SPRECEIVE.DATEREC) >= #" & dFrom.Value & "# And (SPRECEIVE.DATEREC) <= #" & dTo.Value & "#)) ORDER BY SPRECEIVE.DATEREC, SPRECEIVE.PACKINGNO DESC;"
                    'mySQl
                    Qry = "SELECT SPRECEIVE.DATEREC, SPRECEIVE.PACKINGNO, SPRECEIVE.SUPPLIER, SPRECEIVE.SHIPPER, SPRECEIVE.WAYBILLNO, SPRECEIVE.PACKDATE, SPRECEIVE.REMARKS, SPRECEIVE.CHECKEDBY, SPRECEIVE.PRINTCOPY, SPRECEIVE.SPRRNO " & _
                          "From SPRECEIVE Where SPRECEIVE.DATEREC >= '" & Format(dFrom.Value, "yyyy-mm-dd") & "' And SPRECEIVE.DATEREC <= '" & Format(dTo.Value, "yyyy-mm-dd") & "' ORDER BY SPRECEIVE.DATEREC, SPRECEIVE.PACKINGNO DESC;"
                    'Stop here
                Else
                    '***Access
                    'Qry = "SELECT SPRECEIVE.DATEREC, SPRECEIVE.PACKINGNO, SPRECEIVE.SUPPLIER, SPRECEIVE.SHIPPER, SPRECEIVE.WAYBILLNO, SPRECEIVE.PACKDATE, SPRECEIVE.REMARKS, SPRECEIVE.CHECKEDBY, SPRECEIVE.PRINTCOPY, SPRECEIVE.SPRRNO " & _
                          "From SPRECEIVE Where (((SPRECEIVE.DATEREC) = #" & dFrom.Value & "#)) ORDER BY SPRECEIVE.DATEREC, SPRECEIVE.PACKINGNO DESC;"
                    'mySQL
                    Qry = "SELECT SPRECEIVE.DATEREC, SPRECEIVE.PACKINGNO, SPRECEIVE.SUPPLIER, SPRECEIVE.SHIPPER, SPRECEIVE.WAYBILLNO, SPRECEIVE.PACKDATE, SPRECEIVE.REMARKS, SPRECEIVE.CHECKEDBY, SPRECEIVE.PRINTCOPY, SPRECEIVE.SPRRNO " & _
                          "From SPRECEIVE Where SPRECEIVE.DATEREC = '" & Format(dFrom.Value, "yyyy-mm-dd") & "' ORDER BY SPRECEIVE.DATEREC, SPRECEIVE.PACKINGNO DESC;"
                    
                    dCov = "Date Covered: " & dFrom.Value
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
                          "Where SRVCJOBORDER.DATEPROMISED = '" & Format(dFrom.Value, "yyyy-mm-dd") & "' And SRVCJOBORDER.PRINTCOPY > 0 And SRVCJOBORDER.REMARKS <> 'CANCELLED' ORDER BY SRVCJOBORDER.MECHANIC, SRVCJOBORDER.DATEPROMISED;"
                    
                Else
                    '***Access
                    'Qry = "SELECT SRVCJOBORDER.MECHANIC, SRVCJOBORDER.JOBORDERNO, SRVCJOBORDER.DATEPROMISED, SRVCJOBORDER.JOTYPE, SRVCJOBORDER.TOTALCOST From SRVCJOBORDER " & _
                          "Where (((SRVCJOBORDER.DATEPROMISED) >= #" & dFrom.Value & "# And (SRVCJOBORDER.DATEPROMISED) <= #" & dTo.Value & "#) And ((SRVCJOBORDER.PRINTCOPY) > 0)) ORDER BY SRVCJOBORDER.MECHANIC, SRVCJOBORDER.DATEPROMISED;"
                    'mySQL
                    Qry = "SELECT SRVCJOBORDER.MECHANIC, SRVCJOBORDER.JOBORDERNO, SRVCJOBORDER.DATEPROMISED, SRVCJOBORDER.JOTYPE, SRVCJOBORDER.TOTALCOST From SRVCJOBORDER " & _
                          "Where SRVCJOBORDER.DATEPROMISED >= '" & Format(dFrom.Value, "yyyy-mm-dd") & "' And SRVCJOBORDER.DATEPROMISED <= '" & Format(dTo.Value, "yyyy-mm-dd") & "' And SRVCJOBORDER.PRINTCOPY > 0 AND SRVCJOBORDER.REMARKS <> 'CANCELLED' ORDER BY SRVCJOBORDER.MECHANIC, SRVCJOBORDER.DATEPROMISED;"
                    
                End If
                
                PRNmc.QryPrn = Qry
                PRNmc.ParamValue1 = "Date Covered: " & Format(dFrom.Value, "Mmm. dd, yyyy") & IIf(Val(cmdNext.Tag) = 1, "", " to " & Format(dTo.Value, "Mmm. dd, yyyy"))
                PRNmc.ParamValue2 = Val(txtFSCamt.Text)
                PRNmc.MySource = "SRVCMECH"
                PRNmc.Show
            ElseIf mySrc = "DMCM [A/R]" Then
                If Val(cmdNext.Tag) = 1 Then
                
                    Qry = "SELECT " & _
                        "ADJDMCM.ADJDATE, " & _
                        "ADJDMCM.ADJNO, " & _
                        "ADJDMCM.CUSTOMER, " & _
                        "ADJDMCM.TRANSTYPE, " & _
                        "ADJDMCM.REFNO, " & _
                        "spinvoice.INVDATE, " & _
                        "ADJDMCM.DEBIT, " & _
                        "ADJDMCM.CREDIT, " & _
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
                    Qry = "SELECT " & _
                        "ADJDMCM.ADJDATE, " & _
                        "ADJDMCM.ADJNO, " & _
                        "ADJDMCM.CUSTOMER, " & _
                        "ADJDMCM.TRANSTYPE, " & _
                        "ADJDMCM.REFNO, " & _
                        "spinvoice.INVDATE, " & _
                        "ADJDMCM.DEBIT, " & _
                        "ADJDMCM.CREDIT, " & _
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
                               
                PRNmc.QryPrn = Qry
                PRNmc.ParamValue1 = "Date Covered: " & Format(dFrom.Value, "Mmm. dd, yyyy") & IIf(Val(cmdNext.Tag) = 1, "", " to " & Format(dTo.Value, "Mmm. dd, yyyy"))
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

            Else
                MsgBox "Not yet available", vbInformation + vbOKOnly, "Print"
            End If
        Case 3  'Exit
            Unload Me
            MotoRep.tbMain.Visible = True
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

Private Sub flxSum_EnterCell()
    ActRow = flxSum.Row
End Sub
Private Sub flxSum_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If flxSum.TextMatrix(ActRow, 1) = "" Or mySrc = "Received Spare Parts" Or mySrc = "Sales Order" Or mySrc = "Mechanic Income" Or mySrc = "Top Selling Spare Parts" Or mySrc = "Cashier Cash Flow" Then Exit Sub
    If Not IsDate(flxSum.TextMatrix(ActRow, 1)) Then
        Exit Sub
    Else
        If DateValue(flxSum.TextMatrix(ActRow, 1)) <> Date Then Exit Sub
    End If
    If Button = vbRightButton Then
        mnuCancel.Caption = "Cancel " & mySrc
       Call PopupMenu(mnuSum)
    End If
End Sub
Private Sub Form_Load()
    Dim Qry As String
    
    dFrom.Value = Date
    cmdNext.Tag = "1"
    If mySrc = "Job Order" Then
        TCol = 14: ColPrn = 12: AccID = 43
    ElseIf mySrc = "Clerical Adjustment" Then
        TCol = 11: ColPrn = 11: AccID = 56
    ElseIf mySrc = "DMCM [A/R]" Then
        TCol = 11: ColPrn = 11: AccID = 54
    ElseIf mySrc = "Inventory Adjustment" Then
        TCol = 8: ColPrn = 8: AccID = 55
    ElseIf mySrc = "Customer P.O." Then
        TCol = 8: ColPrn = 0: AccID = 40
    ElseIf mySrc = "Top Selling Spare Parts" Then
        TCol = 6: AccID = 97
    ElseIf mySrc = "Purchase Return" Then
        TCol = 9: ColPrn = 0: AccID = 24
        Call cmdNext_Click
        If Day(Date) <= 7 Then
            dFrom.Value = DateAdd("m", -1, DateSerial(Year(Date), Month(Date), 1))
        Else
            dFrom.Value = DateSerial(Year(Date), Month(Date), 1)
        End If
    ElseIf mySrc = "Requisition" Then
        TCol = 9: ColPrn = 0: AccID = 39
    ElseIf mySrc = "Sales Return" Then
        TCol = 9: ColPrn = 0: AccID = 29
    ElseIf mySrc = "Location Adjustment" Then
        TCol = 8: ColPrn = 0: AccID = 35
        Call cmdNext_Click
        If Day(Date) <= 7 Then
            dFrom.Value = DateAdd("m", -1, DateSerial(Year(Date), Month(Date), 1))
        Else
            dFrom.Value = DateSerial(Year(Date), Month(Date), 1)
        End If
    ElseIf mySrc = "Received Spare Parts" Then
        TCol = 10: ColPrn = 0: AccID = 23
        Call cmdNext_Click
        If Day(Date) <= 7 Then
            dFrom.Value = DateAdd("m", -1, DateSerial(Year(Date), Month(Date), 1))
        Else
            dFrom.Value = DateSerial(Year(Date), Month(Date), 1)
        End If
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
        dFrom.Value = DateSerial(Year(Date), Month(Date), 1)
    ElseIf mySrc = "Cashier Cash Flow" Then
        TCol = 6: ColPrn = 0: AccID = 48
        Call cmdNext_Click
        dFrom.Value = DateSerial(Year(Date), Month(Date), 1)
        dTo.Value = Date
        'ShowDateRange = True
    ElseIf mySrc = "Manager's Cash Flow" Then
        TCol = 6: ColPrn = 0: AccID = 48
        Call cmdNext_Click
        dFrom.Value = DateSerial(Year(Date), Month(Date), 1)
        dTo.Value = Date
        'ShowDateRange = True
        
    ElseIf mySrc = "Advances for Liquidation" Then
        TCol = 9: ColPrn = 0: AccID = 65
        dFrom = Date
        dTo.Value = Date
        cmdNext.Tag = "1"
        'ShowDateRange = False
    ElseIf mySrc = "Cash Return" Then
        TCol = 7: ColPrn = 0: AccID = 64
        Call cmdNext_Click
        dFrom.Value = DateSerial(Year(Date), Month(Date), 1)
        dTo.Value = Date
        'ShowDateRange = True
    ElseIf mySrc = "Branch Purchase Order" Then
        TCol = 12: ColPrn = 0
        'AccID =0
        dFrom = Date
        dTo.Value = Date
    End If
    If mySrc = "Purchase Return" Or mySrc = "Location Adjustment" Or mySrc = "Cash Voucher" Or mySrc = "Received Spare Parts" Or mySrc = "Cashier Cash Flow" Or mySrc = "Cash Return" Or mySrc = "Manager's Cash Flow" Then
        Call RefreshRecord(2)
    Else
        Call RefreshRecord(1)
    End If
    ActRow = 0
    MotoRep.sBar.Panels(6).Text = mySrc & " Summary"
    lblHeading(2).Caption = mySrc & " Summary"
    lblHeading(3).Caption = mySrc & " Summary"
    Me.Caption = mySrc & " Summary"
    MotoRep.tbMain.Visible = False
    ActRow = 1
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    MotoRep.sBar.Panels(6).Text = ""
    If rssummary.State = adStateOpen Then rssummary.Close
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
    If mySrc = "Job Order" Then
        Call CancelJO
    ElseIf mySrc = "Clerical Adjustment" Then
        Call CancelClericalAdj
    ElseIf mySrc = "DMCM [A/R]" Then
        Call CancelDMCM
    ElseIf mySrc = "Inventory Adjustment" Then
        Call CancelInvAdj
    ElseIf mySrc = "Customer P.O." Then
        Call CancelCustPO
    ElseIf mySrc = "Purchase Return" Then
        Call CancelPurchaseReturn
    ElseIf mySrc = "Requisition" Then
        Call CancelRequisition
    ElseIf mySrc = "Sales Return" Then
        Call CancelSalesReturn
    ElseIf mySrc = "Location Adjustment" Then
        Call CancelLocAdj
    ElseIf mySrc = "Cash Voucher" Then
        Call CancelCashVoucher
    End If
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
                FlxPrev.AddItem !SPAREPARTID & vbTab & !PrevLoc
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
                flxCancelSalRet.AddItem !RECEIVE & vbTab & !PartNumber & vbTab & !Description & vbTab & !Price & vbTab & !Amount & vbTab & !SPAREPARTID
                If InStr(UCase$(!Description), "4T OIL") > 0 Then
                    MyUnit = IIf(!RECEIVE > 1, "QRTS", "QRT")
                Else
                    MyUnit = IIf(!RECEIVE > 1, "PCS", "PC")
                End If
                If Len(!Description) < 30 Then
                    MyDesc = !Description & Space$(30 - Len(!Description))
                Else
                    MyDesc = Left$(!Description, 30)
                End If
                
                MyDetails = IIf(MyDetails <> "", MyDetails & vbCrLf, "") & !RECEIVE & vbTab & MyUnit & vbTab & !PartNumber & Space$(20 - Len(!PartNumber)) & vbTab & MyDesc & vbTab & Format(!Price, "#,##0.00") & vbTab & Format(!Amount, "#,##0.00")
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
                    MyCredit = IIf(IsNull(!Credit), 0, !Credit)
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
    Dim rsSPitr As New ADODB.Recordset
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
    Dim MyUnit As String, MyDesc As String, MyInv As String, MyProc As String
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
                flxCancelPuRet.AddItem !OUT & vbTab & !PartNumber & vbTab & !Description & vbTab & !Price & vbTab & !Amount & vbTab & !SPAREPARTID
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
                MyDetails = IIf(MyDetails <> "", MyDetails & vbCrLf, "") & !OUT & vbTab & MyUnit & vbTab & !PartNumber & Space$(20 - Len(!PartNumber)) & vbTab & MyDesc & vbTab & Format(!Price, "#,##0.00") & vbTab & Format(!Amount, "#,##0.00")
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
                flxCancel.AddItem !PartNumber & vbTab & IIf(IsNull(!Remarks), "", !Remarks) & vbTab & !RECEIVE & vbTab & !OUT & vbTab & IIf(IsNull(!Source), "", !Source) & vbTab & !SPAREPARTID
                MyDetails = IIf(MyDetails <> "", MyDetails & vbCrLf, "") & !PartNumber & vbTab & flxCancel.TextMatrix(flxCancel.Rows - 1, 1) & Space$(2) & vbTab & IIf(InStr(UCase$(flxCancel.TextMatrix(flxCancel.Rows - 1, 1)), "DEBIT") > 0, !RECEIVE, !OUT) & vbTab & flxCancel.TextMatrix(flxCancel.Rows - 1, 4)
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
        Qry = "SELECT SPINVOICE.INVTYPE From SPINVOICE WHERE SPINVOICE.SPINVNO='" & flxSum.TextMatrix(ActRow, 4) & "';"
        
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
                        MyCredit = IIf(IsNull(!Credit), 0, !Credit)
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
    LYRcon.Execute "UPDATE Ledger SET Ledger.REFTYPE = 'CANCELLED', Ledger.CUSTOMERID = 0 WHERE (((Ledger.REFTYPE)='DMCM') AND ((Ledger.REFNO)='" & MyRefNo & "'));"
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

Private Sub RXResizer1_ResizeComplete()
    Dim i As Integer, j As Integer
    lblHeading(2).Left = Frame1.Left + (Frame1.Width - lblHeading(2).Width) - 30
    lblHeading(3).Left = Frame1.Left + (Frame1.Width - lblHeading(3).Width)
    
    With flxSum
        If mySrc = "Job Order" Then
            
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
        ElseIf mySrc = "DMCM [A/R]" Then
                            '       1       2           3       4         5         6       7        8       9        10      11      12
            .FormatString = "   |^ DATE|^ADJ. NO.|<CUSTOMER|<REFERENCE|^REF. NO.|^INV. DATE|>DEBIT |>CREDIT|>BALANCE|<STATUS|<REASON|<PRN COPY"
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
            .ColWidth(10) = (.Width - 300) * 0.12
            .ColWidth(11) = (.Width - 300) * 0.2
            .ColWidth(12) = (.Width - 300) * 0.12
            
            For i = 1 To .Rows - 1
                .TextMatrix(i, 1) = Format$(.TextMatrix(i, 1), "mm/dd/yyyy")
                .TextMatrix(i, 6) = Format$(.TextMatrix(i, 6), "mm/dd/yyyy")
                .TextMatrix(i, 7) = Format$(.TextMatrix(i, 7), "#,##0.00")
                .TextMatrix(i, 8) = Format$(.TextMatrix(i, 8), "#,##0.00")
                .TextMatrix(i, 9) = Format$(.TextMatrix(i, 9), "#,##0.00")
                If Val(.TextMatrix(i, 12)) = 0 Then
                    '.Col = 0: .Row = i: .CellFontName = "Webdings": .CellFontSize = 10: .TextMatrix(i, 0) = lblP.Caption
                    .Col = 1: .Row = i: .ColSel = 12: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 250, 180)
                    .Col = 2
                End If
                If .TextMatrix(i, 11) = "CANCELLED" Then
                    '.Col = 0: .Row = i: .CellFontName = "Webdings": .CellFontSize = 10: .TextMatrix(i, 0) = lblP.Caption
                    .Col = 1: .Row = i: .ColSel = 12: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(255, 0, 0)
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
        ElseIf mySrc = "Received Spare Parts" Then
                                      '1          2            3         4        5             6           7       8        9
            .FormatString = "   |^DATE REC.|<PACKING NO.|<SUPPLIER|<SHIPPER|<WAY BILL NO.|^PACK DATE|<REMARKS|<CHECKER|<PRINT COPY"
            .ColWidth(0) = 300
            .ColWidth(1) = (.Width - 300) * 0.12
            .ColWidth(2) = (.Width - 300) * 0.12
            .ColWidth(3) = (.Width - 300) * 0.17
            .ColWidth(4) = (.Width - 300) * 0.17
            .ColWidth(5) = (.Width - 300) * 0.12
            .ColWidth(6) = (.Width - 300) * 0.12
            .ColWidth(7) = (.Width - 300) * 0.15
            .ColWidth(8) = (.Width - 300) * 0.15
            .ColWidth(9) = (.Width - 300) * 0.15
            .ColWidth(10) = 0
            For i = 1 To .Rows - 1
                .TextMatrix(i, 1) = Format$(.TextMatrix(i, 1), "mm/dd/yyyy")
                .TextMatrix(i, 6) = Format$(.TextMatrix(i, 6), "mm/dd/yyyy")
                If Val(.TextMatrix(i, 9)) = 0 Then 'Not yet printed
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
            .FormatString = "   |^ORDER DATE|^ORDER NO.|<FRONTLINE|>TOTAL PRICE|^INVOICE NO.|<REFERENCE"
            .ColWidth(0) = 300
            .ColWidth(1) = (.Width - 300) * 0.15
            .ColWidth(2) = (.Width - 300) * 0.15
            .ColWidth(3) = (.Width - 300) * 0.2
            .ColWidth(4) = (.Width - 300) * 0.15
            .ColWidth(5) = (.Width - 300) * 0.15
            .ColWidth(6) = (.Width - 300) * 0.2
            .ColWidth(7) = 0
            For i = 1 To .Rows - 1
                .TextMatrix(i, 1) = Format$(.TextMatrix(i, 1), "mm/dd/yyyy")
                .TextMatrix(i, 4) = Format$(.TextMatrix(i, 4), "###,##0.00")
            Next i
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
            .FormatString = "  |^Cash Flow Date |<Prepared By |< Approved By  |<Remarks "
            .ColWidth(0) = 300
            .ColWidth(1) = (.Width - 400) * 0.15
            .ColWidth(2) = (.Width - 400) * 0.25
            .ColWidth(3) = (.Width - 400) * 0.25
            .ColWidth(4) = (.Width - 400) * 0.3
            .ColWidth(5) = 0
        
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
                .TextMatrix(i, 2) = Format$(Val(.TextMatrix(i, 2)), "0000")
                .TextMatrix(i, 6) = Format$(.TextMatrix(i, 6), "#,##0.00")
                If Val(.TextMatrix(i, 8)) = 0 Then 'Not yet printed
                    .Col = 1: .Row = i: .ColSel = TCol - 1: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 250, 200)
                    .Col = 2
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
        End If
    End With
    
    cmdFind.Height = cmdRefresh.Height
    cmdFind.Top = cmdRefresh.Top
    txtFind.Top = cmdFind.Top
    
End Sub

Private Sub txtFind_Change()
    fRow = 1
    fCol = 1
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then Call cmdFind_Click
End Sub

