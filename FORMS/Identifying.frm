VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form BranchIdentifying 
   Caption         =   "Parts Matrix Order"
   ClientHeight    =   9150
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   19230
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9150
   ScaleWidth      =   19230
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      Height          =   7215
      Left            =   120
      TabIndex        =   17
      Tag             =   "XYF"
      Top             =   2280
      Width           =   18975
      Begin MSComctlLib.ProgressBar pBar 
         Height          =   255
         Left            =   0
         TabIndex        =   24
         Tag             =   "XYF"
         Top             =   6960
         Visible         =   0   'False
         Width           =   18975
         _ExtentX        =   33470
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin MSFlexGridLib.MSFlexGrid flxSPARTS 
         Height          =   6735
         Left            =   120
         TabIndex        =   9
         Tag             =   "XYF"
         Top             =   240
         Width           =   18735
         _ExtentX        =   33046
         _ExtentY        =   11880
         _Version        =   393216
         Rows            =   3
         Cols            =   13
         FixedRows       =   2
         BackColorBkg    =   14737632
         AllowUserResizing=   1
         FormatString    =   "|<MODEL BRAND (PM) |<PARTNUMBER |<DESCRIPTION|<PARTSGROUP|^QH |^STDQ |^SRQ |^ORDER |>UNITPRICE |>TOTALPRICE |<LOCATION"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin MSFlexGridLib.MSFlexGrid flxSRQZero 
      Height          =   1335
      Left            =   19320
      TabIndex        =   23
      Top             =   7440
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   2355
      _Version        =   393216
   End
   Begin VB.TextBox tmpSTATUS 
      Height          =   375
      Left            =   19320
      TabIndex        =   22
      Top             =   7920
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox tmpREFNO 
      Height          =   375
      Left            =   19320
      TabIndex        =   21
      Top             =   7440
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtpmID 
      Height          =   285
      Left            =   19320
      TabIndex        =   20
      Top             =   3360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtALQ 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   19320
      TabIndex        =   11
      Tag             =   "XYF"
      Top             =   3000
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txtHoldWhiteBackColor 
      Height          =   375
      Left            =   19320
      TabIndex        =   19
      Text            =   "HoldBackColor"
      Top             =   2040
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtHoldGrayBackColor 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   19320
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   2520
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   1815
      Left            =   120
      TabIndex        =   12
      Tag             =   "XYF"
      Top             =   480
      Width           =   18975
      Begin VB.TextBox txtProcsBy 
         Height          =   285
         Left            =   1080
         TabIndex        =   26
         Tag             =   "XYF"
         Top             =   1200
         Width           =   5175
      End
      Begin VB.TextBox txtTtlALQ 
         Height          =   285
         Left            =   7080
         TabIndex        =   30
         Tag             =   "XYF"
         Top             =   1200
         Width           =   1095
      End
      Begin MSFlexGridLib.MSFlexGrid flxSummary 
         Height          =   1455
         Left            =   9360
         TabIndex        =   28
         Tag             =   "XYF"
         Top             =   240
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   2566
         _Version        =   393216
         Cols            =   6
         BackColorBkg    =   14737632
         FormatString    =   "|<MODEL BRAND(PM)  |^TOTAL PN |^TOTAL QH |^TOTAL SRQ"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.CheckBox chkFinal 
         Caption         =   "FINAL"
         Height          =   255
         Left            =   7440
         TabIndex        =   27
         Tag             =   "XYF"
         Top             =   840
         Width           =   855
      End
      Begin MSMask.MaskEdBox txtDateFinal 
         Height          =   255
         Left            =   6000
         TabIndex        =   7
         Tag             =   "XYF"
         Top             =   840
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtDatePrepared 
         Height          =   255
         Left            =   3600
         TabIndex        =   6
         Tag             =   "XYF"
         Top             =   840
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtRefNo 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1080
         TabIndex        =   5
         Tag             =   "XYF"
         Top             =   840
         Width           =   1455
      End
      Begin VB.ComboBox cboBranch 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1080
         TabIndex        =   4
         Tag             =   "XYF"
         Top             =   360
         Width           =   7095
      End
      Begin VB.Label lblProcBy 
         Caption         =   "PROCS BY"
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Tag             =   "XYF"
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label lblTtlALQ 
         Caption         =   "TTL ORD"
         Height          =   255
         Left            =   6360
         TabIndex        =   29
         Tag             =   "XYF"
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label lblDateFinal 
         Caption         =   "DATE FINAL"
         Height          =   255
         Left            =   5040
         TabIndex        =   16
         Tag             =   "XYF"
         Top             =   840
         Width           =   975
      End
      Begin VB.Label lblDatePrepared 
         Caption         =   "DATE PREP"
         Height          =   255
         Left            =   2640
         TabIndex        =   15
         Tag             =   "XYF"
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label lblRefNo 
         Caption         =   "REF. NO"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Tag             =   "XYF"
         Top             =   840
         Width           =   855
      End
      Begin VB.Label lblBranch 
         Caption         =   "BRANCH"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Tag             =   "XYF"
         Top             =   360
         Width           =   855
      End
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   1920
      _ExtentX        =   3387
      _ExtentY        =   661
      Caption         =   "&New Identifying"
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
      Image           =   "Identifying.frx":0000
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   1
      Left            =   2040
      TabIndex        =   1
      Top             =   0
      Width           =   1920
      _ExtentX        =   3387
      _ExtentY        =   661
      Caption         =   "&Edit Identifying"
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
      Image           =   "Identifying.frx":059A
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   2
      Left            =   3960
      TabIndex        =   2
      Top             =   0
      Width           =   1920
      _ExtentX        =   3387
      _ExtentY        =   661
      Caption         =   "&View Summary"
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
      Image           =   "Identifying.frx":0B34
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   5
      Left            =   7800
      TabIndex        =   3
      Top             =   0
      Width           =   1920
      _ExtentX        =   3387
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
      Image           =   "Identifying.frx":10CE
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   6
      Left            =   19320
      TabIndex        =   8
      Top             =   1080
      Visible         =   0   'False
      Width           =   1920
      _ExtentX        =   3387
      _ExtentY        =   661
      Caption         =   "Save"
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
      Image           =   "Identifying.frx":1668
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   7
      Left            =   19320
      TabIndex        =   10
      Top             =   1560
      Visible         =   0   'False
      Width           =   1920
      _ExtentX        =   3387
      _ExtentY        =   661
      Caption         =   "Cancel"
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
      Image           =   "Identifying.frx":1C02
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   3
      Left            =   5880
      TabIndex        =   31
      Top             =   0
      Width           =   1920
      _ExtentX        =   3387
      _ExtentY        =   661
      Caption         =   "&Print"
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
      Image           =   "Identifying.frx":219C
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   4
      Left            =   2040
      TabIndex        =   32
      Top             =   0
      Width           =   1920
      _ExtentX        =   3387
      _ExtentY        =   661
      Caption         =   "&Confirm PO"
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
      Image           =   "Identifying.frx":2736
      cBack           =   -2147483633
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   19320
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   19470
      m_MDI_Height    =   9720
   End
   Begin VB.Menu mnuStatus 
      Caption         =   "Popupmenu"
      Visible         =   0   'False
      Begin VB.Menu mnuAdd 
         Caption         =   "AddRow"
      End
   End
End
Attribute VB_Name = "BranchIdentifying"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strPartial As String, strTotal As String
Dim formAction As String
Dim Ans As String
Dim branchCode As String
Dim branchName As String


Dim Tmp() As String
Dim MyErr As String, myMsg As String
Dim Qry As String, Qry2 As String
Dim intID As Integer, strBoolean As String
Dim AVQ As Integer, intQH As Integer, intRO As Integer, intSRQ As Integer, intTOTALSRQ As Integer
Dim ALQ As Integer, intALQ As Integer
Public mySrc As String

Dim arrMODEL() As String
Dim arrBRAND() As String
Dim arrBRANCH() As String

Dim intTotalALQ As Long
Dim intTotalAVQ As Long
Dim i As Long, j As Long
Dim ActRow As Integer, ActCol As Integer
Dim intREFID As Integer
Dim SEQNO As Integer
Dim SUPPLIERID As Integer
'************* CODE UPDATED: 3/14/17 *****************
Dim TtlRecord As Integer, cntZeroLack As Integer
Dim RecordHolder As Integer, RecordHolderRow As Integer
Dim strPMID As String, strFrmSource As String
'********* END CODE UPDATED: 3/14/17 *****************
Dim tmpRs  As New ADODB.Recordset
Dim mainRs As New ADODB.Recordset


Dim StatSrch As Boolean, AddNew As Boolean, ProperExit As Boolean

Public Sub displayPM()
    DoEvents
    'branchCode = arrBRANCH(cboBranch.ListIndex, 0)
    branchCode = "MWH"
    Call modPMBranchIdentifying.sp_brnchIdentify_getDetails_total_retail(branchCode, 0, "TRUE", mainRs)
    With mainRs
        If .RecordCount > 0 Then
            flxSummary.Rows = 1
            Do While Not .EOF
                flxSummary.AddItem vbTab & !MODELBRAND & vbTab & !TOTALPN & vbTab & !TOTALQH & vbTab & !TOTALSRQ & vbTab & !PMATRIXID
                .MoveNext
            Loop
        Else
            flxSummary.Rows = 1
            flxSummary.Rows = 2
        End If
    End With
    
    flxSPARTS.Rows = 2
    flxSPARTS.Rows = 3
End Sub

Public Sub flxWidth()

    With flxSPARTS
        '.Clear
        .Cols = 15
        '.FormatString = "|^MODEL BRAND (PM) |^PARTNUMBER |^DESCRIPTION |^QOH |^STDQ |^SRQ |^QOH |^RO |^AVQ |^ALQ |<LOCATION |^ID"
        .ColWidth(0) = 500
        .ColWidth(1) = (.Width - .ColWidth(0)) * 0.12  'MODEL BRAND (PM)
        .ColWidth(2) = (.Width - .ColWidth(0)) * 0.13 'PARTNUMBER
        .ColWidth(3) = (.Width - .ColWidth(0)) * 0.17 'DESCRIPTION
        .ColWidth(4) = (.Width - .ColWidth(0)) * 0.09 'PARTSGROUP
        .ColWidth(5) = (.Width - .ColWidth(0)) * 0.06 'QH
        .ColWidth(6) = (.Width - .ColWidth(0)) * 0.06 'STDQ
        .ColWidth(7) = (.Width - .ColWidth(0)) * 0.06 'SRQ
        .ColWidth(8) = (.Width - .ColWidth(0)) * 0.06 'ORDER
        .ColWidth(9) = (.Width - .ColWidth(0)) * 0.06 'UNITPRICE
        .ColWidth(10) = (.Width - .ColWidth(0)) * 0.06 'TOTALPRICE
        .ColWidth(11) = (.Width - .ColWidth(0)) * 0.1  ' LOCATION
        .ColWidth(12) = 0 '(.Width - .ColWidth(0)) * 0.06  ' PMATRIXID
        .ColWidth(13) = 0 ' (.Width - .ColWidth(0)) * 0.01 ' SPAREPARTID
        .ColWidth(14) = 0 ' (.Width - .ColWidth(0)) * 0.1 ' SPAREPARTID
    End With
    
    With flxSummary
        .Cols = 6
        '.FormatString = "|<MODEL BRAND(PM)  |^TOTAL PN |^TOTAL SRQ |^TOTAL AVQ"
        .ColWidth(0) = 300
        .ColWidth(1) = (.Width - .ColWidth(0)) * 0.35  'MODEL BRAND (PM)
        .ColWidth(2) = (.Width - .ColWidth(0)) * 0.2 'TOTALPN
        .ColWidth(3) = (.Width - .ColWidth(0)) * 0.2 'TOTALQH
        .ColWidth(4) = (.Width - .ColWidth(0)) * 0.2 'TOTALSRQ
        .ColWidth(5) = 0 '(.Width - .ColWidth(0)) * 0.15 'PMID
        '.ColWidth(6) = 0 '(.Width - .ColWidth(0)) * 0.1 '
    End With
End Sub

Public Sub flxColChanger()

With flxSPARTS
    If mySrc = "From Summary" Then
        'MERGE COL
    .TextMatrix(0, 0) = " "
    .TextMatrix(1, 0) = " "
    'MERGE ROW
    .TextMatrix(0, 1) = "SPAREPARTS"
    .TextMatrix(0, 2) = "SPAREPARTS"
    .TextMatrix(0, 3) = "SPAREPARTS"
    .TextMatrix(0, 4) = "SPAREPARTS"
    .TextMatrix(0, 5) = SysBranch
    .TextMatrix(0, 6) = SysBranch
    .TextMatrix(0, 7) = SysBranch
    .TextMatrix(0, 8) = SysBranch
    .TextMatrix(0, 9) = SysBranch
    .TextMatrix(0, 10) = SysBranch
    .TextMatrix(0, 11) = SysBranch
    
    'MERGE COL =FALSE
    .TextMatrix(1, 1) = "MODEL BRAND (PM)"
    .TextMatrix(1, 2) = "PARTNUMBER"
    .TextMatrix(1, 3) = "DESCRIPTION"
    .TextMatrix(1, 4) = "PARTSGROUP"
    .TextMatrix(1, 5) = "QH"
    .TextMatrix(1, 6) = "STDQ"
    .TextMatrix(1, 7) = "SRQ"
    .TextMatrix(1, 8) = "ORDER"
    .TextMatrix(1, 9) = "UNITPRICE"
    .TextMatrix(1, 10) = "TOTALPRICE"
    .TextMatrix(1, 11) = "LOCATION"

    'SET MERGING
    .MergeCells = 1
    .MergeRow(0) = True
    .MergeCol(0) = True
    .MergeCol(1) = True
    For i = 1 To 7
        .MergeCol(i) = False
    Next i
    'SET THE COLUMN COLOR GREEN
    .Row = 0
    .Col = 0
    .Col = 1
   ' .CellBackColor = vbGreen

    Else
    'MERGE COL
    .TextMatrix(0, 0) = " "
    .TextMatrix(1, 0) = " "
    'MERGE ROW
    .TextMatrix(0, 1) = "SPAREPARTS"
    .TextMatrix(0, 2) = "SPAREPARTS"
    .TextMatrix(0, 3) = "SPAREPARTS"
    .TextMatrix(0, 4) = "SPAREPARTS"
    .TextMatrix(0, 5) = SysBranch
    .TextMatrix(0, 6) = SysBranch
    .TextMatrix(0, 7) = SysBranch
    .TextMatrix(0, 8) = SysBranch
    .TextMatrix(0, 9) = SysBranch
    .TextMatrix(0, 10) = SysBranch
    .TextMatrix(0, 11) = SysBranch
    
    'MERGE COL =FALSE
    .TextMatrix(1, 1) = "MODEL BRAND (PM)"
    .TextMatrix(1, 2) = "PARTNUMBER"
    .TextMatrix(1, 3) = "DESCRIPTION"
    .TextMatrix(1, 4) = "PARTSGROUP"
    .TextMatrix(1, 5) = "QH"
    .TextMatrix(1, 6) = "STDQ"
    .TextMatrix(1, 7) = "SRQ"
    .TextMatrix(1, 8) = "ORDER"
    .TextMatrix(1, 9) = "UNITPRICE"
    .TextMatrix(1, 10) = "TOTALPRICE"
    .TextMatrix(1, 11) = "LOCATION"

    'SET MERGING
    .MergeCells = 1
    .MergeRow(0) = True
    .MergeCol(0) = True
    .MergeCol(1) = True
    For i = 1 To 7
        .MergeCol(i) = False
    Next i
    'SET THE COLUMN COLOR GREEN
    .Row = 0
    .Col = 0
    .Col = 1
   ' .CellBackColor = vbGreen
    End If
End With
    
    Call flxWidth
    
End Sub

Private Sub cboBranch_KeyDown(KeyCode As Integer, Shift As Integer)
'************** CODE UPDATED: 04/07/17 **************************
    If KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then StatSrch = True
'********** END CODE UPDATED: 04/07/17 **************************
End Sub

Private Sub cboBranch_KeyPress(KeyAscii As Integer)
'************** CODE UPDATED: 04/07/17 **************************
Dim cbolistIndex As Integer

    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(cboBranch, KeyAscii, ValidChars)
    If KeyAscii = 13 Then
        cbolistIndex = modObjTools.getListIndex(cboBranch, cboBranch.Text)
        
        If cbolistIndex < 0 Then
            cboBranch.SetFocus
        Else
            cboBranch.ListIndex = cbolistIndex
        End If
    
        'Call cboBranch_Click
    End If
'********** END CODE UPDATED: 04/07/17 **************************
End Sub



Private Sub chkFinal_Click()
If chkFinal.Value = 1 Then
    txtDateFinal.BackColor = txtHoldWhiteBackColor.BackColor
End If
End Sub

Private Sub cmdTrans_Click(Index As Integer)
Dim strDateFinal As String
On Error GoTo er

    If Index = 0 Then 'add
    
        formAction = "add"
        AddNew = True
      
        Call clearform
        Call CboDisplay
        Call LockData(True)
        Call SetButtonVisibility(True)
        txtProcsBy.Text = UserName
    '***************** CODE UPDATED: 04/26/17 *****************
        RecordHolder = 0
        RecordHolderRow = 0
        chkFinal.Enabled = True
        chkFinal.Value = 0
        txtDatePrepared.BackColor = txtHoldWhiteBackColor.BackColor
        cboBranch.Text = "LYR MOTORPARTS CORP. - MAIN WAREHOUSE"
        SUPPLIERID = 146
        Call displayPM
    '************* END CODE UPDATED: 04/26/17 *****************
        
    ElseIf Index = 1 Then 'edit
        If flxSPARTS.TextMatrix(2, 1) = "" Then
            MsgBox "No available PARTS MATRIX to edit", vbOKOnly + vbExclamation, "Edit"
        Else
            Call modPMBranchIdentifying.sp_brnchIdentify_getRefNo_retail(sysBranchCode, intREFID, tmpRs)
            If tmpRs.RecordCount > 0 Then
                With tmpRs
                    If !STATUS = "FINAL" Then
                        MsgBox "Not Available to Edit" & vbCrLf & "This Parts Matrix is already Final", vbOKOnly + vbExclamation, "Edit"
                        Exit Sub
                    End If
                End With
            End If
            formAction = "edit"
            AddNew = False
            chkFinal.Enabled = True
            Call SetButtonVisibility(True)
            Call LockData(True)
            cboBranch.BackColor = txtHoldGrayBackColor.BackColor
            txtProcsBy.Text = UserName
            flxSPARTS.SetFocus
            flxSPARTS.Row = 2
            flxSPARTS.Col = 11
        End If
        
    ElseIf Index = 2 Then 'view summary
        frmSummaryRepl.mySrc = "Branch Identifying"
        'frm_OnProcess.OwnerForm = "View Summary Repl"
        frmSummaryRepl.Show
        Unload Me
        
    ElseIf Index = 5 Then 'exit
        
        formAction = "EXIT"
        ProperExit = True
        Call clearform
        Unload Me
        
    ElseIf Index = 6 Then 'save
    
        If ValidateEntry() = False Then Exit Sub
        
        myMsg = ValidationPartsMatrix
        
        If myMsg <> "" Then
            MsgBox myMsg, vbCritical + vbOKOnly, "Validation Error"
            Exit Sub
        
        End If
        
        Ans = MsgBox("Do you want to save?", vbQuestion + vbYesNo, "Save")
        
        If Ans = vbNo Then Exit Sub
            
        If formAction = "add" Then
            Call CreateTempTable
            DoEvents
            Call modPMBranchIdentifying.sp_brnchIdentify_create_retail(branchCode, Format(txtDatePrepared.Text, "YYYY-mm-dd"), Format(txtDateFinal.Text, "YYYY-mm-dd"), UserID, UserID, Val(chkFinal.Value), SUPPLIERID, mainRs)
            txtRefNo.Text = mainRs!RefNo
            txtpmID.Text = mainRs!REFID
        Else
            Call CreateTempTable
            DoEvents
            If txtDateFinal.Text = "__/__/____" Then
                strDateFinal = ""
            Else
                strDateFinal = Format(txtDateFinal.Text, "YYYY-mm-dd")
            End If
            Call modPMBranchIdentifying.sp_brnchIdentify_update_retail(branchCode, intREFID, txtRefNo.Text, Format(txtDatePrepared.Text, "YYYY-mm-dd"), strDateFinal, UserID, Val(chkFinal.Value), mainRs)
        End If
        If Val(chkFinal.Value) = 1 Then
            Dim TOTALPN As Integer
            frm_OnProcess.OwnerForm = "Identifying"
            frm_OnProcess.Show
            Call ConnectOnline
            TOTALPN = flxSPARTS.Rows - 2
            OnlineCon.Execute "Insert Into sporder( DEALERCODE, DEALERNAME, REFERENCE, ORDERTYPE, ORDERDATE, " & _
            "SUBMITTEDBY, UPDATEDBY, QTYPN, QTYORDER, STATUS)" & _
            "VALUES('" & sysBranchCode & "', '" & cboBranch.Text & "', '" & txtRefNo.Text & "', " & _
            "'REPLENISH','" & Format(txtDateFinal.Text, "yyyy-mm-dd") & "', " & _
            "'" & txtProcsBy.Text & "', " & UserID & ", " & _
            "" & TOTALPN & ", " & Val(Replace(txtTtlALQ.Text, ",", "")) & ",  'NEW' );"
    
            With flxSPARTS
                For i = 2 To .Rows - 2
                    If Val(.TextMatrix(i, 8)) <> 0 Then
                        OnlineCon.Execute "Insert Into sporderdtls( REFERENCE, SPAREPARTID, TRANSDATE, TRANSTYPE, `ORDER`, " & _
                        "SOURCE, PMATRIXID, UPDATEDBY, LASTUPDATE, SEQNO) VALUES( '" & txtRefNo.Text & "', " & _
                        "" & flxSPARTS.TextMatrix(i, 13) & ", '" & Format(txtDateFinal.Text, "yyyy-mm-dd") & "', " & _
                        "'REPLENISH', " & Val(flxSPARTS.TextMatrix(i, 8)) & ", '" & sysBranchCode & "', " & _
                        "'" & Val(flxSPARTS.TextMatrix(i, 12)) & "', " & UserID & ", '" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "', " & i & ");"
                    End If
                Next i
            End With
            OnlineCon.Close
    
            Call Unload(frm_OnProcess)
            LYRcon.Execute "UPDATE sporder SET `STATUS` = 'NEW' Where sporder.REFERENCE = '" & txtRefNo.Text & "';"
            chkFinal.Enabled = False
            MsgBox "Replenishment Request successfully saved online!", vbInformation + vbOKOnly, "Save"
            
        End If
        Call SetButtonVisibility(False)
        Call LockData(False)
        formAction = ""
        
    ElseIf Index = 7 Then 'cancel
        
        Ans = MsgBox("Do you want to cancel?", vbQuestion + vbYesNo)
        
        If Ans = vbNo Then Exit Sub
        
        Call SetButtonVisibility(False)
        Call LockData(False)
        chkFinal.Enabled = False
        formAction = ""
        If AddNew Then
            branchCode = ""
        End If
        
        If formAction = "edit" Then
        'record will back to normal
            flxSPARTS.Rows = 2
            flxSPARTS.Rows = 3
            Call ShowRecord(branchCode, 0, "FALSE", "")
        End If
    End If
    
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->cmdTrans-Click()")
End Sub

Sub CreateTempTable()
On Error GoTo er
        LYRcon.Execute "Drop Temporary Table IF  Exists sp_partsmatrix.TMP_BRNCHIDENTIFY_DTLS;"

         DoEvents
       LYRcon.Execute " Create Temporary Table IF Not Exists sp_partsmatrix.TMP_BRNCHIDENTIFY_DTLS( " & _
        "    PMALQ Int(11)            " & _
        "   ,PMATRIXID Int(11)        " & _
        "   ,SPAREPARTID Int(11)      " & _
        "   ,UNITPRICE DECIMAL(10,2) " & _
        "   ,UPDATEDBY INT(11) " & _
        "   ,SEQNO INT(11)      " & _
        ");"
        
        Qry = "Insert into sp_partsmatrix.TMP_BRNCHIDENTIFY_DTLS(PMALQ,PMATRIXID,SPAREPARTID,UNITPRICE,UPDATEDBY,SEQNO) VALUES "
        SEQNO = 0
        For i = 2 To flxSPARTS.Rows - 1
            If i > 2 Then Qry = Qry & ","
            SEQNO = SEQNO + 1
            Qry = Qry & "(" & Val(flxSPARTS.TextMatrix(i, 8)) & "," & Val(flxSPARTS.TextMatrix(i, 12)) & "," & Val(flxSPARTS.TextMatrix(i, 13)) & "," & Replace(flxSPARTS.TextMatrix(i, 9), ",", "") & ", " & UserID & "," & SEQNO & ") "
        Next i
        
       Call LYRcon.Execute(Qry)
       
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->CreateTempTable()")
End Sub

Private Sub flxSPARTS_Click()
On Error GoTo er

    ActRow = flxSPARTS.Row
    ActCol = flxSPARTS.Col
    
    If formAction = "add" Or formAction = "edit" Then
    
    If ActCol = 0 Or ActRow = 0 Then Exit Sub
        Select Case ActCol
        
        Case 11
            If flxSPARTS.TextMatrix(flxSPARTS.Row, 1) <> "" Then
                With txtALQ
                    If flxSPARTS.TextMatrix(ActRow, 6) <> "" Then .Tag = flxSPARTS.TextMatrix(ActRow, 6)
                    .Top = Frame3.Top + flxSPARTS.CellTop + flxSPARTS.Top
                    .Left = Frame3.Left + flxSPARTS.CellLeft + flxSPARTS.Left
                    .Height = flxSPARTS.CellHeight
                    .Width = flxSPARTS.CellWidth
                    .Text = flxSPARTS.Text
                    .Visible = True
                    .SelStart = 0
                    .SelLength = Len(.Text)
                    .ZOrder
                    .SetFocus
                End With
            Else
                Exit Sub
            End If
            
        End Select
    End If

Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->flxSPARTS-click()")
End Sub

Private Sub flxSPARTS_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then Call flxSPARTS_Click

End Sub

Private Sub flxSPARTS_Scroll()
txtALQ.Visible = False
End Sub

Private Sub flxSummary_DblClick()
If formAction = "add" Or formAction = "edit" Then
    If flxSummary.TextMatrix(1, 1) <> "" Then
    Ans = MsgBox("Do you want to add this partsmatrix?", vbQuestion + vbYesNo, "Parts Matrix")
        If Ans = vbNo Then Exit Sub
    
        ActRow = flxSummary.Row
        strFrmSource = "clckSummary"
        Call ShowRecord(branchCode, Val(flxSummary.TextMatrix(ActRow, 5)), "TRUE", "")
        If flxSummary.Rows = 2 Then
            flxSummary.Rows = 1
            flxSummary.Rows = 2
            Exit Sub
        Else
            flxSummary.RemoveItem (ActRow)
        End If
    Else
        Exit Sub
    End If
Else
    Exit Sub
End If
End Sub

Private Sub Form_Load()
    Call LockData(False)

    txtRefNo.Text = "<Auto>"
    txtDatePrepared.Text = Format(Date, "mm/dd/yyyy")
    txtDateFinal.Text = Format(Date, "mm/dd/yyyy")
    
'***************** CODE UPDATED: 04/26/17 *****************
    txtTtlALQ.Text = 0
    chkFinal.Enabled = False
    ActRow = 1
'************* END CODE UPDATED: 04/26/17 *****************

    '===================BUTTON POSITION===================
    
    If cmdTrans(0).Visible = False Then
            'Align Save Button
            cmdTrans(6).Left = cmdTrans(0).Left
            cmdTrans(6).Top = cmdTrans(0).Top
            'Align Cancel Button
            cmdTrans(7).Left = cmdTrans(1).Left
            cmdTrans(7).Top = cmdTrans(1).Top
    End If
End Sub



Private Sub cboBranch_Change()
On Error GoTo er
    If StatSrch Then
            StatSrch = False
            Exit Sub
    End If

    With cboBranch
        strPartial = .Text
        i = SendMessage(.hwnd, CB_FINDSTRING, -1, ByVal strPartial)  'If match found, append unmatched characters
        If i <> -1 Then
            strTotal = .List(i)
        
            j = Len(strTotal) - Len(strPartial)
           
            If j <> 0 Then
                StatSrch = True
                .SelText = Right$(strTotal, j)
                .SelStart = Len(strPartial)
                .SelLength = j
            End If
        End If
    End With
    
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->cboBranch-Change()")
End Sub

Public Sub CboDisplay()
On Error GoTo er
    If formAction = "add" Then
        'Get Data from database
        DoEvents
        Call modPMBranchIdentifying.sp_get_Branch(tmpRs)
        If tmpRs.State = 1 Then
           Call modObjTools.addToListObjectWithArrayGivenField(cboBranch, tmpRs, arrBRANCH, "BRANCHCODE", "BRANCHNAME")
        End If
        If tmpRs.State <> 0 Then tmpRs.Close
    End If
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->CboDisplay()")
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo er

    If ProperExit Then
        If mainRs.State = adStateOpen Then mainRs.Close
        Call cmdTrans_Click(5)
        ProperExit = False
    Else
        Cancel = True
    End If
    
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->QueryUnload()")
End Sub

Private Sub Form_Resize()
Call RXResizer1_ResizeComplete
End Sub



Private Sub RXResizer1_ResizeComplete()
    Call flxColChanger
    
'************* CODE UPDATED: 3/14/17 *****************
    '    'Frame width
    '    Frame1.Width = Me.Width - 450
    '    Frame2.Width = Me.Width - 450
    '    Frame3.Width = Me.Width - 450
    '
    '    'Frame1 label Left
    '    lblBranch.Left = 120
    '    lblMCBRAND.Left = 120
    '    lblMCMODEL.Left = 120
    '    lblTOTALPN.Left = 2760
    '    lblTOTALSRQ.Left = 2760
    '    lblTOTALAVQ.Left = 5160
    '    lblTOTALALQ.Left = 5160
    '    lblRefNo.Left = Frame1.Width - 3255
    '    lblDatePrepared.Left = Frame1.Width - 3255
    '    lblDateFinal.Left = Frame1.Width - 3255
    '
    '
    '    'Frame1 Label Width
    '    lblBranch.Width = 975
    '    lblTOTALPN.Width = 975
    '    lblTOTALSRQ.Width = 975
    '    lblTOTALAVQ.Width = 975
    '    lblTOTALALQ.Width = 975
    '    lblMCBRAND.Width = 975
    '    lblMCMODEL.Width = 975
    '    lblRefNo.Width = 1455
    '    lblDatePrepared.Width = 1455
    '    lblDateFinal.Width = 1455
    '
    '
    '    'Frame1 Combobox Left
    '    cboBranch.Left = 1200
    '    cboBrand.Left = 1200
    '    cboModel.Left = 1200
    '
    '    'Frame1 Combobox Width
    '    cboBranch.Width = 6255
    '    cboBrand.Width = 1575
    '    cboModel.Width = 1575
    '
    '    'Frame1 Text Left
    '    txtTotalPN.Left = 3840
    '    txtTotalSRQ.Left = 3840
    '
    '    txtTotalAVQ.Left = 6240
    '    txtTotalALQ.Left = 6240
    '    txtRefNo.Left = Frame1.Width - 1695
    '    txtDatePrepared.Left = Frame1.Width - 1695
    '    txtDateFinal.Left = Frame1.Width - 1695
    '
    '    'Frame1 Text Width
    '    txtTotalPN.Width = 1215
    '    txtTotalSRQ.Width = 1215
    '    txtTotalAVQ.Width = 1215
    '    txtTotalALQ.Width = 1215
    '    txtRefNo.Width = 1335
    '    txtDatePrepared.Width = 1335
    '    txtDateFinal.Width = 1335
    '
    '    'Frame2 Label Left
    '    lblProcBy.Left = Frame3.Width - 6255
    '
    '    'Frame2 Label Width
    '    lblProcBy.Width = 1935
    '
    '    'Frame2 Text Left
    '    txtProcsBy.Left = Frame2.Width - 4215
    '
    '    'Frame2 Text Width
    '    txtProcsBy.Width = 3975
    '
    '    'Frame3 FlexGrid Width
    '    flxSPARTS.Width = Frame3.Width - 240
    '
    '    'Frame3 FlexGrid Height
    '    flxSPARTS.Height = Frame3.Height - 360
'**********END CODE UPDATED: 3/14/17 *****************
End Sub


Sub LockData(DataEntry As Boolean)
    DataEntry = Not DataEntry

    cboBranch.Locked = DataEntry
    txtRefNo.Locked = DataEntry
    'txtDatePrepared.Locked = DataEntry
    'txtDateFinal.Locked = DataEntry
    txtProcsBy.Locked = DataEntry
    txtTtlALQ.Locked = DataEntry
   
    If cboBranch.Locked Then
        cboBranch.BackColor = txtHoldGrayBackColor.BackColor
        txtRefNo.BackColor = txtHoldGrayBackColor.BackColor
        txtDatePrepared.BackColor = txtHoldGrayBackColor.BackColor
        txtDateFinal.BackColor = txtHoldGrayBackColor.BackColor
        txtProcsBy.BackColor = txtHoldGrayBackColor.BackColor
        txtTtlALQ.BackColor = txtHoldGrayBackColor.BackColor
        
    Else
        cboBranch.Locked = True
        'txtDatePrepared.BackColor = txtHoldWhiteBackColor.BackColor
        'txtDateFinal.BackColor = txtHoldWhiteBackColor.BackColor
        txtProcsBy.Locked = True
        txtRefNo.Locked = True
        txtTtlALQ.Locked = True
        
    End If
End Sub


Private Sub SetButtonVisibility(DataEntry As Boolean)
'cmdTrans Legend
'0 ->ADD
'1 ->EDIT
'2 ->VIEW SUMM
'3 ->EXIT
'4 ->SAVE
'5 ->CANCEL


    'hide all button
    For i = 0 To 5
        cmdTrans(i).Visible = False
    Next i

    If DataEntry Then
        cmdTrans(6).Visible = True
        cmdTrans(7).Visible = True
    Else
    For i = 0 To 5
        cmdTrans(i).Visible = True
        cmdTrans(6).Visible = False
        cmdTrans(7).Visible = False
    Next i
    
    
    End If
End Sub

Sub clearform()
    
    flxSummary.Rows = 1
    flxSummary.Rows = 2
    flxSPARTS.Rows = 2
    flxSPARTS.Rows = 3
    txtpmID.Text = ""
    cboBranch.Text = ""
    txtRefNo.Text = "<Auto>"
    txtDatePrepared.Text = Format(Date, "mm/dd/yyyy")
    txtDateFinal.Text = Format(Date, "mm/dd/yyyy")
    txtTtlALQ.Text = 0
    branchCode = ""
    
End Sub

Public Sub getREFID(myREFID As Integer)
    intREFID = myREFID
    RecordHolder = 0
    RecordHolderRow = 0
    flxSPARTS.Rows = 2
    flxSPARTS.Rows = 3
End Sub


Public Sub ShowRecord(MyBranchCode As String, MyID As Integer, MyBoolean As String, myStatus As String)
Dim QH As Integer, RO As Integer
On Error GoTo er

    Screen.MousePointer = vbHourglass
    branchCode = MyBranchCode 'use to verify to edit
    txtpmID.Text = MyID
    intID = MyID
    Call ConnectMain
'************* CODE UPDATED: 3/14/17 *****************
 Pbar.Visible = True
 Pbar.Value = 0
 frm_OnProcess.OwnerForm = "Identifying"
 frm_OnProcess.Show
'********** END CODE UPDATED: 3/14/17 ****************

 DoEvents
 Call modPMBranchIdentifying.sp_get_BranchName(MyBranchCode, mainRs)
    cboBranch.Text = mainRs!branchName
    
    
'************* CODE UPDATED: 3/14/17 *****************
    If myStatus = "PENDING" Then
        chkFinal.Value = 0
        chkFinal.Enabled = False
    ElseIf myStatus = "FINAL" Then
        chkFinal.Value = 1
        chkFinal.Enabled = False
    End If
'********* END CODE UPDATED: 3/14/17 *****************
    
    
'***************** CODE UPDATED: 04/26/17 *****************
If MyBoolean = "FALSE" Then 'gkan ang data sa summary
 Call modPMBranchIdentifying.sp_brnchIdentify_getDetails_total_retail(MyBranchCode, intREFID, MyBoolean, mainRs)
    With mainRs
        If .RecordCount > 0 Then
            flxSummary.Rows = 1
            Do While Not .EOF
                flxSummary.AddItem vbTab & !MODELBRAND & vbTab & !TOTALPN & vbTab & !TOTALQH & vbTab & !TOTALSRQ & vbTab & !PMATRIXID
                .MoveNext
            Loop
        Else
            flxSummary.Rows = 1
            flxSummary.Rows = 2
        End If
    End With
    
End If
'************* END CODE UPDATED: 04/26/17 *****************

 DoEvents
 Call modPMBranchIdentifying.sp_brnchIdentify_getDetails_retail(MyBranchCode, MyID, MyBoolean, intREFID, tmpRs)
    
    If tmpRs.State <> 1 Then Exit Sub
    
    If tmpRs.RecordCount <= 0 Then
    '************* CODE UPDATED: 3/18/17 *****************
        Call modPMBranchIdentifying.sp_brnchIdentify_getRefNo_retail(MyBranchCode, intREFID, mainRs)
            If mainRs.RecordCount > 0 Then
                tmpREFNO.Text = mainRs!REFERENCE
                tmpSTATUS.Text = mainRs!STATUS
                MsgBox "This Parts Matrix has been used in REF# " + tmpREFNO.Text + " and the status is " + tmpSTATUS.Text, vbInformation + vbOKOnly, "Show Details"
            Else
                MsgBox "This Parts Matrix is not yet Available", vbInformation + vbOKOnly, "Show Details"
            End If
    '********** END CODE UPDATED: 3/18/17 ****************
        Call Unload(frm_OnProcess)
        Pbar.Visible = False
        Pbar.Value = 0
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    With tmpRs
        If .RecordCount > 0 Then
            .MoveFirst
            'j = 0
            RecordHolder = RecordHolder + .RecordCount
            
            If MyBoolean = "TRUE" Then
                RecordHolderRow = flxSPARTS.Rows
            Else
                RecordHolderRow = flxSPARTS.Rows - 1
            End If
            
            flxSPARTS.Rows = RecordHolderRow + .RecordCount
        
            Pbar.Max = .RecordCount
            
            For i = RecordHolderRow To flxSPARTS.Rows - 1

                If !QtyOnHand < !PMSTDQ Then
                    intSRQ = Abs(!QtyOnHand - !PMSTDQ)
                    ALQ = intSRQ
                ElseIf !QtyOnHand >= !PMSTDQ Then
                    intSRQ = 0
                    ALQ = 0
                End If
                'flxSPARTS.TextMatrix(i, 0) = j
                flxSPARTS.TextMatrix(i, 1) = !MODELBRAND
                flxSPARTS.TextMatrix(i, 2) = !PartNumber
                flxSPARTS.TextMatrix(i, 3) = !Description
                flxSPARTS.TextMatrix(i, 4) = !partsGroup
                flxSPARTS.TextMatrix(i, 5) = !QtyOnHand 'from mrc2retail.spareparts
                flxSPARTS.TextMatrix(i, 6) = !PMSTDQ
                flxSPARTS.TextMatrix(i, 7) = Format$(intSRQ, 0)
                flxSPARTS.TextMatrix(i, 8) = Format$(ALQ, "") '!PMALQ
                flxSPARTS.TextMatrix(i, 9) = Format(IIf(IsNull(!UNITPRICE), "", !UNITPRICE), "###,##0.00")
                flxSPARTS.TextMatrix(i, 11) = Format$(!Location, "")
                flxSPARTS.TextMatrix(i, 12) = !PMATRIXID 'from mrcwarehouse.spareparts
                flxSPARTS.TextMatrix(i, 13) = !SparepartID
                'flxSPARTS.TextMatrix(i, 14) = !SparepartID
                .MoveNext
                Pbar.Value = Pbar.Value + 1
                Next i
        End If
        If MyBoolean = "TRUE" Then
            If flxSPARTS.TextMatrix(2, 1) = "" Then
                flxSPARTS.RemoveItem (2)
            End If
        End If
    End With
    

    Call deleteDuplicatePN
    Call autoSetUpTotalPrice
    Call ComputeTotalAllocation
    Call deleteZeroOrder


    If tmpRs.State <> 0 Then
        Screen.MousePointer = vbDefault
        Call Unload(frm_OnProcess)
        tmpRs.Close
    End If

'***************** CODE UPDATED: 04/26/17 *****************
    If MyBoolean = "FALSE" Then
        Call modPMBranchIdentifying.sp_brnchIdentify_dsplayHeader_retail(intREFID, mainRs)
        With mainRs
            If .RecordCount > 0 Then
                txtRefNo.Text = Format$(!RefNo, "")
                txtDatePrepared.Text = Format$(IIf(IsNull(!DATEPREPARED), "", !DATEPREPARED), "mm/dd/yyyy")
                txtDateFinal.Text = IIf(IsNull(!DATEFINAL), "__/__/____", !DATEFINAL)
                txtProcsBy.Text = Format$(!ProcessedBy, "")
            End If
        End With
        Call ComputeTotalAllocation
        strBoolean = "FALSE"
    End If
'************* END CODE UPDATED: 04/26/17 *****************

    Call Unload(frm_OnProcess)
    Pbar.Visible = False
    Pbar.Value = 0
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->branchIdentifyingShowRecord()")
End Sub


Private Sub txtALQ_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
    Tmp = Split(txtALQ.Tag, "|")
    
    For i = LBound(Tmp) To UBound(Tmp)
        flxSPARTS.TextMatrix(ActRow, i + 1) = Tmp(i)
    Next i
    
    txtALQ.Visible = False
    flxSPARTS.SetFocus
    flxSPARTS.Col = 11
End If
End Sub

Public Sub ComputeTotalAllocation()
    intTotalALQ = 0
    With flxSPARTS
        For i = 2 To flxSPARTS.Rows - 1
            intTotalALQ = intTotalALQ + Val(Replace(.TextMatrix(i, 8), ",", ""))
        Next i
    End With
    txtTtlALQ.Text = Format(intTotalALQ, "0")
End Sub

'************* CODE UPDATED: 4/28/17 *****************
Public Sub deleteDuplicatePN()
Dim lCounter As Long
Dim lCounter2 As Long
Dim lTemp As Long
Dim lAccumulator As Long

lTemp = flxSPARTS.Rows - 1
lAccumulator = 0

For lCounter = 2 To lTemp
    For lCounter2 = lCounter + 1 To lTemp
        ' now compare every item to current item from main loop
        If flxSPARTS.TextMatrix(lCounter, 2) = flxSPARTS.TextMatrix(lCounter2 - lAccumulator, 2) Then
            flxSPARTS.RemoveItem (lCounter2 - lAccumulator)
            lTemp = lTemp - 1
            lAccumulator = lAccumulator + 1
        End If
    Next lCounter2
   lAccumulator = 0
Next lCounter
    
End Sub
'********* END CODE UPDATED: 4/28/17 *****************



'************* CODE UPDATED: 4/28/17 *****************
Public Sub deleteZeroOrder()
Dim lCounter As Long
Dim lCounter2 As Long
Dim lTemp As Long
Dim lAccumulator As Long


With flxSPARTS
    For i = .Rows - 1 To 2 Step -1
        If flxSPARTS.Rows > 3 Then
            If .TextMatrix(i, 8) = 0 Then
                .RemoveItem (i)
            End If
        Else
            If flxSPARTS.TextMatrix(flxSPARTS.Rows - 1, 1) = "" Then
                flxSPARTS.Rows = 2
                flxSPARTS.Rows = 3
            End If
        End If
    Next i
End With
    

    
End Sub
'********* END CODE UPDATED: 4/28/17 *****************

'************* CODE UPDATED: 4/29/17 *****************
Public Sub autoSetUpTotalPrice()
    For i = 2 To flxSPARTS.Rows - 1
        flxSPARTS.TextMatrix(i, 10) = Format(Val(Replace(flxSPARTS.TextMatrix(i, 9), ",", "")) * Val(flxSPARTS.TextMatrix(i, 8)), "###,##0.00")
    Next i
End Sub
'********* END CODE UPDATED: 4/29/17 *****************

Private Sub txtALQ_KeyPress(KeyAscii As Integer)
On Error GoTo er

    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtALQ, KeyAscii, NumbersOnly)

    
    If KeyAscii = 13 Then
        
        flxSPARTS.TextMatrix(ActRow, 11) = txtALQ.Text
        
        If Val(flxSPARTS.TextMatrix(ActRow, 11)) > Val(flxSPARTS.TextMatrix(ActRow, 7)) Then
            MsgBox "Allocated Quantity is Greater Than to Available Quantity", vbInformation + vbOKOnly, "Validation"
            flxSPARTS.TextMatrix(ActRow, 11) = 0
            txtALQ.Text = ""
            txtALQ.Visible = False
            flxSPARTS.SetFocus
            flxSPARTS.Row = ActRow
            flxSPARTS.Col = 11
            Exit Sub
        End If
        
        
        If txtALQ.Text <> "" Then
                
            For i = ActRow To flxSPARTS.Rows - 1
                If flxSPARTS.TextMatrix(i, 11) = "" Or Val(flxSPARTS.TextMatrix(i, 11)) = 0 Then
                    txtALQ.Visible = False
                    flxSPARTS.SetFocus
                    flxSPARTS.Row = i
                    flxSPARTS.Col = 11
                Exit For
                Else
                    txtALQ.Visible = False
                    flxSPARTS.SetFocus
                    flxSPARTS.Row = ActRow
                    flxSPARTS.Col = 11
                End If
                
            Next i
        
        Else
            flxSPARTS.SetFocus
            flxSPARTS.Col = 11
            txtALQ.Visible = False
        End If
        Call ComputeTotalAllocation
        
    End If

Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->txtALQ-KeyPress()")
End Sub

Private Sub txtALQ_LostFocus()
On Error GoTo er

    If txtALQ.Tag = "skiplostfocus" Then
        txtALQ.Tag = ""
        Exit Sub
    End If
    
    If flxSPARTS.TextMatrix(ActRow, 11) = "" Or txtALQ.Text = "" Then
        MsgBox "Wrong Entry of Allocate Quantity!" & vbCrLf & "Please Enter a Correct Allocate Quantity.", vbExclamation + vbOKOnly, "Warning"
        txtALQ.Visible = False
        flxSPARTS.TextMatrix(ActRow, 11) = ""
        flxSPARTS.SetFocus
        flxSPARTS.Col = 11
        
        Exit Sub
    End If
    
    flxSPARTS.TextMatrix(ActRow, 11) = txtALQ.Text
    txtALQ.Visible = False
    
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->txtALQ-LostFocus()")
End Sub


Private Sub txtDateFinal_GotFocus()
    txtDateFinal.SelStart = 0
    txtDateFinal.SelLength = Len(txtDateFinal.Text)
End Sub

Private Sub txtDateFinal_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtDateFinal, KeyAscii, NumbersOnly)
    If KeyAscii = 13 Then
        flxSPARTS.SetFocus
        flxSPARTS.Row = 2
        flxSPARTS.Col = 11
    End If
End Sub

Private Sub txtDatePrepared_GotFocus()
    txtDatePrepared.SelStart = 0
    txtDatePrepared.SelLength = Len(txtDatePrepared.Text)
End Sub

Private Sub txtDatePrepared_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtDatePrepared, KeyAscii, NumbersOnly)
    If KeyAscii = 13 Then txtDateFinal.SetFocus
End Sub


Private Sub txtPROCSBY_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtProcsBy, KeyAscii, ValidChars)
End Sub

Private Sub txtrefno_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtRefNo, KeyAscii, ValidChars)
    If KeyAscii = 13 Then txtDatePrepared.SetFocus
End Sub


Function ValidationPartsMatrix() As String

    MyErr = ""
    If flxSPARTS.TextMatrix(flxSPARTS.Rows - 1, 1) = "" Then
        If flxSPARTS.Rows = 2 Then
            flxSPARTS.Rows = 1
            flxSPARTS.Rows = 2
        Else
            flxSPARTS.RemoveItem flxSPARTS.Rows - 1
        End If
    End If
        
    ValidationPartsMatrix = MyErr

End Function

Function ValidateEntry() As Boolean
  ValidateEntry = False
    
    If Trim(cboBranch.Text) = "" Then
        MsgBox "BRANCH IS REQUIRED!", vbExclamation, "Validation"
        cboBranch.SetFocus
        Exit Function
    End If
    
    ValidateEntry = True

End Function
