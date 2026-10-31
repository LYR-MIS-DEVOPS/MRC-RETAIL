VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form BranchSQTY 
   Caption         =   "Branch Quantity"
   ClientHeight    =   9150
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   19170
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Branch SQTY.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9150
   ScaleWidth      =   19170
   WindowState     =   2  'Maximized
   Begin VB.ComboBox cboOrderFrequency 
      Height          =   315
      ItemData        =   "Branch SQTY.frx":57E2
      Left            =   19320
      List            =   "Branch SQTY.frx":57EC
      TabIndex        =   42
      Tag             =   "YX"
      Top             =   6120
      Visible         =   0   'False
      Width           =   975
   End
   Begin lvButton.lvButtons_H cmdLoadData 
      Height          =   375
      Left            =   5880
      TabIndex        =   38
      Tag             =   "YX"
      Top             =   0
      Width           =   1920
      _ExtentX        =   3387
      _ExtentY        =   661
      Caption         =   "&Update Parts Matrix"
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
      Image           =   "Branch SQTY.frx":57F8
      cBack           =   -2147483633
   End
   Begin MSFlexGridLib.MSFlexGrid flxPG 
      Height          =   1815
      Left            =   19320
      TabIndex        =   36
      Top             =   4200
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   3201
      _Version        =   393216
   End
   Begin VB.TextBox txtHoldGrayBackColor 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   19320
      Locked          =   -1  'True
      TabIndex        =   28
      Top             =   3720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtHoldWhiteBackColor 
      Height          =   375
      Left            =   19320
      TabIndex        =   27
      Text            =   "HoldBackColor"
      Top             =   3240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   120
      TabIndex        =   22
      Tag             =   "YX"
      Top             =   9480
      Width           =   18975
      Begin VB.TextBox txtTOTALAMOUNT 
         Height          =   285
         Left            =   17040
         TabIndex        =   44
         Tag             =   "YX"
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox txtUpdated 
         Height          =   285
         Left            =   1440
         TabIndex        =   6
         Tag             =   "YX"
         Top             =   240
         Width           =   3975
      End
      Begin VB.Label lblTtlAmount 
         Caption         =   "TOTAL AMOUNT"
         Height          =   255
         Left            =   15720
         TabIndex        =   45
         Tag             =   "YX"
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblUP 
         Caption         =   "UPDATED BY:"
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Tag             =   "YX"
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   19320
      Top             =   720
   End
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   120
      TabIndex        =   10
      Tag             =   "YX"
      Top             =   480
      Width           =   18975
      Begin VB.TextBox txtEmpty 
         Height          =   285
         Left            =   17400
         TabIndex        =   7
         Tag             =   "YX"
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox txtLacking 
         Height          =   285
         Left            =   17400
         TabIndex        =   8
         Tag             =   "YX"
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox txtExcess 
         Height          =   285
         Left            =   17400
         TabIndex        =   9
         Tag             =   "YX"
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox txtPG 
         Height          =   285
         Left            =   14880
         TabIndex        =   34
         Tag             =   "YX"
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox txtTTSTDQ 
         Height          =   285
         Left            =   14880
         TabIndex        =   33
         Tag             =   "YX"
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox txtTTPN 
         Height          =   285
         Left            =   14880
         TabIndex        =   32
         Tag             =   "YX"
         Top             =   240
         Width           =   1455
      End
      Begin VB.ComboBox cmbModel 
         Height          =   315
         Left            =   1800
         TabIndex        =   4
         Tag             =   "YX"
         Top             =   840
         Width           =   1935
      End
      Begin VB.ComboBox cmbBrand 
         Height          =   315
         Left            =   1800
         TabIndex        =   3
         Tag             =   "YX"
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label lblEmpty 
         Alignment       =   1  'Right Justify
         Caption         =   "EMPTY"
         Height          =   255
         Left            =   16560
         TabIndex        =   20
         Tag             =   "YX"
         Top             =   960
         Width           =   735
      End
      Begin VB.Label lblLacking 
         Alignment       =   1  'Right Justify
         Caption         =   "LACKING"
         Height          =   255
         Left            =   16560
         TabIndex        =   19
         Tag             =   "YX"
         Top             =   600
         Width           =   735
      End
      Begin VB.Label lblExcess 
         Alignment       =   1  'Right Justify
         Caption         =   "EXCESS"
         Height          =   255
         Left            =   16560
         TabIndex        =   18
         Tag             =   "YX"
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lblTotalPG 
         Alignment       =   1  'Right Justify
         Caption         =   "TOTAL PARTS GROUP"
         Height          =   255
         Left            =   12840
         TabIndex        =   17
         Tag             =   "YX"
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label lblTotalSTDQ 
         Alignment       =   1  'Right Justify
         Caption         =   "TOTAL STDQ"
         Height          =   255
         Left            =   13440
         TabIndex        =   16
         Tag             =   "YX"
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label lblTotalPN 
         Alignment       =   1  'Right Justify
         Caption         =   "TOTAL PN"
         Height          =   255
         Left            =   13920
         TabIndex        =   15
         Tag             =   "YX"
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lblModel 
         Caption         =   "SELECT MC MODELS"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Tag             =   "YX"
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label lblBrand 
         Caption         =   "SELECT MC BRAND"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Tag             =   "YX"
         Top             =   360
         Width           =   1575
      End
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   1
      Left            =   2040
      TabIndex        =   0
      Tag             =   "YX"
      Top             =   0
      Width           =   1920
      _ExtentX        =   3387
      _ExtentY        =   661
      Caption         =   "&Edit Parts Matrix"
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
      Image           =   "Branch SQTY.frx":5D92
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   4
      Left            =   19320
      TabIndex        =   25
      Top             =   2280
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
      Image           =   "Branch SQTY.frx":632C
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   5
      Left            =   19320
      TabIndex        =   26
      Top             =   2760
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
      Image           =   "Branch SQTY.frx":68C6
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   2
      Left            =   3960
      TabIndex        =   1
      Tag             =   "YX"
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
      Image           =   "Branch SQTY.frx":6E60
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   3
      Left            =   7800
      TabIndex        =   2
      Tag             =   "YX"
      Top             =   0
      Width           =   1920
      _ExtentX        =   3387
      _ExtentY        =   661
      Caption         =   "E&xit Parts Matrix"
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
      Image           =   "Branch SQTY.frx":73FA
      cBack           =   -2147483633
   End
   Begin VB.Frame Frame3 
      Height          =   7695
      Left            =   120
      TabIndex        =   24
      Tag             =   "YX"
      Top             =   1800
      Width           =   18975
      Begin TabDlg.SSTab tabProcessing 
         Height          =   975
         Left            =   7800
         TabIndex        =   39
         Tag             =   "YX"
         Top             =   1920
         Visible         =   0   'False
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   1720
         _Version        =   393216
         Tabs            =   1
         TabsPerRow      =   1
         TabHeight       =   520
         TabCaption(0)   =   "Load Data"
         TabPicture(0)   =   "Branch SQTY.frx":7994
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label21"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Image2"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "lblProcess"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).ControlCount=   3
         Begin VB.Label lblProcess 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Updating Parts Matrix ..."
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
            TabIndex        =   41
            Tag             =   "YX"
            Top             =   600
            Width           =   3975
         End
         Begin VB.Image Image2 
            Height          =   720
            Left            =   240
            Picture         =   "Branch SQTY.frx":79B0
            Stretch         =   -1  'True
            Tag             =   "xyf"
            Top             =   120
            Width           =   720
         End
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
            TabIndex        =   40
            Tag             =   "YX"
            Top             =   360
            Width           =   3255
         End
      End
      Begin MSFlexGridLib.MSFlexGrid flxBSQTY 
         Height          =   6975
         Left            =   120
         TabIndex        =   11
         Tag             =   "YX"
         Top             =   600
         Width           =   18735
         _ExtentX        =   33046
         _ExtentY        =   12303
         _Version        =   393216
         Cols            =   13
         FixedCols       =   0
         RowHeightMin    =   300
         BackColorBkg    =   14737632
         AllowUserResizing=   1
         FormatString    =   " |<PARTNUMBER |<DESCRIPTION |^PARTSGROUP |^MC MODEL  |^QOH |^AVQ 3MSQ |^NBS |^STDQ |^LACKING |^EXCESS |^ORDER FREQ"
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
      Begin lvButton.lvButtons_H btnFind 
         Height          =   375
         Left            =   15600
         TabIndex        =   12
         Tag             =   "YX"
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
         Image           =   "Branch SQTY.frx":83A8
         cBack           =   -2147483633
      End
      Begin VB.TextBox txtFind 
         Height          =   285
         Left            =   16920
         TabIndex        =   5
         Tag             =   "YX"
         Top             =   195
         Width           =   1935
      End
      Begin VB.CommandButton btnAltA 
         Caption         =   "&A"
         Height          =   315
         Left            =   9360
         TabIndex        =   30
         Top             =   2280
         Width           =   495
      End
      Begin VB.CommandButton btnAltD 
         Caption         =   "&D"
         Height          =   375
         Left            =   9120
         TabIndex        =   31
         Top             =   1800
         Width           =   495
      End
      Begin lvButton.lvButtons_H cmdCompute 
         Height          =   375
         Left            =   13680
         TabIndex        =   43
         Tag             =   "YX"
         Top             =   150
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         Caption         =   "COMPUTE STDQ"
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
         cFHover         =   255
         LockHover       =   2
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         ImgAlign        =   1
         Image           =   "Branch SQTY.frx":8942
         cBack           =   -2147483633
      End
      Begin VB.Label lblPartsGroup 
         Caption         =   "Parts Group"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   240
         TabIndex        =   37
         Top             =   240
         Width           =   11055
      End
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   35
      Tag             =   "YX"
      Top             =   0
      Width           =   1920
      _ExtentX        =   3387
      _ExtentY        =   661
      Caption         =   "&Add Parts Matrix"
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
      Image           =   "Branch SQTY.frx":8EDC
      cBack           =   -2147483633
   End
   Begin VB.TextBox txtPMID 
      Height          =   375
      Left            =   19320
      TabIndex        =   29
      Top             =   1800
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtSTDQ 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   19320
      TabIndex        =   21
      Tag             =   "YX"
      Top             =   1440
      Visible         =   0   'False
      Width           =   495
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   19320
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   19410
      m_MDI_Height    =   9720
   End
   Begin VB.Menu mnuStatus 
      Caption         =   "PopupMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuAdd 
         Caption         =   "AltA"
      End
   End
End
Attribute VB_Name = "BranchSQTY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Strings variable
Dim formAction As String
Dim strPartial As String, strTotal As String
Dim MyErr As String, myMsg As String
Dim Qry As String
Dim strTtlExcess As String
Dim strTtlLacking As String
Dim strTtlEmpty As String
Dim strTtlQOH As String
Dim Tmp() As String
Dim Ans As String
Dim SixMSQ As String, Lacking As String, Excess As String, STDQ As String
Dim pPMSO As String
Dim strBoolean As String
Dim strPG As String
Dim intPG As Integer
Dim NewPrtsGroup As String
Dim OldPrtsGroup As String
Dim cmbstr As String

Dim arrMODEL() As String
Dim arrBRAND() As String

Public mySrc As String

'recordset variable
Dim mainRs As New ADODB.Recordset
Dim tmpRs  As New ADODB.Recordset
Dim tmpLDRs As New ADODB.Recordset

'integer variable
Dim i As Long, j As Long, a As Integer
Dim ActRow As Integer, ActCol As Integer, fRow As Integer, fCol As Integer
Dim LoopSTDQ As Integer
Dim intTotalSTDQ As Integer
Dim LoopLACKING As Integer
Dim LoopExcess As Integer
Dim LoopEmpty As Integer
Dim LoopQOH As Integer
Dim flxLack As Double
Dim matrixID As Integer
Dim pSPAREPARTID As Integer

'******************* CODE UPDATED: 04/08/17 *************************
Dim intComputeTtlLacking As Integer, intComputeTtlExcess As Integer
Dim intComputeTtlEmpty As Integer, intComputeTtlQtyOnHand As Integer, intComputeTtlAmount As Double
Dim intComputeTtlSTDQ As Integer, intComputeTtlPN As Integer, intComputePG As Integer
'*************** END CODE UPDATED: 04/08/17 *************************

Dim intQOH As Double
Dim intExcess As Double
Dim intLacking As Double
Dim intEmpty As Double
Dim intSTDQ As Double
Dim intFlxEmpty As Double
Dim intTotlQOH As Double
Dim pPMSTDQ As Double
Dim pPMLACKING As Double
Dim pPMEXCESS As Double

'Boolean variable
Dim StatSrch As Boolean, AddNew As Boolean, ProperExit As Boolean
Dim PrtBln As Boolean, PrtSrch As Boolean, cboBln As Boolean, blnFound As Boolean

Public Sub CreateTempTable()
On Error GoTo er
        LYRcon.Execute "Drop Temporary Table IF  Exists sp_partsmatrix.TMP_BRNCHQTY_DTLS;"

         DoEvents
       LYRcon.Execute " Create Temporary Table IF Not Exists sp_partsmatrix.TMP_BRNCHQTY_DTLS( " & _
        "   PMATRIXID int(11),           " & _
        "   SPAREPARTID Int(11),          " & _
        "   PMSTDQ int(11),              " & _
        "   PMORDERFREQ varchar(20),            " & _
        "   UPDATEDBY int(11)            " & _
        ");"

      ' "   PMATRIXID Int(11),           "
        Qry = "insert into sp_partsmatrix.TMP_BRNCHQTY_DTLS(PMATRIXID,SPAREPARTID,PMSTDQ,PMORDERFREQ,UPDATEDBY) VALUES "
        
        For i = 1 To flxBSQTY.Rows - 1
            If i > 1 Then Qry = Qry & ","
        
           Qry = Qry & "(" & txtpmID.Text & "," & Val(flxBSQTY.TextMatrix(i, 12)) & "," & Val(flxBSQTY.TextMatrix(i, 8)) & ",'" & flxBSQTY.TextMatrix(i, 11) & "', " & UserID & ")"
    
        Next

       Call LYRcon.Execute(Qry)
        
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->CreateTempTable()")
End Sub

Function ValidationPartsMatrix() As String

    MyErr = ""
    If flxBSQTY.TextMatrix(flxBSQTY.Rows - 1, 1) = "" Then
        If flxBSQTY.Rows = 2 Then
            flxBSQTY.Rows = 1
            flxBSQTY.Rows = 2
        Else
            flxBSQTY.RemoveItem flxBSQTY.Rows - 1
            For i = 1 To flxBSQTY.Rows - 1
                If flxBSQTY.TextMatrix(i, 1) <> "" Then
                    txtTTPN.Text = i
                End If
            Next i
        End If
    End If
    
    For i = 1 To flxBSQTY.Rows - 1
        If flxBSQTY.TextMatrix(i, 8) = "" Then MyErr = MyErr & vbCrLf & "There is no Standard Quantity @ Row " & i & "!"
    Next i


    ValidationPartsMatrix = MyErr

End Function



'Public Sub SavingData()
'
'
'    If formAction = "edit" Or formAction = "add" Then
'
'    For i = 1 To flxBSQTY.Rows - 1
'        LYRcon.Execute "UPDATE mrc2retail.partsmatrix_dtls SET partsmatrix_dtls.PMSTDQ = '" & Val(flxBSQTY.TextMatrix(i, 5)) & "', partsmatrix_dtls.PMLACKING = '" & Val(flxBSQTY.TextMatrix(i, 6)) & "', partsmatrix_dtls.PMEXCESS = '" & Val(flxBSQTY.TextMatrix(i, 7)) & "', partsmatrix_dtls.PMSO = '" & flxBSQTY.TextMatrix(i, 8) & "' where partsmatrix_dtls.SPAREPARTID = '" & flxBSQTY.TextMatrix(i, 9) & "' and partsmatrix_dtls.PMATRIXID = '" & matrixID & "'"
'
'        Next i
'    End If
'
'End Sub


'===================BUTTON VISIBILITY===================
Private Sub SetButtonVisibility(DataEntry As Boolean)
'cmdTrans Legend
'0 ->ADD
'1 ->EDIT
'2 ->VIEW SUMM
'3 ->EXIT
'4 ->SAVE
'5 ->CANCEL


    'hide all button
    For i = 0 To 3
        cmdTrans(i).Visible = False
        cmdLoadData.Visible = False
    Next i

    If DataEntry Then
        cmdTrans(4).Visible = True
        cmdTrans(5).Visible = True
    Else
    For i = 0 To 3
        cmdTrans(i).Visible = True
        cmdTrans(4).Visible = False
        cmdTrans(5).Visible = False
        cmdLoadData.Visible = True
    Next i

    End If
End Sub


Public Sub flxChanger()
    With flxBSQTY
        .Cols = 16
        .FormatString = " |<PARTNUMBER |<DESCRIPTION |^PARTSGROUP |^MC MODEL  |^QOH |^AVQ 3MSQ |^NBS |^STDQ |^LACKING |^EXCESS |^ORDFREQ || |^UNIT COST |^UNIT PRICE "
    
        .ColWidth(0) = 300
        .ColWidth(1) = (.Width - .ColWidth(0)) * 0.12   'PARTS NUMBER
        .ColWidth(2) = (.Width - .ColWidth(0)) * 0.25  'DESCRIPTION
        .ColWidth(3) = (.Width - .ColWidth(0)) * 0.1 'PARTSGROUP
        .ColWidth(4) = (.Width - .ColWidth(0)) * 0.1 'MC MODEL
        .ColWidth(5) = (.Width - .ColWidth(0)) * 0.06 'QUANTITY ON HAND
        .ColWidth(6) = (.Width - .ColWidth(0)) * 0.06 '6MONTHS SALE QUANTITY
        .ColWidth(7) = (.Width - .ColWidth(0)) * 0.06 'NEW BRANCH STOCKING
        .ColWidth(8) = (.Width - .ColWidth(0)) * 0.06 'STANDARD QUANTITY
        .ColWidth(9) = (.Width - .ColWidth(0)) * 0.06 'LACKING
        .ColWidth(10) = (.Width - .ColWidth(0)) * 0.06 'EXCESS
        .ColWidth(11) = (.Width - .ColWidth(0)) * 0.06 'ORD FREQ
        .ColWidth(12) = 0 '(.Width - .ColWidth(0)) * 0.03  'SPAREPARTID
        .ColWidth(13) = 0 '(.Width - .ColWidth(0)) * 0.03  ' PG
        .ColWidth(14) = (.Width - .ColWidth(0)) * 0.06  ' UNIT COST
        .ColWidth(15) = (.Width - .ColWidth(0)) * 0.06  ' UNIT PRICE
    End With
End Sub


'Private Sub btnAltD_Click()
'    Call mnuDelete_Click
'End Sub
'
'Private Sub btnAltA_Click()
'    Call mnuAdd_Click
'End Sub


Private Sub btnFind_Click()

    If txtFind.Text = "" Then Exit Sub
    With flxBSQTY
        For i = fRow To .Rows - 1
            For j = fCol To 10
                If InStr(.TextMatrix(i, j), txtFind.Text) > 0 Then
                    .Col = 0: .Row = i
                    .ColSel = 10:  .RowSel = i
                    .TopRow = i
                    fRow = IIf(j = 10, i + 1, i)
                    fCol = IIf(j = 10, 1, j + 1)
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

Private Sub cboOrderFrequency_Change()

    If StatSrch Then
        StatSrch = False
        Exit Sub
    End If
    With cboOrderFrequency
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
 
End Sub

Private Sub cboOrderFrequency_Click()
    Call cboOrderFrequency_Change
    flxBSQTY.TextMatrix(ActRow, 11) = cboOrderFrequency.Text
End Sub

Private Sub cboOrderFrequency_GotFocus()
Dim Tmp
    Const CB_SHOWDROPDOWN = &H14F
    Tmp = SendMessage(cboOrderFrequency.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub

Private Sub cboOrderFrequency_KeyDown(KeyCode As Integer, Shift As Integer)
 
    If KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then
        StatSrch = True
    End If
 
End Sub

Private Sub cboOrderFrequency_KeyPress(KeyAscii As Integer)
On Error GoTo er

    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(cboOrderFrequency, KeyAscii, ValidChars)
    
    If KeyAscii = 13 Then
        flxBSQTY.SetFocus
         
        If flxBSQTY.TextMatrix(flxBSQTY.Row, 1) <> "" Then
            cmbstr = cboOrderFrequency.Text
        
            For i = 0 To cboOrderFrequency.ListCount - 1
                If UCase$(cboOrderFrequency.List(i)) = UCase$(cmbstr) Then
                    cboOrderFrequency.ListIndex = i
                    blnFound = True
                    Exit For
                End If
            Next i
    
            If Not blnFound Then
                If cboOrderFrequency.Text = "" Then
                    flxBSQTY.SetFocus
                    flxBSQTY.Col = 11
                    cboOrderFrequency.Visible = False
                    flxBSQTY.TextMatrix(ActRow, 11) = ""
                    Exit Sub
                End If
                MsgBox "Item was not found in the list.", _
                vbExclamation, _
                "Warning"
                flxBSQTY.SetFocus
                flxBSQTY.Col = 11
                cboOrderFrequency.Visible = False
                flxBSQTY.TextMatrix(ActRow, 11) = ""
                Exit Sub
            End If
            
            flxBSQTY.TextMatrix(ActRow, 11) = cboOrderFrequency.Text
            blnFound = False
            Exit Sub
        End If
    End If

Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->cboOrderFrequency-KeyPress()")
End Sub

Private Sub cboOrderFrequency_LostFocus()
If flxBSQTY.TextMatrix(flxBSQTY.Row, 11) <> "" Then
    flxBSQTY.SetFocus
    If ActRow = flxBSQTY.Rows - 1 Then
        cboOrderFrequency.Visible = False
        Exit Sub
    End If
    
    If Val(flxBSQTY.TextMatrix(ActRow + 1, 7)) = 0 Then
        flxBSQTY.Row = ActRow + 1
        flxBSQTY.Col = 7
    Else
        flxBSQTY.Row = ActRow
        flxBSQTY.Col = 11
    End If
End If
cboOrderFrequency.Visible = False
End Sub


Private Sub cmbBrand_Change()
On Error GoTo er
    If StatSrch Then
            StatSrch = False
            Exit Sub
    End If

    With cmbBrand
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
    Call modObjTools.ErrorMessage(Err.Description, "Error->cmbBrand-Change()")
End Sub



Private Sub cmbBrand_Click()
    If flxBSQTY.TextMatrix(1, 1) <> "" Then
        Ans = MsgBox("Do you want to change the MC Brand?", vbQuestion + vbYesNo, "Validation")
        If Ans = vbNo Then Exit Sub
        Call CboModelDisplay
        flxBSQTY.Rows = 1
        flxBSQTY.Rows = 2
        cmbModel.Text = ""
        txtTTPN.Text = ""
        txtTTSTDQ.Text = ""
        txtPG.Text = ""
        txtExcess.Text = "0.0%"
        txtLacking.Text = "0.0%"
        txtEmpty.Text = "0.0%"
        txtpmID.Text = ""
        txtTTSTDQ.Text = ""
        lblPartsGroup.Caption = "Parts Group"
    Else
        Call CboModelDisplay
    End If
End Sub


Private Sub cmbBrand_GotFocus()
Dim Tmp
    Const CB_SHOWDROPDOWN = &H14F
    Tmp = SendMessage(cmbBrand.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub

Private Sub cmbBrand_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then StatSrch = True
End Sub

Private Sub cmbBrand_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(cmbBrand, KeyAscii, ValidChars)
    If KeyAscii = 13 Then
        Call cmbBrand_Click
        cmbModel.SetFocus
    End If
End Sub

Private Sub cmbModel_Change()
On Error GoTo er

    If StatSrch Then
        StatSrch = False
        Exit Sub
    End If
    
    With cmbModel
        strPartial = .Text
        i = SendMessage(.hwnd, CB_FINDSTRING, -1, ByVal strPartial)  'If match found, append unmatched characters
        If i <> -1 Then
            strTotal = .List(i)
            
            j = Len(strTotal) - Len(strPartial)
            
            matrixID = arrMODEL(i, 0)
           
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
    Call modObjTools.ErrorMessage(Err.Description, "Error->cmbModel()")
End Sub

Private Sub cmbModel_Click()
    
    DoEvents
    
    matrixID = arrMODEL(cmbModel.ListIndex, 0)
    txtpmID.Text = matrixID
    Call ShowRecord(matrixID, "TRUE")

    flxBSQTY.SetFocus
    flxBSQTY.Col = 8

End Sub

Private Sub cmbModel_GotFocus()
Dim Tmp
    Const CB_SHOWDROPDOWN = &H14F
    Tmp = SendMessage(cmbModel.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub

Private Sub cmbModel_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then StatSrch = True
End Sub

Private Sub cmbModel_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(cmbModel, KeyAscii, ValidChars)
    If KeyAscii = 13 Then
        Call cmbModel_Click
        flxBSQTY.SetFocus
        flxBSQTY.Col = 8
    End If
End Sub


Private Sub cmdCompute_Click()
Dim cTtlAmount As Double

If formAction = "add" Or formAction = "edit" Then
    If txtpmID.Text <> "" Then
        Ans = MsgBox("Do you want to Auto compute STDQ?", vbQuestion + vbYesNo, "STDQ")
        If Ans = vbNo Then Exit Sub
        
        cTtlAmount = 0
        For i = 1 To flxBSQTY.Rows - 1
            'NBS > AVQ3MSQ = NBS
            If Val(flxBSQTY.TextMatrix(i, 7)) >= Val(flxBSQTY.TextMatrix(i, 6)) Then
                flxBSQTY.TextMatrix(i, 8) = Val(flxBSQTY.TextMatrix(i, 7))
            'NBS < AVQ3MSQ = AVQ3MSQ
            ElseIf Val(flxBSQTY.TextMatrix(i, 7)) <= Val(flxBSQTY.TextMatrix(i, 6)) Then
                flxBSQTY.TextMatrix(i, 8) = Val(flxBSQTY.TextMatrix(i, 6))
            End If
            
            'QOH > STDQ = EXCESS
            If Val(flxBSQTY.TextMatrix(i, 5)) >= Val(flxBSQTY.TextMatrix(i, 8)) Then
                flxLack = Val(flxBSQTY.TextMatrix(i, 5)) - Val(flxBSQTY.TextMatrix(i, 8))
                flxBSQTY.TextMatrix(i, 10) = flxLack
                flxBSQTY.TextMatrix(i, 9) = 0
            Else
            'QOH < STDQ = LACKING
                flxLack = Val(flxBSQTY.TextMatrix(i, 8)) - Val(flxBSQTY.TextMatrix(i, 5))
                flxBSQTY.TextMatrix(i, 10) = 0
                flxBSQTY.TextMatrix(i, 9) = flxLack
            End If
            
            flxBSQTY.TextMatrix(i, 15) = Val(flxBSQTY.TextMatrix(i, 8)) * Val(flxBSQTY.TextMatrix(i, 14))
            cTtlAmount = cTtlAmount + Val(flxBSQTY.TextMatrix(i, 15))
        Next i
        Call DataPercentage
        txtTOTALAMOUNT.Text = Format(cTtlAmount, "###,##0.00")
    Else
        MsgBox "No Parts Matrix Selected", vbInformation + vbOKOnly, "Information"
    End If
End If
End Sub

Private Sub cmdLoadData_Click()
On Error GoTo LocalErr
    Call ConnectMain
    If modConnection.TestConnection() = False Then Exit Sub
    
    Ans = MsgBox("Do you want to Update Parts Matrix?", vbQuestion + vbYesNo, "Update Parts Matrix")
    If Ans = vbNo Then
        Exit Sub
    Else
        Me.Enabled = False
        tabProcessing.Visible = True
        
        Call modMotoR.sp_insertNewPMpgroup(tmpLDRs)
        If tmpLDRs.State <> 1 Then
            MsgBox "Unable to Update Parts Matrix", vbInformation + vbOKOnly, "Update Parts Matrix"
            Me.Enabled = True
            tabProcessing.Visible = False
            tmpLDRs.Close
            Exit Sub
        Else
            Screen.MousePointer = vbDefault
            tmpLDRs.Close
        End If
        
        Call modMotoR.sp_UpdatePartsmatrix(mainRs)
        Screen.MousePointer = vbDefault
        If mainRs.State <> 1 Then
            MsgBox "Unable to Update Parts Matrix", vbInformation + vbOKOnly, "Update Parts Matrix"
            Me.Enabled = True
            tabProcessing.Visible = False
            mainRs.Close
            Exit Sub
        Else
            Screen.MousePointer = vbDefault
            mainRs.Close
        End If
        
        Call modMotoR.sp_UpdatePartsmatrix_dtls(tmpRs)
        If tmpRs.State <> 1 Then
            MsgBox "Unable to Update Parts Matrix", vbInformation + vbOKOnly, "Update Parts Matrix"
            Me.Enabled = True
            tabProcessing.Visible = False
            tmpRs.Close
            Exit Sub
        Else
            Screen.MousePointer = vbDefault
            tmpRs.Close
        End If
    End If
    Me.Enabled = True
    tabProcessing.Visible = False
    MsgBox "Parts Matrix was successfully updated", vbInformation + vbOKOnly, "Update Parts Matrix"
    Exit Sub
    
LocalErr:
    MsgBox Err.Description & vbCrLf & Err.Number & " " & Err.Description & vbCrLf & vbCrLf & "Please inform M.I.S. Department", vbCritical + vbOKOnly, "Error"
    
End Sub

Private Sub cmdTrans_Click(Index As Integer)
On Error GoTo er

    If Index = 0 Then 'add
    
        formAction = "add"
        AddNew = True
        Call clearform
        txtUpdated.Text = UserName
        Call CboDisplay
        
    ElseIf Index = 1 Then 'edit
    
        If txtpmID = "" Then
            MsgBox "No available PARTS MATRIX to edit", vbOKOnly + vbExclamation, "Edit"
        
        Else
            formAction = "edit"
            AddNew = False
            Call SetButtonVisibility(True)
            Call LockData(False)
            flxBSQTY.SetFocus
            flxBSQTY.Col = 8
        End If
        
    ElseIf Index = 2 Then 'view summary
        frmSummaryRepl.mySrc = "Branch Quantity"
        frmSummaryRepl.Show
        Unload Me
    ElseIf Index = 3 Then 'exit
        formAction = "EXIT"
        ProperExit = True
        Call clearform
        Unload Me
        'MotoRep.imgBG.Visible = True
        MotoRep.toolBarButton.Visible = True
    ElseIf Index = 4 Then 'save
        myMsg = ValidationPartsMatrix
        If myMsg <> "" Then
            MsgBox myMsg, vbCritical + vbOKOnly, "Validation Error"
            Exit Sub
        End If
        Ans = MsgBox("Do you want to save?", vbQuestion + vbYesNo, "Save")
        If Ans = vbNo Then Exit Sub
'            For i = 1 To flxBSQTY.Rows - 1
'                If flxBSQTY.TextMatrix(i, 1) <> "" Then
'                    txtTTPN.Text = i
'                    intQOH = intQOH + Val(flxBSQTY.TextMatrix(i, 3))
'                    strTtlQOH = intQOH
'                End If
'
'            Next i
        If AddNew Then
            Call CreateTempTable
            DoEvents
            Call modMotoR.sp_brnchQty_create(txtpmID.Text, mainRs)
        Else
            Call CreateTempTable
            DoEvents
            Call modMotoR.sp_brnchQty_update(txtpmID.Text, mainRs)
            Call ShowRecord(matrixID, strBoolean) 'to Refresh
        End If
            Call SetButtonVisibility(False)
            Call LockData(False)
            formAction = ""
    ElseIf Index = 5 Then 'cancel
        Ans = MsgBox("Do you want to cancel?", vbQuestion + vbYesNo)
        If Ans = vbNo Then Exit Sub
        Call SetButtonVisibility(False)
        Call LockData(False)
        formAction = ""
    End If
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->cmdTrans-Click()")
End Sub

Private Sub flxBSQTY_DblClick()
    If formAction = "add" Or formAction = "edit" Then
        If flxBSQTY.TextMatrix(flxBSQTY.Row, 11) = "SO" Then
            flxBSQTY.TextMatrix(ActRow, 11) = ""
        End If
    End If
End Sub

Private Sub flxBSQTY_Scroll()
txtSTDQ.Visible = False
cboOrderFrequency.Visible = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo er
    If ProperExit Then
        If mainRs.State = adStateOpen Then mainRs.Close
        Call cmdTrans_Click(3)
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
    Call flxChanger
End Sub

Private Sub Form_Load()
    Call LockData(False)
    
    txtFind.Text = ""
    txtExcess.Text = "0.0%"
    txtLacking.Text = "0.0%"
    txtEmpty.Text = "0.0%"
    
    '===================BUTTON POSITION===================
    If cmdTrans(0).Visible = False Then
        'Align Save Button
        cmdTrans(4).Left = cmdTrans(0).Left
        cmdTrans(4).Top = cmdTrans(0).Top
        'Align Cancel Button
        cmdTrans(5).Left = cmdTrans(1).Left
        cmdTrans(5).Top = cmdTrans(1).Top
    End If
    
End Sub

Public Sub CboDisplay()
On Error GoTo er

    If formAction = "add" Then
        'Get Data from database
        DoEvents
        Call modMotoR.get_Brand(tmpRs)
        If tmpRs.State = 1 Then
    '***************** CODE UPDATED: 04/18/17 *********************
            If tmpRs.RecordCount > 0 Then
                Call modObjTools.fillArray(modMotoR.AllArrBrand, tmpRs)
            Else
                MsgBox "No Record Found", vbInformation + vbOKOnly, "Validation"
                Exit Sub
            End If
        End If
        If tmpRs.State <> 0 Then tmpRs.Close
        
        'Load Data to combo box
        Call SetCboListData
        
        Call LockData(True)
        Call SetButtonVisibility(True)
     '************ END CODE UPDATED: 04/18/17 *********************
    End If

Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->CboDisplay()")
End Sub


Public Sub CboModelDisplay()
On Error GoTo er
    If formAction = "add" Then
        DoEvents
        Call modMotoR.get_Model(cmbBrand.Text, tmpRs)
        If tmpRs.State = 1 Then
            Call modObjTools.addToListObjectWithArrayGivenField(cmbModel, tmpRs, arrMODEL, "PMATRIXID", "PMMODEL")
        End If
        If tmpRs.State <> 0 Then tmpRs.Close
    End If

Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->CboModelDisplay()")
End Sub


Sub SetCboListData()
On Error GoTo er

    If cmbBrand.Text = "" Then
        DoEvents
        Call modObjTools.addToListObjectUsingArrayWithCols(cmbBrand, arrBRAND, modMotoR.AllArrBrand, 0, 0)
    End If

Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->SetCboListData()")
End Sub


Sub clearform()

    flxBSQTY.Rows = 1
    flxBSQTY.Rows = 2
   
    cmbBrand.Text = ""
    cmbModel.Text = ""
    txtTTPN.Text = ""
    txtTTSTDQ.Text = ""
    txtPG.Text = ""
    txtExcess.Text = "0.0%"
    txtLacking.Text = "0.0%"
    txtEmpty.Text = "0.0%"
    txtpmID.Text = ""
    txtFind.Text = ""
End Sub

Private Sub flxBSQTY_Click()
On Error GoTo er

    ActRow = flxBSQTY.Row
    ActCol = flxBSQTY.Col
    
    If formAction = "add" Or formAction = "edit" Then 'clickable all
    
    If ActCol = 0 Or ActRow = 0 Then Exit Sub
        Select Case ActCol
        
        Case 8
            If flxBSQTY.TextMatrix(flxBSQTY.Row, 1) <> "" Then
                With txtSTDQ
                    If flxBSQTY.TextMatrix(ActRow, 8) <> "" Then .Tag = flxBSQTY.TextMatrix(ActRow, 8)
                    .Top = Frame3.Top + flxBSQTY.CellTop + flxBSQTY.Top
                    .Left = Frame3.Left + flxBSQTY.CellLeft + flxBSQTY.Left
                    .Height = flxBSQTY.CellHeight
                    .Width = flxBSQTY.CellWidth
                    .Text = flxBSQTY.Text
                    .Visible = True
                    .SelStart = 0
                    .SelLength = Len(.Text)
                    .ZOrder
                    .SetFocus
                End With
            Else
                Exit Sub
            End If
            
        Case 11
            If flxBSQTY.TextMatrix(flxBSQTY.Row, 1) <> "" Then
                With cboOrderFrequency
                    If flxBSQTY.TextMatrix(ActRow, 11) = "" Then .Tag = flxBSQTY.TextMatrix(ActRow, 11)
                    .Top = Frame3.Top + flxBSQTY.CellTop + flxBSQTY.Top
                    .Left = Frame3.Left + flxBSQTY.CellLeft + flxBSQTY.Left
                    .Width = flxBSQTY.CellWidth
                    .Text = flxBSQTY.Text
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
    Call modObjTools.ErrorMessage(Err.Description, "Error->flxBSQTY-click()")
End Sub


Private Sub flxBSQTY_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call flxBSQTY_Click
End Sub

'****************** CODE UPDATED: 04/08/17 ************************
Public Sub ComputeTotal()
Dim cTLacking As Integer, cTExcess As Integer, cTSTDQ As Integer, cTEmpty As Integer, cTQtyOnHand As Integer
Dim cTPN As Integer, cTAmount As Double

    cTPN = 0
    intComputeTtlPN = 0
    cTQtyOnHand = 0
    intComputeTtlQtyOnHand = 0
    cTLacking = 0
    intComputeTtlLacking = 0
    cTExcess = 0
    intComputeTtlExcess = 0
    cTSTDQ = 0
    intComputeTtlSTDQ = 0
    cTEmpty = 0
    intComputeTtlEmpty = 0
    cTAmount = 0
    intComputeTtlAmount = 0
    
    For i = 1 To flxBSQTY.Rows - 1
        If Val(flxBSQTY.TextMatrix(i, 5)) = 0 Then
            cTEmpty = cTEmpty + 1
        End If
        If flxBSQTY.TextMatrix(i, 1) <> "" Then
            cTPN = cTPN + 1
        End If
        cTQtyOnHand = cTQtyOnHand + Val(Replace(flxBSQTY.TextMatrix(i, 5), ",", ""))
        cTSTDQ = cTSTDQ + Val(Replace(flxBSQTY.TextMatrix(i, 8), ",", ""))
        cTLacking = cTLacking + Val(Replace(flxBSQTY.TextMatrix(i, 9), ",", ""))
        cTExcess = cTExcess + Val(Replace(flxBSQTY.TextMatrix(i, 10), ",", ""))
        cTAmount = cTAmount + Val(Replace(flxBSQTY.TextMatrix(i, 15), ",", ""))
    Next i
    intComputeTtlQtyOnHand = cTQtyOnHand
    intComputeTtlSTDQ = cTSTDQ
    intComputeTtlLacking = cTLacking
    intComputeTtlExcess = cTExcess
    intComputeTtlEmpty = cTEmpty
    intComputeTtlPN = cTPN
    intComputeTtlAmount = cTAmount
End Sub
'************** END CODE UPDATED: 04/08/17 ************************

Public Sub ShowRecord(MyID As Integer, MyBoolean As String)
On Error GoTo er
'************** CODE UPDATED: 04/08/17 ******************
Dim intComputeLacking As Integer, intComputeExcess As Integer
Dim intSTDQty As Integer, strORDERFREQ As String
'********** END CODE UPDATED: 04/08/17 ******************
    On Error GoTo er
   
    txtpmID.Text = MyID 'used for Edit verification
    matrixID = MyID 'saving data
    strBoolean = MyBoolean
    Screen.MousePointer = vbHourglass
    
    DoEvents
    '************** CODE UPDATED: 04/18/17 ******************
    Call modMotoR.sp_brnchQty_getDetails(matrixID, MyBoolean, mainRs)
    '********** END CODE UPDATED: 04/18/17 ******************
    
    If mainRs.State <> 1 Then Exit Sub
    
    If mainRs.RecordCount <= 0 Then
        ' & vbCrLf & " If you want to Edit Please go to" & vbCrLf & " View Summary -> Show Form -> Edit"
        MsgBox "There are no Available Parts Group has assigned" & vbCrLf & " If you want to Edit Please go to" & vbCrLf & " View Summary -> Show Form -> Edit", vbInformation + vbOKOnly, "Show Details"
        cmbModel.Text = ""
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
        With mainRs
            If .RecordCount > 0 Then
                .MoveFirst
                flxBSQTY.Rows = .RecordCount + 1
                Call resetVar
                strORDERFREQ = ""
                For i = 1 To flxBSQTY.Rows - 1
                intTotalSTDQ = intTotalSTDQ + IIf(IsNull(!PMSTDQ), 0, !PMSTDQ)
        '************** CODE UPDATED: 04/08/17 ******************
                intComputeLacking = 0
                intComputeExcess = 0
                intSTDQty = 0
'
'                If myBoolean = "TRUE" Then
'                    'get STDQ
'                    If IIf(IsNull(!PMNBS), 0, !PMNBS) >= !PM3MSQty Then
'                        intSTDQty = IIf(IsNull(!PMNBS), 0, !PMNBS)
'                    ElseIf IIf(IsNull(!PMNBS), 0, !PMNBS) <= !PM3MSQty Then
'                        intSTDQty = !PM3MSQty
'                    End If
'
'                    'get LACKING and EXCESS
'                    If !QtyOnHand <= IIf(IsNull(intSTDQty), 0, intSTDQty) Then
'                        intComputeLacking = IIf(IsNull(intSTDQty), 0, intSTDQty) - !QtyOnHand
'                    ElseIf !QtyOnHand >= IIf(IsNull(intSTDQty), 0, intSTDQty) Then
'                        intComputeExcess = !QtyOnHand - IIf(IsNull(intSTDQty), 0, intSTDQty)
'                    End If
'                Else
                    'get STDQ
'                    intSTDQty = IIf(IsNull(!PMSTDQ), 0, !PMSTDQ)
                    
                    'get LACKING and EXCESS
                    If !QtyOnHand <= IIf(IsNull(!PMSTDQ), 0, !PMSTDQ) Then
                        intComputeLacking = IIf(IsNull(!PMSTDQ), 0, !PMSTDQ) - !QtyOnHand
                    ElseIf !QtyOnHand >= IIf(IsNull(!PMSTDQ), 0, !PMSTDQ) Then
                        intComputeExcess = !QtyOnHand - IIf(IsNull(!PMSTDQ), 0, !PMSTDQ)
                    End If
'                End If
                
                    If IIf(IsNull(!PMORDERFREQ), "", !PMORDERFREQ) = "SO" Then
                        strORDERFREQ = "SPECIAL ORDER"
                    Else
                        strORDERFREQ = IIf(IsNull(!PMORDERFREQ), "", !PMORDERFREQ)
                    End If
        '********** END CODE UPDATED: 04/08/17 ******************
                NewPrtsGroup = !partsGroup
                
                If NewPrtsGroup <> OldPrtsGroup Then
                    intPG = intPG + 1
                    flxPG.Rows = intPG + 1
                    flxPG.TextMatrix(intPG, 1) = NewPrtsGroup
                End If
                
                 flxBSQTY.TextMatrix(i, 1) = !PartNumber
                 flxBSQTY.TextMatrix(i, 2) = !Description
                 flxBSQTY.TextMatrix(i, 3) = !partsGroup
                 flxBSQTY.TextMatrix(i, 4) = !PMMODEL
                 flxBSQTY.TextMatrix(i, 5) = !QtyOnHand
                 flxBSQTY.TextMatrix(i, 6) = !PM3MSQty
                 flxBSQTY.TextMatrix(i, 7) = IIf(IsNull(!PMNBS), 0, !PMNBS)
                 flxBSQTY.TextMatrix(i, 8) = IIf(IsNull(!PMSTDQ), 0, !PMSTDQ)
            '************** CODE UPDATED: 04/08/17 ******************
                 flxBSQTY.TextMatrix(i, 9) = Format$(intComputeLacking, 0)
                 flxBSQTY.TextMatrix(i, 10) = Format$(intComputeExcess, 0)
            '********** END CODE UPDATED: 04/08/17 ******************
                 flxBSQTY.TextMatrix(i, 11) = IIf(IsNull(strORDERFREQ), "", strORDERFREQ)
                 flxBSQTY.TextMatrix(i, 12) = !SparepartID
                 flxBSQTY.TextMatrix(i, 13) = !partsGroup
                 flxBSQTY.TextMatrix(i, 14) = IIf(IsNull(!Cost), "#.##00", !Cost)
                 flxBSQTY.TextMatrix(i, 15) = Val(flxBSQTY.TextMatrix(i, 8)) * Val(flxBSQTY.TextMatrix(i, 14))
                .MoveNext
                OldPrtsGroup = NewPrtsGroup
                Next i
                
'************** CODE UPDATED: 04/18/17 ******************
'            ElseIf !NULLs = 0 Then
'                For i = 1 To flxBSQTY.Rows - 1
'            '************** CODE UPDATED: 04/08/17 ******************
'                intComputeLacking = 0
'                intComputeExcess = 0
'                If !QtyOnHand <= !PMSTDQ Then
'                    intComputeLacking = !PMSTDQ - !QtyOnHand
'                ElseIf !QtyOnHand >= !PMSTDQ Then
'                    intComputeExcess = !QtyOnHand - !PMSTDQ
'                End If
'            '********** END CODE UPDATED: 04/08/17 ******************
'                NewPrtsGroup = !partsGroup
'                If NewPrtsGroup <> OldPrtsGroup Then
'                    intPG = intPG + 1
'                    flxPG.Rows = intPG + 1
'                    flxPG.TextMatrix(intPG, 1) = NewPrtsGroup
'                End If
'
'                 flxBSQTY.TextMatrix(i, 1) = !PartNumber
'                 flxBSQTY.TextMatrix(i, 2) = !Description
'                 flxBSQTY.TextMatrix(i, 3) = !PMBRAND
'                 flxBSQTY.TextMatrix(i, 4) = !PMMODEL
'                 flxBSQTY.TextMatrix(i, 5) = !QtyOnHand
'                 flxBSQTY.TextMatrix(i, 6) = !PM3MSQty
'                 flxBSQTY.TextMatrix(i, 7) = IIf(IsNull(!PMSTDQ), 0, !PMSTDQ)
'                 flxBSQTY.TextMatrix(i, 8) = Format$(intComputeLacking, 0)
'                 flxBSQTY.TextMatrix(i, 9) = Format$(intComputeExcess, 0)
'                 flxBSQTY.TextMatrix(i, 10) = ""
'                 flxBSQTY.TextMatrix(i, 11) = !SparepartID
'                 flxBSQTY.TextMatrix(i, 12) = !partsGroup
'                .MoveNext
'                OldPrtsGroup = NewPrtsGroup
'                Next i
'             End If
'********** END CODE UPDATED: 04/18/17 ******************
                For i = 1 To flxPG.Rows - 1
                    If i > 0 Then
                        strPG = strPG & flxPG.TextMatrix(i, 1) & "/"
                        intComputePG = intComputePG + 1
                    End If
                Next
                strPG = Left(strPG, Len(strPG) - 1) ' To remove the last slash
                lblPartsGroup.Caption = strPG & " " & "ONLY"
            
            
        End If
        End With
        
    If mainRs.State <> 0 Then
        Screen.MousePointer = vbDefault
        mainRs.Close
    End If
    
    ' Showrecord of BRAND,MODEL TOTAL PN and etc.
    DoEvents
    Call modMotoR.sp_brnchQty_getDetails_total(matrixID, MyBoolean, tmpRs)

    If tmpRs.State <> 1 Then Exit Sub

    If tmpRs.RecordCount <= 0 Then
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

        With tmpRs
            If .RecordCount > 0 Then
                .MoveFirst
        '****************** CODE UPDATED: 04/08/17 ************************
                Call ComputeTotal
                'Percentage Calculation
                Call calculatePercentageShowRecord
        
                cmbBrand.Text = !PMBRAND
                cmbModel.Text = !PMMODEL
                txtPG.Text = Format$(intComputePG, 0)
                txtTTPN.Text = Format$(intComputeTtlPN, 0)
                txtTTSTDQ.Text = Format$(intComputeTtlSTDQ, 0)
                txtTOTALAMOUNT.Text = Format(intComputeTtlAmount, "###,##0.00")
                txtEmpty.Text = Format$(strTtlEmpty, "0.0%")
                txtExcess.Text = Format$(strTtlExcess, "0.0%")
                txtLacking.Text = Format$(strTtlLacking, "0.0%")
                txtUpdated.Text = IIf(IsNull(!UPDATEDBY), UserName, !UPDATEDBY)
            End If
        '************** END CODE UPDATED: 04/08/17 ************************
        End With
     If tmpRs.State <> 0 Then
        Screen.MousePointer = vbDefault
        tmpRs.Close
    End If
    
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->branchQtyShowRecord()")
        
End Sub

'****************** CODE UPDATED: 04/08/17 ************************
Public Sub calculatePercentageShowRecord()
    strTtlQOH = intComputeTtlQtyOnHand
    'Excess Percentage
    intExcess = Format$(intComputeTtlExcess, 0)
    If intExcess <> 0 Then
        If intExcess <= strTtlQOH Then
            strTtlExcess = intExcess / strTtlQOH * 100 & "%"
        ElseIf intExcess >= strTtlQOH Then
            strTtlExcess = strTtlQOH / intExcess * 100 & "%"
        End If
     Else
        strTtlExcess = "0.0%"
    End If
    
    'Lacking Percentage
    intLacking = Format$(intComputeTtlLacking, 0)
    If intLacking <> 0 Then
        If Val(intLacking) <= intComputeTtlSTDQ Then
            strTtlLacking = intLacking / Format$(intComputeTtlSTDQ, 0) * 100 & "%"
        ElseIf Val(intLacking) >= intComputeTtlSTDQ Then
            strTtlLacking = Format$(intComputeTtlSTDQ, 0) / intLacking * 100 & "%"
        End If
    Else
        strTtlLacking = "0.0%"
    End If
    
    'Empty Percentage
    intEmpty = Format$(intComputeTtlEmpty, 0)
    If intEmpty <> 0 Then
        If intEmpty <= intComputeTtlPN Then
            strTtlEmpty = intEmpty / intComputeTtlPN * 100 & "%"
        ElseIf intEmpty >= intComputeTtlPN Then
            strTtlEmpty = intComputeTtlPN / intEmpty * 100 & "%"
        End If
    Else
        strTtlEmpty = "0.0%"
    End If
    
End Sub
'************** END CODE UPDATED: 04/08/17 ************************

Public Sub resetVar()
    strPG = ""
    intPG = 0
    flxPG.Rows = 1
    flxPG.Rows = 2
    NewPrtsGroup = ""
    OldPrtsGroup = ""
    intTotalSTDQ = 0
    intComputePG = 0
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)

    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtFind, KeyAscii, ValidChars)
    
    If KeyAscii = 13 Then
        Call btnFind_Click
    End If
    
End Sub

'Private Sub flxBSQTY_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
'
'    If Button = vbRightButton Then
'        PopupMenu mnuStatus, vbPopupMenuRightButton
'    End If
'
'End Sub


'Private Sub mnuAdd_Click()
'
'    If flxBSQTY.TextMatrix(actrow, 1) <> flxBSQTY.TextMatrix(flxBSQTY.Rows - 1, 1) Then
'        flxBSQTY.SetFocus
'        Exit Sub
'    End If
'
'    If flxBSQTY.TextMatrix(flxBSQTY.Rows - 1, 5) = "" Or flxBSQTY.TextMatrix(flxBSQTY.Rows - 1, 9) = "" Then
'        MsgBox "Incomplete entry on last row!", vbCritical + vbOKOnly, "Add Row failed"
'        flxBSQTY.SetFocus
'        'PrtAdd = False
'        Exit Sub
'    End If
'
'    flxBSQTY.AddItem ""
'    flxBSQTY.SetFocus
'    flxBSQTY.Row = flxBSQTY.Rows - 1
'    flxBSQTY.Col = 1
'    txtTTPN.Text = flxBSQTY.Rows - 1
'
'End Sub


'Private Sub mnuDelete_Click()
'
'    Ans = MsgBox("Do you want to Delete Row?", vbQuestion + vbYesNo, "Delete")
'    If Ans = vbYes Then
'
'        If flxBSQTY.Rows = 2 Then
'            flxBSQTY.Rows = 1
'            flxBSQTY.Rows = 2
'        Else
'            flxBSQTY.RemoveItem actrow
'            For i = 1 To flxBSQTY.Rows - 1
'                  If flxBSQTY.TextMatrix(i, 1) <> "" Then
'                  txtTTPN.Text = i
'                 End If
'             Next i
'        End If
'    End If
'
'End Sub


Private Sub txtSTDQ_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Tmp = Split(txtSTDQ.Tag, "|")
    
        flxBSQTY.SetFocus
        flxBSQTY.Col = 8
        txtSTDQ.Visible = False
    End If
End Sub

Public Sub DataPercentage()
On Error GoTo er

    intSTDQ = 0
    intLacking = 0
    intExcess = 0
    intEmpty = 0
    intQOH = 0
    strTtlEmpty = 0
    strTtlExcess = 0
    strTtlLacking = 0
    'Getting the total of Standard Quantity
    For LoopSTDQ = 1 To flxBSQTY.Rows - 1
        intSTDQ = intSTDQ + Val(flxBSQTY.TextMatrix(LoopSTDQ, 8))
    Next LoopSTDQ
    txtTTSTDQ.Text = intSTDQ
    
    
    'Getting the Total of Quantity of Hand
    For LoopQOH = 1 To flxBSQTY.Rows - 1
        intQOH = intQOH + Val(flxBSQTY.TextMatrix(LoopQOH, 5))
    Next LoopQOH
    intTotlQOH = intQOH
    
    
    'Getting the Total Percentage of Lacking
    For LoopLACKING = 1 To flxBSQTY.Rows - 1
        intLacking = intLacking + Val(flxBSQTY.TextMatrix(LoopLACKING, 9))
    Next LoopLACKING
    
    strTtlLacking = intLacking / intSTDQ * 100 & "%"
    txtLacking.Text = Format$(strTtlLacking, "0.0%")
    
    
    'Getting the Total Percentage of Empty
    For LoopEmpty = 1 To flxBSQTY.Rows - 1
              
        If flxBSQTY.TextMatrix(LoopEmpty, 5) <> "" Then
            If flxBSQTY.TextMatrix(LoopEmpty, 5) = 0 Then
            intEmpty = intEmpty + 1
            
            Else
            
            GoTo SKIP
            End If
            
        End If
SKIP:
intFlxEmpty = LoopEmpty
    Next LoopEmpty
    
    strTtlEmpty = intEmpty / intFlxEmpty * 100 & "%"
    txtEmpty.Text = Format$(strTtlEmpty, "0.0%")
    
    'Getting the Total Percentage of Excess
    For LoopExcess = 1 To flxBSQTY.Rows - 1
        intExcess = intExcess + Val(flxBSQTY.TextMatrix(LoopExcess, 10))
    Next LoopExcess
     If intExcess <> intTotlQOH Then
    strTtlExcess = intExcess / intTotlQOH * 100 & "%"
    txtExcess.Text = Format$(strTtlExcess, "0.0%")
      Else
       Exit Sub
     End If
     
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->DataPercentage()")
End Sub

Private Sub txtSTDQ_KeyPress(KeyAscii As Integer)
On Error GoTo er

    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtSTDQ, KeyAscii, NumbersOnly)

    If KeyAscii = 13 Then

        If flxBSQTY.TextMatrix(ActRow, 1) <> "" Then
            flxBSQTY.TextMatrix(ActRow, 8) = txtSTDQ
            
            If Val(txtSTDQ.Text) <> 0 Then
                Call cmputeSTDQty
                Call DataPercentage
                flxBSQTY.SetFocus
                flxBSQTY.Row = ActRow
                flxBSQTY.Col = 11
                txtSTDQ.Visible = False
     
            ElseIf Val(txtSTDQ.Text) = 0 Or txtSTDQ.Text = "" Then
                flxBSQTY.SetFocus
                flxBSQTY.Col = 8
                txtSTDQ.Visible = False
            End If
        End If
    End If

Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->txtSTDQ-KeyPress()")
End Sub

Public Sub cmputeSTDQty()
Dim cPTtlAmount As Double

    'QOH > STDQ = EXCESS
    If Val(flxBSQTY.TextMatrix(ActRow, 5)) >= Val(flxBSQTY.TextMatrix(ActRow, 8)) Then
        flxLack = Val(flxBSQTY.TextMatrix(ActRow, 5)) - Val(flxBSQTY.TextMatrix(ActRow, 8))
        flxBSQTY.TextMatrix(ActRow, 10) = flxLack
        flxBSQTY.TextMatrix(ActRow, 9) = 0
    Else
    'QOH < STDQ = LACKING
        flxLack = Val(flxBSQTY.TextMatrix(ActRow, 8)) - Val(flxBSQTY.TextMatrix(ActRow, 5))
        flxBSQTY.TextMatrix(ActRow, 10) = 0
        flxBSQTY.TextMatrix(ActRow, 9) = flxLack
    End If
    
    'displaying Unit Price
    flxBSQTY.TextMatrix(ActRow, 15) = Val(flxBSQTY.TextMatrix(ActRow, 8)) * Val(flxBSQTY.TextMatrix(ActRow, 14))
    
    'getting TOTAL AMOUNT
    cPTtlAmount = 0
    For i = 1 To flxBSQTY.Rows - 1
        cPTtlAmount = cPTtlAmount + Val(flxBSQTY.TextMatrix(i, 15))
    Next i
    
    txtTOTALAMOUNT.Text = Format(cPTtlAmount, "###,##0.00")
End Sub

Sub LockData(DataEntry As Boolean)
    DataEntry = Not DataEntry

    cmbBrand.Locked = DataEntry
    cmbModel.Locked = DataEntry
    txtTTPN.Locked = DataEntry
    txtTTSTDQ.Locked = DataEntry
    txtPG.Locked = DataEntry
    txtExcess.Locked = DataEntry
    txtLacking.Locked = DataEntry
    txtEmpty.Locked = DataEntry
    txtUpdated.Locked = DataEntry
    txtTOTALAMOUNT.Locked = DataEntry
    
    If cmbBrand.Locked Then
        cmbBrand.BackColor = txtHoldGrayBackColor.BackColor
        cmbModel.BackColor = txtHoldGrayBackColor.BackColor
        txtTTPN.BackColor = txtHoldGrayBackColor.BackColor
        txtTTSTDQ.BackColor = txtHoldGrayBackColor.BackColor
        txtPG.BackColor = txtHoldGrayBackColor.BackColor
        txtExcess.BackColor = txtHoldGrayBackColor.BackColor
        txtLacking.BackColor = txtHoldGrayBackColor.BackColor
        txtEmpty.BackColor = txtHoldGrayBackColor.BackColor
        txtUpdated.BackColor = txtHoldGrayBackColor.BackColor
        txtTOTALAMOUNT.BackColor = txtHoldGrayBackColor.BackColor
        
    Else
        cmbBrand.BackColor = txtHoldWhiteBackColor.BackColor
        cmbModel.BackColor = txtHoldWhiteBackColor.BackColor
'        txtTTPN.BackColor = txtHoldWhiteBackColor.BackColor
'        txtTTSTDQ.BackColor = txtHoldWhiteBackColor.BackColor
'        txtPG.BackColor = txtHoldWhiteBackColor.BackColor
'        txtExcess.BackColor = txtHoldWhiteBackColor.BackColor
'        txtLacking.BackColor = txtHoldWhiteBackColor.BackColor
'        txtEmpty.BackColor = txtHoldWhiteBackColor.BackColor
'        txtUpdated.BackColor = txtHoldWhiteBackColor.BackColor
        
    End If
End Sub

Private Sub txtSTDQ_LostFocus()
On Error GoTo er

    If txtSTDQ.Tag = "skiplostfocus" Then
        txtSTDQ.Tag = ""
        Exit Sub
    End If
    
    If flxBSQTY.TextMatrix(ActRow, 8) = "" Or txtSTDQ.Text = "" Then
        MsgBox "Wrong Entry of Standard Quantity!" & vbCrLf & "Please Enter a Correct Standard Quantity.", vbExclamation + vbOKOnly, "Warning"

        flxBSQTY.TextMatrix(ActRow, 8) = 0
        flxBSQTY.SetFocus
        flxBSQTY.Col = 8
        txtSTDQ.Visible = False
        Exit Sub
    End If
    
    flxBSQTY.TextMatrix(ActRow, 8) = txtSTDQ.Text
    txtSTDQ.Visible = False
    
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->txtSTDQ-LostFocus()")
End Sub

