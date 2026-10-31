VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form SPLocAdj 
   Caption         =   "Parts Location Adjustments"
   ClientHeight    =   8580
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11520
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8580
   ScaleWidth      =   11520
   Tag             =   "XYF"
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtNewLoc 
      Height          =   285
      Left            =   5760
      TabIndex        =   34
      Top             =   8760
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtTransLocID 
      Height          =   285
      Left            =   4560
      MaxLength       =   20
      TabIndex        =   33
      Top             =   8760
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSFlexGridLib.MSFlexGrid flxPRN 
      Height          =   975
      Left            =   3120
      TabIndex        =   32
      Top             =   9840
      Visible         =   0   'False
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   1720
      _Version        =   393216
      Cols            =   5
   End
   Begin VB.CommandButton cmdA 
      Caption         =   "&A"
      Height          =   375
      Left            =   6480
      TabIndex        =   31
      Top             =   5760
      Width           =   135
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   0
      Top             =   1800
   End
   Begin VB.TextBox txtPartNo 
      Height          =   285
      Left            =   3240
      TabIndex        =   30
      Top             =   9120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ComboBox cboType 
      Height          =   315
      Left            =   3360
      TabIndex        =   29
      Top             =   8760
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxParts 
      Height          =   1815
      Left            =   480
      TabIndex        =   28
      Top             =   8640
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   3201
      _Version        =   393216
      BackColor       =   16777215
      FixedRows       =   0
      BackColorBkg    =   16777215
      GridColor       =   16777215
      GridColorFixed  =   16777215
      GridColorUnpopulated=   16777215
      ScrollBars      =   2
      SelectionMode   =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Frame Frame2 
      Height          =   1455
      Left            =   480
      TabIndex        =   26
      Tag             =   "XYF"
      Top             =   5685
      Width           =   5895
      Begin VB.TextBox txtReason 
         Height          =   765
         Left            =   240
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Tag             =   "XYF"
         Top             =   480
         Width           =   5535
      End
      Begin VB.Label Label3 
         Caption         =   "REASON FOR ADJUSTMENT:"
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
         Left            =   240
         TabIndex        =   27
         Tag             =   "XYF"
         Top             =   240
         Width           =   3615
      End
   End
   Begin VB.TextBox txtQtyTip 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Height          =   285
      Left            =   360
      TabIndex        =   25
      Text            =   "Qty. Purchased: 0"
      Top             =   8160
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtPrnCopy 
      Height          =   285
      Left            =   6960
      TabIndex        =   24
      Top             =   8760
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Timer tmrFloat 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   840
      Top             =   8760
   End
   Begin VB.Timer tmrDrop 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   840
      Top             =   9240
   End
   Begin VB.Timer tmrFloatRev 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1440
      Top             =   8760
   End
   Begin VB.Timer tmrDropRev 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1440
      Top             =   9240
   End
   Begin MSFlexGridLib.MSFlexGrid flxAdj 
      Height          =   3855
      Left            =   480
      TabIndex        =   2
      Tag             =   "XYF"
      Top             =   1800
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   6800
      _Version        =   393216
      Cols            =   8
      FixedCols       =   0
      RowHeightMin    =   315
      BackColorBkg    =   -2147483633
      ScrollBars      =   2
      FormatString    =   "^    |<PART NUMBER      |<DESCRIPTION          |^TYPE         |^PREV. LOCATION   |^ NEW LOCATION      "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   1000
      Left            =   480
      TabIndex        =   7
      Tag             =   "XYF"
      Top             =   720
      Width           =   10695
      Begin VB.TextBox txtRefNo 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1800
         MaxLength       =   50
         TabIndex        =   0
         Tag             =   "XYF"
         Top             =   538
         Width           =   4935
      End
      Begin VB.TextBox txtAdjDate 
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
         Height          =   315
         Left            =   8520
         TabIndex        =   6
         Tag             =   "XYF"
         Top             =   200
         Width           =   1815
      End
      Begin VB.TextBox txtAdjNo 
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
         Height          =   315
         Left            =   1800
         TabIndex        =   1
         Tag             =   "XYF"
         Top             =   200
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "ADJUSTMENT DATE"
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
         Index           =   4
         Left            =   6720
         TabIndex        =   10
         Tag             =   "XYF"
         Top             =   230
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "CONTROL NO."
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
         Index           =   3
         Left            =   240
         TabIndex        =   9
         Tag             =   "XYF"
         Top             =   230
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "REFERENCE NO."
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
         Index           =   1
         Left            =   240
         TabIndex        =   8
         Tag             =   "XYF"
         Top             =   568
         Width           =   1575
      End
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   661
      Caption         =   "&Add New"
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
      Image           =   "SPLocAdjust.frx":0000
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   2
      Left            =   3000
      TabIndex        =   16
      Top             =   0
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   661
      Caption         =   "Su&mmary"
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
      Image           =   "SPLocAdjust.frx":059A
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   3
      Left            =   4500
      TabIndex        =   17
      Top             =   0
      Width           =   1500
      _ExtentX        =   2646
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
      Image           =   "SPLocAdjust.frx":0B34
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   5
      Left            =   7500
      TabIndex        =   18
      Top             =   0
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   661
      Caption         =   "E&xit"
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
      Image           =   "SPLocAdjust.frx":10CE
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   6
      Left            =   9000
      TabIndex        =   19
      Top             =   0
      Visible         =   0   'False
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   661
      Caption         =   "&Save"
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
      Image           =   "SPLocAdjust.frx":1668
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   7
      Left            =   10500
      TabIndex        =   20
      Top             =   0
      Visible         =   0   'False
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   661
      Caption         =   "&Cancel"
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
      Image           =   "SPLocAdjust.frx":1C02
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   1
      Left            =   1500
      TabIndex        =   21
      Top             =   0
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   661
      Caption         =   "&Edit"
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
      Image           =   "SPLocAdjust.frx":219C
      cBack           =   -2147483633
   End
   Begin VB.TextBox txtPrepBy 
      Height          =   285
      Left            =   7800
      TabIndex        =   4
      Tag             =   "XYF"
      Top             =   6120
      Width           =   3255
   End
   Begin VB.TextBox txtAprvBy 
      Height          =   285
      Left            =   7800
      TabIndex        =   5
      Tag             =   "XYF"
      Top             =   6525
      Width           =   3255
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   4
      Left            =   6000
      TabIndex        =   35
      Top             =   0
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   661
      Caption         =   "&New Location"
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
      Image           =   "SPLocAdjust.frx":2736
      cBack           =   -2147483633
   End
   Begin VB.Label Label2 
      Caption         =   "APPROVED BY:"
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
      Index           =   2
      Left            =   6480
      TabIndex        =   23
      Tag             =   "XYF"
      Top             =   6600
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "PREPARED BY:"
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
      Index           =   1
      Left            =   6480
      TabIndex        =   22
      Tag             =   "XYF"
      Top             =   6120
      Width           =   1335
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   120
      Top             =   4320
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   11640
      m_MDI_Height    =   9090
   End
   Begin VB.Label lblHeading 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PARTS LOCATION ADJUSTMENT FORM"
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
      Height          =   240
      Index           =   2
      Left            =   7215
      TabIndex        =   13
      Tag             =   "TRF"
      Top             =   450
      Width           =   3720
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
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   11
      Tag             =   "TLF"
      Top             =   450
      Width           =   2900
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
      Height          =   255
      Index           =   1
      Left            =   510
      TabIndex        =   12
      Tag             =   "TLF"
      Top             =   475
      Width           =   2900
   End
   Begin VB.Label lblHeading 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PARTS LOCATION ADJUSTMENT FORM"
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
      Height          =   240
      Index           =   3
      Left            =   7245
      TabIndex        =   14
      Tag             =   "TRF"
      Top             =   480
      Width           =   3720
   End
   Begin VB.Menu mnuADJ 
      Caption         =   "SpareParts"
      Visible         =   0   'False
      Begin VB.Menu mnuAdd 
         Caption         =   "Add Row         Alt+A"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete Row"
      End
   End
End
Attribute VB_Name = "SPLocAdj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsParts As New ADODB.Recordset
Dim ActRow As Integer, ActCol As Integer, BtnIndx As Integer
Dim AddNew As Boolean, LockFlx As Boolean, DispOnly As Boolean, PrtSrch As Boolean, PrtBln As Boolean, CustSrch As Boolean
'mySQL Variable for Editing Record
Dim MyTRANSNO As Double, MyDATETRANS As String

Private Sub ClearEntries()
    txtRefNo.Text = ""
    txtAdjNo.Text = ""
    txtAdjDate.Text = ""
    txtReason.Text = ""
    txtPrepBy.Text = ""
    txtAprvBy.Text = ""
    flxAdj.Rows = 1
    flxAdj.Rows = 2
    txtPrnCopy.Text = ""
End Sub

Private Sub LockData(bln As Boolean)
    Frame1.Enabled = Not bln
    Frame2.Enabled = Not bln
    txtPrepBy.Locked = bln
    txtAprvBy.Locked = bln
    LockFlx = bln
End Sub


Private Sub cboType_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        cboType.Text = Empty
        flxAdj.TextMatrix(ActRow, 3) = cboType.Tag
        cboType.Tag = "skiplostfocus"
        flxAdj.SetFocus
        flxAdj.Col = 3
        cboType.Visible = False
    End If
End Sub

Private Sub cboType_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        flxAdj.SetFocus
        flxAdj.Row = ActRow
        flxAdj.Col = 5
    End If
End Sub

Private Sub cboType_LostFocus()
    If cboType.Tag = "skiplostfocus" Then
        cboType.Tag = ""
        Exit Sub
    End If
    If cboType.ListIndex = -1 Then
        flxAdj.TextMatrix(ActRow, 3) = ""
        cboType.Visible = False
        Exit Sub
    End If
    flxAdj.TextMatrix(ActRow, 3) = cboType.Text
    cboType.Visible = False
End Sub


Private Sub cmdA_Click()
    Call mnuAdd_Click
End Sub


Private Sub LoadSpareParts()
    Dim Qry As String
        
    
    Qry = "SELECT  SPAREPARTS.SPAREPARTID, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPAREPARTS.LOCATION " & _
          "From SPAREPARTS ORDER BY SPAREPARTS.PARTNUMBER;"
    rsParts.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    Set flxParts.DataSource = rsParts
    With flxParts
        .Width = flxAdj.ColWidth(4)
        .ColWidth(0) = 0
        .ColWidth(1) = 0    'SPid
        .ColWidth(2) = flxAdj.ColWidth(4)   'PartNo
        .ColWidth(3) = 0    'Description
        .ColWidth(4) = 0    'Location
    End With
End Sub

Private Sub cmdTrans_Click(Index As Integer)
    Dim Ans As Integer, i As Integer
    Dim myMsg As String, PrnVal(0 To 10) As String, Qry As String, MyAcc As String
    
    Select Case Index
        Case 0  'New
            '***Access
            'Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE (((USERACCESS.EMPID)=" & UserID & ") AND ((USERACCESS.MENUINDX)=35));"
            'mySQL
            Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE USERACCESS.EMPID=" & UserID & " AND USERACCESS.MENUINDX=35;"
            
            MyAcc = GetFieldValue(Qry)
            If MyAcc <> "FULL ACCESS" Then
                MsgBox "You don't have full access to this transaction." & vbCrLf & "Please contact your System Administrator", vbExclamation + vbOKOnly, "Access Denied"
                Exit Sub
            End If
        
            BtnIndx = 5
            tmrFloat.Enabled = True
            Call ClearEntries
            Call LockData(False)
            txtAdjDate.Text = Date
            AddNew = True
            txtAdjNo.Text = GetAdjustmentNo
            txtPrepBy.Text = GetSettings("WAREHOUSE MAN")
            txtAprvBy.Text = GetSettings("OIC")
            If rsParts.State = adStateClosed Then Call LoadSpareParts
            txtRefNo.SetFocus
        Case 1  'Edit
            '***Access
            'Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE (((USERACCESS.EMPID)=" & UserID & ") AND ((USERACCESS.MENUINDX)=35));"
            'mySQL
            Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE USERACCESS.EMPID=" & UserID & " AND USERACCESS.MENUINDX=35;"
            
            MyAcc = GetFieldValue(Qry)
            If MyAcc <> "FULL ACCESS" Then
                MsgBox "You don't have full access to this transaction." & vbCrLf & "Please contact your System Administrator", vbExclamation + vbOKOnly, "Access Denied"
                Exit Sub
            End If
            If Val(txtPrnCopy.Text) > 0 Then
                MsgBox "Editing is not allowed because Adjustment was already printed.", vbInformation + vbOKOnly, "Edit Denied"
                Exit Sub
            ElseIf Val(txtAdjNo.Text) = 0 Then
                MsgBox "No available Adjustments to Edit", vbInformation + vbOKOnly, "Confirm P.O."
                Exit Sub
            End If
            BtnIndx = 5
            tmrFloat.Enabled = True
            AddNew = False
            For i = 1 To flxAdj.Rows - 1
                If flxAdj.TextMatrix(i, 3) = "ADDT`L" Then
                    flxAdj.TextMatrix(i, 5) = Replace(flxAdj.TextMatrix(i, 5), flxAdj.TextMatrix(i, 4) & ", ", "")
                End If
            Next i
            'mySQL
            MyTRANSNO = Val(txtAdjNo.Text)
            MyDATETRANS = Format(txtAdjDate.Text, "yyyy-mm-dd")
            '---------
            Call LockData(False)
        Case 2  'Summary
            FrmSummary.mySrc = "Location Adjustment"
            FrmSummary.Show
            'SPLocAdjSumm.Show
            Unload Me
            Exit Sub
        Case 3  'Print
            Ans = MsgBox("Do you want to Print Inventory Adjustment?", vbQuestion + vbYesNo, "Print")
            If Ans = vbNo Then Exit Sub
            flxPRN.Rows = 1
            For i = 1 To flxAdj.Rows - 1
                flxPRN.AddItem vbTab & flxAdj.TextMatrix(i, 1) & vbTab & flxAdj.TextMatrix(i, 3) & vbTab & flxAdj.TextMatrix(i, 4) & vbTab & flxAdj.TextMatrix(i, 5)
            Next i
            PrnVal(0) = "DATE: " & txtAdjDate.Text
            PrnVal(1) = "PART NUMBER"
            PrnVal(2) = "TYPE"
            PrnVal(3) = "PREV. LOCATION"
            PrnVal(4) = "NEW LOCATION"
            PrnVal(5) = txtPrepBy.Text
            PrnVal(6) = txtAprvBy.Text
            PrnVal(7) = IIf(txtRefNo.Text <> "", "REFERENCE: " & txtRefNo.Text, "")
            PrnVal(8) = "Warehouse Man"
            PrnVal(9) = "O.I.C."
            PrnVal(10) = "REASON: " & txtReason.Text
            Call PrintMyForm(PrnVal, "SP LOCATION ADJ", flxPRN)
            If Val(txtPrnCopy.Text) = 0 Then
                txtPrnCopy.Text = Val(txtPrnCopy.Text) + 1
                Call UpdateSPLoc
            Else
                txtPrnCopy.Text = Val(txtPrnCopy.Text) + 1
            End If
            '***Access
            'LYRcon.Execute "UPDATE TRANSFERLOC SET TRANSFERLOC.PRINTCOPY = [TRANSFERLOC].[PRINTCOPY]+1 WHERE (((TRANSFERLOC.TRANSLOCID)=" & txtTransLocID.Text & "));"
            'mySQL
            LYRcon.Execute "UPDATE TRANSFERLOC SET TRANSFERLOC.PRINTCOPY = TRANSFERLOC.PRINTCOPY+1 WHERE TRANSFERLOC.TRANSLOCID=" & txtTransLocID.Text & ";"
            
        Case 4  'New Location
            UpdateLoc.Show vbModal
            
        Case 5  'Exit
            'MotoRep.tbMain.Visible = True
            'MotoRep.imgBG.Visible = True
            MotoRep.toolBarButton.Visible = True
    
            Unload Me
            Exit Sub
        Case 6  'Save
            myMsg = ValidateLocAdj
            If myMsg <> "" Then
                MsgBox myMsg, vbCritical + vbOKOnly, "Validation Error"
                Exit Sub
            End If
            Ans = MsgBox("Do you want to Save Location Adjustment?", vbQuestion + vbYesNo, "Save")
            If Ans = vbNo Then Exit Sub
            Call SaveLocAdj
            MsgBox "Parts Location Adjustment Successfully Saved!", vbInformation + vbOKOnly, "Save"
                
            BtnIndx = 7
            tmrFloatRev.Enabled = True
            Call LockData(True)
            
        Case 7  'Cancel
            Ans = MsgBox("Do you want to Cancel Adjustment?", vbQuestion + vbYesNo, "Cancel")
            If Ans = vbNo Then Exit Sub
            BtnIndx = 7
            tmrFloatRev.Enabled = True
            If AddNew Then
                Call ClearEntries
            Else
                Call DisplayLocAdj(txtTransLocID.Text)
            End If
            Call LockData(True)
    End Select

End Sub

Private Sub flxAdj_Click()
    ActRow = flxAdj.Row
    ActCol = flxAdj.Col
    
    If ActCol = 0 Or ActRow = 0 Or LockFlx Then Exit Sub
    Select Case ActCol
        Case 1 'PartNumber
            With txtPartNo
                If flxAdj.TextMatrix(ActRow, 1) <> "" Then flxAdj.TextMatrix(ActRow, 7) = flxAdj.TextMatrix(ActRow, 1) & "|" & flxAdj.TextMatrix(ActRow, 2) & "|" & flxAdj.TextMatrix(ActRow, 3) & "|" & flxAdj.TextMatrix(ActRow, 4) & "|" & flxAdj.TextMatrix(ActRow, 5) & "|" & flxAdj.TextMatrix(ActRow, 6)
                .Top = flxAdj.CellTop + flxAdj.Top
                .Left = flxAdj.CellLeft + flxAdj.Left
                .Width = flxAdj.CellWidth
                .Text = flxAdj.Text
                .Visible = True
                .SelStart = 0
                .SelLength = Len(.Text)
                .ZOrder
                .SetFocus
            End With
        Case 3    'Adj. Type
            With cboType
                If flxAdj.TextMatrix(ActRow, 3) <> "" Then .Tag = flxAdj.TextMatrix(ActRow, 3)
                .Top = flxAdj.CellTop + flxAdj.Top
                .Left = flxAdj.CellLeft + flxAdj.Left
                .Width = flxAdj.CellWidth
                .Text = flxAdj.Text
                .Visible = True
                .SelStart = 0
                .SelLength = Len(.Text)
                .ZOrder
                .SetFocus
            End With
        Case 5  'New Location
            With txtNewLoc
                If flxAdj.TextMatrix(ActRow, 5) <> "" Then .Tag = flxAdj.TextMatrix(ActRow, 5)
                .Top = flxAdj.CellTop + flxAdj.Top
                .Left = flxAdj.CellLeft + flxAdj.Left
                .Width = flxAdj.CellWidth
                .Text = flxAdj.Text
                .Visible = True
                .SelStart = 0
                .SelLength = Len(.Text)
                .ZOrder
                .SetFocus
            End With
    End Select
End Sub

Private Sub flxAdj_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call flxAdj_Click
End Sub

Private Sub flxAdj_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    If LockFlx Then Exit Sub
    If Button = vbRightButton And ActCol = 0 Then
        Call PopupMenu(mnuADJ)
    End If
End Sub

Private Sub flxAdj_Scroll()
    txtPartNo.Visible = False
    cboType.Visible = False
    txtNewLoc.Visible = False
End Sub

Private Sub flxParts_Click()
    Dim TtlPrice As Currency
    
    With flxParts
        txtPartNo.Visible = True
        txtPartNo.ZOrder
        PrtSrch = True
        txtPartNo.Text = .TextMatrix(.Row, 2)
        'Insert Additional retrieval here
        flxAdj.TextMatrix(ActRow, 2) = .TextMatrix(.Row, 3) 'Description
        flxAdj.TextMatrix(ActRow, 4) = .TextMatrix(.Row, 4) 'SPid
        flxAdj.TextMatrix(ActRow, 6) = .TextMatrix(.Row, 1) 'SPid
    End With
End Sub

Private Sub flxParts_EnterCell()
    Dim TtlPrice As Currency
    
    With flxParts
        txtPartNo.Visible = True
        txtPartNo.ZOrder
        PrtSrch = True
        txtPartNo.Text = .TextMatrix(.Row, 2)
        'Insert Additional retrieval here
        flxAdj.TextMatrix(ActRow, 2) = .TextMatrix(.Row, 3) 'Description
        flxAdj.TextMatrix(ActRow, 4) = .TextMatrix(.Row, 4) 'SPid
        flxAdj.TextMatrix(ActRow, 6) = .TextMatrix(.Row, 1) 'SPid
    End With
End Sub

Private Sub flxParts_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim Tmp() As String
    Dim i As Integer
    
    If KeyCode = vbKeyEscape Then
        If flxAdj.TextMatrix(ActRow, 7) <> "" Then
            Tmp = Split(flxAdj.TextMatrix(ActRow, 7), "|")
            For i = LBound(Tmp) To UBound(Tmp)
                flxAdj.TextMatrix(ActRow, i + 1) = Tmp(i)
            Next i
        Else
            For i = 1 To 7
                flxAdj.TextMatrix(ActRow, i) = ""
            Next i
        End If
        txtPartNo.Visible = False
        flxParts.Visible = False
        flxAdj.SetFocus
        flxAdj.Col = 1
    End If

End Sub

Private Sub flxParts_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call txtPartNo_LostFocus
        flxAdj.SetFocus
        flxAdj.Row = ActRow
        flxAdj.Col = 3
    End If
End Sub

Private Sub Form_Load()
    Dim i As Integer
    
    Call LockData(True)
    cboType.AddItem "ADDT`L"
    cboType.AddItem "CHANGE"
    cmdA.Top = 15000
    MotoRep.sBar.Panels(6).Text = "Spare Parts Location Adjustment Form"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    MotoRep.sBar.Panels(6).Text = ""
    If rsParts.State = adStateOpen Then rsParts.Close
End Sub

Private Sub mnuAdd_Click()
    If flxAdj.TextMatrix(ActRow, 6) = "" Or flxAdj.TextMatrix(ActRow, 3) = "" Or (flxAdj.TextMatrix(ActRow, 3) = "ADDT`L" And flxAdj.TextMatrix(ActRow, 5) = "") Then
        MsgBox "Incomplete entry on last row!", vbCritical + vbOKOnly, "Add Row failed"
        flxAdj.SetFocus
        flxAdj.Col = ActCol
        Exit Sub
    Else
        If flxAdj.Rows = 6 Then
            MsgBox "You can only add upto 5 Part Numbers per Adjustment!", vbExclamation + vbOKOnly, "Limit Exceed"
            Exit Sub
        End If
        If ActRow = flxAdj.Rows - 1 Then flxAdj.AddItem ""
        flxAdj.SetFocus
        flxAdj.Col = 1
        flxAdj.Row = flxAdj.Rows - 1
    End If
End Sub


Private Sub mnuDelete_Click()
    Dim Ans As Integer, i As Integer
    
    Ans = MsgBox("Do you want to Delete Row?", vbQuestion + vbYesNo, "Delete")
    If Ans = vbYes Then
        If ActRow = 1 And flxAdj.Rows = 2 Then
            flxAdj.Rows = 1
            flxAdj.Rows = 2
        Else
            flxAdj.RemoveItem ActRow
        End If
    End If
End Sub

Private Sub RXResizer1_ResizeComplete()
    With flxAdj
        .ColWidth(0) = 350
        .ColWidth(1) = (.Width - 380) * 0.2
        .ColWidth(2) = (.Width - 380) * 0.32
        .ColWidth(3) = (.Width - 380) * 0.12
        .ColWidth(4) = (.Width - 380) * 0.17
        .ColWidth(5) = (.Width - 380) * 0.17
        .ColWidth(6) = 0
        .ColWidth(7) = 0
    End With
End Sub


Private Sub SaveLocAdj()
    Dim rsLocAdj As New ADODB.Recordset
    Dim Qry As String, NLoc As String, i As Integer
    
    'Saving Inventory Adjustment
    '***Access
    'With rsLocAdj
    '    If AddNew Then
    '        .Open "TransferLoc", LYRcon, adOpenKeyset, adLockOptimistic
    '        .AddNew
    '       Else
            '***Access
            'Qry = "SELECT TRANSFERLOC.* From TRANSFERLOC WHERE (((TRANSFERLOC.TRANSLOCID)=" & txtTransLocID.Text & "));"
            
    '        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    '        .MoveFirst
    '    End If
    '    !TRANSNO = Val(txtAdjNo.Text)
    '    !REFERENCE = txtRefNo.Text
    '    !DATETRANS = txtAdjDate.Text
    '    !REASON = txtReason.Text
    '    !PrepBy = txtPrepBy.Text
    '    !AprvBy = txtAprvBy.Text
    '    .Update
    '    txtTransLocID.Text = !TRANSLOCID
    '    .Close
    'End With
    
    'mySQL
    If AddNew Then
        'mySQL Statement
        LYRcon.Execute "INSERT INTO transferloc(TRANSNO, REFERENCE, DATETRANS, REASON, PREPBY, APRVBY)" & _
            "VALUES(" & Val(txtAdjNo.Text) & ", " & txtRefNo.Text & ", '" & Format(txtAdjDate.Text, "yyyy-mm-dd") & "', " & _
            "'" & txtReason.Text & "', '" & txtPrepBy.Text & "', '" & txtAprvBy.Text & "');"
    Else
        'Add unique Tag for Update record
        LYRcon.Execute "UPDATE transferloc SET TRANSNO = " & Val(txtAdjNo.Text) & ", REFERENCE = " & txtRefNo.Text & ", " & _
            "DATETRANS = '" & Format(txtAdjDate.Text, "yyyy-mm-dd") & "', REASON = '" & txtReason.Text & "', " & _
            "PREPBY = '" & txtPrepBy.Text & "', APRVBY = '" & txtAprvBy.Text & "'" & _
            " Where transferloc.TRANSNO = " & MyTRANSNO & " AND transferloc.DATETRANS = '" & MyDATETRANS & "';"
        
    End If
    'SAVE Details
    If Not AddNew Then
        '***Access
        'LYRcon.Execute "DELETE TRANSLOCDTLS.* From TRANSLOCDTLS WHERE (((TRANSLOCDTLS.TRANSLOCID)=" & txtTransLocID.Text & "));"
        'mySQL
        LYRcon.Execute "DELETE From TRANSLOCDTLS WHERE TRANSLOCDTLS.TRANSLOCID=" & txtTransLocID.Text & ";"
        
    End If
    With flxAdj
        For i = 1 To .Rows - 1
            If .TextMatrix(i, 3) = "CHANGE" Then
                NLoc = IIf(.TextMatrix(i, 5) <> "", .TextMatrix(i, 5), Empty)
            Else
                NLoc = .TextMatrix(i, 4) & ", " & .TextMatrix(i, 5)
            End If
            LYRcon.Execute "INSERT INTO TRANSLOCDTLS ( TRANSLOCID, SPAREPARTID, ATYPE, PREVLOC, NEWLOC) " & _
                           "VALUES (" & txtTransLocID.Text & ", " & .TextMatrix(i, 6) & ", '" & .TextMatrix(i, 3) & "', '" & .TextMatrix(i, 4) & "', '" & Left$(NLoc, 50) & "');"
        Next i
    End With
    
End Sub


Private Sub UpdateSPLoc()
    Dim rsITR As New ADODB.Recordset
    Dim i As Integer, NLoc As String
    
    With flxAdj
        For i = 1 To .Rows - 1
            'Update SpareParts Location
            If .TextMatrix(i, 3) = "CHANGE" Then
                NLoc = IIf(.TextMatrix(i, 5) <> "", .TextMatrix(i, 5), Empty)
            Else
                NLoc = .TextMatrix(i, 4) & ", " & .TextMatrix(i, 5)
            End If
            '***Access
            'LYRcon.Execute "UPDATE SPAREPARTS SET SPAREPARTS.LOCATION = '" & NLoc & "'" & _
                           " WHERE (((SPAREPARTS.SPAREPARTID)=" & .TextMatrix(i, 6) & "));"
            'mySQL
            LYRcon.Execute "UPDATE SPAREPARTS SET SPAREPARTS.LOCATION = '" & NLoc & "'" & _
                           " WHERE SPAREPARTS.SPAREPARTID=" & .TextMatrix(i, 6) & ";"
            
        Next i
    End With
End Sub


Function GetAdjustmentNo() As String
    Dim rsAdjCntr As New ADODB.Recordset
    Dim Qry As String
    Dim MyNo As Long
    
    Qry = "SELECT Max(TRANSFERLOC.TRANSNO) AS MaxOfTRANSNO FROM TRANSFERLOC;"

    With rsAdjCntr
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            MyNo = IIf(IsNull(.Fields(0)), 1, .Fields(0) + 1)
        Else
            MyNo = 1
        End If
        .Close
    End With
    GetAdjustmentNo = Format(MyNo, "0000")
End Function


Function ValidateLocAdj() As String
    Dim vMsg As String, Qry As String
    Dim i As Integer, Cntr As Integer
    
    vMsg = ""

    If Not IsDate(txtAdjDate.Text) Then vMsg = "Invalid Adjustment Date"
    If txtAdjNo.Text = "" Then
        vMsg = vMsg & vbCrLf & "Incomplete entry on field Control No."
    Else
        If AddNew Then
            '***Access
            'Qry = "SELECT Val([TRANSNO]) AS myTransNo From TRANSFERLOC WHERE (((Val([TRANSNO]))=" & Val(txtAdjNo.Text) & "));"
            'mySQL
            Qry = "SELECT round(TRANSFERLOC.TRANSNO) AS myTransNo From TRANSFERLOC WHERE round(TRANSFERLOC.TRANSNO)=" & Val(txtAdjNo.Text) & ";"
            
            If CtrlNumExist(Qry) Then
                vMsg = vMsg & vbCrLf & "Control No. " & txtAdjNo.Text & " already exist!"
            End If
        End If
    End If
    If txtReason.Text = "" Then vMsg = vMsg & vbCrLf & "Incomplete entry on field Reason"
    If txtAprvBy.Text = "" Then vMsg = vMsg & vbCrLf & "Incomplete entry on field Approved By"
    If txtPrepBy.Text = "" Then vMsg = vMsg & vbCrLf & "Incomplete entry on field Prepared By"
    If flxAdj.TextMatrix(flxAdj.Rows - 1, 6) = "" Or flxAdj.TextMatrix(flxAdj.Rows - 1, 3) = "" Or (flxAdj.TextMatrix(flxAdj.Rows - 1, 3) = "ADDT`L" And flxAdj.TextMatrix(flxAdj.Rows - 1, 5) = "") Then
        If flxAdj.Rows <= 2 Then
            vMsg = vMsg & vbCrLf & "Incomplete entry on last row of Adjusted Spare Parts"
        Else
            If flxAdj.TextMatrix(flxAdj.Rows - 1, 6) = "" Then
                flxAdj.RemoveItem flxAdj.Rows - 1
            Else
                vMsg = vMsg & vbCrLf & "Incomplete entry on last row of Adjusted Spare Parts"
            End If
        End If
    End If
    ValidateLocAdj = vMsg
    
End Function

Public Sub DisplayLocAdj(MyRef As String)
    Dim rsAdj As New ADODB.Recordset
    Dim rsDtls As New ADODB.Recordset
    Dim Qry As String, AdjType As String
    
    Call ClearEntries
    If rsParts.State = adStateClosed Then Call LoadSpareParts
    '***Access
    'Qry = "SELECT TRANSFERLOC.* From TRANSFERLOC WHERE (((TRANSFERLOC.TRANSLOCID)=" & MyRef & "));"
    
    'mySQL
    Qry = "SELECT TRANSFERLOC.* From TRANSFERLOC WHERE TRANSFERLOC.TRANSLOCID=" & MyRef & ";"
    
    With rsAdj
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            txtAdjNo.Text = Format(!TransNo, "0000")
            txtRefNo.Text = IIf(IsNull(!REFERENCE), "", !REFERENCE)
            txtAdjDate.Text = !DATETRANS
            txtReason.Text = !REASON
            txtPrepBy.Text = !PREPBY
            txtAprvBy.Text = !APRVBY
            txtPrnCopy.Text = !PRINTCOPY
            txtTransLocID.Text = MyRef
        End If
        .Close
    End With
    
    Qry = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, TRANSLOCDTLS.ATYPE, TRANSLOCDTLS.PREVLOC, TRANSLOCDTLS.NEWLOC, TRANSLOCDTLS.SPAREPARTID " & _
          "FROM TRANSLOCDTLS INNER JOIN SPAREPARTS ON TRANSLOCDTLS.SPAREPARTID = SPAREPARTS.SPAREPARTID WHERE TRANSLOCDTLS.TRANSLOCID=" & MyRef & ";"
    
    With rsDtls
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            flxAdj.Rows = 1
            Do While Not .EOF
                flxAdj.AddItem vbTab & !PartNumber & vbTab & !Description & vbTab & !aType & vbTab & !PrevLoc & vbTab & IIf(IsNull(!NewLoc), "", !NewLoc) & vbTab & !SparepartID
                .MoveNext
            Loop
        End If
        .Close
    End With
    Call LockData(True)
End Sub


Private Sub tmrDrop_Timer()
    cmdTrans(BtnIndx).Top = cmdTrans(BtnIndx).Top + 50
    If cmdTrans(BtnIndx).Top >= 0 Then
        cmdTrans(BtnIndx).Top = 0
        BtnIndx = BtnIndx + 1
        If BtnIndx = 8 Then
            tmrDrop.Enabled = False
        End If
    End If
End Sub

Private Sub tmrDropRev_Timer()
    cmdTrans(BtnIndx).Top = cmdTrans(BtnIndx).Top + 50
    If cmdTrans(BtnIndx).Top >= 0 Then
        cmdTrans(BtnIndx).Top = 0
        BtnIndx = BtnIndx + 1
        If BtnIndx = 6 Then
            tmrDropRev.Enabled = False
            'cmdTrans(0).SetFocus
        End If
    End If
End Sub

Private Sub tmrFloat_Timer()
    cmdTrans(BtnIndx).Top = cmdTrans(BtnIndx).Top - 50
    If cmdTrans(BtnIndx).Top <= -375 Then
        cmdTrans(BtnIndx).Visible = False
        BtnIndx = BtnIndx - 1
        If BtnIndx = -1 Then
            tmrFloat.Enabled = False
            BtnIndx = 6 'Index of Button Save
            cmdTrans(6).Top = -375
            cmdTrans(6).Left = 0
            cmdTrans(6).Visible = True
            cmdTrans(7).Top = -375
            cmdTrans(7).Left = cmdTrans(4).Width
            cmdTrans(7).Visible = True
            tmrDrop.Enabled = True
        End If
    End If
End Sub

Private Sub tmrFloatRev_Timer()
    Dim i As Integer
    
    cmdTrans(BtnIndx).Top = cmdTrans(BtnIndx).Top - 50
    If cmdTrans(BtnIndx).Top <= -375 Then
        cmdTrans(BtnIndx).Visible = False
        BtnIndx = BtnIndx - 1
        If BtnIndx = 5 Then 'Index of Button Exit
            tmrFloatRev.Enabled = False
            BtnIndx = 0
            For i = 0 To 5
                cmdTrans(i).Visible = True
            Next i
            tmrDropRev.Enabled = True
        End If
    End If
End Sub


Private Sub txtAdjDate_KeyPress(KeyAscii As Integer)
    Dim strValid As String
    
    strValid = "1234567890/-"
    If KeyAscii > 26 Then
        If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
    If KeyAscii = 13 Then
        flxAdj.SetFocus
        flxAdj.Row = 1
        flxAdj.Col = 1
    End If
End Sub

Private Sub txtAdjNo_GotFocus()
    If Val(txtAdjNo.Text) > 0 Then
        txtAdjNo.SelStart = 0
        txtAdjNo.SelLength = Len(txtAdjNo.Text)
    End If
End Sub

Private Sub txtAdjNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtRefNo.SetFocus
End Sub

Private Sub txtAdjNo_LostFocus()
    Dim Qry As String
    
    'Verify AdjNo if existing
    If Val(txtAdjNo.Text) > 0 And AddNew Then
        '***Access
        'Qry = "SELECT Val([TRANSNO]) AS myTransNo From TRANSFERLOC WHERE (((Val([TRANSNO]))=" & Val(txtAdjNo.Text) & "));"
        'mySQL
        Qry = "SELECT Round(TRANSFERLOC.TRANSNO) AS myTransNo From TRANSFERLOC WHERE Round(TRANSFERLOC.TRANSNO)=" & Val(txtAdjNo.Text) & ";"
        
        If CtrlNumExist(Qry) Then
            MsgBox "Control No. " & txtAdjNo.Text & " already exist!", vbExclamation + vbOKOnly, "Validation Error"
        End If
    End If
    
End Sub

Private Sub txtaprvby_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        If LockFlx Then
            cmdTrans(0).SetFocus
        Else
            cmdTrans(6).SetFocus
        End If
    End If
End Sub
                                            

Private Sub txtNewLoc_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        txtNewLoc.Text = Empty
        flxAdj.TextMatrix(ActRow, 5) = txtNewLoc.Tag
        txtNewLoc.Tag = "skiplostfocus"
        flxAdj.SetFocus
        flxAdj.Col = 5
        txtNewLoc.Visible = False
    End If
End Sub

Private Sub txtNewLoc_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        flxAdj.SetFocus
        flxAdj.Row = ActRow
        flxAdj.Col = 0
    End If
End Sub

Private Sub txtNewLoc_LostFocus()
    Dim Qry As String, Ans As Integer
    
    If txtNewLoc.Tag = "skiplostfocus" Then
        txtNewLoc.Tag = ""
        Exit Sub
    End If
    If txtNewLoc.Text <> "" Then
        If txtNewLoc.Text = flxAdj.TextMatrix(ActRow, 4) Then
            MsgBox "New Location is the same with Previous Location", vbCritical + vbOKOnly, "Validation Error"
            txtNewLoc.SetFocus
            txtNewLoc.SelStart = 0
            txtNewLoc.SelLength = Len(txtNewLoc.Text)
            Exit Sub
        End If
        '***Access
        'Qry = "SELECT SPAREPARTS.* From SPAREPARTS WHERE (((SPAREPARTS.LOCATION) Like '%" & txtNewLoc.Text & "%'));"
        
        'mySQL
        Qry = "SELECT SPAREPARTS.* From SPAREPARTS WHERE (((SPAREPARTS.LOCATION) Like '%" & txtNewLoc.Text & "%'));"
        
        If CtrlNumExist(Qry) Then
            Ans = MsgBox("Location " & txtNewLoc.Text & " is already used!" & vbCrLf & "Do you want to Continue?", vbQuestion + vbYesNo, "Location Exist")
            If Ans = vbNo Then
                txtNewLoc.Tag = ""
                txtNewLoc.Visible = False
                flxAdj.SetFocus
                flxAdj.Row = flxAdj.Rows - 1
                flxAdj.Col = 1
                Exit Sub
            End If
        End If
    End If
    flxAdj.TextMatrix(ActRow, 5) = txtNewLoc.Text
    txtNewLoc.Visible = False
    If Ans = vbYes Then
        flxAdj.SetFocus
        flxAdj.Row = flxAdj.Rows - 1
        flxAdj.Col = 1
    End If
End Sub

Private Sub txtPartNo_Change()
    Dim TtlPrice As Currency, DiscPrice As Currency
    Dim PctDisc As Single
    
    With flxParts
        If PrtSrch Then
            PrtSrch = False
            Exit Sub
        End If
        .Top = txtPartNo.Height + txtPartNo.Top
        .Left = txtPartNo.Left
        .Width = txtPartNo.Width
        .Visible = True
        .ZOrder
        If txtPartNo.Text = "" Then .Visible = False: Exit Sub
        rsParts.Find "[PARTNUMBER] LIKE '" + txtPartNo.Text + "%'", , adSearchForward, 1
        If Not rsParts.EOF Then
            If Not PrtBln Then
                txtPartNo.Tag = Len(txtPartNo.Text)
                PrtBln = True
            End If
            .TopRow = rsParts.AbsolutePosition - 1
            .Row = rsParts.AbsolutePosition - 1
            .Col = 0: .ColSel = 4
            txtPartNo.Text = rsParts!PartNumber
            'Insert Additional retrieval here
            flxAdj.TextMatrix(ActRow, 6) = flxParts.TextMatrix(flxParts.Row, 1) 'SP id
            flxAdj.TextMatrix(ActRow, 2) = flxParts.TextMatrix(flxParts.Row, 3) 'Description
            flxAdj.TextMatrix(ActRow, 4) = flxParts.TextMatrix(flxParts.Row, 4) 'Description
            
            txtPartNo.SelStart = Val(txtPartNo.Tag)
            txtPartNo.SelLength = Len(txtPartNo.Text)
            PrtBln = False
        Else
'            rsParts.Requery
            flxAdj.TextMatrix(ActRow, 2) = ""
            flxAdj.TextMatrix(ActRow, 6) = ""
            flxAdj.TextMatrix(ActRow, 4) = ""
            .Visible = False
        End If

    End With
End Sub

Private Sub txtPartNo_GotFocus()
    With flxParts
        .Top = txtPartNo.Height + txtPartNo.Top
        .Left = txtPartNo.Left
        .Width = txtPartNo.Width
    End With
End Sub

Private Sub txtPartNo_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    Dim Tmp() As String
    
    Select Case KeyCode
        Case vbKeyDown
            flxParts.Tag = "skiplostfocus"
            flxParts.Visible = True
            flxParts.SetFocus
        Case vbKeyEscape
            txtPartNo.Text = Empty
            If flxAdj.TextMatrix(ActRow, 7) <> "" Then
                Tmp = Split(flxAdj.TextMatrix(ActRow, 7), "|")
                For i = LBound(Tmp) To UBound(Tmp)
                    flxAdj.TextMatrix(ActRow, i + 1) = Tmp(i)
                Next i
            Else
                For i = 1 To 7
                    flxAdj.TextMatrix(ActRow, i) = ""
                Next i
            End If
            flxParts.Tag = "skiplostfocus"
            flxAdj.SetFocus
            flxAdj.Col = 1
            txtPartNo.Visible = False
            flxParts.Visible = False
        Case vbKeyDelete
            PrtSrch = True
        Case vbKeyBack
            PrtSrch = True
    End Select
End Sub

Private Sub txtPartNo_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        If flxAdj.TextMatrix(ActRow, 6) = "" Then   'No SPID
            flxAdj.SetFocus
            flxAdj.Row = ActRow
            flxAdj.Col = 1
        Else
            flxAdj.SetFocus
            flxAdj.Row = ActRow
            flxAdj.Col = 3
        End If
        flxParts.Visible = False
    End If
End Sub

Private Sub txtPartNo_LostFocus()
    Dim i As Integer
    
    If flxParts.Tag = "skiplostfocus" Then
        flxParts.Tag = ""
        Exit Sub
    End If
    If flxAdj.TextMatrix(ActRow, 6) = "" Then
        MsgBox "Unrecognized Part Number!" & vbCrLf & "Please select a Part Number from list.", vbExclamation + vbOKOnly, "Warning"
        txtPartNo.Visible = False
        flxParts.Visible = False
        For i = 2 To 7
            flxAdj.TextMatrix(ActRow, i) = ""
        Next i
        Exit Sub
    End If
    If flxAdj.TextMatrix(ActRow, 4) = "" Then
        MsgBox "This Part Number has no existing Location" & vbCrLf & "Please use NEW LOCATION to process this Part Number", vbExclamation + vbOKOnly, "Warning"
        txtPartNo.Visible = False
        flxParts.Visible = False
        For i = 2 To 7
            flxAdj.TextMatrix(ActRow, i) = ""
        Next i
        flxAdj.SetFocus
        flxAdj.Col = 0
        flxAdj.Row = ActRow
        Exit Sub
    End If
    
    flxAdj.TextMatrix(ActRow, 1) = txtPartNo.Text
    txtPartNo.Visible = False
    flxParts.Visible = False
End Sub

Private Sub txtPrepBy_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtAprvBy.SetFocus
End Sub

Private Sub txtReason_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtPrepBy.SetFocus
End Sub


Private Sub txtrefno_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtAdjDate.SetFocus
End Sub

