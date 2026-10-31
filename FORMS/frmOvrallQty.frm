VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmOvrallQty 
   Caption         =   "Overall Quantity"
   ClientHeight    =   6720
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8670
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6720
   ScaleWidth      =   8670
   WindowState     =   2  'Maximized
   Begin lvButton.lvButtons_H cmdExit 
      Height          =   375
      Left            =   0
      TabIndex        =   4
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
      Image           =   "frmOvrallQty.frx":0000
      cBack           =   -2147483633
   End
   Begin TabDlg.SSTab tabProcessing 
      Height          =   975
      Left            =   2400
      TabIndex        =   5
      Tag             =   "xyf"
      Top             =   2880
      Visible         =   0   'False
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   1720
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "SP PRICING"
      TabPicture(0)   =   "frmOvrallQty.frx":059A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label24"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Image2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      Begin VB.Image Image2 
         Height          =   720
         Left            =   240
         Picture         =   "frmOvrallQty.frx":05B6
         Stretch         =   -1  'True
         Top             =   120
         Width           =   705
      End
      Begin VB.Label Label24 
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
         Left            =   120
         TabIndex        =   6
         Tag             =   "xyf"
         Top             =   360
         Width           =   3855
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxOverallQty 
      Height          =   5295
      Left            =   120
      TabIndex        =   3
      Tag             =   "xyf"
      Top             =   1200
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   9340
      _Version        =   393216
      FillStyle       =   1
      ScrollBars      =   2
      SelectionMode   =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   6960
      Top             =   120
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Tag             =   "xyf"
      Top             =   240
      Width           =   8415
      Begin VB.TextBox txtFind 
         Height          =   375
         Left            =   1440
         TabIndex        =   1
         Top             =   360
         Width           =   3375
      End
      Begin lvButton.lvButtons_H cmdFind 
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   360
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
         Image           =   "frmOvrallQty.frx":0FAE
         cBack           =   -2147483633
      End
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   7560
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   8910
      m_MDI_Height    =   7290
      OriginalBorder  =   10
      OriginalCaption =   31
   End
   Begin VB.Menu mnu 
      Caption         =   " "
      Visible         =   0   'False
   End
End
Attribute VB_Name = "frmOvrallQty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Qry     As String
Dim i       As Integer
Dim j       As Integer
Dim fRow    As Integer
Dim fCol    As Integer
Dim TCol    As Integer
Dim ActRow  As Integer


Private Sub cmdExit_Click()
    Unload Me
    'MotoRep.tbMain.Visible = True
        MotoRep.imgBG.Visible = True
    MotoRep.toolBarButton.Visible = True
End Sub

Private Sub cmdFind_Click()
    
    If txtFind.Text = "" Then Exit Sub
    With flxOverallQty
        For i = fRow To .Rows - 1
            For j = fCol To TCol - 1
                If InStr(.TextMatrix(i, j), UCase(txtFind.Text)) > 0 Then
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

Private Sub flxOverallQty_DblClick()
    frmBrkDwnQty.txtSPID.Text = flxOverallQty.TextMatrix(flxOverallQty.Row, 4)
    frmBrkDwnQty.Caption = flxOverallQty.TextMatrix(flxOverallQty.Row, 1) & " (" & flxOverallQty.TextMatrix(flxOverallQty.Row, 2) & ")"
    frmBrkDwnQty.Show vbModal
End Sub

Private Sub Form_Load()
    Timer1.Enabled = True
    TCol = 4
    
End Sub

Private Sub RXResizer1_ResizeComplete()
    With flxOverallQty
        .FormatString = " |<PARTNUMBER|<DESCRIPTION|>QUANTITY"
        .ColWidth(0) = 300
        .ColWidth(1) = (.Width - 300) * 0.2
        .ColWidth(2) = (.Width - 300) * 0.5
        .ColWidth(3) = (.Width - 300) * 0.1
        .ColWidth(4) = 0
    End With
End Sub

Private Sub Timer1_Timer()
    On Error GoTo LocalErr
        tabProcessing.Visible = True
        DoEvents
        Qry = "call sp_mrcretails.sp_fetch_OverallQty(" & SysBranchID & ");"
        If rsAnyUse.State <> 0 Then rsAnyUse.Close
        Set rsAnyUse = LYRcon.Execute(Qry)
        Call loadFlxRecord(rsAnyUse, flxOverallQty)
        Timer1.Enabled = False
        Call RXResizer1_ResizeComplete
        tabProcessing.Visible = False
    Exit Sub
LocalErr:
    MsgBox Err.Description, vbCritical, "Overall Quantity"
    Timer1.Enabled = False
End Sub


Private Sub txtFind_Change()
    fRow = 1
    fCol = 1
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call cmdFind_Click
End Sub
