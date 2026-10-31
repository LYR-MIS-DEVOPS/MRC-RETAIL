VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form SPorderSum 
   Caption         =   "Spare Parts P.O. Summary"
   ClientHeight    =   9090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12930
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9090
   ScaleWidth      =   12930
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab tabCancel 
      Height          =   1935
      Left            =   4080
      TabIndex        =   33
      Top             =   9360
      Visible         =   0   'False
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   3413
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "SPOrderSum.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Image1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Shape1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdTabCancel"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Pbar"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      Begin MSComctlLib.ProgressBar Pbar 
         Height          =   135
         Left            =   120
         TabIndex        =   36
         Top             =   960
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   238
         _Version        =   393216
         Appearance      =   0
      End
      Begin lvButton.lvButtons_H cmdTabCancel 
         Height          =   375
         Left            =   1320
         TabIndex        =   35
         Top             =   1320
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   661
         Caption         =   "OK"
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
         Image           =   "SPOrderSum.frx":001C
         cBack           =   -2147483633
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H80000002&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0C0C0&
         Height          =   255
         Left            =   20
         Top             =   20
         Width           =   3700
      End
      Begin VB.Label Label2 
         Caption         =   "Please wait cancel on progress..."
         Height          =   255
         Left            =   960
         TabIndex        =   34
         Top             =   600
         Width           =   2415
      End
      Begin VB.Image Image1 
         Height          =   360
         Left            =   360
         Picture         =   "SPOrderSum.frx":05B6
         Stretch         =   -1  'True
         Top             =   480
         Width           =   360
      End
   End
   Begin VB.TextBox txtSuppID 
      Height          =   285
      Left            =   240
      TabIndex        =   29
      Top             =   9240
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   50
      TabIndex        =   14
      Tag             =   "XYF"
      Top             =   8040
      Width           =   12855
      Begin VB.TextBox txtRemarks 
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
         Left            =   9480
         TabIndex        =   26
         Tag             =   "XYF"
         Top             =   503
         Width           =   3255
      End
      Begin VB.TextBox txtBackOrder 
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
         Left            =   9480
         TabIndex        =   25
         Tag             =   "XYF"
         Top             =   173
         Width           =   1095
      End
      Begin VB.TextBox txtRec 
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
         Left            =   6120
         TabIndex        =   24
         Tag             =   "XYF"
         Top             =   503
         Width           =   1095
      End
      Begin VB.TextBox txtOrder 
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
         Left            =   6120
         TabIndex        =   23
         Tag             =   "XYF"
         Top             =   173
         Width           =   1095
      End
      Begin VB.ComboBox cboPO 
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
         Left            =   1440
         TabIndex        =   22
         Tag             =   "XYF"
         Top             =   480
         Width           =   2895
      End
      Begin VB.ComboBox cboSupplier 
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
         Left            =   1440
         TabIndex        =   21
         Tag             =   "XYF"
         Top             =   150
         Width           =   2895
      End
      Begin VB.Label lblPctBO 
         Caption         =   "0.00%"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   10680
         TabIndex        =   28
         Tag             =   "XYF"
         Top             =   203
         Width           =   855
      End
      Begin VB.Label lblPctRec 
         Caption         =   "0.00%"
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
         Left            =   7320
         TabIndex        =   27
         Tag             =   "XYF"
         Top             =   533
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "REMARKS"
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
         Index           =   5
         Left            =   8280
         TabIndex        =   20
         Tag             =   "XYF"
         Top             =   518
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "BACK ORDER:"
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
         Left            =   8280
         TabIndex        =   19
         Tag             =   "XYF"
         Top             =   188
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "QTY. RECEIVED"
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
         Left            =   4800
         TabIndex        =   18
         Tag             =   "XYF"
         Top             =   518
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "QTY. ORDERED"
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
         Left            =   4800
         TabIndex        =   17
         Tag             =   "XYF"
         Top             =   188
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "P.O. NUMBER"
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
         TabIndex        =   16
         Tag             =   "XYF"
         Top             =   518
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "SUPPLIER"
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
         Index           =   0
         Left            =   240
         TabIndex        =   15
         Tag             =   "XYF"
         Top             =   188
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Height          =   700
      Left            =   50
      TabIndex        =   9
      Tag             =   "XYF"
      Top             =   400
      Width           =   12855
      Begin VB.TextBox txtFind 
         Height          =   315
         Left            =   10080
         TabIndex        =   10
         Tag             =   "XYF"
         Top             =   230
         Width           =   2655
      End
      Begin lvButton.lvButtons_H cmdFind 
         Height          =   375
         Left            =   8760
         TabIndex        =   11
         Tag             =   "XY"
         Top             =   195
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
         Image           =   "SPOrderSum.frx":0B40
         cBack           =   -2147483633
      End
      Begin VB.Label lblHeading 
         BackStyle       =   0  'Transparent
         Caption         =   "MOTO-R REPUBLIC CORP."
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
         TabIndex        =   12
         Tag             =   "XYF"
         Top             =   240
         Width           =   3495
      End
      Begin VB.Label lblHeading 
         BackStyle       =   0  'Transparent
         Caption         =   "MOTO-R REPUBLIC CORP."
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
         TabIndex        =   13
         Tag             =   "XYF"
         Top             =   270
         Width           =   3495
      End
   End
   Begin MSFlexGridLib.MSFlexGrid flxOrder 
      Height          =   6975
      Left            =   0
      TabIndex        =   7
      Tag             =   "XY"
      Top             =   1080
      Width           =   12855
      _ExtentX        =   22675
      _ExtentY        =   12303
      _Version        =   393216
      Cols            =   14
      RowHeightMin    =   275
      BackColorBkg    =   -2147483633
      SelectionMode   =   1
      AllowUserResizing=   1
      FormatString    =   $"SPOrderSum.frx":10DA
   End
   Begin VB.Timer tmrDropRev 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   4560
      Top             =   8280
   End
   Begin VB.Timer tmrFloatRev 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   4560
      Top             =   7800
   End
   Begin VB.Timer tmrDrop 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3960
      Top             =   8280
   End
   Begin VB.Timer tmrFloat 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3960
      Top             =   7800
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
      Caption         =   "&View Order"
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
      Image           =   "SPOrderSum.frx":11E1
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   1
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
      Image           =   "SPOrderSum.frx":177B
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   1
      Left            =   1680
      TabIndex        =   2
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
      Image           =   "SPOrderSum.frx":1D15
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   5
      Left            =   8400
      TabIndex        =   3
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
      Image           =   "SPOrderSum.frx":22AF
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   6
      Left            =   10080
      TabIndex        =   4
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
      Image           =   "SPOrderSum.frx":2849
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   3
      Left            =   5040
      TabIndex        =   8
      Top             =   0
      Visible         =   0   'False
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
      Caption         =   "&Back Order"
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
      Image           =   "SPOrderSum.frx":2DE3
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   4
      Left            =   6720
      TabIndex        =   30
      Top             =   0
      Visible         =   0   'False
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
      Caption         =   "&Refresh"
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
      Image           =   "SPOrderSum.frx":337D
      cBack           =   -2147483633
   End
   Begin MSFlexGridLib.MSFlexGrid flxOrd 
      Height          =   1095
      Left            =   240
      TabIndex        =   31
      Top             =   9600
      Visible         =   0   'False
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   1931
      _Version        =   393216
      Rows            =   1
      Cols            =   3
      FixedCols       =   0
      FormatString    =   "Date   |<PO Number    |<Orders"
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxSP 
      Height          =   3015
      Left            =   7800
      TabIndex        =   32
      Top             =   9480
      Visible         =   0   'False
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   5318
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label lblRow 
      Height          =   255
      Left            =   10320
      TabIndex        =   6
      Top             =   120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblP 
      Caption         =   "è"
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
      Left            =   240
      TabIndex        =   5
      Top             =   8280
      Width           =   255
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   480
      Top             =   7680
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   13170
      m_MDI_Height    =   9660
   End
   Begin VB.Menu mnuOrder 
      Caption         =   "Sales Report"
      Visible         =   0   'False
      Begin VB.Menu mnuCancel 
         Caption         =   "Cancel P.O."
      End
   End
End
Attribute VB_Name = "SPorderSum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsRPT As New ADODB.Recordset
Dim rsSupplier As New ADODB.Recordset
Dim rsSP As New ADODB.Recordset
Dim BtnIndx As Integer, OSrow As Integer, fRow As Integer, fCol As Integer
Dim Qry As String, dFrm As Date, dTo As Date
Dim SuppSrch As Boolean

Private Sub cboPO_Change()
    Dim rsPO As New ADODB.Recordset
    Dim Qry As String, i As Integer, TtlRec As Long, TtlOrd As Long
    
    
    If cboPO.Text <> "" Then
    
        With flxOrder
            For i = 1 To .Rows - 1
                If .TextMatrix(i, 3) = cboPO.Text Then
                    txtOrder.Text = Format(.TextMatrix(i, 10), "#,##0")
                    txtRec.Text = Format(.TextMatrix(i, 11), "#,##0")
                    TtlOrd = PureNum(.TextMatrix(i, 10))
                    TtlRec = PureNum(.TextMatrix(i, 11))
                    txtBackOrder.Text = Format(TtlOrd - TtlRec, "#,##0")
                    If TtlOrd > 0 And TtlRec > 0 Then
                        lblPctRec.Caption = Format(TtlRec / TtlOrd, "#0.00%")
                        lblPctBO.Caption = Format((TtlOrd - TtlRec) / TtlOrd, "#0.00%")
                    Else
                        lblPctRec.Caption = ""
                        lblPctBO.Caption = ""
                    End If
                    txtRemarks.Text = "Updated " & .TextMatrix(i, 13)
                    .Col = 1: .Row = i: .ColSel = 13: .RowSel = i
                    Exit Sub
                End If
            Next i
        End With
        MsgBox "P.O. Number not found!", vbExclamation + vbOKOnly, "Back Order"
        txtOrder.Text = ""
        txtRec.Text = ""
        lblPctRec.Caption = ""
        txtBackOrder.Text = ""
        lblPctBO.Caption = ""
        txtRemarks.Text = ""
    End If
    
    
End Sub

Private Sub cboPO_Click()
    Call cboPO_Change
End Sub

Private Sub cboSupplier_Change()
    Dim i As Long, j As Long
    Dim strPartial As String, strTotal As String
    Dim X As Integer
    
    If SuppSrch Then
        SuppSrch = False
        Exit Sub
    End If
    With cboSupplier
        strPartial = .Text
        i = SendMessage(.hwnd, CB_FINDSTRING, -1, ByVal strPartial) 'If match found, append unmatched characters
        If i <> -1 Then
            strTotal = .List(i)
            rsSupplier.MoveFirst
            rsSupplier.Move i
            txtSuppID.Text = rsSupplier!SUPPLIERID
            j = Len(strTotal) - Len(strPartial)
            If j <> 0 Then
                SuppSrch = True
                .SelText = Right$(strTotal, j)
                .SelStart = Len(strPartial)
                .SelLength = j
            End If
        Else
            txtSuppID.Text = ""
        End If
    End With
End Sub

Private Sub cboSupplier_Click()
    Dim rsOrder As New ADODB.Recordset
    Dim Qry As String
    
    Call cboSupplier_Change
    If txtSuppID.Text <> "" Then
        'Display P.O. of Suppliers
        '***Access
        'Qry = "SELECT SPORDER.ORDERDATE, SPORDER.REFERENCE, SUPPLIER.SUPPLIERNAME, SPORDER.ORDERTYPE, SPORDER.CONFIRMNO, SPORDER.SUBMITTEDBY, SPORDER.APPROVEDBY, SPORDER.PRINTCOPY, SPORDER.QTYORDER, SPORDER.QTYREC, SPORDER.BACKORDER, SPORDER.BOLASTUPDATE " & _
              "FROM SPORDER INNER JOIN SUPPLIER ON SPORDER.SUPPLIERID = SUPPLIER.SUPPLIERID Where (((SUPPLIER.SUPPLIERID) = " & txtSuppID.Text & ")) ORDER BY SPORDER.ORDERDATE DESC;"
        'mySQL
        Qry = "SELECT SPORDER.ORDERDATE, SPORDER.REFERENCE, SUPPLIER.SUPPLIERNAME, SPORDER.ORDERTYPE, SPORDER.CONFIRMNO, SPORDER.SUBMITTEDBY, SPORDER.APPROVEDBY, SPORDER.PRINTCOPY, SPORDER.QTYORDER, SPORDER.QTYREC, SPORDER.BACKORDER, SPORDER.BOLASTUPDATE " & _
              "FROM SPORDER INNER JOIN SUPPLIER ON SPORDER.SUPPLIERID = SUPPLIER.SUPPLIERID Where SUPPLIER.SUPPLIERID = " & txtSuppID.Text & " ORDER BY SPORDER.ORDERDATE DESC;"
        
        With rsOrder
            .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
            If .RecordCount > 0 Then
                .MoveFirst
                flxOrder.Rows = 1
                cboPO.Clear
                TtlOrd = 0
                TtlRec = 0
                Do While Not .EOF
                    flxOrder.AddItem vbTab & vbTab & Format$(!ORDERDATE, "MM/DD/YYYY") & vbTab & !REFERENCE & vbTab & !SUPPLIERNAME & vbTab & !ORDERTYPE & vbTab & !CONFIRMNO & vbTab & !SubmittedBy & vbTab & !ApprovedBy & vbTab & !PRINTCOPY & vbTab & IIf(IsNull(!QtyOrder), 0, !QtyOrder) & vbTab & IIf(IsNull(!QtyRec), 0, Format(!QtyRec, "#,##0")) & vbTab & IIf(IsNull(!BACKORDER), 0, Format(!BACKORDER, "#,##0")) & vbTab & IIf(IsNull(!BOLastUpdate), "", !BOLastUpdate)
                    cboPO.AddItem !REFERENCE
                    TtlOrd = TtlOrd + IIf(IsNull(!QtyOrder), 0, !QtyOrder)
                    TtlRec = TtlRec + IIf(IsNull(!QtyRec), 0, !QtyRec)
                    .MoveNext
                Loop
            Else
                flxOrder.Rows = 1
                flxOrder.Rows = 2
                TtlOrd = 0
                TtlRec = 0
            End If
            .Close
        End With
        txtOrder.Text = Format(TtlOrd, "#,##0")
        txtRec.Text = Format(TtlRec, "#,##0")
        txtBackOrder.Text = Format(TtlOrd - TtlRec, "#,##0")
        If TtlOrd > 0 And TtlRec > 0 Then
            lblPctRec.Caption = Format(TtlRec / TtlOrd, "#0.00%")
            lblPctBO.Caption = Format((TtlOrd - TtlRec) / TtlOrd, "#0.00%")
        Else
            lblPctRec.Caption = ""
            lblPctBO.Caption = ""
        End If
        With flxOrder
            For i = 1 To .Rows - 1
                If Val(.TextMatrix(i, 9)) = 0 Then
                    '.Col = 0: .Row = i: .CellFontName = "Webdings": .CellFontSize = 10: .TextMatrix(i, 0) = lblP.Caption
                    .Col = 1: .Row = i: .ColSel = 12: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 250, 180)
                    .Col = 2
                End If
            Next i
        End With
        
        
    End If
End Sub

Private Sub cboSupplier_GotFocus()
   Const CB_SHOWDROPDOWN = &H14F
   Dim Tmp
   Tmp = SendMessage(cboSupplier.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub

Private Sub cboSupplier_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then SuppSrch = True
End Sub

Private Sub cboSupplier_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then Call cboSupplier_Click
End Sub

Private Sub cmdFind_Click()
    Dim i As Integer, j As Integer
        
    If txtFind.Text = "" Then Exit Sub
    With flxOrder
        For i = fRow To .Rows - 1
            For j = fCol To 13
                If InStr(.TextMatrix(i, j), txtFind.Text) > 0 Then
                    .Col = 1: .Row = i
                    .ColSel = 13:  .RowSel = i
                    .TopRow = i
                    fRow = IIf(j = 13, i + 1, i)
                    fCol = IIf(j = 13, 1, j + 1)
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
    
    
Private Sub cmdTrans_Click(Index As Integer)
    Dim MyQry As String
    
    Select Case Index
        Case 0  'Grid View
            BtnIndx = 1
            tmrFloat.Enabled = True
            Call PrintPreView(False)
           ' Call cboMonth_Click
        Case 1  'Exit
            ''MotoRep.imgBG.Visible = True
            MotoRep.toolBarButton.Visible = True
            Unload Me
        Case 2  'SP P.O.
            With flxOrder
                If .TextMatrix(OSrow, 3) <> "" Then
                    If OSrow = 0 And .Rows > 2 Then OSrow = 1
                    SPordering.Show
                    Call SPordering.DisplaySPOrder(.TextMatrix(OSrow, 3))
                End If
            End With
            Unload Me
        Case 3
            With flxOrder
                If .TextMatrix(OSrow, 3) <> "" Then
                    If OSrow = 0 And .Rows > 2 Then OSrow = 1
                    SPBackOrder.Show
                    Call SPBackOrder.DisplaySPOrder(.TextMatrix(OSrow, 3))
                End If
            End With
            Unload Me
        Case 4  'Refresh
            Call Form_Load
        Case 5  'Print Preview
'            BtnIndx = 4
'            tmrFloatRev.Enabled = True
'            Call PrintPreView(True)
'            Call DisplayReport
        Case 6  'Exit
            'MotoRep.tbMain.Visible = True
            ''MotoRep.imgBG.Visible = True
            MotoRep.toolBarButton.Visible = True
            Unload Me
    End Select
End Sub

Private Sub flxOrder_Click()
    OSrow = flxOrder.Row
End Sub

Private Sub flxOrder_EnterCell()
    OSrow = flxOrder.Row
End Sub

Private Sub flxORDER_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
    If InStr(flxOrder.TextMatrix(OSrow, 6), "CANCELLED") > 0 Then Exit Sub
    If Button = vbRightButton Then
        Call PopupMenu(mnuOrder)
    End If
End Sub

Private Sub Form_Load()
    Dim rsOrder As New ADODB.Recordset
    Dim Qry As String
    Dim i As Integer, TtlOrd As Long, TtlRec As Long

    Qry = "SELECT SPORDER.ORDERDATE, SPORDER.REFERENCE, SUPPLIER.SUPPLIERNAME, SPORDER.ORDERTYPE, SPORDER.CONFIRMNO, SPORDER.SUBMITTEDBY, SPORDER.APPROVEDBY, SPORDER.PRINTCOPY, SPORDER.QTYORDER, SPORDER.QTYREC, SPORDER.BACKORDER, SPORDER.BOLASTUPDATE " & _
          "FROM SPORDER INNER JOIN SUPPLIER ON SPORDER.SUPPLIERID = SUPPLIER.SUPPLIERID ORDER BY SPORDER.ORDERDATE DESC, SPORDER.LASTUPDATE DESC;"
    
    With rsOrder
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            flxOrder.Rows = 1
            TtlOrd = 0
            TtlRec = 0
            Do While Not .EOF
                flxOrder.AddItem vbTab & vbTab & Format$(!ORDERDATE, "MM/DD/YYYY") & vbTab & !REFERENCE & vbTab & !SUPPLIERNAME & vbTab & !ORDERTYPE & vbTab & !CONFIRMNO & vbTab & !SubmittedBy & vbTab & !ApprovedBy & vbTab & !PRINTCOPY & vbTab & IIf(IsNull(!QtyOrder), 0, !QtyOrder) & vbTab & IIf(IsNull(!QtyRec), 0, Format(!QtyRec, "#,##0")) & vbTab & IIf(IsNull(!BACKORDER), 0, Format(!BACKORDER, "#,##0")) & vbTab & IIf(IsNull(!BOLastUpdate), "", !BOLastUpdate)
                TtlOrd = TtlOrd + IIf(IsNull(!QtyOrder), 0, !QtyOrder)
                TtlRec = TtlRec + IIf(IsNull(!QtyRec), 0, !QtyRec)
                .MoveNext
            Loop
        End If
        .Close
    End With
    txtOrder.Text = Format(TtlOrd, "#,##0")
    txtRec.Text = Format(TtlRec, "#,##0")
    txtBackOrder.Text = Format(TtlOrd - TtlRec, "#,##0")
    If TtlRec > 0 And TtlOrd > 0 Then
        lblPctRec.Caption = Format(TtlRec / TtlOrd, "#0.00%")
        lblPctBO.Caption = Format((TtlOrd - TtlRec) / TtlOrd, "#0.00%")
    Else
        lblPctRec.Caption = ""
        lblPctBO.Caption = ""
    End If
    
    With flxOrder
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, 9)) = 0 Then
                '.Col = 0: .Row = i: .CellFontName = "Webdings": .CellFontSize = 10: .TextMatrix(i, 0) = lblP.Caption
                .Col = 1: .Row = i: .ColSel = 12: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 250, 180)
                .Col = 2
            End If
            If .TextMatrix(i, 6) = "CANCELLED" Then 'REMARKS
                .Col = 1: .Row = i: .ColSel = 12: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(240, 0, 0)
                .Col = 2
            End If
        Next i
    End With
    
    With rsSupplier
        If .State = adStateOpen Then .Close
        Qry = "SELECT SUPPLIER.SUPPLIERNAME, SUPPLIER.SUPPLIERID " & _
              "From SUPPLIER ORDER BY SUPPLIER.SUPPLIERNAME;"
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            cboSupplier.Clear
            Do While Not .EOF
                cboSupplier.AddItem !SUPPLIERNAME
                .MoveNext
            Loop
        End If
    End With
    cboPO.Clear
    BtnIndx = 1
    tmrFloat.Enabled = True
    MotoRep.sBar.Panels(6).Text = "Spare Parts Order Summary"
End Sub
Private Sub PrintPreView(Ppvw As Boolean)
    If Ppvw Then
'        Frame1.Visible = False
        flxOrder.Visible = False
        vwrDmCm.Visible = True
    Else
        'vwrDmCm.Visible = False
 '       Frame1.Visible = True
        flxOrder.Visible = True
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If rsRPT.State = adStateOpen Then rsRPT.Close
    If rsSupplier.State = adStateOpen Then rsSupplier.Close
End Sub

Private Sub mnuCancel_Click()
    Dim Ans As Integer, MyBO As Integer
    Dim Qry As String, MyAcc As String
    
    '***Access
    'Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE (((USERACCESS.EMPID)=" & UserID & ") AND ((USERACCESS.MENUINDX)=19));"
    'mySQL
    Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE USERACCESS.EMPID=" & UserID & " AND USERACCESS.MENUINDX=19;"
    MyAcc = GetFieldValue(Qry)
    If MyAcc <> "FULL ACCESS" Then
        MsgBox "You don't have full access to this transaction." & vbCrLf & "Please contact your System Administrator", vbExclamation + vbOKOnly, "Access Denied"
        Exit Sub
    End If
    
    Ans = MsgBox("Do you want to Cancel P.O. " & flxOrder.TextMatrix(OSrow, 3) & "?", vbQuestion + vbYesNo, "Cancel")
    If Ans = vbNo Then Exit Sub
        
    'UPDATE ORDER FIELDS TO CANCELLED
    '***Access
    'LYRcon.Execute "UPDATE SPORDER SET SPORDER.CONFIRMNO = 'CANCELLED', SPORDER.REMARKS = 'CANCELLED' WHERE (((SPORDER.REFERENCE)='" & flxOrder.TextMatrix(OSrow, 3) & "'));"
    'LYRcon.Execute "UPDATE SPITR SET SPITR.TRANSTYPE = 'CANCELLED' WHERE (((SPITR.REFERENCE)='" & flxOrder.TextMatrix(OSrow, 3) & "'));"
    
    'mySQL
    LYRcon.Execute "UPDATE SPORDER SET SPORDER.CONFIRMNO = 'CANCELLED', SPORDER.REMARKS = 'CANCELLED' WHERE SPORDER.REFERENCE='" & flxOrder.TextMatrix(OSrow, 3) & "';"
    LYRcon.Execute "UPDATE SPITR SET SPITR.TRANSTYPE = 'CANCELLED' WHERE SPITR.REFERENCE='" & flxOrder.TextMatrix(OSrow, 3) & "';"
    
    'Update BackOrder of items in P.O.
    '***Access
    'Qry = "SELECT SPITR.SPAREPARTID, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPITR.ORDER " & _
          "FROM SPITR INNER JOIN SPAREPARTS ON SPITR.SPAREPARTID = SPAREPARTS.SPAREPARTID WHERE (((SPITR.REFERENCE)='" & flxOrder.TextMatrix(OSrow, 3) & "'));"
    'mySQL
    Qry = "SELECT SPITR.SPAREPARTID, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPITR.`ORDER` " & _
          "FROM SPITR INNER JOIN SPAREPARTS ON SPITR.SPAREPARTID = SPAREPARTS.SPAREPARTID WHERE SPITR.REFERENCE='" & flxOrder.TextMatrix(OSrow, 3) & "';"
    
    If rsSP.State = adStateOpen Then rsSP.Close
    rsSP.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    Set flxSP.DataSource = rsSP
    flxSP.Refresh
    tabCancel.Top = flxOrder.Top + 800
    tabCancel.Left = (Me.ScaleWidth - tabCancel.Width) / 2
    tabCancel.Visible = True
    tabCancel.ZOrder
    Frame2.Enabled = False
    For i = 0 To 6
        cmdTrans(i).Enabled = False
    Next i
    cmdTabCancel.Enabled = False
    With flxSP
        pBar.Value = 0
        pBar.Max = .Rows - 1
        For i = 1 To .Rows - 1
            'Start Progress bar
            MyBO = GetBackOrder(.TextMatrix(i, 1))
            '***Access
            'LYRcon.Execute "UPDATE SPAREPARTS SET SPAREPARTS.BACKORDER = " & MyBO & _
                           " WHERE (((SPAREPARTS.SPAREPARTID)=" & .TextMatrix(i, 1) & "));"
            'mySQL
            LYRcon.Execute "UPDATE SPAREPARTS SET SPAREPARTS.BACKORDER = " & MyBO & _
                           " WHERE SPAREPARTS.SPAREPARTID=" & .TextMatrix(i, 1) & ";"
            pBar.Value = pBar.Value + 1
            DoEvents
            pPct = (pBar.Value / pBar.Max)
            DoEvents
        Next i
        DoEvents
    End With
    Frame2.Enabled = True
    For i = 0 To 6
        cmdTrans(i).Enabled = True
    Next i
    tabCancel.Visible = False
    With flxOrder
        .TextMatrix(OSrow, 6) = "CANCELLED"
        .Col = 1: .Row = OSrow: .ColSel = 11: .RowSel = OSrow: .FillStyle = flexFillRepeat: .CellBackColor = RGB(240, 0, 0)
        .Col = 2
    End With
    
    MsgBox "P.O. " & flxOrder.TextMatrix(OSrow, 3) & " Cancelled!", vbInformation + vbOKOnly
End Sub

Function GetBackOrder(MyID As Integer) As Integer
    Dim rsOrder As New ADODB.Recordset
    Dim rsRec As New ADODB.Recordset
    Dim Qry As String, MyPO As String
    Dim xBal As Integer, i As Integer
    
    'Clear FlxStock
    '***Access
    'Qry = "SELECT SPITR.TRANSDATE, SPITR.REFERENCE, SPITR.ORDER From SPITR " & _
          "WHERE (((SPITR.TRANSTYPE)='ORDER') AND ((SPITR.SPAREPARTID)=" & MyId & ")) ORDER BY SPITR.TRANSDATE;"
    'mySQL
    Qry = "SELECT SPITR.TRANSDATE, SPITR.REFERENCE, SPITR.`ORDER` From SPITR " & _
          "WHERE SPITR.TRANSTYPE='ORDER' AND SPITR.SPAREPARTID=" & MyID & " ORDER BY SPITR.TRANSDATE;"
    
    flxOrd.Rows = 1
    With rsOrder
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                flxOrd.AddItem Format(!TRANSDATE, "mm/dd/yyyy") & vbTab & !REFERENCE & vbTab & !Order
                .MoveNext
            Loop
        End If
        .Close
    End With
    xBal = 0
    With flxOrd
        For i = 1 To .Rows - 1
            'Display PO in Stock Card and obtain Balance
            xBal = xBal + Val(.TextMatrix(i, 2))
            MyPO = .TextMatrix(i, 1)
            
            '***Access
            'Qry = "SELECT SPITR.TRANSDATE, SPITR.REFERENCE, SPITR.RECEIVE, SPITR.PACKINGLIST From SPITR " & _
                  "WHERE (((SPITR.TRANSTYPE)='RECEIVE') AND ((SPITR.SPAREPARTID)=" & MyId & ") AND ((SPITR.PACKINGLIST)='" & MyPO & "')) ORDER BY SPITR.TRANSDATE;"
            'mySQL
            Qry = "SELECT SPITR.TRANSDATE, SPITR.REFERENCE, SPITR.RECEIVE, SPITR.PACKINGLIST From SPITR " & _
                  "WHERE SPITR.TRANSTYPE='RECEIVE' AND SPITR.SPAREPARTID=" & MyID & " AND SPITR.PACKINGLIST='" & MyPO & "' ORDER BY SPITR.TRANSDATE;"
            
            rsRec.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
            If rsRec.RecordCount > 0 Then
                rsRec.MoveFirst
                Do While Not rsRec.EOF
                    xBal = xBal - rsRec!RECEIVE
                    rsRec.MoveNext
                Loop
            End If
            rsRec.Close
        Next i
    End With
    GetBackOrder = xBal
End Function


Private Sub RXResizer1_ResizeComplete()
    Dim i As Integer
    With flxOrder
        .ColWidth(0) = 300
        .ColWidth(1) = 0
        .ColWidth(2) = (.Width - 600) * 0.12
        .ColWidth(3) = (.Width - 600) * 0.12
        .ColWidth(4) = (.Width - 600) * 0.2
        .ColWidth(5) = (.Width - 600) * 0.12
        .ColWidth(6) = (.Width - 600) * 0.12
        .ColWidth(7) = (.Width - 600) * 0.16
        .ColWidth(8) = (.Width - 600) * 0.16
        .ColWidth(9) = (.Width - 600) * 0.16
        .ColWidth(10) = (.Width - 600) * 0.16
        .ColWidth(11) = (.Width - 600) * 0.16
    End With
End Sub

Private Sub tmrDrop_Timer()
    cmdTrans(BtnIndx).Top = cmdTrans(BtnIndx).Top + 50
    If cmdTrans(BtnIndx).Top >= 0 Then
        cmdTrans(BtnIndx).Top = 0
        BtnIndx = BtnIndx + 1
        If BtnIndx = 7 Then     'last index + 1
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
            For i = 2 To 6
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


