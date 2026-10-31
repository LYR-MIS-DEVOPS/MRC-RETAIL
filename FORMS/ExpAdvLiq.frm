VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form ExpAdvLiq 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Advances for Liquidation"
   ClientHeight    =   5700
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   8760
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   8760
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Height          =   2175
      Left            =   120
      TabIndex        =   27
      Top             =   1800
      Width           =   8535
      Begin VB.TextBox txtAFLamt 
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
         Left            =   3960
         MaxLength       =   50
         TabIndex        =   35
         Tag             =   "XYF"
         Top             =   1680
         Width           =   1695
      End
      Begin VB.TextBox txtAmtRet 
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
         Left            =   3960
         MaxLength       =   50
         TabIndex        =   32
         Tag             =   "XYF"
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox txtAmtRec 
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
         Left            =   3960
         MaxLength       =   50
         TabIndex        =   29
         Tag             =   "XYF"
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "AFL AMOUNT:"
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
         Left            =   2280
         TabIndex        =   34
         Tag             =   "XYF"
         Top             =   1710
         Width           =   1575
      End
      Begin VB.Label lblTimeRet 
         Caption         =   "Time Returned:"
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
         Left            =   3960
         TabIndex        =   33
         Tag             =   "XYF"
         Top             =   1320
         Width           =   2295
      End
      Begin VB.Label lblAmtRet 
         Alignment       =   1  'Right Justify
         Caption         =   "AMOUNT RETURNED:"
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
         Left            =   2280
         TabIndex        =   31
         Tag             =   "XYF"
         Top             =   990
         Width           =   1575
      End
      Begin VB.Label lblTimeRec 
         Caption         =   "Time Received:"
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
         Left            =   3960
         TabIndex        =   30
         Tag             =   "XYF"
         Top             =   720
         Width           =   2295
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "AMOUNT RECEIVED:"
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
         Left            =   2280
         TabIndex        =   28
         Tag             =   "XYF"
         Top             =   390
         Width           =   1575
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   240
      Top             =   6120
   End
   Begin VB.ComboBox cboCashier 
      Height          =   315
      Left            =   1200
      TabIndex        =   26
      Top             =   4905
      Width           =   2775
   End
   Begin MSFlexGridLib.MSFlexGrid flxPrn 
      Height          =   1095
      Left            =   6720
      TabIndex        =   24
      Top             =   9120
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   1931
      _Version        =   393216
      Cols            =   4
   End
   Begin VB.TextBox txtAFLid 
      Height          =   285
      Left            =   7080
      TabIndex        =   23
      Top             =   7080
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtPrnCopy 
      Height          =   285
      Left            =   6480
      TabIndex        =   22
      Top             =   7080
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtAprvBy 
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
      Left            =   5880
      MaxLength       =   50
      TabIndex        =   21
      Tag             =   "XYF"
      Top             =   4905
      Width           =   2775
   End
   Begin VB.Timer tmrFloat 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1200
      Top             =   9480
   End
   Begin VB.Timer tmrDrop 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1200
      Top             =   9120
   End
   Begin VB.Timer tmrFloatRev 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1800
      Top             =   9480
   End
   Begin VB.Timer tmrDropRev 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1800
      Top             =   9120
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4560
      Top             =   9120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ExpAdvLiq.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ExpAdvLiq.frx":0EDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ExpAdvLiq.frx":1DB4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ExpAdvLiq.frx":2C8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ExpAdvLiq.frx":3B68
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ExpAdvLiq.frx":4A42
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtRemarks 
      BackColor       =   &H00FFFFFF&
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
      Left            =   1200
      Locked          =   -1  'True
      MaxLength       =   250
      TabIndex        =   0
      Tag             =   "XYF"
      Top             =   4080
      Width           =   7455
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   120
      TabIndex        =   1
      Tag             =   "XYF"
      Top             =   720
      Width           =   8535
      Begin VB.TextBox txtRecBy 
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
         Left            =   1320
         MaxLength       =   50
         TabIndex        =   25
         Tag             =   "XYF"
         Top             =   240
         Width           =   3375
      End
      Begin VB.TextBox txtAFLDate 
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
         Left            =   6360
         TabIndex        =   20
         Tag             =   "XYF"
         Top             =   604
         Width           =   2055
      End
      Begin VB.TextBox txtAFLno 
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
         Left            =   6360
         MaxLength       =   10
         TabIndex        =   19
         Tag             =   "XYF"
         Top             =   248
         Width           =   2055
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "DISBURSER"
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
         Left            =   1320
         TabIndex        =   36
         Tag             =   "XYF"
         Top             =   600
         Width           =   3375
      End
      Begin VB.Label Label1 
         Caption         =   "AFL. DATE:"
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
         Left            =   5280
         TabIndex        =   18
         Tag             =   "XYF"
         Top             =   630
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "AFL. NO."
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
         Left            =   5280
         TabIndex        =   17
         Tag             =   "XYF"
         Top             =   285
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "RECEIVED BY:"
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
         Left            =   120
         TabIndex        =   2
         Tag             =   "XYF"
         Top             =   270
         Width           =   1455
      End
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
      Caption         =   "Add &New"
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
      Image           =   "ExpAdvLiq.frx":591C
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   2
      Left            =   3360
      TabIndex        =   8
      Top             =   0
      Width           =   1680
      _ExtentX        =   2963
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
      Image           =   "ExpAdvLiq.frx":5EB6
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   3
      Left            =   5040
      TabIndex        =   9
      Top             =   0
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
      Caption         =   "&Print Form"
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
      Image           =   "ExpAdvLiq.frx":6450
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   4
      Left            =   6720
      TabIndex        =   10
      Top             =   0
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
      Caption         =   "E&xit Form"
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
      Image           =   "ExpAdvLiq.frx":69EA
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   5
      Left            =   8400
      TabIndex        =   11
      Top             =   0
      Visible         =   0   'False
      Width           =   1680
      _ExtentX        =   2963
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
      Image           =   "ExpAdvLiq.frx":6F84
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   6
      Left            =   10080
      TabIndex        =   12
      Top             =   0
      Visible         =   0   'False
      Width           =   1680
      _ExtentX        =   2963
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
      Image           =   "ExpAdvLiq.frx":751E
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   1
      Left            =   1680
      TabIndex        =   16
      Top             =   0
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
      Caption         =   "&Return Excess"
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
      Image           =   "ExpAdvLiq.frx":7AB8
      cBack           =   -2147483633
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "APPROVED BY"
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
      Left            =   5880
      TabIndex        =   15
      Tag             =   "XYF"
      Top             =   5280
      Width           =   2775
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "PROCESSED BY"
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
      Left            =   1200
      TabIndex        =   14
      Tag             =   "XYF"
      Top             =   5280
      Width           =   2775
   End
   Begin VB.Label Label2 
      Caption         =   "REMARKS:"
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
      Left            =   120
      TabIndex        =   13
      Tag             =   "XYF"
      Top             =   4110
      Width           =   1455
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   4320
      Top             =   6840
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   8850
      m_MDI_Height    =   6120
   End
   Begin VB.Label lblHeading 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ADVANCES FOR LIQUIDATION"
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
      Left            =   5160
      TabIndex        =   5
      Top             =   450
      Width           =   3420
   End
   Begin VB.Label lblHeading 
      BackStyle       =   0  'Transparent
      Caption         =   "LYRGCOM"
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
      TabIndex        =   3
      Top             =   480
      Width           =   2415
   End
   Begin VB.Label lblHeading 
      BackStyle       =   0  'Transparent
      Caption         =   "LYRGCOM"
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
      TabIndex        =   4
      Top             =   480
      Width           =   2415
   End
   Begin VB.Label lblHeading 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ADVANCES FOR LIQUIDATION"
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
      Left            =   5160
      TabIndex        =   6
      Top             =   480
      Width           =   3420
   End
   Begin VB.Menu mnuFlx 
      Caption         =   "FlxMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete Cash Voucher Details"
      End
   End
End
Attribute VB_Name = "ExpAdvLiq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim BtnIndx As Integer, AddNew As Boolean
Public SrcSumm As Boolean

Sub ClearEntries()
    
    txtRecBy.Text = ""
    txtAFLno.Text = ""
    txtAFLDate.Text = ""
    txtAmtRet.Text = ""
    txtAmtRec.Text = ""
    txtAFLamt.Text = ""
    'cboCashier.Clear
    txtRemarks.Text = ""
    'cboCashier.Text = ""
    txtAprvby.Text = ""
    txtPrnCopy.Text = ""
    txtAFLid.Text = ""
End Sub

Sub LockData(bln As Boolean)
    Dim i As Integer
    
    Frame1.Enabled = Not bln
    Frame3.Enabled = Not bln
    cbocashier.Locked = bln
    txtAprvby.Locked = bln
    txtRemarks.Locked = bln
End Sub

Public Sub ShowRecord()
    Dim rsAdvLiqurn As New ADODB.Recordset
    Dim Qry As String, i As Integer, MyID As Integer
    
    'Call ClearEntries
    AddNew = False
    Qry = "SELECT advliquidation.* From advliquidation Where advliquidation.AFLID = " & Val(txtAFLid.Text) & ";"
    With rsAdvLiqurn
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            txtAFLno.Text = !AFLNO
            txtAFLDate.Text = Format(!AFLDATE, "mm/dd/yyyy")
            txtRecBy.Text = !RECBY
            txtAmtRec.Text = Format(!AMTREC, "###0.00")
            lblTimeRec.Caption = "Time Received: " & Format(!TIMEREC, "hh:mm ampm")
            If Not IsNull(!TIMERETURN) Then lblTimeRet.Caption = "Time Returned: " & Format(!TIMERETURN, "hh:mm ampm")
            txtAmtRet.Text = Format(!AMTRETURN, "###0.00")
            txtAFLamt.Text = Format(!AFLAMOUNT, "###0.00")
            txtRemarks.Text = IIf(IsNull(!REMARKS), "", !REMARKS)
            cbocashier.Text = !PREPAREDBY
            txtAprvby.Text = !APPROVEDBY
            txtAFLid.Text = !AFLID
            txtPrnCopy.Text = !PRINTCOPY
            For i = 0 To cbocashier.ListCount - 1
                If cbocashier.List(i) = !PREPAREDBY Then
                    cbocashier.ListIndex = i
                    cbocashier.ItemData(i) = !UPDATEDBY
                End If
            Next i
        End If
        .Close
    End With
End Sub

Private Sub cboCashier_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtAprvby.SetFocus
End Sub


Private Sub cmdTrans_Click(Index As Integer)
    Dim Ans As Integer, i As Integer, RetID As Long, MyAFLno As Long
    Dim myMsg As String, Qry As String, MyAcc As String, PrnVal(0 To 13) As String
    Dim flx As MSFlexGrid
    
    Select Case Index
        Case 0      'addnew
            Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE (EMPID=" & UserID & " AND MENUINDX=" & 98 & ");"
            MyAcc = GetFieldValue(Qry)
            If MyAcc <> "FULL ACCESS" Then
                MsgBox "You don't have full access to this transaction." & vbCrLf & "Please contact your System Administrator", vbExclamation + vbOKOnly, "Access Denied"
                Exit Sub
            End If
            BtnIndx = 4
            tmrFloat.Enabled = True
            Call ClearEntries
            Call LockData(False)
            'Qry = "SELECT transignatory.EMPLOYEE From transignatory Where transignatory.TRANSACTION = 'Advances for Liquidation (1)';"
            'txtRecBy.Text = GetFieldValue(Qry)
            txtRecBy.Text = GetSettings("DISBURSER")
            Qry = "SELECT max(round(AFLNo)) As MyAFL From advliquidation;"
            MyAFLno = Val(GetFieldValue(Qry))
            txtAFLno.Text = IIf(Len(str(MyAFLno)) > 3, MyAFLno, Format(MyAFLno + 1, "0000"))
            'Qry = "SELECT TRANSIGNATORY.EMPLOYEE From TRANSIGNATORY WHERE (TRANSACTION='Advances for Liquidation (2)');"
            txtAprvby.Text = GetSettings("MANAGER")
            txtAFLDate.Text = Format(Date, "mm/dd/yyyy")
            txtAmtRet.Enabled = False
            txtAFLamt.Enabled = False
            txtAmtRec.SetFocus
            AddNew = True
        Case 1      'Return Excess of Advances
            Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE (EMPID=" & UserID & " AND MENUINDX=" & 98 & ");"
            MyAcc = GetFieldValue(Qry)
            If MyAcc <> "FULL ACCESS" Then
                MsgBox "You don't have full access to this transaction." & vbCrLf & "Please contact your System Administrator", vbExclamation + vbOKOnly, "Access Denied"
                Exit Sub
            End If
            If Val(txtAFLid.Text) = 0 Then
                MsgBox "No available Advances for Liquidition!", vbExclamation + vbOKOnly, "Report Empty"
                Exit Sub
            End If
            Qry = "SELECT cashflow.CASHFLOWID From cashflow Where cashflow.CFDATE = '" & Format(DateValue(txtAFLDate.Text), "YYYY-MM-DD") & "'  AND cashflow.UPDATEDBY = " & cbocashier.ItemData(cbocashier.ListIndex) & ";"
            If CtrlNumExist(Qry) Then
                MsgBox "Return of Excess is not allowed becase a Cashflow already exist!", vbExclamation + vbOKOnly, "Validation Error"
                Exit Sub
            End If
            BtnIndx = 4
            tmrFloat.Enabled = True
            'Call LockData(False)
            Frame3.Enabled = True
            txtAmtRec.Locked = True
            txtAmtRet.Locked = False
            txtAFLamt.Locked = True
            AddNew = False
            txtAmtRet.SelStart = 0
            txtAmtRet.SelLength = Len(txtAmtRet.Text)
            'txtAmtRet.SetFocus
        Case 5      'save
            myMsg = ValidateEntry
            If myMsg <> "" Then
                MsgBox myMsg, vbOKOnly + vbCritical, "Validation Error"
                Exit Sub
            End If
            Ans = MsgBox("Do you want to save entry?", vbYesNo + vbQuestion, "SAVE")
            If Ans = vbNo Then Exit Sub
            'save here
            
            Call SaveEntries
            BtnIndx = 6
            tmrFloatRev.Enabled = True
            
            MsgBox "Entry successfully saved!", vbOKOnly + vbInformation, "SAVE"
            Call LockData(True)
        Case 6      'cancel
            Ans = MsgBox("Are you sure you want to cancel?", vbQuestion + vbYesNo, "Cancel")
            If Ans = vbNo Then Exit Sub
            BtnIndx = 6
            tmrFloatRev.Enabled = True
            RetID = Val(txtAFLid.Text)
            Call ClearEntries
            txtAFLid.Text = RetID
            If Not AddNew Then Call ShowRecord
            Call LockData(True)
        
        Case 2      'View Sumary
            Unload Me
            FrmSummary.mySrc = "Advances for Liquidation"
            If SrcSumm Then
                'MotoRep.tbMain.Visible = False
                MotoRep.imgBG.Visible = False
                MotoRep.toolBarButton.Visible = False
                Call FrmSummary.RefreshRecord(Val(FrmSummary.cmdNext.Tag))
            Else
                FrmSummary.Show
                'MotoRep.tbMain.Visible = False
                MotoRep.imgBG.Visible = False
            MotoRep.toolBarButton.Visible = False
            End If
        Case 3      '"print"
            If Val(txtAFLid.Text) = 0 Then
                MsgBox "No Advances for Liquidition to print!", vbExclamation + vbOKOnly, "Report Empty"
                Exit Sub
            End If
            Ans = MsgBox("Do you want to Print Advances for Liquidition?", vbQuestion + vbYesNo, "Print")
            If Ans = vbNo Then Exit Sub
            Qry = "SELECT branch.BRANCHNAME From branch Where branch.SYSTEMBRANCH = 1"
            PrnVal(0) = GetFieldValue(Qry)
            PrnVal(1) = "Advances for Liquidation" '
            PrnVal(2) = txtRecBy.Text
            PrnVal(3) = txtAFLno.Text   '
            PrnVal(4) = txtAFLDate.Text '"ADVANCES FOR LIQUIDITION"
            PrnVal(5) = "AMOUNT RECEIVED: " & txtAmtRec.Text
            PrnVal(6) = lblTimeRec.Caption
            PrnVal(7) = IIf(Val(txtAmtRet.Text) > 0, "AMOUNT RETURNED: " & txtAmtRet.Text, "")
            PrnVal(8) = IIf(Val(txtAmtRet.Text) > 0, lblTimeRet.Caption, "")
            PrnVal(9) = txtAFLamt.Text
            PrnVal(10) = "Prepared By:"
            PrnVal(11) = "Approved By:"
            PrnVal(12) = cbocashier.Text
            PrnVal(13) = txtAprvby.Text
            Call PrintMyForm(PrnVal, "Advances for Liquidation", flx)
            
            
         '   Qry = "SELECT Branch.BRANCHNAME From Branch Where Branch.BRANCHID = " & Val(MotoRep.Tag) & ";"
         '   PrnVal(0) = GetFieldValue(Qry)
         '   PrnVal(1) = txtAFLno.Text   '
         '   PrnVal(2) = txtRecBy.Text
         '   PrnVal(3) = txtAFLDate.Text
         '   PrnVal(4) = "ADVANCES FOR LIQUIDITION"
         '   PrnVal(5) = txtAFLamt.Text
         '   PrnVal(6) = "AMOUNT RECEIVED: " & txtAmtRec.Text
         '   PrnVal(7) = lblTimeRec.Caption
         '   PrnVal(8) = IIf(Val(txtAmtRet.Text) > 0, "AMOUNT RETURNED: " & txtAmtRet.Text, "")
         '   PrnVal(9) = IIf(Val(txtAmtRet.Text) > 0, lblTimeRet.Caption, "")
         '   PrnVal(10) = "Prepared By:"
         '   PrnVal(11) = cboCashier.Text
         '   PrnVal(12) = txtAprvBy.Text
         '   Call PrintMyForm(PrnVal, "AFL TEMP PCV", flx)
            If Val(txtPrnCopy.Text) = 0 Then
                LYRcon.Execute "UPDATE advliquidation SET advliquidation.PRINTCOPY = 1 WHERE (AFLID=" & txtAFLid.Text & ");"
            Else
                txtPrnCopy.Text = Val(txtPrnCopy.Text) + 1
                LYRcon.Execute "UPDATE advliquidation SET advliquidation.PRINTCOPY = " & Val(txtPrnCopy.Text) & " WHERE (AFLID=" & txtAFLid.Text & ");"
            End If
        Case 4      '"quit"
            Unload Me
            If SrcSumm Then
                'MotoRep.tbMain.Visible = False
                MotoRep.imgBG.Visible = False
                MotoRep.toolBarButton.Visible = False
                Call FrmSummary.RefreshRecord(Val(FrmSummary.cmdNext.Tag))
            End If
    End Select
End Sub



Private Sub Form_Load()
    Dim rsCashier As New ADODB.Recordset
    Dim MyID As Integer
        
    Call ClearEntries
    MyID = 0
    With rsCashier
        Qry = "SELECT CONCAT(employees.LASTNAME, ', ', employees.FIRSTNAME) AS CASHIERS, Employees.EmpId " & _
              "From Employees Where employees.DEPARTMENT = 'ADMIN' Order By CASHIERS"
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            cbocashier.Clear
            Do While Not .EOF
                cbocashier.AddItem !Cashiers
                cbocashier.ItemData(cbocashier.ListCount - 1) = !EMPID
                If UserID = !EMPID Then MyID = cbocashier.ListCount - 1
                .MoveNext
            Loop
        End If
        .Close
    End With
    If cbocashier.ListCount > 0 Then cbocashier.ListIndex = MyID
    Call LockData(True)
    
    MotoRep.sBar.Panels(6).Text = "Advances for Liquidation"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    MotoRep.sBar.Panels(6).Text = ""
End Sub

Private Sub txtAmtRec_Change()
    txtAFLamt.Text = txtAmtRec.Text
End Sub

Private Sub txtAmtRec_KeyPress(KeyAscii As Integer)
    Dim strValid As String
    
    strValid = "1234567890."
    If KeyAscii > 26 Then
        If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
    If KeyAscii = 13 Then cbocashier.SetFocus
End Sub

Private Sub txtAmtRet_Change()
    txtAFLamt.Text = Val(txtAmtRec.Text) - Val(txtAmtRet.Text)
End Sub

Private Sub txtAmtRet_KeyPress(KeyAscii As Integer)
    Dim strValid As String
    
    strValid = "1234567890."
    If KeyAscii > 26 Then
        If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
    If KeyAscii = 13 Then txtAFLamt.SetFocus
End Sub

Private Sub txtAmtRet_LostFocus()
    If Val(txtAmtRet) > Val(txtAmtRec.Text) Then
        MsgBox "Invalid Amount Returned!!!", vbExclamation + vbOKOnly, "Warning"
        txtAmtRet.Text = ""
        txtAFLamt.Text = txtAmtRec.Text
    End If
End Sub

Private Sub txtaprvby_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
End Sub

Private Sub txtAFLdate_KeyPress(KeyAscii As Integer)
    Dim strValid As String
    
    strValid = "1234567890/-"
    If KeyAscii > 26 Then
        If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
    If KeyAscii = 13 Then txtAmtRec.SetFocus

End Sub

Private Sub txtAFLdate_LostFocus()
    Dim Qry As String
    
    If IsDate(txtAFLDate.Text) Then
        Qry = ""    'Validate if Cashflow for the date if already made.
        'If IsDate(cmdTrans(0).Tag) Then
        '    If DateValue(txtAFLdate.Text) <> DateValue(cmdTrans(0).Tag) Then Call DisplayAcctDistribution
        'Else
        '    Call DisplayAcctDistribution
        'End If
        'cmdTrans(0).Tag = txtAFLdate.Text
    ElseIf txtAFLDate.Text <> "" Then
        MsgBox "Invalid Date!", vbExclamation + vbOKOnly, "Date"
    End If
End Sub

Private Sub txtCVno_KeyPress(KeyAscii As Integer)
    Dim strValid As String
    
    strValid = "1234567890"
    If KeyAscii > 26 Then
        If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
    If KeyAscii = 13 Then txtAFLDate.SetFocus
End Sub

Function ValidateEntry() As String
    Dim vMsg As String, vQry As String, vModel As String
    Dim y As Integer, i As Integer, CVamt As Currency
    
    vMsg = ""
    If txtRecBy.Text = "" Then vMsg = "Incomplete entry on field RECEIVED BY"
    If txtAFLno.Text = "" Then
        vMsg = vMsg & vbCrLf & "Incomplete entry on field AFL NO."
    ElseIf AddNew Then
        vQry = "SELECT advliquidation.AFLNO From advliquidation Where advliquidation.AFLNO = '" & txtAFLno.Text & "';"
        If CtrlNumExist(vQry) Then vMsg = vMsg & vbCrLf & "AFL No. " & txtAFLno.Text & " already exist"
    End If
    
    If Not IsDate(txtAFLDate.Text) Then
        vMsg = vMsg & vbCrLf & "Invalid entry on field C.R. DATE"
    Else
        vQry = "SELECT cashflow.CASHFLOWID From cashflow Where cashflow.CFDATE = '" & Format(DateValue(txtAFLDate.Text), "YYYY-MM-DD") & "' AND cashflow.UPDATEDBY = " & cbocashier.ItemData(cbocashier.ListIndex) & ";"
        If CtrlNumExist(vQry) Then vMsg = vMsg & vbCrLf & "Cannot continue with transaction because a Cash Flow already Exist"
    End If
    If cbocashier.ListIndex = -1 Then vMsg = vMsg & vbCrLf & "Incomplete entry on field PREPARED BY"
    If txtAprvby.Text = "" Then vMsg = vMsg & vbCrLf & "Incomplete entry on field APPROVED BY"
    
    ValidateEntry = vMsg
End Function

Private Sub SaveEntries()
    Dim rsAdvLiq As New ADODB.Recordset
    Dim Qry As String, ErrLevel As String, Tmp As String
    Dim i As Integer
    
    'On Error GoTo LocalErr
    If AddNew Then
        ErrLevel = "Saving new entries on table Advances for Liquidation"
        rsAdvLiq.Open "AdvLiquidation", LYRcon, adOpenKeyset, adLockOptimistic
        rsAdvLiq.AddNew
    Else
        ErrLevel = "Saving edited entries on table Advances for Liquidation"
        Qry = "SELECT AdvLiquidation.* From AdvLiquidation WHERE (AFLID=" & txtAFLid.Text & ");"
        rsAdvLiq.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If rsAdvLiq.RecordCount > 0 Then rsAdvLiq.MoveFirst
    End If
    With rsAdvLiq
        !AFLNO = txtAFLno.Text
        !AFLDATE = txtAFLDate.Text
        !RECBY = txtRecBy.Text
        !AMTREC = Val(txtAmtRec.Text)
        If AddNew Then
            !TIMEREC = Format(Now, "yyyy-mm-dd hh:mm:ss")
            lblTimeRec.Caption = "Time Received: " & Format(Now, "hh:mm ampm")
            !UPDATEDBY = cbocashier.ItemData(cbocashier.ListIndex)
        Else
            !TIMERETURN = Format(Now, "yyyy-mm-dd hh:mm:ss")
            lblTimeRet.Caption = "Time Returned: " & Format(Now, "hh:mm ampm")
        End If
        !AMTRETURN = Val(txtAmtRet.Text)
        !AFLAMOUNT = Val(txtAFLamt.Text)
        !REMARKS = txtRemarks.Text
        !PREPAREDBY = cbocashier.Text
        !APPROVEDBY = txtAprvby.Text
        !LASTUPDATE = Now
        .Update
        If AddNew Then txtAFLid.Text = !AFLID
        .Close
    End With
    
    Exit Sub
LocalErr:
    MsgBox ErrLevel & vbCrLf & Err.Description, vbCritical + vbOKOnly, "Error"
        
End Sub

Private Sub tmrDrop_Timer()
    cmdTrans(BtnIndx).Top = cmdTrans(BtnIndx).Top + 50
    If cmdTrans(BtnIndx).Top >= 0 Then
        cmdTrans(BtnIndx).Top = 0
        BtnIndx = BtnIndx + 1
        If BtnIndx = 7 Then 'last index + 1
            tmrDrop.Enabled = False
        End If
    End If
End Sub

Private Sub tmrDropRev_Timer()
    cmdTrans(BtnIndx).Top = cmdTrans(BtnIndx).Top + 50
    If cmdTrans(BtnIndx).Top >= 0 Then
        cmdTrans(BtnIndx).Top = 0
        BtnIndx = BtnIndx + 1
        If BtnIndx = 5 Then
            tmrDropRev.Enabled = False
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
            BtnIndx = 5
            cmdTrans(5).Top = -375
            cmdTrans(5).Left = 15
            cmdTrans(5).Visible = True
            cmdTrans(6).Top = -375
            cmdTrans(6).Left = cmdTrans(4).Width + 15
            cmdTrans(6).Visible = True
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
        If BtnIndx = 4 Then
            tmrFloatRev.Enabled = False
            BtnIndx = 0
            For i = 0 To 4
                cmdTrans(i).Visible = True
            Next i
            tmrDropRev.Enabled = True
        End If
    End If
End Sub

