VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form BranchReplPO 
   Caption         =   "Weekly Branch Replenishment P.O"
   ClientHeight    =   9150
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13665
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9150
   ScaleWidth      =   13665
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame4 
      Height          =   735
      Left            =   120
      TabIndex        =   39
      Tag             =   "XYF"
      Top             =   8280
      Width           =   13455
      Begin VB.TextBox txtAprvby 
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
         Left            =   8640
         TabIndex        =   42
         Tag             =   "XYF"
         Top             =   240
         Width           =   4695
      End
      Begin VB.TextBox txtPrepby 
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
         Left            =   1800
         TabIndex        =   40
         Tag             =   "XYF"
         Top             =   240
         Width           =   4695
      End
      Begin VB.Label Label10 
         Caption         =   "APPROVED BY:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7080
         TabIndex        =   43
         Tag             =   "XYF"
         Top             =   255
         Width           =   1575
      End
      Begin VB.Label Label8 
         Caption         =   "PREPARED BY:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   41
         Tag             =   "XYF"
         Top             =   255
         Width           =   1575
      End
   End
   Begin VB.TextBox txtPartnumber 
      Height          =   285
      Left            =   12000
      TabIndex        =   38
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Frame Frame2 
      Height          =   5895
      Left            =   120
      TabIndex        =   3
      Tag             =   "XYF"
      Top             =   1680
      Width           =   13455
      Begin VB.ComboBox cboOrdType 
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
         Left            =   11520
         TabIndex        =   45
         Tag             =   "XYF"
         Top             =   240
         Width           =   1695
      End
      Begin MSFlexGridLib.MSFlexGrid flxReplOrd 
         Height          =   4695
         Left            =   120
         TabIndex        =   13
         Tag             =   "XYF"
         Top             =   720
         Width           =   13215
         _ExtentX        =   23310
         _ExtentY        =   8281
         _Version        =   393216
         Rows            =   3
         Cols            =   9
         FixedRows       =   2
         FixedCols       =   0
         GridLinesFixed  =   1
         MergeCells      =   1
         FormatString    =   "^|<PARTNUMBER |<DESCRIPTION |<LAST DATE IN |^QOH |^SALES QTY |^MQTY |^P.O QTY"
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
      Begin VB.CommandButton cmdAltD 
         Caption         =   "&D"
         Height          =   255
         Left            =   6840
         TabIndex        =   37
         Top             =   2640
         Width           =   375
      End
      Begin VB.CommandButton cmdAltA 
         Caption         =   "&A"
         Height          =   255
         Left            =   6360
         TabIndex        =   36
         Top             =   2640
         Width           =   375
      End
      Begin VB.TextBox txtOvrAllTtlQOH 
         Alignment       =   2  'Center
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
         Left            =   2160
         TabIndex        =   35
         Text            =   "0"
         Top             =   5400
         Width           =   1455
      End
      Begin VB.TextBox txtOvrAllTtlPOQ 
         Alignment       =   2  'Center
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
         Left            =   5760
         TabIndex        =   33
         Text            =   "0"
         Top             =   5400
         Width           =   1695
      End
      Begin VB.TextBox txtOvrAllTtlSQ 
         Alignment       =   2  'Center
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
         Left            =   3840
         TabIndex        =   32
         Text            =   "0"
         Top             =   5400
         Width           =   1695
      End
      Begin MSComCtl2.DTPicker dtFrom 
         Height          =   375
         Left            =   2040
         TabIndex        =   27
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   147128321
         CurrentDate     =   43168
      End
      Begin MSComCtl2.DTPicker dtTO 
         Height          =   375
         Left            =   4200
         TabIndex        =   28
         Top             =   240
         Width           =   1710
         _ExtentX        =   3016
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   147128321
         CurrentDate     =   43168
      End
      Begin lvButton.lvButtons_H cmdRefresh 
         Height          =   375
         Left            =   6000
         TabIndex        =   31
         ToolTipText     =   "Execute"
         Top             =   240
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
         Image           =   "BracnhReplPO.frx":0000
         cBack           =   -2147483633
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "ORDER TYPE:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9960
         TabIndex        =   44
         Tag             =   "XYF"
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label7 
         Caption         =   "TO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3840
         TabIndex        =   30
         Top             =   270
         Width           =   375
      End
      Begin VB.Label Label6 
         Caption         =   "SALES COVERED:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   29
         Top             =   270
         Width           =   1935
      End
      Begin VB.Label lblOvrAllTtl 
         Caption         =   "OVERALL TOTAL:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Tag             =   "XYF"
         Top             =   5520
         Width           =   1815
      End
   End
   Begin VB.TextBox txtPOQ 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   12840
      TabIndex        =   34
      Top             =   0
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Frame Frame3 
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Tag             =   "XYF"
      Top             =   7560
      Width           =   13455
      Begin VB.TextBox txtRemarks 
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
         Left            =   1440
         TabIndex        =   18
         Tag             =   "XYF"
         Top             =   240
         Width           =   5775
      End
      Begin VB.Label Label9 
         Caption         =   "REMARKS:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Tag             =   "XYF"
         Top             =   260
         Width           =   1095
      End
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   3
      Left            =   4620
      TabIndex        =   19
      Top             =   0
      Width           =   1620
      _ExtentX        =   2858
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
      Image           =   "BracnhReplPO.frx":059A
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   20
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
      Image           =   "BracnhReplPO.frx":0B34
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   2
      Left            =   3120
      TabIndex        =   21
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
      Image           =   "BracnhReplPO.frx":10CE
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   4
      Left            =   6240
      TabIndex        =   22
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
      Image           =   "BracnhReplPO.frx":1668
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   5
      Left            =   7740
      TabIndex        =   23
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
      Image           =   "BracnhReplPO.frx":1C02
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   6
      Left            =   9240
      TabIndex        =   24
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
      Image           =   "BracnhReplPO.frx":219C
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   1
      Left            =   1620
      TabIndex        =   25
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
      Image           =   "BracnhReplPO.frx":2736
      cBack           =   -2147483633
   End
   Begin VB.TextBox txtHoldGrayBackColor 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   13920
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   2400
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtHoldWhiteBackColor 
      Height          =   375
      Left            =   13920
      TabIndex        =   0
      Text            =   "HoldBackColor"
      Top             =   1920
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   120
      TabIndex        =   2
      Tag             =   "XYF"
      Top             =   360
      Width           =   13455
      Begin VB.CheckBox chkFinal 
         Caption         =   "FINAL"
         Height          =   255
         Left            =   6840
         TabIndex        =   26
         Tag             =   "XYF"
         Top             =   840
         Width           =   855
      End
      Begin MSMask.MaskEdBox txtDateFinal 
         Height          =   375
         Left            =   5160
         TabIndex        =   16
         Tag             =   "XYF"
         Top             =   675
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtDatePrep 
         Height          =   375
         Left            =   11520
         TabIndex        =   15
         Tag             =   "XYF"
         Top             =   675
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.ComboBox cboCategory 
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
         Left            =   1800
         TabIndex        =   12
         Tag             =   "XYF"
         Top             =   675
         Width           =   1935
      End
      Begin VB.TextBox txtPO 
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
         Left            =   11520
         TabIndex        =   11
         Tag             =   "XYF"
         Top             =   300
         Width           =   1695
      End
      Begin VB.TextBox txtBranchName 
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
         Left            =   1800
         TabIndex        =   10
         Tag             =   "XYF"
         Top             =   300
         Width           =   4935
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "DATE FINAL:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3720
         TabIndex        =   9
         Tag             =   "XYF"
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "DATE PREP:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   10110
         TabIndex        =   8
         Tag             =   "XYF"
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "P.O NO:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   10230
         TabIndex        =   7
         Tag             =   "XYF"
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "CATEGORY:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Tag             =   "XYF"
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "BRANCH NAME:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Tag             =   "XYF"
         Top             =   360
         Width           =   1575
      End
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   11880
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   13905
      m_MDI_Height    =   9720
   End
   Begin VB.Menu mnuStatus 
      Caption         =   "Popupmenu"
      Visible         =   0   'False
      Begin VB.Menu mnuAdd 
         Caption         =   "Add Row"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete Row"
      End
   End
End
Attribute VB_Name = "BranchReplPO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mainRs As New ADODB.Recordset
Dim strPartial As String, strTotal As String
Dim formAction As String
Dim Ans As String, Tmp() As String
Dim MyErr As String, myMsg As String, Qry As String
Dim i As Long, j As Long
Dim ActRow As Integer, ActCol As Integer
Dim StatSrch As Boolean, AddNew As Boolean, ProperExit As Boolean
Dim intTotalQOH As Long, intTotalSQ As Long, intTotalPOQ As Long
Dim arrCATEGORY() As String
Dim strDateFrom As String, strDateTo As String
Dim DateHolder As String
Dim DateEndHolder As String
Dim firstDate As String
Dim intRecordTtl As Long
Dim strCategory As String

Sub clearform()
    
    flxReplOrd.Rows = 2
    flxReplOrd.Rows = 3
    
    txtBranchName.Text = "LYR MOTORPARTS CORP. - MAIN WAREHOUSE"
    cboCategory.Text = ""
    txtRemarks.Text = ""
    txtPO.Text = "<Auto>"
    txtDatePrep.Text = Format(Now, "mm/dd/yyyy")
    txtDateFinal.Text = Format(Now, "mm/dd/yyyy")
    txtOvrAllTtlPOQ.Text = 0
    txtOvrAllTtlQOH.Text = 0
    txtOvrAllTtlSQ.Text = 0
    cboOrdType.Text = ""
    
End Sub


Public Sub DisplaySPOrder(MyRef As String)
    Dim rsOrder As New ADODB.Recordset
    Dim rsITR As New ADODB.Recordset
    Dim Qry As String, i As Integer
    
    Call CboDisplay
    chkFinal.Enabled = False
    
    Qry = "call sp_mrcretails.sp_brnchReplPO_header('" & sysBranchCode & "', '" & MyRef & "')"
    
    With rsOrder
        If .State = adStateOpen Then .Close
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            txtBranchName.Text = !BNAME
            txtDatePrep.Text = Format(!DATEPREPARED, "mm/dd/yyyy")
            txtDateFinal.Text = IIf(IsNull(!DATEFINAL), Format$(Now, "mm/dd/yyyy"), Format(!DATEFINAL, "mm/dd/yyyy"))
            txtRemarks.Text = IIf(IsNull(!Remarks), "", !Remarks)
            txtPO.Text = !RefNo
            cboCategory.Text = !Category
            dtFrom.Value = !DATEFROM
            dtTO.Value = !DATETO
            txtPrepby.Text = IIf(IsNull(!PREPBY), "", !PREPBY)
            txtAprvby.Text = IIf(IsNull(!APRVBY), "", !PREPBY)
            cboOrdType.Text = IIf(IsNull(!ORDERTYPE), "", !ORDERTYPE)
            If !STATUS = "PENDING" Then
                chkFinal.Value = 0
            ElseIf !STATUS <> "PENDING" Then
                chkFinal.Value = 1
            End If
        End If
        .Close
    End With
    
    Qry = "call sp_mrcretails.sp_brnchReplPO_getPODetails('" & sysBranchCode & "','" & MyRef & "','" & Format(dtFrom.Value, "YYYY-MM-DD") & "', '" & Format(dtTO.Value, "YYYY-MM-DD") & "')"
    
    
    With rsITR
        If .State = adStateOpen Then .Close
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            flxReplOrd.Rows = 2
            .MoveFirst
            flxReplOrd.Rows = .RecordCount + 2
            For i = 2 To flxReplOrd.Rows - 1
                flxReplOrd.TextMatrix(i, 0) = i - 1
                flxReplOrd.TextMatrix(i, 1) = !PartNumber
                flxReplOrd.TextMatrix(i, 2) = !Description
                flxReplOrd.TextMatrix(i, 3) = IIf(IsNull(!LASTIN), "", !LASTIN)
                flxReplOrd.TextMatrix(i, 4) = IIf(IsNull(!QOH), 0, !QOH)
                flxReplOrd.TextMatrix(i, 5) = IIf(IsNull(!SQ), 0, !SQ)
                flxReplOrd.TextMatrix(i, 6) = IIf(IsNull(!MAXALLOW), 0, !MAXALLOW)
                flxReplOrd.TextMatrix(i, 7) = IIf(IsNull(!POQ), 0, !POQ)
                flxReplOrd.TextMatrix(i, 8) = !SparepartID
                .MoveNext
            Next i
        End If
        .Close
    End With

    Call ComputeTotal
    dtFrom.Enabled = False
    dtTO.Enabled = False
    cmdRefresh.Enabled = False
End Sub

Sub CreateTempTable()
Dim SEQNO As Integer
On Error GoTo er

        LYRcon.Execute "Drop Temporary Table IF  Exists sp_mrcretails.TMP_REPLPO_DTLS;"

         DoEvents
        LYRcon.Execute " Create Temporary Table IF Not Exists sp_mrcretails.TMP_REPLPO_DTLS( " & _
        "    ORDERQTY Int(11)            " & _
        "   ,UPDATEDBY Int(11)        " & _
        "   ,SPAREPARTID Int(11)      " & _
        "   ,SEQNO INT(11)      " & _
        ");"
        
        Qry = "Insert into sp_mrcretails.TMP_REPLPO_DTLS(ORDERQTY,UPDATEDBY,SPAREPARTID,SEQNO) VALUES "
        SEQNO = 0
        For i = 2 To flxReplOrd.Rows - 1
            If i > 2 Then Qry = Qry & ","
            SEQNO = SEQNO + 1
            Qry = Qry & "(" & Val(flxReplOrd.TextMatrix(i, 7)) & "," & UserID & "," & Val(flxReplOrd.TextMatrix(i, 8)) & "," & SEQNO & ") "
        Next i
        
       Call LYRcon.Execute(Qry)
       
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->CreateTempTable()")
End Sub

Public Sub ComputeTotal()
    intTotalQOH = 0
    intTotalSQ = 0
    intTotalPOQ = 0
    For i = 2 To flxReplOrd.Rows - 1
        intTotalQOH = Val(intTotalQOH) + Val(flxReplOrd.TextMatrix(i, 4))
        intTotalSQ = Val(intTotalSQ) + Val(flxReplOrd.TextMatrix(i, 5))
        intTotalPOQ = Val(intTotalPOQ) + Val(flxReplOrd.TextMatrix(i, 7))
    Next i
    
    txtOvrAllTtlQOH.Text = Format(intTotalQOH, "0")
    txtOvrAllTtlSQ.Text = Format(intTotalSQ, "0")
    txtOvrAllTtlPOQ.Text = Format(intTotalPOQ, "0")
End Sub

Public Sub CboDisplay()
On Error GoTo er
    'Get Data from database
    DoEvents
    Qry = "Call sp_mrcretails.sp_getClassification('" & sysBranchCode & "')"
    With mainRs
        .Open Qry, LYRcon, adOpenStatic, adLockReadOnly
        If .State = 1 Then
           Call modObjTools.addToListObjectWithArrayGivenField(cboCategory, mainRs, arrCATEGORY, "", "CLASSIFICATION")
        End If
        If .State <> 0 Then .Close
    End With
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->CboDisplay()")
End Sub

Public Sub flxColChanger()

With flxReplOrd
    
    'MERGE ROW
    .TextMatrix(0, 0) = "PURCHASE ORDER"
    .TextMatrix(0, 1) = "PURCHASE ORDER"
    .TextMatrix(0, 2) = "PURCHASE ORDER"
    .TextMatrix(0, 3) = "PURCHASE ORDER"
    .TextMatrix(0, 4) = "PURCHASE ORDER"
    .TextMatrix(0, 5) = "PURCHASE ORDER"
    .TextMatrix(0, 6) = "PURCHASE ORDER"
    .TextMatrix(0, 7) = "PURCHASE ORDER"
    
    'MERGE COL =FALSE

    .TextMatrix(1, 1) = "PART NUMBER"
    .TextMatrix(1, 2) = "DESCRIPTION"
    .TextMatrix(1, 3) = "LAST DATE-IN"
    .TextMatrix(1, 4) = "QOH"
    .TextMatrix(1, 5) = "SALES QTY"
    .TextMatrix(1, 6) = "MAX QTY"
    .TextMatrix(1, 7) = "P.O QTY"

    'SET MERGING
    .MergeCells = 1
    .MergeRow(0) = True
    .MergeCol(0) = True
    .MergeCol(1) = True
    For i = 0 To 7
        .MergeCol(i) = False
    Next i
    'SET THE COLUMN COLOR GREEN
    .Row = 0
    .Col = 0
    .Col = 1
    '.CellBackColor = vbGreen
End With
    
    With flxReplOrd
        '.Clear
        .Cols = 9
        '.FormatString = "|<PARTNUMBER |<DESCRIPTION |<LAST DATE IN |^QOH |^SALES QTY |^MAX QTY |^P.O QTY |<SPID"
        .ColWidth(0) = 500
        .ColWidth(1) = (.Width - .ColWidth(0)) * 0.2   'PART NUMBER
        .ColWidth(2) = (.Width - .ColWidth(0)) * 0.36 'DESCRIPTION
        .ColWidth(3) = (.Width - .ColWidth(0)) * 0.1 'LAST DATE IN
        .ColWidth(4) = (.Width - .ColWidth(0)) * 0.08 'QOH
        .ColWidth(5) = (.Width - .ColWidth(0)) * 0.08 'SALES QTY
        .ColWidth(6) = (.Width - .ColWidth(0)) * 0.08 'MAX QTY
        .ColWidth(7) = (.Width - .ColWidth(0)) * 0.08 'P.O QTY
        .ColWidth(8) = 0 '(.Width - .ColWidth(0)) * 0.06 'SPID
        
        txtOvrAllTtlQOH.Top = .Top + .Height + 50
        txtOvrAllTtlQOH.Left = .Left + .ColWidth(0) + .ColWidth(1) + .ColWidth(2) + .ColWidth(3)
        'txtOvrAllTtlQOH.Height = txtOvrAllTtlSQ.Height
        txtOvrAllTtlQOH.Width = .ColWidth(4)
        
        txtOvrAllTtlSQ.Top = .Top + .Height + 50
        txtOvrAllTtlSQ.Left = .Left + .ColWidth(0) + .ColWidth(1) + .ColWidth(2) + .ColWidth(3) + .ColWidth(4)
        txtOvrAllTtlSQ.Width = .ColWidth(5)
        
        txtOvrAllTtlPOQ.Top = .Top + .Height + 50
        txtOvrAllTtlPOQ.Left = .Left + .ColWidth(0) + .ColWidth(1) + .ColWidth(2) + .ColWidth(3) + .ColWidth(4) + .ColWidth(5) + .ColWidth(6)
        txtOvrAllTtlPOQ.Width = .ColWidth(6)
        
        lblOvrAllTtl.Top = .Top + .Height + 50
        lblOvrAllTtl.Left = .Left + .ColWidth(0) + .ColWidth(1) + .ColWidth(2)
        lblOvrAllTtl.Width = .ColWidth(3)
    End With
End Sub
    
Private Sub SetButtonVisibility(DataEntry As Boolean)
'cmdTrans Legend
'0 ->ADD
'1 ->EDIT
'2 ->VIEW SUMM
'3 ->PRINT
'4 ->EXIT
'5 ->SAVE
'6 ->CANCEL

    'hide all button
    For i = 0 To 4
        cmdTrans(i).Visible = False
    Next i

    If DataEntry Then
        cmdTrans(5).Visible = True
        cmdTrans(6).Visible = True
    Else
    For i = 0 To 4
        cmdTrans(i).Visible = True
        cmdTrans(5).Visible = False
        cmdTrans(6).Visible = False
    Next i
    
    End If
End Sub

Private Sub cboCategory_Change()
On Error GoTo er
    If StatSrch Then
            StatSrch = False
            Exit Sub
    End If

    With cboCategory
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
    Call modObjTools.ErrorMessage(Err.Description, "Error->cboCategory-Change()")
End Sub

Private Sub cboCategory_Click()
    Call cmdRefresh_Click
End Sub

Private Sub cmdAltA_Click()
'Call mnuAdd_Click
End Sub

Private Sub cmdAltD_Click()
Call mnuDelete_Click
End Sub

Private Sub cmdRefresh_Click()
On Error GoTo er
Dim strCrit As String
Dim rsDetails As New ADODB.Recordset

    intRecordTtl = 0
    With rsAnyUse
        Qry = "call sp_mrcretails.sp_brnchReplPO_chkExistsDateCov('" & sysBranchCode & "', '" & Format(dtFrom.Value, "YYYY-MM-DD") & "', '" & Format(dtTO.Value, "YYYY-MM-DD") & "', '" & Trim(cboCategory.Text) & "')"
        If .State = adStateOpen Then .Close
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            If !level = "Error" Then
                MsgBox "SORRY I CANNOT PROCESS THE TRANSACTION! Please Select a Sales Coverage that does not contain Sales from '" & !DATEFROM & "' TO '" & !DATETO & "' "
                Exit Sub
            End If
            .Close
        End If
    End With
    If cboCategory.Text <> "" Then
        If strDateFrom <> dtFrom.Value Or strDateTo <> dtTO.Value Or strCategory <> cboCategory.Text Then
            Qry = "call sp_mrcretails.sp_brnchReplPO_getDetails('" & sysBranchCode & "','" & cboCategory.Text & "','" & Format(dtFrom.Value, "yyyy-mm-dd") & "','" & Format(dtTO.Value, "yyyy-mm-dd") & "')"
            If rsDetails.State = adStateOpen Then rsDetails.Close
            With rsDetails
                .Open Qry, LYRcon, adOpenStatic, adLockReadOnly
                If .RecordCount > 0 Then
                    .MoveFirst
                    If !TtlRecord > 50 Then
                        MsgBox "You Have " & !TtlRecord & " items to P.O. We Can Only Fetch 50 Records Per P.O! Please Click OK to continue!", vbInformation + vbOKOnly, "INFORMATION"
                        
                    End If
                    
                    intRecordTtl = !TtlRecord
                    flxReplOrd.Rows = .RecordCount + 2
                    For i = 2 To flxReplOrd.Rows - 1
                        flxReplOrd.TextMatrix(i, 0) = i - 1
                        flxReplOrd.TextMatrix(i, 1) = !PartNumber
                        flxReplOrd.TextMatrix(i, 2) = !Description
                        flxReplOrd.TextMatrix(i, 3) = !LASTIN
                        flxReplOrd.TextMatrix(i, 4) = !QOH
                        flxReplOrd.TextMatrix(i, 5) = !SQ
                        flxReplOrd.TextMatrix(i, 6) = !MAXALLOW
                        If !MAXALLOW > 0 Then
                            If !SQ >= !MAXALLOW Then
                                flxReplOrd.TextMatrix(i, 7) = !MAXALLOW
                            Else
                                flxReplOrd.TextMatrix(i, 7) = !SQ
                            End If
                        Else
                            flxReplOrd.TextMatrix(i, 7) = !SQ
                        End If
                        flxReplOrd.TextMatrix(i, 8) = !SparepartID
                        .MoveNext
                    Next i
                End If
            
                If .RecordCount <= 0 Then
                    MsgBox "No Record Found! Please Select another Category.", vbInformation + vbOKOnly, "VALIDATION"
                    strCategory = cboCategory.Text
                    .Close
                    Exit Sub
                End If
                .Close
            End With
         End If
    Else
        MsgBox "Please Select a Category!", vbInformation + vbOKOnly, "VALIDATION"
        Exit Sub
    End If
    
    txtRemarks.Text = ""
    If strDateFrom <> "" Then
        If firstDate = "" Then
            firstDate = strDateFrom
        End If
        If CDate(firstDate) <= CDate(dtFrom.Value) Then
            DateHolder = firstDate
        Else
            DateHolder = dtFrom.Value
        End If
        If DateEndHolder <> "" Then
            If CDate(DateEndHolder) >= CDate(dtTO.Value) Then
                DateEndHolder = DateEndHolder
            Else
                DateEndHolder = dtTO.Value
            End If
        Else
            DateEndHolder = dtTO.Value
        End If
        txtRemarks.Text = txtRemarks.Text & " SALES COVERED:" & DateHolder & " TO " & DateEndHolder
    Else
        strDateFrom = dtFrom.Value
        strDateTo = dtTO.Value
        If firstDate = "" Then
            firstDate = strDateFrom
        End If
        If CDate(firstDate) <= CDate(dtFrom.Value) Then
            DateHolder = firstDate
        Else
            DateHolder = dtFrom.Value
        End If
        If DateEndHolder <> "" Then
            If CDate(DateEndHolder) >= CDate(dtTO.Value) Then
                DateEndHolder = DateEndHolder
            Else
                DateEndHolder = dtTO.Value
            End If
        Else
            DateEndHolder = dtTO.Value
        End If
        txtRemarks.Text = txtRemarks.Text & " SALES COVERED:" & DateHolder & " TO " & DateEndHolder
    End If
    
    strDateFrom = dtFrom.Value
    strDateTo = dtTO.Value
    strCategory = cboCategory.Text
    Call ComputeTotal
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->RefreshRecord)")
End Sub

Public Sub clearform_unalloc()
    formAction = "add"
    flxReplOrd.Rows = 2
    flxReplOrd.Rows = 3
    txtPO.Text = "<Auto>"
    txtOvrAllTtlPOQ.Text = 0
    txtOvrAllTtlQOH.Text = 0
    txtOvrAllTtlSQ.Text = 0
    cboOrdType.Text = ""
    
    cmdRefresh.Enabled = False
    cmdTrans(6).Enabled = False
    chkFinal.Enabled = True
    chkFinal.Value = 0

    dtFrom.Enabled = False
    dtTO.Enabled = False
    cboCategory.Enabled = False
    
    txtPrepby.Text = UserName
    txtAprvby.Text = UCase(SysBranchOIC)
End Sub
Private Sub cmdTrans_Click(Index As Integer)
Dim strDateFinal As String
Dim rsTrans As New ADODB.Recordset
Dim rsUnAllocPO As New ADODB.Recordset

On Error GoTo er

    If Index = 0 Then 'add
        formAction = "add"
        
        Call clearform
        Call LockData(True)
        Call SetButtonVisibility(True)
        cmdRefresh.Enabled = True
        chkFinal.Enabled = True
        chkFinal.Value = 0
        cboCategory.Enabled = True
        cmdTrans(6).Enabled = True
        dtFrom.Enabled = True
        dtTO.Enabled = True
        Call CboDisplay
        
        txtPrepby.Text = UserName
        txtAprvby.Text = UCase(SysBranchOIC)
        
    ElseIf Index = 1 Then 'edit
        formAction = "edit"
        If chkFinal.Value = 1 Then
            MsgBox " '" & txtPO.Text & "' is already final!", vbInformation + vbOKOnly, "Edit"
            Exit Sub
        End If
        Call LockData(True)
        Call SetButtonVisibility(True)
        chkFinal.Enabled = True
        cmdTrans(6).Enabled = True
        cboCategory.BackColor = txtHoldGrayBackColor.BackColor
        cboCategory.Locked = True
        dtFrom.Enabled = False
        dtTO.Enabled = False
        cmdRefresh.Enabled = False
    ElseIf Index = 2 Then 'view summary
        FrmSummary.mySrc = "Branch Order"
        FrmSummary.Show
        Unload Me
    ElseIf Index = 3 Then 'view summary
        Ans = MsgBox("Do you want to Print P.O.#" & txtPO.Text & "?", vbQuestion + vbYesNo, "Print")
        If Ans = vbNo Then Exit Sub
        Qry = "call sp_mrcretails.sp_brnchReplPO_getPODetails('" & sysBranchCode & "', '" & txtPO.Text & "','" & Format(dtFrom.Value, "YYYY-MM-DD") & "', '" & Format(dtTO.Value, "YYYY-MM-DD") & "')"
        frmPrint.QryPrn = Qry
        frmPrint.ParamValue1 = SysBranch
        frmPrint.ParamValue2 = txtBranchName.Text
        frmPrint.ParamValue3 = cboCategory.Text
        frmPrint.ParamValue4 = txtPO.Text
        frmPrint.ParamValue5 = txtDatePrep.Text
        If chkFinal.Value = 0 Then
            frmPrint.ParamValue6 = ""
            frmPrint.ParamValue7 = "PENDING"
        Else
            frmPrint.ParamValue6 = txtDateFinal.Text
            frmPrint.ParamValue7 = "FINAL"
        End If
        frmPrint.ParamValue8 = txtRemarks.Text
        frmPrint.ParamValue9 = txtPrepby.Text
        frmPrint.ParamValue10 = txtAprvby.Text
        frmPrint.ParamValue11 = cboOrdType.Text
        frmPrint.MySource = "Branch Order"
        frmPrint.Show
    ElseIf Index = 4 Then 'exit
        MotoRep.imgBG.Visible = False
        MotoRep.toolBarButton.Visible = True
        formAction = "EXIT"
        ProperExit = True
        Unload Me
    ElseIf Index = 5 Then 'save
        
        If flxReplOrd.Rows > 2 And cboOrdType.Text <> "" Then
        
            Ans = MsgBox("Do you want to save?", vbQuestion + vbYesNo, "Save")
            
            If Ans = vbNo Then Exit Sub
                
            If formAction = "add" Then
                Call CreateTempTable
                Qry = "call sp_mrcretails.sp_brnchReplPO_create('" & sysBranchCode & "','" & cboCategory.Text & "','" & txtRemarks.Text & "','" & Format(txtDatePrep.Text, "YYYY-MM-DD") & "'," & Val(chkFinal.Value) & ", " & UserID & ",'" & Format(dtFrom.Value, "yyyy-mm-dd") & "','" & Format(dtTO.Value, "yyyy-mm-dd") & "','" & txtAprvby.Text & "','" & Trim(cboOrdType.Text) & "')"
            Else
                Call CreateTempTable
                Qry = "call sp_mrcretails.sp_brnchReplPO_update('" & sysBranchCode & "','" & txtPO.Text & "','" & cboCategory.Text & "','" & txtRemarks.Text & "'," & Val(chkFinal.Value) & ", " & UserID & ",'" & Format(dtFrom.Value, "yyyy-mm-dd") & "','" & Format(dtTO.Value, "yyyy-mm-dd") & "','" & Trim(cboOrdType.Text) & "')"
            End If
            
            With rsTrans
                .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
                If .RecordCount > 0 Then
                    If !level = "Granted" Then
                        MsgBox " '" & !message & "' ", vbInformation + vbOKOnly, "SUCCESSFUL"
                        txtPO.Text = !RefNo
                    Else
                        MsgBox " '" & !message & "' ", vbCritical + vbOKOnly, "UNSUCCESSFUL"
                        Exit Sub
                    End If
                End If
                .Close
            End With
            
            If Val(chkFinal.Value) = 1 Then
                Dim TOTALPN As Long
                Dim ProcessedBy As String
                frm_OnProcess.OwnerForm = "Identifying"
                frm_OnProcess.Show
                Call ConnectOnline
                TOTALPN = 0
                
                Qry = " SELECT CONCAT(FIRSTNAME, ' ', LASTNAME) From employees Where EMPID = " & UserID & " "
                ProcessedBy = GetFieldValue(Qry)
                
                With flxReplOrd
                    For i = 2 To .Rows - 1
                        If Val(.TextMatrix(i, 7)) <> 0 Then
                            TOTALPN = TOTALPN + 1
                            OnlineCon.Execute "Insert Into sporderdtls( REFERENCE, SPAREPARTID, TRANSDATE, TRANSTYPE, `ORDER`, " & _
                            "SOURCE, UPDATEDBY, LASTUPDATE, SEQNO) VALUES( '" & txtPO.Text & "', " & _
                            "" & Val(flxReplOrd.TextMatrix(i, 8)) & ", '" & Format(Now, "yyyy-mm-dd") & "', " & _
                            "'" & cboOrdType.Text & "', " & Val(flxReplOrd.TextMatrix(i, 7)) & ", '" & sysBranchCode & "', " & _
                            "" & UserID & ", '" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "', " & i & ");"
                        End If
                    Next i
                End With
                
                OnlineCon.Execute "Insert Into sporder( DEALERCODE, DEALERNAME, REFERENCE, ORDERTYPE, ORDERDATE, " & _
                "SUBMITTEDBY, UPDATEDBY, QTYPN, QTYORDER, STATUS, CATEGORY, REMARKS)" & _
                "VALUES('" & sysBranchCode & "', '" & txtBranchName.Text & "', '" & txtPO.Text & "', " & _
                "'" & cboOrdType.Text & "','" & Format(txtDateFinal.Text, "yyyy-mm-dd") & "', " & _
                "'" & ProcessedBy & "', " & UserID & ", " & _
                "" & TOTALPN & ", " & Val(Replace(txtOvrAllTtlPOQ.Text, ",", "")) & ",  'NEW', '" & cboCategory.Text & "', '" & txtRemarks.Text & "');"
        
                OnlineCon.Close
    
                Call Unload(frm_OnProcess)
                LYRcon.Execute "UPDATE sporder SET `STATUS` = 'NEW' Where sporder.REFERENCE = '" & txtPO.Text & "';"
                chkFinal.Enabled = False
                MsgBox "Replenishment Request successfully saved online!", vbInformation + vbOKOnly, "Save"
            End If
            
            If intRecordTtl > 50 Then
                Qry = "SELECT COUNT(SPAREPARTID) as TTLSPID FROM sp_mrcretails.TMP_data2"
                With rsAnyUse
                    If .State = adStateOpen Then .Close
                    .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
                    If .RecordCount > 0 Then
                        If !TTLSPID > 0 Then
                            MsgBox "There is " & !TTLSPID & " Part Number to Replenish!", vbInformation + vbOKOnly, "Validation"
                            Call clearform_unalloc
                            Qry = "call sp_mrcretails.sp_brnchReplPO_getDetailsUnAllocPO()"
                            With rsUnAllocPO
                                If .State = adStateOpen Then .Close
                                .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
                                If .RecordCount > 0 Then
                                    flxReplOrd.Rows = .RecordCount + 2
                                    For i = 2 To flxReplOrd.Rows - 1
                                        flxReplOrd.TextMatrix(i, 0) = i - 1
                                        flxReplOrd.TextMatrix(i, 1) = !PartNumber
                                        flxReplOrd.TextMatrix(i, 2) = !Description
                                        flxReplOrd.TextMatrix(i, 3) = !LASTIN
                                        flxReplOrd.TextMatrix(i, 4) = !QOH
                                        flxReplOrd.TextMatrix(i, 5) = !SQ
                                        flxReplOrd.TextMatrix(i, 6) = !MAXALLOW
                                        If !MAXALLOW > 0 Then
                                            If !SQ >= !MAXALLOW Then
                                                flxReplOrd.TextMatrix(i, 7) = !MAXALLOW
                                            Else
                                                flxReplOrd.TextMatrix(i, 7) = !SQ
                                            End If
                                        Else
                                            flxReplOrd.TextMatrix(i, 7) = !SQ
                                        End If
                                        flxReplOrd.TextMatrix(i, 8) = !SparepartID
                                        LYRcon.Execute "DELETE FROM sp_mrcretails.TMP_data2 where SPAREPARTID = " & !SparepartID & " "
                                        .MoveNext
                                    Next i
                                End If
                                Call ComputeTotal
                            End With
                        Else
                            Call SetButtonVisibility(False)
                            Call LockData(False)
                            chkFinal.Enabled = False
                            dtFrom.Enabled = False
                            dtTO.Enabled = False
                            cmdRefresh.Enabled = False
                            formAction = ""
                        End If
                    End If
                End With
           Else
                Call SetButtonVisibility(False)
                Call LockData(False)
                chkFinal.Enabled = False
                dtFrom.Enabled = False
                dtTO.Enabled = False
                cmdRefresh.Enabled = False
                formAction = ""
            End If
        Else
            MsgBox "Please Select an Order Type!", vbInformation + vbOKOnly, "Validation"
        End If
    ElseIf Index = 6 Then 'cancel
        
        Ans = MsgBox("Do you want to cancel?", vbQuestion + vbYesNo)
        
        If Ans = vbNo Then Exit Sub
        
        Call SetButtonVisibility(False)
        Call LockData(False)
        chkFinal.Enabled = False
        dtFrom.Enabled = False
        dtTO.Enabled = False
        cmdRefresh.Enabled = False
        formAction = ""
    End If
    
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->cmdTrans-Click()")
End Sub

Private Sub flxReplOrd_Click()
On Error GoTo er

    ActRow = flxReplOrd.Row
    ActCol = flxReplOrd.Col
    
    If formAction = "add" Or formAction = "edit" Then
    
    If ActCol = 0 Or ActRow = 0 Then Exit Sub
        Select Case ActCol
        
        Case 7
            If flxReplOrd.TextMatrix(flxReplOrd.Row, 1) <> "" Then
                With txtPOQ
                    If flxReplOrd.TextMatrix(ActRow, 7) <> "" Then .Tag = flxReplOrd.TextMatrix(ActRow, 7)
                    .Top = Frame2.Top + flxReplOrd.CellTop + flxReplOrd.Top
                    .Left = Frame2.Left + flxReplOrd.CellLeft + flxReplOrd.Left
                    .Height = flxReplOrd.CellHeight
                    .Width = flxReplOrd.CellWidth
                    .Text = flxReplOrd.Text
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
    Call modObjTools.ErrorMessage(Err.Description, "Error->flxReplOrd-click()")
End Sub

Private Sub flxReplOrd_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Call flxReplOrd_Click
End Sub

Private Sub flxReplOrd_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    If flxReplOrd.Rows > 3 Then
        If Button = vbRightButton Then
            Call PopupMenu(mnuStatus)
        End If
    End If
End Sub

Private Sub flxReplOrd_Scroll()
txtPOQ.Visible = False
End Sub

Private Sub Form_Load()
    Call LockData(False)
    
    txtPO.Text = "<AUTO>"
    txtBranchName.Text = "LYR MOTORPARTS CORP. - MAIN WAREHOUSE"
    txtDatePrep.Text = Format(ServerDate, "mm/dd/yyyy")
    txtDateFinal.Text = Format(ServerDate, "mm/dd/yyyy")
    cboOrdType.AddItem "(01) - URGENT"
    cboOrdType.AddItem "(02) - REGULAR"
    chkFinal.Enabled = False
    
    cmdRefresh.Enabled = False
    dtFrom.Enabled = False
    dtTO.Enabled = False
    dtFrom.Value = DateAdd("d", -7, ServerDate)
    dtTO.Value = Format(ServerDate, "mm/dd/yyyy")
    
    '===================BUTTON POSITION===================
    If cmdTrans(0).Visible = False Then
        'Align Save Button
        cmdTrans(5).Left = cmdTrans(0).Left
        cmdTrans(5).Top = cmdTrans(0).Top
        'Align Cancel Button
        cmdTrans(6).Left = cmdTrans(1).Left
        cmdTrans(6).Top = cmdTrans(1).Top
    End If
End Sub

Sub LockData(DataEntry As Boolean)
    DataEntry = Not DataEntry

    txtBranchName.Locked = DataEntry
    cboCategory.Locked = DataEntry
    txtRemarks.Locked = DataEntry
    txtPO.Locked = DataEntry
    txtOvrAllTtlQOH.Locked = DataEntry
    txtOvrAllTtlSQ.Locked = DataEntry
    txtOvrAllTtlPOQ.Locked = DataEntry
    txtPrepby.Locked = DataEntry
    txtAprvby.Locked = DataEntry
    cboOrdType.Locked = DataEntry
    
    If txtBranchName.Locked Then
        txtBranchName.BackColor = txtHoldGrayBackColor.BackColor
        cboCategory.BackColor = txtHoldGrayBackColor.BackColor
        txtDatePrep.BackColor = txtHoldGrayBackColor.BackColor
        txtDateFinal.BackColor = txtHoldGrayBackColor.BackColor
        txtPO.BackColor = txtHoldGrayBackColor.BackColor
        txtRemarks.BackColor = txtHoldGrayBackColor.BackColor
        txtOvrAllTtlQOH.BackColor = txtHoldGrayBackColor.BackColor
        txtOvrAllTtlSQ.BackColor = txtHoldGrayBackColor.BackColor
        txtOvrAllTtlPOQ.BackColor = txtHoldGrayBackColor.BackColor
        txtPrepby.BackColor = txtHoldGrayBackColor.BackColor
        txtAprvby.BackColor = txtHoldGrayBackColor.BackColor
        cboOrdType.BackColor = txtHoldGrayBackColor.BackColor
    Else
        txtBranchName.Locked = True
        cboCategory.BackColor = txtHoldWhiteBackColor.BackColor
        txtPO.Locked = True
        txtRemarks.BackColor = txtHoldWhiteBackColor.BackColor
        txtOvrAllTtlQOH.Locked = True
        txtOvrAllTtlSQ.Locked = True
        txtOvrAllTtlPOQ.Locked = True
        txtPrepby.Locked = True
        txtAprvby.Locked = True
        cboOrdType.BackColor = txtHoldWhiteBackColor.BackColor
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo er
    If ProperExit Then
        If mainRs.State = adStateOpen Then mainRs.Close
        Call cmdTrans_Click(4)
        ProperExit = False
    Else
        Cancel = True
    End If
    
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->QueryUnload()")
End Sub

Private Sub mnuAdd_Click()
    If flxReplOrd.TextMatrix(flxReplOrd.Rows - 1, 7) = "" Or flxReplOrd.TextMatrix(flxReplOrd.Rows - 1, 1) = "" Or flxReplOrd.TextMatrix(flxReplOrd.Rows - 1, 8) = "" Then
        MsgBox "Incomplete entry on last row!", vbCritical + vbOKOnly, "Add Row failed"
        flxReplOrd.SetFocus
        flxReplOrd.Col = ActCol
    Else
        flxReplOrd.AddItem ""
        flxReplOrd.SetFocus
        flxReplOrd.Col = 1
        flxReplOrd.Row = flxReplOrd.Rows - 1
    End If
End Sub

Private Sub mnuDelete_Click()
On Error GoTo er

    Ans = MsgBox("Do you want to Delete Row?", vbQuestion + vbYesNo, "Delete")
    If Ans = vbYes Then
        If flxReplOrd.Rows = 3 Then
            flxReplOrd.Rows = 2
            flxReplOrd.Rows = 3
        Else
            flxReplOrd.RemoveItem ActRow
        End If
    End If
    Call ComputeTotal
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->mnuDelete()")
End Sub

Private Sub RXResizer1_ResizeComplete()
    Call flxColChanger
End Sub




Private Sub txtPOQ_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
    Tmp = Split(txtPOQ.Tag, "|")
    
    For i = LBound(Tmp) To UBound(Tmp)
        flxReplOrd.TextMatrix(ActRow, i + 1) = Tmp(i)
    Next i
    
    txtPOQ.Visible = False
    flxReplOrd.SetFocus
    flxReplOrd.Col = 7
End If
End Sub

Private Sub txtPOQ_KeyPress(KeyAscii As Integer)
On Error GoTo er

    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtPOQ, KeyAscii, NumbersOnly)

    
    If KeyAscii = 13 Then
        
        flxReplOrd.TextMatrix(ActRow, 7) = txtPOQ.Text
        
'        If Val(flxReplOrd.TextMatrix(ActRow, 11)) > Val(flxReplOrd.TextMatrix(ActRow, 7)) Then
'            MsgBox "Allocated Quantity is Greater Than to Available Quantity", vbInformation + vbOKOnly, "Validation"
'            flxReplOrd.TextMatrix(ActRow, 11) = 0
'            txtALQ.Text = ""
'            txtALQ.Visible = False
'            flxReplOrd.SetFocus
'            flxReplOrd.Row = ActRow
'            flxReplOrd.Col = 11
'            Exit Sub
'        End If
'
'
'        If txtPOQ.Text <> "" Then
'
'            For i = ActRow To flxReplOrd.Rows - 1
'                If flxReplOrd.TextMatrix(i, 6) = "" Or Val(flxReplOrd.TextMatrix(i, 11)) = 0 Then
'                    txtPOQ.Visible = False
'                    flxReplOrd.SetFocus
'                    flxReplOrd.Row = i
'                    flxReplOrd.Col = 6
'                Exit For
'                Else
'                    txtALQ.Visible = False
'                    flxReplOrd.SetFocus
'                    flxReplOrd.Row = ActRow
'                    flxReplOrd.Col = 6
'                End If
'
'            Next i
'
'        Else
        flxReplOrd.SetFocus
        flxReplOrd.Col = 7
        txtPOQ.Visible = False
'        End If
        Call ComputeTotal
        
    End If

Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->txtALQ-KeyPress()")
End Sub

Private Sub txtPOQ_LostFocus()
On Error GoTo er

    If txtPOQ.Tag = "skiplostfocus" Then
        txtPOQ.Tag = ""
        Exit Sub
    End If
    
    If flxReplOrd.TextMatrix(ActRow, 7) = "" Or txtPOQ.Text = "" Then
        MsgBox "Wrong Entry of Purchase Order Quantity!" & vbCrLf & "Please Enter a Correct Purchase Order Quantity.", vbExclamation + vbOKOnly, "Warning"
        txtPOQ.Visible = False
        flxReplOrd.TextMatrix(ActRow, 7) = ""
        flxReplOrd.SetFocus
        flxReplOrd.Col = 7
        Exit Sub
    ElseIf flxReplOrd.TextMatrix(ActRow, 6) > 0 Then
        If flxReplOrd.TextMatrix(ActRow, 7) > flxReplOrd.TextMatrix(ActRow, 6) Then
            MsgBox "P.O Quantity is Greater than the Max Allowable Qty!" & vbCrLf & "Please Enter a Correct Purchase Order Quantity.", vbExclamation + vbOKOnly, "Warning"
            If flxReplOrd.TextMatrix(ActRow, 5) >= flxReplOrd.TextMatrix(ActRow, 6) Then
                flxReplOrd.TextMatrix(ActRow, 7) = flxReplOrd.TextMatrix(ActRow, 6)
                flxReplOrd.SetFocus
                flxReplOrd.Col = 7
                Exit Sub
            Else
                flxReplOrd.TextMatrix(ActRow, 7) = flxReplOrd.TextMatrix(ActRow, 5)
                flxReplOrd.SetFocus
                flxReplOrd.Col = 7
                Exit Sub
            End If
        End If
    End If
    
    flxReplOrd.TextMatrix(ActRow, 7) = txtPOQ.Text
    txtPOQ.Visible = False
    
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->txtPOQ-LostFocus()")
End Sub

Private Sub txtRemarks_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtRemarks, KeyAscii, ValidChars)
End Sub
