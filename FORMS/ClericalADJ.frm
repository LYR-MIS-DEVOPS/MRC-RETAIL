VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form ClericalADJ 
   Caption         =   "Clerical Adjustments"
   ClientHeight    =   5385
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11160
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5385
   ScaleWidth      =   11160
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtJobID 
      Height          =   285
      Left            =   120
      TabIndex        =   50
      Top             =   7800
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtJobAmt 
      Height          =   285
      Left            =   120
      TabIndex        =   49
      Top             =   7560
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtJobDate 
      Height          =   285
      Left            =   120
      TabIndex        =   48
      Top             =   7320
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtCustID 
      Height          =   285
      Left            =   120
      TabIndex        =   46
      Top             =   7080
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtPrnCopy 
      Height          =   285
      Left            =   120
      TabIndex        =   38
      Top             =   6840
      Visible         =   0   'False
      Width           =   855
   End
   Begin TabDlg.SSTab tabParts 
      Height          =   2535
      Left            =   1920
      TabIndex        =   34
      Top             =   5640
      Visible         =   0   'False
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   4471
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Select Spare Parts"
      TabPicture(0)   =   "ClericalADJ.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "cmdCancelt"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdOKt"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "flxParts"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      Begin MSFlexGridLib.MSFlexGrid flxParts 
         Height          =   1575
         Left            =   120
         TabIndex        =   35
         Top             =   360
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   2778
         _Version        =   393216
         Cols            =   5
         BackColorBkg    =   -2147483633
         SelectionMode   =   1
         FormatString    =   "     |<  PART NUMBER        |<DESCRIPTION                              |>PRICE         "
      End
      Begin lvButton.lvButtons_H cmdOKt 
         Height          =   375
         Left            =   2640
         TabIndex        =   36
         Top             =   2040
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         Caption         =   "&OK"
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
         Image           =   "ClericalADJ.frx":001C
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdCancelt 
         Height          =   375
         Left            =   4080
         TabIndex        =   37
         Top             =   2040
         Width           =   1335
         _ExtentX        =   2355
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
         Image           =   "ClericalADJ.frx":05B6
         cBack           =   -2147483633
      End
   End
   Begin VB.TextBox txtSPid 
      Height          =   285
      Left            =   2280
      TabIndex        =   33
      Top             =   5640
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Timer tmrDropRev 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1680
      Top             =   6000
   End
   Begin VB.Timer tmrFloatRev 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1680
      Top             =   5520
   End
   Begin VB.Timer tmrDrop 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1080
      Top             =   6000
   End
   Begin VB.Timer tmrFloat 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1080
      Top             =   5520
   End
   Begin VB.Frame Frame2 
      Height          =   3495
      Left            =   120
      TabIndex        =   15
      Tag             =   "XYF"
      Top             =   1800
      Width           =   10935
      Begin VB.ComboBox cboINVTYPE 
         Height          =   315
         ItemData        =   "ClericalADJ.frx":0B50
         Left            =   7200
         List            =   "ClericalADJ.frx":0B5A
         TabIndex        =   53
         Top             =   360
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.ComboBox cboOnCust 
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
         Left            =   7200
         TabIndex        =   51
         Tag             =   "XYF"
         Top             =   720
         Width           =   3615
      End
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
         Height          =   330
         Left            =   7200
         MaxLength       =   50
         TabIndex        =   40
         Tag             =   "XYF"
         Top             =   720
         Width           =   1815
      End
      Begin VB.TextBox txtAprvBy 
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
         Left            =   2040
         MaxLength       =   50
         TabIndex        =   31
         Tag             =   "XYF"
         Top             =   2880
         Width           =   8775
      End
      Begin VB.TextBox txtPrepBy 
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
         Left            =   2040
         MaxLength       =   50
         TabIndex        =   30
         Tag             =   "XYF"
         Top             =   2520
         Width           =   8775
      End
      Begin VB.TextBox txtRemarks 
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
         Left            =   2040
         MaxLength       =   200
         TabIndex        =   29
         Tag             =   "XYF"
         Top             =   2160
         Width           =   8775
      End
      Begin VB.TextBox txtReason 
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
         Left            =   2040
         MaxLength       =   200
         TabIndex        =   28
         Tag             =   "XYF"
         Top             =   1800
         Width           =   8775
      End
      Begin VB.TextBox txtEncoder 
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
         Left            =   2040
         MaxLength       =   50
         TabIndex        =   27
         Tag             =   "XYF"
         Top             =   1440
         Width           =   8775
      End
      Begin VB.TextBox txtNewValue 
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
         Left            =   2040
         MaxLength       =   50
         TabIndex        =   26
         Tag             =   "XYF"
         Top             =   1080
         Width           =   2895
      End
      Begin VB.TextBox txtOldValue 
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
         Left            =   2040
         MaxLength       =   50
         TabIndex        =   25
         Tag             =   "XYF"
         Top             =   720
         Width           =   2895
      End
      Begin VB.ComboBox cboField 
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
         Left            =   2040
         TabIndex        =   17
         Tag             =   "XYF"
         Top             =   360
         Width           =   2895
      End
      Begin VB.Label lblInvType 
         Alignment       =   1  'Right Justify
         Caption         =   "INVOICE TYPE:"
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
         Left            =   5160
         TabIndex        =   54
         Tag             =   "XYF"
         Top             =   360
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label lblCustomer 
         Alignment       =   1  'Right Justify
         Caption         =   "CUSTOMER:"
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
         Left            =   5880
         TabIndex        =   52
         Tag             =   "XYF"
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label2 
         Height          =   255
         Left            =   5040
         TabIndex        =   47
         Top             =   1080
         Width           =   4215
      End
      Begin VB.Label lblRefNo 
         Alignment       =   1  'Right Justify
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
         Left            =   5160
         TabIndex        =   39
         Tag             =   "XYF"
         Top             =   765
         Width           =   1935
      End
      Begin VB.Label lblDesc 
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
         Left            =   5040
         TabIndex        =   32
         Tag             =   "XYF"
         Top             =   765
         Visible         =   0   'False
         Width           =   4215
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
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
         Index           =   7
         Left            =   360
         TabIndex        =   24
         Tag             =   "XYF"
         Top             =   2918
         Width           =   1575
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
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
         Index           =   6
         Left            =   360
         TabIndex        =   23
         Tag             =   "XYF"
         Top             =   2558
         Width           =   1575
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
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
         Index           =   5
         Left            =   360
         TabIndex        =   22
         Tag             =   "XYF"
         Top             =   2198
         Width           =   1575
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "REASON:"
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
         Left            =   360
         TabIndex        =   21
         Tag             =   "XYF"
         Top             =   1838
         Width           =   1575
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "ENCODED BY:"
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
         Left            =   360
         TabIndex        =   20
         Tag             =   "XYF"
         Top             =   1478
         Width           =   1575
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "NEW VALUE:"
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
         Left            =   360
         TabIndex        =   19
         Tag             =   "XYF"
         Top             =   1118
         Width           =   1575
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "OLD VALUE:"
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
         Left            =   360
         TabIndex        =   18
         Tag             =   "XYF"
         Top             =   758
         Width           =   1575
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "SELECT FIELD TO EDIT:"
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
         TabIndex        =   16
         Tag             =   "XYF"
         Top             =   405
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1000
      Left            =   120
      TabIndex        =   2
      Tag             =   "XYF"
      Top             =   750
      Width           =   10935
      Begin VB.ComboBox cboItem 
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
         Left            =   2040
         TabIndex        =   14
         Tag             =   "XYF"
         Top             =   538
         Width           =   2895
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
         Left            =   9360
         TabIndex        =   5
         Tag             =   "XYF"
         Top             =   200
         Width           =   1455
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
         Left            =   9360
         TabIndex        =   4
         Tag             =   "XYF"
         Top             =   538
         Width           =   1455
      End
      Begin VB.TextBox txtAdjType 
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
         Left            =   2040
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   3
         Tag             =   "XYF"
         Top             =   200
         Width           =   2895
      End
      Begin VB.Label Label1 
         Caption         =   "ADJUSTMENT TYPE:"
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
         TabIndex        =   9
         Tag             =   "XYF"
         Top             =   230
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "SELECT ITEM TO ADJ.:"
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
         Top             =   570
         Width           =   1695
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "ADJUSTMENT NO."
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
         Left            =   7560
         TabIndex        =   7
         Tag             =   "XYF"
         Top             =   225
         Width           =   1695
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
         Left            =   7560
         TabIndex        =   6
         Tag             =   "XYF"
         Top             =   570
         Width           =   1695
      End
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   5
      Left            =   7500
      TabIndex        =   0
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
      Image           =   "ClericalADJ.frx":0B6C
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   6
      Left            =   9000
      TabIndex        =   1
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
      Image           =   "ClericalADJ.frx":1106
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   41
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
      Image           =   "ClericalADJ.frx":16A0
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   2
      Left            =   3000
      TabIndex        =   42
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
      Image           =   "ClericalADJ.frx":1C3A
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   3
      Left            =   4500
      TabIndex        =   43
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
      Image           =   "ClericalADJ.frx":21D4
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   4
      Left            =   6000
      TabIndex        =   44
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
      Image           =   "ClericalADJ.frx":276E
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   1
      Left            =   1500
      TabIndex        =   45
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
      Image           =   "ClericalADJ.frx":2D08
      cBack           =   -2147483633
   End
   Begin VB.Label lblHeading 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ADJUSTMENTS"
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
      Left            =   9375
      TabIndex        =   10
      Tag             =   "TRF"
      Top             =   480
      Width           =   1440
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
      Left            =   120
      TabIndex        =   11
      Tag             =   "TLF"
      Top             =   480
      Width           =   2895
   End
   Begin VB.Label lblHeading 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ADJUSTMENTS"
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
      Left            =   9405
      TabIndex        =   13
      Tag             =   "TRF"
      Top             =   510
      Width           =   1440
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
      Left            =   150
      TabIndex        =   12
      Tag             =   "TLF"
      Top             =   510
      Width           =   2895
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   360
      Top             =   5400
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   11400
      m_MDI_Height    =   5955
   End
End
Attribute VB_Name = "ClericalADJ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim BtnIndx As Integer
Dim AddNew As Boolean
Dim PrevItem As String

Public Sub DisplayClrADJ(MyRef As String)
    Dim rsClr As New ADODB.Recordset
    Dim Qry As String
    
    '***Access
    'Qry = "SELECT ADJCLERICAL.* From ADJCLERICAL WHERE (((ADJCLERICAL.ADJNO)='" & MyRef & "'));"
    'mySQL
    Qry = "SELECT ADJCLERICAL.* From ADJCLERICAL WHERE ADJCLERICAL.ADJNO='" & MyRef & "';"
    With rsClr
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            Call ClearEntries
            txtAdjType.Text = "CLERICAL"
            txtAdjNo.Text = !ADJNO
            txtAdjDate.Text = !ADJDATE
            cboItem.Text = !ITEMADJUSTED
            cboField.Text = IIf(IsNull(!FIELDADJUSTED), "", !FIELDADJUSTED)
            txtRefNo.Text = IIf(IsNull(!RefNo), "", !RefNo)
            If txtRefNo.Text <> "" Then
                txtRefNo.Visible = True
                lblRefNo.Visible = True
                If cboField.Text = "INVOICE DATE" Then lblRefNo.Caption = "INVOICE NO."
                If cboField.Text = "JOB ORDER" Then lblRefNo.Caption = "JOB ORDER NO."
                If cboField.Text = "PAYMENT TYPE" Then
                    lblRefNo.Top = lblInvType.Top
                    txtRefNo.Top = cboINVTYPE.Top
                End If
            Else
                txtRefNo.Visible = False
                lblRefNo.Visible = False
            End If
            txtOldValue.Text = !OLDVALUE
            txtNewValue.Text = !NEWVALUE
            txtEncoder.Text = IIf(IsNull(!ENCODER), "", !ENCODER)
            txtReason.Text = !REASON
            txtRemarks.Text = IIf(IsNull(!Remarks), "", !Remarks)
            txtPrepby.Text = !PreparedBy
            txtAprvby.Text = !APPROVEDBY
            txtPrnCopy.Text = IIf(IsNull(!PRINTCOPY), 0, !PRINTCOPY)
            cboOnCust.Text = IIf(IsNull(!CustomerName), "", !CustomerName)
            
            If txtRefNo.Text <> "" And cboItem.Text = "JOB ORDER" Then 'Display Charge to : Customer Name @ label2
                'Qry = "SELECT CUSTOMER.CUSTOMERNAME From CUSTOMER WHERE (((CUSTOMER.CUSTOMERID)=" & txtCustID.Text & "));"
                'mySQL
                Label2.Visible = True
                Qry = "SELECT CUSTOMERID FROM SRVCJOBORDER WHERE JOBORDERNO = '" & txtRefNo.Text & "' "
                txtCustID.Text = GetFieldValue(Qry)
                
                Qry = "SELECT CUSTOMER.CUSTOMERNAME From CUSTOMER WHERE CUSTOMER.CUSTOMERID=" & txtCustID.Text & ";"
                Label2.Caption = "Charge To: " & GetFieldValue(Qry)
            Else
                Label2.Caption = ""
            End If
        End If
        .Close
    End With
    If cboItem.Text = "PART NUMBER" Then
        
        '***Access
        'Qry = "SELECT SPAREPARTS.DESCRIPTION, SPAREPARTS.PARTNUMBER " & _
              "From SPAREPARTS WHERE (((SPAREPARTS.PARTNUMBER)='" & IIf(Val(txtPrnCopy.Text) > 0, txtNewValue.Text, txtOldValue.Text) & "'));"
        'mySQL
        Qry = "SELECT SPAREPARTS.DESCRIPTION, SPAREPARTS.PARTNUMBER " & _
              "From SPAREPARTS WHERE SPAREPARTS.PARTNUMBER='" & IIf(Val(txtPrnCopy.Text) > 0, txtNewValue.Text, txtOldValue.Text) & "';"
        lblDesc.Caption = GetFieldValue(Qry)
        '***Access
        'Qry = "SELECT SPAREPARTS.SPAREPARTID, SPAREPARTS.PARTNUMBER " & _
              "From SPAREPARTS WHERE (((SPAREPARTS.PARTNUMBER)='" & IIf(Val(txtPrnCopy.Text) > 0, txtNewValue.Text, txtOldValue.Text) & "'));"
        'mySQL
        Qry = "SELECT SPAREPARTS.SPAREPARTID, SPAREPARTS.PARTNUMBER " & _
              "From SPAREPARTS WHERE SPAREPARTS.PARTNUMBER='" & IIf(Val(txtPrnCopy.Text) > 0, txtNewValue.Text, txtOldValue.Text) & "';"
        txtSPID.Text = GetFieldValue(Qry)
    End If
    'If cboItem.Text = "JOB ORDER" And cboField.Text = "PAYMENT TYPE" Then Call txtrefno_LostFocus
End Sub

Private Sub ClearEntries()
    txtAdjType.Text = ""
    cboItem.Text = ""
    txtAdjNo.Text = ""
    txtAdjDate.Text = ""
    cboField.Text = ""
    txtRefNo.Text = ""
    txtOldValue.Text = ""
    txtNewValue.Text = ""
    txtEncoder.Text = ""
    txtReason.Text = ""
    txtRemarks.Text = ""
    txtAprvby.Text = ""
    txtPrepby.Text = ""
    txtSPID.Text = ""
    lblDesc.Caption = ""
    txtPrnCopy.Text = ""
    txtCustID.Text = ""
    txtJobDate.Text = ""
    txtJobAmt.Text = ""
    txtJobID.Text = ""
End Sub


Function GetAdjustmentNo() As String
    Dim rsAdjCntr As New ADODB.Recordset
    Dim Qry As String
    Dim MyNo As Long
    
    '***Access
    'Qry = "SELECT Max(ADJCNTR.ADJNO) AS MaxOfADJNO FROM ADJCNTR;"
    'mySQL
    Qry = "SELECT Max(round(ADJCNTR.ADJNO)) AS MaxOfADJNO FROM ADJCNTR;"
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

Function ValidateClericalAdj() As String
    Dim vMsg As String, Qry As String
    Dim i As Integer, Cntr As Integer
    
    vMsg = ""

    If Not IsDate(txtAdjDate.Text) Then vMsg = "Invalid Adjustment Date"
    If txtAdjNo.Text = "" Then
        vMsg = vMsg & vbCrLf & "Incomplete entry on field Adjustment No."
    Else
        If AddNew Then
            '***Access
            'Qry = "SELECT ADJCNTR.ADJNO From ADJCNTR WHERE (((ADJCNTR.ADJNO)=" & Val(txtAdjNo.Text) & "));"
            'mySQL
            Qry = "SELECT ADJCNTR.ADJNO From ADJCNTR WHERE ADJCNTR.ADJNO='" & txtAdjNo.Text & "';"
            If CtrlNumExist(Qry) Then
                vMsg = vMsg & vbCrLf & "Adjustment No. " & txtAdjNo.Text & " already exist!"
            End If
        End If
    End If
    Select Case cboItem.Text
        Case ""
            vMsg = vMsg & vbCrLf & "No selected item to Adjust"
        Case "PART NUMBER"
            If txtOldValue.Text = "" Then
                vMsg = vMsg & vbCrLf & "Incomplete entry on Old Value (Part Number)"
            Else
                If Val(txtSPID.Text) = 0 Then vMsg = vMsg & vbCrLf & "Invalid old value of Part Number"
            End If
            If txtNewValue.Text = "" Then
                vMsg = vMsg & vbCrLf & "Incomplete entry on New Value (Part Number)"
            Else
                '***Access
                'Qry = "SELECT SPAREPARTS.* From SPAREPARTS WHERE (((SPAREPARTS.PARTNUMBER)='" & txtNewValue.Text & "'));"
                'mySQL
                Qry = "SELECT SPAREPARTS.* From SPAREPARTS WHERE SPAREPARTS.PARTNUMBER='" & txtNewValue.Text & "';"
                If CtrlNumExist(Qry) Then vMsg = vMsg & "Cannot replace Part Number with " & txtNewValue.Text & " because it is already existing"
            End If
        Case "INVOICE"
            If cboField.Text <> "INVOICE NUMBER" And cboField.Text <> "INVOICE DATE" Then vMsg = vMsg & vbCrLf & "Invalid Value on Field to Edit"
            If txtOldValue.Text = "" Then vMsg = vMsg & vbCrLf & "Incomplete entry on field Old Value"
            If txtNewValue.Text = "" Then vMsg = vMsg & vbCrLf & "Incomplete entry on field New Value"
            If cboField.Text = "INVOICE DATE" And txtRefNo.Text = "" Then vMsg = vMsg & vbCrLf & "Incomplete entry on field Invoice Number"
        Case "JOB ORDER"
            If cboField.Text <> "PAYMENT TYPE" Then
                vMsg = vMsg & vbCrLf & "Invalid Value on Field to Edit"
            Else
                If txtRefNo.Text = "" Then vMsg = vMsg & vbCrLf & "No Reference numbere specified"
                If txtOldValue.Text = "CASH" And txtNewValue.Text <> "ON ACCOUNT" Then vMsg = vMsg & vbCrLf & "field New Value must only be 'ON ACCOUNT'"
                If txtOldValue.Text = "ON ACCOUNT" And txtNewValue.Text <> "CASH" Then vMsg = vMsg & vbCrLf & "field New Value must only be 'CASH'"
                If Val(txtCustID.Text) = 0 And txtNewValue.Text <> "ON ACCOUNT" Then vMsg = vMsg & vbCrLf & "Customer not valid for ON ACCOUNT!"
            End If
    End Select
    If txtReason.Text = "" Then vMsg = vMsg & vbCrLf & "Incomplete entry on field Reason"
    If txtAprvby.Text = "" Then vMsg = vMsg & vbCrLf & "Incomplete entry on field Approved By"
    If txtPrepby.Text = "" Then vMsg = vMsg & vbCrLf & "Incomplete entry on field Prepared By"
                                  
    ValidateClericalAdj = vMsg
End Function

Function ValidateSavingClericalAdj() As String
    Dim myCustID As Long
    Dim MyCustName As String, myErrMsg As String, Qry As String
    Dim myJOType As String
    
    MyCustName = ""
    myCustID = 0
    myErrMsg = ""
    myJOType = ""
    
    If Trim(txtNewValue.Text) = "ON ACCOUNT" Then
        Qry = "Select SRVCJOBORDER.JOTYPE FROM SRVCJOBORDER WHERE SRVCJOBORDER.JOBORDERNO='" & txtRefNo.Text & "';"
        myJOType = GetFieldValue(Qry)
        
        Qry = "Select SRVCJOBORDER.CUSTOMERID From SRVCJOBORDER WHERE SRVCJOBORDER.JOBORDERNO='" & txtRefNo.Text & "';"
        myCustID = Val(GetFieldValue(Qry))
        If Val(myCustID) = 0 Then
            Qry = "Select SRVCJOBORDER.CUSTOMER From SRVCJOBORDER WHERE SRVCJOBORDER.JOBORDERNO='" & txtRefNo.Text & "';"
            MyCustName = GetFieldValue(Qry)
            Qry = "Select CUSTOMER.CUSTOMERID FROM CUSTOMER WHERE CUSTOMER.CUSTOMERNAME = '" & MyCustName & "'"
            myCustID = Val(GetFieldValue(Qry))
            
            If Val(myCustID) = 0 Then
                myErrMsg = "Record Not Found for '" & MyCustName & "' on On-Account List!, Please Create an On-Account Only to proceed with this Adjustment"
            End If
        End If
    End If

    ValidateSavingClericalAdj = myErrMsg

End Function
Private Sub SaveClericalAdj()
    Dim rsClrAdj As New ADODB.Recordset
    Dim Qry As String, i As Integer
    
    'Saving Inventory Adjustment
    '***Access
    'With rsClrAdj
    '    If AddNew Then
    '        .Open "AdjClerical", LYRcon, adOpenKeyset, adLockOptimistic
    '        .AddNew
    '    Else
    '        Qry = "SELECT ADJCLERICAL.* From ADJCLERICAL WHERE ADJCLERICAL.ADJNO='" & cmdTrans(1).Tag & "';"
    '        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    '        .MoveFirst
    '    End If
    '    !AdjNo = txtAdjNo.Text
    '    !AdjDate = txtAdjDate.Text
    '    !ItemAdjusted = cboItem.Text
    '    !FieldAdjusted = cboField.Text
    '    !RefNo = txtRefNo.Text
    '    !OldValue = txtOldValue.Text
    '    !NewValue = txtNewValue.Text
    '    !Encoder = txtEncoder.Text
    '    !REASON = txtReason.Text
    '    !Remarks = txtRemarks.Text
    '    !PreparedBy = txtPrepBy.Text
    '    !ApprovedBy = txtAprvBy.Text
    '    !UPDATEDBY = UserID
    '    !LASTUPDATE = Now
    '    .Update
    '    .Close
    'End With
    'mySQL
    If AddNew Then
        LYRcon.Execute "INSERT INTO adjclerical( ADJNO, ADJDATE, ITEMADJUSTED, FIELDADJUSTED, REFNO, " & _
            "OLDVALUE, NEWVALUE, ENCODER, REASON, REMARKS, PREPAREDBY, APPROVEDBY, UPDATEDBY, LASTUPDATE, CUSTOMERNAME) " & _
            "VALUES('" & txtAdjNo.Text & "', '" & Format(txtAdjDate.Text, "yyyy-mm-dd") & "', " & _
            "'" & cboItem.Text & "', '" & cboField.Text & "', '" & txtRefNo.Text & "', '" & txtOldValue.Text & "', " & _
            "'" & txtNewValue.Text & "', '" & txtEncoder.Text & "', '" & txtReason.Text & "', " & _
            "'" & txtRemarks.Text & "', '" & txtPrepby.Text & "', '" & txtAprvby.Text & "', " & UserID & ", " & _
            "'" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "', '" & Trim(cboOnCust.Text) & "');"
        LYRcon.Execute "INSERT INTO ADJCNTR (ADJNO, ADJTYPE) VALUES (" & Val(txtAdjNo.Text) & ", 'CLERICAL');"
    Else
        LYRcon.Execute "UPDATE adjclerical SET ADJNO = '" & txtAdjNo.Text & "', ADJDATE = '" & Format(txtAdjDate.Text, "yyyy-mm-dd") & "', " & _
            "ITEMADJUSTED = '" & cboItem.Text & "', FIELDADJUSTED = '" & cboField.Text & "', REFNO = '" & txtRefNo.Text & "', " & _
            "OLDVALUE = '" & txtOldValue.Text & "', NEWVALUE = '" & txtNewValue.Text & "', ENCODER = '" & txtEncoder.Text & "', " & _
            "REASON = '" & txtReason.Text & "', REMARKS = '" & txtRemarks.Text & "', PREPAREDBY = '" & txtPrepby.Text & "', " & _
            "APPROVEDBY = '" & txtAprvby.Text & "', UPDATEDBY = " & UserID & ", LASTUPDATE = '" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "', CUSTOMERNAME = '" & Trim(cboOnCust.Text) & "' " & _
            "Where adjclerical.ADJNO = '" & cmdTrans(1).Tag & "';"
   End If
   'If AddNew Then
        '***Access
        'LYRcon.Execute "INSERT INTO ADJCNTR ( ADJNO, ADJTYPE) VALUES (" & Val(txtAdjNo.Text) & ", 'CLERICAL');"
   'End If
End Sub

Private Sub UpdateClericalAdj()
    Dim Qry As String, eMsg As String, MyRefType As String, ReqID As Integer, Bal As Currency
    Dim myCustID As Long, myAcctNO As Long, myJOType As String
    Dim MyCustName As String, myErrMsg As String
    Dim myPayID As Long, myAccNo As String
    
    On Error GoTo LocalErr
    Select Case cboItem
        Case "PART NUMBER"
            eMsg = "Line 168: Updating Spareparts Part Number"
            '***Access
            'LYRcon.Execute "UPDATE SPAREPARTS SET SPAREPARTS.PARTNUMBER = '" & txtNewValue.Text & "' " & _
                           "WHERE (((SPAREPARTS.PARTNUMBER)='" & txtOldValue.Text & "'));"
            'mySQL
            LYRcon.Execute "UPDATE SPAREPARTS SET SPAREPARTS.PARTNUMBER = '" & Trim(txtNewValue.Text) & "' " & _
                           "WHERE SPAREPARTS.PARTNUMBER='" & txtOldValue.Text & "';"
            eMsg = "Line 170: Updating Source SP Part Number"
            '***Access
            'LYRcon.Execute "UPDATE SOURCESP SET SOURCESP.PARTNUMBER = '" & txtNewValue.Text & "' " & _
                           "WHERE (((SOURCESP.PARTNUMBER)='" & txtOldValue.Text & "'));"
            'mySQL
            LYRcon.Execute "UPDATE SOURCESP SET SOURCESP.PARTNUMBER = '" & Trim(txtNewValue.Text) & "' " & _
                           "WHERE SOURCESP.PARTNUMBER='" & txtOldValue.Text & "';"
        Case "INVOICE"
            If cboField.Text = "INVOICE NUMBER" Then
                '***Access
                'Qry = "SELECT SPINVOICE.INVTYPE From SPINVOICE WHERE (((SPINVOICE.SPINVNO)='" & txtOldValue.Text & "'));"
                'mySQL
                Qry = "SELECT SPINVOICE.INVTYPE From SPINVOICE WHERE SPINVOICE.SPINVNO='" & Trim(txtOldValue.Text) & "' AND SPINVOICE.INVTYPE = '" & Trim(cboINVTYPE.Text) & "';"
                MyRefType = GetFieldValue(Qry)
                eMsg = "Line 177: Updating Invoice Number"
                'Access
                'LYRcon.Execute "UPDATE SPINVOICE SET SPINVOICE.SPINVNO = '" & txtNewValue.Text & "', SPINVOICE.TRANSNO = 'SI#-" & txtNewValue.Text & "' " & _
                               "WHERE (((SPINVOICE.SPINVNO)='" & txtOldValue.Text & "'));"
                'mySQL
                LYRcon.Execute "UPDATE SPINVOICE SET SPINVOICE.SPINVNO = '" & Trim(txtNewValue.Text) & "', SPINVOICE.TRANSNO = 'SI#-" & Trim(txtNewValue.Text) & "' " & _
                               "WHERE SPINVOICE.SPINVNO='" & Trim(txtOldValue.Text) & "' AND SPINVOICE.INVTYPE = '" & Trim(cboINVTYPE.Text) & "';"
                
                eMsg = "Line 180: Updating Reference of Transaction Register"
                '***Access
                'LYRcon.Execute "UPDATE SPITR SET SPITR.REFERENCE = 'SI#-" & txtNewValue.Text & "' " & _
                               "WHERE (((SPITR.REFERENCE)='SI#-" & txtOldValue.Text & "'));"
                'mySQL
                LYRcon.Execute "UPDATE SPITR SET SPITR.REFERENCE = 'SI#-" & txtNewValue.Text & "' " & _
                               "WHERE SPITR.REFERENCE='SI#-" & Trim(txtOldValue.Text) & "' AND SPITR.REMARKS = '" & Trim(cboINVTYPE.Text) & "';"
                
                If MyRefType = "CASH" Then
                    eMsg = "Line 180: Updating Invoice Number on table Payments"
                    '***Access
                    'LYRcon.Execute "UPDATE PAYMENTS SET PAYMENTS.INVOICENO = '" & txtNewValue.Text & "' WHERE (((PAYMENTS.INVOICENO)='" & txtOldValue.Text & "'));"
                    'mySQL
                    LYRcon.Execute "UPDATE PAYMENTS SET PAYMENTS.INVOICENO = '" & txtNewValue.Text & "' WHERE PAYMENTS.INVOICENO='" & txtOldValue.Text & "';"
                Else
                    eMsg = "Line 187: Updating Invoice Number on table On Account Terms"
                    '***Access
                    'LYRcon.Execute "UPDATE ONACCTERMS SET ONACCTERMS.SPINVNO = '" & txtNewValue.Text & "' WHERE (((ONACCTERMS.SPINVNO)='" & txtOldValue.Text & "'));"
                    'mysql
                    LYRcon.Execute "UPDATE ONACCTERMS SET ONACCTERMS.SPINVNO = '" & txtNewValue.Text & "' WHERE ONACCTERMS.SPINVNO='" & txtOldValue.Text & "';"
                    '***Access
                    'Qry = "SELECT LEDGER.SPINVNO From Ledger WHERE (((LEDGER.SPINVNO)='" & txtOldValue.Text & "') AND ((LEDGER.JOBID)=0));"
                    'mySQL
                    Qry = "SELECT LEDGER.SPINVNO From Ledger WHERE LEDGER.SPINVNO='" & txtOldValue.Text & "' AND LEDGER.JOBID=0;"
                    If CtrlNumExist(Qry) Then
                        eMsg = "Line 191: Updating Invoice Number on table Ledger"
                        '***Access
                        'LYRcon.Execute "UPDATE LEDGER SET LEDGER.SPINVNO = '" & txtNewValue.Text & "' " & _
                                "WHERE (((LEDGER.SPINVNO)='" & txtOldValue.Text & "') AND ((LEDGER.JOBID)=0));"
                        'mySQL
                        LYRcon.Execute "UPDATE LEDGER SET LEDGER.SPINVNO = '" & txtNewValue.Text & "' " & _
                                "WHERE LEDGER.SPINVNO='" & txtOldValue.Text & "' AND LEDGER.JOBID=0;"
                    End If
                End If
                '***Access
                'Qry = "SELECT SPINVOICE.REQID From SPINVOICE WHERE (((SPINVOICE.SPINVNO)='" & txtOldValue.Text & "'));"
                'mySQL
                Qry = "SELECT SPINVOICE.REQID From SPINVOICE WHERE SPINVOICE.SPINVNO='" & txtOldValue.Text & "' AND SPINVOICE.INVTYPE = '" & Trim(cboINVTYPE.Text) & "';"
                ReqID = Val(GetFieldValue(Qry))
                If ReqID > 0 Then
                    eMsg = "Line 199: Updating Invoice Number on table Requisition"
                    '***Access
                    'LYRcon.Execute "UPDATE REQUISITION SET REQUISITION.INVOICENO = '" & txtNewValue.Text & "' WHERE (((REQUISITION.REQID)=" & ReqID & "));"
                    'mysql
                    LYRcon.Execute "UPDATE REQUISITION SET REQUISITION.INVOICENO = '" & txtNewValue.Text & "' WHERE REQUISITION.REQID=" & ReqID & ";"
                End If
            ElseIf cboField.Text = "INVOICE DATE" Then
                '***Access
                'Qry = "SELECT SPINVOICE.INVTYPE From SPINVOICE WHERE (((SPINVOICE.SPINVNO)='" & txtRefNo.Text & "'));"
                'mySQL
                Qry = "SELECT SPINVOICE.INVTYPE From SPINVOICE WHERE SPINVOICE.SPINVNO='" & txtRefNo.Text & "' AND SPINVOICE.INVTYPE = '" & Trim(cboINVTYPE.Text) & "';"
                MyRefType = GetFieldValue(Qry)
                eMsg = "Line 205: Updating Invoice Date"
                '***Access
                'LYRcon.Execute "UPDATE SPINVOICE SET SPINVOICE.INVDATE = #" & txtNewValue.Text & "# WHERE (((SPINVOICE.SPINVNO)='" & txtRefNo.Text & "'));"
                'mySQL
                LYRcon.Execute "UPDATE SPINVOICE SET SPINVOICE.INVDATE = '" & Format(txtNewValue.Text, "yyyy-mm-dd") & "' WHERE SPINVOICE.SPINVNO='" & txtRefNo.Text & "' AND SPINVOICE.INVTYPE = '" & Trim(cboINVTYPE.Text) & "';"
                eMsg = "Line 207: Updating Date on Transaction Register"
                '***Access
                'LYRcon.Execute "UPDATE SPITR SET SPITR.TRANSDATE = #" & txtNewValue.Text & "# WHERE (((SPITR.REFERENCE)='SI#-" & txtRefNo.Text & "'));"
                'mySQL
                LYRcon.Execute "UPDATE SPITR SET SPITR.TRANSDATE = '" & Format(txtNewValue.Text, "yyyy-mm-dd") & "' WHERE SPITR.REFERENCE='SI#-" & txtRefNo.Text & "' AND SPITR.REMARKS = '" & Trim(cboINVTYPE.Text) & "';"
                If MyRefType = "CASH" Then
                    eMsg = "Line 210: Updating Invoice Date on table Payments"
                    '***Access
                    'LYRcon.Execute "UPDATE PAYMENTS SET PAYMENTS.PAYDATE= #" & txtNewValue.Text & "# WHERE (((PAYMENTS.INVOICENO)='" & txtRefNo.Text & "'));"
                    'mySQL
                    LYRcon.Execute "UPDATE PAYMENTS SET PAYMENTS.PAYDATE= '" & Format(txtNewValue.Text, "yyyy-mm-dd") & "' WHERE PAYMENTS.INVOICENO='" & txtRefNo.Text & "';"
                Else
                    eMsg = "Line 213: Updating Invoice Date on table On Account Terms"
                    '***Access
                    'LYRcon.Execute "UPDATE ONACCTERMS SET ONACCTERMS.TRANSDATE = #" & txtNewValue.Text & "# WHERE (((ONACCTERMS.SPINVNO)='" & txtRefNo.Text & "'));"
                    'mySQL
                    LYRcon.Execute "UPDATE ONACCTERMS SET ONACCTERMS.TRANSDATE = '" & Format(txtNewValue.Text, "yyyy-mm-dd") & "' WHERE ONACCTERMS.SPINVNO='" & txtRefNo.Text & "';"
                End If
            End If
        Case "JOB ORDER"
            If cboField.Text = "PAYMENT TYPE" Then
                eMsg = "Line 238: Updating Job Order field Payment Type"
                '***Access
                'LYRcon.Execute "UPDATE SRVCJOBORDER SET SRVCJOBORDER.PAYTYPE = '" & txtNewValue.Text & "' WHERE (((SRVCJOBORDER.JOBORDERNO)='" & txtRefNo.Text & "'));"
                'mySQL
                myCustID = 0
                myAcctNO = 0
                MyCustName = ""
                myJOType = ""
                If Trim(txtNewValue.Text) = "ON ACCOUNT" Then
                    Qry = "Select SRVCJOBORDER.REMARKS From SRVCJOBORDER WHERE SRVCJOBORDER.JOBORDERNO='" & txtRefNo.Text & "';"
                    myJOType = GetFieldValue(Qry)
                    
                    If InStr(myJOType, "FSC") > 0 Then
                        Qry = "Select CUSTOMERID FROM CUSTOMER WHERE CUSTOMERNAME = '" & Trim(cboOnCust.Text) & "'"
                        myCustID = Val(GetFieldValue(Qry))
                        LYRcon.Execute "UPDATE SRVCJOBORDER SET SRVCJOBORDER.CUSTOMERID = " & Val(myCustID) & " WHERE SRVCJOBORDER.JOBORDERNO='" & txtRefNo.Text & "'; "
                        
                        Qry = "SELECT SRVCJOBORDER.ACCOUNTNO FROM SRVCJOBORDER WHERE SRVCJOBORDER.JOBORDERNO='" & txtRefNo.Text & "';"
                        myAcctNO = Val(GetFieldValue(Qry))
                        If myAcctNO > 0 Then
                            LYRcon.Execute "UPDATE SRVCUSTOMER SET SRVCUSTOMER.CUSTOMERID = " & Val(myCustID) & " WHERE SRVCUSTOMER.ACCOUNTNO =" & Val(myAcctNO) & "; "
                        End If
                    Else
                        Qry = "Select SRVCJOBORDER.CUSTOMERID From SRVCJOBORDER WHERE SRVCJOBORDER.JOBORDERNO='" & txtRefNo.Text & "';"
'                        MyCustName = GetFieldValue(Qry)
'
'                        Qry = "Select CUSTOMER.CUSTOMERID FROM CUSTOMER WHERE CUSTOMER.CUSTOMERNAME = '" & MyCustName & "'"
                        myCustID = Val(GetFieldValue(Qry))
                        LYRcon.Execute "UPDATE SRVCJOBORDER SET SRVCJOBORDER.CUSTOMERID = " & Val(myCustID) & " WHERE SRVCJOBORDER.JOBORDERNO='" & txtRefNo.Text & "'; "
                        
                        Qry = "SELECT SRVCJOBORDER.ACCOUNTNO FROM SRVCJOBORDER WHERE SRVCJOBORDER.JOBORDERNO='" & txtRefNo.Text & "';"
                        myAcctNO = Val(GetFieldValue(Qry))
                        If myAcctNO > 0 Then
                            LYRcon.Execute "UPDATE SRVCUSTOMER SET SRVCUSTOMER.CUSTOMERID = " & Val(myCustID) & " WHERE SRVCUSTOMER.ACCOUNTNO =" & Val(myAcctNO) & "; "
                        End If
                    End If
                End If
                LYRcon.Execute "UPDATE SRVCJOBORDER SET SRVCJOBORDER.PAYTYPE = '" & txtNewValue.Text & "' WHERE SRVCJOBORDER.JOBORDERNO='" & txtRefNo.Text & "';"
                Call UpdateJOBonAcct(myCustID)
                'Update Customer Outstanding Balance
                '***Access
                'Qry = "SELECT Sum(ONACCTERMS.BALANCE) AS SumOfBALANCE From ONACCTERMS WHERE (((ONACCTERMS.CUSTOMERID)=" & txtCustID.Text & ") AND ((ONACCTERMS.STATUS)='OPEN'));"
                'mySQL
                If Trim(txtNewValue.Text) = "ON ACCOUNT" Then
                    Qry = "SELECT Sum(ONACCTERMS.BALANCE) AS SumOfBALANCE From ONACCTERMS WHERE ONACCTERMS.CUSTOMERID=" & Val(myCustID) & " AND ONACCTERMS.STATUS='OPEN';"
                    Bal = Val(GetFieldValue(Qry))
                    LYRcon.Execute "UPDATE CUSTOMER SET CUSTOMER.OSBAL = " & Bal & _
                                   " WHERE CUSTOMER.CUSTOMERID=" & Val(myCustID) & ";"
                Else
                    Qry = "SELECT Sum(ONACCTERMS.BALANCE) AS SumOfBALANCE From ONACCTERMS WHERE ONACCTERMS.CUSTOMERID=" & Val(txtCustID.Text) & " AND ONACCTERMS.STATUS='OPEN';"
                    Bal = Val(GetFieldValue(Qry))
                    LYRcon.Execute "UPDATE CUSTOMER SET CUSTOMER.OSBAL = " & Bal & _
                                   " WHERE CUSTOMER.CUSTOMERID=" & Val(txtCustID.Text) & ";"
                End If
                '***Access
                'LYRcon.Execute "UPDATE CUSTOMER SET CUSTOMER.OSBAL = " & Bal & _
                               " WHERE (((CUSTOMER.CUSTOMERID)=" & txtCustID.Text & "));"
                'mySQL

            End If
        Case "OFFICIAL RECEIPT"
            If cboField.Text = "O.R NUMBER" Then
                Qry = "call sp_mrcretails.sp_ClericalAdj_OR('" & sysBranchCode & "','" & Trim(txtOldValue.Text) & "','" & Trim(txtNewValue.Text) & "')"
                With rsAnyUse
                    If .State = adStateOpen Then .Close
                    .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
                    If .RecordCount > 0 Then
                        .MoveFirst
                        If !level = "Error" Then
                            MsgBox " " & !message & " " ', vbInformation + vbOKOnly, "Validation"
                            Exit Sub
                        End If
                    End If
                    .Close
                End With
            End If
        Case "COLLECTION RECEIPT"
            If cboField.Text = "C.R NUMBER" Then
                Qry = "call sp_mrcretails.sp_ClericalAdj_CR('" & sysBranchCode & "','" & Trim(txtOldValue.Text) & "','" & Trim(txtNewValue.Text) & "')"
                With rsAnyUse
                    If .State = adStateOpen Then .Close
                    .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
                    If .RecordCount > 0 Then
                        .MoveFirst
                        If !level = "Error" Then
                            MsgBox " " & !message & " ", vbInformation + vbOKOnly, "Validation"
                            Exit Sub
                        End If
                    End If
                    .Close
                End With
            End If
    End Select
    Exit Sub
LocalErr:
    MsgBox eMsg & vbCrLf & Err.Number & ": " & Err.Description, vbCritical + vbOKOnly, "Error"
End Sub

Private Sub UpdateJOBonAcct(myCustID As Long)
    Dim rsOnAccount As New ADODB.Recordset
    Dim rsTerm As New ADODB.Recordset
    Dim nPay As Integer, nTrm As Integer
    Dim Qry As String, nSkm As String
    
    'Remove all On account record of Job Order
    '***Access
    'LYRcon.Execute "DELETE ONACCTERMS.* From ONACCTERMS WHERE (((ONACCTERMS.CUSTOMERID)=" & txtCustID.Text & ") AND ((ONACCTERMS.JOBORDERNO)='" & txtRefNo.Text & "') AND ((ONACCTERMS.SPINVNO) Is Null));"
    'mySQL
    LYRcon.Execute "DELETE From ONACCTERMS WHERE (((ONACCTERMS.CUSTOMERID)=" & Val(myCustID) & ") AND ((ONACCTERMS.JOBORDERNO)='" & txtRefNo.Text & "') AND ((ONACCTERMS.SPINVNO) Is Null));"

    If txtNewValue.Text = "ON ACCOUNT" Then
        '***Access
        'Qry = "SELECT ONACCTERMS.* From ONACCTERMS " & _
              "WHERE (((ONACCTERMS.STATUS)='Default') AND ((ONACCTERMS.CUSTOMERID)=" & txtCustID.Text & "));"
        ''mySQL
        Qry = "SELECT ONACCTERMS.* From ONACCTERMS " & _
              "WHERE ONACCTERMS.STATUS='Default' AND ONACCTERMS.CUSTOMERID=" & Val(myCustID) & ";"
        With rsTerm
            .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
            If .RecordCount > 0 Then
                .MoveFirst
                nTrm = !Terms
                nSkm = !PayScheme
                nPay = !Npays
            Else
                nTrm = 30
                nSkm = "MONTHLY"
                nPay = 1
            End If
            .Close
        End With
        
        'With rsOnAccount
        '    .Open "ONACCTERMS", LYRcon, adOpenKeyset, adLockOptimistic
        '    .AddNew
        '    !CUSTOMERID = txtCustID.Text
        '    !PayScheme = nSkm
        '    !Terms = nTrm
        '    !Npays = nPay
        '    If IsDate(txtJobDate.Text) Then
        '        !PayStart = DateAdd("M", 1, DateValue(txtJobDate.Text))
        '        !TRANSDATE = DateValue(txtJobDate.Text)
        '    Else
        '        !PayStart = DateAdd("M", 1, Date)
        '        !TRANSDATE = Date
        '    End If
        '    !BILLINGAMT = Val(txtJobAmt.Text) / nPay
        '    !JOBORDERNO = txtRefNo.Text
        '    !JOBORDERID = Val(txtJobID.Text)
        '    !BALANCE = Val(txtJobAmt.Text)
        '    !AcctAmount = Val(txtJobAmt.Text)
        '    !Status = "OPEN"
        '    !LASTUPDATE = Now
        '    .Update
        '    .Close
        'end with
        
        'mySQL
        If IsDate(txtJobDate.Text) Then
            AdjPayStart = DateAdd("M", 1, DateValue(txtJobDate.Text))
            AdjTRANSDATE = DateValue(txtJobDate.Text)
        Else
            AdjPayStart = DateAdd("M", 1, Date)
            AdjTRANSDATE = Date
        End If

        If Val(txtJobAmt.Text) > 0 Then
            LYRcon.Execute "Insert Into onaccterms(CUSTOMERID, PAYSCHEME, TERMS, NPAYS, PAYSTART, TRANSDATE, BILLINGAMT, JOBORDERNO, JOBORDERID, " & _
                "BALANCE, ACCTAMOUNT, `STATUS`, LASTUPDATE)VALUES(" & Val(myCustID) & ", '" & nSkm & "', " & nTrm & ", " & nPay & ", " & _
                "'" & Format(AdjPayStart, "yyyy-mm-dd") & "', '" & Format(AdjTRANSDATE, "yyyy-mm-dd") & "', " & Val(txtJobAmt.Text) / nPay & ", " & _
                "'" & txtRefNo.Text & "', " & Val(txtJobID.Text) & ", " & Val(txtJobAmt.Text) & ", " & Val(txtJobAmt.Text) & ", 'OPEN', " & _
                "'" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "');"
        Else
            LYRcon.Execute "Insert Into onaccterms(CUSTOMERID, PAYSCHEME, TERMS, NPAYS, PAYSTART, TRANSDATE, BILLINGAMT, JOBORDERNO, JOBORDERID, " & _
                "BALANCE, ACCTAMOUNT, `STATUS`, LASTUPDATE)VALUES(" & Val(myCustID) & ", '" & nSkm & "', " & nTrm & ", " & nPay & ", " & _
                "'" & Format(AdjPayStart, "yyyy-mm-dd") & "', '" & Format(AdjTRANSDATE, "yyyy-mm-dd") & "', " & Val(txtJobAmt.Text) / nPay & ", " & _
                "'" & txtRefNo.Text & "', " & Val(txtJobID.Text) & ", " & Val(txtJobAmt.Text) & ", " & Val(txtJobAmt.Text) & ", 'CLOSED', " & _
                "'" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "');"
        End If
    End If
End Sub
Private Sub cboField_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        If cboItem.Text = "INVOICE" Then
            If cboField.Text = "INVOICE NUMBER" Then
                txtOldValue.SetFocus
            Else
                txtRefNo.Visible = True
                txtRefNo.SetFocus
            End If
        ElseIf cboItem.Text = "JOB ORDER" Then
            txtRefNo.Visible = True
            txtRefNo.SetFocus
        End If
    End If
End Sub

Private Sub cboField_LostFocus()
    If cboItem.Text = "INVOICE" Then
        If cboField.Text = "INVOICE NUMBER" Then
            lblRefNo.Visible = False
            txtRefNo.Visible = False
            txtRefNo.Text = ""
            txtOldValue.Locked = False
            cboOnCust.Visible = False
            lblCustomer.Visible = False
            lblInvType.Visible = True
            cboINVTYPE.Visible = True
        Else
            lblRefNo.Top = lblCustomer.Top
            txtRefNo.Top = cboOnCust.Top
            lblRefNo.Visible = True
            txtRefNo.Visible = True
            txtOldValue.Locked = True
            cboOnCust.Visible = False
            lblCustomer.Visible = False
            lblInvType.Visible = True
            cboINVTYPE.Visible = True
        End If
    ElseIf cboItem.Text = "JOB ORDER" Then
        If cboField.Text = "PAYMENT TYPE" Then
            lblRefNo.Visible = True
            txtRefNo.Visible = True
            txtRefNo.Text = ""
            lblRefNo.Top = lblInvType.Top
            txtRefNo.Top = cboINVTYPE.Top
            txtOldValue.Locked = True
            cboOnCust.Visible = True
            lblCustomer.Visible = True
            lblInvType.Visible = False
            cboINVTYPE.Visible = False
        End If
    ElseIf cboItem.Text = "OFFICIAL RECEIPT" Then
        If cboField.Text = "O.R NUMBER" Then
            lblRefNo.Visible = False
            txtRefNo.Visible = False
            txtRefNo.Text = ""
            txtOldValue.Locked = False
            cboOnCust.Visible = False
            lblCustomer.Visible = False
            lblInvType.Visible = False
            cboINVTYPE.Visible = False
        End If
    ElseIf cboItem.Text = "COLLECTION RECEIPT" Then
        If cboField.Text = "C.R NUMBER" Then
            lblRefNo.Visible = False
            txtRefNo.Visible = False
            txtRefNo.Text = ""
            txtOldValue.Locked = False
            cboOnCust.Visible = False
            lblCustomer.Visible = False
            lblInvType.Visible = False
            cboINVTYPE.Visible = False
        End If
    End If
End Sub



Private Sub cboItem_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtAdjNo.SetFocus
End Sub

Private Sub cboItem_LostFocus()
    If PrevItem = cboItem.Text Then Exit Sub
    If cboItem.Text = "PART NUMBER" Then
        lbl(0).ForeColor = QBColor(8)
        cboField.Text = ""
        cboField.Enabled = False
    ElseIf cboItem.Text = "INVOICE" Then
        lbl(0).ForeColor = vbBlack
        cboField.Text = ""
        cboField.Enabled = True
        cboField.Clear
        cboField.AddItem "INVOICE NUMBER"
        cboField.AddItem "INVOICE DATE"
    ElseIf cboItem.Text = "JOB ORDER" Then
        lbl(0).ForeColor = vbBlack
        cboField.Text = ""
        cboField.Enabled = True
        cboField.Clear
        cboField.AddItem "PAYMENT TYPE"
    ElseIf cboItem.Text = "OFFICIAL RECEIPT" Then
        lbl(0).ForeColor = vbBlack
        cboField.Text = ""
        cboField.Enabled = True
        cboField.Clear
        cboField.AddItem "O.R NUMBER"
    ElseIf cboItem.Text = "COLLECTION RECEIPT" Then
        lbl(0).ForeColor = vbBlack
        cboField.Text = ""
        cboField.Enabled = True
        cboField.Clear
        cboField.AddItem "C.R NUMBER"
    Else
        lbl(0).ForeColor = vbBlack
        cboField.Text = ""
        cboField.Enabled = True
    End If
    PrevItem = cboItem.Text
End Sub

Private Sub cmdCancelt_Click()
    Frame1.Enabled = True
    Frame2.Enabled = True
    cmdTrans(5).Enabled = True
    cmdTrans(6).Enabled = True
    tabParts.Top = 8000
    tabParts.Visible = False
End Sub

Private Sub cmdOKt_Click()
    Dim r As Integer
    
    With flxParts
        r = IIf(.Row = 0, 1, .Row)
        txtOldValue.Text = .TextMatrix(r, 1)
        lblDesc.Caption = .TextMatrix(r, 2)
        txtSPID.Text = .TextMatrix(r, 4)
    End With
    Frame1.Enabled = True
    Frame2.Enabled = True
    cmdTrans(5).Enabled = True
    cmdTrans(6).Enabled = True
    tabParts.Top = 8000
    tabParts.Visible = False
End Sub

Private Sub cmdTrans_Click(Index As Integer)
    Dim Ans As Integer, i As Integer
    Dim myMsg As String, PrnVal(0 To 18) As String
    Dim flx As MSFlexGrid
    Dim Qry As String, MyAcc As String
    
    Select Case Index
        Case 0  'New
            '***Access
            'Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE (((USERACCESS.EMPID)=" & UserID & ") AND ((USERACCESS.MENUINDX)=53));"
            'mySQL
            Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE USERACCESS.EMPID=" & UserID & " AND USERACCESS.MENUINDX=53;"
            
            MyAcc = GetFieldValue(Qry)
            If MyAcc <> "FULL ACCESS" Then
                MsgBox "You don't have full access to this transaction." & vbCrLf & "Please contact your System Administrator", vbExclamation + vbOKOnly, "Access Denied"
                Exit Sub
            End If
        
            BtnIndx = 4
            tmrFloat.Enabled = True
            Call ClearEntries
            Frame1.Enabled = True
            Frame2.Enabled = True
            txtAdjType.Text = "CLERICAL"
            txtAdjDate.Text = Date
            AddNew = True
            PrevItem = ""
            txtAdjNo.Text = GetAdjustmentNo
            txtPrepby.Text = GetUserName
            cboItem.SetFocus
        Case 1  'Edit
            '***Access
            'Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE (((USERACCESS.EMPID)=" & UserID & ") AND ((USERACCESS.MENUINDX)=53));"
            'mySQL
            Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE USERACCESS.EMPID=" & UserID & " AND USERACCESS.MENUINDX=53;"
            MyAcc = GetFieldValue(Qry)
            If MyAcc <> "FULL ACCESS" Then
                MsgBox "You don't have full access to this transaction." & vbCrLf & "Please contact your System Administrator", vbExclamation + vbOKOnly, "Access Denied"
                Exit Sub
            End If
            If Val(txtPrnCopy.Text) > 0 Then
                MsgBox "Editing is not allowed because adjustment was already printed.", vbInformation + vbOKOnly, "Edit Denied"
                Exit Sub
            ElseIf Val(txtAdjNo.Text) = 0 Then
                MsgBox "No available Adjustments to Edit", vbInformation + vbOKOnly, "Confirm P.O."
                Exit Sub
            End If
            BtnIndx = 4
            tmrFloat.Enabled = True
            AddNew = False
            cmdTrans(1).Tag = txtAdjNo.Text
            Frame1.Enabled = True
            Frame2.Enabled = True
        Case 2  'Summary
            Unload Me
            FrmSummary.mySrc = "Clerical Adjustment"
            FrmSummary.Show
            'ClrAdjSumm.Show
        Case 3  'Print
            Ans = MsgBox("Do you want to Print Clerical Adjustment?", vbQuestion + vbYesNo, "Print")
            If Ans = vbNo Then Exit Sub
            myMsg = ""
            If Val(txtPrnCopy.Text) = 0 Then
                myMsg = ValidateSavingClericalAdj
                If myMsg <> "" Then
                    MsgBox myMsg, vbCritical + vbOKOnly, "Validation Error"
                    Exit Sub
                End If
                txtPrnCopy.Text = Val(txtPrnCopy.Text) + 1
                Call UpdateClericalAdj
            Else
                txtPrnCopy.Text = Val(txtPrnCopy.Text) + 1
            End If
            
            PrnVal(0) = "TYPE: CLERICAL ADJUSTMENT"
            PrnVal(1) = "DATE: " & txtAdjDate.Text
            PrnVal(2) = "ITEM ADJUSTED: " & cboItem.Text
            If cboItem.Text = "PART NUMBER" Then
                PrnVal(3) = ""
                PrnVal(4) = ""
                'PrnVal(5) = ""
            ElseIf cboItem.Text = "INVOICE" Then
                PrnVal(3) = "FIELD TO EDIT: " & cboField.Text
                If cboField.Text = "INVOICE DATE" Then
                    PrnVal(4) = "INV. NO. " & txtRefNo.Text
                    'PrnVal(5) = ""
                Else
                    PrnVal(4) = ""
                    'PrnVal(5) = ""
                End If
            ElseIf cboItem.Text = "JOB ORDER" Then
                PrnVal(3) = "FIELD TO EDIT: " & cboField.Text
                PrnVal(4) = "JOB ORDER NO. " & txtRefNo.Text
                'PrnVal(5) = ""
            End If
            PrnVal(5) = "OLD VALUE:"
            PrnVal(6) = txtOldValue.Text
            PrnVal(7) = "NEW VALUE:"
            PrnVal(8) = txtNewValue.Text
            PrnVal(9) = "ENCODED BY:"
            PrnVal(10) = txtEncoder.Text
            PrnVal(11) = "REASON:"
            PrnVal(12) = txtReason.Text
            If txtRemarks.Text <> "" Then
                PrnVal(13) = "REMARKS: " & txtRemarks.Text
            Else
                PrnVal(13) = ""
            End If
            If cboItem.Text = "JOB ORDER" Then PrnVal(13) = "REMARKS: " & Label2.Caption & " " & txtRemarks.Text
            PrnVal(14) = "PREPARED BY:"
            PrnVal(15) = "APPROVED BY:"
            PrnVal(16) = txtPrepby.Text
            PrnVal(17) = txtAprvby.Text
            Call PrintMyForm(PrnVal, "CLERICAL ADJ", flx)
            '***Access
            'LYRcon.Execute "UPDATE ADJCLERICAL SET ADJCLERICAL.PRINTCOPY = " & Val(txtPrnCopy.Text) & " WHERE (((ADJCLERICAL.ADJNO)='" & txtAdjNo.Text & "'));"
            'mySQL
            LYRcon.Execute "UPDATE ADJCLERICAL SET ADJCLERICAL.PRINTCOPY = " & Val(txtPrnCopy.Text) & " WHERE ADJCLERICAL.ADJNO='" & txtAdjNo.Text & "';"
        Case 4  'Exit
            Unload Me
            Exit Sub
        Case 5  'Save
            myMsg = ""
            myMsg = ValidateClericalAdj
            If myMsg <> "" Then
                MsgBox myMsg, vbCritical + vbOKOnly, "Validation Error"
                Exit Sub
            End If
            Ans = MsgBox("Do you want to Save Adjustment?", vbQuestion + vbYesNo, "Save")
            If Ans = vbNo Then Exit Sub
            Call SaveClericalAdj
            MsgBox "Adjustment Successfully Saved!", vbInformation + vbOKOnly, "Save"
                
            BtnIndx = 6
            tmrFloatRev.Enabled = True
            Frame1.Enabled = False
            Frame2.Enabled = False
            
        Case 6  'Cancel
            Ans = MsgBox("Do you want to Cancel Adjustment?", vbQuestion + vbYesNo, "Cancel")
            If Ans = vbNo Then Exit Sub
            BtnIndx = 6
            tmrFloatRev.Enabled = True
            If AddNew Then
                Call ClearEntries
            Else
                'Call DisplaySalesReturn(txtReturnID.Text)
            End If
            Frame1.Enabled = False
            Frame2.Enabled = False
    End Select

End Sub

Private Sub Form_Load()
    
    Frame1.Enabled = False
    Frame2.Enabled = False
    cboItem.AddItem "INVOICE"
    'cboItem.AddItem "PART NUMBER"
    cboItem.AddItem "JOB ORDER"
    cboItem.AddItem "OFFICIAL RECEIPT"
    cboItem.AddItem "COLLECTION RECEIPT"
    
    MotoRep.sBar.Panels(6).Text = "Clerical Adjustment Form"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    MotoRep.sBar.Panels(6).Text = ""
End Sub

Private Sub tmrDrop_Timer()
    cmdTrans(BtnIndx).Top = cmdTrans(BtnIndx).Top + 50
    If cmdTrans(BtnIndx).Top >= 0 Then
        cmdTrans(BtnIndx).Top = 0
        BtnIndx = BtnIndx + 1
        If BtnIndx = 7 Then
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
            BtnIndx = 5 'Index of Button Save
            cmdTrans(5).Top = -375
            cmdTrans(5).Left = 0
            cmdTrans(5).Visible = True
            cmdTrans(6).Top = -375
            cmdTrans(6).Left = cmdTrans(4).Width
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
        If BtnIndx = 4 Then 'Index of Button Exit
            tmrFloatRev.Enabled = False
            BtnIndx = 0
            For i = 0 To 4
                cmdTrans(i).Visible = True
            Next i
            tmrDropRev.Enabled = True
        End If
    End If
End Sub

Private Sub txtAdjDate_KeyPress(KeyAscii As Integer)
    Dim strValid As String
    
    strValid = "0123456789-/"
    If KeyAscii > 26 Then
        If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
    If KeyAscii = 13 Then
        If cboItem.Text = "PART NUMBER" Then
            txtOldValue.SetFocus
        Else
            cboField.SetFocus
        End If
    End If
End Sub

Private Sub txtAdjNo_GotFocus()
    If txtAdjNo.Text = "" Then
        txtAdjNo.SelStart = 0
        txtAdjNo.SelLength = Len(txtAdjNo.Text)
    End If
End Sub

Private Sub txtAdjNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtAdjDate.SetFocus
End Sub

Private Sub txtaprvby_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtPrepby.SetFocus
End Sub

Private Sub txtEncoder_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtReason.SetFocus
End Sub

Private Sub txtNewValue_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtEncoder.SetFocus
End Sub

Private Sub txtNewValue_LostFocus()
    Dim Q As String
    
    If txtNewValue.Text <> "" Then
        Select Case cboItem.Text
            Case "PART NUMBER"
                '***Access
                'Q = "SELECT SPAREPARTS.* From SPAREPARTS WHERE (((SPAREPARTS.PARTNUMBER)='" & txtNewValue.Text & "'));"
                'mySQL
                Q = "SELECT SPAREPARTS.* From SPAREPARTS WHERE SPAREPARTS.PARTNUMBER='" & txtNewValue.Text & "';"
                If CtrlNumExist(Q) Then
                    MsgBox "Part Number " & txtNewValue.Text & " already exist!" & vbCrLf & "Cannot Replace an existing Part Number!", vbCritical + vbOKOnly, "Spare Part Adjustment"
                    txtNewValue.Text = ""
                End If
            Case "INVOICE"
                If cboField.Text = "INVOICE NUMBER" Then
                    '***Access
                    'Q = "SELECT SPINVOICE.SPINVNO From SPINVOICE WHERE (((SPINVOICE.SPINVNO)='" & txtNewValue.Text & "'));"
                    'mySQL
                    Q = "SELECT SPINVOICE.SPINVNO From SPINVOICE WHERE SPINVOICE.SPINVNO='" & txtNewValue.Text & "';"
                    If CtrlNumExist(Q) Then
                        MsgBox "Invoice No. " & txtNewValue.Text & " already exist!", vbExclamation + vbOKOnly, "Invoice"
                        txtNewValue.Text = ""
                        txtNewValue.SetFocus
                    End If
                Else
                    If Not IsDate(txtNewValue.Text) Then
                        MsgBox "Invalid New Invoice Date value!", vbExclamation + vbOKOnly, "Invoice"
                        txtNewValue.Text = ""
                        txtNewValue.SetFocus
                    End If
                End If
        End Select
    End If
End Sub

Private Sub txtOldValue_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtNewValue.SetFocus
End Sub

Private Sub DisplayPartNumber()
    Dim rsSP As New ADODB.Recordset
    Dim Qry As String, M1 As Boolean
    
    '***Access
    'Qry = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.SPAREPARTID, SPAREPARTS.DESCRIPTION, SPAREPARTS.STOREPRICE " & _
          "From SPAREPARTS WHERE (((SPAREPARTS.PARTNUMBER)='" & txtOldValue.Text & "'));"
    'mySQL
    Qry = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.SPAREPARTID, SPAREPARTS.DESCRIPTION, SPAREPARTS.STOREPRICE " & _
          "From SPAREPARTS WHERE SPAREPARTS.PARTNUMBER='" & txtOldValue.Text & "';"
    
    M1 = False
    With rsSP
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount = 1 Then
            .MoveFirst
            txtSPID.Text = !SparepartID
            lblDesc.Caption = !Description
        ElseIf .RecordCount > 1 Then
            .MoveFirst
            flxParts.Rows = 1
            Do While Not .EOF
                flxParts.AddItem vbTab & !PartNumber & vbTab & !Description & vbTab & Format(!RETAILPRICE, "#,##0.00") & vbTab & !SparepartID
                .MoveNext
            Loop
            flxParts.ColWidth(4) = 0
            M1 = True
        Else
            MsgBox "Part Number " & txtOldValue.Text & " does not exist!", vbExclamation + vbOKOnly, "Edit Part Number"
            txtOldValue.Text = ""
            txtSPID.Text = ""
            lblDesc.Caption = ""
        End If
        .Close
    End With
    If M1 Then
        Frame1.Enabled = False
        Frame2.Enabled = False
        cmdTrans(5).Enabled = False
        cmdTrans(6).Enabled = False
        With tabParts
            .Move (Me.Width - .Width) \ 2, (Me.Height - .Height) \ 2
            .Visible = True
        End With
    End If
End Sub

Private Sub txtOldValue_LostFocus()
    Dim Qry As String
    
    If txtOldValue.Text <> "" Then
        Select Case cboItem.Text
            Case "PART NUMBER"
                Call DisplayPartNumber
            Case "INVOICE"
                If cboField.Text = "INVOICE NUMBER" Then
                    '***Access
                    'Qry = "SELECT SPINVOICE.SPINVNO From SPINVOICE WHERE (((SPINVOICE.SPINVNO)='" & txtOldValue.Text & "'));"
                    'mySQL
                    Qry = "SELECT SPINVOICE.SPINVNO From SPINVOICE WHERE SPINVOICE.SPINVNO='" & txtOldValue.Text & "';"
                    If Not CtrlNumExist(Qry) Then
                        MsgBox "Invoice No. " & txtOldValue.Text & " not found!", vbExclamation + vbOKOnly, "Invoice No."
                        txtOldValue.Text = ""
                        txtOldValue.SetFocus
                    Else
                        '***Access
                        'Qry = "SELECT [LASTNAME] & ', ' & [FIRSTNAME] & IIf(IsNull([MIDDLENAME]),'',' ' & Left$([MIDDLENAME],1) & '.') AS EMPNAME " & _
                              "FROM SPINVOICE INNER JOIN EMPLOYEES ON SPINVOICE.EMPID = EMPLOYEES.EMPID WHERE (((SPINVOICE.SPINVNO)='" & txtOldValue.Text & "'));"
                        'mySQL
                        Qry = "SELECT concat(EMPLOYEES.LASTNAME,', ',EMPLOYEES.FIRSTNAME,' ',IF(ISNULL(LEFT(EMPLOYEES.MIDDLENAME,1)),'',CONCAT(LEFT(EMPLOYEES.MIDDLENAME,1), '.'))) AS EMPNAME " & _
                              "FROM SPINVOICE INNER JOIN EMPLOYEES ON SPINVOICE.EMPID = EMPLOYEES.EMPID WHERE SPINVOICE.SPINVNO='" & txtOldValue.Text & "';"
                        txtEncoder.Text = GetFieldValue(Qry)
                    End If
                End If
            Case "OFFICIAL RECEIPT"
                If cboField.Text = "O.R NUMBER" Then
                    Qry = "SELECT PAYMENTS.ORNO FROM PAYMENTS WHERE PAYMENTS.ORNO = '" & txtOldValue.Text & "' ORDER BY PAYMENTS.PAYDATE DESC LIMIT 1"
                    If Not CtrlNumExist(Qry) Then
                        MsgBox "O.R No. " & txtOldValue.Text & " not found!", vbExclamation + vbOKOnly, "Official Receipt"
                        txtOldValue.Text = ""
                        txtOldValue.SetFocus
                    Else
                        Qry = "SELECT concat(EMPLOYEES.LASTNAME,', ',EMPLOYEES.FIRSTNAME,' ',IF(ISNULL(LEFT(EMPLOYEES.MIDDLENAME,1)),'',CONCAT(LEFT(EMPLOYEES.MIDDLENAME,1), '.'))) AS EMPNAME " & _
                              "FROM PAYMENTS INNER JOIN EMPLOYEES ON PAYMENTS.UPDATEDBY = EMPLOYEES.EMPID WHERE PAYMENTS.ORNO='" & txtOldValue.Text & "';"
                        txtEncoder.Text = GetFieldValue(Qry)
                    End If
                End If
            Case "COLLECTION RECEIPT"
                If cboField.Text = "C.R NUMBER" Then
                    Qry = "SELECT PAYMENTS.CRNO FROM PAYMENTS WHERE PAYMENTS.CRNO = '" & txtOldValue.Text & "' ORDER BY PAYMENTS.PAYDATE DESC LIMIT 1"
                    If Not CtrlNumExist(Qry) Then
                        MsgBox "C.R No. " & txtOldValue.Text & " not found!", vbExclamation + vbOKOnly, "Collection Receipt"
                        txtOldValue.Text = ""
                        txtOldValue.SetFocus
                    Else
                        Qry = "SELECT concat(EMPLOYEES.LASTNAME,', ',EMPLOYEES.FIRSTNAME,' ',IF(ISNULL(LEFT(EMPLOYEES.MIDDLENAME,1)),'',CONCAT(LEFT(EMPLOYEES.MIDDLENAME,1), '.'))) AS EMPNAME " & _
                              "FROM PAYMENTS INNER JOIN EMPLOYEES ON PAYMENTS.UPDATEDBY = EMPLOYEES.EMPID WHERE PAYMENTS.CRNO='" & txtOldValue.Text & "';"
                        txtEncoder.Text = GetFieldValue(Qry)
                    End If
                End If
        End Select
    End If
End Sub

Private Sub txtPrepBy_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtAprvby.SetFocus
End Sub

Private Sub txtReason_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtRemarks.SetFocus
End Sub

Private Sub txtrefno_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        If cboField.Text = "PAYMENT TYPE" Then
            txtReason.SetFocus
        Else
            txtOldValue.SetFocus
        End If
    End If
End Sub

Private Sub txtrefno_LostFocus()
    Dim rsJO As New ADODB.Recordset
    Dim Qry As String
    
    If txtRefNo.Text <> "" Then
        If cboItem.Text = "INVOICE" Then
            'Qry = "SELECT SPINVOICE.SPINVNO From SPINVOICE WHERE (((SPINVOICE.SPINVNO)='" & txtRefNo.Text & "'));"
            'mySQL
            Qry = "SELECT SPINVOICE.SPINVNO From SPINVOICE WHERE SPINVOICE.SPINVNO='" & txtRefNo.Text & "';"
            If Not CtrlNumExist(Qry) Then
                MsgBox "Invoice No. " & txtRefNo.Text & " not found!", vbExclamation + vbOKOnly, "Invoice No."
                txtRefNo.Text = ""
                txtRefNo.SetFocus
            Else
                Qry = "SELECT SPINVOICE.INVDATE From SPINVOICE WHERE (((SPINVOICE.SPINVNO)='" & txtRefNo.Text & "'));"
                txtOldValue.Text = GetFieldValue(Qry)
                'Qry = "SELECT [LASTNAME] & ', ' & [FIRSTNAME] & IIf(IsNull([MIDDLENAME]),'',Left$([MIDDLENAME],1) & '.') AS EMPNAME " & _
                      "FROM SPINVOICE INNER JOIN EMPLOYEES ON SPINVOICE.EMPID = EMPLOYEES.EMPID WHERE (((SPINVOICE.SPINVNO)='" & txtRefNo.Text & "'));"
                'mySQL
                Qry = "SELECT concat(EMPLOYEES.LASTNAME,', ',EMPLOYEES.FIRSTNAME,' ',IF(ISNULL(LEFT(EMPLOYEES.MIDDLENAME,1)),'',CONCAT(LEFT(EMPLOYEES.MIDDLENAME,1), '.'))) AS EMPNAME " & _
                      "FROM SPINVOICE INNER JOIN EMPLOYEES ON SPINVOICE.EMPID = EMPLOYEES.EMPID WHERE SPINVOICE.SPINVNO='" & txtRefNo.Text & "';"
                txtEncoder.Text = GetFieldValue(Qry)
            End If
        ElseIf cboItem.Text = "JOB ORDER" And cboField.Text = "PAYMENT TYPE" Then
            'Qry = "SELECT SRVCJOBORDER.* From SRVCJOBORDER WHERE (((SRVCJOBORDER.JOBORDERNO)='" & txtRefNo.Text & "'));"
            'mySQL
            Qry = "SELECT SRVCJOBORDER.* From SRVCJOBORDER WHERE SRVCJOBORDER.JOBORDERNO='" & txtRefNo.Text & "';"
            With rsJO
                .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
                If .RecordCount > 0 Then
                    txtCustID.Text = !CUSTOMERID
                    txtOldValue.Text = !PayType
                    If txtOldValue.Text = "CASH" Then
                        txtNewValue.Text = "ON ACCOUNT"
                    ElseIf txtOldValue.Text = "ON ACCOUNT" Then
                        txtNewValue.Text = "CASH"
                    End If
                    lblDesc.Caption = "Customer: " & !CUSTOMER
                    txtEncoder.Text = !PreparedBy
                    txtJobDate.Text = !DatePromised
                    txtJobAmt.Text = !Balance
                    txtJobID.Text = !JOBORDERID
                    If InStr(IIf(IsNull(!Remarks), "", !Remarks), "YAMAHA") > 0 Then
                        Qry = "Select CUSTOMERNAME FROM CUSTOMER WHERE (CUSTOMERNAME LIKE '" & "%YAMAHA%" & "' OR CUSTOMERNAME LIKE '" & "%MOTORHUB%" & "') "
                        With rsAnyUse
                            If .State = adStateOpen Then .Close
                            .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
                            If .RecordCount > 0 Then
                                .MoveFirst
                                cboOnCust.Clear
                                Do While Not .EOF
                                    cboOnCust.AddItem !CustomerName
                                    .MoveNext
                                Loop
                            End If
                            .Close
                        End With
                    ElseIf InStr(IIf(IsNull(!Remarks), "", !Remarks), "HONDA") > 0 Then
                        Qry = "Select CUSTOMERNAME FROM CUSTOMER WHERE (CUSTOMERNAME LIKE '" & "%BINOS%" & "' OR CUSTOMERNAME LIKE '" & "%HONDA%" & "') "
                        With rsAnyUse
                            If .State = adStateOpen Then .Close
                            .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
                            If .RecordCount > 0 Then
                                .MoveFirst
                                cboOnCust.Clear
                                Do While Not .EOF
                                    cboOnCust.AddItem !CustomerName
                                    .MoveNext
                                Loop
                            End If
                            .Close
                        End With
                    ElseIf InStr(IIf(IsNull(!Remarks), "", !Remarks), "KAWASAKI") > 0 Then
                        Qry = "Select CUSTOMERNAME FROM CUSTOMER WHERE (CUSTOMERNAME LIKE '" & "%KAWASAKI%" & "' OR CUSTOMERNAME LIKE '" & "%KAWAZAKI%" & "') "
                        With rsAnyUse
                            If .State = adStateOpen Then .Close
                            .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
                            If .RecordCount > 0 Then
                                .MoveFirst
                                cboOnCust.Clear
                                Do While Not .EOF
                                    cboOnCust.AddItem !CustomerName
                                    .MoveNext
                                Loop
                            End If
                            .Close
                        End With
                    ElseIf InStr(IIf(IsNull(!Remarks), "", !Remarks), "SUZUKI") > 0 Then
                        Qry = "Select CUSTOMERNAME FROM CUSTOMER WHERE (CUSTOMERNAME LIKE '" & "%SUZUKI%" & "' OR CUSTOMERNAME LIKE '" & "%SUSUKI%" & "') "
                        With rsAnyUse
                            If .State = adStateOpen Then .Close
                            .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
                            If .RecordCount > 0 Then
                                .MoveFirst
                                cboOnCust.Clear
                                Do While Not .EOF
                                    cboOnCust.AddItem !CustomerName
                                    .MoveNext
                                Loop
                            End If
                            .Close
                        End With
                    ElseIf Val(txtCustID) > 0 Then
                        Qry = "SELECT CUSTOMER FROM srvcustomer WHERE CUSTOMERID = " & Val(txtCustID) & " "
                        With rsAnyUse
                            If .State = adStateOpen Then .Close
                            .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
                            If .RecordCount > 0 Then
                                .MoveFirst
                                cboOnCust.Clear
                                Do While Not .EOF
                                    cboOnCust.AddItem !CUSTOMER
                                    .MoveNext
                                Loop
                            End If
                            .Close
                        End With
                    Else
                        Qry = "Select CUSTOMERNAME FROM CUSTOMER"
                        With rsAnyUse
                            If .State = adStateOpen Then .Close
                            .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
                            If .RecordCount > 0 Then
                                .MoveFirst
                                cboOnCust.Clear
                                Do While Not .EOF
                                    cboOnCust.AddItem !CustomerName
                                    .MoveNext
                                Loop
                            End If
                            .Close
                        End With
                    End If
                Else
                    txtCustID.Text = ""
                    txtOldValue.Text = ""
                    txtNewValue.Text = ""
                    lblDesc.Caption = ""
                    txtEncoder.Text = ""
                    txtJobDate.Text = ""
                    txtJobAmt.Text = ""
                    txtJobID.Text = ""
                End If
                .Close
            End With
            If Val(txtCustID.Text) > 0 Then 'Display Charge to : Customer Name @ label2
                'Qry = "SELECT CUSTOMER.CUSTOMERNAME From CUSTOMER WHERE (((CUSTOMER.CUSTOMERID)=" & txtCustID.Text & "));"
                'mySQL
                Qry = "SELECT CUSTOMER.CUSTOMERNAME From CUSTOMER WHERE CUSTOMER.CUSTOMERID=" & txtCustID.Text & ";"
                Label2.Caption = "Charge To: " & GetFieldValue(Qry)
            Else
                Label2.Caption = ""
            End If
            If txtJobID.Text = "" Then MsgBox "Job Order No. " & txtRefNo.Text & " does not exist!", vbExclamation + vbOKOnly, "Adjustment"
        End If
    End If
End Sub

Private Sub txtRemarks_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtPrepby.SetFocus
End Sub
