VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form AcctCustomer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "On-Account Customer"
   ClientHeight    =   6780
   ClientLeft      =   45
   ClientTop       =   150
   ClientWidth     =   13470
   Icon            =   "AcctCustomer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6780
   ScaleWidth      =   13470
   StartUpPosition =   2  'CenterScreen
   Begin lvButton.lvButtons_H cmdSrvcMember 
      Height          =   375
      Left            =   4920
      TabIndex        =   53
      Top             =   6120
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "Show MC"
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
      Image           =   "AcctCustomer.frx":0CCA
      cBack           =   -2147483633
   End
   Begin VB.CheckBox chkCC 
      Caption         =   "CREDIT CARD"
      Height          =   255
      Left            =   360
      TabIndex        =   90
      Top             =   6120
      Width           =   1455
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   1920
      Top             =   6600
   End
   Begin VB.TextBox txtEditCustID 
      Height          =   375
      Left            =   240
      TabIndex        =   65
      Top             =   6600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CheckBox ChkApply 
      Caption         =   "APPLY IN SERVICE"
      Height          =   255
      Left            =   1920
      TabIndex        =   54
      Top             =   6120
      Width           =   1815
   End
   Begin VB.Timer tmrDropRev 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3960
      Top             =   6600
   End
   Begin VB.Timer tmrFloatRev 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3480
      Top             =   6600
   End
   Begin VB.Timer tmrDrop 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2880
      Top             =   6600
   End
   Begin VB.Timer tmrFloat 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2400
      Top             =   6600
   End
   Begin VB.Frame Frame1 
      Height          =   5175
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   6045
      Begin VB.TextBox txtAgent 
         Alignment       =   2  'Center
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   2160
         MaxLength       =   50
         TabIndex        =   78
         Top             =   4500
         Width           =   3510
      End
      Begin VB.ComboBox txtCity 
         Height          =   315
         Left            =   2160
         TabIndex        =   76
         Top             =   1406
         Width           =   3510
      End
      Begin VB.ComboBox cboAcctCust 
         Height          =   315
         Left            =   2160
         TabIndex        =   4
         Top             =   720
         Visible         =   0   'False
         Width           =   3495
      End
      Begin VB.TextBox txtBranch 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   2160
         MaxLength       =   50
         TabIndex        =   8
         Top             =   2445
         Width           =   3510
      End
      Begin VB.TextBox txtRegion 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   2160
         MaxLength       =   50
         TabIndex        =   7
         Top             =   2115
         Width           =   3510
      End
      Begin VB.TextBox txtProvince 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   2160
         MaxLength       =   50
         TabIndex        =   6
         Top             =   1770
         Width           =   3510
      End
      Begin VB.TextBox txtContactNo 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   2160
         MaxLength       =   50
         TabIndex        =   9
         Top             =   2790
         Width           =   3495
      End
      Begin VB.ComboBox cboCustType 
         Height          =   315
         Left            =   2160
         TabIndex        =   2
         Top             =   270
         Width           =   3495
      End
      Begin VB.TextBox txtCLimit 
         Alignment       =   2  'Center
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   2160
         MaxLength       =   25
         TabIndex        =   12
         Top             =   3825
         Width           =   3510
      End
      Begin VB.TextBox txtCustTerm 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   2160
         MaxLength       =   25
         TabIndex        =   13
         Top             =   4170
         Width           =   3510
      End
      Begin VB.TextBox txtCustID 
         Height          =   285
         Left            =   120
         TabIndex        =   23
         Top             =   1920
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtAddress 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   2160
         MaxLength       =   50
         TabIndex        =   5
         Top             =   1063
         Width           =   3510
      End
      Begin VB.TextBox txtCustomer 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   2160
         MaxLength       =   50
         TabIndex        =   3
         Top             =   720
         Width           =   3510
      End
      Begin VB.TextBox txtTIN 
         Alignment       =   2  'Center
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   2160
         MaxLength       =   25
         TabIndex        =   11
         Top             =   3480
         Width           =   3510
      End
      Begin VB.TextBox txtConPers 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   2160
         MaxLength       =   25
         TabIndex        =   10
         Top             =   3135
         Width           =   3510
      End
      Begin VB.Label lblI_fld 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "AGENT"
         ForeColor       =   &H80000007&
         Height          =   195
         Index           =   2
         Left            =   1530
         TabIndex        =   77
         Top             =   4560
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "CITY/TOWN"
         Height          =   255
         Index           =   4
         Left            =   60
         TabIndex        =   75
         Top             =   1440
         Width           =   2025
      End
      Begin VB.Label lblCustName 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CUSTOMER NAME"
         ForeColor       =   &H80000007&
         Height          =   195
         Left            =   660
         TabIndex        =   74
         Top             =   780
         Width           =   1425
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "BRANCHES"
         Height          =   255
         Index           =   3
         Left            =   420
         TabIndex        =   73
         Top             =   2475
         Width           =   1665
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "REGION"
         Height          =   255
         Index           =   2
         Left            =   780
         TabIndex        =   72
         Top             =   2160
         Width           =   1305
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "PROVINCE"
         Height          =   255
         Index           =   1
         Left            =   660
         TabIndex        =   71
         Top             =   1800
         Width           =   1425
      End
      Begin VB.Label lblI_fld 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CUSTOMER TYPE:"
         ForeColor       =   &H80000007&
         Height          =   195
         Index           =   1
         Left            =   540
         TabIndex        =   30
         Top             =   330
         Width           =   1545
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "ADDRESS"
         Height          =   255
         Index           =   0
         Left            =   60
         TabIndex        =   24
         Top             =   1093
         Width           =   2025
      End
      Begin VB.Label lblI_fld 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TERM"
         ForeColor       =   &H80000007&
         Height          =   195
         Index           =   0
         Left            =   1620
         TabIndex        =   22
         Top             =   4230
         Width           =   465
      End
      Begin VB.Label lblI_fld 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CREDIT LIMIT"
         ForeColor       =   &H80000007&
         Height          =   195
         Index           =   8
         Left            =   540
         TabIndex        =   17
         Top             =   3885
         Width           =   1545
      End
      Begin VB.Label lblContactPerson 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MANAGER"
         ForeColor       =   &H80000007&
         Height          =   195
         Left            =   60
         TabIndex        =   16
         Top             =   3195
         Width           =   2025
      End
      Begin VB.Label lblI_fld 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CONTACT NOS."
         ForeColor       =   &H80000007&
         Height          =   195
         Index           =   6
         Left            =   60
         TabIndex        =   15
         Top             =   2850
         Width           =   2025
      End
      Begin VB.Label lblI_fld 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "T.I.N."
         ForeColor       =   &H80000007&
         Height          =   195
         Index           =   7
         Left            =   1680
         TabIndex        =   14
         Top             =   3540
         Width           =   405
      End
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   6120
      Top             =   2400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AcctCustomer.frx":96D1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AcctCustomer.frx":97BF6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AcctCustomer.frx":99D30
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AcctCustomer.frx":9AC0A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AcctCustomer.frx":9BAE4
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AcctCustomer.frx":9C9BE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   1440
      _ExtentX        =   2540
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
      Image           =   "AcctCustomer.frx":9FA40
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   1
      Left            =   1440
      TabIndex        =   25
      Top             =   0
      Width           =   1440
      _ExtentX        =   2540
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
      Image           =   "AcctCustomer.frx":9FFDA
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   3
      Left            =   4320
      TabIndex        =   26
      Top             =   0
      Width           =   1440
      _ExtentX        =   2540
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
      Image           =   "AcctCustomer.frx":A0574
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   4
      Left            =   5760
      TabIndex        =   27
      Top             =   0
      Visible         =   0   'False
      Width           =   1440
      _ExtentX        =   2540
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
      Image           =   "AcctCustomer.frx":A0B0E
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   5
      Left            =   7200
      TabIndex        =   28
      Top             =   0
      Visible         =   0   'False
      Width           =   1440
      _ExtentX        =   2540
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
      Image           =   "AcctCustomer.frx":A10A8
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   2
      Left            =   2880
      TabIndex        =   29
      Top             =   0
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   661
      Caption         =   "&Find"
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
      Image           =   "AcctCustomer.frx":A1642
      cBack           =   -2147483633
   End
   Begin MSFlexGridLib.MSFlexGrid flxAllMC 
      Height          =   1215
      Left            =   8280
      TabIndex        =   68
      Top             =   5640
      Visible         =   0   'False
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   2143
      _Version        =   393216
      Cols            =   14
      FixedCols       =   0
   End
   Begin lvButton.lvButtons_H cmdCC 
      Height          =   375
      Left            =   3840
      TabIndex        =   91
      Top             =   6120
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Caption         =   "Show CC"
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
      Image           =   "AcctCustomer.frx":A1BDC
      cBack           =   -2147483633
   End
   Begin TabDlg.SSTab ssCreditTab 
      Height          =   6015
      Left            =   6480
      TabIndex        =   79
      Top             =   480
      Visible         =   0   'False
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   10610
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Credit Card Information"
      TabPicture(0)   =   "AcctCustomer.frx":A1F2C
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "flxCreditInfo"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Credit Card Entry Form"
      TabPicture(1)   =   "AcctCustomer.frx":A1F48
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "btnCC(2)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "btnCC(1)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "btnCC(0)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Frame6"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).ControlCount=   4
      Begin MSFlexGridLib.MSFlexGrid flxCreditInfo 
         Height          =   5535
         Left            =   -74940
         TabIndex        =   92
         Top             =   360
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   9763
         _Version        =   393216
         Cols            =   5
         BackColorBkg    =   14737632
         FormatString    =   "|<BANK NAME |<HOLDER NAME |<ACCT NO."
      End
      Begin VB.Frame Frame6 
         Height          =   1575
         Left            =   120
         TabIndex        =   80
         Tag             =   "XYF"
         Top             =   840
         Width           =   6615
         Begin VB.TextBox txtCardHolder 
            Height          =   285
            Left            =   1920
            MaxLength       =   50
            TabIndex        =   83
            Tag             =   "XYF"
            Top             =   360
            Width           =   4455
         End
         Begin VB.TextBox txtCardNo 
            Height          =   285
            Left            =   1920
            TabIndex        =   82
            Tag             =   "XYF"
            Top             =   720
            Width           =   4455
         End
         Begin VB.TextBox txtCreditBank 
            Height          =   285
            Left            =   1920
            TabIndex        =   81
            Tag             =   "XYF"
            Top             =   1080
            Width           =   4455
         End
         Begin VB.Label Label12 
            Caption         =   "CARD HOLDER NAME:"
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
            TabIndex        =   86
            Tag             =   "XYF"
            Top             =   360
            Width           =   1695
         End
         Begin VB.Label Label13 
            Caption         =   "CARD NO.:"
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
            TabIndex        =   85
            Tag             =   "XYF"
            Top             =   720
            Width           =   1575
         End
         Begin VB.Label Label14 
            Caption         =   "NAME OF  BANK:"
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
            TabIndex        =   84
            Tag             =   "XYF"
            Top             =   1080
            Width           =   1575
         End
      End
      Begin lvButton.lvButtons_H btnCC 
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   87
         Top             =   360
         Width           =   1440
         _ExtentX        =   2540
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
         Image           =   "AcctCustomer.frx":A1F64
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H btnCC 
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   88
         Top             =   360
         Visible         =   0   'False
         Width           =   1440
         _ExtentX        =   2540
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
         Image           =   "AcctCustomer.frx":A24FE
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H btnCC 
         Height          =   375
         Index           =   2
         Left            =   1560
         TabIndex        =   89
         Top             =   360
         Visible         =   0   'False
         Width           =   1440
         _ExtentX        =   2540
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
         Image           =   "AcctCustomer.frx":A2A98
         cBack           =   -2147483633
      End
   End
   Begin TabDlg.SSTab tabMC 
      Height          =   6015
      Left            =   6480
      TabIndex        =   31
      Top             =   480
      Visible         =   0   'False
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   10610
      _Version        =   393216
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "Accredited MC"
      TabPicture(0)   =   "AcctCustomer.frx":A3032
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "flxMC"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Data Entry Form"
      TabPicture(1)   =   "AcctCustomer.frx":A304E
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "cmdMC(2)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "cmdMC(1)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "cmdMC(0)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Frame2"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Terms"
      TabPicture(2)   =   "AcctCustomer.frx":A306A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame3"
      Tab(2).Control(1)=   "cmdTerm(0)"
      Tab(2).Control(2)=   "cmdTerm(1)"
      Tab(2).Control(3)=   "cmdTerm(2)"
      Tab(2).ControlCount=   4
      Begin VB.Frame Frame3 
         Enabled         =   0   'False
         Height          =   3210
         Left            =   -74880
         TabIndex        =   56
         Top             =   780
         Width           =   6615
         Begin VB.TextBox txtGives 
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
            TabIndex        =   61
            Tag             =   "XYF"
            Top             =   615
            Width           =   1335
         End
         Begin VB.TextBox txtTerm 
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
            TabIndex        =   60
            Tag             =   "XYF"
            Top             =   975
            Width           =   1335
         End
         Begin VB.ComboBox cboScheme 
            Height          =   315
            Left            =   2040
            Style           =   2  'Dropdown List
            TabIndex        =   59
            Top             =   240
            Width           =   4215
         End
         Begin VB.Label lblGives 
            Caption         =   "NO. OF PAYMENTS"
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
            Left            =   120
            TabIndex        =   64
            Tag             =   "XYF"
            Top             =   645
            Width           =   1695
         End
         Begin VB.Label Label5 
            Caption         =   "TERMS [DAYS]"
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
            Left            =   120
            TabIndex        =   63
            Tag             =   "XYF"
            Top             =   1005
            Width           =   1695
         End
         Begin VB.Label Label4 
            Caption         =   "PAYMENT SCHEME"
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
            Left            =   120
            TabIndex        =   62
            Tag             =   "XYF"
            Top             =   270
            Width           =   1695
         End
      End
      Begin VB.Frame Frame2 
         Height          =   3315
         Left            =   120
         TabIndex        =   33
         Top             =   780
         Width           =   6615
         Begin VB.ComboBox cboEngine 
            Height          =   315
            Left            =   2400
            TabIndex        =   42
            Top             =   3480
            Visible         =   0   'False
            Width           =   2895
         End
         Begin VB.TextBox txtOAcustID 
            Height          =   285
            Left            =   1440
            TabIndex        =   67
            Top             =   3480
            Visible         =   0   'False
            Width           =   615
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxCust 
            Height          =   1455
            Left            =   1920
            TabIndex        =   66
            Top             =   3840
            Visible         =   0   'False
            Width           =   4455
            _ExtentX        =   7858
            _ExtentY        =   2566
            _Version        =   393216
            FixedRows       =   0
            FixedCols       =   0
            BackColorBkg    =   16777215
            GridColor       =   16777215
            GridColorFixed  =   16777215
            GridColorUnpopulated=   16777215
            ScrollBars      =   2
            SelectionMode   =   1
            Appearance      =   0
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin VB.TextBox txtAcct 
            Height          =   285
            Left            =   480
            TabIndex        =   52
            Top             =   3480
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.TextBox txtModel 
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
            Left            =   1920
            TabIndex        =   40
            Tag             =   "XYF"
            Top             =   1230
            Width           =   4455
         End
         Begin VB.TextBox txtEngine 
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
            Left            =   1920
            TabIndex        =   41
            Tag             =   "XYF"
            Top             =   1560
            Width           =   4455
         End
         Begin VB.TextBox txtChassis 
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
            Left            =   1920
            TabIndex        =   43
            Tag             =   "XYF"
            Top             =   1890
            Width           =   4455
         End
         Begin VB.TextBox txtPlateNo 
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
            Left            =   1920
            TabIndex        =   44
            Tag             =   "XYF"
            Top             =   2220
            Width           =   4455
         End
         Begin VB.TextBox txtBrand 
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
            Left            =   1920
            TabIndex        =   39
            Tag             =   "XYF"
            Top             =   900
            Width           =   4455
         End
         Begin VB.TextBox txtDesig 
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
            Left            =   1920
            TabIndex        =   38
            Tag             =   "XYF"
            Top             =   570
            Width           =   4455
         End
         Begin VB.TextBox txtMCust 
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
            Left            =   1920
            TabIndex        =   37
            Tag             =   "XYF"
            Top             =   240
            Width           =   4455
         End
         Begin VB.Label lblNote 
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
            Height          =   615
            Left            =   1920
            TabIndex        =   70
            Tag             =   "XYF"
            Top             =   2550
            Width           =   4455
         End
         Begin VB.Label lblAvlMC 
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
            Left            =   240
            TabIndex        =   69
            Tag             =   "XYF"
            Top             =   2640
            Width           =   1575
         End
         Begin VB.Label Label2 
            Caption         =   "MODEL"
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
            TabIndex        =   51
            Tag             =   "XYF"
            Top             =   1260
            Width           =   1815
         End
         Begin VB.Label Label2 
            Caption         =   "ENGINE NO."
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
            Left            =   240
            TabIndex        =   50
            Tag             =   "XYF"
            Top             =   1590
            Width           =   1815
         End
         Begin VB.Label Label2 
            Caption         =   "CHASSIS NO."
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
            Index           =   8
            Left            =   240
            TabIndex        =   49
            Tag             =   "XYF"
            Top             =   1920
            Width           =   1815
         End
         Begin VB.Label Label2 
            Caption         =   "PLATE NO."
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
            Left            =   240
            TabIndex        =   48
            Tag             =   "XYF"
            Top             =   2250
            Width           =   1815
         End
         Begin VB.Label Label2 
            Caption         =   "MC BRAND"
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
            Index           =   9
            Left            =   240
            TabIndex        =   47
            Tag             =   "XYF"
            Top             =   930
            Width           =   1815
         End
         Begin VB.Label Label2 
            Caption         =   "CUSTOMER"
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
            TabIndex        =   46
            Tag             =   "XYF"
            Top             =   270
            Width           =   1815
         End
         Begin VB.Label lblDealer 
            Caption         =   "DESIGNATION"
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
            TabIndex        =   45
            Tag             =   "XYF"
            Top             =   600
            Width           =   1815
         End
      End
      Begin lvButton.lvButtons_H cmdMC 
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   34
         Top             =   360
         Width           =   1440
         _ExtentX        =   2540
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
         Image           =   "AcctCustomer.frx":A3086
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdMC 
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   35
         Top             =   360
         Visible         =   0   'False
         Width           =   1440
         _ExtentX        =   2540
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
         Image           =   "AcctCustomer.frx":A3620
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdMC 
         Height          =   375
         Index           =   2
         Left            =   1560
         TabIndex        =   36
         Top             =   360
         Visible         =   0   'False
         Width           =   1440
         _ExtentX        =   2540
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
         Image           =   "AcctCustomer.frx":A3BBA
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdTerm 
         Height          =   375
         Index           =   0
         Left            =   -74880
         TabIndex        =   55
         Top             =   360
         Width           =   1440
         _ExtentX        =   2540
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
         Image           =   "AcctCustomer.frx":A4154
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdTerm 
         Height          =   375
         Index           =   1
         Left            =   -74880
         TabIndex        =   57
         Top             =   360
         Visible         =   0   'False
         Width           =   1440
         _ExtentX        =   2540
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
         Image           =   "AcctCustomer.frx":A46EE
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdTerm 
         Height          =   375
         Index           =   2
         Left            =   -73440
         TabIndex        =   58
         Top             =   360
         Visible         =   0   'False
         Width           =   1440
         _ExtentX        =   2540
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
         Image           =   "AcctCustomer.frx":A4C88
         cBack           =   -2147483633
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxMC 
         Height          =   5535
         Left            =   -74940
         TabIndex        =   32
         Top             =   360
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   9763
         _Version        =   393216
         SelectionMode   =   1
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.Label lblHeading 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ON-ACCOUNT"
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
      Left            =   3195
      TabIndex        =   20
      Top             =   540
      Width           =   3000
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
      TabIndex        =   18
      Top             =   540
      Width           =   3495
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
      TabIndex        =   19
      Top             =   570
      Width           =   3495
   End
   Begin VB.Label lblHeading 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ON-ACCOUNT"
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
      Left            =   3225
      TabIndex        =   21
      Top             =   570
      Width           =   3000
   End
   Begin VB.Menu mnuMC 
      Caption         =   "Service Customer"
      Visible         =   0   'False
      Begin VB.Menu mnuEditSrvCust 
         Caption         =   "Edit Service Customer"
      End
   End
   Begin VB.Menu mnuCC 
      Caption         =   "Credit Card"
      Visible         =   0   'False
      Begin VB.Menu mnuEditCreditCard 
         Caption         =   "Edit Credit Card"
      End
   End
End
Attribute VB_Name = "AcctCustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsCustomer As New ADODB.Recordset
Dim rsOnAcct As New ADODB.Recordset
Dim AddNew As Boolean, NewMC As Boolean, CustSrch As Boolean
Dim AcctSrch As Boolean, CustBln As Boolean, StatSrch As Boolean
Dim BtnIndx As Integer, ActRow As Integer
Dim rsAcctCust As New ADODB.Recordset
Dim CreditCardID As Long
Dim formAction As String

Sub ClearEntries()
    
    txtCustID.Text = ""
    txtCustomer.Text = ""
    txtAddress.Text = ""
    txtCity.Text = ""
    txtContactNo.Text = ""
    txtConPers.Text = ""
    txtTIN.Text = ""
    txtCLimit.Text = ""
    txtCustTerm.Text = ""
    txtProvince.Text = ""
    txtRegion.Text = ""
    txtBranch.Text = ""
    txtAgent.Text = ""
    
End Sub

Sub ClearEntriesCreditCard()
    txtCardHolder.Text = ""
    txtCardNo.Text = ""
    txtCreditBank.Text = ""
    CreditCardID = 0
End Sub
Sub LockData(bln As Boolean)
    txtCustID.Locked = bln
    txtCustomer.Locked = bln
    txtAddress.Locked = bln
    txtCity.Locked = bln
    txtProvince.Locked = bln
    txtRegion.Locked = bln
    txtBranch.Locked = bln
    txtContactNo.Locked = bln
    txtConPers.Locked = bln
    txtTIN.Locked = bln
    txtCLimit.Locked = bln
    ChkApply.Enabled = Not bln
    cboCustType.Locked = bln
    txtAgent.Locked = bln
    chkCC.Enabled = Not bln
End Sub

Sub SaveCustomer()
    Dim rsTerm As New ADODB.Recordset
    Dim X As Integer
    Dim Qry As String


If Trim(cboCustType.Text) = "3S SPAREPARTS" Then
    If InStr(cboAcctCust.Text, "SPAREPARTS") > 0 Or InStr(cboAcctCust.Text, "PIAGGIO") > 0 Then
        If AddNew Then
            myLastUpdt = ""
            myLastUpdt = Format(Now, "yyyy-mm-dd hh:mm:ss")
            LYRcon.Execute "INSERT INTO customer(CTYPE, CUSTOMERNAME, ADDRESS, CITYTOWN, PROVINCE, REGION, BRANCHES, CONTACTNO, CONTACTPERSON, " & _
                "CRLIMIT, UPDATEDBY, LASTUPDATE, TIN, APPLYSERVICE, AGENT, TERMS, APPLYCREDIT) VALUES('" & cboCustType.Text & "', '" & cboAcctCust.Text & "', " & _
                "'" & txtAddress.Text & "', '" & txtCity.Text & "', '" & txtProvince.Text & "', '" & txtRegion.Text & "', '" & txtBranch.Text & "', '" & txtContactNo.Text & "', '" & txtConPers.Text & "', " & Val(txtCLimit.Text) & ", " & _
                "" & UserID & ", '" & myLastUpdt & "', '" & txtTIN.Text & "', " & ChkApply.Value & ", '" & txtAgent.Text & "', " & Val(txtCustTerm.Text) & "," & Val(chkCC.Value) & ");"
        
                Qry = "SELECT CUSTOMER.* From CUSTOMER WHERE CUSTOMER.LASTUPDATE='" & myLastUpdt & "' AND CUSTOMER.CUSTOMERNAME='" & cboAcctCust.Text & "';"
                rsCustomer.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
                txtCustID.Text = rsCustomer!CUSTOMERID
        Else
            LYRcon.Execute "UPDATE CUSTOMER SET CTYPE = '" & cboCustType.Text & "', CUSTOMERNAME = '" & cboAcctCust.Text & "', " & _
                "ADDRESS = '" & txtAddress.Text & "', CITYTOWN = '" & txtCity.Text & "', PROVINCE = '" & txtProvince.Text & "', REGION = '" & txtRegion.Text & "', BRANCHES = '" & txtBranch.Text & "', CONTACTNO = '" & txtContactNo.Text & "', CONTACTPERSON = '" & txtConPers.Text & "', " & _
                "TIN = '" & txtTIN.Text & "', CRLIMIT = " & Val(txtCLimit.Text) & ", APPLYSERVICE = " & ChkApply.Value & ", " & _
                "UPDATEDBY = " & UserID & ", LASTUPDATE = '" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "', AGENT = '" & txtAgent.Text & "', TERMS = " & Val(txtCustTerm.Text) & " , APPLYCREDIT = " & Val(chkCC.Value) & " Where CUSTOMER.CUSTOMERID = " & txtCustID.Text & ";"
        End If
        If ChkApply.Value = 1 Then
            Qry = "SELECT ONACCTERMS.CUSTOMERID, ONACCTERMS.STATUS From ONACCTERMS " & _
                  "WHERE (((ONACCTERMS.CUSTOMERID) =" & txtCustID.Text & ") AND ((ONACCTERMS.STATUS)='Default'));"
            If Not CtrlNumExist(Qry) Then
                LYRcon.Execute "Insert Into onaccterms(CUSTOMERID,PAYSCHEME,TERMS,NPAYS,SPSORDERID,STATUS)VALUES(" & txtCustID.Text & ",'MONTHLY',30,1,0,'Default');"
            End If
        End If
    Else
        MsgBox "Invalid Customer Name!Please Try Another Customer Name", vbInformation + vbOKOnly, "Validation"
        Exit Sub
    End If
Else
    If AddNew Then
        myLastUpdt = ""
        myLastUpdt = Format(Now, "yyyy-mm-dd hh:mm:ss")
        LYRcon.Execute "INSERT INTO customer(CTYPE, CUSTOMERNAME, ADDRESS, CITYTOWN, PROVINCE, REGION, BRANCHES, CONTACTNO, CONTACTPERSON, " & _
            "CRLIMIT, UPDATEDBY, LASTUPDATE, TIN, APPLYSERVICE, AGENT, TERMS, APPLYCREDIT) VALUES('" & cboCustType.Text & "', '" & cboAcctCust.Text & "', " & _
            "'" & txtAddress.Text & "', '" & txtCity.Text & "', '" & txtProvince.Text & "', '" & txtRegion.Text & "', '" & txtBranch.Text & "', '" & txtContactNo.Text & "', '" & txtConPers.Text & "', " & Val(txtCLimit.Text) & ", " & _
            "" & UserID & ", '" & myLastUpdt & "', '" & txtTIN.Text & "', " & ChkApply.Value & ", '" & txtAgent.Text & "', " & Val(txtCustTerm.Text) & "," & Val(chkCC.Value) & ");"
    
            Qry = "SELECT CUSTOMER.* From CUSTOMER WHERE CUSTOMER.LASTUPDATE='" & myLastUpdt & "' AND CUSTOMER.CUSTOMERNAME='" & cboAcctCust.Text & "';"
            rsCustomer.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
            txtCustID.Text = rsCustomer!CUSTOMERID
    Else
        LYRcon.Execute "UPDATE CUSTOMER SET CTYPE = '" & cboCustType.Text & "', CUSTOMERNAME = '" & cboAcctCust.Text & "', " & _
            "ADDRESS = '" & txtAddress.Text & "', CITYTOWN = '" & txtCity.Text & "', PROVINCE = '" & txtProvince.Text & "', REGION = '" & txtRegion.Text & "', BRANCHES = '" & txtBranch.Text & "', CONTACTNO = '" & txtContactNo.Text & "', CONTACTPERSON = '" & txtConPers.Text & "', " & _
            "TIN = '" & txtTIN.Text & "', CRLIMIT = " & Val(txtCLimit.Text) & ", APPLYSERVICE = " & ChkApply.Value & ", " & _
            "UPDATEDBY = " & UserID & ", LASTUPDATE = '" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "', AGENT = '" & txtAgent.Text & "', TERMS = " & Val(txtCustTerm.Text) & " , APPLYCREDIT = " & Val(chkCC.Value) & " Where CUSTOMER.CUSTOMERID = " & txtCustID.Text & ";"
    End If
    If ChkApply.Value = 1 Then
        Qry = "SELECT ONACCTERMS.CUSTOMERID, ONACCTERMS.STATUS From ONACCTERMS " & _
              "WHERE (((ONACCTERMS.CUSTOMERID) =" & txtCustID.Text & ") AND ((ONACCTERMS.STATUS)='Default'));"
        If Not CtrlNumExist(Qry) Then
            LYRcon.Execute "Insert Into onaccterms(CUSTOMERID,PAYSCHEME,TERMS,NPAYS,SPSORDERID,STATUS)VALUES(" & txtCustID.Text & ",'MONTHLY',30,1,0,'Default');"
        End If
    End If
End If
    If SysBranch = "MOTO-R REPUBLIC CORP." Then
        LYRcon.Execute "UPDATE agingcurrent SET CREDITLIMIT = " & Val(txtCLimit.Text) & ", AGENT = '" & txtAgent.Text & "' " & _
        "Where agingcurrent.CUSTID = " & Val(txtCustID.Text) & " AND agingcurrent.BRANCH = 'MRC1';"
    ElseIf SysBranch = "LYR MOTORPARTS CORP. - TAGUM" Then
        LYRcon.Execute "UPDATE agingcurrent SET CREDITLIMIT = " & Val(txtCRlimit.Text) & ", AGENT = '" & txtAgent.Text & "' " & _
        "Where agingcurrent.CUSTID = " & Val(txtCustID.Text) & " AND agingcurrent.BRANCH = 'MRC2';"
    ElseIf SysBranch = "LYR MOTORPARTS CORP. - VAN SELLING I" Then
        LYRcon.Execute "UPDATE agingcurrent SET CREDITLIMIT = " & Val(txtCRlimit.Text) & ", AGENT = '" & txtAgent.Text & "' " & _
        "Where agingcurrent.CUSTID = " & Val(txtCustID.Text) & " AND agingcurrent.BRANCH = 'VAN1';"
    ElseIf SysBranch = "LYR MOTORPARTS CORP. - VAN SELLING II" Then
        LYRcon.Execute "UPDATE agingcurrent SET CREDITLIMIT = " & Val(txtCRlimit.Text) & ", AGENT = '" & txtAgent.Text & "' " & _
        "Where agingcurrent.CUSTID = " & Val(txtCustID.Text) & " AND agingcurrent.BRANCH = 'VAN2';"
    End If

End Sub

Sub ShowRecord(myCustID As Integer)
    Dim rsCust As New ADODB.Recordset
    Dim Qry As String
    Dim X As Integer, i As Integer
    '***Access
    'Qry = "SELECT CUSTOMER.* From CUSTOMER WHERE (((CUSTOMER.CUSTOMERID)=" & MyCustID & "));"
    'mySQL
    Qry = "SELECT CUSTOMER.* From CUSTOMER WHERE CUSTOMER.CUSTOMERID=" & myCustID & ";"
   
    rsCust.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    With rsCust
        If .RecordCount > 0 Then
            .MoveFirst
            cboCustType.Text = IIf(IsNull(!Ctype), "", !Ctype)
            For i = 0 To cboCustType.ListCount - 1
                If cboCustType.List(i) = !Ctype Then cboCustType.ListIndex = i
            Next i
            txtCustID.Text = !CUSTOMERID
            txtCustomer.Text = !CustomerName
            txtAddress.Text = !Address
            txtCity.Text = IIf(IsNull(!CITYTOWN), "", !CITYTOWN)
            txtProvince.Text = IIf(IsNull(!province), "", !province)
            txtRegion.Text = IIf(IsNull(!REGION), "", !REGION)
            txtBranch.Text = IIf(IsNull(!BRANCHES), "", !BRANCHES)
            txtContactNo.Text = IIf(IsNull(!CONTACTNO), "", !CONTACTNO)
            txtConPers.Text = IIf(IsNull(!ContactPerson), "", !ContactPerson)
            txtTIN.Text = IIf(IsNull(!TIN), "", !TIN)
            txtCLimit.Text = Format$(!CRLimit, "#,###,##0.00")
            txtCustTerm.Text = IIf(IsNull(!Terms), "", !Terms)
            txtAgent.Text = IIf(IsNull(!Agent), "", !Agent)
            ChkApply.Value = IIf(!APPLYSERVICE, 1, 0)
            chkCC.Value = IIf(!APPLYCREDIT, 1, 0)
        End If
        .Close
    End With
    Call ClearMCust
    Call HideSrvcCustomer
    If ChkApply.Value = 1 Then
        cmdSrvcMember.Visible = True
    Else
        cmdSrvcMember.Visible = False
    End If
    
    If chkCC.Value = 1 Then
        cmdCC.Visible = True
    Else
        cmdCC.Visible = False
    End If
End Sub
Function ValidateEntry() As String
    Dim myMsg As String
    
    myMsg = ""
    If AddNew Then
        If cboAcctCust.Text = "" Then myMsg = "Incomplete entry on " & lblCustName.Caption
        If Val(txtCustID.Text) > 0 Then myMsg = myMsg & vbCrLf & "Customer Name already exist"
    Else
        If Val(txtEditCustID.Text) > 0 Then
            If Val(txtCustID.Text) <> Val(txtEditCustID.Text) Then myMsg = myMsg & vbCrLf & "Customer Name already exist"
        End If
    End If
    If cboCustType.ListIndex = -1 Then myMsg = myMsg & vbCrLf & "Invalid Customer Type! Please select from the list."
    If txtAddress.Text = "" Then myMsg = myMsg & vbCrLf & "Incomplete entry on field ADDRESS"
    
    ValidateEntry = myMsg
End Function

Private Sub BtnSaveCancel(bln As Boolean)
    Dim i As Integer
    
    If bln Then
        For i = 0 To 3
            cmdTrans(i).Visible = False
        Next i
        cmdTrans(4).Visible = True
        cmdTrans(5).Visible = True
        cmdTrans(4).Left = 0
        cmdTrans(5).Left = cmdTrans(4).Width
    Else
        For i = 0 To 3
            cmdTrans(i).Visible = True
        Next i
        cmdTrans(4).Visible = False
        cmdTrans(5).Visible = False
    End If
End Sub

Private Sub ClearMCust()
    txtMCust.Text = ""
    txtDesig.Text = ""
    txtModel.Text = ""
    txtEngine.Text = ""
    txtChassis.Text = ""
    txtPlateNo.Text = ""
    txtBrand.Text = ""
    cboEngine.Text = ""
    cboEngine.Clear
    cboEngine.Visible = False
    txtEngine.Visible = True
    txtOAcustID.Text = ""
    lblNote.Caption = ""
    lblAvlMC.Caption = ""
End Sub


Private Sub btnCC_Click(Index As Integer)
Dim rsSaveCC As New ADODB.Recordset
Dim Ans As String
Select Case Index
    Case 0  'add new
        txtCardHolder.Text = ""
        txtCardNo.Text = ""
        txtCreditBank.Text = ""
        
        btnCC(0).Visible = False
        btnCC(1).Visible = True
        btnCC(2).Visible = True
        For i = 0 To 3
            cmdTrans(i).Enabled = False
        Next i
        ssCreditTab.TabEnabled(0) = False
        ssCreditTab.TabEnabled(1) = False
        cmdCC.Enabled = False
        cmdSrvcMember.Enabled = False
        formAction = "addnew"
        txtCardHolder.SetFocus
    Case 1  'save
        If formAction = "addnew" Then
            Ans = MsgBox("Do you want to Save Credit Card Details?", vbQuestion + vbYesNo, "Save")
            If Ans = vbNo Then Exit Sub
            Qry = "call sp_mrcretails.sp_creditcard_create('" & sysBranchCode & "', '" & txtCardHolder.Text & "', '" & txtCardNo & "','" & txtCreditBank.Text & "','" & txtCustID.Text & "'," & Val(UserID) & ")"
        ElseIf formAction = "edit" Then
            Ans = MsgBox("Do you want to Update Credit Card Details?", vbQuestion + vbYesNo, "Save")
            If Ans = vbNo Then Exit Sub
            Qry = "call sp_mrcretails.sp_creditcard_update('" & sysBranchCode & "', '" & txtCardHolder.Text & "', '" & txtCardNo & "','" & txtCreditBank.Text & "'," & Val(flxCreditInfo.TextMatrix(ActRow, 4)) & "," & Val(UserID) & ")"
        End If
        
        If rsSaveCC.State = adStateOpen Then rsSaveCC.Close
        With rsSaveCC
            .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
            
            If .RecordCount <= 0 Then
                .Close
            End If
            
            If .RecordCount > 0 Then
                If !level = "Granted" Then
                    CreditCardID = IIf(IsNull(!CCID), 0, !CCID)
                    MsgBox "'" & !message & "'", vbInformation + vbOKOnly, "'" & !level & "'"
                Else
                    MsgBox "'" & !message & "'", vbInformation + vbOKOnly, "'" & !level & "'"
                End If
            End If
            .Close
        End With
        btnCC(0).Visible = True
        btnCC(1).Visible = False
        btnCC(2).Visible = False
        For i = 0 To 3
            cmdTrans(i).Enabled = True
        Next i
        
        ssCreditTab.TabEnabled(0) = True
        ssCreditTab.TabEnabled(1) = True
        cmdCC.Enabled = True
        cmdSrvcMember.Enabled = True
        ssCreditTab.Tab = 0
        Call ClearEntriesCreditCard
        Call DisplayCreditCardAcct
    Case 2  'cancel
        Ans = MsgBox("Do you want to Cancel " & IIf(NewCC, "Entry", "Editing") & "?", vbQuestion + vbYesNo, "Cancel")
        If Ans = vbNo Then Exit Sub
        
        txtCardHolder.Text = ""
        txtCardNo.Text = ""
        txtCreditBank.Text = ""
        
        ssCreditTab.Tab = 0
        btnCC(0).Visible = True
        btnCC(1).Visible = False
        btnCC(2).Visible = False
        
        For i = 0 To 3
            cmdTrans(i).Enabled = True
        Next i
        
        ssCreditTab.TabEnabled(0) = True
        ssCreditTab.TabEnabled(1) = True
        cmdCC.Enabled = True
        cmdSrvcMember.Enabled = True
End Select
End Sub

Private Sub cboAcctCust_Change()
    Dim i As Long, j As Long
    Dim strPartial As String, strTotal As String
    Dim X As Integer
    
    If CustSrch Then
        CustSrch = False
        Exit Sub
    End If
    With cboAcctCust
        strPartial = .Text
        i = SendMessage(.hwnd, CB_FINDSTRING, -1, ByVal strPartial) 'If match found, append unmatched characters
        If i <> -1 Then
            strTotal = .List(i)
            rsAcctCust.MoveFirst
            rsAcctCust.Move i
            If AddNew Then
                txtCustID.Text = rsAcctCust!CUSTOMERID
                txtAddress.Text = rsAcctCust!Address
                txtContactNo.Text = IIf(IsNull(rsAcctCust!CONTACTNO), "", rsAcctCust!CONTACTNO)
                txtConPers.Text = IIf(IsNull(rsAcctCust!ContactPerson), "", rsAcctCust!ContactPerson)
                txtTIN.Text = IIf(IsNull(rsAcctCust!TIN), "", rsAcctCust!TIN)
                txtCLimit.Text = IIf(IsNull(rsAcctCust!CRLimit), "", rsAcctCust!CRLimit)
                txtCustTerm.Text = IIf(IsNull(rsAcctCust!Terms), "", rsAcctCust!Terms)
            Else
                txtEditCustID.Text = rsAcctCust!CUSTOMERID
            End If
            j = Len(strTotal) - Len(strPartial)
            If j <> 0 Then
                CustSrch = True
                .SelText = Right$(strTotal, j)
                .SelStart = Len(strPartial)
                .SelLength = j
            End If
        Else
            If AddNew Then
                txtCustID.Text = ""
                txtAddress.Text = ""
                txtContactNo.Text = ""
                txtConPers.Text = ""
                txtTIN.Text = ""
                txtCLimit.Text = ""
                txtCustTerm.Text = ""
            Else
                txtEditCustID.Text = ""
            End If
        End If
    End With

End Sub

Private Sub cboAcctCust_Click()
    Call cboAcctCust_Change
End Sub

Private Sub cboAcctCust_GotFocus()
   Dim Tmp
    
   Const CB_SHOWDROPDOWN = &H14F
   Tmp = SendMessage(cboAcctCust.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub

Private Sub cboAcctCust_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then CustSrch = True
End Sub

Private Sub cboAcctCust_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        txtAddress.SetFocus
    End If
End Sub


Private Sub cboCustType_Change()
    Dim i As Long, j As Long
    Dim strPartial As String, strTotal As String
    Dim X As Integer
    
    If CustSrch Then
        CustSrch = False
        Exit Sub
    End If
    With cboCustType
        strPartial = .Text
        i = SendMessage(.hwnd, CB_FINDSTRING, -1, ByVal strPartial) 'If match found, append unmatched characters
        If i <> -1 Then
            strTotal = .List(i)
            j = Len(strTotal) - Len(strPartial)
            If InStr(.Text, "LGC") > 0 Or InStr(.Text, "CORPORATE") > 0 Then
                lblCustName.Caption = "COMPANY NAME"
                lblContactPerson.Caption = "MANAGER"
            Else
                lblCustName.Caption = "CUSTOMER NAME"
                lblContactPerson.Caption = "EMPLOYER"
            End If
            If j <> 0 Then
                CustSrch = True
                .SelText = Right$(strTotal, j)
                .SelStart = Len(strPartial)
                .SelLength = j
            End If
        End If
    End With

End Sub

Private Sub cboCustType_Click()
    Call cboCustType_Change
End Sub

Private Sub cboCustType_GotFocus()
   Dim Tmp
    
   Const CB_SHOWDROPDOWN = &H14F
   Tmp = SendMessage(cboCustType.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub

Private Sub cboCustType_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then cboAcctCust.SetFocus
End Sub

Private Sub cboEngine_Change()
    Dim i As Integer, xRow As Integer
    
    With flxAllMC
        xRow = 0
        For i = 1 To .Rows - 1
            If cboEngine.Text = .TextMatrix(i, 4) Then
                xRow = i
                Exit For
            End If
        Next i
        If xRow = 0 Then
            txtAcct.Text = ""
            txtBrand.Text = ""
            txtOAcustID.Text = ""
            txtModel.Text = ""
            txtEngine.Text = cboEngine.Text
            txtChassis.Text = ""
            txtPlateNo.Text = ""
            Exit Sub
        End If
        txtAcct.Text = .TextMatrix(xRow, 0)
        txtBrand.Text = .TextMatrix(xRow, 2) 'Brand
        txtOAcustID.Text = .TextMatrix(xRow, 8)    'CustomerID
        txtModel.Text = .TextMatrix(xRow, 3)  'Model
        txtEngine.Text = .TextMatrix(xRow, 4)  'Engine
        txtChassis.Text = .TextMatrix(xRow, 5)  'chassis
        txtPlateNo.Text = .TextMatrix(xRow, 6)  'Plate
        If .TextMatrix(xRow, 7) = "True" Then
            'Engine is accredited to other customer
            lblNote.Caption = "*Note: The Motorcyle is already accredited as on account to " & GetCustomerName(Val(.TextMatrix(xRow, 8)))
        Else
            lblNote.Caption = ""
        End If
    End With
End Sub

Function GetCustomerName(MyID As Integer) As String
    Dim rsCName As New ADODB.Recordset
    Dim Qry As String
    
    '***Access
    'Qry = "SELECT CUSTOMER.CUSTOMERNAME, CUSTOMER.CUSTOMERID From CUSTOMER " & _
          "WHERE (((CUSTOMER.CUSTOMERID)=" & MyId & "));"
    'mySQL
    Qry = "SELECT CUSTOMER.CUSTOMERNAME, CUSTOMER.CUSTOMERID From CUSTOMER " & _
          "WHERE CUSTOMER.CUSTOMERID=" & MyID & ";"
 
    With rsCName
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            GetCustomerName = !CustomerName
        Else
            GetCustomerName = ""
        End If
        .Close
    End With
    
End Function

Private Sub cboEngine_Click()
    Call cboEngine_Change
End Sub

Private Sub cmdCC_Click()
    Call DisplayCreditCardAcct
    ssCreditTab.Visible = True
    tabMC.Visible = False
    Me.Width = 13560
    ssCreditTab.Tab = 0
    Me.Left = (MotoRep.ScaleWidth - Me.Width) \ 2
'MsgBox "Not Yet Available", vbInformation + vbOKOnly, "Validation"
End Sub

Public Sub DisplayCreditCardAcct()
    Dim rsAcctCredit As New ADODB.Recordset
    Dim rsTerm As New ADODB.Recordset
    Dim Qry As String, i As Integer
    '***Access
    'Qry = "SELECT SRVCUSTOMER.CUSTOMER, SRVCUSTOMER.DESIGNATION, SRVCUSTOMER.BRAND, SRVCUSTOMER.MODEL, SRVCUSTOMER.ENGINE, SRVCUSTOMER.CHASSIS, SRVCUSTOMER.PLATENO, SRVCUSTOMER.ACCOUNTNO " & _
          "From SRVCUSTOMER Where (((SRVCUSTOMER.CUSTOMERID) = " & Val(txtCustID.Text) & ")) ORDER BY SRVCUSTOMER.CUSTOMER;"
    'mySQL
    Qry = "call sp_mrcretails.sp_creditcard_Summary('" & sysBranchCode & "'," & Val(txtCustID.Text) & ")"
    
    rsAcctCredit.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    If rsAcctCredit.RecordCount = 0 Then
        flxCreditInfo.Clear
        flxCreditInfo.Rows = 2
        flxCreditInfo.FormatString = "   |<CCBANKNAME |<CCNAME  |^CCNO |^CCID "
    Else
        With rsAcctCredit
            If rsAcctCredit.RecordCount > 0 Then
                .MoveFirst
                flxCreditInfo.Rows = .RecordCount + 1
                
                For i = 1 To flxCreditInfo.Rows - 1
                    flxCreditInfo.TextMatrix(i, 1) = !CCBANKNAME
                    flxCreditInfo.TextMatrix(i, 2) = !CCNAME
                    flxCreditInfo.TextMatrix(i, 3) = !CCNO
                    flxCreditInfo.TextMatrix(i, 4) = !CCID
                    .MoveNext
                Next i
            End If
            .Close
        End With
    End If
    flxCreditInfo.ColWidth(0) = 300
    For i = 1 To 4
        flxCreditInfo.ColWidth(i) = 2000
    Next i
    flxCreditInfo.ColWidth(4) = 0
    
End Sub
Private Sub cmdMC_Click(Index As Integer)
    Dim i As Integer, Ans As Integer
    Dim Qry As String, MyAcc As String
    
    Select Case Index
        Case 0  'Addnew
            '***Access
            'Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE (((USERACCESS.EMPID)=" & UserID & ") AND ((USERACCESS.MENUINDX)=10));"
            Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE USERACCESS.EMPID=" & UserID & " AND USERACCESS.MENUINDX=10;"
            MyAcc = GetFieldValue(Qry)
            If MyAcc <> "FULL ACCESS" Then
                MsgBox "You don't have full access to this transaction." & vbCrLf & "Please contact your System Administrator", vbExclamation + vbOKOnly, "Access Denied"
                Exit Sub
            End If
            Frame2.Enabled = True
            Call ClearMCust
            '***Access
            'Qry = "SELECT SRVCUSTOMER.ACCOUNTNO, SRVCUSTOMER.CUSTOMER, SRVCUSTOMER.ENGINE " & _
                  "From SRVCUSTOMER Where (((SRVCUSTOMER.ADDTLMC) = False)) ORDER BY SRVCUSTOMER.CUSTOMER;"
            Qry = "SELECT SRVCUSTOMER.ACCOUNTNO, SRVCUSTOMER.CUSTOMER, SRVCUSTOMER.ENGINE " & _
                  "From SRVCUSTOMER Where SRVCUSTOMER.ADDTLMC= 0 ORDER BY SRVCUSTOMER.CUSTOMER;"
            If rsOnAcct.State = adStateOpen Then rsOnAcct.Close
            rsOnAcct.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
            Set flxCust.DataSource = rsOnAcct
            With flxCust
                .Width = txtMCust.Width
                .ColWidth(0) = 0
                .ColWidth(1) = txtMCust.Width
                .ColWidth(2) = 0
            End With
            cmdMC(0).Visible = False
            cmdMC(1).Visible = True
            cmdMC(2).Visible = True
            For i = 0 To 3
                cmdTrans(i).Enabled = False
            Next i
            tabMC.TabEnabled(0) = False
            tabMC.TabEnabled(2) = False
            cmdSrvcMember.Enabled = False
            NewMC = True
            If cboCustType.Text = "INDIVIDUAL" Then 'OptType(1).Value = True Then    'Individual Customer type
                If flxMC.TextMatrix(1, 1) = "" Then
                    txtMCust.Text = txtCustomer.Text
                    flxCust.Visible = False
                    If Val(txtAcct.Text) > 0 Then
                        'SQLAcctCustomer04
                        '***Access
                        'Qry = "SELECT SRVCUSTOMER.CUSTOMER, SRVCUSTOMER.ONACCOUNT From SRVCUSTOMER " & _
                              "WHERE (((SRVCUSTOMER.CUSTOMER)='" & txtMCust.Text & "') AND ((SRVCUSTOMER.ONACCOUNT)=False));"
                        Qry = "SELECT SRVCUSTOMER.CUSTOMER, SRVCUSTOMER.ONACCOUNT From SRVCUSTOMER " & _
                              "WHERE SRVCUSTOMER.CUSTOMER='" & txtMCust.Text & "' AND SRVCUSTOMER.ONACCOUNT=0;"
                        If CtrlNumExist(Qry) Then Call txtMCust_LostFocus
                    End If
                End If
            End If
            txtMCust.SetFocus
            
        Case 1  'Save
            
            vMsg = ValidateSrvCust
            If vMsg <> "" Then
                MsgBox vMsg, vbCritical + vbOKOnly, "Validation Error"
                Exit Sub
            End If
            Ans = MsgBox("Do you want to Save Service Customer?", vbQuestion + vbYesNo, "Save")
            If Ans = vbNo Then Exit Sub
            If Val(txtAcct.Text) > 0 Then
                'SQLAcctCustomer05
                '***Access
                'Qry = "SELECT SRVCUSTOMER.ENGINE, SRVCUSTOMER.ACCOUNTNO From SRVCUSTOMER " & _
                      "WHERE (((SRVCUSTOMER.ENGINE)='" & txtEngine.Text & "') AND ((SRVCUSTOMER.ACCOUNTNO)=" & txtAcct.Text & "));"
                Qry = "SELECT SRVCUSTOMER.ENGINE, SRVCUSTOMER.ACCOUNTNO From SRVCUSTOMER " & _
                      "WHERE SRVCUSTOMER.ENGINE='" & txtEngine.Text & "' AND SRVCUSTOMER.ACCOUNTNO=" & txtAcct.Text & ";"
                If CtrlNumExist(Qry) Then
                    '***Access
                    'LYRcon.Execute "UPDATE SRVCUSTOMER SET SRVCUSTOMER.ONACCOUNT = True, SRVCUSTOMER.CUSTOMERID = " & txtCustID.Text & " WHERE (((SRVCUSTOMER.ACCOUNTNO)=" & txtAcct.Text & "));"
                    LYRcon.Execute "UPDATE SRVCUSTOMER SET SRVCUSTOMER.ONACCOUNT = 1, SRVCUSTOMER.CUSTOMERID = " & txtCustID.Text & " WHERE SRVCUSTOMER.ACCOUNTNO=" & txtAcct.Text & ";"
                Else
                Call SaveSrvcCustomer
                End If
            Else
                Call SaveSrvcCustomer
            End If
            MsgBox "Service On-Account Customer Successfully Saved!", vbInformation + vbOKOnly
            Frame2.Enabled = False
            cmdMC(0).Visible = True
            cmdMC(1).Visible = False
            cmdMC(2).Visible = False
            For i = 0 To 3
                cmdTrans(i).Enabled = True
            Next i
            tabMC.TabEnabled(0) = True
            tabMC.TabEnabled(2) = True
            cmdSrvcMember.Enabled = True
            Call DisplaySrvcAcct
        Case 2  'Cancel
            Ans = MsgBox("Do you want to Cancel " & IIf(NewMC, "Entry", "Editing") & "?", vbQuestion + vbYesNo, "Cancel")
            If Ans = vbNo Then Exit Sub
            If NewMC Then
                Call ClearMCust
            Else
                Call ClearMCust
                tabMC.Tab = 0
            End If
            cmdMC(0).Visible = True
            cmdMC(1).Visible = False
            cmdMC(2).Visible = False
            For i = 0 To 3
                cmdTrans(i).Enabled = True
            Next i
            tabMC.TabEnabled(0) = True
            tabMC.TabEnabled(2) = True
            cmdSrvcMember.Enabled = True
    End Select
End Sub
Private Sub SaveSrvcCustomer()
    Dim rsCust As New ADODB.Recordset
    Dim Qry As String, myAddtlMC As Integer
    
    'SQLAcctCustomer06
    '***Access
    'If NewMC Then
    '    rsCust.Open "SrvCustomer", LYRcon, adOpenKeyset, adLockOptimistic
    '    rsCust.AddNew
    'Else
    '    Qry = "SELECT SRVCUSTOMER.* From SRVCUSTOMER WHERE (((SRVCUSTOMER.ACCOUNTNO)=" & txtOAcustID.Text & "));"
    '    rsCust.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    '    rsCust.MoveFirst
    'End If
    'rsCust!CUSTOMER = txtMCust.Text
    'rsCust!Address = txtAddress.Text ' IIf(OptType(0).Value = True, txtCustomer.Text, txtAddress.Text)  'Company
    'rsCust!CONTACTNO = txtContactNo.Text
    'rsCust!DESIGNATION = txtDesig.Text
    'rsCust!MODEL = txtModel.Text
    'rsCust!ENGINE = txtEngine.Text
    'rsCust!CHASSIS = txtChassis.Text
    'rsCust!PLATENO = txtPlateNo.Text
    'rsCust!BRAND = txtBrand.Text
    'rsCust!Ctype = cboCustType.Text ' IIf(OptType(0).Value = True, "COMPANY", "INDIVIDUAL")
    'Access
    'rsCust!ONACCOUNT = True
    
    'rsCust!ONACCOUNT = 1
    'rsCust!CUSTOMERID = txtCustID.Text
    'If Val(txtAcct.Text) > 0 And flxAllMC.Rows > 1 Then
    '    rsCust!ADDTLMC = True
    'End If
    'rsCust.Update
    'rsCust.Close
    
    'mySQL
    
    If NewMC Then
        If Val(txtAcct.Text) > 0 And flxAllMC.Rows > 1 Then
            myAddtlMC = 1
        Else
            myAddtlMC = 0
        End If
        
        LYRcon.Execute "INSERT INTO srvcustomer(CUSTOMER, ADDRESS, CONTACTNO, DESIGNATION, MODEL, " & _
            "ENGINE, CHASSIS, PLATENO, BRAND, CTYPE, ONACCOUNT, CUSTOMERID, ADDTLMC)" & _
            "VALUES('" & txtMCust.Text & "', '" & txtAddress.Text & "', '" & txtContactNo.Text & "', " & _
            "'" & txtDesig.Text & "', '" & txtModel.Text & "', '" & txtEngine.Text & "', '" & txtChassis.Text & "', " & _
            "'" & txtPlateNo.Text & "', '" & txtBrand.Text & "', '" & cboCustType.Text & "', 1, " & _
            "" & txtCustID.Text & ", " & myAddtlMC & ");"
    Else
        LYRcon.Execute "UPDATE srvcustomer SET CUSTOMER = '" & txtMCust.Text & "', ADDRESS = '" & txtAddress.Text & "', " & _
            "CONTACTNO = '" & txtContactNo.Text & "', DESIGNATION = '" & txtDesig.Text & "', MODEL = '" & txtModel.Text & "', " & _
            "ENGINE = '" & txtEngine.Text & "', CHASSIS = '" & txtChassis.Text & "', PLATENO = '" & txtPlateNo.Text & "', " & _
            "BRAND = '" & txtBrand.Text & "', CTYPE = '" & cboCustType.Text & "', ONACCOUNT = 1, CUSTOMERID = " & txtCustID.Text & ", " & _
            "ADDTLMC = " & myAddtlMC & " where SRVCUSTOMER.ACCOUNTNO=" & txtOAcustID.Text & ";"
    End If

End Sub

Function ValidateSrvCust() As String
    Dim eMsg As String, vQry As String
    
    eMsg = ""
    If txtEngine.Text = "" Then
        eMsg = "Incomplete Entry on field Engine"
    Else
        If NewMC Then
            If Val(txtAcct.Text) = 0 Then
                '***Access
                'vQry = "SELECT SRVCUSTOMER.ENGINE From SRVCUSTOMER WHERE (((SRVCUSTOMER.ENGINE)='" & txtEngine.Text & "'));"
                vQry = "SELECT SRVCUSTOMER.ENGINE From SRVCUSTOMER WHERE SRVCUSTOMER.ENGINE='" & txtEngine.Text & "';"
            Else
                'SQLAcctCustomer07
                '***Access
                'vQry = "SELECT SRVCUSTOMER.ENGINE, SRVCUSTOMER.ACCOUNTNO From SRVCUSTOMER " & _
                         "WHERE (((SRVCUSTOMER.ENGINE)='" & txtEngine.Text & "') AND ((SRVCUSTOMER.ACCOUNTNO)<>" & txtAcct.Text & "));"
                vQry = "SELECT SRVCUSTOMER.ENGINE, SRVCUSTOMER.ACCOUNTNO From SRVCUSTOMER " & _
                         "WHERE SRVCUSTOMER.ENGINE='" & txtEngine.Text & "' AND SRVCUSTOMER.ACCOUNTNO<>" & txtAcct.Text & ";"

            End If
            If CtrlNumExist(vQry) Then
                eMsg = "Engine number " & txtEngine.Text & " already exist!" & vbCrLf & GetEngineOwner(txtEngine.Text)
            End If
        End If
    End If
    If txtMCust.Text = "" Then eMsg = eMsg & vbCrLf & "Incomplete entry on field Customer"
    If InStr(cboCustType.Text, "LGC") > 0 Then
        If txtDesig.Text = "" Then eMsg = eMsg & vbCrLf & "Incomplete entry on field Designation"
    End If
    If txtBrand.Text = "" Then eMsg = eMsg & vbCrLf & "Incomplete entry on field Brand"
    If txtModel.Text = "" Then eMsg = eMsg & vbCrLf & "Incomplete entry on field Model"
    If txtChassis.Text = "" Then eMsg = eMsg & vbCrLf & "Incomplete entry on field Chassis"
    
    ValidateSrvCust = eMsg
End Function
Function GetEngineOwner(MyEN As String) As String
    Dim rsEngine As New ADODB.Recordset
    Dim Qry As String
    '***Access
    'Qry = "SELECT SRVCUSTOMER.* From SRVCUSTOMER WHERE (((SRVCUSTOMER.ENGINE)='" & MyEN & "'));"
    Qry = "SELECT SRVCUSTOMER.* From SRVCUSTOMER WHERE SRVCUSTOMER.ENGINE='" & MyEN & "';"
    With rsEngine
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            GetEngineOwner = "Customer: " & !CUSTOMER & vbCrLf & "Address: " & !Address & vbCrLf & "Account No. " & !ACCOUNTNO
        Else
            GetEngineOwner = ""
        End If
        .Close
    End With
    
End Function
Private Sub cmdSrvcMember_Click()
    Call DisplaySrvcAcct
    tabMC.Visible = True
    ssCreditTab.Visible = False
    Me.Width = 13560
    tabMC.Tab = 0
    Me.Left = (MotoRep.ScaleWidth - Me.Width) \ 2
End Sub

Private Sub DisplaySrvcAcct()
    Dim rsAcctSrvc As New ADODB.Recordset
    Dim rsTerm As New ADODB.Recordset
    Dim Qry As String, i As Integer
    '***Access
    'Qry = "SELECT SRVCUSTOMER.CUSTOMER, SRVCUSTOMER.DESIGNATION, SRVCUSTOMER.BRAND, SRVCUSTOMER.MODEL, SRVCUSTOMER.ENGINE, SRVCUSTOMER.CHASSIS, SRVCUSTOMER.PLATENO, SRVCUSTOMER.ACCOUNTNO " & _
          "From SRVCUSTOMER Where (((SRVCUSTOMER.CUSTOMERID) = " & Val(txtCustID.Text) & ")) ORDER BY SRVCUSTOMER.CUSTOMER;"
    'mySQL
    Qry = "SELECT SRVCUSTOMER.CUSTOMER, SRVCUSTOMER.DESIGNATION, SRVCUSTOMER.BRAND, SRVCUSTOMER.MODEL, SRVCUSTOMER.ENGINE, SRVCUSTOMER.CHASSIS, SRVCUSTOMER.PLATENO, SRVCUSTOMER.ACCOUNTNO " & _
          "From SRVCUSTOMER Where SRVCUSTOMER.CUSTOMERID = " & Val(txtCustID.Text) & " ORDER BY SRVCUSTOMER.CUSTOMER;"
    
    
    rsAcctSrvc.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    If rsAcctSrvc.RecordCount = 0 Then
        flxMC.Clear
        flxMC.Rows = 2
        flxMC.FormatString = "   |^CUSTOMER |^DESIGNATION  |^BRAND |^MODEL |^ENGINE |^CHASSIS|^PLATENO "
    Else
        Set flxMC.DataSource = rsAcctSrvc
    End If
    flxMC.ColWidth(0) = 300
    For i = 1 To 7
        flxMC.ColWidth(i) = 2000
    Next i
    flxMC.ColWidth(8) = 0
    
    
    'Display Term
    'SQLAcctCustomer08
    '***Access
    'Qry = "SELECT ONACCTERMS.*, ONACCTERMS.CUSTOMERID, ONACCTERMS.STATUS From ONACCTERMS " & _
          "WHERE (((ONACCTERMS.CUSTOMERID)=" & Val(txtCustID.Text) & ") AND ((ONACCTERMS.STATUS)='Default'));"
    Qry = "SELECT ONACCTERMS.*, ONACCTERMS.CUSTOMERID, ONACCTERMS.STATUS From ONACCTERMS " & _
          "WHERE ONACCTERMS.CUSTOMERID=" & Val(txtCustID.Text) & " AND ONACCTERMS.STATUS='Default';"
    
    
    With rsTerm
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            If Not IsNull(!PayScheme) Then cboScheme.Text = !PayScheme
            txtGives.Text = !Npays
            txtTerm.Text = !Terms
        End If
        .Close
    End With
End Sub

Private Sub HideSrvcCustomer()
    tabMC.Visible = False
    Me.Width = 6420
End Sub

Private Sub cmdTerm_Click(Index As Integer)
    Dim rsTerm As New ADODB.Recordset
    Dim Qry As String
    Dim i As Integer, Ans As Integer
        
    Select Case Index
        Case 0  'Edit
            Frame3.Enabled = True
            cmdTerm(0).Visible = False
            cmdTerm(1).Visible = True
            cmdTerm(2).Visible = True
            For i = 0 To 3
                cmdTrans(i).Enabled = False
            Next i
            tabMC.TabEnabled(0) = False
            tabMC.TabEnabled(1) = False
            cmdSrvcMember.Enabled = False
            NewMC = True
            cboScheme.SetFocus
            
        Case 1  'Save
            If txtGives.Text = "" Then MsgBox "Incomplete entry on Number of Payments", vbOKOnly + vbExclamation, "Validation Error"
            If txtTerm.Text = "" Then MsgBox "Incomplete entry on Terms [No. of Days]", vbOKOnly + vbExclamation, "Validation Error"
            Ans = MsgBox("Do you want to Save Changes?", vbQuestion + vbYesNo, "Save")
            If Ans = vbNo Then Exit Sub
            'SQLAcctCustomer09
            '***Access
            'LYRcon.Execute "UPDATE ONACCTERMS SET ONACCTERMS.PAYSCHEME = '" & cboScheme.Text & "', ONACCTERMS.NPAYS = " & txtGives.Text & ", ONACCTERMS.TERMS = " & txtTerm.Text & _
                           " WHERE (((ONACCTERMS.CUSTOMERID)=" & txtCustID.Text & ") AND ((ONACCTERMS.STATUS)='Default'));"
            LYRcon.Execute "UPDATE ONACCTERMS SET ONACCTERMS.PAYSCHEME = '" & cboScheme.Text & "', ONACCTERMS.NPAYS = " & txtGives.Text & ", ONACCTERMS.TERMS = " & txtTerm.Text & _
                           " WHERE ONACCTERMS.CUSTOMERID=" & txtCustID.Text & " AND ONACCTERMS.STATUS='Default';"
            MsgBox "On-Account Term Successfully Saved!", vbInformation + vbOKOnly
            Frame3.Enabled = False
            cmdTerm(0).Visible = True
            cmdTerm(1).Visible = False
            cmdTerm(2).Visible = False
            For i = 0 To 3
                cmdTrans(i).Enabled = True
            Next i
            tabMC.TabEnabled(0) = True
            tabMC.TabEnabled(1) = True
            cmdSrvcMember.Enabled = True
        Case 2  'Cancel
            Ans = MsgBox("Do you want to Cancel " & IIf(NewMC, "Entry", "Editing") & "?", vbQuestion + vbYesNo, "Cancel")
            If Ans = vbNo Then Exit Sub
            'Display Term
            'SQLAcctCustomer10
            '***Access
            'Qry = "SELECT ONACCTERMS.*, ONACCTERMS.CUSTOMERID, ONACCTERMS.STATUS From ONACCTERMS " & _
                  "WHERE (((ONACCTERMS.CUSTOMERID)=" & Val(txtCustID.Text) & ") AND ((ONACCTERMS.STATUS)='Default'));"
            Qry = "SELECT ONACCTERMS.*, ONACCTERMS.CUSTOMERID, ONACCTERMS.STATUS From ONACCTERMS " & _
                  "WHERE ONACCTERMS.CUSTOMERID=" & Val(txtCustID.Text) & " AND ONACCTERMS.STATUS='Default';"
            With rsTerm
                .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
                If .RecordCount > 0 Then
                    .MoveFirst
                    If Not IsNull(!PayScheme) Then cboScheme.Text = !PayScheme
                    txtGives.Text = !Npays
                    txtTerm.Text = !Terms
                End If
                .Close
            End With
            cmdTerm(0).Visible = True
            cmdTerm(1).Visible = False
            cmdTerm(2).Visible = False
            For i = 0 To 3
                cmdTrans(i).Enabled = True
            Next i
            tabMC.TabEnabled(0) = True
            tabMC.TabEnabled(1) = True
            cmdSrvcMember.Enabled = True
            Frame3.Enabled = False
    End Select

End Sub
Private Sub cmdTrans_Click(Index As Integer)
    Dim vMsg As String, QrySrch As String, MyAcc As String
    Dim myCustID As Integer, Ans As Integer
    
    Select Case Index
        Case 0  '   "new"
            '***Access
            'QrySrch = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE (((USERACCESS.EMPID)=" & UserID & ") AND ((USERACCESS.MENUINDX)=10));"
            QrySrch = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE USERACCESS.EMPID=" & UserID & " AND USERACCESS.MENUINDX=10;"
            MyAcc = GetFieldValue(QrySrch)
            If MyAcc <> "FULL ACCESS" Then
                MsgBox "You don't have full access to this transaction." & vbCrLf & "Please contact your System Administrator", vbExclamation + vbOKOnly, "Access Denied"
                Exit Sub
            End If
            If DispBtn Then
                Call BtnSaveCancel(True)
            Else
                BtnIndx = 3
                tmrFloat.Enabled = True
            End If
                
            Call LockData(False)
            Call ClearEntries
            cmdSrvcMember.Visible = False
            cmdCC.Visible = False
            Call HideSrvcCustomer
            AddNew = True
            txtCustomer.Visible = False
            cboAcctCust.Visible = True
            cboCustType.SetFocus
        Case 1      '"edit"
            '***Access
            'QrySrch = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE (((USERACCESS.EMPID)=" & UserID & ") AND ((USERACCESS.MENUINDX)=10));"
            QrySrch = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE USERACCESS.EMPID=" & UserID & " AND USERACCESS.MENUINDX=10;"
            MyAcc = GetFieldValue(QrySrch)
            If MyAcc <> "FULL ACCESS" Then
                MsgBox "You don't have full access to this transaction." & vbCrLf & "Please contact your System Administrator", vbExclamation + vbOKOnly, "Access Denied"
                Exit Sub
            End If
            If txtCustID.Text = "" Then
                MsgBox "No records found!", vbExclamation + vbOKOnly, "Warning"
                txtCustomer.SetFocus
                Exit Sub
            End If
            If DispBtn Then
                Call BtnSaveCancel(True)
            Else
                BtnIndx = 3
                tmrFloat.Enabled = True
            End If
            Call LockData(False)
            AddNew = False
            cmdSrvcMember.Visible = False
            cmdCC.Visible = False
            Call HideSrvcCustomer
            txtCustomer.Visible = False
            CustSrch = True
            cboAcctCust.Text = txtCustomer.Text
            cboAcctCust.Visible = True
            txtEditCustID.Text = ""
        Case 2          '"find"
            Set frmSearch.FrmSource = AcctCustomer
            QrySrch = "SELECT CUSTOMER.CUSTOMERID, CUSTOMER.CUSTOMERNAME " & _
                      "From CUSTOMER ORDER BY CUSTOMER.CUSTOMERNAME;"
            frmSearch.QrySearch = QrySrch
            frmSearch.lblCap.Caption = "Search for Customer"
            frmSearch.Show vbModal
        Case 3          '"quit"
            
            'MotoRep.tbMain.Visible = True
            MotoRep.toolBarButton.Visible = True
            'MotoRep.imgBG.Visible = True
            Unload Me
        
        Case 4              '"save"
            vMsg = ValidateEntry
            If vMsg <> "" Then
                MsgBox vMsg, vbOKOnly + vbCritical, "Validation Error"
                Exit Sub
            End If
            Ans = MsgBox("Do you want to Save Customer?", vbQuestion + vbYesNo, "Save")
            If Ans = vbNo Then Exit Sub
            Call SaveCustomer
            MsgBox "Record successfully saved!", vbInformation + vbOKOnly, "Save"
            Call LockData(True)
            'txtCustomer.SetFocus
            If DispBtn Then
                Call BtnSaveCancel(False)
            Else
                BtnIndx = 5
                tmrFloatRev.Enabled = True
            End If
            If ChkApply.Value = 1 Then
                cmdSrvcMember.Visible = True
            Else
                cmdSrvcMember.Visible = False
            End If
            txtCustomer.Text = cboAcctCust.Text
            txtCustomer.Visible = True
            cboAcctCust.Visible = False
        Case 5          '"cancel"
            Ans = MsgBox("Do you want to cancel Customer?", vbQuestion + vbYesNo, "Cancel")
            If Ans = vbNo Then Exit Sub
            If AddNew Then
                Call ClearEntries
                txtCustomer.Visible = True
                cboAcctCust.Visible = False
            Else
                myCustID = txtCustID.Text
                Call ClearEntries
                Call ShowRecord(myCustID)
                txtCustomer.Visible = True
                cboAcctCust.Visible = False
            End If
            Call LockData(True)
            txtCustomer.SetFocus
            If DispBtn Then
                Call BtnSaveCancel(False)
            Else
                BtnIndx = 5
                tmrFloatRev.Enabled = True
            End If
    End Select
End Sub

Private Sub flxCreditInfo_Click()
    ActRow = flxCreditInfo.Row
End Sub

Private Sub flxCreditInfo_EnterCell()
    ActRow = flxCreditInfo.Row
End Sub

Private Sub flxCreditInfo_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
    If Button = vbRightButton And ActRow > 0 Then
        If flxCreditInfo.TextMatrix(ActRow, 4) <> "" Then
            Call PopupMenu(mnuCC)
        End If
    End If
End Sub

Private Sub flxCust_Click()
    txtMCust.Text = flxCust.TextMatrix(flxCust.Row, 1)
End Sub
Private Sub flxCust_EnterCell()
    CustSrch = True
    txtMCust.Text = flxCust.TextMatrix(flxCust.Row, 1)
End Sub

Private Sub flxCust_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        flxCust.Visible = False
        txtMCust.SetFocus
    End If
End Sub
Private Sub flxCust_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call txtMCust_Change
        Call txtMCust_LostFocus
        flxCust.Visible = False
        txtDesig.SetFocus
    End If
End Sub
Private Sub flxMC_Click()
    ActRow = flxMC.Row
End Sub

Private Sub flxMC_EnterCell()
    ActRow = flxMC.Row
End Sub

Private Sub flxMC_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
    If Button = vbRightButton And ActRow > 0 Then
        If flxMC.TextMatrix(ActRow, 8) <> "" Then
            Call PopupMenu(mnuMC)
        End If
    End If
End Sub

Private Sub Form_Load()
    Dim Qry As String
    
    cboScheme.AddItem "WEEKLY"
    cboScheme.AddItem "BI-MONTHLY"
    cboScheme.AddItem "MONTHLY"
    Call LockData(True)
    'Call OptType_Click(1)
    Call HideSrvcCustomer
    
    
    
    Timer1.Enabled = True
    MotoRep.sBar.Panels(6).Text = "On-Account Customer Entry Form"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    MotoRep.sBar.Panels(6).Text = ""
    If rsAcctCust.State = adStateOpen Then rsAcctCust.Close
End Sub


Private Sub mnuEditCreditCard_Click()
    Dim rsEditCC As New ADODB.Recordset
    Dim Ans As Integer, i As Integer
    Dim Qry As String
    
    Ans = MsgBox("Do you want to Edit Credit Card # " & flxCreditInfo.TextMatrix(ActRow, 3) & "?", vbQuestion + vbYesNo, "Edit")
    If Ans = vbYes Then
        '
        Qry = "call sp_mrcretails.sp_credicard_details('" & sysBranchCode & "'," & Val(flxCreditInfo.TextMatrix(ActRow, 4)) & ")"
        If rsEditCC.State = adStateOpen Then rsEditCC.Close
        rsEditCC.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        
        With rsEditCC
            If .RecordCount > 0 Then
                CreditCardID = Val(flxCreditInfo.TextMatrix(ActRow, 4))
                txtCardHolder.Text = !CCNAME
                txtCardNo.Text = !CCNO
                txtCreditBank.Text = !CCBANKNAME
            End If
            .Close
        End With
    End If
    ssCreditTab.Tab = 1
    btnCC(0).Visible = False
    btnCC(1).Visible = True
    btnCC(2).Visible = True
    For i = 0 To 3
        cmdTrans(i).Enabled = False
    Next i
    ssCreditTab.TabEnabled(0) = False
    cmdCC.Enabled = False
    formAction = "edit"
    txtCardHolder.SetFocus
End Sub

Private Sub mnuEditSrvCust_Click()
    Dim Ans As Integer, i As Integer
    Dim Qry As String
    
    Ans = MsgBox("Do you want to Edit Customer " & flxMC.TextMatrix(ActRow, 1) & "?", vbQuestion + vbYesNo, "Edit")
    If Ans = vbYes Then
        
        '***Access
        'Qry = "SELECT SRVCUSTOMER.ACCOUNTNO, SRVCUSTOMER.CUSTOMER, SRVCUSTOMER.ENGINE " & _
              "From SRVCUSTOMER Where SRVCUSTOMER.ADDTLMC = False ORDER BY SRVCUSTOMER.CUSTOMER;"
        'mySQL
        Qry = "SELECT SRVCUSTOMER.ACCOUNTNO, SRVCUSTOMER.CUSTOMER, SRVCUSTOMER.ENGINE " & _
              "From SRVCUSTOMER Where SRVCUSTOMER.ADDTLMC = 0 ORDER BY SRVCUSTOMER.CUSTOMER;"
        If rsOnAcct.State = adStateOpen Then rsOnAcct.Close
        rsOnAcct.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        Set flxCust.DataSource = rsOnAcct
        With flxCust
            .Width = txtMCust.Width
            .ColWidth(0) = 0
            .ColWidth(1) = txtMCust.Width
            .ColWidth(2) = 0
        End With
        With flxMC
            AcctSrch = True
            txtMCust.Text = .TextMatrix(ActRow, 1)
            txtDesig.Text = .TextMatrix(ActRow, 2)
            txtBrand.Text = .TextMatrix(ActRow, 3)
            txtModel.Text = .TextMatrix(ActRow, 4)
            txtEngine.Text = .TextMatrix(ActRow, 5)
            txtChassis.Text = .TextMatrix(ActRow, 6)
            txtPlateNo.Text = .TextMatrix(ActRow, 7)
            txtAcct.Text = .TextMatrix(ActRow, 8)
            txtOAcustID.Text = .TextMatrix(ActRow, 8)
        End With
    End If
    tabMC.Tab = 1
    cmdMC(0).Visible = False
    cmdMC(1).Visible = True
    cmdMC(2).Visible = True
    For i = 0 To 3
        cmdTrans(i).Enabled = False
    Next i
    tabMC.TabEnabled(0) = False
    tabMC.TabEnabled(2) = False
    cmdSrvcMember.Enabled = False
End Sub


Private Sub Text1_Change()

End Sub


Private Sub Timer1_Timer()
    Dim rsCType As New ADODB.Recordset
    Dim rsAdrs As New ADODB.Recordset
    Dim Qry As String
    
    With rsCType
        '***Access
        'Qry = "SELECT SETTINGS.* From SETTINGS " & _
              "WHERE (((SETTINGS.DESCRIPTION)='Customer Type')) ORDER BY SETTINGS.DESCRIPTION;"
        Qry = "SELECT SETTINGS.* From SETTINGS " & _
              "WHERE SETTINGS.DESCRIPTION='Customer Type' ORDER BY SETTINGS.DESCRIPTION;"
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                cboCustType.AddItem !Value
                .MoveNext
            Loop
        End If
        .Close
    End With
    Qry = "SELECT SETTINGS.REMARKS From SETTINGS " & _
              "WHERE SETTINGS.DESCRIPTION = 'BM DISCOUNT' AND SETTINGS.VALUE = '8/01/2016';"
    If UserID = Val(GetFieldValue(Qry)) Then cboCustType.AddItem "WHOLESALE"
    
    
    Qry = "SELECT settings.VALUE as Address From settings Where settings.DESCRIPTION LIKE 'REGION%' Order By settings.Value"
    With rsAdrs
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                txtCity.AddItem !Address
                .MoveNext
            Loop
        End If
        .Close
    End With
    

    Qry = "SELECT CUSTOMER.CUSTOMERID, CUSTOMER.CUSTOMERNAME, CUSTOMER.ADDRESS, CUSTOMER.CONTACTNO, CUSTOMER.CONTACTPERSON, CUSTOMER.TIN, CUSTOMER.CRLIMIT, CUSTOMER.TERMS, CUSTOMER.CTYPE " & _
          "From CUSTOMER ORDER BY CUSTOMER.CUSTOMERNAME;"
    With rsAcctCust
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            cboAcctCust.Clear
            Do While Not .EOF
                cboAcctCust.AddItem !CustomerName
                .MoveNext
            Loop
        End If
    End With
    Timer1.Enabled = False
End Sub

Private Sub tmrDrop_Timer()
    cmdTrans(BtnIndx).Top = cmdTrans(BtnIndx).Top + 50
    If cmdTrans(BtnIndx).Top >= 0 Then
        cmdTrans(BtnIndx).Top = 0
        BtnIndx = BtnIndx + 1
        If BtnIndx = 6 Then 'last index + 1
            tmrDrop.Enabled = False
        End If
    End If
End Sub

Private Sub tmrDropRev_Timer()
    cmdTrans(BtnIndx).Top = cmdTrans(BtnIndx).Top + 50
    If cmdTrans(BtnIndx).Top >= 0 Then
        cmdTrans(BtnIndx).Top = 0
        BtnIndx = BtnIndx + 1
        If BtnIndx = 4 Then
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
            BtnIndx = 4
            cmdTrans(4).Top = -375
            cmdTrans(4).Left = 0
            cmdTrans(4).Visible = True
            cmdTrans(5).Top = -375
            cmdTrans(5).Left = cmdTrans(4).Width
            cmdTrans(5).Visible = True
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
        If BtnIndx = 3 Then
            tmrFloatRev.Enabled = False
            BtnIndx = 0
            For i = 0 To 3
                cmdTrans(i).Visible = True
            Next i
            tmrDropRev.Enabled = True
        End If
    End If
End Sub

Private Sub txtAddress_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 40 Then
        SendKeys "{TAB}"
    ElseIf KeyCode = 38 Then
        SendKeys "+{TAB}"
    End If
End Sub

Private Sub txtAddress_KeyPress(KeyAscii As Integer)
    If Chr(KeyAscii) = "'" Or Chr(KeyAscii) = "\" Then KeyAscii = 0
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub



Private Sub txtAgent_KeyPress(KeyAscii As Integer)
    If Chr(KeyAscii) = "'" Or Chr(KeyAscii) = "\" Then KeyAscii = 0
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
End Sub

Private Sub txtBranch_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 40 Then
        SendKeys "{TAB}"
    ElseIf KeyCode = 38 Then
        SendKeys "+{TAB}"
    End If

End Sub

Private Sub txtBranch_KeyPress(KeyAscii As Integer)
    If Chr(KeyAscii) = "'" Or Chr(KeyAscii) = "\" Then KeyAscii = 0
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then SendKeys "{TAB}"

End Sub

Private Sub txtBrand_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtModel.SetFocus
End Sub


Private Sub txtCardHolder_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtCardHolder, KeyAscii, ValidChars)
    If KeyAscii = 13 Then
        txtCardNo.SetFocus
    End If
End Sub

Private Sub txtCardNo_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtCardNo, KeyAscii, NumbersOnly)
    If KeyAscii = 13 Then
        txtCreditBank.SetFocus
    End If
End Sub

Private Sub txtChassis_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtPlateNo.SetFocus
End Sub


Private Sub txtCity_Change()
    Dim i As Long, j As Long
    Dim strPartial As String, strTotal As String
    
    If StatSrch Then
        StatSrch = False
        Exit Sub
    End If
    With txtCity
        strPartial = .Text
        i = SendMessage(.hwnd, CB_FINDSTRING, -1, ByVal strPartial) 'If match found, append unmatched characters
        If i <> -1 Then
            strTotal = .List(i)
            'txtSrcID.Tag = ""
            j = Len(strTotal) - Len(strPartial)
            If j <> 0 Then
                StatSrch = True
                .SelText = Right$(strTotal, j)
                .SelStart = Len(strPartial)
                .SelLength = j
            End If
        Else
            'txtSrcID.Tag = "invalid"
        End If
    End With

End Sub

Private Sub txtCity_Click()
    Call txtCity_Change
End Sub

Private Sub txtCity_GotFocus()
   Dim Tmp

   Const CB_SHOWDROPDOWN = &H14F
   Tmp = SendMessage(txtCity.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)

End Sub

Private Sub txtCity_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 40 Then
        SendKeys "{TAB}"
    ElseIf KeyCode = 38 Then
        SendKeys "+{TAB}"
    End If
End Sub

Private Sub txtCity_KeyPress(KeyAscii As Integer)
    If Chr(KeyAscii) = "'" Or Chr(KeyAscii) = "\" Then KeyAscii = 0
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then SendKeys "{TAB}"

End Sub

Private Sub txtCity_LostFocus()
    Dim MyAdrs As New ADODB.Recordset
    Dim Qry As String

    Qry = "SELECT settings.VALUE, settings.REMARKS, settings.DESCRIPTION From settings Where settings.VALUE = '" & txtCity.Text & "';"
    With MyAdrs
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            txtProvince.Text = !Remarks
            txtRegion.Text = !Description
        Else
            txtProvince.Text = ""
            txtRegion.Text = ""
        End If
        .Close
    End With

End Sub

Private Sub txtCLimit_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 40 Then
        SendKeys "{TAB}"
    ElseIf KeyCode = 38 Then
        SendKeys "+{TAB}"
    End If
End Sub

Private Sub txtCLimit_KeyPress(KeyAscii As Integer)
    Dim strValid As String
    
    strValid = "1234567890."
    If KeyAscii > 26 Then
        If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
        
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtConPers_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 40 Then
        SendKeys "{TAB}"
    ElseIf KeyCode = 38 Then
        SendKeys "+{TAB}"
    End If
End Sub

Private Sub txtConPers_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtContactNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 40 Then
        SendKeys "{TAB}"
    ElseIf KeyCode = 38 Then
        SendKeys "+{TAB}"
    End If
End Sub

Private Sub txtContactNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub



Private Sub txtCreditBank_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtCreditBank, KeyAscii, ValidChars)
    If KeyAscii = 13 Then
        Exit Sub
    End If
End Sub

Private Sub txtCustomer_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 40 Then
        SendKeys "{TAB}"
    ElseIf KeyCode = 38 Then
        SendKeys "+{TAB}"
    End If
End Sub

Private Sub txtCustomer_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub



Private Sub txtCustTerm_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 40 Then
        SendKeys "{TAB}"
    ElseIf KeyCode = 38 Then
        SendKeys "+{TAB}"
    End If

End Sub

Private Sub txtCustTerm_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtAgent.SetFocus
End Sub

Private Sub txtDesig_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtBrand.SetFocus
End Sub


Private Sub txtEngine_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtChassis.SetFocus
End Sub


Private Sub txtGives_KeyPress(KeyAscii As Integer)
    Dim strValid As String
    
    strValid = "1234567890"
    If keysacii > 26 Then
        If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
    If KeyAscii = 13 Then txtTerm.SetFocus
End Sub

Private Sub txtGives_LostFocus()
    Dim iPS As Integer, TPrice As Currency, sPay As Date
    
    If cboScheme.Text = "" Or Val(txtGives.Text) = 0 Then Exit Sub
    If cboScheme.Text = "WEEKLY" Then iPS = 7
    If cboScheme.Text = "BI-MONTHLY" Then iPS = 15
    If cboScheme.Text = "MONTHLY" Then iPS = 30
        'Compute for term base on scheme
    If iPS = 7 Then
        m = 0
        If Val(txtGives.Text) >= 4 Then
            m = Val(txtGives.Text) \ 4
            w = Val(txtGives.Text) Mod 4
        Else
            w = Val(txtGives.Text)
        End If
        txtTerm.Text = (m * 30) + (w * 7)
    Else
        txtTerm.Text = Val(txtGives.Text) * iPS
    End If

End Sub

Private Sub txtMCust_Change()
    
    With flxCust
        If AcctSrch Then
            AcctSrch = False
            Exit Sub
        End If
        If txtMCust.Text = "" Then .Visible = False: Exit Sub
        rsOnAcct.Find "[CUSTOMER] LIKE '" + txtMCust.Text + "%'", , adSearchForward, 1
            If Not rsOnAcct.EOF Then
            .Top = txtMCust.Top + txtMCust.Height + 30
            .Left = txtMCust.Left
            .Width = txtMCust.Width
            .Visible = True
            If Not CustBln Then
                txtMCust.Tag = Len(txtMCust.Text)
                CustBln = True
            End If
            .TopRow = rsOnAcct.AbsolutePosition - 1
            .Row = rsOnAcct.AbsolutePosition - 1
            .Col = 0: .ColSel = 2
            AcctSrch = True
            txtMCust.Text = rsOnAcct!CUSTOMER
            txtAcct.Text = flxCust.TextMatrix(flxCust.Row, 0)
            
            txtMCust.SelStart = Val(txtMCust.Tag)
            txtMCust.SelLength = Len(txtMCust.Text)
            CustBln = False
        Else
            txtAcct.Text = ""
            flxCust.Visible = False
        End If
    End With
End Sub
                                                                   
Private Sub txtMCust_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDown
            If flxCust.Visible Then
                flxCust.SetFocus
            Else
                txtAddress.SetFocus
            End If
        Case vbKeyEscape
            txtMCust.Text = ""
            flxCust.Visible = False
        Case vbKeyDelete
            AcctSrch = True
        Case vbKeyBack
            AcctSrch = True
    End Select
End Sub

Private Sub txtMCust_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        flxCust.Visible = False
        txtDesig.SetFocus
    End If
End Sub

Private Sub txtMCust_LostFocus()
    Dim rsAllMC As New ADODB.Recordset
    Dim Qry As String
    
    If Val(txtAcct.Text) = 0 Then
        flxAllMC.Rows = 1
        cboEngine.Clear
        Exit Sub
    End If
    '***Access
    'Qry = "SELECT SRVCUSTOMER.ACCOUNTNO, SRVCUSTOMER.CUSTOMER, SRVCUSTOMER.BRAND, SRVCUSTOMER.MODEL, SRVCUSTOMER.ENGINE, SRVCUSTOMER.CHASSIS, SRVCUSTOMER.PLATENO, SRVCUSTOMER.ONACCOUNT, SRVCUSTOMER.CUSTOMERID, SRVCUSTOMER.ADDTLMC " & _
          "From SRVCUSTOMER WHERE (((SRVCUSTOMER.CUSTOMER)='" & txtMCust.Text & "')) ORDER BY SRVCUSTOMER.ONACCOUNT DESC"
    
    Qry = "SELECT SRVCUSTOMER.ACCOUNTNO, SRVCUSTOMER.CUSTOMER, SRVCUSTOMER.BRAND, SRVCUSTOMER.MODEL, SRVCUSTOMER.ENGINE, SRVCUSTOMER.CHASSIS, SRVCUSTOMER.PLATENO, SRVCUSTOMER.ONACCOUNT, SRVCUSTOMER.CUSTOMERID, SRVCUSTOMER.ADDTLMC " & _
          "From SRVCUSTOMER WHERE SRVCUSTOMER.CUSTOMER='" & txtMCust.Text & "' ORDER BY SRVCUSTOMER.ONACCOUNT DESC"

    With rsAllMC
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        flxAllMC.Rows = 1
        cboEngine.Clear
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                '***Access
                'flxAllMC.AddItem !ACCOUNTNO & vbTab & !CUSTOMER & vbTab & IIf(IsNull(!BRAND), "", !BRAND) & vbTab & !MODEL & vbTab & !ENGINE & vbTab & !CHASSIS & vbTab & IIf(IsNull(!PLATENO), "", !PLATENO) & vbTab & IIf(IsNull(!ONACCOUNT), "", !ONACCOUNT) & vbTab & IIf(IsNull(!CUSTOMERID), "", !CUSTOMERID) & vbTab & IIf(IsNull(!ADDTLMC), "False", !ADDTLMC)
                'mySQL
                flxAllMC.AddItem !ACCOUNTNO & vbTab & !CUSTOMER & vbTab & IIf(IsNull(!BRAND), "", !BRAND) & vbTab & !MODEL & vbTab & !ENGINE & vbTab & !CHASSIS & vbTab & IIf(IsNull(!PLATENO), "", !PLATENO) & vbTab & IIf(IsNull(!ONACCOUNT), "", !ONACCOUNT) & vbTab & IIf(IsNull(!CUSTOMERID), "", !CUSTOMERID) & vbTab & IIf(IsNull(!ADDTLMC), 0, !ADDTLMC)
                cboEngine.AddItem !ENGINE
                .MoveNext
            Loop
        End If
        .Close
    End With
    With flxAllMC
        If .Rows = 1 Then Exit Sub
        If flxAllMC.Rows = 2 And flxAllMC.TextMatrix(1, 7) = "True" Then Exit Sub
        txtAcct.Text = .TextMatrix(1, 0)
        txtBrand.Text = .TextMatrix(1, 2) 'Brand
        txtOAcustID.Text = .TextMatrix(1, 8)    'CustomerID
        txtModel.Text = .TextMatrix(1, 3)  'Model
        txtEngine.Text = .TextMatrix(1, 4)  'Engine
        txtChassis.Text = .TextMatrix(1, 5)  'chassis
        txtPlateNo.Text = .TextMatrix(1, 6)  'Plate
    End With
    If cboEngine.ListCount > 1 Then
        lblAvlMC.Caption = cboEngine.ListCount & " Available MC"
        With cboEngine
            .Top = txtEngine.Top
            .Left = txtEngine.Left
            .Width = txtEngine.Width
            '.Height = txtEngine.Height
            .Visible = True
            .Text = flxAllMC.TextMatrix(1, 4)
            txtEngine.Visible = False
            txtEngine.Text = .Text
        End With
    Else
        lblAvlMC.Caption = ""
        cboEngine.Visible = False
        txtEngine.Visible = True
    End If
End Sub

Private Sub txtModel_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then SendKeys "+{TAB}" 'txtEngine.SetFocus
End Sub



Private Sub txtPlateNo_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then cmdMC(1).SetFocus
End Sub

Private Sub txtProvince_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 40 Then
        SendKeys "{TAB}"
    ElseIf KeyCode = 38 Then
        SendKeys "+{TAB}"
    End If
End Sub

Private Sub txtProvince_KeyPress(KeyAscii As Integer)
    If Chr(KeyAscii) = "'" Or Chr(KeyAscii) = "\" Then KeyAscii = 0
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtRegion_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 40 Then
        SendKeys "{TAB}"
    ElseIf KeyCode = 38 Then
        SendKeys "+{TAB}"
    End If
End Sub

Private Sub txtRegion_KeyPress(KeyAscii As Integer)
    If Chr(KeyAscii) = "'" Or Chr(KeyAscii) = "\" Then KeyAscii = 0
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

'mysql
Private Sub txtTerm_KeyPress(KeyAscii As Integer)
    Dim strValid As String
    
    strValid = "1234567890"
    If keysacii > 26 Then
        If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
    If KeyAscii = 13 Then cmdTerm(1).SetFocus
End Sub

Private Sub txtTerm_LostFocus()
    Dim iPS As Integer
    
    If cboScheme.Text = "" Then Exit Sub
    If cboScheme.Text = "WEEKLY" Then iPS = 7
    If cboScheme.Text = "BI-MONTHLY" Then iPS = 15
    If cboScheme.Text = "MONTHLY" Then iPS = 30
    If Val(txtTerm.Text) >= iPS Then
        txtGives.Text = Val(txtTerm.Text) \ iPS
    Else
        txtGives.Text = 1
    End If
End Sub

Private Sub txtTIN_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 40 Then
        SendKeys "{TAB}"
    ElseIf KeyCode = 38 Then
        SendKeys "+{TAB}"
    End If
End Sub

Private Sub txtTIN_KeyPress(KeyAscii As Integer)
    
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub


