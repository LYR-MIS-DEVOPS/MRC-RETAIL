VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form ExpCashRet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cash Return"
   ClientHeight    =   6840
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   8685
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6840
   ScaleWidth      =   8685
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtAmount 
      Height          =   285
      Left            =   6480
      TabIndex        =   45
      Top             =   8040
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtParticular 
      Height          =   285
      Left            =   6480
      MaxLength       =   30
      TabIndex        =   44
      Top             =   7560
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   240
      Top             =   6120
   End
   Begin VB.ComboBox cboCashier 
      Height          =   315
      Left            =   1080
      TabIndex        =   43
      Top             =   6105
      Width           =   2775
   End
   Begin VB.ListBox lstVoucher 
      Height          =   645
      Left            =   4320
      TabIndex        =   41
      Top             =   6000
      Visible         =   0   'False
      Width           =   855
   End
   Begin TabDlg.SSTab tabCV 
      Height          =   1935
      Left            =   840
      TabIndex        =   30
      Top             =   7440
      Visible         =   0   'False
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   3413
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "ExpCashRet.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label4"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "txtCVno"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtExpAmt"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtCVamt"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtCVdate"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtAcctDesc"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Frame2"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      Begin VB.Frame Frame2 
         Height          =   615
         Left            =   2760
         TabIndex        =   38
         Top             =   1080
         Width           =   3015
         Begin lvButton.lvButtons_H cmdOK 
            Height          =   390
            Left            =   120
            TabIndex        =   39
            Top             =   165
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   688
            Caption         =   "&Ok"
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
            Image           =   "ExpCashRet.frx":001C
            cBack           =   -2147483633
         End
         Begin lvButton.lvButtons_H cmdCancel 
            Height          =   390
            Left            =   1560
            TabIndex        =   40
            Top             =   160
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   688
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
            Image           =   "ExpCashRet.frx":05B6
            cBack           =   -2147483633
         End
      End
      Begin VB.TextBox txtAcctDesc 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   37
         Top             =   1440
         Width           =   2295
      End
      Begin VB.TextBox txtCVdate 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   36
         Text            =   "C.V. Date :"
         Top             =   1080
         Width           =   2295
      End
      Begin VB.TextBox txtCVamt 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   35
         Text            =   "C.V. Amount :"
         Top             =   720
         Width           =   2295
      End
      Begin VB.TextBox txtExpAmt 
         Height          =   285
         Left            =   4440
         TabIndex        =   34
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox txtCVno 
         Height          =   285
         Left            =   960
         TabIndex        =   32
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Actual Expense:"
         Height          =   255
         Left            =   3000
         TabIndex        =   33
         Top             =   375
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "C.V.  No."
         Height          =   255
         Left            =   240
         TabIndex        =   31
         Top             =   375
         Width           =   855
      End
   End
   Begin VB.TextBox txtTotal 
      Alignment       =   1  'Right Justify
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
      Left            =   6720
      MaxLength       =   50
      TabIndex        =   26
      Tag             =   "XYF"
      Top             =   4920
      Width           =   1815
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
   Begin VB.TextBox txtCRid 
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
      Left            =   5760
      MaxLength       =   50
      TabIndex        =   21
      Tag             =   "XYF"
      Top             =   6105
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
            Picture         =   "ExpCashRet.frx":0B50
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ExpCashRet.frx":1A2A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ExpCashRet.frx":2904
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ExpCashRet.frx":37DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ExpCashRet.frx":46B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ExpCashRet.frx":5592
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
      Left            =   1080
      Locked          =   -1  'True
      MaxLength       =   250
      TabIndex        =   0
      Tag             =   "XYF"
      Top             =   5520
      Width           =   7455
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   120
      TabIndex        =   1
      Tag             =   "XYF"
      Top             =   720
      Width           =   8535
      Begin VB.ComboBox cboType 
         Height          =   315
         Left            =   1200
         TabIndex        =   47
         Top             =   604
         Width           =   3615
      End
      Begin VB.TextBox txtBranch 
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
         MaxLength       =   50
         TabIndex        =   28
         Tag             =   "XYF"
         Top             =   240
         Width           =   3615
      End
      Begin VB.TextBox txtCRDate 
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
      Begin VB.TextBox txtCRno 
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
         Caption         =   "TYPE:"
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
         TabIndex        =   46
         Tag             =   "XYF"
         Top             =   634
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "RET. DATE"
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
         Caption         =   "RET. NO."
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
         Caption         =   "BRANCH:"
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
         Left            =   360
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
      Image           =   "ExpCashRet.frx":646C
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
      Image           =   "ExpCashRet.frx":6A06
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
      Image           =   "ExpCashRet.frx":6FA0
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
      Image           =   "ExpCashRet.frx":753A
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
      Image           =   "ExpCashRet.frx":7AD4
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
      Image           =   "ExpCashRet.frx":806E
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
      Image           =   "ExpCashRet.frx":8608
      cBack           =   -2147483633
   End
   Begin MSFlexGridLib.MSFlexGrid flxCR 
      Height          =   3015
      Left            =   120
      TabIndex        =   25
      Tag             =   "XYF"
      Top             =   1920
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   5318
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      RowHeightMin    =   315
      BackColorBkg    =   16777215
      ScrollBars      =   2
      FormatString    =   "     |<  P A R T I C U L A R S                         |>          A M O U N T         "
   End
   Begin lvButton.lvButtons_H cmdAddCV 
      Height          =   375
      Left            =   120
      TabIndex        =   29
      Tag             =   "xyf"
      Top             =   5040
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
      Caption         =   "Add Voucher"
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
      ImgAlign        =   3
      Image           =   "ExpCashRet.frx":8BA2
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdPrnAtt 
      Height          =   375
      Left            =   1800
      TabIndex        =   42
      Tag             =   "xyf"
      Top             =   5040
      Visible         =   0   'False
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
      Caption         =   "Attachement"
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
      ImgAlign        =   3
      Image           =   "ExpCashRet.frx":913C
      cBack           =   -2147483633
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      Caption         =   "T O T A L:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5040
      TabIndex        =   27
      Tag             =   "XYF"
      Top             =   4980
      Width           =   1575
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
      Left            =   5760
      TabIndex        =   15
      Tag             =   "XYF"
      Top             =   6480
      Width           =   2775
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "RECEIVED BY"
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
      Left            =   1080
      TabIndex        =   14
      Tag             =   "XYF"
      Top             =   6480
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
      Top             =   5550
      Width           =   1455
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   6240
      Top             =   8640
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   8775
      m_MDI_Height    =   7260
   End
   Begin VB.Label lblHeading 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "CASH RETURN"
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
      Left            =   6210
      TabIndex        =   5
      Top             =   450
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
      Caption         =   "CASH RETURN"
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
      Left            =   6240
      TabIndex        =   6
      Top             =   480
      Width           =   2415
   End
   Begin VB.Menu mnuFlx 
      Caption         =   "FlxMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuAdd 
         Caption         =   "Add Row"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete Row"
      End
   End
End
Attribute VB_Name = "ExpCashRet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim ActRow As Integer, ActCol As Integer, BtnIndx As Integer
Dim JustScroll As Boolean, AddNew As Boolean
Public SrcSumm As Boolean

Sub ClearEntries()
    
    txtBranch.Text = ""
    txtCRno.Text = ""
    txtCRDate.Text = ""
    txtRemarks.Text = ""
    txtTotal.Text = ""
    cboCashier.Text = ""
    txtAprvBy.Text = ""
    flxCR.Rows = 1
    flxCR.Rows = 2
    txtPrnCopy.Text = ""
    txtCRid.Text = ""
End Sub

Sub LockData(bln As Boolean)
    Dim i As Integer
    
    Frame1.Enabled = Not bln
    cboCashier.Locked = bln
    txtAprvBy.Locked = bln
    txtRemarks.Locked = bln
    flxCR.Enabled = Not bln
    cmdAddCV.Enabled = Not bln
End Sub

Public Sub ShowRecord()
    Dim rsCashReturn As New ADODB.Recordset
    Dim rsCRdtls As New ADODB.Recordset
    Dim rsCV As New ADODB.Recordset
    Dim Qry As String
    Dim i As Integer, Ttl As Currency, Expns As Currency
    
    'Call ClearEntries
    AddNew = False
    Qry = "SELECT cashreturn.* From cashreturn WHERE (CRID=" & Val(txtCRid.Text) & ");"
    With rsCashReturn
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            txtCRno.Text = !CashRetNO
            txtCRDate.Text = Format(!crDate, "mm/dd/yyyy")
            txtBranch.Text = !Branch
            txtRemarks.Text = !Remarks
            cboCashier.Text = !PreparedBy
            txtAprvBy.Text = !APPROVEDBY
            txtCRid.Text = !CRID
            txtPrnCopy.Text = !PRINTCOPY
            cboType.ListIndex = !CRType
        End If
        .Close
    End With
    If cboType.ListIndex = 0 Then
        Qry = "SELECT cashretdtls.* From cashretdtls Where cashretdtls.CRID = " & Val(txtCRid.Text) & ";"
        lstVoucher.Clear
        With rsCRdtls
            .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
            If .RecordCount > 0 Then
                .MoveFirst
                Do While Not .EOF
                    lstVoucher.AddItem !VOUCHERID
                    lstVoucher.ItemData(lstVoucher.ListCount - 1) = !ExpAmt
                    .MoveNext
                Loop
                .Close
            End If
        End With
        Qry = "SELECT disbursement.CVNUMBER, disbursement.CVDATE, disbursement.ACCTDESC, disbursement.AMOUNT, cashretdtls.EXPAMT, cashretdtls.VOUCHERID " & _
              "From cashretdtls INNER JOIN disbursement ON (cashretdtls.VOUCHERID = disbursement.VOUCHERID) Where cashretdtls.CRID = " & Val(txtCRid.Text) & ";"
        With rsCV
            .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
            If .RecordCount > 0 Then
                .MoveFirst
                flxCR.Rows = 1
                Ttl = 0
                Do While Not .EOF
                    flxCR.AddItem vbTab & Space$(10) & "CV No. " & !CVNUMBER & " / " & "Date:" & Format(!CVDate, "mm/dd/yyyy") & vbTab & Format(!Amount, "#,##0.00") & Space$(8) & vbTab & !VOUCHERID
                    flxCR.AddItem vbTab & Space$(10) & !AcctDesc & vbTab & vbTab & !VOUCHERID
                    Expns = !Amount - !ExpAmt
                    flxCR.AddItem vbTab & Space$(20) & "Actual Expense" & vbTab & Format(Expns, "#,##0.00") & Space$(8) & vbTab & !VOUCHERID
                    Ttl = Ttl + !ExpAmt
                    .MoveNext
                Loop
            End If
            .Close
        End With
        txtTotal.Text = Format(Ttl, "#,##0.00") & Space(8)
        If lstVoucher.ListCount > 2 And Val(txtPrnCopy.Text) > 0 Then
            cmdPrnAtt.Visible = True
        Else
            cmdPrnAtt.Visible = False
        End If
    Else
        Qry = "SELECT * From cashretdtls Where cashretdtls.CRID = " & Val(txtCRid.Text) & " Order By cashretdtls.VOUCHERID;"
        With rsCV
            .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
            If .RecordCount > 0 Then
                flxCR.Rows = 1
                Ttl = 0
                Do While Not .EOF
                    flxCR.AddItem vbTab & !Particulars & vbTab & Format(!ExpAmt, "#,##0.00")
                    Ttl = Ttl + !ExpAmt
                    .MoveNext
                Loop
                txtTotal.Text = Format(Ttl, "#,##0.00")
            End If
            .Close
        End With
    End If
End Sub

Private Sub cboCashier_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtAprvBy.SetFocus
End Sub

Private Sub cboType_Click()
    flxCR.Rows = 1
    flxCR.Rows = 2
    If cboType.ListIndex = 0 Then
        cmdAddCV.Enabled = True
    ElseIf cboType.ListIndex = 1 Then
        cmdAddCV.Enabled = False
    End If
End Sub

Private Sub cmdAddCV_Click()
    With tabCV
        .Top = (Me.Height - .Height) \ 2
        .Left = (Me.Width - .Width) \ 2
        .Visible = True
    End With
    Call LockData(True)
    cmdTrans(5).Enabled = False
    cmdTrans(6).Enabled = False
    txtCVno.Text = ""
    txtCVno.Tag = ""
    txtExpAmt.Text = ""
    txtCVamt.Text = "C.V. Amount: "
    txtCVamt.Tag = ""
    txtCVdate.Text = "C.V. Date: "
    txtAcctDesc.Text = ""
    txtCVno.Tag = ""
    txtCVno.SetFocus
End Sub

Private Sub cmdCancel_Click()
    With tabCV
        .Visible = False
        .Top = 10000
    End With
    Call LockData(False)
    cmdTrans(5).Enabled = True
    cmdTrans(6).Enabled = True
    cmdAddCV.SetFocus
End Sub

Private Sub cmdOK_Click()
    Dim vMsg As String, myTtl As Currency, cRetAmt As Currency
    
    vMsg = ""
    'Validate if CV is correct
    If txtCVno.Tag = "" Then
        vMsg = "Please indicate a valid C.V. number!"
    ElseIf txtExpAmt.Text = "" Then
        vMsg = vMsg & vbCrLf & "No Actual expense specified!"
    ElseIf Val(txtExpAmt.Text) > Val(txtCVamt.Tag) Then
        vMsg = vMsg & vbCrLf & "Actual Expense must not be greater than " & txtCVamt.Text
    End If
    If vMsg <> "" Then
        MsgBox vMsg, vbExclamation + vbOKOnly, "Warning"
        Exit Sub
    End If
    If flxCR.Rows = 2 Then flxCR.Rows = 1
    flxCR.AddItem vbTab & Space$(10) & "PCV No. " & txtCVno.Text & " / " & txtCVdate.Text & vbTab & Format(Val(txtCVamt.Tag), "#,##0.00") & Space$(8) & vbTab & txtCVno.Tag
    flxCR.AddItem vbTab & Space$(10) & txtAcctDesc.Text & vbTab & vbTab & txtCVno.Tag
    flxCR.AddItem vbTab & Space$(20) & "Actual Expense" & vbTab & Format(Val(txtExpAmt.Text), "#,##0.00") & Space$(8) & vbTab & txtCVno.Tag
    myTtl = Val(Replace(txtTotal.Text, ",", ""))
    cRetAmt = Val(txtCVamt.Tag) - Val(txtExpAmt.Text)
    myTtl = myTtl + cRetAmt
    txtTotal.Text = Format(myTtl, "#,##0.00") & Space(8)
    lstVoucher.AddItem txtCVno.Tag
    lstVoucher.ItemData(lstVoucher.ListCount - 1) = cRetAmt
    Call cmdCancel_Click
End Sub

Private Sub cmdTrans_Click(Index As Integer)
    Dim Ans As Integer, i As Integer, RetID As Long, MyCRno As Long
    Dim myMsg As String, Qry As String, MyAcc As String, PrnVal(0 To 3) As String
    
    Select Case Index
        Case 0      'addnew
            Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE (EMPID=" & UserID & " AND MENUINDX=" & 99 & ");"
            MyAcc = GetFieldValue(Qry)
            If MyAcc <> "FULL ACCESS" Then
                MsgBox "You don't have full access to this transaction." & vbCrLf & "Please contact your System Administrator", vbExclamation + vbOKOnly, "Access Denied"
                Exit Sub
            End If
            BtnIndx = 4
            tmrFloat.Enabled = True
            Call ClearEntries
            Call LockData(False)
            cmdPrnAtt.Visible = False
            Qry = "SELECT branchname From branch WHERE   SYSTEMbranch = 1;"
            txtBranch.Text = GetFieldValue(Qry)
            Qry = "SELECT max(round(CashRetNo)) As MyCR From cashreturn;"
            MyCRno = Val(GetFieldValue(Qry))
            txtCRno.Text = IIf(Len(str(MyCRno)) > 3, MyCRno, Format(MyCRno + 1, "0000"))
            'Qry = "SELECT TRANSIGNATORY.EMPLOYEE From TRANSIGNATORY WHERE (TRANSACTION='CASH RETURN (1)');"
            'cbocashier.Text = GetFieldValue(Qry)
            'Qry = "SELECT TRANSIGNATORY.EMPLOYEE From TRANSIGNATORY WHERE (TRANSACTION='CASH RETURN (2)');"
            'txtAprvBy.Text = GetFieldValue(Qry)
            txtCRDate.Text = Format(Date, "mm/dd/yyyy")
            cmdAddCV.SetFocus
            AddNew = True
            JustScroll = False
            cboType.ListIndex = 0
        Case 1      'edit
            Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE (EMPID=" & UserID & " AND MENUINDX=" & 99 & ");"
            MyAcc = GetFieldValue(Qry)
            If MyAcc <> "FULL ACCESS" Then
                MsgBox "You don't have full access to this transaction." & vbCrLf & "Please contact your System Administrator", vbExclamation + vbOKOnly, "Access Denied"
                Exit Sub
            End If
            If Val(txtPrnCopy.Text) > 0 Then
                MsgBox "Editing is not allowed because the transaction was already printed", vbExclamation + vbOKOnly, "Print Receiving Receipt"
                Exit Sub
            End If
            
            BtnIndx = 4
            tmrFloat.Enabled = True
            Call LockData(False)
            cmdTrans(1).Tag = txtCRno.Text
            AddNew = False
            JustScroll = False
            cmdPrnAtt.Visible = False
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
            RetID = Val(txtCRid.Text)
            Call ClearEntries
            txtCRid.Text = RetID
            If Not AddNew Then Call ShowRecord
            Call LockData(True)
        
        Case 2      'View Sumary
            Unload Me
            FrmSummary.mySrc = "Cash Return"
            FrmSummary.Show
            'MotoRep.tbMain.Visible = False
            MotoRep.imgBG.Visible = False
            MotoRep.toolBarButton.Visible = False
        Case 3      '"print"
            If Val(txtCRid.Text) = 0 Then
                MsgBox "No Cash Return to print!", vbExclamation + vbOKOnly, "Report Empty"
                Exit Sub
            End If
            Ans = MsgBox("Do you want to Print Cash Return Slip?", vbQuestion + vbYesNo, "Print")
            If Ans = vbNo Then Exit Sub
            flxPrn.Rows = 1
            If lstVoucher.ListCount <= 2 Then
                With flxCR
                    For i = 1 To .Rows - 1
                        flxPrn.AddItem vbTab & Trim(.TextMatrix(i, 1)) & vbTab & RTrim(.TextMatrix(i, 2))
                    Next i
                End With
            Else
                flxPrn.Rows = 3
                flxPrn.AddItem vbTab & "TOTAL CASH RETURN" & vbTab & txtTotal.Text
                flxPrn.AddItem vbTab & "Please see attachment (" & lstVoucher.ListCount & " CVs)"
            End If
            PrnVal(0) = txtBranch.Text
            PrnVal(1) = txtCRDate.Text
            PrnVal(2) = txtTotal.Text
            PrnVal(3) = Format(Now, "mm/dd/yy hh:ss ampm")
            Call PrintMyForm(PrnVal, "Cash Return Slip", flxPrn)
            If Val(txtPrnCopy.Text) = 0 Then
                LYRcon.Execute "UPDATE cashreturn SET cashreturn.PRINTCOPY = 1 WHERE (CRID=" & txtCRid.Text & ");"
            Else
                txtPrnCopy.Text = Val(txtPrnCopy.Text) + 1
                LYRcon.Execute "UPDATE cashreturn SET cashreturn.PRINTCOPY = " & Val(txtPrnCopy.Text) & " WHERE (CRID=" & txtCRid.Text & ");"
            End If
            If lstVoucher.ListCount > 2 Then    'Print Attachment
                cmdPrnAtt.Visible = True
                Ans = MsgBox("Do you want to Preview Attachment?", vbQuestion + vbYesNo, "Attachment")
                If Ans = vbNo Then Exit Sub
                
                Qry = "SELECT disbursement.CVNUMBER, disbursement.CVDATE, disbursement.ACCTDESC, disbursement.PAYEE, disbursement.AMOUNT, cashretdtls.EXPAMT " & _
                      "From cashretdtls INNER JOIN disbursement ON (cashretdtls.VOUCHERID = disbursement.VOUCHERID) " & _
                      "Where cashretdtls.CRID = " & Val(txtCRid.Text) & " Order By disbursement.CVNUMBER;"
                FrmPrn.QryPrn = Qry
                FrmPrn.ParamValue1 = txtBranch.Text
                FrmPrn.ParamValue2 = "Ret. Date: " & txtCRDate.Text
                FrmPrn.ParamValue3 = "Cash Ret. No: " & txtCRno.Text
                FrmPrn.MySource = "CASH RETURN ATTACHMENT"
                Unload Me
                FrmPrn.Show
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

Private Sub flxCR_Click()
    ActRow = flxCR.Row
    ActCol = flxCR.Col
    
    If ActCol = 0 Or ActRow = 0 Or JustScroll Or cboType.ListIndex = 0 Then Exit Sub
    Select Case ActCol
        Case 1       'Particulars
            With txtParticular
                .Tag = flxCR.TextMatrix(ActRow, ActCol)
                .Top = flxCR.CellTop + flxCR.Top
                .Left = flxCR.CellLeft + flxCR.Left
                .Width = flxCR.CellWidth
                .Text = flxCR.Text
                .Visible = True
                .SelStart = 0
                .SelLength = Len(.Text)
                .ZOrder
                .SetFocus
              End With
        Case 2      'Amount
            With txtAmount
                .Tag = flxCR.TextMatrix(ActRow, 2)
                .Top = flxCR.CellTop + flxCR.Top
                .Left = flxCR.CellLeft + flxCR.Left
                .Width = flxCR.CellWidth
                .Text = flxCR.Text
                .Visible = True
                .SelStart = 0
                .SelLength = Len(.Text)
                .ZOrder
                .SetFocus
            End With
    End Select
End Sub

Private Sub flxCR_EnterCell()
    ActRow = flxCR.Row
End Sub

Private Sub flxCR_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call flxCR_Click
End Sub

Private Sub flxCR_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If JustScroll Then Exit Sub
    If Button = vbRightButton And ActRow > 0 And ActCol = 0 Then
        If cboType.ListIndex = 0 Then
            mnuAdd.Enabled = False
        Else
            mnuAdd.Enabled = True
        End If
        Call PopupMenu(mnuFlx)
    End If
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
            cboCashier.Clear
            Do While Not .EOF
                cboCashier.AddItem !Cashiers
                cboCashier.ItemData(cboCashier.ListCount - 1) = !EMPID
                If UserID = !EMPID Then MyID = cboCashier.ListCount - 1
                .MoveNext
            Loop
        End If
        .Close
    End With
    cboType.AddItem "PREVIOUS VOUCHER"
    cboType.AddItem "OTHERS"
    If cboCashier.ListCount > 0 Then cboCashier.ListIndex = MyID
    Call LockData(True)
    
    MotoRep.sBar.Panels(6).Text = "Cash Return Form"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    MotoRep.sBar.Panels(6).Text = ""
End Sub


Private Sub mnuAdd_Click()
    
    ActRow = flxCR.Rows - 1
    If flxCR.TextMatrix(ActRow, 1) = "" Then
        MsgBox "Incomplete entry on last row!", vbExclamation + vbOKOnly, "Validation Error"
        flxCR.SetFocus
        flxCR.Col = 0
    Else
        flxCR.AddItem ""
        flxCR.SetFocus
        flxCR.Row = flxCR.Rows - 1
        flxCR.Col = 0
    End If
End Sub

Private Sub mnuDelete_Click()
    Dim Ans As Integer, GrpRow As Long, myTtl As Currency
    
    Ans = MsgBox("Do you want to delete voucher details?", vbQuestion + vbYesNo, "Delete")
    If Ans = vbNo Then Exit Sub
    If cboType.ListIndex = 1 Then
        If flxCR.Rows = 2 Then
            flxCR.Rows = 1
            flxCR.Rows = 2
        Else
            flxCR.RemoveItem ActRow
        End If
    ElseIf cboType.ListIndex = 0 Then
        GrpRow = flxCR.TextMatrix(ActRow, 2)
        For Ans = flxCR.Rows - 1 To 1 Step -1
            If Val(flxCR.TextMatrix(Ans, 2)) = GrpRow Then
                If Ans = 1 And flxCR.Rows = 2 Then
                    flxCR.Rows = 1
                    flxCR.Rows = 2
                Else
                    flxCR.RemoveItem Ans
                End If
            End If
        Next Ans
        myTtl = Val(Replace(txtTotal.Text, ",", ""))
        For Ans = 0 To lstVoucher.ListCount - 1
            If GrpRow = Val(lstVoucher.List(Ans)) Then
                myTtl = myTtl - lstVoucher.ItemData(Ans)
                txtTotal.Text = IIf(myTtl > 0, Format(myTtl, "#,##0.00") & Space(8), "")
                lstVoucher.RemoveItem Ans
                Exit For
            End If
        Next Ans
    End If
End Sub

Private Sub RXResizer1_ResizeComplete()
    With flxCR
        .ColWidth(0) = 300
        .ColWidth(1) = .Width * 0.7
        .ColWidth(2) = .Width * 0.25
        .ColWidth(3) = 0
        .Row = 0: .Col = 1
        .Text = "P A R T I C U L A R S"
        .CellAlignment = 4
        .CellFontBold = True
        .Row = 0: .Col = 2
        .Text = "A M O U N T"
        .CellAlignment = 4
        .CellFontBold = True
        txtTotal.Left = .Left + .ColWidth(0) + .ColWidth(1)
        txtTotal.Width = .ColWidth(2)
    End With
    lblTotal.Left = txtTotal.Left - lblTotal.Width
    lblHeading(2).Left = Frame1.Left + Frame1.Width - lblHeading(2).Width
    lblHeading(3).Left = lblHeading(2).Left + 30
End Sub

Private Sub txtaprvby_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtChkBy.SetFocus
End Sub

Private Sub txtAprvBy_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        If JustScroll Then
            cmdTrans(4).SetFocus
        Else
            cmdTrans(5).SetFocus
        End If
    End If
End Sub

Private Sub txtcrdate_KeyPress(KeyAscii As Integer)
    Dim strValid As String
    
    strValid = "1234567890/-"
    If KeyAscii > 26 Then
        If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
    If KeyAscii = 13 Then cmdAddCV.SetFocus

End Sub

Private Sub txtcrdate_LostFocus()
    Dim Qry As String
    
    If IsDate(txtCRDate.Text) Then
        Qry = ""    'Validate if Cashflow for the date if already made.
        'If IsDate(cmdTrans(0).Tag) Then
        '    If DateValue(txtCRDate.Text) <> DateValue(cmdTrans(0).Tag) Then Call DisplayAcctDistribution
        'Else
        '    Call DisplayAcctDistribution
        'End If
        'cmdTrans(0).Tag = txtCRDate.Text
    ElseIf txtCRDate.Text <> "" Then
        MsgBox "Invalid Date!", vbExclamation + vbOKOnly, "Date"
    End If
End Sub


Function ValidateEntry() As String
    Dim vMsg As String, vQry As String, vModel As String
    Dim y As Integer, i As Integer, CVamt As Currency
    
    vMsg = ""
    If txtBranch.Text = "" Then vMsg = "Incomplete entry on field BRANCH"
    If txtCRno.Text = "" Then
        vMsg = vMsg & vbCrLf & "Incomplete entry on field CASH RETURN NO."
    ElseIf AddNew Or (AddNew = False And cmdTrans(1).Tag <> txtCRno.Text) Then
        vQry = "SELECT cashreturn.CASHRETNO From cashreturn WHERE (cashreturn.CASHRETNO ='" & txtCRno.Text & "');"
        If CtrlNumExist(vQry) Then vMsg = vMsg & vbCrLf & "Cash Return No. " & txtCRno.Text & " already exist"
    End If
    If Not IsDate(txtCRDate.Text) Then vMsg = vMsg & vbCrLf & "Invalid entry on field C.R. DATE"
    If cboCashier.ListIndex = -1 Then vMsg = vMsg & vbCrLf & "Incomplete entry on field RECEIVED BY"
    If txtAprvBy.Text = "" Then vMsg = vMsg & vbCrLf & "Incomplete entry on field APPROVED BY"
    If cboType.ListIndex = 0 Then
        If lstVoucher.ListCount = 0 Then vMsg = vMsg & vbCrLf & "No Cash Voucher specified in Particulars"
    ElseIf cboType.ListIndex = 1 Then
        If flxCR.TextMatrix(1, 1) = "" Then vMsg = vMsg & vbCrLf & "Incomplete entry on Cash Return Particulars"
        If flxCR.TextMatrix(1, 2) = "" Then vMsg = vMsg & vbCrLf & "Incomplete entry on Cash Return Amount"
    Else
        vMsg = vMsg & vbCrLf & "Invalid Cash Return Type"
    End If
        
    ValidateEntry = vMsg
End Function

Private Sub SaveEntries()
    Dim rsCashRet As New ADODB.Recordset
    Dim Qry As String, ErrLevel As String, Tmp As String
    Dim i As Integer
    
    On Error GoTo LocalErr
    If AddNew Then
        ErrLevel = "Saving new entries on table Cash Return"
        LYRcon.Execute "Insert Into cashreturn(CASHRETNO, CRDATE, BRANCH, CRAMOUNT, REMARKS, PREPAREDBY, APPROVEDBY, CRTYPE, UPDATEDBY, LASTUPDATE) " & _
                       "VALUES('" & txtCRno.Text & "', '" & Format(txtCRDate.Text, "yyyy-mm-dd") & "', '" & Left$(txtBranch.Text, 50) & "', " & Val(Replace(txtTotal.Text, ",", "")) & ", '" & txtRemarks.Text & "', '" & cboCashier.Text & "', '" & txtAprvBy.Text & "', " & cboType.ListIndex & ", " & cboCashier.ItemData(cboCashier.ListIndex) & ", '" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "');"
        Qry = "SELECT CashReturn.CRID From CashReturn Where cashreturn.CASHRETNO = '" & txtCRno.Text & "' AND cashreturn.CRDATE = '" & Format(txtCRDate.Text, "yyyy-mm-dd") & "';"
        txtCRid.Text = Val(GetFieldValue(Qry))
    Else
        ErrLevel = "Saving edited entries on table Check Voucher"
        LYRcon.Execute "UPDATE CashReturn SET CASHRETNO = '" & txtCRno.Text & "', CRDATE = '" & Format(txtCRDate.Text, "yyyy-mm-dd") & "', BRANCH = '" & Left$(txtBranch.Text, 50) & "', CRAMOUNT = " & Val(Replace(txtTotal.Text, ",", "")) & ", REMARKS = '" & txtRemarks.Text & "', " & _
                       "PREPAREDBY = '" & cboCashier.Text & "', APPROVEDBY = '" & txtAprvBy.Text & "', UPDATEDBY = " & cboCashier.ItemData(cboCashier.ListIndex) & ", CRTYPE = " & cboType.ListIndex & ", LASTUPDATE = '" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "' " & _
                        "Where CashReturn.CRID = " & txtCRid.Text & ";"
    End If
    
    If Not AddNew Then
        ErrLevel = "Clearing Cash Return Details from edited record"
        LYRcon.Execute "DELETE cashretdtls.* From cashretdtls Where cashretdtls.CRID = " & Val(txtCRid.Text) & ";"
    End If
    If cboType.ListIndex = 0 Then
        With lstVoucher
            ErrLevel = "Inserting Cash Return Details records"
            For i = 0 To .ListCount - 1
                LYRcon.Execute "INSERT INTO cashretdtls (CRID, VOUCHERID, EXPAMT) VALUES(" & Val(txtCRid.Text) & ", " & Val(.List(i)) & ", " & .ItemData(i) & ");"
            Next i
        End With
    ElseIf cboType.ListIndex = 1 Then
        With flxCR
            ErrLevel = "Inserting Cash Return Details records"
            For i = 1 To .Rows - 1
                LYRcon.Execute "INSERT INTO cashretdtls (CRID, VOUCHERID, EXPAMT, PARTICULARS) VALUES(" & Val(txtCRid.Text) & ", 0, " & Val(Replace(.TextMatrix(i, 2), ",", "")) & ", '" & .TextMatrix(i, 1) & "');"
            Next i
        End With
    End If
    Exit Sub
LocalErr:
    MsgBox ErrLevel & vbCrLf & Err.Description, vbCritical + vbOKOnly, "Error"
        
End Sub

Private Sub FinalizeVoucher()
    Dim i As Integer
    
    'rsCV.Open "CVDetails", LYRcon, adOpenKeyset, adLockOptimistic
    With flxCV
        For i = 1 To .Rows - 1
            LYRcon.Execute "Insert Into cvdetails (CVID, ACCTDIST, DEBIT, CREDIT) " & _
                           "VALUES (" & Val(txtCVid.Text) & ", '" & Trim(.TextMatrix(i, 0)) & "', " & Val(Replace(.TextMatrix(i, 1), ",", "")) & ", " & Val(Replace(.TextMatrix(i, 2), ",", "")) & ");"
            If InStr(.TextMatrix(i, 0), "Cash On Bank") > 0 Then Exit For
        Next i
    End With
    LYRcon.Execute "UPDATE checkvoucher SET AMOUNT = " & Val(Replace(txtTotal.Text, ",", "")) & ", PRINTCOPY = 1 " & _
                   "Where (checkvoucher.CVID = " & Val(txtCVid.Text) & ");"
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

Private Sub txtCVno_LostFocus()
    Dim rsCV As New ADODB.Recordset
    Dim Qry As String
    
    If txtCVno.Text = "" Then Exit Sub
    Qry = "SELECT disbursement.VOUCHERID, disbursement.CVDATE, disbursement.ACCTDESC, disbursement.AMOUNT, disbursement.CVNUMBER " & _
          "From disbursement Where disbursement.CVNUMBER = '" & txtCVno.Text & "';"
    With rsCV
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            txtCVamt.Text = "C.V. Amount: " & Format(!Amount, "#,##0.00")
            txtCVamt.Tag = !Amount
            txtCVdate.Text = "C.V. Date: " & Format(!CVDate, "mm/dd/yyyy")
            txtAcctDesc.Text = !AcctDesc
            txtCVno.Tag = !VOUCHERID
        Else
            txtCVamt.Text = "C.V. Amount: "
            txtCVdate.Text = "C.V. Date: "
            txtAcctDesc.Text = ""
            txtCVno.Tag = ""
        End If
        .Close
    End With
End Sub

Private Sub txtExpAmt_KeyPress(KeyAscii As Integer)
    Dim strValid As String
    
    strValid = "1234567890."
    If KeyAscii > 26 Then
        If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
    If KeyAscii = 13 Then cmdOK.SetFocus

End Sub

Private Sub txtExpAmt_LostFocus()
    If Val(txtExpAmt.Text) > 0 And txtCVno.Tag <> "" Then
        If Val(txtExpAmt.Text) > Val(txtCVamt.Tag) Then
            MsgBox "Actual Expense must not be greater than " & txtCVamt.Text, vbExclamation + vbOKOnly, "Invalid Amount"
            txtExpAmt.Text = ""
            txtExpAmt.SetFocus
        End If
    End If
End Sub


Private Sub txtParticular_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyEscape Then
        txtParticular.Text = Empty
        If txtParticular.Tag <> "" Then flxCR.TextMatrix(ActRow, 1) = txtParticular.Tag
        txtParticular.Tag = "skiplostfocus"
        flxCR.SetFocus
        flxCR.Col = 1
        txtParticular.Visible = False
    End If
End Sub

Private Sub txtParticular_KeyPress(KeyAscii As Integer)
    
    If Chr(KeyAscii) = "'" Then KeyAscii = 0
    If KeyAscii = 13 Then
        flxCR.SetFocus
        flxCR.Col = 2
        flxCR.Row = ActRow
    End If
End Sub

Private Sub txtParticular_LostFocus()
    Dim i As Integer
    
    If txtParticular.Tag = "skiplostfocus" Then
        txtParticular.Tag = ""
        Exit Sub
    End If
    flxCR.TextMatrix(ActRow, 1) = txtParticular.Text
    txtParticular.Visible = False
End Sub

Private Sub txtAmount_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        txtAmount.Text = Empty
        flxCR.TextMatrix(ActRow, 2) = txtAmount.Tag
        txtAmount.Tag = "skiplostfocus"
        flxCR.SetFocus
        flxCR.Col = 2
        txtAmount.Visible = False
    End If
End Sub

Private Sub txtAmount_KeyPress(KeyAscii As Integer)
    Dim strValid As String
    
    strValid = "1234567890.,"
    If KeyAscii > 26 Then
        If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
    If KeyAscii = 13 Then
        flxCR.SetFocus
        flxCR.Col = 2
    End If
End Sub

Private Sub txtAmount_LostFocus()
    Dim i As Integer, TtlAmt As Currency
    
    If txtAmount.Tag = "skiplostfocus" Then
        txtAmount.Tag = ""
        Exit Sub
    End If
    flxCR.TextMatrix(ActRow, 2) = txtAmount.Text
    txtAmount.Visible = False
    TtlAmt = 0
    For i = 1 To flxCR.Rows - 1
        TtlAmt = TtlAmt + Val(Replace(flxCR.TextMatrix(i, 2), ",", ""))
    Next i
    txtTotal.Text = Format(TtlAmt, "#,##0.00")
End Sub
