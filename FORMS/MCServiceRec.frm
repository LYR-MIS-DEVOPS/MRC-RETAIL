VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form SRVCrecord 
   Caption         =   "Customer Service Record"
   ClientHeight    =   9240
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10980
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9240
   ScaleWidth      =   10980
   WindowState     =   2  'Maximized
   Begin lvButton.lvButtons_H cmdMCStat 
      Height          =   375
      Left            =   9000
      TabIndex        =   59
      Tag             =   "XYF"
      Top             =   2880
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Caption         =   "Additional MC"
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
      cBack           =   -2147483633
   End
   Begin VB.TextBox txtCustID 
      Height          =   285
      Left            =   480
      TabIndex        =   58
      Top             =   10080
      Visible         =   0   'False
      Width           =   855
   End
   Begin TabDlg.SSTab tabCust 
      Height          =   2895
      Left            =   1920
      TabIndex        =   39
      Top             =   9480
      Visible         =   0   'False
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   5106
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "MCServiceRec.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Shape1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label5"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label4(1)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label6"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1(1)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label4(2)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label4(3)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdCancel"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmdOK"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtNewAdrs"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cboCompany"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtDesignation"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "flxNCust"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtNewCust"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtNewID"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtNewAcct"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtDesig"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "cboCType"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).ControlCount=   18
      Begin VB.ComboBox cboCType 
         Height          =   315
         Left            =   1920
         TabIndex        =   56
         Top             =   1140
         Width           =   4215
      End
      Begin VB.TextBox txtDesig 
         Height          =   285
         Left            =   2160
         TabIndex        =   55
         Top             =   3360
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtNewAcct 
         Height          =   285
         Left            =   1200
         TabIndex        =   54
         Top             =   3360
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtNewID 
         Height          =   285
         Left            =   360
         TabIndex        =   53
         Top             =   3360
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtNewCust 
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
         TabIndex        =   52
         Top             =   480
         Width           =   4215
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxNCust 
         Height          =   1215
         Left            =   2640
         TabIndex        =   51
         Top             =   3840
         Visible         =   0   'False
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   2143
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
      Begin VB.TextBox txtDesignation 
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
         MaxLength       =   30
         TabIndex        =   47
         Top             =   1800
         Width           =   4215
      End
      Begin VB.ComboBox cboCompany 
         Height          =   315
         Left            =   1920
         TabIndex        =   46
         Top             =   1470
         Width           =   4215
      End
      Begin VB.TextBox txtNewAdrs 
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
         MaxLength       =   100
         TabIndex        =   41
         Top             =   810
         Width           =   4215
      End
      Begin lvButton.lvButtons_H cmdOK 
         Height          =   375
         Left            =   1950
         TabIndex        =   49
         Top             =   2280
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
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
         Image           =   "MCServiceRec.frx":001C
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdCancel 
         Height          =   375
         Left            =   3390
         TabIndex        =   50
         Top             =   2280
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
         Image           =   "MCServiceRec.frx":05B6
         cBack           =   -2147483633
      End
      Begin VB.Label Label4 
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
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   3
         Left            =   360
         TabIndex        =   48
         Tag             =   "XYF"
         Top             =   1830
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "COMPANY NAME:"
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
         TabIndex        =   45
         Tag             =   "XYF"
         Top             =   1485
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "NEW CUSTOMER"
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
         TabIndex        =   44
         Tag             =   "XYF"
         Top             =   510
         Width           =   1815
      End
      Begin VB.Label Label6 
         Caption         =   "NEW ADDRESS"
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
         Left            =   360
         TabIndex        =   43
         Tag             =   "XYF"
         Top             =   870
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "CUSTOMER TYPE"
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
         TabIndex        =   42
         Tag             =   "XYF"
         Top             =   1178
         Width           =   1455
      End
      Begin VB.Label Label5 
         BackColor       =   &H80000002&
         Caption         =   "   CHANGE MOTORCYCLE OWNER"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   0
         TabIndex        =   40
         Top             =   60
         Width           =   6255
      End
      Begin VB.Shape Shape1 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H80000002&
         FillStyle       =   0  'Solid
         Height          =   315
         Left            =   0
         Top             =   0
         Width           =   6495
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxCust 
      Height          =   1335
      Left            =   8640
      TabIndex        =   35
      Top             =   9480
      Visible         =   0   'False
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   2355
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
   Begin MSFlexGridLib.MSFlexGrid flxSrvc 
      Height          =   5775
      Left            =   120
      TabIndex        =   11
      Tag             =   "XYF"
      Top             =   3240
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   10186
      _Version        =   393216
      Cols            =   13
      RowHeightMin    =   300
      BackColorBkg    =   16777215
      SelectionMode   =   1
      AllowUserResizing=   1
      FormatString    =   $"MCServiceRec.frx":0B50
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   2175
      Left            =   120
      TabIndex        =   12
      Tag             =   "XYF"
      Top             =   720
      Width           =   10815
      Begin VB.TextBox txtCType 
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
         Left            =   1680
         TabIndex        =   57
         Tag             =   "XYF"
         Top             =   180
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
         Left            =   8520
         TabIndex        =   6
         Tag             =   "XYF"
         Top             =   495
         Width           =   2175
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
         Left            =   8520
         TabIndex        =   10
         Tag             =   "XYF"
         Top             =   1755
         Width           =   2175
      End
      Begin VB.TextBox txtCustomer 
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
         Left            =   1680
         TabIndex        =   0
         Tag             =   "XYF"
         Top             =   495
         Width           =   4455
      End
      Begin VB.TextBox txtDealer 
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
         Left            =   1680
         TabIndex        =   4
         Tag             =   "XYF"
         Top             =   1755
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
         Left            =   8520
         TabIndex        =   9
         Tag             =   "XYF"
         Top             =   1440
         Width           =   2175
      End
      Begin VB.TextBox txtAcct 
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
         Locked          =   -1  'True
         TabIndex        =   5
         Tag             =   "XYF"
         Top             =   180
         Width           =   2175
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
         Left            =   8520
         TabIndex        =   8
         Tag             =   "XYF"
         Top             =   1125
         Width           =   2175
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
         Left            =   8520
         TabIndex        =   7
         Tag             =   "XYF"
         Top             =   810
         Width           =   2175
      End
      Begin VB.TextBox txtAddress 
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
         Left            =   1680
         TabIndex        =   1
         Tag             =   "XYF"
         Top             =   810
         Width           =   4455
      End
      Begin VB.TextBox txtDatePurchased 
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
         Left            =   1680
         TabIndex        =   3
         Tag             =   "XYF"
         Top             =   1440
         Width           =   4455
      End
      Begin VB.TextBox txtContactNo 
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
         Left            =   1680
         TabIndex        =   2
         Tag             =   "XYF"
         Top             =   1125
         Width           =   4455
      End
      Begin VB.Label Label4 
         Caption         =   "CUSTOMER TYPE"
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
         TabIndex        =   38
         Tag             =   "XYF"
         Top             =   210
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
         Left            =   7200
         TabIndex        =   36
         Tag             =   "XYF"
         Top             =   525
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
         Left            =   7200
         TabIndex        =   34
         Tag             =   "XYF"
         Top             =   1785
         Width           =   1815
      End
      Begin VB.Label lblDealer 
         Caption         =   "DEALER"
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
         TabIndex        =   33
         Tag             =   "XYF"
         Top             =   1785
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
         Left            =   7200
         TabIndex        =   20
         Tag             =   "XYF"
         Top             =   1470
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "ACCOUNT NO."
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
         Left            =   7200
         TabIndex        =   19
         Tag             =   "XYF"
         Top             =   210
         Width           =   1815
      End
      Begin VB.Label lblPurDes 
         Caption         =   "DATE PURCHASED"
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
         TabIndex        =   18
         Tag             =   "XYF"
         Top             =   1470
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "CONTACT NOS."
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
         Left            =   120
         TabIndex        =   17
         Tag             =   "XYF"
         Top             =   1155
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
         Left            =   7200
         TabIndex        =   16
         Tag             =   "XYF"
         Top             =   1155
         Width           =   1815
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
         Left            =   7200
         TabIndex        =   15
         Tag             =   "XYF"
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label lblAdrsCom 
         Caption         =   "ADDRESS"
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
         TabIndex        =   14
         Tag             =   "XYF"
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label Label1 
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
         Left            =   120
         TabIndex        =   13
         Tag             =   "XYF"
         Top             =   525
         Width           =   1815
      End
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   3
      Left            =   3360
      TabIndex        =   27
      Top             =   0
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
      Caption         =   "&View Job Order"
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
      Image           =   "MCServiceRec.frx":0C1B
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   4
      Left            =   5040
      TabIndex        =   28
      Top             =   0
      Width           =   1680
      _ExtentX        =   2963
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
      Image           =   "MCServiceRec.frx":11B5
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   5
      Left            =   6720
      TabIndex        =   29
      Top             =   0
      Width           =   1680
      _ExtentX        =   2963
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
      Image           =   "MCServiceRec.frx":174F
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   6
      Left            =   8400
      TabIndex        =   30
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
      Image           =   "MCServiceRec.frx":1CE9
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   7
      Left            =   10080
      TabIndex        =   31
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
      Image           =   "MCServiceRec.frx":2283
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   1
      Left            =   0
      TabIndex        =   32
      Top             =   0
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
      Caption         =   "&Change Owner"
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
      Image           =   "MCServiceRec.frx":281D
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   2
      Left            =   1680
      TabIndex        =   37
      Top             =   0
      Width           =   1680
      _ExtentX        =   2963
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
      Image           =   "MCServiceRec.frx":2DB7
      cBack           =   -2147483633
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   2160
      Top             =   9000
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   11220
      m_MDI_Height    =   9810
      OriginalBorder  =   10
      OriginalCaption =   31
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "CUSTOMER SERVICE RECORD"
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
      Index           =   5
      Left            =   7920
      TabIndex        =   26
      Top             =   480
      Width           =   3015
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "CUSTOMER SERVICE RECORD"
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
      Index           =   4
      Left            =   7950
      TabIndex        =   25
      Top             =   510
      Width           =   3015
   End
   Begin VB.Label Label3 
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
      Height          =   300
      Index           =   3
      Left            =   120
      TabIndex        =   24
      Top             =   480
      Width           =   2900
   End
   Begin VB.Label Label3 
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
      Height          =   300
      Index           =   2
      Left            =   150
      TabIndex        =   23
      Top             =   510
      Width           =   2900
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "SERVICE HISTORY"
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
      Index           =   1
      Left            =   120
      TabIndex        =   22
      Top             =   2985
      Width           =   2055
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "SERVICE HISTORY"
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
      Index           =   0
      Left            =   135
      TabIndex        =   21
      Top             =   3000
      Width           =   2055
   End
   Begin VB.Menu mnuSR 
      Caption         =   "Service Record"
      Visible         =   0   'False
   End
End
Attribute VB_Name = "SRVCrecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsNewCust As New ADODB.Recordset
Dim rsCompany As New ADODB.Recordset
Dim ActRow As Single
Dim ViewOnly As Boolean, AddNew As Boolean, CustSrch As Boolean, CustBln As Boolean, CompSrch As Boolean

Sub ClearEntries()
    txtCustomer.Text = ""
    txtAddress.Text = ""
    txtContactNo.Text = ""
    txtDatePurchased.Text = ""
    txtDealer.Text = ""
    txtAcct.Text = ""
    txtModel.Text = ""
    txtEngine.Text = ""
    txtChassis.Text = ""
    txtPlateNo.Text = ""
    txtBrand.Text = ""
    flxSrvc.Rows = 1
    flxSrvc.Rows = 2
End Sub

Private Sub BtnSaveCancel(bln As Boolean)
    Dim i As Integer
    
    If bln Then
        For i = 1 To 5
            cmdTrans(i).Visible = False
        Next i
        cmdTrans(6).Visible = True
        cmdTrans(7).Visible = True
        cmdTrans(6).Left = 0
        cmdTrans(7).Left = cmdTrans(6).Width
    Else
        For i = 1 To 5
            cmdTrans(i).Visible = True
        Next i
        cmdTrans(6).Visible = False
        cmdTrans(7).Visible = False
    End If
End Sub

Public Sub ShowRecord(MyAcct As String)
    Dim rsCustomer As New ADODB.Recordset
    Dim rsJO As New ADODB.Recordset
    Dim Qry As String, myJR As String
    Dim PrtSale As Single, Ttl As Single
    
    'Display Customer Record
    '***Access
    'Qry = "SELECT SRVCUSTOMER.* From SRVCUSTOMER WHERE (((SRVCUSTOMER.ACCOUNTNO)=" & MyAcct & "));"
    'mySQL
    Qry = "SELECT SRVCUSTOMER.* From SRVCUSTOMER WHERE (((SRVCUSTOMER.ACCOUNTNO)=" & MyAcct & "));"
    
    With rsCustomer
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            txtCustomer.Text = !CUSTOMER
            txtAddress.Text = !Address
            txtContactNo.Text = IIf(IsNull(!CONTACTNO), "", !CONTACTNO)
            txtCustID.Text = IIf(IsNull(!CUSTOMERID), 0, !CUSTOMERID)
            txtCType.Text = IIf(IsNull(!Ctype), "", !Ctype)
            'OptCType(0).Enabled = True
            'OptCType(1).Enabled = True
            If Val(txtCustID.Text) > 0 Then
                'OptCType(1).Value = True
                lblAdrsCom.Caption = "COMPANY"
                lblPurDes.Caption = "DESIGNATION"
                txtDatePurchased.Text = IIf(IsNull(!DESIGNATION), "", !DESIGNATION)
                lblDealer.Visible = False
                txtDealer.Visible = False
            Else
                'OptCType(0).Value = True
                lblAdrsCom.Caption = "ADDRESS"
                lblPurDes.Caption = "DATE PURCHASED"
                lblDealer.Visible = True
                txtDealer.Visible = True
                txtDatePurchased.Text = IIf(IsNull(!DATEPURCHASED), "", Format$(!DATEPURCHASED, "Mmmm dd, yyyy"))
                txtDealer.Text = IIf(IsNull(!DEALER), "", !DEALER)
            End If
            txtAcct.Text = !ACCOUNTNO
            txtModel.Text = IIf(IsNull(!MODEL), "", !MODEL)
            txtEngine.Text = IIf(IsNull(!ENGINE), "", !ENGINE)
            txtChassis.Text = IIf(IsNull(!CHASSIS), "", !CHASSIS)
            txtPlateNo.Text = IIf(IsNull(!PLATENO), "", !PLATENO)
            txtBrand.Text = IIf(IsNull(!BRAND), "", !BRAND)
            'Access
            'If !ADDTLMC = True Then
            'mySQL
            If !ADDTLMC = 1 Then
                cmdMCStat.Caption = "Additional MC"
            Else
                cmdMCStat.Caption = "Current MC"
            End If
        End If
        .Close
    End With
    If txtAcct.Text = "" Then Exit Sub
    '***Access
    'Qry = "SELECT SRVCJOBORDER.*, SRVCJOBORDER.REMARKS From SRVCJOBORDER " & _
          "WHERE (((SRVCJOBORDER.ACCOUNTNO)=" & txtAcct.Text & ") AND ((SRVCJOBORDER.PRINTCOPY)>0) AND ((SRVCJOBORDER.REMARKS)<>'CANCELLED')) ORDER BY SRVCJOBORDER.DATERECEIVED;"
    
    'mySQL
    Qry = "SELECT SRVCJOBORDER.*, SRVCJOBORDER.REMARKS From SRVCJOBORDER " & _
          "WHERE (((SRVCJOBORDER.ACCOUNTNO)=" & txtAcct.Text & ") AND ((SRVCJOBORDER.PRINTCOPY)>0) AND ((SRVCJOBORDER.REMARKS)<>'CANCELLED')) ORDER BY SRVCJOBORDER.DATERECEIVED;"
    
    rsJO.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    If rsJO.RecordCount > 0 Then
        flxSrvc.Rows = 1
        rsJO.MoveFirst
        Do While Not rsJO.EOF
            myJR = JobRequest(rsJO!JOBORDERID)
            PrtSale = 0 'Get Parts Sale
            Ttl = rsJO!TotalCost + PrtSale
            flxSrvc.AddItem vbTab & Format$(rsJO!DateReceived, "mm/dd/yyyy") & vbTab & rsJO!JOBORDERNO & vbTab & rsJO!KMReading & vbTab & myJR & vbTab & Format$(rsJO!TotalCost, "###,##0.00") & vbTab & Format$(PrtSale, "###,##0.00") & vbTab & Format$(Ttl, "###,##0.00") & vbTab & vbTab & IIf(IsNull(rsJO!Remarks), "", rsJO!Remarks) & " " & rsJO!JOType & vbTab & IIf(IsNull(rsJO!NxtSrvc), "", Format$(rsJO!NxtSrvc, "mm/dd/yyyy")) & vbTab & rsJO!Mechanic & vbTab & rsJO!JOBORDERID
            rsJO.MoveNext
        Loop
        ActRow = 1
    End If
    rsJO.Close
End Sub

Function JobRequest(MyJOid As Long) As String
    Dim rsJD As New ADODB.Recordset
    Dim Qry As String, Tmp As String
    
    '***Access
    'Qry = "SELECT SRVJOBDETAILS.JOBREQUEST From SRVJOBDETAILS " & _
          "WHERE (((SRVJOBDETAILS.JOBORDERID)=" & MyJOid & "));"
    'mySQL
    Qry = "SELECT SRVJOBDETAILS.JOBREQUEST From SRVJOBDETAILS " & _
          "WHERE (((SRVJOBDETAILS.JOBORDERID)=" & MyJOid & "));"
    With rsJD
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        Tmp = ""
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                Tmp = IIf(Tmp <> "", Tmp & ", ", "") & IIf(IsNull(!JobRequest), "", !JobRequest)
                .MoveNext
            Loop
        End If
        .Close
    End With
    JobRequest = Tmp
End Function

Private Sub cboCompany_Change()
    Dim i As Long, j As Long
    Dim strPartial As String, strTotal As String
    Dim X As Integer
    
    If CompSrch Then
        CompSrch = False
        Exit Sub
    End If
    With cboCompany
        strPartial = .Text
        i = SendMessage(.hwnd, CB_FINDSTRING, -1, ByVal strPartial) 'If match found, append unmatched characters
        If i <> -1 Then
            strTotal = .List(i)
            rsCompany.MoveFirst
            rsCompany.Move i
            txtNewID.Text = rsCompany!CUSTOMERID
            j = Len(strTotal) - Len(strPartial)
            If j <> 0 Then
                CompSrch = True
                .SelText = Right$(strTotal, j)
                .SelStart = Len(strPartial)
                .SelLength = j
            End If
        Else
            txtNewID.Text = ""
        End If
    End With
End Sub

Private Sub cboCompany_Click()
    Call cboCompany_Change
End Sub

Private Sub cboCompany_GotFocus()
   Dim Tmp
    
   Const CB_SHOWDROPDOWN = &H14F
   Tmp = SendMessage(cboCompany.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub

Private Sub cboCompany_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then CompSrch = True
End Sub

Private Sub cboCompany_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        txtDesignation.SetFocus
    End If
End Sub

Private Sub cboCType_Change()
    Dim i As Long, j As Long
    Dim strPartial As String, strTotal As String
    Dim X As Integer
    
    If CompSrch Then
        CompSrch = False
        Exit Sub
    End If
    With cboCompany
        strPartial = .Text
        i = SendMessage(.hwnd, CB_FINDSTRING, -1, ByVal strPartial) 'If match found, append unmatched characters
        If i <> -1 Then
            strTotal = .List(i)
            j = Len(strTotal) - Len(strPartial)
            If j <> 0 Then
                CompSrch = True
                .SelText = Right$(strTotal, j)
                .SelStart = Len(strPartial)
                .SelLength = j
            End If
        End If
    End With
End Sub

Private Sub cboCType_Click()
    Call cboCType_Change
End Sub

Private Sub cboCType_GotFocus()
   Dim Tmp
    
   Const CB_SHOWDROPDOWN = &H14F
   Tmp = SendMessage(cboCType.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub

Private Sub cboCType_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then CompSrch = True
End Sub

Private Sub cboCType_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        cboCompany.SetFocus
    End If
End Sub

Private Sub cboCType_LostFocus()
    If cboCType.ListIndex = -1 Or cboCType.Text = "INDIVIDUAL" Then
        cboCompany.Clear
        Exit Sub
    End If
    If rsCompany.State = adStateOpen Then rsCompany.Close
    '***Access
    'Qry = "SELECT CUSTOMER.CUSTOMERNAME, CUSTOMER.CUSTOMERID From CUSTOMER " & _
          "WHERE (((CUSTOMER.CTYPE)='" & cboCType.Text & "')) ORDER BY CUSTOMER.CUSTOMERNAME;"
    'mySQL
    Qry = "SELECT CUSTOMER.CUSTOMERNAME, CUSTOMER.CUSTOMERID From CUSTOMER " & _
          "WHERE (((CUSTOMER.CTYPE)='" & cboCType.Text & "')) ORDER BY CUSTOMER.CUSTOMERNAME;"
    
    With rsCompany
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                cboCompany.AddItem !CustomerName
                .MoveNext
            Loop
        End If
    End With
End Sub


Private Sub cmdCancel_Click()
    Dim Ans As Integer, i As Integer
    
    Ans = MsgBox("Do you want to Cancel change of owner?", vbQuestion + vbYesNo, "Cancel")
    If Ans = vbNo Then Exit Sub
    tabCust.Top = 15000
    tabCust.Visible = True
    Frame1.Enabled = True
    flxSrvc.Enabled = True
    For i = 1 To 5
        cmdTrans(i).Enabled = True
    Next i

End Sub

Private Sub cmdMCStat_Click()
    Dim Ans As Integer
    
    Ans = MsgBox("Do you want to Change Status of Customer?", vbQuestion + vbYesNo, "Additional MC")
    If Ans = vbNo Then Exit Sub
    If cmdMCStat.Caption = "Additional MC" Then
        '***Access
        'LYRcon.Execute "UPDATE SRVCUSTOMER SET SRVCUSTOMER.ADDTLMC = False WHERE (((SRVCUSTOMER.ACCOUNTNO)=" & Val(txtAcct.Text) & "));"
        'mySQL
        LYRcon.Execute "UPDATE SRVCUSTOMER SET SRVCUSTOMER.ADDTLMC = 0 WHERE SRVCUSTOMER.ACCOUNTNO=" & Val(txtAcct.Text) & ";"
        
        cmdMCStat.Caption = "Current MC"
    Else
        'LYRcon.Execute "UPDATE SRVCUSTOMER SET SRVCUSTOMER.ADDTLMC = True WHERE (((SRVCUSTOMER.ACCOUNTNO)=" & Val(txtAcct.Text) & "));"
        'mySQL
        LYRcon.Execute "UPDATE SRVCUSTOMER SET SRVCUSTOMER.ADDTLMC = 1 WHERE SRVCUSTOMER.ACCOUNTNO=" & Val(txtAcct.Text) & ";"
        
        cmdMCStat.Caption = "Additional MC"
    End If

End Sub

Private Sub cmdOK_Click()
    Dim rsPrev As New ADODB.Recordset
    Dim rsSrvc As New ADODB.Recordset
    Dim rsMC As New ADODB.Recordset
    Dim Ans As Integer, i As Integer, vMsg As String, Qry As String, nMC As Integer
    
    vMsg = ""
    If txtNewCust.Text = "" Then vMsg = "Incomplete entry on field New Customer"
    If txtNewAdrs.Text = "" Then vMsg = vMsg & vbCrLf & "Incomplete entry on field New Address"
    If cboCType.ListIndex = -1 Then vMsg = vMsg & vbCrLf & "Invalid entry on field New Address"
    'If newCType(0).Value = True Then    'Validate company
    '    If Val(txtNewID.Text) = 0 Then vMsg = vMsg & vbCrLf & "Invalid entry on field Company"
    '    If txtDesignation.Text = "" Then vMsg = vMsg & vbCrLf & "Incomplete entry on field Designation"
    'End If
    If vMsg <> "" Then
        MsgBox vMsg, vbCritical + vbOKOnly, "Validation Error"
        Exit Sub
    End If
    Ans = MsgBox("Do you want to Save changes?", vbQuestion + vbYesNo, "Save")
    If Ans = vbNo Then Exit Sub
    'Save Previous Customer Name
    'LYRcon.Execute "INSERT INTO PREVCUSTOMER ( ACCOUNTNO, PREVCUSTOMER, PREVADDRESS, DATECHANGED) " & _
               "VALUES (" & txtAcct.Text & ", '" & txtCustomer.Text & "', '" & txtAddress.Text & "', #" & Now & "#);"
    'mySQL
    LYRcon.Execute "INSERT INTO PREVCUSTOMER ( ACCOUNTNO, PREVCUSTOMER, PREVADDRESS, DATECHANGED) " & _
               "VALUES (" & txtAcct.Text & ", '" & txtCustomer.Text & "', '" & txtAddress.Text & "', '" & Format(Now, "yyyy-mm-dd") & "');"
    '***Access
    'Qry = "SELECT SRVCUSTOMER.* From SRVCUSTOMER WHERE (((SRVCUSTOMER.ACCOUNTNO)=" & txtAcct.Text & "));"
    'With rsSrvc
    '    .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    '    If .RecordCount > 0 Then
    '        .MoveFirst
    '        !CUSTOMER = txtNewCust.Text
    '        !Address = txtNewAdrs.Text
    '        If Val(txtNewAcct.Text) > 0 Then
                'Access
                '!ADDTLMC = True
                'mySQL
    '           !ADDTLMC = 1
                
    '        Else
    '            !ADDTLMC = False
                'Access
                '!ONACCOUNT = False
                'mySQL
    '           !ONACCOUNT = 0
    '        End If
    '        If Val(txtNewID.Text) > 0 Then 'Company
    '            !CUSTOMERID = Val(txtNewID.Text)
    '            !DESIGNATION = txtDesignation.Text
    '            !Ctype = cboCType.Text
                 'Access
                 '!ONACCOUNT = True
                 'mySQL
    '            !ONACCOUNT = 1
    '            If Val(txtNewAcct.Text) = 0 Then !Address = cboCompany.Text
    '        Else
    '            !CUSTOMERID = 0
    '            !Ctype = cboCType.Text
    '            'Access
    '            '!ONACCOUNT = False
    '            'mySQL
    '            !ONACCOUNT = 0
    '        End If
    '        .Update
    '    End If
    '    .Close
    'End With
            
    'mySQL statement
    If Val(txtNewAcct.Text) > 0 Then
        myAddtlMC = 1
    Else
        myAddtlMC = 0
        myONACCOUNT = 0
    End If
    If Val(txtNewID.Text) > 0 Then 'Company
        myCustoID = Val(txtNewID.Text)
        myDESIG = txtDesignation.Text
        myCTYPE = cboCType.Text
        myOnAcct = 1
        If Val(txtNewAcct.Text) = 0 Then myADDRESS = cboCompany.Text
    Else
        myCustoID = 0
        myCTYPE = cboCType.Text
        myOnAcct = 0
    End If
    
        'LYRcon.Execute "UPDATE srvcustomer SET CUSTOMER = '" & myCUSTOMER & "', ADDRESS = '" & myADDRESS & "', " & _
            "ADDTLMC = " & myAddtlMC & ", ONACCOUNT = " & myOnAcct & ", CUSTOMERID = " & myCustoID & ", " & _
            "DESIGNATION = '" & myDESIG & "', CTYPE = '" & myCTYPE & "' Where srvcustomer.ACCOUNTNO = " & txtAcct.Text & ";"
        LYRcon.Execute "UPDATE srvcustomer SET CUSTOMER = '" & txtNewCust.Text & "', ADDRESS = '" & txtNewAdrs.Text & "', " & _
           "ADDTLMC = " & myAddtlMC & ", ONACCOUNT = " & myOnAcct & ", CUSTOMERID = " & myCustoID & ", " & _
           "DESIGNATION = '" & myDESIG & "', CTYPE = '" & myCTYPE & "' Where srvcustomer.ACCOUNTNO = " & txtAcct.Text & ";"
        'Change MC value to 01234 service record
        Qry = "SELECT srvcustomer.ADDTLMC, srvcustomer.ACCOUNTNO From srvcustomer Where srvcustomer.CUSTOMER = '" & Trim(txtCustomer.Text) & "';"
        rsMC.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If rsMC.RecordCount > 0 Then
            rsMC.MoveFirst
            For nMC = 0 To rsMC.RecordCount - 1
                LYRcon.Execute "Update srvcustomer SET ADDTLMC = " & nMC & " Where srvcustomer.ACCOUNTNO = " & rsMC!ACCOUNTNO & ";"
                rsMC.MoveNext
            Next nMC
        End If
        rsMC.Close
     
     
     
    'Display
    txtCustomer.Text = txtNewCust.Text
    txtCType.Text = cboCType.Text
    If Val(txtNewID.Text) > 0 Then
        txtAddress.Text = cboCompany.Text
        txtCustID.Text = txtNewID.Text
        lblAdrsCom.Caption = "COMPANY"
        lblPurDes.Caption = "DESIGNATION"
        txtDatePurchased.Text = txtDesignation.Text
    Else
        lblAdrsCom.Caption = "ADDRESS"
        lblPurDes.Caption = "DATE PURCHASED"
        txtDatePurchased.Text = ""
        txtAddress.Text = txtNewAdrs.Text
        txtCustID.Text = ""
    End If
    MsgBox "Customer Name successfully changed!", vbInformation + vbOKOnly, "Save"
    tabCust.Top = 15000
    tabCust.Visible = True
    Frame1.Enabled = True
    flxSrvc.Enabled = True
    For i = 1 To 5
        cmdTrans(i).Enabled = True
    Next i
    
End Sub

Private Sub cmdTrans_Click(Index As Integer)
    Dim rsSettings As New ADODB.Recordset
    Dim Qry As String, vMsg As String, MyAcc As String
    Dim Ans As Integer, i As Integer
    Select Case Index
        
        Case 1  'Edit
            '***Access
            'Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE (((USERACCESS.EMPID)=" & UserID & ") AND ((USERACCESS.MENUINDX)=40));"
            'mySQL
            Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE USERACCESS.EMPID=" & UserID & " AND USERACCESS.MENUINDX=40;"
            MyAcc = GetFieldValue(Qry)
            If MyAcc <> "FULL ACCESS" Then
                MsgBox "You don't have full access to this transaction." & vbCrLf & "Please contact your System Administrator", vbExclamation + vbOKOnly, "Access Denied"
                Exit Sub
            End If
            If Val(txtAcct.Text) = 0 Then
                MsgBox "No available record to change!", vbExclamation + vbOKOnly, "Change Owner"
                Exit Sub
            End If
            
            Ans = MsgBox("Do you want to change owner of motorcyle" & vbCrLf & "with engine no. " & txtEngine.Text & "?", vbQuestion + vbYesNo, "Change Owner")
            If Ans = vbNo Then Exit Sub
            With tabCust
                .Caption = ""
                .Top = (Me.ScaleHeight - .Height) / 2
                .Left = (Me.ScaleWidth - .Width) / 2
                .Visible = True
                .Enabled = False
            End With
            Frame1.Enabled = False
            flxSrvc.Enabled = False
            For i = 1 To 5
                cmdTrans(i).Enabled = False
            Next i
            'Load Customer Names
            If rsNewCust.State = adStateClosed Then
                '***Access
                'Qry = "SELECT SRVCUSTOMER.ACCOUNTNO, SRVCUSTOMER.CUSTOMER, SRVCUSTOMER.ADDRESS, SRVCUSTOMER.CUSTOMERID, SRVCUSTOMER.DESIGNATION, SRVCUSTOMER.CTYPE " & _
                      "From SRVCUSTOMER Where (((SRVCUSTOMER.ADDTLMC) = False)) ORDER BY SRVCUSTOMER.CUSTOMER;"
                'mySQL
                Qry = "SELECT SRVCUSTOMER.ACCOUNTNO, SRVCUSTOMER.CUSTOMER, SRVCUSTOMER.ADDRESS, SRVCUSTOMER.CUSTOMERID, SRVCUSTOMER.DESIGNATION, SRVCUSTOMER.CTYPE " & _
                      "From SRVCUSTOMER Where SRVCUSTOMER.ADDTLMC = 0 ORDER BY SRVCUSTOMER.CUSTOMER;"
                
                rsNewCust.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
                Set flxNCust.DataSource = rsNewCust
                With flxNCust
                    .Width = txtNewCust.Width
                    .ColWidth(0) = 0    'AccountNo
                    .ColWidth(1) = txtNewCust.Width
                    .ColWidth(2) = 0    'address
                    .ColWidth(3) = 0    'CustomerId
                    .ColWidth(4) = 0    'Designation
                    .ColWidth(5) = 0    'Ctype
                End With
            End If
            If rsSettings.State = adStateClosed Then
                '***Access
                'Qry = "SELECT SETTINGS.VALUE From SETTINGS WHERE (((SETTINGS.DESCRIPTION)='Customer Type')) ORDER BY SETTINGS.VALUE;"
                'mySQL
                Qry = "SELECT SETTINGS.VALUE From SETTINGS WHERE SETTINGS.DESCRIPTION='Customer Type' ORDER BY SETTINGS.VALUE;"
                
                With rsSettings
                    .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
                    If .RecordCount > 0 Then
                        .MoveFirst
                        Do While Not .EOF
                            cboCType.AddItem !Value
                            .MoveNext
                        Loop
                    End If
                End With
            End If
            tabCust.Enabled = True
            'Call newCType_Click(1)
            txtNewCust.SetFocus
            
        Case 2  'Find
            Set frmSearch.FrmSource = SRVCrecord
            '***Access
            'Qry = "SELECT SRVCUSTOMER.ACCOUNTNO, [CUSTOMER] & ' - ' & [ENGINE] AS ENGCUST From SRVCUSTOMER ORDER BY [CUSTOMER] & ' - ' & [ENGINE];"
            'mySQL
            Qry = "SELECT SRVCUSTOMER.ACCOUNTNO, concat(SRVCUSTOMER.CUSTOMER,' - ',SRVCUSTOMER.ENGINE) From SRVCUSTOMER ORDER BY concat(SRVCUSTOMER.CUSTOMER,' - ', SRVCUSTOMER.ENGINE);"
            
            frmSearch.QrySearch = Qry
            frmSearch.lblCap.Caption = "Search for Customer"
            
            frmSearch.Show vbModal
            
        Case 3  'View J.O
            With flxSrvc
                If ActRow = 0 Or ActRow > .Rows - 1 Then
                    Ans = MsgBox("No Job Order selected!" & vbCrLf & "Do you want to load Job Order anyway?", vbQuestion + vbYesNo, "Job Order")
                    If Ans = vbNo Then Exit Sub
                    JobOrder.Show
                Else
                    Call JobOrder.DisplayJobOrder(flxSrvc.TextMatrix(ActRow, 12))
                End If
            End With
            Unload Me
        Case 4  'Print
        
        Case 5  'Exit
            Unload Me
            'MotoRep.tbMain.Visible = True
            'MotoRep.imgBG.Visible = True
            MotoRep.toolBarButton.Visible = True
        Case 6  'Save
            
            vMsg = ValidateSrvCust
            If vMsg <> "" Then
                MsgBox vMsg, vbCritical + vbOKOnly, "Validation Error"
                Exit Sub
            End If
            Ans = MsgBox("Do you want to Save Customer?", vbQuestion + vbYesNo, "Save")
            If Ans = vbNo Then Exit Sub
            Call SaveIndCustomer
            MsgBox "Customer Successfully Saved!", vbInformation + vbOKOnly
            Frame1.Enabled = False
            Call BtnSaveCancel(False)
        Case 7  'Cancel
            Ans = MsgBox("Do you want to Cancel " & IIf(AddNew, "Entry", "Editing") & "?", vbQuestion + vbYesNo, "Cancel")
            If Ans = vbNo Then Exit Sub
            If AddNew Then
                Call ClearEntries
            Else
                Call ShowRecord(Val(txtAcct.Text))
            End If
            Frame1.Enabled = False
            Call BtnSaveCancel(False)
        
    End Select
End Sub
Private Sub SaveIndCustomer()
    Dim rsCust As New ADODB.Recordset
    Dim Qry As String
    
    'If AddNew Then
    '    rsCust.Open "SrvCustomer", LYRcon, adOpenKeyset, adLockOptimistic
    '    rsCust.AddNew
    'Else
    '    Qry = "SELECT SRVCUSTOMER.* From SRVCUSTOMER WHERE (((SRVCUSTOMER.ACCOUNTNO)=" & txtAcct.Text & "));"
    '    rsCust.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    '    rsCust.MoveFirst
    'End If
    'rsCust!CUSTOMER = txtCustomer.Text
    'rsCust!Address = txtAddress.Text
    'rsCust!CONTACTNO = txtContactNo.Text
    'If IsDate(txtDatePurchased.Text) Then
    '    rsCust!DATEPURCHASED = txtDatePurchased.Text
    'End If
    'rsCust!MODEL = txtModel.Text
    'rsCust!ENGINE = txtEngine.Text
    'rsCust!CHASSIS = txtChassis.Text
    'rsCust!PLATENO = txtPlateNo.Text
    'rsCust!DEALER = txtDealer.Text
    'rsCust!BRAND = txtBrand.Text
    'rsCust!Ctype = "INDIVIDUAL"
    'Access
    'rsCust!ONACCOUNT = False
    'mySQL
    'rsCust!ONACCOUNT = 0
    'rsCust.Update
    'rsCust.Close
    
    'mySQL
    If IsDate(txtDatePurchased.Text) Then
        myDATEPURCHASED = Format(txtDatePurchased.Text, "yyyy-mm-dd")
    End If
    If AddNew Then
        LYRcon.Execute "INSERT INTO srvcustomer(CUSTOMER, ADDRESS, CONTACTNO, DATEPURCHASED, MODEL, ENGINE," & _
            "CHASSIS, PLATENO, DEALER, BRAND, CTYPE, ONACCOUNT)VALUES('" & txtCustomer.Text & "', " & _
            "'" & txtAddress.Text & "', '" & txtContactNo.Text & "', '" & myDATEPURCHASED & "', '" & txtModel.Text & "', " & _
            "'" & txtEngine.Text & "', '" & txtChassis.Text & "', '" & txtPlateNo.Text & "', '" & txtDealer.Text & "', " & _
            "'" & txtBrand.Text & "', 'INDIVIDUAL',0);"
    Else
        LYRcon.Execute "Update srvcustomer SET CUSTOMER = '" & txtCustomer.Text & "', ADDRESS = '" & txtAddress.Text & "', " & _
            "CONTACTNO = '" & txtContactNo.Text & "', DATEPURCHASED = '" & txtDatePurchased.Text & "', MODEL = '" & txtModel.Text & "', " & _
            "ENGINE = '" & txtEngine.Text & "', CHASSIS = '" & txtChassis.Text & "', PLATENO = '" & txtPlateNo.Text & "', " & _
            "DEALER = '" & txtDealer.Text & "', BRAND = '" & txtBrand.Text & "', CTYPE = 'INDIVIDUAL', ONACCOUNT = 0" & _
            " Where srvcustomer.ACCOUNTNO = " & txtAcct.Text & ";"
    End If
End Sub

Function ValidateSrvCust() As String
    Dim eMsg As String, vQry As String
    
    eMsg = ""
    If txtEngine.Text = "" Then
        eMsg = "Incomplete Entry on field Engine"
    Else
        If AddNew Then
            '***Access
            'vQry = "SELECT SRVCUSTOMER.ENGINE From SRVCUSTOMER WHERE (((SRVCUSTOMER.ENGINE)='" & txtEngine.Text & "'));"
            'mySQL
            vQry = "SELECT SRVCUSTOMER.ENGINE From SRVCUSTOMER WHERE SRVCUSTOMER.ENGINE='" & txtEngine.Text & "';"
            
            If CtrlNumExist(vQry) Then
                eMsg = "Engine number " & txtEngine.Text & " already exist"
            End If
        End If
    End If
    If txtCustomer.Text = "" Then eMsg = eMsg & vbCrLf & "Incomplete entry on field Customer Name"
    If txtAddress.Text = "" Then eMsg = eMsg & vbCrLf & "Incomplete entry on field Address"
    If txtBrand.Text = "" Then eMsg = eMsg & vbCrLf & "Incomplete entry on field Brand"
    If txtModel.Text = "" Then eMsg = eMsg & vbCrLf & "Incomplete entry on field Model"
    If txtChassis.Text = "" Then eMsg = eMsg & vbCrLf & "Incomplete entry on field Chassis"
    ValidateSrvCust = eMsg
End Function
Private Sub flxNCust_Click()
    txtNewCust.Text = flxNCust.TextMatrix(flxNCust.Row, 1)
End Sub
Private Sub flxNCust_EnterCell()
    CustSrch = True
    txtNewCust.Text = flxNCust.TextMatrix(flxNCust.Row, 1)
End Sub
Private Sub flxNCust_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        flxNCust.Visible = False
        txtNewCust.SetFocus
    End If
End Sub
Private Sub flxNCust_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call txtNewCust_Change
        flxNCust.Visible = False
        txtNewAdrs.SetFocus
    End If
End Sub
Private Sub flxSrvc_Click()
    ActRow = flxSrvc.Row
End Sub
Private Sub Form_Load()
    Frame1.Enabled = False
    MotoRep.sBar.Panels(6).Text = "Customer Service Record."
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If rsNewCust.State = adStateOpen Then rsNewCust.Close
    If rsCompany.State = adStateOpen Then rsCompany.Close
    MotoRep.sBar.Panels(6).Text = ""
End Sub


Private Sub RXResizer1_ResizeComplete()
    Dim i As Integer
    
    Label3(3).Left = Frame1.Left
    Label3(2).Left = Frame1.Left + 30
    Label3(3).Top = Frame1.Top - (Label3(3).Height + 30)
    Label3(2).Top = Label3(3).Top + 30
    Label3(5).Top = Label3(3).Top
    Label3(4).Top = Label3(2).Top
    Label3(5).Left = Frame1.Left + Frame1.Width - Label3(5).Width
    Label3(4).Left = Label3(5).Left + 30
    Label3(1).Left = flxSrvc.Left
    Label3(0).Left = flxSrvc.Left + 30
    Label3(1).Top = flxSrvc.Top - (Label3(1).Height + 30)
    Label3(0).Top = Label3(1).Top + 30
    With flxSrvc
        .ColWidth(0) = 300
        .ColWidth(12) = 0
        For i = 1 To 11
            If i = 4 Or i = 9 Or i = 11 Then
                .ColWidth(i) = 2500
            Else
                .ColWidth(i) = 1500
            End If
        Next i
    End With
End Sub

Private Sub txtAcct_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtAddress_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtChassis_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtContactNo_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub



Private Sub txtCustomer_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtDatePurchased_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtDealer_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub
Private Sub txtDesignation_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then cmdOK.SetFocus
End Sub
Private Sub txtEngine_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub
Private Sub txtModel_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub
Private Sub txtNewAdrs_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then cboCType.SetFocus
End Sub
Private Sub txtNewCust_Change()
    With flxNCust
        If CustSrch Then
            CustSrch = False
            Exit Sub
        End If
        .Top = txtNewCust.Top + txtNewCust.Height + 30
        .Left = txtNewCust.Left
        .Width = txtNewCust.Width
        .Visible = True
        .ZOrder
        If txtNewCust.Text = "" Then .Visible = False: Exit Sub
        rsNewCust.Find "[CUSTOMER] LIKE '" + txtNewCust.Text + "%'", , adSearchForward, 1
        If Not rsNewCust.EOF Then
            If Not CustBln Then
                txtNewCust.Tag = Len(txtNewCust.Text)
                CustBln = True
            End If
            .TopRow = rsNewCust.AbsolutePosition - 1
            .Row = rsNewCust.AbsolutePosition - 1
            .Col = 0: .ColSel = 4
            CustSrch = True
            txtNewCust.Text = rsNewCust!CUSTOMER
            'Insert Additional retrieval here
            txtNewAdrs.Text = flxNCust.TextMatrix(flxNCust.Row, 2)
            txtNewAcct.Text = flxNCust.TextMatrix(flxNCust.Row, 0)
            txtNewID.Text = flxNCust.TextMatrix(flxNCust.Row, 3)
            txtDesig.Text = flxNCust.TextMatrix(flxNCust.Row, 4)
            
            txtNewCust.SelStart = Val(txtNewCust.Tag)
            txtNewCust.SelLength = Len(txtNewCust.Text)
            CustBln = False
        Else
            txtNewAdrs.Text = ""
            txtNewAcct.Text = ""
            txtNewID.Text = ""
            txtDesig.Text = ""
            .Visible = False
        End If
    End With

End Sub
Private Sub txtNewCust_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDown
            If flxNCust.Visible Then
                flxNCust.SetFocus
            Else
                txtNewAdrs.SetFocus
            End If
        Case vbKeyEscape
            txtNewCust.Text = ""
            flxNCust.Visible = False
        Case vbKeyDelete
            CustSrch = True
        Case vbKeyBack
            CustSrch = True
    End Select
End Sub
Private Sub txtNewCust_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        flxNCust.Visible = False
        txtNewAdrs.SetFocus
    End If
End Sub
Private Sub txtPlateNo_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub
