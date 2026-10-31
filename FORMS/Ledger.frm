VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Ledger 
   Caption         =   "Customer Accounts"
   ClientHeight    =   9555
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12705
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9555
   ScaleWidth      =   12705
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdC 
      Caption         =   "&C"
      Height          =   255
      Left            =   2160
      TabIndex        =   85
      Top             =   10560
      Width           =   855
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   4920
      Top             =   240
   End
   Begin VB.TextBox txtCustomerID 
      Height          =   285
      Left            =   720
      TabIndex        =   57
      Top             =   10320
      Visible         =   0   'False
      Width           =   1095
   End
   Begin TabDlg.SSTab tabDocs 
      Height          =   3615
      Left            =   6600
      TabIndex        =   34
      Tag             =   "XYF"
      Top             =   5760
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   6376
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "INVOICE"
      TabPicture(0)   =   "Ledger.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblGrp2(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblGrp2(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblTtl(3)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblInvNo"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblInvDate"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblInvRef"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblGrp2(12)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblPartsMan"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "flxInv"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtTtlInv"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cmdInvDetails"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).ControlCount=   11
      TabCaption(1)   =   "JOB ORDER"
      TabPicture(1)   =   "Ledger.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblGrp2(2)"
      Tab(1).Control(1)=   "lblGrp2(3)"
      Tab(1).Control(2)=   "lblGrp2(4)"
      Tab(1).Control(3)=   "lblGrp2(5)"
      Tab(1).Control(4)=   "lblGrp2(6)"
      Tab(1).Control(5)=   "lblGrp2(7)"
      Tab(1).Control(6)=   "lblTtl(4)"
      Tab(1).Control(7)=   "lblJobNo"
      Tab(1).Control(8)=   "lblJobDate"
      Tab(1).Control(9)=   "lblType"
      Tab(1).Control(10)=   "lblModel"
      Tab(1).Control(11)=   "lblEngine"
      Tab(1).Control(12)=   "lblChassis"
      Tab(1).Control(13)=   "lblCustomer"
      Tab(1).Control(14)=   "flxJob"
      Tab(1).Control(15)=   "txtTtlJob"
      Tab(1).ControlCount=   16
      TabCaption(2)   =   "PCV"
      TabPicture(2)   =   "Ledger.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txtTtlPCV"
      Tab(2).Control(1)=   "flxPCV"
      Tab(2).Control(2)=   "lblATitle"
      Tab(2).Control(3)=   "lblPayee"
      Tab(2).Control(4)=   "lblPcvDate"
      Tab(2).Control(5)=   "lblPcvNo"
      Tab(2).Control(6)=   "lblTtl(5)"
      Tab(2).Control(7)=   "lblGrp2(11)"
      Tab(2).Control(8)=   "lblGrp2(10)"
      Tab(2).Control(9)=   "lblGrp2(9)"
      Tab(2).Control(10)=   "lblGrp2(8)"
      Tab(2).ControlCount=   11
      Begin lvButton.lvButtons_H cmdInvDetails 
         Height          =   255
         Left            =   120
         TabIndex        =   84
         Tag             =   "XYF"
         Top             =   3270
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   450
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Webdings"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         ImgAlign        =   5
         Image           =   "Ledger.frx":0054
         ImgSize         =   24
         cBack           =   -2147483633
      End
      Begin VB.TextBox txtTtlPCV 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   -70500
         TabIndex        =   54
         Tag             =   "XYF"
         Top             =   3240
         Width           =   1455
      End
      Begin VB.TextBox txtTtlJob 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -70320
         TabIndex        =   47
         Tag             =   "XYF"
         Top             =   3240
         Width           =   1275
      End
      Begin VB.TextBox txtTtlInv 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4500
         TabIndex        =   46
         Tag             =   "XYF"
         Top             =   3240
         Width           =   1455
      End
      Begin MSFlexGridLib.MSFlexGrid flxInv 
         Height          =   2295
         Left            =   60
         TabIndex        =   35
         Tag             =   "XY"
         Top             =   960
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   4048
         _Version        =   393216
         Cols            =   7
         BackColorBkg    =   -2147483633
         AllowUserResizing=   1
         FormatString    =   $"Ledger.frx":05EE
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
      Begin MSFlexGridLib.MSFlexGrid flxJob 
         Height          =   1935
         Left            =   -74940
         TabIndex        =   36
         Tag             =   "XY"
         Top             =   1320
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   3413
         _Version        =   393216
         Cols            =   5
         FixedCols       =   0
         RowHeightMin    =   300
         BackColorBkg    =   -2147483633
         ScrollBars      =   2
         FormatString    =   "^    |< JOB   REQUEST                  |>  LABOR  |> DISCOUNT  |>  T O T A L  "
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
      Begin MSFlexGridLib.MSFlexGrid flxPCV 
         Height          =   2255
         Left            =   -74940
         TabIndex        =   49
         Tag             =   "XY"
         Top             =   1000
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   3969
         _Version        =   393216
         Cols            =   3
         FixedCols       =   0
         RowHeightMin    =   300
         BackColorBkg    =   -2147483633
         ScrollBars      =   2
         FormatString    =   "^    |< P A R T I C U L A R S                         |>  T O T A L  "
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
      Begin VB.Label lblCustomer 
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
         Left            =   -74880
         TabIndex        =   73
         Tag             =   "XYF"
         Top             =   3270
         Width           =   3375
      End
      Begin VB.Label lblPartsMan 
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
         Left            =   600
         TabIndex        =   72
         Tag             =   "XYF"
         Top             =   3270
         Width           =   3135
      End
      Begin VB.Label lblGrp2 
         Alignment       =   1  'Right Justify
         Caption         =   "REFERENCE"
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
         Index           =   12
         Left            =   120
         TabIndex        =   71
         Tag             =   "XYF"
         Top             =   690
         Width           =   1095
      End
      Begin VB.Label lblInvRef 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1320
         TabIndex        =   70
         Tag             =   "XYF"
         Top             =   690
         Width           =   4455
      End
      Begin VB.Label lblATitle 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   -71400
         TabIndex        =   69
         Tag             =   "XYF"
         Top             =   690
         Width           =   2295
      End
      Begin VB.Label lblPayee 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   -71400
         TabIndex        =   68
         Tag             =   "XYF"
         Top             =   420
         Width           =   2295
      End
      Begin VB.Label lblPcvDate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   -73920
         TabIndex        =   67
         Tag             =   "XYF"
         Top             =   690
         Width           =   1215
      End
      Begin VB.Label lblPcvNo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   -73920
         TabIndex        =   66
         Tag             =   "XYF"
         Top             =   420
         Width           =   1215
      End
      Begin VB.Label lblChassis 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   -71040
         TabIndex        =   65
         Tag             =   "XYF"
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label lblEngine 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   -71040
         TabIndex        =   64
         Tag             =   "XYF"
         Top             =   690
         Width           =   1935
      End
      Begin VB.Label lblModel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   -71040
         TabIndex        =   63
         Tag             =   "XYF"
         Top             =   420
         Width           =   1935
      End
      Begin VB.Label lblType 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   -73560
         TabIndex        =   62
         Tag             =   "XYF"
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label lblJobDate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   -73560
         TabIndex        =   61
         Tag             =   "XYF"
         Top             =   690
         Width           =   1215
      End
      Begin VB.Label lblJobNo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   -73560
         TabIndex        =   60
         Tag             =   "XYF"
         Top             =   420
         Width           =   1215
      End
      Begin VB.Label lblInvDate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4560
         TabIndex        =   59
         Tag             =   "XYF"
         Top             =   420
         Width           =   1215
      End
      Begin VB.Label lblInvNo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1320
         TabIndex        =   58
         Tag             =   "XYF"
         Top             =   420
         Width           =   1215
      End
      Begin VB.Label lblTtl 
         Alignment       =   1  'Right Justify
         Caption         =   "TOTAL:"
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
         Left            =   -72000
         TabIndex        =   55
         Tag             =   "XYF"
         Top             =   3300
         Width           =   1455
      End
      Begin VB.Label lblGrp2 
         Alignment       =   1  'Right Justify
         Caption         =   "ACCT. TITLE"
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
         Index           =   11
         Left            =   -72600
         TabIndex        =   53
         Tag             =   "XYF"
         Top             =   690
         Width           =   1095
      End
      Begin VB.Label lblGrp2 
         Alignment       =   1  'Right Justify
         Caption         =   "PAYEE"
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
         Index           =   10
         Left            =   -72600
         TabIndex        =   52
         Tag             =   "XYF"
         Top             =   420
         Width           =   1095
      End
      Begin VB.Label lblGrp2 
         Alignment       =   1  'Right Justify
         Caption         =   "PCV DATE"
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
         Left            =   -74880
         TabIndex        =   51
         Tag             =   "XYF"
         Top             =   690
         Width           =   855
      End
      Begin VB.Label lblGrp2 
         Alignment       =   1  'Right Justify
         Caption         =   "PCV NO."
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
         Left            =   -74880
         TabIndex        =   50
         Tag             =   "XYF"
         Top             =   420
         Width           =   855
      End
      Begin VB.Label lblTtl 
         Alignment       =   1  'Right Justify
         Caption         =   "TOTAL:"
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
         Left            =   -71760
         TabIndex        =   48
         Tag             =   "XYF"
         Top             =   3270
         Width           =   1455
      End
      Begin VB.Label lblTtl 
         Alignment       =   1  'Right Justify
         Caption         =   "TOTAL:"
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
         Left            =   3000
         TabIndex        =   45
         Tag             =   "XYF"
         Top             =   3270
         Width           =   1455
      End
      Begin VB.Label lblGrp2 
         Alignment       =   1  'Right Justify
         Caption         =   "CHASSIS#"
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
         Left            =   -72240
         TabIndex        =   44
         Tag             =   "XYF"
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label lblGrp2 
         Alignment       =   1  'Right Justify
         Caption         =   "ENGINE#"
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
         Left            =   -72240
         TabIndex        =   43
         Tag             =   "XYF"
         Top             =   690
         Width           =   1095
      End
      Begin VB.Label lblGrp2 
         Alignment       =   1  'Right Justify
         Caption         =   "MC MODEL"
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
         Left            =   -72240
         TabIndex        =   42
         Tag             =   "XYF"
         Top             =   420
         Width           =   1095
      End
      Begin VB.Label lblGrp2 
         Alignment       =   1  'Right Justify
         Caption         =   "JOB TYPE"
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
         Left            =   -74880
         TabIndex        =   41
         Tag             =   "XYF"
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label lblGrp2 
         Alignment       =   1  'Right Justify
         Caption         =   "JOB DATE"
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
         Left            =   -74880
         TabIndex        =   40
         Tag             =   "XYF"
         Top             =   690
         Width           =   1095
      End
      Begin VB.Label lblGrp2 
         Alignment       =   1  'Right Justify
         Caption         =   "JOB ORDER#"
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
         Left            =   -74880
         TabIndex        =   39
         Tag             =   "XYF"
         Top             =   420
         Width           =   1095
      End
      Begin VB.Label lblGrp2 
         Alignment       =   1  'Right Justify
         Caption         =   "INV. DATE:"
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
         Left            =   3240
         TabIndex        =   38
         Tag             =   "XYF"
         Top             =   420
         Width           =   1215
      End
      Begin VB.Label lblGrp2 
         Alignment       =   1  'Right Justify
         Caption         =   "INVOICE NO."
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
         TabIndex        =   37
         Tag             =   "XYF"
         Top             =   420
         Width           =   1095
      End
   End
   Begin TabDlg.SSTab tabAccts 
      Height          =   3135
      Left            =   120
      TabIndex        =   14
      Tag             =   "XYF"
      Top             =   2520
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   5530
      _Version        =   393216
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "Spare &Parts"
      TabPicture(0)   =   "Ledger.frx":0677
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblTtl(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "flxSP"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cboSPacct"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtFindSP"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdFindSP"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "flxSPTtl"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Ser&vice"
      TabPicture(1)   =   "Ledger.frx":0693
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtFindSVC"
      Tab(1).Control(1)=   "cboSVCacct"
      Tab(1).Control(2)=   "flxSVC"
      Tab(1).Control(3)=   "cmdFindSVC"
      Tab(1).Control(4)=   "flxSVCttl"
      Tab(1).Control(5)=   "lblTtl(1)"
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "&Others"
      TabPicture(2)   =   "Ledger.frx":06AF
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txtTtlOthers"
      Tab(2).Control(1)=   "cboOtherAcct"
      Tab(2).Control(2)=   "txtFindOthers"
      Tab(2).Control(3)=   "cmdFndOthers"
      Tab(2).Control(4)=   "flxOthers"
      Tab(2).Control(5)=   "lblTtl(2)"
      Tab(2).ControlCount=   6
      TabCaption(3)   =   "P&ayments"
      TabPicture(3)   =   "Ledger.frx":06CB
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "txtFindPay"
      Tab(3).Control(1)=   "cmdFindPay"
      Tab(3).Control(2)=   "flxPay"
      Tab(3).Control(3)=   "flxPAYttl"
      Tab(3).Control(4)=   "lblTtl(6)"
      Tab(3).ControlCount=   5
      TabCaption(4)   =   "&DM/CM"
      TabPicture(4)   =   "Ledger.frx":06E7
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "txtFindDMCM"
      Tab(4).Control(1)=   "cmdFindDMCM"
      Tab(4).Control(2)=   "flxDMCM"
      Tab(4).Control(3)=   "flxDMCMttl"
      Tab(4).Control(4)=   "lblTtl(7)"
      Tab(4).ControlCount=   5
      Begin VB.TextBox txtFindDMCM 
         Height          =   315
         Left            =   -64200
         TabIndex        =   79
         Tag             =   "XYF"
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox txtFindPay 
         Height          =   315
         Left            =   -64200
         TabIndex        =   74
         Tag             =   "XYF"
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox txtTtlOthers 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -66720
         Locked          =   -1  'True
         TabIndex        =   32
         Tag             =   "XYF"
         Top             =   2740
         Width           =   1815
      End
      Begin VB.ComboBox cboOtherAcct 
         Height          =   315
         Left            =   -74880
         TabIndex        =   29
         Tag             =   "XYF"
         Top             =   360
         Width           =   2295
      End
      Begin VB.TextBox txtFindOthers 
         Height          =   315
         Left            =   -64200
         TabIndex        =   28
         Tag             =   "XYF"
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox txtFindSVC 
         Height          =   315
         Left            =   -64200
         TabIndex        =   23
         Tag             =   "XYF"
         Top             =   360
         Width           =   1575
      End
      Begin VB.ComboBox cboSVCacct 
         Height          =   315
         Left            =   -74880
         TabIndex        =   21
         Tag             =   "XYF"
         Top             =   360
         Width           =   2295
      End
      Begin MSFlexGridLib.MSFlexGrid flxSPTtl 
         Height          =   300
         Left            =   5400
         TabIndex        =   20
         Tag             =   "XYF"
         Top             =   2740
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   529
         _Version        =   393216
         Rows            =   1
         Cols            =   5
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   300
         BackColorBkg    =   -2147483633
         ScrollBars      =   0
         BorderStyle     =   0
      End
      Begin lvButton.lvButtons_H cmdFindSP 
         Height          =   315
         Left            =   9600
         TabIndex        =   18
         Tag             =   "XYF"
         Top             =   360
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         Caption         =   "Find"
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
         Image           =   "Ledger.frx":0703
         cBack           =   -2147483633
      End
      Begin VB.TextBox txtFindSP 
         Height          =   315
         Left            =   10800
         TabIndex        =   17
         Tag             =   "XYF"
         Top             =   360
         Width           =   1575
      End
      Begin VB.ComboBox cboSPacct 
         Height          =   315
         Left            =   120
         TabIndex        =   16
         Tag             =   "XYF"
         Top             =   360
         Width           =   2295
      End
      Begin MSFlexGridLib.MSFlexGrid flxSP 
         Height          =   2055
         Left            =   120
         TabIndex        =   15
         Tag             =   "XYF"
         Top             =   675
         Width           =   12255
         _ExtentX        =   21616
         _ExtentY        =   3625
         _Version        =   393216
         Cols            =   11
         FixedCols       =   0
         RowHeightMin    =   280
         BackColorBkg    =   -2147483633
         SelectionMode   =   1
         AllowUserResizing=   1
         FormatString    =   $"Ledger.frx":0C9D
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
      Begin MSFlexGridLib.MSFlexGrid flxSVC 
         Height          =   2055
         Left            =   -74880
         TabIndex        =   22
         Tag             =   "XYF"
         Top             =   675
         Width           =   12255
         _ExtentX        =   21616
         _ExtentY        =   3625
         _Version        =   393216
         Cols            =   11
         FixedCols       =   0
         RowHeightMin    =   280
         BackColorBkg    =   -2147483633
         SelectionMode   =   1
         AllowUserResizing=   1
         FormatString    =   $"Ledger.frx":0D46
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
      Begin lvButton.lvButtons_H cmdFindSVC 
         Height          =   315
         Left            =   -65400
         TabIndex        =   24
         Tag             =   "XYF"
         Top             =   360
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         Caption         =   "Find"
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
         Image           =   "Ledger.frx":0DEE
         cBack           =   -2147483633
      End
      Begin MSFlexGridLib.MSFlexGrid flxSVCttl 
         Height          =   300
         Left            =   -69480
         TabIndex        =   26
         Tag             =   "XYF"
         Top             =   2740
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   529
         _Version        =   393216
         Rows            =   1
         Cols            =   5
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   300
         BackColorBkg    =   -2147483633
         ScrollBars      =   0
         BorderStyle     =   0
      End
      Begin lvButton.lvButtons_H cmdFndOthers 
         Height          =   315
         Left            =   -65400
         TabIndex        =   27
         Tag             =   "XYF"
         Top             =   360
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         Caption         =   "Find"
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
         Image           =   "Ledger.frx":1388
         cBack           =   -2147483633
      End
      Begin MSFlexGridLib.MSFlexGrid flxOthers 
         Height          =   2055
         Left            =   -74880
         TabIndex        =   30
         Tag             =   "XYF"
         Top             =   675
         Width           =   12255
         _ExtentX        =   21616
         _ExtentY        =   3625
         _Version        =   393216
         Cols            =   7
         FixedCols       =   0
         RowHeightMin    =   280
         BackColorBkg    =   -2147483633
         SelectionMode   =   1
         AllowUserResizing=   1
         FormatString    =   $"Ledger.frx":1922
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
      Begin lvButton.lvButtons_H cmdFindPay 
         Height          =   315
         Left            =   -65400
         TabIndex        =   75
         Tag             =   "XYF"
         Top             =   360
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         Caption         =   "Find"
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
         Image           =   "Ledger.frx":19BD
         cBack           =   -2147483633
      End
      Begin MSFlexGridLib.MSFlexGrid flxPay 
         Height          =   2055
         Left            =   -74880
         TabIndex        =   76
         Tag             =   "XYF"
         Top             =   675
         Width           =   12255
         _ExtentX        =   21616
         _ExtentY        =   3625
         _Version        =   393216
         Cols            =   7
         FixedCols       =   0
         RowHeightMin    =   280
         BackColorBkg    =   -2147483633
         SelectionMode   =   1
         AllowUserResizing=   1
         FormatString    =   "       |^O.R. DATE     |^O.R. NO.    |>SPARE PARTS  |> SERVICE            |>OTHERS            |< REMARKS               "
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
      Begin MSFlexGridLib.MSFlexGrid flxPAYttl 
         Height          =   300
         Left            =   -69120
         TabIndex        =   77
         Tag             =   "XYF"
         Top             =   2740
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   529
         _Version        =   393216
         Rows            =   1
         Cols            =   3
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   300
         BackColorBkg    =   -2147483633
         ScrollBars      =   0
         BorderStyle     =   0
      End
      Begin lvButton.lvButtons_H cmdFindDMCM 
         Height          =   315
         Left            =   -65400
         TabIndex        =   80
         Tag             =   "XYF"
         Top             =   360
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         Caption         =   "Find"
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
         Image           =   "Ledger.frx":1F57
         cBack           =   -2147483633
      End
      Begin MSFlexGridLib.MSFlexGrid flxDMCM 
         Height          =   2055
         Left            =   -74880
         TabIndex        =   81
         Tag             =   "XYF"
         Top             =   675
         Width           =   12255
         _ExtentX        =   21616
         _ExtentY        =   3625
         _Version        =   393216
         Cols            =   7
         FixedCols       =   0
         RowHeightMin    =   280
         BackColorBkg    =   -2147483633
         SelectionMode   =   1
         AllowUserResizing=   1
         FormatString    =   "       |^REF DATE     |^REF. NO.    |<TYPE  |> DEBIT            |>CREDIT            |< REMARKS               "
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
      Begin MSFlexGridLib.MSFlexGrid flxDMCMttl 
         Height          =   300
         Left            =   -69120
         TabIndex        =   82
         Tag             =   "XYF"
         Top             =   2745
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   529
         _Version        =   393216
         Rows            =   1
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   300
         BackColorBkg    =   -2147483633
         ScrollBars      =   0
         BorderStyle     =   0
      End
      Begin VB.Label lblTtl 
         Alignment       =   1  'Right Justify
         Caption         =   "TOTAL"
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
         Left            =   -70800
         TabIndex        =   83
         Tag             =   "XYF"
         Top             =   2790
         Width           =   1455
      End
      Begin VB.Label lblTtl 
         Alignment       =   1  'Right Justify
         Caption         =   "TOTAL"
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
         Left            =   -70800
         TabIndex        =   78
         Tag             =   "XYF"
         Top             =   2785
         Width           =   1455
      End
      Begin VB.Label lblTtl 
         Alignment       =   1  'Right Justify
         Caption         =   "TOTAL"
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
         Left            =   -68400
         TabIndex        =   31
         Tag             =   "XYF"
         Top             =   2770
         Width           =   1455
      End
      Begin VB.Label lblTtl 
         Alignment       =   1  'Right Justify
         Caption         =   "TOTAL"
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
         Left            =   -71160
         TabIndex        =   25
         Tag             =   "XYF"
         Top             =   2785
         Width           =   1455
      End
      Begin VB.Label lblTtl 
         Alignment       =   1  'Right Justify
         Caption         =   "TOTAL"
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
         Left            =   3720
         TabIndex        =   19
         Tag             =   "XYF"
         Top             =   2785
         Width           =   1455
      End
   End
   Begin VB.ComboBox cboBillMon 
      Height          =   315
      Left            =   120
      TabIndex        =   13
      Top             =   2040
      Visible         =   0   'False
      Width           =   1695
   End
   Begin MSFlexGridLib.MSFlexGrid flxBill 
      Height          =   735
      Left            =   120
      TabIndex        =   12
      Tag             =   "XYF"
      Top             =   1690
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   1296
      _Version        =   393216
      Cols            =   6
      FixedCols       =   0
      RowHeightMin    =   315
      BackColorBkg    =   -2147483633
      ScrollBars      =   0
      FormatString    =   $"Ledger.frx":24F1
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
      Height          =   975
      Left            =   120
      TabIndex        =   4
      Tag             =   "XYF"
      Top             =   720
      Width           =   12495
      Begin VB.ComboBox cboCustomer 
         Height          =   315
         Left            =   2040
         TabIndex        =   56
         Tag             =   "XYF"
         Top             =   200
         Width           =   4215
      End
      Begin VB.TextBox txtEmpCo 
         Height          =   315
         Left            =   8760
         TabIndex        =   11
         Tag             =   "XYF"
         Top             =   200
         Width           =   3615
      End
      Begin VB.TextBox txtAddress 
         Height          =   315
         Left            =   2040
         TabIndex        =   10
         Tag             =   "xyf"
         Top             =   560
         Width           =   4215
      End
      Begin VB.Label lblGrp1 
         Alignment       =   1  'Right Justify
         Caption         =   "EMPLOYER:"
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
         TabIndex        =   9
         Tag             =   "XYF"
         Top             =   225
         Width           =   1455
      End
      Begin VB.Label lblGrp1 
         Alignment       =   1  'Right Justify
         Caption         =   "CUSTOMER ADDRESS:"
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
         Left            =   120
         TabIndex        =   8
         Tag             =   "XYF"
         Top             =   590
         Width           =   1815
      End
      Begin VB.Label lblGrp1 
         Alignment       =   1  'Right Justify
         Caption         =   "CUSTOMER NAME:"
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
         TabIndex        =   7
         Tag             =   "XYF"
         Top             =   230
         Width           =   1815
      End
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   661
      Caption         =   "&Summary"
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
      Image           =   "Ledger.frx":25B1
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   1
      Left            =   1800
      TabIndex        =   1
      Top             =   0
      Width           =   1800
      _ExtentX        =   3175
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
      Image           =   "Ledger.frx":2B4B
      cBack           =   -2147483633
   End
   Begin MSFlexGridLib.MSFlexGrid flxLedger 
      Height          =   3615
      Left            =   120
      TabIndex        =   33
      Tag             =   "XYF"
      Top             =   5760
      Width           =   6420
      _ExtentX        =   11324
      _ExtentY        =   6376
      _Version        =   393216
      Cols            =   6
      FixedCols       =   0
      BackColorBkg    =   -2147483633
      ScrollBars      =   2
      SelectionMode   =   1
      AllowUserResizing=   1
      FormatString    =   "^ DATE|^REF. NO. |^REF TYPE |>DEBIT |>CREDIT |>BALANCE"
   End
   Begin VB.Label lblHeading 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "CUSTOMER ACCOUNTS"
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
      Height          =   275
      Index           =   2
      Left            =   8040
      TabIndex        =   5
      Top             =   480
      Width           =   4500
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
      Height          =   275
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   2610
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   120
      Top             =   7800
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   12945
      m_MDI_Height    =   10125
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
      Height          =   275
      Index           =   1
      Left            =   150
      TabIndex        =   3
      Top             =   510
      Width           =   2610
   End
   Begin VB.Label lblHeading 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "CUSTOMER ACCOUNTS"
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
      Height          =   275
      Index           =   3
      Left            =   8055
      TabIndex        =   6
      Top             =   510
      Width           =   4500
   End
   Begin VB.Menu mnuPay 
      Caption         =   "Payments"
      Visible         =   0   'False
      Begin VB.Menu mnuAttach 
         Caption         =   "Show Attachments"
      End
   End
   Begin VB.Menu mnuPrn 
      Caption         =   "Print"
      Visible         =   0   'False
      Begin VB.Menu mnuPrnLdgr 
         Caption         =   "Print Ledger"
      End
   End
End
Attribute VB_Name = "Ledger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsAcctCust As New ADODB.Recordset
Dim CustSrch As Boolean
Dim ActRow As Integer, ActTab As Integer, fRow As Integer, fCol As Integer

Private Sub cboCustomer_Change()
    Dim i As Long, j As Long
    Dim strPartial As String, strTotal As String
    Dim X As Integer
    
    If CustSrch Then
        CustSrch = False
        Exit Sub
    End If
    With cboCustomer
        strPartial = .Text
        i = SendMessage(.hwnd, CB_FINDSTRING, -1, ByVal strPartial) 'If match found, append unmatched characters
        If i <> -1 Then
            strTotal = .List(i)
            rsAcctCust.MoveFirst
            rsAcctCust.Move i
            txtCustomerID.Text = rsAcctCust!CUSTOMERID
            txtAddress.Text = rsAcctCust!Address
            txtEmpCo.Text = IIf(IsNull(rsAcctCust!ContactPerson), "", rsAcctCust!ContactPerson)
            If rsAcctCust!Ctype = "INDIVIDUAL" Then
                lblGrp1(2).Caption = "EMPLOYER"
                For j = 0 To 5
                    flxBill.TextMatrix(1, j) = ""
                Next j
            Else
                lblGrp1(2).Caption = "MANAGER"
                Call DisplayBilling
            End If
            j = Len(strTotal) - Len(strPartial)
            If j <> 0 Then
                CustSrch = True
                .SelText = Right$(strTotal, j)
                .SelStart = Len(strPartial)
                .SelLength = j
            End If
        Else
            txtCustomerID.Text = ""
            txtAddress.Text = ""
            txtEmpCo.Text = ""
            lblGrp1(2).Caption = "EMPLOYER"
        End If
        flxLedger.Rows = 1
        flxLedger.Rows = 2
    End With
End Sub

Private Sub cboCustomer_Click()
    Call cboCustomer_Change
End Sub

Private Sub cboCustomer_GotFocus()
   Dim Tmp
    
   Const CB_SHOWDROPDOWN = &H14F
   Tmp = SendMessage(cboCustomer.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub

Private Sub cboCustomer_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then CustSrch = True
End Sub

Private Sub cboCustomer_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtAddress.SetFocus
End Sub

Private Sub cboCustomer_LostFocus()
    If Val(txtCustomerID.Text) > 0 Then
        Call DisplayPayment
        Call DisplayDMCM
        cboOtherAcct.ListIndex = 0
        Call DisplayOpenPCV
        Call ComputeFlxPCVttl
        cboSVCacct.ListIndex = 0
        Call DisplayOpenSVC
        Call ComputeFlxSVCttl
        cboSPacct.ListIndex = 0
        Call DisplayOpenSP
        Call ComputeFlxSPttl
        tabAccts.Tab = 0
    End If
End Sub


Private Sub DisplayCloseSVC()
    Dim rsAccts As New ADODB.Recordset
    Dim Qry As String, MyTerm As String, MyStat As String
    Dim sRow As Integer, eRow As Integer
                                                         
    'Display SP Accts
    'Qry = "SELECT ONACCTERMS.TRANSDATE, ONACCTERMS.SPINVNO, ONACCTERMS.NPAYS, ONACCTERMS.PAYSCHEME, ONACCTERMS.ACCTAMOUNT, ONACCTERMS.BALANCE, ONACCTERMS.TTLAMTPAID, ONACCTERMS.AMTDUE, ONACCTERMS.STATUS, ONACCTERMS.JOBORDERNO From ONACCTERMS " & _
          "WHERE (((ONACCTERMS.SPINVNO) Is Null) AND ((ONACCTERMS.STATUS)='CLOSED') AND (Not (ONACCTERMS.JOBORDERNO) Is Null) AND ((ONACCTERMS.CUSTOMERID)=" & txtCustomerID.Text & ")) ORDER BY ONACCTERMS.TRANSDATE;"
    'mySQL
    Qry = "SELECT ONACCTERMS.TRANSDATE, ONACCTERMS.SPINVNO, ONACCTERMS.NPAYS, ONACCTERMS.PAYSCHEME, ONACCTERMS.ACCTAMOUNT, ONACCTERMS.BALANCE, ONACCTERMS.TTLAMTPAID, ONACCTERMS.AMTDUE, ONACCTERMS.STATUS, ONACCTERMS.JOBORDERNO From ONACCTERMS " & _
          "WHERE (((ONACCTERMS.SPINVNO) Is Null) AND ((ONACCTERMS.STATUS)='CLOSED') AND (Not (ONACCTERMS.JOBORDERNO) Is Null) AND ((ONACCTERMS.CUSTOMERID)=" & txtCustomerID.Text & ")) ORDER BY ONACCTERMS.TRANSDATE;"
    
    With rsAccts
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            If flxSVC.TextMatrix(1, 1) = "" Or cboSVCacct.Text = "CLOSE ACCOUNTS" Then
                flxSVC.Rows = 1
                sRow = 1
                eRow = 0
            Else
                sRow = IIf(flxSP.Rows = 2, 2, flxSP.Rows - 1)
                eRow = flxSVC.Rows - 1
            End If
            If !PayScheme = "WEEKLY" Then
                MyTerm = IIf(!Npays > 1, !Npays & " WKS", !Npays & " WK")
            ElseIf !PayScheme = "MONTHLY" Then
                MyTerm = IIf(!Npays > 1, !Npays & " MOS", !Npays & " MON")
            Else
                MyTerm = IIf(!Npays > 1, !Npays & " BI-MOS", !Npays & "BI-MON")
            End If
            Do While Not .EOF
                flxSVC.AddItem vbTab & Format(!TRANSDATE, "mm/dd/yy") & vbTab & !JOBORDERNO & vbTab & Format(!AcctAmount, "#,##0.00") & vbTab & MyTerm & vbTab & Format(!Balance, "#,##0.00") & vbTab & Format(!TTLAMTPAID, "#,##0.00") & vbTab & "0.00" & vbTab & "0.00" & vbTab & "0.00" & vbTab & "C"
                eRow = eRow + 1
                .MoveNext
            Loop
            flxSVC.Row = sRow: flxSVC.Col = 0: flxSVC.RowSel = eRow: flxSVC.ColSel = flxSVC.Cols - 1: flxSVC.FillStyle = flexFillRepeat: flxSVC.CellForeColor = QBColor(8)
            flxSVC.Col = 1
        Else
            If cboSVCacct.Text = "CLOSE ACCOUNTS" Then
                flxSVC.Rows = 1
                flxSVC.Rows = 2
            End If
        End If
        .Close
    End With
End Sub

Private Sub DisplayClosePCV()
    Dim rsPCV As New ADODB.Recordset
    Dim Qry As String
    Dim eRow As Integer, sRow As Integer
    
    'Qry = "SELECT DISBURSEMENT.CVDATE, DISBURSEMENT.CVNUMBER, CUSTOMER.CUSTOMERNAME, DISBURSEMENT.PARTICULARS, DISBURSEMENT.AMOUNT, ONACCTPCV.BALANCE, ONACCTPCV.CUSTOMERID, DISBURSEMENT.DISBURSETYPE " & _
          "FROM DISBURSEMENT INNER JOIN (ONACCTPCV INNER JOIN CUSTOMER ON ONACCTPCV.CUSTOMERID = CUSTOMER.CUSTOMERID) ON DISBURSEMENT.VOUCHERID = ONACCTPCV.VOUCHERID " & _
          "WHERE (((ONACCTPCV.BALANCE)=0) AND ((ONACCTPCV.CUSTOMERID)=" & txtCustomerID.Text & ") AND ((DISBURSEMENT.DISBURSETYPE)<>'Cancelled Vouchers'));"
    'mySQL
    Qry = "SELECT DISBURSEMENT.CVDATE, DISBURSEMENT.CVNUMBER, CUSTOMER.CUSTOMERNAME, DISBURSEMENT.PARTICULARS, DISBURSEMENT.AMOUNT, ONACCTPCV.BALANCE, ONACCTPCV.CUSTOMERID, DISBURSEMENT.DISBURSETYPE " & _
          "FROM DISBURSEMENT INNER JOIN (ONACCTPCV INNER JOIN CUSTOMER ON ONACCTPCV.CUSTOMERID = CUSTOMER.CUSTOMERID) ON DISBURSEMENT.VOUCHERID = ONACCTPCV.VOUCHERID " & _
          "WHERE (((ONACCTPCV.BALANCE)=0) AND ((ONACCTPCV.CUSTOMERID)=" & txtCustomerID.Text & ") AND ((DISBURSEMENT.DISBURSETYPE)<>'Cancelled Vouchers'));"
    
    With rsPCV
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            If flxOthers.TextMatrix(1, 1) = "" Or cboOtherAcct.Text = "CLOSE ACCOUNTS" Then
                flxOthers.Rows = 1
                sRow = 1
                eRow = 0
            Else
                sRow = IIf(flxOthers.Rows = 2, 2, flxOthers.Rows - 1)
                eRow = flxOthers.Rows - 1
            End If
            Do While Not .EOF
                flxOthers.AddItem vbTab & Format(!CVDate, "mm/dd/yyyy") & vbTab & Format(!CVNUMBER, "0000") & vbTab & !CustomerName & vbTab & !Particulars & vbTab & Format(!Amount, "#,##0.00") & vbTab & "CLOSE"
                eRow = eRow + 1
                .MoveNext
            Loop
            flxOthers.Row = sRow: flxOthers.Col = 0: flxOthers.RowSel = eRow: flxOthers.ColSel = flxOthers.Cols - 1: flxOthers.FillStyle = flexFillRepeat: flxOthers.CellForeColor = QBColor(8)
            flxOthers.Col = 1
        Else
            flxOthers.Rows = 1
            flxOthers.Rows = 2
        End If
        .Close
    End With
End Sub

Private Sub DisplayOpenPCV()
    Dim rsPCV As New ADODB.Recordset
    Dim Qry As String
    
    Qry = "SELECT DISBURSEMENT.CVDATE, DISBURSEMENT.CVNUMBER, CUSTOMER.CUSTOMERNAME, DISBURSEMENT.PARTICULARS, DISBURSEMENT.AMOUNT, ONACCTPCV.BALANCE, ONACCTPCV.CUSTOMERID " & _
          "FROM DISBURSEMENT INNER JOIN (ONACCTPCV INNER JOIN CUSTOMER ON ONACCTPCV.CUSTOMERID = CUSTOMER.CUSTOMERID) ON DISBURSEMENT.VOUCHERID = ONACCTPCV.VOUCHERID " & _
          "WHERE (((ONACCTPCV.BALANCE)>0) AND ((ONACCTPCV.CUSTOMERID)=" & txtCustomerID.Text & "));"
       
    With rsPCV
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            flxOthers.Rows = 1
            Do While Not .EOF
                flxOthers.AddItem vbTab & Format(!CVDate, "mm/dd/yyyy") & vbTab & Format(!CVNUMBER, "0000") & vbTab & !CustomerName & vbTab & !Particulars & vbTab & Format(!Amount, "#,##0.00") & vbTab & "OPEN"
                .MoveNext
            Loop
            flxOthers.Row = 1: flxOthers.Col = 0: flxOthers.RowSel = flxOthers.Rows - 1: flxOthers.ColSel = flxOthers.Cols - 1: flxOthers.FillStyle = flexFillRepeat: flxOthers.CellForeColor = vbBlack
            flxOthers.Col = 1
        Else
            flxOthers.Rows = 1
            flxOthers.Rows = 2
        End If
        .Close
    End With
        
End Sub

Private Sub DisplayOpenSVC()
    Dim rsAccts As New ADODB.Recordset
    Dim Qry As String, MyTerm As String, MyStat As String
    Dim AmtDue As Currency, OverDue As Currency, Payable As Currency, TmpDue As Currency
    Dim myBillDate As Date, myBN As Integer, TmpOD As Currency
    
    'Display SP Accts
    'Qry = "SELECT ONACCTERMS.TRANSDATE, ONACCTERMS.SPINVNO, ONACCTERMS.NPAYS, ONACCTERMS.PAYSCHEME, ONACCTERMS.ACCTAMOUNT, ONACCTERMS.BALANCE, ONACCTERMS.TTLAMTPAID, ONACCTERMS.AMTDUE, ONACCTERMS.STATUS, ONACCTERMS.PAYSTART, ONACCTERMS.TERMS, ONACCTERMS.JOBORDERNO From ONACCTERMS " & _
          "WHERE (((ONACCTERMS.SPINVNO) Is Null) AND ((ONACCTERMS.STATUS)='OPEN') AND ((ONACCTERMS.CUSTOMERID)=" & txtCustomerID.Text & ") AND (Not (ONACCTERMS.JOBORDERNO) Is Null)) ORDER BY ONACCTERMS.TRANSDATE;"
    'mySQL
    Qry = "SELECT ONACCTERMS.TRANSDATE, ONACCTERMS.SPINVNO, ONACCTERMS.NPAYS, ONACCTERMS.PAYSCHEME, ONACCTERMS.ACCTAMOUNT, ONACCTERMS.BALANCE, ONACCTERMS.TTLAMTPAID, ONACCTERMS.AMTDUE, ONACCTERMS.STATUS, ONACCTERMS.PAYSTART, ONACCTERMS.TERMS, ONACCTERMS.JOBORDERNO From ONACCTERMS " & _
          "WHERE (((ONACCTERMS.SPINVNO) Is Null) AND ((ONACCTERMS.STATUS)='OPEN') AND ((ONACCTERMS.CUSTOMERID)=" & txtCustomerID.Text & ") AND (Not (ONACCTERMS.JOBORDERNO) Is Null)) ORDER BY ONACCTERMS.TRANSDATE;"
    
    With rsAccts
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            flxSVC.Rows = 1
            Do While Not .EOF
                TmpDue = !AcctAmount / !Npays
                
                If !PayStart >= Date Then  'wla pa naabot ang Start of Payment date
                    OverDue = 0 'Over Due
                    If !Balance < TmpDue Then     'if Balance<Billing Amt
                        AmtDue = !Balance 'Balance
                    Else
                        AmtDue = TmpDue 'Billing Amt
                    End If
                Else
                    'Process per payment Scheme
                    myBillDate = !PayStart
                    myBN = 0 'Bill no
                    Do While myBillDate <= Date And myBN <= !Npays   'Loop while billing date <current date and Billing No < #payments
                        If !PayScheme = "WEEKLY" Then
                            If Day(myBillDate) = 21 Then
                                myBillDate = DateAdd("d", -1, DateSerial(Year(DateAdd("m", 1, myBillDate)), Month(DateAdd("m", 1, myBillDate)), 1))
                            Else
                                myBillDate = DateAdd("d", 7, myBillDate)
                            End If
                        ElseIf !PayScheme = "BI-MONTHLY" Then
                            myBillDate = DateAdd("d", 15, myBillDate)
                        ElseIf !PayScheme = "MONTHLY" Then
                            myBillDate = DateAdd("m", 1, myBillDate)
                        End If
                        myBN = myBN + 1
                    Loop           'No. of Payments
                    If myBN >= !Npays Then 'Billing = NPays
                        AmtDue = 0
                        OverDue = !Balance 'Balance
                    Else
                        If !Balance < TmpDue Then     'if Balance<Billing Amt
                            AmtDue = !Balance 'Balance
                        Else
                            AmtDue = TmpDue 'Billing Amt
                        End If
                        TmpOD = ((myBN - 1) * TmpDue) - !TTLAMTPAID 'Exact Amt upto Billing less Total Amt. Paid
                        If TmpOD > 0 Then
                            If TmpOD > !Balance - AmtDue Then  'if temp OverDue > (Balance-BillAmt)
                                OverDue = !Balance - TmpOD
                            Else
                                OverDue = TmpOD
                            End If
                        Else
                            OverDue = 0
                        End If
                    End If
                End If
                                                
                Payable = AmtDue + OverDue
                If !PayScheme = "WEEKLY" Then
                    MyTerm = IIf(!Npays > 1, !Npays & " WKS", !Npays & " WK")
                ElseIf !PayScheme = "MONTHLY" Then
                    MyTerm = IIf(!Npays > 1, !Npays & " MOS", !Npays & " MON")
                Else
                    MyTerm = IIf(!Npays > 1, !Npays & " BI-MOS", !Npays & "BI-MON")
                End If
                flxSVC.AddItem vbTab & Format(!TRANSDATE, "mm/dd/yy") & vbTab & !JOBORDERNO & vbTab & Format(!AcctAmount, "#,##0.00") & vbTab & MyTerm & vbTab & Format(!Balance, "#,##0.00") & vbTab & Format(!TTLAMTPAID, "#,##0.00") & vbTab & Format(AmtDue, "#,##0.00") & vbTab & Format(OverDue, "#,##0.00") & vbTab & Format(Payable, "#,##0.00") & vbTab & "OPEN"
                .MoveNext
            Loop
            flxSVC.Row = 1: flxSVC.Col = 0: flxSVC.RowSel = flxSVC.Rows - 1: flxSVC.ColSel = flxSVC.Cols - 1: flxSVC.FillStyle = flexFillRepeat: flxSVC.CellForeColor = vbBlack
            flxSVC.Col = 1
        Else
            flxSVC.Rows = 1
            flxSVC.Rows = 2
        End If
        .Close
    End With
End Sub

Private Sub DisplayOpenSP()
    Dim rsAccts As New ADODB.Recordset
    Dim Qry As String, MyTerm As String, MyStat As String
    Dim AmtDue As Currency, OverDue As Currency, Payable As Currency, TmpDue As Currency
    Dim myBillDate As Date, myBN As Integer, TmpOD As Currency
    
    'Display SP Accts
    'Qry = "SELECT ONACCTERMS.TRANSDATE, ONACCTERMS.SPINVNO, ONACCTERMS.NPAYS, ONACCTERMS.PAYSCHEME, ONACCTERMS.ACCTAMOUNT, ONACCTERMS.BALANCE, ONACCTERMS.TTLAMTPAID, ONACCTERMS.AMTDUE, ONACCTERMS.STATUS, ONACCTERMS.PAYSTART, ONACCTERMS.TERMS From ONACCTERMS " & _
          "WHERE ((Not (ONACCTERMS.SPINVNO) Is Null) AND ((ONACCTERMS.STATUS)='OPEN') AND ((ONACCTERMS.CUSTOMERID)=" & txtCustomerID.Text & ")) ORDER BY ONACCTERMS.TRANSDATE;"
    'mySQL
    Qry = "SELECT ONACCTERMS.TRANSDATE, ONACCTERMS.SPINVNO, ONACCTERMS.NPAYS, ONACCTERMS.PAYSCHEME, ONACCTERMS.ACCTAMOUNT, ONACCTERMS.BALANCE, ONACCTERMS.TTLAMTPAID, ONACCTERMS.AMTDUE, ONACCTERMS.STATUS, ONACCTERMS.PAYSTART, ONACCTERMS.TERMS From ONACCTERMS " & _
          "WHERE ((Not (ONACCTERMS.SPINVNO) Is Null) AND ((ONACCTERMS.STATUS)='OPEN') AND ((ONACCTERMS.CUSTOMERID)=" & txtCustomerID.Text & ")) ORDER BY ONACCTERMS.TRANSDATE;"
    
    With rsAccts
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            flxSP.Rows = 1
            Do While Not .EOF
                                 
                If !Npays = 0 Then
                    TmpDue = !AcctAmount
                Else
                    TmpDue = !AcctAmount / !Npays
                End If
                
                If !PayStart >= Date Then  'wla pa naabot ang Start of Payment date
                    OverDue = 0 'Over Due
                    If !Balance < TmpDue Then     'if Balance<Billing Amt
                        AmtDue = !Balance 'Balance
                    Else
                        AmtDue = TmpDue 'Billing Amt
                    End If
                Else
                    'Process per payment Scheme
                    myBillDate = !PayStart
                    myBN = 0 'Bill no
                    Do While myBillDate <= Date And myBN <= !Npays   'Loop while billing date <current date and Billing No < #payments
                        If !PayScheme = "WEEKLY" Then
                            If Day(myBillDate) = 21 Then
                                myBillDate = DateAdd("d", -1, DateSerial(Year(DateAdd("m", 1, myBillDate)), Month(DateAdd("m", 1, myBillDate)), 1))
                            Else
                                myBillDate = DateAdd("d", 7, myBillDate)
                            End If
                        ElseIf !PayScheme = "BI-MONTHLY" Then
                            myBillDate = DateAdd("d", 15, myBillDate)
                        ElseIf !PayScheme = "MONTHLY" Then
                            myBillDate = DateAdd("m", 1, myBillDate)
                        End If
                        myBN = myBN + 1
                    Loop           'No. of Payments
                    If myBN >= !Npays Then 'Billing = NPays
                        AmtDue = 0
                        OverDue = !Balance 'Balance
                    Else
                        If !Balance < TmpDue Then     'if Balance<Billing Amt
                            AmtDue = !Balance 'Balance
                        Else
                            AmtDue = TmpDue 'Billing Amt
                        End If
                        TmpOD = ((myBN - 1) * TmpDue) - !TTLAMTPAID 'Exact Amt upto Billing less Total Amt. Paid
                        If TmpOD > 0 Then
                            If TmpOD > !Balance - AmtDue Then  'if temp OverDue > (Balance-BillAmt)
                                OverDue = !Balance - TmpOD
                            Else
                                OverDue = TmpOD
                            End If
                        Else
                            OverDue = 0
                        End If
                    End If
                End If
                
                
                Payable = AmtDue + OverDue
                'If !PayScheme = "WEEKLY" Then
                '    MyTerm = IIf(!Npays > 1, !Npays & " WKS", !Npays & " WK")
                'ElseIf !PayScheme = "MONTHLY" Then
                '    MyTerm = IIf(!Npays > 1, !Npays & " MOS", !Npays & " MON")
                'Else
                '    MyTerm = IIf(!Npays > 1, !Npays & " BI-MOS", !Npays & "BI-MON")
                'End If
                
               
                flxSP.AddItem vbTab & Format(!TRANSDATE, "mm/dd/yy") & vbTab & !SPINVNO & vbTab & Format(!AcctAmount, "#,##0.00") & vbTab & !Terms & vbTab & Format(!Balance, "#,##0.00") & vbTab & Format(!TTLAMTPAID, "#,##0.00") & vbTab & Format(AmtDue, "#,##0.00") & vbTab & Format(OverDue, "#,##0.00") & vbTab & Format(Payable, "#,##0.00") & vbTab & "O"
                .MoveNext
            Loop
            flxSP.Row = 1: flxSP.Col = 0: flxSP.RowSel = flxSP.Rows - 1: flxSP.ColSel = flxSP.Cols - 1: flxSP.FillStyle = flexFillRepeat: flxSP.ForeColor = vbBlack
            flxSP.Col = 1
        Else
            flxSP.Rows = 1
            flxSP.Rows = 2
        End If
        .Close
    End With
End Sub

Private Sub DisplayCloseSP()
    Dim rsAccts As New ADODB.Recordset
    Dim Qry As String, MyTerm As String, MyStat As String
    Dim sRow As Integer, eRow As Integer
                                                         
    'Display SP Accts
    'Qry = "SELECT ONACCTERMS.TRANSDATE, ONACCTERMS.SPINVNO, ONACCTERMS.NPAYS, ONACCTERMS.PAYSCHEME, ONACCTERMS.ACCTAMOUNT, ONACCTERMS.BALANCE, ONACCTERMS.TTLAMTPAID, ONACCTERMS.AMTDUE, ONACCTERMS.STATUS From ONACCTERMS " & _
          "WHERE ((Not (ONACCTERMS.SPINVNO) Is Null) AND ((ONACCTERMS.STATUS)='CLOSED') AND ((ONACCTERMS.CUSTOMERID)=" & txtCustomerID.Text & ")) ORDER BY ONACCTERMS.TRANSDATE;"
    'mySQL
    Qry = "SELECT ONACCTERMS.TRANSDATE, ONACCTERMS.SPINVNO, ONACCTERMS.NPAYS, ONACCTERMS.PAYSCHEME, ONACCTERMS.ACCTAMOUNT, ONACCTERMS.BALANCE, ONACCTERMS.TTLAMTPAID, ONACCTERMS.AMTDUE, ONACCTERMS.STATUS From ONACCTERMS " & _
          "WHERE ((Not (ONACCTERMS.SPINVNO) Is Null) AND ((ONACCTERMS.STATUS)='CLOSED') AND ((ONACCTERMS.CUSTOMERID)=" & txtCustomerID.Text & ")) ORDER BY ONACCTERMS.TRANSDATE;"
    
    With rsAccts
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            If flxSP.TextMatrix(1, 1) = "" Or cboSPacct.Text = "CLOSE ACCOUNTS" Then
                flxSP.Rows = 1
                sRow = 1
                eRow = 0
            Else
                sRow = IIf(flxSP.Rows = 2, 2, flxSP.Rows - 1)
                eRow = flxSP.Rows - 1
            End If
            
            'old Code
            'If !PayScheme = "WEEKLY" Then
            '    MyTerm = IIf(!Npays > 1, !Npays & " WKS", !Npays & " WK")
            'ElseIf !PayScheme = "MONTHLY" Then
            '    MyTerm = IIf(!Npays > 1, !Npays & " MOS", !Npays & " MON")
            'Else
            '    MyTerm = IIf(!Npays > 1, !Npays & " BI-MOS", !Npays & "BI-MON")
            'End If
            
            Do While Not .EOF
                
                If !PayScheme = "WEEKLY" Then
                    MyTerm = IIf(!Npays > 1, !Npays & " WKS", !Npays & " WK")
                ElseIf !PayScheme = "MONTHLY" Then
                    MyTerm = IIf(!Npays > 1, !Npays & " MOS", !Npays & " MON")
                Else
                    MyTerm = IIf(!Npays > 1, !Npays & " BI-MOS", !Npays & "BI-MON")
                End If
                '                                 date                             invno                    invAmnt                           terms
                flxSP.AddItem vbTab & Format(!TRANSDATE, "mm/dd/yy") & vbTab & !SPINVNO & vbTab & Format(!AcctAmount, "#,##0.00") & vbTab & MyTerm & vbTab & Format(!Balance, "#,##0.00") & vbTab & Format(!TTLAMTPAID, "#,##0.00") & vbTab & "0.00" & vbTab & "0.00" & vbTab & "0.00" & vbTab & "C"
                eRow = eRow + 1
                .MoveNext
            
            Loop
            
            flxSP.Row = sRow: flxSP.Col = 0: flxSP.RowSel = eRow: flxSP.ColSel = flxSP.Cols - 1: flxSP.FillStyle = flexFillRepeat: flxSP.CellForeColor = QBColor(8)
            flxSP.Col = 1
        Else
            If cboSPacct.Text = "CLOSE ACCOUNTS" Then
                flxSP.Rows = 1
                flxSP.Rows = 2
            End If
        End If
        .Close
    End With
End Sub

Private Sub DisplayBilling()
    Dim rsBill As New ADODB.Recordset
    Dim Qry As String
    
    'Qry = "SELECT BILLING.* From BILLING Where (((BILLING.CUSTOMERID) = " & txtCustomerID.Text & ")) ORDER BY BILLING.BILLYR DESC , BILLING.BILLMO DESC;"
    'mySQL
    Qry = "SELECT BILLING.* From BILLING Where BILLING.CUSTOMERID = " & txtCustomerID.Text & " ORDER BY BILLING.BILLYR DESC , BILLING.BILLMO DESC;"
    
    With rsBill
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            flxBill.TextMatrix(1, 0) = !DUEDATE
            flxBill.TextMatrix(1, 1) = Format(!PrevBal, "#,##0.00")
            flxBill.TextMatrix(1, 2) = Format(!CurrentBill, "#,##0.00")
            flxBill.TextMatrix(1, 3) = Format(!Payments, "#,##0.00")
            flxBill.TextMatrix(1, 4) = Format(!DmCm, "#,##0.00")
            flxBill.TextMatrix(1, 5) = Format(!Balance, "#,##0.00")
        Else
            For j = 0 To 5
                flxBill.TextMatrix(1, j) = ""
            Next j
        End If
        .Close
    End With
End Sub

Private Sub DisplayDMCM()
    Dim rsDmCm As New ADODB.Recordset
    Dim Qry As String, CntRec As Integer
    Dim TtlDM As Currency, TtlCM As Currency
    
    'Display DMCM type of Adjustment
    'Qry = "SELECT ADJDMCM.ADJDATE, ADJDMCM.ADJNO, ADJDMCM.TRANSTYPE, ADJDMCM.DEBIT, ADJDMCM.CREDIT, ADJDMCM.REASON From ADJDMCM " & _
          "WHERE (((ADJDMCM.REASON)<>'CANCELLED') AND ((ADJDMCM.CUSTOMERID)=" & txtCustomerID.Text & ")) ORDER BY ADJDMCM.ADJDATE DESC;"
    'mySQL
    Qry = "SELECT ADJDMCM.ADJDATE, ADJDMCM.ADJNO, ADJDMCM.TRANSTYPE, ADJDMCM.DEBIT, ADJDMCM.CREDIT, ADJDMCM.REASON From ADJDMCM " & _
          "WHERE (((ADJDMCM.REASON)<>'CANCELLED') AND ((ADJDMCM.CUSTOMERID)=" & txtCustomerID.Text & ")) ORDER BY ADJDMCM.ADJDATE DESC;"
    
    With rsDmCm
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            flxDMCM.Rows = 1
            TtlDM = 0: TtlCM = 0: CntRec = 0
            Do While Not .EOF
                flxDMCM.AddItem vbTab & !ADJDATE & vbTab & !ADJNO & vbTab & "DMCM [" & !TRANSTYPE & "]" & vbTab & Format(!Debit, "#,##0.00") & vbTab & Format(!Credit, "#,##0.00") & vbTab & !REASON
                TtlCM = TtlCM + !Credit
                TtlDM = TtlDM + !Debit
                CntRec = CntRec + 1
                .MoveNext
            Loop
        Else
            flxDMCM.Rows = 1
            flxDMCM.Rows = 2
            CntRec = 0
        End If
        .Close
    End With
    
    'Display Sales Return Type of Adjustment
    'Qry = "SELECT SRETURNSP.RETDATE, SRETURNSP.RETNO, SRETURNSP.TOTALAMT, SRETURNSP.REASON FROM SRETURNSP INNER JOIN SPINVOICE ON SRETURNSP.INVOICE = SPINVOICE.SPINVNO " & _
          "WHERE (((SRETURNSP.REMARKS)<>'CANCELLED') AND ((SPINVOICE.CUSTOMERID)=" & txtCustomerID.Text & ")) ORDER BY SRETURNSP.RETDATE DESC;"
    'mySQL
    Qry = "SELECT SRETURNSP.RETDATE, SRETURNSP.RETNO, SRETURNSP.TOTALAMT, SRETURNSP.REASON FROM SRETURNSP INNER JOIN SPINVOICE ON SRETURNSP.INVOICE = SPINVOICE.SPINVNO " & _
          "WHERE (((SRETURNSP.REMARKS)<>'CANCELLED') AND ((SPINVOICE.CUSTOMERID)=" & txtCustomerID.Text & ")) ORDER BY SRETURNSP.RETDATE DESC;"
    
    With rsDmCm
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            If CntRec = 0 Then flxDMCM.Rows = 1
            Do While Not .EOF
                flxDMCM.AddItem vbTab & !RetDate & vbTab & !RetNo & vbTab & "SALES RETURN" & vbTab & vbTab & Format(!TotalAmt, "#,##0.00") & vbTab & !REASON
                TtlCM = TtlCM + !TotalAmt
                .MoveNext
            Loop
        End If
        .Close
    End With
    flxDMCMttl.TextMatrix(0, 0) = Format(TtlDM, "#,##0.00")
    flxDMCMttl.TextMatrix(0, 1) = Format(TtlCM, "#,##0.00")
    
End Sub

Private Sub DisplayPayment()
    Dim rsPay As New ADODB.Recordset
    Dim Qry As String
    Dim TtlSP As Currency, TtlSvc As Currency, TtlOthers As Currency
    
    'Qry = "SELECT PAYMENTS.* From PAYMENTS WHERE (((PAYMENTS.CUSTOMERID)=" & txtCustomerID.Text & ") AND ((PAYMENTS.REMARKS)<>'CANCELLED')) ORDER BY PAYMENTS.PAYDATE DESC;"
    'mySQl
    Qry = "SELECT PAYMENTS.* From PAYMENTS WHERE PAYMENTS.CUSTOMERID=" & txtCustomerID.Text & " AND PAYMENTS.REMARKS<>'CANCELLED' ORDER BY PAYMENTS.PAYDATE DESC;"
    
    With rsPay
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            FlxPay.Rows = 1
            TtlSP = 0: TtlSvc = 0: TtlOthers = 0
            Do While Not .EOF
                FlxPay.AddItem vbTab & !PayDate & vbTab & IIf(IsNull(!ORNO), !CRNO, !ORNO) & vbTab & Format(!SPAREPARTS, "#,##0.00") & vbTab & Format(!Service, "#,##0.00") & vbTab & Format(!Others, "#,##0.00") & vbTab & IIf(IsNull(!Remarks), "", !Remarks)
                TtlSP = TtlSP + IIf(IsNull(!SPAREPARTS), 0, !SPAREPARTS)
                TtlSvc = TtlSvc + IIf(IsNull(!Service), 0, !Service)
                TtlOthers = TtlOthers + IIf(IsNull(!Others), 0, !Others) + IIf(IsNull(!DownPayment), 0, !DownPayment)
                .MoveNext
            Loop
            flxPAYttl.TextMatrix(0, 0) = Format(TtlSP, "#,##0.00")
            flxPAYttl.TextMatrix(0, 1) = Format(TtlSvc, "#,##0.00")
            flxPAYttl.TextMatrix(0, 2) = Format(TtlOthers, "#,##0.00")
        Else
            FlxPay.Rows = 1
            FlxPay.Rows = 2
        End If
        .Close
    End With
End Sub

Private Sub DisplayLedger()
    Dim rsLedger As New ADODB.Recordset
    Dim Qry As String
    
    With rsLedger
        If tabAccts.Tab = 0 Then
            'mySQL
            Qry = "SELECT LEDGER.REFDATE, LEDGER.REFNO, LEDGER.REFTYPE, LEDGER.DEBIT, LEDGER.CREDIT, LEDGER.OSBAL From LEDGER " & _
                  "Where LEDGER.SPINVNO = '" & flxSP.TextMatrix(ActRow, 2) & "' And LEDGER.CUSTOMERID = " & txtCustomerID.Text & " And LEDGER.REFTYPE != 'CANCELLED' ORDER BY LEDGER.SEQNO;"
            
        Else
            'mySQL
            Qry = "SELECT LEDGER.REFDATE, LEDGER.REFNO, LEDGER.REFTYPE, LEDGER.DEBIT, LEDGER.CREDIT, LEDGER.OSBAL From LEDGER " & _
                  "WHERE (((LEDGER.SPINVNO)='" & flxSVC.TextMatrix(ActRow, 2) & "') AND ((LEDGER.CUSTOMERID)=" & txtCustomerID.Text & ") AND ((LEDGER.JOBID)>0)) And LEDGER.REFTYPE != 'CANCELLED' ORDER BY LEDGER.SEQNO;"
            
        End If
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            flxLedger.Rows = 1
            Do While Not .EOF
                flxLedger.AddItem !REFDATE & vbTab & !RefNo & vbTab & !RefType & vbTab & IIf(!Debit <> 0, Format$(!Debit, "#,##0.00"), "") & vbTab & IIf(!Credit <> 0, Format$(!Credit, "#,##0.00"), "") & vbTab & Format$(!OSbal, "#,##0.00")
                .MoveNext
            Loop
        Else
            flxLedger.Rows = 1
            flxLedger.Rows = 2
        End If
        .Close
    End With
End Sub

Private Sub DisplayPCV()
    Dim Dtls As String, Qry As String
    Dim Details As Variant
    With flxOthers
        lblPcvNo.Caption = .TextMatrix(ActRow, 2)
        lblPcvDate.Caption = .TextMatrix(ActRow, 1)
        lblPayee.Caption = .TextMatrix(ActRow, 3)
        'Qry = "SELECT DISBURSEMENT.ACCTDESC From Disbursement WHERE (((DISBURSEMENT.CVNUMBER)=" & Val(.TextMatrix(ActRow, 2)) & "));"
        'mySQL
        Qry = "SELECT DISBURSEMENT.ACCTDESC From Disbursement WHERE DISBURSEMENT.CVNUMBER=" & Val(.TextMatrix(ActRow, 2)) & ";"
        
        lblATitle.Caption = GetFieldValue(Qry)
        'Qry = "SELECT DISBURSEMENT.DETAILS From Disbursement WHERE (((DISBURSEMENT.CVNUMBER)=" & Val(.TextMatrix(ActRow, 2)) & "));"
        'mySQL
        Qry = "SELECT DISBURSEMENT.DETAILS From Disbursement WHERE DISBURSEMENT.CVNUMBER=" & Val(.TextMatrix(ActRow, 2)) & ";"
        
        Dtls = GetFieldValue(Qry)
        flxPCV.Rows = 1
        flxPCV.AddItem vbTab & .TextMatrix(ActRow, 4) & vbTab & .TextMatrix(ActRow, 5)
        If Len(Dtls) > 1 Then
            Details = Split(Dtls, "|")
            For i = LBound(Details) To UBound(Details)
                flxPCV.AddItem vbTab & Details(i)
            Next i
        End If
        txtTtlPCV.Text = .TextMatrix(ActRow, 5)
    End With
    tabDocs.Tab = 2
End Sub

Private Sub DisplayJobOrder()
    Dim rsJob As New ADODB.Recordset
    Dim rsDetails As New ADODB.Recordset
    Dim Qry As String, myDisc As Currency, MyLabor As Currency, myTtl As Currency
    Dim MyJOid As Long, GTtl As Currency, OAdisc As Currency
    
    'Qry = "SELECT SRVCJOBORDER.JOBORDERNO, SRVCJOBORDER.DATERECEIVED, SRVCJOBORDER.JOTYPE, SRVCUSTOMER.MODEL, SRVCUSTOMER.ENGINE, SRVCUSTOMER.CHASSIS, SRVCJOBORDER.CUSTOMER, SRVCJOBORDER.JOBORDERID, SRVCJOBORDER.DISCOUNT " & _
          "FROM SRVCJOBORDER INNER JOIN SRVCUSTOMER ON SRVCJOBORDER.ACCOUNTNO = SRVCUSTOMER.ACCOUNTNO WHERE (((SRVCJOBORDER.JOBORDERNO)='" & flxSVC.TextMatrix(ActRow, 2) & "'));"
    'mySQl
    Qry = "SELECT SRVCJOBORDER.JOBORDERNO, SRVCJOBORDER.DATERECEIVED, SRVCJOBORDER.JOTYPE, SRVCUSTOMER.MODEL, SRVCUSTOMER.ENGINE, SRVCUSTOMER.CHASSIS, SRVCJOBORDER.CUSTOMER, SRVCJOBORDER.JOBORDERID, SRVCJOBORDER.DISCOUNT " & _
          "FROM SRVCJOBORDER INNER JOIN SRVCUSTOMER ON SRVCJOBORDER.ACCOUNTNO = SRVCUSTOMER.ACCOUNTNO WHERE SRVCJOBORDER.JOBORDERNO='" & flxSVC.TextMatrix(ActRow, 2) & "';"
    
    With rsJob
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            lblJobNo.Caption = !JOBORDERNO
            lblJobDate.Caption = !DateReceived
            lblType.Caption = IIf(IsNull(!JOType), "", !JOType)
            lblModel.Caption = !Model
            lblEngine.Caption = !ENGINE
            lblChassis.Caption = !CHASSIS
            lblCustomer.Caption = "Customer: " & !CUSTOMER
            OAdisc = IIf(IsNull(!DISCOUNT), 0, !DISCOUNT)
            MyJOid = !JOBORDERID
        End If
        .Close
    End With
    'Qry = "SELECT SRVJOBDETAILS.* From SRVJOBDETAILS " & _
          "WHERE (((SRVJOBDETAILS.JOBORDERID)=" & MyJOid & "));"
    'mySQL
    Qry = "SELECT SRVJOBDETAILS.* From SRVJOBDETAILS " & _
          "WHERE SRVJOBDETAILS.JOBORDERID=" & MyJOid & ";"
    
    With rsDetails
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            flxJob.Rows = 1
            .MoveFirst
            GTtl = 0
            Do While Not .EOF
                myDisc = IIf(IsNull(!DISCOUNT), 0, !DISCOUNT)
                MyLabor = IIf(IsNull(!Labor), 0, !Labor)
                myTtl = MyLabor - (MyLabor * (myDisc / 100))
                GTtl = GTtl + myTtl
                flxJob.AddItem vbTab & !JobRequest & vbTab & Format(MyLabor, "#,##0.00") & vbTab & IIf(myDisc <> 0, Format(myDisc, "##0.00") & "%", "") & vbTab & Format(myTtl, "#,##0.00")
                .MoveNext
            Loop
            txtTtlJob.Text = Format(GTtl, "#,##0.00")
        End If
        .Close
    End With
    If OAdisc > 0 Then
        lblTtl(4).Caption = Format(GTtl, "#,##0.00") & " Less " & OAdisc & "%:"
        GTtl = GTtl - (GTtl * (OAdisc / 100))
        txtTtlJob.Text = Format(GTtl, "#,##0.00")
    Else
        lblTtl(4).Caption = "Total"
    End If
    tabDocs.Tab = 1
End Sub

Private Sub DisplayInvoice()
    Dim rsITR As New ADODB.Recordset
    Dim Qry As String, MyUnit As String, MyReq As String, MyJo As String, MyPO As String, MyRef As String
    Dim Ttl As Currency
    
    'Qry = "SELECT SPITR.OUT, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPITR.PRICE, [PRICE]*[OUT] AS TTLPRICE, SPITR.SPAREPARTID " & _
          "FROM SPAREPARTS INNER JOIN SPITR ON SPAREPARTS.SPAREPARTID = SPITR.SPAREPARTID WHERE (((SPITR.REFERENCE)='SI#-" & flxSP.TextMatrix(ActRow, 2) & "')) ORDER BY SPAREPARTS.PARTNUMBER;"
    'mySQl
    If Left$(flxSP.TextMatrix(ActRow, 2), 3) <> "TRA" Then
        MyRef = "SI#-" & flxSP.TextMatrix(ActRow, 2)
    Else
        MyRef = flxSP.TextMatrix(ActRow, 2)
    End If
    Qry = "SELECT SPITR.`OUT`, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPITR.PRICE, SPITR.PRICE*SPITR.`OUT` AS TTLPRICE, SPITR.SPAREPARTID " & _
          "FROM SPAREPARTS INNER JOIN SPITR ON SPAREPARTS.SPAREPARTID = SPITR.SPAREPARTID WHERE SPITR.REFERENCE='" & MyRef & "'  AND SPITR.REMARKS = 'CREDIT' ORDER BY SPAREPARTS.PARTNUMBER;"
    
    With rsITR
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            flxInv.Rows = 1
            Ttl = 0
            Do While Not .EOF
                If InStr(!Description, "ENGINE OIL") > 0 Or InStr(!Description, "OIL 4T") > 0 Then
                    MyUnit = IIf(!OUT > 1, "Qrts.", "Qrt.")
                Else
                    MyUnit = IIf(!OUT > 1, "Pcs.", "Pc.")
                End If
                Ttl = Ttl + !TtlPrice
                flxInv.AddItem vbTab & !OUT & vbTab & MyUnit & vbTab & !PartNumber & vbTab & !Description & vbTab & Format(!Price, "#,##0.00") & vbTab & Format(!TtlPrice, "#,##0.00") '& vbTab & !SparepartID
                MyQty = ReturnedQty(!SparepartID)
                If MyQty > 0 Then
                    flxInv.Col = 1: flxInv.Row = flxInv.Rows - 1: flxInv.ColSel = 7: flxInv.RowSel = flxInv.Rows - 1: flxInv.FillStyle = flexFillRepeat: flxInv.CellBackColor = RGB(250, 250, 180)
                    flxInv.Col = 2: flxInv.FillStyle = flexFillSingle
                End If
                .MoveNext
            Loop
        Else
            flxInv.Rows = 1
            flxInv.Rows = 2
        End If
        .Close
    End With
    tabDocs.Tab = 0
    lblInvNo.Caption = flxSP.TextMatrix(ActRow, 2)
    lblInvDate.Caption = flxSP.TextMatrix(ActRow, 1)
    'Qry = "SELECT SPINVOICE.FRONTLINER From SPINVOICE WHERE (((SPINVOICE.SPINVNO)='" & lblInvNo.Caption & "'));"
    'mySQl
    Qry = "SELECT SPINVOICE.FRONTLINER From SPINVOICE WHERE SPINVOICE.SPINVNO='" & lblInvNo.Caption & "';"
    
    lblPartsMan.Caption = "PARTSMAN: " & GetFieldValue(Qry)
    txtTtlInv.Text = Format(Ttl, "#,##0.00")
    'Qry = "SELECT SPSALESORDER.REQID, SPSALESORDER.INVOICENO From SPSALESORDER WHERE (((SPSALESORDER.INVOICENO)='" & lblInvNo.Caption & "'));"
    'mySQL
    Qry = "SELECT SPSALESORDER.REQID, SPSALESORDER.INVOICENO From SPSALESORDER WHERE SPSALESORDER.INVOICENO='" & lblInvNo.Caption & "';"
    
    If Val(GetFieldValue(Qry)) > 0 Then
        'Qry = "SELECT REQUISITION.REQNO, SPSALESORDER.INVOICENO FROM SPSALESORDER INNER JOIN REQUISITION ON SPSALESORDER.REQID = REQUISITION.REQID WHERE (((SPSALESORDER.INVOICENO)='" & lblInvNo.Caption & "'));"
        'mySQl
        Qry = "SELECT REQUISITION.REQNO, SPSALESORDER.INVOICENO FROM SPSALESORDER INNER JOIN REQUISITION ON SPSALESORDER.REQID = REQUISITION.REQID WHERE SPSALESORDER.INVOICENO='" & lblInvNo.Caption & "';"
        
        MyReq = GetFieldValue(Qry)
        'Qry = "SELECT REQUISITION.JOBNO, SPSALESORDER.INVOICENO FROM SPSALESORDER INNER JOIN REQUISITION ON SPSALESORDER.REQID = REQUISITION.REQID WHERE (((SPSALESORDER.INVOICENO)='" & lblInvNo.Caption & "'));"
        
        'mySQL
        Qry = "SELECT REQUISITION.JOBNO, SPSALESORDER.INVOICENO FROM SPSALESORDER INNER JOIN REQUISITION ON SPSALESORDER.REQID = REQUISITION.REQID WHERE SPSALESORDER.INVOICENO='" & lblInvNo.Caption & "';"
        
        MyJo = GetFieldValue(Qry)
        lblInvRef.Caption = "REQ. NO. " & MyReq & ", JOB ORDER NO. " & MyJo
    Else
        'Qry = "SELECT SPSALESORDER.REFERENCE From SPSALESORDER WHERE (((SPSALESORDER.INVOICENO)='" & lblInvNo.Caption & "'));"
        
        'mySQL
        Qry = "SELECT SPSALESORDER.REFERENCE From SPSALESORDER WHERE SPSALESORDER.INVOICENO='" & lblInvNo.Caption & "';"
        
        MyPO = GetFieldValue(Qry)
        If Val(MyPO) > 0 Then
            lblInvRef.Caption = "TRA# " & MyPO
        Else
            lblInvRef.Caption = ""
        End If
    End If
        
End Sub

Function ReturnedQty(SPid As Long) As Integer
    Dim rsRetQ As New ADODB.Recordset
    Dim Qry As String
    
    'Qry = "SELECT SPITR.RECEIVE, SPITR.TRANSTYPE, SPITR.SPAREPARTID, SPITR.PACKINGLIST From SPITR " & _
          "WHERE (((SPITR.TRANSTYPE)='SALES RETURN') AND ((SPITR.SPAREPARTID)=" & SPid & ") AND ((SPITR.PACKINGLIST)='" & MyInv & "'));"
    'mySQl
    Qry = "SELECT SPITR.RECEIVE, SPITR.TRANSTYPE, SPITR.SPAREPARTID, SPITR.PACKINGLIST From SPITR " & _
          "WHERE SPITR.TRANSTYPE='SALES RETURN' AND SPITR.SPAREPARTID=" & SPid & " AND SPITR.PACKINGLIST='" & MyInv & "';"
    
    With rsRetQ
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            ReturnedQty = !RECEIVE
        Else
            ReturnedQty = 0
        End If
        .Close
    End With
End Function
                                
Private Sub ComputeFlxPCVttl()
    Dim Ttl As Currency, i As Integer
    
    With flxOthers
        Ttl = 0
        For i = 1 To .Rows - 1
            Ttl = Ttl + Val(Replace(.TextMatrix(i, 5), ",", ""))
        Next i
        txtTtlOthers.Text = Format(Ttl, "#,##0.00")
    End With
End Sub
                                
Private Sub ComputeFlxSVCttl()
    Dim i As Integer
    Dim TBal As Currency, TAmtPd As Currency, TAmtDue As Currency, TOvrDue As Currency, TPay As Currency
    
    TBal = 0: TAmtPd = 0: TAmtDue = 0: TOvrDue = 0: TPay = 0
    With flxSVC
        For i = 1 To .Rows - 1
            TBal = TBal + Val(Replace(.TextMatrix(i, 5), ",", ""))
            TAmtPd = TAmtPd + Val(Replace(.TextMatrix(i, 6), ",", ""))
            TAmtDue = TAmtDue + Val(Replace(.TextMatrix(i, 7), ",", ""))
            TOvrDue = TOvrDue + Val(Replace(.TextMatrix(i, 8), ",", ""))
            TPay = TPay + Val(Replace(.TextMatrix(i, 9), ",", ""))
        Next i
    End With
    With flxSVCttl
        .TextMatrix(0, 0) = Format(TBal, "#,##0.00")
        .TextMatrix(0, 1) = Format(TAmtPd, "#,##0.00")
        .TextMatrix(0, 2) = Format(TAmtDue, "#,##0.00")
        .TextMatrix(0, 3) = Format(TOvrDue, "#,##0.00")
        .TextMatrix(0, 4) = Format(TPay, "#,##0.00")
    End With
End Sub
                                
Private Sub ComputeFlxSPttl()
    Dim i As Integer
    Dim TBal As Currency, TAmtPd As Currency, TAmtDue As Currency, TOvrDue As Currency, TPay As Currency
    
    TBal = 0: TAmtPd = 0: TAmtDue = 0: TOvrDue = 0: TPay = 0
    With flxSP
        For i = 1 To .Rows - 1
            TBal = TBal + Val(Replace(.TextMatrix(i, 5), ",", ""))
            TAmtPd = TAmtPd + Val(Replace(.TextMatrix(i, 6), ",", ""))
            TAmtDue = TAmtDue + Val(Replace(.TextMatrix(i, 7), ",", ""))
            TOvrDue = TOvrDue + Val(Replace(.TextMatrix(i, 8), ",", ""))
            TPay = TPay + Val(Replace(.TextMatrix(i, 9), ",", ""))
        Next i
    End With
    With flxSPTtl
        .TextMatrix(0, 0) = Format(TBal, "#,##0.00")
        .TextMatrix(0, 1) = Format(TAmtPd, "#,##0.00")
        .TextMatrix(0, 2) = Format(TAmtDue, "#,##0.00")
        .TextMatrix(0, 3) = Format(TOvrDue, "#,##0.00")
        .TextMatrix(0, 4) = Format(TPay, "#,##0.00")
    End With
End Sub

Private Sub cboOtherAcct_Click()
    Select Case cboOtherAcct.ListIndex
        Case 0  'Open
            Call DisplayOpenPCV
            Call ComputeFlxPCVttl
        Case 1  'Close
            Call DisplayClosePCV
            Call ComputeFlxPCVttl
        Case 2
            Call DisplayOpenPCV
            Call ComputeFlxPCVttl
    End Select
End Sub

Private Sub cboSPacct_Click()
    Select Case cboSPacct.ListIndex
        Case 0  'Open
            Call DisplayOpenSP
            Call ComputeFlxSPttl
        Case 1  'Close
            Call DisplayCloseSP
            Call ComputeFlxSPttl
        Case 2
            Call DisplayOpenSP
            Call DisplayCloseSP
            Call ComputeFlxSPttl
    End Select
End Sub

Private Sub cboSVCacct_Click()
    Select Case cboSVCacct.ListIndex
        Case 0  'Open
            Call DisplayOpenSVC
            Call ComputeFlxSVCttl
        Case 1  'Close
            Call DisplayCloseSVC
            Call ComputeFlxSVCttl
        Case 2
            Call DisplayOpenSVC
            Call DisplayCloseSVC
            Call ComputeFlxSVCttl
    End Select
End Sub

Private Sub cmdC_Click()
    cboCustomer.SetFocus
End Sub

Private Sub cmdFindDMCM_Click()
    Dim i As Integer, j As Integer
        
    If txtFindDMCM.Text = "" Then Exit Sub
    With flxDMCM
        For i = fRow To .Rows - 1
            For j = fCol To 6
                If InStr(UCase$(.TextMatrix(i, j)), txtFindDMCM.Text) > 0 Then
                    .Col = 1: .Row = i
                    .ColSel = 6:  .RowSel = i
                    .TopRow = i
                    fRow = IIf(j = 6, i + 1, i)
                    fCol = IIf(j = 6, 1, j + 1)
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

Private Sub cmdFindPay_Click()
    Dim i As Integer, j As Integer
        
    If txtFindPay.Text = "" Then Exit Sub
    With FlxPay
        For i = fRow To .Rows - 1
            For j = fCol To 6
                If InStr(.TextMatrix(i, j), txtFindPay.Text) > 0 Then
                    .Col = 1: .Row = i
                    .ColSel = 6:  .RowSel = i
                    .TopRow = i
                    fRow = IIf(j = 6, i + 1, i)
                    fCol = IIf(j = 6, 1, j + 1)
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

Private Sub cmdFindSP_Click()
    Dim i As Integer, j As Integer
        
    If txtFindSP.Text = "" Then Exit Sub
    With flxSP
        For i = fRow To .Rows - 1
            For j = fCol To 10
                If InStr(.TextMatrix(i, j), txtFindSP.Text) > 0 Then
                    .Col = 1: .Row = i
                    .ColSel = 10:  .RowSel = i
                    .TopRow = i
                    fRow = IIf(j = 10, i + 1, i)
                    fCol = IIf(j = 10, 1, j + 1)
                    Call flxSP_Click
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

Private Sub cmdFindSVC_Click()
    Dim i As Integer, j As Integer
        
    If txtFindSVC.Text = "" Then Exit Sub
    With flxSVC
        For i = fRow To .Rows - 1
            For j = fCol To 10
                If InStr(.TextMatrix(i, j), txtFindSVC.Text) > 0 Then
                    .Col = 1: .Row = i
                    .ColSel = 10:  .RowSel = i
                    .TopRow = i
                    fRow = IIf(j = 10, i + 1, i)
                    fCol = IIf(j = 10, 1, j + 1)
                    Call flxSVC_Click
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

Private Sub cmdFndOthers_Click()
    Dim i As Integer, j As Integer
        
    If txtFindOthers.Text = "" Then Exit Sub
    With flxOthers
        For i = fRow To .Rows - 1
            For j = fCol To 6
                If InStr(UCase$(.TextMatrix(i, j)), txtFindOthers.Text) > 0 Then
                    .Col = 1: .Row = i
                    .ColSel = 6:  .RowSel = i
                    .TopRow = i
                    fRow = IIf(j = 6, i + 1, i)
                    fCol = IIf(j = 6, 1, j + 1)
                    Call flxOthers_Click
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

Private Sub cmdInvDetails_Click()
    InvDetails.MyInv = lblInvNo.Caption
    InvDetails.CancelInv = False
    InvDetails.Show vbModal
End Sub

Private Sub cmdTrans_Click(Index As Integer)
    Select Case Index
        Case 0  'Summary
        
        Case 1  'exit
            Unload Me
            'MotoRep.tbMain.Visible = True
            'MotoRep.imgBG.Visible = True
            MotoRep.toolBarButton.Visible = True
    End Select
End Sub

Private Sub flxOthers_Click()
    ActRow = flxOthers.Row
    ActTab = 0
    If ActRow = 0 Or flxOthers.TextMatrix(1, 1) = "" Then Exit Sub
    flxLedger.Rows = 1
    flxLedger.Rows = 2
    Call DisplayPCV
End Sub

Private Sub flxOthers_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ActRow = flxOthers.Row
        ActTab = 0
        If ActRow = 0 Or flxOthers.TextMatrix(1, 1) = "" Then Exit Sub
        flxLedger.Rows = 1
        flxLedger.Rows = 2
        Call DisplayPCV
    End If
End Sub

Private Sub flxPay_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
    If Button = vbRightButton Then
        If InStr(UCase$(FlxPay.TextMatrix(FlxPay.Row, 6)), "ATTACHMENT") > 0 Then Call PopupMenu(mnuPay)
    End If
End Sub

Private Sub flxSP_Click()
    ActRow = flxSP.Row
    ActTab = 0
    If ActRow = 0 Or flxSP.TextMatrix(1, 1) = "" Then Exit Sub
    Call DisplayLedger
    Call DisplayInvoice
End Sub

Private Sub flxSP_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ActRow = flxSP.Row
        ActTab = 0
        If ActRow = 0 Or flxSP.TextMatrix(1, 1) = "" Then Exit Sub
        Call DisplayLedger
        Call DisplayInvoice
    End If
End Sub

Private Sub flxSP_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    If Button = vbRightButton Then
        If Val(txtCustomerID.Text) > 0 Then Call PopupMenu(mnuPrn)
    End If
End Sub

Private Sub flxSVC_Click()
    ActRow = flxSVC.Row
    ActTab = 1
    If ActRow = 0 Or flxSVC.TextMatrix(1, 1) = "" Then Exit Sub
    Call DisplayLedger
    Call DisplayJobOrder
End Sub

Private Sub flxSVC_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ActRow = flxSVC.Row
        ActTab = 1
        If ActRow = 0 Or flxSP.TextMatrix(1, 1) = "" Then Exit Sub
        Call DisplayLedger
        Call DisplayJobOrder
    End If
End Sub

Private Sub Form_Load()
    
    cboSPacct.AddItem "OPEN ACCOUNTS"
    cboSPacct.AddItem "CLOSE ACCOUNTS"
    cboSPacct.AddItem "ALL ACCOUNTS"
    cboSVCacct.AddItem "OPEN ACCOUNTS"
    cboSVCacct.AddItem "CLOSE ACCOUNTS"
    cboSVCacct.AddItem "ALL ACCOUNTS"
    cboOtherAcct.AddItem "OPEN ACCOUNTS"
    cboOtherAcct.AddItem "CLOSE ACCOUNTS"
    cboOtherAcct.AddItem "ALL ACCOUNTS"
    tabAccts.Tab = 0
    tabDocs.Tab = 0
    cmdC.Top = 15000
    MotoRep.sBar.Panels(6).Text = "Customer Accounts"
    Timer1.Enabled = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If rsAcctCust.State = adStateOpen Then rsAcctCust.Close
End Sub


Private Sub mnuAttach_Click()
    Dim Qry As String
    
    'Qry = "SELECT PAYMENTS.CUSTOMERNAME, [SPAREPARTS]+[SERVICE]+[OTHERS] AS ORAMOUNT, PAYMENTS.ORNO, PAYMENTS.PAYDATE, ORDETAILS.INVOICENO, ORDETAILS.JOBORDERNO, ORDETAILS.TRANSDATE, ORDETAILS.ACCTAMT, ORDETAILS.AMTPAID, ORDETAILS.TTLPAYMENT, ORDETAILS.BALANCE " & _
          "FROM PAYMENTS INNER JOIN ORDETAILS ON PAYMENTS.PAYMENTID = ORDETAILS.PAYMENTID WHERE (((PAYMENTS.ORNO)='" & flxPay.TextMatrix(flxPay.Row, 2) & "') AND ((ORDETAILS.CUSTOMERID)=" & txtCustomerID.Text & ")) ORDER BY ORDETAILS.JOBORDERNO, ORDETAILS.INVOICENO;"
    'mySQL
    Qry = "SELECT PAYMENTS.CUSTOMERNAME, PAYMENTS.SPAREPARTS+PAYMENTS.SERVICE+PAYMENTS.OTHERS AS ORAMOUNT, PAYMENTS.ORNO, PAYMENTS.PAYDATE, ORDETAILS.INVOICENO, ORDETAILS.JOBORDERNO, ORDETAILS.TRANSDATE, ORDETAILS.ACCTAMT, ORDETAILS.AMTPAID, ORDETAILS.TTLPAYMENT, ORDETAILS.BALANCE " & _
          "FROM PAYMENTS INNER JOIN ORDETAILS ON PAYMENTS.PAYMENTID = ORDETAILS.PAYMENTID WHERE PAYMENTS.ORNO='" & FlxPay.TextMatrix(FlxPay.Row, 2) & "' AND ORDETAILS.CUSTOMERID=" & txtCustomerID.Text & " ORDER BY ORDETAILS.JOBORDERNO, ORDETAILS.INVOICENO;"
    
    PRNatt.MySource = "ATTACHMENT"
    PRNatt.QryPrn = Qry
    PRNatt.Show vbModal
End Sub

Private Sub mnuPrnLdgr_Click()
    Dim Qry As String, MyDue As String
    
    If tabAccts.Tab = 0 Then
        'Qry = "SELECT LEDGER.REFDATE, LEDGER.REFNO, LEDGER.DEBIT, LEDGER.CREDIT, LEDGER.OSBAL, LEDGER.REFTYPE From LEDGER " & _
              "Where (((LEDGER.SPINVNO) = '" & flxSP.TextMatrix(ActRow, 2) & "') And ((LEDGER.CUSTOMERID) = " & txtCustomerID.Text & ")) ORDER BY LEDGER.SEQNO;"
        'mySQL
        Qry = "SELECT LEDGER.REFDATE, LEDGER.REFNO, LEDGER.DEBIT, LEDGER.CREDIT, LEDGER.OSBAL, LEDGER.REFTYPE From LEDGER " & _
              "Where LEDGER.SPINVNO = '" & flxSP.TextMatrix(ActRow, 2) & "' And LEDGER.CUSTOMERID = " & txtCustomerID.Text & " ORDER BY LEDGER.SEQNO;"
        
        If InStr(flxSP.TextMatrix(ActRow, 4), "WKS") > 0 Then
            MyDue = DateAdd("d", 7, DateValue(flxSP.TextMatrix(ActRow, 1)))
        ElseIf InStr(flxSP.TextMatrix(ActRow, 4), "BI-MOS") > 0 Then
            MyDue = DateAdd("d", 15, DateValue(flxSP.TextMatrix(ActRow, 1)))
        Else
            MyDue = DateAdd("M", 1, DateValue(flxSP.TextMatrix(ActRow, 1)))
        End If
    Else
        'Qry = "SELECT LEDGER.REFDATE, LEDGER.REFNO, LEDGER.DEBIT, LEDGER.CREDIT, LEDGER.OSBAL, LEDGER.REFTYPE From LEDGER " & _
              "WHERE (((LEDGER.SPINVNO)='" & flxSVC.TextMatrix(ActRow, 2) & "') AND ((LEDGER.CUSTOMERID)=" & txtCustomerID.Text & ") AND ((LEDGER.JOBID)>0)) ORDER BY LEDGER.SEQNO;"
        'mySQL
        Qry = "SELECT LEDGER.REFDATE, LEDGER.REFNO, LEDGER.DEBIT, LEDGER.CREDIT, LEDGER.OSBAL, LEDGER.REFTYPE From LEDGER " & _
              "WHERE (((LEDGER.SPINVNO)='" & flxSVC.TextMatrix(ActRow, 2) & "') AND ((LEDGER.CUSTOMERID)=" & txtCustomerID.Text & ") AND ((LEDGER.JOBID)>0)) ORDER BY LEDGER.SEQNO;"
        
        If InStr(flxSVC.TextMatrix(ActRow, 4), "WKS") > 0 Then
            MyDue = DateAdd("d", 7, DateValue(flxSVC.TextMatrix(ActRow, 1)))
        ElseIf InStr(flxSVC.TextMatrix(ActRow, 4), "BI-MOS") > 0 Then
            MyDue = DateAdd("d", 15, DateValue(flxSVC.TextMatrix(ActRow, 1)))
        Else
            MyDue = DateAdd("M", 1, DateValue(flxSVC.TextMatrix(ActRow, 1)))
        End If
    End If
    PRNatt.QryPrn = Qry
    PRNatt.ParamValue1 = cboCustomer.Text
    PRNatt.ParamValue2 = txtAddress.Text
    PRNatt.ParamValue3 = txtEmpCo.Text
    If tabAccts.Tab = 0 Then
        PRNatt.ParamValue4 = flxSP.TextMatrix(ActRow, 2)
        PRNatt.ParamValue5 = flxSP.TextMatrix(ActRow, 1)
        PRNatt.ParamValue6 = flxSP.TextMatrix(ActRow, 3)
        PRNatt.ParamValue7 = flxSP.TextMatrix(ActRow, 4)
    Else
        PRNatt.ParamValue4 = flxSVC.TextMatrix(ActRow, 2)
        PRNatt.ParamValue5 = flxSVC.TextMatrix(ActRow, 1)
        PRNatt.ParamValue6 = flxSVC.TextMatrix(ActRow, 3)
        PRNatt.ParamValue7 = flxSVC.TextMatrix(ActRow, 4)
    End If
    PRNatt.ParamValue8 = MyDue
    PRNatt.MySource = "LEDGER"
    PRNatt.Top = (Me.Height - PRNatt.Height) \ 2
    PRNatt.Left = (Me.Width - PRNatt.Width) \ 2
    PRNatt.Show vbModal

    
End Sub

Private Sub RXResizer1_ResizeComplete()
    Dim i As Integer
    
    lblHeading(0).Top = Frame1.Top - (lblHeading(0).Height + 50)
    lblHeading(1).Top = Frame1.Top - (lblHeading(1).Height + 20)
    lblHeading(2).Top = Frame1.Top - (lblHeading(0).Height + 100)
    lblHeading(3).Top = Frame1.Top - (lblHeading(1).Height + 70)
    lblHeading(2).Left = (Frame1.Left + Frame1.Width) - (lblHeading(2).Width + 30)
    lblHeading(3).Left = (Frame1.Left + Frame1.Width) - lblHeading(3).Width
    With flxBill
        For i = 0 To .Cols - 1
            .ColWidth(i) = (.Width - 150) / 6
        Next i
    End With
    With flxSP
        .ColWidth(0) = 300
        For i = 1 To 9
            .ColWidth(i) = (.Width - 300) * 0.1
        Next i
        .ColWidth(5) = (.Width - 300) * 0.12
        .ColWidth(6) = (.Width - 300) * 0.12
        .ColWidth(10) = (.Width - 300) * 0.04
        flxSPTtl.Width = .Width - (.ColWidth(0) + .ColWidth(1) + .ColWidth(2) + .ColWidth(3) + .ColWidth(4))
        flxSPTtl.Left = .Left + .ColWidth(0) + .ColWidth(1) + .ColWidth(2) + .ColWidth(3) + .ColWidth(4)
        lblTtl(0).Left = flxSPTtl.Left - (lblTtl(0).Width + 100)
        For i = 0 To 4
            flxSPTtl.ColWidth(i) = .ColWidth(i + 5)
        Next i
    End With
    With flxSVC
        .ColWidth(0) = 300
        For i = 1 To 9
            .ColWidth(i) = (.Width - 300) * 0.1
        Next i
        .ColWidth(5) = (.Width - 300) * 0.12
        .ColWidth(6) = (.Width - 300) * 0.12
        .ColWidth(10) = (.Width - 300) * 0.04
        flxSVCttl.Width = .Width - (.ColWidth(0) + .ColWidth(1) + .ColWidth(2) + .ColWidth(3) + .ColWidth(4))
        flxSVCttl.Left = .Left + .ColWidth(0) + .ColWidth(1) + .ColWidth(2) + .ColWidth(3) + .ColWidth(4)
        lblTtl(1).Left = flxSVCttl.Left - (lblTtl(1).Width + 100)
        For i = 0 To 4
            flxSVCttl.ColWidth(i) = .ColWidth(i + 5)
        Next i
    End With
    With flxOthers
        .ColWidth(0) = 300
        For i = 1 To 6
            .ColWidth(i) = (.Width - 300) * 0.13
        Next i
        .ColWidth(3) = (.Width - 300) * 0.22
        .ColWidth(4) = (.Width - 300) * 0.25
        .ColWidth(6) = (.Width - 300) * 0.1
        txtTtlOthers.Left = .Left + .ColWidth(0) + .ColWidth(1) + .ColWidth(2) + .ColWidth(3) + .ColWidth(4)
        txtTtlOthers.Width = .ColWidth(5)
        lblTtl(2).Left = txtTtlOthers.Left - (lblTtl(2).Width + 100)
    End With
    With FlxPay
        .ColWidth(0) = 300
        For i = 1 To 5
            .ColWidth(i) = (.Width - 300) * 0.14
        Next i
        .ColWidth(6) = (.Width - 300) * 0.27
        flxPAYttl.Width = .Width - (.ColWidth(0) + .ColWidth(1) + .ColWidth(2))
        flxPAYttl.Left = .Left + .ColWidth(0) + .ColWidth(1) + .ColWidth(2)
        lblTtl(6).Left = flxPAYttl.Left - (lblTtl(1).Width + 100)
        For i = 0 To 2
            flxPAYttl.ColWidth(i) = .ColWidth(i + 2)
        Next i
    End With
    With flxDMCM
        .ColWidth(0) = 300
        For i = 1 To 5
            .ColWidth(i) = (.Width - 300) * 0.12
        Next i
        .ColWidth(3) = (.Width - 300) * 0.24
        .ColWidth(6) = (.Width - 300) * 0.24
        flxDMCMttl.Width = .Width - (.ColWidth(0) + .ColWidth(1) + .ColWidth(2) + .ColWidth(3))
        flxDMCMttl.Left = .Left + .ColWidth(0) + .ColWidth(1) + .ColWidth(2) + .ColWidth(3)
        lblTtl(7).Left = flxDMCMttl.Left - (lblTtl(7).Width + 100)
        For i = 0 To 1
            flxDMCMttl.ColWidth(i) = .ColWidth(i + 4)
        Next i
    End With
    With flxLedger
        For i = 0 To .Cols - 1
            .ColWidth(i) = (.Width - 150) / 6
        Next i
    End With
    With flxJob
        .ColWidth(0) = 300
        .ColWidth(1) = (.Width - 300) * 0.4
        .ColWidth(2) = (.Width - 300) * 0.2
        .ColWidth(3) = (.Width - 300) * 0.18
        .ColWidth(4) = (.Width - 300) * 0.2
        txtTtlJob.Width = .ColWidth(4)
        txtTtlJob.Left = .Left + .Width - txtTtlJob.Width
        lblTtl(4).Left = txtTtlJob.Left - (lblTtl(4).Width + 50)
    End With
    With flxPCV
        .ColWidth(0) = 300
        .ColWidth(1) = (.Width - 300) * 0.68
        .ColWidth(2) = (.Width - 300) * 0.3
        txtTtlPCV.Width = .ColWidth(2)
        txtTtlPCV.Left = .Left + .Width - txtTtlPCV.Width
        lblTtl(5).Left = txtTtlPCV.Left - (lblTtl(5).Width + 50)
    End With
End Sub

Private Sub tabAccts_GotFocus()
    Select Case tabAccts.Tab
        Case 0
            cboSPacct.SetFocus
        Case 1
            cboSVCacct.SetFocus
        Case 2
            cboOtherAcct.SetFocus
    End Select
End Sub

Private Sub Timer1_Timer()
    Dim Qry As String
    
    Qry = "SELECT CUSTOMER.CUSTOMERID, CUSTOMER.CUSTOMERNAME, CUSTOMER.ADDRESS, CUSTOMER.CONTACTPERSON, CUSTOMER.CTYPE " & _
          "From CUSTOMER ORDER BY CUSTOMER.CUSTOMERNAME;"
    With rsAcctCust
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            cboCustomer.Clear
            Do While Not .EOF
                cboCustomer.AddItem !CustomerName
                .MoveNext
            Loop
        End If
    End With
    Timer1.Enabled = False
End Sub

Private Sub txtFindDMCM_Change()
    fRow = 1
    fCol = 1
End Sub

Private Sub txtFindDMCM_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then Call cmdFindDMCM_Click
End Sub

Private Sub txtFindOthers_Change()
    fRow = 1
    fCol = 1
End Sub

Private Sub txtFindOthers_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then Call cmdFndOthers_Click
End Sub

Private Sub txtFindPay_Change()
    fRow = 1
    fCol = 1
End Sub

Private Sub txtFindPay_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call cmdFindPay_Click
End Sub

Private Sub txtFindSP_Change()
    fRow = 1
    fCol = 1
End Sub

Private Sub txtFindSP_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call cmdFindSP_Click
End Sub

Private Sub txtFindSVC_Change()
    fRow = 1
    fCol = 1
End Sub

Private Sub txtFindSVC_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call cmdFindSVC_Click
End Sub
