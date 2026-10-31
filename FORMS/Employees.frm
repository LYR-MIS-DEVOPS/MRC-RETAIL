VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form Employees 
   Caption         =   "Employees Profile"
   ClientHeight    =   5370
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11685
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   5370
   ScaleWidth      =   11685
   WindowState     =   2  'Maximized
   Begin VB.CommandButton TBcmd 
      Caption         =   "&S"
      Height          =   495
      Index           =   0
      Left            =   -120
      TabIndex        =   36
      Top             =   8400
      Width           =   75
   End
   Begin VB.Frame Frame4 
      Height          =   4455
      Left            =   6960
      TabIndex        =   29
      Tag             =   "XYF"
      Top             =   720
      Width           =   4455
      Begin VB.ComboBox cboDept 
         Height          =   315
         ItemData        =   "Employees.frx":0000
         Left            =   1560
         List            =   "Employees.frx":0002
         TabIndex        =   47
         Tag             =   "XYF"
         Top             =   3600
         Width           =   2655
      End
      Begin lvButton.lvButtons_H cmdBrowse 
         Height          =   375
         Left            =   1200
         TabIndex        =   37
         Tag             =   "XYF"
         Top             =   1920
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         Caption         =   "Browse Picture"
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
         Image           =   "Employees.frx":0004
         cBack           =   -2147483633
      End
      Begin MSComDlg.CommonDialog cdgPic 
         Left            =   120
         Top             =   4080
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComCtl2.DTPicker Hired 
         Height          =   315
         Left            =   1560
         TabIndex        =   14
         Tag             =   "XYF"
         Top             =   3240
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   103809025
         CurrentDate     =   38841
      End
      Begin VB.ComboBox cboStatus 
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
         Left            =   1560
         TabIndex        =   13
         Tag             =   "XYF"
         Top             =   2880
         Width           =   2655
      End
      Begin VB.ComboBox cboPosition 
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
         Left            =   1560
         TabIndex        =   12
         Tag             =   "XYF"
         Top             =   2520
         Width           =   2655
      End
      Begin VB.Label Label3 
         Caption         =   "DEPARTMENT"
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
         TabIndex        =   46
         Tag             =   "XYF"
         Top             =   3600
         Width           =   1215
      End
      Begin VB.Label lblPicLoc 
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   3960
         Visible         =   0   'False
         Width           =   4215
      End
      Begin VB.Image ImgEmp 
         BorderStyle     =   1  'Fixed Single
         Height          =   1575
         Left            =   1200
         Stretch         =   -1  'True
         Tag             =   "XYF"
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label2 
         Caption         =   "DATE HIRED"
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
         TabIndex        =   32
         Tag             =   "XYF"
         Top             =   3270
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "EMP.  STATUS"
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
         TabIndex        =   31
         Tag             =   "XYF"
         Top             =   2910
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "POSITION"
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
         TabIndex        =   30
         Tag             =   "XYF"
         Top             =   2550
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      Height          =   4455
      Left            =   480
      TabIndex        =   18
      Tag             =   "XYF"
      Top             =   720
      Width           =   6375
      Begin VB.ComboBox cboCivilStat 
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
         TabIndex        =   9
         Tag             =   "XYF"
         Top             =   3015
         Width           =   3975
      End
      Begin VB.ComboBox cboGender 
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
         TabIndex        =   6
         Tag             =   "XYF"
         Top             =   1965
         Width           =   3975
      End
      Begin VB.TextBox txtEntry 
         Height          =   285
         Index           =   0
         Left            =   360
         TabIndex        =   34
         Top             =   4080
         Visible         =   0   'False
         Width           =   1575
      End
      Begin MSComCtl2.DTPicker Birth 
         Height          =   315
         Left            =   2040
         TabIndex        =   7
         Tag             =   "XYF"
         Top             =   2325
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   103809025
         CurrentDate     =   38853
      End
      Begin VB.TextBox txtEntry 
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
         Index           =   7
         Left            =   2040
         MaxLength       =   25
         TabIndex        =   8
         Tag             =   "XYF"
         Top             =   2670
         Width           =   3975
      End
      Begin VB.TextBox txtEntry 
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
         Index           =   10
         Left            =   2040
         MaxLength       =   50
         TabIndex        =   11
         Tag             =   "XYF"
         Top             =   3720
         Width           =   3975
      End
      Begin VB.TextBox txtEntry 
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
         Index           =   5
         Left            =   2040
         MaxLength       =   25
         TabIndex        =   5
         Tag             =   "XYF"
         Top             =   1620
         Width           =   3975
      End
      Begin VB.TextBox txtEntry 
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
         Index           =   4
         Left            =   2040
         MaxLength       =   50
         TabIndex        =   4
         Tag             =   "XYF"
         Top             =   1275
         Width           =   3975
      End
      Begin VB.TextBox txtEntry 
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
         Index           =   9
         Left            =   2040
         MaxLength       =   50
         TabIndex        =   10
         Tag             =   "XYF"
         Top             =   3375
         Width           =   3975
      End
      Begin VB.TextBox txtEntry 
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
         Index           =   1
         Left            =   2040
         MaxLength       =   25
         TabIndex        =   1
         Tag             =   "XYF"
         Top             =   240
         Width           =   3975
      End
      Begin VB.TextBox txtEntry 
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
         Index           =   2
         Left            =   2040
         MaxLength       =   25
         TabIndex        =   2
         Tag             =   "XYF"
         Top             =   585
         Width           =   3975
      End
      Begin VB.TextBox txtEntry 
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
         Index           =   3
         Left            =   2040
         MaxLength       =   25
         TabIndex        =   3
         Tag             =   "XYF"
         Top             =   930
         Width           =   3975
      End
      Begin VB.Label Label1 
         Caption         =   "GENDER"
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
         TabIndex        =   33
         Tag             =   "XYF"
         Top             =   1995
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "RELIGION"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   17
         Left            =   360
         TabIndex        =   28
         Tag             =   "XYF"
         Top             =   2700
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "SPOUSE EMPLOYER"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   16
         Left            =   360
         TabIndex        =   27
         Tag             =   "XYF"
         Top             =   3735
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "CIVIL STATUS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   15
         Left            =   360
         TabIndex        =   26
         Tag             =   "XYF"
         Top             =   3045
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "BIRTH DATE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   13
         Left            =   360
         TabIndex        =   25
         Tag             =   "XYF"
         Top             =   2355
         Width           =   1335
      End
      Begin VB.Label Label1 
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
         Height          =   240
         Index           =   11
         Left            =   360
         TabIndex        =   24
         Tag             =   "XYF"
         Top             =   1650
         Width           =   1335
      End
      Begin VB.Label Label1 
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
         Height          =   240
         Index           =   9
         Left            =   360
         TabIndex        =   23
         Tag             =   "XYF"
         Top             =   1305
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "NAME OF SPOUSE"
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
         TabIndex        =   22
         Tag             =   "XYF"
         Top             =   3390
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "LAST NAME"
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
         TabIndex        =   21
         Tag             =   "XYF"
         Top             =   255
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "MIDDLE NAME"
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
         TabIndex        =   20
         Tag             =   "XYF"
         Top             =   945
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "FIRST NAME"
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
         TabIndex        =   19
         Tag             =   "XYF"
         Top             =   600
         Width           =   1455
      End
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   38
      Top             =   0
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
      Caption         =   "&New Employee"
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
      Image           =   "Employees.frx":06DA
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   3
      Left            =   5040
      TabIndex        =   39
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
      Image           =   "Employees.frx":0A74
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   4
      Left            =   6720
      TabIndex        =   40
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
      Image           =   "Employees.frx":100E
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   5
      Left            =   8400
      TabIndex        =   41
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
      Image           =   "Employees.frx":15A8
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   6
      Left            =   10080
      TabIndex        =   42
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
      Image           =   "Employees.frx":1B42
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   7
      Left            =   11760
      TabIndex        =   43
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
      Image           =   "Employees.frx":20DC
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   1
      Left            =   1680
      TabIndex        =   44
      Top             =   0
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
      Caption         =   "&Edit Employee"
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
      Image           =   "Employees.frx":2676
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   2
      Left            =   3360
      TabIndex        =   45
      Top             =   0
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
      Caption         =   "&Search"
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
      Image           =   "Employees.frx":2C10
      cBack           =   -2147483633
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   1320
      Top             =   6120
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   11925
      m_MDI_Height    =   5940
   End
   Begin VB.Label lblHeading 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Employee's Profile"
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
      Left            =   7410
      TabIndex        =   16
      Top             =   450
      Width           =   4005
   End
   Begin VB.Label lblHeading 
      BackStyle       =   0  'Transparent
      Caption         =   "Moto-R Republic"
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
      Left            =   480
      TabIndex        =   0
      Top             =   450
      Width           =   2415
   End
   Begin VB.Label lblHeading 
      BackStyle       =   0  'Transparent
      Caption         =   "Moto-R Republic"
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
      Left            =   510
      TabIndex        =   15
      Top             =   480
      Width           =   2415
   End
   Begin VB.Label lblHeading 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Employee's Profile"
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
      Left            =   7440
      TabIndex        =   17
      Top             =   480
      Width           =   4005
   End
   Begin VB.Menu mnuAdmin 
      Caption         =   "Admin and Staff"
      Visible         =   0   'False
   End
End
Attribute VB_Name = "Employees"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsAdmin As New ADODB.Recordset
Dim AddNew As Boolean


Sub ClearEntries()
    Dim i As Integer
    
    For i = 0 To 10
        If i <> 6 And i <> 8 Then txtEntry(i).Text = ""
    Next i
    cboGender.Text = ""
    cboCivilStat.Text = ""
    cboDept.Text = ""
    ImgEmp.Picture = LoadPicture("")
    Birth.Value = Date
    Hired.Value = Date
    cboStatus.Text = ""
    cboPosition.Text = ""
    'clear grids on tab
End Sub

Sub LockData(bln As Boolean)
    Dim i As Integer
    
    For i = 1 To 10
        If i <> 6 And i <> 8 Then txtEntry(i).Locked = bln
    Next i
    cboGender.Locked = bln
    cboCivilStat.Locked = bln
    cboDept.Locked = bln
    cmdBrowse.Enabled = Not bln
    Birth.Enabled = Not bln
    Hired.Enabled = Not bln
    cboPosition.Locked = bln
    cboStatus.Locked = bln
End Sub

Sub tbSaveCancel(bln As Boolean)
    Dim i As Integer

    If bln Then
        For i = 0 To 5
            cmdTrans(i).Visible = False
        Next i
        cmdTrans(6).Left = 0
        cmdTrans(7).Left = cmdTrans(6).Width
        cmdTrans(6).Visible = True
        cmdTrans(7).Visible = True
    Else
        For i = 0 To 5
            cmdTrans(i).Visible = True
        Next i
        cmdTrans(6).Visible = False
        cmdTrans(7).Visible = False
    End If
End Sub
Sub SaveAdmin()
    Dim X As Integer
    Dim Qry As String
    
    'Access Code
    If AddNew Then
        rsAdmin.Open "EMPLOYEES", LYRcon, adOpenKeyset, adLockOptimistic
        rsAdmin.AddNew
       
    Else
        Qry = "SELECT EMPLOYEES.* From EMPLOYEES " & _
              "WHERE (((EMPLOYEES.EMPID)=" & txtEntry(0).Text & "));"
           
        rsAdmin.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        rsAdmin.MoveFirst
    End If
    With rsAdmin
        !LASTNAME = txtEntry(1).Text
        !FirstName = txtEntry(2).Text
        !MIDDLENAME = txtEntry(3).Text
        !Address = txtEntry(4).Text
        !CONTACTNOS = txtEntry(5).Text
        !GENDER = cboGender.Text
        !RELIGION = txtEntry(7).Text
        !CIVILSTATUS = cboCivilStat.Text
        !SPOUSE = txtEntry(9).Text
        !SPOUSEEMP = txtEntry(10).Text
        .Fields("BirthDate") = Birth.Value
        .Fields("DateHired") = Hired.Value
        .Fields("Position") = cboPosition.Text
        If cboStatus.Text = "In-Active" Then
            .Fields("Inactive") = 1
        Else
            .Fields("Inactive") = 0
        End If
        .Fields("EmpStatus") = cboStatus.Text
        .Fields("Department") = UCase(cboDept.Text)
        Call SavePic
        .Fields("UpdatedBy") = UserID
        .Fields("LastUpdate") = Date
        .Update
        If AddNew Then txtEntry(0).Text = !EMPID
        .Close
    End With
End Sub
Private Sub SavePic()
    Dim i As Integer
    
    'Save File
    If lblPicLoc.Caption = "" Then Exit Sub
    If (Dir(Trim(lblPicLoc.Caption)) = "") Then Exit Sub
    'Open as Binary
    Open Trim(lblPicLoc.Caption) For Binary Access Read As lngDataFile
    lngLengh = LOF(lngDataFile)    ' Length of data in file
    If lngLengh = 0 Then Close lngDataFile: Exit Sub
    intChunks = lngLengh \ ChunkSize
    intFragment = lngLengh Mod ChunkSize
    'Add New Record in DataBase
        ReDim Chunk(intFragment)
        'Read data from a file into a variable
        Get lngDataFile, , Chunk()
        'Appends data to a large text or binary data Field or Parameter object.
        rsAdmin.Fields("EmpPicture").AppendChunk Chunk()
        ReDim Chunk(ChunkSize)
        For i = 1 To intChunks
            Get lngDataFile, , Chunk()
            rsAdmin.Fields("EmpPicture").AppendChunk Chunk()
        Next i
    'Update
    'Close File
    Close lngDataFile
End Sub

Public Sub ShowPic()
    On Error Resume Next
    
    Open "pictemp" For Binary Access Write As lngDataFile
        lngLengh = rsAdmin.Fields(16).ActualSize
        intChunks = lngLengh \ ChunkSize
        intFragment = lngLengh Mod ChunkSize
        ReDim Chunk(intFragment)
        Chunk() = rsAdmin.Fields(16).GetChunk(intFragment)
        Put lngDataFile, , Chunk()
        For i = 1 To intChunks
           ReDim Buffer(ChunkSize)
            Chunk() = rsAdmin.Fields("EmpPicture").GetChunk(ChunkSize)
            Put lngDataFile, , Chunk()
        Next i
    Close lngDataFile
    
    filename = "pictemp"
    ImgEmp.Picture = LoadPicture(filename)
End Sub
Public Sub ShowRecord(EMPID As String)
    Dim Qry As String
    Dim X As Integer
    'ACCESS
    'Qry = "SELECT EMPLOYEES.* From EMPLOYEES " & _
          "WHERE (((EMPLOYEES.EMPID)=" & EmpId & "));"
    
    'mySQL
    Qry = "SELECT EMPLOYEES.* From EMPLOYEES " & _
          "WHERE EMPLOYEES.EMPID=" & EMPID & ";"
    
    rsAdmin.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    With rsAdmin
        .MoveFirst
        For X = 0 To 10
            If X <> 6 And X <> 8 Then txtEntry(X).Text = IIf(IsNull(.Fields(X)), "", .Fields(X))
        Next X
        cboGender.Text = IIf(IsNull(!GENDER), "", !GENDER)
        cboCivilStat.Text = IIf(IsNull(!CIVILSTATUS), "", !CIVILSTATUS)
        Birth.Value = .Fields("BirthDate")
        Hired.Value = .Fields("DateHired")
        cboPosition.Text = IIf(IsNull(.Fields("Position")), "", .Fields("Position"))
        cboStatus.Text = IIf(IsNull(.Fields("EmpStatus")), "", .Fields("EmpStatus"))
        cboDept.Text = IIf(IsNull(.Fields("Department")), "", .Fields("Department"))
        If Not IsNull(.Fields("EmpPicture")) Then
            Call ShowPic
        Else
            ImgEmp.Picture = LoadPicture("")
        End If
        .Close
    End With
End Sub

Function ValidateEntry() As String
    Dim myMsg As String
    
    myMsg = ""
    If txtEntry(1).Text = "" Then myMsg = "No data on field LAST NAME"
    If txtEntry(2).Text = "" Then myMsg = myMsg & vbCrLf & "No data on field FIRST NAME"
    If txtEntry(4).Text = "" Then myMsg = myMsg & vbCrLf & "No data on field ADDRESS"
    If cboPosition.Text = "" Then myMsg = myMsg & vbCrLf & "Select POSITION from the list."
    If cboDept.Text = "" Then myMsg = myMsg & vbCrLf & "Select DEPARTMENT from the list."
    If AddNew Then
        If EmployeeExist Then myMsg = myMsg & vbCrLf & "Employee record already exist!"
    End If
    ValidateEntry = myMsg
End Function

Private Sub cboCivilStat_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub cboGender_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub cboStatus_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub cboPosition_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub cmdBrowse_Click()
    With cdgPic
        .ShowOpen
        lblPicLoc.Caption = .filename
        ImgEmp.Picture = LoadPicture(.filename)
    End With
End Sub

Private Sub cmdTrans_Click(Index As Integer)
    Dim vMsg As String, QrySrch As String, Qry As String, MyAcc As String
    Dim EmployeeID As Integer, Ans As Integer
    
    Select Case Index
        Case 0  '"addnew"
            'Access
            'Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE (((USERACCESS.EMPID)=" & UserID & ") AND ((USERACCESS.MENUINDX)=8));"
            'mySQL
            Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE USERACCESS.EMPID=" & UserID & " AND USERACCESS.MENUINDX=8;"
            MyAcc = GetFieldValue(Qry)
            If MyAcc <> "FULL ACCESS" Then
                MsgBox "You don't have full access to this transaction." & vbCrLf & "Please contact your System Administrator", vbExclamation + vbOKOnly, "Access Denied"
                Exit Sub
            End If
        
            Call LockData(False)
            Call ClearEntries
            Call tbSaveCancel(True)
            AddNew = True
            txtEntry(1).SetFocus
        
        Case 1  '"edit"
            'Access
            'Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE (((USERACCESS.EMPID)=" & UserID & ") AND ((USERACCESS.MENUINDX)=8));"
            'mySQL
            Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE USERACCESS.EMPID=" & UserID & " AND USERACCESS.MENUINDX=8;"
            
            MyAcc = GetFieldValue(Qry)
            If MyAcc <> "FULL ACCESS" Then
                MsgBox "You don't have full access to this transaction." & vbCrLf & "Please contact your System Administrator", vbExclamation + vbOKOnly, "Access Denied"
                Exit Sub
            End If
            
            If txtEntry(0).Text = "" Then
                MsgBox "No records found!", vbExclamation + vbOKOnly, "Warning"
                txtEntry(1).SetFocus
                Exit Sub
            End If
            Call LockData(False)
            Call tbSaveCancel(True)
            AddNew = False
            txtEntry(1).SetFocus
        Case 6  '"save"
            vMsg = ValidateEntry
            If vMsg <> "" Then
                MsgBox vMsg, vbOKOnly + vbCritical, "Validation Error"
                txtEntry(1).SetFocus
                Exit Sub
            End If
            Ans = MsgBox("Do you want to save employee?", vbQuestion + vbYesNo, "Save")
            If Ans = vbNo Then
                txtEntry(1).SetFocus
                Exit Sub
            End If
            Call SaveAdmin
            MsgBox "Record successfully saved!", vbInformation + vbOKOnly, "Save"
            Call tbSaveCancel(False)
            Call LockData(True)
        Case 7  '"cancel"
            Ans = MsgBox("Are you sure you want to cancel?", vbQuestion + vbYesNo, "Cancel")
            If Ans = vbNo Then Exit Sub
            If AddNew Then
                Call ClearEntries
            Else
                EmployeeID = txtEntry(0).Text
                Call ClearEntries
                Call ShowRecord(CStr(EmployeeID))
            End If
            Call LockData(True)
            Call tbSaveCancel(False)
            txtEntry(1).SetFocus
        Case 2  '"search"
            vMsg = ""
            For Ans = 0 To cboPosition.ListCount - 1
                vMsg = vMsg & "(EMPLOYEES.Position)='" & cboPosition.List(Ans) & "' Or "
            Next Ans
            vMsg = "((" & Left$(vMsg, Len(vMsg) - 3) & ")) "
            
            Set frmSearch.FrmSource = Employees
            'Access
            'QrySrch = "SELECT EMPLOYEES.EMPID, [LASTNAME] & ', ' & [FIRSTNAME] & ' ' & Left$([MIDDLENAME],1) & '.' AS EMPNAME " & _
                      "From Employees ORDER BY [LASTNAME] & ', ' & [FIRSTNAME] & ' ' & Left$([MIDDLENAME],1) & '.';"
            'mySQL
            QrySrch = "SELECT EMPLOYEES.EMPID, concat(EMPLOYEES.LASTNAME,', ',EMPLOYEES.FIRSTNAME,' ',IF(ISNULL(LEFT(EMPLOYEES.MIDDLENAME,1)),'',CONCAT(LEFT(EMPLOYEES.MIDDLENAME,1), '.'))) AS EMPNAME" & _
                    " From Employees ORDER BY concat(EMPLOYEES.LASTNAME,', ',EMPLOYEES.FIRSTNAME,' ',IF(ISNULL(LEFT(EMPLOYEES.MIDDLENAME,1)),'',CONCAT(LEFT(EMPLOYEES.MIDDLENAME,1), '.')))"

            frmSearch.QrySearch = QrySrch
            frmSearch.lblCap.Caption = "Search for employees"
            frmSearch.Show
            Me.Enabled = False
        Case 3   '"view"
            MsgBox ("No print Preview"), vbExclamation, "Employee"
            'EmpSummary.MySource = "Admin"
            'EmpSummary.Show
            'Unload Me
        Case 4  'Print
            MsgBox "Print feature not available!", vbExclamation + vbOKOnly, "Print"
        Case 5     '"quit"
            'MotoRep.tbMain.Visible = True
            MotoRep.toolBarButton.Visible = True
            ''MotoRep.imgBG.Visible = True
            Unload Me
    End Select

End Sub

Private Sub Form_Load()
    Dim rsPos As New ADODB.Recordset
    Dim rsStat As New ADODB.Recordset
    Dim Qry As String, i As Integer
    
    'Access
    'Get CNC position from setting then load to combo position
    'Qry = "SELECT SETTINGS.*, SETTINGS.DESCRIPTION From SETTINGS " & _
          "WHERE (((SETTINGS.DESCRIPTION)='ADM POSITION'));"
    'mySQL
    Qry = "SELECT SETTINGS.*, SETTINGS.DESCRIPTION From SETTINGS " & _
          "WHERE SETTINGS.DESCRIPTION='ADM POSITION';"
    
    rsPos.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    With rsPos
        If .RecordCount <> 0 Then
            .MoveFirst
            Do While Not .EOF
                cboPosition.AddItem .Fields(1)
                .MoveNext
            Loop
        End If
        .Close
    End With
    'Access
    'Get Employee status from settings then load to combo status
    'Qry = "SELECT SETTINGS.*, SETTINGS.DESCRIPTION From SETTINGS " & _
          "WHERE (((SETTINGS.DESCRIPTION)='EMP STATUS'));"
    'mySQL
    Qry = "SELECT SETTINGS.*, SETTINGS.DESCRIPTION From SETTINGS " & _
          "WHERE SETTINGS.DESCRIPTION='EMP STATUS';"
    
    rsStat.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    With rsStat
        If .RecordCount <> 0 Then
            .MoveFirst
            Do While Not .EOF
                cboStatus.AddItem .Fields(1)
                .MoveNext
            Loop
        End If
        .Close
    End With
    With cboCivilStat
        .Clear
        .AddItem "Single"
        .AddItem "Married"
        .AddItem "Separated"
        .AddItem "Widow/Widower"
    End With
    With cboGender
        .Clear
        .AddItem "Male"
        .AddItem "Female"
    End With
    cboDept.AddItem "SPARE PARTS"
    cboDept.AddItem "SERVICE"
    cboDept.AddItem "2H"
    cboDept.AddItem "ADMIN"
    ClearEntries
    Call LockData(True)
    Call tbSaveCancel(False)
    MotoRep.sBar.Panels(6).Text = "Employee's Profile entry form"
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    MotoRep.sBar.Panels(6).Text = ""
End Sub

Private Sub Form_Resize()
    Me.WindowState = 2
End Sub




Function EmployeeExist() As Boolean
    Dim rsEmp As New ADODB.Recordset
    Dim Qry As String
    'Access
    'Qry = "SELECT EMPLOYEES.LASTNAME, EMPLOYEES.FIRSTNAME, EMPLOYEES.MIDDLENAME " & _
         "From EMPLOYEES " & _
         "WHERE (((EMPLOYEES.LASTNAME)='" & txtEntry(1).Text & "') AND ((EMPLOYEES.FIRSTNAME)='" & txtEntry(2).Text & "') AND ((EMPLOYEES.MIDDLENAME)='" & txtEntry(3).Text & "'));"
    'mySQl
    Qry = "SELECT EMPLOYEES.LASTNAME, EMPLOYEES.FIRSTNAME, EMPLOYEES.MIDDLENAME " & _
         "From EMPLOYEES " & _
         "WHERE EMPLOYEES.LASTNAME='" & txtEntry(1).Text & "' AND EMPLOYEES.FIRSTNAME='" & txtEntry(2).Text & "' AND EMPLOYEES.MIDDLENAME='" & txtEntry(3).Text & "';"
    
    With rsEmp
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            EmployeeExist = True
        Else
            EmployeeExist = False
        End If
        .Close
    End With
End Function





Private Sub RXResizer1_ResizeComplete()
    lblHeading(2).Left = Frame4.Left + Frame4.Width - lblHeading(2).Width
    lblHeading(3).Left = lblHeading(2).Left + 30
End Sub

Private Sub txtEntry_GotFocus(Index As Integer)
    If Index = 9 Then
        If cboCivilStat.Text = "Single" Then
            txtEntry(9).Locked = True
            txtEntry(10).Locked = True
        Else
            txtEntry(9).Locked = False
            txtEntry(10).Locked = False
        End If
    End If
End Sub

Private Sub txtEntry_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtEntry_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 40 Then
        SendKeys "{TAB}"
    ElseIf KeyCode = 38 Then
        SendKeys "+{TAB}"
    End If
End Sub



