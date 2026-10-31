VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form SPInsXst 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Insert Existing Spare Parts"
   ClientHeight    =   7770
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   16485
   FillColor       =   &H000000FF&
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "SPInsXst.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7770
   ScaleWidth      =   16485
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   350
      Left            =   9810
      Top             =   135
   End
   Begin MSComCtl2.FlatScrollBar fScrol 
      Height          =   195
      Left            =   6075
      TabIndex        =   54
      ToolTipText     =   "Click to Select Branch"
      Top             =   1260
      Width           =   510
      _ExtentX        =   900
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   2
      Arrows          =   65536
      Min             =   1
      Max             =   3
      Orientation     =   1245185
      Value           =   1
   End
   Begin VB.Frame fmAnotation 
      Caption         =   "ANNOTATION"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   915
      Left            =   45
      TabIndex        =   41
      Top             =   6795
      Width           =   16350
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ST = "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   6255
         TabIndex        =   57
         Top             =   585
         Width           =   435
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Stock Transfer"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   6795
         TabIndex        =   56
         Top             =   270
         Width           =   1245
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ST = "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   6345
         TabIndex        =   55
         Top             =   270
         Width           =   435
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Quantity ON-HAND"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   1485
         TabIndex        =   53
         Top             =   270
         Width           =   1545
      End
      Begin VB.Label Label6 
         Caption         =   " Supplier BACK-ORDER"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1440
         TabIndex        =   52
         Top             =   585
         Width           =   2220
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Sales Quantity"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   4365
         TabIndex        =   51
         Top             =   270
         Width           =   1230
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Quantity ON-HAND"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   4365
         TabIndex        =   50
         Top             =   585
         Width           =   1545
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Purchase Order"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   11610
         TabIndex        =   49
         Top             =   270
         Width           =   1305
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Quantity Taking"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   11610
         TabIndex        =   48
         Top             =   585
         Width           =   1350
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "QH ="
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   900
         TabIndex        =   47
         Top             =   270
         Width           =   420
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "BO ="
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   900
         TabIndex        =   46
         Top             =   585
         Width           =   405
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SQ = "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   3870
         TabIndex        =   45
         Top             =   270
         Width           =   450
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "QH = "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   3870
         TabIndex        =   44
         Top             =   585
         Width           =   465
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PO = "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   11115
         TabIndex        =   43
         Top             =   270
         Width           =   450
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "QT ="
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   11115
         TabIndex        =   42
         Top             =   585
         Width           =   405
      End
   End
   Begin MSComCtl2.UpDown UpDRef 
      Height          =   315
      Left            =   1800
      TabIndex        =   37
      Top             =   240
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   556
      _Version        =   393216
      Max             =   1
      Enabled         =   -1  'True
   End
   Begin MSComctlLib.ProgressBar Pbar 
      Height          =   255
      Left            =   45
      TabIndex        =   21
      Top             =   6450
      Visible         =   0   'False
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.TextBox txtOrder 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3960
      TabIndex        =   20
      Top             =   6360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin lvButton.lvButtons_H cmdReplenish 
      Height          =   495
      Left            =   13860
      TabIndex        =   19
      Top             =   135
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   873
      Caption         =   "Stocks for Replenish"
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
      Image           =   "SPInsXst.frx":058A
      cBack           =   -2147483633
   End
   Begin VB.TextBox txtSuppCode 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   -120
      TabIndex        =   18
      Top             =   6840
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.ComboBox cboFind 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2040
      TabIndex        =   0
      Top             =   240
      Width           =   4335
   End
   Begin lvButton.lvButtons_H cmdInsert 
      Height          =   375
      Left            =   12195
      TabIndex        =   7
      Top             =   6330
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "&Insert"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
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
      Image           =   "SPInsXst.frx":0B24
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdOptHide 
      Height          =   375
      Left            =   3600
      TabIndex        =   5
      Top             =   600
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      Caption         =   "&Hide Option"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
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
      Image           =   "SPInsXst.frx":10BE
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdOptShow 
      Height          =   375
      Left            =   480
      TabIndex        =   4
      Top             =   600
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      Caption         =   "&Show Option"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
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
      Image           =   "SPInsXst.frx":1658
      cBack           =   -2147483633
   End
   Begin VB.Frame FrameOpt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   45
      TabIndex        =   3
      Top             =   720
      Width           =   16365
      Begin VB.ComboBox cbomos 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "SPInsXst.frx":1BF2
         Left            =   6690
         List            =   "SPInsXst.frx":1BF4
         TabIndex        =   38
         Tag             =   "XYF"
         Top             =   720
         Width           =   615
      End
      Begin MSComCtl2.UpDown OptrSQ 
         Height          =   285
         Left            =   13320
         TabIndex        =   36
         Top             =   255
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   503
         _Version        =   393216
         Value           =   1
         Max             =   5
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown OptrBO 
         Height          =   285
         Left            =   11040
         TabIndex        =   35
         Top             =   615
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   503
         _Version        =   393216
         Max             =   5
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtQH 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   11295
         TabIndex        =   32
         Text            =   "0"
         Top             =   255
         Width           =   480
      End
      Begin MSComCtl2.UpDown updBO 
         Height          =   285
         Left            =   11760
         TabIndex        =   28
         Top             =   615
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   503
         _Version        =   393216
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtBO"
         BuddyDispid     =   196629
         OrigLeft        =   9840
         OrigTop         =   600
         OrigRight       =   10095
         OrigBottom      =   855
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtBO 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   11280
         TabIndex        =   27
         Text            =   "0"
         Top             =   615
         Width           =   480
      End
      Begin MSComCtl2.UpDown updSQ 
         Height          =   285
         Left            =   14040
         TabIndex        =   26
         Top             =   255
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   503
         _Version        =   393216
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtSQ"
         BuddyDispid     =   196630
         OrigLeft        =   11280
         OrigTop         =   255
         OrigRight       =   11535
         OrigBottom      =   540
         Max             =   1000
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtSQ 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   13560
         TabIndex        =   25
         Text            =   "0"
         Top             =   255
         Width           =   480
      End
      Begin MSComCtl2.DTPicker SDfrom 
         Height          =   315
         Left            =   1890
         TabIndex        =   15
         Tag             =   "XYF"
         Top             =   720
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   65798145
         CurrentDate     =   39498
      End
      Begin VB.CheckBox chkAllSupp 
         Caption         =   "All Suppliers"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5010
         TabIndex        =   13
         Tag             =   "XYF"
         Top             =   360
         Width           =   1335
      End
      Begin VB.ComboBox cboSupplier 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1890
         TabIndex        =   12
         Tag             =   "XYF"
         Top             =   360
         Width           =   3015
      End
      Begin MSComCtl2.DTPicker SDto 
         Height          =   315
         Left            =   5010
         TabIndex        =   17
         Tag             =   "XYF"
         Top             =   720
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   65798145
         CurrentDate     =   39498
      End
      Begin MSComCtl2.UpDown OptrQH 
         Height          =   285
         Left            =   11040
         TabIndex        =   31
         Top             =   255
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   503
         _Version        =   393216
         Max             =   5
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown updQH 
         Height          =   285
         Left            =   11760
         TabIndex        =   33
         Top             =   255
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   503
         _Version        =   393216
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtQH"
         BuddyDispid     =   196628
         OrigLeft        =   9015
         OrigTop         =   255
         OrigRight       =   9270
         OrigBottom      =   540
         Max             =   1000
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H008080FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Mos."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6780
         TabIndex        =   39
         Tag             =   "XYF"
         Top             =   360
         Width           =   435
      End
      Begin VB.Label lblOprQH 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "="
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   10560
         TabIndex        =   34
         Top             =   240
         Width           =   480
      End
      Begin VB.Label lblOprSQ 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   ">"
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   12840
         TabIndex        =   30
         Top             =   255
         Width           =   480
      End
      Begin VB.Label lblOprBO 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "="
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   10560
         TabIndex        =   29
         Top             =   615
         Width           =   480
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "SQ:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   12360
         TabIndex        =   24
         ToolTipText     =   "SALES QTY."
         Top             =   270
         Width           =   375
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "BO:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   10080
         TabIndex        =   23
         ToolTipText     =   "BACK ORDER"
         Top             =   630
         Width           =   375
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "QH:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   10080
         TabIndex        =   22
         ToolTipText     =   "QTY-ON-HAND"
         Top             =   270
         Width           =   375
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Sales Date To:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3570
         TabIndex        =   16
         Tag             =   "XYF"
         Top             =   750
         Width           =   1335
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Sales Date From:"
         BeginProperty Font 
            Name            =   "Tahoma"
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
         Top             =   750
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Supplier:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
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
         Top             =   390
         Width           =   1215
      End
   End
   Begin lvButton.lvButtons_H cmdFind 
      Height          =   495
      Left            =   6435
      TabIndex        =   2
      Top             =   135
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   4
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
      ImgAlign        =   4
      Image           =   "SPInsXst.frx":1BF6
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdOK 
      Height          =   375
      Left            =   13635
      TabIndex        =   8
      Top             =   6330
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "&OK"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
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
      Image           =   "SPInsXst.frx":2190
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdCancel 
      Height          =   375
      Left            =   15075
      TabIndex        =   9
      Top             =   6330
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "&Cancel"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
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
      Image           =   "SPInsXst.frx":272A
      cBack           =   -2147483633
   End
   Begin MSFlexGridLib.MSFlexGrid flxSP 
      Height          =   3945
      Left            =   45
      TabIndex        =   10
      Top             =   2340
      Width           =   16365
      _ExtentX        =   28866
      _ExtentY        =   6959
      _Version        =   393216
      Rows            =   1
      Cols            =   29
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   300
      BackColorBkg    =   16777215
      FocusRect       =   2
      FillStyle       =   1
      ScrollBars      =   2
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid flxSPH 
      Height          =   510
      Left            =   45
      TabIndex        =   40
      Top             =   1845
      Width           =   16365
      _ExtentX        =   28866
      _ExtentY        =   900
      _Version        =   393216
      Cols            =   27
      FixedRows       =   2
      ScrollBars      =   0
      AllowUserResizing=   3
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblChk 
      Caption         =   "ü"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   90
      TabIndex        =   11
      Top             =   6300
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblRefSearch 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FIND PARTS DESC."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   180
      TabIndex        =   1
      Top             =   315
      Width           =   1485
   End
   Begin VB.Menu mnuSP 
      Caption         =   "SP"
      Visible         =   0   'False
      Begin VB.Menu mnuAct 
         Caption         =   "set to ""ACTIVE"""
      End
      Begin VB.Menu mnuInAct 
         Caption         =   "set to ""IN-ACTIVE"""
      End
      Begin VB.Menu sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMark 
         Caption         =   "Mark"
      End
      Begin VB.Menu mnuUnmark 
         Caption         =   "Unmark"
      End
      Begin VB.Menu sep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMarkAll 
         Caption         =   "Mark All"
      End
      Begin VB.Menu mnuUnmarkAll 
         Caption         =   "UnMark All"
      End
   End
End
Attribute VB_Name = "SPInsXst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsSupplier As New ADODB.Recordset
Dim ShowOpt As Boolean, SuppSrch As Boolean
Dim ActRow As Integer, ActCol As Integer, PrvW As Integer
Public MySource As String, RwSel As Boolean
Private Sub cboFind_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call cmdFind_Click
End Sub

Private Sub cbomos_Change()
    cbomos_Click
End Sub
Private Sub cbomos_Click()
Dim mos As Integer
    mos = Val(cboMos.Text)
    pmos = "-" & mos
    SDfrom.Value = DateAdd("m", pmos, Date)
   
    If cboMos.Text = "All" Then
        SDfrom.Value = SPordering.SalesDateFrom.Value
    End If
End Sub
Private Sub cbomos_DblClick()
    cbomos_Click
End Sub
Private Sub cbomos_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then Exit Sub
    If KeyAscii > 54 Or KeyAscii < 48 Then
        KeyAscii = 0
    End If
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
            txtSuppCode.Text = IIf(IsNull(rsSupplier!SUPPLIERCODE), "", rsSupplier!SUPPLIERCODE)
            chkAllSupp.Value = 0
            j = Len(strTotal) - Len(strPartial)
            If j <> 0 Then
                SuppSrch = True
                .SelText = Right$(strTotal, j)
                .SelStart = Len(strPartial)
                .SelLength = j
            End If
        Else
            txtSuppCode.Text = ""
        End If
    End With

End Sub

Private Sub cboSupplier_Click()
    Call cboSupplier_Change
End Sub

Private Sub cboSupplier_GotFocus()
   Dim Tmp
    
   Const CB_SHOWDROPDOWN = &H14F
   Tmp = SendMessage(cboSupplier.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub

Private Sub cboSupplier_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then SuppSrch = True
End Sub

Private Sub cboSupplier_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then cboFind.SetFocus
End Sub

Private Sub chkAllSupp_Click()
    If chkAllSupp.Value = 1 Then 'All Supplier
        cboSupplier.Text = ""
        txtSuppCode.Text = ""
    Else
        If txtSuppCode.Text = "" Then
            cboSupplier.Text = SPordering.cboSupplier.Text
            txtSuppCode.Text = SPordering.txtScode.Text
        End If
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    Dim rsSP As New ADODB.Recordset
   
    
    Dim MySP() As String, Mdl() As String, Qry As String
    Dim myW As String, myW1 As String, myW2 As String, myW3 As String
    Dim MySel As String, i As Long, r As Long, rM As Long
    
    If cboFind.Text = "" Then Exit Sub
    MySP = Split(cboFind.Text, " ")
    myW1 = ""
    For i = LBound(MySP) To UBound(MySP)
        If myW1 = "" Then
            'myW1 = "(SPAREPARTS.DESCRIPTION) Like '%" & MySP(i) & "%'"
            
            If UpDRef.Value = 0 Then
                myW1 = "(SPAREPARTS.PARTNUMBER) Like '%" & MySP(i) & "%'"
            Else
                myW1 = "(SPAREPARTS.DESCRIPTION) Like '%" & MySP(i) & "%'"
            End If
        Else
            'myW1 = myW1 & " AND (SPAREPARTS.DESCRIPTION) Like '%" & MySP(i) & "%'"
            If UpDRef.Value = 0 Then
                myW1 = myW1 & " AND (SPAREPARTS.PARTNUMBER) Like '%" & MySP(i) & "%'"
            Else
                myW1 = myW1 & " AND (SPAREPARTS.DESCRIPTION) Like '%" & MySP(i) & "%'"
            End If
        End If
    Next i
    myW = myW1
    If txtSuppCode.Text <> "" Then
        myW = "(" & myW & ") AND ((SPAREPARTS.SUPPLIERCODE) = '" & txtSuppCode.Text & "')"
    End If
   Call CollAlign
   MySel = "SELECT  SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPAREPARTS.QTYONHAND, SPAREPARTS.BACKORDER, SPAREPARTS.SALES, SPAREPARTS.COST, SPAREPARTS.SPAREPARTID, SPAREPARTS.APPLICMODEL, SPAREPARTS.DESCRIPTION2 AS MYSTAT From SPAREPARTS WHERE (" & myW & ")  ORDER BY SPAREPARTS.PARTNUMBER;"
    With rsSP
        .Open MySel, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            flxSP.Rows = 0
            pBar.Value = 0
            pBar.Visible = True
            pBar.Max = .RecordCount * 2
            'Do While Not .EOF
                'flxSP.AddItem vbTab & !PartNumber & vbTab & !Description & vbTab & !QtyOnHand & vbTab & IIf(IsNull(!BACKORDER), "", !BACKORDER) & vbTab & IIf(IsNull(!Sales), "", !Sales) & vbTab & vbTab & Format(!Cost, "#,##0.00") & vbTab & vbTab & !SPAREPARTID & vbTab & IIf(IsNull(!ApplicModel), "", !ApplicModel) & vbTab & IIf(IsNull(!MyStat), "", !MyStat)
                'flxSP.AddItem vbTab & !PartNumber & vbTab & !Description & vbTab & !QtyOnHand & vbTab & IIf(IsNull(!BACKORDER), "", !BACKORDER) & vbTab & IIf(IsNull(!Sales), "", !Sales) & vbTab & vbTab & vbTab & vbTab & vbTab & !SPAREPARTID & vbTab & Format(!Cost, "#,##0.00") & vbTab & IIf(IsNull(!ApplicModel), "", !ApplicModel) & vbTab & IIf(IsNull(!MyStat), "", !MyStat)
                'Desinged for Branch P.O
            '    flxSP.AddItem vbTab & !PartNumber & vbTab & !Description & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & !QtyOnHand & vbTab & IIf(IsNull(!BACKORDER), "", !BACKORDER) & vbTab & IIf(IsNull(!Sales), "", !Sales) & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & !SPAREPARTID & vbTab & Format(!Cost, "#,##0.00") & vbTab & IIf(IsNull(!ApplicModel), "", !ApplicModel) & vbTab & IIf(IsNull(!MyStat), "", !MyStat)
            '    Pbar.Value = Pbar.Value + 1
                
                'Hi-lights Column by rows
            '    flxSP.Col = 3
            '    flxSP.Row = Pbar.Value - 1
            '    flxSP.ColSel = 3
            '    flxSP.ColSel = 4
            '    flxSP.ColSel = 5
            '   flxSP.RowSel = Pbar.Value - 1
            '    flxSP.FillStyle = flexFillRepeat
            '    flxSP.CellBackColor = &HC0FFC0
            '
            '    .MoveNext
            '    DoEvents
            'Loop
            
            .MoveFirst
            For rM = 1 To .RecordCount
                    flxSP.Rows = rM
                    flxSP.TextMatrix(rM - 1, 1) = !PartNumber
                    flxSP.TextMatrix(rM - 1, 2) = !Description
                    flxSP.TextMatrix(rM - 1, 19) = !QtyOnHand
                    flxSP.TextMatrix(rM - 1, 20) = IIf(IsNull(!BACKORDER), "", !BACKORDER)
                    'flxSP.TextMatrix(rM - 1, 21) = IIf(IsNull(!Sales), "", !Sales)
                    flxSP.TextMatrix(rM - 1, 25) = !SPAREPARTID
                    flxSP.TextMatrix(rM - 1, 26) = Format(!Cost, "#,##0.00")
                    flxSP.TextMatrix(rM - 1, 27) = IIf(IsNull(!ApplicModel), "", !ApplicModel)
                    flxSP.TextMatrix(rM - 1, 28) = IIf(IsNull(!MyStat), "", !MyStat)
                    pBar.Value = pBar.Value + 1
                .MoveNext
            Next rM
            
            flxSP.Visible = True
            If ShowOpt Then
                'flxSP.Top = 1920
                flxSPH.Top = FrameOpt.Top + FrameOpt.Height
                flxSP.Top = FrameOpt.Top + FrameOpt.Height + flxSPH.Height
            Else
                'flxSP.Top = 1080
                flxSPH.Top = cmdOptHide.Top + cmdOptHide.Height + 100
                flxSP.Top = FrameOpt.Height + flxSPH.Height
            End If
            cmdInsert.Top = flxSP.Top + flxSP.Height + 120
            cmdOk.Top = flxSP.Top + flxSP.Height + 120
            cmdCancel.Top = flxSP.Top + flxSP.Height + 120
            pBar.Top = flxSP.Top + flxSP.Height
            Me.Height = fmAnotation.Height + cmdOk.Top + cmdOk.Height + 580
        Else
        '    With flxSP
                flxSP.Clear
                flxSP.Rows = 1
        '        .FormatString = "  |<PART NUMBER|<DESCRIPTION|^QH |^BO|^SQ|^ST|>COST |>ORDER||<APPLICABLE MODEL"
        '        .ColWidth(0) = 300
        '        .ColWidth(1) = (.Width - 400) * 0.15
        '       .ColWidth(2) = (.Width - 400) * 0.33
        '        .ColWidth(3) = (.Width - 400) * 0.05
        '        .ColWidth(4) = (.Width - 400) * 0.05
        '        .ColWidth(5) = (.Width - 400) * 0.05
        '        .ColWidth(6) = (.Width - 400) * 0.05
        '        .ColWidth(7) = (.Width - 400) * 0.05
        '        .ColWidth(8) = (.Width - 400) * 0.05
        '        .ColWidth(9) = (.Width - 400) * 0.05
        '        '.ColWidth(10) = 0
        '        .ColWidth(11) = (.Width - 400) * 0.05
        '        .ColWidth(12) = (.Width - 400) * 0.1
        '        .ColWidth(13) = 0
        '        .Visible = True
        '    End With
            
        End If
        .Close
    End With
    
    If flxSP.Rows = 1 And flxSP.TextMatrix(0, 1) = "" Then Exit Sub
    With flxSP
        '.ColWidth(10) = 0
        '.ColWidth(13) = 0
        'Previous Code
        'For i = 0 To .Rows - 1
        '    pBar.Value = IIf(pBar.Value = pBar.Max, pBar.Max, pBar.Value + 1)
            'Design for Branch
            'Get Sales of this spareparts
        '    .TextMatrix(i, 8) = GetSalesQty(Val(.TextMatrix(i, 13)))
            'Qry = "SELECT Sum(SPITR.OUT) AS SumOfOUT From SPITR " & _
                  "WHERE (((SPITR.TRANSDATE)>=#" & SDfrom.Value & "# And (SPITR.TRANSDATE)<=#" & SDto.Value & "#) AND ((SPITR.TRANSTYPE)='OUT') AND ((SPITR.REMARKS)='TRANSFER') AND ((SPITR.SPAREPARTID)=" & Val(.TextMatrix(i, 9)) & "));"
            'mySQL
        '    Qry = "SELECT Sum(SPITR.OUT) AS SumOfOUT From SPITR " & _
                  "WHERE SPITR.TRANSDATE>='" & Format(SDfrom.Value, "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(SDto.Value, "yyyy-mm-dd") & "' AND SPITR.TRANSTYPE='OUT' AND SPITR.REMARKS='TRANSFER' AND SPITR.SPAREPARTID=" & Val(.TextMatrix(i, 13)) & ";"
            'Design for Branch
        '    .TextMatrix(i, 9) = GetFieldValue(Qry)
            'Branch Designed
        '    If .TextMatrix(i, 16) = "INACTIVE" Then
        '        .Col = 1: .Row = i: .ColSel = 16: .RowSel = i: .FillStyle = flexFillRepeat: .CellForeColor = RGB(255, 0, 0)
        '    End If
        'Next i
        
        For i = .Rows - 1 To 0 Step -1
            'flxSP.TextMatrix(i, 5) = GetSalesQty(Val(flxSP.TextMatrix(i, 10)))
            'Main Branch
            .TextMatrix(i, 21) = GetSalesQty(Val(.TextMatrix(i, 25)))
            
            'Qry = "SELECT Sum(SPITR.OUT) AS SumOfOUT From SPITR " & _
                  "WHERE (((SPITR.TRANSDATE)>=#" & SDfrom.Value & "# And (SPITR.TRANSDATE)<=#" & SDto.Value & "#) AND ((SPITR.TRANSTYPE)='OUT') AND ((SPITR.REMARKS)='TRANSFER') AND ((SPITR.SPAREPARTID)=" & Val(flxSP.TextMatrix(i, 9)) & "));"
            
            '---------------------------------------------------------
            '** Stock Transfer
            'Moto-r Ulas
            'Qry = "SELECT Sum(SPITR.OUT) AS SumOfOUT From SPITR " & _
                  "WHERE SPITR.TRANSDATE>='" & Format(SDfrom.Value, "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(SDto.Value, "yyyy-mm-dd") & "' AND SPITR.TRANSTYPE='OUT' AND SPITR.REMARKS='TRANSFER' AND SPITR.SPAREPARTID=" & Val(.TextMatrix(i, 28)) & " AND SPITR.SOURCE='MOTO-R ULAS';"
            '.TextMatrix(i, 22) = Val(GetFieldValue(Qry))
            'Yamaha Motorhub
            'Qry = "SELECT Sum(SPITR.OUT) AS SumOfOUT From SPITR " & _
                  "WHERE SPITR.TRANSDATE>='" & Format(SDfrom.Value, "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(SDto.Value, "yyyy-mm-dd") & "' AND SPITR.TRANSTYPE='OUT' AND SPITR.REMARKS='TRANSFER' AND SPITR.SPAREPARTID=" & Val(.TextMatrix(i, 28)) & " AND SPITR.SOURCE='YAMAHA MOTORHUB';"
            '.TextMatrix(i, 23) = Val(GetFieldValue(Qry))
            
            'MC Pasrts Depot
            'Qry = "SELECT Sum(SPITR.OUT) AS SumOfOUT From SPITR " & _
                  "WHERE SPITR.TRANSDATE>='" & Format(SDfrom.Value, "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(SDto.Value, "yyyy-mm-dd") & "' AND SPITR.TRANSTYPE='OUT' AND SPITR.REMARKS='TRANSFER' AND SPITR.SPAREPARTID=" & Val(.TextMatrix(i, 28)) & " AND SPITR.SOURCE='MC PARTS DEPOT CORP.';"
            '.TextMatrix(i, 24) = Val(GetFieldValue(Qry))
            '---------------------------------------------------------
            
            
            pBar.Value = IIf(pBar.Value = pBar.Max, pBar.Max, pBar.Value + 1)
            DoEvents
            If .TextMatrix(i, 28) = "INACTIVE" Then
                .Col = 1: .Row = i: .ColSel = 28: .RowSel = i: .FillStyle = flexFillRepeat: .CellForeColor = RGB(255, 0, 0)
            End If
        Next i
        
    End With
    
    'Pbar.Visible = False
    Call AddInCombo
    'Previous Code
    'pBar.Value = 0
    'pBar.Max = flxSP.Rows
    'For r = 0 To flxSP.Rows - 1
    
    '    pBar.Value = pBar.Value + 1
    '    flxSP.TextMatrix(r, 10) = GetSalesQty6mos(Val(flxSP.TextMatrix(r, 13)))
    '    spSOQ = Val(flxSP.TextMatrix(r, 10)) - Val(flxSP.TextMatrix(r, 6)) - Val(flxSP.TextMatrix(r, 7))
    '    flxSP.TextMatrix(r, 11) = IIf(spSOQ < 0, 0, spSOQ)
    '    Qry = ""
    '    Qry = "SELECT spareparts.QTYONHAND From spareparts Where spareparts.SPAREPARTID = " & flxSP.TextMatrix(r, 13) & ";"
    '    flxSP.TextMatrix(r, 3) = Val(GetFieldValueMain(Qry))
    '    Qry = ""
    '    Qry = "SELECT spareparts.BACKORDER From spareparts Where spareparts.SPAREPARTID = " & flxSP.TextMatrix(r, 13) & ";"
    '    flxSP.TextMatrix(r, 4) = Val(GetFieldValueMain(Qry))
        
    '    flxSP.TextMatrix(r, 5) = GetSalesQtyMain(Val(flxSP.TextMatrix(r, 13)))
    
    'Next r
    pBar.Value = 0
    pBar.Max = flxSP.Rows
    For r = 0 To flxSP.Rows - 1
        pBar.Value = pBar.Value + 1
        
        '**Query Branch SKU
        'Moto-r Ulas
        Qry = ""
        Qry = "SELECT spareparts.QTYONHAND From spareparts Where spareparts.SPAREPARTID = " & flxSP.TextMatrix(r, 25) & ";"
        flxSP.TextMatrix(r, 3) = Val(GetFieldValueMU(Qry))
        Qry = ""
        Qry = "SELECT spareparts.BACKORDER From spareparts Where spareparts.SPAREPARTID = " & flxSP.TextMatrix(r, 25) & ";"
        flxSP.TextMatrix(r, 4) = Val(GetFieldValueMU(Qry))
        flxSP.TextMatrix(r, 5) = GetSalesQtyBrchMU(Val(flxSP.TextMatrix(r, 25)))
        Qry = "SELECT Sum(SPITR.OUT) AS SumOfOUT From SPITR " & _
              "WHERE SPITR.TRANSDATE>='" & Format(SDfrom.Value, "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(SDto.Value, "yyyy-mm-dd") & "' AND SPITR.TRANSTYPE='OUT' AND SPITR.REMARKS='TRANSFER' AND SPITR.SPAREPARTID=" & Val(flxSP.TextMatrix(r, 25)) & " AND SPITR.SOURCE='MOTO-R ULAS';"
        flxSP.TextMatrix(r, 6) = Val(GetFieldValueMU(Qry))
        
        'Yamaha MotorHub
        Qry = ""
        Qry = "SELECT spareparts.QTYONHAND From spareparts Where spareparts.SPAREPARTID = " & flxSP.TextMatrix(r, 25) & ";"
        flxSP.TextMatrix(r, 7) = Val(GetFieldValueYM(Qry))
        Qry = ""
        Qry = "SELECT spareparts.BACKORDER From spareparts Where spareparts.SPAREPARTID = " & flxSP.TextMatrix(r, 25) & ";"
        flxSP.TextMatrix(r, 8) = Val(GetFieldValueYM(Qry))
        flxSP.TextMatrix(r, 9) = GetSalesQtyBrchYM(Val(flxSP.TextMatrix(r, 25)))
        Qry = "SELECT Sum(SPITR.OUT) AS SumOfOUT From SPITR " & _
              "WHERE SPITR.TRANSDATE>='" & Format(SDfrom.Value, "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(SDto.Value, "yyyy-mm-dd") & "' AND SPITR.TRANSTYPE='OUT' AND SPITR.REMARKS='TRANSFER' AND SPITR.SPAREPARTID=" & Val(flxSP.TextMatrix(r, 25)) & " AND SPITR.SOURCE='YAMAHA MOTORHUB';"
        flxSP.TextMatrix(r, 10) = Val(GetFieldValueYM(Qry))
        
        'MC Parts Depot Corporation
        Qry = "SELECT spareparts.QTYONHAND From spareparts Where spareparts.SPAREPARTID = " & flxSP.TextMatrix(r, 25) & ";"
        flxSP.TextMatrix(r, 11) = Val(GetFieldValueMD(Qry))
        Qry = ""
        Qry = "SELECT spareparts.BACKORDER From spareparts Where spareparts.SPAREPARTID = " & flxSP.TextMatrix(r, 25) & ";"
        flxSP.TextMatrix(r, 12) = Val(GetFieldValueMD(Qry))
        flxSP.TextMatrix(r, 13) = GetSalesQtyBrchMD(Val(flxSP.TextMatrix(r, 25)))
        Qry = "SELECT Sum(SPITR.OUT) AS SumOfOUT From SPITR " & _
              "WHERE SPITR.TRANSDATE>='" & Format(SDfrom.Value, "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(SDto.Value, "yyyy-mm-dd") & "' AND SPITR.TRANSTYPE='OUT' AND SPITR.REMARKS='TRANSFER' AND SPITR.SPAREPARTID=" & Val(flxSP.TextMatrix(r, 25)) & " AND SPITR.SOURCE='MC PARTS DEPOT CORP.';"
        flxSP.TextMatrix(r, 14) = Val(GetFieldValueMD(Qry))
       
        'Total Branch
        'SQ
        flxSP.TextMatrix(r, 15) = Val(flxSP.TextMatrix(r, 3)) + Val(flxSP.TextMatrix(r, 7)) + Val(flxSP.TextMatrix(r, 11)) + Val(flxSP.TextMatrix(r, 21))
        'BO
        flxSP.TextMatrix(r, 16) = Val(flxSP.TextMatrix(r, 4)) + Val(flxSP.TextMatrix(r, 8)) + Val(flxSP.TextMatrix(r, 12))
        'SQ
        flxSP.TextMatrix(r, 17) = Val(flxSP.TextMatrix(r, 5)) + Val(flxSP.TextMatrix(r, 9)) + Val(flxSP.TextMatrix(r, 13))
        'ST
        flxSP.TextMatrix(r, 18) = Val(flxSP.TextMatrix(r, 6)) + Val(flxSP.TextMatrix(r, 10)) + Val(flxSP.TextMatrix(r, 14))
        
        '6 months Sales Quantity All Branch
        'Moto-R Main
        
        mrulas6m = GetSalesQty6mosMU(Val(flxSP.TextMatrix(r, 25)))
        ymtgm6m = GetSalesQty6mosYM(Val(flxSP.TextMatrix(r, 25)))
        mcpd6mos = GetSalesQty6mosMD(Val(flxSP.TextMatrix(r, 25)))
        mtrmain6mos = GetSalesQty6mos(Val(flxSP.TextMatrix(r, 25)))
        flxSP.TextMatrix(r, 22) = mrulas6m + ymtgm6m + mcpd6mos + mtrmain6mos
        
        'spSOQ = Val(flxSP.TextMatrix(r, 25)) - Val(flxSP.TextMatrix(r, 15)) - Val(flxSP.TextMatrix(r, 16))
        'flxSP.TextMatrix(r, 26) = IIf(spSOQ < 0, 0, spSOQ)
        
        'randyp                                Branch QTY Total + main                                         Back Order Main Only
        spSOQ = Val(flxSP.TextMatrix(r, 22)) - (Val(flxSP.TextMatrix(r, 15)) + Val(flxSP.TextMatrix(r, 19))) - Val(flxSP.TextMatrix(r, 20))
        flxSP.TextMatrix(r, 23) = IIf(spSOQ < 0, 0, spSOQ)
        
    Next r
    pBar.Visible = False
    'Call FillColorGrd
    'flxSP.Row = 0
    'flxSP.Col = 12
    If flxSP.Rows >= 1 Then
        'If flxSP.TextMatrix(1, 3) <> "" Then
                         'QTY on Hand
        If flxSP.TextMatrix(0, 19) <> "" Then
            flxSP.SetFocus
        Else
            cboFind.SetFocus
        End If
    Else
        cboFind.SetFocus
    End If
 End Sub

Function GetSalesQty(MyID As Long) As Long
    Dim rsQty As New ADODB.Recordset
    Dim Qry As String
    Dim MySales As Long, MyRet As Long
    
    'old Qry = "SELECT Sum(SPITR.OUT) AS SumOfOUT From SPITR " & _
          "WHERE (((SPITR.TRANSDATE)>=#" & SDfrom.Value & "# And (SPITR.TRANSDATE)<=#" & SDto.Value & "#) AND ((SPITR.TRANSTYPE)='OUT') AND ((SPITR.SPAREPARTID)=" & SDto.Value & "));"
    'Qry = "SELECT Sum(SPITR.OUT) AS SumOfOUT From SPITR " & _
          "WHERE (((SPITR.TRANSDATE)>=#" & SDfrom.Value & "# And (SPITR.TRANSDATE)<=#" & SDto.Value & "#) AND ((SPITR.TRANSTYPE)='OUT') AND ((SPITR.REMARKS)<>'TRANSFER') AND ((SPITR.SPAREPARTID)=" & MyID & "));"
    'mySQL
    Qry = "SELECT Sum(SPITR.OUT) AS SumOfOUT From SPITR " & _
          "WHERE SPITR.TRANSDATE>='" & Format(SDfrom.Value, "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(SDto.Value, "yyyy-mm-dd") & "' AND SPITR.TRANSTYPE='OUT' AND SPITR.REMARKS<>'TRANSFER' AND SPITR.SPAREPARTID=" & MyID & ";"
   
    With rsQty
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            MySales = IIf(IsNull(.Fields(0)), 0, .Fields(0))
        Else
            MySales = 0
        End If
        .Close
    End With
    'Qry = "SELECT Sum(SPITR.RECEIVE) AS SumOfRECEIVE From SPITR " & _
          "WHERE (((SPITR.SPAREPARTID)=" & MyID & ") AND ((SPITR.TRANSTYPE)='SALES RETURN') AND ((SPITR.TRANSDATE)>=#" & SDfrom.Value & "# And (SPITR.TRANSDATE)<=#" & SDto.Value & "#));"
    'mySQL
    Qry = "SELECT Sum(SPITR.RECEIVE) AS SumOfRECEIVE From SPITR " & _
          "WHERE SPITR.SPAREPARTID=" & MyID & " AND SPITR.TRANSTYPE='SALES RETURN' AND SPITR.TRANSDATE>='" & Format(SDfrom.Value, "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(SDto.Value, "yyyy-mm-dd") & "';"
  
    MyRet = Val(GetFieldValue(Qry))
    GetSalesQty = MySales - MyRet
End Function

Function GetSalesQty6mos(MyID As Long) As Long
    Dim rsQty As New ADODB.Recordset
    Dim Qry As String
    Dim MySales As Long, MyRet As Long
    
    Qry = "SELECT Sum(SPITR.OUT) AS SumOfOUT From SPITR " & _
          "WHERE SPITR.TRANSDATE>='" & Format(DateAdd("m", -6, Date), "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(Date, "yyyy-mm-dd") & "' AND SPITR.TRANSTYPE='OUT' AND SPITR.REMARKS<>'TRANSFER' AND SPITR.SPAREPARTID=" & MyID & ";"
   
    With rsQty
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            MySales = IIf(IsNull(.Fields(0)), 0, .Fields(0))
        Else
            MySales = 0
        End If
        .Close
    End With
    Qry = "SELECT Sum(SPITR.RECEIVE) AS SumOfRECEIVE From SPITR " & _
          "WHERE SPITR.SPAREPARTID=" & MyID & " AND SPITR.TRANSTYPE='SALES RETURN' AND SPITR.TRANSDATE>='" & Format(DateAdd("m", -6, Date), "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(Date, "yyyy-mm-dd") & "';"
  
    MyRet = Val(GetFieldValue(Qry))
    GetSalesQty6mos = MySales - MyRet
End Function

Function GetSalesQty6mosMU(MyID As Long) As Long
    Dim rsQty As New ADODB.Recordset
    Dim Qry As String
    Dim MySales As Long, MyRet As Long
    
    Qry = "SELECT Sum(SPITR.OUT) AS SumOfOUT From SPITR " & _
          "WHERE SPITR.TRANSDATE>='" & Format(DateAdd("m", -6, Date), "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(Date, "yyyy-mm-dd") & "' AND SPITR.TRANSTYPE='OUT' AND SPITR.REMARKS<>'TRANSFER' AND SPITR.SPAREPARTID=" & MyID & ";"
   
    With rsQty
        .Open Qry, BrchConMU, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            MySales = IIf(IsNull(.Fields(0)), 0, .Fields(0))
        Else
            MySales = 0
        End If
        .Close
    End With
    Qry = "SELECT Sum(SPITR.RECEIVE) AS SumOfRECEIVE From SPITR " & _
          "WHERE SPITR.SPAREPARTID=" & MyID & " AND SPITR.TRANSTYPE='SALES RETURN' AND SPITR.TRANSDATE>='" & Format(DateAdd("m", -6, Date), "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(Date, "yyyy-mm-dd") & "';"
  
    MyRet = Val(GetFieldValueMU(Qry))
    GetSalesQty6mosMU = MySales - MyRet
End Function
Function GetSalesQty6mosYM(MyID As Long) As Long
    Dim rsQty As New ADODB.Recordset
    Dim Qry As String
    Dim MySales As Long, MyRet As Long
    
    Qry = "SELECT Sum(SPITR.OUT) AS SumOfOUT From SPITR " & _
          "WHERE SPITR.TRANSDATE>='" & Format(DateAdd("m", -6, Date), "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(Date, "yyyy-mm-dd") & "' AND SPITR.TRANSTYPE='OUT' AND SPITR.REMARKS<>'TRANSFER' AND SPITR.SPAREPARTID=" & MyID & ";"
   
    With rsQty
        .Open Qry, BrchConYM, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            MySales = IIf(IsNull(.Fields(0)), 0, .Fields(0))
        Else
            MySales = 0
        End If
        .Close
    End With
    Qry = "SELECT Sum(SPITR.RECEIVE) AS SumOfRECEIVE From SPITR " & _
          "WHERE SPITR.SPAREPARTID=" & MyID & " AND SPITR.TRANSTYPE='SALES RETURN' AND SPITR.TRANSDATE>='" & Format(DateAdd("m", -6, Date), "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(Date, "yyyy-mm-dd") & "';"
  
    MyRet = Val(GetFieldValueYM(Qry))
    GetSalesQty6mosYM = MySales - MyRet
End Function
Function GetSalesQty6mosMD(MyID As Long) As Long
    Dim rsQty As New ADODB.Recordset
    Dim Qry As String
    Dim MySales As Long, MyRet As Long
    
    Qry = "SELECT Sum(SPITR.OUT) AS SumOfOUT From SPITR " & _
          "WHERE SPITR.TRANSDATE>='" & Format(DateAdd("m", -6, Date), "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(Date, "yyyy-mm-dd") & "' AND SPITR.TRANSTYPE='OUT' AND SPITR.REMARKS<>'TRANSFER' AND SPITR.SPAREPARTID=" & MyID & ";"
   
    With rsQty
        .Open Qry, BrchConMD, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            MySales = IIf(IsNull(.Fields(0)), 0, .Fields(0))
        Else
            MySales = 0
        End If
        .Close
    End With
    Qry = "SELECT Sum(SPITR.RECEIVE) AS SumOfRECEIVE From SPITR " & _
          "WHERE SPITR.SPAREPARTID=" & MyID & " AND SPITR.TRANSTYPE='SALES RETURN' AND SPITR.TRANSDATE>='" & Format(DateAdd("m", -6, Date), "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(Date, "yyyy-mm-dd") & "';"
  
    MyRet = Val(GetFieldValueMD(Qry))
    GetSalesQty6mosMD = MySales - MyRet
End Function
Function GetSalesQtyMain(MyID As Long) As Long
    Dim rsQty As New ADODB.Recordset
    Dim Qry As String
    Dim MySales As Long, MyRet As Long
    
    Qry = "SELECT Sum(SPITR.OUT) AS SumOfOUT From SPITR " & _
          "WHERE SPITR.TRANSDATE>='" & Format(DateAdd("m", -6, Date), "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(Date, "yyyy-mm-dd") & "' AND SPITR.TRANSTYPE='OUT' AND SPITR.REMARKS<>'TRANSFER' AND SPITR.SPAREPARTID=" & MyID & ";"
   
    With rsQty
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            MySales = IIf(IsNull(.Fields(0)), 0, .Fields(0))
        Else
            MySales = 0
        End If
        .Close
    End With
    Qry = "SELECT Sum(SPITR.RECEIVE) AS SumOfRECEIVE From SPITR " & _
          "WHERE SPITR.SPAREPARTID=" & MyID & " AND SPITR.TRANSTYPE='SALES RETURN' AND SPITR.TRANSDATE>='" & Format(DateAdd("m", -6, Date), "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(Date, "yyyy-mm-dd") & "';"
  
    MyRet = Val(GetFieldValue(Qry))
    GetSalesQtyMain = MySales - MyRet
End Function
Private Sub AddInCombo()
    Dim X As Integer, w As Boolean
    
    If cboFind.Text = "" Then Exit Sub
    w = False
    For X = 0 To cboFind.ListCount - 1
        If cboFind.Text = cboFind.List(X) Then
            If X = 0 Then
                w = True
                Exit For
            Else
                cboFind.RemoveItem X
                Exit For
            End If
        End If
    Next X
    If w = False Then cboFind.AddItem cboFind.Text, 0
End Sub

Private Sub cmdInsert_Click()
    Dim i As Integer
    
    
    If MySource = "SPORDERING" Then
        With SPordering.flxParts
            If .TextMatrix(.Rows - 1, 1) = "" Then  'Remove last row if empty
                If .Rows = 2 Then
                    .Rows = 1
                Else
                    .RemoveItem .Rows - 1
                End If
            End If
            For i = 0 To flxSP.Rows - 1
                If flxSP.TextMatrix(i, 0) <> "" Then
                    If Not SPinFlx(flxSP.TextMatrix(i, 1)) Then
                        '.AddItem .Rows & vbTab & flxSP.TextMatrix(i, 1) & vbTab & flxSP.TextMatrix(i, 2) & vbTab & flxSP.TextMatrix(i, 8) & vbTab & flxSP.TextMatrix(i, 7) & vbTab & vbTab & flxSP.TextMatrix(i, 9) & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & .TextMatrix(i, 13)
                        'Add applicable model
                        '                        1-PARTNUMBER                     2-DESCRIPTION                    3-ORDER                          4-COST                                   6-TOTAL                                                                          13-APPLICABLE MODEL
                        '.AddItem .Rows & vbTab & flxSP.TextMatrix(i, 1) & vbTab & flxSP.TextMatrix(i, 2) & vbTab & flxSP.TextMatrix(i, 8) & vbTab & flxSP.TextMatrix(i, 7) & vbTab & vbTab & flxSP.TextMatrix(i, 9) & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & flxSP.TextMatrix(i, 10)
                        '                        1-PARTNUMBER                     2-DESCRIPTION                    3-ORDER                          4-COST                       6-TOTAL                   10-spid                                                                        12-APPLICABLE MODEL
                        '.AddItem .Rows & vbTab & flxSP.TextMatrix(i, 1) & vbTab & flxSP.TextMatrix(i, 2) & vbTab & flxSP.TextMatrix(i, 9) & vbTab & flxSP.TextMatrix(i, 11) & vbTab & vbTab & flxSP.TextMatrix(i, 10) & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & flxSP.TextMatrix(i, 12)
                        'Design for  BrANCH     1-PARTNUMBER                       2-DESCRIPTION                  3-ORDER                          4-cost                       5-Total                    6-Spid                                                                        12-APPLICABLE MODEL
                        '.AddItem .Rows & vbTab & flxSP.TextMatrix(i, 1) & vbTab & flxSP.TextMatrix(i, 2) & vbTab & flxSP.TextMatrix(i, 12) & vbTab & flxSP.TextMatrix(i, 14) & vbTab & vbTab & flxSP.TextMatrix(i, 13) & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & flxSP.TextMatrix(i, 15)
                        'Design for  Main     1-PARTNUMBER                       2-DESCRIPTION                  3-ORDER                          4-cost                       5-Total                    6-Spid                                                                        12-APPLICABLE MODEL
                        .AddItem .Rows & vbTab & flxSP.TextMatrix(i, 1) & vbTab & flxSP.TextMatrix(i, 2) & vbTab & flxSP.TextMatrix(i, 24) & vbTab & flxSP.TextMatrix(i, 26) & vbTab & vbTab & flxSP.TextMatrix(i, 25) & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & flxSP.TextMatrix(i, 27)
             
                    End If
                End If
            Next i
            If .Rows = 1 Then .AddItem "1"
        End With
    End If
    If MySource = "RECEIVING" Then
        'Unused code because we have separated program for receiving
        'ith SPincoming.flxOrd
        '    If .TextMatrix(.Rows - 1, 7) = "" And .TextMatrix(.Rows - 1, 2) = "" Then 'Remove last row if empty
                
        '        If .Rows = 2 Then
        '            .Rows = 1
        '        Else
        '             .RemoveItem .Rows - 1
        '        End If
        '    End If
        '    For i = 1 To flxSP.Rows - 1
        '        If flxSP.TextMatrix(i, 0) <> "" Then
        '            If Not SPinFlx(flxSP.TextMatrix(i, 9)) Then
                        '.AddItem vbTab & vbTab & vbTab & flxSP.TextMatrix(i, 2) & vbTab & flxSP.TextMatrix(i, 3) & vbTab & flxSP.TextMatrix(i, 5) & vbTab & vbTab & flxSP.TextMatrix(i, 9) & vbTab & flxSP.TextMatrix(i, 1)
                        '
        '                .AddItem vbTab & vbTab & vbTab & flxSP.TextMatrix(i, 2) & vbTab & flxSP.TextMatrix(i, 3) & vbTab & flxSP.TextMatrix(i, 5) & vbTab & vbTab & flxSP.TextMatrix(i, 9) & vbTab & flxSP.TextMatrix(i, 1)
        '
        '            End If
        '        End If
        '    Next i
        '    If .Rows = 1 Then .Rows = 2
        'End With
    End If
End Sub
Function SPinFlx(myPN As String) As Boolean
    Dim i As Integer
    
    If MySource = "SPORDERING" Then
        With SPordering.flxParts
            For i = 1 To .Rows - 1
                If myPN = .TextMatrix(i, 1) Then
                    SPinFlx = True
                    Exit Function
                End If
            Next i
        End With
        SPinFlx = False
        Exit Function
    End If
    'Unused code for receiving
    'If MySource = "RECEIVING" Then
    '    With SPincoming.flxOrd
    '        For i = 1 To .Rows - 1
    '            If MyID = .TextMatrix(i, 7) Then
    '                SPinFlx = True
    '                Exit Function
    '            End If
    '        Next i
    '    End With
    '    SPinFlx = False
    '    Exit Function
    'End If
End Function

Private Sub cmdOK_Click()
    Call cmdInsert_Click
    Call SPordering.ComputeTotal
    Unload Me
End Sub

Private Sub cmdOptHide_Click()
    FrameOpt.Visible = False
    cmdOptHide.Left = cmdOptShow.Left
    cmdOptHide.Visible = False
    cmdOptShow.Visible = True
    flxSPH.Top = cmdOptHide.Top + cmdOptHide.Height + 100
    flxSP.Top = FrameOpt.Height + flxSPH.Height
    cmdInsert.Top = flxSP.Top + flxSP.Height + 120
    cmdOk.Top = flxSP.Top + flxSP.Height + 120
    cmdCancel.Top = flxSP.Top + flxSP.Height + 120
    pBar.Top = flxSP.Top + flxSP.Height
    Me.Height = fmAnotation.Height + cmdOk.Top + cmdOk.Height + 580
    fmAnotation.Top = cmdOk.Top + cmdOk.Height + 900 - fmAnotation.Height + 120
    fScrol.Top = flxSPH.Top + 15
    'txtModel.Text = ""
    ShowOpt = False
End Sub

Private Sub cmdOptShow_Click()
    FrameOpt.Visible = True
    cmdOptHide.Left = cmdOptShow.Left
    cmdOptHide.Visible = True
    cmdOptShow.Visible = False
    flxSP.Visible = True
    flxSPH.Top = FrameOpt.Top + FrameOpt.Height
    flxSP.Top = FrameOpt.Top + FrameOpt.Height + flxSPH.Height
    cmdInsert.Top = flxSP.Top + flxSP.Height + 120
    cmdOk.Top = flxSP.Top + flxSP.Height + 120
    cmdCancel.Top = flxSP.Top + flxSP.Height + 120
    pBar.Top = flxSP.Top + flxSP.Height
    Me.Height = fmAnotation.Height + cmdOk.Top + cmdOk.Height + 580
    fmAnotation.Top = cmdOk.Top + cmdOk.Height + 900 - fmAnotation.Height + 120
    fScrol.Top = flxSPH.Top + 15
    ShowOpt = True
End Sub
Private Sub cmdReplenish_Click()
    Dim rsRpnsh As New ADODB.Recordset
    Dim Qry As String, r As Long, spSOQ As Double, rM As Long, mrulas6m As Double, ymtgm6m As Double, mcpd6mos As Double, mtrmain6mos As Double
    
    cmdReplenish.Enabled = False
    cmdInsert.Enabled = False
    cmdOk.Enabled = False
    cmdCancel.Enabled = False
    'Timer1.Enabled = False
    Call ResizeCompleteHead
    Call CollAlign
    If txtSuppCode.Text = "" Then
        'Qry = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPAREPARTS.QTYONHAND, SPAREPARTS.BACKORDER, SPAREPARTS.SALES, SPAREPARTS.COST, SPAREPARTS.SPAREPARTID, SPAREPARTS.APPLICMODEL, SPAREPARTS.DESCRIPTION2 AS MYSTAT " & _
              "From SPAREPARTS WHERE (((SPAREPARTS.QTYONHAND)" & lblOprQH.Caption & Val(txtQH.Text) & ") AND ((SPAREPARTS.BACKORDER)" & lblOprBO.Caption & Val(txtBO.Text) & ")) ORDER BY SPAREPARTS.PARTNUMBER;"
        'mySQL
        Qry = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPAREPARTS.QTYONHAND, SPAREPARTS.BACKORDER, SPAREPARTS.SALES, SPAREPARTS.COST, SPAREPARTS.SPAREPARTID, SPAREPARTS.APPLICMODEL, SPAREPARTS.DESCRIPTION2 AS MYSTAT " & _
              "From SPAREPARTS WHERE SPAREPARTS.QTYONHAND" & lblOprQH.Caption & Val(txtQH.Text) & " AND SPAREPARTS.BACKORDER" & lblOprBO.Caption & Val(txtBO.Text) & " ORDER BY SPAREPARTS.PARTNUMBER;"

    Else
        'Qry = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPAREPARTS.QTYONHAND, SPAREPARTS.BACKORDER, SPAREPARTS.SALES, SPAREPARTS.COST, SPAREPARTS.SPAREPARTID, SPAREPARTS.APPLICMODEL, SPAREPARTS.DESCRIPTION2 AS MYSTAT " & _
              "From SPAREPARTS WHERE (((SPAREPARTS.QTYONHAND)" & lblOprQH.Caption & Val(txtQH.Text) & ") AND ((SPAREPARTS.BACKORDER)" & lblOprBO.Caption & Val(txtBO.Text) & ") AND ((SPAREPARTS.SUPPLIERCODE)='" & txtSuppCode.Text & "')) ORDER BY SPAREPARTS.PARTNUMBER;"
        'mySQL
        Qry = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPAREPARTS.QTYONHAND, SPAREPARTS.BACKORDER, SPAREPARTS.SALES, SPAREPARTS.COST, SPAREPARTS.SPAREPARTID, SPAREPARTS.APPLICMODEL, SPAREPARTS.DESCRIPTION2 AS MYSTAT " & _
              "From SPAREPARTS WHERE SPAREPARTS.QTYONHAND" & lblOprQH.Caption & Val(txtQH.Text) & " AND SPAREPARTS.BACKORDER" & lblOprBO.Caption & Val(txtBO.Text) & " AND SPAREPARTS.SUPPLIERCODE='" & txtSuppCode.Text & "' ORDER BY SPAREPARTS.PARTNUMBER;"

       
    End If
    
    With rsRpnsh
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            flxSP.Rows = 0
            pBar.Value = 0
            pBar.Visible = True
            pBar.Max = .RecordCount * 2
            
            'Do While Not .EOF
                      '1-PART NUMBER, 2-DESCRIPTION, 3- QH, 4- BO, 5-SQ, 6-COST, 7-ORDER, 8-SPID, 9, applicmodel, 10-MYSTAT
                      '1-PART NUMBER, 2-DESCRIPTION, 3- QH, 4- BO, 5-SQ, 6-ST, 7-COST, 8-ORDER, 9-SPID, 10- applicmodel, 11-MYSTAT
                'Previous code
                'flxSP.AddItem vbTab & !PartNumber & vbTab & !Description & vbTab & !QtyOnHand & vbTab & IIf(IsNull(!BACKORDER), "", !BACKORDER) & vbTab & IIf(IsNull(!Sales), "", !Sales) & vbTab & vbTab & Format(!Cost, "#,##0.00") & vbTab & vbTab & !SPAREPARTID & vbTab & IIf(IsNull(!ApplicModel), "", !ApplicModel) & vbTab & IIf(IsNull(!MyStat), "", !MyStat)
                '                      1-PARTNUMBER         2-DESCRIPTION          3-QH                    4-BO                                         5-SQ                                  6-ST    7-6MOS PSQ     8-SOQ  9-ORDER     10-SPID              11- CWOV                        12-APPMOD                      13-MYSTAT
                'flxSP.AddItem vbTab & !PartNumber & vbTab & !Description & vbTab & !QtyOnHand & vbTab & IIf(IsNull(!BACKORDER), "", !BACKORDER) & vbTab & IIf(IsNull(!Sales), "", !Sales) & vbTab & vbTab & vbTab & vbTab & vbTab & !SPAREPARTID & vbTab & Format(!Cost, "#,##0.00") & vbTab & IIf(IsNull(!ApplicModel), "", !ApplicModel) & vbTab & IIf(IsNull(!MyStat), "", !MyStat)
                'flxSP.AddItem vbTab & !PartNumber & vbTab & !Description & vbTab & !QtyOnHand & vbTab & IIf(IsNull(!BACKORDER), "", !BACKORDER) & vbTab & IIf(IsNull(!Sales), "", !Sales) & vbTab & vbTab & vbTab & vbTab & vbTab & !SPAREPARTID & vbTab & Format(!Cost, "#,##0.00") & vbTab & IIf(IsNull(!ApplicModel), "", !ApplicModel) & vbTab & IIf(IsNull(!MyStat), "", !MyStat)
                
                'Designed for Branch
                          '0-Blank Col 1-SPNUMBER           2-Description          3-QH        4-BO         5-SQ          6-QH                 7-BO                                                        8-SQ                           9-ST         10-6MOS     11-6m        12-SOQ       13-ORDER         14-spid                        15-COST                         16-APPMOD                                       17-STATUS
            '    flxSP.AddItem vbTab & !PartNumber & vbTab & !Description & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & !QtyOnHand & vbTab & IIf(IsNull(!BACKORDER), "", !BACKORDER) & vbTab & IIf(IsNull(!Sales), "", !Sales) & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & !SPAREPARTID & vbTab & Format(!Cost, "#,##0.00") & vbTab & IIf(IsNull(!ApplicModel), "", !ApplicModel) & vbTab & IIf(IsNull(!MyStat), "", !MyStat)
            '    Pbar.Value = Pbar.Value + 1
                
                'Hi-lights Column by rows
            '    flxSP.Col = 3
            '    flxSP.Row = Pbar.Value - 1
            '    flxSP.ColSel = 3
            '    flxSP.ColSel = 4
            '    flxSP.ColSel = 5
            '    flxSP.RowSel = Pbar.Value - 1
            '    flxSP.FillStyle = flexFillRepeat
            '    flxSP.CellBackColor = &HC0FFC0

            '    .MoveNext
            '    DoEvents
            'Loop
            .MoveFirst
            For rM = 1 To .RecordCount
                    flxSP.Rows = rM
                    flxSP.TextMatrix(rM - 1, 1) = !PartNumber
                    flxSP.TextMatrix(rM - 1, 2) = !Description
                    flxSP.TextMatrix(rM - 1, 19) = !QtyOnHand
                    flxSP.TextMatrix(rM - 1, 20) = IIf(IsNull(!BACKORDER), "", !BACKORDER)
                    'flxSP.TextMatrix(rM - 1, 21) = IIf(IsNull(!Sales), "", !Sales)
                    flxSP.TextMatrix(rM - 1, 25) = !SPAREPARTID
                    flxSP.TextMatrix(rM - 1, 26) = Format(!Cost, "#,##0.00")
                    flxSP.TextMatrix(rM - 1, 27) = IIf(IsNull(!ApplicModel), "", !ApplicModel)
                    flxSP.TextMatrix(rM - 1, 28) = IIf(IsNull(!MyStat), "", !MyStat)
                    pBar.Value = pBar.Value + 1
                .MoveNext
            Next rM
            
            flxSP.Visible = True
            If ShowOpt Then
                flxSPH.Top = FrameOpt.Top + FrameOpt.Height
                flxSP.Top = FrameOpt.Top + FrameOpt.Height + flxSPH.Height
            Else
                flxSPH.Top = cmdOptHide.Top + cmdOptHide.Height + 100
                flxSP.Top = FrameOpt.Height + flxSPH.Height
            End If
            cmdInsert.Top = flxSP.Top + flxSP.Height + 120
            cmdOk.Top = flxSP.Top + flxSP.Height + 120
            cmdCancel.Top = flxSP.Top + flxSP.Height + 120
            pBar.Top = flxSP.Top + flxSP.Height
            Me.Height = fmAnotation.Height + cmdOk.Top + cmdOk.Height + 580
           
        Else
        '    With flxSP
                flxSP.Clear
                flxSP.Rows = 1
        '        'Old code
                      '1-PART NUMBER, 2-DESCRIPTION, 3- QH, 4- BO, 5-SQ, 6-ST, 7-COST, 8-ORDER, 9-SPID, 10- applicmodel, 11-MYSTAT
                '.FormatString = "  |<PART NUMBER|<DESCRIPTION|^QH |^BO |^SQ |^ST |>CWOV |>ORDER |  |<APPLICABLE MODEL|"
                      
                'New Designed
                      '1-PART NUMBER, 2-DESCRIPTION, 3- QH, 4- BO, 5-SQ, 6-ST, 7-6MOS, 8-SOQ, 9-ORDER,10-SPID, 11-CWOV 12- applicmodel, 13-MYSTAT
       '         .FormatString = "  |<PART NUMBER|<DESCRIPTION|^QH |^BO |^SQ |^ST |^6m. SQ|^SOQ|>ORDER | |^CWOV |< APPLMOD|"
                
       '         .ColWidth(0) = 300
                'PARTNUMBER
       '         .ColWidth(1) = (.Width - 400) * 0.15
                'DISCRIPTION
       '         .ColWidth(2) = (.Width - 400) * 0.33
                'QH
       '         .ColWidth(3) = (.Width - 400) * 0.05
                'BO
       '         .ColWidth(4) = (.Width - 400) * 0.05
                'SQ
       '         .ColWidth(5) = (.Width - 400) * 0.05
                'ST
       '         .ColWidth(6) = (.Width - 400) * 0.05
                '6MOS-PSQ
       '         .ColWidth(7) = (.Width - 400) * 0.05
                 'SOQ
       '         .ColWidth(8) = (.Width - 400) * 0.05
                'ORDER
       '         .ColWidth(9) = (.Width - 400) * 0.05
                'SPID
       '         .ColWidth(10) = 0
                'CWOV
       '         .ColWidth(11) = (.Width - 400) * 0.05
                'APPLIMOD
       '         .ColWidth(12) = (.Width - 400) * 0.1
                'MYSTAT
       '         .ColWidth(13) = 0
                
                
       '         .Visible = True
       '     End With
           
        End If
        .Close
    End With
    
    If flxSP.Rows = 1 And flxSP.TextMatrix(0, 1) = "" Then Exit Sub
    
    'flxSP.ColWidth(10) = 0
    'flxSP.ColWidth(13) = 0
    'For i = flxSP.Rows - 1 To 1 Step -1
   
    For i = flxSP.Rows - 1 To 0 Step -1
        'flxSP.TextMatrix(i, 5) = GetSalesQty(Val(flxSP.TextMatrix(i, 10)))
        'Main Branch
        flxSP.TextMatrix(i, 21) = GetSalesQty(Val(flxSP.TextMatrix(i, 25)))
        
        'Qry = "SELECT Sum(SPITR.OUT) AS SumOfOUT From SPITR " & _
              "WHERE (((SPITR.TRANSDATE)>=#" & SDfrom.Value & "# And (SPITR.TRANSDATE)<=#" & SDto.Value & "#) AND ((SPITR.TRANSTYPE)='OUT') AND ((SPITR.REMARKS)='TRANSFER') AND ((SPITR.SPAREPARTID)=" & Val(flxSP.TextMatrix(i, 9)) & "));"
        
        '----------------------------------------------------------
        '** Stock Transfer - Order Taking
        'Moto-r Ulas
        'Qry = "SELECT Sum(SPITR.OUT) AS SumOfOUT From SPITR " & _
              "WHERE SPITR.TRANSDATE>='" & Format(SDfrom.Value, "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(SDto.Value, "yyyy-mm-dd") & "' AND SPITR.TRANSTYPE='OUT' AND SPITR.REMARKS='TRANSFER' AND SPITR.SPAREPARTID=" & Val(flxSP.TextMatrix(i, 28)) & " AND SPITR.SOURCE='MOTO-R ULAS';"
        'flxSP.TextMatrix(i, 22) = Val(GetFieldValue(Qry))
        'Yamaha Motorhub
        'Qry = "SELECT Sum(SPITR.OUT) AS SumOfOUT From SPITR " & _
              "WHERE SPITR.TRANSDATE>='" & Format(SDfrom.Value, "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(SDto.Value, "yyyy-mm-dd") & "' AND SPITR.TRANSTYPE='OUT' AND SPITR.REMARKS='TRANSFER' AND SPITR.SPAREPARTID=" & Val(flxSP.TextMatrix(i, 28)) & " AND SPITR.SOURCE='YAMAHA MOTORHUB';"
        'flxSP.TextMatrix(i, 23) = Val(GetFieldValue(Qry))
        
        'MC Pasrts Depot
        'Qry = "SELECT Sum(SPITR.OUT) AS SumOfOUT From SPITR " & _
              "WHERE SPITR.TRANSDATE>='" & Format(SDfrom.Value, "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(SDto.Value, "yyyy-mm-dd") & "' AND SPITR.TRANSTYPE='OUT' AND SPITR.REMARKS='TRANSFER' AND SPITR.SPAREPARTID=" & Val(flxSP.TextMatrix(i, 28)) & " AND SPITR.SOURCE='MC PARTS DEPOT CORP.';"
        'flxSP.TextMatrix(i, 24) = Val(GetFieldValue(Qry))
        '-----------------------------------------------------------
        
        
        
        pBar.Value = IIf(pBar.Value = pBar.Max, pBar.Max, pBar.Value + 1)
        DoEvents
        If flxSP.TextMatrix(i, 28) = "INACTIVE" Then
            flxSP.Col = 1: flxSP.Row = i: flxSP.ColSel = 28: flxSP.RowSel = i: flxSP.FillStyle = flexFillRepeat: flxSP.CellForeColor = RGB(255, 0, 0)
        End If
        '***Option Sales Quantity
        Select Case OptrSQ.Value
            Case 0
                If Val(flxSP.TextMatrix(i, 21)) <> Val(txtSQ.Text) Then 'flxSP.TextMatrix(i, 9) = "x"
                    If flxSP.Rows = 1 Then
                        flxSP.Rows = 1
                        'flxSP.Rows = 2
                    Else
                        flxSP.RemoveItem i
                    End If
                End If
            Case 1
                If Val(flxSP.TextMatrix(i, 21)) <= Val(txtSQ.Text) Then
                    If flxSP.Rows = 1 Then
                        flxSP.Rows = 1
                        'flxSP.Rows = 2
                    Else
                        flxSP.RemoveItem i
                    End If
                End If
            Case 2
                If Val(flxSP.TextMatrix(i, 21)) < Val(txtSQ.Text) Then
                    If flxSP.Rows = 1 Then
                        flxSP.Rows = 1
                        'flxSP.Rows = 2
                    Else
                        flxSP.RemoveItem i
                    End If
                End If
            Case 3
                If Val(flxSP.TextMatrix(i, 21)) >= Val(txtSQ.Text) Then
                    If flxSP.Rows = 1 Then
                        flxSP.Rows = 1
                        'flxSP.Rows = 2
                    Else
                        flxSP.RemoveItem i
                    End If
                End If
            Case 4
                If Val(flxSP.TextMatrix(i, 21)) > Val(txtSQ.Text) Then
                    If flxSP.Rows = 1 Then
                        flxSP.Rows = 1
                        'flxSP.Rows = 2
                    Else
                        flxSP.RemoveItem i
                    End If
                End If
            Case 5
                If Val(flxSP.TextMatrix(i, 21)) = Val(txtSQ.Text) Then
                    If flxSP.Rows = 1 Then
                        flxSP.Rows = 1
                        'flxSP.Rows = 2
                    Else
                        flxSP.RemoveItem i
                    End If
                End If
        End Select
    Next i
    Call AddInCombo
    'Pbar.Visible = False
    'Get 6 mos Sales quantity
    
    'pBar.Max = flxSP.Rows - 1
    'For r = 0 To flxSP.Rows - 1
    pBar.Value = 0
    pBar.Max = flxSP.Rows
    For r = 0 To flxSP.Rows - 1
        pBar.Value = pBar.Value + 1
        
        '**Query Branch SKU
        'Moto-r Ulas
        Qry = ""
        Qry = "SELECT spareparts.QTYONHAND From spareparts Where spareparts.SPAREPARTID = " & flxSP.TextMatrix(r, 25) & ";"
        flxSP.TextMatrix(r, 3) = Val(GetFieldValueMU(Qry))
        Qry = ""
        Qry = "SELECT spareparts.BACKORDER From spareparts Where spareparts.SPAREPARTID = " & flxSP.TextMatrix(r, 25) & ";"
        flxSP.TextMatrix(r, 4) = Val(GetFieldValueMU(Qry))
        flxSP.TextMatrix(r, 5) = GetSalesQtyBrchMU(Val(flxSP.TextMatrix(r, 25)))
        Qry = "SELECT Sum(SPITR.OUT) AS SumOfOUT From SPITR " & _
              "WHERE SPITR.TRANSDATE>='" & Format(SDfrom.Value, "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(SDto.Value, "yyyy-mm-dd") & "' AND SPITR.TRANSTYPE='OUT' AND SPITR.REMARKS='TRANSFER' AND SPITR.SPAREPARTID=" & Val(flxSP.TextMatrix(r, 25)) & " AND SPITR.SOURCE='MOTO-R ULAS';"
        flxSP.TextMatrix(r, 6) = Val(GetFieldValueMU(Qry))
        
        'Yamaha MotorHub
        Qry = ""
        Qry = "SELECT spareparts.QTYONHAND From spareparts Where spareparts.SPAREPARTID = " & flxSP.TextMatrix(r, 25) & ";"
        flxSP.TextMatrix(r, 7) = Val(GetFieldValueYM(Qry))
        Qry = ""
        Qry = "SELECT spareparts.BACKORDER From spareparts Where spareparts.SPAREPARTID = " & flxSP.TextMatrix(r, 25) & ";"
        flxSP.TextMatrix(r, 8) = Val(GetFieldValueYM(Qry))
        flxSP.TextMatrix(r, 9) = GetSalesQtyBrchYM(Val(flxSP.TextMatrix(r, 25)))
        Qry = "SELECT Sum(SPITR.OUT) AS SumOfOUT From SPITR " & _
              "WHERE SPITR.TRANSDATE>='" & Format(SDfrom.Value, "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(SDto.Value, "yyyy-mm-dd") & "' AND SPITR.TRANSTYPE='OUT' AND SPITR.REMARKS='TRANSFER' AND SPITR.SPAREPARTID=" & Val(flxSP.TextMatrix(r, 25)) & " AND SPITR.SOURCE='YAMAHA MOTORHUB';"
        flxSP.TextMatrix(r, 10) = Val(GetFieldValueYM(Qry))
        
        'MC Parts Depot Corporation
        Qry = "SELECT spareparts.QTYONHAND From spareparts Where spareparts.SPAREPARTID = " & flxSP.TextMatrix(r, 25) & ";"
        flxSP.TextMatrix(r, 11) = Val(GetFieldValueMD(Qry))
        Qry = ""
        Qry = "SELECT spareparts.BACKORDER From spareparts Where spareparts.SPAREPARTID = " & flxSP.TextMatrix(r, 25) & ";"
        flxSP.TextMatrix(r, 12) = Val(GetFieldValueMD(Qry))
        flxSP.TextMatrix(r, 13) = GetSalesQtyBrchMD(Val(flxSP.TextMatrix(r, 25)))
        Qry = "SELECT Sum(SPITR.OUT) AS SumOfOUT From SPITR " & _
              "WHERE SPITR.TRANSDATE>='" & Format(SDfrom.Value, "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(SDto.Value, "yyyy-mm-dd") & "' AND SPITR.TRANSTYPE='OUT' AND SPITR.REMARKS='TRANSFER' AND SPITR.SPAREPARTID=" & Val(flxSP.TextMatrix(r, 25)) & " AND SPITR.SOURCE='MC PARTS DEPOT CORP.';"
        flxSP.TextMatrix(r, 14) = Val(GetFieldValueMD(Qry))
       
        'Total Branch
        'QH
        flxSP.TextMatrix(r, 15) = Val(flxSP.TextMatrix(r, 3)) + Val(flxSP.TextMatrix(r, 7)) + Val(flxSP.TextMatrix(r, 11))
        'BO
        flxSP.TextMatrix(r, 16) = Val(flxSP.TextMatrix(r, 4)) + Val(flxSP.TextMatrix(r, 8)) + Val(flxSP.TextMatrix(r, 12))
        'SQ all Branch
        flxSP.TextMatrix(r, 17) = Val(flxSP.TextMatrix(r, 5)) + Val(flxSP.TextMatrix(r, 9)) + Val(flxSP.TextMatrix(r, 13))
        'ST
        flxSP.TextMatrix(r, 18) = Val(flxSP.TextMatrix(r, 6)) + Val(flxSP.TextMatrix(r, 10)) + Val(flxSP.TextMatrix(r, 14))
        
        '6 months Sales Quantity All Branch
        'Moto-R Main
        
        mrulas6m = GetSalesQty6mosMU(Val(flxSP.TextMatrix(r, 25)))
        ymtgm6m = GetSalesQty6mosYM(Val(flxSP.TextMatrix(r, 25)))
        mcpd6mos = GetSalesQty6mosMD(Val(flxSP.TextMatrix(r, 25)))
        mtrmain6mos = GetSalesQty6mos(Val(flxSP.TextMatrix(r, 25)))
        flxSP.TextMatrix(r, 22) = mrulas6m + ymtgm6m + mcpd6mos + mtrmain6mos
        
        'randyp                                Branch QTY Total + main                                         Back Order Main Only
        spSOQ = Val(flxSP.TextMatrix(r, 22)) - (Val(flxSP.TextMatrix(r, 15)) + Val(flxSP.TextMatrix(r, 19))) - Val(flxSP.TextMatrix(r, 20))
        flxSP.TextMatrix(r, 23) = IIf(spSOQ < 0, 0, spSOQ)
    Next r
    pBar.Visible = False
    'Call FillColorGrd
    'flxSP.Row = 0
    'flxSP.Col = 12
    If flxSP.Rows >= 1 Then
        'If flxSP.TextMatrix(0, 6) <> "" Then
        If flxSP.TextMatrix(0, 19) <> "" Then
            flxSP.SetFocus
        Else
            cboFind.SetFocus
        End If
    Else
        cboFind.SetFocus
    End If
        
   cmdReplenish.Enabled = True
   cmdInsert.Enabled = True
   cmdOk.Enabled = True
   cmdCancel.Enabled = True
   'Timer1.Enabled = True
   Call fScrol_Change
   
End Sub

Private Sub flxSP_Click()
    
    ActRow = flxSP.Row
    ActCol = flxSP.Col
    If ActCol <> 24 Then: RwSel = False: Exit Sub
    With txtOrder
        .Top = flxSP.CellTop + flxSP.Top
        .Left = flxSP.CellLeft + flxSP.Left
        .Width = flxSP.CellWidth
        .Text = Trim(flxSP.Text)
        .Visible = True
        .SelStart = 0
        .SelLength = Len(.Text)
        .ZOrder
        .SetFocus
    End With
    RwSel = True
    flxSP.RowSel = flxSP.Row
    flxSP.Col = 0
    flxSP.ColSel = 28
    RwSel = False
    flxSP.FillStyle = flexFillRepeat
    
End Sub
Private Sub flxSP_DblClick()
    If flxSP.TextMatrix(flxSP.Row, 1) = "" And Val(flxSP.TextMatrix(flxSP.Row, 1)) > 0 Then
        Call mnuMark_Click
    Else
        Call mnuUnmark_Click
    End If
End Sub

Private Sub flxSP_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
        If flxSP.TextMatrix(flxSP.Row, 1) = "" And Val(flxSP.TextMatrix(flxSP.Row, 1)) > 0 Then
            Call mnuMark_Click
        Else
            Call mnuUnmark_Click
        End If
    ElseIf KeyCode = vbKeyF11 Then
        flxSP.Col = 24
        Call flxSP_Click
    End If
End Sub
Private Sub flxSP_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call flxSP_Click
End Sub
Private Sub flxSP_LeaveCell()
    If RwSel = True Then Exit Sub
        flxSP.ColSel = 28
        flxSP.FillStyle = flexFillRepeat
End Sub
Private Sub flxSP_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
    If Val(flxSP.TextMatrix(flxSP.Row, 1)) = 0 Then Exit Sub
    If Button = vbRightButton Then
        If flxSP.TextMatrix(flxSP.Row, 16) = "INACTIVE" Then
            mnuInAct.Visible = False
            mnuAct.Visible = True
        Else
            mnuInAct.Visible = True
            mnuAct.Visible = False
        End If
        Call PopupMenu(mnuSP)
    End If
End Sub
Private Sub flxSP_SelChange()

    If flxSP.Col = 24 Then Exit Sub
        RwSel = True
        flxSP.RowSel = flxSP.Row
        flxSP.Col = 0
        flxSP.ColSel = 28
        RwSel = False
        flxSP.FillStyle = flexFillRepeat
    
End Sub
Private Sub Form_Load()
    Dim Qry As String, m As Integer
    
    'With flxSP
    '    .ColWidth(0) = 300
    '    .ColWidth(1) = (.Width - 400) * 0.15
    '    .ColWidth(2) = (.Width - 400) * 0.33
    '    .ColWidth(3) = (.Width - 400) * 0.05
    '    .ColWidth(4) = (.Width - 400) * 0.05
    '    .ColWidth(5) = (.Width - 400) * 0.05
    '    .ColWidth(6) = (.Width - 400) * 0.05
    '    .ColWidth(7) = (.Width - 400) * 0.05
    '    .ColWidth(8) = (.Width - 400) * 0.05
    '    .ColWidth(9) = (.Width - 400) * 0.05
    '    .ColWidth(10) = 0
    '    .ColWidth(11) = (.Width - 400) * 0.05
    '    .ColWidth(12) = (.Width - 400) * 0.1
    '    .ColWidth(13) = 0
        
    'End With
    RwSel = True
    Call cmdOptHide_Click
    
    With SPordering
        If .LockFlx = "True" Then
            cmdInsert.Enabled = False
            cmdOk.Enabled = False
        End If
        'SDfrom.Value = .SalesDateFrom.Value
        SDfrom.Value = DateAdd("m", -6, Date)
        SDto.Value = .SalesDateTo.Value
        cboMos.Text = 6
        For m = 1 To 6
            cboMos.AddItem m
        Next m
        cboMos.AddItem "All"
    End With
    
    Qry = "SELECT SUPPLIER.SUPPLIERID, SUPPLIER.SUPPLIERNAME, SUPPLIER.SUPPLIERCODE " & _
          "From SUPPLIER ORDER BY SUPPLIER.SUPPLIERNAME;"
    With rsSupplier
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
    
    cboSupplier.Text = SPordering.cboSupplier.Text
    txtSuppCode.Text = SPordering.txtScode.Text
    'Set Search Description
    UpDRef.Value = 1
    Call ResizeCompleteHead
    Call ResizeCompleteFlxSP
    
    'Connect use Moto-r Main Program
    'Call ConnectMain
    Call ConnectMU
    Call ConnectYM
    Call ConnectMD
    'Timer1.Enabled = True
    'LYRmain.sBar.Panels(6).Text = "Find Spare Parts"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'LYRmain.sBar.Panels(6).Text = ""
    'MainCon.Close
    BrchConMU.Close
    BrchConYM.Close
    BrchConMD.Close
    'Timer1.Enabled = False
    If rsSupplier.State = adStateOpen Then rsSupplier.Close
End Sub
Private Sub fScrol_Change()
Dim reAlns As Integer, aCol As Integer
    
    With flxSP
       If fScrol.Value = 1 Then
            .ColWidth(3) = (.Width - 300) * 0.036
            .ColWidth(4) = (.Width - 300) * 0.036
            .ColWidth(5) = (.Width - 300) * 0.036
            .ColWidth(6) = (.Width - 300) * 0.036
            
            flxSPH.ColWidth(3) = (.Width - 300) * 0.036
            flxSPH.ColWidth(4) = (.Width - 300) * 0.036
            flxSPH.ColWidth(5) = (.Width - 300) * 0.036
            flxSPH.ColWidth(6) = (.Width - 300) * 0.036
            For reCol = 3 To 6
                flxSPH.ColAlignment(reCol) = 4
                .ColAlignment(reCol) = 4
            Next reCol
            
            For reCol = 7 To 14
                flxSPH.ColWidth(reCol) = 0
                flxSPH.ColAlignment(reCol) = 1
                .ColWidth(reCol) = 0
                .ColAlignment(reCol) = 1
            Next reCol
        ElseIf fScrol.Value = 2 Then
            .ColWidth(7) = (.Width - 300) * 0.036
            .ColWidth(8) = (.Width - 300) * 0.036
            .ColWidth(9) = (.Width - 300) * 0.036
            .ColWidth(10) = (.Width - 300) * 0.036
            
            flxSPH.ColWidth(7) = (.Width - 300) * 0.036
            flxSPH.ColWidth(8) = (.Width - 300) * 0.036
            flxSPH.ColWidth(9) = (.Width - 300) * 0.036
            flxSPH.ColWidth(10) = (.Width - 300) * 0.036
            For reCol = 7 To 10
                 flxSPH.ColAlignment(reCol) = 4
                .ColAlignment(reCol) = 4
            Next reCol
            For reCol = 3 To 6
                flxSPH.ColWidth(reCol) = 0
                flxSPH.ColAlignment(reCol) = 1
                .ColWidth(reCol) = 0
                .ColAlignment(reCol) = 1
            Next reCol
            
            For reCol = 11 To 14
                flxSPH.ColWidth(reCol) = 0
                flxSPH.ColAlignment(reCol) = 1
                .ColWidth(reCol) = 0
                .ColAlignment(reCol) = 1
            Next reCol
        ElseIf fScrol.Value = 3 Then
            .ColWidth(11) = (.Width - 300) * 0.036
            .ColWidth(12) = (.Width - 300) * 0.036
            .ColWidth(13) = (.Width - 300) * 0.036
            .ColWidth(14) = (.Width - 300) * 0.036
            
            flxSPH.ColWidth(11) = (.Width - 300) * 0.036
            flxSPH.ColWidth(12) = (.Width - 300) * 0.036
            flxSPH.ColWidth(13) = (.Width - 300) * 0.036
            flxSPH.ColWidth(14) = (.Width - 300) * 0.036
            For reCol = 11 To 14
                flxSPH.ColAlignment(reCol) = 4
                .ColAlignment(reCol) = 4
            Next reCol
            For reCol = 3 To 10
                flxSPH.ColWidth(reCol) = 0
                flxSPH.ColAlignment(reCol) = 1
                .ColWidth(reCol) = 0
                .ColAlignment(reCol) = 1
            Next reCol
        End If
        
    End With
        
End Sub
Private Sub fScrol_Scroll()
    'Timer1.Enabled = False
    Call fScrol_Change
    'Timer1.Enabled = True
End Sub

Private Sub mnuAct_Click()
    Dim Ans As Integer
    
    Ans = MsgBox("Do you want ACTIVATE Partnumber " & flxSP.TextMatrix(flxSP.Row, 1) & "?", vbQuestion + vbYesNo, "In-Active")
    If Ans = vbNo Then Exit Sub
    
    flxSP.Col = 1:  flxSP.ColSel = 16: flxSP.RowSel = flxSP.Row: flxSP.FillStyle = flexFillRepeat: flxSP.CellForeColor = RGB(0, 0, 0)
    flxSP.TextMatrix(flxSP.Row, 16) = ""
    
    'LYRcon.Execute "UPDATE SPAREPARTS SET SPAREPARTS.DESCRIPTION2 = '' WHERE SPAREPARTS.SPAREPARTID=" & Val(flxSP.TextMatrix(flxSP.Row, 9)) & ";"
    'mySQL
    LYRcon.Execute "UPDATE SPAREPARTS SET SPAREPARTS.DESCRIPTION2 = '' WHERE SPAREPARTS.SPAREPARTID=" & Val(flxSP.TextMatrix(flxSP.Row, 13)) & ";"

End Sub

Private Sub mnuInAct_Click()
    Dim Ans As Integer
    
    Ans = MsgBox("Do you want Partnumber " & flxSP.TextMatrix(flxSP.Row, 1) & " to be IN-ACTIVE?", vbQuestion + vbYesNo, "In-Active")
    If Ans = vbNo Then Exit Sub
    
    flxSP.Col = 1:  flxSP.ColSel = 16: flxSP.RowSel = flxSP.Row: flxSP.FillStyle = flexFillRepeat: flxSP.CellForeColor = RGB(255, 0, 0)
    flxSP.TextMatrix(flxSP.Row, 16) = "INACTIVE"
    
    'LYRcon.Execute "UPDATE SPAREPARTS SET SPAREPARTS.DESCRIPTION2 = 'INACTIVE' WHERE (((SPAREPARTS.SPAREPARTID)=" & Val(flxSP.TextMatrix(flxSP.Row, 9)) & "));"
    'mySQL
    LYRcon.Execute "UPDATE SPAREPARTS SET SPAREPARTS.DESCRIPTION2 = 'INACTIVE' WHERE SPAREPARTS.SPAREPARTID=" & Val(flxSP.TextMatrix(flxSP.Row, 13)) & ";"
    
End Sub

Private Sub mnuMark_Click()
    With flxSP
        .Col = 0
        .CellFontName = "WINGDINGS"
        .CellFontSize = 12
        .Text = lblChk.Caption
        .CellAlignment = 4
    End With
End Sub

Private Sub mnuMarkAll_Click()
    Dim i As Integer
    
    With flxSP
        For i = 1 To .Rows - 1
            .Row = i
            .Col = 0
            .CellFontName = "WINGDINGS"
            .CellFontSize = 12
            .Text = lblChk.Caption
            .CellAlignment = 4
            .TextMatrix(i, 0) = lblChk.Caption
        Next i
    End With
    
End Sub

Private Sub mnuUnmark_Click()
    flxSP.TextMatrix(flxSP.Row, 0) = ""
End Sub

Private Sub mnuUnmarkAll_Click()
    Dim i As Integer
    
    For i = 1 To flxSP.Rows - 1
        flxSP.TextMatrix(i, 0) = ""
    Next i
End Sub


Private Sub OptrBO_Change()
    Select Case OptrBO.Value
        Case 0
            lblOprBO.Caption = "="
        Case 1
            lblOprBO.Caption = ">"
        Case 2
            lblOprBO.Caption = ">="
        Case 3
            lblOprBO.Caption = "<"
        Case 4
            lblOprBO.Caption = "<="
        Case 5
            lblOprBO.Caption = "<>"
    End Select

End Sub
Private Sub OptrQH_Change()
    Select Case OptrQH.Value
        Case 0
            lblOprQH.Caption = "="
        Case 1
            lblOprQH.Caption = ">"
        Case 2
            lblOprQH.Caption = ">="
        Case 3
            lblOprQH.Caption = "<"
        Case 4
            lblOprQH.Caption = "<="
        Case 5
            lblOprQH.Caption = "<>"
    End Select
End Sub
Private Sub OptrSQ_Change()
    Select Case OptrSQ.Value
        Case 0
            lblOprSQ.Caption = "="
        Case 1
            lblOprSQ.Caption = ">"
        Case 2
            lblOprSQ.Caption = ">="
        Case 3
            lblOprSQ.Caption = "<"
        Case 4
            lblOprSQ.Caption = "<="
        Case 5
            lblOprSQ.Caption = "<>"
    End Select
End Sub

Private Sub Timer1_Timer()
    If flxSP.Rows = 1 And flxSP.TextMatrix(0, 1) = "" Then Exit Sub
    For Cw = 0 To 24
        flxSP.ColWidth(Cw) = flxSPH.ColWidth(Cw)
    Next Cw
    fScrol.Left = flxSPH.ColPos(4) + 60
    flxSP.ColWidth(26) = flxSPH.ColWidth(25)
    flxSP.ColWidth(27) = flxSPH.ColWidth(26)
End Sub
Private Sub txtOrder_GotFocus()
    'Timer1.Enabled = False
End Sub

Private Sub txtOrder_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        If Val(txtOrder.Text) > Val(flxSP.TextMatrix(ActRow, 23)) And UserID <> 4 Then
            flxSP.TextMatrix(ActRow, 24) = ""
            flxSP.TextMatrix(ActRow, 0) = ""
            txtOrder.Visible = False
            flxSP.Col = 24
            Exit Sub
        ElseIf Val(txtOrder.Text) > Val(flxSP.TextMatrix(ActRow, 23)) And UserID = 4 Then
            MsgBox ("Order Greater than Suggested Order Quantity (SOQ)"), vbExclamation, "Order"
            txtOrder.SetFocus
            Exit Sub
        ElseIf Val(txtOrder.Text) < Val(flxSP.TextMatrix(ActRow, 23)) And Val(txtOrder.Text) <> 0 Then
            MsgBox ("Order Lesser than Suggested Order Quantity (SOQ)"), vbExclamation, "Order"
            txtOrder.SetFocus
            Exit Sub
        End If
        txtOrder.Text = flxSP.TextMatrix(ActRow, 24)
        flxSP.SetFocus
        flxSP.Row = ActRow
        flxSP.Col = 24
        txtOrder.Visible = False
    End If
End Sub
Private Sub txtOrder_KeyPress(KeyAscii As Integer)
    Dim strValid As String
    
    strValid = "1234567890"
    
    If KeyAscii > 26 Then
        If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
    If KeyAscii = 13 Then
            flxSP.SetFocus
            flxSP.Row = ActRow
            flxSP.Col = 24
   End If
End Sub
Private Sub txtOrder_LostFocus()
    With flxSP
    
        If Val(txtOrder.Text) > Val(.TextMatrix(ActRow, 23)) And UserID <> 4 Then
            MsgBox ("Can't add Order Quantity" & vbCrLf & _
                    "which Greater than Suggested Order Quantity (SOQ)"), vbExclamation, "Order"
            flxSP.TextMatrix(ActRow, 24) = ""
            flxSP.TextMatrix(ActRow, 0) = ""
            If txtOrder.Visible = True Then txtOrder.SetFocus: SendKeys "{Home}+{End}"
            Exit Sub
        ElseIf Val(txtOrder.Text) > Val(flxSP.TextMatrix(ActRow, 23)) And UserID = 4 Then
            MsgBox ("Order Greater than Suggested Order Quantity (SOQ)"), vbExclamation, "Order"
        ElseIf Val(txtOrder.Text) < Val(flxSP.TextMatrix(ActRow, 23)) And Val(txtOrder.Text) <> 0 Then
            MsgBox ("Order Lesser than Suggested Order Quantity (SOQ)"), vbExclamation, "Order"
        End If
        
        .TextMatrix(ActRow, 24) = txtOrder.Text
        txtOrder.Visible = False
        If Val(.TextMatrix(ActRow, 24)) > 0 Then
            .Row = ActRow
            .Col = 0
            .CellFontName = "WINGDINGS"
            .CellFontSize = 12
            .Text = lblChk.Caption
            .CellAlignment = 4
            .Col = 24
        Else
            .TextMatrix(ActRow, 24) = ""
            .TextMatrix(ActRow, 0) = ""
            .Col = 24
        End If
        
        'Timer1.Enabled = True
    End With
End Sub
Private Sub UpDRef_Change()
    If UpDRef.Value = 0 Then
        lblRefSearch.Caption = "FIND PARTNUMBER"
    Else
        lblRefSearch.Caption = "FIND PARTS DESC."
    End If
   
End Sub
Private Sub ResizeCompleteHead()
Dim aCol As Integer, aCel As Integer, reAln As Integer
    
    With flxSPH
        
        For aCol = 1 To 26
            .ColAlignment(aCol) = 4
        Next aCol
        For reAln = 7 To 14
            .ColAlignment(reAln) = 1
        Next
        
        .TextMatrix(0, 1) = "S P A R E P A R T S"
        .TextMatrix(0, 2) = "S P A R E P A R T S"
        
        .TextMatrix(0, 3) = "MOTO-R ULAS"
        .TextMatrix(0, 4) = "MOTO-R ULAS"
        .TextMatrix(0, 5) = "MOTO-R ULAS"
        .TextMatrix(0, 6) = "MOTO-R ULAS"
        
        .TextMatrix(0, 7) = "YM MOTORHUB"
        .TextMatrix(0, 8) = "YM MOTORHUB"
        .TextMatrix(0, 9) = "YM MOTORHUB"
        .TextMatrix(0, 10) = "YM MOTORHUB"
        
        .TextMatrix(0, 11) = "MCP DEPOT"
        .TextMatrix(0, 12) = "MCP DEPOT"
        .TextMatrix(0, 13) = "MCP DEPOT"
        .TextMatrix(0, 14) = "MCP DEPOT"
             
        .TextMatrix(0, 15) = "BRANCH TOTAL"
        .TextMatrix(0, 16) = "BRANCH TOTAL"
        .TextMatrix(0, 17) = "BRANCH TOTAL"
        .TextMatrix(0, 18) = "BRANCH TOTAL"
             
        .TextMatrix(0, 19) = "MOTO-R MAIN"
        .TextMatrix(0, 20) = "MOTO-R MAIN"
        .TextMatrix(0, 21) = "MOTO-R MAIN"
             
        '.TextMatrix(0, 22) = "ORDER TAKING"
        '.TextMatrix(0, 23) = "ORDER TAKING"
        '.TextMatrix(0, 24) = "ORDER TAKING"
        
        .TextMatrix(0, 22) = "PURCHASE ORDER"
        .TextMatrix(0, 23) = "PURCHASE ORDER"
        .TextMatrix(0, 24) = "PURCHASE ORDER"
        .TextMatrix(0, 25) = "PURCHASE ORDER"
        .TextMatrix(0, 26) = "PURCHASE ORDER"
        
        'moto-r Ulas
        .TextMatrix(1, 1) = "PARTNUMBER"
        .TextMatrix(1, 2) = "DESCRIPTION"
        .TextMatrix(1, 3) = "QH"
        .TextMatrix(1, 4) = "BO"
        .TextMatrix(1, 5) = "SQ"
        .TextMatrix(1, 6) = "ST"
        'Yamaha Motorhub
        .TextMatrix(1, 7) = "QH"
        .TextMatrix(1, 8) = "BO"
        .TextMatrix(1, 9) = "SQ"
        .TextMatrix(1, 10) = "ST"
        'MC Parts Depot
        .TextMatrix(1, 11) = "QH"
        .TextMatrix(1, 12) = "BO"
        .TextMatrix(1, 13) = "SQ"
        .TextMatrix(1, 14) = "ST"
        'Branch Total
        .TextMatrix(1, 15) = "QH"
        .TextMatrix(1, 16) = "BO"
        .TextMatrix(1, 17) = "SQ"
        .TextMatrix(1, 18) = "ST"
        'Moto-r Main
        .TextMatrix(1, 19) = "QH"
        .TextMatrix(1, 20) = "BO"
        .TextMatrix(1, 21) = cboMos.Text & "mSQ"
        'Stock Order Taking
        '.TextMatrix(1, 22) = "MU"
        '.TextMatrix(1, 23) = "YH"
        '.TextMatrix(1, 24) = "MD"
        'PurChase Order
        '.TextMatrix(1, 25) = "6mSQ"
        '.TextMatrix(1, 26) = "SOQ"
        '.TextMatrix(1, 27) = "ORDER"
        '.TextMatrix(1, 28) = "CWOV"
        '.TextMatrix(1, 29) = "APM"
        'New Without Order Taking
        .TextMatrix(1, 22) = "6mSQ"
        .TextMatrix(1, 23) = "SOQ"
        .TextMatrix(1, 24) = "ORDER"
        .TextMatrix(1, 25) = "CWOV"
        .TextMatrix(1, 26) = "APM"
        
        .MergeCells = flexMergeFree
        
        
        .MergeRow(0) = True
        .MergeCol(1) = True
        .MergeCol(3) = True
        .MergeCol(4) = True
        .MergeCol(5) = True
        .MergeCol(6) = True
        .MergeCol(7) = True
        .MergeCol(8) = True
        .MergeCol(9) = True
        .MergeCol(10) = True
        .MergeCol(11) = True
        .MergeCol(12) = True
        .MergeCol(13) = True
        .MergeCol(14) = True
        .MergeCol(15) = True
        .MergeCol(16) = True
        .MergeCol(17) = True
        .MergeCol(18) = True
        .MergeCol(19) = True
        .MergeCol(20) = True
        .MergeCol(21) = True
        .MergeCol(22) = True
        .MergeCol(23) = True
        .MergeCol(24) = True
        .MergeCol(25) = True
        .MergeCol(26) = True
        '.MergeCol(27) = True
        '.MergeCol(28) = True
        '.MergeCol(29) = True
        
        
        
        .ColWidth(0) = 300
        .ColWidth(1) = (.Width - 300) * 0.13
        .ColWidth(2) = (.Width - 300) * 0.25
        .ColWidth(3) = (.Width - 300) * 0.036
        .ColWidth(4) = (.Width - 300) * 0.036
        .ColWidth(5) = (.Width - 300) * 0.036
        .ColWidth(6) = (.Width - 300) * 0.036
        .ColWidth(7) = (.Width - 300) * 0.036
        .ColWidth(8) = (.Width - 300) * 0.036
        .ColWidth(9) = (.Width - 300) * 0.036
        .ColWidth(10) = (.Width - 300) * 0.036
        .ColWidth(11) = (.Width - 300) * 0.036
        .ColWidth(12) = (.Width - 300) * 0.036
        .ColWidth(13) = (.Width - 300) * 0.036
        .ColWidth(14) = (.Width - 300) * 0.036
        .ColWidth(15) = (.Width - 300) * 0.036
        .ColWidth(16) = (.Width - 300) * 0.036
        .ColWidth(17) = (.Width - 300) * 0.036
        .ColWidth(18) = (.Width - 300) * 0.036
        .ColWidth(19) = (.Width - 300) * 0.036
        .ColWidth(20) = (.Width - 300) * 0.036
        .ColWidth(21) = (.Width - 300) * 0.036
        .ColWidth(22) = (.Width - 300) * 0.036
        .ColWidth(23) = (.Width - 300) * 0.036
        .ColWidth(24) = (.Width - 300) * 0.04
        .ColWidth(25) = (.Width - 300) * 0.07
        .ColWidth(26) = (.Width - 300) * 0.05
        
        '.ColWidth(27) = (.Width - 300) * 0.035
        '.ColWidth(28) = (.Width - 300) * 0.05
        '.ColWidth(29) = (.Width - 300) * 0.04
        
        For reCol = 7 To 14
            .ColWidth(reCol) = 0
        Next reCol
        
        
       
        fScrol.Top = flxSPH.Top + 15
        fScrol.Height = .RowHeight(0)
        fScrol.Left = flxSPH.ColPos(4) + 80
    End With
End Sub
Sub ResizeCompleteFlxSP()
Dim reCol As Integer
 With flxSP
        .ColWidth(0) = 300
        .ColWidth(1) = (.Width - 300) * 0.13
        .ColWidth(2) = (.Width - 300) * 0.25
        
        .ColWidth(3) = (.Width - 300) * 0.036
        .ColWidth(4) = (.Width - 300) * 0.036
        .ColWidth(5) = (.Width - 300) * 0.036
        .ColWidth(6) = (.Width - 300) * 0.036
            
        For reCol = 7 To 14
            .ColWidth(reCol) = 0
            .ColAlignment(reCol) = 4
        Next reCol
            
        .ColWidth(15) = (.Width - 300) * 0.036
        .ColWidth(16) = (.Width - 300) * 0.036
        .ColWidth(17) = (.Width - 300) * 0.036
        .ColWidth(18) = (.Width - 300) * 0.036
        .ColWidth(19) = (.Width - 300) * 0.036
        .ColWidth(20) = (.Width - 300) * 0.036
        .ColWidth(21) = (.Width - 300) * 0.036
        '.ColWidth(22) = (.Width - 300) * 0.035
        '.ColWidth(23) = (.Width - 300) * 0.035
        '.ColWidth(24) = (.Width - 300) * 0.035
        '.ColWidth(25) = (.Width - 300) * 0.035
        '.ColWidth(26) = (.Width - 300) * 0.035
        '.ColWidth(27) = (.Width - 300) * 0.035
        '.ColWidth(28) = 0
        '.ColWidth(29) = (.Width - 300) * 0.05
        '.ColWidth(30) = (.Width - 300) * 0.04
        '.ColWidth(31) = 0
        .ColWidth(22) = (.Width - 300) * 0.036
        .ColWidth(23) = (.Width - 300) * 0.036
        .ColWidth(24) = (.Width - 300) * 0.04
        .ColWidth(25) = 0
        .ColWidth(26) = (.Width - 300) * 0.07
        .ColWidth(27) = (.Width - 300) * 0.05
        .ColWidth(28) = 0
        
 End With
End Sub
Private Sub CollAlign()
Dim nCols As Integer
     For nCols = 3 To 12
            flxSP.ColAlignment(1) = 2
            flxSP.ColAlignment(2) = 2
            
            flxSP.ColAlignment(nCols) = 4
     Next nCols
End Sub
Private Sub FillColorGrd()
Dim ccol As Integer, cRow As Integer
    flxSP.Col = 3
    flxSP.Row = 0
    For ccol = 3 To 5
        flxSP.ColSel = ccol
        For cRow = 0 To flxSP.Rows - 1
            flxSP.RowSel = cRow
            flxSP.FillStyle = flexFillRepeat: flxSP.CellBackColor = &HC0FFC0
        Next cRow
    Next ccol
End Sub


