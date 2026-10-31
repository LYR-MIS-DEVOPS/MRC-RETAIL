VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form SPInsXstWHAgdao 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Insert Existing Spare Parts"
   ClientHeight    =   7770
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15000
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
   Icon            =   "SPInsXstWHAgdao.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7770
   ScaleWidth      =   15000
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fmAnotation 
      Caption         =   "ANNOTATION"
      ForeColor       =   &H000000C0&
      Height          =   915
      Left            =   135
      TabIndex        =   40
      Top             =   6795
      Width           =   14730
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SOQ ="
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   11385
         TabIndex        =   56
         Top             =   585
         Width           =   570
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ST = "
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   11520
         TabIndex        =   55
         Top             =   270
         Width           =   480
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Suggested Order Quantity"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   12015
         TabIndex        =   54
         Top             =   585
         Width           =   2205
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Stock Transfer"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   12015
         TabIndex        =   53
         Top             =   270
         Width           =   1275
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Quantity ON-HAND"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   1260
         TabIndex        =   52
         Top             =   270
         Width           =   1380
      End
      Begin VB.Label Label6 
         Caption         =   " Supplier BACK-ORDER"
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1215
         TabIndex        =   51
         Top             =   585
         Width           =   2220
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Sales Quantity"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   5220
         TabIndex        =   50
         Top             =   270
         Width           =   1245
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Quantity ON-HAND"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   5220
         TabIndex        =   49
         Top             =   585
         Width           =   1635
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Purchase Order"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   8775
         TabIndex        =   48
         Top             =   270
         Width           =   1335
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Quantity Taking"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   8775
         TabIndex        =   47
         Top             =   585
         Width           =   1365
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "QH ="
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   675
         TabIndex        =   46
         Top             =   270
         Width           =   510
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "BO ="
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   675
         TabIndex        =   45
         Top             =   585
         Width           =   435
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SQ = "
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   4725
         TabIndex        =   44
         Top             =   270
         Width           =   405
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "QH = "
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   4725
         TabIndex        =   43
         Top             =   585
         Width           =   420
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PO = "
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   8280
         TabIndex        =   42
         Top             =   270
         Width           =   405
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "QT ="
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   8280
         TabIndex        =   41
         Top             =   585
         Width           =   450
      End
   End
   Begin MSComCtl2.UpDown UpDRef 
      Height          =   315
      Left            =   1800
      TabIndex        =   36
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
      Left            =   120
      TabIndex        =   20
      Top             =   6450
      Visible         =   0   'False
      Width           =   10440
      _ExtentX        =   18415
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
      TabIndex        =   19
      Top             =   6360
      Visible         =   0   'False
      Width           =   1215
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
      Left            =   10665
      TabIndex        =   7
      Top             =   6360
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "&Insert"
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
      Image           =   "SPInsXstWHAgdao.frx":058A
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
      Image           =   "SPInsXstWHAgdao.frx":0B24
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
      Image           =   "SPInsXstWHAgdao.frx":10BE
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
      Left            =   135
      TabIndex        =   3
      Top             =   720
      Width           =   14760
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
         ItemData        =   "SPInsXstWHAgdao.frx":1658
         Left            =   7320
         List            =   "SPInsXstWHAgdao.frx":165A
         TabIndex        =   37
         Top             =   720
         Width           =   615
      End
      Begin MSComCtl2.UpDown OptrSQ 
         Height          =   285
         Left            =   13320
         TabIndex        =   35
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
         TabIndex        =   34
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
         TabIndex        =   31
         Text            =   "0"
         Top             =   255
         Width           =   480
      End
      Begin MSComCtl2.UpDown updBO 
         Height          =   285
         Left            =   11760
         TabIndex        =   27
         Top             =   615
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   503
         _Version        =   393216
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtBO"
         BuddyDispid     =   196632
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
         TabIndex        =   26
         Text            =   "0"
         Top             =   615
         Width           =   480
      End
      Begin MSComCtl2.UpDown updSQ 
         Height          =   285
         Left            =   14040
         TabIndex        =   25
         Top             =   255
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   503
         _Version        =   393216
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtSQ"
         BuddyDispid     =   196633
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
         TabIndex        =   24
         Text            =   "0"
         Top             =   255
         Width           =   480
      End
      Begin MSComCtl2.DTPicker SDfrom 
         Height          =   315
         Left            =   2220
         TabIndex        =   15
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
         Format          =   104726529
         CurrentDate     =   39498
      End
      Begin VB.CheckBox chkAllSupp 
         Caption         =   "All Codes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
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
         Top             =   360
         Visible         =   0   'False
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
         Left            =   2220
         TabIndex        =   12
         Top             =   360
         Width           =   2685
      End
      Begin MSComCtl2.DTPicker SDto 
         Height          =   315
         Left            =   5010
         TabIndex        =   17
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
         Format          =   104726529
         CurrentDate     =   39498
      End
      Begin MSComCtl2.UpDown OptrQH 
         Height          =   285
         Left            =   11040
         TabIndex        =   30
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
         TabIndex        =   32
         Top             =   255
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   503
         _Version        =   393216
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtQH"
         BuddyDispid     =   196631
         OrigLeft        =   9015
         OrigTop         =   255
         OrigRight       =   9270
         OrigBottom      =   540
         Max             =   1000
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown UpDsearch 
         Height          =   285
         Left            =   1920
         TabIndex        =   57
         Top             =   360
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   503
         _Version        =   393216
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtSQ"
         BuddyDispid     =   196633
         OrigLeft        =   11280
         OrigTop         =   255
         OrigRight       =   11535
         OrigBottom      =   540
         Max             =   1000
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin lvButton.lvButtons_H cmdReplenish 
         Height          =   375
         Left            =   12480
         TabIndex        =   58
         Top             =   660
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
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
         Image           =   "SPInsXstWHAgdao.frx":165C
         cBack           =   -2147483633
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Mos."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   6800
         TabIndex        =   38
         Top             =   720
         Width           =   495
      End
      Begin VB.Label lblOprQH 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "="
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   10560
         TabIndex        =   33
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
         TabIndex        =   29
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
         TabIndex        =   28
         Top             =   615
         Width           =   480
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "SQ:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
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
         TabIndex        =   23
         ToolTipText     =   "SALES QTY."
         Top             =   270
         Width           =   375
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "BO:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
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
         TabIndex        =   22
         ToolTipText     =   "BACK ORDER"
         Top             =   630
         Width           =   375
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "QH:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
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
         TabIndex        =   21
         ToolTipText     =   "QTY-ON-HAND"
         Top             =   270
         Width           =   375
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Sales Date To:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
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
         Top             =   750
         Width           =   1335
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Sales Date From:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   14
         Top             =   750
         Width           =   1455
      End
      Begin VB.Label lblSrchCode 
         Alignment       =   1  'Right Justify
         Caption         =   "Classification"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   80
         TabIndex        =   6
         Top             =   390
         Width           =   1815
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
      Image           =   "SPInsXstWHAgdao.frx":1BF6
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdOK 
      Height          =   375
      Left            =   12105
      TabIndex        =   8
      Top             =   6375
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
      Image           =   "SPInsXstWHAgdao.frx":2190
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdCancel 
      Height          =   375
      Left            =   13545
      TabIndex        =   9
      Top             =   6375
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
      Image           =   "SPInsXstWHAgdao.frx":272A
      cBack           =   -2147483633
   End
   Begin MSFlexGridLib.MSFlexGrid flxSP 
      Height          =   3945
      Left            =   135
      TabIndex        =   10
      Top             =   2340
      Width           =   14760
      _ExtentX        =   26035
      _ExtentY        =   6959
      _Version        =   393216
      Rows            =   1
      Cols            =   19
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   300
      BackColorBkg    =   16777215
      FocusRect       =   2
      ScrollBars      =   2
      AllowUserResizing=   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
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
      Left            =   135
      TabIndex        =   39
      Top             =   1845
      Width           =   14760
      _ExtentX        =   26035
      _ExtentY        =   900
      _Version        =   393216
      Cols            =   14
      FixedRows       =   2
      ScrollBars      =   0
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
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
      Left            =   0
      TabIndex        =   11
      Top             =   6360
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
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuInAct 
         Caption         =   "set to ""IN-ACTIVE"""
         Enabled         =   0   'False
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
Attribute VB_Name = "SPInsXstWHAgdao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsSupplier As New ADODB.Recordset
Dim ShowOpt As Boolean, SuppSrch As Boolean
Dim ActRow As Integer, ActCol As Integer
Public MySource As String

Private Sub cboFind_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call cmdFind_Click
End Sub

Private Sub cbomos_Change()
    cbomos_Click
End Sub
Private Sub cbomos_Click()
Dim mos As Integer
    mos = Val(cbomos.Text)
    pmos = "-" & mos
    SDfrom.Value = DateAdd("m", pmos, Date)
    If cbomos.Text = "All" Then
        SDfrom.Value = SPorderingWHAgdao.SalesDateFrom.Value
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

'Private Sub cboSupplier_Change()
'    Dim i As Long, j As Long
'    Dim strPartial As String, strTotal As String
'    Dim x As Integer
'
'    If SuppSrch Then
'        SuppSrch = False
'        Exit Sub
'    End If
'    With cboSupplier
'        strPartial = .Text
'        i = SendMessage(.hwnd, CB_FINDSTRING, -1, ByVal strPartial) 'If match found, append unmatched characters
'        If i <> -1 Then
'            strTotal = .List(i)
'            rsSupplier.MoveFirst
 '           rsSupplier.Move i
'            txtSuppCode.Text = IIf(IsNull(rsSupplier!SUPPLIERCODE), "", rsSupplier!SUPPLIERCODE)
'            chkAllSupp.Value = 0
'            j = Len(strTotal) - Len(strPartial)
'            If j <> 0 Then
'                SuppSrch = True
'                .SelText = Right$(strTotal, j)
'                .SelStart = Len(strPartial)
'                .SelLength = j
'            End If
'        Else
'            txtSuppCode.Text = ""
'        End If
'    End With

'End Sub

'Private Sub cboSupplier_Click()
'    Call cboSupplier_Change
'End Sub'

'Private Sub cboSupplier_GotFocus()
'   Dim Tmp
    
'   Const CB_SHOWDROPDOWN = &H14F
'   Tmp = SendMessage(cboSupplier.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
'End Sub

'Private Sub cboSupplier_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then SuppSrch = True
'End Sub

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
            cboSupplier.Text = "AA"
            txtSuppCode.Text = "AA"
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
    Dim MySel As String, i As Long, r As Long
    
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
   MySel = "SELECT  SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPAREPARTS.QTYONHAND, SPAREPARTS.BACKORDER, SPAREPARTS.SALES, SPAREPARTS.COST, SPAREPARTS.SPAREPARTID, SPAREPARTS.APPLICMODEL, SPAREPARTS.DESCRIPTION2 AS MYSTAT, SPAREPARTS.DLRPRICE, SPAREPARTS.MAXALLOW From SPAREPARTS WHERE (" & myW & ") AND REMARKS IN('ENGINE OIL','OIL','STIRE','TIRE','TIRES') ORDER BY SPAREPARTS.PARTNUMBER;"
    With rsSP
        .Open MySel, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            flxSP.Rows = 0
            Pbar.Value = 0
            Pbar.Visible = True
            Pbar.Max = .RecordCount * 2
            Do While Not .EOF
                'flxSP.AddItem vbTab & !PartNumber & vbTab & !Description & vbTab & !QtyOnHand & vbTab & IIf(IsNull(!BACKORDER), "", !BACKORDER) & vbTab & IIf(IsNull(!Sales), "", !Sales) & vbTab & vbTab & Format(!Cost, "#,##0.00") & vbTab & vbTab & !SPAREPARTID & vbTab & IIf(IsNull(!ApplicModel), "", !ApplicModel) & vbTab & IIf(IsNull(!MyStat), "", !MyStat)
                'flxSP.AddItem vbTab & !PartNumber & vbTab & !Description & vbTab & !QtyOnHand & vbTab & IIf(IsNull(!BACKORDER), "", !BACKORDER) & vbTab & IIf(IsNull(!Sales), "", !Sales) & vbTab & vbTab & vbTab & vbTab & vbTab & !SPAREPARTID & vbTab & Format(!Cost, "#,##0.00") & vbTab & IIf(IsNull(!ApplicModel), "", !ApplicModel) & vbTab & IIf(IsNull(!MyStat), "", !MyStat)
                'Desinged for Branch P.O
                '               0           1                   2                   3            4            5             6                           7                                               8                                   9            10          11           12           13         14                             15                          16                                                   17                                        17                  18
                flxSP.AddItem vbTab & !PartNumber & vbTab & !Description & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & !QtyOnHand & vbTab & IIf(IsNull(!BACKORDER), "", !BACKORDER) & vbTab & IIf(IsNull(!Sales), "", !Sales) & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & !SparepartID & vbTab & Format(!Cost, "#,##0.00") & vbTab & IIf(IsNull(!ApplicModel), "", !ApplicModel) & vbTab & IIf(IsNull(!MyStat), "", !MyStat) & vbTab & !DlrPrice & vbTab & !MAXALLOW
                Pbar.Value = Pbar.Value + 1
                
                'Hi-lights Column by rows
                flxSP.Col = 3
                flxSP.Row = Pbar.Value - 1
                flxSP.ColSel = 3
                flxSP.ColSel = 4
                flxSP.ColSel = 5
                flxSP.RowSel = Pbar.Value - 1
                flxSP.FillStyle = flexFillRepeat
                flxSP.CellBackColor = &HC0FFC0
                
                .MoveNext
                DoEvents
            Loop
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
            cmdOK.Top = flxSP.Top + flxSP.Height + 120
            cmdCancel.Top = flxSP.Top + flxSP.Height + 120
            Pbar.Top = flxSP.Top + flxSP.Height
            Me.Height = fmAnotation.Height + cmdOK.Top + cmdOK.Height + 580
        Else
            MsgBox "No Record Found!. Only Oil And Tires can be search!", vbInformation + vbOKOnly, "Validation"
            flxSP.Clear
            flxSP.Rows = 1
        End If
        .Close
    End With
    
    If flxSP.Rows = 1 And flxSP.TextMatrix(0, 1) = "" Then Exit Sub
    With flxSP
        '.ColWidth(10) = 0
        '.ColWidth(13) = 0
        For i = 0 To .Rows - 1
            Pbar.Value = IIf(Pbar.Value = Pbar.Max, Pbar.Max, Pbar.Value + 1)
            'Design for Branch
            'Get Sales of this spareparts
            .TextMatrix(i, 8) = GetSalesQty(Val(.TextMatrix(i, 13)))
            'Qry = "SELECT Sum(SPITR.OUT) AS SumOfOUT From SPITR " & _
                  "WHERE (((SPITR.TRANSDATE)>=#" & SDfrom.Value & "# And (SPITR.TRANSDATE)<=#" & SDto.Value & "#) AND ((SPITR.TRANSTYPE)='OUT') AND ((SPITR.REMARKS)='TRANSFER') AND ((SPITR.SPAREPARTID)=" & Val(.TextMatrix(i, 9)) & "));"
            'mySQL
            Qry = "SELECT Sum(SPITR.OUT) AS SumOfOUT From SPITR " & _
                  "WHERE SPITR.TRANSDATE>='" & Format(SDfrom.Value, "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(SDto.Value, "yyyy-mm-dd") & "' AND SPITR.TRANSTYPE='OUT' AND SPITR.REMARKS='TRANSFER' AND SPITR.SPAREPARTID=" & Val(.TextMatrix(i, 13)) & ";"
            'Design for Branch
            .TextMatrix(i, 9) = GetFieldValue(Qry)
            'Branch Designed
            If .TextMatrix(i, 16) = "INACTIVE" Then
                .Col = 1: .Row = i: .ColSel = 16: .RowSel = i: .FillStyle = flexFillRepeat: .CellForeColor = RGB(255, 0, 0)
            End If
        Next i
    End With
    'Pbar.Visible = False
    Call AddInCombo
    
    Pbar.Value = 0
    Pbar.Max = flxSP.Rows
    For r = 0 To flxSP.Rows - 1
    
        Pbar.Value = Pbar.Value + 1
        flxSP.TextMatrix(r, 10) = GetSalesQty6mos(Val(flxSP.TextMatrix(r, 13)))
        spSOQ = Val(flxSP.TextMatrix(r, 10)) - Val(flxSP.TextMatrix(r, 6)) - Val(flxSP.TextMatrix(r, 7))
        flxSP.TextMatrix(r, 11) = IIf(spSOQ < 0, 0, spSOQ)
        Qry = ""
        'Qry = "SELECT CASE WHEN (spareparts.QTYONHAND-IFNULL(spareparts.WSQTYONHAND,0)) <= 0 THEN 0 ELSE (spareparts.QTYONHAND-IFNULL(spareparts.WSQTYONHAND,0)) END as QTYONHAND From spareparts Where spareparts.SPAREPARTID = " & flxSP.TextMatrix(r, 13) & ";"
        Qry = "SELECT spareparts.QTYONHAND From spareparts Where spareparts.SPAREPARTID = " & flxSP.TextMatrix(r, 13) & ";"
        flxSP.TextMatrix(r, 3) = Val(GetFieldValueDVOWH(Qry))
        Qry = ""
        Qry = "SELECT spareparts.BACKORDER From spareparts Where spareparts.SPAREPARTID = " & flxSP.TextMatrix(r, 13) & ";"
        flxSP.TextMatrix(r, 4) = Val(GetFieldValueDVOWH(Qry))
        
        flxSP.TextMatrix(r, 5) = GetSalesQtyMain(Val(flxSP.TextMatrix(r, 13)))
    
    Next r

    Pbar.Visible = False
    'Call FillColorGrd
    'flxSP.Row = 0
    'flxSP.Col = 12
    If flxSP.Rows >= 0 Then
        'If flxSP.TextMatrix(1, 3) <> "" Then
        'fOR Branch
        If flxSP.TextMatrix(0, 6) <> "" Then
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

Function GetSalesQtyMain(MyID As Long) As Long
    Dim rsQty As New ADODB.Recordset
    Dim Qry As String
    Dim MySales As Long, MyRet As Long
    
    Qry = "SELECT Sum(SPITR.OUT) AS SumOfOUT From SPITR " & _
          "WHERE SPITR.TRANSDATE>='" & Format(DateAdd("m", -6, Date), "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(Date, "yyyy-mm-dd") & "' AND SPITR.TRANSTYPE='OUT' AND SPITR.REMARKS<>'TRANSFER' AND SPITR.SPAREPARTID=" & MyID & ";"
   
    With rsQty
        .Open Qry, DavaoWHCon, adOpenKeyset, adLockOptimistic
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
  
    MyRet = Val(GetFieldValueDVOWH(Qry))
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
        With SPorderingWHAgdao.flxParts
            If .TextMatrix(.Rows - 1, 1) = "" Then 'Remove last row if empty
                If .Rows = 2 Then
                    .Rows = 1
                Else
                    .RemoveItem .Rows - 1
                End If
            End If
            For i = 0 To flxSP.Rows - 1
                If flxSP.TextMatrix(i, 0) <> "" Then
                    If Not SPinFlx(flxSP.TextMatrix(i, 1)) Then
                        .AddItem .Rows & vbTab & flxSP.TextMatrix(i, 1) & vbTab & flxSP.TextMatrix(i, 2) & vbTab & Val(flxSP.TextMatrix(i, 18)) & vbTab & Val(flxSP.TextMatrix(i, 12))      '& vbTab & flxSP.TextMatrix(i, 14) & vbTab & vbTab & flxSP.TextMatrix(i, 13) & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & flxSP.TextMatrix(i, 15)
                        .TextMatrix(.Rows - 1, 5) = flxSP.TextMatrix(i, 14) 'cost
                        .TextMatrix(.Rows - 1, 6) = Format(Val(flxSP.TextMatrix(i, 12)) * flxSP.TextMatrix(i, 14), "###,###,#00.00")  'spid
                        .TextMatrix(.Rows - 1, 7) = flxSP.TextMatrix(i, 14) 'cost
                        .TextMatrix(.Rows - 1, 8) = flxSP.TextMatrix(i, 17)  'DlrPrice
                        .TextMatrix(.Rows - 1, 9) = Val(flxSP.TextMatrix(i, 13))  'spid
                        
                        If SysBranch = "LYR MOTORPARTS CORP. - WHOLESALE" Then
                            .TextMatrix(.Rows - 1, 4) = flxSP.TextMatrix(i, 17)
                        End If
                    End If
                End If
            Next i
            If .Rows = 1 Then .AddItem "1"
        End With
    End If
End Sub

Function SPinFlx(myPN As String) As Boolean
    Dim i As Integer
    
    If MySource = "SPORDERING" Then
        With SPorderingWHAgdao.flxParts
            For i = 0 To .Rows - 1
                'If myPN = .TextMatrix(i, 1) Then
                
                If myPN = .TextMatrix(i, 1) Then
                    SPinFlx = True
                    Exit Function
                End If
            Next i
        End With
        SPinFlx = False
        Exit Function
    End If
End Function

Private Sub cmdOK_Click()
    Call cmdInsert_Click
    Call SPorderingWHAgdao.ComputeTotal
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
    cmdOK.Top = flxSP.Top + flxSP.Height + 120
    cmdCancel.Top = flxSP.Top + flxSP.Height + 120
    Pbar.Top = flxSP.Top + flxSP.Height
    Me.Height = fmAnotation.Height + cmdOK.Top + cmdOK.Height + 580
    fmAnotation.Top = cmdOK.Top + cmdOK.Height + 1000 - fmAnotation.Height
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
    cmdOK.Top = flxSP.Top + flxSP.Height + 120
    cmdCancel.Top = flxSP.Top + flxSP.Height + 120
    Pbar.Top = flxSP.Top + flxSP.Height
    Me.Height = fmAnotation.Height + cmdOK.Top + cmdOK.Height + 580
    fmAnotation.Top = cmdOK.Top + cmdOK.Height + 1000 - fmAnotation.Height
    ShowOpt = True
    cboFind.Text = ""
End Sub

Private Sub cmdReplenish_Click()
    Dim rsRpnsh As New ADODB.Recordset
    Dim Qry As String, r As Long, spSOQ As Double
    
    cmdReplenish.Enabled = False
    cmdInsert.Enabled = False
    cmdOK.Enabled = False
    cmdCancel.Enabled = False
    Call CollAlign
    If cboSupplier.Text = "" Then
        'Qry = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPAREPARTS.QTYONHAND, SPAREPARTS.BACKORDER, SPAREPARTS.SALES, SPAREPARTS.COST, SPAREPARTS.SPAREPARTID, SPAREPARTS.APPLICMODEL, SPAREPARTS.DESCRIPTION2 AS MYSTAT " & _
              "From SPAREPARTS WHERE (((SPAREPARTS.QTYONHAND)" & lblOprQH.Caption & Val(txtQH.Text) & ") AND ((SPAREPARTS.BACKORDER)" & lblOprBO.Caption & Val(txtBO.Text) & ")) ORDER BY SPAREPARTS.PARTNUMBER;"
        'mySQL
        Qry = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPAREPARTS.QTYONHAND, SPAREPARTS.BACKORDER, SPAREPARTS.SALES, SPAREPARTS.COST, SPAREPARTS.SPAREPARTID, SPAREPARTS.APPLICMODEL, SPAREPARTS.DESCRIPTION2 AS MYSTAT, SPAREPARTS.DLRPRICE, SPAREPARTS.MAXALLOW " & _
              "From SPAREPARTS WHERE SPAREPARTS.QTYONHAND" & lblOprQH.Caption & Val(txtQH.Text) & " AND SPAREPARTS.BACKORDER" & lblOprBO.Caption & Val(txtBO.Text) & " AND REMARKS IN('ENGINE OIL','OIL','STIRE','TIRE','TIRES') ORDER BY SPAREPARTS.PARTNUMBER;"

    Else
        If UpDsearch.Value = 1 Then 'CLASSIFICATION
            Qry = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPAREPARTS.QTYONHAND, SPAREPARTS.BACKORDER, SPAREPARTS.SALES, SPAREPARTS.COST, SPAREPARTS.SPAREPARTID, SPAREPARTS.APPLICMODEL, SPAREPARTS.DESCRIPTION2 AS MYSTAT, SPAREPARTS.DLRPRICE, SPAREPARTS.MAXALLOW " & _
                  "From SPAREPARTS WHERE SPAREPARTS.QTYONHAND" & lblOprQH.Caption & Val(txtQH.Text) & " AND SPAREPARTS.BACKORDER" & lblOprBO.Caption & Val(txtBO.Text) & " AND SPAREPARTS.REMARKS='" & cboSupplier.Text & "' ORDER BY SPAREPARTS.PARTNUMBER;"
        
        ElseIf UpDsearch.Value = 2 Then 'ABC APPROACH
            'Qry = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPAREPARTS.QTYONHAND, SPAREPARTS.BACKORDER, SPAREPARTS.SALES, SPAREPARTS.COST, SPAREPARTS.SPAREPARTID, SPAREPARTS.APPLICMODEL, SPAREPARTS.DESCRIPTION2 AS MYSTAT, SPAREPARTS.DLRPRICE " & _
                  "From SPAREPARTS WHERE SPAREPARTS.QTYONHAND" & lblOprQH.Caption & Val(txtQH.Text) & " AND SPAREPARTS.BACKORDER" & lblOprBO.Caption & Val(txtBO.Text) & " AND SPAREPARTS.SUPPLIERCODE='" & txtSuppCode.Text & "' ORDER BY SPAREPARTS.PARTNUMBER;"
            
            Qry = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPAREPARTS.QTYONHAND, SPAREPARTS.BACKORDER, SPAREPARTS.SALES, SPAREPARTS.COST, SPAREPARTS.SPAREPARTID, SPAREPARTS.APPLICMODEL, SPAREPARTS.DESCRIPTION2 AS MYSTAT, SPAREPARTS.DLRPRICE, SPAREPARTS.MAXALLOW " & _
                  "From SPAREPARTS WHERE SPAREPARTS.QTYONHAND" & lblOprQH.Caption & Val(txtQH.Text) & " AND SPAREPARTS.BACKORDER" & lblOprBO.Caption & Val(txtBO.Text) & " AND REMARKS IN('ENGINE OIL','OIL','STIRE','TIRE','TIRES') ORDER BY SPAREPARTS.PARTNUMBER;"
        ElseIf UpDsearch.Value = 3 Then 'APPLICABLE MODEL
            Qry = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPAREPARTS.QTYONHAND, SPAREPARTS.BACKORDER, SPAREPARTS.SALES, SPAREPARTS.COST, SPAREPARTS.SPAREPARTID, SPAREPARTS.APPLICMODEL, SPAREPARTS.DESCRIPTION2 AS MYSTAT, SPAREPARTS.DLRPRICE, SPAREPARTS.MAXALLOW " & _
                  "From SPAREPARTS WHERE SPAREPARTS.QTYONHAND" & lblOprQH.Caption & Val(txtQH.Text) & " AND SPAREPARTS.BACKORDER" & lblOprBO.Caption & Val(txtBO.Text) & " AND SPAREPARTS.APPLICMODEL='" & cboSupplier.Text & "' AND REMARKS IN('ENGINE OIL','OIL','STIRE','TIRE','TIRES') ORDER BY SPAREPARTS.PARTNUMBER;"
        End If
    End If
    
    With rsRpnsh
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            flxSP.Rows = 0
            Pbar.Value = 0
            Pbar.Visible = True
            Pbar.Max = .RecordCount * 2
            Do While Not .EOF
                      '1-PART NUMBER, 2-DESCRIPTION, 3- QH, 4- BO, 5-SQ, 6-COST, 7-ORDER, 8-SPID, 9, applicmodel, 10-MYSTAT
                      '1-PART NUMBER, 2-DESCRIPTION, 3- QH, 4- BO, 5-SQ, 6-ST, 7-COST, 8-ORDER, 9-SPID, 10- applicmodel, 11-MYSTAT
                'Previous code
                'flxSP.AddItem vbTab & !PartNumber & vbTab & !Description & vbTab & !QtyOnHand & vbTab & IIf(IsNull(!BACKORDER), "", !BACKORDER) & vbTab & IIf(IsNull(!Sales), "", !Sales) & vbTab & vbTab & Format(!Cost, "#,##0.00") & vbTab & vbTab & !SPAREPARTID & vbTab & IIf(IsNull(!ApplicModel), "", !ApplicModel) & vbTab & IIf(IsNull(!MyStat), "", !MyStat)
                '                      1-PARTNUMBER         2-DESCRIPTION          3-QH                    4-BO                                         5-SQ                                  6-ST    7-6MOS PSQ     8-SOQ  9-ORDER     10-SPID              11- CWOV                        12-APPMOD                      13-MYSTAT
                'flxSP.AddItem vbTab & !PartNumber & vbTab & !Description & vbTab & !QtyOnHand & vbTab & IIf(IsNull(!BACKORDER), "", !BACKORDER) & vbTab & IIf(IsNull(!Sales), "", !Sales) & vbTab & vbTab & vbTab & vbTab & vbTab & !SPAREPARTID & vbTab & Format(!Cost, "#,##0.00") & vbTab & IIf(IsNull(!ApplicModel), "", !ApplicModel) & vbTab & IIf(IsNull(!MyStat), "", !MyStat)
                'flxSP.AddItem vbTab & !PartNumber & vbTab & !Description & vbTab & !QtyOnHand & vbTab & IIf(IsNull(!BACKORDER), "", !BACKORDER) & vbTab & IIf(IsNull(!Sales), "", !Sales) & vbTab & vbTab & vbTab & vbTab & vbTab & !SPAREPARTID & vbTab & Format(!Cost, "#,##0.00") & vbTab & IIf(IsNull(!ApplicModel), "", !ApplicModel) & vbTab & IIf(IsNull(!MyStat), "", !MyStat)
                
                'Designed for Branch
                          '0-Blank Col 1-SPNUMBER           2-Description          3-QH        4-BO         5-SQ          6-QH                 7-BO                                                        8-SQ                           9-ST         10-6MOS     11-6m        12-SOQ       13-ORDER         14-spid                        15-COST                         16-APPMOD                                       17-STATUS
                flxSP.AddItem vbTab & !PartNumber & vbTab & !Description & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & !QtyOnHand & vbTab & IIf(IsNull(!BACKORDER), "", !BACKORDER) & vbTab & IIf(IsNull(!Sales), "", !Sales) & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & !SparepartID & vbTab & Format(!Cost, "#,##0.00") & vbTab & IIf(IsNull(!ApplicModel), "", !ApplicModel) & vbTab & IIf(IsNull(!MyStat), "", !MyStat) & vbTab & !DlrPrice & vbTab & !MAXALLOW
                Pbar.Value = Pbar.Value + 1
                
                'Hi-lights Column by rows
                flxSP.Col = 3
                flxSP.Row = Pbar.Value - 1
                flxSP.ColSel = 3
                flxSP.ColSel = 4
                flxSP.ColSel = 5
                flxSP.RowSel = Pbar.Value - 1
                flxSP.FillStyle = flexFillRepeat
                flxSP.CellBackColor = &HC0FFC0
                
                .MoveNext
                DoEvents
            Loop
            flxSP.Visible = True
            If ShowOpt Then
                flxSPH.Top = FrameOpt.Top + FrameOpt.Height
                flxSP.Top = FrameOpt.Top + FrameOpt.Height + flxSPH.Height
            Else
                flxSPH.Top = cmdOptHide.Top + cmdOptHide.Height + 100
                flxSP.Top = FrameOpt.Height + flxSPH.Height
            End If
            cmdInsert.Top = flxSP.Top + flxSP.Height + 120
            cmdOK.Top = flxSP.Top + flxSP.Height + 120
            cmdCancel.Top = flxSP.Top + flxSP.Height + 120
            Pbar.Top = flxSP.Top + flxSP.Height
            Me.Height = fmAnotation.Height + cmdOK.Top + cmdOK.Height + 580
           
        Else
        '    With flxSP
                flxSP.Clear
                flxSP.Rows = 1
                cmdReplenish.Enabled = True
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
       
    For i = flxSP.Rows - 1 To 0 Step -1
        'flxSP.TextMatrix(i, 5) = GetSalesQty(Val(flxSP.TextMatrix(i, 10)))
        'Designed for Branch
        flxSP.TextMatrix(i, 8) = GetSalesQty(Val(flxSP.TextMatrix(i, 13)))
        
        'Qry = "SELECT Sum(SPITR.OUT) AS SumOfOUT From SPITR " & _
              "WHERE (((SPITR.TRANSDATE)>=#" & SDfrom.Value & "# And (SPITR.TRANSDATE)<=#" & SDto.Value & "#) AND ((SPITR.TRANSTYPE)='OUT') AND ((SPITR.REMARKS)='TRANSFER') AND ((SPITR.SPAREPARTID)=" & Val(flxSP.TextMatrix(i, 9)) & "));"
        'mySQL
        Qry = "SELECT Sum(SPITR.OUT) AS SumOfOUT From SPITR " & _
              "WHERE SPITR.TRANSDATE>='" & Format(SDfrom.Value, "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(SDto.Value, "yyyy-mm-dd") & "' AND SPITR.TRANSTYPE='OUT' AND SPITR.REMARKS='TRANSFER' AND SPITR.SPAREPARTID=" & Val(flxSP.TextMatrix(i, 13)) & ";"
        
        flxSP.TextMatrix(i, 9) = GetFieldValue(Qry)
        Pbar.Value = IIf(Pbar.Value = Pbar.Max, Pbar.Max, Pbar.Value + 1)
        DoEvents
        If flxSP.TextMatrix(i, 16) = "INACTIVE" Then
            flxSP.Col = 1: flxSP.Row = i: flxSP.ColSel = 16: flxSP.RowSel = i: flxSP.FillStyle = flexFillRepeat: flxSP.CellForeColor = RGB(255, 0, 0)
        End If
        
        Select Case OptrSQ.Value
            Case 0
                If Val(flxSP.TextMatrix(i, 8)) <> Val(txtSQ.Text) Then 'flxSP.TextMatrix(i, 9) = "x"
                    If flxSP.Rows = 1 Then
                        flxSP.Rows = 1
                    Else
                        flxSP.RemoveItem i
                    End If
                End If
            Case 1
                If Val(flxSP.TextMatrix(i, 8)) <= Val(txtSQ.Text) Then
                    If flxSP.Rows = 1 Then
                        flxSP.Rows = 1
                    Else
                        flxSP.RemoveItem i
                    End If
                End If
            Case 2
                If Val(flxSP.TextMatrix(i, 8)) < Val(txtSQ.Text) Then
                    If flxSP.Rows = 1 Then
                        flxSP.Rows = 1
                      
                    Else
                        flxSP.RemoveItem i
                    End If
                End If
            Case 3
                If Val(flxSP.TextMatrix(i, 8)) >= Val(txtSQ.Text) Then
                    If flxSP.Rows = 1 Then
                        flxSP.Rows = 1
              
                    Else
                        flxSP.RemoveItem i
                    End If
                End If
            Case 4
                If Val(flxSP.TextMatrix(i, 8)) > Val(txtSQ.Text) Then
                    If flxSP.Rows = 1 Then
                        flxSP.Rows = 1
                  
                    Else
                        flxSP.RemoveItem i
                    End If
                End If
            Case 5
                If Val(flxSP.TextMatrix(i, 8)) = Val(txtSQ.Text) Then
                    If flxSP.Rows = 1 Then
                        flxSP.Rows = 1
                     
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
    
    Pbar.Value = 0
    Pbar.Max = flxSP.Rows
    For r = 0 To flxSP.Rows - 1
    
        Pbar.Value = Pbar.Value + 1
        flxSP.TextMatrix(r, 10) = GetSalesQty6mos(Val(flxSP.TextMatrix(r, 13)))
        spSOQ = Val(flxSP.TextMatrix(r, 10)) - Val(flxSP.TextMatrix(r, 6)) - Val(flxSP.TextMatrix(r, 7))
        flxSP.TextMatrix(r, 11) = IIf(spSOQ < 0, 0, spSOQ)
        Qry = ""
        Qry = "SELECT spareparts.QTYONHAND From spareparts Where spareparts.SPAREPARTID = " & flxSP.TextMatrix(r, 13) & ";"
        flxSP.TextMatrix(r, 3) = Val(GetFieldValueDVOWH(Qry))
        Qry = ""
        Qry = "SELECT spareparts.BACKORDER From spareparts Where spareparts.SPAREPARTID = " & flxSP.TextMatrix(r, 13) & ";"
        flxSP.TextMatrix(r, 4) = Val(GetFieldValueDVOWH(Qry))
        
        flxSP.TextMatrix(r, 5) = GetSalesQtyMain(Val(flxSP.TextMatrix(r, 13)))
        
    Next r
    Pbar.Visible = False
    
    'flxSP.Row = 0
    'flxSP.Col = 12
    If flxSP.Rows >= 1 Then
        If flxSP.TextMatrix(0, 6) <> "" Then
            flxSP.SetFocus
        Else
            cboFind.SetFocus
        End If
    Else
        cboFind.SetFocus
    End If
       
   cmdReplenish.Enabled = True
   cmdInsert.Enabled = True
   cmdOK.Enabled = True
   cmdCancel.Enabled = True
End Sub

Private Sub flxSP_Click()
Dim ClSel As Integer

    ActRow = flxSP.Row
    ActCol = flxSP.Col
    'If ActRow = 0 Or ActCol <> 12 Then Exit Sub
    If ActCol <> 12 Then Exit Sub
    
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
   
     'HiLights If Select
    flxSP.Col = 0
    flxSP.ColSel = 11
    flxSP.RowSel = ActRow
    flxSP.FillStyle = flexFillRepeat
    
End Sub

Private Sub flxSP_DblClick()
        If flxSP.TextMatrix(flxSP.Row, 0) = "" And flxSP.TextMatrix(flxSP.Row, 1) <> "" Then
            Call mnuMark_Click
        Else
            Call mnuUnmark_Click
        End If
End Sub

Private Sub flxSP_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
        If flxSP.TextMatrix(flxSP.Row, 0) = "" And flxSP.TextMatrix(flxSP.Row, 1) <> "" Then
            Call mnuMark_Click
        Else
            Call mnuUnmark_Click
        End If
    End If
End Sub


Private Sub flxSP_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call flxSP_Click
End Sub

Private Sub flxSP_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
    If flxSP.TextMatrix(flxSP.Row, 1) = "" Then Exit Sub
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

Private Sub flxSP_Scroll()
    txtOrder.Visible = False
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
    Call cmdOptHide_Click
    
    With SPorderingWHAgdao
        If .LockFlx = "True" Then
            cmdInsert.Enabled = False
            cmdOK.Enabled = False
        End If
        'SDfrom.Value = .SalesDateFrom.Value
        SDfrom.Value = DateAdd("m", -6, Date)
        SDto.Value = .SalesDateTo.Value
        cbomos.Text = 6
        For m = 1 To 6
            cbomos.AddItem m
        Next m
        cbomos.AddItem "All"
    End With
    
'    Qry = "SELECT SUPPLIER.SUPPLIERID, SUPPLIER.SUPPLIERNAME, SUPPLIER.SUPPLIERCODE " & _
'          "From SUPPLIER ORDER BY SUPPLIER.SUPPLIERNAME;"
'    With rsSupplier
'        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
'        If .RecordCount > 0 Then
'            .MoveFirst
'            cboSupplier.Clear
'            Do While Not .EOF
'                'cboSupplier.AddItem !SUPPLIERNAME
'                cboSupplier.AddItem !SUPPLIERCODE
'                .MoveNext
'            Loop
'        End If
'    End With
    '----------------------------------------
    'Don't Display Code
    'cboSupplier.Text = SPordering.cboSupplier.Text
    'txtSuppCode.Text = SPordering.txtScode.Text
    '----------------------------------------
    'Set Search Description
    UpDsearch.Value = 1
    UpDRef.Value = 1
    Call ResizeComplete
    'Connect use Moto-r Main Program
    Call ConnectDavaoWH
    'LYRmain.sBar.Panels(6).Text = "Find Spare Parts"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'LYRmain.sBar.Panels(6).Text = ""
    DavaoWHCon.Close
    If rsSupplier.State = adStateOpen Then rsSupplier.Close
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

Private Sub txtOrder_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        '---------------------------------------------------------------------------------------------
        'Lock Ordering from Suggested Order Quantity
        'If Val(txtOrder.Text) > Val(flxSP.TextMatrix(ActRow, 11)) And UserID <> 4 And VerifySP6Mos = True Then
        '   If VerifySP6Mos = True Then
        '        flxSP.TextMatrix(ActRow, 12) = ""
        '        flxSP.TextMatrix(ActRow, 0) = ""
        '        txtOrder.Visible = False
        '        flxSP.Col = 12
        '        Exit Sub
        '    Else
        '        MsgBox ("Less than 6 mos. Sparepart"), vbExclamation, "Order"
        '    End If
        'ElseIf Val(txtOrder.Text) > Val(flxSP.TextMatrix(ActRow, 11)) And UserID = 4 Then
        '    If VerifySP6Mos = True Then
        '        MsgBox ("Order Greater than Suggested Order Quantity (SOQ)"), vbExclamation, "Order"
        '        txtOrder.SetFocus
        '        Exit Sub
        '    Else
        '        MsgBox ("Less than 6 mos. Sparepart"), vbExclamation, "Order"
        '    End If
        'ElseIf Val(txtOrder.Text) < Val(flxSP.TextMatrix(ActRow, 11)) And Val(txtOrder.Text) <> 0 Then
        '    MsgBox ("Order Lesser than Suggested Order Quantity (SOQ)"), vbExclamation, "Order"
        '    txtOrder.SetFocus
        '    Exit Sub
        'End If
        'txtOrder.Text = flxSP.TextMatrix(ActRow, 12)
        'flxSP.SetFocus
        'flxSP.Row = ActRow
        'flxSP.Col = 12
        'txtOrder.Visible = False
        
        '---------------------------------------------------------------------------------------------
        'UnLock Ordering from Suggested Order Quantity
        
        If Val(txtOrder.Text) > Val(flxSP.TextMatrix(ActRow, 11)) Then
            MsgBox ("Order Greater than Suggested Order Quantity (SOQ)"), vbExclamation, "Order"
            txtOrder.SetFocus
            Exit Sub
        ElseIf Val(txtOrder.Text) < Val(flxSP.TextMatrix(ActRow, 11)) And Val(txtOrder.Text) <> 0 Then
            MsgBox ("Order Lesser than Suggested Order Quantity (SOQ)"), vbExclamation, "Order"
            txtOrder.SetFocus
            Exit Sub
        End If
        txtOrder.Text = flxSP.TextMatrix(ActRow, 12)
        flxSP.SetFocus
        flxSP.Row = ActRow
        flxSP.Col = 12
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
            flxSP.Col = 12
   End If
End Sub
Private Sub txtOrder_LostFocus()
    With flxSP
        '-----------------------------------------------------------------------------------------------
        'Lock Ordering for Suggested Order Quantity
        
        'If Val(txtOrder.Text) > Val(.TextMatrix(ActRow, 11)) And UserID <> 4 Then
        '    If VerifySP6Mos = True Then
        '        MsgBox ("Can't add Order Quantity" & vbCrLf & _
        '                "which Greater than Suggested Order Quantity (SOQ)"), vbExclamation, "Order"
        '        flxSP.TextMatrix(ActRow, 12) = ""
        '        flxSP.TextMatrix(ActRow, 0) = ""
        '        If txtOrder.Visible = True Then txtOrder.SetFocus: SendKeys "{Home}+{End}"
        '        Exit Sub
        '    Else
        '        MsgBox ("Less than 6 mos. Sparepart"), vbExclamation, "Order"
        '    End If
        'ElseIf Val(txtOrder.Text) > Val(flxSP.TextMatrix(ActRow, 11)) And UserID = 4 Then
        '    If VerifySP6Mos = True Then
        '        MsgBox ("Order Greater than Suggested Order Quantity (SOQ)"), vbExclamation, "Order"
        '    Else
        '        MsgBox ("Less than 6 mos. Sparepart"), vbExclamation, "Order"
        '    End If
        'ElseIf Val(txtOrder.Text) < Val(flxSP.TextMatrix(ActRow, 11)) And Val(txtOrder.Text) <> 0 Then
        '    MsgBox ("Order Lesser than Suggested Order Quantity (SOQ)"), vbExclamation, "Order"
        'End If
        
        '.TextMatrix(ActRow, 12) = txtOrder.Text
        'txtOrder.Visible = False
        'If Val(.TextMatrix(ActRow, 12)) > 0 Then
        '    .Row = ActRow
        '    .Col = 0
        '    .CellFontName = "WINGDINGS"
        '    .CellFontSize = 12
        '    .Text = lblChk.Caption
        '    .CellAlignment = 4
        '    .Col = 12
        'Else
        '    .TextMatrix(ActRow, 12) = ""
        '    .TextMatrix(ActRow, 0) = ""
        '    .Col = 12
        'End If
        
        '-----------------------------------------------------------------------------------------------
        'UnLock Ordering for Suggested Order Quantity
        If Val(txtOrder.Text) > Val(flxSP.TextMatrix(ActRow, 11)) Then
            MsgBox ("Order Greater than Suggested Order Quantity (SOQ)"), vbExclamation, "Order"
        ElseIf Val(txtOrder.Text) < Val(flxSP.TextMatrix(ActRow, 11)) And Val(txtOrder.Text) <> 0 Then
            MsgBox ("Order Lesser than Suggested Order Quantity (SOQ)"), vbExclamation, "Order"
        End If
        
        .TextMatrix(ActRow, 12) = txtOrder.Text
        txtOrder.Visible = False
        If Val(.TextMatrix(ActRow, 12)) > 0 Then
            .Row = ActRow
            .Col = 0
            .CellFontName = "WINGDINGS"
            .CellFontSize = 12
            .Text = lblChk.Caption
            .CellAlignment = 4
            .Col = 12
        Else
            .TextMatrix(ActRow, 12) = ""
            .TextMatrix(ActRow, 0) = ""
            .Col = 12
        End If
     End With
End Sub
Private Sub UpDRef_Change()
    If UpDRef.Value = 0 Then
        lblRefSearch.Caption = "FIND PARTNUMBER"
    Else
        lblRefSearch.Caption = "FIND PARTS DESC."
    End If
End Sub
Private Sub ResizeComplete()
    Dim aCol As Integer
    
    For aCol = 1 To 12
        flxSPH.ColAlignment(aCol) = 4
    Next aCol
    With flxSPH
       
        
        .TextMatrix(0, 1) = "S P A R E P A R T S"
        .TextMatrix(0, 2) = "S P A R E P A R T S"
        .TextMatrix(0, 3) = "MRC WAREHOUSE AGDAO"
        .TextMatrix(0, 4) = "MRC WAREHOUSE AGDAO"
        .TextMatrix(0, 5) = "MRC WAREHOUSE AGDAO"
        
        'FOR MC PARTS DEPOT
        .TextMatrix(0, 6) = SysBranch '"M C  P A R T S  D E P 0 T"
        .TextMatrix(0, 7) = SysBranch '"M C  P A R T S  D E P 0 T"
        .TextMatrix(0, 8) = SysBranch '"M C  P A R T S  D E P 0 T"
        .TextMatrix(0, 9) = SysBranch '"M C  P A R T S  D E P 0 T"
        .TextMatrix(0, 10) = SysBranch '"M C  P A R T S  D E P 0 T"
        
        
        'FOR MOTO-R van
'        .TextMatrix(0, 6) = "V A N  S E L L I N G II"
'        .TextMatrix(0, 7) = "V A N  S E L L I N G II"
'        .TextMatrix(0, 8) = "V A N  S E L L I N G II"
'        .TextMatrix(0, 9) = "V A N  S E L L I N G II"
'        .TextMatrix(0, 10) = "V A N  S E L L I N G II"
        
        'FOR MOTO-R ULAS
'        .TextMatrix(0, 6) = "M O T O - R  U L A S"
'        .TextMatrix(0, 7) = "M O T O - R  U L A S"
'        .TextMatrix(0, 8) = "M O T O - R  U L A S"
'        .TextMatrix(0, 9) = "M O T O - R  U L A S"
'        .TextMatrix(0, 10) = "M O T O - R  U L A S"
     
        'FOR MRC PANABO
'        .TextMatrix(0, 6) = "M R C  P A N A B O"
'        .TextMatrix(0, 7) = "M R C  P A N A B O"
'        .TextMatrix(0, 8) = "M R C  P A N A B O"
'        .TextMatrix(0, 9) = "M R C  P A N A B O"
'        .TextMatrix(0, 10) = "M R C  P A N A B O"
     
        'FOR MRC SOBRECARY
'        .TextMatrix(0, 6) = "M R C 1 - SOBRECARY"
'        .TextMatrix(0, 7) = "M R C 1 - SOBRECARY"
'        .TextMatrix(0, 8) = "M R C 1 - SOBRECARY"
'        .TextMatrix(0, 9) = "M R C 1 - SOBRECARY"
'        .TextMatrix(0, 10) = "M R C 1 - SOBRECARY"
     
        'FOR MOTO-R van
'        .TextMatrix(0, 6) = "M R C   G E N S A N"
'        .TextMatrix(0, 7) = "M R C   G E N S A N"
'        .TextMatrix(0, 8) = "M R C   G E N S A N"
'        .TextMatrix(0, 9) = "M R C   G E N S A N"
'        .TextMatrix(0, 10) = "M R C   G E N S A N"
     
     
        .TextMatrix(0, 11) = "PURCHASE ORDER"
        .TextMatrix(0, 12) = "PURCHASE ORDER"
        .TextMatrix(0, 13) = "PURCHASE ORDER"
        
        .TextMatrix(1, 1) = "PARTNUMBER"
        .TextMatrix(1, 2) = "DESCRIPTION"
        .TextMatrix(1, 3) = "QH"
        .TextMatrix(1, 4) = "BO"
        .TextMatrix(1, 5) = "SQ"
        .TextMatrix(1, 6) = "QH"
        .TextMatrix(1, 7) = "BO"
        .TextMatrix(1, 8) = "SQ"
        .TextMatrix(1, 9) = "ST"
        .TextMatrix(1, 10) = "6m.SQ"
        .TextMatrix(1, 11) = "SOQ"
        .TextMatrix(1, 12) = "ORDER"
        .TextMatrix(1, 13) = "APPMOD"
        
        
        .MergeCells = flexMergeFree
        .MergeRow(0) = True
        .MergeCol(1) = True
        .MergeCol(2) = True
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
        
        .ColWidth(0) = 300
        .ColWidth(1) = (.Width - 400) * 0.13
        .ColWidth(2) = (.Width - 400) * 0.3
        .ColWidth(3) = (.Width - 400) * 0.05
        .ColWidth(4) = (.Width - 400) * 0.05
        .ColWidth(5) = (.Width - 400) * 0.05
        .ColWidth(6) = (.Width - 400) * 0.05
        .ColWidth(7) = (.Width - 400) * 0.05
        .ColWidth(8) = (.Width - 400) * 0.05
        .ColWidth(9) = (.Width - 400) * 0.05
        .ColWidth(10) = (.Width - 400) * 0.05
        .ColWidth(11) = (.Width - 400) * 0.05
        .ColWidth(12) = (.Width - 400) * 0.065
        .ColWidth(13) = (.Width - 400) * 0.06
        
        flxSP.ColWidth(0) = 300
        flxSP.ColWidth(1) = (.Width - 400) * 0.13
        flxSP.ColWidth(2) = (.Width - 400) * 0.3
        flxSP.ColWidth(3) = (.Width - 400) * 0.05
        flxSP.ColWidth(4) = (.Width - 400) * 0.05
        flxSP.ColWidth(5) = (.Width - 400) * 0.05
        flxSP.ColWidth(6) = (.Width - 400) * 0.05
        flxSP.ColWidth(7) = (.Width - 400) * 0.05
        flxSP.ColWidth(8) = (.Width - 400) * 0.05
        flxSP.ColWidth(9) = (.Width - 400) * 0.05
        flxSP.ColWidth(10) = (.Width - 400) * 0.05
        flxSP.ColWidth(11) = (.Width - 400) * 0.05
        flxSP.ColWidth(12) = (.Width - 400) * 0.065
        flxSP.ColWidth(15) = (.Width - 400) * 0.06
        flxSP.TextMatrix(0, 13) = ""
        flxSP.TextMatrix(0, 14) = ""
        'flxSP.TextMatrix(0, 15) = ""
        flxSP.TextMatrix(0, 16) = ""
       flxSP.TextMatrix(0, 17) = ""
       
        flxSP.ColWidth(13) = 0
        flxSP.ColWidth(14) = 0
        'flxSP.ColWidth(15) = 0
        flxSP.ColWidth(16) = 0
       flxSP.ColWidth(17) = 0
        flxSP.ColWidth(18) = 0 '(.Width - 400) * 0.06
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
Private Sub FillColorL6MOS(cRow As Long)

Dim ccol As Integer
    flxSP.Col = 0
    flxSP.Row = 0
    
    For ccol = 0 To 16
        flxSP.ColSel = ccol
    Next ccol
    
    flxSP.RowSel = cRow
    flxSP.FillStyle = flexFillRepeat: flxSP.CellBackColor = &H8080FF
    
End Sub
Private Sub ClearFRow()
Dim cRw As Integer
    For cRw = 0 To 18
        flxSP.TextMatrix(0, cRw) = ""
    Next cRw
End Sub
Function VerifySP6Mos() As Boolean

    Dim Qry As String, FDate As String, VerDate As String

    Qry = "SELECT MIN(spitr.TRANSDATE) As FIRSTORDER From spitr Where spitr.SPAREPARTID = " & Val(flxSP.TextMatrix(flxSP.Row, 13)) & " AND spitr.TRANSTYPE = 'RECEIVE';"
    FDate = GetFieldValue(Qry)
        If FDate = "" Then: VerifySP6Mos = False: Exit Function
        VerDate = DateAdd("m", 6, FDate)
            If CDate(VerDate) <= Date Then
                VerifySP6Mos = True
            Else
                VerifySP6Mos = False
            End If
    
End Function

Private Sub UpDsearch_Change()
    Dim rsSrch As New ADODB.Recordset
    Dim Qry As String
    
    If UpDsearch.Value = 0 Then
        lblSrchCode.Caption = "All Spareparts"
        cboSupplier.Clear
    ElseIf UpDsearch.Value = 1 Then
        lblSrchCode.Caption = "Classification"
        Qry = "SELECT DISTINCT spareparts.REMARKS From spareparts Where REMARKS IN ('ENGINE OIL','OIL','STIRE','TIRE','TIRES') Order By spareparts.REMARKS;"
        With rsSrch
            .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
            If .RecordCount > 0 Then
                .MoveFirst
                cboSupplier.Clear
                Do While Not .EOF
                    cboSupplier.AddItem !Remarks
                    .MoveNext
                Loop
            Else
                cboSupplier.Clear
            End If
        End With
    ElseIf UpDsearch.Value = 2 Then
        lblSrchCode = "ABC Approach"
        cboSupplier.Clear
        cboSupplier.AddItem "A-APPROACH"
        cboSupplier.AddItem "B-APPROACH"
        cboSupplier.AddItem "C-APPROACH"
        cboSupplier.AddItem "D-APPROACH"
        cboSupplier.AddItem "E-APPROACH"
    ElseIf UpDsearch.Value = 3 Then
        lblSrchCode = "Applicable Model"
        Qry = "SELECT DISTINCT spareparts.APPLICMODEL From spareparts Where spareparts.APPLICMODEL IS NOT NULL Order By spareparts.APPLICMODEL"
        With rsSrch
            .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
            If .RecordCount > 0 Then
                .MoveFirst
                cboSupplier.Clear
                Do While Not .EOF
                    cboSupplier.AddItem !ApplicModel
                    .MoveNext
                Loop
            Else
                cboSupplier.Clear
            End If
        End With
    
    End If
End Sub
