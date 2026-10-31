VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form DMCMadj 
   Caption         =   "DMCM - Accounts Receivables"
   ClientHeight    =   7935
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   13095
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7935
   ScaleWidth      =   13095
   Tag             =   "XYF"
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtRefType 
      Height          =   285
      Left            =   2280
      TabIndex        =   97
      Top             =   8280
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtJOid 
      Height          =   285
      Left            =   1440
      TabIndex        =   96
      Top             =   8280
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtCustID 
      Height          =   285
      Left            =   360
      TabIndex        =   91
      Top             =   8280
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtPrepBy 
      Height          =   315
      Left            =   2040
      MaxLength       =   50
      TabIndex        =   88
      Tag             =   "XYF"
      Top             =   6960
      Width           =   3615
   End
   Begin VB.TextBox txtAprvBy 
      Height          =   315
      Left            =   2040
      MaxLength       =   50
      TabIndex        =   87
      Tag             =   "XYF"
      Top             =   7320
      Width           =   3615
   End
   Begin VB.Frame Frame3 
      Height          =   1455
      Left            =   480
      TabIndex        =   22
      Tag             =   "XYF"
      Top             =   1800
      Width           =   12375
      Begin VB.TextBox txtNewStatus 
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
         Left            =   9720
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   94
         Tag             =   "XYF"
         Top             =   600
         Width           =   2295
      End
      Begin VB.TextBox txtNewBal 
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
         Left            =   9720
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   93
         Tag             =   "XYF"
         Top             =   240
         Width           =   2295
      End
      Begin MSComCtl2.UpDown UpdDmCm 
         Height          =   315
         Left            =   2040
         TabIndex        =   29
         Tag             =   "XYF"
         Top             =   600
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   556
         _Version        =   393216
         Value           =   1
         Max             =   2
         Min             =   1
         Enabled         =   -1  'True
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
         Height          =   315
         Left            =   2280
         MaxLength       =   250
         TabIndex        =   27
         Tag             =   "XYF"
         Top             =   960
         Width           =   9735
      End
      Begin VB.TextBox txtAmount 
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
         Left            =   2280
         TabIndex        =   25
         Tag             =   "XYF"
         Top             =   600
         Width           =   2415
      End
      Begin VB.TextBox txtRefNo 
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
         Left            =   2280
         MaxLength       =   50
         TabIndex        =   24
         Tag             =   "XYF"
         Top             =   240
         Width           =   2415
      End
      Begin MSComCtl2.UpDown UpdInv 
         Height          =   315
         Left            =   2040
         TabIndex        =   98
         Tag             =   "XYF"
         Top             =   240
         Visible         =   0   'False
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   556
         _Version        =   393216
         Value           =   1
         Max             =   1
         Enabled         =   -1  'True
      End
      Begin VB.Label lblCrntStat 
         Alignment       =   1  'Right Justify
         Caption         =   "CURRENT STATUS:"
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
         Left            =   7800
         TabIndex        =   95
         Tag             =   "XYF"
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label lblNewBal 
         Alignment       =   1  'Right Justify
         Caption         =   "NEW BALANCE:"
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
         Left            =   7800
         TabIndex        =   92
         Tag             =   "XYF"
         Top             =   270
         Width           =   1815
      End
      Begin VB.Label lblDmCm 
         Alignment       =   1  'Right Justify
         Caption         =   "CREDIT"
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
         Left            =   720
         TabIndex        =   28
         Tag             =   "XYF"
         Top             =   630
         Width           =   1215
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "REASON"
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
         Left            =   720
         TabIndex        =   26
         Tag             =   "XYF"
         Top             =   990
         Width           =   1215
      End
      Begin VB.Label lblRefNo 
         Alignment       =   1  'Right Justify
         Caption         =   "REF. NO."
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
         TabIndex        =   23
         Tag             =   "XYF"
         Top             =   270
         Width           =   1695
      End
   End
   Begin MSFlexGridLib.MSFlexGrid flxPRN 
      Height          =   975
      Left            =   3120
      TabIndex        =   20
      Top             =   9840
      Visible         =   0   'False
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   1720
      _Version        =   393216
      Cols            =   5
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   120
      Top             =   1800
   End
   Begin VB.TextBox txtPrnCopy 
      Height          =   285
      Left            =   3240
      TabIndex        =   19
      Top             =   8280
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Timer tmrFloat 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   840
      Top             =   8760
   End
   Begin VB.Timer tmrDrop 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   840
      Top             =   9240
   End
   Begin VB.Timer tmrFloatRev 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1440
      Top             =   8760
   End
   Begin VB.Timer tmrDropRev 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1440
      Top             =   9240
   End
   Begin VB.Frame Frame1 
      Height          =   1005
      Left            =   480
      TabIndex        =   3
      Tag             =   "XYF"
      Top             =   720
      Width           =   12375
      Begin VB.ComboBox cboTrans 
         Height          =   315
         Left            =   2280
         TabIndex        =   21
         Tag             =   "XYF"
         Top             =   600
         Width           =   3615
      End
      Begin VB.TextBox txtAdjType 
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
         Left            =   2280
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   0
         Tag             =   "XYF"
         Top             =   240
         Width           =   3615
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
         Left            =   9720
         TabIndex        =   2
         Tag             =   "XYF"
         Top             =   540
         Width           =   2295
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
         Left            =   9720
         TabIndex        =   1
         Tag             =   "XYF"
         Top             =   195
         Width           =   2295
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
         Left            =   7920
         TabIndex        =   7
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
         Left            =   7920
         TabIndex        =   6
         Tag             =   "XYF"
         Top             =   225
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "TRANSACTION:"
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
         TabIndex        =   5
         Tag             =   "XYF"
         Top             =   615
         Width           =   2055
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
         Left            =   120
         TabIndex        =   4
         Tag             =   "XYF"
         Top             =   270
         Width           =   2055
      End
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   12
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
      Image           =   "DMCMadj.frx":0000
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   2
      Left            =   3000
      TabIndex        =   13
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
      Image           =   "DMCMadj.frx":059A
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   3
      Left            =   4500
      TabIndex        =   14
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
      Image           =   "DMCMadj.frx":0B34
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   4
      Left            =   6000
      TabIndex        =   15
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
      Image           =   "DMCMadj.frx":10CE
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   5
      Left            =   7500
      TabIndex        =   16
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
      Image           =   "DMCMadj.frx":1668
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   6
      Left            =   9000
      TabIndex        =   17
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
      Image           =   "DMCMadj.frx":1C02
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   1
      Left            =   1500
      TabIndex        =   18
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
      Image           =   "DMCMadj.frx":219C
      cBack           =   -2147483633
   End
   Begin TabDlg.SSTab tabDocs 
      Height          =   4455
      Left            =   6840
      TabIndex        =   30
      Tag             =   "XYF"
      Top             =   3360
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   7858
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "INVOICE"
      TabPicture(0)   =   "DMCMadj.frx":2736
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblPartsMan"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblGrp2(12)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblInvRef"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblInvDate"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblInvNo"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblTtl(3)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblGrp2(1)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblGrp2(0)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lblInvCust"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lblGrp2(13)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "lblInvStatus"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "lblInvBal"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "lblGrp2(14)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "lblGrp2(15)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "flxInv"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtTtlInv"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).ControlCount=   16
      TabCaption(1)   =   "JOB ORDER"
      TabPicture(1)   =   "DMCMadj.frx":2752
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblChassis"
      Tab(1).Control(1)=   "lblEngine"
      Tab(1).Control(2)=   "lblModel"
      Tab(1).Control(3)=   "lblType"
      Tab(1).Control(4)=   "lblJobDate"
      Tab(1).Control(5)=   "lblJobNo"
      Tab(1).Control(6)=   "lblTtl(4)"
      Tab(1).Control(7)=   "lblGrp2(7)"
      Tab(1).Control(8)=   "lblGrp2(6)"
      Tab(1).Control(9)=   "lblGrp2(5)"
      Tab(1).Control(10)=   "lblGrp2(4)"
      Tab(1).Control(11)=   "lblGrp2(3)"
      Tab(1).Control(12)=   "lblGrp2(2)"
      Tab(1).Control(13)=   "lblMCdriver"
      Tab(1).Control(14)=   "lblJOCust"
      Tab(1).Control(15)=   "lblGrp2(16)"
      Tab(1).Control(16)=   "lblJOstatus"
      Tab(1).Control(17)=   "lblJOBal"
      Tab(1).Control(18)=   "lblGrp2(17)"
      Tab(1).Control(19)=   "lblGrp2(18)"
      Tab(1).Control(20)=   "flxJob"
      Tab(1).Control(21)=   "txtTtlJob"
      Tab(1).ControlCount=   22
      TabCaption(2)   =   "PCV"
      TabPicture(2)   =   "DMCMadj.frx":276E
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txtTtlPCV"
      Tab(2).Control(1)=   "flxPCV"
      Tab(2).Control(2)=   "lblPCVstatus"
      Tab(2).Control(3)=   "lblPcvBal"
      Tab(2).Control(4)=   "lblGrp2(21)"
      Tab(2).Control(5)=   "lblGrp2(20)"
      Tab(2).Control(6)=   "lblPcvCust"
      Tab(2).Control(7)=   "lblGrp2(19)"
      Tab(2).Control(8)=   "lblGrp2(8)"
      Tab(2).Control(9)=   "lblGrp2(9)"
      Tab(2).Control(10)=   "lblGrp2(10)"
      Tab(2).Control(11)=   "lblGrp2(11)"
      Tab(2).Control(12)=   "lblTtl(5)"
      Tab(2).Control(13)=   "lblPcvNo"
      Tab(2).Control(14)=   "lblPcvDate"
      Tab(2).Control(15)=   "lblPayee"
      Tab(2).Control(16)=   "lblATitle"
      Tab(2).ControlCount=   17
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
         TabIndex        =   33
         Tag             =   "XYF"
         Top             =   3960
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
         Left            =   -70500
         TabIndex        =   32
         Tag             =   "XYF"
         Top             =   3960
         Width           =   1455
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
         TabIndex        =   31
         Tag             =   "XYF"
         Top             =   3960
         Width           =   1455
      End
      Begin MSFlexGridLib.MSFlexGrid flxInv 
         Height          =   2295
         Left            =   60
         TabIndex        =   34
         Tag             =   "XY"
         Top             =   1680
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   4048
         _Version        =   393216
         Cols            =   7
         BackColorBkg    =   -2147483633
         AllowUserResizing=   1
         FormatString    =   $"DMCMadj.frx":278A
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
         Height          =   2055
         Left            =   -74940
         TabIndex        =   35
         Tag             =   "XY"
         Top             =   1920
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   3625
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
         Height          =   2250
         Left            =   -74940
         TabIndex        =   36
         Tag             =   "XY"
         Top             =   1725
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
      Begin VB.Label lblPCVstatus 
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
         TabIndex        =   86
         Tag             =   "XYF"
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label lblPcvBal 
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
         TabIndex        =   85
         Tag             =   "XYF"
         Top             =   1320
         Width           =   2295
      End
      Begin VB.Label lblGrp2 
         Alignment       =   1  'Right Justify
         Caption         =   "STATUS"
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
         Index           =   21
         Left            =   -74880
         TabIndex        =   84
         Tag             =   "XYF"
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label lblGrp2 
         Alignment       =   1  'Right Justify
         Caption         =   "BALANCE"
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
         Index           =   20
         Left            =   -72600
         TabIndex        =   83
         Tag             =   "XYF"
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label lblPcvCust 
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
         TabIndex        =   82
         Tag             =   "XYF"
         Top             =   420
         Width           =   4815
      End
      Begin VB.Label lblGrp2 
         Alignment       =   1  'Right Justify
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
         Index           =   19
         Left            =   -74880
         TabIndex        =   81
         Tag             =   "XYF"
         Top             =   420
         Width           =   855
      End
      Begin VB.Label lblGrp2 
         Alignment       =   1  'Right Justify
         Caption         =   "BALANCE"
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
         Index           =   18
         Left            =   -72240
         TabIndex        =   80
         Tag             =   "XYF"
         Top             =   1590
         Width           =   1095
      End
      Begin VB.Label lblGrp2 
         Alignment       =   1  'Right Justify
         Caption         =   "STATUS"
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
         Index           =   17
         Left            =   -74880
         TabIndex        =   79
         Tag             =   "XYF"
         Top             =   1590
         Width           =   1095
      End
      Begin VB.Label lblJOBal 
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
         TabIndex        =   78
         Tag             =   "XYF"
         Top             =   1590
         Width           =   1935
      End
      Begin VB.Label lblJOstatus 
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
         TabIndex        =   77
         Tag             =   "XYF"
         Top             =   1590
         Width           =   1215
      End
      Begin VB.Label lblGrp2 
         Alignment       =   1  'Right Justify
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
         Index           =   16
         Left            =   -74760
         TabIndex        =   76
         Tag             =   "XYF"
         Top             =   420
         Width           =   1095
      End
      Begin VB.Label lblJOCust 
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
         TabIndex        =   75
         Tag             =   "XYF"
         Top             =   420
         Width           =   4455
      End
      Begin VB.Label lblGrp2 
         Alignment       =   1  'Right Justify
         Caption         =   "BALANCE"
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
         Index           =   15
         Left            =   3360
         TabIndex        =   74
         Tag             =   "XYF"
         Top             =   713
         Width           =   1095
      End
      Begin VB.Label lblGrp2 
         Alignment       =   1  'Right Justify
         Caption         =   "STATUS"
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
         Index           =   14
         Left            =   120
         TabIndex        =   73
         Tag             =   "XYF"
         Top             =   713
         Width           =   1095
      End
      Begin VB.Label lblInvBal 
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
         TabIndex        =   72
         Tag             =   "XYF"
         Top             =   713
         Width           =   1215
      End
      Begin VB.Label lblInvStatus 
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
         TabIndex        =   71
         Tag             =   "XYF"
         Top             =   713
         Width           =   1215
      End
      Begin VB.Label lblMCdriver 
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
         TabIndex        =   70
         Tag             =   "XYF"
         Top             =   3960
         Width           =   3735
      End
      Begin VB.Label lblGrp2 
         Alignment       =   1  'Right Justify
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
         Index           =   13
         Left            =   120
         TabIndex        =   69
         Tag             =   "XYF"
         Top             =   1006
         Width           =   1095
      End
      Begin VB.Label lblInvCust 
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
         TabIndex        =   68
         Tag             =   "XYF"
         Top             =   1006
         Width           =   4455
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
         TabIndex        =   66
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
         TabIndex        =   65
         Tag             =   "XYF"
         Top             =   420
         Width           =   1215
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
         TabIndex        =   64
         Tag             =   "XYF"
         Top             =   780
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
         TabIndex        =   63
         Tag             =   "XYF"
         Top             =   1050
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
         TabIndex        =   62
         Tag             =   "XYF"
         Top             =   1320
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
         TabIndex        =   61
         Tag             =   "XYF"
         Top             =   780
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
         TabIndex        =   60
         Tag             =   "XYF"
         Top             =   1050
         Width           =   1095
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
         TabIndex        =   59
         Tag             =   "XYF"
         Top             =   1320
         Width           =   1095
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
         Left            =   3600
         TabIndex        =   58
         Tag             =   "XYF"
         Top             =   4020
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
         Left            =   -72000
         TabIndex        =   57
         Tag             =   "XYF"
         Top             =   4020
         Width           =   1455
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
         TabIndex        =   56
         Tag             =   "XYF"
         Top             =   780
         Width           =   855
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
         TabIndex        =   55
         Tag             =   "XYF"
         Top             =   1050
         Width           =   855
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
         TabIndex        =   54
         Tag             =   "XYF"
         Top             =   780
         Width           =   1095
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
         Top             =   1050
         Width           =   1095
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
         TabIndex        =   52
         Tag             =   "XYF"
         Top             =   4020
         Width           =   1455
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
         TabIndex        =   51
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
         TabIndex        =   50
         Tag             =   "XYF"
         Top             =   420
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
         TabIndex        =   49
         Tag             =   "XYF"
         Top             =   780
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
         TabIndex        =   48
         Tag             =   "XYF"
         Top             =   1050
         Width           =   1215
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
         TabIndex        =   47
         Tag             =   "XYF"
         Top             =   1320
         Width           =   1215
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
         TabIndex        =   46
         Tag             =   "XYF"
         Top             =   780
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
         TabIndex        =   45
         Tag             =   "XYF"
         Top             =   1050
         Width           =   1935
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
         TabIndex        =   44
         Tag             =   "XYF"
         Top             =   1320
         Width           =   1935
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
         TabIndex        =   43
         Tag             =   "XYF"
         Top             =   780
         Width           =   1215
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
         TabIndex        =   42
         Tag             =   "XYF"
         Top             =   1050
         Width           =   1215
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
         TabIndex        =   41
         Tag             =   "XYF"
         Top             =   780
         Width           =   2295
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
         TabIndex        =   40
         Tag             =   "XYF"
         Top             =   1050
         Width           =   2295
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
         TabIndex        =   39
         Tag             =   "XYF"
         Top             =   1300
         Width           =   4455
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
         TabIndex        =   38
         Tag             =   "XYF"
         Top             =   1300
         Width           =   1095
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
         Left            =   120
         TabIndex        =   37
         Tag             =   "XYF"
         Top             =   3960
         Width           =   3735
      End
   End
   Begin MSFlexGridLib.MSFlexGrid flxLedger 
      Height          =   3375
      Left            =   480
      TabIndex        =   67
      Tag             =   "XYF"
      Top             =   3360
      Width           =   6300
      _ExtentX        =   11113
      _ExtentY        =   5953
      _Version        =   393216
      Cols            =   6
      FixedCols       =   0
      BackColorBkg    =   -2147483633
      ScrollBars      =   2
      SelectionMode   =   1
      FormatString    =   "^ DATE|^REF. NO. |^REF TYPE|>DEBIT |>CREDIT |>BALANCE"
   End
   Begin VB.Label Label2 
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
      Index           =   2
      Left            =   480
      TabIndex        =   90
      Tag             =   "XYF"
      Top             =   7335
      Width           =   1335
   End
   Begin VB.Label Label2 
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
      Index           =   1
      Left            =   480
      TabIndex        =   89
      Tag             =   "XYF"
      Top             =   6975
      Width           =   1335
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   120
      Top             =   4320
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   13335
      m_MDI_Height    =   8505
   End
   Begin VB.Label lblHeading 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DEBIT/CREDIT MEMO"
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
      Left            =   10740
      TabIndex        =   10
      Tag             =   "TRF"
      Top             =   450
      Width           =   1995
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
      Left            =   480
      TabIndex        =   8
      Tag             =   "TLF"
      Top             =   450
      Width           =   2900
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
      Left            =   510
      TabIndex        =   9
      Tag             =   "TLF"
      Top             =   475
      Width           =   2900
   End
   Begin VB.Label lblHeading 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DEBIT/CREDIT MEMO"
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
      Left            =   10770
      TabIndex        =   11
      Tag             =   "TRF"
      Top             =   480
      Width           =   1995
   End
   Begin VB.Menu mnuADJ 
      Caption         =   "SpareParts"
      Visible         =   0   'False
   End
End
Attribute VB_Name = "DMCMadj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsParts As New ADODB.Recordset
Dim rsEmployee As New ADODB.Recordset
Dim ActRow As Integer, ActCol As Integer, BtnIndx As Integer
Dim AddNew As Boolean, LockFlx As Boolean, DispOnly As Boolean, PrtSrch As Boolean, PrtBln As Boolean, CustSrch As Boolean
Public ViewTransType As String

Private Sub ClearEntries()
    txtAdjType.Text = ""
    cboTrans.Text = ""
    txtAdjNo.Text = ""
    txtAdjDate.Text = ""
    txtRefNo.Text = ""
    txtAmount.Text = ""
    txtNewBal.Text = ""
    txtNewStatus.Text = ""
    txtReason.Text = ""
    txtPrepby.Text = ""
    txtAprvby.Text = ""
    flxLedger.Text = ""
    txtCustID.Text = ""
    txtJOid.Text = ""
    txtRefType.Text = ""
    flxLedger.Rows = 1
    
    lblInvNo.Caption = ""
    lblInvStatus.Caption = ""
    lblInvCust.Caption = ""
    lblInvRef.Caption = ""
    flxInv.Rows = 1
    flxInv.Rows = 2
    lblInvDate.Caption = ""
    lblInvBal.Caption = ""
    txtTtlInv.Text = ""
    lblPartsMan.Caption = ""
    
    lblJOCust.Caption = ""
    lblJobNo.Caption = ""
    lblJobDate.Caption = ""
    lblType.Caption = ""
    lblJOstatus.Caption = ""
    lblModel.Caption = ""
    lblEngine.Caption = ""
    lblChassis.Caption = ""
    lblJOBal.Caption = ""
    flxJob.Rows = 1
    flxJob.Rows = 2
    txtTtlJob.Text = ""
    lblMCdriver.Caption = ""
    
    lblPcvCust.Caption = ""
    lblPcvNo.Caption = ""
    lblPcvDate.Caption = ""
    lblPCVstatus.Caption = ""
    lblPayee.Caption = ""
    lblATitle.Caption = ""
    lblPcvBal.Caption = ""
    flxPCV.Rows = 1
    flxPCV.Rows = 2
    txtTtlPCV.Text = ""

    txtPrnCopy.Text = ""
End Sub

Private Sub LockData(bln As Boolean)
    Frame1.Enabled = Not bln
    Frame3.Enabled = Not bln
    txtPrepby.Locked = bln
    txtAprvby.Locked = bln
    tabDocs.Enabled = Not bln
End Sub

Public Sub ComputeTotal()
    Dim GrdTtl As Single, i As Integer
    Dim TtlPrice As Currency
    
    GrdTtl = 0
    For i = 1 To flxAdj.Rows - 1
        TtlPrice = Val(flxAdj.TextMatrix(i, 5)) * PureNum(flxAdj.TextMatrix(i, 6))
        GrdTtl = GrdTtl + TtlPrice
        flxAdj.TextMatrix(i, 7) = Format(TtlPrice, "#,##0.00")
    Next i
    txtTotal.Text = Format$(GrdTtl, "#,##0.00")
End Sub

Private Sub cboTrans_Click()
    Select Case cboTrans.ListIndex
        Case 0  'Invoice
            lblRefNo.Caption = "INVOICE NO."
            tabDocs.Tab = 0
            tabDocs.TabEnabled(0) = True
            tabDocs.TabEnabled(1) = False
            tabDocs.TabEnabled(2) = False
            UpdInv.Visible = True
            UpdInv.Value = 0
        Case 1  'JO
            lblRefNo.Caption = "JOB ORDER NO."
            tabDocs.Tab = 1
            tabDocs.TabEnabled(0) = False
            tabDocs.TabEnabled(1) = True
            tabDocs.TabEnabled(2) = False
            UpdInv.Visible = False
        Case 2  'pcv
            lblRefNo.Caption = "PCV NO."
            tabDocs.TabEnabled(2) = True
            tabDocs.TabEnabled(1) = False
            tabDocs.TabEnabled(0) = False
            tabDocs.Tab = 2
            UpdInv.Visible = False
    End Select
End Sub

Private Sub cboTrans_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtAdjNo.SetFocus
End Sub
Private Sub cmdTrans_Click(Index As Integer)
    Dim Ans As Integer, i As Integer
    Dim myMsg As String, PrnVal(0 To 17) As String, JOrem As String, Qry As String
    Dim flx As MSFlexGrid
    
    Select Case Index
        Case 0  'New
            '***Access
            'Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE (((USERACCESS.EMPID)=" & UserID & ") AND ((USERACCESS.MENUINDX)=51));"
            'mySQL
            Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE USERACCESS.EMPID=" & UserID & " AND USERACCESS.MENUINDX=51;"
            MyAcc = GetFieldValue(Qry)
            If MyAcc <> "FULL ACCESS" Then
                MsgBox "You don't have full access to this transaction." & vbCrLf & "Please contact your System Administrator", vbExclamation + vbOKOnly, "Access Denied"
                Exit Sub
            End If
        
            BtnIndx = 4
            tmrFloat.Enabled = True
            Call ClearEntries
            Call LockData(False)
            txtAdjType.Text = "DMCM [A/R]"
            txtAdjDate.Text = Date
            AddNew = True
            txtAdjNo.Text = GetAdjustmentNo
            txtPrepby.Text = GetSettings("OIC")
            txtAprvby.Text = GetSettings("MANAGER")
            
            'New Code RVP
            cmdTrans(1).Tag = txtAdjNo.Text
            
            cboTrans.SetFocus
        Case 1  'Edit
            '***Access
            'Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE (((USERACCESS.EMPID)=" & UserID & ") AND ((USERACCESS.MENUINDX)=51));"
            'mySQL
            Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE USERACCESS.EMPID=" & UserID & " AND USERACCESS.MENUINDX=51;"
            MyAcc = GetFieldValue(Qry)
            If MyAcc <> "FULL ACCESS" Then
                MsgBox "You don't have full access to this transaction." & vbCrLf & "Please contact your System Administrator", vbExclamation + vbOKOnly, "Access Denied"
                Exit Sub
            End If
            If Val(txtPrnCopy.Text) > 0 Then
                MsgBox "Editing is not allowed because Adjustment was already printed.", vbInformation + vbOKOnly, "Edit Denied"
                Exit Sub
            ElseIf txtAdjNo.Text = "" Then
                MsgBox "No available Adjustments to Edit", vbInformation + vbOKOnly, "Confirm P.O."
                Exit Sub
            End If
            BtnIndx = 4
            tmrFloat.Enabled = True
            AddNew = False
            cmdTrans(1).Tag = txtAdjNo.Text
            Call LockData(False)
        Case 2  'Summary
            ViewTransType = cboTrans.Text
            FrmSummary.mySrc = "DMCM [A/R]"
            FrmSummary.Show
            'DmCmSumm.Show
            Unload Me
            Exit Sub
        Case 3  'Print
            Ans = MsgBox("Do you want to Print DMCM Adjustment?", vbQuestion + vbYesNo, "Print")
            If Ans = vbNo Then Exit Sub
            PrnVal(0) = "TYPE: DMCM [A/R]"
            PrnVal(1) = "DATE: " & txtAdjDate.Text
            PrnVal(2) = "TRANSACTION: " & cboTrans.Text
            If cboTrans.Text = "INVOICE" Then
                PrnVal(3) = "CUSTOMER: " & lblInvCust.Caption
                PrnVal(5) = "DATE: " & lblInvDate.Caption
            ElseIf cboTrans.Text = "JOB ORDER" Then
                PrnVal(3) = "CUSTOMER: " & lblJOCust.Caption
                PrnVal(5) = "DATE: " & lblJobDate.Caption
            End If
            PrnVal(4) = lblRefNo.Caption & " " & txtRefNo.Text
            PrnVal(6) = lblDmCm.Caption & " AMOUNT:"
            PrnVal(7) = txtAmount.Text
            PrnVal(8) = "BALANCE:"
            PrnVal(9) = txtNewBal.Text
            If txtRefType.Text = "CREDIT" Or txtRefType.Text = "ON ACCOUNT" Then
                PrnVal(10) = "CURRENT STATUS:"
                PrnVal(11) = txtNewStatus.Text
            Else
                PrnVal(10) = ""
                PrnVal(11) = ""
            End If
            PrnVal(12) = "REASON:"
            PrnVal(13) = txtReason.Text
            PrnVal(14) = txtPrepby.Text
            PrnVal(15) = txtAprvby.Text
            PrnVal(16) = "OFFICE CLERK"
            PrnVal(17) = "OIC"
            If Val(txtPrnCopy.Text) = 0 Then
                txtPrnCopy.Text = Val(txtPrnCopy.Text) + 1
                If txtRefType.Text = "CREDIT" Or txtRefType.Text = "ON ACCOUNT" Then
                    Call UpdateLedger
                ElseIf txtRefType.Text = "CASH" Then
                    If cboTrans.Text = "INVOICE" Then
                        '***Access
                        'LYRcon.Execute "UPDATE SPINVOICE SET SPINVOICE.AMOUNT = " & Val(Replace(txtNewBal.Text, ",", "")) & ", SPINVOICE.REMARKS = 'DMCM ADJ# " & txtAdjNo.Text & "' " & _
                            "WHERE (((SPINVOICE.SPINVNO)='" & txtRefNo.Text & "'));"
                        'mySQL
                        LYRcon.Execute "UPDATE SPINVOICE SET SPINVOICE.AMOUNT = " & Val(Replace(txtNewBal.Text, ",", "")) & ", SPINVOICE.REMARKS = 'DMCM ADJ# " & txtAdjNo.Text & "' " & _
                            "WHERE SPINVOICE.SPINVNO='" & txtRefNo.Text & "';"
                    Else
                        '***Access
                        'Qry = "SELECT SRVCJOBORDER.REMARKS From SRVCJOBORDER WHERE (((SRVCJOBORDER.JOBORDERNO)='" & txtRefNo.Text & "'));"
                        'mySQL
                        Qry = "SELECT SRVCJOBORDER.REMARKS From SRVCJOBORDER WHERE SRVCJOBORDER.JOBORDERNO='" & txtRefNo.Text & "';"
                        JOrem = GetFieldValue(Qry)
                        '***Access
                        'LYRcon.Execute "UPDATE SRVCJOBORDER SET SRVCJOBORDER.BALANCE = " & Val(Replace(txtNewBal.Text, ",", "")) & ", SRVCJOBORDER.REMARKS = '" & JOrem & " DMCM ADJ# " & txtAdjNo.Text & "' " & _
                                       "WHERE (((SRVCJOBORDER.JOBORDERNO)='" & txtRefNo.Text & "'));"
                        'mySQL
                        LYRcon.Execute "UPDATE SRVCJOBORDER SET SRVCJOBORDER.BALANCE = " & Val(Replace(txtNewBal.Text, ",", "")) & ", SRVCJOBORDER.REMARKS = '" & JOrem & " DMCM ADJ# " & txtAdjNo.Text & "' " & _
                                       "WHERE SRVCJOBORDER.JOBORDERNO='" & txtRefNo.Text & "';"
                    End If
                End If
            Else
                txtPrnCopy.Text = Val(txtPrnCopy.Text) + 1
            End If
            '***Access
            'LYRcon.Execute "UPDATE ADJDMCM SET ADJDMCM.PRINTCOPY = [ADJDMCM].[PRINTCOPY]+1 WHERE (((ADJDMCM.ADJNO)='" & txtAdjNo.Text & "'));"
            'mysql
            LYRcon.Execute "UPDATE ADJDMCM SET ADJDMCM.PRINTCOPY = ADJDMCM.PRINTCOPY+1 WHERE ADJDMCM.ADJNO='" & txtAdjNo.Text & "';"
            'Print process follows
            
            Call PrintMyForm(PrnVal, "DMCM ADJ", flx)
            
        Case 4  'Exit
            'MotoRep.tbMain.Visible = True
            'MotoRep.imgBG.Visible = True
            MotoRep.toolBarButton.Visible = True
            Unload Me
            Exit Sub
        Case 5  'Save
            myMsg = ValidateDmCm
            If myMsg <> "" Then
                MsgBox myMsg, vbCritical + vbOKOnly, "Validation Error"
                Exit Sub
            End If
            Ans = MsgBox("Do you want to Save Adjustment?", vbQuestion + vbYesNo, "Save")
            If Ans = vbNo Then Exit Sub
            Call SaveDmCm
            MsgBox "DMCM Adjustment Successfully Saved!", vbInformation + vbOKOnly, "Save"
                
            BtnIndx = 6
            tmrFloatRev.Enabled = True
            Call LockData(True)
            Call cmdTrans_Click(3)
        Case 6  'Cancel
            Ans = MsgBox("Do you want to Cancel DMCM?", vbQuestion + vbYesNo, "Cancel")
            If Ans = vbNo Then Exit Sub
            BtnIndx = 6
            tmrFloatRev.Enabled = True
            If AddNew Then
                Call ClearEntries
            Else
                'Call DisplayInvAdj(cmdTrans(1).Tag)
            End If
            Call LockData(True)
    End Select
End Sub


Private Sub Form_Load()
    
    Call LockData(True)
    cboTrans.AddItem "INVOICE"
    cboTrans.AddItem "JOB ORDER"
    cboTrans.AddItem "PETTY CASH VOUCHER"
    'cmdA.Top = 15000
    MotoRep.sBar.Panels(6).Text = "DMCM - Accounts Receivables"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    MotoRep.sBar.Panels(6).Text = ""
End Sub

Private Sub lblDmCm_Change()
    Dim MyBal As Currency
    
    If txtRefNo.Text = "" Then Exit Sub
    If cboTrans.Text = "INVOICE" Then MyBal = IIf(txtRefType.Text <> "CASH", Val(Replace(lblInvBal.Caption, ",", "")), Val(Replace(txtTtlInv.Text, ",", "")))
    If lblDmCm.Caption = "CREDIT" Then
        If Val(Replace(txtAmount.Text, ",", "")) > 0 Then
            txtNewBal.Text = Format(MyBal - Val(Replace(txtAmount.Text, ",", "")), "#,##0.00")
        End If
    Else
        If Val(Replace(txtAmount.Text, ",", "")) > 0 Then
            txtNewBal.Text = Format(MyBal + Val(Replace(txtAmount.Text, ",", "")), "#,##0.00")
        End If
    End If
End Sub


Private Sub RXResizer1_ResizeComplete()
    Dim i As Integer
    
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
Private Sub SaveDmCm()
    Dim rsAdjDMCM As New ADODB.Recordset
    Dim Qry As String, eMsg As String, i As Integer, cAdjno As String
    
    On Error GoTo LocalErr
    'Saving Inventory Adjustment
    
    '***Access
    'With rsAdjDMCM
    '    eMsg = "Error Saving DMCM adjustment"
    '    If AddNew Then
    '        .Open "AdjDMCM", LYRcon, adOpenKeyset, adLockOptimistic
    '        .AddNew
        
    '    Else
    '        Qry = "SELECT ADJDMCM.* From ADJDMCM WHERE (((ADJDMCM.ADJNO)='" & cmdTrans(1).Tag & "'));"
    '        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    '        .MoveFirst
         
    '    End If
    '    !AdjNo = txtAdjNo.Text
    '    !AdjDate = txtAdjDate.Text
    '    !TRANSTYPE = cboTrans.Text
    '    !RefNo = txtRefNo.Text
    '    If lblDmCm.Caption = "CREDIT" Then
    '        !Credit = Val(Replace(txtAmount.Text, ",", ""))
    '        !Debit = 0
    '    Else
    '        !Credit = 0
    '        !Debit = Val(Replace(txtAmount.Text, ",", ""))
    '    End If
    '    !BALANCE = Val(Replace(txtNewBal.Text, ",", ""))
    '    !Status = txtNewStatus.Text
    '    !REASON = txtReason.Text
    '    !PrepBy = txtPrepBy.Text
    '    !AprvBy = txtAprvBy.Text
    '    !UPDATEDBY = UserID
    '    !LASTUPDATE = Now
    '    If cboTrans.Text = "INVOICE" Then
    '        !CUSTOMER = lblInvCust.Caption
    '        If txtRefType.Text = "CREDIT" Then !CUSTOMERID = txtCustID.Text
    '    ElseIf cboTrans.Text = "JOB ORDER" Then
    '        !CUSTOMER = lblJOCust.Caption
    '        If txtRefType.Text = "ON ACCOUNT" Then !CUSTOMERID = txtCustID.Text
    '    End If
    '    .Update
    '    .Close
    'End With
    
    'mySQL
    If lblDmCm.Caption = "CREDIT" Then
        AdjCredit = Val(Replace(txtAmount.Text, ",", ""))
        AdjDebit = 0
    Else
        AdjCredit = 0
        AdjDebit = Val(Replace(txtAmount.Text, ",", ""))
    End If
    If AddNew Then
        LYRcon.Execute "Insert Into adjdmcm(ADJNO, ADJDATE, TRANSTYPE, REFNO, CREDIT, " & _
            "DEBIT, BALANCE, `STATUS`, REASON, PREPBY, APRVBY, UPDATEDBY, LASTUPDATE) " & _
            "VALUES('" & txtAdjNo.Text & "', '" & Format(txtAdjDate.Text, "yyyy-mm-dd") & "', " & _
            "'" & cboTrans.Text & "', '" & txtRefNo.Text & "', " & AdjCredit & ", " & _
            "" & AdjDebit & ", " & Val(Replace(txtNewBal.Text, ",", "")) & ", " & _
            "'" & txtNewStatus.Text & "', '" & txtReason.Text & "', '" & txtPrepby.Text & "', " & _
            "'" & txtAprvby.Text & "', " & UserID & ", '" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "');"
            
    Else
        LYRcon.Execute "Update adjdmcm SET ADJNO = '" & txtAdjNo.Text & "', " & _
            "ADJDATE = '" & Format(txtAdjDate.Text, "yyyy-mm-dd") & "', TRANSTYPE = '" & cboTrans.Text & "', " & _
            "REFNO = '" & txtRefNo.Text & "', CREDIT = " & AdjCredit & ", DEBIT = " & AdjDebit & ", " & _
            "BALANCE = " & Val(Replace(txtNewBal.Text, ",", "")) & ", `STATUS` = '" & txtNewStatus.Text & "', " & _
            "REASON = '" & txtReason.Text & "', PREPBY = '" & txtPrepby.Text & "', APRVBY = '" & txtAprvby.Text & "', " & _
            "UPDATEDBY = " & UserID & ", LASTUPDATE = '" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "' " & _
            "Where adjdmcm.ADJNO = '" & cmdTrans(1).Tag & "';"
    End If
    
    If AddNew Then
        cAdjno = txtAdjNo.Text
    Else
        cAdjno = cmdTrans(1).Tag
    End If
    
    If cboTrans.Text = "INVOICE" Then
        LYRcon.Execute "Update adjdmcm SET CUSTOMER = '" & lblInvCust.Caption & "' " & _
            "Where adjdmcm.ADJNO = '" & cAdjno & "';"
                If txtRefType.Text = "CREDIT" Then
                    LYRcon.Execute "UPDATE adjdmcm SET CUSTOMERID = " & txtCustID.Text & " " & _
                        "Where adjdmcm.ADJNO = '" & cAdjno & "';"
                End If
    ElseIf cboTrans.Text = "JOB ORDER" Then
        LYRcon.Execute "Update adjdmcm SET CUSTOMER = '" & lblJOCust.Caption & "' " & _
            "Where adjdmcm.ADJNO = '" & cAdjno & "';"
                If txtRefType.Text = "ON ACCOUNT" Then
                    LYRcon.Execute "UPDATE adjdmcm SET CUSTOMERID = " & txtCustID.Text & " " & _
                        "Where adjdmcm.ADJNO = '" & cAdjno & "';"
                End If
    End If
    'Add Control Number
    If AddNew Then
        eMsg = "Error Saving Adjustment Counter"
        LYRcon.Execute "INSERT INTO ADJCNTR ( ADJNO, ADJTYPE) VALUES ('" & txtAdjNo.Text & "', 'DMCM');"
    Else
        eMsg = "Error Update Adjustment Counter"
        '***Access
        'LYRcon.Execute "UPDATE ADJCNTR SET ADJCNTR.ADJNO = " & Val(txtAdjNo.Text) & _
                       " WHERE (((ADJCNTR.ADJNO)=" & Val(cmdTrans(1).Tag) & ") AND ((ADJCNTR.ADJTYPE)='DMCM'));"
        'mySQL
        LYRcon.Execute "UPDATE ADJCNTR SET ADJCNTR.ADJNO = '" & txtAdjNo.Text & "' WHERE ADJCNTR.ADJNO=" & Val(cmdTrans(1).Tag) & " AND ADJCNTR.ADJTYPE='DMCM';"
    End If
    Exit Sub
LocalErr:
    MsgBox eMsg & vbCrLf & Err.Number & ": " & Err.Description, vbCritical + vbOKOnly, "Error"
End Sub

Function MaxSeqNo(MyRef As String) As Integer
    Dim rsSeq As New ADODB.Recordset
    Dim Qry As String
    
    If cboTrans.Text = "INVOICE" Then
        '***Access
        'Qry = "SELECT Max(LEDGER.SEQNO) AS MaxOfSEQNO From LEDGER WHERE (((LEDGER.SPINVNO)='" & MyRef & "'));"
        'mySQL
        Qry = "SELECT Max(round(LEDGER.SEQNO)) AS MaxOfSEQNO From LEDGER WHERE LEDGER.SPINVNO='" & MyRef & "';"
    Else
        '***Access
        'Qry = "SELECT Max(LEDGER.SEQNO) AS MaxOfSEQNO From Ledger WHERE (((LEDGER.SPINVNO)='" & MyRef & "') AND ((LEDGER.JOBID)>0));"
        'mySQL
        Qry = "SELECT Max(round(LEDGER.SEQNO)) AS MaxOfSEQNO From Ledger WHERE LEDGER.SPINVNO='" & MyRef & "' AND LEDGER.JOBID>0;"
    End If
    With rsSeq
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            MaxSeqNo = IIf(IsNull(.Fields(0)), 0, .Fields(0)) + 1
        Else
            MaxSeqNo = 1
        End If
        .Close
    End With
End Function

Private Sub UpdateLedger()
    Dim rsLedger As New ADODB.Recordset
    Dim i As Integer, myTtl As Currency, OldBal As Currency
    Dim Qry As String, JOrem As String
    'mySQL Variable
    Dim CreditAmnt, DebitAmnt As Currency
    Dim MaxSqVal, JobIDNo As Long
    
    '***Access
    'With rsLedger
    '    .Open "Ledger", LYRcon, adOpenKeyset, adLockOptimistic
    '    .AddNew
    '    !SPINVNO = txtRefNo.Text
    '    !RefType = "DMCM"
    '    !RefNo = txtAdjNo.Text
    '    !RefDate = txtAdjDate.Text
    '    If lblDmCm.Caption = "CREDIT" Then
    '        !Credit = Val(Replace(txtAmount.Text, ",", ""))
    '    Else
    '        !Debit = Val(Replace(txtAmount.Text, ",", ""))
    '    End If
    '    !OSbal = Val(Replace(txtNewBal.Text, ",", ""))
    '    !CUSTOMERID = txtCustID.Text
    '    !SeqNo = MaxSeqNo(txtRefNo.Text)
    '    If cboTrans.Text = "INVOICE" Then
    '        !JOBID = 0
    '    Else
    '        !JOBID = Val(txtJOid.Text)
    '    End If
    '    .Update
    '    .Close
    'End With
    
    'MySQL
    If cboTrans.Text = "INVOICE" Then
        JobIDNo = 0
    Else
        JobIDNo = Val(txtJOid.Text)
    End If
    LYRcon.Execute "Insert Into ledger(SPINVNO, REFTYPE, REFNO, REFDATE, OSBAL, CUSTOMERID, SEQNO, JOBID)" & _
                   "VALUES('" & txtRefNo.Text & "', 'DMCM', '" & txtAdjNo.Text & "', '" & Format(txtAdjDate.Text, "yyyy-mm-dd") & "', " & _
                   "" & Val(Replace(txtNewBal.Text, ",", "")) & ", " & txtCustID.Text & ", " & _
                   "" & MaxSeqNo(txtRefNo.Text) & ", " & JobIDNo & ");"

    
    'Error update all Credit the same amount
    'If lblDmCm.Caption = "CREDIT" Then
    '     LYRcon.Execute "UPDATE Ledger SET CREDIT = " & Val(Replace(txtAmount.Text, ",", "")) & " Where ledger.SPINVNO = '" & txtRefNo.Text & "'"
    'Else
    '     LYRcon.Execute "UPDATE Ledger SET DEBIT = " & Val(Replace(txtAmount.Text, ",", "")) & " Where ledger.SPINVNO = '" & txtRefNo.Text & "'"
    'End If
    
    'Fixed 10/29/2010
    If lblDmCm.Caption = "CREDIT" Then
         LYRcon.Execute "UPDATE Ledger SET CREDIT = " & Val(Replace(txtAmount.Text, ",", "")) & " Where ledger.RefNo = '" & Trim(txtAdjNo.Text) & "';"
    Else
         LYRcon.Execute "UPDATE Ledger SET DEBIT = " & Val(Replace(txtAmount.Text, ",", "")) & " Where ledger.RefNo = '" & Trim(txtAdjNo.Text) & "';"
    End If
    
    'mySQL
    'Update Table OnAccount
    If cboTrans.Text = "INVOICE" Then
        '***Access
        'LYRcon.Execute "UPDATE ONACCTERMS SET ONACCTERMS.BALANCE = " & Val(Replace(txtNewBal.Text, ",", "")) & ", ONACCTERMS.STATUS = '" & txtNewStatus.Text & "' " & _
                       "WHERE (((ONACCTERMS.SPINVNO)='" & txtRefNo.Text & "'));"
        'mySQL
        LYRcon.Execute "UPDATE ONACCTERMS SET ONACCTERMS.BALANCE = " & Val(Replace(txtNewBal.Text, ",", "")) & ", ONACCTERMS.STATUS = '" & txtNewStatus.Text & "' " & _
                       "WHERE ONACCTERMS.SPINVNO='" & txtRefNo.Text & "';"
        '***Access
        'LYRcon.Execute "UPDATE SPINVOICE SET SPINVOICE.REMARKS = 'DMCM ADJ# " & txtAdjNo.Text & "' " & _
                       "WHERE (((SPINVOICE.SPINVNO)='" & txtRefNo.Text & "'));"
        'mySQL
        LYRcon.Execute "UPDATE SPINVOICE SET SPINVOICE.REMARKS = 'DMCM ADJ# " & txtAdjNo.Text & "' " & _
                       "WHERE SPINVOICE.SPINVNO='" & txtRefNo.Text & "';"
        
    Else
        '***Access
        'LYRcon.Execute "UPDATE ONACCTERMS SET ONACCTERMS.BALANCE = " & Val(Replace(txtNewBal.Text, ",", "")) & ", ONACCTERMS.STATUS = '" & txtNewStatus.Text & "' " & _
                       "WHERE (((ONACCTERMS.STATUS)<>'TEMP') AND ((ONACCTERMS.JOBORDERNO)='" & txtRefNo.Text & "') AND ((ONACCTERMS.SPINVNO) Is Null));"
        'Qry = "SELECT SRVCJOBORDER.REMARKS From SRVCJOBORDER WHERE (((SRVCJOBORDER.JOBORDERNO)='" & txtRefNo.Text & "'));"
        'JOrem = GetFieldValue(Qry)
        'LYRcon.Execute "UPDATE SRVCJOBORDER SET SRVCJOBORDER.REMARKS = '" & JOrem & " DMCM ADJ# " & txtAdjNo.Text & "' " & _
                       "WHERE (((SRVCJOBORDER.JOBORDERNO)='" & txtRefNo.Text & "'));"
        'mySQL
        LYRcon.Execute "UPDATE ONACCTERMS SET ONACCTERMS.BALANCE = " & Val(Replace(txtNewBal.Text, ",", "")) & ", ONACCTERMS.STATUS = '" & txtNewStatus.Text & "' " & _
                       "WHERE ONACCTERMS.STATUS<>'TEMP' AND ONACCTERMS.JOBORDERNO='" & txtRefNo.Text & "' AND (ONACCTERMS.SPINVNO) Is Null;"
        Qry = "SELECT SRVCJOBORDER.REMARKS From SRVCJOBORDER WHERE SRVCJOBORDER.JOBORDERNO='" & txtRefNo.Text & "';"
        JOrem = GetFieldValue(Qry)
        LYRcon.Execute "UPDATE SRVCJOBORDER SET SRVCJOBORDER.REMARKS = '" & JOrem & " DMCM ADJ# " & txtAdjNo.Text & "' " & _
                       "WHERE SRVCJOBORDER.JOBORDERNO='" & txtRefNo.Text & "';"
    End If
    If lblDmCm.Caption = "CREDIT" Then
        '***Access
        'LYRcon.Execute "UPDATE CUSTOMER SET CUSTOMER.OSBAL = [CUSTOMER].[OSBAL]-" & Val(Replace(txtAmount.Text, ",", "")) & _
                       " WHERE (((CUSTOMER.CUSTOMERID)=" & txtCustID.Text & "));"
        'mySQL
        LYRcon.Execute "UPDATE CUSTOMER SET CUSTOMER.OSBAL = CUSTOMER.OSBAL-" & Val(Replace(txtAmount.Text, ",", "")) & _
                       " WHERE CUSTOMER.CUSTOMERID=" & txtCustID.Text & ";"
    Else
        '***Access
        'LYRcon.Execute "UPDATE CUSTOMER SET CUSTOMER.OSBAL = [CUSTOMER].[OSBAL]+" & Val(Replace(txtAmount.Text, ",", "")) & _
                       " WHERE (((CUSTOMER.CUSTOMERID)=" & txtCustID.Text & "));"
        'mySQL
        LYRcon.Execute "UPDATE CUSTOMER SET CUSTOMER.OSBAL = CUSTOMER.OSBAL+" & Val(Replace(txtAmount.Text, ",", "")) & _
                       " WHERE CUSTOMER.CUSTOMERID=" & txtCustID.Text & ";"
    End If
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

Function ValidateDmCm() As String
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
    If txtRefType.Text = "CREDIT" Or txtRefType.Text = "ON ACCOUNT" Then
        If Val(txtCustID.Text) = 0 Then vMsg = vMsg & vbCrLf & "Invalid Customer [No ID found]"
    End If
    If txtAprvby.Text = "" Then vMsg = vMsg & vbCrLf & "Incomplete entry on field Approved By"
    If txtPrepby.Text = "" Then vMsg = vMsg & vbCrLf & "Incomplete entry on field Prepared By"
    If Val(Replace(txtAmount.Text, ",", "")) = 0 Then vMsg = vMsg & vbCrLf & "Invalid " & lblDmCm.Caption & " Amount"
    If txtReason.Text = "" Then vMsg = vMsg & vbCrLf & "Incomplete entry on field Reason"
    If cboTrans.Text <> "INVOICE" And cboTrans.Text <> "JOB ORDER" And cboTrans.Text <> "PETTY CASH VOUCHER" Then
        vMsg = vMsg & vbCrLf & "Invalid Transaction"
    Else
        If txtRefNo.Text = "" Then
             vMsg = vMsg & vbCrLf & "Incomplete entry on " & cboTrans.Text & " number"
        Else
            If cboTrans.Text = "INVOICE" Then
                If lblInvNo.Caption = "" Then vMsg = vMsg & vbCrLf & "Invalid Invoice number"
                If Val(txtNewBal.Text) < 0 Then vMsg = vMsg & vbCrLf & "New Balance must not be less than Zero (0)."
            ElseIf cboTrans.Text = "JOB ORDER" Then
                If lblJobNo.Caption = "" Then vMsg = vMsg & vbCrLf & "Invalid Job Order number"
            End If
        End If
    End If
    ValidateDmCm = vMsg
    
End Function

Public Sub ShowRecord(MyRef As String)
    Dim rsDmCm As New ADODB.Recordset
    Dim Qry As String, AdjType As String, MyStat As String
    Dim AdjQ As Integer, MyAmt As Currency, MyBal As Currency
    
    Call ClearEntries
    '***Access
    'Qry = "SELECT ADJDMCM.* From ADJDMCM WHERE ADJDMCM.ADJNO='" & MyRef & "';"
    'mySQL
    Qry = "SELECT ADJDMCM.* From ADJDMCM WHERE ADJDMCM.ADJNO='" & MyRef & "';"
    
    With rsDmCm
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            txtAdjType.Text = "DMCM [A/R]"
            txtAdjNo.Text = !ADJNO
            txtAdjDate.Text = !ADJDATE
            If !TRANSTYPE = "INVOICE" Then
                cboTrans.ListIndex = 0
            ElseIf !TRANSTYPE = "JOB ORDER" Then
                cboTrans.ListIndex = 1
            Else
                cboTrans.ListIndex = 2
            End If
            txtRefNo.Text = !RefNo
            If !CREDIT > 0 Then
                MyAmt = !CREDIT
                AdjType = "CREDIT"
            Else
                MyAmt = !Debit
                AdjType = "DEBIT"
            End If
            MyBal = !Balance
            MyStat = !STATUS
            txtReason.Text = !REASON
            txtPrepby.Text = !PREPBY
            txtAprvby.Text = !APRVBY
            txtCustID.Text = IIf(IsNull(!CUSTOMERID), 0, !CUSTOMERID)
            txtPrnCopy.Text = IIf(IsNull(!PRINTCOPY), 0, !PRINTCOPY)
        End If
        .Close
    End With
    Call txtrefno_LostFocus
    Call LockData(True)
    
    txtAmount.Text = Format(MyAmt, "#,##0.00")
    lblDmCm.Caption = AdjType
    txtNewBal.Text = Format(MyBal, "#,##0.00")
    txtNewStatus.Text = MyStat
    
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
    
    strValid = "1234567890/-"
    If KeyAscii > 26 Then
        If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
    If KeyAscii = 13 Then txtRefNo.SetFocus
End Sub

Private Sub txtAdjNo_GotFocus()
    If txtAdjNo.Text <> "" Then
        txtAdjNo.SelStart = 0
        txtAdjNo.SelLength = Len(txtAdjNo.Text)
    End If
End Sub

Private Sub txtAdjNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtAdjDate.SetFocus
End Sub

Private Sub txtAdjNo_LostFocus()
    Dim Qry As String
    
    'Verify AdjNo if existing
    If txtAdjNo.Text <> "" And AddNew Then
        '***Access
        'Qry = "SELECT ADJCNTR.ADJNO From ADJCNTR WHERE (((ADJCNTR.ADJNO)=" & Val(txtAdjNo.Text) & "));"
        'mySQL
        Qry = "SELECT ADJCNTR.ADJNO From ADJCNTR WHERE ADJCNTR.ADJNO='" & txtAdjNo.Text & "';"
        
        If CtrlNumExist(Qry) Then
            MsgBox "Adjustment No. " & txtAdjNo.Text & " already exist!", vbExclamation + vbOKOnly, "Validation Error"
        End If
    End If
    
End Sub

Private Sub txtAdjType_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cboTrans.SetFocus
End Sub


Private Sub txtAmount_KeyPress(KeyAscii As Integer)
    Dim strValid As String
    
    strValid = "1234567890."
    If KeyAscii > 26 Then
        If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
    If KeyAscii = 13 Then txtReason.SetFocus
End Sub
Private Sub txtAmount_LostFocus()
    Dim Qry As String, MyType As String, MyBal As Currency
    'identify transaction
    If cboTrans.Text = "INVOICE" And txtRefNo.Text <> "" Then
        If lblDmCm.Caption = "CREDIT" Then  'Adj Type
            MyBal = IIf(txtRefType.Text <> "CASH", Val(Replace(lblInvBal.Caption, ",", "")), Val(Replace(txtTtlInv.Text, ",", "")))
            If Val(Replace(txtAmount.Text, ",", "")) > MyBal Then
                MsgBox "Invalid Amount entered!" & vbCrLf & "Credit Amount cannot be greater than Balance", vbExclamation + vbOKOnly, "DMCM - CREDIT"
                Exit Sub
            Else
                txtNewBal.Text = Format(MyBal - Val(Replace(txtAmount.Text, ",", "")), "#,##0.00")
            End If
        Else
            MyBal = IIf(txtRefType.Text <> "CASH", Val(Replace(lblInvBal.Caption, ",", "")), Val(Replace(txtTtlInv.Text, ",", "")))
            txtNewBal.Text = Format(MyBal + Val(Replace(txtAmount.Text, ",", "")), "#,##0.00")
        End If
        If txtRefType.Text = "CREDIT" Then txtNewStatus.Text = IIf(Val(Replace(txtNewBal.Text, ",", "")) > 0, "OPEN", "CLOSED")
    End If
    If cboTrans.Text = "JOB ORDER" And txtRefNo.Text <> "" Then
        MyBal = Val(Replace(lblJOBal.Caption, ",", ""))
        If lblDmCm.Caption = "CREDIT" Then  'Adj Type
            If Val(Replace(txtAmount.Text, ",", "")) > MyBal Then
                MsgBox "Invalid Amount entered!" & vbCrLf & "Credit Amount cannot be greater than Balance", vbExclamation + vbOKOnly, "DMCM - CREDIT"
                Exit Sub
            Else
                txtNewBal.Text = Format(MyBal - Val(Replace(txtAmount.Text, ",", "")), "#,##0.00")
            End If
        Else
            txtNewBal.Text = Format(MyBal + Val(Replace(txtAmount.Text, ",", "")), "#,##0.00")
        End If
        If txtRefType.Text = "ON ACCOUNT" Then txtNewStatus.Text = IIf(Val(Replace(txtNewBal.Text, ",", "")) > 0, "OPEN", "CLOSED")
    End If
    
    txtAmount.Text = Format(txtAmount.Text, "#,##0.00")
End Sub

Private Sub txtaprvby_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        If LockFlx Then
            cmdTrans(0).SetFocus
        Else
            cmdTrans(5).SetFocus
        End If
    End If
End Sub

Private Sub txtPrepBy_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtAprvby.SetFocus
End Sub

Private Sub txtReason_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(Chr(KeyAscii))
    Call ObjKeyTrapping.Validation(txtReason, KeyAscii, ValidSQLChars)
    If KeyAscii = 13 Then txtPrepby.SetFocus
End Sub

Private Sub txtrefno_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtAmount.SetFocus
End Sub
Private Sub txtrefno_LostFocus()
    If txtRefNo.Text = "" Then Exit Sub
    If cboTrans.ListIndex = -1 Or cboTrans.Text = "" Then
        MsgBox "Please select Transaction from the list provided!", vbExclamation + vbOKOnly, "No Transaction Type"
        Exit Sub
    End If
    Select Case cboTrans.ListIndex
        Case 0  '
            Call DisplayInvoice
        Case 1  '
            Call DisplayService
        Case 2  '
            MsgBox "Cannot Process Adjustment for PCV!" & vbCrLf & "Please contact MIS Department", vbExclamation + vbOKOnly
            
    End Select
End Sub
Private Sub DisplayPCV()

End Sub

Private Sub DisplayService()
    Dim rsJob As New ADODB.Recordset
    Dim rsDetails As New ADODB.Recordset
    Dim rsLedger As New ADODB.Recordset
    Dim myDisc As Currency, MyLabor As Currency, myTtl As Currency
    Dim MyJOid As Long, GTtl As Currency
    Dim Qry As String
    Dim MyBal As Currency, MyCreditLedger As Currency, MyDebitLedger As Currency, MyLedger As Currency, MyJOAmnt As Currency
    
    MyBal = 0
    MyCreditLedger = 0
    MyDebitLedger = 0
    MyLedger = 0
    MyJOAmnt = 0
    
    '***Access
    'Qry = "SELECT SRVCJOBORDER.JOBORDERNO, SRVCJOBORDER.DATERECEIVED, SRVCJOBORDER.JOTYPE, SRVCUSTOMER.MODEL, SRVCUSTOMER.ENGINE, SRVCUSTOMER.CHASSIS, SRVCJOBORDER.CUSTOMER, SRVCJOBORDER.JOBORDERID, SRVCJOBORDER.PAYTYPE, SRVCJOBORDER.BALANCE, SRVCUSTOMER.CUSTOMERID " & _
          "FROM SRVCJOBORDER INNER JOIN SRVCUSTOMER ON SRVCJOBORDER.ACCOUNTNO = SRVCUSTOMER.ACCOUNTNO WHERE (((SRVCJOBORDER.JOBORDERNO)='" & txtRefNo.Text & "'));"
    'mySQL
    Qry = "SELECT SRVCJOBORDER.JOBORDERNO, SRVCJOBORDER.DATERECEIVED, SRVCJOBORDER.JOTYPE, SRVCUSTOMER.MODEL, SRVCUSTOMER.ENGINE, SRVCUSTOMER.CHASSIS, SRVCJOBORDER.CUSTOMER, SRVCJOBORDER.JOBORDERID, SRVCJOBORDER.PAYTYPE, SRVCJOBORDER.BALANCE, SRVCUSTOMER.CUSTOMERID " & _
          "FROM SRVCJOBORDER INNER JOIN SRVCUSTOMER ON SRVCJOBORDER.ACCOUNTNO = SRVCUSTOMER.ACCOUNTNO WHERE SRVCJOBORDER.JOBORDERNO='" & txtRefNo.Text & "';"

    txtRefType.Text = ""
    With rsJob
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
'            If Month(!DateReceived) = Month(Now()) And Year(!DateReceived) = Year(Now()) Then
                lblJOCust.Caption = !CUSTOMER
                lblJobNo.Caption = !JOBORDERNO
                lblJobDate.Caption = Format(!DateReceived, "mm/dd/yyyy")
                lblType.Caption = IIf(IsNull(!JOType), "", !JOType)
                lblModel.Caption = !Model
                lblEngine.Caption = !ENGINE
                lblChassis.Caption = !CHASSIS
                MyJOid = !JOBORDERID
                txtJOid.Text = MyJOid
                txtRefType.Text = !PayType
                txtCustID.Text = !CUSTOMERID
                lblJOBal.Caption = Format(!Balance, "#,##0.00")
'            Else
'                If DateDiff("d", !DateReceived, CDate(Now())) > 2 Then
'                    MsgBox "DMCM is not allowed!", vbInformation + vbOKOnly, "Validation"
'                    .Close
'                    Exit Sub
'                Else
'                    lblJOCust.Caption = !CUSTOMER
'                    lblJobNo.Caption = !JOBORDERNO
'                    lblJobDate.Caption = Format(!DateReceived, "mm/dd/yyyy")
'                    lblType.Caption = IIf(IsNull(!JOType), "", !JOType)
'                    lblModel.Caption = !Model
'                    lblEngine.Caption = !ENGINE
'                    lblChassis.Caption = !CHASSIS
'                    MyJOid = !JOBORDERID
'                    txtJOid.Text = MyJOid
'                    txtRefType.Text = !PayType
'                    txtCustID.Text = !CUSTOMERID
'                    lblJOBal.Caption = Format(!Balance, "#,##0.00")
'                End If
'            End If
        End If
        .Close
    End With
    If txtRefType.Text = "" Then
        MsgBox "Job Order Number not found!", vbExclamation + vbOKOnly, "DMCM"
        Exit Sub
    ElseIf txtRefType.Text = "ON ACCOUNT" Then
        lblCrntStat.Visible = True
        txtNewStatus.Visible = True
        '***Access
        'Qry = "SELECT ONACCTERMS.STATUS From ONACCTERMS " & _
              "WHERE (((ONACCTERMS.JOBORDERNO)='" & txtRefNo.Text & "') AND ((ONACCTERMS.SPINVNO) Is Null));"
        'lblJOstatus.Caption = GetFieldValue(Qry)
        'Qry = "SELECT ONACCTERMS.BALANCE From ONACCTERMS " & _
              "WHERE (((ONACCTERMS.JOBORDERNO)='" & txtRefNo.Text & "') AND ((ONACCTERMS.SPINVNO) Is Null) AND ((ONACCTERMS.STATUS)<>'TEMP'));"
        'lblJOBal.Caption = Format(Val(GetFieldValue(Qry)), "#,##0.00")
        'Qry = "SELECT CUSTOMER.CUSTOMERNAME From CUSTOMER WHERE (((CUSTOMER.CUSTOMERID)=" & txtCustID.Text & "));"
        
        'mySQL
        Qry = "SELECT ONACCTERMS.STATUS From ONACCTERMS " & _
              "WHERE ONACCTERMS.JOBORDERNO='" & txtRefNo.Text & "' AND (ONACCTERMS.SPINVNO) Is Null;"
        lblJOstatus.Caption = GetFieldValue(Qry)
        
        Qry = "SELECT ONACCTERMS.ACCTAMOUNT From ONACCTERMS " & _
              "WHERE (((ONACCTERMS.JOBORDERNO)='" & txtRefNo.Text & "') AND ((ONACCTERMS.SPINVNO) Is Null) AND ((ONACCTERMS.STATUS)<>'TEMP'));"
        MyJOAmnt = Val(GetFieldValue(Qry))
        
        Qry = "Select IFNULL(SUM(CREDIT),0) as TTLCREDIT From ledger where spinvno = '" & txtRefNo.Text & "' And CUSTOMERID = " & txtCustID.Text & " And JOBID = " & MyJOid & " group by spinvno"
        MyCreditLedger = Val(GetFieldValue(Qry))
        
        Qry = "Select IFNULL(SUM(DEBIT),0) as TTLDEBIT From ledger where spinvno = '" & txtRefNo.Text & "' And CUSTOMERID = " & txtCustID.Text & " And JOBID  = " & MyJOid & " group by spinvno"
        MyDebitLedger = Val(GetFieldValue(Qry))
        
        MyLedger = MyCreditLedger - MyDebitLedger
        
        If MyLedger > 0 Then
            MyBal = MyJOAmnt - MyLedger
        Else
            MyBal = MyJOAmnt
        End If
        
        lblJOBal.Caption = Format(Val(MyBal), "#,##0.00")
        
        Qry = "SELECT CUSTOMER.CUSTOMERNAME From CUSTOMER WHERE CUSTOMER.CUSTOMERID=" & txtCustID.Text & ";"
        lblMCdriver.Caption = "ON ACCOUNT Charge to: " & GetFieldValue(Qry)
        If Val(Replace(lblJOBal.Caption, ",", "")) = 0 Then 'if Balance=0 only Debit Adj type is allowed
            UpdDmCm.Enabled = False
            lblDmCm.Caption = "DEBIT"
        Else
            UpdDmCm.Enabled = True
        End If
    ElseIf txtRefType.Text = "CASH" Then
        lblCrntStat.Visible = False
        txtNewStatus.Visible = False
        lblJOstatus.Caption = ""
        lblMCdriver.Caption = ""
        UpdDmCm.Enabled = True
        If lblJobDate.Caption = Format(Date, "mm/dd/yyyy") Then
            MsgBox "Job Order " & txtRefNo.Text & " was issued just today," & vbCrLf & "It is recommended to Cancel the invoice", vbExclamation + vbOKOnly, "DMCM"
            Call ClearEntries
            Exit Sub
        End If
    End If
        
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
    
    With rsLedger
        '***Access
        'Qry = "SELECT LEDGER.REFDATE, LEDGER.REFNO, LEDGER.DEBIT, LEDGER.CREDIT, LEDGER.OSBAL From LEDGER " & _
              "WHERE (((LEDGER.SPINVNO)='" & txtRefNo.Text & "') AND ((LEDGER.CUSTOMERID)=" & txtCustID.Text & ") AND ((LEDGER.JOBID)>0)) ORDER BY LEDGER.SEQNO;"
        'mySQL
        Qry = "SELECT LEDGER.REFDATE, LEDGER.REFNO, LEDGER.REFTYPE, LEDGER.DEBIT, LEDGER.CREDIT, LEDGER.OSBAL From LEDGER " & _
              "WHERE (((LEDGER.SPINVNO)='" & txtRefNo.Text & "') AND ((LEDGER.CUSTOMERID)=" & txtCustID.Text & ") AND ((LEDGER.JOBID)>0)) ORDER BY LEDGER.SEQNO;"
        
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            flxLedger.Rows = 1
            Do While Not .EOF
                flxLedger.AddItem !REFDATE & vbTab & !RefNo & vbTab & !RefType & vbTab & IIf(!Debit <> 0, Format$(!Debit, "#,##0.00"), "") & vbTab & IIf(!CREDIT <> 0, Format$(!CREDIT, "#,##0.00"), "") & vbTab & Format$(!OSbal, "#,##0.00")
                .MoveNext
            Loop
        Else
            flxLedger.Rows = 1
            flxLedger.Rows = 2
        End If
        .Close
    End With
End Sub

Private Sub DisplayInvoice()
    Dim rsInv As New ADODB.Recordset
    Dim rsITR As New ADODB.Recordset
    Dim rsLedger As New ADODB.Recordset
    Dim Qry As String, MyStat As String, MyUnit As String
    Dim Ttl As Currency, RetQty As Integer
    Dim MyBal As Currency, MyCreditLedger As Currency, MyDebitLedger As Currency, MyLedger As Currency, MyInvAmnt As Currency
    
    MyBal = 0
    MyCreditLedger = 0
    MyDebitLedger = 0
    MyLedger = 0
    MyInvAmnt = 0
    
    '***Access
    'Qry = "SELECT SPINVOICE.* From SPINVOICE WHERE (((SPINVOICE.SPINVNO)='" & txtRefNo.Text & "'));"
    'mySQL
    If UpdInv.Value = 0 And lblRefNo.Caption = "CREDIT INVOICE" Then
        Qry = "SELECT SPINVOICE.* From SPINVOICE WHERE SPINVOICE.SPINVNO='" & txtRefNo.Text & "' AND spinvoice.INVTYPE = 'CREDIT';"
    ElseIf UpdInv.Value = 1 And lblRefNo.Caption = "CASH INVOICE" Then
        Qry = "SELECT SPINVOICE.* From SPINVOICE WHERE SPINVOICE.SPINVNO='" & txtRefNo.Text & "' AND spinvoice.INVTYPE = 'CASH';"
    End If
    txtRefType.Text = ""
    With rsInv
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
'   =========DMCM validation cannot be adjusted if the month and year are not equal to the current month and date.========
'            If Month(!InvDate) = Month(Now()) And Year(!InvDate) = Year(Now()) Then
'                lblInvCust.Caption = !ChargeTo
'                lblInvNo.Caption = !SPINVNO
'                lblInvDate.Caption = Format(!InvDate, "mm/dd/yyyy")
'                txtCustID.Text = !CUSTOMERID
'                lblPartsMan.Caption = "PARTSMAN: " & !Frontliner
'                txtRefType.Text = !InvType
'            Else
'                If DateDiff("d", !InvDate, CDate(Now())) > 2 Then
'                    MsgBox "DMCM is not allowed!", vbInformation + vbOKOnly, "Validation"
'                    .Close
'                    Exit Sub
'                Else
                    'MsgBox DateDiff("d", !InvDate, CDate(Now())), vbInformation + vbOKOnly, "Validation"
                    lblInvCust.Caption = !chargeTo
                    lblInvNo.Caption = !SPINVNO
                    lblInvDate.Caption = Format(!InvDate, "mm/dd/yyyy")
                    txtCustID.Text = !CUSTOMERID
                    lblPartsMan.Caption = "PARTSMAN: " & !Frontliner
                    txtRefType.Text = !InvType
'                End If
'            End If
        End If
        .Close
    End With
    If txtRefType.Text = "" Then
        MsgBox "Invoice Number not found!", vbExclamation + vbOKOnly, "DMCM"
        Exit Sub
    ElseIf txtRefType.Text = "CREDIT" Then
        lblCrntStat.Visible = True
        txtNewStatus.Visible = True
        lblRefNo.Caption = "CREDIT INVOICE"
        
        '***Access
        'Qry = "SELECT ONACCTERMS.STATUS From ONACCTERMS WHERE (((ONACCTERMS.SPINVNO)='" & txtRefNo.Text & "'));"
        'lblInvStatus.Caption = GetFieldValue(Qry)
        'Qry = "SELECT ONACCTERMS.BALANCE From ONACCTERMS WHERE (((ONACCTERMS.SPINVNO)='" & txtRefNo.Text & "'));"
        
        'mySQL
        Qry = "SELECT ONACCTERMS.STATUS From ONACCTERMS WHERE ONACCTERMS.SPINVNO='" & txtRefNo.Text & "';"
        lblInvStatus.Caption = GetFieldValue(Qry)
        
        Qry = "SELECT ONACCTERMS.BALANCE FROM ONACCTERMS WHERE ONACCTERMS.SPINVNO='" & txtRefNo.Text & "';"
        MyBal = Val(GetFieldValue(Qry))
        
        Qry = "SELECT SPINVOICE.AMOUNT FROM spinvoice WHERE SPINVOICE.SPINVNO='" & txtRefNo.Text & "' AND spinvoice.INVTYPE = 'CREDIT';"
        MyInvAmnt = Val(GetFieldValue(Qry))
        
        If (cboTrans.Text = "INVOICE") Then
            Qry = "Select IFNULL(SUM(CREDIT),0) as TTLCREDIT From ledger where spinvno = '" & txtRefNo.Text & "' And CUSTOMERID = " & txtCustID.Text & " ;"
            MyCreditLedger = Val(GetFieldValue(Qry))
        Else
            Qry = "Select IFNULL(SUM(CREDIT),0) as TTLCREDIT From ledger where spinvno = '" & txtRefNo.Text & "' And CUSTOMERID = " & txtCustID.Text & " group by spinvno"
            MyCreditLedger = Val(GetFieldValue(Qry))
        End If
        Qry = "Select IFNULL(SUM(DEBIT),0) as TTLDEBIT From ledger where spinvno = '" & txtRefNo.Text & "' And CUSTOMERID = " & txtCustID.Text & " group by spinvno"
        MyDebitLedger = Val(GetFieldValue(Qry))
        
        MyLedger = MyCreditLedger - MyDebitLedger
        
        If MyLedger > 0 Then
            MyBal = MyInvAmnt - MyLedger
        End If
        
'
'        Qry = "SELECT ABS((inv.AMOUNT-(acct.TTLCREDIT-acct.TTLDEBIT))) as BALANCE  From spinvoice as inv " & _
'              "Left Join (Select SUM(CREDIT) as TTLCREDIT,SPINVNO, SUM(DEBIT) as TTLDEBIT From ledger where spinvno = '" & txtRefNo.Text & "' group by spinvno) " & _
'                          "as acct ON acct.SPINVNO = inv.SPINVNO " & _
'              "Where inv.spinvno = '" & txtRefNo.Text & "' and inv.invtype='CREDIT';"
'
        lblInvBal.Caption = Format(Val(MyBal), "#,##0.00")
        If Val(Replace(lblInvBal.Caption, ",", "")) = 0 Then 'if Balance=0 only Debit Adj type is allowed
            UpdDmCm.Enabled = False
            lblDmCm.Caption = "DEBIT"
        Else
            UpdDmCm.Enabled = True
        End If
    ElseIf txtRefType.Text = "CASH" Then
        lblCrntStat.Visible = False
        txtNewStatus.Visible = False
        lblRefNo.Caption = "CASH INVOICE"
        lblInvStatus.Caption = ""
        lblInvBal.Caption = ""
        UpdDmCm.Enabled = True
        If lblInvDate.Caption = Format(Date, "mm/dd/yyyy") Then
            MsgBox "Cash Invoice " & txtRefNo.Text & " was issued just today," & vbCrLf & "It is recommended to Cancel the invoice", vbExclamation + vbOKOnly, "DMCM"
            Call ClearEntries
            Exit Sub
        End If
    End If
    '***Access
    'Qry = "SELECT SPSALESORDER.REQID, SPSALESORDER.INVOICENO From SPSALESORDER WHERE (((SPSALESORDER.INVOICENO)='" & txtRefNo.Text & "'));"
    'mySQL
    Qry = "SELECT SPSALESORDER.REQID, SPSALESORDER.INVOICENO From SPSALESORDER WHERE SPSALESORDER.INVOICENO='" & txtRefNo.Text & "';"
    
    If Val(GetFieldValue(Qry)) > 0 Then
        '***Access
        'Qry = "SELECT REQUISITION.REQNO, SPSALESORDER.INVOICENO FROM SPSALESORDER INNER JOIN REQUISITION ON SPSALESORDER.REQID = REQUISITION.REQID WHERE (((SPSALESORDER.INVOICENO)='" & txtRefNo.Text & "'));"
        'MyReq = GetFieldValue(Qry)
        'Qry = "SELECT REQUISITION.JOBNO, SPSALESORDER.INVOICENO FROM SPSALESORDER INNER JOIN REQUISITION ON SPSALESORDER.REQID = REQUISITION.REQID WHERE (((SPSALESORDER.INVOICENO)='" & txtRefNo.Text & "'));"
        
        'mySQL
        Qry = "SELECT REQUISITION.REQNO, SPSALESORDER.INVOICENO FROM SPSALESORDER INNER JOIN REQUISITION ON SPSALESORDER.REQID = REQUISITION.REQID WHERE SPSALESORDER.INVOICENO='" & txtRefNo.Text & "';"
        MyReq = GetFieldValue(Qry)
        Qry = "SELECT REQUISITION.JOBNO, SPSALESORDER.INVOICENO FROM SPSALESORDER INNER JOIN REQUISITION ON SPSALESORDER.REQID = REQUISITION.REQID WHERE SPSALESORDER.INVOICENO='" & txtRefNo.Text & "';"
        MyJo = GetFieldValue(Qry)
        lblInvRef.Caption = "JOB ORDER NO. " & MyJo & ", REQ. NO. " & MyReq
    End If
    
    If UpdInv.Value = 0 And lblRefNo.Caption = "CREDIT INVOICE" Then
        Qry = "SELECT SPITR.`OUT`, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPITR.PRICE, SPITR.PRICE*SPITR.`OUT` AS TTLPRICE, SPITR.SPAREPARTID " & _
              "FROM SPAREPARTS INNER JOIN SPITR ON SPAREPARTS.SPAREPARTID = SPITR.SPAREPARTID WHERE SPITR.REFERENCE='SI#-" & txtRefNo.Text & "' AND SPITR.REMARKS = 'CREDIT';"
    ElseIf UpdInv.Value = 1 And lblRefNo.Caption = "CASH INVOICE" Then
        Qry = "SELECT SPITR.`OUT`, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPITR.PRICE, SPITR.PRICE*SPITR.`OUT` AS TTLPRICE, SPITR.SPAREPARTID " & _
              "FROM SPAREPARTS INNER JOIN SPITR ON SPAREPARTS.SPAREPARTID = SPITR.SPAREPARTID WHERE SPITR.REFERENCE='SI#-" & txtRefNo.Text & "' AND SPITR.REMARKS = 'CASH';"
    End If
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
                RetQty = ReturnedQty(!SparepartID)
                If RetQty > 0 Then
                    flxInv.Col = 1: flxInv.Row = flxInv.Rows - 1: flxInv.ColSel = 7: flxInv.RowSel = flxInv.Rows - 1: flxInv.FillStyle = flexFillRepeat: flxInv.CellBackColor = RGB(250, 250, 180)
                    flxInv.Col = 2: flxInv.FillStyle = flexFillSingle
                End If
                .MoveNext
            Loop
            txtTtlInv.Text = Format(Ttl, "#,##0.00")
        Else
            flxInv.Rows = 1
            flxInv.Rows = 2
        End If
        .Close
    End With
        
    With rsLedger
        '***Access
        'Qry = "SELECT LEDGER.REFDATE, LEDGER.REFNO, LEDGER.DEBIT, LEDGER.CREDIT, LEDGER.OSBAL From LEDGER " & _
              "Where (((LEDGER.SPINVNO) = '" & txtRefNo.Text & "') And ((LEDGER.CUSTOMERID) = " & txtCustID.Text & ")) ORDER BY LEDGER.SEQNO;"
        'mySQL
        Qry = "SELECT LEDGER.REFDATE, LEDGER.REFNO, LEDGER.REFTYPE, LEDGER.DEBIT, LEDGER.CREDIT, LEDGER.OSBAL From LEDGER " & _
              "Where LEDGER.SPINVNO = '" & txtRefNo.Text & "' And LEDGER.CUSTOMERID = " & txtCustID.Text & " And LEDGER.REFTYPE != 'CANCELLED' And (LEDGER.JOBID IS NULL or LEDGER.JOBID = '') ORDER BY LEDGER.SEQNO;"
        
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            flxLedger.Rows = 1
            Do While Not .EOF
                flxLedger.AddItem !REFDATE & vbTab & !RefNo & vbTab & !RefType & vbTab & IIf(!Debit <> 0, Format$(!Debit, "#,##0.00"), "") & vbTab & IIf(!CREDIT <> 0, Format$(!CREDIT, "#,##0.00"), "") & vbTab & Format$(!OSbal, "#,##0.00")
                .MoveNext
            Loop
        Else
            flxLedger.Rows = 1
            flxLedger.Rows = 2
        End If
        .Close
    End With
End Sub

Function ReturnedQty(SPid As Long) As Integer
    Dim rsRetQ As New ADODB.Recordset
    Dim Qry As String
    
    '***Access
    'Qry = "SELECT SPITR.RECEIVE, SPITR.TRANSTYPE, SPITR.SPAREPARTID, SPITR.PACKINGLIST From SPITR " & _
          "WHERE (((SPITR.TRANSTYPE)='SALES RETURN') AND ((SPITR.SPAREPARTID)=" & SPid & ") AND ((SPITR.PACKINGLIST)='" & MyInv & "'));"
    'mySQL
    Qry = "SELECT SPITR.RECEIVE, SPITR.TRANSTYPE, SPITR.SPAREPARTID, SPITR.PACKINGLIST From SPITR " & _
          "WHERE (((SPITR.TRANSTYPE)='SALES RETURN') AND ((SPITR.SPAREPARTID)=" & SPid & ") AND ((SPITR.PACKINGLIST)='" & MyInv & "'));"
    
    With rsRetQ
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            ReturnedQty = !Receive
        Else
            ReturnedQty = 0
        End If
        .Close
    End With
End Function


Private Sub UpdDmCm_DownClick()
    If lblDmCm.Caption = "CREDIT" Then
        lblDmCm.Caption = "DEBIT"
    Else
        lblDmCm.Caption = "CREDIT"
    End If
    Call txtAmount_LostFocus
End Sub

Private Sub UpdDmCm_UpClick()
    Call UpdDmCm_DownClick
End Sub

Private Sub UpdInv_Change()
    If UpdInv.Value = 0 Then
        lblRefNo.Caption = "CREDIT INVOICE"
    ElseIf UpdInv.Value = 1 Then
        lblRefNo.Caption = "CASH INVOICE"
    End If
    If txtRefNo.Text <> "" Then Call DisplayInvoice
End Sub
