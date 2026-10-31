VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form SalesReturn 
   Caption         =   "Spare Parts Sales Return"
   ClientHeight    =   8505
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   13515
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8505
   ScaleWidth      =   13515
   Tag             =   "XYF"
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtCustID 
      Height          =   285
      Left            =   5520
      TabIndex        =   48
      Top             =   9360
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   480
      TabIndex        =   42
      Tag             =   "XYF"
      Top             =   6400
      Width           =   6495
      Begin VB.CheckBox chkReplace 
         Caption         =   "REPLACE ITEM"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2880
         TabIndex        =   49
         Tag             =   "XY"
         Top             =   600
         Visible         =   0   'False
         Width           =   2055
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
         Left            =   5160
         TabIndex        =   46
         Tag             =   "XYF"
         Top             =   195
         Width           =   1215
      End
      Begin VB.OptionButton OptProcess 
         Caption         =   "REPLACE ITEM"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   44
         Tag             =   "XY"
         Top             =   240
         Value           =   -1  'True
         Width           =   2415
      End
      Begin VB.OptionButton OptProcess 
         Caption         =   "REFUND PAYMENT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   43
         Tag             =   "XY"
         Top             =   600
         Width           =   2415
      End
      Begin VB.Label lblRefAmt 
         Alignment       =   2  'Center
         Caption         =   "0.00"
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
         TabIndex        =   47
         Tag             =   "XYF"
         Top             =   540
         Width           =   1215
      End
      Begin VB.Label lblRef 
         Alignment       =   1  'Right Justify
         Caption         =   "NEW INVOICE NO."
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
         Left            =   3360
         TabIndex        =   45
         Tag             =   "XYF"
         Top             =   210
         Width           =   1695
      End
   End
   Begin VB.TextBox txtQtyTip 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Height          =   285
      Left            =   360
      TabIndex        =   38
      Text            =   "Qty. Purchased: 0"
      Top             =   8160
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtInvType 
      Height          =   285
      Left            =   6360
      TabIndex        =   37
      Top             =   9360
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtReturnID 
      Height          =   285
      Left            =   7320
      TabIndex        =   36
      Top             =   9360
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.ComboBox cboReason 
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
      Left            =   2520
      TabIndex        =   6
      Tag             =   "XYF"
      Top             =   7320
      Width           =   4455
   End
   Begin lvButton.lvButtons_H cmdReturnAll 
      Height          =   375
      Left            =   11520
      TabIndex        =   4
      Tag             =   "XYF"
      Top             =   2115
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      Caption         =   "Return All"
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
      Image           =   "SalesReturn.frx":0000
      cBack           =   -2147483633
   End
   Begin VB.TextBox txtPrnCopy 
      Height          =   285
      Left            =   8040
      TabIndex        =   34
      Top             =   9360
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtTotal 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   10320
      Locked          =   -1  'True
      TabIndex        =   32
      Tag             =   "XYF"
      Top             =   6400
      Width           =   2055
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
   Begin VB.TextBox txtFlx 
      Height          =   285
      Left            =   8760
      TabIndex        =   22
      Top             =   9360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtRemarks 
      Height          =   645
      Left            =   2520
      MaxLength       =   250
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Tag             =   "XYF"
      Top             =   7680
      Width           =   4455
   End
   Begin MSFlexGridLib.MSFlexGrid flxParts 
      Height          =   3855
      Left            =   480
      TabIndex        =   5
      Tag             =   "XYF"
      Top             =   2520
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   6800
      _Version        =   393216
      Cols            =   10
      FixedCols       =   0
      RowHeightMin    =   315
      BackColorBkg    =   -2147483633
      ScrollBars      =   2
      FormatString    =   "^    |^QTY.   |^UNIT       |<PART NUMBER   |<DESCRIPTION|>UNIT PRICE |>TOTAL PRICE"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   1290
      Left            =   480
      TabIndex        =   11
      Tag             =   "XYF"
      Top             =   720
      Width           =   12495
      Begin MSComCtl2.UpDown UpDnINV 
         Height          =   255
         Left            =   1560
         TabIndex        =   50
         Tag             =   "xy"
         Top             =   240
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         _Version        =   393216
         Max             =   1
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtPartsman 
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
         Left            =   10560
         TabIndex        =   40
         Tag             =   "XYF"
         Top             =   900
         Width           =   1815
      End
      Begin VB.TextBox txtInvNo 
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
         Left            =   1800
         MaxLength       =   15
         TabIndex        =   0
         Tag             =   "XYF"
         Top             =   215
         Width           =   1695
      End
      Begin VB.TextBox txtInvDate 
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
         Left            =   1800
         MaxLength       =   50
         TabIndex        =   10
         Tag             =   "XYF"
         Top             =   538
         Width           =   1695
      End
      Begin VB.TextBox txtDateRet 
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
         Left            =   10560
         TabIndex        =   2
         Tag             =   "XYF"
         Top             =   538
         Width           =   1815
      End
      Begin VB.TextBox txtRetNo 
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
         Left            =   10560
         TabIndex        =   1
         Tag             =   "XYF"
         Top             =   200
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "PARTSMAN"
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
         Left            =   9000
         TabIndex        =   41
         Tag             =   "XYF"
         Top             =   930
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "DATE RETURNED"
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
         Left            =   9000
         TabIndex        =   16
         Tag             =   "XYF"
         Top             =   570
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "SALES RETURN NO."
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
         Left            =   8760
         TabIndex        =   15
         Tag             =   "XYF"
         Top             =   225
         Width           =   1695
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
         Left            =   360
         TabIndex        =   14
         Tag             =   "XYF"
         Top             =   900
         Width           =   5175
      End
      Begin VB.Label Label1 
         Caption         =   "INVOICE DATE"
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
         TabIndex        =   13
         Tag             =   "XYF"
         Top             =   568
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "CHARGE INV.#"
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
         TabIndex        =   12
         Tag             =   "XYF"
         Top             =   225
         Width           =   1215
      End
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   23
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
      Image           =   "SalesReturn.frx":059A
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   2
      Left            =   3000
      TabIndex        =   24
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
      Image           =   "SalesReturn.frx":0B34
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   3
      Left            =   4500
      TabIndex        =   25
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
      Image           =   "SalesReturn.frx":10CE
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   4
      Left            =   6000
      TabIndex        =   26
      Top             =   0
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   661
      Caption         =   "E&xit Order"
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
      Image           =   "SalesReturn.frx":1668
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   5
      Left            =   7500
      TabIndex        =   27
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
      Image           =   "SalesReturn.frx":1C02
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   6
      Left            =   9000
      TabIndex        =   28
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
      Image           =   "SalesReturn.frx":219C
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   1
      Left            =   1500
      TabIndex        =   29
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
      Image           =   "SalesReturn.frx":2736
      cBack           =   -2147483633
   End
   Begin VB.TextBox txtPrepBy 
      Height          =   285
      Left            =   10200
      TabIndex        =   8
      Tag             =   "XYF"
      Top             =   7560
      Width           =   2775
   End
   Begin VB.TextBox txtAprvBy 
      Height          =   285
      Left            =   10200
      TabIndex        =   9
      Tag             =   "XYF"
      Top             =   7920
      Width           =   2775
   End
   Begin lvButton.lvButtons_H cmdSelect 
      Height          =   375
      Left            =   10080
      TabIndex        =   3
      Tag             =   "XYF"
      Top             =   2115
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      Caption         =   "Select Item"
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
      Image           =   "SalesReturn.frx":2CD0
      cBack           =   -2147483633
   End
   Begin VB.Label lblAddress 
      Height          =   135
      Left            =   600
      TabIndex        =   39
      Top             =   2160
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "REASON FOR RETURN:"
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
      TabIndex        =   35
      Tag             =   "XYF"
      Top             =   7320
      Width           =   2055
   End
   Begin VB.Label lblTotal 
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
      Left            =   9240
      TabIndex        =   33
      Tag             =   "XYF"
      Top             =   6420
      Width           =   975
   End
   Begin VB.Label Label2 
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
      Index           =   2
      Left            =   8640
      TabIndex        =   31
      Tag             =   "XYF"
      Top             =   7935
      Width           =   1335
   End
   Begin VB.Label Label2 
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
      Index           =   1
      Left            =   8640
      TabIndex        =   30
      Tag             =   "XYF"
      Top             =   7560
      Width           =   1335
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   120
      Top             =   4320
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   13755
      m_MDI_Height    =   9075
      OriginalBorder  =   10
      OriginalCaption =   31
   End
   Begin VB.Label lblHeading 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SALES RETURN"
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
      Left            =   11490
      TabIndex        =   20
      Tag             =   "TRF"
      Top             =   450
      Width           =   1485
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
      TabIndex        =   18
      Tag             =   "TLF"
      Top             =   450
      Width           =   2900
   End
   Begin VB.Label lbl2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
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
      Height          =   210
      Left            =   960
      TabIndex        =   17
      Tag             =   "XYF"
      Top             =   7680
      Width           =   1395
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
      TabIndex        =   19
      Tag             =   "TLF"
      Top             =   475
      Width           =   2900
   End
   Begin VB.Label lblHeading 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SALES RETURN"
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
      Left            =   11520
      TabIndex        =   21
      Tag             =   "TRF"
      Top             =   480
      Width           =   1485
   End
   Begin VB.Menu mnuRows 
      Caption         =   "SpareParts"
      Visible         =   0   'False
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete Row"
      End
   End
End
Attribute VB_Name = "SalesReturn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ActRow As Integer, ActCol As Integer, BtnIndx As Integer
Dim AddNew As Boolean, LockFlx As Boolean, DispOnly As Boolean

Private Sub ClearEntries()
    txtInvNo.Text = ""
    txtInvDate.Text = ""
    lblCustomer.Caption = ""
    txtRetNo.Text = ""
    txtDateRet.Text = ""
    txtTotal.Text = ""
    cboReason.Text = ""
    txtRemarks.Text = ""
    txtPrepBy.Text = ""
    txtAprvBy.Text = ""
    txtPrnCopy.Text = ""
    txtCustID.Text = ""
    txtRefNo.Text = ""
    lblRefAmt.Caption = "0.00"
    flxParts.Rows = 1
    flxParts.Rows = 2
End Sub

Private Sub LockData(Bln As Boolean)
    Frame1.Enabled = Not Bln
    Frame2.Enabled = Not Bln
    cboReason.Locked = Bln
    txtRemarks.Locked = Bln
    txtPrepBy.Locked = Bln
    txtAprvBy.Locked = Bln
    LockFlx = Bln
End Sub

Public Sub ComputeTotal()
    Dim Ttl As Single, i As Integer
    
    Ttl = 0
    For i = 1 To flxParts.Rows - 1
        Ttl = Ttl + PureNum(flxParts.TextMatrix(i, 6))
    Next i
    txtTotal.Text = Format$(Ttl, "#,##0.00")
    
End Sub

Private Sub cboReason_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtRemarks.SetFocus
End Sub

Private Sub cmdReturnAll_Click()
    Dim rsITR As New ADODB.Recordset
    Dim Qry As String
    Dim MyQty As Integer
    
    '***Access
    'Qry = "SELECT SPITR.OUT, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPITR.PRICE, [PRICE]*[OUT] AS TTLAMT, SPITR.SPAREPARTID, SPITR.REFERENCE " & _
          "FROM SPITR INNER JOIN SPAREPARTS ON SPITR.SPAREPARTID = SPAREPARTS.SPAREPARTID WHERE (((SPITR.REFERENCE)='SI#-" & txtInvNo.Text & "'));"
    'mySQL
    Qry = "SELECT SPITR.OUT, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPITR.PRICE, SPITR.PRICE*SPITR.OUT AS TTLAMT, SPITR.SPAREPARTID, SPITR.REFERENCE " & _
          "FROM SPITR INNER JOIN SPAREPARTS ON SPITR.SPAREPARTID = SPAREPARTS.SPAREPARTID WHERE SPITR.REFERENCE='SI#-" & txtInvNo.Text & "' AND SPITR.REMARKS = '" & IIf(UpDnINV.Value = 1, "CASH", "CREDIT") & "';"
    
    With rsITR
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 1 Then
            .MoveFirst
            flxParts.Rows = 1
            Do While Not .EOF
                If InStr(!Description, "ENGINE OIL") > 0 Or InStr(!Description, "4T OIL") > 0 Then
                    MyUnit = IIf(!OUT > 1, "QRTS.", "QRT.")
                Else
                    MyUnit = IIf(!OUT > 1, "PCS.", "PC.")
                End If
                MyQty = !OUT - ReturnedQty(!SPAREPARTID)
                If MyQty > 0 Then
                    flxParts.AddItem vbTab & MyQty & vbTab & MyUnit & vbTab & !PartNumber & vbTab & !Description & vbTab & Format(!Price, "###,##0.00") & vbTab & Format(!TtlAmt, "###,##0.00") & vbTab & !SPAREPARTID & vbTab & MyQty
                End If
                .MoveNext
            Loop
        End If
        .Close
    End With
    If flxParts.Rows = 1 Then
        flxParts.Rows = 2
        MsgBox "No available items for return!", vbExclamation + vbOKOnly, "Validation Error"
    Else
        Call ComputeTotal
    End If
End Sub

Private Sub cmdSelect_Click()
    SalesRetInv.Show vbModal
End Sub

Private Sub cmdTrans_Click(Index As Integer)
    Dim Ans As Integer, i As Integer
    Dim myMsg As String, PrnVal(0 To 8) As String, Qry As String, MyAcc As String
    
    Select Case Index
        Case 0  'New
            '***Access
            'Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE (((USERACCESS.EMPID)=" & UserID & ") AND ((USERACCESS.MENUINDX)=27));"
            'mySQl
            Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE USERACCESS.EMPID=" & UserID & " AND USERACCESS.MENUINDX=27;"
            
            MyAcc = GetFieldValue(Qry)
            If MyAcc <> "FULL ACCESS" Then
                MsgBox "You don't have full access to this transaction." & vbCrLf & "Please contact your System Administrator", vbExclamation + vbOKOnly, "Access Denied"
                Exit Sub
            End If
        
            BtnIndx = 4
            tmrFloat.Enabled = True
            Call ClearEntries
            Call LockData(False)
            UpDnINV.Value = 0
            txtDateRet.Text = Date
            AddNew = True
            txtRetNo.Text = GetSalesRetNo
            txtPrepBy.Text = GetUserName
            txtInvNo.SetFocus
        Case 1  'Edit
            '***Access
            'Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE (((USERACCESS.EMPID)=" & UserID & ") AND ((USERACCESS.MENUINDX)=27));"
            'mySQl
            Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE USERACCESS.EMPID=" & UserID & " AND USERACCESS.MENUINDX=27;"
            
            MyAcc = GetFieldValue(Qry)
            If MyAcc <> "FULL ACCESS" Then
                MsgBox "You don't have full access to this transaction." & vbCrLf & "Please contact your System Administrator", vbExclamation + vbOKOnly, "Access Denied"
                Exit Sub
            End If
            If Val(txtPrnCopy.Text) > 0 Then
                MsgBox "Editing is not allowed because P.O. was already printed.", vbInformation + vbOKOnly, "Edit Denied"
                Exit Sub
            ElseIf txtReturnID.Text = "" Then
                MsgBox "No available Sales Return to Edit", vbInformation + vbOKOnly, "Confirm P.O."
                Exit Sub
            End If
            AddNew = False
            BtnIndx = 4
            tmrFloat.Enabled = True
            AddNew = False
            Call LockData(False)
        Case 2  'Summary
            FrmSummary.mySrc = "Sales Return"
            FrmSummary.Show
            'SalesRetSum.Show
            Unload Me
            Exit Sub
        Case 3  'Print
            
            Ans = MsgBox("Do you want to Print Sales Return?", vbQuestion + vbYesNo, "Print")
            If Ans = vbNo Then Exit Sub
            PrnVal(0) = lblCustomer.Caption
            PrnVal(1) = txtInvNo.Text
            PrnVal(2) = txtDateRet.Text
            PrnVal(3) = lblAddress.Caption
            PrnVal(4) = txtInvDate.Text
            PrnVal(5) = txtPartsman.Text
            PrnVal(6) = "TOTAL:"
            PrnVal(7) = txtTotal.Text
            PrnVal(8) = IIf(txtRemarks.Text <> "", "NOTE: ", "") & txtRemarks.Text
            Call PrintMyForm(PrnVal, "SALES RETURN", flxParts)
            If Val(txtPrnCopy.Text) = 0 Then
                txtPrnCopy.Text = Val(txtPrnCopy.Text) + 1
                Call UpdateSPITR
            Else
                txtPrnCopy.Text = Val(txtPrnCopy.Text) + 1
            End If
            '***Access
            'LYRcon.Execute "UPDATE SRETURNSP SET SRETURNSP.PRINTCOPY = [SRETURNSP].[PRINTCOPY]+1 WHERE (((SRETURNSP.RETURNID)=" & txtReturnID.Text & "));"
            
            'mySQL
            LYRcon.Execute "UPDATE SRETURNSP SET SRETURNSP.PRINTCOPY = SRETURNSP.PRINTCOPY+1 WHERE SRETURNSP.RETURNID=" & txtReturnID.Text & ";"
            
            'Print process follows
            
        Case 4  'Exit
            MotoRep.tbMain.Visible = True
            Unload Me
            Exit Sub
        Case 5  'Save
            myMsg = ValidateSalesReturn
            If myMsg <> "" Then
                MsgBox myMsg, vbCritical + vbOKOnly, "Validation Error"
                Exit Sub
            End If
            Ans = MsgBox("Do you want to Save Sales Return?", vbQuestion + vbYesNo, "Save")
            If Ans = vbNo Then Exit Sub
                Call SaveSalesReturn
                MsgBox "Sales Return Successfully Saved!", vbInformation + vbOKOnly, "Save"
                
            BtnIndx = 6
            tmrFloatRev.Enabled = True
            Call LockData(True)
            
        Case 6  'Cancel
            Ans = MsgBox("Do you want to Cancel Sales Return?", vbQuestion + vbYesNo, "Cancel")
            If Ans = vbNo Then Exit Sub
            BtnIndx = 6
            tmrFloatRev.Enabled = True
            If AddNew Then
                Call ClearEntries
            Else
                Call DisplaySalesReturn(txtReturnID.Text)
            End If
            Call LockData(True)
    End Select
    cmdSelect.Enabled = False
    cmdReturnAll.Enabled = False

End Sub

Private Sub flxParts_Click()
    ActRow = flxParts.Row
    ActCol = flxParts.Col
    
    'FlxParts Columns: 1-PartNumber, 2-Desc, 3-Qty, 4-Price, 5-Total, 6-ID, 7-Remarks
    If ActCol <> 1 Or ActRow = 0 Or LockFlx Then Exit Sub
    
    With txtFlx
        .Top = flxParts.CellTop + flxParts.Top
        .Left = flxParts.CellLeft + flxParts.Left
        .Width = flxParts.CellWidth
        .Text = flxParts.Text
        .Visible = True
        .SelStart = 0
        .SelLength = Len(.Text)
        .ZOrder
        .SetFocus
    End With
    Call QtyTip(True)
End Sub

Private Sub QtyTip(Disp As Boolean)
    If Disp Then
        With txtQtyTip
            .Top = flxParts.CellTop + flxParts.Top + 290
            .Left = flxParts.Left + flxParts.ColWidth(0) + 300
            .Text = "Qty. Purchased: " & flxParts.TextMatrix(ActRow, 8)
            .Visible = True
            .ZOrder
        End With
    Else
        txtQtyTip.Visible = False
    End If
End Sub

Private Sub flxParts_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call flxParts_Click
End Sub

Private Sub flxParts_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If LockFlx Then Exit Sub
    If Button = vbRightButton Then
        If flxParts.TextMatrix(ActRow, 7) = "" Then Exit Sub
        Call PopupMenu(mnuRows)
    End If
End Sub


Private Sub Form_Load()
    Dim i As Integer
    
    Call LockData(True)
    cmdSelect.Enabled = False
    cmdReturnAll.Enabled = False
    
    MotoRep.sBar.Panels(6).Text = "Spare Parts Sales Return Form"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    MotoRep.sBar.Panels(6).Text = ""
End Sub


Private Sub mnuDelete_Click()
    Dim Ans As Integer, i As Integer
    
    Ans = MsgBox("Do you want to Delete Row?", vbQuestion + vbYesNo, "Delete")
    If Ans = vbYes Then
        If ActRow = 1 And flxParts.Rows = 2 Then
            flxParts.Rows = 1
            flxParts.Rows = 2
        Else
            flxParts.RemoveItem ActRow
        End If
        Call ComputeTotal
    End If
End Sub

Private Sub OptProcess_Click(Index As Integer)
    If Index = 0 Then
        lblRef.Caption = "NEW INVOICE NO."
    Else
        lblRef.Caption = "PCV NO."
    End If
    txtRefNo.Text = ""
    lblRefAmt.Caption = "0.00"
End Sub

Private Sub RXResizer1_ResizeComplete()
    With flxParts
        .ColWidth(0) = 350
        .ColWidth(1) = (.Width - 380) * 0.1
        .ColWidth(2) = (.Width - 380) * 0.1
        .ColWidth(3) = (.Width - 380) * 0.25
        .ColWidth(4) = (.Width - 380) * 0.25
        .ColWidth(5) = (.Width - 380) * 0.14
        .ColWidth(6) = (.Width - 380) * 0.14
        .ColWidth(7) = 0
        .ColWidth(8) = 0
        .ColWidth(9) = 0
        txtTotal.Left = .Left + .ColWidth(0) + .ColWidth(1) + .ColWidth(2) + .ColWidth(3) + .ColWidth(4) + .ColWidth(5)
    End With
    txtTotal.Top = flxParts.Height + flxParts.Top
    txtTotal.Left = flxParts.Left + (flxParts.Width - (flxParts.ColWidth(6) + 400))
    txtTotal.Width = flxParts.ColWidth(5)
    lblTotal.Top = txtTotal.Top
    lblTotal.Left = txtTotal.Left - lblTotal.Width 'flxParts.Left + (flxParts.Width - (flxParts.ColWidth(4) + flxParts.ColWidth(5)))
    lbl2.Top = txtRemarks.Top
    lbl2.Left = txtRemarks.Left - lbl2.Width
    lbl2.FontSize = txtRemarks.FontSize

End Sub
Private Sub SaveSalesReturn()
    Dim rsReturn As New ADODB.Recordset
    Dim rsReturnDtls As New ADODB.Recordset
    Dim Qry As String, i As Integer
    
    'Saving Sales Return
    '***Access
    'With rsReturn
    '    If AddNew Then
    '        .Open "SReturnSP", LYRcon, adOpenKeyset, adLockOptimistic
    '        .AddNew
    '    Else
            '***Access
            'Qry = "SELECT SRETURNSP.* From SRETURNSP WHERE (((SRETURNSP.RETURNID)=" & txtReturnID.Text & "));"
            
            'mySQL
    '        Qry = "SELECT SRETURNSP.* From SRETURNSP WHERE SRETURNSP.RETURNID=" & txtReturnID.Text & ";"
            
    '        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    '        .MoveFirst
    '    End If
    '    !RetNo = txtRetNo.Text
    '    !RetDate = txtDateRet.Text
    '    !REASON = cboReason.Text
    '    !Remarks = txtRemarks.Text
    '    !Invoice = txtInvNo.Text
    '    !TotalAmt = Val(Replace(txtTotal.Text, ",", ""))
    '    If txtInvType.Text = "CASH" Then
    '        !PROCESS = IIf(OptProcess(0).Value = True, "REPLACE", "REFUND")
    '    Else
    '        !PROCESS = IIf(chkReplace.Value = 1, "REPLACE", "RETURN")
    '    End If
    '    !ProRef = txtRefNo.Text
    '    !PreparedBy = txtPrepBy.Text
    '    !ApprovedBy = txtAprvBy.Text
    '    !UPDATEDBY = UserID
    '    !LASTUPDATE = Now
    '    !PrintCopy = 0
    '    .Update
    '    If AddNew Then txtReturnID.Text = !RETURNID
    '    .Close
        
   'End With
        'mySQL Statement
        If AddNew Then
            If txtInvType.Text = "CASH" Then
                MyOptnProc = IIf(OptProcess(0).Value = True, "REPLACE", "REFUND")
            Else
                MyOptnProc = IIf(chkReplace.Value = 1, "REPLACE", "RETURN")
            End If
            
            LYRcon.Execute "INSERT INTO sreturnsp(RETNO, RETDATE, REASON, REMARKS, INVOICE, TOTALAMT, `PROCESS`, " & _
                "PROREF, PREPAREDBY, APPROVEDBY, UPDATEDBY, LASTUPDATE, PRINTCOPY)" & _
                "VALUES('" & txtRetNo.Text & "', '" & Format(txtDateRet.Text, "yyyy-mm-dd") & "', '" & cboReason.Text & "', '" & txtRemarks.Text & "', " & _
                "'" & txtInvNo.Text & "', " & Val(Replace(txtTotal.Text, ",", "")) & ", '" & MyOptnProc & "', " & _
                "'" & txtRefNo.Text & "', '" & txtPrepBy.Text & "', '" & txtAprvBy.Text & "', " & UserID & ", " & _
                "'" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "', 0);"
            
            Qry = "SELECT sreturnsp.RETURNID, sreturnsp.Invoice From sreturnsp Where sreturnsp.RETNO = '" & txtRetNo.Text & "' AND sreturnsp.INVOICE = '" & txtInvNo.Text & "';"
            txtReturnID.Text = GetFieldValue(Qry)
        Else
            LYRcon.Execute "Insert Into sreturnsp( RETNO, RETDATE, REASON, REMARKS, INVOICE, TOTALAMT, `PROCESS`, PROREF, " & _
                "PREPAREDBY, APPROVEDBY, UPDATEDBY, LASTUPDATE, PRINTCOPY)VALUES('" & txtRetNo.Text & "', '" & Format(txtDateRet.Text, "yyyy-mm-dd") & "', " & _
                "'" & cboReason.Text & "', '" & txtRemarks.Text & "', '" & txtInvNo.Text & "', " & Val(Replace(txtTotal.Text, ",", "")) & ", '" & MyOptnProc & "', " & _
                "'" & txtRefNo.Text & "', '" & txtPrepBy.Text & "', '" & txtAprvBy.Text & "', " & UserID & ", '" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "', 0);"
        End If
    'Sales Return Details
    If Not AddNew Then
        '***Access
        'LYRcon.Execute "DELETE SRETURNDTLS.* From SRETURNDTLS WHERE (((SRETURNDTLS.RETURNID)=" & txtReturnID.Text & "));"
        'mySQl
        LYRcon.Execute "DELETE From SRETURNDTLS WHERE SRETURNDTLS.RETURNID=" & txtReturnID.Text & ";"
    
    End If
    
    With flxParts
        For i = 1 To .Rows - 1
            LYRcon.Execute "INSERT INTO SRETURNDTLS ( RETURNID, SPAREPARTID, UNIT, QTY, PRICE, QTYPURCHASED ) " & _
                            "VALUES (" & txtReturnID.Text & ", " & .TextMatrix(i, 7) & ", '" & .TextMatrix(i, 2) & "', " & .TextMatrix(i, 1) & ",  " & Val(Replace(.TextMatrix(i, 5), ",", "")) & ", " & .TextMatrix(i, 8) & ");"
        Next i
    End With
End Sub

Private Sub UpdateSPITR()
    Dim rsITR As New ADODB.Recordset
    Dim rsLedger As New ADODB.Recordset
    Dim i As Integer, myTtl As Currency, OldBal As Currency
    Dim Qry As String, MyStat As String, myRem As String
    
    With flxParts
        For i = 1 To .Rows - 1
            'Update SpareParts Qty on hand
            '***Access
            'LYRcon.Execute "UPDATE SPAREPARTS SET SPAREPARTS.QTYONHAND = [SPAREPARTS].[QTYONHAND]+" & Val(.TextMatrix(i, 1)) & _
                           " WHERE (((SPAREPARTS.SPAREPARTID)=" & .TextMatrix(i, 7) & "));"
            'mySQL
            LYRcon.Execute "UPDATE SPAREPARTS SET SPAREPARTS.QTYONHAND = SPAREPARTS.QTYONHAND+" & Val(.TextMatrix(i, 1)) & _
                           " WHERE SPAREPARTS.SPAREPARTID=" & .TextMatrix(i, 7) & ";"
            
            'Saving Inventory Transaction Register
            '***Access
            'rsITR.Open "SPITR", LYRcon, adOpenKeyset, adLockOptimistic
            'rsITR.AddNew
            'rsITR!REFERENCE = "RET#" & txtRetNo.Text
            'rsITR!SparepartID = .TextMatrix(i, 7)
            'rsITR!RECEIVE = Val(.TextMatrix(i, 1))
            'rsITR!TRANSDATE = txtDateRet.Text
            'rsITR!TRANSTYPE = "SALES RETURN"
            'rsITR!Price = PureNum(.TextMatrix(i, 5))
            'rsITR!PACKINGLIST = txtInvNo.Text
            'rsITR!Source = lblCustomer.Caption
            'rsITR!Remarks = txtInvType.Text
            'rsITR!UPDATEDBY = UserID
            'rsITR!LASTUPDATE = Now
            'rsITR.Update
            'rsITR.Close
            
            'mySQL
            LYRcon.Execute "Insert spitr(REFERENCE, SPAREPARTID, RECEIVE, TRANSDATE, TRANSTYPE," & _
                "PRICE, PACKINGLIST, SOURCE, REMARKS, UPDATEDBY, LASTUPDATE, COST)" & _
                "VALUES('RET#" & txtRetNo.Text & "', " & .TextMatrix(i, 7) & ", " & Val(.TextMatrix(i, 1)) & ", " & _
                "'" & Format(txtDateRet.Text, "yyyy-mm-dd") & "', 'SALES RETURN', " & _
                "" & Val(Replace(.TextMatrix(i, 5), ",", "")) & ", '" & txtInvNo.Text & "', '" & lblCustomer.Caption & "', " & _
                "'" & txtInvType.Text & "', " & UserID & ", '" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "', " & Val(Format(.TextMatrix(i, 9), "####0.00")) & ");"
            
            
        Next i
    End With
    
    'Update Invoice
    If (txtInvType.Text = "CASH" And OptProcess(0).Value = True) Or txtInvType.Text = "CREDIT" And chkReplace.Value = 1 Then
        '***Access
        'LYRcon.Execute "UPDATE SPINVOICE SET SPINVOICE.REMARKS = 'REPLACE W/ INV.# " & txtRefNo.Text & "' " & _
                       "WHERE (((SPINVOICE.SPINVNO)='" & txtInvNo.Text & "'));"
        'mySQL
        LYRcon.Execute "UPDATE SPINVOICE SET SPINVOICE.REMARKS = 'REPLACE W/ INV.# " & txtRefNo.Text & "' " & _
                       "WHERE SPINVOICE.SPINVNO='" & txtInvNo.Text & "';"
        '***Access
        'LYRcon.Execute "UPDATE SPINVOICE SET SPINVOICE.REMARKS = 'SALES RET# " & txtRetNo & " REPLACEMENT' " & _
                       "WHERE (((SPINVOICE.SPINVNO)='" & txtRefNo.Text & "'));"
        'mySQL
        LYRcon.Execute "UPDATE SPINVOICE SET SPINVOICE.REMARKS = 'SALES RET# " & txtRetNo & " REPLACEMENT' " & _
                       "WHERE SPINVOICE.SPINVNO='" & txtRefNo.Text & "';"
        
    ElseIf txtInvType.Text = "CASH" And OptProcess(1).Value = True Then    'REFUND ITEM
        '***Access
        'LYRcon.Execute "UPDATE SPINVOICE SET SPINVOICE.REMARKS = 'SALES RET# " & txtRetNo & " REFUNDED PCV# " & txtRefNo.Text & "' " & _
                       "WHERE (((SPINVOICE.SPINVNO)='" & txtRefNo.Text & "'));"
        'mySQL
        LYRcon.Execute "UPDATE SPINVOICE SET SPINVOICE.REMARKS = 'SALES RET# " & txtRetNo & " REFUNDED PCV# " & txtRefNo.Text & "' " & _
                       "WHERE SPINVOICE.SPINVNO='" & txtRefNo.Text & "';"
        '***Access
        'Qry = "SELECT PAYMENTS.REMARKS From PAYMENTS WHERE (((PAYMENTS.INVOICENO)='" & txtInvNo.Text & "'));"
        'mySQL
        Qry = "SELECT PAYMENTS.REMARKS From PAYMENTS WHERE PAYMENTS.INVOICENO='" & txtInvNo.Text & "';"
        
        myRem = GetFieldValue(Qry)
        '***Access
        'LYRcon.Execute "UPDATE PAYMENTS SET PAYMENTS.REMARKS = '" & IIf(myRem <> "", myRem & " ", "") & "REFUNDED W/ PCV# " & txtRefNo.Text & "' " & _
                       "WHERE (((PAYMENTS.INVOICENO)='" & txtInvNo.Text & "'));"
        'mySQL
        LYRcon.Execute "UPDATE PAYMENTS SET PAYMENTS.REMARKS = '" & IIf(myRem <> "", myRem & " ", "") & "REFUNDED W/ PCV# " & txtRefNo.Text & "' " & _
                       "WHERE PAYMENTS.INVOICENO='" & txtInvNo.Text & "';"
    
    Else
        '***Access
        'LYRcon.Execute "UPDATE SPINVOICE SET SPINVOICE.REMARKS = 'SALES RET# " & txtRetNo & " RETURNED' " & _
                       "WHERE (((SPINVOICE.SPINVNO)='" & txtInvNo.Text & "'));"
        'mySQL
        LYRcon.Execute "UPDATE SPINVOICE SET SPINVOICE.REMARKS = 'SALES RET# " & txtRetNo & " RETURNED' " & _
                       "WHERE SPINVOICE.SPINVNO='" & txtInvNo.Text & "';"
        
    End If
    If txtInvType.Text = "CASH" Then
        '***Access
        'Qry = "SELECT PAYMENTS.REMARKS From PAYMENTS WHERE (((PAYMENTS.INVOICENO)='" & txtRefNo.Text & "'));"
        'myRem = GetFieldValue(Qry)
        'LYRcon.Execute "UPDATE PAYMENTS SET PAYMENTS.REMARKS = '" & IIf(myRem <> "", myRem & " ", "") & "SALES RET# " & txtRetNo.Text & "' " & _
                       "WHERE (((PAYMENTS.INVOICENO)='" & txtRefNo.Text & "'));"
        'Qry = "SELECT PAYMENTS.REMARKS From PAYMENTS WHERE (((PAYMENTS.INVOICENO)='" & txtInvNo.Text & "'));"
        'myRem = GetFieldValue(Qry)
        'LYRcon.Execute "UPDATE PAYMENTS SET PAYMENTS.REMARKS = '" & IIf(myRem <> "", myRem & " ", "") & "REPLACE W/ INV.# " & txtRefNo.Text & "' " & _
                       "WHERE (((PAYMENTS.INVOICENO)='" & txtInvNo.Text & "'));"
    
        'mySQL
        Qry = "SELECT PAYMENTS.REMARKS From PAYMENTS WHERE PAYMENTS.INVOICENO='" & txtRefNo.Text & "';"
        myRem = GetFieldValue(Qry)
        LYRcon.Execute "UPDATE PAYMENTS SET PAYMENTS.REMARKS = '" & IIf(myRem <> "", myRem & " ", "") & "SALES RET# " & txtRetNo.Text & "' " & _
                       "WHERE PAYMENTS.INVOICENO='" & txtRefNo.Text & "';"
        Qry = "SELECT PAYMENTS.REMARKS From PAYMENTS WHERE PAYMENTS.INVOICENO='" & txtInvNo.Text & "';"
        myRem = GetFieldValue(Qry)
        LYRcon.Execute "UPDATE PAYMENTS SET PAYMENTS.REMARKS = '" & IIf(myRem <> "", myRem & " ", "") & "REPLACE W/ INV.# " & txtRefNo.Text & "' " & _
                       "WHERE PAYMENTS.INVOICENO='" & txtInvNo.Text & "';"
    
    
    Else    'Debit Sales Return amount to customer's account
        'Add on Ledger
        '***Access
        'Qry = "SELECT ONACCTERMS.BALANCE From ONACCTERMS WHERE (((ONACCTERMS.SPINVNO)='" & txtInvNo.Text & "'));"
        
        'OldBal = Val(GetFieldValue(Qry))
        'myTtl = Val(Replace(txtTotal.Text, ",", ""))
        'With rsLedger
        '    .Open "Ledger", LYRcon, adOpenKeyset, adLockOptimistic
        '    .AddNew
        '    !SPINVNO = txtInvNo.Text
        '    !RefType = "Sales Return"
        '    !RefNo = txtRetNo.Text
        '    !RefDate = txtDateRet.Text
        '    !Credit = myTtl
        '    !OSbal = OldBal - myTtl
        '    !CUSTOMERID = txtCustID.Text
        '    !SeqNo = MaxSeqNo(txtInvNo.Text)
        '    .Update
        '    .Close
        'End With
        
        'mySQL
        'Qry = "SELECT PAYMENTS.REMARKS From PAYMENTS WHERE PAYMENTS.INVOICENO='" & txtRefNo.Text & "';"
        Qry = "SELECT ONACCTERMS.BALANCE From ONACCTERMS WHERE ONACCTERMS.SPINVNO='" & txtInvNo.Text & "';"
    
        OldBal = Val(GetFieldValue(Qry))
        myTtl = Val(Replace(txtTotal.Text, ",", ""))
        LYRcon.Execute "INSERT INTO ledger(SPINVNO, REFTYPE, REFNO, REFDATE, CREDIT, OSBAL, " & _
            "CUSTOMERID, SEQNO)VALUES('" & txtInvNo.Text & "', 'Sales Return', '" & txtRetNo.Text & "', " & _
            "'" & Format(txtDateRet.Text, "yyyy-mm-dd") & "', " & myTtl & ", " & OldBal - myTtl & ", " & _
            "" & txtCustID.Text & ", " & MaxSeqNo(txtInvNo.Text) & ")"
    
        'End With
        'Update Customer and Invoice Balance
        MyStat = IIf(OldBal - myTtl >= 1, "OPEN", "CLOSED")
        
        '***Access
        'LYRcon.Execute "UPDATE ONACCTERMS SET ONACCTERMS.BALANCE = [ONACCTERMS].[BALANCE]-" & myTtl & ", ONACCTERMS.STATUS = '" & MyStat & "' " & _
                       "WHERE (((ONACCTERMS.SPINVNO)='" & txtInvNo.Text & "'));"
        'mySQL
        LYRcon.Execute "UPDATE ONACCTERMS SET ONACCTERMS.BALANCE = ONACCTERMS.BALANCE-" & myTtl & ", ONACCTERMS.STATUS = '" & MyStat & "' " & _
                       "WHERE ONACCTERMS.SPINVNO='" & txtInvNo.Text & "';"
        
        'Update Customer Balance
        '***Access
        'LYRcon.Execute "UPDATE CUSTOMER SET CUSTOMER.OSBAL = [CUSTOMER].[OSBAL]-" & myTtl & _
                       " WHERE (((CUSTOMER.CUSTOMERID)=" & txtCustID.Text & "));"
    
        'MySQL
        LYRcon.Execute "UPDATE CUSTOMER SET CUSTOMER.OSBAL = CUSTOMER.OSBAL-" & myTtl & _
                       " WHERE CUSTOMER.CUSTOMERID=" & txtCustID.Text & ";"
    
        
    End If
End Sub


Function MaxSeqNo(MyInv As String) As Integer
    Dim rsSeq As New ADODB.Recordset
    Dim Qry As String
    
    '***Access
    'Qry = "SELECT Max(LEDGER.SEQNO) AS MaxOfSEQNO From LEDGER WHERE (((LEDGER.SPINVNO)='" & MyInv & "'));"
    
    'mysql
    Qry = "SELECT Max(LEDGER.SEQNO) AS MaxOfSEQNO From LEDGER WHERE LEDGER.SPINVNO='" & MyInv & "';"
    
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

Function GetSalesRetNo() As String
    Dim rsInvNo As New ADODB.Recordset
    Dim Qry As String
    Dim MyNo As Long
    
    '***Access
    'Qry = "SELECT Max(round(SRETURNSP.RETNO)) AS MaxNo FROM SRETURNSP;"
    'mySQL
    Qry = "SELECT Max(round(SRETURNSP.RETNO)) AS MaxNo FROM SRETURNSP;"
    
    With rsInvNo
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            MyNo = IIf(IsNull(.Fields(0)), 1, .Fields(0) + 1)
        Else
            MyNo = 1
        End If
        .Close
    End With
    GetSalesRetNo = Format(MyNo, "0000")
End Function


Function ValidateSalesReturn() As String
    Dim vMsg As String
    Dim i As Integer, Cntr As Integer
    
    vMsg = ""

    If Not IsDate(txtDateRet.Text) Then vMsg = "Invalid Return Date"
    If txtRetNo.Text = "" Then
        vMsg = vMsg & vbCrLf & "Incomplete entry on field Sales Return No."
    Else
        If AddNew Then
            '***Access
            'If CtrlNumExist("SELECT SRETURNSP.RETNO From SRETURNSP WHERE (((SRETURNSP.RETNO)='" & txtRetNo.Text & "'));") Then
            'mySQL
            If CtrlNumExist("SELECT SRETURNSP.RETNO From SRETURNSP WHERE SRETURNSP.RETNO='" & txtRetNo.Text & "';") Then
                vMsg = vMsg & vbCrLf & "Sales Return No. already exist!"
            End If
        End If
    End If
    If txtInvNo.Text = "" Then
        vMsg = vMsg & vbCrLf & "Incomplete entry on field Invoice Number"
    Else
        '***Access
        'If Not CtrlNumExist("SELECT SPINVOICE.SPINVNO From SPINVOICE WHERE (((SPINVOICE.SPINVNO)='" & txtInvNo.Text & "'));") Then
        'mySQL
        If Not CtrlNumExist("SELECT SPINVOICE.SPINVNO From SPINVOICE WHERE SPINVOICE.SPINVNO='" & txtInvNo.Text & "' AND SPINVOICE.INVTYPE = '" & IIf(UpDnINV.Value = 1, "CASH", "CREDIT") & "';") Then
            vMsg = vMsg & vbCrLf & "Invoice No. does not exist!"
        End If
    End If
    If cboReason.Text = "" Then vMsg = vMsg & vbCrLf & "Incomplete entry on field Reason"
    If txtAprvBy.Text = "" Then vMsg = vMsg & vbCrLf & "Incomplete entry on field Approved By"
    If txtPrepBy.Text = "" Then vMsg = vMsg & vbCrLf & "Incomplete entry on field Prepared By"
    If txtInvType.Text = "CASH" Then
        If OptProcess(0).Value = True Then
            If txtRefNo.Text = "" Or Val(lblRefAmt.Caption) = 0 Then vMsg = vMsg & vbCrLf & "Incomplete entry on " & lblRef.Caption
        End If
    Else
        If chkReplace.Value = 1 Then If txtRefNo.Text = "" Or Val(lblRefAmt.Caption) = 0 Then vMsg = vMsg & vbCrLf & "Incomplete entry on " & lblRef.Caption
    End If
    If flxParts.TextMatrix(flxParts.Rows - 1, 1) = "" Or flxParts.TextMatrix(flxParts.Rows - 1, 7) = "" Then
        vMsg = vMsg & vbCrLf & "Incomplete entry on last row of Spare Parts for return"
    Else
        With flxParts
            For i = 1 To .Rows - 1
                If Val(.TextMatrix(i, 1)) > Val(.TextMatrix(i, 8)) Then
                    vMsg = vMsg & vbCrLf & "Qty. Returned must not be greater than Qty. Purchased at row " & i
                End If
                If Val(.TextMatrix(i, 1)) = 0 Then
                    vMsg = vMsg & vbCrLf & "No Qty. Returned specified at row " & i
                End If
            Next i
        End With
    End If
    ValidateSalesReturn = vMsg
    
End Function

Public Sub DisplaySalesReturn(MyRef As String)
    Dim rsRet As New ADODB.Recordset
    Dim rsDtls As New ADODB.Recordset
    Dim Qry As String, MyUnit As String
    
    Call ClearEntries
    '***Access
    'Qry = "SELECT SRETURNSP.* From SRETURNSP WHERE (((SRETURNSP.RETURNID)=" & MyRef & "));"
    
    'mySQL
    Qry = "SELECT SRETURNSP.* From SRETURNSP WHERE SRETURNSP.RETURNID=" & MyRef & ";"
    
    DispOnly = True

    With rsRet
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            txtRetNo.Text = !RetNo
            txtDateRet.Text = Format(!RetDate, "mm/dd/yyyy")
            cboReason.Text = IIf(IsNull(!REASON), "", !REASON)
            txtRemarks.Text = IIf(IsNull(!Remarks), "", !Remarks)
            txtInvNo.Text = !Invoice
            txtTotal.Text = Format$(!TotalAmt, "###,##0.00")
            txtPrepBy.Text = !PreparedBy
            txtAprvBy.Text = !ApprovedBy
            txtPrnCopy.Text = IIf(IsNull(!PRINTCOPY), 0, !PRINTCOPY)
            txtReturnID.Text = !RETURNID
            txtInvType.Text = IIf(UpDnINV.Value = 1, "CASH", "CREDIT")
            If txtInvType.Text = "CASH" Then
                OptProcess(0).Visible = True
                OptProcess(1).Visible = True
                If !PROCESS = "REPLACE" Then
                    OptProcess(0).Value = True
                    Call OptProcess_Click(0)
                Else
                    OptProcess(1).Value = True
                    Call OptProcess_Click(1)
                End If
            Else
                chkReplace.Visible = True
                chkReplace.Top = OptProcess(0).Top
                chkReplace.Left = OptProcess(0).Left
                OptProcess(0).Visible = False
                OptProcess(1).Visible = False
                chkReplace.Value = IIf(!PROCESS = "REPLACE", 1, 0)
            End If
            
            txtRefNo.Text = IIf(IsNull(!ProRef), "", !ProRef)
            Call txtrefno_LostFocus
        End If
        .Close
    End With
    '***Access
    'Qry = "SELECT SPINVOICE.CUSTOMERID From SPINVOICE WHERE (((SPINVOICE.SPINVNO)='" & txtInvNo.Text & "'));"
    
    'mySQL
    Qry = "SELECT SPINVOICE.CUSTOMERID From SPINVOICE WHERE SPINVOICE.SPINVNO='" & txtInvNo.Text & "' AND SPINVOICE.INVTYPE = '" & IIf(UpDnINV.Value = 1, "CASH", "CREDIT") & "';"
    txtCustID.Text = GetFieldValue(Qry)
    DispOnly = True
    Call DisplaySalesInvoice
    If txtPrnCopy.Text = 0 Then
        '***Access
        'Qry = "SELECT SRETURNDTLS.QTY, SRETURNDTLS.UNIT, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SRETURNDTLS.PRICE, [PRICE]*[QTY] AS TTLAMT, SRETURNDTLS.SPAREPARTID, SRETURNDTLS.QTYPURCHASED " & _
              "FROM SRETURNDTLS INNER JOIN SPAREPARTS ON SRETURNDTLS.SPAREPARTID = SPAREPARTS.SPAREPARTID WHERE (((SRETURNDTLS.RETURNID)=" & MyRef & "));"
        
        'mySQL
        Qry = "SELECT SRETURNDTLS.QTY, SRETURNDTLS.UNIT, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SRETURNDTLS.PRICE, SRETURNDTLS.PRICE*SRETURNDTLS.QTY AS TTLAMT, SRETURNDTLS.SPAREPARTID, SRETURNDTLS.QTYPURCHASED " & _
              "FROM SRETURNDTLS INNER JOIN SPAREPARTS ON SRETURNDTLS.SPAREPARTID = SPAREPARTS.SPAREPARTID WHERE SRETURNDTLS.RETURNID=" & MyRef & ";"
        
        With rsDtls
            .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
            If .RecordCount > 0 Then
                .MoveFirst
                flxParts.Rows = 1
                Do While Not .EOF
                    flxParts.AddItem vbTab & !Qty & vbTab & !Unit & vbTab & !PartNumber & vbTab & !Description & vbTab & Format(!Price, "###,##0.00") & vbTab & Format(!TtlAmt, "###,##0.00") & vbTab & !SPAREPARTID & vbTab & !QtyPurchased
                    .MoveNext
                Loop
            End If
            .Close
        End With
    Else
        '***Access
        'Qry = "SELECT SPITR.RECEIVE, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPITR.PRICE, [PRICE]*[RECEIVE] AS TOTALAMT, SPAREPARTS.SPAREPARTID, SPITR.REMARKS " & _
              "FROM SPITR INNER JOIN SPAREPARTS ON SPITR.SPAREPARTID = SPAREPARTS.SPAREPARTID WHERE (((SPITR.REFERENCE)='RET#" & txtRetNo.Text & "') AND ((SPITR.TRANSTYPE)='SALES RETURN'));"
        'mySQL
        Qry = "SELECT SPITR.RECEIVE, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPITR.PRICE, SPITR.PRICE*SPITR.RECEIVE AS TOTALAMT, SPAREPARTS.SPAREPARTID, SPITR.REMARKS, SPITR.COST " & _
              "FROM SPITR INNER JOIN SPAREPARTS ON SPITR.SPAREPARTID = SPAREPARTS.SPAREPARTID WHERE SPITR.REFERENCE='RET#" & txtRetNo.Text & "' AND SPITR.TRANSTYPE='SALES RETURN';"
        
        With rsDtls
            .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
            If .RecordCount > 0 Then
                .MoveFirst
                txtInvType.Text = IIf(IsNull(!Remarks), "", !Remarks)
                flxParts.Rows = 1
                Do While Not .EOF
                    If InStr(!Description, "ENGINE OIL") > 0 Or InStr(!Description, "4T OIL") > 0 Then
                        MyUnit = IIf(!RECEIVE > 1, "QRTS.", "QRT.")
                    Else
                        MyUnit = IIf(!RECEIVE > 1, "PCS.", "PC.")
                    End If
                    flxParts.AddItem vbTab & !RECEIVE & vbTab & MyUnit & vbTab & !PartNumber & vbTab & !Description & vbTab & Format(!Price, "###,##0.00") & vbTab & Format(!TotalAmt, "###,##0.00") & vbTab & !SPAREPARTID & vbTab & vbTab & Format(!Cost, "###,##0.00") '& vbTab & !QtyPurchased
                    .MoveNext
                Loop
            End If
            .Close
        End With
    End If
    Call LockData(True)
    Call ComputeTotal
    
End Sub

Function GetInvType()
    Dim rsInv As New ADODB.Recordset
    Dim Qry As String
    
    If txtInvNo.Text = "" Then
        GetInvType = ""
        Exit Function
    End If
    '***Access
    'Qry = "SELECT SPINVOICE.SPINVNO, SPINVOICE.INVTYPE " & _
          "From SPINVOICE WHERE (((SPINVOICE.SPINVNO)='" & txtInvNo.Text & "'));"
    
    'mySQL
    Qry = "SELECT SPINVOICE.SPINVNO, SPINVOICE.INVTYPE " & _
          "From SPINVOICE WHERE SPINVOICE.SPINVNO='" & txtInvNo.Text & "' AND SPINVOICE.INVTYPE = '" & IIf(UpDnINV.Value = 1, "CASH", "CREDIT") & "';"
    
    With rsInv
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            GetInvType = IIf(IsNull(!InvType), "", !InvType)
        Else
            GetInvType = ""
        End If
        .Close
    End With
End Function

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

Private Sub txtDateRet_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtFlx_KeyPress(KeyAscii As Integer)
    Dim strValid As String
    
    strValid = "0123456789"
    If KeyAscii > 26 Then
        If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
    If KeyAscii = 13 Then
        Call txtFlx_LostFocus
        flxParts.SetFocus
        flxParts.Col = ActCol + 1
    End If
End Sub

Private Sub txtFlx_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        txtFlx.Tag = "skiplostfocus"
        txtFlx.Visible = False
        flxParts.SetFocus
        flxParts.Col = ActCol
    End If
End Sub

Private Sub txtFlx_LostFocus()
    Dim Price As Single
        
    If txtFlx.Tag = "skiplostfocus" Then
        txtFlx.Tag = ""
        Exit Sub
    End If
    If Val(txtFlx.Text) > Val(flxParts.TextMatrix(ActRow, 8)) Then
        MsgBox "Qty. Return must not greater than Qty. Purchased!", vbExclamation + vbOKOnly, "Validation Error"
        Exit Sub
    End If
    flxParts.TextMatrix(ActRow, 1) = txtFlx.Text
    txtFlx.Visible = False
    If flxParts.TextMatrix(ActRow, 5) <> "" Then
        Price = PureNum(flxParts.TextMatrix(ActRow, 5))
        flxParts.TextMatrix(ActRow, 6) = Format$(Val(txtFlx.Text) * Price, "#,##0.00")
        Call ComputeTotal
    End If
    Call QtyTip(False)
End Sub
                                             
Private Sub txtInvNo_GotFocus()
    If txtInvNo.Text <> "" Then
        txtInvNo.SelStart = 0
        txtInvNo.SelLength = Len(txtInvNo.Text)
    End If
End Sub

Private Sub txtInvNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtInvNo_LostFocus()
    Dim Qry As String
    
    If txtInvNo.Text <> "" Then
        If UpDnINV.Value = 1 Then
            Qry = "SELECT SPINVOICE.SPINVNO From SPINVOICE WHERE SPINVOICE.SPINVNO='" & txtInvNo.Text & "' AND SPINVOICE.INVTYPE = 'CASH';"
        Else
            Qry = "SELECT SPINVOICE.SPINVNO From SPINVOICE WHERE SPINVOICE.SPINVNO='" & txtInvNo.Text & "' AND SPINVOICE.INVTYPE = 'CREDIT';"
        End If
        If Not CtrlNumExist(Qry) Then
            MsgBox "Invoice No. does not exist!", vbExclamation + vbOKOnly, "Invoice not found"
            Exit Sub
        Else
            Call DisplaySalesInvoice
        End If
    End If
End Sub

Private Sub txtPrepBy_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtAprvBy.SetFocus
End Sub

Private Sub txtrefno_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cboReason.SetFocus
End Sub

Private Sub txtrefno_LostFocus()
    Dim rsRef As New ADODB.Recordset
    Dim Qry As String, vMsg As String
    
    If txtRefNo.Text <> "" Then
        If OptProcess(0).Value = True Then
            '***Access
            'Qry = "SELECT SPINVOICE.SPINVNO, SPINVOICE.AMOUNT From SPINVOICE WHERE (((SPINVOICE.SPINVNO)='" & txtRefNo.Text & "'));"
            'mySQL
            Qry = "SELECT SPINVOICE.SPINVNO, SPINVOICE.AMOUNT From SPINVOICE WHERE SPINVOICE.SPINVNO='" & txtRefNo.Text & "';"
            
            rsRef.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
            If rsRef.RecordCount > 0 Then
                rsRef.MoveFirst
                lblRefAmt.Caption = Format$(rsRef!Amount, "#,##0.00")
                vMsg = ""
            Else
                vMsg = "Sales Invoice not found!"
            End If
            rsRef.Close
        Else
            '***Access
            'Qry = "SELECT DISBURSEMENT.CVNUMBER, DISBURSEMENT.AMOUNT From Disbursement WHERE (((DISBURSEMENT.CVNUMBER)=" & txtRefNo.Text & "));"
            
            'mySQL
            Qry = "SELECT DISBURSEMENT.CVNUMBER, DISBURSEMENT.AMOUNT From Disbursement WHERE DISBURSEMENT.CVNUMBER=" & txtRefNo.Text & ";"
            
            rsRef.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
            If rsRef.RecordCount > 0 Then
                rsRef.MoveFirst
                lblRefAmt.Caption = Format$(rsRef!Amount, "#,##0.00")
                vMsg = ""
            Else
                vMsg = "Petty Cash Voucher not found!"
            End If
            rsRef.Close
        End If
    End If
    If vMsg <> "" Then
        MsgBox vMsg, vbCritical + vbOKOnly, "Validation Error"
    Else
        If Not DispOnly Then
            If OptProcess(0).Value = True Then
                txtRemarks.Text = "Items replaced with Invoice No. " & txtRefNo.Text
            Else
                txtRemarks.Text = "Amount refunded thru P.C.V. No. " & txtRefNo.Text
            End If
        End If
    End If
End Sub


Private Sub txtRetNo_GotFocus()
    If txtRetNo.Text <> "" Then
        txtRetNo.SelStart = 0
        txtRetNo.SelLength = Len(txtRetNo.Text)
    End If
End Sub

Private Sub txtRetNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtRetNo_LostFocus()
    If txtRetNo.Text <> "" Then
        '***Access
        'If CtrlNumExist("SELECT SRETURNSP.RETNO From SRETURNSP WHERE (((SRETURNSP.RETNO)='" & txtRetNo.Text & "'));") Then
        'mySQL
        If CtrlNumExist("SELECT SRETURNSP.RETNO From SRETURNSP WHERE SRETURNSP.RETNO='" & txtRetNo.Text & "';") Then
            MsgBox "Sales Return No. already exist!", vbExclamation + vbOKOnly, "Sales Return"
        End If
    End If
End Sub

Function GetPartsMan(MyID As Integer) As String
    Dim rsUser As New ADODB.Recordset
    Dim Qry As String
    
    '***Access
    'Qry = "SELECT PCUSER.USERNAME, PCUSER.EMPID From PCUSER WHERE (((PCUSER.EMPID)=" & MyId & "));"
    
    'mySQL
    Qry = "SELECT PCUSER.USERNAME, PCUSER.EMPID From PCUSER WHERE PCUSER.EMPID=" & MyID & ";"
    With rsUser
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            GetPartsMan = IIf(IsNull(!UserName), "", UCase$(!UserName))
        Else
            GetPartsMan = ""
        End If
        .Close
    End With
End Function

Private Sub DisplaySalesInvoice()
    Dim rsInv As New ADODB.Recordset
    Dim rsITR As New ADODB.Recordset
    Dim Qry As String, MyUnit As String, MyVal As String
    Dim MyQty As Integer
    
    If UpDnINV.Value = 1 Then
        Qry = "SELECT SPINVOICE.* From SPINVOICE WHERE SPINVOICE.SPINVNO='" & txtInvNo.Text & "' AND SPINVOICE.INVTYPE = 'CASH';"
    Else
        Qry = "SELECT SPINVOICE.* From SPINVOICE WHERE SPINVOICE.SPINVNO='" & txtInvNo.Text & "' AND SPINVOICE.INVTYPE = 'CREDIT';"
    End If
    
    With rsInv
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            txtInvDate.Text = Format$(!InvDate, "mm/dd/yyyy")
            If UCase$(!ChargeTo) = "CASH" Then
                lblCustomer.Caption = "WALK-IN"
            Else
                lblCustomer.Caption = !ChargeTo
            End If
            txtInvType.Text = !InvType
            lblAddress.Caption = IIf(IsNull(!Address), "", !Address)
            txtPartsman.Text = GetPartsMan(!EmpId)
            txtCustID.Text = IIf(IsNull(!CUSTOMERID), 0, !CUSTOMERID)
            If Not DispOnly Then txtRemarks.Text = IIf(IsNull(!Remarks), "", !Remarks)
        End If
        .Close
    End With
    If DispOnly Then
        DispOnly = False
        Exit Sub
    End If
    cmdSelect.Enabled = True
    cmdReturnAll.Enabled = True
    'Auto Display sparepart in flx if ONLY 1 item was purchased
    'Qry = "SELECT SPITR.OUT, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPITR.PRICE, [PRICE]*[OUT] AS TTLAMT, SPITR.SPAREPARTID, SPITR.REFERENCE " & _
          "FROM SPITR INNER JOIN SPAREPARTS ON SPITR.SPAREPARTID = SPAREPARTS.SPAREPARTID WHERE (((SPITR.REFERENCE)='SI#-" & txtInvNo.Text & "'));"
    'mySQL
'    If UCase$(Left$(txtInvNo.Text, 4)) = "TRA#" Then
'        Qry = "SELECT SPITR.OUT, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPITR.PRICE, SPITR.PRICE*SPITR.OUT AS TTLAMT, SPITR.SPAREPARTID, SPITR.REFERENCE, SPITR.COST " & _
'                "FROM SPITR INNER JOIN SPAREPARTS ON SPITR.SPAREPARTID = SPAREPARTS.SPAREPARTID WHERE SPITR.REFERENCE='" & txtInvNo.Text & "' AND SPITR.REMARKS = '" & IIf(UpDnINV.Value = 1, "CASH", "CREDIT") & "';"
'    Else
        Qry = "SELECT SPITR.OUT, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPITR.PRICE, SPITR.PRICE*SPITR.OUT AS TTLAMT, SPITR.SPAREPARTID, SPITR.REFERENCE, SPITR.COST " & _
                "FROM SPITR INNER JOIN SPAREPARTS ON SPITR.SPAREPARTID = SPAREPARTS.SPAREPARTID WHERE SPITR.REFERENCE='SI#-" & txtInvNo.Text & "' AND SPITR.REMARKS = '" & IIf(UpDnINV.Value = 1, "CASH", "CREDIT") & "';"
   
    MyVal = ""
    With rsITR
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount = 1 Then
            .MoveFirst
            flxParts.Rows = 1
            If InStr(!Description, "ENGINE OIL") > 0 Or InStr(!Description, "4T OIL") > 0 Then
                MyUnit = IIf(!OUT > 1, "QRTS.", "QRT.")
            Else
                MyUnit = IIf(!OUT > 1, "PCS.", "PC.")
            End If
            MyQty = !OUT - ReturnedQty(!SPAREPARTID)
            If MyQty > 0 Then
                flxParts.AddItem vbTab & MyQty & vbTab & MyUnit & vbTab & !PartNumber & vbTab & !Description & vbTab & Format(!Price, "###,##0.00") & vbTab & Format(!TtlAmt, "###,##0.00") & vbTab & !SPAREPARTID & vbTab & MyQty & vbTab & Format(!Cost, "###,##0.00")
            Else
                flxParts.AddItem ""
                MyVal = "No items available for return!"
            End If
        End If
        .Close
    End With
    Call ComputeTotal
    If txtInvType.Text = "CASH" Then
        OptProcess(0).Visible = True
        OptProcess(1).Visible = True
    Else
        chkReplace.Visible = True
        chkReplace.Top = OptProcess(0).Top
        chkReplace.Left = OptProcess(0).Left
        OptProcess(0).Visible = False
        OptProcess(1).Visible = False
    End If
    If MyVal <> "" Then
        MsgBox MyVal, vbCritical + vbOKOnly, "Item already Returned"
        BtnIndx = 6
        tmrFloatRev.Enabled = True
        Call LockData(True)
    End If
End Sub

Function ReturnedQty(SPid As Long) As Integer
    Dim rsRetQ As New ADODB.Recordset
    Dim Qry As String
    
    '***Access
    'Qry = "SELECT SPITR.RECEIVE, SPITR.TRANSTYPE, SPITR.SPAREPARTID, SPITR.PACKINGLIST From SPITR " & _
          "WHERE (((SPITR.TRANSTYPE)='SALES RETURN') AND ((SPITR.SPAREPARTID)=" & SPid & ") AND ((SPITR.PACKINGLIST)='" & txtInvNo.Text & "'));"
    'mySQL
    Qry = "SELECT SPITR.RECEIVE, SPITR.TRANSTYPE, SPITR.SPAREPARTID, SPITR.PACKINGLIST From SPITR " & _
          "WHERE SPITR.TRANSTYPE='SALES RETURN' AND SPITR.SPAREPARTID=" & SPid & " AND SPITR.PACKINGLIST='" & txtInvNo.Text & "';"
    
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

Private Sub UpDnINV_Change()
    If UpDnINV.Value = 1 Then
        Label1(0).Caption = "CASH INV.#"
    Else
        Label1(0).Caption = "CHARGE INV.#"
    End If
End Sub
