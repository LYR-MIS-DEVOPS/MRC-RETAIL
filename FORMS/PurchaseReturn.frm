VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form PurchaseReturn 
   Caption         =   "Spare Parts Purchase Return"
   ClientHeight    =   8355
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11520
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8355
   ScaleWidth      =   11520
   Tag             =   "XYF"
   WindowState     =   2  'Maximized
   Begin MSFlexGridLib.MSFlexGrid flxPrn 
      Height          =   735
      Left            =   3960
      TabIndex        =   44
      Top             =   8280
      Visible         =   0   'False
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   1296
      _Version        =   393216
      Rows            =   1
      Cols            =   7
      FixedRows       =   0
      FixedCols       =   0
   End
   Begin VB.TextBox txtPuRtrnID 
      Height          =   285
      Left            =   2880
      TabIndex        =   43
      Top             =   8160
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtPackingID 
      Height          =   285
      Left            =   2040
      TabIndex        =   42
      Top             =   8160
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtRem 
      Height          =   285
      Left            =   6600
      TabIndex        =   37
      Top             =   9360
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtQtyTip 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Height          =   285
      Left            =   360
      TabIndex        =   34
      Text            =   "Qty. Purchased: 0"
      Top             =   8160
      Visible         =   0   'False
      Width           =   1575
   End
   Begin lvButton.lvButtons_H cmdReturnAll 
      Height          =   375
      Left            =   8280
      TabIndex        =   4
      Tag             =   "XYF"
      Top             =   2475
      Visible         =   0   'False
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
      Image           =   "PurchaseReturn.frx":0000
      cBack           =   -2147483633
   End
   Begin VB.TextBox txtPrnCopy 
      Height          =   285
      Left            =   8040
      TabIndex        =   33
      Top             =   9360
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtTotal 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   7920
      Locked          =   -1  'True
      TabIndex        =   31
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
      TabIndex        =   21
      Top             =   9360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtRemarks 
      Height          =   645
      Left            =   1320
      MaxLength       =   200
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Tag             =   "XYF"
      Top             =   7320
      Width           =   4935
   End
   Begin MSFlexGridLib.MSFlexGrid flxParts 
      Height          =   3495
      Left            =   120
      TabIndex        =   5
      Tag             =   "XY"
      Top             =   2880
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   6165
      _Version        =   393216
      Cols            =   11
      FixedCols       =   0
      RowHeightMin    =   315
      BackColorBkg    =   -2147483633
      ScrollBars      =   2
      FormatString    =   "^    |^QTY.   |^UNIT       |<PART NUMBER   |<DESCRIPTION|>UNIT COST |>AMOUNT|<P.O. No.|<REMARKS  "
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
      Height          =   1650
      Left            =   120
      TabIndex        =   10
      Tag             =   "XYF"
      Top             =   720
      Width           =   11295
      Begin VB.TextBox txtWayBill 
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
         Left            =   9240
         Locked          =   -1  'True
         TabIndex        =   40
         Tag             =   "XYF"
         Top             =   1200
         Width           =   1815
      End
      Begin VB.TextBox txtShipper 
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
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   38
         Tag             =   "XYF"
         Top             =   1200
         Width           =   4815
      End
      Begin VB.TextBox txtSupplier 
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
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   35
         Tag             =   "XYF"
         Top             =   871
         Width           =   4815
      End
      Begin VB.TextBox txtPacking 
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
         Left            =   1560
         MaxLength       =   15
         TabIndex        =   0
         Tag             =   "XYF"
         Top             =   215
         Width           =   2175
      End
      Begin VB.TextBox txtDateRec 
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
         Left            =   1560
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   9
         Tag             =   "XYF"
         Top             =   543
         Width           =   2175
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
         Left            =   9240
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
         Left            =   9240
         TabIndex        =   1
         Tag             =   "XYF"
         Top             =   200
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "WAY BILL NO."
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
         Left            =   7680
         TabIndex        =   41
         Tag             =   "XYF"
         Top             =   1230
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "SHIPPER:"
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
         Left            =   240
         TabIndex        =   39
         Tag             =   "XYF"
         Top             =   1230
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "SUPPLIER:"
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
         TabIndex        =   36
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
         Left            =   7680
         TabIndex        =   15
         Tag             =   "XYF"
         Top             =   570
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "PURCHASE RETURN NO."
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
         Left            =   7080
         TabIndex        =   14
         Tag             =   "XYF"
         Top             =   225
         Width           =   2055
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
         Left            =   240
         TabIndex        =   13
         Tag             =   "XYF"
         Top             =   900
         Width           =   5175
      End
      Begin VB.Label Label1 
         Caption         =   "DATE RECEIVED:"
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
         TabIndex        =   12
         Tag             =   "XYF"
         Top             =   568
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "PACKING LIST:"
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
         TabIndex        =   11
         Tag             =   "XYF"
         Top             =   230
         Width           =   1455
      End
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   22
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
      Image           =   "PurchaseReturn.frx":059A
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   2
      Left            =   3000
      TabIndex        =   23
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
      Image           =   "PurchaseReturn.frx":0B34
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   3
      Left            =   4500
      TabIndex        =   24
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
      Image           =   "PurchaseReturn.frx":10CE
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   4
      Left            =   6000
      TabIndex        =   25
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
      Image           =   "PurchaseReturn.frx":1668
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   5
      Left            =   7500
      TabIndex        =   26
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
      Image           =   "PurchaseReturn.frx":1C02
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   6
      Left            =   9000
      TabIndex        =   27
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
      Image           =   "PurchaseReturn.frx":219C
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   1
      Left            =   1500
      TabIndex        =   28
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
      Image           =   "PurchaseReturn.frx":2736
      cBack           =   -2147483633
   End
   Begin VB.TextBox txtPrepBy 
      Height          =   285
      Left            =   8400
      TabIndex        =   7
      Tag             =   "XYF"
      Top             =   7320
      Width           =   2775
   End
   Begin VB.TextBox txtAprvBy 
      Height          =   285
      Left            =   8400
      TabIndex        =   8
      Tag             =   "XYF"
      Top             =   7680
      Width           =   2775
   End
   Begin lvButton.lvButtons_H cmdSelect 
      Height          =   375
      Left            =   9720
      TabIndex        =   3
      Tag             =   "XYF"
      Top             =   2475
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
      Image           =   "PurchaseReturn.frx":2CD0
      cBack           =   -2147483633
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
      Left            =   6840
      TabIndex        =   32
      Tag             =   "XYF"
      Top             =   6415
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
      Left            =   6840
      TabIndex        =   30
      Tag             =   "XYF"
      Top             =   7695
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
      Left            =   6840
      TabIndex        =   29
      Tag             =   "XYF"
      Top             =   7320
      Width           =   1335
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   120
      Top             =   4680
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   11640
      m_MDI_Height    =   8865
   End
   Begin VB.Label lblHeading 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PURCHASE RETURN"
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
      Left            =   9270
      TabIndex        =   19
      Tag             =   "TRF"
      Top             =   450
      Width           =   1905
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
      Left            =   240
      TabIndex        =   17
      Tag             =   "TLF"
      Top             =   450
      Width           =   2895
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
      Left            =   240
      TabIndex        =   16
      Tag             =   "XYF"
      Top             =   7320
      Width           =   1035
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
      Left            =   270
      TabIndex        =   18
      Tag             =   "TLF"
      Top             =   480
      Width           =   2895
   End
   Begin VB.Label lblHeading 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PURCHASE RETURN"
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
      Left            =   9300
      TabIndex        =   20
      Tag             =   "TRF"
      Top             =   480
      Width           =   1905
   End
   Begin VB.Menu mnuRows 
      Caption         =   "SpareParts"
      Visible         =   0   'False
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete Row"
      End
   End
End
Attribute VB_Name = "PurchaseReturn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ActRow As Integer, ActCol As Integer, BtnIndx As Integer
Dim AddNew As Boolean, LockFlx As Boolean, DispOnly As Boolean

Private Sub ClearEntries()
    txtPacking.Text = ""
    txtDateRec.Text = ""
    txtSupplier.Text = ""
    txtShipper.Text = ""
    txtRetNo.Text = ""
    txtDateRet.Text = ""
    txtWayBill.Text = ""
    txtTotal.Text = ""
    txtRemarks.Text = ""
    txtPrepBy.Text = ""
    txtAprvBy.Text = ""
    txtPrnCopy.Text = ""
    txtPackingID.Text = ""
    txtPuRtrnID.Text = ""
    txtPuRtrnID.Tag = ""
    flxParts.Rows = 1
    flxParts.Rows = 2
End Sub

Private Sub LockData(bln As Boolean)
    Frame1.Enabled = Not bln
    txtRemarks.Locked = bln
    txtPrepBy.Locked = bln
    txtAprvBy.Locked = bln
    LockFlx = bln
    cmdSelect.Enabled = Not bln
End Sub

Public Sub ComputeTotal()
    Dim Ttl As Currency, i As Integer
    
    Ttl = 0
    For i = 1 To flxParts.Rows - 1
        Ttl = Ttl + PureNum(flxParts.TextMatrix(i, 6))
    Next i
    txtTotal.Text = Format$(Ttl, "#,##0.00")
    
End Sub


Private Sub cmdSelect_Click()
    If txtPacking.Text <> "" And txtPackingID.Text <> "" Then
        PurchRetPacking.Show vbModal
    Else
        MsgBox "Invalid Packing List Number!", vbExclamation + vbOKOnly, "Purchase Return"
    End If
End Sub

Private Sub cmdTrans_Click(Index As Integer)
    Dim Ans As Integer, i As Integer
    Dim myMsg As String, PrnVal(0 To 5) As String
    Dim Qry As String, MyAcc As String
    
    Select Case Index
        Case 0  'New
            '***Access
            'Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE (((USERACCESS.EMPID)=" & UserID & ") AND ((USERACCESS.MENUINDX)=22));"
            
            'mySQL
            Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE USERACCESS.EMPID=" & UserID & " AND USERACCESS.MENUINDX=22;"
            
            MyAcc = GetFieldValue(Qry)
            If MyAcc <> "FULL ACCESS" Then
                MsgBox "You don't have full access to this transaction." & vbCrLf & "Please contact your System Administrator", vbExclamation + vbOKOnly, "Access Denied"
                Exit Sub
            End If
        
            BtnIndx = 4
            tmrFloat.Enabled = True
            Call ClearEntries
            Call LockData(False)
            txtDateRet.Text = Date
            AddNew = True
            txtRetNo.Text = GetPurchaseRetNo
            txtPrepBy.Text = GetUserName
            txtPacking.Locked = False
            txtPacking.SetFocus
        Case 1  'Edit
            '***Access
            'Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE (((USERACCESS.EMPID)=" & UserID & ") AND ((USERACCESS.MENUINDX)=22));"
            'mySQL
            Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE USERACCESS.EMPID=" & UserID & " AND USERACCESS.MENUINDX=22;"
            
            MyAcc = GetFieldValue(Qry)
            If MyAcc <> "FULL ACCESS" Then
                MsgBox "You don't have full access to this transaction." & vbCrLf & "Please contact your System Administrator", vbExclamation + vbOKOnly, "Access Denied"
                Exit Sub
            End If
            
            If Val(txtPrnCopy.Text) > 0 Then
                MsgBox "Editing is not allowed because Purchase Return was already printed.", vbInformation + vbOKOnly, "Edit Denied"
                Exit Sub
            ElseIf txtPuRtrnID.Text = "" Then
                MsgBox "No available Purchase Return to Edit", vbInformation + vbOKOnly, "Edit"
                Exit Sub
            End If
            BtnIndx = 4
            tmrFloat.Enabled = True
            AddNew = False
            txtPacking.Locked = True
            txtPuRtrnID.Tag = txtRetNo.Text
            cmdSelect.Enabled = True
            Call LockData(False)
        Case 2  'Summary
            FrmSummary.mySrc = "Purchase Return"
            FrmSummary.Show
            'PurchRetSum.Show
            Unload Me
            Exit Sub
        Case 3  'Print
            If flxParts.TextMatrix(flxParts.Rows - 1, 1) = "" Then
                MsgBox "No available data to print", vbCritical, "Print rror"
                Exit Sub
            End If
            Ans = MsgBox("Do you want to Print Purchase Return?", vbQuestion + vbYesNo, "Print")
            If Ans = vbNo Then Exit Sub
            PrnVal(0) = txtSupplier.Text
            PrnVal(1) = txtPacking.Text
            PrnVal(2) = txtDateRet.Text
            PrnVal(3) = txtShipper.Text
            PrnVal(4) = txtDateRec.Text
            PrnVal(5) = IIf(txtRemarks.Text <> "", "Note: " & txtRemarks.Text, "")
            FlxPrn.Rows = 1
            With flxParts
                For i = 1 To .Rows - 1
                    FlxPrn.AddItem vbTab & .TextMatrix(i, 1) & vbTab & .TextMatrix(i, 2) & vbTab & Left$(.TextMatrix(i, 4), 18) & vbTab & .TextMatrix(i, 8) & vbTab & .TextMatrix(i, 5) & vbTab & .TextMatrix(i, 6)
                Next i
            End With
            Call PrintMyForm(PrnVal, "PURCHASE RETURN", FlxPrn)
            If Val(txtPrnCopy.Text) = 0 Then
                txtPrnCopy.Text = Val(txtPrnCopy.Text) + 1
                Call UpdateSPITR
            Else
                txtPrnCopy.Text = Val(txtPrnCopy.Text) + 1
            End If
            'old LYRcon.Execute "UPDATE SRETURNSP SET SRETURNSP.PRINTCOPY = [SRETURNSP].[PRINTCOPY]+1 WHERE (((SRETURNSP.RETURNID)=" & txtReturnID.Text & "));"
            '***Access
            'LYRcon.Execute "UPDATE PURCHASERET SET PURCHASERET.PRINTCOPY = [PURCHASERET].[PRINTCOPY]+1 WHERE (((PURCHASERET.PURTRNID)=" & txtPuRtrnID.Text & " ));"
            
            'mySQL
            LYRcon.Execute "UPDATE PURCHASERET SET PURCHASERET.PRINTCOPY = PURCHASERET.PRINTCOPY+1 WHERE PURCHASERET.PURTRNID=" & txtPuRtrnID.Text & ";"
            
            'Print process follows
            
        Case 4  'Exit
            'MotoRep.tbMain.Visible = True
            'MotoRep.imgBG.Visible = True
            MotoRep.toolBarButton.Visible = True
            
            Unload Me
            Exit Sub
        Case 5  'Save
            myMsg = ValidatePurchaseReturn
            If myMsg <> "" Then
                MsgBox myMsg, vbCritical + vbOKOnly, "Validation Error"
                Exit Sub
            End If
            Ans = MsgBox("Do you want to Save Purchase Return?", vbQuestion + vbYesNo, "Save")
            If Ans = vbNo Then Exit Sub
            Call SavePurchaseReturn
            MsgBox "Purchase Return Successfully Saved!", vbInformation + vbOKOnly, "Save"
                
            BtnIndx = 6
            tmrFloatRev.Enabled = True
            Call LockData(True)
            
        Case 6  'Cancel
            txtRem.Visible = False
            txtFlx.Visible = False
            Ans = MsgBox("Do you want to Cancel Purchase Return?", vbQuestion + vbYesNo, "Cancel")
            If Ans = vbNo Then Exit Sub
            BtnIndx = 6
            tmrFloatRev.Enabled = True
            If AddNew Then
                Call ClearEntries
            Else
                Call DisplayPurchaseReturn(txtPuRtrnID.Text)
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
    If (ActCol <> 1 And ActCol <> 8) Or LockFlx Then Exit Sub
    
    If ActCol = 1 Then
        With txtFlx
            If flxParts.TextMatrix(ActRow, 1) <> "" Then txtFlx.Tag = flxParts.TextMatrix(ActRow, 1)
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
    ElseIf ActCol = 8 Then
        With txtRem
            If flxParts.TextMatrix(ActRow, 8) <> "" Then txtRem.Tag = flxParts.TextMatrix(ActRow, 8)
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
    End If
End Sub

Private Sub QtyTip(Disp As Boolean)
    If Disp Then
        With txtQtyTip
            .Top = flxParts.CellTop + flxParts.Top + 290
            .Left = flxParts.Left + flxParts.ColWidth(0) + 300
            .Text = "Qty. Purchased: " & flxParts.TextMatrix(ActRow, 10)
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

Private Sub flxParts_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
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
    
    MotoRep.sBar.Panels(6).Text = "Spare Parts Purchase Return"
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

Private Sub RXResizer1_ResizeComplete()
    With flxParts
        .ColWidth(0) = 350
        .ColWidth(1) = (.Width - 380) * 0.06
        .ColWidth(2) = (.Width - 380) * 0.06
        .ColWidth(3) = (.Width - 380) * 0.18
        .ColWidth(4) = (.Width - 380) * 0.2
        .ColWidth(5) = (.Width - 380) * 0.1
        .ColWidth(6) = (.Width - 380) * 0.1
        .ColWidth(7) = (.Width - 380) * 0.1
        .ColWidth(8) = (.Width - 380) * 0.18
        .ColWidth(9) = 0
        .ColWidth(10) = 0
        
        txtTotal.Left = .Left + .ColWidth(0) + .ColWidth(1) + .ColWidth(2) + .ColWidth(3) + .ColWidth(4) + .ColWidth(5)
    End With
    txtTotal.Top = flxParts.Height + flxParts.Top
    txtTotal.Width = flxParts.ColWidth(5)
    lblTotal.Top = txtTotal.Top
    lblTotal.Left = txtTotal.Left - lblTotal.Width 'flxParts.Left + (flxParts.Width - (flxParts.ColWidth(4) + flxParts.ColWidth(5)))
    lbl2.Top = txtRemarks.Top
    lbl2.Left = txtRemarks.Left - lbl2.Width
    lbl2.FontSize = txtRemarks.FontSize
End Sub

Private Sub SavePurchaseReturn()
    Dim rsReturn As New ADODB.Recordset
    Dim rsReturnDtls As New ADODB.Recordset
    Dim Qry As String, i As Integer
    
    'Saving Sales Return
    '***Access
    'With rsReturn
    '    If AddNew Then
    '        .Open "PURCHASERET", LYRcon, adOpenKeyset, adLockOptimistic
    '        .AddNew
    '    Else
    '        Qry = "SELECT PURCHASERET.*, PURCHASERET.PURTRNID From PURCHASERET " & _
    '              "WHERE (((PURCHASERET.PURTRNID)=" & txtPuRtrnID.Text & "));"
    '
    '        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    '        .MoveFirst
    '    End If
    '    !PURTRNO = txtRetNo.Text
    '    !DATERET = txtDateRet.Text
    '    !PACKINGNO = txtPacking.Text
    '    !AMOUNT = Val(Replace(txtTotal.Text, ",", ""))
    '    !REMARKS = txtRemarks.Text
    '    !PREPBY = txtPrepBy.Text
    '    !APRVBY = txtAprvBy.Text
    '    !UPDATEDBY = UserID
    '    !LASTUPDATE = Now
    '    .Update
    '    If AddNew Then txtPuRtrnID.Text = !PURTRNID
    '    .Close
    'End With
    
    'mySQL
    If AddNew Then
        myLastUpdt = Format(Now, "yyyy-mm-dd hh:mm:ss")
        Qry = "INSERT INTO purchaseret(PURTRNO, DATERET, PACKINGNO, AMOUNT, " & _
            "REMARKS, PREPBY, APRVBY, UPDATEDBY, LASTUPDATE)VALUES('" & txtRetNo.Text & "', " & _
            "'" & Format(txtDateRet.Text, "yyyy-mm-dd") & "', '" & txtPacking.Text & "', " & _
            "" & Val(Replace(txtTotal.Text, ",", "")) & ", '" & txtRemarks.Text & "', '" & txtPrepBy.Text & "', " & _
            "'" & txtAprvBy.Text & "', " & UserID & ", '" & myLastUpdt & "');"
        LYRcon.Execute Qry
        
        Qry = "select purchaseret.PURTRNID from purchaseret Where purchaseret.PURTRNO='" & txtRetNo.Text & "' AND purchaseret.LASTUPDATE='" & myLastUpdt & "';"
        txtPuRtrnID.Text = GetFieldValue(Qry)
    Else
        LYRcon.Execute "UPDATE purchaseret SET PURTRNO = '" & txtRetNo.Text & "', DATERET = '" & Format(txtDateRet.Text, "yyyy-mm-dd") & "', " & _
            "PACKINGNO = '" & txtPacking.Text & "', AMOUNT = " & Val(Replace(txtTotal.Text, ",", "")) & ", REMARKS = '" & txtRemarks.Text & "', " & _
            "PREPBY = '" & txtPrepBy.Text & "', APRVBY = '" & txtAprvBy.Text & "', UPDATEDBY = " & UserID & " Where purchaseret.PURTRNID = " & txtPuRtrnID.Text & ";"
    End If

    'Sales Return Details
    If Not AddNew Then
        '***Access
        'LYRcon.Execute "DELETE PURCHRETDTLS.* From PURCHRETDTLS WHERE (((PURCHRETDTLS.PURTRNID)=" & txtPuRtrnID.Text & "));"
        'mySQL
        LYRcon.Execute "DELETE From PURCHRETDTLS WHERE PURCHRETDTLS.PURTRNID=" & txtPuRtrnID.Text & ";"
    End If
    With flxParts
        For i = 1 To .Rows - 1
            '***Access
            'LYRcon.Execute "INSERT INTO PURCHRETDTLS ( PURTRNID, SPAREPARTID, QTY, UNIT, PRICE, PONO, REMARKS ) " & _
                           "VALUES (" & txtPuRtrnID.Text & ", " & .TextMatrix(i, 9) & ", " & Val(.TextMatrix(i, 1)) & ", '" & .TextMatrix(i, 2) & "', " & Val(Replace(.TextMatrix(i, 5), ",", "")) & ",  '" & .TextMatrix(i, 7) & "', '" & .TextMatrix(i, 8) & "');"
            'mySQL
            LYRcon.Execute "INSERT INTO PURCHRETDTLS ( PURTRNID, SPAREPARTID, QTY, UNIT, PRICE, PONO, REMARKS ) " & _
                           "VALUES (" & txtPuRtrnID.Text & ", " & .TextMatrix(i, 9) & ", " & Val(.TextMatrix(i, 1)) & ", '" & .TextMatrix(i, 2) & "', " & Val(Replace(.TextMatrix(i, 5), ",", "")) & ",  '" & .TextMatrix(i, 7) & "', '" & .TextMatrix(i, 8) & "');"
        Next i
    End With
End Sub

Private Sub UpdateSPITR()
    Dim rsITR As New ADODB.Recordset
    Dim i As Integer, myTtl As Currency, OldBal As Currency
    
    With flxParts
        For i = 1 To .Rows - 1
            'Update SpareParts Qty on hand
            '***Access
            'LYRcon.Execute "UPDATE SPAREPARTS SET SPAREPARTS.QTYONHAND = [SPAREPARTS].[QTYONHAND]-" & Val(.TextMatrix(i, 1)) & _
                           " WHERE (((SPAREPARTS.SPAREPARTID)=" & .TextMatrix(i, 9) & "));"
            
            'mySQL
            LYRcon.Execute "UPDATE SPAREPARTS SET SPAREPARTS.QTYONHAND = SPAREPARTS.QTYONHAND-" & Val(.TextMatrix(i, 1)) & _
                           " WHERE SPAREPARTS.SPAREPARTID=" & .TextMatrix(i, 9) & ";"
            
            'Saving Inventory Transaction Register
            '***Access
            'rsITR.Open "SPITR", LYRcon, adOpenKeyset, adLockOptimistic
            'rsITR.AddNew
            'rsITR!REFERENCE = "RET#" & txtRetNo.Text
            'rsITR!SPAREPARTID = .TextMatrix(i, 9)
            'rsITR!OUT = Val(.TextMatrix(i, 1))
            'rsITR!TRANSDATE = txtDateRet.Text
            'rsITR!TRANSTYPE = "PURCHASE RETURN"
            'rsITR!Price = Val(Replace(.TextMatrix(i, 5), ",", ""))
            'rsITR!PACKINGLIST = txtPacking.Text
            'rsITR!Source = txtSupplier.Text
            'rsITR!REMARKS = .TextMatrix(i, 8)
            'rsITR!UPDATEDBY = UserID
            'rsITR!LASTUPDATE = Now
            'rsITR.Update
            'rsITR.Close
            'mySQL
            LYRcon.Execute "Insert Into spitr( REFERENCE, SPAREPARTID, `OUT`, TRANSDATE, TRANSTYPE, PRICE, " & _
                "PACKINGLIST, SOURCE, REMARKS, UPDATEDBY, LASTUPDATE)VALUES('RET#" & txtRetNo.Text & "', " & _
                "" & .TextMatrix(i, 9) & ", " & Val(.TextMatrix(i, 1)) & ", '" & Format(txtDateRet.Text, "yyyy-mm-dd") & "', " & _
                "'PURCHASE RETURN', " & Val(Replace(.TextMatrix(i, 5), ",", "")) & ", '" & txtPacking.Text & "', " & _
                "'" & txtSupplier.Text & "', '" & .TextMatrix(i, 8) & "', " & UserID & ", '" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "');"
        Next i
    End With
    
End Sub


Function GetPurchaseRetNo() As String
    Dim rsInvNo As New ADODB.Recordset
    Dim Qry As String
    Dim MyNo As Long
    
    '***Access
    'Qry = "SELECT Max(Val([PURTRNO])) AS RETNO FROM PURCHASERET;"
    'mySQL
    Qry = "SELECT Max(Round(PURCHASERET.PURTRNO)) AS RETNO FROM PURCHASERET;"
    
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
    GetPurchaseRetNo = Format(MyNo, "0000")
End Function


Function ValidatePurchaseReturn() As String
    Dim vMsg As String, Qry As String
    Dim i As Integer, Cntr As Integer
    
    vMsg = ""

    If Not IsDate(txtDateRet.Text) Then vMsg = "Invalid Return Date"
    If txtRetNo.Text = "" Then
        vMsg = vMsg & vbCrLf & "Incomplete entry on field Purchase Return No."
    Else
        If AddNew Then
            '***Access
            'Qry = "SELECT PURCHASERET.PURTRNO From PURCHASERET WHERE (((PURCHASERET.PURTRNO)='" & txtRetNo.Text & "'));"
            
            'mySQL
            Qry = "SELECT PURCHASERET.PURTRNO From PURCHASERET WHERE PURCHASERET.PURTRNO='" & txtRetNo.Text & "';"
            
            If CtrlNumExist(Qry) Then
                vMsg = vMsg & vbCrLf & "Purchase Return No. " & txtRetNo.Text & " already exist!"
            End If
        End If
    End If
    If txtPacking.Text = "" Then
        vMsg = vMsg & vbCrLf & "No Packing List specified"
    Else
        '***Access
        'Qry = "SELECT SPRECEIVE.* From SPRECEIVE WHERE (((SPRECEIVE.PACKINGNO)='" & txtPacking.Text & "'));"
        
        'mySQL
        Qry = "SELECT SPRECEIVE.* From SPRECEIVE WHERE SPRECEIVE.PACKINGNO='" & txtPacking.Text & "';"
        
        If Not CtrlNumExist(Qry) Then
            vMsg = vMsg & vbCrLf & "Packing list No. " & txtPacking.Text & " does not exist!"
        End If
    End If
    If txtAprvBy.Text = "" Then vMsg = vMsg & vbCrLf & "Incomplete entry on field Approved By"
    If txtPrepBy.Text = "" Then vMsg = vMsg & vbCrLf & "Incomplete entry on field Prepared By"
    If flxParts.TextMatrix(flxParts.Rows - 1, 1) = "" Or flxParts.TextMatrix(flxParts.Rows - 1, 9) = "" Then
        vMsg = vMsg & vbCrLf & "Incomplete entry on last row of Spare Parts for return"
    Else
        With flxParts
            For i = 1 To .Rows - 1
                If Val(.TextMatrix(i, 1)) > Val(.TextMatrix(i, 10)) Then
                    vMsg = vMsg & vbCrLf & "Qty. Returned must not be greater than Qty. Purchased at row " & i
                End If
                If Val(.TextMatrix(i, 1)) = 0 Then
                    vMsg = vMsg & vbCrLf & "No Qty. Returned specified at row " & i
                End If
            Next i
        End With
    End If
    ValidatePurchaseReturn = vMsg
    
End Function

Public Sub DisplayPurchaseReturn(MyRef As String)
    Dim rsRet As New ADODB.Recordset
    Dim rsDtls As New ADODB.Recordset
    Dim Qry As String, MyUnit As String
    Dim OrigQty As Integer
    
    Call ClearEntries
    
    '***Access
    'Qry = "SELECT PURCHASERET.* From PURCHASERET WHERE (((PURCHASERET.PURTRNID)=" & MyRef & "));"
    
    'mySQL
    Qry = "SELECT PURCHASERET.* From PURCHASERET WHERE PURCHASERET.PURTRNID=" & MyRef & ";"
    
    With rsRet
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            txtRetNo.Text = !PURTRNO
            txtDateRet.Text = !DATERET
            txtPacking.Text = !packingNO
            txtTotal.Text = Format(!Amount, "#,##0.00")
            txtRemarks.Text = IIf(IsNull(!Remarks), "", !Remarks)
            txtPrepBy.Text = !PREPBY
            txtAprvBy.Text = !APRVBY
            txtPuRtrnID.Text = !PURTRNID
            txtPrnCopy.Text = !PRINTCOPY
        End If
        .Close
    End With
    If txtPacking.Text <> "" Then Call DisplayPackingList(txtPacking.Text)
    If Val(txtPrnCopy.Text) = 0 Then
        '***Access
        'Qry = "SELECT PURCHRETDTLS.QTY, PURCHRETDTLS.UNIT, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, PURCHRETDTLS.PRICE, [QTY]*[PRICE] AS TTLAMT, PURCHRETDTLS.PONO, PURCHRETDTLS.REMARKS, PURCHRETDTLS.SPAREPARTID " & _
              "FROM PURCHRETDTLS INNER JOIN SPAREPARTS ON PURCHRETDTLS.SPAREPARTID = SPAREPARTS.SPAREPARTID WHERE (((PURCHRETDTLS.PURTRNID)=" & MyRef & "));"
        'mySQL
        Qry = "SELECT PURCHRETDTLS.QTY, PURCHRETDTLS.UNIT, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, PURCHRETDTLS.PRICE, PURCHRETDTLS.QTY*PURCHRETDTLS.PRICE AS TTLAMT, PURCHRETDTLS.PONO, PURCHRETDTLS.REMARKS, PURCHRETDTLS.SPAREPARTID " & _
              "FROM PURCHRETDTLS INNER JOIN SPAREPARTS ON PURCHRETDTLS.SPAREPARTID = SPAREPARTS.SPAREPARTID WHERE PURCHRETDTLS.PURTRNID=" & MyRef & ";"
        
        With rsDtls
            .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
            If .RecordCount > 0 Then
                .MoveFirst
                flxParts.Rows = 1
                Do While Not .EOF
                    '***Access
                    'Qry = "SELECT SPITR.RECEIVE From SPITR WHERE (((SPITR.REFERENCE)='" & txtPacking.Text & "') AND ((SPITR.TRANSTYPE)='RECEIVE') AND ((SPITR.PACKINGLIST)='" & !PoNo & "') AND ((SPITR.SPAREPARTID)=" & !SparepartID & "));"
                    'mySQL
                    Qry = "SELECT SPITR.RECEIVE From SPITR WHERE (((SPITR.REFERENCE)='" & txtPacking.Text & "') AND ((SPITR.TRANSTYPE)='RECEIVE') AND ((SPITR.PACKINGLIST)='" & !PONO & "') AND ((SPITR.SPAREPARTID)=" & !SparepartID & "));"
                    
                    OrigQty = Val(GetFieldValue(Qry))
                    flxParts.AddItem vbTab & !Qty & vbTab & !Unit & vbTab & !PartNumber & vbTab & !Description & vbTab & Format(!Price, "###,##0.00") & vbTab & Format(!TtlAmt, "###,##0.00") & vbTab & !PONO & vbTab & IIf(IsNull(!Remarks), "", !Remarks) & vbTab & !SparepartID & vbTab & OrigQty
                    .MoveNext
                Loop
            End If
            .Close
        End With
    Else
        '***Access
        'Qry = "SELECT SPITR.OUT, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPITR.PRICE, [PRICE]*[OUT] AS TOTALAMT, SPITR.PACKINGLIST, SPITR.REMARKS, SPAREPARTS.SPAREPARTID " & _
              "FROM SPITR INNER JOIN SPAREPARTS ON SPITR.SPAREPARTID = SPAREPARTS.SPAREPARTID " & _
              "WHERE (((SPITR.REFERENCE)='RET#" & txtRetNo.Text & "') AND ((SPITR.TRANSTYPE)='PURCHASE RETURN'));"
        'mySQL
        Qry = "SELECT SPITR.`OUT`, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPITR.PRICE, SPITR.PRICE*SPITR.`OUT` AS TOTALAMT, SPITR.PACKINGLIST, SPITR.REMARKS, SPAREPARTS.SPAREPARTID " & _
              "FROM SPITR INNER JOIN SPAREPARTS ON SPITR.SPAREPARTID = SPAREPARTS.SPAREPARTID " & _
              "WHERE SPITR.REFERENCE='RET#" & txtRetNo.Text & "' AND SPITR.TRANSTYPE='PURCHASE RETURN';"
        
        With rsDtls
            .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
            If .RecordCount > 0 Then
                .MoveFirst
                flxParts.Rows = 1
                Do While Not .EOF
                    If InStr(!Description, "ENGINE OIL") > 0 Or InStr(!Description, "4T OIL") > 0 Then
                        MyUnit = IIf(!OUT > 1, "QRTS.", "QRT.")
                    Else
                        MyUnit = IIf(!OUT > 1, "PCS.", "PC.")
                    End If
                    flxParts.AddItem vbTab & !OUT & vbTab & MyUnit & vbTab & !PartNumber & vbTab & !Description & vbTab & Format(!Price, "#,##0.00") & vbTab & Format(!TotalAmt, "#,##0.00") & vbTab & !PACKINGLIST & vbTab & IIf(IsNull(!Remarks), "", !Remarks)
                    .MoveNext
                Loop
            End If
            .Close
        End With
    End If
    Call LockData(True)
    Call ComputeTotal
    
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

Private Sub txtDateRec_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtSupplier.SetFocus
End Sub

Private Sub txtDateRet_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtWayBill.SetFocus
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
        flxParts.TextMatrix(ActRow, 1) = txtFlx.Tag
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
    If Val(txtFlx.Text) > Val(flxParts.TextMatrix(ActRow, 10)) Then
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
                                             
Private Sub txtPacking_GotFocus()
    If txtPacking.Text <> "" Then
        txtPacking.SelStart = 0
        txtPacking.SelLength = Len(txtPacking.Text)
    End If
End Sub

Private Sub txtPacking_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtRetNo.SetFocus
End Sub

Private Sub txtPacking_LostFocus()
    Dim Qry As String
    
    If txtPacking.Text <> "" Then
        '***Access
        'Qry = "SELECT SPRECEIVE.PACKINGNO From SPRECEIVE WHERE (((SPRECEIVE.PACKINGNO)='" & txtPacking.Text & "'));"
        'mySQL
        Qry = "SELECT SPRECEIVE.PACKINGNO From SPRECEIVE WHERE SPRECEIVE.PACKINGNO='" & txtPacking.Text & "';"
        
        If Not CtrlNumExist(Qry) Then
            MsgBox "Packing List does not exist!", vbExclamation + vbOKOnly, "File not found"
            txtSupplier.Text = ""
            txtShipper.Text = ""
            txtDateRec.Text = ""
            txtWayBill.Text = ""
            txtPackingID.Text = ""
            Exit Sub
        Else
            Call DisplayPackingList(txtPacking.Text)
            cmdSelect.Enabled = True
        End If
    End If
End Sub

Private Sub DisplayPackingList(MyPL As String)
    Dim rsPack As New ADODB.Recordset
    Dim Qry As String
    
    '***Access
    'Qry = "SELECT SPRECEIVE.* From SPRECEIVE WHERE (((SPRECEIVE.PACKINGNO)='" & MyPL & "'));"
    
    'mySQL
    Qry = "SELECT SPRECEIVE.* From SPRECEIVE WHERE SPRECEIVE.PACKINGNO='" & MyPL & "';"
    
    With rsPack
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            txtSupplier.Text = IIf(IsNull(!Supplier), "", !Supplier)
            txtShipper.Text = IIf(IsNull(!Shipper), "", !Shipper)
            txtDateRec.Text = IIf(IsNull(!DATEREC), "", Format(!DATEREC, "mm/dd/yyyy"))
            txtWayBill.Text = IIf(IsNull(!WayBillno), "", !WayBillno)
            txtPackingID.Text = !SPRRNO
        Else
            txtSupplier.Text = ""
            txtShipper.Text = ""
            txtDateRec.Text = ""
            txtWayBill.Text = ""
            txtPackingID.Text = ""
        End If
    End With
End Sub

Private Sub txtPrepBy_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtAprvBy.SetFocus
End Sub


Private Sub txtRem_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        flxParts.SetFocus
        If ActRow < flxParts.Rows - 1 Then flxParts.Row = ActRow + 1
        flxParts.Col = 1
    End If
End Sub

Private Sub txtRem_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        flxParts.TextMatrix(ActRow, 8) = txtRem.Tag
        txtRem.Tag = "skiplostfocus"
        txtRem.Visible = False
        flxParts.SetFocus
        flxParts.Col = ActCol
    End If
End Sub

Private Sub txtRem_LostFocus()
    If txtRem.Tag = "skiplostfocus" Then
        txtRem.Tag = ""
        Exit Sub
    End If
    flxParts.TextMatrix(ActRow, 8) = txtRem.Text
    txtRem.Visible = False
End Sub

Private Sub txtRetNo_GotFocus()
    If txtRetNo.Text <> "" Then
        txtRetNo.SelStart = 0
        txtRetNo.SelLength = Len(txtRetNo.Text)
    End If
End Sub

Private Sub txtRetNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtDateRet.SetFocus
End Sub

Private Sub txtRetNo_LostFocus()
    
    If txtRetNo.Text <> "" Then
        If AddNew Or (Not AddNew And txtRetNo.Text <> txtPuRtrnID.Tag) Then
            '***Access
            'If CtrlNumExist("SELECT PURCHASERET.PURTRNO From PURCHASERET WHERE (((PURCHASERET.PURTRNO)='" & txtRetNo.Text & "'));") Then
            'mySQL
            If CtrlNumExist("SELECT PURCHASERET.PURTRNO From PURCHASERET WHERE PURCHASERET.PURTRNO='" & txtRetNo.Text & "';") Then
                MsgBox "Purchase Return No. already exist!", vbExclamation + vbOKOnly, "Purchase Return"
            End If
        End If
    End If
End Sub

Private Sub txtShipper_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtRetNo.SetFocus
End Sub

Private Sub txtSupplier_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtShipper.SetFocus
End Sub

Private Sub txtWayBill_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSelect.SetFocus
End Sub
