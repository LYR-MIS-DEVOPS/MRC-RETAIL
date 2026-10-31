VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form EMPuser 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User Profile"
   ClientHeight    =   7095
   ClientLeft      =   45
   ClientTop       =   150
   ClientWidth     =   9105
   Icon            =   "EMPuser.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7095
   ScaleWidth      =   9105
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab tabAccess 
      Height          =   4095
      Left            =   240
      TabIndex        =   27
      Top             =   2880
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   7223
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "Applications"
      TabPicture(0)   =   "EMPuser.frx":3072
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame2"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Main System"
      TabPicture(1)   =   "EMPuser.frx":308E
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "lblFormName"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label3"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label4"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "lblRem"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "TreeView1"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "cboAccess"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "cboAType"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "optFull"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "optView"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "optAll"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "flxFrm"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).ControlCount=   11
      Begin MSFlexGridLib.MSFlexGrid flxFrm 
         Height          =   1815
         Left            =   5160
         TabIndex        =   46
         Top             =   2160
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   3201
         _Version        =   393216
         Rows            =   0
         Cols            =   3
         FixedRows       =   0
         FixedCols       =   0
         BackColorBkg    =   16777215
         GridColor       =   16777215
         GridColorFixed  =   16777215
         GridLines       =   0
         ScrollBars      =   2
         SelectionMode   =   1
      End
      Begin VB.OptionButton optAll 
         Caption         =   "All Forms"
         Height          =   375
         Left            =   7500
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   1800
         Width           =   1165
      End
      Begin VB.OptionButton optView 
         Caption         =   "View Only"
         Height          =   375
         Left            =   6330
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   1800
         Width           =   1165
      End
      Begin VB.OptionButton optFull 
         Caption         =   "Full Access"
         Height          =   375
         Left            =   5160
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   1800
         Width           =   1165
      End
      Begin VB.ComboBox cboAType 
         Height          =   315
         Left            =   6600
         TabIndex        =   40
         Top             =   1200
         Width           =   2055
      End
      Begin VB.ComboBox cboAccess 
         Height          =   315
         Left            =   6600
         Style           =   2  'Dropdown List
         TabIndex        =   39
         Top             =   825
         Width           =   2055
      End
      Begin VB.Frame Frame2 
         Height          =   3615
         Left            =   -74880
         TabIndex        =   28
         Top             =   360
         Width           =   8535
         Begin VB.CheckBox chkRebuiltPO 
            Caption         =   "REBUILT P.O."
            Height          =   255
            Left            =   840
            TabIndex        =   48
            Top             =   2760
            Width           =   3015
         End
         Begin VB.CheckBox chkCashier 
            Caption         =   "CASHIER"
            Enabled         =   0   'False
            Height          =   255
            Left            =   840
            TabIndex        =   35
            Top             =   720
            Width           =   3015
         End
         Begin VB.CheckBox chkDisburser 
            Caption         =   "DISBURSER"
            Enabled         =   0   'False
            Height          =   255
            Left            =   840
            TabIndex        =   34
            Top             =   1050
            Width           =   3015
         End
         Begin VB.CheckBox chkPartsman 
            Caption         =   "PARTSMAN"
            Height          =   255
            Left            =   840
            TabIndex        =   33
            Top             =   1395
            Width           =   3015
         End
         Begin VB.CheckBox chkService 
            Caption         =   "SERVICE CENTER"
            Enabled         =   0   'False
            Height          =   255
            Left            =   840
            TabIndex        =   32
            Top             =   2070
            Width           =   3015
         End
         Begin VB.CheckBox chkWarehouse 
            Caption         =   "WAREHOUSEMAN"
            Height          =   255
            Left            =   840
            TabIndex        =   31
            Top             =   2400
            Width           =   3015
         End
         Begin VB.CheckBox chkReceiver 
            Caption         =   "RECEIVER"
            Height          =   255
            Left            =   840
            TabIndex        =   30
            Top             =   1725
            Width           =   3015
         End
         Begin VB.CheckBox chkMotoR 
            Caption         =   "MOTO-R [MAIN SYSTEM]"
            Height          =   255
            Left            =   840
            TabIndex        =   29
            Top             =   360
            Width           =   3015
         End
      End
      Begin MSComctlLib.TreeView TreeView1 
         Height          =   3615
         Left            =   120
         TabIndex        =   36
         Top             =   360
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   6376
         _Version        =   393217
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   6
         Appearance      =   1
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
      Begin VB.Label lblRem 
         Alignment       =   2  'Center
         Height          =   255
         Left            =   5160
         TabIndex        =   42
         Top             =   1560
         Width           =   3495
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Access Type:"
         Height          =   255
         Left            =   5160
         TabIndex        =   41
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Allow Access:"
         Height          =   255
         Left            =   5160
         TabIndex        =   38
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label lblFormName 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   5160
         TabIndex        =   37
         Top             =   480
         Width           =   3495
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxEmp 
      Height          =   1815
      Left            =   2280
      TabIndex        =   24
      Top             =   7680
      Visible         =   0   'False
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   3201
      _Version        =   393216
      FixedRows       =   0
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
   Begin VB.Frame Frame1 
      Height          =   1935
      Left            =   240
      TabIndex        =   6
      Top             =   840
      Width           =   8775
      Begin VB.TextBox txtUserID 
         Height          =   285
         Left            =   1800
         TabIndex        =   1
         Top             =   560
         Width           =   2055
      End
      Begin VB.TextBox txtEmpID 
         Height          =   285
         Left            =   5040
         TabIndex        =   25
         Top             =   600
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.OptionButton optAccess 
         Caption         =   "Operator"
         Height          =   255
         Index           =   1
         Left            =   7080
         TabIndex        =   23
         Top             =   575
         Width           =   1095
      End
      Begin VB.OptionButton optAccess 
         Caption         =   "Administrator"
         Height          =   255
         Index           =   0
         Left            =   7080
         TabIndex        =   22
         Top             =   210
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.TextBox txtEmployee 
         Height          =   315
         Left            =   1800
         TabIndex        =   0
         Top             =   180
         Width           =   3975
      End
      Begin VB.TextBox txtcPW 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   6360
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   5
         Top             =   1440
         Width           =   2295
      End
      Begin VB.ComboBox cboUType 
         Height          =   315
         ItemData        =   "EMPuser.frx":30AA
         Left            =   1800
         List            =   "EMPuser.frx":30AC
         TabIndex        =   3
         Top             =   1440
         Width           =   2805
      End
      Begin VB.TextBox txtPW 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   6360
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   1020
         Width           =   2295
      End
      Begin VB.TextBox txtUser 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1800
         MaxLength       =   10
         TabIndex        =   2
         Top             =   1020
         Width           =   2775
      End
      Begin VB.Label lblI_fld 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "EMPLOYEE ID"
         ForeColor       =   &H80000007&
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   26
         Top             =   605
         Width           =   1200
      End
      Begin VB.Label Label1 
         Caption         =   "DESIGNATION"
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label lblI_fld 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Confirm Password:"
         ForeColor       =   &H80000007&
         Height          =   195
         Index           =   8
         Left            =   4920
         TabIndex        =   10
         Top             =   1440
         Width           =   1305
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000011&
         Index           =   2
         X1              =   -120
         X2              =   8760
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   1
         X1              =   -120
         X2              =   6960
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Label lblI_fld 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "EMPLOYEE NAME"
         ForeColor       =   &H80000007&
         Height          =   195
         Index           =   4
         Left            =   240
         TabIndex        =   9
         Top             =   240
         Width           =   1380
      End
      Begin VB.Label lblI_fld 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "USER NAME"
         ForeColor       =   &H80000007&
         Height          =   195
         Index           =   6
         Left            =   240
         TabIndex        =   8
         Top             =   1080
         Width           =   1320
      End
      Begin VB.Label lblI_fld 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password:"
         ForeColor       =   &H80000007&
         Height          =   195
         Index           =   7
         Left            =   5250
         TabIndex        =   7
         Top             =   1080
         Width           =   975
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
            Picture         =   "EMPuser.frx":30AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "EMPuser.frx":3F88
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "EMPuser.frx":60C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "EMPuser.frx":6F9C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "EMPuser.frx":7E76
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "EMPuser.frx":8D50
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   661
      Caption         =   "&New User"
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
      Image           =   "EMPuser.frx":AA5A
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   1
      Left            =   1440
      TabIndex        =   16
      Top             =   0
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   661
      Caption         =   "&Edit User"
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
      Image           =   "EMPuser.frx":AFF4
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   3
      Left            =   4320
      TabIndex        =   17
      Top             =   0
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   661
      Caption         =   "E&xit Profile"
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
      Image           =   "EMPuser.frx":B58E
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   4
      Left            =   5760
      TabIndex        =   18
      Top             =   0
      Visible         =   0   'False
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   661
      Caption         =   "&Save User"
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
      Image           =   "EMPuser.frx":BB28
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   5
      Left            =   7200
      TabIndex        =   19
      Top             =   0
      Visible         =   0   'False
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   661
      Caption         =   "&Cancel User"
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
      Image           =   "EMPuser.frx":C0C2
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   2
      Left            =   2880
      TabIndex        =   20
      Top             =   0
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   661
      Caption         =   "&Find User"
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
      Image           =   "EMPuser.frx":C65C
      cBack           =   -2147483633
   End
   Begin MSFlexGridLib.MSFlexGrid flxN 
      Height          =   2775
      Left            =   600
      TabIndex        =   47
      Top             =   7440
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   4895
      _Version        =   393216
      Rows            =   102
      Cols            =   5
      FormatString    =   $"EMPuser.frx":CBF6
   End
   Begin VB.Label lblHeading 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "USER PROFILE"
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
      Left            =   6555
      TabIndex        =   13
      Top             =   540
      Width           =   2415
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
      Left            =   240
      TabIndex        =   11
      Top             =   545
      Width           =   2615
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
      Left            =   270
      TabIndex        =   12
      Top             =   570
      Width           =   2615
   End
   Begin VB.Label lblHeading 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "USER PROFILE"
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
      Left            =   6585
      TabIndex        =   14
      Top             =   570
      Width           =   2415
   End
End
Attribute VB_Name = "EMPuser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim AddNew As Boolean, LockFlx As Boolean
Dim rsEmployee As New ADODB.Recordset
Dim CustSrch As Boolean, CustBln As Boolean, xRow As Integer
Dim nodX As Node


'User defined procedures

Sub LockData(Bln As Boolean)
    Dim i As Integer
    
    txtEmployee.Locked = Bln
    txtUserID.Locked = Bln
    cboUType.Locked = Bln
    txtUser.Locked = Bln
    txtPW.Locked = Bln
    txtcPW.Locked = Bln
    Frame2.Enabled = Not Bln
    optAccess(0).Enabled = Not Bln
    optAccess(1).Enabled = Not Bln
    tabAccess.TabEnabled(1) = Not Bln
    tabAccess.Tab = 0
End Sub

Sub ClearEntries()
    Dim i As Integer
    
    txtEmployee.Text = ""
    txtUserID.Text = ""
    cboUType.Text = ""
    txtUser.Text = ""
    txtPW.Text = ""
    txtcPW.Text = ""
    chkMotoR.Value = 0
    chkCashier.Value = 0
    chkDisburser.Value = 0
    chkPartsman.Value = 0
    chkService.Value = 0
    chkWarehouse.Value = 0
    chkRebuiltPO.Value = 0
    For i = 1 To flxN.Rows - 1
        flxN.TextMatrix(i, 2) = ""
        TreeView1.Nodes(i).ForeColor = vbBlack
    Next i
    tabAccess.Tab = 0
End Sub


Public Sub ShowRecord(EmpId As Integer)
    Dim rsPcUser As New ADODB.Recordset
    Dim rsUAccess As New ADODB.Recordset
    Dim Qry As String
    Dim i As Integer
    'Access
    'Qry = "SELECT [LASTNAME] & ', ' & [FIRSTNAME] & ' ' & Left$([MIDDLENAME],1) & '.' AS CUSTOMER, PCUSER.PCUSERID, PCUSER.USERNAME, PCUSER.USERTYPE, PCUSER.PASSWORD, PCUSER.EMPID,  PCUSER.FRMPERS, PCUSER.FRMPARTS, PCUSER.FRMSRVC, PCUSER.FRMTRANS, PCUSER.FRMREPORTS, PCUSER.FRMTOOLS, PCUSER.DESIGNATION, PCUSER.PARTSMAN, PCUSER.WAREHOUSE, PCUSER.CASHIER, PCUSER.SERVICE, PCUSER.DISBURSER, PCUSER.RECEIVER, PCUSER.MOTOR, PCUSER.REBUILTPO " & _
          "FROM EMPLOYEES INNER JOIN PCUSER ON EMPLOYEES.EMPID = PCUSER.EMPID WHERE (((PCUSER.EMPID)=" & EmpId & "));"
    'mySQL
    Qry = "SELECT concat(EMPLOYEES.LASTNAME,', ',EMPLOYEES.FIRSTNAME,' ',IF(ISNULL(LEFT(EMPLOYEES.MIDDLENAME,1)),'',CONCAT(LEFT(EMPLOYEES.MIDDLENAME,1), '.'))) AS CUSTOMER," & _
          "PCUSER.PCUSERID, PCUSER.USERNAME, PCUSER.USERTYPE, PCUSER.PASSWORD, PCUSER.EMPID,  PCUSER.FRMPERS, PCUSER.FRMPARTS, PCUSER.FRMSRVC, PCUSER.FRMTRANS, PCUSER.FRMREPORTS, PCUSER.FRMTOOLS, PCUSER.DESIGNATION, PCUSER.PARTSMAN, PCUSER.WAREHOUSE, PCUSER.CASHIER, PCUSER.SERVICE, PCUSER.DISBURSER, PCUSER.RECEIVER, PCUSER.MOTOR, PCUSER.REBUILTPO " & _
          "FROM EMPLOYEES INNER JOIN PCUSER ON EMPLOYEES.EMPID = PCUSER.EMPID WHERE PCUSER.EMPID=" & EmpId & ";"
    
    Call ClearEntries
    With rsPcUser
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        .MoveFirst
        CustSrch = True
        txtEmployee.Text = !CUSTOMER
        txtUserID.Text = IIf(IsNull(!PCUSERID), "", !PCUSERID)
        cboUType.Text = !DESIGNATION
        txtUser.Text = !UserName
        txtPW.Text = !Password
        txtcPW.Text = !Password
        txtEmpID.Text = EmpId
        'Access
        'chkCashier.Value = IIf(!CASHIER = True, 1, 0)
        'mySQL
        chkCashier.Value = IIf(!Cashier = 1, 1, 0)
        'Access
        'chkDisburser.Value = IIf(!Disburser = True, 1, 0)
        'mySQL
        chkDisburser.Value = IIf(!Disburser = 1, 1, 0)
        'Access
        'chkPartsman.Value = IIf(!PARTSMAN = True, 1, 0)
        'mySQL
        chkPartsman.Value = IIf(!PARTSMAN = 1, 1, 0)
        'Access
        'chkService.Value = IIf(!Service = True, 1, 0)
        'mySQL
        chkService.Value = IIf(!Service = 1, 1, 0)
        'Access
        'chkWarehouse.Value = IIf(!WAREHOUSE = True, 1, 0)
        'mySQL
        chkWarehouse.Value = IIf(!WAREHOUSE = 1, 1, 0)
        'Access
        'chkReceiver.Value = IIf(!RECEIVER = True, 1, 0)
        'mySQL
        chkReceiver.Value = IIf(!RECEIVER = 1, 1, 0)
        'Access
        'chkRebuiltPO.Value = IIf(!RebuiltPO = True, 1, 0)
        'mySQL
        chkRebuiltPO.Value = IIf(!RebuiltPO = 1, 1, 0)
        'Access
        'If !Motor = True Then
        If !Motor = 1 Then
            chkMotoR.Value = 1
            tabAccess.TabEnabled(1) = False
        End If
        If !UserType = "Administrator" Then
            optAccess(0).Value = True
            Call BtnAdmin(False)
        Else
            optAccess(1).Value = True
            tabAccess.Tab = 1
        End If
        .Close
    End With
    'Display User Access
    'Access
    'Qry = "SELECT USERACCESS.* From USERACCESS WHERE (((USERACCESS.EMPID)=" & EmpId & "));"
    'mySQl
    
    Qry = "SELECT USERACCESS.* From USERACCESS WHERE USERACCESS.EMPID=" & EmpId & ";"
    
    With rsUAccess
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                flxN.TextMatrix(!MenuIndx, 2) = !AccessType
                .MoveNext
            Loop
        End If
        .Close
    End With
    optAll.Value = True
    Call optAll_Click
    For i = 1 To flxN.Rows - 1
        If flxN.TextMatrix(i, 2) = "FULL ACCESS" Then TreeView1.Nodes(i).ForeColor = vbRed
        If flxN.TextMatrix(i, 2) = "VIEW ONLY" Then TreeView1.Nodes(i).ForeColor = vbGreen
    Next i
    tabAccess.Tab = 0
End Sub

Sub SaveUser()
    Dim rsMyUser As New ADODB.Recordset
    Dim Qry As String
    Dim i As Integer
    '****ACCESS
    'If AddNew Then
    '    rsMyUser.Open "PcUser", LYRcon, adOpenKeyset, adLockOptimistic
    '    rsMyUser.AddNew
    '    rsMyUser!EmpId = txtEmpID.Text
    'Else
    '    Qry = "SELECT PCUSER.* From PCUSER WHERE (((PCUSER.EMPID)=" & txtEmpID.Text & "));"
    '    rsMyUser.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    '    rsMyUser.MoveFirst
    'End If
    'If optAccess(0).Value = True Then
    '    rsMyUser!UserType = "Administrator"
    'Else
    '    rsMyUser!UserType = "Operator"
    'End If
    'rsMyUser!PCUSERID = txtUserID.Text
    'rsMyUser!UserName = txtUser.Text
    'rsMyUser!DESIGNATION = cboUType.Text
    'rsMyUser!Password = txtPW.Text
    'rsMyUser!UPDATEDBY = UserID
    'rsMyUser!LASTUPDATE = Date
    'Access
    'rsMyUser!CASHIER = IIf(chkCashier.Value = 1, True, False)
    'mySQL
    'rsMyUser!CASHIER = IIf(chkCashier.Value = 1, 1, 0)
    'Access
    'rsMyUser!Disburser = IIf(chkDisburser.Value = 1, True, False)
    'mySQL
    'rsMyUser!Disburser = IIf(chkDisburser.Value = 1, 1, 0)
    'Access
    'rsMyUser!PARTSMAN = IIf(chkPartsman.Value = 1, True, False)
    'mySQL
    'rsMyUser!PARTSMAN = IIf(chkPartsman.Value = 1, 1, 0)
    'Access
    'rsMyUser!Service = IIf(chkService.Value = 1, True, False)
    'mySQL
    'rsMyUser!Service = IIf(chkService.Value = 1, 1, 0)
    'Access
    'rsMyUser!WAREHOUSE = IIf(chkWarehouse.Value = 1, True, False)
    'mySQL
    'rsMyUser!WAREHOUSE = IIf(chkWarehouse.Value = 1, 1, 0)
    'Access
    'rsMyUser!RECEIVER = IIf(chkReceiver.Value = 1, True, False)
    'mySQL
    'rsMyUser!RECEIVER = IIf(chkReceiver.Value = 1, 1, 0)
    'Access
    'rsMyUser!Motor = IIf(chkMotoR.Value = 1, True, False)
    'mySQL
    'rsMyUser!Motor = IIf(chkMotoR.Value = 1, 1, 0)
    'Access
    'rsMyUser!RebuiltPO = IIf(chkRebuiltPO.Value = 1, True, False)
    'mySQL
    'rsMyUser!RebuiltPO = IIf(chkRebuiltPO.Value = 1, 1, 0)
    
    'rsMyUser.Update
    'If AddNew Then txtEmpID.Text = rsMyUser!EmpId
    'rsMyUser.Close
    
    'Update table Employees field PCuser
    'LYRcon.Execute "UPDATE EMPLOYEES SET EMPLOYEES.PCUSER = '" & cboUType.Text & "' WHERE (((EMPLOYEES.EMPID)=" & txtEmpID & "));"
    
    'If Not AddNew Then
    '    LYRcon.Execute "DELETE USERACCESS.*, USERACCESS.EMPID From USERACCESS WHERE (((USERACCESS.EMPID)=" & txtEmpID & "));"
    'End If
    'With flxN
    '    For i = 1 To .Rows - 1
    '        If .TextMatrix(i, 2) <> "" Then
    '            LYRcon.Execute "INSERT INTO USERACCESS ( EMPID, MENUINDX, MENUID, ACCESSTYPE )" & _
    '                "VALUES (" & txtEmpID & ", " & i & ", " & .TextMatrix(i, 1) & ", '" & .TextMatrix(i, 2) & "');"
    '        End If
    '    Next i
    'End With
    'mySQL Statement
    
    If optAccess(0).Value = True Then
        myUserType = "Administrator"
    Else
        myUserType = "Operator"
    End If
     If AddNew Then
        LYRcon.Execute "INSERT INTO pcuser(EMPID, UserType, PCUSERID, USERNAME, DESIGNATION, PASSWORD, UPDATEDBY, " & _
            "LASTUPDATE, CASHIER, DISBURSER, PARTSMAN, SERVICE, WAREHOUSE, RECEIVER, MOTOR, REBUILTPO)" & _
            "VALUES('" & txtEmpID.Text & "', '" & myUserType & "', '" & txtUserID.Text & "', '" & txtUser.Text & "', '" & cboUType.Text & "', " & _
            "'" & txtPW.Text & "', " & UserID & ", '" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "', " & _
            "" & IIf(chkCashier.Value = 1, 1, 0) & ", " & IIf(chkDisburser.Value = 1, 1, 0) & ", " & _
            "" & IIf(chkPartsman.Value = 1, 1, 0) & ", " & IIf(chkService.Value = 1, 1, 0) & ", " & _
            "" & IIf(chkWarehouse.Value = 1, 1, 0) & ", " & IIf(chkReceiver.Value = 1, 1, 0) & ", " & _
            "" & IIf(chkMotoR.Value = 1, 1, 0) & ", " & IIf(chkRebuiltPO.Value = 1, 1, 0) & ");"
    '    Display EmpID
         Qry = "SELECT pcuser.EmpId From pcuser Where pcuser.PCUSERID = '" & txtUserID.Text & "'"
         myEmpID = GetFieldValue(Qry)
         rsMyUser.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
         txtEmpID.Text = myEmpID
     Else
        LYRcon.Execute "Update pcuser SET EMPID = '" & txtEmpID.Text & "', UserType = '" & myUserType & "', PCUSERID = '" & txtUserID.Text & "', " & _
            "USERNAME = '" & txtUser.Text & "', DESIGNATION = '" & cboUType.Text & "', PASSWORD = '" & txtPW.Text & "', " & _
            "UPDATEDBY = " & UserID & ", LASTUPDATE = '" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "', CASHIER = " & IIf(chkCashier.Value = 1, 1, 0) & ", " & _
            "DISBURSER = " & IIf(chkDisburser.Value = 1, 1, 0) & ", PARTSMAN = " & IIf(chkPartsman.Value = 1, 1, 0) & ", " & _
            "SERVICE = " & IIf(chkService.Value = 1, 1, 0) & ", WAREHOUSE = " & IIf(chkWarehouse.Value = 1, 1, 0) & ", " & _
            "RECEIVER = " & IIf(chkReceiver.Value = 1, 1, 0) & ", MOTOR = " & IIf(chkMotoR.Value = 1, 1, 0) & ", " & _
            "REBUILTPO = " & IIf(chkRebuiltPO.Value = 1, 1, 0) & " Where pcuser.EMPID = " & txtEmpID.Text & ";"
        
        LYRcon.Execute "DELETE From USERACCESS WHERE USERACCESS.EMPID=" & txtEmpID & ";"
    End If
    
       
    'Update table Employees field PCuser
    LYRcon.Execute "UPDATE EMPLOYEES SET EMPLOYEES.PCUSER = '" & cboUType.Text & "' WHERE EMPLOYEES.EMPID=" & txtEmpID & ";"
    With flxN
        For i = 1 To .Rows - 1
            If .TextMatrix(i, 2) <> "" Then
                LYRcon.Execute "INSERT INTO USERACCESS ( EMPID, MENUINDX, MENUID, ACCESSTYPE )" & _
                    "VALUES (" & txtEmpID & ", " & i & ", " & .TextMatrix(i, 1) & ", '" & .TextMatrix(i, 2) & "');"
            End If
        Next i
    End With
End Sub

Function ValidateEntries() As String
    Dim myMsg As String, vQry As String
    Dim i As Integer, arrCnt As Integer
    
    myMsg = ""

    If txtUser.Text = "" Then myMsg = "No User Name defined."
    If txtEmployee.Text = "" Then myMsg = myMsg & vbCrLf & "Incomplete entry on field Employee"
    If txtUserID.Text = "" Then
        myMsg = myMsg & vbCrLf & "Incomplete entry on field User ID"
    ElseIf txtUserID.Text <> "" And AddNew Then
        '****Access
        'vQry = "SELECT PCUSER.PCUSERID From PCUSER WHERE (((PCUSER.PCUSERID)='" & txtUserID.Text & "'));"
        'mySQL
        vQry = "SELECT PCUSER.PCUSERID From PCUSER WHERE PCUSER.PCUSERID='" & txtUserID.Text & "';"
        
        'If CtrlNumExist(vQry) Then myMsg = myMsg & vbCrLf & "User ID already Exist"
    End If
    If txtPW.Text = "" Or txtcPW.Text = "" Then
        myMsg = myMsg & vbCrLf & "No password defined."
    ElseIf txtPW.Text <> txtcPW.Text Then
        myMsg = myMsg & vbCrLf & "Incorrect Password."
    End If
    
    ValidateEntries = myMsg
    
End Function

Private Sub ClearAccessSubMenu(MyR As Integer)
    Dim i As Integer
        
    With flxN
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, 3)) = MyR Then
                .TextMatrix(i, 2) = ""
                TreeView1.Nodes(i).ForeColor = vbBlack
            End If
        Next i
    End With
End Sub


Private Sub cboAccess_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If cboAccess.Text = "YES" And flxN.TextMatrix(xRow, 4) <> "" Then
            SendKeys "{TAB}"
        Else
            TreeView1.SetFocus
        End If
    End If
End Sub

Private Sub cboAccess_LostFocus()
    
    If cboAccess.Text = "NO" Then
        flxN.TextMatrix(xRow, 2) = ""
        TreeView1.Nodes(xRow).ForeColor = vbBlack
        cboAType.Text = ""
        cboAType.Enabled = False
        Call ClearAccessSubMenu(xRow)
        If optAll.Value = True Then
            Call optAll_Click
        ElseIf optFull.Value = True Then
            Call optFull_Click
        ElseIf optView.Value = True Then
            Call optView_Click
        End If
    Else
        If flxN.TextMatrix(xRow, 4) = "" Then
            flxN.TextMatrix(xRow, 2) = "TOP LEVEL MENU"
            cboAType.Enabled = False
            lblRem.Caption = "Top Level Menu"
        ElseIf flxN.TextMatrix(xRow, 4) = "FALSE" Then
            cboAType.Text = "VIEW ONLY"
            cboAType.Enabled = False
            lblRem.Caption = "Display Type Application"
            If optAll.Value = True Then
                Call optAll_Click
            ElseIf optFull.Value = True Then
                Call optFull_Click
            ElseIf optView.Value = True Then
                Call optView_Click
            End If
        Else
            cboAType.Enabled = True
            lblRem.Caption = "Data Processing Type Application"
        End If
    End If
End Sub

Private Sub cboAType_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then TreeView1.SetFocus
End Sub

Private Sub cboAType_LostFocus()
    
    If cboAType.Text <> "FULL ACCESS" And cboAType.Text <> "VIEW ONLY" Then
        If cboAccess.Text = "YES" And flxN.TextMatrix(xRow, 4) <> "" Then
            MsgBox "Invalid Access type... Please select from the list!", vbExclamation + vbOKOnly, "Access Type"
            cboAType.SetFocus
        End If
        Exit Sub
    End If
    flxN.TextMatrix(xRow, 2) = cboAType.Text
    If cboAType.Text = "VIEW ONLY" Then
        TreeView1.Nodes(xRow).ForeColor = vbGreen
    Else
        TreeView1.Nodes(xRow).ForeColor = vbRed
    End If
    If optAll.Value = True Then
        Call optAll_Click
    ElseIf optFull.Value = True Then
        Call optFull_Click
    ElseIf optView.Value = True Then
        Call optView_Click
    End If
End Sub

Private Sub cboUType_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub chkMotoR_Click()
    If chkMotoR.Value = 1 Then 'true
        tabAccess.TabEnabled(1) = True
    Else
        tabAccess.TabEnabled(1) = False
    End If
End Sub

Private Sub cmdTrans_Click(Index As Integer)
    Dim Ans As Integer, myMsg As String
    Dim Qry As String, MyAcc As String
    
    Select Case Index
        Case 0  'New
            '**Access
            'Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE (((USERACCESS.EMPID)=" & UserID & ") AND ((USERACCESS.MENUINDX)=9));"
            'mySQL
            Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE USERACCESS.EMPID=" & UserID & " AND USERACCESS.MENUINDX=9;"
            
            MyAcc = GetFieldValue(Qry)
            If MyAcc <> "FULL ACCESS" Then
                MsgBox "You don't have full access to this transaction." & vbCrLf & "Please contact your System Administrator", vbExclamation + vbOKOnly, "Access Denied"
                Exit Sub
            End If
            Call ClearEntries
            Call LockData(False)
            Call BtnAdmin(True)
            AddNew = True
        Case 1  'Edit
            '**Access
            'Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE (((USERACCESS.EMPID)=" & UserID & ") AND ((USERACCESS.MENUINDX)=9));"
            'mySQL
            Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE USERACCESS.EMPID=" & UserID & " AND USERACCESS.MENUINDX=9;"
            
            MyAcc = GetFieldValue(Qry)
            If MyAcc <> "FULL ACCESS" Then
                MsgBox "You don't have full access to this transaction." & vbCrLf & "Please contact your System Administrator", vbExclamation + vbOKOnly, "Access Denied"
                Exit Sub
            End If
            If txtEmpID.Text = "" Then
                MsgBox "No available employee to edit.", vbInformation + vbOKOnly, "Edit"
                Exit Sub
            End If
            LockFlx = False
            AddNew = False
            Call LockData(False)
            Call BtnAdmin(True)
        Case 2  'Find
            Set frmSearch.FrmSource = EMPuser
            '**Access
            'Qry = "SELECT PCUSER.EMPID, [LASTNAME] & ', ' & [FIRSTNAME] & ' ' & Left$([MIDDLENAME],1) & '.' AS EMPLOYEE " & _
                  "FROM PCUSER INNER JOIN EMPLOYEES ON PCUSER.EMPID = EMPLOYEES.EMPID " & _
                  "ORDER BY [LASTNAME] & ', ' & [FIRSTNAME] & ' ' & Left$([MIDDLENAME],1) & '.';"

            
            'mySQL
            Qry = "SELECT PCUSER.EMPID, concat(EMPLOYEES.LASTNAME,', ',EMPLOYEES.FIRSTNAME,' ',IF(ISNULL(LEFT(EMPLOYEES.MIDDLENAME,1)),'',CONCAT(LEFT(EMPLOYEES.MIDDLENAME,1), '.'))) AS EMPLOYEE " & _
                "From PCUSER INNER JOIN EMPLOYEES ON PCUSER.EMPID = EMPLOYEES.EMPID" & _
                " ORDER BY concat(EMPLOYEES.LASTNAME,', ',EMPLOYEES.FIRSTNAME,' ',IF(ISNULL(LEFT(EMPLOYEES.MIDDLENAME,1)),'',CONCAT(LEFT(EMPLOYEES.MIDDLENAME,1), '.')));"
                
                    
                
            frmSearch.QrySearch = Qry
            frmSearch.lblCap.Caption = "Search for User Name"
            frmSearch.Show vbModal
            
        Case 3  'Exit
            Unload Me
        Case 4  'Save
            myMsg = ValidateEntries
            If myMsg <> "" Then
                MsgBox myMsg, vbExclamation + vbOKOnly, "Validation Error"
                Exit Sub
            End If
            Ans = MsgBox("Do you want to Save User Profile?", vbYesNo + vbQuestion, "Save")
            If Ans = vbNo Then Exit Sub
            Call SaveUser
            If Val(txtEmpID.Text) = UserID Then Call UserAccess
            MsgBox "User Profile successfully saved!"
            Call LockData(True)
            Call BtnAdmin(False)
            LockFlx = True
        Case 5  'Cancel
            Ans = MsgBox("Do you want to Cancel User Profile?", vbYesNo + vbQuestion, "Save")
            If Ans = vbNo Then Exit Sub
            Call LockData(True)
            If AddNew Then
                Call ClearEntries
            Else
                Call ShowRecord(Val(txtEmpID.Text))
            End If
            Call BtnAdmin(False)
            LockFlx = True
    End Select
End Sub

Private Sub flxEmp_Click()
    txtEmployee.Text = flxEmp.TextMatrix(flxEmp.Row, 1)
End Sub

Private Sub flxEmp_EnterCell()
    CustSrch = True
    txtEmployee.Text = flxEmp.TextMatrix(flxEmp.Row, 1)
End Sub

Private Sub flxEmp_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        flxEmp.Visible = False
        txtEmployee.SetFocus
    End If
End Sub

Private Sub flxEmp_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        flxEmp.Visible = False
        txtUser.SetFocus
    End If
End Sub

Private Sub flxFrm_Click()
    If Val(flxFrm.TextMatrix(flxFrm.Row, 2)) > 0 Then
        TreeView1.SelectedItem = TreeView1.Nodes(Val(flxFrm.TextMatrix(flxFrm.Row, 2)))
    End If
End Sub

Private Sub Form_Load()
    Dim i As Integer
    

    cboUType.AddItem "Cashier"
    cboUType.AddItem "Checker"
    cboUType.AddItem "Disburser"
    cboUType.AddItem "Encoder"
    cboUType.AddItem "Manager"
    cboUType.AddItem "Partsman"
    cboUType.AddItem "Receiver"
    cboUType.AddItem "Service Advisor"
    cboUType.AddItem "System Admin."
    cboUType.AddItem "Warehouseman"
    
    TreeView1.Nodes.Clear
   Set nodX = TreeView1.Nodes.Add(, , "R", "Moto-R Republic Corporation")
   Set nodX = TreeView1.Nodes.Add("R", tvwChild, "M1", "Personnel Information")
   Set nodX = TreeView1.Nodes.Add("R", tvwChild, "M2", "Spare Parts")
   Set nodX = TreeView1.Nodes.Add("R", tvwChild, "M3", "Service Center")
   Set nodX = TreeView1.Nodes.Add("R", tvwChild, "M4", "Transactions")
   Set nodX = TreeView1.Nodes.Add("R", tvwChild, "M5", "Reports")
   Set nodX = TreeView1.Nodes.Add("R", tvwChild, "M6", "Tools")
   'Add child to Menu 1
   Set nodX = TreeView1.Nodes.Add("M1", tvwChild, "M1A", "Employees Profile")
   Set nodX = TreeView1.Nodes.Add("M1", tvwChild, "M1B", "User Profile")
   Set nodX = TreeView1.Nodes.Add("M1", tvwChild, "M1C", "On-Account Customer")
   Set nodX = TreeView1.Nodes.Add("M1", tvwChild, "M1D", "Suppliers")
   Set nodX = TreeView1.Nodes.Add("M1", tvwChild, "M1E", "Shipper")
    'Add child to Menu 2
    Set nodX = TreeView1.Nodes.Add("M2", tvwChild, "M2A", "Ordering")
    Set nodX = TreeView1.Nodes.Add("M2", tvwChild, "M2B", "Receiving")
    Set nodX = TreeView1.Nodes.Add("M2", tvwChild, "M2C", "Sales")
    Set nodX = TreeView1.Nodes.Add("M2", tvwChild, "M2D", "Stock Monitoring")
    Set nodX = TreeView1.Nodes.Add("M2", tvwChild, "M2E", "Updates")
    'Add child to Menu 2-A [Ordering]
    Set nodX = TreeView1.Nodes.Add("M2A", tvwChild, "M2AI", "Purchase Order")
    Set nodX = TreeView1.Nodes.Add("M2A", tvwChild, "M2AII", "Summary of P.O.")
    Set nodX = TreeView1.Nodes.Add("M2A", tvwChild, "M2AIII", "Back Order Summary")
    'Add child to Menu 2-B [Receiving]
    Set nodX = TreeView1.Nodes.Add("M2B", tvwChild, "M2BI", "Receiving Form")
    Set nodX = TreeView1.Nodes.Add("M2B", tvwChild, "M2BII", "Purchase Return")
    Set nodX = TreeView1.Nodes.Add("M2B", tvwChild, "M2BIII", "Receiving Summary")
    Set nodX = TreeView1.Nodes.Add("M2B", tvwChild, "M2BIV", "Purchase Return Summary")
    'Add child to Menu 2-C [Sales]
    Set nodX = TreeView1.Nodes.Add("M2C", tvwChild, "M2CI", "Sales Order")
    Set nodX = TreeView1.Nodes.Add("M2C", tvwChild, "M2CII", "Warehouse")
    Set nodX = TreeView1.Nodes.Add("M2C", tvwChild, "M2CIII", "Sales Return")
    Set nodX = TreeView1.Nodes.Add("M2C", tvwChild, "M2CIV", "Sales Order Summary")
    Set nodX = TreeView1.Nodes.Add("M2C", tvwChild, "M2CV", "Sales Return Summary")
    'Add child to Menu 3-D [Stock Monitory]
    Set nodX = TreeView1.Nodes.Add("M2D", tvwChild, "M2DI", "Cardex")
    Set nodX = TreeView1.Nodes.Add("M2D", tvwChild, "M2DII", "Cardex Red")
    Set nodX = TreeView1.Nodes.Add("M2D", tvwChild, "M2DIII", "Merchandise Inventory")
    Set nodX = TreeView1.Nodes.Add("M2D", tvwChild, "M2DIV", "Monthly Inventory Summary")
    Set nodX = TreeView1.Nodes.Add("M2D", tvwChild, "M2DV", "Spare Parts Price List")
    'Add child to Menu 4-E [Updates]
    Set nodX = TreeView1.Nodes.Add("M2E", tvwChild, "M2EI", "Update Location")
    Set nodX = TreeView1.Nodes.Add("M2E", tvwChild, "M2EII", "Update Spare Parts")
    Set nodX = TreeView1.Nodes.Add("M2E", tvwChild, "M2EIII", "Price Adjustment")
    'Add child to Menu 3
    Set nodX = TreeView1.Nodes.Add("M3", tvwChild, "M3A", "Job Order")
    Set nodX = TreeView1.Nodes.Add("M3", tvwChild, "M3B", "Requisition Summary")
    Set nodX = TreeView1.Nodes.Add("M3", tvwChild, "M3C", "Purchase Order Summary")
    Set nodX = TreeView1.Nodes.Add("M3", tvwChild, "M3D", "Customer Service Record")
    Set nodX = TreeView1.Nodes.Add("M3", tvwChild, "M3E", "Mechanic Service Income")
    Set nodX = TreeView1.Nodes.Add("M3", tvwChild, "M3F", "Service Income Summary")
    'Add child to Menu 4
    Set nodX = TreeView1.Nodes.Add("M4", tvwChild, "M4A", "Sales Invoice")
    Set nodX = TreeView1.Nodes.Add("M4", tvwChild, "M4B", "Official Receipt")
    Set nodX = TreeView1.Nodes.Add("M4", tvwChild, "M4C", "Provisional Receipt")
    Set nodX = TreeView1.Nodes.Add("M4", tvwChild, "M4D", "Adjustment")
    Set nodX = TreeView1.Nodes.Add("M4", tvwChild, "M4E", "Disbursement")
    Set nodX = TreeView1.Nodes.Add("M4", tvwChild, "M4F", "Cash Flow")
    Set nodX = TreeView1.Nodes.Add("M4", tvwChild, "M4G", "Accounts Receivable")
    'Add child to Menu 4-D [Adjustments]
    Set nodX = TreeView1.Nodes.Add("M4D", tvwChild, "M4DI", "A/R DMCM")
    Set nodX = TreeView1.Nodes.Add("M4D", tvwChild, "M4DII", "Inventory Adjustment")
    Set nodX = TreeView1.Nodes.Add("M4D", tvwChild, "M4DIII", "Clerical Adjustment")
    Set nodX = TreeView1.Nodes.Add("M4D", tvwChild, "M4DIV", "A/R DMCM Summary")
    Set nodX = TreeView1.Nodes.Add("M4D", tvwChild, "M4DV", "Inv. Adj. Summary")
    Set nodX = TreeView1.Nodes.Add("M4D", tvwChild, "M4DVI", "Clerical Adj. Summary")
    'Add child to Menu 4-E [Disbursement]
    Set nodX = TreeView1.Nodes.Add("M4E", tvwChild, "M4EI", "Petty Cash Voucher")
    Set nodX = TreeView1.Nodes.Add("M4E", tvwChild, "M4EII", "Cash Voucher")
    Set nodX = TreeView1.Nodes.Add("M4E", tvwChild, "M4EIII", "PCV Summary")
    Set nodX = TreeView1.Nodes.Add("M4E", tvwChild, "M4EIV", "Cash Voucher Summary")
    'Add child to menu 4-F [Cash Flow]
    Set nodX = TreeView1.Nodes.Add("M4F", tvwChild, "M4FI", "Cashier's Cash Flow")
    Set nodX = TreeView1.Nodes.Add("M4F", tvwChild, "M4FII", "Manager's Cash Flow")
    Set nodX = TreeView1.Nodes.Add("M4F", tvwChild, "M4FV", "Checks Summary")
    'Add child to menu 4-G [Accounts Receivables]
    Set nodX = TreeView1.Nodes.Add("M4G", tvwChild, "M4GI", "Aging of Account Receivables")
    Set nodX = TreeView1.Nodes.Add("M4G", tvwChild, "M4GII", "Statement of Accounts")
    Set nodX = TreeView1.Nodes.Add("M4G", tvwChild, "M4GIII", "Ledger")
    Set nodX = TreeView1.Nodes.Add("M4G", tvwChild, "M4GIV", "Down Payment Summary")
    Set nodX = TreeView1.Nodes.Add("M4G", tvwChild, "M4GV", "Cash QckSvc w/ Credit Invoice")
    'Add child to menu 5
    Set nodX = TreeView1.Nodes.Add("M5", tvwChild, "M5A", "Daily Transactions")
    Set nodX = TreeView1.Nodes.Add("M5", tvwChild, "M5B", "Spare Parts Inventory")
    Set nodX = TreeView1.Nodes.Add("M5", tvwChild, "M5C", "Service Center")
    Set nodX = TreeView1.Nodes.Add("M5", tvwChild, "M5D", "Monthly Reports")
    'Add child to menu 5-A [Daily Transactions]
    Set nodX = TreeView1.Nodes.Add("M5A", tvwChild, "M5AI", "Daily Sales Report")
    Set nodX = TreeView1.Nodes.Add("M5A", tvwChild, "M5AII", "Daily Cash Collection")
    Set nodX = TreeView1.Nodes.Add("M5A", tvwChild, "M5AIII", "Daily Disbursement Report")
    'Add child to menu 5-B [Spare Parts Inventory]
    Set nodX = TreeView1.Nodes.Add("M5B", tvwChild, "M5BI", "Merchandise Inventory")
    Set nodX = TreeView1.Nodes.Add("M5B", tvwChild, "M5BII", "Detailed Receiving Report")
    Set nodX = TreeView1.Nodes.Add("M5B", tvwChild, "M5BIII", "Detailed Sales Report At Selling Price")
    Set nodX = TreeView1.Nodes.Add("M5B", tvwChild, "M5BIV", "Summary of Inv. Adjustment")
    Set nodX = TreeView1.Nodes.Add("M5B", tvwChild, "M5BV", "Summary of Sales Return")
    Set nodX = TreeView1.Nodes.Add("M5B", tvwChild, "M5BVI", "Summary of Purchase Return")
    'Add child to menu 5-C [Service Center]
    Set nodX = TreeView1.Nodes.Add("M5C", tvwChild, "M5CI", "Daily Service Income")
    Set nodX = TreeView1.Nodes.Add("M5C", tvwChild, "M5CII", "Mechanic's Service Income")
    Set nodX = TreeView1.Nodes.Add("M5C", tvwChild, "M5CIII", "Service Income Summary")
    'Add child to menu 5-D [Monthly Reports]
    Set nodX = TreeView1.Nodes.Add("M5D", tvwChild, "M5DI", "Monthly Sales Report")
    Set nodX = TreeView1.Nodes.Add("M5D", tvwChild, "M5DII", "Monthly Collection Report")
    Set nodX = TreeView1.Nodes.Add("M5D", tvwChild, "M5DIII", "Monthly Disbursement Report")
    Set nodX = TreeView1.Nodes.Add("M5D", tvwChild, "M5DIV", "Monthly Service Report")
    Set nodX = TreeView1.Nodes.Add("M5D", tvwChild, "M5DV", "Monthly Inventory Summary")
    'Add child to menu 6
    Set nodX = TreeView1.Nodes.Add("M6", tvwChild, "M6A", "Backup")
    Set nodX = TreeView1.Nodes.Add("M6", tvwChild, "M6B", "Restore")
    Set nodX = TreeView1.Nodes.Add("M6", tvwChild, "M6C", "Audit Tools")
    Set nodX = TreeView1.Nodes.Add("M6", tvwChild, "M6D", "Settings")
    'Add child to menu 6-C [Audit Tools]
    Set nodX = TreeView1.Nodes.Add("M6C", tvwChild, "M6CI", "Inventory Audit")
    'Add child to menu 6-D [Settings]
    Set nodX = TreeView1.Nodes.Add("M6D", tvwChild, "M6DI", "Printing Preference")
    Set nodX = TreeView1.Nodes.Add("M6D", tvwChild, "M6DII", "Options")
    'Late comers
    Set nodX = TreeView1.Nodes.Add("M2A", tvwChild, "M2AIV", "Top Selling Spare Parts")
    Set nodX = TreeView1.Nodes.Add("M4F", tvwChild, "M4FIII", "Advances for Liquidation")
    Set nodX = TreeView1.Nodes.Add("M4F", tvwChild, "M4FIV", "Cash Return")
    Set nodX = TreeView1.Nodes.Add("M5B", tvwChild, "M5BVII", "Detailed Sales Report At Cost Price")
    Set nodX = TreeView1.Nodes.Add("M5D", tvwChild, "M5DVI", "Monthly Sales Report w/Adjustment")
    
    
    'Cash Quick Service w/ Credit Invoice
    'Set nodX = TreeView1.Nodes.Add("M4G", tvwChild, "M4GVI", "Cash QckSvc w/ Credit Invoice")
   nodX.EnsureVisible
   'TreeView1.Style = tvwTreelinesText ' Style 4.
   TreeView1.BorderStyle = vbFixedSingle
        
        
'    Call ShowRecord(UserID)
    LockFlx = True
    MotoRep.sBar.Panels(6).Text = "System User's Profile"
    
    'Load Employee's records
    '**Access
    'Qry = "SELECT UCase$([LASTNAME] & ', ' & [FIRSTNAME] & ' ' & Left$([MIDDLENAME],1) & '.') AS EMPLOYEE, EMPLOYEES.EMPID From EMPLOYEES " & _
          "WHERE (((EMPLOYEES.DEPARTMENT)<>'Default') AND ((EMPLOYEES.INACTIVE)=False) AND ((EMPLOYEES.PCUSER) Is Null)) " & _
          "ORDER BY UCase$([LASTNAME] & ', ' & [FIRSTNAME] & ' ' & Left$([MIDDLENAME],1) & '.');"
    'mySQL
    Qry = "SELECT concat(EMPLOYEES.LASTNAME,', ',EMPLOYEES.FIRSTNAME,' ',IF(ISNULL(LEFT(EMPLOYEES.MIDDLENAME,1)),'',CONCAT(LEFT(EMPLOYEES.MIDDLENAME,1), '.'))) AS EMPLOYEE, EMPLOYEES.EMPID From EMPLOYEES " & _
          "WHERE EMPLOYEES.DEPARTMENT<>'Default' AND EMPLOYEES.INACTIVE=0 AND (EMPLOYEES.PCUSER) Is Null " & _
          "ORDER BY concat(EMPLOYEES.LASTNAME,', ',EMPLOYEES.FIRSTNAME,' ',IF(LEFT(EMPLOYEES.MIDDLENAME,1)='','',CONCAT(LEFT(EMPLOYEES.MIDDLENAME,1), '.')));"
    
    rsEmployee.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    Set flxEmp.DataSource = rsEmployee
    With flxEmp
        .Width = txtEmployee.Width
        .ColWidth(0) = 0
        .ColWidth(1) = txtEmployee.Width - 300
        .ColWidth(2) = 0
    End With
    
    
    'Default Menu values
    With flxN
        .TextMatrix(8, 1) = 1: .TextMatrix(8, 4) = "TRUE"
        .TextMatrix(9, 1) = 2: .TextMatrix(9, 4) = "TRUE"
        .TextMatrix(10, 1) = 4: .TextMatrix(10, 4) = "TRUE"
        .TextMatrix(11, 1) = 6: .TextMatrix(11, 4) = "TRUE"
        .TextMatrix(12, 1) = 7: .TextMatrix(12, 4) = "TRUE"
        .TextMatrix(13, 1) = 1
        .TextMatrix(18, 1) = 1: .TextMatrix(18, 3) = 13: .TextMatrix(18, 4) = "TRUE"
        .TextMatrix(19, 1) = 3: .TextMatrix(19, 3) = 13: .TextMatrix(19, 4) = "TRUE"
        .TextMatrix(20, 1) = 4: .TextMatrix(20, 3) = 13: .TextMatrix(20, 4) = "FALSE"
        .TextMatrix(14, 1) = 2
        .TextMatrix(21, 1) = 1: .TextMatrix(21, 3) = 14: .TextMatrix(21, 4) = "TRUE"
        .TextMatrix(22, 1) = 2: .TextMatrix(22, 3) = 14: .TextMatrix(22, 4) = "TRUE"
        .TextMatrix(23, 1) = 4: .TextMatrix(23, 3) = 14: .TextMatrix(23, 4) = "FALSE"
        .TextMatrix(24, 1) = 5: .TextMatrix(24, 3) = 14: .TextMatrix(24, 4) = "TRUE"
        .TextMatrix(15, 1) = 3
        .TextMatrix(25, 1) = 1: .TextMatrix(25, 3) = 15: .TextMatrix(25, 4) = "TRUE"
        .TextMatrix(26, 1) = 2: .TextMatrix(26, 3) = 15: .TextMatrix(26, 4) = "FALSE"
        .TextMatrix(27, 1) = 3: .TextMatrix(27, 3) = 15: .TextMatrix(27, 4) = "TRUE"
        .TextMatrix(28, 1) = 5: .TextMatrix(28, 3) = 15: .TextMatrix(28, 4) = "FALSE"
        .TextMatrix(29, 1) = 6: .TextMatrix(29, 3) = 15: .TextMatrix(29, 4) = "TRUE"
        .TextMatrix(16, 1) = 4
        .TextMatrix(30, 1) = 1: .TextMatrix(30, 3) = 16: .TextMatrix(30, 4) = "FALSE"
        .TextMatrix(31, 1) = 2: .TextMatrix(31, 3) = 16: .TextMatrix(31, 4) = "FALSE"
        .TextMatrix(32, 1) = 4: .TextMatrix(32, 3) = 16: .TextMatrix(32, 4) = "FALSE"
        .TextMatrix(33, 1) = 5: .TextMatrix(33, 3) = 16: .TextMatrix(33, 4) = "FALSE"
        .TextMatrix(34, 1) = 6: .TextMatrix(34, 3) = 16: .TextMatrix(34, 4) = "FALSE"
        .TextMatrix(17, 1) = 5
        .TextMatrix(35, 1) = 1: .TextMatrix(35, 3) = 17: .TextMatrix(35, 4) = "TRUE"
        .TextMatrix(36, 1) = 2: .TextMatrix(36, 3) = 17: .TextMatrix(36, 4) = "TRUE"
        .TextMatrix(37, 1) = 3: .TextMatrix(37, 3) = 17: .TextMatrix(37, 4) = "TRUE"
        .TextMatrix(38, 1) = 1: .TextMatrix(38, 4) = "TRUE"
        .TextMatrix(39, 1) = 2: .TextMatrix(39, 4) = "TRUE"
        .TextMatrix(40, 1) = 3: .TextMatrix(40, 4) = "TRUE"
        .TextMatrix(41, 1) = 4: .TextMatrix(41, 4) = "TRUE"
        .TextMatrix(42, 1) = 5: .TextMatrix(42, 4) = "FALSE"
        .TextMatrix(43, 1) = 6: .TextMatrix(43, 4) = "TRUE"
        .TextMatrix(44, 1) = 1: .TextMatrix(44, 4) = "TRUE"
        .TextMatrix(45, 1) = 2: .TextMatrix(45, 4) = "TRUE"
        .TextMatrix(46, 1) = 3: .TextMatrix(46, 4) = "TRUE"
        .TextMatrix(47, 1) = 4
        .TextMatrix(51, 1) = 1: .TextMatrix(51, 3) = 47: .TextMatrix(51, 4) = "TRUE"
        .TextMatrix(52, 1) = 2: .TextMatrix(52, 3) = 47: .TextMatrix(52, 4) = "TRUE"
        .TextMatrix(53, 1) = 3: .TextMatrix(53, 3) = 47: .TextMatrix(53, 4) = "TRUE"
        .TextMatrix(54, 1) = 5: .TextMatrix(54, 3) = 47: .TextMatrix(54, 4) = "TRUE"
        .TextMatrix(55, 1) = 6: .TextMatrix(55, 3) = 47: .TextMatrix(55, 4) = "TRUE"
        .TextMatrix(56, 1) = 7: .TextMatrix(56, 3) = 47: .TextMatrix(56, 4) = "TRUE"
        .TextMatrix(48, 1) = 5
        .TextMatrix(57, 1) = 1: .TextMatrix(57, 3) = 48: .TextMatrix(57, 4) = "TRUE"
        .TextMatrix(58, 1) = 2: .TextMatrix(58, 3) = 48: .TextMatrix(58, 4) = "TRUE"
        .TextMatrix(59, 1) = 4: .TextMatrix(59, 3) = 48: .TextMatrix(59, 4) = "TRUE"
        .TextMatrix(60, 1) = 5: .TextMatrix(60, 3) = 48: .TextMatrix(60, 4) = "TRUE"
        .TextMatrix(49, 1) = 6
        .TextMatrix(61, 1) = 1: .TextMatrix(61, 3) = 49: .TextMatrix(61, 4) = "TRUE"
        .TextMatrix(62, 1) = 2: .TextMatrix(62, 3) = 49: .TextMatrix(62, 4) = "TRUE"
        .TextMatrix(63, 1) = 6: .TextMatrix(63, 3) = 49: .TextMatrix(63, 4) = "TRUE"
        .TextMatrix(50, 1) = 7
        .TextMatrix(64, 1) = 1: .TextMatrix(64, 3) = 50: .TextMatrix(64, 4) = "FALSE"
        .TextMatrix(65, 1) = 2: .TextMatrix(65, 3) = 50: .TextMatrix(65, 4) = "TRUE"
        .TextMatrix(66, 1) = 3: .TextMatrix(66, 3) = 50: .TextMatrix(66, 4) = "FALSE"
        .TextMatrix(67, 1) = 4: .TextMatrix(67, 3) = 50: .TextMatrix(67, 4) = "TRUE"
        .TextMatrix(68, 1) = 5: .TextMatrix(68, 3) = 50: .TextMatrix(68, 4) = "TRUE"
        '.TextMatrix(97, 1) = 6: .TextMatrix(97, 3) = 50: .TextMatrix(97, 4) = "TRUE"
        .TextMatrix(69, 1) = 1
        .TextMatrix(73, 1) = 1: .TextMatrix(73, 3) = 69: .TextMatrix(73, 4) = "TRUE"
        .TextMatrix(74, 1) = 2: .TextMatrix(74, 3) = 69: .TextMatrix(74, 4) = "TRUE"
        .TextMatrix(75, 1) = 3: .TextMatrix(75, 3) = 69: .TextMatrix(75, 4) = "TRUE"
        .TextMatrix(70, 1) = 2
        .TextMatrix(76, 1) = 1: .TextMatrix(76, 3) = 70: .TextMatrix(76, 4) = "FALSE"
        .TextMatrix(77, 1) = 2: .TextMatrix(77, 3) = 70: .TextMatrix(77, 4) = "FALSE"
        .TextMatrix(78, 1) = 3: .TextMatrix(78, 3) = 70: .TextMatrix(78, 4) = "FALSE"
        .TextMatrix(79, 1) = 4: .TextMatrix(79, 3) = 70: .TextMatrix(79, 4) = "TRUE"
        .TextMatrix(80, 1) = 5: .TextMatrix(80, 3) = 70: .TextMatrix(80, 4) = "TRUE"
        .TextMatrix(81, 1) = 6: .TextMatrix(81, 3) = 70: .TextMatrix(81, 4) = "TRUE"
        .TextMatrix(71, 1) = 3
        .TextMatrix(82, 1) = 1: .TextMatrix(82, 3) = 71: .TextMatrix(82, 4) = "FALSE"
        .TextMatrix(83, 1) = 2: .TextMatrix(83, 3) = 71: .TextMatrix(83, 4) = "FALSE"
        .TextMatrix(84, 1) = 3: .TextMatrix(84, 3) = 71: .TextMatrix(84, 4) = "FALSE"
        .TextMatrix(72, 1) = 4
        .TextMatrix(85, 1) = 1: .TextMatrix(85, 3) = 72: .TextMatrix(85, 4) = "TRUE"
        .TextMatrix(86, 1) = 2: .TextMatrix(86, 3) = 72: .TextMatrix(86, 4) = "TRUE"
        .TextMatrix(87, 1) = 3: .TextMatrix(87, 3) = 72: .TextMatrix(87, 4) = "TRUE"
        .TextMatrix(88, 1) = 4: .TextMatrix(88, 3) = 72: .TextMatrix(88, 4) = "FALSE"
        .TextMatrix(89, 1) = 6: .TextMatrix(89, 3) = 72: .TextMatrix(89, 4) = "FALSE"
        .TextMatrix(90, 1) = 1: .TextMatrix(90, 4) = "TRUE"
        .TextMatrix(91, 1) = 2: .TextMatrix(91, 4) = "FALSE"
        .TextMatrix(92, 1) = 4
        .TextMatrix(94, 1) = 1: .TextMatrix(94, 3) = 92: .TextMatrix(94, 4) = "TRUE"
        .TextMatrix(93, 1) = 5
        .TextMatrix(95, 1) = 1: .TextMatrix(95, 3) = 93: .TextMatrix(95, 4) = "TRUE"
        .TextMatrix(96, 1) = 1: .TextMatrix(96, 3) = 93: .TextMatrix(96, 4) = "TRUE"
        .TextMatrix(97, 1) = 5: .TextMatrix(97, 3) = 13: .TextMatrix(97, 4) = "TRUE"
        .TextMatrix(98, 1) = 3: .TextMatrix(98, 3) = 49: .TextMatrix(98, 4) = "TRUE"
        .TextMatrix(99, 1) = 4: .TextMatrix(99, 3) = 49: .TextMatrix(99, 4) = "TRUE"
        .TextMatrix(100, 1) = 7: .TextMatrix(100, 3) = 70: .TextMatrix(100, 4) = "TRUE"
        .TextMatrix(101, 1) = 7: .TextMatrix(101, 3) = 72: .TextMatrix(101, 4) = "TRUE"
        
    End With
    cboAccess.AddItem "YES"
    cboAccess.AddItem "NO"
    cboAccess.ListIndex = 0
    cboAType.AddItem "FULL ACCESS"
    cboAType.AddItem "VIEW ONLY"
    tabAccess.Tab = 0
    tabAccess.TabEnabled(1) = False
    optAll.Value = True
    With flxFrm
        .ColWidth(0) = 300
        .ColWidth(1) = .Width - 300
        .ColWidth(2) = 0
    End With
    Call LockData(True)
End Sub


Private Sub BtnAdmin(SaveCancel As Boolean)
    Dim i As Integer
    
    If SaveCancel Then
        For i = 0 To 3
            cmdTrans(i).Visible = False
        Next i
        cmdTrans(4).Visible = True
        cmdTrans(4).Left = 0
        cmdTrans(5).Visible = True
        cmdTrans(5).Left = cmdTrans(4).Width
    Else
        cmdTrans(0).Left = 0
        cmdTrans(0).Visible = True
        For i = 1 To 3
            cmdTrans(i).Left = cmdTrans(0).Width * i
            cmdTrans(i).Visible = True
        Next i
        cmdTrans(4).Visible = False
        cmdTrans(5).Visible = False
    End If
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    rsEmployee.Close
    MotoRep.sBar.Panels(6).Text = ""
'    MotoRep.tbMain.Visible = True
End Sub

Private Sub optAll_Click()
    Dim i As Integer
    
    If optAll.Value = True Then
        flxFrm.Rows = 0
        For i = 1 To flxN.Rows - 1
            If flxN.TextMatrix(i, 2) = "FULL ACCESS" Or flxN.TextMatrix(i, 2) = "VIEW ONLY" Then flxFrm.AddItem vbTab & TreeView1.Nodes(i).Text & vbTab & i
        Next i
    End If

End Sub

Private Sub optFull_Click()
    Dim i As Integer
    
    If optFull.Value = True Then
        flxFrm.Rows = 0
        For i = 1 To flxN.Rows - 1
            If flxN.TextMatrix(i, 2) = "FULL ACCESS" Then flxFrm.AddItem vbTab & TreeView1.Nodes(i).Text & vbTab & i
        Next i
    End If
End Sub

Private Sub optView_Click()
    Dim i As Integer
    
    If optView.Value = True Then
        flxFrm.Rows = 0
        For i = 1 To flxN.Rows - 1
            If flxN.TextMatrix(i, 2) = "VIEW ONLY" Then flxFrm.AddItem vbTab & TreeView1.Nodes(i).Text & vbTab & i
        Next i
    End If
End Sub

Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
    Dim mTop As Integer
    
    lblFormName.Caption = Node.Text
    xRow = Node.Index
    lblRem.Caption = ""
    If flxN.TextMatrix(xRow, 1) = "" Then       'If no Menu ID Disable all
        cboAccess.Text = "YES"
        cboAccess.Enabled = False
        cboAType.Enabled = False
        Exit Sub
    Else
        cboAccess.Enabled = True
        cboAType.Enabled = True
    End If
    If flxN.TextMatrix(xRow, 3) <> "" Then  'Current Node is a Sub Menu
        mTop = Val(flxN.TextMatrix(xRow, 3))
        If flxN.TextMatrix(mTop, 2) = "" Then   'Top Menu is not yet Accessible
            cboAccess.Enabled = False
            cboAType.Enabled = False
            lblRem.Caption = "User has no access to " & TreeView1.Nodes(mTop).Text
        End If
    End If
    If flxN.TextMatrix(xRow, 2) <> "" Then      'Display current value of Access Type
        If flxN.TextMatrix(xRow, 2) = "TOP LEVEL MENU" Then
            cboAType.Text = ""
            cboAType.Enabled = False
            lblRem.Caption = "Top Level Menu"
        Else
            cboAType.Text = flxN.TextMatrix(xRow, 2)
            If flxN.TextMatrix(xRow, 4) = "FALSE" Then
                cboAType.Enabled = False
                lblRem.Caption = "Display Type Application"
            Else
                lblRem.Caption = "Data Processing Type Application"
            End If
        End If
        cboAccess.Text = "YES"
    Else
        cboAccess.Text = "NO"
        cboAType.Text = ""
    End If
End Sub

Private Sub txtcPW_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtEmployee_Change()
    With flxEmp
        If CustSrch Or Not AddNew Then
            CustSrch = False
            Exit Sub
        End If
        If txtEmployee.Text = "" Then .Visible = False: Exit Sub
        rsEmployee.Find "[EMPLOYEE] LIKE '" + txtEmployee.Text + "%'", , adSearchForward, 1
        If Not rsEmployee.EOF Then
            .Top = Frame1.Top + txtEmployee.Top + txtEmployee.Height + 30
            .Left = Frame1.Left + txtEmployee.Left
            .Width = txtEmployee.Width
            .Visible = True
            If Not CustBln Then
                txtEmployee.Tag = Len(txtEmployee.Text)
                CustBln = True
            End If
            .TopRow = rsEmployee.AbsolutePosition - 1
            .Row = rsEmployee.AbsolutePosition - 1
            .Col = 0: .ColSel = 2
            txtEmployee.Text = rsEmployee!EMPLOYEE
            
            'Insert Additional retrieval here
            txtEmpID.Text = flxEmp.TextMatrix(.Row, 2)
            
            txtEmployee.SelStart = Val(txtEmployee.Tag)
            txtEmployee.SelLength = Len(txtEmployee.Text)
            CustBln = False
        Else
            'rsEmployee.Requery
            txtEmpID.Text = ""
            .Visible = False
        End If
    End With
End Sub
                                                           
Private Sub txtEmployee_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
        Case vbKeyDown
            If flxEmp.Visible Then
                flxEmp.SetFocus
            Else
                txtUser.SetFocus
            End If
        Case vbKeyEscape
            txtEmployee.Text = ""
            flxEmp.Visible = False
        Case vbKeyDelete
            CustSrch = True
        Case vbKeyBack
            CustSrch = True
    End Select

End Sub

Private Sub txtEmployee_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        flxEmp.Visible = False
        txtUserID.SetFocus
    End If
End Sub

Private Sub txtPW_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtUser_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtUserID_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtUser.SetFocus
End Sub
