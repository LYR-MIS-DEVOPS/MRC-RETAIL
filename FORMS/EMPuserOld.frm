VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form EMPuser 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User Profile"
   ClientHeight    =   6570
   ClientLeft      =   45
   ClientTop       =   150
   ClientWidth     =   7065
   Icon            =   "EMPuserOld.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6570
   ScaleWidth      =   7065
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab tabAccess 
      Height          =   3135
      Left            =   240
      TabIndex        =   27
      Top             =   3360
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   5530
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Administrator's Access"
      TabPicture(0)   =   "EMPuserOld.frx":3072
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblCheck"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblX"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "flxTrans"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Option1(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Option1(1)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Option1(2)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Option1(3)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Option1(4)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Option1(5)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Option1(6)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      TabCaption(1)   =   "Single Module Access"
      TabPicture(1)   =   "EMPuserOld.frx":308E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame2 
         Height          =   2415
         Left            =   -74760
         TabIndex        =   38
         Top             =   480
         Width           =   6135
         Begin VB.CheckBox chkReceiver 
            Caption         =   "RECEIVER"
            Height          =   255
            Left            =   840
            TabIndex        =   44
            Top             =   1248
            Width           =   3015
         End
         Begin VB.CheckBox chkWarehouse 
            Caption         =   "WAREHOUSEMAN"
            Height          =   255
            Left            =   840
            TabIndex        =   43
            Top             =   1920
            Width           =   3015
         End
         Begin VB.CheckBox chkService 
            Caption         =   "SERVICE CENTER"
            Height          =   255
            Left            =   840
            TabIndex        =   42
            Top             =   1584
            Width           =   3015
         End
         Begin VB.CheckBox chkPartsman 
            Caption         =   "PARTSMAN"
            Height          =   255
            Left            =   840
            TabIndex        =   41
            Top             =   912
            Width           =   3015
         End
         Begin VB.CheckBox chkDisburser 
            Caption         =   "DISBURSER"
            Height          =   255
            Left            =   840
            TabIndex        =   40
            Top             =   576
            Width           =   3015
         End
         Begin VB.CheckBox chkCashier 
            Caption         =   "CASHIER"
            Height          =   255
            Left            =   840
            TabIndex        =   39
            Top             =   240
            Width           =   3015
         End
      End
      Begin VB.OptionButton Option1 
         Height          =   255
         Index           =   6
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   2760
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Tools"
         Height          =   375
         Index           =   5
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   2355
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Reports"
         Height          =   375
         Index           =   4
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   1980
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Transactions"
         Height          =   375
         Index           =   3
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   1605
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Service Center"
         Height          =   375
         Index           =   2
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   1230
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Spare Parts"
         Height          =   375
         Index           =   1
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   855
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Personnel"
         Height          =   375
         Index           =   0
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   480
         Width           =   1575
      End
      Begin MSFlexGridLib.MSFlexGrid flxTrans 
         Height          =   2295
         Left            =   1680
         TabIndex        =   34
         Top             =   480
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   4048
         _Version        =   393216
         Rows            =   1
         Cols            =   5
         FixedRows       =   0
         FixedCols       =   0
         BackColorBkg    =   -2147483633
         ScrollBars      =   0
      End
      Begin VB.Label lblX 
         Caption         =   "ý"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   36
         Top             =   2760
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblCheck 
         Caption         =   "þ"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   35
         Top             =   2760
         Visible         =   0   'False
         Width           =   255
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxEmp 
      Height          =   1815
      Left            =   2280
      TabIndex        =   24
      Top             =   6600
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
      Height          =   2415
      Left            =   240
      TabIndex        =   6
      Top             =   840
      Width           =   6615
      Begin VB.TextBox txtUserID 
         Height          =   285
         Left            =   1800
         TabIndex        =   1
         Top             =   560
         Width           =   3135
      End
      Begin VB.TextBox txtEmpID 
         Height          =   285
         Left            =   5640
         TabIndex        =   25
         Top             =   960
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.OptionButton optAccess 
         Caption         =   "Operator"
         Height          =   255
         Index           =   1
         Left            =   5160
         TabIndex        =   23
         Top             =   575
         Width           =   1095
      End
      Begin VB.OptionButton optAccess 
         Caption         =   "Administrator"
         Height          =   255
         Index           =   0
         Left            =   5160
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
         Width           =   3150
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
         Left            =   3720
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   5
         Top             =   2040
         Width           =   2655
      End
      Begin VB.ComboBox cboUType 
         Height          =   315
         ItemData        =   "EMPuserOld.frx":30AA
         Left            =   240
         List            =   "EMPuserOld.frx":30AC
         TabIndex        =   3
         Top             =   2040
         Width           =   2805
      End
      Begin VB.TextBox txtPW 
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
         Left            =   3720
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   1320
         Width           =   2655
      End
      Begin VB.TextBox txtUser 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   240
         MaxLength       =   10
         TabIndex        =   2
         Top             =   1320
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
         Caption         =   "Designation"
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label lblI_fld 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Confirm Password:"
         ForeColor       =   &H80000007&
         Height          =   195
         Index           =   8
         Left            =   3720
         TabIndex        =   10
         Top             =   1800
         Width           =   1305
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000011&
         Index           =   2
         X1              =   -120
         X2              =   6960
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
         Caption         =   "User Name"
         ForeColor       =   &H80000007&
         Height          =   195
         Index           =   6
         Left            =   240
         TabIndex        =   8
         Top             =   1080
         Width           =   795
      End
      Begin VB.Label lblI_fld 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password:"
         ForeColor       =   &H80000007&
         Height          =   195
         Index           =   7
         Left            =   3720
         TabIndex        =   7
         Top             =   1080
         Width           =   735
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
            Picture         =   "EMPuserOld.frx":30AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "EMPuserOld.frx":3F88
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "EMPuserOld.frx":60C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "EMPuserOld.frx":6F9C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "EMPuserOld.frx":7E76
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "EMPuserOld.frx":8D50
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
      Image           =   "EMPuserOld.frx":AA5A
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
      Image           =   "EMPuserOld.frx":AFF4
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
      Image           =   "EMPuserOld.frx":B58E
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
      Image           =   "EMPuserOld.frx":BB28
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
      Image           =   "EMPuserOld.frx":C0C2
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
      Image           =   "EMPuserOld.frx":C65C
      cBack           =   -2147483633
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
      Left            =   4395
      TabIndex        =   13
      Top             =   545
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
      Left            =   4425
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
Dim rsPcUser As New ADODB.Recordset
Dim AddNew As Boolean, LockFlx As Boolean
Dim rsEmployee As New ADODB.Recordset
Dim CustSrch As Boolean, CustBln As Boolean

Dim UA(7) As Variant
Dim LastX As Integer
Dim ArrPers() As Variant
Dim ArrPersX() As Variant
Dim ArrParts() As Variant
Dim ArrPartsX() As Variant
Dim ArrSrvc() As Variant
Dim ArrSrvcX() As Variant
Dim ArrTrans() As Variant
Dim ArrTransX() As Variant
Dim ArrReport() As Variant
Dim ArrReportX() As Variant
Dim ArrTools() As Variant
Dim ArrToolsX() As Variant
Dim ArrTls() As Variant
Dim ArrTlsX() As Variant

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
    flxTrans.Enabled = Not Bln
    For i = 0 To 6
        Option1(i).Enabled = Not Bln
    Next i
    optAccess(0).Enabled = Not Bln
    optAccess(1).Enabled = Not Bln
End Sub

Sub ClearEntries()
    Dim i As Integer
    
    txtEmployee.Text = ""
    txtUserID.Text = ""
    cboUType.Text = ""
    txtUser.Text = ""
    txtPW.Text = ""
    txtcPW.Text = ""
    flxTrans.Rows = 0
    For i = Option1.LBound To Option1.UBound
        Option1(i).Value = False
        UA(i) = ""
    Next i
    chkCashier.Value = 0
    chkDisburser.Value = 0
    chkPartsman.Value = 0
    chkService.Value = 0
    chkWarehouse.Value = 0
End Sub


Public Sub ShowRecord(EmpId As Integer)
    Dim rsPcUser As New ADODB.Recordset
    Dim Qry As String
    Dim i As Integer
    
    Qry = "SELECT [LASTNAME] & ', ' & [FIRSTNAME] & ' ' & Left$([MIDDLENAME],1) & '.' AS CUSTOMER, PCUSER.PCUSERID, PCUSER.USERNAME, PCUSER.USERTYPE, PCUSER.PASSWORD, PCUSER.EMPID,  PCUSER.FRMPERS, PCUSER.FRMPARTS, PCUSER.FRMSRVC, PCUSER.FRMTRANS, PCUSER.FRMREPORTS, PCUSER.FRMTOOLS, PCUSER.DESIGNATION, PCUSER.PARTSMAN, PCUSER.WAREHOUSE, PCUSER.CASHIER, PCUSER.SERVICE, PCUSER.DISBURSER, PCUSER.RECEIVER " & _
          "FROM EMPLOYEES INNER JOIN PCUSER ON EMPLOYEES.EMPID = PCUSER.EMPID WHERE (((PCUSER.EMPID)=" & EmpId & "));"
    
    With rsPcUser
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        .MoveFirst
        CustSrch = True
        txtEmployee.Text = !CUSTOMER
        txtUserID.Text = IIf(IsNull(!PCUserID), "", !PCUserID)
        cboUType.Text = !Designation
        txtUser.Text = !UserName
        txtPW.Text = !Password
        txtcPW.Text = !Password
        txtEmpID.Text = EmpId
        chkCashier.Value = IIf(!CASHIER = True, 1, 0)
        chkDisburser.Value = IIf(!Disburser = True, 1, 0)
        chkPartsman.Value = IIf(!Partsman = True, 1, 0)
        chkService.Value = IIf(!Service = True, 1, 0)
        chkWarehouse.Value = IIf(!Warehouse = True, 1, 0)
        chkReceiver.Value = IIf(!Receiver = True, 1, 0)
        If !UserType = "Administrator" Then
            optAccess(0).Value = True
            Call BtnAdmin(False)
        Else
            optAccess(1).Value = True
'            Call BtnOperator(False)
'            optAccess(0).Visible = False
'            optAccess(1).Visible = False
'            tabAccess.TabEnabled(0) = False
            tabAccess.Tab = 1
        End If
        For i = 0 To 5
            Option1(i).Enabled = False
            UA(i) = IIf(IsNull(.Fields(i + 6)), "", .Fields(i + 6))
        Next i
        .Close
    End With
        
End Sub

Sub SaveUser()
    Dim rsMyUser As New ADODB.Recordset
    Dim Qry As String
    Dim i As Integer
    
    If AddNew Then
        rsMyUser.Open "PcUser", LYRcon, adOpenKeyset, adLockOptimistic
        rsMyUser.AddNew
        rsMyUser!EmpId = txtEmpID.Text
    Else
        Qry = "SELECT PCUSER.* From PCUSER WHERE (((PCUSER.EMPID)=" & txtEmpID.Text & "));"
        rsMyUser.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        rsMyUser.MoveFirst
    End If
    If optAccess(0).Value = True Then
        rsMyUser!UserType = "Administrator"
    Else
        rsMyUser!UserType = "Operator"
    End If
    rsMyUser!PCUserID = txtUserID.Text
    rsMyUser!UserName = txtUser.Text
    rsMyUser!Designation = cboUType.Text
    rsMyUser!Password = txtPW.Text
    rsMyUser!UpdatedBy = UserID
    rsMyUser!LastUpdate = Date
    rsMyUser!CASHIER = IIf(chkCashier.Value = 1, True, False)
    rsMyUser!Disburser = IIf(chkDisburser.Value = 1, True, False)
    rsMyUser!Partsman = IIf(chkPartsman.Value = 1, True, False)
    rsMyUser!Service = IIf(chkService.Value = 1, True, False)
    rsMyUser!Warehouse = IIf(chkWarehouse.Value = 1, True, False)
    rsMyUser!Receiver = IIf(chkReceiver.Value = 1, True, False)
    For i = 0 To 5
        rsMyUser.Fields(i + 4) = UA(i)
    Next i
    rsMyUser.Update
    If AddNew Then txtEmpID.Text = rsMyUser!EmpId
    rsMyUser.Close
    
    'Update table Employees field PCuser
    LYRcon.Execute "UPDATE EMPLOYEES SET EMPLOYEES.PCUSER = '" & cboUType.Text & "' WHERE (((EMPLOYEES.EMPID)=" & txtEmpID & "));"
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
        vQry = "SELECT PCUSER.PCUSERID From PCUSER WHERE (((PCUSER.PCUSERID)='" & txtUserID.Text & "'));"
        If CtrlNumExist(vQry) Then myMsg = myMsg & vbCrLf & "User ID already Exist"
    End If
    
    If txtPW.Text = "" Or txtcPW.Text = "" Then
        myMsg = myMsg & vbCrLf & "No password defined."
    ElseIf txtPW.Text <> txtcPW.Text Then
        myMsg = myMsg & vbCrLf & "Incorrect Password."
    End If
'    arrCnt = 0
'    For i = 0 To 6
'        If UA(i) <> "" Or Not IsEmpty(UA(i)) Then arrCnt = arrCnt + 1
'    Next i
'    If arrCnt = 0 Then myMsg = myMsg & vbCrLf & "No accessible forms defined."
    
    ValidateEntries = myMsg
    
End Function


Private Sub cboUType_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub cmdTrans_Click(Index As Integer)
    Dim Ans As Integer, myMsg As String
    Dim Qry As String
    
    Select Case Index
        Case 0  'New
            Call ClearEntries
            Call LockData(False)
            Call BtnAdmin(True)
            AddNew = True
        Case 1  'Edit
            If txtEmpID.Text = "" Then
                MsgBox "No available employee to edit.", vbInformation + vbOKOnly, "Edit"
                Exit Sub
            End If
            LockFlx = False
            AddNew = False
            Call LockData(False)
            Call BtnAdmin(True)
        Case 2  'Find
            Qry = "SELECT PCUSER.EMPID, [LASTNAME] & ', ' & [FIRSTNAME] & ' ' & Left$([MIDDLENAME],1) & '.' AS EMPLOYEE " & _
                  "FROM PCUSER INNER JOIN EMPLOYEES ON PCUSER.EMPID = EMPLOYEES.EMPID " & _
                  "ORDER BY [LASTNAME] & ', ' & [FIRSTNAME] & ' ' & Left$([MIDDLENAME],1) & '.';"

            frmSearch.QrySearch = Qry
            frmSearch.lblCap.Caption = "Search for User Name"
            Set frmSearch.FrmSource = EMPuser
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
            MsgBox "User Profile successfully saved!"
            Call LockData(True)
            If UserType = "Administrator" Then
                Call BtnAdmin(False)
            Else
                Call BtnOperator(False)
            End If
            LockFlx = True
        Case 5  'Cancel
            Ans = MsgBox("Do you want to Cancel User Profile?", vbYesNo + vbQuestion, "Save")
            If Ans = vbNo Then Exit Sub
            Call LockData(True)
            If UserType = "Administrator" Then
                Call BtnAdmin(False)
                Call ClearEntries
            Else
                Call BtnOperator(False)
            End If
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


Private Sub flxTrans_Click()
    If LockFlx Then Exit Sub
    With flxTrans
        If .Col = 0 Then
            If .TextMatrix(.Row, 0) = "" Then
                .CellFontName = "Wingdings"
                .CellFontSize = 10
                .TextMatrix(.Row, 0) = lblCheck.Caption
                .Col = 2
                .CellFontName = "Wingdings"
                .CellAlignment = vbCenter
                .CellFontSize = 10
                .TextMatrix(.Row, 2) = lblCheck.Caption
                .TextMatrix(.Row, 3) = "Full Access"
            Else
                .TextMatrix(.Row, 0) = ""
                .TextMatrix(.Row, 2) = ""
                .TextMatrix(.Row, 3) = "Access Type"
            End If
        ElseIf .Col = 2 Then
            If .TextMatrix(.Row, 0) = "" Then Exit Sub
            If .TextMatrix(.Row, 3) = "Full Access" Then
                .TextMatrix(.Row, 2) = lblX.Caption
                .TextMatrix(.Row, 3) = "View Only"
            Else
                .TextMatrix(.Row, 2) = lblCheck.Caption
                .TextMatrix(.Row, 3) = "Full Access"
            End If
        End If
    End With
End Sub

Sub ShowUA()
    Dim temp() As String
    Dim i As Integer, j As Integer
    
    If UA(LastX) = "" Then Exit Sub
    temp = Split(UA(LastX), ",")
    For i = LBound(temp) To UBound(temp)
        With flxTrans
        For j = 0 To .Rows - 1
            If Left$(temp(i), Len(temp(i)) - 1) = .TextMatrix(j, 4) Then
                .Col = 0
                .Row = j
                .CellFontName = "Wingdings"
                .CellFontSize = 10
                .TextMatrix(j, 0) = lblCheck.Caption
                If Right$(temp(i), 1) = "F" Then
                    .Col = 2
                    .Row = j
                    .CellFontName = "Wingdings"
                    .CellFontSize = 10
                    .TextMatrix(j, 2) = lblCheck.Caption
                    .TextMatrix(j, 3) = "Full Access"
                Else
                    .Col = 2
                    .Row = j
                    .CellFontName = "Wingdings"
                    .CellFontSize = 10
                    .TextMatrix(j, 2) = lblX.Caption
                    .TextMatrix(j, 3) = "View Only"
                End If
            End If
        Next j
        End With
    Next i
End Sub
Private Sub flxTrans_LostFocus()
    Dim y As Integer
    Dim temp As String
    
    temp = ""
    For y = 0 To flxTrans.Rows - 1
        If flxTrans.TextMatrix(y, 0) <> "" Then
           temp = IIf(temp = "", flxTrans.TextMatrix(y, 4) & Left$(flxTrans.TextMatrix(y, 3), 1), temp & "," & flxTrans.TextMatrix(y, 4) & Left$(flxTrans.TextMatrix(y, 3), 1))
        End If
    Next y
    UA(LastX) = temp
End Sub

Private Sub Form_Load()
    Dim i As Integer
    

    ArrPers = Array("Employees", "User Profile", "On-Account", "Supplier", "Shipper")
    ArrPersX = Array(1, 2, 4, 6, 7)
    ArrParts = Array("Ordering", "Receiving", "Sales", "Stock Monitoring", "Adjustment")
    ArrPartsX = Array(1, 2, 3, 4, 5)
    ArrSrvc = Array("Job Orders", "Requisition", "Customer Service Record", "Daily Service Income", "Mechanic Service Income", "Service Income Summary")
    ArrSrvcX = Array(1, 2, 3, 4, 5, 6)
    ArrTrans = Array("Sales Invoice", "Official Receipt", "Provisional Receipt", "DM/DM Adjustment", "Disbursement", "Cash Flow", "Accounts Receivable")
    ArrTransX = Array(1, 2, 3, 4, 5, 6, 7)
    ArrReport = Array("Daily Transaction", "Spare Parts Inventory", "Service Center", "Accounts Receivable")
    ArrReportX = Array(1, 2, 3, 4)
    ArrTools = Array("Backup", "Restore", "Beginning Inventory", "Settings")
    ArrToolsX = Array(1, 2, 4, 5)
    With flxTrans
        .ColWidth(0) = 300
        .ColWidth(1) = 2700
        .ColWidth(2) = 300
        .ColWidth(3) = 1700
        .ColWidth(4) = 500
    End With

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
    
        
'    Call ShowRecord(UserID)
    LockFlx = True
    Call LockData(True)
    MotoRep.sBar.Panels(6).Text = "System User's Profile"
    
    'Load Employee's records
    
    Qry = "SELECT UCase$([LASTNAME] & ', ' & [FIRSTNAME] & ' ' & Left$([MIDDLENAME],1) & '.') AS EMPLOYEE, EMPLOYEES.EMPID From EMPLOYEES " & _
          "WHERE (((EMPLOYEES.DEPARTMENT)<>'Default') AND ((EMPLOYEES.INACTIVE)=False) AND ((EMPLOYEES.PCUSER) Is Null)) " & _
          "ORDER BY UCase$([LASTNAME] & ', ' & [FIRSTNAME] & ' ' & Left$([MIDDLENAME],1) & '.');"

    rsEmployee.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    Set flxEmp.DataSource = rsEmployee
    With flxEmp
        .Width = txtEmployee.Width
        .ColWidth(0) = 0
        .ColWidth(1) = txtEmployee.Width - 300
        .ColWidth(2) = 0
    End With
End Sub

Private Sub BtnOperator(SaveCancel As Boolean)
    
    If SaveCancel Then
        For i = 0 To 3
            cmdTrans(i).Visible = False
        Next i
        cmdTrans(4).Visible = True
        cmdTrans(4).Left = 0
        cmdTrans(5).Visible = True
        cmdTrans(5).Left = cmdTrans(4).Width
    Else
        For i = 0 To 5
            cmdTrans(i).Visible = False
        Next i
        cmdTrans(1).Left = 0
        cmdTrans(1).Visible = True
        cmdTrans(3).Left = cmdTrans(1).Width
        cmdTrans(3).Visible = True
    End If
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
    If rsPcUser.State = adStateOpen Then rsPcUser.Close
    rsEmployee.Close
    MotoRep.sBar.Panels(6).Text = ""
    MotoRep.tbMain.Visible = True
End Sub

Private Sub Option1_Click(Index As Integer)
    Dim i As Integer
    
    flxTrans.Rows = 0
    Select Case Index
        Case 0
            For i = LBound(ArrPers) To UBound(ArrPers)
                flxTrans.AddItem vbTab & ArrPers(i) & vbTab & vbTab & "Access Type" & vbTab & ArrPersX(i)
            Next i
        Case 1
            For i = LBound(ArrParts) To UBound(ArrParts)
                flxTrans.AddItem vbTab & ArrParts(i) & vbTab & vbTab & "Access Type" & vbTab & ArrPartsX(i)
            Next i
        Case 2
            For i = LBound(ArrSrvc) To UBound(ArrSrvc)
                flxTrans.AddItem vbTab & ArrSrvc(i) & vbTab & vbTab & "Access Type" & vbTab & ArrSrvcX(i)
            Next i
        Case 3
            For i = LBound(ArrTrans) To UBound(ArrTrans)
                flxTrans.AddItem vbTab & ArrTrans(i) & vbTab & vbTab & "Access Type" & vbTab & ArrTransX(i)
            Next i
        Case 4
            For i = LBound(ArrReport) To UBound(ArrReport)
                flxTrans.AddItem vbTab & ArrReport(i) & vbTab & vbTab & "Access Type" & vbTab & ArrReportX(i)
            Next i
        Case 5
            For i = LBound(ArrTools) To UBound(ArrTools)
                flxTrans.AddItem vbTab & ArrTools(i) & vbTab & vbTab & "Access Type" & vbTab & ArrToolsX(i)
            Next i
        Case 6
            For i = LBound(ArrTls) To UBound(ArrTls)
                flxTrans.AddItem vbTab & ArrTls(i) & vbTab & vbTab & "Access Type" & vbTab & ArrTlsX(i)
            Next i
    End Select
    LastX = Index
    Call ShowUA

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
            txtEmployee.Text = rsEmployee!Employee
            
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
