VERSION 5.00
Begin VB.Form frmHoldBackground 
   Caption         =   "frmHoldBackground"
   ClientHeight    =   5550
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11430
   LinkTopic       =   "Form1"
   ScaleHeight     =   5550
   ScaleWidth      =   11430
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picBackground 
      Height          =   975
      Index           =   38
      Left            =   7440
      MouseIcon       =   "frmHoldBackground.frx":0000
      Picture         =   "frmHoldBackground.frx":57E2
      ScaleHeight     =   915
      ScaleWidth      =   1035
      TabIndex        =   38
      Tag             =   "MCB"
      Top             =   3840
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.PictureBox picBackground 
      Height          =   975
      Index           =   32
      Left            =   9480
      MouseIcon       =   "frmHoldBackground.frx":4ACF7
      Picture         =   "frmHoldBackground.frx":504D9
      ScaleHeight     =   915
      ScaleWidth      =   1035
      TabIndex        =   36
      Tag             =   "BSPT"
      Top             =   2400
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.PictureBox picBackground 
      Height          =   975
      Index           =   33
      Left            =   5760
      MouseIcon       =   "frmHoldBackground.frx":9B2B4
      Picture         =   "frmHoldBackground.frx":A0A96
      ScaleHeight     =   915
      ScaleWidth      =   1035
      TabIndex        =   35
      Tag             =   "BSPT"
      Top             =   3960
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.PictureBox picBackground 
      Height          =   975
      Index           =   29
      Left            =   29
      ScaleHeight     =   915
      ScaleWidth      =   1035
      TabIndex        =   32
      Tag             =   "MWH"
      Top             =   3720
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.PictureBox picBackground 
      Height          =   975
      Index           =   30
      Left            =   1560
      MouseIcon       =   "frmHoldBackground.frx":CF8E1
      Picture         =   "frmHoldBackground.frx":D50C3
      ScaleHeight     =   915
      ScaleWidth      =   1035
      TabIndex        =   30
      Tag             =   "MCB"
      Top             =   3720
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.PictureBox picDefault 
      Height          =   975
      Left            =   3240
      Picture         =   "frmHoldBackground.frx":18B008
      ScaleHeight     =   915
      ScaleWidth      =   1875
      TabIndex        =   28
      Top             =   4080
      Width           =   1935
   End
   Begin VB.PictureBox picBackground 
      Height          =   975
      Index           =   22
      Left            =   240
      MouseIcon       =   "frmHoldBackground.frx":19E9A7
      Picture         =   "frmHoldBackground.frx":1A4189
      ScaleHeight     =   915
      ScaleWidth      =   1035
      TabIndex        =   20
      Tag             =   "MV2"
      Top             =   2400
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.PictureBox picBackground 
      Height          =   975
      Index           =   23
      Left            =   1560
      MouseIcon       =   "frmHoldBackground.frx":1D4682
      Picture         =   "frmHoldBackground.frx":21CF0C
      ScaleHeight     =   915
      ScaleWidth      =   1035
      TabIndex        =   19
      Tag             =   "MG"
      Top             =   2400
      Width           =   1095
   End
   Begin VB.PictureBox picBackground 
      Height          =   975
      Index           =   24
      Left            =   2880
      MouseIcon       =   "frmHoldBackground.frx":27E893
      Picture         =   "frmHoldBackground.frx":2850E5
      ScaleHeight     =   915
      ScaleWidth      =   1035
      TabIndex        =   18
      Tag             =   "BST"
      Top             =   2400
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.PictureBox picBackground 
      Height          =   975
      Index           =   25
      Left            =   4200
      MouseIcon       =   "frmHoldBackground.frx":2C7CEE
      Picture         =   "frmHoldBackground.frx":2C86F0
      ScaleHeight     =   915
      ScaleWidth      =   1035
      TabIndex        =   17
      Tag             =   "BBS"
      Top             =   2400
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.PictureBox picBackground 
      Height          =   975
      Index           =   26
      Left            =   5520
      MouseIcon       =   "frmHoldBackground.frx":31C807
      Picture         =   "frmHoldBackground.frx":31D209
      ScaleHeight     =   915
      ScaleWidth      =   1035
      TabIndex        =   16
      Tag             =   "YMK"
      Top             =   2400
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.PictureBox picBackground 
      Height          =   975
      Index           =   27
      Left            =   6840
      MouseIcon       =   "frmHoldBackground.frx":335AEC
      Picture         =   "frmHoldBackground.frx":33B2CE
      ScaleHeight     =   915
      ScaleWidth      =   1035
      TabIndex        =   15
      Tag             =   "MV3"
      Top             =   2400
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.PictureBox picBackground 
      Height          =   975
      Index           =   28
      Left            =   8160
      MouseIcon       =   "frmHoldBackground.frx":3A0759
      Picture         =   "frmHoldBackground.frx":3A115B
      ScaleHeight     =   915
      ScaleWidth      =   1035
      TabIndex        =   14
      Tag             =   "MRT"
      Top             =   2400
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.PictureBox picBackground 
      Height          =   975
      Index           =   20
      Left            =   8160
      MouseIcon       =   "frmHoldBackground.frx":438968
      Picture         =   "frmHoldBackground.frx":43936A
      ScaleHeight     =   915
      ScaleWidth      =   1035
      TabIndex        =   12
      Tag             =   "BSP"
      Top             =   480
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.PictureBox picBackground 
      Height          =   975
      Index           =   19
      Left            =   6840
      MouseIcon       =   "frmHoldBackground.frx":486974
      Picture         =   "frmHoldBackground.frx":487FBE
      ScaleHeight     =   915
      ScaleWidth      =   1035
      TabIndex        =   10
      Tag             =   "MP"
      Top             =   480
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.PictureBox picBackground 
      Height          =   975
      Index           =   16
      Left            =   5520
      MouseIcon       =   "frmHoldBackground.frx":4C8F1A
      Picture         =   "frmHoldBackground.frx":4CE6FC
      ScaleHeight     =   915
      ScaleWidth      =   1035
      TabIndex        =   8
      Tag             =   "MV1"
      Top             =   480
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.PictureBox picBackground 
      Height          =   975
      Index           =   15
      Left            =   4200
      MouseIcon       =   "frmHoldBackground.frx":4FDED6
      Picture         =   "frmHoldBackground.frx":515C00
      ScaleHeight     =   915
      ScaleWidth      =   1035
      TabIndex        =   6
      Tag             =   "MPT"
      Top             =   480
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.PictureBox picBackground 
      Height          =   975
      Index           =   14
      Left            =   2880
      MouseIcon       =   "frmHoldBackground.frx":534C5F
      Picture         =   "frmHoldBackground.frx":53A441
      ScaleHeight     =   915
      ScaleWidth      =   1035
      TabIndex        =   2
      Tag             =   "YMT"
      Top             =   480
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.PictureBox picBackground 
      Height          =   975
      Index           =   13
      Left            =   1560
      MouseIcon       =   "frmHoldBackground.frx":5D4495
      Picture         =   "frmHoldBackground.frx":5EC1BF
      ScaleHeight     =   915
      ScaleWidth      =   1035
      TabIndex        =   1
      Tag             =   "MUD"
      Top             =   480
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.PictureBox picBackground 
      Height          =   975
      Index           =   12
      Left            =   240
      MouseIcon       =   "frmHoldBackground.frx":6A73E6
      Picture         =   "frmHoldBackground.frx":6BF110
      ScaleHeight     =   915
      ScaleWidth      =   1035
      TabIndex        =   0
      Tag             =   "MST"
      Top             =   480
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label19 
      Caption         =   "MCPR"
      Height          =   255
      Left            =   9720
      TabIndex        =   37
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label Label18 
      Caption         =   "ALL"
      Height          =   255
      Left            =   6120
      TabIndex        =   34
      Top             =   3600
      Width           =   375
   End
   Begin VB.Label Label17 
      Caption         =   "MWH"
      Height          =   255
      Left            =   600
      TabIndex        =   33
      Top             =   3480
      Width           =   375
   End
   Begin VB.Label Label16 
      Caption         =   "MCB"
      Height          =   255
      Left            =   1920
      TabIndex        =   31
      Top             =   3480
      Width           =   375
   End
   Begin VB.Label Label15 
      Caption         =   "DEFAULT"
      Height          =   255
      Left            =   3840
      TabIndex        =   29
      Top             =   3840
      Width           =   855
   End
   Begin VB.Label Label14 
      Caption         =   "MV2"
      Height          =   255
      Left            =   600
      TabIndex        =   27
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label Label13 
      Caption         =   "MG"
      Height          =   255
      Left            =   1920
      TabIndex        =   26
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label Label12 
      Caption         =   "BST"
      Height          =   255
      Left            =   3240
      TabIndex        =   25
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label Label11 
      Caption         =   "BBS"
      Height          =   255
      Left            =   4560
      TabIndex        =   24
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label Label10 
      Caption         =   "YMK"
      Height          =   255
      Left            =   5880
      TabIndex        =   23
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label Label9 
      Caption         =   "MV3"
      Height          =   255
      Left            =   7200
      TabIndex        =   22
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label Label8 
      Caption         =   "MRT"
      Height          =   255
      Left            =   8520
      TabIndex        =   21
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label Label7 
      Caption         =   "BSP"
      Height          =   255
      Left            =   8520
      TabIndex        =   13
      Top             =   240
      Width           =   375
   End
   Begin VB.Label Label6 
      Caption         =   "MP"
      Height          =   255
      Left            =   7200
      TabIndex        =   11
      Top             =   240
      Width           =   375
   End
   Begin VB.Label Label5 
      Caption         =   "MV1"
      Height          =   255
      Left            =   5880
      TabIndex        =   9
      Top             =   240
      Width           =   375
   End
   Begin VB.Label Label4 
      Caption         =   "MPT"
      Height          =   255
      Left            =   4560
      TabIndex        =   7
      Top             =   240
      Width           =   375
   End
   Begin VB.Label Label3 
      Caption         =   "YMT"
      Height          =   255
      Left            =   3240
      TabIndex        =   5
      Top             =   240
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "MUD"
      Height          =   255
      Left            =   1920
      TabIndex        =   4
      Top             =   240
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "MST"
      Height          =   255
      Left            =   600
      TabIndex        =   3
      Top             =   240
      Width           =   375
   End
End
Attribute VB_Name = "frmHoldBackground"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

