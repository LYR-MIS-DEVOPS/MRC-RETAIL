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
      Index           =   33
      Left            =   5640
      MouseIcon       =   "frmHoldBackground.frx":0000
      Picture         =   "frmHoldBackground.frx":57E2
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
      MouseIcon       =   "frmHoldBackground.frx":3462D
      Picture         =   "frmHoldBackground.frx":39E0F
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
      Picture         =   "frmHoldBackground.frx":EFD54
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
      MouseIcon       =   "frmHoldBackground.frx":1036F3
      Picture         =   "frmHoldBackground.frx":108ED5
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
      MouseIcon       =   "frmHoldBackground.frx":1393CE
      Picture         =   "frmHoldBackground.frx":181C58
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
      MouseIcon       =   "frmHoldBackground.frx":1E35DF
      Picture         =   "frmHoldBackground.frx":1E9E31
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
      MouseIcon       =   "frmHoldBackground.frx":22CA3A
      Picture         =   "frmHoldBackground.frx":22D43C
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
      MouseIcon       =   "frmHoldBackground.frx":281553
      Picture         =   "frmHoldBackground.frx":281F55
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
      MouseIcon       =   "frmHoldBackground.frx":29A838
      Picture         =   "frmHoldBackground.frx":2A001A
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
      MouseIcon       =   "frmHoldBackground.frx":3054A5
      Picture         =   "frmHoldBackground.frx":305EA7
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
      MouseIcon       =   "frmHoldBackground.frx":39D6B4
      Picture         =   "frmHoldBackground.frx":39E0B6
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
      MouseIcon       =   "frmHoldBackground.frx":3EB6C0
      Picture         =   "frmHoldBackground.frx":3ECD0A
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
      MouseIcon       =   "frmHoldBackground.frx":42DC66
      Picture         =   "frmHoldBackground.frx":433448
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
      MouseIcon       =   "frmHoldBackground.frx":462C22
      Picture         =   "frmHoldBackground.frx":47A94C
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
      MouseIcon       =   "frmHoldBackground.frx":4999AB
      Picture         =   "frmHoldBackground.frx":49F18D
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
      MouseIcon       =   "frmHoldBackground.frx":5391E1
      Picture         =   "frmHoldBackground.frx":550F0B
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
      MouseIcon       =   "frmHoldBackground.frx":60C132
      Picture         =   "frmHoldBackground.frx":623E5C
      ScaleHeight     =   915
      ScaleWidth      =   1035
      TabIndex        =   0
      Tag             =   "MST"
      Top             =   480
      Visible         =   0   'False
      Width           =   1095
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

