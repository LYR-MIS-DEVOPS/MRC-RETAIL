VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Invoice 
   Caption         =   "Sales Invoice"
   ClientHeight    =   8100
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11865
   ControlBox      =   0   'False
   Icon            =   "Invoice.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8100
   ScaleWidth      =   11865
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtTotPen 
      Height          =   315
      Left            =   7020
      TabIndex        =   86
      Top             =   8520
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtJobID 
      Height          =   285
      Left            =   7320
      TabIndex        =   83
      Top             =   8160
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtNTerms 
      Height          =   285
      Left            =   6840
      TabIndex        =   82
      Top             =   8160
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtPScheme 
      Height          =   285
      Left            =   6360
      TabIndex        =   81
      Top             =   8160
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtNPay 
      Height          =   285
      Left            =   5880
      TabIndex        =   80
      Top             =   8160
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtCType 
      Height          =   285
      Left            =   4560
      TabIndex        =   79
      Top             =   8160
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtDueAmt 
      Height          =   285
      Left            =   3720
      TabIndex        =   63
      Top             =   8160
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtCustomerID 
      Height          =   285
      Left            =   3120
      TabIndex        =   61
      Top             =   8160
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtReqID 
      Height          =   285
      Left            =   2640
      TabIndex        =   60
      Top             =   8160
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtPrnCopy 
      Height          =   285
      Left            =   2040
      TabIndex        =   54
      Top             =   8160
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtCashierID 
      Height          =   285
      Left            =   1320
      TabIndex        =   53
      Top             =   8160
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtSPSorderID 
      Height          =   285
      Left            =   720
      TabIndex        =   52
      Top             =   8160
      Visible         =   0   'False
      Width           =   615
   End
   Begin ResizeXtra.RXSplitter RXSplitter1 
      Height          =   7575
      Left            =   120
      TabIndex        =   14
      Tag             =   "XYF"
      Top             =   360
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   13361
      BorderStyle     =   0
      PanelAMinSize   =   3000
      PanelASize      =   3063
      PanelBSize      =   8492
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      Begin VB.Frame Frame3 
         Height          =   7575
         Left            =   0
         TabIndex        =   15
         Top             =   0
         Width           =   3063
         Begin VB.ComboBox cboCashier 
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   84
            Top             =   6840
            Width           =   2775
         End
         Begin VB.ComboBox cboOthers 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   240
            TabIndex        =   44
            Top             =   8400
            Visible         =   0   'False
            Width           =   2535
         End
         Begin VB.Frame Frame2 
            Height          =   1335
            Left            =   120
            TabIndex        =   41
            Top             =   7200
            Visible         =   0   'False
            Width           =   2775
            Begin lvButton.lvButtons_H cmdCancel 
               Height          =   495
               Left            =   120
               TabIndex        =   43
               Top             =   720
               Width           =   2535
               _ExtentX        =   4471
               _ExtentY        =   873
               Caption         =   "&Cancel Payment"
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
               Image           =   "Invoice.frx":27A2
               cBack           =   -2147483633
            End
            Begin lvButton.lvButtons_H cmdSave 
               Height          =   495
               Left            =   120
               TabIndex        =   42
               Top             =   240
               Width           =   2535
               _ExtentX        =   4471
               _ExtentY        =   873
               Caption         =   "&Save Payment  "
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
               Image           =   "Invoice.frx":2D3C
               cBack           =   -2147483633
            End
         End
         Begin VB.Frame Frame1 
            Height          =   2175
            Left            =   120
            TabIndex        =   40
            Top             =   4560
            Width           =   2775
            Begin VB.Image imgOptr 
               BorderStyle     =   1  'Fixed Single
               Height          =   1695
               Left            =   360
               Stretch         =   -1  'True
               Top             =   240
               Width           =   1935
            End
         End
         Begin lvButton.lvButtons_H cmdQuit 
            Height          =   495
            Left            =   240
            TabIndex        =   24
            TabStop         =   0   'False
            Top             =   3840
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   873
            Caption         =   "&Quit"
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
            Image           =   "Invoice.frx":32D6
            cBack           =   -2147483633
         End
         Begin lvButton.lvButtons_H cmdCashier 
            Height          =   495
            Index           =   0
            Left            =   240
            TabIndex        =   25
            TabStop         =   0   'False
            Top             =   240
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   873
            Caption         =   "&1] Cash Invoice  "
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
            Image           =   "Invoice.frx":3870
            cBack           =   -2147483633
         End
         Begin lvButton.lvButtons_H cmdCashier 
            Height          =   495
            Index           =   1
            Left            =   240
            TabIndex        =   26
            Top             =   735
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   873
            Caption         =   "&2] Quick Service "
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
            Image           =   "Invoice.frx":9062
            cBack           =   -2147483633
         End
         Begin lvButton.lvButtons_H cmdCashier 
            Height          =   495
            Index           =   2
            Left            =   240
            TabIndex        =   27
            TabStop         =   0   'False
            Top             =   1230
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   873
            Caption         =   "&3] Charge Invoice"
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
            Image           =   "Invoice.frx":993C
            cBack           =   -2147483633
         End
         Begin lvButton.lvButtons_H cmdCashier 
            Height          =   495
            Index           =   3
            Left            =   240
            TabIndex        =   28
            TabStop         =   0   'False
            Top             =   1680
            Visible         =   0   'False
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   873
            Caption         =   "&4] Others            "
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
            Image           =   "Invoice.frx":A216
            cBack           =   -2147483633
         End
      End
      Begin VB.Frame Frame4 
         Height          =   7575
         Left            =   3178
         TabIndex        =   16
         Tag             =   "XYF"
         Top             =   0
         Width           =   8477
         Begin VB.ComboBox cboPayType 
            Height          =   315
            Left            =   4680
            TabIndex        =   85
            Tag             =   "XYF"
            Top             =   3360
            Width           =   3495
         End
         Begin TabDlg.SSTab tabTerm 
            Height          =   1695
            Left            =   120
            TabIndex        =   64
            Tag             =   "XYF"
            Top             =   1980
            Visible         =   0   'False
            Width           =   4455
            _ExtentX        =   7858
            _ExtentY        =   2990
            _Version        =   393216
            Tab             =   2
            TabHeight       =   520
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            TabCaption(0)   =   "Term"
            TabPicture(0)   =   "Invoice.frx":B0F0
            Tab(0).ControlEnabled=   0   'False
            Tab(0).Control(0)=   "FraTerms"
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "Down Payment"
            TabPicture(1)   =   "Invoice.frx":B10C
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "Frame5"
            Tab(1).ControlCount=   1
            TabCaption(2)   =   "Promo Card"
            TabPicture(2)   =   "Invoice.frx":B128
            Tab(2).ControlEnabled=   -1  'True
            Tab(2).Control(0)=   "Frame6"
            Tab(2).Control(0).Enabled=   0   'False
            Tab(2).ControlCount=   1
            Begin VB.Frame Frame6 
               Height          =   1260
               Left            =   120
               TabIndex        =   88
               Tag             =   "XYF"
               Top             =   360
               Width           =   4215
               Begin VB.TextBox txtPCPoints 
                  Alignment       =   2  'Center
                  BackColor       =   &H8000000B&
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
                  Left            =   960
                  Locked          =   -1  'True
                  TabIndex        =   94
                  Tag             =   "XYF"
                  Top             =   885
                  Width           =   3135
               End
               Begin VB.TextBox txtPCType 
                  Alignment       =   2  'Center
                  BackColor       =   &H8000000B&
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
                  Left            =   960
                  Locked          =   -1  'True
                  TabIndex        =   93
                  Tag             =   "XYF"
                  Top             =   532
                  Width           =   3135
               End
               Begin VB.TextBox txtPCCardNo 
                  Alignment       =   2  'Center
                  BackColor       =   &H8000000B&
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
                  Left            =   960
                  Locked          =   -1  'True
                  TabIndex        =   91
                  Tag             =   "XYF"
                  Top             =   180
                  Width           =   3135
               End
               Begin VB.Label Label8 
                  Caption         =   "TYPE:"
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
                  TabIndex        =   92
                  Tag             =   "XYF"
                  Top             =   555
                  Width           =   735
               End
               Begin VB.Label Label7 
                  Caption         =   "EARN:"
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
                  TabIndex        =   90
                  Tag             =   "XYF"
                  Top             =   915
                  Width           =   735
               End
               Begin VB.Label Label4 
                  Caption         =   "CARD NO.:"
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
                  TabIndex        =   89
                  Tag             =   "XYF"
                  Top             =   210
                  Width           =   975
               End
            End
            Begin VB.Frame Frame5 
               Height          =   1260
               Left            =   -74880
               TabIndex        =   72
               Tag             =   "XYF"
               Top             =   375
               Width           =   4215
               Begin VB.TextBox txtORamt 
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
                  Left            =   1680
                  Locked          =   -1  'True
                  TabIndex        =   75
                  Tag             =   "XYF"
                  Top             =   532
                  Width           =   2415
               End
               Begin VB.TextBox txtORno 
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
                  Left            =   1680
                  TabIndex        =   74
                  Tag             =   "XYF"
                  Top             =   180
                  Width           =   2415
               End
               Begin VB.TextBox txtORdate 
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
                  Left            =   1680
                  Locked          =   -1  'True
                  TabIndex        =   73
                  Tag             =   "XYF"
                  Top             =   885
                  Width           =   2415
               End
               Begin VB.Label lblORamt 
                  Caption         =   "O.R. AMOUNT"
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
                  TabIndex        =   78
                  Tag             =   "XYF"
                  Top             =   555
                  Width           =   1575
               End
               Begin VB.Label lblORno 
                  Caption         =   "O.R. NUMBER"
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
                  TabIndex        =   77
                  Tag             =   "XYF"
                  Top             =   210
                  Width           =   1215
               End
               Begin VB.Label lblORdate 
                  Caption         =   "O.R. DATE"
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
                  TabIndex        =   76
                  Tag             =   "XYF"
                  Top             =   915
                  Width           =   1695
               End
            End
            Begin VB.Frame FraTerms 
               Height          =   1260
               Left            =   -74880
               TabIndex        =   65
               Tag             =   "XYF"
               Top             =   375
               Width           =   4215
               Begin VB.TextBox txtPayStart 
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
                  Left            =   1680
                  Locked          =   -1  'True
                  TabIndex        =   68
                  Tag             =   "XYF"
                  Top             =   885
                  Width           =   2415
               End
               Begin VB.TextBox txtTerm 
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
                  Left            =   1680
                  Locked          =   -1  'True
                  TabIndex        =   67
                  Tag             =   "XYF"
                  Top             =   180
                  Width           =   2415
               End
               Begin VB.TextBox txtAmtDue 
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
                  Left            =   1680
                  Locked          =   -1  'True
                  TabIndex        =   66
                  Tag             =   "XYF"
                  Top             =   532
                  Width           =   2415
               End
               Begin VB.Label Label3 
                  Caption         =   "PAYMENT START"
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
                  TabIndex        =   71
                  Tag             =   "XYF"
                  Top             =   915
                  Width           =   1695
               End
               Begin VB.Label Label5 
                  Caption         =   "TERMS "
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
                  TabIndex        =   70
                  Tag             =   "XYF"
                  Top             =   210
                  Width           =   1215
               End
               Begin VB.Label Label6 
                  Caption         =   "AMOUNT DUE"
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
                  TabIndex        =   69
                  Tag             =   "XYF"
                  Top             =   562
                  Width           =   1215
               End
            End
         End
         Begin TabDlg.SSTab tabPreview 
            Height          =   3735
            Left            =   120
            TabIndex        =   49
            Tag             =   "XYF"
            Top             =   3720
            Width           =   8175
            _ExtentX        =   14420
            _ExtentY        =   6588
            _Version        =   393216
            Tabs            =   1
            TabsPerRow      =   2
            TabHeight       =   520
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            TabCaption(0)   =   "Sales Order"
            TabPicture(0)   =   "Invoice.frx":B144
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "lblRemarks"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "flxOrder"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).Control(2)=   "txtTotal"
            Tab(0).Control(2).Enabled=   0   'False
            Tab(0).ControlCount=   3
            Begin VB.TextBox txtTotal 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   6720
               TabIndex        =   51
               Tag             =   "XYF"
               Top             =   3360
               Width           =   1335
            End
            Begin MSFlexGridLib.MSFlexGrid flxOrder 
               Height          =   3015
               Left            =   120
               TabIndex        =   50
               Tag             =   "XYF"
               Top             =   360
               Width           =   7935
               _ExtentX        =   13996
               _ExtentY        =   5318
               _Version        =   393216
               Cols            =   9
               RowHeightMin    =   300
               BackColorBkg    =   16777215
               ScrollBars      =   2
               FormatString    =   "     |^QTY.  |^UNIT   |<PART NUMBER             |<DESCRIPTION                    |>UNIT PRICE   |>TOTAL PRICE "
            End
            Begin VB.Label lblRemarks 
               Height          =   255
               Left            =   240
               TabIndex        =   87
               Tag             =   "xyf"
               Top             =   3400
               Width           =   3015
            End
         End
         Begin VB.Frame FraPayment 
            Height          =   1380
            Left            =   4680
            TabIndex        =   45
            Tag             =   "XYF"
            Top             =   1920
            Width           =   3615
            Begin VB.TextBox txtAmtChange 
               Alignment       =   1  'Right Justify
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
               Locked          =   -1  'True
               TabIndex        =   9
               Tag             =   "XYF"
               Top             =   885
               Width           =   1695
            End
            Begin VB.TextBox txtAmtRec 
               Alignment       =   1  'Right Justify
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
               TabIndex        =   8
               Tag             =   "XYF"
               Top             =   540
               Width           =   1695
            End
            Begin VB.TextBox txtPayable 
               Alignment       =   1  'Right Justify
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
               Locked          =   -1  'True
               TabIndex        =   7
               Tag             =   "XYF"
               Top             =   180
               Width           =   1695
            End
            Begin VB.Label Label11 
               Caption         =   "CHANGE"
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
               TabIndex        =   48
               Tag             =   "XYF"
               Top             =   915
               Width           =   1335
            End
            Begin VB.Label Label10 
               Caption         =   "AMT. RECEIVED"
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
               TabIndex        =   47
               Tag             =   "XYF"
               Top             =   570
               Width           =   1335
            End
            Begin VB.Label Label9 
               Caption         =   "TOTAL PAYABLE"
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
               TabIndex        =   46
               Tag             =   "XYF"
               Top             =   255
               Width           =   1575
            End
         End
         Begin VB.Frame FraCheck 
            Height          =   1475
            Left            =   4680
            TabIndex        =   29
            Tag             =   "XYF"
            Top             =   7560
            Visible         =   0   'False
            Width           =   3495
            Begin VB.TextBox txtChkNo 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1200
               Locked          =   -1  'True
               TabIndex        =   34
               Tag             =   "XYF"
               Top             =   120
               Width           =   2055
            End
            Begin VB.TextBox txtChkDate 
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
               Left            =   1200
               Locked          =   -1  'True
               TabIndex        =   33
               Tag             =   "XYF"
               Top             =   440
               Width           =   2055
            End
            Begin VB.TextBox txtChkAmt 
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
               Left            =   1200
               Locked          =   -1  'True
               TabIndex        =   32
               Tag             =   "XYF"
               Top             =   760
               Width           =   2055
            End
            Begin VB.TextBox txtChkBank 
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
               Left            =   1200
               Locked          =   -1  'True
               TabIndex        =   31
               Tag             =   "XYF"
               Top             =   1100
               Width           =   2055
            End
            Begin VB.ComboBox cboChecks 
               Height          =   315
               Left            =   1200
               TabIndex        =   30
               Top             =   1680
               Visible         =   0   'False
               Width           =   2055
            End
            Begin VB.Label lblChk 
               AutoSize        =   -1  'True
               Caption         =   "CHECK NO.  "
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
               Index           =   3
               Left            =   120
               TabIndex        =   39
               Tag             =   "XYF"
               Top             =   135
               Width           =   915
            End
            Begin VB.Label lblChk 
               AutoSize        =   -1  'True
               Caption         =   "CHECK DATE   "
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
               Index           =   0
               Left            =   120
               TabIndex        =   38
               Tag             =   "XYF"
               Top             =   465
               Width           =   1095
            End
            Begin VB.Label lblChk 
               AutoSize        =   -1  'True
               Caption         =   "AMOUNT   "
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
               Index           =   1
               Left            =   120
               TabIndex        =   37
               Tag             =   "XYF"
               Top             =   795
               Width           =   795
            End
            Begin VB.Label lblChk 
               AutoSize        =   -1  'True
               Caption         =   "BANK   "
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
               Index           =   2
               Left            =   120
               TabIndex        =   36
               Tag             =   "XYF"
               Top             =   1125
               Width           =   570
            End
            Begin VB.Label lblDisable 
               Caption         =   "Label3"
               ForeColor       =   &H80000011&
               Height          =   255
               Left            =   240
               TabIndex        =   35
               Top             =   1680
               Visible         =   0   'False
               Width           =   735
            End
         End
         Begin VB.Frame FraHeading 
            Height          =   1815
            Left            =   120
            TabIndex        =   17
            Tag             =   "XYF"
            Top             =   120
            Width           =   8175
            Begin VB.ComboBox cboAcctCust 
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
               Left            =   1680
               TabIndex        =   62
               Tag             =   "XYF"
               Top             =   600
               Visible         =   0   'False
               Width           =   3015
            End
            Begin MSComCtl2.UpDown updRef 
               Height          =   255
               Left            =   1440
               TabIndex        =   59
               Tag             =   "XYF"
               Top             =   270
               Visible         =   0   'False
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   450
               _Version        =   393216
               OrigLeft        =   1440
               OrigTop         =   270
               OrigRight       =   1695
               OrigBottom      =   525
               Max             =   2
               Enabled         =   -1  'True
            End
            Begin VB.TextBox Text2 
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
               Left            =   6360
               Locked          =   -1  'True
               TabIndex        =   56
               Tag             =   "XYF"
               Top             =   1320
               Visible         =   0   'False
               Width           =   1695
            End
            Begin VB.TextBox Text1 
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
               Left            =   6360
               Locked          =   -1  'True
               TabIndex        =   55
               Tag             =   "XYF"
               Top             =   960
               Visible         =   0   'False
               Width           =   1695
            End
            Begin MSMask.MaskEdBox txtDate 
               Height          =   315
               Left            =   6360
               TabIndex        =   4
               Tag             =   "XYF"
               Top             =   600
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   556
               _Version        =   393216
               MaxLength       =   10
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Mask            =   "##/##/####"
               PromptChar      =   "_"
            End
            Begin VB.TextBox txtOrderNo 
               Alignment       =   2  'Center
               Height          =   315
               Left            =   1680
               TabIndex        =   1
               Tag             =   "XYF"
               Top             =   240
               Width           =   1095
            End
            Begin VB.TextBox txtSoldTo 
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
               Left            =   1680
               TabIndex        =   3
               Tag             =   "XYF"
               Top             =   600
               Width           =   3015
            End
            Begin VB.TextBox txtAddress 
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
               Left            =   1680
               TabIndex        =   5
               Tag             =   "XYF"
               Top             =   960
               Width           =   3015
            End
            Begin VB.TextBox txtInvNo 
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
               Left            =   6360
               TabIndex        =   2
               Tag             =   "XYF"
               Top             =   240
               Width           =   1695
            End
            Begin VB.TextBox txtPartsMan 
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
               Left            =   1680
               Locked          =   -1  'True
               TabIndex        =   6
               Tag             =   "XYF"
               Top             =   1320
               Width           =   3015
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
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
               Left            =   4920
               TabIndex        =   58
               Tag             =   "XYF"
               Top             =   1372
               Width           =   1335
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
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
               Left            =   4920
               TabIndex        =   57
               Tag             =   "XYF"
               Top             =   1012
               Width           =   1335
            End
            Begin VB.Label lblSoldTo 
               Caption         =   "SOLD TO"
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
               TabIndex        =   23
               Tag             =   "XYF"
               Top             =   630
               Width           =   1095
            End
            Begin VB.Label lblOrderNo 
               Caption         =   "SALES ORDER#"
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
               TabIndex        =   22
               Tag             =   "XYF"
               Top             =   270
               Width           =   1335
            End
            Begin VB.Label lblAddress 
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
               Height          =   255
               Left            =   120
               TabIndex        =   21
               Tag             =   "XYF"
               Top             =   960
               Width           =   1095
            End
            Begin VB.Label lblDate 
               Alignment       =   1  'Right Justify
               Caption         =   "DATE"
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
               Left            =   4920
               TabIndex        =   20
               Tag             =   "XYF"
               Top             =   652
               Width           =   1335
            End
            Begin VB.Label lblInvNo 
               Alignment       =   1  'Right Justify
               Caption         =   "SALES INVOICE NO."
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
               Left            =   4440
               TabIndex        =   19
               Tag             =   "XYF"
               Top             =   270
               Width           =   1815
            End
            Begin VB.Label lblPartsMan 
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
               Left            =   120
               TabIndex        =   18
               Tag             =   "XYF"
               Top             =   1350
               Width           =   1215
            End
         End
         Begin ResizeXtra.RXResizer RXResizer1 
            Left            =   720
            Top             =   7560
            _ExtentX        =   847
            _ExtentY        =   847
            RegistryKey     =   "RXB49E"
            Registered      =   -1  'True
            m_MDI_Width     =   12105
            m_MDI_Height    =   8670
         End
      End
   End
   Begin VB.Timer MoveUp 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3000
      Top             =   3000
   End
   Begin VB.Timer HideBtn 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3000
      Top             =   3480
   End
   Begin VB.Timer ReArrange 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3000
      Top             =   3960
   End
   Begin VB.TextBox txtEmpID 
      Height          =   285
      Left            =   120
      TabIndex        =   13
      Top             =   8160
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblHeading 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "SALES INVOICE"
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
      Left            =   9240
      TabIndex        =   11
      Top             =   90
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
      Left            =   120
      TabIndex        =   0
      Top             =   60
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
      Height          =   375
      Index           =   1
      Left            =   150
      TabIndex        =   10
      Top             =   90
      Width           =   2900
   End
   Begin VB.Label lblHeading 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "SALES INVOICE"
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
      Left            =   9225
      TabIndex        =   12
      Top             =   120
      Width           =   2415
   End
   Begin VB.Menu mnuPayment 
      Caption         =   "Payment"
      Visible         =   0   'False
   End
End
Attribute VB_Name = "Invoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsOnAcct        As New ADODB.Recordset
Dim myInvType       As String
Dim ActiveBtn       As Integer
Dim PayType         As Integer
Dim CurrentBtn      As Integer
Dim CmdTop(0 To 3)  As Integer
Dim CustSrch        As Boolean
Dim NoPrice         As Boolean
Dim Cashr_Idx       As Integer
Dim InvldOrd        As Boolean
Dim wCost           As Variant

Private Sub ClearEntries()
    'clear content of tabs
    txtOrderNo.Text = ""
    txtSoldTo.Text = ""
    txtAddress.Text = ""
    txtPartsMan.Text = ""
    txtInvNo.Text = ""
'    txtDate.Text = Format$(Date, "mm/dd/yyyy")
    txtPayable.Text = ""
    txtAmtRec.Text = ""
    txtAmtChange.Text = ""
    'txtEmpID.Text = ""
    'flxOrder.Rows = 1
    'flxOrder.Rows = 2
    'Clear Terms
    txtTerm.Text = ""
    txtAmtDue.Text = ""
    txtPayStart.Text = ""
    txtTotal.Text = ""
    cboAcctCust.Text = ""
    cboAcctCust.Clear
    txtDueAmt.Text = ""
    
    Label1.Caption = ""
    Text1.Text = ""
    Text1.Visible = False
    Label2.Caption = ""
    Text2.Text = ""
    Text2.Visible = False
    
    txtORno.Text = ""
    txtORamt.Text = ""
    txtORdate.Text = ""
    lblORno.ToolTipText = ""
    
    'chkOnLine.Visible = False
    'chkOnLine.Value = 0
    NoPrice = False
    
    flxOrder.Rows = 1
    txtCType.Text = ""
    
    txtPScheme.Text = ""
    txtNTerms.Text = ""
    txtNPay.Text = ""
    txtDueAmt.Text = ""
    txtPayStart.Text = ""
    txtTerm.Text = ""
    txtAmtDue.Text = ""
    txtJobID.Text = ""
    
    cboPayType.Text = ""
    
End Sub


Private Sub LockData(bln As Boolean)
    Dim i As Integer
    
    updRef.Enabled = Not bln
    cboAcctCust.Locked = bln
    txtOrderNo.Locked = bln
    txtSoldTo.Locked = bln
    txtAddress.Locked = bln
    txtPartsMan.Locked = bln
    txtInvNo.Locked = bln
    FraPayment.Enabled = Not bln
    tabPreview.Enabled = Not bln
End Sub
Public Sub ShowPic()
    On Error Resume Next
    Dim rsORinChrg As New ADODB.Recordset
    Qry = "SELECT EMPLOYEES.* From EMPLOYEES " & _
          "WHERE EMPLOYEES.EMPID=" & UserID & ";"
    If rsORinChrg.State = adStateOpen Then rsORinChrg.Close
    rsORinChrg.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
   
    If Not IsNull(rsORinChrg.Fields("EmpPicture")) Then
        Open "pictemp" For Binary Access Write As lngDataFile
            lngLengh = rsORinChrg.Fields("EmpPicture").ActualSize
            intChunks = lngLengh \ ChunkSize
            intFragment = lngLengh Mod ChunkSize
            ReDim Chunk(intFragment)
            Chunk() = rsORinChrg.Fields("EmpPicture").GetChunk(intFragment)
            Put lngDataFile, , Chunk()
            For i = 1 To intChunks
               ReDim Buffer(ChunkSize)
                Chunk() = rsORinChrg.Fields("EmpPicture").GetChunk(ChunkSize)
                Put lngDataFile, , Chunk()
            Next i
        Close lngDataFile
        
        filename = "pictemp"
        imgOptr.Picture = LoadPicture(filename)
    Else
        ImgEmp.Picture = LoadPicture("")
    End If
End Sub

Private Sub cboAcctCust_Change()
    Dim i As Long, j As Long
    Dim strPartial As String, strTotal As String
    Dim x As Integer
    
    If CustSrch Then
        CustSrch = False
        Exit Sub
    End If
    With cboAcctCust
        strPartial = .Text
        i = SendMessage(.hwnd, CB_FINDSTRING, -1, ByVal strPartial) 'If match found, append unmatched characters
        If i <> -1 Then
            strTotal = .List(i)
            rsOnAcct.MoveFirst
            rsOnAcct.Move i
            txtCustomerID.Text = rsOnAcct!CUSTOMERID
            txtOrderNo.Text = rsOnAcct!ORDERNO
            j = Len(strTotal) - Len(strPartial)
            If j <> 0 Then
                CustSrch = True
                .SelText = Right$(strTotal, j)
                .SelStart = Len(strPartial)
                .SelLength = j
            End If
        Else
            txtCustomerID.Text = ""
            txtOrderNo.Text = ""
        End If
    End With
    If updRef.Value <> 0 Then updRef.Value = 0     'make sure nka Order# mode
    
End Sub

Private Sub cboAcctCust_Click()
    Call cboAcctCust_Change
End Sub

Private Sub cboAcctCust_GotFocus()
   Dim Tmp
    
   Const CB_SHOWDROPDOWN = &H14F
   Tmp = SendMessage(cboAcctCust.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub

Private Sub cboAcctCust_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then CustSrch = True
End Sub

Private Sub cboAcctCust_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        'Call cboAcctCust_LostFocus
        txtAddress.SetFocus
    End If
End Sub

Private Sub cboAcctCust_LostFocus()
    If txtCustomerID.Text <> "" And txtOrderNo.Text <> "" Then
        Call DisplayOrder
    End If
End Sub


Private Sub cbocashier_Click()
Dim Ans As Integer
    'Previous Code
    'If cboCashier.ListIndex = -1 Then Exit Sub
    'txtEmpID.Text = cboCashier.ItemData(cboCashier.ListIndex)
    
    If cboCashier.ListIndex = -1 Then Exit Sub
    If UserID <> cboCashier.ItemData(cboCashier.ListIndex) Then
        Ans = MsgBox("You want to Change Cashier`s Name?", vbQuestion + vbYesNo + vbDefaultButton2, "Cashier")
        If Ans = vbYes Then
            txtEmpID.Text = cboCashier.ItemData(cboCashier.ListIndex)
        Else
            cboCashier.ListIndex = Cashr_Idx
            txtEmpID.Text = cboCashier.ItemData(cboCashier.ListIndex)
        End If
    Else
            txtEmpID.Text = cboCashier.ItemData(cboCashier.ListIndex)
    End If
End Sub

Private Sub cmdCancel_Click()
    Dim Ans As Integer
    
    Ans = MsgBox("Do you want to cancel payment?", vbQuestion + vbYesNo, "Cancel")
    If Ans = vbNo Then Exit Sub
    Frame2.Visible = False
    tabTerm.Visible = False
    cmdQuit.Visible = True
    cmdCashier(ActiveBtn).Enabled = True
    cboAcctCust.Visible = False
    txtSoldTo.Visible = True
    Call ClearEntries
    flxOrder.Rows = 1
    flxOrder.Rows = 2
    updRef.Visible = False
    lblOrderNo.Caption = "SALES ORDER#"
    Call LockData(True)
    Call BtnDisplayAll(True)
End Sub

Private Sub BtnDisplayAll(bln As Boolean)
    Dim i As Integer
    
    If bln Then
        For i = 0 To 3
            cmdCashier(i).Visible = True
            cmdCashier(i).Enabled = True
            If i = 0 Then
                cmdCashier(0).Top = 240
            Else
                cmdCashier(i).Top = 240 + cmdCashier(i).Height * i
            End If
        Next i
        Frame2.Visible = False
        cmdQuit.Visible = True
        cmdCashier(3).Visible = False   'Others button for reservation
    Else
        For i = 0 To 3
            cmdCashier(i).Visible = False
        Next i
        Frame2.Top = 3360
        Frame2.Visible = True
        cmdQuit.Visible = False
    End If
End Sub

Private Sub cmdCashier_Click(Index As Integer)
    Dim Tmp
    Dim i As Integer
    Dim Qry As String, CustName As String, MyAcc As String
    Const CB_SHOWDROPDOWN = &H14F
    '***Access
    'Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE (((USERACCESS.EMPID)=" & UserID & ") AND ((USERACCESS.MENUINDX)=44));"
    'mySQl
    Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE USERACCESS.EMPID=" & UserID & " AND USERACCESS.MENUINDX=44;"
    
    MyAcc = GetFieldValue(Qry)
    If MyAcc <> "FULL ACCESS" Then
        MsgBox "You don't have full access to this transaction." & vbCrLf & "Please contact your System Administrator", vbExclamation + vbOKOnly, "Access Denied"
        Exit Sub
    End If
    
    ActiveBtn = Index
    For i = 0 To 3
        cmdCashier(i).Visible = False
    Next i
    Frame2.Top = 3360
    Frame2.Visible = True
    cmdQuit.Visible = False
    cmdCashier(Index).Top = 240
    cmdCashier(Index).Visible = True
    cmdCashier(Index).Enabled = False
    Call ClearEntries
    cboPayType.Visible = True
    txtInvNo = GetInvoiceNo
    flxOrder.Rows = 1
    flxOrder.Rows = 2
    'txtDate.Text = Format$(Date, "mm/dd/yyyy")
    Select Case Index
        Case 0  'Walk In
            txtSoldTo.Visible = True
            cboAcctCust.Visible = False
            txtSoldTo.Text = "CASH"
            If SysBranch = "MOTO-R REPUBLIC CORP." Or SysBranch = "MOTO-R VAN SELLING" Then
                txtAddress.Text = "Tagum City"
            Else
                txtAddress.Text = "Davao City"
            End If
            
            'txtAddress.Text = "Tagum City"  '"Davao City"  '
            Call LockData(False)
            txtOrderNo.SetFocus
            myInvType = "CASH"
            lblInvNo.Caption = "SALES INVOICE NO."
            tabTerm.Visible = False
            txtAmtRec.Enabled = True
            txtAmtChange.Enabled = True
            lblRemarks.Caption = ""
            'If InStr(SysBranch, "PANABO") > 0 Then
            '    Qry = "SELECT MAX(Abs(spinvoice.SPINVNO)) As FIELD_1 From spinvoice;"
            '    txtInvNo.Text = Format(Val(GetFieldValue(Qry)) + 1, "#0000")
            'End If
                
            'chkOnLine.Visible = True
            'chkOnLine.Value = 0
            
            txtPCCardNo.Text = ""
            txtPCType.Text = ""
            txtPCPoints.Text = ""
        Case 1    'Quick Service
            txtSoldTo.Visible = True
            cboAcctCust.Visible = False
            Call LockData(False)
            updRef.Visible = True
            updRef.Max = 2
            updRef.Value = 0
            lblOrderNo.Caption = "REQUISITION#"
            lblInvNo.Caption = "CHARGE INV. NO."
            txtOrderNo.SetFocus
            myInvType = "CREDIT"
            txtAmtRec.Enabled = False
            txtAmtChange.Enabled = False
            tabTerm.Visible = False
            lblRemarks.Caption = ""
        Case 2    'ON account
            If Not IsDate(txtDate.Text) Then
                MsgBox "Invalid Date Format!", vbExclamation + vbOKOnly, "Error"
                Exit Sub
            End If
            '***Access
            'Qry = "SELECT CUSTOMER.CUSTOMERNAME, SPSALESORDER.ORDERNO, CUSTOMER.CUSTOMERID, SPSALESORDER.ORDERDATE " & _
                  "FROM CUSTOMER INNER JOIN SPSALESORDER ON CUSTOMER.CUSTOMERID = SPSALESORDER.CUSTOMERID " & _
                  "Where (((SPSALESORDER.ORDERDATE) = #" & txtDate.Text & "#) And ((SPSALESORDER.INVOICENO) Is Null)) ORDER BY CUSTOMER.CUSTOMERNAME;"
            'mySQL
            Qry = "SELECT CUSTOMER.CUSTOMERNAME, SPSALESORDER.ORDERNO, CUSTOMER.CUSTOMERID, SPSALESORDER.ORDERDATE " & _
                  "FROM CUSTOMER INNER JOIN SPSALESORDER ON CUSTOMER.CUSTOMERID = SPSALESORDER.CUSTOMERID " & _
                  "Where SPSALESORDER.ORDERDATE = '" & Format(txtDate.Text, "yyyy-mm-dd") & "' And (SPSALESORDER.INVOICENO) Is Null ORDER BY CUSTOMER.CUSTOMERNAME;"
            With rsOnAcct
                If .State = adStateOpen Then .Close
                .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
                If .RecordCount > 0 Then
                    .MoveFirst
                    cboAcctCust.Clear
                    Do While Not .EOF
                        cboAcctCust.AddItem !CustomerName
                        .MoveNext
                    Loop
                Else
                    cboAcctCust.Clear
                End If
            End With
            Call LockData(False)
            updRef.Visible = True
            updRef.Max = 1
            updRef.Value = 0
            txtInvNo = GetInvoiceNoCredit
            lblOrderNo.Caption = "SALES ORDER#"
            lblInvNo.Caption = "CHARGE INV. NO."
            txtSoldTo.Visible = False
            cboAcctCust.Visible = True
            txtOrderNo.SetFocus
            myInvType = "CREDIT"
            txtAmtRec.Enabled = False
            txtAmtChange.Enabled = False
            tabTerm.Visible = True
            cboPayType.Visible = False
            lblRemarks.Caption = ""
            txtPCCardNo.Text = ""
            txtPCType.Text = ""
            txtPCPoints.Text = ""
        Case 3    'Others
        
    End Select
    PayType = Index
End Sub

Function GetInvoiceNo() As String
    Dim rsInvNo As New ADODB.Recordset
    Dim Qry As String
    Dim MyNo As Long
    
    '***Access
    'Qry = "SELECT Max(Val([SPINVNO])) AS MaxNo FROM SPINVOICE;"
    'mySQl
    'Qry = "SELECT Max(round(SPINVOICE.SPINVNO)) AS MaxNo FROM SPINVOICE WHERE SPINVOICE.INVDATE = '" & Format(Date, "YYYY-MM-DD") & "' AND SPINVOICE.INVTYPE = 'CASH';"
    Qry = "SELECT s.SPINVNO AS MaxNo FROM SPINVOICE s WHERE s.lastupdate = (SELECT MAX(s2.lastupdate) FROM SPINVOICE s2) AND s.INVTYPE = 'CASH' AND S.INVDATE = '" & Format(Date, "YYYY-MM-DD") & "';"
    With rsInvNo
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            MyNo = IIf(IsNull(.Fields(0)), 1, Val(Replace(Replace(.Fields(0), getStoreCode(sysBranchCode) & "-", ""), " ", "")) + 1)
        Else
            MyNo = 0
        End If
        .Close
    End With
    GetInvoiceNo = getStoreCode(sysBranchCode) & "-" & MyNo
    
End Function

Public Function getStoreCode(brcode As String) As String
    
    Dim Qry As String
    Qry = "select s.storecode from branch s where branchcode = '" & brcode & "'"
     getStoreCode = GetFieldValue(Qry)
'    If brcode = "LMPC" Then        'COMPOSTELA
'        getStoreCode = "F010"
'    ElseIf brcode = "LMPB" Then    'BAYUGAN
'        getStoreCode = "F006"
'    ElseIf brcode = "MST" Then     'SOBRECAREY
'        getStoreCode = "F014"
'    ElseIf brcode = "LMPD" Then    'DIGOS
'        getStoreCode = "F016"
'    ElseIf brcode = "MP" Then      'PANABO
'        getStoreCode = "F013"
'    ElseIf brcode = "MLA" Then     'AGDAO - LAPULAPU
'        getStoreCode = "F008"
'    ElseIf brcode = "MRT" Then     'GANTE
'        getStoreCode = "F007"
'    ElseIf brcode = "MCD" Then      'CALINAN
'        getStoreCode = "F011"
'    ElseIf brcode = "LMSF" Then     'SAN FRANS
'        getStoreCode = "F012"
'    ElseIf brcode = "MUD" Then      'ULAS
'        getStoreCode = "F004"
'    ElseIf brcode = "MG" Then       'GENSAN
'        getStoreCode = "F009"
'    ElseIf brcode = "MCPR" Then     'RIZAL
'        getStoreCode = "F018"
'    ElseIf brcode = "MPT" Then      'TORIL
'        getStoreCode = "F017"
'    ElseIf brcode = "MMK" Then      'MONKAYO
'        getStoreCode = "F009"
'    End If
 
End Function

Function GetInvoiceNoCredit() As String
    Dim rsInvNo As New ADODB.Recordset
    Dim Qry As String
    Dim MyNo As Long
    
    '***Access
    'Qry = "SELECT Max(Val([SPINVNO])) AS MaxNo FROM SPINVOICE;"
    'mySQl
    'Qry = "SELECT Max(round(SPINVOICE.SPINVNO)) AS MaxNo FROM SPINVOICE WHERE SPINVOICE.INVDATE = '" & Format(Date, "YYYY-MM-DD") & "' AND SPINVOICE.INVTYPE = 'CREDIT';"
    Qry = "SELECT s.SPINVNO AS MaxNo FROM SPINVOICE s WHERE s.lastupdate = (SELECT MAX(s2.lastupdate) FROM SPINVOICE s2) AND s.INVTYPE = 'CREDIT' AND S.INVDATE = '" & Format(Date, "YYYY-MM-DD") & "';"
    With rsInvNo
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            MyNo = IIf(IsNull(.Fields(0)), 1, Replace(.Fields(0), getStoreCode(sysBranchCode) & "-", "") + 1)
        Else
            MyNo = 0
        End If
        .Close
    End With
    'GetInvoiceNoCredit = Format(MyNo, "00000")
    GetInvoiceNoCredit = getStoreCode(sysBranchCode) & "-" & MyNo
End Function
Private Sub cmdCashier_GotFocus(Index As Integer)
    cmdCashier(Index).ForeColor = vbRed
    cmdCashier(Index).Font.Bold = True
End Sub

Private Sub cmdCashier_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        If Index = 6 Then
            cmdQuit.SetFocus
        Else
            cmdCashier(Index + 1).SetFocus
        End If
    ElseIf KeyCode = vbKeyUp Then
        If Index = 0 Then
            cmdQuit.SetFocus
        Else
            cmdCashier(Index - 1).SetFocus
        End If
    End If
End Sub

Private Sub cmdCashier_LostFocus(Index As Integer)
    cmdCashier(Index).ForeColor = vbBlack
    cmdCashier(Index).Font.Bold = False
End Sub

Private Sub cmdQuit_Click()
    'MotoRep.tbMain.Visible = True
    ''MotoRep.imgBG.Visible = True
    MotoRep.toolBarButton.Visible = True
    Unload Me
End Sub

Private Sub cmdQuit_GotFocus()
    cmdQuit.ForeColor = vbRed
    cmdQuit.Font.Bold = True
End Sub

Private Sub cmdQuit_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Then
        cmdCashier(6).SetFocus
    ElseIf KeyCode = cbkeydown Then
        cmdCashier(0).SetFocus
    End If
End Sub

Private Sub cmdQuit_LostFocus()
    cmdQuit.ForeColor = vbBlack
    cmdQuit.Font.Bold = False
End Sub

Private Sub cmdSave_Click()
    Dim Ans As Integer, i As Integer
    Dim myMsg As String, MsgDone As String
        
        
    If InvldOrd = True Then MsgBox "Check some items Quantity not Greater than Available Stock!" & vbCrLf & "Please inform your Manager to update the order!", vbExclamation + vbOKOnly, "Warning": Exit Sub
    'Walkin
    If PayType = 0 Then
        myMsg = ValidateWalkIn
        MsgDone = "Payment Successfully Saved!"
    Else
    'Charge Invoice
        myMsg = ValidateChargeInvoice
        MsgDone = "Charge Invoice Successfully Saved!"
    End If
    If myMsg <> "" Then
        MsgBox myMsg, vbCritical + vbOKOnly, "Validation Error"
        txtOrderNo.SetFocus
        Exit Sub
    End If
    Ans = MsgBox("Do you want to save invoice?", vbQuestion + vbYesNo)
    If Ans = vbNo Then
        txtOrderNo.SetFocus
        Exit Sub
    End If
    Call SaveInvoice(MsgDone)
    Call LockData(True)
    'Call tbSaveCancel(False)
    
    Call BtnDisplayAll(True)
    cmdQuit.Visible = True
    cmdQuit.SetFocus
End Sub

Private Sub Form_Load()
    Dim rsCashier As New ADODB.Recordset
    Dim rsOthers As New ADODB.Recordset
    Dim i As Integer, MyID As Integer
    Dim Qry As String
    
    flxOrder.ToolTipText = "v15.7.11 wCst"
     
    For i = 0 To 3
        CmdTop(i) = cmdCashier(i).Top
    Next i
    BtnOnHold = False
    txtDate.Text = Format(Date, "mm/dd/yyyy")
    
    cboPayType.AddItem "CREDIT CARD PAYMENT"
    cboPayType.AddItem "INTER-BRANCH PAYMENT"
    cboPayType.AddItem "ON-LINE PAYMENT"
    
    MyID = -1
    Qry = "SELECT CONCAT(employees.LASTNAME, ', ', employees.FIRSTNAME) AS CASHIERS, Employees.EmpId " & _
          "From Employees Where employees.DEPARTMENT = 'ADMIN' Order By CASHIERS"
    With rsCashier
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                cboCashier.AddItem !Cashiers
                cboCashier.ItemData(cboCashier.ListCount - 1) = !EMPID
                'If !EmpId = UserID Then MyID = cboCashier.ListCount - 1
                'Add Original
                If !EMPID = UserID Then MyID = cboCashier.ListCount - 1: Cashr_Idx = cboCashier.ListCount - 1
                .MoveNext
            Loop
        End If
        .Close
    End With
    
    If cboCashier.ListCount > 0 And MyID >= 0 Then
        cboCashier.ListIndex = MyID
    End If
    Call LockData(True)
    Call ShowPic
    
    MotoRep.sBar.Panels(6).Text = "Process Customer Sales Invoice"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If rsOnAcct.State = adStateOpen Then rsOnAcct.Close
    MotoRep.sBar.Panels(6).Text = ""
End Sub



Private Sub RXResizer1_ResizeComplete()
    Dim i As Integer
                                  
    RXSplitter1.PanelASize = 3063
    lblHeading(2).Left = Me.Width - (lblHeading(2).Width + 200)
    lblHeading(3).Left = Me.Width - (lblHeading(3).Width + 200)
    With flxOrder
        .ColWidth(0) = 300
        .ColWidth(1) = (.Width - 300) * 0.1
        .ColWidth(2) = (.Width - 300) * 0.09
        .ColWidth(3) = (.Width - 300) * 0.25
        .ColWidth(4) = (.Width - 300) * 0.25
        .ColWidth(5) = (.Width - 300) * 0.15
        .ColWidth(6) = (.Width - 300) * 0.15
        .ColWidth(7) = 0
        txtTotal.Width = .ColWidth(6)
        txtTotal.Top = .Top + .Height
        txtTotal.Left = (.Left + .Width) - txtTotal.Width
    End With
    
End Sub

Private Sub txtAddress_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        SendKeys "{TAB}"
    ElseIf KeyCode = vbKeyUp Then
        SendKeys "+{TAB}"
    End If
End Sub

Private Sub txtAddress_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtAmtChange_KeyPress(KeyAscii As Integer)
    Dim strValid As String
    
    strValid = "1234567890"
    
    If KeyAscii > 26 Then
        If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
    If KeyAscii = 13 Then
        Call cmdSave_Click
    End If
End Sub

Private Sub txtAmtRec_KeyPress(KeyAscii As Integer)
    Dim strValid As String
    
    strValid = "1234567890."
    If KeyAscii > 26 Then
        If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
    If KeyAscii = 13 Then
        Call txtAmtRec_LostFocus
        If txtAmtChange.Text <> "" Then SendKeys "{TAB}"
    End If
End Sub

Private Sub txtAmtRec_LostFocus()
    If Val(txtPayable.Text) > Val(txtAmtRec.Text) Then
        MsgBox "Insufficient Amount Received!", vbInformation + vbOKOnly, "Invoice"
        txtAmtChange.Text = ""
        txtAmtRec.SelStart = 0
        txtAmtRec.SelLength = Len(txtAmtRec.Text)
        txtAmtRec.SetFocus
        Exit Sub
    End If
    If txtAmtRec.Text <> "" Then txtAmtRec.Text = Format$(txtAmtRec.Text, "#####0.00")
    txtAmtChange.Text = Format(Val(txtAmtRec.Text) - Val(txtPayable.Text), "#####0.00")
End Sub

Private Sub txtChkAmt_KeyPress(KeyAscii As Integer)
    Dim strValid As String
    
    strValid = "1234567890."
    
    If KeyAscii > 26 Then
        If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub txtDate_GotFocus()
    txtDate.SelStart = 0
    txtDate.SelLength = Len(txtDate.Text)
End Sub

Private Sub txtDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtInvNo_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtInvNo, KeyAscii, ValidChars)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtInvNo_LostFocus()
    Dim MyNo As String
    If txtInvNo.Text <> "" Then
        '***Access
        'MyNo = "SELECT SPINVOICE.SPINVNO From SPINVOICE WHERE (((SPINVOICE.SPINVNO)='" & txtInvNo.Text & "'));"
        'mySQL
        If myInvType = "CASH" Then
            MyNo = "SELECT SPINVOICE.SPINVNO From SPINVOICE WHERE spinvoice.INVTYPE = 'CASH' AND SPINVOICE.SPINVNO='" & Trim(txtInvNo.Text) & "';"
        Else
            MyNo = "SELECT SPINVOICE.SPINVNO From SPINVOICE WHERE spinvoice.INVTYPE = 'CREDIT' AND SPINVOICE.SPINVNO='" & Trim(txtInvNo.Text) & "';"
        End If
        If CtrlNumExist(MyNo) Then
            MsgBox "Invoice No. " & txtInvNo.Text & " already exist!", vbExclamation + vbOKOnly, "Invoice Exist"
            txtInvNo.Text = ""
            txtInvNo.SetFocus
        End If
    End If
End Sub

Private Sub txtOrderNo_KeyPress(KeyAscii As Integer)
    Dim Qry As String, strValid As String
    
    'Add code module 04-24-2010
    strValid = "1234567890."
    If KeyAscii > 26 Then
        If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
        
    
    
    If KeyAscii = 13 Then
        If txtOrderNo.Text = "" Then
            If PayType = 2 Then cboAcctCust.SetFocus
            Exit Sub
        End If
        'Validate PO number here!           to be continued
        If PayType = 2 And updRef.Value = 1 Then 'Check if PO is valid
            '***Access
            Qry = "SELECT CUSTPO.* From CUSTPO WHERE (((CUSTPO.CPONO)='" & txtOrderNo.Text & "') AND ((CUSTPO.INVOICENO) Is Null) AND ((CUSTPO.REMARKS) Not Like '%CANCELLED%'));"
            If Not CtrlNumExist(Qry) Then
                MsgBox "Invalid P.O. number! Its either Cancelled or already in used", vbExclamation + vbOKOnly, "Invalid P.O."
                Call ClearEntries
                Exit Sub
            End If
        End If
        Call DisplayOrder
        If PayType = 2 Then
            txtPayable.SetFocus
        Else
            txtSoldTo.SetFocus
        End If
    End If
End Sub


Private Sub txtORno_KeyPress(KeyAscii As Integer)
    Dim strValid As String
    
    strValid = "1234567890"
    If KeyAscii > 26 Then
        If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
    If KeyAscii = 13 Then txtORamt.SetFocus
End Sub

Private Sub txtORno_LostFocus()
    Dim rsOR As New ADODB.Recordset
    Dim rsLedger As New ADODB.Recordset
    Dim Qry As String, vMsg As String, myRem As String
    Dim ORttlamt As Single
    
    If txtORno.Text = "" Then Exit Sub
    vMsg = ""
    '***Access
    'Qry = "SELECT DPLEDGER.* From DPLEDGER WHERE (((DPLEDGER.ORNO)='" & txtORno.Text & "'));"
    'mySQL
    Qry = "SELECT DPLEDGER.* From DPLEDGER WHERE DPLEDGER.ORNO='" & txtORno.Text & "';"
    With rsOR
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            myRem = IIf(IsNull(!STATUS), "", !STATUS)
            If myRem = "OPEN" Then
                txtORdate.Text = Format(!ORdate, "mm/dd/yyyy")
                txtORamt.Text = Format(!Balance, "#,##0.00")
            ElseIf myRem = "CLOSED" Then
                vMsg = "O.R.# " & txtORno.Text & " already in used!"
                Select Case !TRANSTYPE
                    Case 1  'SP
                        If Not IsNull(!InvRef) Then vMsg = vMsg & vbCrLf & "Refer to Invoice No. " & !InvRef
                    Case 2  'SP
                        If Not IsNull(!JORef) Then vMsg = vMsg & vbCrLf & "Refer to Job Order No. " & !JORef
                    Case 3  'SP & JO
                        If Not IsNull(!InvRef) Then vMsg = vMsg & vbCrLf & "Refer to Invoice No. " & !InvRef
                        If Not IsNull(!JORef) Then vMsg = vMsg & vbCrLf & "Refer to Job Order No. " & !JORef
                End Select
            ElseIf myRem = "CANCELLED" Then
                vMsg = "O.R.# " & txtORno.Text & " was Cancelled!"
            ElseIf myRem = "REFUNDED" Then
                vMsg = "O.R.# " & txtORno.Text & " was refunded!"
            Else
                vMsg = "No Status of Down Payment found in O.R.# " & txtORno.Text
            End If
        Else
            vMsg = "No Down Payment found in O.R.# " & txtORno.Text
        End If
        .Close
    End With
    If vMsg <> "" Then
        MsgBox vMsg, vbExclamation + vbOKOnly, "Down Payment"
        txtORno.Text = ""
        txtORdate.Text = ""
        txtORamt.Text = ""
    End If
End Sub

Private Sub txtPartsMan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtPayable_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub



Private Sub txtSoldTo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtPayable.SetFocus
End Sub

Private Sub DisplayOrder()
    Dim rsOrder As New ADODB.Recordset
    Dim rsOrdDtls As New ADODB.Recordset
    Dim rsTerm As New ADODB.Recordset
    Dim rsDP As New ADODB.Recordset
    Dim Qry As String, myRem As String, CQry As String
    Dim Ttl As Single, CustID As Integer
    '***Access
    'If PayType = 0 Then         'walkin / cash
    '    Qry = "SELECT SPSALESORDER.ORDERNO, SPSALESORDER.ORDERDATE, SPSALESORDER.SPSORDERID, SPSALESORDER.FRONTLINER, SPSALESORDER.EMPID, SPSALESORDER.TOTALPRICE, SPSALESORDER.CUSTOMERID, SPSALESORDER.CTYPE, SPSALESORDER.REMARKS, SPSALESORDER.PAYSCHEME, SPSALESORDER.TERMS, SPSALESORDER.NPAYS, SPSALESORDER.PAYSTART, SPSALESORDER.BILLAMT From SPSALESORDER " & _
    '          "WHERE (((SPSALESORDER.ORDERNO)=" & txtOrderno.Text & ") AND ((SPSALESORDER.ORDERDATE)=#" & txtDate.Text & "#) AND ((SPSALESORDER.INVOICENO) Is Null));"
    'ElseIf PayType = 1 And updRef.Value = 0 Then    'Quick Service w/ Requisition
    '    Qry = "SELECT SPSALESORDER.ORDERNO, SPSALESORDER.ORDERDATE, SPSALESORDER.SPSORDERID, SPSALESORDER.FRONTLINER, SPSALESORDER.EMPID, SPSALESORDER.TOTALPRICE, REQUISITION.JOBNO, REQUISITION.REQID, REQUISITION.REQNO, REQUISITION.CUSTOMER, SPSALESORDER.CUSTOMERID, SPSALESORDER.CTYPE, SPSALESORDER.REMARKS, SPSALESORDER.PAYSCHEME, SPSALESORDER.TERMS, SPSALESORDER.NPAYS, SPSALESORDER.PAYSTART, SPSALESORDER.BILLAMT " & _
    '          "FROM SPSALESORDER INNER JOIN REQUISITION ON SPSALESORDER.REQID = REQUISITION.REQID " & _
    '          "WHERE (((SPSALESORDER.INVOICENO) Is Null) AND ((REQUISITION.REQNO)='" & txtOrderno.Text & "'));"
    'ElseIf PayType = 1 And updRef.Value = 1 Then    'Quick Service using Sales Order
    '    Qry = "SELECT SPSALESORDER.ORDERNO, SPSALESORDER.ORDERDATE, SPSALESORDER.SPSORDERID, SPSALESORDER.FRONTLINER, SPSALESORDER.EMPID, SPSALESORDER.TOTALPRICE, REQUISITION.JOBNO, REQUISITION.REQID, REQUISITION.REQNO, REQUISITION.CUSTOMER, SPSALESORDER.CUSTOMERID, SPSALESORDER.CTYPE, SPSALESORDER.REMARKS, SPSALESORDER.PAYSCHEME, SPSALESORDER.TERMS, SPSALESORDER.NPAYS, SPSALESORDER.PAYSTART, SPSALESORDER.BILLAMT " & _
    '          "FROM SPSALESORDER INNER JOIN REQUISITION ON SPSALESORDER.REQID = REQUISITION.REQID " & _
    '          "WHERE (((SPSALESORDER.ORDERNO)=" & txtOrderno.Text & ") AND ((SPSALESORDER.ORDERDATE)=#" & txtDate.Text & "#) AND ((SPSALESORDER.INVOICENO) Is Null));"
    'ElseIf PayType = 1 And updRef.Value = 2 Then    'Quick Service using Job Order
    '    Qry = "SELECT SPSALESORDER.ORDERNO, SPSALESORDER.ORDERDATE, SPSALESORDER.SPSORDERID, SPSALESORDER.FRONTLINER, SPSALESORDER.EMPID, SPSALESORDER.TOTALPRICE, REQUISITION.JOBNO, REQUISITION.REQID, REQUISITION.REQNO, REQUISITION.CUSTOMER, SPSALESORDER.CUSTOMERID, SPSALESORDER.CTYPE, SPSALESORDER.REMARKS, SPSALESORDER.PAYSCHEME, SPSALESORDER.TERMS, SPSALESORDER.NPAYS, SPSALESORDER.PAYSTART, SPSALESORDER.BILLAMT " & _
    '          "FROM SPSALESORDER INNER JOIN REQUISITION ON SPSALESORDER.REQID = REQUISITION.REQID " & _
    '          "WHERE (((REQUISITION.JOBNO)='" & txtOrderno.Text & "') AND ((SPSALESORDER.INVOICENO) Is Null));"
    'ElseIf PayType = 2 And updRef.Value = 1 Then    'On Acct using P.O. number
    '    Qry = "SELECT SPSALESORDER.ORDERNO, SPSALESORDER.ORDERDATE, SPSALESORDER.SPSORDERID, SPSALESORDER.FRONTLINER, SPSALESORDER.EMPID, CUSTOMER.CUSTOMERNAME, CUSTOMER.ADDRESS, SPSALESORDER.TOTALPRICE, SPSALESORDER.REFERENCE, SPSALESORDER.CUSTOMERID, SPSALESORDER.REQID, SPSALESORDER.CTYPE, SPSALESORDER.REMARKS, SPSALESORDER.PAYSCHEME, SPSALESORDER.TERMS, SPSALESORDER.NPAYS, SPSALESORDER.PAYSTART, SPSALESORDER.BILLAMT " & _
    '          "FROM SPSALESORDER INNER JOIN CUSTOMER ON SPSALESORDER.CUSTOMERID = CUSTOMER.CUSTOMERID " & _
    '          "WHERE (((SPSALESORDER.REFERENCE)='" & txtOrderno.Text & "') AND ((SPSALESORDER.ORDERDATE)=#" & txtDate.Text & "#) AND ((SPSALESORDER.INVOICENO) Is Null)) ORDER BY SPSALESORDER.REQID DESC;"
    'ElseIf PayType = 2 And updRef.Value = 0 Then    'On Acct using Sales Order
    '    Qry = "SELECT SPSALESORDER.ORDERNO, SPSALESORDER.ORDERDATE, SPSALESORDER.SPSORDERID, SPSALESORDER.FRONTLINER, SPSALESORDER.EMPID, CUSTOMER.CUSTOMERNAME, CUSTOMER.ADDRESS, SPSALESORDER.TOTALPRICE, SPSALESORDER.REFERENCE, SPSALESORDER.CUSTOMERID, SPSALESORDER.REQID, SPSALESORDER.CTYPE, SPSALESORDER.REMARKS, SPSALESORDER.PAYSCHEME, SPSALESORDER.TERMS, SPSALESORDER.NPAYS, SPSALESORDER.PAYSTART, SPSALESORDER.BILLAMT " & _
    '          "FROM SPSALESORDER INNER JOIN CUSTOMER ON SPSALESORDER.CUSTOMERID = CUSTOMER.CUSTOMERID " & _
    '          "WHERE (((SPSALESORDER.ORDERNO)=" & txtOrderno.Text & ") AND ((SPSALESORDER.ORDERDATE)=#" & txtDate.Text & "#) AND ((SPSALESORDER.INVOICENO) Is Null)) ORDER BY SPSALESORDER.REQID DESC;"
    'End If
    
    'mySQL
    InvldOrd = False
    
    If PayType = 0 Then         'walkin / cash
        Qry = "SELECT SPSALESORDER.ORDERNO, SPSALESORDER.ORDERDATE, SPSALESORDER.SPSORDERID, SPSALESORDER.FRONTLINER, SPSALESORDER.EMPID, SPSALESORDER.TOTALPRICE, SPSALESORDER.CUSTOMERID, SPSALESORDER.CTYPE, SPSALESORDER.REMARKS, SPSALESORDER.PAYSCHEME, SPSALESORDER.TERMS, SPSALESORDER.NPAYS, SPSALESORDER.PAYSTART, SPSALESORDER.BILLAMT, SPSALESORDER.CUSTOMER, SPSALESORDER.ADDRESS, SPSALESORDER.REFERENCE From SPSALESORDER " & _
              "WHERE (((SPSALESORDER.ORDERNO)=" & txtOrderNo.Text & ") AND ((SPSALESORDER.ORDERDATE)='" & Format(txtDate.Text, "yyyy-mm-dd") & "') AND ((SPSALESORDER.INVOICENO) Is Null));"
    ElseIf PayType = 1 And updRef.Value = 0 Then    'Quick Service w/ Requisition
        Qry = "SELECT SPSALESORDER.ORDERNO, SPSALESORDER.ORDERDATE, SPSALESORDER.SPSORDERID, SPSALESORDER.FRONTLINER, SPSALESORDER.EMPID, SPSALESORDER.TOTALPRICE, REQUISITION.JOBNO, REQUISITION.REQID, REQUISITION.REQNO, REQUISITION.CUSTOMER, SPSALESORDER.CUSTOMERID, SPSALESORDER.CTYPE, SPSALESORDER.REMARKS, SPSALESORDER.PAYSCHEME, SPSALESORDER.TERMS, SPSALESORDER.NPAYS, SPSALESORDER.PAYSTART, SPSALESORDER.BILLAMT, SPSALESORDER.ADDRESS " & _
              "FROM SPSALESORDER INNER JOIN REQUISITION ON SPSALESORDER.REQID = REQUISITION.REQID " & _
              "WHERE (((SPSALESORDER.INVOICENO) Is Null) AND ((REQUISITION.REQNO)='" & txtOrderNo.Text & "'));"
    ElseIf PayType = 1 And updRef.Value = 1 Then    'Quick Service using Sales Order
       ' Qry = "SELECT SPSALESORDER.ORDERNO, SPSALESORDER.ORDERDATE, SPSALESORDER.SPSORDERID, SPSALESORDER.FRONTLINER, SPSALESORDER.EMPID, SPSALESORDER.TOTALPRICE, REQUISITION.JOBNO, REQUISITION.REQID, REQUISITION.REQNO, REQUISITION.CUSTOMER, SPSALESORDER.CUSTOMERID, SPSALESORDER.CTYPE, SPSALESORDER.REMARKS, SPSALESORDER.PAYSCHEME, SPSALESORDER.TERMS, SPSALESORDER.NPAYS, SPSALESORDER.PAYSTART, SPSALESORDER.BILLAMT " & _
              "FROM SPSALESORDER INNER JOIN REQUISITION ON SPSALESORDER.REQID = REQUISITION.REQID " & _
              "WHERE (((SPSALESORDER.ORDERNO)=" & txtOrderno.Text & ") AND ((SPSALESORDER.ORDERDATE)='" & Format(txtDate.Text, "yyyy-mm-dd") & ") AND ((SPSALESORDER.INVOICENO) Is Null));"
        Qry = "SELECT SPSALESORDER.ORDERNO, SPSALESORDER.ORDERDATE, SPSALESORDER.SPSORDERID, SPSALESORDER.FRONTLINER, SPSALESORDER.EMPID, SPSALESORDER.TOTALPRICE, REQUISITION.JOBNO, REQUISITION.REQID, REQUISITION.REQNO, REQUISITION.CUSTOMER, SPSALESORDER.CUSTOMERID, SPSALESORDER.CTYPE, SPSALESORDER.REMARKS, SPSALESORDER.PAYSCHEME, SPSALESORDER.TERMS, SPSALESORDER.NPAYS, SPSALESORDER.PAYSTART, SPSALESORDER.BILLAMT, SPSALESORDER.ADDRESS " & _
              "FROM SPSALESORDER INNER JOIN REQUISITION ON SPSALESORDER.REQID = REQUISITION.REQID " & _
              "WHERE (((SPSALESORDER.ORDERNO)=" & txtOrderNo.Text & ") AND ((SPSALESORDER.ORDERDATE)='" & Format(txtDate.Text, "yyyy-mm-dd") & "') AND ((SPSALESORDER.INVOICENO) Is Null));"
    
    ElseIf PayType = 1 And updRef.Value = 2 Then    'Quick Service using Job Order
        Qry = "SELECT SPSALESORDER.ORDERNO, SPSALESORDER.ORDERDATE, SPSALESORDER.SPSORDERID, SPSALESORDER.FRONTLINER, SPSALESORDER.EMPID, SPSALESORDER.TOTALPRICE, REQUISITION.JOBNO, REQUISITION.REQID, REQUISITION.REQNO, REQUISITION.CUSTOMER, SPSALESORDER.CUSTOMERID, SPSALESORDER.CTYPE, SPSALESORDER.REMARKS, SPSALESORDER.PAYSCHEME, SPSALESORDER.TERMS, SPSALESORDER.NPAYS, SPSALESORDER.PAYSTART, SPSALESORDER.BILLAMT, SPSALESORDER.ADDRESS " & _
              "FROM SPSALESORDER INNER JOIN REQUISITION ON SPSALESORDER.REQID = REQUISITION.REQID " & _
              "WHERE (((REQUISITION.JOBNO)='" & txtOrderNo.Text & "') AND ((SPSALESORDER.INVOICENO) Is Null));"
    ElseIf PayType = 2 And updRef.Value = 1 Then    'On Acct using P.O. number
        Qry = "SELECT SPSALESORDER.ORDERNO, SPSALESORDER.ORDERDATE, SPSALESORDER.SPSORDERID, SPSALESORDER.FRONTLINER, SPSALESORDER.EMPID, CUSTOMER.CUSTOMERNAME, SPSALESORDER.TOTALPRICE, SPSALESORDER.REFERENCE, SPSALESORDER.CUSTOMERID, SPSALESORDER.REQID, SPSALESORDER.CTYPE, SPSALESORDER.REMARKS, SPSALESORDER.PAYSCHEME, SPSALESORDER.TERMS, SPSALESORDER.NPAYS, SPSALESORDER.PAYSTART, SPSALESORDER.BILLAMT, SPSALESORDER.ADDRESS " & _
              "FROM SPSALESORDER INNER JOIN CUSTOMER ON SPSALESORDER.CUSTOMERID = CUSTOMER.CUSTOMERID " & _
              "WHERE (((SPSALESORDER.REFERENCE)='" & txtOrderNo.Text & "') AND ((SPSALESORDER.ORDERDATE)='" & Format(txtDate.Text, "yyyy-mm-dd") & "') AND ((SPSALESORDER.INVOICENO) Is Null)) ORDER BY SPSALESORDER.REQID DESC;"
    ElseIf PayType = 2 And updRef.Value = 0 Then    'On Acct using Sales Order
        Qry = "SELECT SPSALESORDER.ORDERNO, SPSALESORDER.ORDERDATE, SPSALESORDER.SPSORDERID, SPSALESORDER.FRONTLINER, SPSALESORDER.EMPID, CUSTOMER.CUSTOMERNAME, SPSALESORDER.TOTALPRICE, SPSALESORDER.REFERENCE, SPSALESORDER.CUSTOMERID, SPSALESORDER.REQID, SPSALESORDER.CTYPE, SPSALESORDER.REMARKS, SPSALESORDER.PAYSCHEME, SPSALESORDER.TERMS, SPSALESORDER.NPAYS, SPSALESORDER.PAYSTART, SPSALESORDER.BILLAMT, SPSALESORDER.ADDRESS " & _
              "FROM SPSALESORDER INNER JOIN CUSTOMER ON SPSALESORDER.CUSTOMERID = CUSTOMER.CUSTOMERID " & _
              "WHERE (((SPSALESORDER.ORDERNO)=" & txtOrderNo.Text & ") AND ((SPSALESORDER.ORDERDATE)='" & Format(txtDate.Text, "yyyy-mm-dd") & "') AND ((SPSALESORDER.INVOICENO) Is Null)) ORDER BY SPSALESORDER.REQID DESC;"
    End If

    
    With rsOrder
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            'txtOrderNo.Text = !ORDERNO
            'txtDate.Text = !ORDERDATE
            txtPartsMan.Text = !Frontliner
            txtSPSorderID.Text = !SPSORDERID
            txtPayable.Text = Format$(!TOTALPRICE, "#####0.00")
            'txtEmpID.Text = UserID
            txtCType.Text = IIf(IsNull(!Ctype), "", !Ctype)
            Text1.Visible = True
            Text2.Visible = True
            If PayType = 0 Then
                CustID = IIf(IsNull(!CUSTOMERID), 0, !CUSTOMERID)
                myRem = IIf(IsNull(!Remarks), "", !Remarks)

                If CustID > 0 And InStr(myRem, "POINTS") > 0 Then
                    txtSoldTo.Text = IIf(IsNull(!CUSTOMER), "", !CUSTOMER) '"CASH"
                    txtAddress.Text = IIf(IsNull(!Address), "", !Address) ' "Tagum City"  '"Davao City"
                    If IIf(IsNull(!REFERENCE), "", !REFERENCE) <> "" Then
                        tabTerm.Visible = True
                        tabTerm.Tab = 2
                        txtPCCardNo.Text = IIf(IsNull(!REFERENCE), "", !REFERENCE)
                        If InStr(!Remarks, "CUSTOMER POINTS") > 0 Then
                            txtPCType.Text = "CUSTOMER POINTS"
                        Else
                            txtPCType.Text = "CLUB POINTS"
                        End If
                        
                        txtPCPoints.Text = Format((!TOTALPRICE / 100), "#,###,###.#0")
                    End If
                    
                    If InStr(myRem, "WHOLESALE") > 0 And InStr(myRem, "REDEEM") > 0 Then
                        lblRemarks.Caption = "WHOLESALE DISCOUNT/REDEEM POINTS"
                        txtAmtRec.Text = Format$(!TOTALPRICE, "#####0.00")
                        Call txtAmtRec_LostFocus
                    ElseIf InStr(myRem, "WHOLESALE") > 0 Then
                        lblRemarks.Caption = "WHOLESALE DISCOUNT"
                    ElseIf InStr(myRem, "REDEEM") > 0 Then
                        lblRemarks.Caption = "REDEEM POINTS"
                        txtAmtRec.Text = Format$(!TOTALPRICE, "#####0.00")
                        Call txtAmtRec_LostFocus
                    End If
                ElseIf CustID = 0 Then
                    txtSoldTo.Text = IIf(IsNull(!CUSTOMER), "", !CUSTOMER) '"CASH"
                    txtAddress.Text = IIf(IsNull(!Address), "", !Address) ' "Tagum City"  '"Davao City"
                    'If SysBranch = "MOTO-R REPUBLIC CORP." Or SysBranch = "MOTO-R VAN SELLING" Then
                    '    txtAddress.Text = "Tagum City"
                    'Else
                    '    txtAddress.Text = "Davao City"
                    'End If
                ElseIf CustID > 0 And txtCType.Text = "WHOLESALE" And InStr(myRem, "CASH") > 0 Then
                    txtSoldTo.Text = !CUSTOMER '"CASH"
                    txtAddress.Text = !Address ' "Tagum City"  '"Davao City"
                    lblRemarks.Caption = IIf(IsNull(!REFERENCE), "", !REFERENCE)
                    'mySQL
                    'CQry = "SELECT CUSTOMER.CUSTOMERNAME From CUSTOMER WHERE CUSTOMER.CUSTOMERID=" & CustID & ";"
                    'txtSoldTo.Text = GetFieldValue(CQry)
                    'mySQL
                    'CQry = "SELECT CUSTOMER.ADDRESS From CUSTOMER WHERE CUSTOMER.CUSTOMERID=" & CustID & ";"
                    'txtAddress.Text = GetFieldValue(CQry)
                Else
                    MsgBox "Invalid Cash Sales Order!", vbExclamation + vbOKOnly, "Cash Only"
                    txtPartsMan.Text = ""
                    txtSPSorderID.Text = ""
                    txtPayable.Text = ""
                    'txtEmpID.Text = ""
                    txtCType.Text = ""
                    lblRemarks.Caption = ""
                    Text1.Visible = False
                    Text2.Visible = False
                    .Close
                    Exit Sub
                End If
            ElseIf PayType = 1 And updRef.Value = 0 Then    'Requisition
                txtSoldTo.Text = !CUSTOMER
                txtAddress.Text = !Address
                Label1.Caption = "REQ.#"
                Text1.Text = !REQNO
                Label2.Caption = "JOB ORDER#"
                Text2.Text = !JobNo
                txtReqID.Text = !ReqID
                txtCustomerID.Text = !CUSTOMERID
                'txtAddress.Text = GetAddress
                '***Access
                'txtJobID.Text = GetFieldValue("SELECT JOBORDERID From SRVCJOBORDER WHERE (JOBORDERNO='" & Text2.Text & "');")
                'mySQl
                txtJobID.Text = GetFieldValue("SELECT JOBORDERID From SRVCJOBORDER WHERE JOBORDERNO='" & Text2.Text & "';")
                
            ElseIf PayType = 1 And updRef.Value = 1 Then    'Sales Order
                txtSoldTo.Text = !CUSTOMER
                Label1.Caption = "REQ.#"
                Text1.Text = !REQNO
                Label2.Caption = "JOB ORDER#"
                Text2.Text = !JobNo
                txtReqID.Text = !ReqID
                txtCustomerID.Text = !CUSTOMERID
                txtAddress.Text = !Address
            ElseIf PayType = 1 And updRef.Value = 2 Then    'Job Order
                txtSoldTo.Text = !CUSTOMER
                Label1.Caption = "REQ.#"
                Text1.Text = !REQNO
                Label2.Caption = "JOB ORDER#"
                Text2.Text = !JobNo
                txtReqID.Text = !ReqID
                txtCustomerID.Text = !CUSTOMERID
                txtAddress.Text = !Address
            End If
            If PayType = 2 And updRef.Value = 1 Then   'On Acct - Order#
                txtSoldTo.Text = !CustomerName
                CustSrch = True
                cboAcctCust.Text = !CustomerName
                Label1.Caption = "SALES ORDER#"
                Text1.Text = !ORDERNO
                Label2.Caption = ""
                Text2.Visible = False
                txtAddress.Text = !Address
                txtCustomerID.Text = !CUSTOMERID
                txtReqID.Text = IIf(IsNull(!ReqID), "", !ReqID)
                If Val(txtCustomerID.Text) > 0 And Val(txtReqID.Text) > 0 Then
                    Label1.Caption = "REQUISITION#:"
                    Text1.Text = GetReqNo(txtReqID.Text)
                    Text1.Visible = True
                End If
            End If
            If PayType = 2 And updRef.Value = 0 Then   'On Acct. - P.O.
                myRem = IIf(IsNull(!Remarks), "", !Remarks)
                If txtCType.Text = "WHOLESALE" And myRem = "CASH" Then
                    MsgBox "Please use CASH in processing Sales Order " & txtOrderNo.Text, vbExclamation + vbOKOnly, "Invalid Charge Invoice"
                    txtPartsMan.Text = ""
                    txtSPSorderID.Text = ""
                    txtPayable.Text = ""
                    'txtEmpID.Text = ""
                    txtCType.Text = ""
                    Text1.Visible = False
                    Text2.Visible = False
                    .Close
                    Exit Sub
                End If
                txtSoldTo.Text = !CustomerName
                CustSrch = True
                cboAcctCust.Text = !CustomerName
                Label2.Caption = ""
                Text2.Visible = False
                txtAddress.Text = IIf(IsNull(!Address), "", !Address)
                txtCustomerID.Text = !CUSTOMERID
                txtReqID.Text = IIf(IsNull(!ReqID), "", !ReqID)
                If Val(txtCustomerID.Text) > 0 And Val(txtReqID.Text) > 0 Then
                    Label1.Caption = "REQUISITION#:"
                    Text1.Text = GetReqNo(txtReqID.Text)
                End If
                
                If IIf(IsNull(!REFERENCE), "", !REFERENCE) <> "" And InStr(!Remarks, "CUSTOMER POINTS") > 0 Then
                    tabTerm.Visible = True
                    tabTerm.Tab = 2
                    txtPCCardNo.Text = IIf(IsNull(!REFERENCE), "", !REFERENCE)
                    If InStr(!Remarks, "CUSTOMER POINTS") > 0 Then
                        txtPCType.Text = "CUSTOMER POINTS"
                    Else
                        txtPCType.Text = "CLUB POINTS"
                    End If
                    txtPCPoints.Text = Format((!TOTALPRICE / 100), "#,###,###.#0")
                End If
                
                If txtPCCardNo.Text = "" Then
                    Label1.Caption = "P.O.#"
                    Text1.Text = IIf(IsNull(!REFERENCE), "", !REFERENCE)
                End If
            End If
            'On Acct Info
            txtPScheme.Text = IIf(IsNull(!PayScheme), "", !PayScheme)
            txtNTerms.Text = IIf(IsNull(!Terms), "", !Terms)
            txtNPay.Text = IIf(IsNull(!Npays), "", !Npays)
            txtDueAmt.Text = IIf(IsNull(!BILLAMT), "", !BILLAMT)
            txtPayStart.Text = IIf(IsNull(!PayStart), "", Format(!PayStart, "mm/dd/yyyy"))
            If txtPScheme.Text <> "" Then
                txtTerm.Text = txtNTerms.Text & " DAYS  [" & txtNPay.Text & IIf(Val(txtNPay.Text) > 1, " Payments]", " Payment]")
                txtAmtDue.Text = Format$(Val(txtDueAmt.Text), "#,##0.00") & " " & txtPScheme.Text
                If PayType = 1 Then tabTerm.Visible = True: tabTerm.Tab = 0
            End If
        Else
            txtOrderNo.Text = ""
            txtPartsMan.Text = ""
            txtSPSorderID.Text = ""
            txtPayable.Text = ""
            'txtEmpID.Text = ""
            txtReqID.Text = ""
            txtCustomerID.Text = ""
            txtSoldTo.Text = ""
            txtAddress.Text = ""
            Label1.Caption = ""
            Text1.Text = ""
            Text1.Visible = False
            Label2.Caption = ""
            Text2.Text = ""
            Text2.Visible = False
            flxOrder.Rows = 1
            flxOrder.Rows = 2
            txtTotal.Text = ""
            If PayType = 2 Then
                txtTerm.Text = ""
                txtAmtDue.Text = ""
                txtPayStart.Text = ""
                cboAcctCust.Text = ""
            End If
            txtPScheme.Text = ""
            txtNTerms.Text = ""
            txtNPay.Text = ""
            txtDueAmt.Text = ""
            txtPayStart.Text = ""
            txtTerm.Text = ""
            txtAmtDue.Text = ""
            lblRemarks.Caption = ""
        End If
        .Close
    End With
    If txtSPSorderID.Text = "" Then Exit Sub
    '***Access
    'Qry = "SELECT SPSORDERDETAIL.QTY, SPSORDERDETAIL.UNIT, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPSORDERDETAIL.PRICE, [QTY]*[PRICE] AS TTLPRICE, SPSORDERDETAIL.SPAREPARTID, SPAREPARTS.QTYONHAND " & _
          "FROM SPSORDERDETAIL INNER JOIN SPAREPARTS ON SPSORDERDETAIL.SPAREPARTID = SPAREPARTS.SPAREPARTID WHERE (((SPSORDERDETAIL.QTY)>0) AND ((SPSORDERDETAIL.SPSORDERID)=" & txtSPSorderID.Text & ")) ORDER BY SPAREPARTS.PARTNUMBER;"
    'mySQL
    Qry = "SELECT SPSORDERDETAIL.QTY, SPSORDERDETAIL.UNIT, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPSORDERDETAIL.PRICE, SPSORDERDETAIL.QTY*SPSORDERDETAIL.PRICE AS TTLPRICE, SPSORDERDETAIL.SPAREPARTID, SPAREPARTS.QTYONHAND " & _
          "FROM SPSORDERDETAIL INNER JOIN SPAREPARTS ON SPSORDERDETAIL.SPAREPARTID = SPAREPARTS.SPAREPARTID WHERE SPSORDERDETAIL.QTY>0 AND SPSORDERDETAIL.SPSORDERID=" & txtSPSorderID.Text & " ORDER BY SPAREPARTS.PARTNUMBER;"
    
    NoPrice = False
    With rsOrdDtls
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            flxOrder.Rows = 1
            Ttl = 0
            Do While Not .EOF
                flxOrder.AddItem vbTab & !Qty & vbTab & !Unit & vbTab & !PartNumber & vbTab & !Description & vbTab & Format$(!Price, "###,##0.00") & vbTab & Format$(!TtlPrice, "###,##0.00") & vbTab & !SparepartID & vbTab & !QtyOnHand
                Call TOT_PENDINGORDER(Format(txtDate.Text, "yyyy-mm-dd"), flxOrder.TextMatrix(flxOrder.Row, 7))
                If flxOrder.TextMatrix(flxOrder.Row, 8) - Val(txtTotPen.Text) < flxOrder.TextMatrix(flxOrder.Row, 1) Then
                    For ccol = 0 To 8
                        flxOrder.ColSel = ccol
                        flxOrder.RowSel = flxOrder.Row
                        flxOrder.FillStyle = flexFillRepeat: flxOrder.CellBackColor = &HC0FFC0
                    Next ccol
                    InvldOrd = True
                End If
                If !Price = 0 Then NoPrice = True
                Ttl = Ttl + !TtlPrice
                .MoveNext
            Loop
            txtTotal.Text = Format$(Ttl, "###,##0.00")
            txtPayable.Text = Format$(Ttl, "#####0.00")
        Else
            flxOrder.Rows = 1
            flxOrder.Rows = 2
        End If
        .Close
    End With
    If PayType = 2 Then
        '***Access
        'Qry = "SELECT DPLEDGER.* From DPLEDGER WHERE (((DPLEDGER.CUSTOMERID)=" & txtCustomerID.Text & ") AND ((DPLEDGER.STATUS)='OPEN'));"
        'mySQL
        Qry = "SELECT DPLEDGER.* From DPLEDGER WHERE DPLEDGER.CUSTOMERID=" & txtCustomerID.Text & " AND DPLEDGER.STATUS='OPEN';"
        With rsDP
            .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
            If .RecordCount > 0 Then
                lblORno.ToolTipText = IIf(IsNull(!ORNO), "", "OR#: " & !ORNO & " OR Amt:" & Format(!ORAmount, "#,##0.00") & " OR Date:" & Format(!ORdate, "mm/dd/yy"))
                If Text1.Text = !ORNO Then
                    Label1.Caption = "O.R. NO."
                Else
                    Text1.Text = ""
                    Label1.Caption = ""
                End If
            End If
            .Close
        End With
    End If
    If NoPrice Then MsgBox "Some items contain zero (0) price!" & vbCrLf & "Please inform your Manager to update the order!", vbExclamation + vbOKOnly, "Warning"
    If InvldOrd = True Then MsgBox "Check some items Quantity not Greater than Available Stock!" & vbCrLf & "Please inform your Manager to update the order!", vbExclamation + vbOKOnly, "Warning"
End Sub
'Validate Pending Order
'Checked if the invoice already printed
Private Sub TOT_PENDINGORDER(SP_ORDERDATE As String, SP_ID As Long)
    Dim rsPOrder As New ADODB.Recordset, TotPTrans As Long
    Dim QRY_ORDER As String, PENDINGORDER As Long, S_Order As Long, s_Col As Long
    
    
    QRY_ORDER = "SELECT SUM(spsorderdetail.QTY) AS TOTALQTY, spsorderdetail.SPAREPARTID, spsalesorder.INVOICENO " & _
        " From spsalesorder RIGHT OUTER JOIN spsorderdetail ON (spsalesorder.SPSORDERID = spsorderdetail.SPSORDERID) " & _
        " Where spsalesorder.ORDERNO <> " & Val(txtOrderNo.Text) & " AND spsalesorder.ORDERDATE = '" & SP_ORDERDATE & "' AND spsorderdetail.SPAREPARTID = " & SP_ID & " AND spsalesorder.INVOICENO IS NULL Group By " & _
        " spsorderdetail.SPAREPARTID, spsalesorder.INVOICENO IS NULL;"
    
    PENDINGORDER = Val(GetFieldValue(QRY_ORDER))
    
    QRY_ORDER = "SELECT SUM(stocktransdtl.QTY) As TOTTRANS From stocktransfer RIGHT OUTER JOIN stocktransdtl ON (stocktransfer.TRANSFERID = stocktransdtl.TRANSFERID) " & _
                 " Where stocktransfer.`STATUS` = 'PENDING' AND stocktransdtl.SPAREPARTID = " & SP_ID & ";"
    TotPTrans = Val(GetFieldValue(QRY_ORDER))
    txtTotPen.Text = PENDINGORDER + TotPTrans
End Sub

Function GetReqNo(ReqID As Long) As String
    Dim rsReq As New ADODB.Recordset
    Dim Qry As String
    
    '***Access
    'Qry = "SELECT REQUISITION.REQNO From Requisition WHERE (((REQUISITION.REQID)=" & ReqID & "));"
    'mySQL
    Qry = "SELECT REQUISITION.REQNO From Requisition WHERE REQUISITION.REQID=" & ReqID & ";"
    With rsReq
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            GetReqNo = IIf(IsNull(!REQNO), "", !REQNO)
        Else
            GetReqNo = ""
        End If
        .Close
    End With
End Function

Function GetAddress() As String
    Dim rsAdrs As New ADODB.Recordset
    Dim Qry As String, MyRef As String
    
    If txtSoldTo.Text = "" Then
        GetAddress = ""
        Exit Function
    End If
    
    If Val(txtCustomerID.Text) > 0 Then
        '***Access
        'Qry = "SELECT CUSTOMER.ADDRESS, CUSTOMER.CUSTOMERID From CUSTOMER " & _
              "WHERE (((CUSTOMER.CUSTOMERID)=" & txtCustomerID.Text & "));"
        'mySQl
        Qry = "SELECT CUSTOMER.ADDRESS, CUSTOMER.CUSTOMERID From CUSTOMER " & _
              "WHERE CUSTOMER.CUSTOMERID=" & txtCustomerID.Text & ";"
    
    Else
        '***Access
        'Qry = "SELECT SRVCUSTOMER.ADDRESS From SRVCUSTOMER " & _
              "WHERE (((SRVCUSTOMER.CUSTOMER)='" & txtSoldTo.Text & "'));"
        'mySQL
        Qry = "SELECT SRVCUSTOMER.ADDRESS From SRVCUSTOMER " & _
              "WHERE SRVCUSTOMER.CUSTOMER='" & txtSoldTo.Text & "';"
    
    End If
    With rsAdrs
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            GetAddress = IIf(IsNull(!Address), "", !Address)
        Else
            GetAddress = ""
        End If
        .Close
    End With
End Function

Function ValidateWalkIn() As String
    Dim eMsg As String, vQry As String
    
    eMsg = ""
    If Val(txtPayable.Text) <> PureNum(txtTotal.Text) Then eMsg = "Amount Payable did not match Sales Order Total!"
    If txtInvNo.Text = "" Then
        eMsg = eMsg & vbCrLf & "Incomplete entry on field Invoice No."
    Else
        'vQry = "SELECT SPINVOICE.SPINVNO From SPINVOICE WHERE (((SPINVOICE.SPINVNO)='" & txtInvNo.Text & "'));"
        'mySQL
        vQry = "SELECT SPINVOICE.SPINVNO From SPINVOICE WHERE spinvoice.INVTYPE = 'CASH' AND SPINVOICE.SPINVNO = '" & txtInvNo.Text & "';"
     
        If CtrlNumExist(vQry) Then eMsg = eMsg & vbCrLf & "Invoice No. already exist"
    End If
    If Val(txtEmpID.Text) = 0 Then eMsg = eMsg & vbCrLf & "Please select assigned Cashier in list"
    If Not IsDate(txtDate.Text) Then eMsg = eMsg & vbCrLf & "Invalid entry on field Date"
    If txtSoldTo.Text = "" Then eMsg = eMsg & vbCrLf & "Incomplete entry on field Sold To"
    If txtAddress.Text = "" Then eMsg = eMsg & vbCrLf & "Incomplete entry on field Address"
    If Val(txtAmtRec.Text) < Val(txtPayable.Text) Then eMsg = eMsg & vbCrLf & "Insufficient amount received!"
    If NoPrice Then eMsg = eMsg & vbCrLf & "Some items contain zero(0) price!"
    ValidateWalkIn = eMsg
End Function

Function ValidateChargeInvoice() As String
    Dim eMsg As String, vQry As String, oQry As String
    
    
    eMsg = ""
    If txtInvNo.Text = "" Then
        eMsg = eMsg & vbCrLf & "Incomplete entry on field Invoice No."
    Else
        'vQry = "SELECT SPINVOICE.SPINVNO From SPINVOICE WHERE (((SPINVOICE.SPINVNO)='" & txtInvNo.Text & "'));"
        'mySQl
        vQry = "SELECT SPINVOICE.SPINVNO From SPINVOICE WHERE spinvoice.INVTYPE = 'CREDIT' AND SPINVOICE.SPINVNO='" & txtInvNo.Text & "';"
 
        If CtrlNumExist(vQry) Then eMsg = eMsg & vbCrLf & "Invoice No. already exist"
    End If
    If Not IsDate(txtDate.Text) Then eMsg = eMsg & vbCrLf & "Invalid entry on field Date"
    If txtSoldTo.Text = "" Then eMsg = eMsg & vbCrLf & "Incomplete entry on field Sold To"
    If txtAddress.Text = "" Then eMsg = eMsg & vbCrLf & "Incomplete entry on field Address"
    
    If PayType = 1 Then
        If txtReqID.Text = "" Then eMsg = eMsg & vbCrLf & "Unknown Requisition Number"
        If txtPScheme.Text = "" Or Val(txtNPay.Text) = 0 Then eMsg = eMsg & vbCrLf & "No Terms Specified"
    ElseIf PayType = 2 Then
        If txtCustomerID.Text = "" Then eMsg = eMsg & vbCrLf & "Invalid On Account Customer"
'        oQry = "SELECT spsalesorder.CTYPE From spsalesorder WHERE spsalesorder.ORDERNO = '" & txtOrderNo.Text & "' AND spsalesorder.CTYPE='" & txtCType.Text & "';"
        If txtCType.Text = "KEY ACCOUNTS" Then
            txtTerm.Text = "30"
            txtPScheme.Text = "MONTHLY"
        Else
            If txtTerm.Text = "" Then eMsg = eMsg & vbCrLf & "No Terms Specified"
        End If
        If txtTerm.Text = "" And txtORno.Text = "" Then
            eMsg = eMsg & vbCrLf & "No Terms Specified"
        ElseIf txtPScheme.Text = "" And Val(txtNPay.Text) = 0 Then
            eMsg = eMsg & vbCrLf & "No Terms Specified"
        End If
    End If
    If NoPrice Then eMsg = eMsg & vbCrLf & "Some items contain zero(0) price!"
    ValidateChargeInvoice = eMsg
End Function

Private Sub SaveInvoice(sMsg As String)
    Dim Qry As String, PrnVal(0 To 18) As String, ErrMsg As String, CItype As String, DPref As String, POref As String
    Dim LSeq As Integer, myTtl As Currency, Ans As Integer, vReqID As Long, vJOid As Long
    Dim myVat As Currency, AmtLessVat As Currency, DPamt As Currency, MyBal As Currency
    Dim myRem As String, MyPO As String, MyPStart As String, myPROMOTITLE As String
    
    On Error GoTo LocalErr
    If SysBranch = "LYR MOTOR SHOP AND SPAREPARTS DAVAO" Then
        Ans = MsgBox("Do you want to Save Invoice?", vbQuestion + vbYesNo, "Print")
    'ElseIf InStr(SysBranch, "PANABO") > 0 Then
        'no questions
    '    Ans = vbYes
    Else
        Ans = MsgBox("Do you want to print Invoice?", vbQuestion + vbYesNo, "Print")
    End If
    'Ans = MsgBox("Do you want to print Invoice?", vbQuestion + vbYesNo, "Print")
    If Ans = vbNo Then Exit Sub
    'Saving Table Spare Parts Invoice
    ErrMsg = "[787] Error adding record in SP Invoice"
    myTtl = Val(Replace(txtTotal.Text, ",", ""))
    MyBal = myTtl
    myPROMOTITLE = ""
    
    If cboPayType.Text <> "" Then myRem = cboPayType.Text
    If PayType = 0 Then
        If lblRemarks.Caption <> "" Then myRem = lblRemarks.Caption
    ElseIf PayType = 1 Then
        vReqID = Val(txtReqID.Text)
        myRem = myRem & " QUICK SERVICE"
        MyPO = ""
    ElseIf PayType = 2 Then
        If updRef.Value = 0 Then MyPO = Text1.Text
        If updRef.Value = 1 Then MyPO = txtOrderNo.Text
        If txtORno.Text <> "" And txtORdate.Text <> "" And txtORamt.Text <> "" Then myRem = "ORDER W/ DP"
    End If
    
    If InStr(lblRemarks.Caption, "REDEEM") > 0 Then
        LYRcon.Execute "INSERT INTO SPInvoice (SPINVNO, Transno, InvDate, ChargeTo, InvType, Address, Amount, PrintCopy, Frontliner, EmpId, SPSORDERID, Ctype, ReqID, Remarks, CUSTOMERID, SPPono, UpdatedBy, LastUpdate) " & _
        "VALUES ('" & txtInvNo.Text & "', '" & "SI#-" & Trim(txtInvNo.Text) & "', '" & Format(txtDate.Text, "yyyy-mm-dd") & "', '" & txtSoldTo.Text & "', '" & "REDEEM" & "', '" & Replace(txtAddress.Text, "'", "") & "', " & myTtl & ", " & IIf(Ans <> vbYes, 0, 1) & ", '" & txtPartsMan.Text & "', " & Val(txtEmpID.Text) & ", " & Val(txtSPSorderID.Text) & ", '" & txtCType.Text & "', " & vReqID & ", '" & myRem & "', " & Val(txtCustomerID.Text) & ", '" & MyPO & "', " & UserID & ", '" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "');"
    Else
        LYRcon.Execute "INSERT INTO SPInvoice (SPINVNO, Transno, InvDate, ChargeTo, InvType, Address, Amount, PrintCopy, Frontliner, EmpId, SPSORDERID, Ctype, ReqID, Remarks, CUSTOMERID, SPPono, UpdatedBy, LastUpdate) " & _
        "VALUES ('" & txtInvNo.Text & "', '" & "SI#-" & Trim(txtInvNo.Text) & "', '" & Format(txtDate.Text, "yyyy-mm-dd") & "', '" & txtSoldTo.Text & "', '" & myInvType & "', '" & Replace(txtAddress.Text, "'", "") & "', " & myTtl & ", " & IIf(Ans <> vbYes, 0, 1) & ", '" & txtPartsMan.Text & "', " & Val(txtEmpID.Text) & ", " & Val(txtSPSorderID.Text) & ", '" & txtCType.Text & "', " & vReqID & ", '" & myRem & "', " & Val(txtCustomerID.Text) & ", '" & MyPO & "', " & UserID & ", '" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "');"
    End If
    
    If PayType = 0 Then 'Walkin
        ErrMsg = "[818] Error adding record to Payments"
        myRem = cboPayType.Text  ' IIf(chkOnLine.Value = 1, "ON-LINE PAYMENT", "")
        
        If InStr(lblRemarks.Caption, "WHOLESALE") > 0 And InStr(lblRemarks.Caption, "REDEEM") > 0 Then
            myRem = myRem & " WHOLESALE DISCOUNT/REDEEM POINTS"
        ElseIf lblRemarks.Caption = "WHOLESALE DISCOUNT" Then
            myRem = myRem & " WHOLESALE DISCOUNT"
        ElseIf InStr(lblRemarks.Caption, "REDEEM") > 0 Then
            myRem = myRem & " REDEEM POINTS"
        Else
            If txtCType.Text = "WHOLESALE" Then
                myRem = myRem & " WHOLESALE"
            End If
        End If
        
        If InStr(myRem, "REDEEM") <= 0 Then
            LYRcon.Execute "INSERT INTO Payments (PayDate, InvoiceNO, SpareParts, CustomerName, Remarks, EmpId, UpdatedBy, LastUpdate) " & _
            "VALUES ('" & Format(txtDate.Text, "yyyy-mm-dd") & "', '" & Trim(txtInvNo.Text) & "', " & Val(txtPayable.Text) & ", '" & txtSoldTo.Text & "', '" & myRem & "', " & Val(txtEmpID.Text) & ", " & UserID & ", '" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "');"
        End If
        
        If txtPCCardNo.Text <> "" Then
            Qry = "call sp_promotionalonline.sp_promocard_create_earn('" & sysBranchCode & "','" & txtPCCardNo.Text & "','" & Trim(txtSoldTo.Text) & "','" & txtPCType.Text & "', '" & Trim(txtInvNo.Text) & "','" & Replace(txtPayable.Text, ",", "") & "','" & Replace(txtPCPoints.Text, ",", "") & "'," & Val(UserID) & ")"
            With rsAnyUse
                If .State = adStateOpen Then .Close
                .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
            End With
            
            If InStr(myRem, "REDEEM") > 0 Then
                Qry = "call sp_promotionalonline.sp_promocard_create_redeem('" & sysBranchCode & "','" & txtPCCardNo.Text & "','" & Trim(txtSoldTo.Text) & "','" & txtPCType.Text & "', '" & Trim(txtInvNo.Text) & "','" & Replace(txtPayable.Text, ",", "") & "','" & Replace(txtAmtRec.Text, ",", "") & "'," & Val(UserID) & ")"
                With rsAnyUse
                    If .State = adStateOpen Then .Close
                    .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
                End With
            End If
        End If
    End If
    If PayType = 1 Then
        ErrMsg = "[834] Error updating Invoice in Requisition"
        'mySQL
        LYRcon.Execute "UPDATE REQUISITION SET REQUISITION.INVOICENO = '" & txtInvNo.Text & "' " & _
                       "WHERE REQUISITION.REQID=" & txtReqID.Text & ";"
        CItype = "QUICK SERVICE"
    End If
    If txtORno.Text <> "" And txtORdate.Text <> "" And txtORamt.Text <> "" Then
        CItype = "ORDER W/ DP"
        DPamt = Val(Replace(txtORamt.Text, ",", ""))
        If DPamt > myTtl Then DPamt = myTtl
        MyBal = myTtl - DPamt
        'Update Customer O.S. Bal
        ErrMsg = "[843] Error updating Customer Outstanding Balance"
        'mySQL
        LYRcon.Execute "UPDATE CUSTOMER SET CUSTOMER.OSBAL = CUSTOMER.OSBAL+" & myTtl - DPamt & _
                       " WHERE CUSTOMER.CUSTOMERID=" & txtCustomerID.Text & ";"
        'Update DP Ledger
        ErrMsg = "[966] Error updating Down Payment Ledger"
        Qry = "SELECT DPLEDGER.INVREF From DPLEDGER WHERE DPLEDGER.ORNO='" & txtORno.Text & "' AND DPLEDGER.CUSTOMERID=" & txtCustomerID.Text & ";"
        
        DPref = GetFieldValue(Qry)
        DPref = IIf(DPref <> "", DPref & ", " & txtInvNo.Text, txtInvNo.Text)
        '***Access
        'LYRcon.Execute "UPDATE DPLEDGER SET DPLEDGER.STATUS = '" & IIf(DPamt <= myTtl, "CLOSED", "OPEN") & "', DPLEDGER.BALANCE = " & IIf(DPamt <= myTtl, 0, DPamt - myTtl) & ", DPLEDGER.INVREF = '" & DPref & "' " & _
                       "WHERE (((DPLEDGER.ORNO)='" & txtORno.Text & "') AND ((DPLEDGER.CUSTOMERID)=" & txtCustomerID.Text & "));"
        'mySQL
        LYRcon.Execute "UPDATE DPLEDGER SET DPLEDGER.STATUS = '" & IIf(DPamt <= myTtl, "CLOSED", "OPEN") & "', DPLEDGER.BALANCE = " & IIf(DPamt <= myTtl, 0, DPamt - myTtl) & ", DPLEDGER.INVREF = '" & DPref & "' " & _
                       "WHERE DPLEDGER.ORNO='" & txtORno.Text & "' AND DPLEDGER.CUSTOMERID=" & txtCustomerID.Text & ";"
        
        ErrMsg = "[975] Error updating Remarks in O.R. of Down Payment "
        LYRcon.Execute "UPDATE PAYMENTS SET PAYMENTS.REMARKS = 'INV.# " & txtInvNo.Text & "' WHERE PAYMENTS.ORNO='" & txtORno.Text & "';"
        
        'Add OR in Ledger
        ErrMsg = "[953] Error new record for Ledger with Down Payment"
        
        LYRcon.Execute "INSERT INTO LEDGER (SPINVNO, RefType, RefNo, RefDate, Credit, OSbal, CUSTOMERID, SeqNo) " & _
        "VALUES ('" & txtInvNo.Text & "', 'Payment', '" & Trim(txtORno.Text) & "', '" & Format(txtORdate.Text, "yyyy-mm-dd") & "', " & DPamt & ", " & MyBal & ", " & Val(txtCustomerID.Text) & ", 1);"
    End If
    
    If txtPScheme.Text <> "" And Val(txtNPay.Text) > 0 Then 'Save On Account
        If CItype <> "ORDER W/ DP" Then CItype = "ON-ACCOUNT"
        ErrMsg = "[925] Error new record for On Account with Down Payment"
        MyPStart = IIf(IsDate(txtPayStart.Text), txtPayStart.Text, DateAdd("m", 1, DateValue(txtDate.Text)))
        If PayType = 1 Then 'Display Referece
            If CItype = "ORDER W/ DP" Then CItype = "ON-ACCT. QUICK SERVICE"
            MyPO = Text2.Text
            vJOid = Val(txtJobID.Text)
            vReqID = Val(txtReqID.Text)
        Else
            MyPO = ""
            vJOid = 0
            vReqID = 0
        End If

        'mySQL
        LYRcon.Execute "INSERT INTO ONACCTERMS (SPINVNO, Terms, PayScheme, Npays, PayStart, BillingAmt, NxtBill, AmtDue, Status, TtlAmtPaid, BALANCE, SPSORDERID, CUSTOMERID, UpdatedBy, LastUpdate, AcctAmount, TransDate, JOBORDERNO, JOBORDERID, RequisitionID) " & _
        "VALUES ('" & Trim(txtInvNo.Text) & "', " & Val(txtNTerms.Text) & ", '" & txtPScheme.Text & "', " & Val(txtNPay.Text) & ", '" & Format(MyPStart, "yyyy-mm-dd") & "', " & Val(txtDueAmt.Text) & ", '" & Format(MyPStart, "yyyy-mm-dd") & "', " & Val(txtDueAmt.Text) & ", '" & IIf(MyBal = 0, "CLOSED", "OPEN") & "', " & DPamt & ", " & MyBal & ", " & Val(txtSPSorderID.Text) & ", " & Val(txtCustomerID.Text) & ", " & UserID & ", '" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "', " & myTtl & ", '" & Format(txtDate.Text, "yyyy-mm-dd") & "', '" & MyPO & "', " & vJOid & ", " & vReqID & ");"
        
        LYRcon.Execute "UPDATE CUSTOMER SET CUSTOMER.OSBAL = CUSTOMER.OSBAL+" & MyBal & _
                       " WHERE CUSTOMER.CUSTOMERID=" & txtCustomerID.Text & ";"
    End If
    
       
    'Saving Table Spare Parts Inventory Transaction Register
    With flxOrder
        ErrMsg = "[849] Error adding record in SPITR"
        For i = 1 To .Rows - 1
            MyPO = ""
            If PayType = 0 Or (PayType = 1 And updRef.Value = 1) Or (PayType = 2 And updRef.Value = 0) Then
                MyPO = Left$(txtOrderNo.Text, 20) & IIf(lblRemarks.Caption <> "WHOLESALE DISCOUNT", "", " WS Discount")
            ElseIf (PayType = 1 And updRef.Value = 0) Or (PayType = 2 And updRef.Value = 1) Then
                MyPO = Left$(Text1.Text, 20)
            ElseIf PayType = 1 And updRef.Value = 2 Then
                MyPO = Left$(Text2.Text, 20)
            End If
            
'------------------------- w/ COST
            
            wCost = GetFieldValue("SELECT spareparts.cost From spareparts Where spareparts.SPAREPARTID = " & Val(.TextMatrix(i, 7)) & ";")
            myPROMOTITLE = GetFieldValue("SELECT spareparts.PROMOTITLE From spareparts Where spareparts.SPAREPARTID = " & Val(.TextMatrix(i, 7)) & ";")
            
            If InStr(myRem, "REDEEM") > 0 Then
                LYRcon.Execute "INSERT INTO SPITR (Reference, TRANSTYPE, Remarks, TransDate, SparepartID, `OUT`, Price, PACKINGLIST, Source, UpdatedBy, LastUpdate, cost, PROMOTITLE) " & _
                "VALUES ('" & "SI#-" & Trim(txtInvNo.Text) & "', 'REDEEM', '" & IIf(Ans <> vbYes, "Temp", txtPCCardNo.Text) & "', '" & Format(txtDate.Text, "yyyy-mm-dd") & "', " & Val(.TextMatrix(i, 7)) & ", " & Val(.TextMatrix(i, 1)) & ", " & Val(Replace(.TextMatrix(i, 5), ",", "")) & ", '" & MyPO & "', '" & Left$(txtSoldTo.Text, 50) & "', " & UserID & ", '" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "', " & wCost & ", '" & myPROMOTITLE & "');"
            Else
                LYRcon.Execute "INSERT INTO SPITR (Reference, TRANSTYPE, Remarks, TransDate, SparepartID, `OUT`, Price, PACKINGLIST, Source, UpdatedBy, LastUpdate, cost, PROMOTITLE) " & _
                "VALUES ('" & "SI#-" & Trim(txtInvNo.Text) & "', 'OUT', '" & IIf(Ans <> vbYes, "Temp", myInvType) & "', '" & Format(txtDate.Text, "yyyy-mm-dd") & "', " & Val(.TextMatrix(i, 7)) & ", " & Val(.TextMatrix(i, 1)) & ", " & Val(Replace(.TextMatrix(i, 5), ",", "")) & ", '" & MyPO & "', '" & Left$(txtSoldTo.Text, 50) & "', " & UserID & ", '" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "', " & wCost & ", '" & myPROMOTITLE & "');"
            End If
            
            ErrMsg = "[874] Error updating SP Qty on hand"
            LYRcon.Execute "UPDATE SPAREPARTS SET SPAREPARTS.QTYONHAND = SPAREPARTS.QTYONHAND-" & Val(.TextMatrix(i, 1)) & ", SPAREPARTS.UPDATEDBY = " & UserID & ", SPAREPARTS.LASTUPDATE = '" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "' " & _
                           "WHERE SPAREPARTS.SPAREPARTID=" & Val(.TextMatrix(i, 7)) & ";"
        Next i
    End With
    'Print Here
     If Ans = vbYes And SysBranch <> "LYR MOTOR SHOP AND SPAREPARTS DAVAO" And SysBranch <> "BINOS SPAREPARTS STO. TOMAS" And SysBranch <> "BINOS SPAREPARTS SAMAL" And SysBranch <> "MOTORHUB SPAREPARTS MARBEL" And SysBranch <> "MOTORHUB SPAREPARTS TAGUM" Then 'PRINT USING preprinted invoice 'And SysBranch <> "BINOS SPAREPARTS POLOMOLOK"
        PrnVal(0) = IIf(PayType = 0, "x", "") 'Cash
        PrnVal(1) = IIf(PayType <> 0, "x", "") 'Credit
        PrnVal(2) = txtSoldTo.Text  'Customer
        PrnVal(3) = CItype  'Type
        PrnVal(4) = txtDate.Text    'date
        PrnVal(5) = txtAddress.Text 'address
        PrnVal(6) = txtTerm.Text    'term
        PrnVal(7) = MotoRep.sBar.Panels(3).Text 'Cashier
        PrnVal(8) = ""  'TIN
        If PayType = 1 Then 'Display Referece
            PrnVal(9) = "JOB ORDER# " & Text2.Text & ", REQ.# " & Text1.Text
        Else
            If CItype = "ORDER W/ DP" Then
                PrnVal(9) = ""
            Else
                If PayType = 2 And updRef.Value = 1 Then
                    PrnVal(9) = "P.O.# " & txtOrderNo.Text
                Else
                    PrnVal(9) = IIf(Text1.Text <> "", Label1.Caption & " " & Text1.Text, "") 'Reference
                End If
            End If
        End If
        PrnVal(10) = txtPartsMan.Text 'Partsman
        PrnVal(11) = "AMOUNT:"
        AmtLessVat = myTtl / 1.12
        PrnVal(12) = Format(AmtLessVat, "#,##0.00") 'Total less Vat
        PrnVal(13) = "12% VAT:"
        myVat = myTtl - AmtLessVat
        PrnVal(14) = Format(myVat, "#,##0.00") 'Total
        PrnVal(15) = "TOTAL:"
        PrnVal(16) = txtTotal.Text 'Total
        PrnVal(17) = Format(Time, "hh:mm ampm") & IIf(CItype <> "ORDER W/ DP", "", "   Note: DP O.R.#" & txtORno.Text & ", O.R. Date:" & txtORdate.Text & ", O.R. Amt:" & txtORamt.Text)
        PrnVal(18) = txtAmtRec.Text 'amount recieved
        
        Call PrintMyForm(PrnVal, "INVOICE", flxOrder)
        
    End If
    
    'Update SPSalesOrder
    ErrMsg = "[814] Error Updating Invoice No. in SP Sales Order"
    '***Access
    'LYRcon.Execute "UPDATE SPSALESORDER SET SPSALESORDER.INVOICENO = '" & txtInvNo.Text & "' " & _
                   "WHERE (((SPSALESORDER.SPSORDERID)=" & Val(txtSPSorderID.Text) & "));"
    'mySQL
    LYRcon.Execute "UPDATE SPSALESORDER SET SPSALESORDER.INVOICENO = '" & Trim(txtInvNo.Text) & "' " & _
                   "WHERE SPSALESORDER.SPSORDERID=" & Val(txtSPSorderID.Text) & ";"
    
    'Update PO
    POref = ""
    ErrMsg = "[1124] Error Updating Invoice No. in PURCHASE ORDER"
    If PayType = 2 And updRef.Value = 1 Then POref = txtOrderNo.Text
    If PayType = 2 And updRef.Value = 0 Then POref = Text1.Text
    If POref <> "" Then
        '***Access
        'LYRcon.Execute "UPDATE CUSTPO SET CUSTPO.INVOICENO = '" & txtInvNo.Text & "' WHERE (((CUSTPO.CPONO)='" & POref & "'));"
        'mySQL
        LYRcon.Execute "UPDATE CUSTPO SET CUSTPO.INVOICENO = '" & Trim(txtInvNo.Text) & "' WHERE CUSTPO.CPONO='" & POref & "';"
        
    End If
    If Ans = vbYes And (SysBranch = "BINOS SPAREPARTS STO. TOMAS" Or SysBranch = "BINOS SPAREPARTS SAMAL" Or SysBranch = "MOTORHUB SPAREPARTS MARBEL" Or SysBranch = "MOTORHUB SPAREPARTS TAGUM") Then 'Or SysBranch = "BINOS SPAREPARTS POLOMOLOK"
        Qry = "SELECT SPITR.`OUT`, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPITR.PRICE " & _
              "FROM SPAREPARTS INNER JOIN SPITR ON SPAREPARTS.SPAREPARTID = SPITR.SPAREPARTID WHERE SPITR.REFERENCE='SI#-" & txtInvNo.Text & "';"
        
        PRNatt.MySource = "SPINVOICE"
        PRNatt.QryPrn = Qry
        PRNatt.ParamValue1 = myInvType
        PRNatt.ParamValue2 = txtSoldTo.Text
        PRNatt.ParamValue3 = Replace(txtAddress.Text, "'", "")
        PRNatt.ParamValue4 = txtInvNo.Text
        PRNatt.ParamValue5 = Format(txtDate.Text, "yyyy-mm-dd")
        PRNatt.ParamValue6 = myRem
        PRNatt.Show vbModal
    Else
        MsgBox sMsg, vbInformation + vbOKOnly, "Print S.I."
    End If

    Exit Sub
LocalErr:
    If InStr(ErrMsg, "[814] Error Updating Invoice No. in SP Sales Order") > 0 Then Exit Sub
    If InStr(ErrMsg, "[849] Error adding record in SPITR") > 0 Or InStr(ErrMsg, "[874] Error updating SP Qty on hand") > 0 Then
        '***Access
        'LYRcon.Execute "DELETE SPINVOICE.*, SPINVOICE.SPINVNO From SPINVOICE WHERE (((SPINVOICE.SPINVNO)='" & txtInvNo.Text & "'));"
        'LYRcon.Execute "DELETE SPITR.*, SPITR.REFERENCE From SPITR WHERE (((SPITR.REFERENCE)='SI#-" & txtInvNo.Text & "'));"
        'If PayType = 0 Then 'Walkin
        '    LYRcon.Execute "DELETE PAYMENTS.*, PAYMENTS.INVOICENO From PAYMENTS WHERE (((PAYMENTS.INVOICENO)='" & txtInvNo.Text & "'));"
        'ElseIf PayType = 1 Then    '
        '    LYRcon.Execute "UPDATE REQUISITION SET REQUISITION.INVOICENO = Null WHERE (((REQUISITION.REQID)=" & txtReqID.Text & "));"
        '    LYRcon.Execute "UPDATE ONACCTERMS SET ONACCTERMS.STATUS = 'Temp' WHERE (((ONACCTERMS.SPSORDERID)=" & txtSPSorderID.Text & "));"
        'ElseIf PayType = 2 And CItype = "ON-ACCOUNT" Then
        '    LYRcon.Execute "UPDATE ONACCTERMS SET ONACCTERMS.STATUS = 'Temp' WHERE (((ONACCTERMS.SPSORDERID)=" & txtSPSorderID.Text & ") And ((ONACCTERMS.CUSTOMERID)=" & txtCustomerID.Text & "));"
        'End If
        
        'mySQL
        
        
        If PayType = 0 Then 'Walkin
            LYRcon.Execute "DELETE From SPINVOICE WHERE (((SPINVOICE.SPINVNO)='" & Trim(txtInvNo.Text) & "')) AND SPINVOICE.INVTYPE = 'CASH';"
            LYRcon.Execute "DELETE From SPITR WHERE (((SPITR.REFERENCE)='SI#-" & Trim(txtInvNo.Text) & "')) AND REMARKS = 'CASH';"
            LYRcon.Execute "DELETE From PAYMENTS WHERE (((PAYMENTS.INVOICENO)='" & Trim(txtInvNo.Text) & "'));"
        ElseIf PayType = 1 Then    '
            LYRcon.Execute "DELETE From SPINVOICE WHERE (((SPINVOICE.SPINVNO)='" & Trim(txtInvNo.Text) & "')) AND SPINVOICE.INVTYPE = 'CREDIT';"
            LYRcon.Execute "DELETE From SPITR WHERE (((SPITR.REFERENCE)='SI#-" & Trim(txtInvNo.Text) & "')) AND REMARKS = 'CREDIT';"
            LYRcon.Execute "UPDATE REQUISITION SET REQUISITION.INVOICENO = Null WHERE (((REQUISITION.REQID)=" & txtReqID.Text & "));"
            LYRcon.Execute "UPDATE ONACCTERMS SET ONACCTERMS.STATUS = 'Temp' WHERE (((ONACCTERMS.SPSORDERID)=" & txtSPSorderID.Text & "));"
        ElseIf PayType = 2 And CItype = "ON-ACCOUNT" Then
            LYRcon.Execute "DELETE From SPINVOICE WHERE (((SPINVOICE.SPINVNO)='" & Trim(txtInvNo.Text) & "')) AND SPINVOICE.INVTYPE = 'CREDIT';"
            LYRcon.Execute "DELETE From SPITR WHERE (((SPITR.REFERENCE)='SI#-" & Trim(txtInvNo.Text) & "')) AND REMARKS = 'CREDIT';"
            LYRcon.Execute "UPDATE ONACCTERMS SET ONACCTERMS.STATUS = 'Temp' WHERE (((ONACCTERMS.SPSORDERID)=" & txtSPSorderID.Text & ") And ((ONACCTERMS.CUSTOMERID)=" & txtCustomerID.Text & "));"
        End If
                
    End If
    
    MsgBox ErrMsg & vbCrLf & Err.Description, vbCritical + vbOKOnly, "Error " & Err.Number
End Sub


Function MaxSeqB(MyID As Integer)
    Dim rsSeq As New ADODB.Recordset
    Dim Qry As String
    
    '***Access
    'Qry = "SELECT Max(LEDGER.SEQNO) AS MaxOfSEQNO From LEDGER " & _
          "WHERE (((LEDGER.BRANCHID)=" & MyId & "));"
    'mySQL
    Qry = "SELECT Max(round(LEDGER.SEQNO)) AS MaxOfSEQNO From LEDGER " & _
          "WHERE LEDGER.BRANCHID=" & MyID & ";"
    
    With rsSeq
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            MaxSeqB = IIf(IsNull(.Fields(0)), 0, .Fields(0))
        Else
            MaxSeqB = 0
        End If
        .Close
    End With
End Function

Function LastBalance(MyID As Integer, MySeq As Integer) As Single
    Dim rsBal As New ADODB.Recordset
    Dim Qry As String
    
    'Qry = "SELECT LEDGER.OSBAL From LEDGER " & _
          "WHERE (((LEDGER.SEQNO)=" & MySeq & ") AND ((LEDGER.BRANCHID)=" & MyId & "));"
    'mySQL
    Qry = "SELECT LEDGER.OSBAL From LEDGER " & _
          "WHERE LEDGER.SEQNO=" & MySeq & " AND LEDGER.BRANCHID=" & MyID & ";"
    
    
    With rsBal
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            LastBalance = IIf(IsNull(.Fields(0)), 0, .Fields(0))
        Else
            LastBalance = 0
        End If
        .Close
    End With
End Function

Private Sub UpDRef_Change()
    If txtOrderNo.Text <> "" Then Call ClearEntries
    txtOrderNo.Text = ""
    If PayType = 1 Then 'Quick Service
        If updRef.Value = 0 Then
            lblOrderNo.Caption = "REQUISITION#"
        ElseIf updRef.Value = 1 Then
            lblOrderNo.Caption = "SALES ORDER#"
        ElseIf updRef.Value = 2 Then
            lblOrderNo.Caption = "JOB ORDER#"
        End If
    ElseIf PayType = 2 Then 'Charge Invoice
        If updRef.Value = 0 Then
            lblOrderNo.Caption = "SALES ORDER#"
        ElseIf updRef.Value = 1 Then
            lblOrderNo.Caption = "P.O. NUMBER"
        End If
    End If
End Sub
