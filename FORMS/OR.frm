VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form Receipt 
   Caption         =   "Official Receipt"
   ClientHeight    =   11490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15210
   ControlBox      =   0   'False
   Icon            =   "OR.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11490
   ScaleWidth      =   15210
   WindowState     =   2  'Maximized
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxJOdtls 
      Height          =   2535
      Left            =   3360
      TabIndex        =   130
      Tag             =   "XYF"
      Top             =   3360
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   4471
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.TextBox txtTmpGrayColor 
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Left            =   15600
      TabIndex        =   113
      Top             =   600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Timer MoveUp 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   0
      Top             =   4320
   End
   Begin VB.Timer tmrLoadCustomer 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   2280
      Top             =   7800
   End
   Begin VB.ComboBox cboCashier 
      Height          =   315
      Left            =   120
      TabIndex        =   103
      Top             =   6840
      Width           =   3015
   End
   Begin MSFlexGridLib.MSFlexGrid flxAcct 
      Height          =   1215
      Left            =   240
      TabIndex        =   102
      Top             =   10320
      Visible         =   0   'False
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   2143
      _Version        =   393216
      Cols            =   17
      FormatString    =   $"OR.frx":27A2
   End
   Begin TabDlg.SSTab tabWait 
      Height          =   975
      Left            =   5640
      TabIndex        =   99
      Top             =   6840
      Visible         =   0   'False
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   1720
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "OR.frx":2867
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Image1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label4"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      Begin VB.Label Label4 
         Caption         =   "Please wait.... your request is being processed"
         Height          =   375
         Left            =   360
         TabIndex        =   100
         Top             =   360
         Width           =   3615
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   3600
         Picture         =   "OR.frx":2883
         Top             =   120
         Width           =   240
      End
   End
   Begin VB.CheckBox OnAcctPCV 
      Caption         =   "Check1"
      Height          =   255
      Left            =   360
      TabIndex        =   98
      Top             =   8280
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox txtPaymentID 
      Height          =   285
      Left            =   360
      TabIndex        =   92
      Top             =   9120
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CheckBox chkInvoice 
      Caption         =   "INVOICE"
      Height          =   255
      Left            =   360
      TabIndex        =   65
      Tag             =   "XYF"
      Top             =   6120
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.CheckBox chkJobOrder 
      Caption         =   "JOB ORDER"
      Height          =   255
      Left            =   360
      TabIndex        =   64
      Tag             =   "XYF"
      Top             =   6360
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Frame Frame4 
      Height          =   5295
      Left            =   3240
      TabIndex        =   35
      Tag             =   "XYF"
      Top             =   720
      Width           =   11880
      Begin VB.ComboBox cboAcctTitle 
         Height          =   315
         Left            =   8400
         TabIndex        =   128
         Tag             =   "XYF"
         Top             =   2920
         Width           =   3255
      End
      Begin VB.ComboBox cboBranch 
         Height          =   315
         Left            =   120
         TabIndex        =   109
         Top             =   4800
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.ComboBox cboPayType 
         Height          =   315
         Left            =   8040
         TabIndex        =   105
         Tag             =   "XYF"
         Top             =   2610
         Width           =   3615
      End
      Begin VB.TextBox txtTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   95
         Tag             =   "XYF"
         Top             =   4860
         Width           =   1455
      End
      Begin VB.Frame Frame2 
         Height          =   1935
         Left            =   7080
         TabIndex        =   67
         Tag             =   "XYF"
         Top             =   3240
         Width           =   4575
         Begin VB.Frame Frame6 
            Height          =   1935
            Left            =   0
            TabIndex        =   121
            Tag             =   "XYF"
            Top             =   0
            Visible         =   0   'False
            Width           =   4575
            Begin VB.TextBox txtCreditBank 
               Height          =   285
               Left            =   1800
               TabIndex        =   126
               Tag             =   "XYF"
               Top             =   1320
               Width           =   2535
            End
            Begin VB.TextBox txtCardNo 
               Height          =   285
               Left            =   1800
               TabIndex        =   125
               Tag             =   "XYF"
               Top             =   960
               Width           =   2535
            End
            Begin VB.TextBox txtCardHolder 
               Height          =   285
               Left            =   1800
               MaxLength       =   50
               TabIndex        =   123
               Tag             =   "XYF"
               Top             =   600
               Width           =   2535
            End
            Begin VB.Label Label14 
               Alignment       =   1  'Right Justify
               Caption         =   "NAME OF  BANK"
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
               TabIndex        =   127
               Tag             =   "XYF"
               Top             =   1320
               Width           =   1575
            End
            Begin VB.Label Label13 
               Alignment       =   1  'Right Justify
               Caption         =   "CARD NO."
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
               TabIndex        =   124
               Tag             =   "XYF"
               Top             =   960
               Width           =   1575
            End
            Begin VB.Label Label12 
               Alignment       =   1  'Right Justify
               Caption         =   "CARD HOLDER NAME"
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
               TabIndex        =   122
               Tag             =   "XYF"
               Top             =   600
               Width           =   1575
            End
         End
         Begin VB.ComboBox cboCheck 
            Height          =   315
            Left            =   1680
            TabIndex        =   97
            Top             =   2040
            Visible         =   0   'False
            Width           =   2655
         End
         Begin VB.TextBox txtChkPRid 
            Height          =   285
            Left            =   360
            TabIndex        =   96
            Tag             =   "XYF"
            Top             =   2040
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.TextBox txtChkAcct 
            Height          =   285
            Left            =   1680
            TabIndex        =   77
            Tag             =   "XYF"
            Top             =   1230
            Width           =   2655
         End
         Begin VB.TextBox txtChkAmt 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   1680
            TabIndex        =   76
            Tag             =   "XYF"
            Top             =   900
            Width           =   2655
         End
         Begin VB.TextBox txtChkDate 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   1680
            TabIndex        =   75
            Tag             =   "XYF"
            Top             =   570
            Width           =   2655
         End
         Begin VB.TextBox txtChkBank 
            Height          =   285
            Left            =   1680
            TabIndex        =   74
            Tag             =   "XYF"
            Top             =   1560
            Width           =   2655
         End
         Begin VB.TextBox txtCheckNo 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   1680
            TabIndex        =   73
            Tag             =   "XYF"
            Top             =   240
            Width           =   2655
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "NAME OF  BANK"
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
            Left            =   120
            TabIndex        =   72
            Tag             =   "XYF"
            Top             =   1560
            Width           =   1455
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "ACCOUNT NAME"
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
            Left            =   120
            TabIndex        =   71
            Tag             =   "XYF"
            Top             =   1230
            Width           =   1455
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "AMOUNT [BAL.]"
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
            Left            =   120
            TabIndex        =   70
            Tag             =   "XYF"
            Top             =   900
            Width           =   1455
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "CHECK DATE"
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
            TabIndex        =   69
            Tag             =   "XYF"
            Top             =   570
            Width           =   1455
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "CHECK NO."
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
            TabIndex        =   68
            Tag             =   "XYF"
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.CheckBox chkCheck 
         Caption         =   "Check"
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
         Left            =   7080
         TabIndex        =   66
         Tag             =   "XYF"
         Top             =   2640
         Width           =   1215
      End
      Begin VB.Frame FraHeading 
         Height          =   975
         Left            =   120
         TabIndex        =   56
         Tag             =   "XYF"
         Top             =   120
         Width           =   11535
         Begin VB.Frame fraAllocation 
            BorderStyle     =   0  'None
            Height          =   375
            Left            =   120
            TabIndex        =   111
            Tag             =   "XYF"
            Top             =   120
            Visible         =   0   'False
            Width           =   4455
            Begin VB.ComboBox cboAllocationNo 
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
               Left            =   1440
               Style           =   2  'Dropdown List
               TabIndex        =   2
               Tag             =   "XYF"
               Top             =   60
               Width           =   2775
            End
            Begin VB.Label lblALLOCATIONNO 
               Caption         =   "ALLOCATION NO."
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
               Left            =   0
               TabIndex        =   112
               Tag             =   "XYF"
               Top             =   60
               Width           =   1335
            End
         End
         Begin MSComCtl2.UpDown udORCR 
            Height          =   255
            Left            =   9480
            TabIndex        =   108
            Tag             =   "XYF"
            Top             =   210
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   450
            _Version        =   393216
            OrigLeft        =   1440
            OrigTop         =   270
            OrigRight       =   1695
            OrigBottom      =   525
            Max             =   1
            Enabled         =   -1  'True
         End
         Begin VB.TextBox txtORNo 
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
            TabIndex        =   3
            Tag             =   "XYF"
            Top             =   180
            Width           =   1695
         End
         Begin VB.TextBox txtRefNo 
            Alignment       =   2  'Center
            Height          =   315
            Left            =   1680
            TabIndex        =   60
            Tag             =   "XYF"
            Top             =   180
            Width           =   1575
         End
         Begin MSComCtl2.UpDown updRef 
            Height          =   255
            Left            =   1440
            TabIndex        =   59
            Tag             =   "XYF"
            Top             =   210
            Visible         =   0   'False
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   450
            _Version        =   393216
            OrigLeft        =   1440
            OrigTop         =   270
            OrigRight       =   1695
            OrigBottom      =   525
            Max             =   1
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox txtDate 
            Height          =   315
            Left            =   9720
            TabIndex        =   4
            Tag             =   "XYF"
            Top             =   540
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
         Begin VB.Frame fraAllocationOthrData 
            BorderStyle     =   0  'None
            Height          =   495
            Left            =   120
            TabIndex        =   114
            Tag             =   "XYF"
            Top             =   450
            Visible         =   0   'False
            Width           =   7575
            Begin VB.TextBox txtTaxAmount 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0C0C0&
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
               Left            =   6255
               Locked          =   -1  'True
               TabIndex        =   120
               Tag             =   "XYF"
               Top             =   120
               Width           =   1215
            End
            Begin VB.TextBox txtTax 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0C0C0&
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
               Left            =   4000
               Locked          =   -1  'True
               TabIndex        =   118
               Tag             =   "XYF"
               Top             =   120
               Width           =   855
            End
            Begin VB.TextBox txtTTLAllocation 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0C0C0&
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
               Left            =   1440
               Locked          =   -1  'True
               TabIndex        =   116
               Tag             =   "XYF"
               Top             =   120
               Width           =   1215
            End
            Begin VB.Label Label8 
               Caption         =   "TAX w/Held Amt"
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
               Left            =   5040
               TabIndex        =   119
               Tag             =   "XYF"
               Top             =   120
               Width           =   1215
            End
            Begin VB.Label Label7 
               Caption         =   "TAX w/Held (%)"
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
               Left            =   2760
               TabIndex        =   117
               Tag             =   "XYF"
               Top             =   120
               Width           =   1215
            End
            Begin VB.Label Label6 
               Caption         =   "TOTAL AMOUNT"
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
               Left            =   0
               TabIndex        =   115
               Tag             =   "XYF"
               Top             =   120
               Width           =   1455
            End
         End
         Begin VB.CheckBox chkQckSrvc 
            Caption         =   "QUICK SERVICE"
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
            TabIndex        =   57
            Tag             =   "XYF"
            Top             =   640
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.CheckBox chkCash 
            Caption         =   "Cash"
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
            Left            =   1680
            TabIndex        =   58
            Tag             =   "XYF"
            Top             =   640
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.Label lblORNo 
            Alignment       =   1  'Right Justify
            Caption         =   "OFFICIAL RECEIPT NO."
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
            Left            =   7440
            TabIndex        =   63
            Tag             =   "XYF"
            Top             =   210
            Width           =   1935
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
            Left            =   8280
            TabIndex        =   62
            Tag             =   "XYF"
            Top             =   585
            Width           =   1335
         End
         Begin VB.Label lblRefNo 
            Caption         =   "REFERENCE NO."
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
            TabIndex        =   61
            Tag             =   "XYF"
            Top             =   210
            Width           =   1455
         End
      End
      Begin VB.Frame FraPayment 
         Height          =   1425
         Left            =   8040
         TabIndex        =   49
         Tag             =   "XYF"
         Top             =   1080
         Width           =   3615
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
            TabIndex        =   52
            Tag             =   "XYF"
            Top             =   180
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
            TabIndex        =   51
            Tag             =   "XYF"
            Top             =   540
            Width           =   1695
         End
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
            TabIndex        =   50
            Tag             =   "XYF"
            Top             =   885
            Width           =   1695
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
            TabIndex        =   55
            Tag             =   "XYF"
            Top             =   255
            Width           =   1575
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
            TabIndex        =   54
            Tag             =   "XYF"
            Top             =   570
            Width           =   1335
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
            TabIndex        =   53
            Tag             =   "XYF"
            Top             =   915
            Width           =   1335
         End
      End
      Begin VB.Frame Frame1 
         Height          =   1425
         Left            =   120
         TabIndex        =   38
         Tag             =   "XYF"
         Top             =   1080
         Width           =   7815
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
            TabIndex        =   42
            Tag             =   "XYF"
            Top             =   600
            Width           =   6015
         End
         Begin VB.TextBox txtSumOf 
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
            TabIndex        =   41
            Tag             =   "XYF"
            Top             =   960
            Width           =   6015
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
            Left            =   5880
            TabIndex        =   40
            Tag             =   "XYF"
            Top             =   1560
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.TextBox txtTIN 
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
            Left            =   6240
            TabIndex        =   39
            Tag             =   "XYF"
            Top             =   240
            Width           =   1455
         End
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
            TabIndex        =   1
            Tag             =   "XYF"
            Top             =   240
            Visible         =   0   'False
            Width           =   3615
         End
         Begin VB.TextBox txtRecFrom 
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
            TabIndex        =   43
            Tag             =   "XYF"
            Top             =   240
            Width           =   3615
         End
         Begin VB.Label lblSoldTo 
            Caption         =   "RECEIVED FROM"
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
            TabIndex        =   48
            Tag             =   "XYF"
            Top             =   270
            Width           =   1455
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
            Left            =   240
            TabIndex        =   47
            Tag             =   "XYF"
            Top             =   600
            Width           =   1095
         End
         Begin VB.Label lblPartsMan 
            Caption         =   "THE SUM OF"
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
            TabIndex        =   46
            Tag             =   "XYF"
            Top             =   960
            Width           =   1215
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
            Left            =   4680
            TabIndex        =   45
            Tag             =   "XYF"
            Top             =   1680
            Width           =   1335
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "TIN NO."
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
            Left            =   5160
            TabIndex        =   44
            Tag             =   "XYF"
            Top             =   285
            Width           =   975
         End
      End
      Begin VB.TextBox txtAmtPd 
         Height          =   285
         Left            =   6240
         TabIndex        =   36
         Top             =   4800
         Visible         =   0   'False
         Width           =   615
      End
      Begin MSFlexGridLib.MSFlexGrid flxBal 
         Height          =   1935
         Left            =   120
         TabIndex        =   37
         Tag             =   "XYF"
         Top             =   2640
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   3413
         _Version        =   393216
         Cols            =   5
         FixedCols       =   0
         RowHeightMin    =   300
         BackColorBkg    =   -2147483633
         ScrollBars      =   2
         AllowUserResizing=   1
         FormatString    =   "^INV. NO.       |^J.O. NO.       |>BALANCE      |>AMT. PAID   |NEW BAL.      "
      End
      Begin MSFlexGridLib.MSFlexGrid flxOrder 
         Height          =   2655
         Left            =   12480
         TabIndex        =   107
         Tag             =   "XYF"
         Top             =   2520
         Visible         =   0   'False
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   4683
         _Version        =   393216
         Cols            =   9
         RowHeightMin    =   300
         BackColorBkg    =   16777215
         FormatString    =   "     |^QTY.  |^UNIT   |<PART NUMBER             |<DESCRIPTION                    |>UNIT PRICE   |>TOTAL PRICE "
      End
      Begin VB.Label Label15 
         Caption         =   "ACCOUNT TITLE:"
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
         Left            =   7080
         TabIndex        =   129
         Tag             =   "XYF"
         Top             =   3000
         Width           =   1335
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         Caption         =   "TOTAL AMT. PAID: "
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
         Left            =   1680
         TabIndex        =   94
         Tag             =   "XYF"
         Top             =   4875
         Width           =   1575
      End
      Begin ResizeXtra.RXResizer RXResizer1 
         Left            =   720
         Top             =   7560
         _ExtentX        =   847
         _ExtentY        =   847
         RegistryKey     =   "RXB49E"
         Registered      =   -1  'True
         m_MDI_Width     =   15450
         m_MDI_Height    =   12060
      End
   End
   Begin VB.Frame Frame3 
      Height          =   5295
      Left            =   120
      TabIndex        =   30
      Tag             =   "XYF"
      Top             =   720
      Width           =   3000
      Begin VB.TextBox txtDP 
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
         Left            =   240
         TabIndex        =   32
         Tag             =   "XYF"
         Top             =   1440
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.ComboBox cboOthers 
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
         Left            =   240
         TabIndex        =   31
         Tag             =   "XYF"
         Top             =   1080
         Visible         =   0   'False
         Width           =   2535
      End
      Begin lvButton.lvButtons_H cmdDetails 
         Height          =   375
         Left            =   120
         TabIndex        =   33
         Tag             =   "XYF"
         Top             =   240
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   661
         Caption         =   "Details"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
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
         cBack           =   -2147483633
      End
      Begin MSFlexGridLib.MSFlexGrid flxDetails 
         Height          =   4575
         Left            =   120
         TabIndex        =   34
         Tag             =   "XYF"
         Top             =   600
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   8070
         _Version        =   393216
         Rows            =   1
         Cols            =   3
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   300
         BackColorBkg    =   16777215
         GridColor       =   16777215
         GridColorFixed  =   16777215
         ScrollBars      =   2
      End
   End
   Begin TabDlg.SSTab tabOR 
      Height          =   3855
      Left            =   3240
      TabIndex        =   29
      Tag             =   "XYF"
      Top             =   6240
      Visible         =   0   'False
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   6800
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "LEDGER"
      TabPicture(0)   =   "OR.frx":2E0D
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "flxLedger"
      Tab(0).Control(1)=   "Frame5"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "ATTACHMENT"
      TabPicture(1)   =   "OR.frx":2E29
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "CRViewer1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin CRVIEWERLibCtl.CRViewer CRViewer1 
         Height          =   3375
         Left            =   120
         TabIndex        =   93
         Tag             =   "XYF"
         Top             =   360
         Width           =   11535
         DisplayGroupTree=   -1  'True
         DisplayToolbar  =   -1  'True
         EnableGroupTree =   0   'False
         EnableNavigationControls=   -1  'True
         EnableStopButton=   -1  'True
         EnablePrintButton=   -1  'True
         EnableZoomControl=   -1  'True
         EnableCloseButton=   -1  'True
         EnableProgressControl=   -1  'True
         EnableSearchControl=   -1  'True
         EnableRefreshButton=   -1  'True
         EnableDrillDown =   -1  'True
         EnableAnimationControl=   -1  'True
         EnableSelectExpertButton=   0   'False
         EnableToolbar   =   -1  'True
         DisplayBorder   =   -1  'True
         DisplayTabs     =   -1  'True
         DisplayBackgroundEdge=   -1  'True
         SelectionFormula=   ""
         EnablePopupMenu =   -1  'True
         EnableExportButton=   0   'False
         EnableSearchExpertButton=   0   'False
         EnableHelpButton=   0   'False
      End
      Begin VB.Frame Frame5 
         Height          =   2775
         Left            =   -74880
         TabIndex        =   78
         Tag             =   "XYF"
         Top             =   360
         Width           =   3495
         Begin VB.TextBox txtInvNo 
            Alignment       =   2  'Center
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
            Left            =   1560
            TabIndex        =   90
            Tag             =   "XYF"
            Top             =   300
            Width           =   1815
         End
         Begin VB.TextBox txtInvDate 
            Alignment       =   2  'Center
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
            Left            =   1560
            TabIndex        =   89
            Tag             =   "XYF"
            Top             =   672
            Width           =   1815
         End
         Begin VB.TextBox txtInvAmt 
            Alignment       =   2  'Center
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
            Left            =   1560
            TabIndex        =   82
            Tag             =   "XYF"
            Top             =   1044
            Width           =   1815
         End
         Begin VB.TextBox txtTerms 
            Alignment       =   2  'Center
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
            Left            =   1560
            TabIndex        =   81
            Tag             =   "XYF"
            Top             =   1788
            Width           =   1815
         End
         Begin VB.TextBox txtDue 
            Alignment       =   2  'Center
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
            Left            =   1560
            TabIndex        =   80
            Tag             =   "XYF"
            Top             =   2160
            Width           =   1815
         End
         Begin VB.TextBox txtBal 
            Alignment       =   2  'Center
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
            Left            =   1560
            TabIndex        =   79
            Tag             =   "XYF"
            Top             =   1416
            Width           =   1815
         End
         Begin VB.Label lblInvNo 
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
            Left            =   120
            TabIndex        =   88
            Tag             =   "XYF"
            Top             =   330
            Width           =   1335
         End
         Begin VB.Label lblInvDate 
            Alignment       =   1  'Right Justify
            Caption         =   "INV. DATE"
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
            TabIndex        =   87
            Tag             =   "XYF"
            Top             =   702
            Width           =   1335
         End
         Begin VB.Label lblInvAmt 
            Alignment       =   1  'Right Justify
            Caption         =   "INV. AMOUNT"
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
            TabIndex        =   86
            Tag             =   "XYF"
            Top             =   1074
            Width           =   1335
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "TERMS"
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
            Left            =   120
            TabIndex        =   85
            Tag             =   "XYF"
            Top             =   1818
            Width           =   1335
         End
         Begin VB.Label lblDue 
            Alignment       =   1  'Right Justify
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
            TabIndex        =   84
            Tag             =   "XYF"
            Top             =   2190
            Width           =   1335
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "O.S. BAL."
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
            Left            =   120
            TabIndex        =   83
            Tag             =   "XYF"
            Top             =   1440
            Width           =   1335
         End
      End
      Begin MSFlexGridLib.MSFlexGrid flxLedger 
         Height          =   3255
         Left            =   -71280
         TabIndex        =   91
         Tag             =   "XYF"
         Top             =   480
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   5741
         _Version        =   393216
         Cols            =   5
         FixedCols       =   0
         BackColorBkg    =   16777215
         ScrollBars      =   2
         SelectionMode   =   1
         AllowUserResizing=   1
         FormatString    =   "^ DATE|^REF. NO.|>DEBIT |>CREDIT |>BALANCE"
      End
   End
   Begin VB.ListBox lstChrgInvNo 
      Height          =   450
      Left            =   12240
      TabIndex        =   28
      Top             =   10200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtJOamt 
      Height          =   285
      Left            =   7560
      TabIndex        =   27
      Top             =   10320
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtPartsman 
      Height          =   285
      Left            =   6480
      TabIndex        =   26
      Top             =   10320
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtTotalDue 
      Height          =   285
      Left            =   5760
      TabIndex        =   24
      Top             =   10320
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtTotalBal 
      Height          =   285
      Left            =   4920
      TabIndex        =   23
      Top             =   10320
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtCustomerID 
      Height          =   285
      Left            =   4320
      TabIndex        =   22
      Top             =   10320
      Visible         =   0   'False
      Width           =   615
   End
   Begin lvButton.lvButtons_H cmdPrint 
      Height          =   375
      Left            =   11880
      TabIndex        =   21
      Top             =   0
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   661
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
      ImgAlign        =   4
      Image           =   "OR.frx":2E45
      cBack           =   -2147483633
   End
   Begin VB.TextBox txtAcctNo 
      Height          =   285
      Left            =   3960
      TabIndex        =   20
      Top             =   10320
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtPrevRef 
      Height          =   285
      Left            =   3120
      TabIndex        =   19
      Top             =   10320
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtJobID 
      Height          =   285
      Left            =   2640
      TabIndex        =   12
      Top             =   10320
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtPrnCopy 
      Height          =   285
      Left            =   2040
      TabIndex        =   11
      Top             =   10320
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtCashierID 
      Height          =   285
      Left            =   1320
      TabIndex        =   10
      Top             =   10320
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtSPSorderID 
      Height          =   285
      Left            =   720
      TabIndex        =   9
      Top             =   10320
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Timer HideBtn 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   0
      Top             =   4800
   End
   Begin VB.Timer ReArrange 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   0
      Top             =   5280
   End
   Begin VB.TextBox txtEmpID 
      Height          =   285
      Left            =   120
      TabIndex        =   8
      Top             =   10320
      Visible         =   0   'False
      Width           =   615
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   0
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      Caption         =   "&Quick Service"
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
      Image           =   "OR.frx":33DF
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   1
      Left            =   1695
      TabIndex        =   14
      Top             =   0
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      Caption         =   "&Down Payment"
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
      Image           =   "OR.frx":3979
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   4
      Left            =   6780
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   0
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      Caption         =   "E&xit O.R."
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
      Image           =   "OR.frx":3F13
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   2
      Left            =   3390
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   0
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      Caption         =   "&Charge Invoice"
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
      Image           =   "OR.frx":44AD
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   5
      Left            =   8475
      TabIndex        =   17
      Top             =   0
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      Caption         =   "&Save O.R."
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
      Image           =   "OR.frx":4D87
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   6
      Left            =   10170
      TabIndex        =   18
      Top             =   0
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      Caption         =   "&Cancel O.R."
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
      Image           =   "OR.frx":5321
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   3
      Left            =   5085
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   0
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      Caption         =   "&Others"
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
      Image           =   "OR.frx":58BB
      cBack           =   -2147483633
   End
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   255
      Left            =   3240
      TabIndex        =   101
      Tag             =   "XY"
      Top             =   6000
      Visible         =   0   'False
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   7
      Left            =   12360
      TabIndex        =   106
      Top             =   0
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      Caption         =   "Cash Invoice"
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
      Image           =   "OR.frx":6795
      cBack           =   -2147483633
   End
   Begin MSFlexGridLib.MSFlexGrid FlxAging 
      Height          =   975
      Left            =   15480
      TabIndex        =   110
      Tag             =   "xyf"
      Top             =   3120
      Visible         =   0   'False
      Width           =   14415
      _ExtentX        =   25426
      _ExtentY        =   1720
      _Version        =   393216
      Cols            =   24
      FixedCols       =   0
      RowHeightMin    =   300
      BackColorBkg    =   -2147483633
      FormatString    =   $"OR.frx":BF87
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "ASSIGNED CASHIER"
      Height          =   255
      Left            =   120
      TabIndex        =   104
      Tag             =   "XYF"
      Top             =   7200
      Width           =   3015
   End
   Begin VB.Label lblHeading 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "OFFICIAL RECEIPT"
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
      Left            =   12600
      TabIndex        =   6
      Top             =   480
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
      Top             =   420
      Width           =   2895
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
      TabIndex        =   5
      Top             =   450
      Width           =   2895
   End
   Begin VB.Label lblHeading 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "OFFICIAL RECEIPT"
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
      Left            =   12585
      TabIndex        =   7
      Top             =   510
      Width           =   2415
   End
   Begin VB.Menu mnuDetails 
      Caption         =   "Details"
      Visible         =   0   'False
      Begin VB.Menu mnuLedger 
         Caption         =   "Display Ledger"
      End
   End
End
Attribute VB_Name = "Receipt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Dim CustSrch As Boolean, ResizeFlx As Boolean, BtnOnHold As Boolean, CustBln As Boolean
'PostPaymentType As Integer
'Dim MyAcct As String

Dim rsPrn As New ADODB.Recordset
Dim rsAcctCust As New ADODB.Recordset
Dim rsAcct As New ADODB.Recordset
Dim myInvType As String, ThisBranch As String
Dim ActiveBtn As Integer, PayType As Integer, ActRow As Integer, ActCol As Integer, isOthers As Integer
Dim CurrentBtn As Integer, PrtlPay As Boolean
Dim CmdTop(0 To 3) As Integer
Dim CustSrch As Boolean, BalRow As Integer, BalCol As Integer
Dim Cashr_Idx As Integer
Dim JONoAcct As String


'New Variable Declaration
'By: SMC
'Date: 2016-08-05
Dim DidRecordEntry As Boolean
Dim AllocationNo As String
Dim TRANSTYPE As String
Dim PrevCustomer As String
Dim StrPAYTYPE As String

'New Variable Declaration
'By: ARL
'Date: 2020-08-28
Dim rsAcctTitle As New ADODB.Recordset
Dim arrAcctTitle() As String
Dim Qry As String
Dim myAcctCode As String
Dim AcctSrch As Boolean

Public Sub CboDisplay()
On Error GoTo er
    'Get Data from database
    DoEvents
    Qry = "call accounting_corp.sp_chart_of_accts_list()"
    If rsAcctTitle.State = adStateOpen Then rsAcctTitle.Close
    rsAcctTitle.Open Qry, LYRcon, adOpenStatic, adLockPessimistic
    If rsAcctTitle.State = 1 Then
       Call modObjTools.addToListObjectWithArrayGivenField(cboAcctTitle, rsAcctTitle, arrAcctTitle, "Acct_Code", "Acct_Title")
    End If
    If rsAcctTitle.State <> 0 Then rsAcctTitle.Close
    
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->CboDisplay()")
End Sub


Private Sub ClearEntries()

    flxJOdtls.Visible = False
    flxJOdtls.Clear
    txtPayable.Locked = False
    txtAmtRec.Locked = False
    txtAmtChange.Locked = False
    txtPayable.BackColor = txtTIN.BackColor
    txtAmtRec.BackColor = txtTIN.BackColor
    txtAmtChange.BackColor = txtTIN.BackColor
    
    txtCheckNo.Locked = False
    txtChkDate.Locked = False
    txtChkAmt.Locked = False
    txtChkAcct.Locked = False
    txtChkBank.Locked = False
    txtCheckNo.BackColor = txtTIN.BackColor
    txtChkDate.BackColor = txtTIN.BackColor
    txtChkAmt.BackColor = txtTIN.BackColor
    txtChkAcct.BackColor = txtTIN.BackColor
    txtChkBank.BackColor = txtTIN.BackColor
    
    udORCR.Enabled = True
    cboAllocationNo.Clear
    
    chkCheck.Enabled = True
    
    txtDate.Enabled = True
    txtORno.Enabled = True
    
    txtTTLAllocation.Text = ""
    txtTax.Text = ""
    txtTaxAmount.Text = ""
    
    
    PrevCustomer = ""
    txtRefNo.Text = ""
    txtRefNo.Visible = True
    lblRefNo.Caption = "REFERENCE NO."
    lblRefNo.Visible = True
    txtORno.Text = ""
    txtDate.Text = "__/__/____"
    txtRecFrom.Text = ""
    txtAddress.Text = ""
    txtSumOf.Text = ""
    txtTIN.Text = ""
    txtPayable.Text = ""
    txtAmtRec.Text = ""
    txtAmtChange.Text = ""
    'txtEmpID.Text = ""
    txtJobID.Text = ""
    txtPrevRef.Text = ""
    txtCustomerID.Text = ""
    cboAcctCust.Text = ""
    cmdDetails.Caption = "Details"
    flxDetails.Rows = 0
    flxDetails.Rows = 1
    'flxBal.Visible = False
    flxBal.Rows = 1
    flxBal.TextMatrix(0, 0) = "INV. NO."
    flxBal.TextMatrix(0, 1) = "J.O. NO."
    txtTotalBal.Text = ""
    txtTotalDue.Text = ""
    cboOthers.Text = ""
    cboOthers.Visible = False
    txtJOamt.Text = ""
    txtPartsMan.Text = ""
    chkQckSrvc.Visible = False
    lstChrgInvNo.Clear
    txtAmtPd.Visible = False
    chkInvoice.Tag = "SKIP"
    chkInvoice.Value = 0
    chkJobOrder.Tag = "SKIP"
    chkJobOrder.Value = 0
    chkInvoice.Visible = False
    chkJobOrder.Visible = False
    txtPaymentID.Text = ""
    txtPaymentID.Tag = ""
    txtTotal.Text = ""
    lblTotal.Caption = "TOTAL AMT. PAID: "
    lblTotal.Visible = False
    txtTotal.Visible = False
    txtTotal.Text = ""
    flxBal.Visible = True
    flxOrder.Visible = False
    
    'CLEAR LEDGER
    txtInvNo.Text = ""
    txtInvDate.Text = ""
    txtInvAmt.Text = ""
    txtTerms.Text = ""
    lblDue.Caption = "AMOUNT DUE"
    txtDue.Text = ""
    txtBal.Text = ""
    flxLedger.Rows = 1
    flxLedger.Rows = 2
    
    'Clear Checks
    chkCheck.Value = 0
    txtCheckNo.Text = ""
    txtChkDate.Text = ""
    txtChkAmt.Text = ""
    txtChkAcct.Text = ""
    txtChkBank.Text = ""
    txtChkPRid.Text = ""
    cboCheck.Clear
    cboCheck.Text = ""
    isOthers = 1
    
End Sub

Private Sub LockData(bln As Boolean)
    Frame3.Enabled = Not bln
    Frame4.Enabled = Not bln
End Sub

'Commented by SMC
'2016-08-04
'Public Sub ShowPic()
'    On Error Resume Next
'
'    Open "pictemp" For Binary Access Write As lngDataFile
'        lngLengh = rsORinChrg.Fields(2).ActualSize
'        intChunks = lngLengh \ ChunkSize
'        intFragment = lngLengh Mod ChunkSize
'        ReDim Chunk(intFragment)
'        Chunk() = rsORinChrg.Fields(2).GetChunk(intFragment)
'        Put lngDataFile, , Chunk()
'        For i = 1 To intChunks
'           ReDim Buffer(ChunkSize)
'            Chunk() = rsORinChrg.Fields("EmpPicture").GetChunk(ChunkSize)
'            Put lngDataFile, , Chunk()
'        Next i
'    Close lngDataFile
'
'    filename = "pictemp"
'    imgOptr.Picture = LoadPicture(filename)
'End Sub


Private Function NumToText(dblVal As Double) As String
    Static Ones(0 To 9) As String
    Static Teens(0 To 9) As String
    Static Tens(0 To 9) As String
    Static Thousands(0 To 4) As String
    Static bInit As Boolean
    Dim i As Integer, bAllZeros As Boolean, bShowThousands As Boolean
    Dim strVal As String, strBuff As String, strTemp As String
    Dim nCol As Integer, nChar As Integer
    Debug.Assert dblVal >= 0

    If bInit = False Then
       
        bInit = True
        Ones(0) = "zero"
        Ones(1) = "one"
        Ones(2) = "two"
        Ones(3) = "three"
        Ones(4) = "four"
        Ones(5) = "five"
        Ones(6) = "six"
        Ones(7) = "seven"
        Ones(8) = "eight"
        Ones(9) = "nine"
        Teens(0) = "ten"
        Teens(1) = "eleven"
        Teens(2) = "twelve"
        Teens(3) = "thirteen"
        Teens(4) = "fourteen"
        Teens(5) = "fifteen"
        Teens(6) = "sixteen"
        Teens(7) = "seventeen"
        Teens(8) = "eighteen"
        Teens(9) = "nineteen"
        Tens(0) = ""
        Tens(1) = "ten"
        Tens(2) = "twenty"
        Tens(3) = "thirty"
        Tens(4) = "forty"
        Tens(5) = "fifty"
        Tens(6) = "sixty"
        Tens(7) = "seventy"
        Tens(8) = "eighty"
        Tens(9) = "ninety"
        Thousands(0) = ""
        Thousands(1) = "thousand"   'US numbering
        Thousands(2) = "million"
        Thousands(3) = "billion"
        Thousands(4) = "trillion"
    End If
   On Error GoTo NumToTextError
    If dblVal - Int(dblVal) > 0 Then
        strBuff = "PESOS and " & Format((dblVal - Int(dblVal)) * 100, "00") & "/100"
    Else
        strBuff = "PESOS ONLY"
    End If
     strVal = CStr(Int(dblVal))
    bAllZeros = True
    For i = Len(strVal) To 1 Step -1
        nChar = Val(Mid$(strVal, i, 1))
        nCol = (Len(strVal) - i) + 1
        Select Case (nCol Mod 3)
            Case 1
                bShowThousands = True
                If i = 1 Then
                 
                    strTemp = Ones(nChar) & " "
                ElseIf Mid$(strVal, i - 1, 1) = "1" Then
                    
                    strTemp = Teens(nChar) & " "
                    i = i - 1
                ElseIf nChar > 0 Then
                    
                    strTemp = Ones(nChar) & " "
                Else
                     bShowThousands = False
                    If Mid$(strVal, i - 1, 1) <> "0" Then
                        bShowThousands = True
                    ElseIf i > 2 Then
                        If Mid$(strVal, i - 2, 1) <> "0" Then
                            bShowThousands = True
                        End If
                    End If
                    strTemp = ""
                End If
                 If bShowThousands Then
                    If nCol > 1 Then
                        strTemp = strTemp & Thousands(nCol \ 3)
                        If bAllZeros Then
                            strTemp = strTemp & " "
                        Else
                            strTemp = strTemp & " "
                         End If
                    End If
                     bAllZeros = False
                End If
                strBuff = strTemp & strBuff
            Case 2
                If nChar > 0 Then
                    If Mid$(strVal, i + 1, 1) <> "0" Then
                        strBuff = Tens(nChar) & " " & strBuff
                    Else
                        strBuff = Tens(nChar) & " " & strBuff
                    End If
                End If
            Case 0
                If nChar > 0 Then
                    strBuff = Ones(nChar) & " hundred " & strBuff
                End If
        End Select
    Next i
    'strBuff = UCase$(Left$(strBuff, 1)) & Mid$(strBuff, 2)
    strBuff = UCase$(strBuff) '& Mid$(strBuff, 2)
EndNumToText:
    NumToText = strBuff
    Exit Function
NumToTextError:
    strBuff = "#Error#"
    Resume EndNumToText
End Function


Private Sub cmdSaveCancel(bln As Boolean)
    Dim i As Integer
    
    'If InStr(SysBranch, "PANABO") > 0 Then
    '    If Bln Then
    '        For i = 0 To 4
    '            cmdTrans(i).Visible = False
    '        Next i
    '        cmdTrans(7).Visible = False
    '        cmdTrans(6).Visible = True
    '        cmdTrans(5).Visible = True
    '        cmdTrans(5).Left = 0
    '        cmdTrans(6).Left = cmdTrans(4).Width
    '    Else
    '        For i = 0 To 4
    '            cmdTrans(i).Visible = True
    '        Next i
    '        cmdTrans(6).Visible = False
    '        cmdTrans(5).Visible = False
    '        cmdTrans(7).Visible = True
    '        cmdTrans(7).Left = cmdTrans(2).Left + cmdTrans(2).Width
    '        cmdTrans(3).Left = cmdTrans(7).Left + cmdTrans(7).Width
    '        cmdTrans(4).Left = cmdTrans(3).Left + cmdTrans(3).Width
    ' '   End If
    'Else
    
        If bln Then
            For i = 0 To 4
                cmdTrans(i).Visible = False
            Next i
            cmdTrans(6).Visible = True
            cmdTrans(5).Visible = True
            cmdTrans(5).Left = 0
            cmdTrans(6).Left = cmdTrans(4).Width
        Else
            For i = 0 To 4
                cmdTrans(i).Visible = True
            Next i
            cmdTrans(6).Visible = False
            cmdTrans(5).Visible = False
        End If
    'End If
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
            rsAcctCust.MoveFirst
            rsAcctCust.Move i
            txtCustomerID.Text = rsAcctCust!CUSTOMERID
            txtAddress.Text = rsAcctCust!Address
            txtPayable.Text = ""
            
            If PayType = 2 Then 'Clear Ledger
                txtPayable.Text = ""
                txtAmtRec.Text = ""
                txtAmtChange.Text = ""
                isOthers = 0
                
                txtTTLAllocation.Text = ""
                txtTax.Text = ""
                txtTaxAmount.Text = ""
                
                
                txtCheckNo.Text = ""
                txtChkDate.Text = ""
                txtChkAmt.Text = ""
                txtChkAcct.Text = ""
                txtChkBank.Text = ""
                

                txtInvNo.Text = ""
                txtInvDate.Text = ""
                txtInvAmt.Text = ""
                txtTerms.Text = ""
                lblDue.Caption = "AMOUNT DUE"
                txtDue.Text = ""
                txtBal.Text = ""
                flxLedger.Rows = 1
                flxLedger.Rows = 2
            End If
            j = Len(strTotal) - Len(strPartial)
            If j <> 0 Then
                CustSrch = True
                .SelText = Right$(strTotal, j)
                .SelStart = Len(strPartial)
                .SelLength = j
            End If
        Else
            txtCustomerID.Text = ""
            txtAddress.Text = ""
            txtRefNo.Text = ""
            txtPayable.Text = ""
        End If
    End With
    If PayType <> 1 Then
        If Not (PayType = 3 And OnAcctPCV.Value = 0) Then flxDetails.Rows = 1 'Clear Details xcept DP
    End If
    flxBal.Rows = 1
    txtTotal.Text = ""
    'flxBal.Visible = False
End Sub

Private Sub cboAcctCust_Click()
'Modify by SMC
'2016-08-05
On Error GoTo er
    Call cboAcctCust_Change
    
    DoEvents
    
    If PayType = 2 Then
        txtAmtRec.SetFocus
    End If
Exit Sub
er:
    MsgBox Err.Description, vbCritical, "Error->cboAcctCust_Click()"
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
On Error GoTo er
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        If PayType = 1 Or (PayType = 2 And chkQckSrvc.Value = 0) Then
            txtPayable.SetFocus
        Else
            txtAmtRec.SetFocus
        End If
    End If

Exit Sub
er:
    MsgBox Err.Description, vbCritical, "Error->cboAcctCust_KeyPress()"
End Sub

Function RemarksOnAcct() As String
    Dim i As Integer
    Dim myRem As String
    
    myRem = IIf(cboPayType.Text <> "", cboPayType.Text, "")
    With flxBal
        For i = 1 To .Rows - 1
            If Val(flxAcct.TextMatrix(i, 3)) > 0 Then
                If Val(.TextMatrix(i, 4)) = 0 Then 'new balance=0
                    myRem = IIf(myRem <> "", myRem & ", ", "") & IIf(.TextMatrix(i, 0) = "", "J.O.# ", "Inv.# ") & .TextMatrix(i, 0) & "-Full"
                Else
                    myRem = IIf(myRem <> "", myRem & ", ", "") & IIf(.TextMatrix(i, 0) = "", "J.O.# ", "Inv.# ") & .TextMatrix(i, 0) & "-Prtl"
                End If
            End If
        Next i
    End With
    RemarksOnAcct = myRem
End Function

Private Sub DisplayPCV()
    Dim rsPCV As New ADODB.Recordset
    Dim Qry As String
    Dim TtlPCV As Currency, Ans As Integer
    
    flxBal.TextMatrix(0, 0) = "PCV DATE"
    flxBal.TextMatrix(0, 1) = "PCV NO."
    TtlPCV = 0
    '***Access
    'Qry = "SELECT DISBURSEMENT.CVDATE, DISBURSEMENT.CVNUMBER, ONACCTPCV.BALANCE, ONACCTPCV.CUSTOMERID " & _
          "FROM DISBURSEMENT INNER JOIN ONACCTPCV ON DISBURSEMENT.VOUCHERID = ONACCTPCV.VOUCHERID Where (((ONACCTPCV.BALANCE) > 0) And ((ONACCTPCV.CUSTOMERID) = " & txtCustomerID.Text & ")) ORDER BY DISBURSEMENT.CVDATE;"
    'mySQL
    Qry = "SELECT DISBURSEMENT.CVDATE, DISBURSEMENT.CVNUMBER, ONACCTPCV.BALANCE, ONACCTPCV.CUSTOMERID " & _
          "FROM DISBURSEMENT INNER JOIN ONACCTPCV ON DISBURSEMENT.VOUCHERID = ONACCTPCV.VOUCHERID Where ONACCTPCV.BALANCE > 0 And ONACCTPCV.CUSTOMERID = " & txtCustomerID.Text & " ORDER BY DISBURSEMENT.CVDATE;"

'    Qry = "SELECT Disbursement.CVDATE, Disbursement.CVNUMBER,  Disbursement.AMOUNT, Disbursement.CUSTOMERID " & _
'            "From Disbursement Where Disbursement.CUSTOMERID = " & txtCustomerID.Text & " and amount <> 0;"
'
    With rsPCV
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            flxBal.Rows = 1
            Do While Not .EOF
                flxBal.AddItem !CVDate & vbTab & Format(!CVNUMBER, "0000") & vbTab & Format(!Amount, "#,##0.00") & vbTab & Format(!Amount, "#,##0.00")
                TtlPCV = TtlPCV + !Amount
                .MoveNext
            Loop
        End If
        .Close
    End With
    txtTotal.Text = Format(TtlPCV, "#,##0.00")
    flxBal.Visible = True
    txtTotal.Visible = True
    lblTotal.Visible = True
    If TtlPCV > 0 Then
        Ans = MsgBox("Do you want to pay On-Account from PCV?", vbQuestion + vbYesNo, "Pay On-Account")
        If Ans = vbYes Then
            txtPayable.Text = Format(TtlPCV, "###0.00")
            OnAcctPCV.Value = 1
            Exit Sub
        End If
    End If
    OnAcctPCV.Value = 0
    flxBal.Rows = 1
    flxBal.Rows = 2
    txtTotal.Text = ""
    
End Sub

Private Sub cboAcctCust_LostFocus()
On Error GoTo er
    Dim Qry As String, MyRef As String, RetAmt As Currency
    Dim i As Integer
    
    If PrevCustomer = "" Or PrevCustomer <> Trim(cboAcctCust.Text) Then
        cboAllocationNo.Clear
    End If
    
    If txtCustomerID.Text = "" Or PayType = 1 Then
        Exit Sub
    End If
    
    Me.Enabled = False
    Screen.MousePointer = vbHourglass

    
    DoEvents
    
    If PayType = 3 Then
        Call DisplayPCV
        Me.Enabled = True
        Screen.MousePointer = vbDefault
        
        txtPayable.SetFocus
        Exit Sub
    End If
    tabWait.Caption = ""
    tabWait.Visible = True
    'old Qry = "SELECT IIf(IsNull([ONACCTERMS.SPINVNO]),[JOBORDERNO],[ONACCTERMS.SPINVNO]) AS REFNO, ONACCTERMS.TERMS, ONACCTERMS.PAYSCHEME, ONACCTERMS.NPAYS, ONACCTERMS.PAYSTART, ONACCTERMS.BILLINGAMT, ONACCTERMS.STATUS, ONACCTERMS.TTLAMTPAID, ONACCTERMS.BALANCE, ONACCTERMS.CUSTOMERID, SPINVOICE.ORNO, SPINVOICE.ORNO, ONACCTERMS.TRANSDATE, ONACCTERMS.ACCTAMOUNT, ONACCTERMS.JOBORDERID, SPINVOICE.ORNO " & _
          "FROM ONACCTERMS INNER JOIN SPINVOICE ON ONACCTERMS.SPINVNO = SPINVOICE.SPINVNO WHERE (((ONACCTERMS.STATUS)='OPEN') AND ((ONACCTERMS.CUSTOMERID)=" & txtCustomerID.Text & ")) ORDER BY ONACCTERMS.JOBORDERID, ONACCTERMS.REFNO;"
    'old Qry = "SELECT ONACCTERMS.SPINVNO, ONACCTERMS.JOBORDERNO, ONACCTERMS.TERMS, ONACCTERMS.PAYSCHEME, ONACCTERMS.NPAYS, ONACCTERMS.PAYSTART, ONACCTERMS.BILLINGAMT, ONACCTERMS.STATUS, ONACCTERMS.TTLAMTPAID, ONACCTERMS.BALANCE, ONACCTERMS.CUSTOMERID, ONACCTERMS.NPAYS, ONACCTERMS.NPAYS, ONACCTERMS.TRANSDATE, ONACCTERMS.ACCTAMOUNT, ONACCTERMS.JOBORDERID, ONACCTERMS.NPAYS " & _
          "From ONACCTERMS WHERE (((ONACCTERMS.STATUS)='OPEN') AND ((ONACCTERMS.CUSTOMERID)=" & txtCustomerID.Text & ")) ORDER BY ONACCTERMS.JOBORDERID, ONACCTERMS.SPINVNO;"
    
    '***Access
    'Qry = "SELECT ONACCTERMS.SPINVNO, ONACCTERMS.TERMS, ONACCTERMS.PAYSCHEME, ONACCTERMS.NPAYS, ONACCTERMS.PAYSTART, ONACCTERMS.BILLINGAMT, ONACCTERMS.STATUS, ONACCTERMS.TTLAMTPAID, ONACCTERMS.BALANCE, ONACCTERMS.CUSTOMERID, ONACCTERMS.TRANSDATE, ONACCTERMS.ACCTAMOUNT, ONACCTERMS.JOBORDERID " & _
          "From ONACCTERMS WHERE ((Not (ONACCTERMS.SPINVNO) Is Null) AND ((ONACCTERMS.STATUS)='OPEN') AND ((ONACCTERMS.CUSTOMERID)=" & txtCustomerID.Text & ")) ORDER BY ONACCTERMS.SPINVNO, ONACCTERMS.TRANSDATE;"
    'mySQl
    Qry = "SELECT ONACCTERMS.SPINVNO, ONACCTERMS.TERMS, ONACCTERMS.PAYSCHEME, ONACCTERMS.NPAYS, ONACCTERMS.PAYSTART, ONACCTERMS.BILLINGAMT, ONACCTERMS.STATUS, ONACCTERMS.TTLAMTPAID, ONACCTERMS.BALANCE, ONACCTERMS.CUSTOMERID, ONACCTERMS.TRANSDATE, ONACCTERMS.ACCTAMOUNT, ONACCTERMS.JOBORDERID " & _
          "From ONACCTERMS WHERE Not (ONACCTERMS.SPINVNO) Is Null AND ONACCTERMS.STATUS='OPEN' AND ONACCTERMS.CUSTOMERID=" & txtCustomerID.Text & " ORDER BY ONACCTERMS.SPINVNO, ONACCTERMS.TRANSDATE;"

    If rsAcct.State = adStateOpen Then rsAcct.Close
    flxAcct.Rows = 1
    With rsAcct         'REtrieve Invoice Only
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                MyRef = IIf(IsNull(!SPINVNO), "", !SPINVNO)
                If MyRef <> "" Then
                                        '0-Invoice No.| 1-Terms | 2-Pay Scheme | 3-Npays | 4-Pay Start | 5-Billing Amount |6- Status |7-Total Amt Pd |8-Balance |9-CustomerID|10-Current AmtPaid|11-New Balance|12-Inv. Date |13-Inv Amount | 14 - JobID | 15-Ret amt
                    flxAcct.AddItem MyRef & vbTab & !Terms & vbTab & !PayScheme & vbTab & !Npays & vbTab & !PayStart & vbTab & !BILLINGAMT & vbTab & !STATUS & vbTab & !TTLAMTPAID & vbTab & !Balance & vbTab & !CUSTOMERID & vbTab & vbTab & vbTab & !TRANSDATE & vbTab & !AcctAmount & vbTab & 0
                End If
                .MoveNext
            Loop
        End If
        .Close
    End With
    
    '***Access
    'Qry = "SELECT ONACCTERMS.JOBORDERNO, ONACCTERMS.TERMS, ONACCTERMS.PAYSCHEME, ONACCTERMS.NPAYS, ONACCTERMS.PAYSTART, ONACCTERMS.BILLINGAMT, ONACCTERMS.STATUS, ONACCTERMS.TTLAMTPAID, ONACCTERMS.BALANCE, ONACCTERMS.CUSTOMERID, ONACCTERMS.TRANSDATE, ONACCTERMS.ACCTAMOUNT, ONACCTERMS.JOBORDERID " & _
          "From ONACCTERMS WHERE (((ONACCTERMS.SPINVNO) Is Null) AND ((ONACCTERMS.STATUS)='OPEN') AND ((ONACCTERMS.CUSTOMERID)=" & txtCustomerID.Text & ")) ORDER BY ONACCTERMS.JOBORDERNO, ONACCTERMS.TRANSDATE;"
    'mySQL
    Qry = "SELECT ONACCTERMS.JOBORDERNO, ONACCTERMS.TERMS, ONACCTERMS.PAYSCHEME, ONACCTERMS.NPAYS, ONACCTERMS.PAYSTART, ONACCTERMS.BILLINGAMT, ONACCTERMS.STATUS, ONACCTERMS.TTLAMTPAID, ONACCTERMS.BALANCE, ONACCTERMS.CUSTOMERID, ONACCTERMS.TRANSDATE, ONACCTERMS.ACCTAMOUNT, ONACCTERMS.JOBORDERID " & _
          "From ONACCTERMS WHERE (ONACCTERMS.SPINVNO) Is Null AND ONACCTERMS.STATUS='OPEN' AND ONACCTERMS.CUSTOMERID=" & txtCustomerID.Text & " ORDER BY ONACCTERMS.JOBORDERNO, ONACCTERMS.TRANSDATE;"
    
    With rsAcct         'REtrieve JOB ORDER Only
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                MyRef = IIf(IsNull(!JOBORDERNO), "", !JOBORDERNO)
                If MyRef <> "" Then
                                        '0-Invoice No.| 1-Terms | 2-Pay Scheme | 3-Npays | 4-Pay Start | 5-Billing Amount |6- Status |7-Total Amt Pd |8-Balance |9-CustomerID|10-Current AmtPaid|11-New Balance|12-Inv. Date |13-Inv Amount | 14 - JobID | 15-Ret amt
                    flxAcct.AddItem MyRef & vbTab & !Terms & vbTab & !PayScheme & vbTab & !Npays & vbTab & !PayStart & vbTab & !BILLINGAMT & vbTab & !STATUS & vbTab & !TTLAMTPAID & vbTab & !Balance & vbTab & !CUSTOMERID & vbTab & vbTab & vbTab & !TRANSDATE & vbTab & !AcctAmount & vbTab & !JOBORDERID
                End If
                .MoveNext
            Loop
        End If
        .Close
    End With
    
    Pbar.Visible = False
    If flxAcct.Rows = 1 Then
        tabWait.Visible = False
        
        DoEvents
        GoTo LastGate
    Else
        If flxAcct.Rows >= 2 And flxDetails.Rows > 1 Then
            If flxDetails.TextMatrix(1, 1) = "Inv.# " & flxAcct.TextMatrix(1, 0) Then
                tabWait.Visible = False
                
                DoEvents
                GoTo LastGate
            End If
        End If
    End If
    
    'Process Billing here
    Call DisplayAcctDetails
    If chkInvoice.Value = 1 And chkJobOrder.Value = 0 Then
        chkInvoice.Enabled = False
        chkJobOrder.Enabled = False
    ElseIf chkInvoice.Value = 0 And chkJobOrder.Value = 1 Then
        chkJobOrder.Enabled = False
        chkInvoice.Enabled = False
    ElseIf chkInvoice.Value = 1 And chkJobOrder.Value = 1 Then
        chkJobOrder.Enabled = True
        chkInvoice.Enabled = True
    End If
    
    DoEvents
    
    ' Display Current Customer Allocation Record
LastGate:
    If DidRecordEntry = True Then
        If PayType = 2 Then 'Charge Invoice`
            If PrevCustomer = "" Or PrevCustomer <> Trim(cboAcctCust.Text) Or cboAllocationNo.ListCount = 0 Then
                txtCustomerID.Text = Trim(txtCustomerID.Text)
                If txtCustomerID.Text = "" Then
                    Me.Enabled = True
                    Screen.MousePointer = vbDefault
                    
                    MsgBox "Customer ID is missing!", vbExclamation, "Validation!"
                    cboAcctCust.SetFocus
                    Exit Sub
                End If
                
                Dim tmpRs As New ADODB.Recordset
                Call modAllocation.sp_customer_list_active_allocation(txtCustomerID.Text, tmpRs)
                Call modTools.addToListObject(cboAllocationNo, tmpRs, , , , 1)
            End If
        End If
    End If
    PrevCustomer = Trim(cboAcctCust.Text)
    
    tabWait.Visible = False
    chkInvoice.Tag = ""
    DoEvents
    
'    If Val(Replace(txtAmtRec.Text, ",", "")) > 0 Then
    If Val(Replace(txtPayable.Text, ",", "")) > 0 Then
        txtSumOf.Text = NumToText(CDbl(txtPayable.Text))
    Else
        txtSumOf.Text = ""
    End If
    Me.Enabled = True
    Screen.MousePointer = vbDefault
Exit Sub
er:
    Me.Enabled = True
    Screen.MousePointer = vbDefault
    MsgBox Err.Description, vbCritical, "Error->cboAcctCust_LostFocus()"
End Sub

Private Sub DisplayAcctDetails()
    Dim i As Integer, myBN As Integer, myOD As Currency, MyBillAmt As Currency, TmpOD As Currency
    Dim myBillDate As Date, GrndTtl As Currency, TtlBal As Currency
    Dim cInv As Integer, cJob As Integer
    
    With flxAcct
        GrndTtl = 0
        TtlBal = 0
        Pbar.Visible = True
        Pbar.Value = 0
        Pbar.Max = .Rows - 1
        cInv = 0
        cJob = 0
        For i = 1 To .Rows - 1
                            'PayStart                   'Date
            If IsDate(.TextMatrix(i, 4)) Then
                If DateValue(.TextMatrix(i, 4)) >= DateValue(txtDate.Text) Then  'wla pa naabot ang Start of Payment date
                    myBillDate = .TextMatrix(i, 4) 'PayStart
                    myOD = 0 'Over Due
                    myBN = 1 'Bill no
                    If Val(.TextMatrix(i, 8)) < Val(.TextMatrix(i, 5)) Then     'if Balance<Billing Amt
                        MyBillAmt = Val(.TextMatrix(i, 8)) 'Balance
                    Else
                        MyBillAmt = Val(.TextMatrix(i, 5)) 'Billing Amt
                    End If
                    'Display in flxDetails
                    flxDetails.AddItem vbTab & IIf(Val(.TextMatrix(i, 14)) > 0, "J.O.# ", "Inv.# ") & .TextMatrix(i, 0)
                    flxDetails.AddItem vbTab & Space$(4) & "Balance" & vbTab & Format$(Val(.TextMatrix(i, 8)), "#,##0.00") & Space$(5)
                    flxDetails.AddItem vbTab & Space$(4) & "Amt. Due" & vbTab & Format$(MyBillAmt, "#,##0.00") & Space$(5)
                    .TextMatrix(i, 10) = MyBillAmt
                    .TextMatrix(i, 5) = MyBillAmt
                    .TextMatrix(i, 11) = Val(.TextMatrix(i, 8)) - MyBillAmt
                Else
                    'Process per payment Scheme
                    myBillDate = .TextMatrix(i, 4) 'PayStart
                    myBN = 0 'Bill no
                    Do While myBillDate <= DateValue(txtDate.Text) And myBN <= Val(.TextMatrix(i, 3))   'Loop while billing date <current date and Billing No < #payments
                        If .TextMatrix(i, 2) = "WEEKLY" Then
                            If Day(myBillDate) = 21 Then
                                myBillDate = DateAdd("d", -1, DateSerial(Year(DateAdd("m", 1, myBillDate)), Month(DateAdd("m", 1, myBillDate)), 1))
                            Else
                                myBillDate = DateAdd("d", 7, myBillDate)
                            End If
                        ElseIf .TextMatrix(i, 2) = "BI-MONTHLY" Then
                            myBillDate = DateAdd("d", 15, myBillDate)
                        ElseIf .TextMatrix(i, 2) = "MONTHLY" Then
                            myBillDate = DateAdd("m", 1, myBillDate)
                        End If
                        myBN = myBN + 1
                    Loop           'No. of Payments
                    If myBN >= Val(.TextMatrix(i, 3)) Then 'Billing = NPays
                        MyBillAmt = 0
                        myOD = Val(.TextMatrix(i, 8)) 'Balance
                    Else
                        If Val(.TextMatrix(i, 8)) < Val(.TextMatrix(i, 5)) Then     'if Balance<Billing Amt
                            MyBillAmt = Val(.TextMatrix(i, 8)) 'Balance
                        Else
                            MyBillAmt = Val(.TextMatrix(i, 5)) 'Billing Amt
                        End If
                        TmpOD = ((myBN - 1) * Val(.TextMatrix(i, 5))) - Val(.TextMatrix(i, 7)) 'Exact Amt upto Billing less Total Amt. Paid
                        If TmpOD > 0 Then
                            If TmpOD > Val(.TextMatrix(i, 8)) - MyBillAmt Then  'if temp OverDue > (Balance-BillAmt)
                                myOD = Val(.TextMatrix(i, 8)) - TmpOD
                            Else
                                myOD = TmpOD
                            End If
                        Else
                            myOD = 0
                        End If
                    End If
                    'Display in flxDetails
                    flxDetails.AddItem vbTab & IIf(Val(.TextMatrix(i, 14)) > 0, "J.O.# ", "Inv.# ") & .TextMatrix(i, 0)
                    flxDetails.AddItem vbTab & Space$(4) & "Balance" & vbTab & Format$(Val(.TextMatrix(i, 8)), "#,##0.00") & Space$(5)
                    flxDetails.AddItem vbTab & Space$(4) & "Amt. Due" & vbTab & Format$(MyBillAmt, "#,##0.00") & Space$(5)
                    If myOD > 0 Then
                        flxDetails.AddItem vbTab & Space$(4) & "Over Due" & vbTab & Format$(myOD, "#,##0.00") & Space$(5)
                        flxDetails.AddItem vbTab & "Total:" & vbTab & Format$(MyBillAmt + myOD, "#,##0.00") & Space$(5)
                    End If
                    .TextMatrix(i, 10) = MyBillAmt + myOD
                    .TextMatrix(i, 5) = MyBillAmt + myOD
                    .TextMatrix(i, 11) = Val(.TextMatrix(i, 8)) - (MyBillAmt + myOD)
                End If
            End If
            If Val(.TextMatrix(i, 14)) = 0 Then
                cInv = cInv + 1
            Else
                cJob = cJob + 1
            End If
            GrndTtl = GrndTtl + MyBillAmt + myOD
            TtlBal = TtlBal + Val(.TextMatrix(i, 8))
            Pbar.Value = Pbar.Value + 1
            DoEvents
        Next i
        chkInvoice.Caption = "INVOICE (" & cInv & ")"
        chkJobOrder.Caption = "JOB ORDER (" & cJob & ")"
        If Val(Replace(txtPayable.Text, ",", "")) > 0 Then
            txtSumOf.Text = NumToText(CDbl(txtPayable.Text))
        Else
            txtSumOf.Text = ""
        End If
        If PayType <> 2 Then
            txtPayable.Text = Format$(GrndTtl, "#####0.00")
        End If
        
        txtTotalDue.Text = GrndTtl
        txtTotalBal.Text = TtlBal
        
        If Not (DidRecordEntry = True And PayType = 2) Then
            txtPayable.Locked = False
        End If
        
        If .Rows > 2 Then
            flxDetails.AddItem ""
            flxDetails.AddItem vbTab & "Total Balance" & vbTab & Format$(TtlBal, "#,##0.00") & Space$(5)
            flxDetails.AddItem vbTab & "Total Amt Due" & vbTab & Format$(GrndTtl, "#,##0.00") & Space$(5)
            'txtPayable.Locked = True
        End If
    End With
    Pbar.Visible = False
End Sub

Private Sub cboAcctTitle_Change()
    Dim i As Long, j As Long
    Dim strPartial As String, strTotal As String
    Dim x As Integer
    
    If AcctSrch Then
        AcctSrch = False
        Exit Sub
    End If
    With cboAcctTitle
        strPartial = .Text
        i = SendMessage(.hwnd, CB_FINDSTRING, -1, ByVal strPartial) 'If match found, append unmatched characters
        If i <> -1 Then
            strTotal = .List(i)
            j = Len(strTotal) - Len(strPartial)
            If j <> 0 Then
                AcctSrch = True
                .SelText = Right$(strTotal, j)
                .SelStart = Len(strPartial)
                .SelLength = j
            End If
        End If
    End With
End Sub

Private Sub cboAcctTitle_Click()
Call cboAcctTitle_Change
End Sub

Private Sub cboAcctTitle_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then AcctSrch = True
End Sub

Private Sub cboAcctTitle_KeyPress(KeyAscii As Integer)
Dim cbolistIndex As Integer
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then
    
        If cboAcctTitle.Text = "" Then
            Exit Sub
        End If
        
        cbolistIndex = modObjTools.getListIndex(cboAcctTitle, cboAcctTitle.Text)
        
        If cbolistIndex < 0 Then
            cboAcctTitle.SetFocus
        Else
            cboAcctTitle.ListIndex = cbolistIndex
        End If
        txtAmtRec.SetFocus
    End If
End Sub

Private Sub cboAcctTitle_LostFocus()
    myAcctCode = ""
    If cboAcctTitle.Text = "" Then Exit Sub
    myAcctCode = arrAcctTitle(cboAcctTitle.ListIndex, 0)
End Sub

Private Sub cboAllocationNo_Click()
On Error GoTo er
    'Allocation  Data Format
    '2 -> ( JOB ORDER )
    Dim strArr() As String
    Dim strAllocationNo As String
    
    If DidRecordEntry = False Then Exit Sub
    
    txtDate.Enabled = True
    txtORno.Enabled = True
    
    txtPayable.Text = ""
    txtAmtRec.Text = ""
    txtAmtChange.Text = ""
    
    txtTTLAllocation.Text = ""
    txtTax.Text = ""
    txtTaxAmount.Text = ""
    
    AllocationNo = ""
    TRANSTYPE = ""
    
    strAllocationNo = Trim(cboAllocationNo.Text)
    If strAllocationNo = "" Then Exit Sub
    
    Me.Enabled = False
    Screen.MousePointer = vbHourglass
    
    strArr() = Split(strAllocationNo, "->")
    'Retrieve Allocation Number
    AllocationNo = Trim(strArr(0))
    
    'Retrieve Transtype
    TRANSTYPE = strArr(1)
    TRANSTYPE = Replace(TRANSTYPE, "(", "")
    TRANSTYPE = Replace(TRANSTYPE, ")", "")
    TRANSTYPE = UCase(Trim(TRANSTYPE))
    
    If TRANSTYPE = "" Or AllocationNo = "" Then
        AllocationNo = ""
        TRANSTYPE = ""
        MsgBox "Allocation number or transaction type is not presented!", vbExclamation, "Validation"
        
        Me.Enabled = True
        Screen.MousePointer = vbDefault
        
        cboAllocationNo.SetFocus
        Exit Sub
    End If
    
    txtDate.Enabled = True
    txtORno.Enabled = True
    StrPAYTYPE = ""
    If TRANSTYPE = "SPARE PARTS" Then
        udORCR.Value = 1
        lblORno.Caption = "COLLECTION RECEIPT NO."
        
    ElseIf TRANSTYPE = "SERVICE" Then
       
        udORCR.Value = 0
        lblORno.Caption = "OFFICIAL RECEIPT NO."
        
    Else
        Dim Qry As String
        Qry = "SELECT PAYTYPE FROM  allocation Where ALLOCATIONID=" & AllocationNo & ";"
        StrPAYTYPE = GetFieldValue(Qry)
        
        If StrPAYTYPE <> "AKLDGMT" Then 'ACKNOWLEDGEMENT
            AllocationNo = ""
            TRANSTYPE = ""
            MsgBox "Allocation number or transaction type is not presented!", vbExclamation, "Validation"
            Me.Enabled = True
            Screen.MousePointer = vbDefault
            
            cboAllocationNo.SetFocus
            Exit Sub
        Else
            txtDate.Enabled = False
            txtORno.Enabled = False
            txtORno.Text = ""
        End If
        
    End If
    
    Dim tmpRs As New ADODB.Recordset
    Call modAllocation.sp_payments_list_allocation(CInt(AllocationNo), tmpRs)
    'ALLOCATIONID, CUSTOMERID, CUSTOMERNAME, ALLOCATEDDATE, TAXWHELDPERCENT, TAXWHELDAMOUNT, TTLALLOC_AMOUNT
    ',NETPAYABLE, ALLOCATIONSTATUS, ACCOUNTSTATUS, ACCOUNTTYPE, REFNO, REFAMOUNT, REFBALANCE, ALLOC_AMOUNT
    ',BALANCEDURINGALLOCATION, BILLAMOUNT, BILLINGNO, DUEDATE, PRNO, PRDATE, CHECKAMOUNT
    ',USABLEBAL, BANK, CHECKNO, TTLDEBIT, TTLCREDIT, CHECKBAL
    
    flxBal.Rows = 1
    If tmpRs.State = 1 Then
        If tmpRs.RecordCount > 0 Then
            Dim tmpINVNo As String
            Dim tmpJONo As String
            tmpRs.MoveFirst
            
            txtPayable.Text = tmpRs!NetPayable
            txtAmtRec.Text = tmpRs!NetPayable
            txtAmtChange.Text = "0.00"
            
            txtTTLAllocation.Text = tmpRs!TTLALLOC_AMOUNT
            txtTax.Text = tmpRs!TAXWHELDPERCENT
            txtTaxAmount.Text = tmpRs!TAXWHELDAMOUNT
            
            txtCheckNo.Text = tmpRs!CHECKNO
            txtChkDate.Text = tmpRs!CHECKDATE
            txtChkAmt.Text = tmpRs!CHECKBAL
            txtChkAcct.Text = tmpRs!CHKACCTNAME
            txtChkBank.Text = tmpRs!BANK
            
            While Not tmpRs.EOF
                '|^INV. NO.  |^J.O. NO.   |>BALANCE  |>AMT. PAID  |NEW BAL.
                tmpINVNo = ""
                tmpJONo = ""
                
                If Trim(tmpRs!ACCOUNTTYPE) = "SPARE PARTS" Then
                    tmpINVNo = tmpRs!RefNo
                Else
                    tmpJONo = tmpRs!RefNo
                End If
                JONoAcct = tmpJONo
                flxBal.AddItem _
                    tmpINVNo & _
                    vbTab & tmpJONo & _
                    vbTab & tmpRs!REFBALANCE & _
                    vbTab & tmpRs!ALLOC_AMOUNT & _
                    vbTab & tmpRs!BALANCEDURINGALLOCATION
                tmpRs.MoveNext
            Wend
             Call displayJOdetails
        End If
    End If
    If tmpRs.State <> 0 Then tmpRs.Close
    
    Me.Enabled = True
    Screen.MousePointer = vbDefault
Exit Sub
er:
    Me.Enabled = True
    Screen.MousePointer = vbDefault
    
    AllocationNo = ""
    TRANSTYPE = ""
    cboAllocationNo.ListIndex = -1
    MsgBox Err.Description, vbCritical, "Error->cboAllocationNo_Click()"
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

Private Sub cboCheck_Click()
    Dim rsChk As New ADODB.Recordset
    Dim Qry As String
    
    If cboCheck.ListIndex = -1 Or cboCheck.Text = "" Then Exit Sub
        
    '***Access
    'Qry = "SELECT CHECKS.CHECKNO, CHECKS.CHECKDATE, CHECKS.CHECKBAL, CHECKS.CHKACCTNAME, CHECKS.BANK, CHECKS.CHKSPRID " & _
          "From CHECKS WHERE (((CHECKS.CHECKNO)='" & cboCheck.Text & "'));"
    'mySQL
    Qry = "SELECT CHECKS.CHECKNO, CHECKS.CHECKDATE, CHECKS.CHECKBAL, CHECKS.CHKACCTNAME, CHECKS.BANK, CHECKS.CHKSPRID " & _
          "From CHECKS WHERE CHECKS.CHECKNO='" & cboCheck.Text & "' AND CHECKS.STATUS <> 'CANCELLED';"
    
    
    With rsChk
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            txtCheckNo.Text = !CHECKNO
            txtChkDate.Text = Format(!CHECKDATE, "mm/dd/yyyy")
            txtChkAmt.Text = Format(IIf(IsNull(!CHECKBAL), 0, !CHECKBAL), "#,##0.00")
            txtChkAcct.Text = !CHKACCTNAME
            txtChkBank.Text = !BANK
            txtChkPRid.Text = !CHKSPRID
            txtAmtRec.Text = Format(IIf(IsNull(!CHECKBAL), 0, !CHECKBAL), "#,##0.00")
        End If
        .Close
    End With
    If Val(txtChkPRid.Text) > 0 And PayType <> 3 Then
        Qry = "SELECT SUM(checkalloc.SPSORDERID + checkalloc.JOBORDERID) AS TTLID FROM checkalloc WHERE checkalloc.CHKSPRID = " & Val(txtChkPRid.Text) & ";"
        If Val(GetFieldValue(Qry)) > 0 Then Call CheckAllocation
    End If
End Sub

Private Sub cboOthers_GotFocus()
   Dim Tmp
    
   Const CB_SHOWDROPDOWN = &H14F
   Tmp = SendMessage(cboOthers.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub

Private Sub cboOthers_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        cboOthers.Text = flxDetails.TextMatrix(1, 1)
        cboAcctCust.SetFocus
        cboOthers.Visible = False
    End If
End Sub

Private Sub cboOthers_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        'If InStr(SysBranch, "PANABO") > 0 Then
        '    If PayType = 7 And cboOthers.Text <> "" Then
        '        Call DisplayCashInvoice
        '    End If
        '    txtAmtRec.SetFocus
        'Else
            cboAcctCust.SetFocus
        'End If
    End If
End Sub

Private Sub cboOthers_LostFocus()
    flxDetails.TextMatrix(1, 1) = Space(10) & cboOthers.Text
    cboOthers.Visible = False
End Sub

Private Sub DisplayCashInvoice()
    Dim rsCash As New ADODB.Recordset
    Dim rsOrder As New ADODB.Recordset
    Dim Qry As String
    Dim MySalesID As Long, Ttl As Currency
    
    Qry = "SELECT * From spinvoice Where spinvoice.INVDATE = '" & Format(Date, "yyyy-mm-dd") & "' AND " & _
          "spinvoice.INVTYPE = 'CASH' AND spinvoice.ORNO IS NULL AND spinvoice.SPINVNO = '" & cboOthers.Text & "';"
    With rsCash
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            txtRecFrom.Text = !chargeTo
            txtAddress.Text = !Address
            txtPayable.Text = Format(!Amount, "###0.00")
            MySalesID = !SPSORDERID
        End If
        .Close
    End With
    Qry = "SELECT SPSORDERDETAIL.QTY, SPSORDERDETAIL.UNIT, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPSORDERDETAIL.PRICE, SPSORDERDETAIL.QTY*SPSORDERDETAIL.PRICE AS TTLPRICE, SPSORDERDETAIL.SPAREPARTID, SPAREPARTS.QTYONHAND " & _
          "FROM SPSORDERDETAIL INNER JOIN SPAREPARTS ON SPSORDERDETAIL.SPAREPARTID = SPAREPARTS.SPAREPARTID WHERE SPSORDERDETAIL.QTY>0 AND SPSORDERDETAIL.SPSORDERID=" & MySalesID & " ORDER BY SPAREPARTS.PARTNUMBER;"
'    NoPrice = False
    With rsOrder
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            flxOrder.Rows = 1
            Ttl = 0
            Do While Not .EOF
                flxOrder.AddItem vbTab & !Qty & vbTab & !Unit & vbTab & !PartNumber & vbTab & !Description & vbTab & Format$(!Price, "###,##0.00") & vbTab & Format$(!TtlPrice, "###,##0.00") & vbTab & !SparepartID & vbTab & !QtyOnHand
                'If !Price = 0 Then NoPrice = True
                Ttl = Ttl + !TtlPrice
                .MoveNext
            Loop
            txtTotal.Text = Format$(Ttl, "###,##0.00")
            txtPayable.Text = Format$(Ttl, "#####0.00")
            txtSumOf.Text = NumToText(CDbl(Val(Replace(txtPayable.Text, ",", ""))))
        Else
            flxOrder.Rows = 1
            flxOrder.Rows = 2
            txtTotal.Text = ""
            txtPayable.Text = ""
            txtSumOf.Text = ""
        End If
        .Close
    End With

End Sub

Private Sub cboPayType_Click()
If cboPayType.Text = "CREDIT CARD PAYMENT" Then
    Frame6.Visible = True
End If

End Sub

Private Sub chkCheck_Click()
    Dim rsChk As New ADODB.Recordset
    Dim Qry As String
    
    If chkCheck.Value = 0 Then
        txtCheckNo.Visible = True
        cboCheck.Visible = False
        txtCheckNo.Text = ""
        txtChkDate.Text = ""
        txtChkAmt.Text = ""
        txtChkAcct.Text = ""
        txtChkBank.Text = ""
        Frame2.Enabled = False
        cboCheck.Clear
        txtChkPRid.Text = ""
        If txtChkPRid.Tag = "CHECK ALLOCATED" Then
            txtChkPRid.Tag = ""
            flxBal.Cols = 5
            flxBal.Rows = 1
            flxBal.TextMatrix(0, 1) = "J.O. NO."
        End If
    Else
        Frame2.Enabled = True
        '***Access
        'Qry = "SELECT CHECKS.CHECKNO, CHECKS.CHECKDATE, CHECKS.CHECKBAL, CHECKS.CHKACCTNAME, CHECKS.BANK, CHECKALLOC.CHKSPRID " & _
              "FROM CHECKS INNER JOIN CHECKALLOC ON CHECKS.CHKSPRID = CHECKALLOC.CHKSPRID " & _
              "WHERE (((CHECKS.CHECKBAL)>0) AND ((CHECKALLOC.CUSTOMERID)=" & Val(txtCustomerID.Text) & ") AND ((CHECKS.STATUS) Like '%CLEARED%'));"
        'mySQL
        'Qry = "SELECT CHECKS.CHECKNO, CHECKS.CHECKDATE, CHECKS.CHECKBAL, CHECKS.CHKACCTNAME, CHECKS.BANK, CHECKALLOC.CHKSPRID " & _
              "FROM CHECKS INNER JOIN CHECKALLOC ON CHECKS.CHKSPRID = CHECKALLOC.CHKSPRID " & _
              "WHERE CHECKS.CHECKBAL>=1 AND CHECKALLOC.CUSTOMERID=" & Val(txtCustomerID.Text) & " AND (CHECKS.STATUS Like '%CLEARED%' OR CHECKS.STATUS = 'PDC DUE');"

        If ThisBranch = "MRC2" Then
            Qry = "SELECT CHECKS.CHECKNO, CHECKS.CHECKDATE, CHECKS.CHECKBAL, CHECKS.CHKACCTNAME, CHECKS.BANK, CHECKS.CHKSPRID " & _
                  "FROM CHECKS WHERE CHECKS.CUSTOMERID=" & Val(txtCustomerID.Text) & " AND CHECKS.CHECKBAL>=1 AND (CHECKS.STATUS Like '%CLEARED%' OR CHECKS.STATUS = 'PDC DUE');"
        Else
            Qry = "SELECT DISTINCT checks.CHECKNO, checks.CHECKDATE, checks.CHECKBAL, checks.CHKACCTNAME, checks.BANK, checkalloc.CHKSPRID " & _
                  "From checkalloc INNER JOIN checks ON (checkalloc.CHKSPRID = checks.CHKSPRID) Where checkalloc.CUSTOMERID = " & Val(txtCustomerID.Text) & " AND checks.CHECKBAL > 1 AND (CHECKS.`STATUS` LIKE '%CLEARED%' OR CHECKS.`STATUS` = 'PDC DUE')"
        End If
        With rsChk
            If .State = adStateOpen Then .Close
            .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
            If .RecordCount = 1 Then
                .MoveFirst
                cboCheck.Visible = False
                txtCheckNo.Visible = True
                txtCheckNo.Text = !CHECKNO
                txtChkDate.Text = Format(!CHECKDATE, "mm/dd/yyyy")
                txtChkAmt.Text = Format(IIf(IsNull(!CHECKBAL), 0, !CHECKBAL), "#,##0.00")
                txtChkAcct.Text = !CHKACCTNAME
                txtChkBank.Text = !BANK
                txtChkPRid.Text = !CHKSPRID
                txtAmtRec.Text = Format(IIf(IsNull(!CHECKBAL), 0, !CHECKBAL), "###0.00")
            ElseIf .RecordCount > 1 Then
                .MoveFirst
                Do While Not .EOF
                    cboCheck.AddItem !CHECKNO
                    .MoveNext
                Loop
                txtCheckNo.Visible = False
                cboCheck.Top = txtCheckNo.Top
                cboCheck.Left = txtCheckNo.Left
                cboCheck.Width = txtCheckNo.Width
                cboCheck.Visible = True
                cboCheck.ListIndex = 0
                txtChkPRid.Text = ""
            Else
                txtCheckNo.Text = ""
                txtChkDate.Text = ""
                txtChkAmt.Text = ""
                txtChkAcct.Text = ""
                txtChkBank.Text = ""
                txtChkPRid.Text = ""
                txtCheckNo.Visible = True
                cboCheck.Visible = False
            End If
            .Close
        End With
    End If
    If Val(txtChkPRid.Text) > 0 And PayType <> 3 Then
        Qry = "SELECT checkalloc.SPSORDERID From checkalloc Where checkalloc.CHKSPRID = " & Val(txtChkPRid.Text) & ";"
        If Val(GetFieldValue(Qry)) > 0 Then Call CheckAllocation
    End If
        
        
End Sub

Private Sub CheckAllocation()
    Dim rsAlloc As New ADODB.Recordset
    Dim Qry As String, MyRef As String
    Dim i As Integer, TtlAmtPd As Currency
    
    flxBal.Cols = 7
    flxBal.ColWidth(5) = 0
    flxBal.TextMatrix(0, 1) = "REF. NO"
    flxBal.TextMatrix(0, 1) = "BRANCH"
    flxBal.TextMatrix(0, 3) = "AMT.ALLOC."
    flxBal.Rows = 1
    If ThisBranch = "MRC2" Then
        Qry = "SELECT checkalloc.CUSTOMERID, checkalloc.CHKSPRID, checkalloc.SPSORDERID, checkalloc.ALLOCAMT, checkalloc.BRANCH, checkalloc.RefNo, checkalloc.JOBORDERID  " & _
              "From checkalloc Where checkalloc.CHKSPRID = " & Val(txtChkPRid.Text) & " AND checkalloc.ORNO IS NULL ORDER BY checkalloc.REFNO;"
    Else
        If udORCR.Value = 0 Then
        'lblORno.Caption = "OFFICIAL RECEIPT NO."
            Qry = "SELECT checkalloc.CUSTOMERID, checkalloc.CHKSPRID, checkalloc.SPSORDERID, checkalloc.ALLOCAMT, checkalloc.BRANCH, checkalloc.RefNo, checkalloc.JOBORDERID  " & _
              "From checkalloc Where checkalloc.CHKSPRID = " & Val(txtChkPRid.Text) & " AND checkalloc.ORNO IS NULL AND checkalloc.JOBORDERID > 0 ORDER BY checkalloc.REFNO;"
        
        Else
        'lblORno.Caption = "COLLECTION RECEIPT NO."
            Qry = "SELECT checkalloc.CUSTOMERID, checkalloc.CHKSPRID, checkalloc.SPSORDERID, checkalloc.ALLOCAMT, checkalloc.BRANCH, checkalloc.RefNo, checkalloc.JOBORDERID  " & _
              "From checkalloc Where checkalloc.CHKSPRID = " & Val(txtChkPRid.Text) & " AND checkalloc.ORNO IS NULL AND checkalloc.SPSORDERID > 0 ORDER BY checkalloc.REFNO;"
        
        
        End If
        
        
'        Qry = "SELECT checkalloc.CUSTOMERID, checkalloc.CHKSPRID, checkalloc.SPSORDERID, checkalloc.ALLOCAMT, checkalloc.BRANCH, checkalloc.RefNo, checkalloc.JOBORDERID " & _
              "From checkalloc Where checkalloc.CHKSPRID = " & Val(txtChkPRid.Text) & " AND checkalloc.CUSTOMERID = " & Val(txtCustomerID.Text) & " AND checkalloc.ORNO IS NULL ORDER BY checkalloc.REFNO;"
    End If
    With rsAlloc
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                flxBal.AddItem !RefNo & vbTab & IIf(IsNull(!Branch), "", !Branch) & vbTab & vbTab & Format(!AllocAmt, "###0.00") & vbTab & vbTab & IIf(UCase$(Left$(!RefNo, 3)) <> "JO#", !SPSORDERID, !JOBORDERID) & vbTab & !CUSTOMERID
                .MoveNext
            Loop
        End If
        .Close
    End With
    With flxBal
        TtlAmtPd = 0
        For i = .Rows - 1 To 1 Step -1
            If .TextMatrix(i, 1) = ThisBranch Then
                If UCase$(Left$(.TextMatrix(i, 0), 3)) = "JO#" Then
                    Qry = "SELECT onaccterms.Balance From onaccterms Where onaccterms.SPINVNO IS NULL AND onaccterms.JOBORDERID = " & Val(.TextMatrix(i, 5)) & ";"
                Else
                    Qry = "SELECT onaccterms.Balance From onaccterms Where onaccterms.SPINVNO = '" & .TextMatrix(i, 0) & "' AND onaccterms.SPSORDERID = " & Val(.TextMatrix(i, 5)) & ";"
                End If
                .TextMatrix(i, 2) = Format(Val(GetFieldValue(Qry)), "###0.00")
                .TextMatrix(i, 4) = Val(Replace(.TextMatrix(i, 2), ",", "")) - Val(Replace(.TextMatrix(i, 3), ",", ""))
                If Val(Replace(.TextMatrix(i, 2), ",", "")) = 0 Then
                    If .Rows = 2 And i = 1 Then
                        .Rows = 1
                        .Rows = 2
                    Else
                        .RemoveItem i
                    End If
                End If
            Else
                If ThisBranch = "MRC2" Then 'Branch MRC2 Gante is the only capable of interbranch payment
                    If .TextMatrix(i, 1) = "MRC1" Then
                        If ConnectMRC1 Then
                            Qry = "SELECT onaccterms.Balance From onaccterms Where onaccterms.SPINVNO = '" & .TextMatrix(i, 0) & "' AND onaccterms.SPSORDERID = " & Val(.TextMatrix(i, 5)) & ";"
                            .TextMatrix(i, 2) = Format(Val(GetFieldValueBranch(Qry, "MRC1")), "###0.00")
                            .TextMatrix(i, 4) = Val(Replace(.TextMatrix(i, 2), ",", "")) - Val(Replace(.TextMatrix(i, 3), ",", ""))
                            If Val(Replace(.TextMatrix(i, 2), ",", "")) = 0 Then
                                If .Rows = 2 And i = 1 Then
                                    .Rows = 1
                                    .Rows = 2
                                Else
                                    .RemoveItem i
                                End If
                            End If
                        Else
                            MsgBox "Unable to connect to MRC1 Sobrecary!" & vbCrLf & "Allocation of check payment cannot continue.", vbExclamation + vbOKOnly, "MRC1 Failed"
                            flxBal.Cols = 5
                            flxBal.Rows = 1
                            flxBal.Rows = 2
                            flxBal.TextMatrix(0, 1) = "J.O. NO."
                            Exit Sub
                        End If
                    ElseIf .TextMatrix(i, 1) = "VAN1" Then
                        If ConnectVAN1 Then
                            Qry = "SELECT onaccterms.Balance From onaccterms Where onaccterms.SPINVNO = '" & .TextMatrix(i, 0) & "' AND onaccterms.SPSORDERID = " & Val(.TextMatrix(i, 5)) & ";"
                            .TextMatrix(i, 2) = Format(Val(GetFieldValueBranch(Qry, "VAN1")), "###0.00")
                            .TextMatrix(i, 4) = Val(Replace(.TextMatrix(i, 2), ",", "")) - Val(Replace(.TextMatrix(i, 3), ",", ""))
                            If Val(Replace(.TextMatrix(i, 2), ",", "")) = 0 Then
                                If .Rows = 2 And i = 1 Then
                                    .Rows = 1
                                    .Rows = 2
                                Else
                                    .RemoveItem i
                                End If
                            End If
                        Else
                            MsgBox "Unable to connect to VAN1! Please contact Network Administrator" & vbCrLf & "Allocation of check payment cannot continue.", vbExclamation + vbOKOnly, "VAN 1 Failed"
                            flxBal.Cols = 5
                            flxBal.Rows = 1
                            flxBal.Rows = 2
                            flxBal.TextMatrix(0, 1) = "J.O. NO."
                            Exit Sub
                        End If
        
                    ElseIf .TextMatrix(i, 1) = "VAN2" Then
                        If ConnectVAN2 Then
                            Qry = "SELECT onaccterms.Balance From onaccterms Where onaccterms.SPINVNO = '" & .TextMatrix(i, 0) & "' AND onaccterms.SPSORDERID = " & Val(.TextMatrix(i, 5)) & ";"
                            .TextMatrix(i, 2) = Format(Val(GetFieldValueBranch(Qry, "VAN2")), "###0.00")
                            .TextMatrix(i, 4) = Val(Replace(.TextMatrix(i, 2), ",", "")) - Val(Replace(.TextMatrix(i, 3), ",", ""))
                            If Val(Replace(.TextMatrix(i, 2), ",", "")) = 0 Then
                                If .Rows = 2 And i = 1 Then
                                    .Rows = 1
                                    .Rows = 2
                                Else
                                    .RemoveItem i
                                End If
                            End If
                        Else
                            MsgBox "Unable to connect to VAN2! Please contact Network Admin." & vbCrLf & "Allocation of check payment cannot continue.", vbExclamation + vbOKOnly, "VAN 2 Failed"
                            flxBal.Cols = 5
                            flxBal.Rows = 1
                            flxBal.Rows = 2
                            flxBal.TextMatrix(0, 1) = "J.O. NO."
                            Exit Sub
                        End If
                    ElseIf .TextMatrix(i, 1) = "VAN3" Then
                        If ConnectVan3 Then
                            Qry = "SELECT onaccterms.Balance From onaccterms Where onaccterms.SPINVNO = '" & .TextMatrix(i, 0) & "' AND onaccterms.SPSORDERID = " & Val(.TextMatrix(i, 5)) & ";"
                            .TextMatrix(i, 2) = Format(Val(GetFieldValueBranch(Qry, "VAN3")), "###0.00")
                            .TextMatrix(i, 4) = Val(Replace(.TextMatrix(i, 2), ",", "")) - Val(Replace(.TextMatrix(i, 3), ",", ""))
                            If Val(Replace(.TextMatrix(i, 2), ",", "")) = 0 Then
                                If .Rows = 2 And i = 1 Then
                                    .Rows = 1
                                    .Rows = 2
                                Else
                                    .RemoveItem i
                                End If
                            End If
                        Else
                            MsgBox "Unable to connect to VAN3! Please contact Network Admin." & vbCrLf & "Allocation of check payment cannot continue.", vbExclamation + vbOKOnly, "VAN 2 Failed"
                            flxBal.Cols = 5
                            flxBal.Rows = 1
                            flxBal.Rows = 2
                            flxBal.TextMatrix(0, 1) = "J.O. NO."
                            Exit Sub
                        End If
                    
                    Else
                        flxBal.Cols = 5
                        flxBal.Rows = 1
                        flxBal.Rows = 2
                        flxBal.TextMatrix(0, 1) = "J.O. NO."
                        txtChkPRid.Tag = ""
                        Exit Sub
                    End If
                Else
                    flxBal.Cols = 5
                    flxBal.Rows = 1
                    flxBal.Rows = 2
                    flxBal.TextMatrix(0, 1) = "J.O. NO."
                    txtChkPRid.Tag = ""
                    Exit Sub
                End If
            End If
            If i <= .Rows - 1 Then
                TtlAmtPd = TtlAmtPd + Val(Replace(.TextMatrix(i, 3), ",", ""))
            End If
        Next i
        flxDetails.Rows = 0
        flxDetails.Rows = 1
        flxDetails.AddItem vbTab & "Check Allocation:"
        For i = 1 To .Rows - 1
            If UCase$(Left$(.TextMatrix(i, 0), 3)) = "JO#" Then
                flxDetails.AddItem vbTab & Space$(5) & .TextMatrix(i, 0) & vbTab & .TextMatrix(i, 3)
            Else
                If Left$(.TextMatrix(i, 0), 3) = "TRA" Then
                    MyRef = .TextMatrix(i, 0)
                Else
                    MyRef = "Inv.#:" & .TextMatrix(i, 0)
                End If
                flxDetails.AddItem vbTab & "   " & MyRef & vbTab & .TextMatrix(i, 3)
            End If
        Next i
        flxDetails.AddItem vbTab & "TOTAL" & vbTab & Format(TtlAmtPd, "#,##0.00")
        If flxBal.Rows > 10 Then flxDetails.TextMatrix(1, 1) = "Check Allocation (See Attachment):"
    End With
    txtPayable.Text = TtlAmtPd
    txtTotal.Text = Format(TtlAmtPd, "#,##0.00")
    txtChkPRid.Tag = "CHECK ALLOCATED"
End Sub


Private Sub chkInvoice_Click()
    Dim Qry As String, RetAmt As Currency
    
    If chkInvoice.Tag = "SKIP" Then Exit Sub
    
    If chkInvoice.Value = 1 Then
        If chkJobOrder.Value = 1 Then
            flxDetails.Rows = 1
            Call cboAcctCust_LostFocus 'BOTH JO & INV NAAY CHECK
            Call txtPayable_LostFocus
            Exit Sub
        End If
        '***Access
        'Qry = "SELECT ONACCTERMS.SPINVNO, ONACCTERMS.TERMS, ONACCTERMS.PAYSCHEME, ONACCTERMS.NPAYS, ONACCTERMS.PAYSTART, ONACCTERMS.BILLINGAMT, ONACCTERMS.STATUS, ONACCTERMS.TTLAMTPAID, ONACCTERMS.BALANCE, ONACCTERMS.CUSTOMERID, ONACCTERMS.NPAYS, ONACCTERMS.NPAYS, ONACCTERMS.TRANSDATE, ONACCTERMS.ACCTAMOUNT, ONACCTERMS.JOBORDERID, ONACCTERMS.NPAYS " & _
              "From ONACCTERMS WHERE (((ONACCTERMS.STATUS)='OPEN') AND ((ONACCTERMS.CUSTOMERID)=" & txtCustomerID.Text & ") AND ((ONACCTERMS.JOBORDERID) Is Null Or (ONACCTERMS.JOBORDERID)=0)) ORDER BY ONACCTERMS.JOBORDERID, ONACCTERMS.SPINVNO;"
        'mySQL
        Qry = "SELECT ONACCTERMS.SPINVNO, ONACCTERMS.TERMS, ONACCTERMS.PAYSCHEME, ONACCTERMS.NPAYS, ONACCTERMS.PAYSTART, ONACCTERMS.BILLINGAMT, ONACCTERMS.STATUS, ONACCTERMS.TTLAMTPAID, ONACCTERMS.BALANCE, ONACCTERMS.CUSTOMERID, ONACCTERMS.NPAYS, ONACCTERMS.NPAYS, ONACCTERMS.TRANSDATE, ONACCTERMS.ACCTAMOUNT, ONACCTERMS.JOBORDERID, ONACCTERMS.NPAYS " & _
              "From ONACCTERMS WHERE ONACCTERMS.STATUS='OPEN' AND ONACCTERMS.CUSTOMERID=" & txtCustomerID.Text & " AND (ONACCTERMS.JOBORDERID Is Null Or ONACCTERMS.JOBORDERID=0) ORDER BY ONACCTERMS.JOBORDERID, ONACCTERMS.SPINVNO;"
        
        With rsAcct
            .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
            If .RecordCount > 0 Then
                .MoveFirst
                chkInvoice.Caption = "INVOICE (" & .RecordCount & ")"
                flxAcct.Rows = 1
                Pbar.Value = 0
                Pbar.Max = .RecordCount
                Pbar.Visible = True
                Do While Not .EOF
                    '***Access
                    'Qry = "SELECT Sum(SPITR.RECEIVE*SPITR.PRICE) AS AMTRET From SPITR " & _
                          "WHERE SPITR.TRANSTYPE='SALES RETURN') AND ((SPITR.PACKINGLIST)='" & !SPINVNO & "'));"
                    'mySQL
                    Qry = "SELECT Sum(SPITR.RECEIVE*SPITR.PRICE) AS AMTRET From SPITR " & _
                          "WHERE SPITR.TRANSTYPE='SALES RETURN' AND SPITR.PACKINGLIST='" & !SPINVNO & "';"
                    
                    RetAmt = Val(GetFieldValue(Qry))
                    '0-Invoice No.| 1-Terms | 2-Pay Scheme | 3-Npays | 4-Pay Start | 5-Billing Amount |6- Status |7-Total Amt Pd |8-Balance |9-CustomerID|10-Current AmtPaid|11-New Balance|12-Inv. Date |13-Inv Amount| 14 - JobID | 15-Ret amt
                    flxAcct.AddItem !SPINVNO & vbTab & !Terms & vbTab & !PayScheme & vbTab & !Npays & vbTab & !PayStart & vbTab & !BILLINGAMT & vbTab & !STATUS & vbTab & IIf(IsNull(!TTLAMTPAID), 0, !TTLAMTPAID) & vbTab & !Balance & vbTab & !CUSTOMERID & vbTab & vbTab & vbTab & !TRANSDATE & vbTab & !AcctAmount & vbTab & vbTab & RetAmt
                    .MoveNext
                    Pbar.Value = Pbar.Value + 1
                    DoEvents
                Loop
                Pbar.Visible = False
            End If
            .Close
        End With
        
    Else
        If chkJobOrder.Value = 0 Then
            chkJobOrder.Value = 1
            Exit Sub
        End If
        '***Access
        'Qry = "SELECT ONACCTERMS.JOBORDERNO, ONACCTERMS.TERMS, ONACCTERMS.PAYSCHEME, ONACCTERMS.NPAYS, ONACCTERMS.PAYSTART, ONACCTERMS.BILLINGAMT, ONACCTERMS.STATUS, ONACCTERMS.TTLAMTPAID, ONACCTERMS.BALANCE, ONACCTERMS.CUSTOMERID, ONACCTERMS.NPAYS, ONACCTERMS.NPAYS, ONACCTERMS.TRANSDATE, ONACCTERMS.ACCTAMOUNT, ONACCTERMS.JOBORDERID, ONACCTERMS.NPAYS " & _
              "From ONACCTERMS WHERE (((ONACCTERMS.STATUS)='OPEN') AND ((ONACCTERMS.CUSTOMERID)=" & txtCustomerID.Text & ") AND ((ONACCTERMS.SPINVNO) Is Null)) ORDER BY ONACCTERMS.JOBORDERNO, ONACCTERMS.JOBORDERID, ONACCTERMS.SPINVNO;"
        'mySQL
        Qry = "SELECT ONACCTERMS.JOBORDERNO, ONACCTERMS.TERMS, ONACCTERMS.PAYSCHEME, ONACCTERMS.NPAYS, ONACCTERMS.PAYSTART, ONACCTERMS.BILLINGAMT, ONACCTERMS.STATUS, ONACCTERMS.TTLAMTPAID, ONACCTERMS.BALANCE, ONACCTERMS.CUSTOMERID, ONACCTERMS.NPAYS, ONACCTERMS.NPAYS, ONACCTERMS.TRANSDATE, ONACCTERMS.ACCTAMOUNT, ONACCTERMS.JOBORDERID, ONACCTERMS.NPAYS " & _
              "From ONACCTERMS WHERE ONACCTERMS.STATUS='OPEN' AND ONACCTERMS.CUSTOMERID=" & txtCustomerID.Text & " AND (ONACCTERMS.SPINVNO) Is Null ORDER BY ONACCTERMS.JOBORDERNO, ONACCTERMS.JOBORDERID, ONACCTERMS.SPINVNO;"
        With rsAcct
            .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
            If .RecordCount > 0 Then
                .MoveFirst
                chkJobOrder.Caption = "JOB ORDER (" & .RecordCount & ")"
                flxAcct.Rows = 1
                Pbar.Visible = True
                Pbar.Max = .RecordCount + 1
                Pbar.Value = 0
                Do While Not .EOF
                    '0-Job No.| 1-Terms | 2-Pay Scheme | 3-Npays | 4-Pay Start | 5-Billing Amount |6- Status |7-Total Amt Pd |8-Balance |9-CustomerID|10-Current AmtPaid|11-New Balance|12-Inv. Date |13-Inv Amount|14-Job Order ID| 15-Ret amt
                    flxAcct.AddItem !JOBORDERNO & vbTab & !Terms & vbTab & !PayScheme & vbTab & !Npays & vbTab & !PayStart & vbTab & !BILLINGAMT & vbTab & !STATUS & vbTab & IIf(IsNull(!TTLAMTPAID), 0, !TTLAMTPAID) & vbTab & !Balance & vbTab & !CUSTOMERID & vbTab & vbTab & vbTab & !TRANSDATE & vbTab & !AcctAmount & vbTab & IIf(IsNull(!JOBORDERID), "", !JOBORDERID)
                    .MoveNext
                    Pbar.Value = Pbar.Value + 1
                    DoEvents
                Loop
                Pbar.Visible = False
            End If
            .Close
        End With
        
    End If
    flxDetails.Rows = 1
    Call DisplayAcctDetails
    Call txtPayable_LostFocus
End Sub

Private Sub chkJobOrder_Click()
    Dim rsAcct As New ADODB.Recordset
    Dim Qry As String, RetAmt As Currency
    
    If chkJobOrder.Tag = "SKIP" Then Exit Sub
    
    If chkJobOrder.Value = 1 Then
        If chkInvoice.Value = 1 Then
            flxDetails.Rows = 1
            Call cboAcctCust_LostFocus
            Call txtPayable_LostFocus
            Exit Sub
        End If
        '***Access
        'Qry = "SELECT ONACCTERMS.JOBORDERNO, ONACCTERMS.TERMS, ONACCTERMS.PAYSCHEME, ONACCTERMS.NPAYS, ONACCTERMS.PAYSTART, ONACCTERMS.BILLINGAMT, ONACCTERMS.STATUS, ONACCTERMS.TTLAMTPAID, ONACCTERMS.BALANCE, ONACCTERMS.CUSTOMERID, ONACCTERMS.TRANSDATE, ONACCTERMS.ACCTAMOUNT, ONACCTERMS.JOBORDERID " & _
              "From ONACCTERMS WHERE (((ONACCTERMS.STATUS)='OPEN') AND ((ONACCTERMS.CUSTOMERID)=" & txtCustomerID.Text & ") AND ((ONACCTERMS.JOBORDERID)>0));"
        'mySQL
        Qry = "SELECT ONACCTERMS.JOBORDERNO, ONACCTERMS.TERMS, ONACCTERMS.PAYSCHEME, ONACCTERMS.NPAYS, ONACCTERMS.PAYSTART, ONACCTERMS.BILLINGAMT, ONACCTERMS.STATUS, ONACCTERMS.TTLAMTPAID, ONACCTERMS.BALANCE, ONACCTERMS.CUSTOMERID, ONACCTERMS.TRANSDATE, ONACCTERMS.ACCTAMOUNT, ONACCTERMS.JOBORDERID " & _
              "From ONACCTERMS WHERE ONACCTERMS.STATUS='OPEN' AND ONACCTERMS.CUSTOMERID=" & txtCustomerID.Text & " AND ONACCTERMS.JOBORDERID>0;"
        
        With rsAcct
            .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
            If .RecordCount > 0 Then
                .MoveFirst
                chkJobOrder.Caption = "JOB ORDER (" & .RecordCount & ")"
                flxAcct.Rows = 1
                Do While Not .EOF
                    '0-Job No.| 1-Terms | 2-Pay Scheme | 3-Npays | 4-Pay Start | 5-Billing Amount |6- Status |7-Total Amt Pd |8-Balance |9-CustomerID|10-Current AmtPaid|11-New Balance|12-Inv. Date |13-Inv Amount|14-Job Order ID| 15-Ret amt
                    flxAcct.AddItem !JOBORDERNO & vbTab & !Terms & vbTab & !PayScheme & vbTab & !Npays & vbTab & !PayStart & vbTab & !BILLINGAMT & vbTab & !STATUS & vbTab & IIf(IsNull(!TTLAMTPAID), 0, !TTLAMTPAID) & vbTab & !Balance & vbTab & !CUSTOMERID & vbTab & vbTab & vbTab & !TRANSDATE & vbTab & !AcctAmount & vbTab & IIf(IsNull(!JOBORDERID), "", !JOBORDERID)
                    .MoveNext
                Loop
            End If
            .Close
        End With
    Else
        If chkInvoice.Value = 0 Then
            chkInvoice.Value = 1
            Exit Sub
        End If
        '***Access
        'Qry = "SELECT ONACCTERMS.*, SPINVOICE.INVDATE, SPINVOICE.AMOUNT FROM ONACCTERMS INNER JOIN SPINVOICE ON ONACCTERMS.SPINVNO = SPINVOICE.SPINVNO " & _
              "WHERE (((ONACCTERMS.CUSTOMERID)=" & txtCustomerID.Text & ") AND ((ONACCTERMS.STATUS)='OPEN'));"
        'mySQl
        Qry = "SELECT ONACCTERMS.*, SPINVOICE.INVDATE, SPINVOICE.AMOUNT FROM ONACCTERMS INNER JOIN SPINVOICE ON ONACCTERMS.SPINVNO = SPINVOICE.SPINVNO " & _
              "WHERE ONACCTERMS.CUSTOMERID=" & txtCustomerID.Text & " AND ONACCTERMS.STATUS='OPEN';"
        
        With rsAcct
            .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
            If .RecordCount > 0 Then
                .MoveFirst
                chkInvoice.Caption = "INVOICE (" & .RecordCount & ")"
                flxAcct.Rows = 1
                Do While Not .EOF
                    '***Access
                    'Qry = "SELECT Sum(SPITR.RECEIVE*SPITR.PRICE) AS AMTRET From SPITR " & _
                          "WHERE SPITR.TRANSTYPE='SALES RETURN') AND ((SPITR.PACKINGLIST)='" & !SPINVNO & "'));"
                    'mySQL
                    Qry = "SELECT Sum(SPITR.RECEIVE*SPITR.PRICE) AS AMTRET From SPITR " & _
                          "WHERE SPITR.TRANSTYPE='SALES RETURN' AND SPITR.PACKINGLIST='" & !SPINVNO & "';"
                    
                    RetAmt = Val(GetFieldValue(Qry))
                    '0-Invoice No.| 1-Terms | 2-Pay Scheme | 3-Npays | 4-Pay Start | 5-Billing Amount |6- Status |7-Total Amt Pd |8-Balance |9-CustomerID|10-Current AmtPaid|11-New Balance|12-Inv. Date |13-Inv Amount| 14 - JobID | 15-Ret amt
                    flxAcct.AddItem !SPINVNO & vbTab & !Terms & vbTab & !PayScheme & vbTab & !Npays & vbTab & !PayStart & vbTab & !BILLINGAMT & vbTab & !STATUS & vbTab & IIf(IsNull(!TTLAMTPAID), 0, !TTLAMTPAID) & vbTab & !Balance & vbTab & !CUSTOMERID & vbTab & vbTab & vbTab & !InvDate & vbTab & !Amount & vbTab & vbTab & RetAmt
                    .MoveNext
                Loop
            End If
            .Close
        End With
    End If
    flxDetails.Rows = 1
    Call DisplayAcctDetails
    Call txtPayable_LostFocus
End Sub

Private Sub chkQckSrvc_Click()
    txtRefNo.Text = ""
    If chkQckSrvc.Value = 1 Then
        lblRefNo.Visible = True
        txtRefNo.Visible = True
        updRef.Visible = True
        updRef.Max = 1
        updRef.Value = 0
        txtRefNo.SetFocus
        cmdDetails.Caption = "QUICK SERVICE On-Account"
        
        fraAllocation.Visible = False
        fraAllocationOthrData.Visible = False
    Else
        lblRefNo.Visible = False
        txtRefNo.Visible = False
        updRef.Visible = False
        cmdDetails.Caption = "ON-ACCOUNT"
        
        fraAllocation.Visible = True
        fraAllocationOthrData.Visible = True
    End If
End Sub

Private Sub cmdPrint_Click()
    Dim mcM2Rep As CRAXDRT.Application
    Dim mcReport As CRAXDRT.Report

    Dim Qry As String, z As Integer
    Dim i As Integer, j As Integer, NxtRow As Integer, InsertHere As Integer
    Dim MyDesc() As String, myTop() As String, myLeft() As String, myLen() As String
    Dim myHeight() As String, myFsize() As String, myFbold() As String, myAlign() As String
    Dim mySubLeft() As String, mySubLen() As String, mySubBold() As String, mySubAlign() As String
    Dim txtObj As TextObject
    Dim BoxObj As BoxObject
    Dim LinObj As LineObject
    Set mcM2Rep = New CRAXDRT.Application
    Set mcReport = mcM2Rep.NewReport
    
    mcReport.PaperSize = crDefaultPaperSize
    mcReport.TopMargin = 0
    mcReport.BottomMargin = 0
    mcReport.Sections(1).Height = 0
    mcReport.Sections(2).Height = 0
    With mcReport.Sections(3)
        For i = 500 To 12000 Step 500
            Set txtObj = .AddTextObject("|", i, 100)
            txtObj.Width = 300: txtObj.HorAlignment = crLeftAlign: txtObj.Font.Name = "Arial": txtObj.Font.Size = 10: txtObj.Height = 250
        Next i
        For i = 100 To 8000 Step 100
            Set txtObj = .AddTextObject(IIf(i Mod 500 = 0, "---", "-"), 100, i)
            txtObj.Width = 300: txtObj.HorAlignment = crLeftAlign: txtObj.Font.Name = "Arial": txtObj.Font.Size = 10: txtObj.Height = 250
        Next i
    End With
    mcReport.Sections(4).Height = 0
    mcReport.Sections(5).Height = 0
    
    mcReport.PrintOut False, 1, , 1, 1
    
End Sub

Function GetReceiptNo() As String
    Dim rsInvNo As New ADODB.Recordset
    Dim Qry As String
    Dim MyNo As Long
    
    '***Access
    'Qry = "SELECT PAYMENTS.ORNO From PAYMENTS Where ((Not (PAYMENTS.ORNO) Is Null)) ORDER BY PAYMENTS.ORNO DESC;"
    'mySQL
    On Error GoTo LocalErr
    If udORCR.Value = 0 Then
        Qry = "SELECT payments.ORNO From payments Where payments.PAYDATE = '" & Format(Date, "yyyy-mm-dd") & "' AND payments.ORNO IS NOT NULL Order By payments.ORNO Desc;"
    Else
        Qry = "SELECT payments.CRNO From payments Where payments.PAYDATE = '" & Format(Date, "yyyy-mm-dd") & "' AND payments.ORNO IS NOT NULL Order By payments.CRNO Desc;"
    End If
    With rsInvNo
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            MyNo = IIf(IsNull(.Fields(0)), 1, Val(.Fields(0)) + 1)
        Else
            MyNo = 0
        End If
        .Close
    End With
    GetReceiptNo = IIf(Len(str(MyNo)) >= 5, MyNo, Format(MyNo, "00000"))
    Exit Function
LocalErr:
    MsgBox Err.Description & vbCrLf & "Please contact System Administrator", vbCritical + vbOKOnly, "Error"
    
End Function

Private Sub UpdateCheck()
    Dim MyBal As Currency, Qry As String
    
    'Update Check Allocation
    'mySQL
    LYRcon.Execute "UPDATE CHECKALLOC SET CHECKALLOC.ORAMT = " & Val(Replace(txtPayable.Text, ",", "")) & ", CHECKALLOC.ORNO = '" & txtORno.Text & "', CHECKALLOC.ORDATE = '" & Format(txtDate.Text, "yyyy-mm-dd") & "' WHERE (CHECKALLOC.ORDATE) Is Null AND CHECKALLOC.CHKSPRID=" & Val(txtChkPRid.Text) & " AND CHECKALLOC.CUSTOMERID=" & Val(txtCustomerID.Text) & ";"
    
    'Update Check Balance
    MyBal = Val(Replace(txtChkAmt.Text, ",", "")) - (txtPayable.Text)
    
    'mySQL
    LYRcon.Execute "UPDATE CHECKS SET CHECKS.CHECKBAL = CHECKS.CHECKBAL-" & Val(Replace(txtPayable.Text, ",", "")) & " WHERE CHECKS.CHKSPRID=" & Val(txtChkPRid.Text) & ";"
    
    If MyBal > 0 Then
        'mySQL
        Qry = "SELECT CHECKALLOC.CHKSPRID, CHECKALLOC.ORDATE From CHECKALLOC " & _
              "WHERE CHECKALLOC.CHKSPRID=" & Val(txtChkPRid.Text) & " AND (CHECKALLOC.ORDATE) Is Null;"
        If Not CtrlNumExist(Qry) Then
            LYRcon.Execute "INSERT INTO CHECKALLOC ( CHKSPRID, CUSTOMERID ) VALUES (" & Val(txtChkPRid.Text) & ", " & Val(txtCustomerID.Text) & ");"
        End If
    End If
End Sub

Private Sub cmdTrans_Click(Index As Integer)
On Error GoTo er
    Dim rsCash As New ADODB.Recordset
    Dim Ans As Integer, i As Integer
    Dim Qry As String, PrnVal(0 To 10) As String, myRem As String, DPdtls() As String
    Dim OSbal As Single, MyAcc As String, PrnPrf As String, myCNattach As Integer
    Dim tmpRs As New ADODB.Recordset
    
    If Index <> 4 Then
        '***Access
        'Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE (((USERACCESS.EMPID)=" & UserID & ") AND ((USERACCESS.MENUINDX)=45));"
        'mySQL
        Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE USERACCESS.EMPID=" & UserID & " AND USERACCESS.MENUINDX=45;"
        MyAcc = GetFieldValue(Qry)
        If MyAcc <> "FULL ACCESS" Then
            MsgBox "You don't have full access to this transaction." & vbCrLf & "Please contact your System Administrator", vbExclamation + vbOKOnly, "Access Denied"
            Exit Sub
        End If
    End If
    
    Select Case Index
        Case 0  'Quick Service
            fraAllocation.Visible = False
            fraAllocationOthrData.Visible = False
            DidRecordEntry = True
            
            Call ClearEntries
            lblRefNo.Caption = "JOB ORDER NO."
            txtORno.Text = GetReceiptNo
            txtDate.Text = Format$(Date, "mm/dd/yyyy")
            Call LockData(False)
            txtRefNo.SetFocus
            updRef.Visible = False
            chkCash.Value = 1
            cmdDetails.Caption = "QUICK SERVICE-CASH"
            Call cmdSaveCancel(True)
            txtRecFrom.Visible = True
            cboAcctCust.Visible = False
            PrtlPay = False
            tabOR.Visible = False
            flxBal.Visible = False
            PayType = 0         'NEW entry use for MRC Panabo
            udORCR.Value = 0
            
        Case 1  'Down Payment
            fraAllocation.Visible = False
            fraAllocationOthrData.Visible = False
            DidRecordEntry = True
            
            Call ClearEntries
            txtORno.Text = GetReceiptNo
            txtDate.Text = Format$(Date, "mm/dd/yyyy")
            cmdDetails.Caption = "DOWN PAYMENT"
            Call LockData(False)
            Call cmdSaveCancel(True)
            lblRefNo.Caption = "SPARE PARTS"    'Default
            txtRefNo.Visible = False    'set to true if updref = JOB ORDER
            updRef.Visible = True   'toggle btween SP and JO, hide txtrefno if SP
            updRef.Max = 3
            txtRecFrom.Visible = False
            cboAcctCust.Visible = True
            txtPayable.Locked = False
            tabOR.Visible = False
            flxBal.Visible = False
            With flxDetails
                .ColWidth(0) = 0
                .ColWidth(1) = .Width * 0.98
                .ColWidth(2) = 0
                .Rows = 2
                .TextMatrix(0, 1) = "Down Payment for:"
                .Col = 1
                .Row = 1
            End With
            udORCR.Value = 1
            
            Call flxDetails_Click
            
        Case 2  'Charge Invoice
            
            fraAllocation.Visible = True
            fraAllocationOthrData.Visible = True
            DidRecordEntry = True
            
            Call ClearEntries
            
            chkCheck.Enabled = False
            udORCR.Enabled = False
            
            
            txtPayable.Locked = True
            txtAmtRec.Locked = True
            txtAmtChange.Locked = True
            txtPayable.BackColor = txtTmpGrayColor.BackColor
            txtAmtRec.BackColor = txtTmpGrayColor.BackColor
            txtAmtChange.BackColor = txtTmpGrayColor.BackColor
            
            
            txtCheckNo.Locked = True
            txtChkDate.Locked = True
            txtChkAmt.Locked = True
            txtChkAcct.Locked = True
            txtChkBank.Locked = True
            txtCheckNo.BackColor = txtTmpGrayColor.BackColor
            txtChkDate.BackColor = txtTmpGrayColor.BackColor
            txtChkAmt.BackColor = txtTmpGrayColor.BackColor
            txtChkAcct.BackColor = txtTmpGrayColor.BackColor
            txtChkBank.BackColor = txtTmpGrayColor.BackColor
 
 
            txtDate.Text = Format$(Date, "mm/dd/yyyy")
            txtORno.Text = GetReceiptNo
            cmdDetails.Caption = "ON-ACCOUNT"
            Call LockData(False)
            Call cmdSaveCancel(True)
            lblRefNo.Caption = "INVOICE NO."
            lblRefNo.Visible = False
            txtRefNo.Visible = False
            chkQckSrvc.Value = 0
            chkQckSrvc.Visible = True
            updRef.Visible = False
            txtRecFrom.Visible = False
            cboAcctCust.Visible = True
            
            chkInvoice.Visible = False
            chkJobOrder.Visible = False
            udORCR.Value = 1
            
            chkQckSrvc.Visible = False
            
        Case 3  'Others
            fraAllocation.Visible = False
            fraAllocationOthrData.Visible = False
            DidRecordEntry = True
            Call ClearEntries
            txtORno.Text = GetReceiptNo
            txtDate.Text = Format$(Date, "mm/dd/yyyy")
            cmdDetails.Caption = "OTHERS"
            Call LockData(False)
            Call cmdSaveCancel(True)
            lblRefNo.Visible = False
            txtRefNo.Visible = False    'set to true if updref = JOB ORDER
            updRef.Visible = False   'toggle btween SP and JO, hide txtrefno if SP
            txtRecFrom.Visible = False
            cboAcctCust.Visible = True
            txtPayable.Locked = False
            OnAcctPCV.Value = 0
            PayType = 3
            isOthers = 1
            With flxDetails
                .ColWidth(0) = 0
                .ColWidth(1) = .Width * 0.98
                .ColWidth(2) = 0
                .Rows = 2
                .TextMatrix(0, 1) = "Payment for:"
                .Col = 1
                .Row = 1
            End With
            flxBal.Visible = False
            Call flxDetails_Click
            udORCR.Value = 0
            
        Case 4    'EXIT
            Call Unload(Receipt)
            
            Set Receipt = Nothing
            'MotoRep.imgBG.Visible = True
            MotoRep.toolBarButton.Visible = True
            Exit Sub
        Case 5      'Save
            Dim myMsg As String
            myMsg = ValidateReceipt
            If myMsg <> "" Then
                MsgBox myMsg, vbCritical + vbOKOnly, "Validation Error"
                'txtRefNo.SetFocus
                Exit Sub
            End If
            Ans = MsgBox("Do you want to save payment?", vbQuestion + vbYesNo)
            If Ans = vbNo Then
                'txtRefNo.SetFocus
                Exit Sub
                'MotoRep.tbMain.Visible = True
                'MotoRep.imgBG.Visible = True
                MotoRep.toolBarButton.Visible = True
            End If
            
            Select Case PayType
                Case 0  'Quick Service
                    Call SaveJOpay
                    If Val(txtChkPRid.Text) > 0 Then Call UpdateCheck
                    txtPayable.Locked = True
                Case 1  'DP
                    Call SaveDP_Others
                Case 2  'Charge Invoice
                    
                    If chkQckSrvc.Value = 0 Then ' Credit Transaction
                        
                        'Comment By SMC
                        '2016-08-08
                        'Call SaveOnAccount
                        'If Val(txtChkPRid.Text) > 0 And txtChkPRid.Tag = "" Then Call UpdateCheck
                        'txtPayable.Locked = True
                        'If txtPaymentID.Tag = "ATTACHMENT" Then
                        '    flxDetails.Rows = 0
                        '    flxDetails.Rows = 1
                        '    flxDetails.AddItem vbTab & "See Attachment"
                        '
                        'ElseIf txtChkPRid.Tag = "" Then
                        '    flxDetails.Rows = 0
                        '    flxDetails.Rows = 1
                        '    OSbal = 0
                        '    For i = 1 To flxAcct.Rows - 1
                        '        If Val(flxAcct.TextMatrix(i, 10)) > 0 Then
                        '            flxDetails.AddItem vbTab & IIf(Val(flxAcct.TextMatrix(i, 14)) > 0, "J.O.# ", "Inv.# ") & flxAcct.TextMatrix(i, 0)
                        '            flxDetails.AddItem vbTab & Space$(5) & "Total Amt.Pd." & vbTab & Format(Val(flxAcct.TextMatrix(i, 7)) + Val(flxAcct.TextMatrix(i, 10)), "#,##0.00")
                        '            If Val(flxAcct.TextMatrix(i, 11)) = 0 Then
                        '                flxDetails.AddItem vbTab & Space$(5) & "FULLY PAID"
                        '            Else
                        '                flxDetails.AddItem vbTab & Space$(5) & "Balance" & vbTab & Format$(flxAcct.TextMatrix(i, 11), "#,##0.00")
                        '            End If
                        '        End If
                        '        OSbal = OSbal + Val(flxAcct.TextMatrix(i, 11))
                        '    Next i
                        '    If OSbal > 0 Then
                        '        flxDetails.AddItem ""
                        '        flxDetails.AddItem vbTab & "Outstanding Bal." & vbTab & Format(OSbal, "#,##0.00")
                        '    End If
                        'End If
                        Dim pREFDATE As String
                        Dim pORNO As String
                        Dim pCRNO As String
                        pORNO = ""
                        pCRNO = ""
                        
                        If Trim(cboAllocationNo.Text) = "" Then
                            MsgBox "ALLOCATION NO. IS REQUIRED!", vbExclamation, "Validation!"
                            cboAllocationNo.SetFocus
                            Exit Sub
                        End If
                        
                        If Trim(txtTTLAllocation.Text) = "" Then
                            MsgBox "TOTAL AMOUNT IS REQUIRED!", vbExclamation, "Validation!"
                            Exit Sub
                        End If
                        
                        If Trim(cboAcctCust.Text) = "" Then
                            MsgBox "RECEIVED FROM IS REQUIRED!", vbExclamation, "Validation!"
                            cboAcctCust.SetFocus
                            Exit Sub
                        End If
                        
                        If flxBal.Rows <= 1 Then
                            MsgBox "ACCOUNT LIST IS EMPTY!", vbCritical, "Validation!"
                            Exit Sub
                        End If
                        
                        If cboPayType.Text = "CREDIT CARD PAYMENT" Then
                            If Trim(txtCardHolder.Text) = "" Then
                                MsgBox "CARD HOLDER NAME IS REQUIRED!", vbExclamation, "Validation!"
                                txtCardHolder.SetFocus
                                Exit Sub
                            End If
                            If Trim(txtCardNo.Text) = "" Then
                                MsgBox "CARD NUMBER IS REQUIRED!", vbExclamation, "Validation!"
                                txtCardNo.SetFocus
                                Exit Sub
                            End If
                            If Trim(txtCreditBank.Text) = "" Then
                                MsgBox "NAME OF BANK IS REQUIRED!", vbExclamation, "Validation!"
                                txtCreditBank.SetFocus
                                Exit Sub
                            End If
                            
                            LYRcon.Execute "Drop Temporary Table IF  Exists sp_mrcretails.TMP_CREDITCARD_DTLS;"
                            
                            DoEvents
                            LYRcon.Execute " Create Temporary Table IF Not Exists sp_mrcretails.TMP_CREDITCARD_DTLS( " & _
                            "   CCNAME Varchar(100),         " & _
                            "   CCNO Varchar(100),                " & _
                            "   CCBANKNAME Varchar(100)      " & _
                            ");"
                    
                            Qry = "insert into sp_mrcretails.TMP_CREDITCARD_DTLS(CCNAME,CCNO,CCBANKNAME) VALUES ('" & txtCardHolder.Text & "','" & txtCardNo.Text & "', '" & txtCreditBank.Text & "')"
                            
                            Call LYRcon.Execute(Qry)
                        End If
                        pREFDATE = txtDate.Text
                        
                        TRANSTYPE = Trim(TRANSTYPE)
                        If TRANSTYPE = "SPARE PARTS" Then
                            pCRNO = txtORno.Text
                            TRANSTYPE = "SPAREPARTS"
                        Else
                            pORNO = txtORno.Text
                        End If
                        

                        ' For PayType ACKNOWLEDGEMENT
                        ' Payment Document is not required
                        If StrPAYTYPE = "AKLDGMT" Then
                            pREFDATE = ""
                            pORNO = ""
                            pCRNO = ""
                        End If
                        
                        If Trim(pREFDATE) <> "" Then
                            pREFDATE = Format(CDate(pREFDATE), "YYYY-m-dd")
                        End If
                        
                        DoEvents
                        Screen.MousePointer = vbHourglass
                        Me.Enabled = False
                        Call modAllocation.sp_payments_create(AllocationNo, txtCustomerID.Text, pORNO, pCRNO, pREFDATE, TRANSTYPE, myAcctCode, tmpRs)
                        Screen.MousePointer = vbDefault
                        Me.Enabled = True
                        DoEvents
                         If tmpRs.State = 1 Then
                            myCNattach = 0
                            If tmpRs.RecordCount > 0 Then
                                If tmpRs!level = "Error" Or tmpRs!level = "Warning" Then
                                    MsgBox tmpRs!message, vbCritical, "Error-> ( FLAG:" & tmpRs!Flag & ") "
                                    txtORno.SetFocus
                                    If tmpRs.State <> 0 Then tmpRs.Close
                                    Exit Sub
                                End If
                                txtPaymentID.Text = tmpRs!PaymentID
                                If tmpRs!CNTAttachment > 1 Then
                                    txtPaymentID.Tag = "ATTACHMENT"
                                    myCNattach = tmpRs!CNTAttachment
                                Else
                                    txtPaymentID.Tag = ""
                                End If
                            End If
                        Else
                            MsgBox "PAYMENT CREATION WAS UNSUCCESSFUL EXPECTED PARAMETER WAS NOT ARRIVED." & vbCrLf & "PLEASE INFORM THE MIS DEPARTMENT.", vbCritical, "Error->cmdTrans_Click()"
                            Exit Sub
                        End If
                        If tmpRs.State <> 0 Then tmpRs.Close

                        
                        If txtPaymentID.Tag = "ATTACHMENT" Then
                            flxDetails.Rows = 0
                            flxDetails.Rows = 1
                            flxDetails.AddItem vbTab & "See Attachment"
                        End If
                    Else
                        Call SaveOnAcctQS
                        If Val(txtChkPRid.Text) > 0 Then Call UpdateCheck
                    End If
                    DidRecordEntry = False
                    
                Case 3  'Others
                    Call SaveDP_Others
                    If Val(txtChkPRid.Text) > 0 Then Call UpdateCheck
                    If txtPaymentID.Tag = "ATTACHMENT" Then
                        flxDetails.Rows = 0
                        flxDetails.Rows = 1
                        flxDetails.AddItem vbTab & "See Attachment"
                    End If
                    DidRecordEntry = False
                    
                Case 7  'Cash Invoice
                    DidRecordEntry = False
                    LYRcon.Execute "UPDATE SPINVOICE SET SPINVOICE.ORNO = '" & txtORno.Text & "', REMARKS = 'Receipt# " & txtORno.Text & "' " & _
                        "WHERE SPINVOICE.SPINVNO = '" & Trim(flxDetails.TextMatrix(1, 1)) & "';"
                    LYRcon.Execute "UPDATE payments SET REMARKS = 'Receipt# " & txtORno.Text & "' Where payments.INVOICENO = '" & Trim(flxDetails.TextMatrix(1, 1)) & "';"
            End Select
            
            Call LockData(True)
            Call cmdSaveCancel(False)
            'If InStr(SysBranch, "PANABO") > 0 And PayType = 7 Then " & myCNattach & "
                
            If SysBranch <> "LYR MOTOR SHOP AND SPAREPARTS DAVAO" Then
                Ans = MsgBox("Payment successfully Saved!" & vbCrLf & vbCrLf & "Do you want to Print Official Receipt?", vbQuestion + vbYesNo, "Print O.R")
                If Ans = vbNo Then Exit Sub
                 PrnVal(0) = txtDate.Text
                 PrnVal(1) = IIf(PayType < 2 Or PayType = 7, txtRecFrom.Text, cboAcctCust.Text)
                 PrnVal(2) = txtAddress.Text
                 PrnVal(3) = txtTIN.Text
                 PrnVal(4) = txtSumOf.Text
                 PrnVal(5) = ""  'Text Amount 2
                 PrnVal(6) = Format$(txtPayable.Text, "P#,##0.00")
                 PrnVal(7) = txtCheckNo.Text
                 PrnVal(8) = txtChkBank.Text
                 If PayType = 1 Or (PayType = 3 And OnAcctPCV.Value = 0) Then 'Down Payment and Others with Details
                     PrnVal(9) = PrnRemarks(PayType)
                     myRem = ""
                     For i = 0 To flxDetails.Rows - 1
                         If Trim(flxDetails.TextMatrix(i, 1)) <> "" Then
                             myRem = myRem & IIf(myRem <> "", "|", "") & Trim(flxDetails.TextMatrix(i, 1))
                         End If
                     Next i
                
                     flxDetails.Rows = 0
                 'ElseIf PayType = 7 And InStr(SysBranch, "PANABO") > 0 Then
                 '   PrnVal(2) = "SPAREPARTS " & IIf(Trim(flxDetails.TextMatrix(1, 1)) <> "", "REF. NO. " & Trim(flxDetails.TextMatrix(1, 1)), "")
                 '   PrnVal(9) = ""  'cmdDetails.Caption
                 'ElseIf PayType = 0 And InStr(SysBranch, "PANABO") > 0 Then
                 '   PrnVal(2) = "SERVICE " & IIf(Trim(flxDetails.TextMatrix(1, 1)) <> "", Trim(flxDetails.TextMatrix(1, 1)), "")
                 '   PrnVal(9) = ""  'cmdDetails.Caption
                 Else
                    PrnVal(9) = ""
                 End If
                 If udORCR.Value = 0 Then 'O.R.
                    If isOthers = 1 Then ' added on 01/12/2026
                          Qry = "SELECT settings.Value From settings Where settings.DESCRIPTION = 'PRINT RECEIPT' AND settings.REMARKS = 'O.R.(OTHERS)'"
                    Else
                     Qry = "SELECT settings.Value From settings Where settings.DESCRIPTION = 'PRINT RECEIPT' AND settings.REMARKS = 'O.R.'"
                     End If
                 Else
                     Qry = "SELECT settings.Value From settings Where settings.DESCRIPTION = 'PRINT RECEIPT' AND settings.REMARKS = 'C.R.'"
                 End If
                 PrnPrf = GetFieldValue(Qry) 'PURPOSE DATI ==> Call PrintMyFormNew(PrnVal, PrnPrf, flxJOdtls)
                  PrnVal(10) = txtAmtRec.Text
                  If PrnPrf = "RECEIPT-OR" Then
                    Call PrintMyFormNew(PrnVal, PrnPrf, flxJOdtls)
                 Else
                    Call PrintMyForm(PrnVal, PrnPrf, flxBal)
                 End If
            Else
                 MsgBox ("Payment successfully Saved!"), vbInformation, "Save"
            End If
            
            If (PayType = 1 Or PayType = 3) And myRem <> "" Then
                DPdtls = Split(myRem, "|")
                For i = LBound(DPdtls) To UBound(DPdtls)
                    flxDetails.AddItem vbTab & DPdtls(i)
                Next i
            End If
            
            If PayType = 2 And txtPaymentID.Tag = "ATTACHMENT" Then
                Dim crAttch As New rptORattachment
                
                If udORCR.Value = 0 Then 'O.R.
                    Qry = " SELECT PAYMENTS.CUSTOMERNAME, PAYMENTS.SPAREPARTS+PAYMENTS.SERVICE+PAYMENTS.OTHERS AS ORAMOUNT, PAYMENTS.ORNO " & _
                          " ,PAYMENTS.PAYDATE, ORDETAILS.INVOICENO, ORDETAILS.JOBORDERNO, ORDETAILS.TRANSDATE, ORDETAILS.ACCTAMT, ORDETAILS.AMTPAID " & _
                          " ,ORDETAILS.TTLPAYMENT, ORDETAILS.BALANCE, round(ifnull(allocation.TAXWHELDAMOUNT,0),2) as  TAXWHELDAMOUNT " & _
                          " FROM PAYMENTS INNER JOIN ORDETAILS ON PAYMENTS.PAYMENTID = ORDETAILS.PAYMENTID  " & _
                          " Left Join  allocation  on payments.ALLOCATIONID =  allocation.ALLOCATIONID " & _
                          " Where PAYMENTS.PAYMENTID = " & txtPaymentID.Text & " ORDER BY ORDETAILS.JOBORDERNO, ORDETAILS.INVOICENO;"
                Else
                    Qry = " SELECT PAYMENTS.CUSTOMERNAME, PAYMENTS.SPAREPARTS+PAYMENTS.SERVICE+PAYMENTS.OTHERS AS ORAMOUNT, PAYMENTS.CRNO " & _
                          " ,PAYMENTS.PAYDATE, ORDETAILS.INVOICENO, ORDETAILS.JOBORDERNO, ORDETAILS.TRANSDATE, ORDETAILS.ACCTAMT, ORDETAILS.AMTPAID " & _
                          " ,ORDETAILS.TTLPAYMENT, ORDETAILS.BALANCE , round(ifnull(allocation.TAXWHELDAMOUNT,0),2) as  TAXWHELDAMOUNT " & _
                          " FROM PAYMENTS INNER JOIN ORDETAILS ON PAYMENTS.PAYMENTID = ORDETAILS.PAYMENTID " & _
                          " Left Join  allocation  on payments.ALLOCATIONID =  allocation.ALLOCATIONID " & _
                          " Where PAYMENTS.PAYMENTID = " & txtPaymentID.Text & " ORDER BY ORDETAILS.JOBORDERNO, ORDETAILS.INVOICENO;"
                End If
                
                If rsPrn.State = adStateOpen Then rsPrn.Close
                rsPrn.Open Qry, LYRcon, adOpenDynamic, adLockOptimistic
                crAttch.DiscardSavedData
                crAttch.Database.SetDataSource rsPrn
                crAttch.ParameterFields(1).ClearCurrentValueAndRange
                crAttch.ParameterFields(1).AddCurrentValue SysBranch
                crAttch.ParameterFields(2).ClearCurrentValueAndRange
                crAttch.ParameterFields(2).AddCurrentValue SysAddress
                CRViewer1.ReportSource = crAttch
                CRViewer1.ViewReport
                CRViewer1.Zoom 75
                tabOR.TabEnabled(0) = False
                tabOR.TabEnabled(1) = True
                tabOR.Tab = 1
                tabOR.Visible = True
                Ans = MsgBox("Do you want to Print Attachment?", vbQuestion + vbYesNo, "O.R. Attachment")
                If Ans = vbYes Then CRViewer1.PrintReport
            End If
            
            
            DidRecordEntry = False
        Case 6      'Cancel
            Ans = MsgBox("Do you want to cancel payment?", vbQuestion + vbYesNo, "Cancel")
            If Ans = vbNo Then Exit Sub
            fraAllocation.Visible = False
            fraAllocationOthrData.Visible = False
            DidRecordEntry = False
            Frame6.Visible = False
            Call ClearEntries
            updRef.Visible = False
            Call LockData(True)
            Call cmdSaveCancel(False)
            Call RXResizer1_ResizeComplete
            
        Case 7  'Cash
            fraAllocation.Visible = False
            fraAllocationOthrData.Visible = False
            DidRecordEntry = True
            Call ClearEntries
            txtORno.Text = GetReceiptNo
            txtDate.Text = Format$(Date, "mm/dd/yyyy")
            cmdDetails.Caption = "Cash Invoice"
            Call LockData(False)
            Call cmdSaveCancel(True)
            lblRefNo.Visible = False
            txtRefNo.Visible = False    'set to true if updref = JOB ORDER
            updRef.Visible = False   'toggle btween SP and JO, hide txtrefno if SP
            txtRecFrom.Visible = True
            cboAcctCust.Visible = False
            txtPayable.Locked = False
            OnAcctPCV.Value = 0
            PayType = 7
            With flxDetails
                .ColWidth(0) = 0
                .ColWidth(1) = .Width * 0.98
                .ColWidth(2) = 0
                .Rows = 2
                .TextMatrix(0, 1) = "Select Invoice No:"
                .Col = 1
                .Row = 1
            End With
            flxOrder.Left = flxBal.Left
            flxOrder.Top = flxBal.Top
            flxOrder.Width = flxBal.Width
            flxOrder.Height = flxBal.Height
            flxOrder.Rows = 1
            flxOrder.Rows = 2
            flxOrder.Visible = True
            flxBal.Visible = False
            lblTotal.Caption = "TOTAL"
            lblTotal.Visible = True
            txtTotal.Visible = True
            udORCR.Value = 1
            'Display Cash Invoices for the day without OR
            Qry = "SELECT spinvoice.SPINVNO From spinvoice Where spinvoice.INVDATE = '" & Format(Date, "YYYY-MM-DD") & "' AND " & _
                  "spinvoice.INVTYPE = 'CASH' AND spinvoice.ORNO IS NULL Order By spinvoice.LASTUPDATE Desc;"
            With rsCash
                .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
                If .RecordCount > 0 Then
                    cboOthers.Clear
                    Do While Not .EOF
                        cboOthers.AddItem !SPINVNO
                        .MoveNext
                    Loop
                Else
                    cboOthers.Clear
                End If
                .Close
            End With
            Call flxDetails_Click
            
    End Select
    
    PayType = Index
    
    Screen.MousePointer = vbDefault
    Me.Enabled = True
Exit Sub
er:
    Screen.MousePointer = vbDefault
    Me.Enabled = True
    MsgBox Err.Description, vbCritical, "Error->cmdTrans_Click()"
End Sub

Function PrnRemarks(PT As Integer) As String
    Dim i As Integer, Tmp As String
    
    With flxDetails
        If Trim(.TextMatrix(1, 1)) = "" Then
            PrnRemarks = ""
            Exit Function
        End If
        Tmp = IIf(PT <> 1, "", "DOWN ") & "PAYMENT FOR" & vbCrLf & vbCrLf
        If Trim(.TextMatrix(.Rows - 1, 1)) = "" Then .RemoveItem .Rows - 1
        For i = 1 To .Rows - 1
            Tmp = Tmp & Trim(.TextMatrix(i, 1)) & IIf(i <> .Rows - 1, ", ", "")
        Next i
    End With
    PrnRemarks = Tmp
End Function

Private Sub flxBal_Click()
    If txtChkPRid.Tag = "CHECK ALLOCATED" Then Exit Sub
    
    If OnAcctPCV.Value = 1 Then ' If PayType = 2 Or OnAcctPCV.Value = 1 Then 'DP
                                ' Updated By SMC
                                ' 2016-08-05
        BalRow = flxBal.Row
        BalCol = flxBal.Col
        If BalCol <> 3 Or flxBal.Rows < 3 Then Exit Sub
        With txtAmtPd
            .Top = flxBal.CellTop + flxBal.Top
            .Left = flxBal.CellLeft + flxBal.Left
            .Width = flxBal.CellWidth
            .Text = Trim(flxBal.Text)
            .Visible = True
            .SelStart = 0
            .SelLength = Len(.Text)
            .ZOrder
            .SetFocus
        End With
    End If
End Sub

Private Sub flxBal_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call flxBal_Click
End Sub

Private Sub flxBal_Scroll()
    txtAmtPd.Visible = False
End Sub

Private Sub flxDetails_Click()
    Dim MyInv As String
    
    With flxDetails
        If .Col = 1 And PayType = 2 Then
            If .TextMatrix(.Row, 1) <> "" Then
                If InStr(.TextMatrix(.Row, 1), "Inv.#") > 0 Then
                    MyInv = Trim(Mid$(.TextMatrix(.Row, 1), 6, Len(.TextMatrix(.Row, 1)) - 5))
                    txtRefNo.Text = MyInv
                    lblRefNo.Caption = "INVOICE NO."
                    Call PopupMenu(mnuDetails)
                End If
                If InStr(.TextMatrix(.Row, 1), "J.O.#") > 0 Then
                    MyInv = Trim(Mid$(.TextMatrix(.Row, 1), 6, Len(.TextMatrix(.Row, 1)) - 5))
                    txtRefNo.Text = MyInv
                    lblRefNo.Caption = "JOB ORDER NO."
                    Call PopupMenu(mnuDetails)
                End If
            End If
        End If
    End With
    
    If PayType = 1 Then 'DP
        ActRow = flxDetails.Row
        ActCol = flxDetails.Col
        If ActRow = 0 Then Exit Sub
        With txtDP
            .Top = flxDetails.CellTop + flxDetails.Top
            .Left = flxDetails.CellLeft + flxDetails.Left
            .Width = flxDetails.CellWidth
            .Text = Trim(flxDetails.Text)
            .Visible = True
            .SelStart = 0
            .SelLength = Len(.Text)
            .ZOrder
            .SetFocus
        End With
    End If
    
    If PayType = 3 And OnAcctPCV.Value = 0 Then 'Others
        flxDetails.Col = 1
        flxDetails.Row = 1
        With cboOthers
            .Top = flxDetails.CellTop + flxDetails.Top
            .Left = flxDetails.CellLeft + flxDetails.Left
            .Width = flxDetails.CellWidth
            .Text = flxDetails.Text
            .Visible = True
            .SelStart = 0
            .SelLength = Len(.Text)
            .ZOrder
            .SetFocus
        End With
    End If
    If PayType = 7 Then 'Cash
        flxDetails.Col = 1
        flxDetails.Row = 1
        With cboOthers
            .Top = flxDetails.CellTop + flxDetails.Top
            .Left = flxDetails.CellLeft + flxDetails.Left
            .Width = flxDetails.CellWidth
            .Text = flxDetails.Text
            .Visible = True
            .SelStart = 0
            .SelLength = Len(.Text)
            .ZOrder
            .SetFocus
        End With
    End If

End Sub

Private Sub Form_Load()
    Dim rsCashier As New ADODB.Recordset
    Dim Qry As String
    Dim MyID As Integer
    
    Call LockData(True)
    Call cmdSaveCancel(False)
    JONoAcct = ""
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
                If !EMPID = UserID Then MyID = cboCashier.ListCount - 1: Cashr_Idx = cboCashier.ListCount - 1
                .MoveNext
            Loop
        End If
        .Close
    End With
    Call CboDisplay
    
    If cboCashier.ListCount > 0 And MyID >= 0 Then
        cboCashier.ListIndex = MyID
    End If
    
    If SysBranch = "LYR MOTORPARTS CORP. - SOBRECARY" Then
        ThisBranch = "MRC1"
        cboBranch.AddItem "MRC1"
    ElseIf SysBranch = "LYR MOTORPARTS CORP. - GANTE" Then
        ThisBranch = "MRC2 RETAIL"
        cboBranch.AddItem "MRC2 RETAIL"
    ElseIf SysBranch = "LYR MOTORPARTS CORP. - WHOLESALE" Then
        ThisBranch = "MRC2"
        cboBranch.AddItem "MRC1"
        cboBranch.AddItem "MRC2"
        cboBranch.AddItem "VAN1"
        cboBranch.AddItem "VAN2"
    ElseIf SysBranch = "LYR MOTORPARTS CORP. - VAN SELLING" Then
        ThisBranch = "VAN1"
        cboBranch.AddItem "VAN1"
    ElseIf SysBranch = "LYR MOTORPARTS CORP. - VAN 2" Then
        ThisBranch = "VAN2"
        cboBranch.AddItem "VAN2"
    ElseIf SysBranch = "LYR MOTORPARTS CORP. - ULAS" Then
        ThisBranch = "ULAS"
        cboBranch.AddItem "ULAS"
    ElseIf SysBranch = "MC PARTS DEPOT TORIL" Then
        ThisBranch = "TORIL"
        cboBranch.AddItem "TORIL"
    ElseIf SysBranch = "LYR MOTORPARTS CORP. - PANABO" Then
        ThisBranch = "PANABO"
        cboBranch.AddItem "PANABO"
    ElseIf SysBranch = "LYR MOTOR SHOP AND SPAREPARTS DAVAO" Then
        ThisBranch = "LYR SHOP DAVAO"
        cboBranch.AddItem "LYR SHOP DAVAO"
    End If
    
    tmrLoadCustomer.Enabled = True
    MotoRep.sBar.Panels(6).Text = "Process Official Receipt"
End Sub

 




Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    If rsAcctCust.State = adStateOpen Then rsAcctCust.Close
    If rsPrn.State = adStateOpen Then rsPrn.Close
    If rsAcct.State = adStateOpen Then rsAcct.Close
    If conVAN1.State = adStateOpen Then conVAN1.Close
    If conVAN2.State = adStateOpen Then conVAN2.Close
    If conMRC1.State = adStateOpen Then conMRC1.Close
    
    Set Receipt = Nothing
End Sub










Private Sub mnuLedger_Click()
    Dim rsLedger As New ADODB.Recordset
    Dim Qry As String
    Dim i As Integer, JOid As Long

    txtInvNo.Text = txtRefNo.Text
    With flxAcct
        For i = 1 To .Rows - 1
            If .TextMatrix(i, 0) = txtInvNo.Text Then
                txtInvDate.Text = .TextMatrix(i, 12)
                txtInvAmt.Text = Format$(.TextMatrix(i, 13), "#,##0.00")
                txtTerms.Text = .TextMatrix(i, 3) & IIf(Val(.TextMatrix(i, 3)) > 1, " Payments ", " Payment ") & LCase(.TextMatrix(i, 2))
                lblDue.Caption = .TextMatrix(i, 2) & " DUE"
                txtDue.Text = Format$(.TextMatrix(i, 5), "#,##0.00")
                txtBal.Text = Format$(.TextMatrix(i, 8), "#,##0.00")
                If lblRefNo.Caption = "INVOICE NO." Then
                    lblInvNo.Caption = "INVOICE NO."
                    lblInvDate.Caption = "INV. DATE"
                    lblInvAmt.Caption = "INV. AMOUNT"
                Else
                    lblInvNo.Caption = "JOB ORDER NO."
                    lblInvDate.Caption = "J.O. DATE"
                    lblInvAmt.Caption = "J.O. AMT"
                    JOid = Val(.TextMatrix(i, 14))
                End If
                Exit For
            End If
        Next i
    End With
    If lblRefNo.Caption = "INVOICE NO." Then
        '***Access
        'Qry = "SELECT LEDGER.REFDATE, LEDGER.REFNO, LEDGER.DEBIT, LEDGER.CREDIT, LEDGER.OSBAL From LEDGER " & _
              "Where (((LEDGER.SPINVNO) = '" & txtRefNo.Text & "') And ((LEDGER.CUSTOMERID) = " & txtCustomerID.Text & ")) ORDER BY LEDGER.SEQNO;"
        'mySQL
        Qry = "SELECT LEDGER.REFDATE, LEDGER.REFNO, LEDGER.DEBIT, LEDGER.CREDIT, LEDGER.OSBAL From LEDGER " & _
              "Where LEDGER.SPINVNO = '" & txtRefNo.Text & "' And LEDGER.CUSTOMERID = " & txtCustomerID.Text & " ORDER BY LEDGER.SEQNO;"
    Else
        '***Access
        'Qry = "SELECT LEDGER.REFDATE, LEDGER.REFNO, LEDGER.DEBIT, LEDGER.CREDIT, LEDGER.OSBAL From LEDGER " & _
              "WHERE (((LEDGER.SPINVNO)='" & txtRefNo.Text & "') AND ((LEDGER.CUSTOMERID)=" & txtCustomerID.Text & ") AND ((LEDGER.JOBID)=" & JOid & ")) ORDER BY LEDGER.SEQNO;"
        'mySQL
        Qry = "SELECT LEDGER.REFDATE, LEDGER.REFNO, LEDGER.DEBIT, LEDGER.CREDIT, LEDGER.OSBAL From LEDGER " & _
              "WHERE LEDGER.SPINVNO='" & txtRefNo.Text & "' AND LEDGER.CUSTOMERID=" & txtCustomerID.Text & " AND LEDGER.JOBID=" & JOid & " ORDER BY LEDGER.SEQNO;"
    End If
    With rsLedger
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            flxLedger.Rows = 1
            Do While Not .EOF
                flxLedger.AddItem !REFDATE & vbTab & !RefNo & vbTab & IIf(!Debit <> 0, Format$(!Debit, "#,##0.00"), "") & vbTab & IIf(!CREDIT <> 0, Format$(!CREDIT, "#,##0.00"), "") & vbTab & Format$(!OSbal, "#,##0.00")
                
                .MoveNext
            Loop
        End If
        .Close
    End With
    tabOR.Visible = True
    tabOR.TabEnabled(1) = False
    tabOR.Tab = 0
End Sub

Private Sub RXResizer1_ResizeComplete()
    Dim i As Integer
                                  
    'RXSplitter1.PanelASize = 3063
    lblHeading(2).Left = Me.Width - (lblHeading(2).Width + 200)
    lblHeading(3).Left = Me.Width - (lblHeading(3).Width + 200)
    With flxDetails
        .ColWidth(0) = 0
        .ColWidth(1) = .Width * 0.49
        .ColWidth(2) = .Width * 0.44
    End With
    With flxBal
        For i = 0 To 4
            .ColWidth(i) = (.Width - 350) / 5
        Next i
        txtTotal.Top = .Top + .Height
        txtTotal.Left = .Left + .ColWidth(0) + .ColWidth(1) + .ColWidth(2)
        lblTotal.Left = txtTotal.Left - lblTotal.Width
    End With
    With flxLedger
        For i = 0 To 4
            .ColWidth(i) = (.Width - 300) / 5
        Next i
    End With
    Frame3.Top = Frame4.Top
    Pbar.Top = Frame4.Top + Frame4.Height
    chkInvoice.Top = Frame3.Top + Frame3.Height + 10
    chkJobOrder.Top = chkInvoice.Top + 260
    
    tabOR.Top = Frame4.Top + Frame4.Height + 10
    tabOR.Left = Frame4.Left
End Sub


Private Sub tmrLoadCustomer_Timer()
On Error GoTo er
    Dim Qry As String
    Qry = "SELECT CUSTOMER.CUSTOMERID, CUSTOMER.CUSTOMERNAME, CUSTOMER.ADDRESS, CUSTOMER.OSBAL " & _
          "From CUSTOMER ORDER BY CUSTOMER.CUSTOMERNAME;"
          
    cboAcctCust.Clear
    With rsAcctCust
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            cboAcctCust.Clear
            Do While Not .EOF
                cboAcctCust.AddItem !CustomerName
                .MoveNext
            Loop
        End If
    End With

    tmrLoadCustomer.Enabled = False
Exit Sub
er:
    Call Unload(Me)
    MsgBox Err.Description, vbCritical, "Error->tmrLoadCustomer_Timer()"
End Sub

Private Sub txtAddress_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then SendKeys "{TAB}"
    If KeyCode = vbKeyUp Then SendKeys "+{TAB}"
End Sub

Private Sub txtAddress_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub


Private Sub txtAmtChange_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Then SendKeys "+{TAB}"
End Sub

Private Sub txtAmtChange_KeyPress(KeyAscii As Integer)
    Dim strValid As String
    
    strValid = "1234567890"
    
    If KeyAscii > 26 Then
        If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
    If KeyAscii = 13 Then Call cmdTrans_Click(5)
End Sub


Private Sub txtAmtPd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        txtAmtPd.Text = flxBal.TextMatrix(BalRow, 3)
        flxBal.SetFocus
        flxBal.Row = flxBal.Rows - 1
        flxBal.Col = 3
        txtAmtPd.Visible = False
    End If
End Sub

Private Sub txtAmtPd_KeyPress(KeyAscii As Integer)
    Dim strValid As String
    
    strValid = "1234567890."
    If KeyAscii > 26 Then
        If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
    If KeyAscii = 13 Then
        With flxBal
            .SetFocus
            If BalRow = .Rows - 1 Then
                .Row = .Rows - 1
            Else
                .Row = BalRow + 1
            End If
            .Col = 3
        End With
    End If
End Sub

Private Sub txtAmtPd_LostFocus()
    Dim TtlFlxAmt As Currency, MyBal As Currency
    Dim i As Integer
    
    'Validate entry if exceeds txtpayable
    If Val(txtAmtPd.Text) > Val(Replace(txtPayable.Text, ",", "")) Then
        MsgBox "Amount Paid must not exceed Total Payable!", vbExclamation + vbOKOnly, "Validation Error"
        txtAmtPd.Visible = False
        Exit Sub
    End If
    If Val(Replace(flxBal.TextMatrix(BalRow, 2), ",", "")) < Val(txtAmtPd.Text) Then 'Bal<Amt pd
        MsgBox "Amount Paid must not exceed Invoice Balance!", vbExclamation + vbOKOnly, "Validation Error"
        txtAmtPd.Visible = False
        Exit Sub
    End If
        
    flxBal.TextMatrix(BalRow, 3) = Format(txtAmtPd.Text, "###0.00")
    txtAmtPd.Visible = False
    'update related tables
    flxBal.TextMatrix(BalRow, 4) = Format(Val(Replace(flxBal.TextMatrix(BalRow, 2), ",", "")) - Val(flxBal.TextMatrix(BalRow, 3)), "###0.00")
    If PayType = 2 Then
        flxAcct.TextMatrix(BalRow, 10) = txtAmtPd.Text
        flxAcct.TextMatrix(BalRow, 11) = flxBal.TextMatrix(BalRow, 4)
    End If
    With flxBal
        'Adjust Balance to Payable if 2 invoice only
        If .Rows = 3 Then
            If BalRow = 1 Then
                .TextMatrix(2, 3) = Format(Val(Replace(txtPayable.Text, ",", "")) - Val(.TextMatrix(1, 3)), "###0.00")
                .TextMatrix(2, 4) = Format(PureNum(.TextMatrix(2, 2)) - Val(.TextMatrix(2, 3)), "###0.00")
                If PayType = 2 Then
                    flxAcct.TextMatrix(2, 10) = .TextMatrix(2, 3)
                    flxAcct.TextMatrix(2, 11) = .TextMatrix(2, 4)
                End If
            Else
                .TextMatrix(1, 3) = Format(Val(Replace(txtPayable.Text, ",", "")) - Val(.TextMatrix(2, 3)), "###0.00")
                .TextMatrix(1, 4) = Format(PureNum(.TextMatrix(1, 2)) - Val(.TextMatrix(1, 3)), "###0.00")
                If PayType = 2 Then
                    flxAcct.TextMatrix(1, 10) = .TextMatrix(1, 3)
                    flxAcct.TextMatrix(1, 11) = .TextMatrix(1, 4)
                End If
            End If
            txtTotal.Text = Format(Val(Replace(.TextMatrix(1, 3), ",", "")) + Val(Replace(.TextMatrix(2, 3), ",", "")), "#,##0.00")
        Else
            TtlFlxAmt = 0
            For i = 1 To .Rows - 1
                TtlFlxAmt = TtlFlxAmt + Val(Replace(.TextMatrix(i, 3), ",", ""))
            Next i
            txtTotal.Text = Format(TtlFlxAmt, "#,##0.00")
            If TtlFlxAmt <> Val(Replace(txtPayable.Text, ",", "")) Then
                MsgBox "Amount Paid is not equal to Total Payable!", vbExclamation + vbOKOnly, "Validation Error"
            End If
        End If
    End With
    With flxBal
        .SetFocus
        If BalRow = .Rows - 1 Then
            .Row = .Rows - 1
        Else
            .Row = BalRow + 1
        End If
        .Col = 3
    End With
End Sub

Private Sub txtAmtRec_GotFocus()
    If txtAmtRec.Text <> "" Then
        txtAmtRec.SelStart = 0
        txtAmtRec.SelLength = Len(txtAmtRec.Text)
    End If
End Sub

Private Sub txtAmtRec_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then SendKeys "+{TAB}"   'txtAddress.SetFocus  'txtAmtChange.SetFocus
    If KeyCode = vbKeyUp Then SendKeys "+{TAB}"
End Sub

Private Sub txtAmtRec_KeyPress(KeyAscii As Integer)
    Dim strValid As String
    
    strValid = "1234567890."
    If KeyAscii > 26 Then
        If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
    If KeyAscii = 13 Then txtAmtChange.SetFocus
End Sub

Private Sub txtAmtRec_LostFocus()
    If Val(Replace(txtPayable.Text, ",", "")) > Val(Replace(txtAmtRec.Text, ",", "")) And txtAmtRec.Text <> "" Then
        MsgBox "Insufficient Amount Received!", vbInformation + vbOKOnly, "Invoice"
        txtAmtChange.SetFocus
        Exit Sub
    End If
    If txtAmtRec.Text <> "" Then
        txtAmtRec.Text = Format$(txtAmtRec.Text, "#,###,##0.00")
        txtAmtChange.Text = Format(Val(Replace(txtAmtRec.Text, ",", "")) - Val(Replace(txtPayable.Text, ",", "")), "#,###,##0.00")
        txtSumOf.Text = NumToText(CDbl(txtPayable.Text))
    Else
        txtAmtChange.Text = ""
        txtSumOf.Text = ""
    End If

End Sub

Private Sub txtCheckNo_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtChkDate.SetFocus
End Sub

Private Sub txtChkAcct_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtChkBank.SetFocus
End Sub


Private Sub txtChkAmt_KeyPress(KeyAscii As Integer)
    Dim strValid As String
    
    strValid = "1234567890."
    
    If KeyAscii > 26 Then
        If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
    If KeyAscii = 13 Then txtChkAcct.SetFocus
End Sub

Private Sub txtChkBank_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
End Sub

Private Sub txtChkDate_KeyPress(KeyAscii As Integer)
    Dim strValid As String
    
    strValid = "1234567890/-"
    
    If KeyAscii > 26 Then
        If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
    If KeyAscii = 13 Then txtChkAmt.SetFocus
End Sub

Private Sub txtDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then SendKeys "{TAB}"
    If KeyCode = vbKeyUp Then SendKeys "+{TAB}"
End Sub

Private Sub txtDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub


Private Sub txtDP_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        txtDP.Text = flxDetails.TextMatrix(ActRow, ActCol)
        flxDetails.SetFocus
        flxDetails.Row = flxDetails.Rows - 1
        flxDetails.Col = 1
        txtDP.Visible = False
    End If
End Sub

Private Sub txtDP_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        With flxDetails
            .SetFocus
            .Row = .Rows - 1
            .Col = 1
        End With
    End If
End Sub

Private Sub txtDP_LostFocus()
    flxDetails.TextMatrix(ActRow, 1) = Space(10) & txtDP.Text
    txtDP.Visible = False
    If Trim(flxDetails.TextMatrix(ActRow, 1)) = "" And ActRow < flxDetails.Row Then
        flxDetails.RemoveItem ActRow
        ActRow = flxDetails.Rows - 1
    ElseIf Trim(flxDetails.TextMatrix(ActRow, 1)) <> "" And ActRow = flxDetails.Rows - 1 Then
        flxDetails.AddItem ""
    End If
End Sub



Private Sub txtORno_GotFocus()
    If txtORno.Text <> "" Then
        txtORno.SelStart = 0
        txtORno.SelLength = Len(txtORno.Text)
    End If
End Sub

Private Sub txtORNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then SendKeys "{TAB}"
    If KeyCode = vbKeyUp Then SendKeys "+{TAB}"
End Sub

Private Sub txtORno_KeyPress(KeyAscii As Integer)
On Error Resume Next
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtORno, KeyAscii, ValidChars)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub
Private Sub txtORno_LostFocus()
    Dim Qry As String
    If udORCR.Value = 0 Then 'O.R.
        Qry = "SELECT PAYMENTS.ORNO From PAYMENTS WHERE PAYMENTS.ORNO='" & Trim(txtORno.Text) & "';"
    Else
        Qry = "SELECT PAYMENTS.CRNO From PAYMENTS WHERE PAYMENTS.CRNO='" & Trim(txtORno.Text) & "';"
    End If
    If CtrlNumExist(Qry) Then
        MsgBox lblORno.Caption & " " & txtORno.Text & " already exist!", vbExclamation + vbOKOnly, "Warning"
    End If
End Sub

Private Sub AutoAdjustPCV()
    Dim i As Integer
    Dim AmtPay As Currency, TtlPay As Currency
    
    AmtPay = Val(Replace(txtPayable.Text, ",", ""))
    With flxBal
        For i = 1 To .Rows - 1
            If Val(Replace(.TextMatrix(i, 2), ",", "")) <= AmtPay Then
                .TextMatrix(i, 3) = .TextMatrix(i, 2)
                AmtPay = AmtPay - Val(Replace(.TextMatrix(i, 2), ",", ""))
                TtlPay = TtlPay + Val(Replace(.TextMatrix(i, 3), ",", ""))
                .TextMatrix(i, 4) = Format(Val(Replace(.TextMatrix(i, 2), ",", "")) - Val(Replace(.TextMatrix(i, 3), ",", "")), "#,##0.00")
            ElseIf Val(Replace(.TextMatrix(i, 2), ",", "")) > AmtPay And AmtPay > 0 Then
                .TextMatrix(i, 3) = Format(AmtPay, "#,##0.00")
                TtlPay = TtlPay + AmtPay
                AmtPay = 0
                .TextMatrix(i, 4) = Format(Val(Replace(.TextMatrix(i, 2), ",", "")) - Val(Replace(.TextMatrix(i, 3), ",", "")), "#,##0.00")
            ElseIf AmtPay = 0 Then
                .TextMatrix(i, 3) = "0.00"
                .TextMatrix(i, 4) = .TextMatrix(i, 2)
            End If
        Next i
    End With
    txtTotal.Text = Format(TtlPay, "#,##0.00")
    
End Sub

Private Sub txtPayable_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then SendKeys "{TAB}"
    If KeyCode = vbKeyUp Then SendKeys "+{TAB}"
End Sub

Private Sub txtPayable_KeyPress(KeyAscii As Integer)
    Dim strValid As String
    
    strValid = "1234567890."
    If KeyAscii > 26 Then
        If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
    If KeyAscii = 13 Then txtAmtRec.SetFocus
End Sub

Private Sub txtPayable_LostFocus()
On Error GoTo er
    Dim i As Integer, D As Integer, Ans As Integer
    Dim NewBal As Currency, rBal As Currency
    Dim vFlxAcct10 As Currency, vFlxAcct11 As Currency
    
    'Exit Block Code
    'Added By SMC
    '2016-08-05
    If PayType = 2 Then 'Charge Invoice
        If DidRecordEntry = True Then
            ' If code execution goes here
            ' It will call the exit sub instruction since pay type is required allocation record
            Exit Sub
        End If
    End If
    
    
    If PayType = 0 And Val(Replace(txtPayable.Text, ",", "")) > Val(txtJOamt.Text) Then
        MsgBox "Invalid Amount entered!" & vbCrLf & "Amount Paid must not exceed Job Order Total Balance!", vbExclamation + vbOKOnly, "Quick Service"
        txtPayable.Text = txtJOamt.Text
        Exit Sub
    End If
    
    If PayType = 3 And OnAcctPCV.Value = 1 Then
        If Val(Replace(txtPayable.Text, ",", "")) > Val(Replace(txtTotal.Text, ",", "")) Then
            MsgBox "Invalid Amount entered!" & vbCrLf & "Amount Paid must not exceed Total Account Payable", vbExclamation + vbOKOnly, "Other Payment"
            txtPayable.Text = Val(Replace(txtTotal.Text, ",", ""))
        Else
            txtSumOf.Text = NumToText(CDbl(Val(Replace(txtPayable.Text, ",", ""))))
            If Val(Replace(txtPayable.Text, ",", "")) <> Val(Replace(txtTotal.Text, ",", "")) Then
                Ans = MsgBox("Do you want to Auto-Adjust Payments for PCV?", vbQuestion + vbYesNo, "Auto-Adjust")
                If Ans = vbYes Then Call AutoAdjustPCV
            End If
        End If
        Exit Sub
    End If
    
    txtSumOf.Text = NumToText(CDbl(Val(Replace(txtPayable.Text, ",", ""))))
    txtPayable.Text = Format$(txtPayable.Text, "#,###,##0.00")
    
    If PayType = 2 And chkQckSrvc.Value = 0 Then 'If AmtPd > Balance
        If txtChkPRid.Tag = "CHECK ALLOCATED" Then Exit Sub
        If flxAcct.Rows > 2 Then
            If flxBal.Rows > 1 Then
                Ans = MsgBox("Do you want to Auto-Adjust Balance?", vbQuestion + vbYesNo, "O.R.")
                If Ans = vbNo Then Exit Sub
            End If
            If Val(Replace(txtPayable.Text, ",", "")) > Val(Format(txtTotalBal.Text, "###0.00")) Then 'Amt Paid Greater than Balance
                MsgBox "Invalid Amount entered!" & vbCrLf & "Amount Paid must not exceed Total Balance!"
                txtPayable.Text = ""
                Exit Sub
            End If
            flxBal.Rows = 1
            NewBal = Val(Replace(txtPayable.Text, ",", ""))
            If NewBal = Val(txtTotalBal.Text) Then    'Adjust using Balance
                For i = 1 To flxAcct.Rows - 1
                    If NewBal >= Val(flxAcct.TextMatrix(i, 8)) Then
                        vFlxAcct10 = Val(flxAcct.TextMatrix(i, 8))
                        vFlxAcct11 = 0
                        NewBal = NewBal - Val(flxAcct.TextMatrix(i, 8))
                    Else
                        vFlxAcct10 = NewBal
                        vFlxAcct11 = Val(flxAcct.TextMatrix(i, 8)) - NewBal
                        NewBal = 0
                    End If
                    flxBal.AddItem IIf(Val(flxAcct.TextMatrix(i, 14)) > 0, "", flxAcct.TextMatrix(i, 0)) & vbTab & IIf(Val(flxAcct.TextMatrix(i, 14)) > 0, flxAcct.TextMatrix(i, 0), "") & vbTab & Format$(flxAcct.TextMatrix(i, 8), "#,##0.00") & vbTab & Format$(vFlxAcct10, "#,##0.00") & vbTab & Format$(vFlxAcct11, "#,##0.00")
                Next i
            ElseIf NewBal <= Val(txtTotalDue.Text) Then 'Adjust using Billing Amt
                For i = 1 To flxAcct.Rows - 1
                    vFlxAcct10 = Val(flxAcct.TextMatrix(i, 10))
                    If NewBal >= vFlxAcct10 Then
                        'As is AmtPd:  flxAcct.TextMatrix(i, 10) = Val(flxAcct.TextMatrix(i, 8))
                        vFlxAcct11 = Val(flxAcct.TextMatrix(i, 8)) - vFlxAcct10
                        NewBal = NewBal - vFlxAcct10
                    Else
                        vFlxAcct10 = NewBal
                        vFlxAcct11 = Val(flxAcct.TextMatrix(i, 8)) - NewBal
                        NewBal = 0
                    End If
                    flxBal.AddItem IIf(Val(flxAcct.TextMatrix(i, 14)) > 0, "", flxAcct.TextMatrix(i, 0)) & vbTab & IIf(Val(flxAcct.TextMatrix(i, 14)) > 0, flxAcct.TextMatrix(i, 0), "") & vbTab & Format$(flxAcct.TextMatrix(i, 8), "#,##0.00") & vbTab & Format$(vFlxAcct10, "#,##0.00") & vbTab & Format$(vFlxAcct11, "#,##0.00")
                Next i
            ElseIf NewBal > Val(txtTotalDue.Text) Then 'Adjust using Billing Amt
                D = 1   'ihapon kung kapila nabalikan ang invoice usually kung payment>amt due
                Do While NewBal > 0
                    For i = 1 To flxAcct.Rows - 1
                        If D = 1 Then
                            If NewBal >= Val(flxAcct.TextMatrix(i, 5)) Then
                                vFlxAcct10 = Val(flxAcct.TextMatrix(i, 5))
                                vFlxAcct11 = Val(flxAcct.TextMatrix(i, 8)) - Val(flxAcct.TextMatrix(i, 5))
                                NewBal = NewBal - vFlxAcct10
                            Else
                                vFlxAcct10 = NewBal
                                vFlxAcct11 = Val(flxAcct.TextMatrix(i, 8)) - NewBal
                                NewBal = 0
                            End If
                            flxBal.AddItem IIf(Val(flxAcct.TextMatrix(i, 14)) > 0, "", flxAcct.TextMatrix(i, 0)) & vbTab & IIf(Val(flxAcct.TextMatrix(i, 14)) > 0, flxAcct.TextMatrix(i, 0), "") & vbTab & Format$(flxAcct.TextMatrix(i, 8), "#,##0.00") & vbTab & Format$(vFlxAcct10, "#,##0.00") & vbTab & Format$(vFlxAcct11, "#,##0.00")
                        Else
                            rBal = Val(flxAcct.TextMatrix(i, 8)) - Val(flxBal.TextMatrix(i, 3)) 'NEW
                            If NewBal >= rBal Then
                                vFlxAcct10 = rBal + Val(flxBal.TextMatrix(i, 3)) 'NEW
                                vFlxAcct11 = Val(flxAcct.TextMatrix(i, 8)) - vFlxAcct10
                                NewBal = NewBal - rBal
                            Else
                                vFlxAcct10 = Val(flxBal.TextMatrix(i, 3)) + NewBal  'NEW
                                vFlxAcct11 = Val(flxAcct.TextMatrix(i, 8)) - vFlxAcct10
                                NewBal = 0
                            End If
                            flxBal.TextMatrix(i, 3) = Format$(vFlxAcct10, "#,##0.00")   'NEW
                            flxBal.TextMatrix(i, 4) = Format$(vFlxAcct11, "#,##0.00")   'NEW
                        End If
                        If NewBal = 0 Then Exit For
                    Next i
                    D = D + 1
                Loop
            End If
            flxBal.Visible = True
        Else
            If flxAcct.Rows = 1 Then
                MsgBox "No On-account record available!", vbOKOnly + vbExclamation, "Charge Inv."
                txtPayable.Text = ""
                Exit Sub
            End If
            If Val(Replace(txtPayable.Text, ",", "")) > Format(Val(flxAcct.TextMatrix(1, 8)), "###0.00") Then 'Amt Paid Greater than Balance
                MsgBox "Invalid Amount entered!" & vbCrLf & "Amount Paid must not exceed Balance!"
                txtPayable.Text = ""
                Exit Sub
            End If
            If flxAcct.Rows = 2 Then ' Val(Replace(txtTotal.Text, ",", "")
                flxAcct.TextMatrix(1, 10) = Val(Replace(txtPayable.Text, ",", ""))
                flxAcct.TextMatrix(1, 11) = Format(Val(flxAcct.TextMatrix(1, 8)), "###0.00") - Val(Replace(txtPayable.Text, ",", ""))
                flxBal.Rows = 1
                flxBal.AddItem IIf(Val(flxAcct.TextMatrix(1, 14)) > 0, "", flxAcct.TextMatrix(1, 0)) & vbTab & IIf(Val(flxAcct.TextMatrix(1, 14)) > 0, flxAcct.TextMatrix(1, 0), "") & vbTab & Format$(flxAcct.TextMatrix(1, 8), "#,##0.00") & vbTab & Format$(flxAcct.TextMatrix(1, 10), "#,##0.00") & vbTab & Format$(flxAcct.TextMatrix(1, 11), "#,##0.00")
            End If
            flxBal.Visible = True
        End If
    End If
    rBal = 0
    For i = 1 To flxBal.Rows - 1
        rBal = rBal + PureNum(flxBal.TextMatrix(i, 3))
    Next i
    txtTotal.Text = Format(rBal, "#,##0.00")
    lblTotal.Visible = True
    txtTotal.Visible = True
    
Exit Sub
er:
    MsgBox Err.Description, vbCritical, "Error->txtPayable_LostFocus()"
End Sub

Private Sub txtRecFrom_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then SendKeys "{TAB}"
    If KeyCode = vbKeyUp Then SendKeys "+{TAB}"
End Sub

Private Sub txtRecFrom_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtrefno_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtPayable.SetFocus
End Sub

Private Sub txtrefno_LostFocus()
    Dim Qry As String
    
    If txtRefNo.Text = "" Then Exit Sub
    If PayType = 0 Then
        '***Access
        'Qry = "SELECT SRVCJOBORDER.JOBORDERNO From SRVCJOBORDER " & _
              "WHERE (((SRVCJOBORDER.JOBORDERNO)='" & txtRefNo.Text & "'));"
        'mySQL
        
        Qry = "SELECT SRVCJOBORDER.JOBORDERNO From SRVCJOBORDER " & _
              "WHERE SRVCJOBORDER.JOBORDERNO='" & Trim(txtRefNo.Text) & "' AND SRVCJOBORDER.PAYTYPE='CASH';"
        
        If CtrlNumExist(Qry) Then
            Call DisplayQuickService
        Else
            MsgBox lblRefNo.Caption & " " & txtRefNo.Text & " does not exist!", vbExclamation + vbOKOnly, "Payment"
            txtRefNo.Text = ""
            Exit Sub
        End If
    ElseIf PayType = 1 Then
        '***Access
        'Qry = "SELECT SRVCJOBORDER.JOBORDERNO From SRVCJOBORDER " & _
              "WHERE (((SRVCJOBORDER.JOBORDERNO)='" & txtRefNo.Text & "'));"
        'mySQL
        Qry = "SELECT SRVCJOBORDER.JOBORDERNO From SRVCJOBORDER " & _
              "WHERE SRVCJOBORDER.JOBORDERNO='" & Trim(txtRefNo.Text) & "';"
        
        If Not CtrlNumExist(Qry) Then
            MsgBox lblRefNo.Caption & " " & txtRefNo.Text & " does not exist!", vbExclamation + vbOKOnly, "Payment"
            txtRefNo.Text = ""
            Exit Sub
        End If
    ElseIf PayType = 2 Then 'On-Account
        If txtRefNo.Text = txtPrevRef.Text Then Exit Sub
        If updRef.Value = 0 Then
            Call DisplayInvoice
        Else
            Call DisplayQuickService
        End If
    End If
End Sub

Private Sub DisplayInvoice()
    Dim rsInv As New ADODB.Recordset
    Dim Qry As String, eMsg As String
    Dim MyAmt As Double, MyRet As Currency
    
    '***Access
    'Qry = "SELECT SPINVOICE.CHARGETO, SPINVOICE.ADDRESS, SPINVOICE.AMOUNT, SPINVOICE.SPINVNO, SPINVOICE.ORNO, SPINVOICE.EMPID, SPINVOICE.CUSTOMERID " & _
          "From SPINVOICE WHERE (((SPINVOICE.SPINVNO)='" & txtRefNo.Text & "'));"
    'mySQL
    Qry = "SELECT SPINVOICE.CHARGETO, SPINVOICE.ADDRESS, SPINVOICE.AMOUNT, SPINVOICE.SPINVNO, SPINVOICE.ORNO, SPINVOICE.EMPID, SPINVOICE.CUSTOMERID " & _
          "From SPINVOICE WHERE SPINVOICE.SPINVNO='" & Trim(txtRefNo.Text) & "';"
    
    With rsInv
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        eMsg = ""
        If .RecordCount > 0 Then
            .MoveFirst
            txtPrevRef.Text = !SPINVNO
            If PayType = 2 And !CUSTOMERID > 0 Then
                cboAcctCust.Text = !chargeTo
            Else
                If PayType = 2 Then
                    CustSrch = True
                    cboAcctCust.Text = !chargeTo
                End If
                txtRecFrom.Text = !chargeTo
                txtAddress.Text = !Address
                lstChrgInvNo.AddItem !SPINVNO
                '***Access
                'Qry = "SELECT Sum(SRETURNSP.TOTALAMT) AS SumOfTOTALAMT From SRETURNSP WHERE (((SRETURNSP.INVOICE)='" & txtRefNo.Text & "'));"
                'mySQL
                Qry = "SELECT Sum(SRETURNSP.TOTALAMT) AS SumOfTOTALAMT From SRETURNSP WHERE SRETURNSP.INVOICE='" & txtRefNo.Text & "';"
                
                MyRet = Val(GetFieldValue(Qry))
                MyAmt = !Amount - MyRet
                flxDetails.AddItem vbTab & "Inv.# " & !SPINVNO & " -" & vbTab & Format$(MyAmt, "###,##0.00") & Space$(5)
                txtPayable.Text = Format$(MyAmt, "#####0.00")
                txtSumOf.Text = NumToText(MyAmt)
                'txtEmpID.Text = !EmpId
                If Not IsNull(!ORNO) Then
                    If Len(!ORNO) >= 1 Then eMsg = "Invoice already paid thru OR#:" & !ORNO
                End If
            End If
        Else
            txtRecFrom.Text = ""
            txtAddress.Text = ""
            txtPayable.Text = ""
            txtSumOf.Text = ""
            txtPrevRef.Text = ""
            flxDetails.Rows = 1
            eMsg = "Invoice not found!"
        End If
        .Close
    End With
    If eMsg <> "" Then
        MsgBox eMsg, vbCritical + vbOKOnly, "Invoice Status"
        Call ClearEntries
        txtDate.Text = Format$(Date, "mm/dd/yyyy")
        txtORno.Text = GetReceiptNo
        cmdDetails.Caption = "ON-ACCOUNT"
    End If
End Sub

Private Sub DisplayQuickService()
    Dim rsJob As New ADODB.Recordset
    Dim rsJobDtls As New ADODB.Recordset
    Dim Qry As String, Ttl As Double
    Dim eMsg As String, MyOR As String, ORAmt As Currency, InvAmt As Currency, MyRet As Currency
    flxJOdtls.Visible = True
    '***Access
    'Qry = "SELECT SRVCJOBORDER.BALANCE, SRVCUSTOMER.CUSTOMER, SRVCUSTOMER.ADDRESS, SRVCJOBORDER.JOBORDERID, SRVCJOBORDER.JOBORDERNO, SRVCJOBORDER.ACCOUNTNO, SRVCJOBORDER.ORNO, SRVCJOBORDER.EMPID " & _
          "FROM SRVCJOBORDER INNER JOIN SRVCUSTOMER ON SRVCJOBORDER.ACCOUNTNO = SRVCUSTOMER.ACCOUNTNO WHERE (((SRVCJOBORDER.JOBORDERNO)='" & txtRefNo.Text & "'));"
    'mySQL
    Qry = "SELECT SRVCJOBORDER.BALANCE, SRVCUSTOMER.CUSTOMER, SRVCUSTOMER.ADDRESS, SRVCJOBORDER.JOBORDERID, SRVCJOBORDER.JOBORDERNO, SRVCJOBORDER.ACCOUNTNO, SRVCJOBORDER.ORNO, SRVCJOBORDER.EMPID " & _
          "FROM SRVCJOBORDER INNER JOIN SRVCUSTOMER ON SRVCJOBORDER.ACCOUNTNO = SRVCUSTOMER.ACCOUNTNO WHERE SRVCJOBORDER.JOBORDERNO='" & txtRefNo.Text & "';"
    
    With rsJob
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            eMsg = ""
            If txtJobID.Text <> !JOBORDERID Then
                If PayType = 2 Then
                    txtCustomerID.Text = ""
                    CustSrch = True
                    cboAcctCust.Text = !CUSTOMER
                End If
                txtRecFrom.Text = !CUSTOMER
                txtAddress.Text = !Address
                If PayType = 0 Then
                    flxDetails.AddItem vbTab & "JO#:" & !JOBORDERNO & vbTab & Format$(!Balance, "###,##0.00") & Space$(5)
                    txtPayable.Text = Format$(!Balance, "#####0.00")
                    txtSumOf.Text = NumToText(!Balance)
                    If !Balance <= 0 Then eMsg = "Job Order already paid!" & vbCrLf & "Customer: " & !CUSTOMER & vbCrLf & "OR# " & !ORNO
                End If
                txtJobID.Text = !JOBORDERID
                txtAcctNo.Text = !ACCOUNTNO
                'txtEmpID.Text = !EmpId
                txtJOamt.Text = Format(!Balance, "###0.00")
                
                If Not (DidRecordEntry = True And PayType = 2) Then
                    txtPayable.Locked = False
                End If
                
            Else
                eMsg = "Same Entry"
            End If
            
        Else
            txtRecFrom.Text = ""
            txtAddress.Text = ""
            txtJobID.Text = ""
            txtPayable.Text = ""
            txtSumOf.Text = ""
            txtAcctNo.Text = ""
            flxDetails.Rows = 1
            eMsg = "Job Order not found!"
        End If
        .Close
    End With
    If eMsg <> "" Then
        If eMsg = "Same Entry" Then Exit Sub
        MsgBox eMsg, vbExclamation + vbOKOnly
        Call ClearEntries
        txtDate.Text = Format$(Date, "mm/dd/yyyy")
        txtORno.Text = GetReceiptNo
        cmdDetails.Caption = IIf(PayType <> 0, "ON-ACCOUNT", "QUICK SERVICE")
        Exit Sub
    End If
    'NEW validation of Job Order with Requisition and Invoices
    '***Access
    'Qry = "SELECT REQUISITION.INVOICENO FROM SRVCJOBORDER INNER JOIN REQUISITION ON SRVCJOBORDER.JOBORDERID = REQUISITION.JOBORDERID " & _
          "WHERE (((SRVCJOBORDER.JOBORDERNO)='" & txtRefNo.Text & "') AND (Not (REQUISITION.INVOICENO) Is Null));"
    'mySQL
    Qry = "SELECT REQUISITION.INVOICENO FROM SRVCJOBORDER INNER JOIN REQUISITION ON SRVCJOBORDER.JOBORDERID = REQUISITION.JOBORDERID " & _
          "WHERE SRVCJOBORDER.JOBORDERNO='" & txtRefNo.Text & "' AND Not (REQUISITION.INVOICENO) Is Null;"
    Call displayJOdetails
    With rsJobDtls
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            Ttl = 0
            lstChrgInvNo.Clear
            Do While Not .EOF
                Qry = "SELECT SPINVOICE.ORNO From SPINVOICE WHERE (((SPINVOICE.SPINVNO)='" & !InvoiceNO & "'));"
                'mySQL
                'Qry = "SELECT SPINVOICE.ORNO From SPINVOICE WHERE SPINVOICE.SPINVNO='" & !InvoiceNO & "';"
                MyOR = GetFieldValue(Qry)
                '***Access
                'Qry = "SELECT SPINVOICE.AMOUNT From SPINVOICE WHERE (((SPINVOICE.SPINVNO)='" & !InvoiceNO & "'));"
                'mySQL
                Qry = "SELECT SPINVOICE.AMOUNT From SPINVOICE WHERE SPINVOICE.SPINVNO='" & !InvoiceNO & "';"
                InvAmt = Val(GetFieldValue(Qry))
                If MyOR <> "" Then  'Invoice with OR
                    '***Access
                    'Qry = "SELECT [SPAREPARTS]+[DOWNPAYMENT] AS ORAMT FROM PAYMENTS WHERE (((PAYMENTS.ORNO)='" & MyOR & "'));"
                    'mySQL
                    Qry = "SELECT PAYMENTS.SPAREPARTS+DOWNPAYMENT AS ORAMT FROM PAYMENTS WHERE PAYMENTS.ORNO='" & MyOR & "';"
                    ORAmt = Val(GetFieldValue(Qry))
                    '***Access
                    'Qry = "SELECT Sum([RECEIVE]*[PRICE]) AS AMTRET From SPITR " & _
                          "WHERE (((SPITR.TRANSTYPE)='SALES RETURN') AND ((SPITR.PACKINGLIST)='" & !InvoiceNO & "'));"
                    'mySQL
                    Qry = "SELECT Sum(SPITR.RECEIVE*SPITR.PRICE) AS AMTRET From SPITR " & _
                          "WHERE SPITR.TRANSTYPE='SALES RETURN' AND SPITR.PACKINGLIST='" & !InvoiceNO & "';"
                    MyRet = Val(GetFieldValue(Qry))
                    If (ORAmt + MyRet) < InvAmt Then 'OR Amt less than invoice amt
                        flxDetails.AddItem vbTab & "S.I.#" & !InvoiceNO & vbTab & Format$(InvAmt - (ORAmt + MyRet), "#,##0.00") & Space$(5)
                        lstChrgInvNo.AddItem !InvoiceNO
                        Ttl = Ttl + (InvAmt - (ORAmt + MyRet))
                    End If
                Else
                    flxDetails.AddItem vbTab & "S.I.#" & !InvoiceNO & vbTab & Format$(InvAmt, "#,##0.00") & Space$(5)
                    lstChrgInvNo.AddItem !InvoiceNO
                    Ttl = Ttl + InvAmt
                End If
                .MoveNext
            Loop
        End If
        .Close
    End With
    If Ttl > 0 And PayType = 0 Then
        eMsg = "Please process payment for Invoice before paying Job Order!"
    End If
    If PayType = 2 Then     'Down Payment for JO & SP
        If Ttl = 0 Then
            eMsg = "No available unpaid Charge Invoice for JO# " & txtRefNo.Text
        Else
            txtPayable.Text = Format(Ttl, "###0.00")
            Call txtPayable_LostFocus
        End If
    End If
    If eMsg <> "" Then
        MsgBox eMsg, vbCritical + vbOKOnly, "Warning"
        Call ClearEntries
    End If
End Sub

Private Sub displayJOdetails()
    Dim rsJodtls As New ADODB.Recordset
    If (JONoAcct <> "") Then
        Qry = "SELECT s.JOBREQUEST,FSC,LABOR,DISCOUNT,TOTAL FROM srvjobdetails s WHERE s.JOBORDERID = (SELECT s.JOBORDERID FROM srvcjoborder s WHERE s.JOBORDERNO = '" & JONoAcct & "');"
    Else
        Qry = "SELECT s.JOBREQUEST,FSC,LABOR,DISCOUNT,TOTAL FROM srvjobdetails s WHERE s.JOBORDERID = (SELECT s.JOBORDERID FROM srvcjoborder s WHERE s.JOBORDERNO = '" & txtRefNo.Text & "');"
    End If
    rsJodtls.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    Set flxJOdtls.DataSource = rsJodtls 'MSHFLXGRID LANG NI MAG EFFECT DILI MSFLXGRID
    rsJodtls.Close
    flxJOdtls.Refresh
    
End Sub
'Commented by SMC
'2016-08-04
'Function GetAddress() As String
'    Dim rsAdrs As New ADODB.Recordset
'    Dim Qry As String, MyRef As String
'
'    If txtSoldTo.Text = "" Then
'        GetAddress = ""
'        Exit Function
'    End If
'    '***Access
'    'Qry = "SELECT SRVCUSTOMER.ADDRESS From SRVCUSTOMER " & _
'          "WHERE (((SRVCUSTOMER.CUSTOMER)='" & txtSoldTo.Text & "'));"
'    'mySQL
'    Qry = "SELECT SRVCUSTOMER.ADDRESS From SRVCUSTOMER " & _
'          "WHERE SRVCUSTOMER.CUSTOMER='" & txtSoldTo.Text & "';"
'    With rsAdrs
'        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
'        If .RecordCount > 0 Then
'            .MoveFirst
'            GetAddress = IIf(IsNull(!Address), "", !Address)
'        Else
'            GetAddress = ""
'        End If
'        .Close
'    End With
'End Function


'Commented By SMC
'2016-08-04
'Function ValidateWalkIn() As String
'    Dim eMsg As String, vQry As String
'    eMsg = ""
'    If Val(Replace(txtPayable.Text, ",", "")) <> PureNum(txtTotal.Text) Then eMsg = "Amount Payable did not match Sales Order Total!"
'    If txtInvNo.Text = "" Then
'        eMsg = eMsg & vbCrLf & "Incomplete entry on field Invoice No."
'    Else
'        '***Access
'        'vQry = "SELECT SPINVOICE.SPINVNO From SPINVOICE WHERE (((SPINVOICE.SPINVNO)='" & txtInvNo.Text & "'));"
'        'mySQL
'        vQry = "SELECT SPINVOICE.SPINVNO From SPINVOICE WHERE SPINVOICE.SPINVNO='" & txtInvNo.Text & "';"
'        If CtrlNumExist(vQry) Then eMsg = eMsg & vbCrLf & "Invoice No. already exist"
'    End If
'    If Not IsDate(txtDate.Text) Then eMsg = eMsg & vbCrLf & "Invalid entry on field Date"
'    If txtSoldTo.Text = "" Then eMsg = eMsg & vbCrLf & "Incomplete entry on field Sold To"
'    If txtAddress.Text = "" Then eMsg = eMsg & vbCrLf & "Incomplete entry on field Address"
'    If Val(Replace(txtAmtRec.Text, ",", "")) < Val(Replace(txtPayable.Text, ",", "")) Then eMsg = eMsg & vbCrLf & "Insufficient amount received!"
'    ValidateWalkIn = eMsg
'End Function
                                    
Function ValidateReceipt() As String
    Dim TaxAmount As Double
    Dim eMsg As String, vQry As String, myRem As String
    Dim Ans As Integer, i As Integer, MyBal As Currency, cPay As Currency, OSbal As Currency
    
    eMsg = ""
    If txtORno.Text = "" Then
        eMsg = eMsg & vbCrLf & "Incomplete entry on field " & lblORno.Caption & "!"
    Else
        'mySQL
        If udORCR.Value = 0 Then 'OR
            vQry = "SELECT PAYMENTS.ORNO From PAYMENTS WHERE PAYMENTS.ORNO='" & txtORno.Text & "';"
        Else
            vQry = "SELECT PAYMENTS.CRNO From PAYMENTS WHERE PAYMENTS.CRNO='" & txtORno.Text & "';"
        End If
        
        If CtrlNumExist(vQry) Then eMsg = eMsg & vbCrLf & lblORno.Caption & " already exist"
    End If
    If Not IsDate(txtDate.Text) Then eMsg = eMsg & vbCrLf & "Invalid entry on field Date"
    If Val(Replace(txtAmtRec.Text, ",", "")) = 0 Then eMsg = eMsg & vbCrLf & "No Amount Received"
    Select Case PayType
        Case 0  'Quick Service
            If txtJobID.Text = "" Then eMsg = eMsg & vbCrLf & "Invalid Job Order No."
            If Val(Replace(txtPayable.Text, ",", "")) = 0 Then
                eMsg = eMsg & vbCrLf & "No Amount Payable specified"
            ElseIf Val(Replace(txtPayable.Text, ",", "")) > Val(txtJOamt.Text) Then
                eMsg = eMsg & vbCrLf & "Amount Paid must not exceed Job Order Total Balance!"
            End If
        Case 1  'Down Payment
            If txtCustomerID.Text = "" Then eMsg = eMsg & vbCrLf & "Please select a Customer from list."
            If Val(Replace(txtPayable.Text, ",", "")) = 0 Then eMsg = eMsg & vbCrLf & "Please specify DP Amount in Total Payable"
        Case 2
            If chkQckSrvc.Value = 0 Then
                If txtCustomerID.Text = "" Then eMsg = eMsg & vbCrLf & "Please select a Customer from list."
                
                If Val(txtChkPRid.Text) = 0 Then    'encoded 1/13/2014 because of error in validation, flxAcct not used in check process
                    With flxAcct
                        For i = 1 To .Rows - 1
                            'Balance < Amount Paid
                            If Val(Format(.TextMatrix(i, 8), "###0.00")) < Val(Format(.TextMatrix(i, 10), "###0.00")) Then eMsg = eMsg & vbCrLf & "Excess payment for Inv. No. " & .TextMatrix(i, 0)
                        Next i
                    End With
                End If
                With flxBal
                    If .Rows > 2 Then
                        cPay = 0
                        For i = 1 To .Rows - 1
                            cPay = cPay + Val(Replace(.TextMatrix(i, 3), ",", ""))
                        Next i
                        'If Format(cPay, "###0.00") <> Val(Replace(txtPayable.Text, ",", "")) Then eMsg = eMsg & vbCrLf & "Allocation of Amount Paid not balance with Total Payable!"
                        TaxAmount = Replace(txtTaxAmount.Text, ",", "")
                        If Val(Format(cPay - TaxAmount, "###0.00")) <> Val(Replace(txtPayable.Text, ",", "")) Then eMsg = eMsg & vbCrLf & "Allocation of Amount Paid not balance with Total Payable!"
                    End If
                End With
            End If
            If Val(Replace(txtPayable.Text, ",", "")) = 0 Then eMsg = eMsg & vbCrLf & "Invalid Amount of Total Payable!"
            If Val(Replace(txtAmtRec.Text, ",", "")) < Val(Replace(txtPayable.Text, ",", "")) Then eMsg = eMsg & vbCrLf & "Insufficient amount received!"
            
        Case 3
            If cboAcctCust.Text = "" Then eMsg = eMsg & vbCrLf & "No Customer Name specified"
            If txtAddress.Text = "" Then eMsg = eMsg & vbCrLf & "Incomplete entry on field Address"
            If OnAcctPCV.Value = 1 Then
                If Val(Replace(txtTotal.Text, ",", "")) <> Val(Replace(txtPayable.Text, ",", "")) Then eMsg = eMsg & vbCrLf & "PCV total amount paid not equal to Total Payable!"
                If Val(Replace(txtAmtRec.Text, ",", "")) < Val(Replace(txtPayable.Text, ",", "")) Then eMsg = eMsg & vbCrLf & "Insufficient amount received!"
            Else
                If cboOthers.Text = "" Then eMsg = eMsg & vbCrLf & "Please select type of Payment"
            End If
    End Select
    If chkCheck.Value = 1 Then
        If Val(txtChkPRid.Text) = 0 Then eMsg = eMsg & vbCrLf & "Invalid Check Details"
        If Val(Replace(txtChkAmt.Text, ",", "")) < Val(Replace(txtAmtRec.Text, ",", "")) Then eMsg = eMsg & vbCrLf & "Amount Received must not be greater than Check Amount"
    End If
    If myAcctCode = "" Then eMsg = eMsg & vbCrLf & "Please select Account Title!"
    ValidateReceipt = eMsg
End Function

Private Sub SaveJOpay()
    Dim rsPay As New ADODB.Recordset
    Dim rsJO As New ADODB.Recordset
    Dim Qry As String, AllInv As String, i As Integer
    Dim MyOR As String, myRem As String, MyBal As Single
    Dim myORNO As String
    Dim QryB As String
    Dim myBALANCE As Double
    Dim myPayDone As Integer
    'mySQL
    myRem = "JO# " & txtRefNo.Text & "  " & cboPayType.Text
    If udORCR.Value = 0 Then 'OR
        LYRcon.Execute "Insert Into payments( PAYDATE, ORNO, ACCOUNTNO, SERVICE, REMARKS, CUSTOMERNAME, EMPID, UPDATEDBY, LASTUPDATE, ACCT_CODE)" & _
                "VALUES('" & Format(txtDate.Text, "yyyy-mm-dd") & "', '" & txtORno.Text & "', " & txtAcctNo.Text & ", " & Val(Replace(txtPayable.Text, ",", "")) & ", " & _
                "'" & myRem & "', '" & txtRecFrom.Text & "', " & Val(txtEmpID.Text) & ", " & UserID & ", '" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "','" & myAcctCode & "');"
    Else
        LYRcon.Execute "Insert Into payments( PAYDATE, CRNO, ACCOUNTNO, SERVICE, REMARKS, CUSTOMERNAME, EMPID, UPDATEDBY, LASTUPDATE, ACCT_CODE)" & _
                "VALUES('" & Format(txtDate.Text, "yyyy-mm-dd") & "', '" & txtORno.Text & "', " & txtAcctNo.Text & ", " & Val(Replace(txtPayable.Text, ",", "")) & ", " & _
                "'" & myRem & "', '" & txtRecFrom.Text & "', " & Val(txtEmpID.Text) & ", " & UserID & ", '" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "','" & myAcctCode & "');"
    End If
    
    
    'mySQL
    Qry = "SELECT SRVCJOBORDER.ORNO From SRVCJOBORDER WHERE SRVCJOBORDER.JOBORDERNO='" & txtRefNo.Text & "';"
    myORNO = GetFieldValue(Qry)
    If myORNO = "" Then
        myORNO = txtORno.Text
    Else
        myORNO = MyOR & ", " & txtORno.Text
    End If
    
    QryB = "SELECT SRVCJOBORDER.BALANCE From SRVCJOBORDER WHERE SRVCJOBORDER.JOBORDERNO='" & txtRefNo.Text & "';"
    rsJO.Open QryB, LYRcon, adOpenKeyset, adLockOptimistic
        MyBal = IIf(IsNull(rsJO!Balance), 0, rsJO!Balance)
        myBALANCE = MyBal - Val(Replace(txtPayable.Text, ",", ""))
        If MyBal - Val(Replace(txtPayable.Text, ",", "")) <= 0 Then
            myPayDone = 1
        Else
            myPayDone = 0
        End If
    rsJO.Close
    LYRcon.Execute "UPDATE SRVCJOBORDER SET ORNO = '" & myORNO & "', BALANCE = " & myBALANCE & ", PAYMENTDONE = " & myPayDone & " " & _
                   "WHERE SRVCJOBORDER.JOBORDERNO='" & txtRefNo.Text & "';"
End Sub

Private Sub SaveDP_Others()  'DP
    Dim rsPay As New ADODB.Recordset
    Dim rsDP As New ADODB.Recordset
    Dim PCVID As Long, CntAtt As Integer
    Dim Qry As String, AllInv As String, i As Integer, myRem As String
    Dim OSbal As Double
    Dim MyDPayments As Double
    Dim MyRemarks As String
    Dim OthersPay As Double
    
    CntAtt = 0
    If PayType = 3 And OnAcctPCV = 1 Then
        'Display Summary of payment in FlxDetails
        With flxDetails
            .ColWidth(0) = 0
            .ColWidth(1) = .Width * 0.49
            .ColWidth(2) = .Width * 0.44
            .Rows = 0
        End With
        myRem = ""
        flxDetails.Rows = 1
        For i = 1 To flxBal.Rows - 1
            If Val(flxBal.TextMatrix(i, 3)) > 0 Then
                flxDetails.AddItem vbTab & "PCV# " & flxBal.TextMatrix(i, 1) & ":" & vbTab & Format(Val(flxBal.TextMatrix(i, 3)), "#,##0.00")
                If Val(flxBal.TextMatrix(i, 4)) = 0 Then
                    flxDetails.AddItem vbTab & vbTab & "FULLY PAID"
                    myRem = IIf(myRem <> "", myRem & ", ", "") & "PCV#: " & flxBal.TextMatrix(i, 0) & " Full"
                Else
                    flxDetails.AddItem vbTab & vbTab & "Bal: " & Format$(flxBal.TextMatrix(i, 4), "#,##0.00")
                    myRem = IIf(myRem <> "", myRem & ", ", "") & "PCV#: " & flxBal.TextMatrix(i, 0) & " Prtl"
                End If
                CntAtt = CntAtt + 1
            End If
            OSbal = OSbal + Val(Replace(flxBal.TextMatrix(i, 4), ",", ""))
        Next i
        If OSbal > 0 Then
            flxDetails.AddItem ""
            flxDetails.AddItem vbTab & "Outstanding Bal." & vbTab & Format(OSbal, "#,##0.00")
        End If
    End If
    
    'mySQL
    If PayType = 1 Then
        MyDPayments = Val(Replace(txtPayable.Text, ",", ""))
        txtRecFrom.Text = cboAcctCust.Text
        MyRemarks = lblRefNo.Caption & " " & cboPayType.Text
    Else
        OthersPay = Val(Replace(txtPayable.Text, ",", ""))
        If OnAcctPCV.Value = 1 Then
            MyRemarks = "Other Accts: " & myRem & " " & cboPayType.Text
        Else
            MyRemarks = cboOthers.Text & " " & cboPayType.Text
        End If
    End If
    
    If Len(MyRemarks) > 250 Then MyRemarks = Left$(MyRemarks, 250)
    If udORCR.Value = 0 Then  'OR
       LYRcon.Execute "insert into payments(PAYDATE, ORNO, CUSTOMERNAME, DOWNPAYMENT, REMARKS, OTHERS, EMPID, UPDATEDBY, CUSTOMERID, LASTUPDATE, ACCT_CODE)" & _
               "VALUES('" & Format(txtDate.Text, "yyyy-mm-dd") & "', '" & txtORno.Text & "', '" & cboAcctCust.Text & "', " & Val(MyDPayments) & ", '" & MyRemarks & "', " & _
               "" & Val(OthersPay) & ", " & Val(txtEmpID.Text) & ", " & UserID & ", " & Val(txtCustomerID.Text) & ", '" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "', '" & myAcctCode & "');"
    Else  'Cr
       LYRcon.Execute "insert into payments(PAYDATE, CRNO, CUSTOMERNAME, DOWNPAYMENT, REMARKS, OTHERS, EMPID, UPDATEDBY, CUSTOMERID, LASTUPDATE, ACCT_CODE)" & _
               "VALUES('" & Format(txtDate.Text, "yyyy-mm-dd") & "', '" & txtORno.Text & "', '" & cboAcctCust.Text & "', " & Val(MyDPayments) & ", '" & MyRemarks & "', " & _
               "" & Val(OthersPay) & ", " & Val(txtEmpID.Text) & ", " & UserID & ", " & Val(txtCustomerID.Text) & ", '" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "', '" & myAcctCode & "');"
    End If
    
    '***Access
    'If PayType = 1 Then
    '    With rsDP
    '        .Open "DPLedger", LYRcon, adOpenKeyset, adLockOptimistic
    '        .AddNew
    '        !ORNO = txtORno.Text
    '        !ORDATE = txtDate.Text
    '        !ORAmount = Val(Replace(txtPayable.Text, ",", ""))
    '        !Status = "OPEN"
    '        !BALANCE = Val(Replace(txtPayable.Text, ",", ""))
    '        !CUSTOMERID = Val(txtCustomerID.Text)
    '        If flxDetails.Rows >= 2 Then
    '            If Trim(flxDetails.TextMatrix(1, 1)) <> "" Then
    '                myRem = ""
    '                For i = 1 To flxDetails.Rows - 1
    '                    If Trim(flxDetails.TextMatrix(i, 1)) <> "" Then
    '                        myRem = myRem & IIf(myRem <> "", ", ", "DP for ") & Trim(flxDetails.TextMatrix(i, 1))
    '                    End If
    '                Next i
    '            End If
    '        End If
    '        !Remarks = myRem
    '        .Update
    '        .Close
    '    End With
    'End If
      
      
    'mySQL
    If flxDetails.Rows >= 2 Then
        If Trim(flxDetails.TextMatrix(1, 1)) <> "" Then
            myRem = ""
            For i = 1 To flxDetails.Rows - 1
                If Trim(flxDetails.TextMatrix(i, 1)) <> "" Then
                    myRem = myRem & IIf(myRem <> "", ", ", "DP for ") & Trim(flxDetails.TextMatrix(i, 1))
                End If
            Next i
        End If
    End If
    If PayType = 1 Then
        LYRcon.Execute "INSERT INTO dpledger(ORNO, ORDATE, ORAMOUNT, `STATUS`, BALANCE, CUSTOMERID, REMARKS)" & _
            "VALUES('" & txtORno.Text & "', '" & Format(txtDate.Text, "yyyy-mm-dd") & "', " & Val(Replace(txtPayable.Text, ",", "")) & ", 'OPEN', " & _
            "" & Val(Replace(txtPayable.Text, ",", "")) & ", " & Val(txtCustomerID.Text) & ", '" & myRem & "')"
    End If
    If PayType = 3 And OnAcctPCV.Value = 1 Then
        With flxBal
            For i = 1 To .Rows - 1
                If Val(.TextMatrix(i, 3)) > 0 Then
                    Qry = "SELECT Disbursement.VOUCHERID From Disbursement " & _
                          "WHERE (((Disbursement.CVNUMBER)=" & Val(.TextMatrix(i, 1)) & ") AND ((Disbursement.DISBURSETYPE) Not Like 'Cancelled%'));"
                    PCVID = Val(GetFieldValue(Qry))
                    'mySQL
                    LYRcon.Execute "UPDATE ONACCTPCV SET ONACCTPCV.BALANCE = " & Val(Replace(.TextMatrix(i, 4), ",", "")) & ", ONACCTPCV.ORNO = '" & txtORno.Text & "' " & _
                                   "WHERE ONACCTPCV.CUSTOMERID=" & txtCustomerID.Text & " AND ONACCTPCV.VOUCHERID=" & PCVID & ";"
                End If
            Next i
        End With
        If CntAtt > 6 Then txtPaymentID.Tag = "ATTACHMENT"
    End If
End Sub

Function MaxSeqNo(MyInv As String) As Integer
    Dim rsSeq As New ADODB.Recordset
    Dim Qry As String
    
    '***Access
    'Qry = "SELECT Max(LEDGER.SEQNO) AS MaxOfSEQNO From LEDGER WHERE (((LEDGER.SPINVNO)='" & MyInv & "'));"
    'mySQL
    Qry = "SELECT Max(round(LEDGER.SEQNO)) AS MaxOfSEQNO From LEDGER WHERE LEDGER.SPINVNO='" & MyInv & "';"
    
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

Function GetSPttlAmtPd() As Currency
    Dim myTtl As Currency
    Dim i As Integer
    
    myTtl = 0
    With flxBal
        For i = 1 To .Rows - 1
            If .TextMatrix(i, 0) <> "" Then
                myTtl = myTtl + Val(Replace(.TextMatrix(i, 3), ",", ""))
            End If
        Next i
    End With
    GetSPttlAmtPd = myTtl
End Function
Function GetSRVCttlAmtPd() As Currency
    Dim myTtl As Currency
    Dim i As Integer
    
    myTtl = 0
    With flxBal
        For i = 1 To .Rows - 1
            If .TextMatrix(i, 1) <> "" Then
                myTtl = myTtl + Val(Replace(.TextMatrix(i, 3), ",", ""))
            End If
        Next i
    End With
    GetSRVCttlAmtPd = myTtl
End Function

Private Sub SaveOnAccount()
    Dim MyDate As Date
    Dim checkPayVal As Integer
    Dim rsPay As New ADODB.Recordset
    Dim rsLedger As New ADODB.Recordset
    Dim rsAttDtls As New ADODB.Recordset
    Dim Qry As String, Qry2 As String, jobOR As String, myRem As String, InvDate As String, MyJo As String, MyRef As String
    Dim i As Integer, CntAtt As Integer, TtlPay As Currency, SPamtPd As Currency, SVCamtPd As Currency
    Dim CustID As Long, SeqN As Integer
    
    'MySQL
    MyDate = Now
    If chkCheck.Value = 1 Then
       checkPayVal = 1
       
       If txtChkPRid.Tag = "CHECK ALLOCATED" Then
            SPamtPd = GetSPttlAmtPd
            SVCamtPd = 0
            myRem = "Check#:" & txtCheckNo.Text & " " & txtChkBank.Text
        Else
            SPamtPd = GetSPttlAmtPd
            SVCamtPd = GetSRVCttlAmtPd
            myRem = RemarksOnAcct
        End If
        
        If udORCR.Value = 0 Then 'OR
            LYRcon.Execute "INSERT INTO payments(PAYDATE, " & _
                "ORNO, SPAREPARTS, SERVICE, REMARKS, CUSTOMERNAME," & _
                "EMPID, UPDATEDBY, LASTUPDATE, CUSTOMERID," & _
                "CHECKPAYMENT, ACCT_CODE)VALUES('" & Format(txtDate.Text, "yyyy-mm-dd") & "'," & _
                "'" & txtORno.Text & "', " & SPamtPd & "," & _
                "" & SVCamtPd & ",'" & Left$(myRem, 150) & "'," & _
                "'" & cboAcctCust.Text & "'," & Val(txtEmpID.Text) & "," & _
                "" & UserID & ",'" & Format(MyDate, "yyyy-mm-dd hh:mm:ss") & "'," & _
                "" & Val(txtCustomerID.Text) & "," & checkPayVal & ",'" & myAcctCode & "');"
                txtPaymentID.Text = GetFieldValue("SELECT payments.PAYMENTID From payments Where payments.ORNO = '" & txtORno.Text & "' AND payments.LASTUPDATE = '" & Format(MyDate, "yyyy-mm-dd hh:mm:ss") & "';")
        Else
            LYRcon.Execute "INSERT INTO payments(PAYDATE, " & _
                "CRNO, SPAREPARTS, SERVICE, REMARKS, CUSTOMERNAME," & _
                "EMPID, UPDATEDBY, LASTUPDATE, CUSTOMERID," & _
                "CHECKPAYMENT, ACCT_CODE)VALUES('" & Format(txtDate.Text, "yyyy-mm-dd") & "'," & _
                "'" & txtORno.Text & "', " & SPamtPd & "," & _
                "" & SVCamtPd & ",'" & Left$(myRem, 150) & "'," & _
                "'" & cboAcctCust.Text & "'," & Val(txtEmpID.Text) & "," & _
                "" & UserID & ",'" & Format(MyDate, "yyyy-mm-dd hh:mm:ss") & "'," & _
                "" & Val(txtCustomerID.Text) & "," & Val(txtChkPRid.Text) & ",'" & myAcctCode & "');"
                txtPaymentID.Text = GetFieldValue("SELECT payments.PAYMENTID From payments Where payments.CRNO = '" & txtORno.Text & "' AND payments.LASTUPDATE = '" & Format(MyDate, "yyyy-mm-dd hh:mm:ss") & "';")
        End If
        
    Else
        If udORCR.Value = 0 Then 'OR
            LYRcon.Execute "INSERT INTO payments(PAYDATE, " & _
                "ORNO, SPAREPARTS, SERVICE, REMARKS, CUSTOMERNAME," & _
                "EMPID, UPDATEDBY, LASTUPDATE," & _
                "CUSTOMERID, ACCT_CODE) VALUES('" & Format(txtDate.Text, "yyyy-mm-dd") & "'," & _
                "'" & txtORno.Text & "', " & GetSPttlAmtPd & "," & _
                "" & GetSRVCttlAmtPd & ",'" & Left$(RemarksOnAcct, 150) & "'," & _
                "'" & cboAcctCust.Text & "'," & Val(txtEmpID.Text) & "," & _
                "" & UserID & ",'" & Format(MyDate, "yyyy-mm-dd hh:mm:ss") & "'," & _
                "" & Val(txtCustomerID.Text) & ",'" & myAcctCode & "');"
                txtPaymentID.Text = GetFieldValue("SELECT payments.PAYMENTID From payments Where payments.ORNO = '" & txtORno.Text & "' AND payments.LASTUPDATE = '" & Format(MyDate, "yyyy-mm-dd hh:mm:ss") & "';")
        Else
            LYRcon.Execute "INSERT INTO payments(PAYDATE, " & _
                "CRNO, SPAREPARTS, SERVICE, REMARKS, CUSTOMERNAME," & _
                "EMPID, UPDATEDBY, LASTUPDATE," & _
                "CUSTOMERID, ACCT_CODE) VALUES('" & Format(txtDate.Text, "yyyy-mm-dd") & "'," & _
                "'" & txtORno.Text & "', " & GetSPttlAmtPd & "," & _
                "" & GetSRVCttlAmtPd & ",'" & Left$(RemarksOnAcct, 150) & "'," & _
                "'" & cboAcctCust.Text & "'," & Val(txtEmpID.Text) & "," & _
                "" & UserID & ",'" & Format(MyDate, "yyyy-mm-dd hh:mm:ss") & "'," & _
                "" & Val(txtCustomerID.Text) & ",'" & myAcctCode & "');"
            txtPaymentID.Text = GetFieldValue("SELECT payments.PAYMENTID From payments Where payments.CRNO = '" & txtORno.Text & "' AND payments.LASTUPDATE = '" & Format(MyDate, "yyyy-mm-dd hh:mm:ss") & "';")
        End If
        
    End If
    
    If txtChkPRid.Tag = "CHECK ALLOCATED" Then
        With flxBal
            For i = 1 To .Rows - 1
                If .TextMatrix(i, 1) = ThisBranch Then
                    If UCase$(Left$(.TextMatrix(i, 0), 3)) = "JO#" Then
                        MyJo = Replace$(.TextMatrix(i, 0), "JO#", "")
                        'save ledger
                        LYRcon.Execute "INSERT INTO ledger(SPINVNO, REFTYPE, REFNO, REFDATE, CREDIT, OSBAL, CUSTOMERID, SEQNO, JOBID) " & _
                        "VALUES('" & MyJo & "', 'Payment', '" & txtORno.Text & "', '" & Format(txtDate.Text, "yyyy-mm-dd") & "', " & _
                        Val(.TextMatrix(i, 3)) & ", " & Val(.TextMatrix(i, 4)) & ", " & Val(.TextMatrix(i, 6)) & ", " & MaxSeqNo(MyJo) & ", " & Val(.TextMatrix(i, 5)) & ");"
                    
                        'update onacctTerms
                        LYRcon.Execute "UPDATE ONACCTERMS SET ONACCTERMS.TTLAMTPAID = ONACCTERMS.TTLAMTPAID + " & Val(.TextMatrix(i, 3)) & ", ONACCTERMS.BALANCE = " & Val(.TextMatrix(i, 4)) & ", ONACCTERMS.CHKSPRID = " & Val(txtChkPRid.Text) & ", " & _
                        "ONACCTERMS.STATUS = '" & IIf(Val(.TextMatrix(i, 4)) = 0, "CLOSED", "OPEN") & "', ONACCTERMS.UPDATEDBY = " & UserID & ", ONACCTERMS.LASTUPDATE = '" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "' " & _
                        "WHERE ONACCTERMS.SPINVNO IS NULL AND ONACCTERMS.CUSTOMERID = " & Val(.TextMatrix(i, 6)) & " AND ONACCTERMS.JOBORDERID = " & Val(.TextMatrix(i, 5)) & ";"
                        
                        'Update Check Allocation
                        LYRcon.Execute "UPDATE checkalloc SET ORAMT = " & Val(Replace(txtPayable.Text, ",", "")) & ", ORNO = '" & txtORno.Text & "', ORDATE = '" & Format(txtDate.Text, "yyyy-mm-dd") & "' " & _
                        "Where checkalloc.CUSTOMERID = " & Val(.TextMatrix(i, 6)) & " AND checkalloc.CHKSPRID = " & Val(txtChkPRid.Text) & " AND checkalloc.JOBORDERID = " & Val(.TextMatrix(i, 5)) & ";"
                        Qry2 = "SELECT onaccterms.TRANSDATE From onaccterms Where onaccterms.JOBORDERID = " & Val(.TextMatrix(i, 5)) & " AND onaccterms.SPINVNO IS NULL;"
                        InvDate = GetFieldValue(Qry2)
                    Else
                        LYRcon.Execute "INSERT INTO ledger(SPINVNO, REFTYPE, REFNO, REFDATE, CREDIT, OSBAL, CUSTOMERID, SEQNO) " & _
                        "VALUES('" & .TextMatrix(i, 0) & "', 'Payment', '" & txtORno.Text & "', '" & Format(txtDate.Text, "yyyy-mm-dd") & "', " & _
                        Val(.TextMatrix(i, 3)) & ", " & Val(.TextMatrix(i, 4)) & ", " & Val(.TextMatrix(i, 6)) & ", " & MaxSeqNo(.TextMatrix(i, 0)) & ");"
                    
                        'update onacctTerms
                        LYRcon.Execute "UPDATE ONACCTERMS SET ONACCTERMS.TTLAMTPAID = ONACCTERMS.TTLAMTPAID + " & Val(.TextMatrix(i, 3)) & ", ONACCTERMS.BALANCE = " & Val(.TextMatrix(i, 4)) & ", ONACCTERMS.CHKSPRID = " & Val(txtChkPRid.Text) & ", " & _
                        "ONACCTERMS.STATUS = '" & IIf(Val(.TextMatrix(i, 4)) = 0, "CLOSED", "OPEN") & "', ONACCTERMS.UPDATEDBY = " & UserID & ", ONACCTERMS.LASTUPDATE = '" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "' " & _
                        "WHERE ONACCTERMS.SPINVNO='" & .TextMatrix(i, 0) & "' AND ONACCTERMS.CUSTOMERID=" & Val(.TextMatrix(i, 6)) & " AND ONACCTERMS.SPSORDERID = " & Val(.TextMatrix(i, 5)) & ";"
                                        
                        'Update Check Allocation
                        LYRcon.Execute "UPDATE checkalloc SET ORAMT = " & Val(Replace(txtPayable.Text, ",", "")) & ", ORNO = '" & txtORno.Text & "', ORDATE = '" & Format(txtDate.Text, "yyyy-mm-dd") & "' " & _
                        "Where checkalloc.CUSTOMERID = " & Val(.TextMatrix(i, 6)) & " AND checkalloc.CHKSPRID = " & Val(txtChkPRid.Text) & " AND checkalloc.SPSORDERID = " & Val(.TextMatrix(i, 5)) & ";"
                        Qry2 = "SELECT onaccterms.TRANSDATE From onaccterms Where onaccterms.SPSORDERID = " & Val(.TextMatrix(i, 5)) & " AND onaccterms.SPINVNO = '" & .TextMatrix(i, 0) & "';"
                        InvDate = GetFieldValue(Qry2)
                    
                    End If
                         
                End If
                
                If ThisBranch = "MRC2" Then
                    If .TextMatrix(i, 1) = "MRC1" Then
                             
                        Qry = "SELECT customer.CUSTOMERID From customer Where customer.CUSTOMERNAME = '" & cboAcctCust.Text & "';"
                        CustID = Val(GetFieldValueBranch(Qry, "MRC1"))
                        If CustID = 0 Then CustID = Val(txtCustomerID.Text)
                        'save ledger
                        Qry = "SELECT Max(round(LEDGER.SEQNO)) AS MaxOfSEQNO From LEDGER WHERE LEDGER.SPINVNO='" & .TextMatrix(i, 0) & "';"
                        SeqN = Val(GetFieldValueBranch(Qry, "MRC1")) + 1
                        conMRC1.Execute "INSERT INTO ledger(SPINVNO, REFTYPE, REFNO, REFDATE, CREDIT, OSBAL, CUSTOMERID, SEQNO) " & _
                        "VALUES('" & .TextMatrix(i, 0) & "', 'Payment', '" & txtORno.Text & "', '" & Format(txtDate.Text, "yyyy-mm-dd") & "', " & _
                        Val(.TextMatrix(i, 3)) & ", " & Val(.TextMatrix(i, 4)) & ", " & CustID & ", " & SeqN & ");"
                        
                        'update onacctTerms
                        conMRC1.Execute "UPDATE ONACCTERMS SET ONACCTERMS.TTLAMTPAID = ONACCTERMS.TTLAMTPAID + " & Val(.TextMatrix(i, 3)) & ", ONACCTERMS.BALANCE = " & Val(.TextMatrix(i, 4)) & ", ONACCTERMS.CHKSPRID = " & Val(txtChkPRid.Text) & ", " & _
                        "ONACCTERMS.STATUS = '" & IIf(Val(.TextMatrix(i, 4)) = 0, "CLOSED", "OPEN") & "', ONACCTERMS.UPDATEDBY = " & UserID & ", ONACCTERMS.LASTUPDATE = '" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "' " & _
                        "WHERE ONACCTERMS.SPINVNO='" & .TextMatrix(i, 0) & "' AND ONACCTERMS.CUSTOMERID=" & CustID & " AND ONACCTERMS.SPSORDERID = " & Val(.TextMatrix(i, 5)) & ";"
                        
                        'Update Check Allocation
                        conMRC1.Execute "UPDATE checkalloc SET ORAMT = " & Val(Replace(txtPayable.Text, ",", "")) & ", ORNO = '" & txtORno.Text & "', ORDATE = '" & Format(txtDate.Text, "yyyy-mm-dd") & "' " & _
                        "Where checkalloc.CUSTOMERID = " & CustID & " AND checkalloc.CHKSPRID = " & Val(txtChkPRid.Text) & " AND checkalloc.SPSORDERID = " & Val(.TextMatrix(i, 5)) & ";"
                        
                        Qry2 = "SELECT onaccterms.TRANSDATE From onaccterms Where onaccterms.SPSORDERID = " & Val(.TextMatrix(i, 5)) & " AND onaccterms.SPINVNO = '" & .TextMatrix(i, 0) & "';"
                        InvDate = GetFieldValueBranch(Qry2, "MRC1")
                        
                    ElseIf .TextMatrix(i, 1) = "VAN1" Then
                        Qry = "SELECT customer.CUSTOMERID From customer Where customer.CUSTOMERNAME = '" & cboAcctCust.Text & "';"
                        CustID = Val(GetFieldValueBranch(Qry, "VAN1"))
                        If CustID = 0 Then CustID = Val(txtCustomerID.Text)
                        'save ledger
                        Qry = "SELECT Max(round(LEDGER.SEQNO)) AS MaxOfSEQNO From LEDGER WHERE LEDGER.SPINVNO='" & .TextMatrix(i, 0) & "';"
                        SeqN = Val(GetFieldValueBranch(Qry, "VAN1")) + 1
                        conVAN1.Execute "INSERT INTO ledger(SPINVNO, REFTYPE, REFNO, REFDATE, CREDIT, OSBAL, CUSTOMERID, SEQNO) " & _
                        "VALUES('" & .TextMatrix(i, 0) & "', 'Payment', '" & txtORno.Text & "', '" & Format(txtDate.Text, "yyyy-mm-dd") & "', " & _
                        Val(.TextMatrix(i, 3)) & ", " & Val(.TextMatrix(i, 4)) & ", " & CustID & ", " & SeqN & ");"
                        
                        'update onacctTerms
                        conVAN1.Execute "UPDATE ONACCTERMS SET ONACCTERMS.TTLAMTPAID = ONACCTERMS.TTLAMTPAID + " & Val(.TextMatrix(i, 3)) & ", ONACCTERMS.BALANCE = " & Val(.TextMatrix(i, 4)) & ", ONACCTERMS.CHKSPRID = " & Val(txtChkPRid.Text) & ", " & _
                        "ONACCTERMS.STATUS = '" & IIf(Val(.TextMatrix(i, 4)) = 0, "CLOSED", "OPEN") & "', ONACCTERMS.UPDATEDBY = " & UserID & ", ONACCTERMS.LASTUPDATE = '" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "' " & _
                        "WHERE ONACCTERMS.SPINVNO='" & .TextMatrix(i, 0) & "' AND ONACCTERMS.CUSTOMERID=" & CustID & " AND ONACCTERMS.SPSORDERID = " & Val(.TextMatrix(i, 5)) & ";"
                        
                        'Update Check Allocation
                        conVAN1.Execute "UPDATE checkalloc SET ORAMT = " & Val(Replace(txtPayable.Text, ",", "")) & ", ORNO = '" & txtORno.Text & "', ORDATE = '" & Format(txtDate.Text, "yyyy-mm-dd") & "' " & _
                        "Where checkalloc.CUSTOMERID = " & CustID & " AND checkalloc.CHKSPRID = " & Val(txtChkPRid.Text) & " AND checkalloc.SPSORDERID = " & Val(.TextMatrix(i, 5)) & ";"
                        
                        Qry2 = "SELECT onaccterms.TRANSDATE From onaccterms Where onaccterms.SPSORDERID = " & Val(.TextMatrix(i, 5)) & " AND onaccterms.SPINVNO = '" & .TextMatrix(i, 0) & "';"
                        InvDate = GetFieldValueBranch(Qry2, "VAN1")
                    ElseIf .TextMatrix(i, 1) = "VAN2" Then
                        Qry = "SELECT customer.CUSTOMERID From customer Where customer.CUSTOMERNAME = '" & cboAcctCust.Text & "';"
                        CustID = Val(GetFieldValueBranch(Qry, "VAN2"))
                        If CustID = 0 Then CustID = Val(txtCustomerID.Text)
                        'save ledger
                        Qry = "SELECT Max(round(LEDGER.SEQNO)) AS MaxOfSEQNO From LEDGER WHERE LEDGER.SPINVNO='" & .TextMatrix(i, 0) & "';"
                        SeqN = Val(GetFieldValueBranch(Qry, "VAN2")) + 1
                        conVAN2.Execute "INSERT INTO ledger(SPINVNO, REFTYPE, REFNO, REFDATE, CREDIT, OSBAL, CUSTOMERID, SEQNO) " & _
                        "VALUES('" & .TextMatrix(i, 0) & "', 'Payment', '" & txtORno.Text & "', '" & Format(txtDate.Text, "yyyy-mm-dd") & "', " & _
                        Val(.TextMatrix(i, 3)) & ", " & Val(.TextMatrix(i, 4)) & ", " & CustID & ", " & SeqN & ");"
                        
                        'update onacctTerms
                        conVAN2.Execute "UPDATE ONACCTERMS SET ONACCTERMS.TTLAMTPAID = ONACCTERMS.TTLAMTPAID + " & Val(.TextMatrix(i, 3)) & ", ONACCTERMS.BALANCE = " & Val(.TextMatrix(i, 4)) & ", ONACCTERMS.CHKSPRID = " & Val(txtChkPRid.Text) & ", " & _
                        "ONACCTERMS.STATUS = '" & IIf(Val(.TextMatrix(i, 4)) = 0, "CLOSED", "OPEN") & "', ONACCTERMS.UPDATEDBY = " & UserID & ", ONACCTERMS.LASTUPDATE = '" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "' " & _
                        "WHERE ONACCTERMS.SPINVNO='" & .TextMatrix(i, 0) & "' AND ONACCTERMS.CUSTOMERID=" & CustID & " AND ONACCTERMS.SPSORDERID = " & Val(.TextMatrix(i, 5)) & ";"
                        
                        'Update Check Allocation
                        conVAN2.Execute "UPDATE checkalloc SET ORAMT = " & Val(Replace(txtPayable.Text, ",", "")) & ", ORNO = '" & txtORno.Text & "', ORDATE = '" & Format(txtDate.Text, "yyyy-mm-dd") & "' " & _
                        "Where checkalloc.CUSTOMERID = " & CustID & " AND checkalloc.CHKSPRID = " & Val(txtChkPRid.Text) & " AND checkalloc.SPSORDERID = " & Val(.TextMatrix(i, 5)) & ";"
                        
                        Qry2 = "SELECT onaccterms.TRANSDATE From onaccterms Where onaccterms.SPSORDERID = " & Val(.TextMatrix(i, 5)) & " AND onaccterms.SPINVNO = '" & .TextMatrix(i, 0) & "';"
                        InvDate = GetFieldValueBranch(Qry2, "VAN2")
                        
                    ElseIf .TextMatrix(i, 1) = "VAN3" Then
                        Qry = "SELECT customer.CUSTOMERID From customer Where customer.CUSTOMERNAME = '" & cboAcctCust.Text & "';"
                        CustID = Val(GetFieldValueBranch(Qry, "VAN3"))
                        If CustID = 0 Then CustID = Val(txtCustomerID.Text)
                        'save ledger
                        Qry = "SELECT Max(round(LEDGER.SEQNO)) AS MaxOfSEQNO From LEDGER WHERE LEDGER.SPINVNO='" & .TextMatrix(i, 0) & "';"
                        SeqN = Val(GetFieldValueBranch(Qry, "VAN3")) + 1
                        conVAN3.Execute "INSERT INTO ledger(SPINVNO, REFTYPE, REFNO, REFDATE, CREDIT, OSBAL, CUSTOMERID, SEQNO) " & _
                        "VALUES('" & .TextMatrix(i, 0) & "', 'Payment', '" & txtORno.Text & "', '" & Format(txtDate.Text, "yyyy-mm-dd") & "', " & _
                        Val(.TextMatrix(i, 3)) & ", " & Val(.TextMatrix(i, 4)) & ", " & CustID & ", " & SeqN & ");"
                        
                        'update onacctTerms
                        conVAN3.Execute "UPDATE ONACCTERMS SET ONACCTERMS.TTLAMTPAID = ONACCTERMS.TTLAMTPAID + " & Val(.TextMatrix(i, 3)) & ", ONACCTERMS.BALANCE = " & Val(.TextMatrix(i, 4)) & ", ONACCTERMS.CHKSPRID = " & Val(txtChkPRid.Text) & ", " & _
                        "ONACCTERMS.STATUS = '" & IIf(Val(.TextMatrix(i, 4)) = 0, "CLOSED", "OPEN") & "', ONACCTERMS.UPDATEDBY = " & UserID & ", ONACCTERMS.LASTUPDATE = '" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "' " & _
                        "WHERE ONACCTERMS.SPINVNO='" & .TextMatrix(i, 0) & "' AND ONACCTERMS.CUSTOMERID=" & CustID & " AND ONACCTERMS.SPSORDERID = " & Val(.TextMatrix(i, 5)) & ";"
                        
                        'Update Check Allocation
                        conVAN3.Execute "UPDATE checkalloc SET ORAMT = " & Val(Replace(txtPayable.Text, ",", "")) & ", ORNO = '" & txtORno.Text & "', ORDATE = '" & Format(txtDate.Text, "yyyy-mm-dd") & "' " & _
                        "Where checkalloc.CUSTOMERID = " & CustID & " AND checkalloc.CHKSPRID = " & Val(txtChkPRid.Text) & " AND checkalloc.SPSORDERID = " & Val(.TextMatrix(i, 5)) & ";"
                        
                        Qry2 = "SELECT onaccterms.TRANSDATE From onaccterms Where onaccterms.SPSORDERID = " & Val(.TextMatrix(i, 5)) & " AND onaccterms.SPINVNO = '" & .TextMatrix(i, 0) & "';"
                        InvDate = GetFieldValueBranch(Qry2, "VAN3")
                        
                    End If
                    
                End If
                'Update Attachment
                LYRcon.Execute "INSERT INTO ordetails(CUSTOMERID, PAYMENTID, INVOICENO, JOBORDERNO, TRANSDATE, ACCTAMT, AMTPAID, BALANCE) " & _
                "VALUES(" & Val(txtCustomerID.Text) & ", " & Val(txtPaymentID.Text) & ", '" & IIf(MyJo <> "", "", .TextMatrix(i, 0)) & "', '" & IIf(MyJo <> "", .TextMatrix(i, 0), "") & "', '" & Format(DateValue(InvDate), "yyyy-mm-dd") & "', " & Val(Replace(.TextMatrix(i, 2), ",", "")) & ", " & Val(Replace(.TextMatrix(i, 3), ",", "")) & ", " & Val(Replace(.TextMatrix(i, 4), ",", "")) & ")"
                'Update Aging
                If Left$(.TextMatrix(i, 0), 3) = "TRA" Then
                    MyRef = .TextMatrix(i, 0)
                Else
                    MyRef = "Inv# " & .TextMatrix(i, 0)
                End If
                Qry = "SELECT * From agingcurrent Where agingcurrent.REFNO = '" & MyRef & "' AND agingcurrent.CustID = " & Val(txtCustomerID.Text) & ";"
                
            Next i
        End With
        LYRcon.Execute "UPDATE CHECKS SET CHECKS.CHECKBAL = CHECKS.CHECKBAL-" & Val(Replace(txtPayable.Text, ",", "")) & " WHERE CHECKS.CHKSPRID=" & Val(txtChkPRid.Text) & ";"
        'Update Remarks
        If flxBal.Rows > 4 Then txtPaymentID.Tag = "ATTACHMENT" 'UPDATE KEVIN
    Else
        'Add New to table Ledger
        With flxAcct
            If .Rows > 2 Then
                For i = 1 To .Rows - 1
                    .TextMatrix(i, 10) = Val(Replace(flxBal.TextMatrix(i, 3), ",", ""))
                    .TextMatrix(i, 11) = Val(Replace(flxBal.TextMatrix(i, 4), ",", ""))
                Next i
            End If
            
            'mySQL
            CntAtt = 0
            Dim MyJobID As Integer
            For i = 1 To .Rows - 1
                If Val(.TextMatrix(i, 10)) > 0 Then
                    
                    If Val(.TextMatrix(i, 14)) > 0 Then
                        MyJobID = Val(.TextMatrix(i, 14))
                        LYRcon.Execute "INSERT INTO ledger(SPINVNO, REFTYPE, REFNO, REFDATE, CREDIT, " & _
                            "OSBAL, CUSTOMERID, SEQNO, JOBID)VALUES('" & .TextMatrix(i, 0) & "', 'Payment', " & _
                            "'" & txtORno.Text & "', '" & Format(txtDate.Text, "yyyy-mm-dd") & "', " & Val(.TextMatrix(i, 10)) & ", " & _
                            "" & Val(.TextMatrix(i, 11)) & ", " & txtCustomerID.Text & ", " & MaxSeqNo(.TextMatrix(i, 0)) & ", " & _
                            "" & MyJobID & ");"
                            CntAtt = CntAtt + 1
                    Else
                        LYRcon.Execute "INSERT INTO ledger(SPINVNO, REFTYPE, REFNO, REFDATE, CREDIT, " & _
                            "OSBAL, CUSTOMERID, SEQNO)VALUES('" & .TextMatrix(i, 0) & "', 'Payment', " & _
                            "'" & txtORno.Text & "', '" & Format(txtDate.Text, "yyyy-mm-dd") & "', " & Val(.TextMatrix(i, 10)) & ", " & _
                            "" & Val(.TextMatrix(i, 11)) & ", " & txtCustomerID.Text & ", " & MaxSeqNo(.TextMatrix(i, 0)) & ");"
                            CntAtt = CntAtt + 1
                    End If
                End If
            Next i
            
            'Update Table OnAcctTerms
            For i = 1 To .Rows - 1
                If Val(.TextMatrix(i, 10)) > 0 Then
                    If Val(.TextMatrix(i, 14)) > 0 Then 'Update On Account for J.O.
                        'mySQL
                        LYRcon.Execute "UPDATE ONACCTERMS SET ONACCTERMS.TTLAMTPAID = " & Val(.TextMatrix(i, 10)) + Val(.TextMatrix(i, 7)) & ", ONACCTERMS.BALANCE = " & Val(.TextMatrix(i, 11)) & ", ONACCTERMS.STATUS = '" & IIf(Val(.TextMatrix(i, 11)) = 0, "CLOSED", "OPEN") & "', ONACCTERMS.UPDATEDBY = " & UserID & ", ONACCTERMS.LASTUPDATE = '" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "' " & _
                                       "WHERE ONACCTERMS.JOBORDERID=" & Val(.TextMatrix(i, 14)) & " AND ONACCTERMS.CUSTOMERID=" & txtCustomerID.Text & " AND ONACCTERMS.REQUISITIONID=0;"
                        jobOR = GetJobOR(Val(.TextMatrix(i, 14)))
                        LYRcon.Execute "UPDATE SRVCJOBORDER SET SRVCJOBORDER.BALANCE = " & Val(.TextMatrix(i, 11)) & ", SRVCJOBORDER.PAYMENTDONE = " & IIf(Val(.TextMatrix(i, 11)) = 0, 1, 0) & ", SRVCJOBORDER.ORNO = '" & IIf(Len(Trim(jobOR)) > 1, jobOR & ", " & txtORno.Text, txtORno.Text) & "' " & _
                                       "WHERE SRVCJOBORDER.CUSTOMERID=" & txtCustomerID.Text & " AND SRVCJOBORDER.JOBORDERID=" & Val(.TextMatrix(i, 14)) & ";"
                    Else    'Update On Account for Invoice
                        'mySQL
                        LYRcon.Execute "UPDATE ONACCTERMS SET ONACCTERMS.TTLAMTPAID = " & Val(.TextMatrix(i, 10)) + Val(.TextMatrix(i, 7)) & ", ONACCTERMS.BALANCE = " & Val(.TextMatrix(i, 11)) & ", ONACCTERMS.STATUS = '" & IIf(Val(.TextMatrix(i, 11)) = 0, "CLOSED", "OPEN") & "', ONACCTERMS.UPDATEDBY = " & UserID & ", ONACCTERMS.LASTUPDATE = '" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "' " & _
                            "WHERE ONACCTERMS.SPINVNO='" & .TextMatrix(i, 0) & "' AND ONACCTERMS.CUSTOMERID=" & txtCustomerID.Text & ";"
                    End If
                End If
            Next i
            'Update Customer OS Bal
            If Val(txtCustomerID.Text) > 0 Then
                'mySQl
                LYRcon.Execute "UPDATE CUSTOMER SET CUSTOMER.OSBAL = CUSTOMER.OSBAL-" & Val(Replace(txtPayable.Text, ",", "")) & " WHERE CUSTOMER.CUSTOMERID=" & txtCustomerID.Text & ";"
            End If
            If CntAtt > 1 Then   'Apply Attachments
                'mySQL
                myRem = cboPayType.Text & " See various attachment"
                LYRcon.Execute "UPDATE PAYMENTS SET PAYMENTS.REMARKS = '" & myRem & "' WHERE PAYMENTS.PAYMENTID=" & txtPaymentID.Text & ";"
                
                
                'mySQL
                Dim MyJONo As String
                Dim MyInvoiceNo As String
                For i = 1 To .Rows - 1
                    If Val(.TextMatrix(i, 10)) > 0 Then
                        If Val(.TextMatrix(i, 14)) > 0 Then
                            MyJONo = .TextMatrix(i, 0)
                        Else
                            MyInvoiceNo = .TextMatrix(i, 0)
                        End If
                        
                        If Val(.TextMatrix(i, 7)) + Val(flxAcct.TextMatrix(i, 10)) > Val(.TextMatrix(i, 13)) Then
                             TtlPay = Val(.TextMatrix(i, 10))
                        Else
                             TtlPay = Val(.TextMatrix(i, 7)) + Val(flxAcct.TextMatrix(i, 10))
                        End If
                        
                        LYRcon.Execute "INSERT INTO ordetails(CUSTOMERID, PAYMENTID, JOBORDERNO, INVOICENO, TRANSDATE, " & _
                            "ACCTAMT, AMTPAID, TTLPAYMENT, BALANCE)VALUES(" & Val(txtCustomerID.Text) & "," & _
                            "" & Val(txtPaymentID.Text) & ", '" & MyJONo & "', '" & MyInvoiceNo & "', " & _
                            "'" & Format(DateValue(.TextMatrix(i, 12)), "yyyy-mm-dd") & "', " & Val(.TextMatrix(i, 13)) & ", " & Val(.TextMatrix(i, 10)) & ", " & _
                            "" & TtlPay & ", " & Val(.TextMatrix(i, 11)) & ");"
                    End If
                Next i
                txtPaymentID.Tag = "ATTACHMENT"
            End If
        End With
    End If
End Sub
Function GetJobOR(MyID As Double) As String
    Dim rsJobOR As New ADODB.Recordset
    Dim Qry As String
    
    '***Access
    'Qry = "SELECT SRVCJOBORDER.ORNO From SRVCJOBORDER WHERE (((SRVCJOBORDER.JOBORDERID)=" & MyId & "));"
    'mySQL
    Qry = "SELECT SRVCJOBORDER.ORNO From SRVCJOBORDER WHERE SRVCJOBORDER.JOBORDERID=" & MyID & ";"
    
    With rsJobOR
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            GetJobOR = IIf(IsNull(!ORNO), "", !ORNO)
        Else
            GetJobOR = ""
        End If
        .Close
    End With

End Function

Private Sub SaveOnAcctQS()   'Use in Charge Invoice On-Account
    Dim rsPay As New ADODB.Recordset
    Dim MyInv As String, myRem As String, i As Integer
    
    For i = 0 To lstChrgInvNo.ListCount - 1
        MyInv = MyInv & IIf(i > 0, ", ", "") & lstChrgInvNo.List(i)
    Next i
    'mySQL
    myRem = "Charge Inv.# " & MyInv & " " & cboPayType.Text
    If udORCR.Value = 0 Then
        LYRcon.Execute "INSERT INTO payments(PAYDATE, ORNO, SPAREPARTS, CUSTOMERNAME, REMARKS," & _
            "EMPID, UPDATEDBY, LASTUPDATE, ACCT_CODE)VALUES('" & Format(txtDate.Text, "yyyy-mm-dd") & "', '" & txtORno.Text & "', " & _
            "" & Val(Replace(txtPayable.Text, ",", "")) & ", '" & txtRecFrom.Text & "', '" & myRem & "', " & _
            "" & Val(txtEmpID.Text) & ", " & UserID & ", '" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "','" & myAcctCode & "');"
    Else
        LYRcon.Execute "INSERT INTO payments(PAYDATE, CRNO, SPAREPARTS, CUSTOMERNAME, REMARKS," & _
            "EMPID, UPDATEDBY, LASTUPDATE, ACCT_CODE)VALUES('" & Format(txtDate.Text, "yyyy-mm-dd") & "', '" & txtORno.Text & "', " & _
            "" & Val(Replace(txtPayable.Text, ",", "")) & ", '" & txtRecFrom.Text & "', '" & myRem & "', " & _
            "" & Val(txtEmpID.Text) & ", " & UserID & ", '" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "','" & myAcctCode & "');"
    End If
    'Update Invoice for OR#
    
    With lstChrgInvNo
        For i = 0 To .ListCount - 1
            'mySQL
            LYRcon.Execute "UPDATE SPINVOICE SET SPINVOICE.ORNO = '" & txtORno.Text & "' " & _
                "WHERE SPINVOICE.SPINVNO='" & .List(i) & "';"
        Next i
    End With
End Sub

Private Sub txtSumOf_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then SendKeys "{TAB}"
    If KeyCode = vbKeyUp Then SendKeys "+{TAB}"
End Sub
Private Sub txtSumOf_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtTIN_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then SendKeys "{TAB}"
    If KeyCode = vbKeyUp Then SendKeys "+{TAB}"
End Sub

Private Sub txtTIN_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub udORCR_Change()
    If udORCR.Value = 0 Then
        lblORno.Caption = "OFFICIAL RECEIPT NO."
    Else
        lblORno.Caption = "COLLECTION RECEIPT NO."
    End If
End Sub

Private Sub UpDRef_Change()
    If txtRefNo.Text <> "" Then Call ClearEntries
    txtRefNo.Text = ""
    Select Case PayType
        Case 1  'Down Payment
            If updRef.Value = 0 Then
                lblRefNo.Caption = "SPARE PARTS"
                txtRefNo.Visible = False
            ElseIf updRef.Value = 1 Then
                lblRefNo.Caption = "JOB ORDER"
                txtRefNo.Visible = True
            ElseIf updRef.Value = 2 Then
                lblRefNo.Caption = "SP && JO"
                txtRefNo.Visible = False
            End If
        Case 2
            updRef.Visible = True
            If updRef.Value = 0 Then    'INVOICE# FOR CHARGE INVOICE PAYING FOR FULL PAYMENT
                lblRefNo.Caption = "INVOICE NO."
            ElseIf updRef.Value = 1 Then
                lblRefNo.Caption = "JOB ORDER NO."
            End If
    End Select
End Sub


