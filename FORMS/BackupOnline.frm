VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form BackupOnline 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Backup Utilities"
   ClientHeight    =   7680
   ClientLeft      =   2985
   ClientTop       =   2400
   ClientWidth     =   11280
   FillColor       =   &H000000FF&
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "BackupOnline.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7680
   ScaleWidth      =   11280
   Begin VB.Frame Frame1 
      Height          =   6855
      Left            =   150
      TabIndex        =   2
      Top             =   750
      Width           =   11055
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshFlxHis 
         Height          =   1425
         Left            =   3240
         TabIndex        =   60
         Top             =   5280
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   2514
         _Version        =   393216
         FixedCols       =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin TabDlg.SSTab stabBk 
         Height          =   4815
         Left            =   3120
         TabIndex        =   3
         Top             =   180
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   8493
         _Version        =   393216
         Style           =   1
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "Remote Backup"
         TabPicture(0)   =   "BackupOnline.frx":058A
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "frRmBackup(1)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "frRmBackup(2)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "frRmBackup(0)"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).ControlCount=   3
         TabCaption(1)   =   "Update Databases"
         TabPicture(1)   =   "BackupOnline.frx":05A6
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "fmUpdate(1)"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "fmUpdate(2)"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "fmUpdate(3)"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).Control(3)=   "fmUpdate(0)"
         Tab(1).Control(3).Enabled=   0   'False
         Tab(1).ControlCount=   4
         Begin VB.Frame frRmBackup 
            Height          =   4095
            Index           =   0
            Left            =   120
            TabIndex        =   55
            Top             =   570
            Width           =   7575
            Begin lvButton.lvButtons_H cmdNextRm 
               Height          =   375
               Left            =   5400
               TabIndex        =   56
               Top             =   3480
               Width           =   1935
               _ExtentX        =   3413
               _ExtentY        =   661
               Caption         =   "   Next    >>"
               CapAlign        =   2
               BackStyle       =   2
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
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
               Image           =   "BackupOnline.frx":05C2
               cBack           =   -2147483633
            End
            Begin VB.Label Label10 
               Caption         =   "Online backup from branch Databases"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   2400
               TabIndex        =   58
               Top             =   1080
               Width           =   3615
            End
            Begin VB.Label Label11 
               Caption         =   "Would you like to run Backup? Clik Next Button.."
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   2340
               TabIndex        =   57
               Top             =   1650
               Width           =   4215
            End
            Begin VB.Image Image6 
               Height          =   720
               Left            =   600
               Picture         =   "BackupOnline.frx":0B5C
               Top             =   600
               Width           =   720
            End
         End
         Begin VB.Frame fmUpdate 
            Height          =   4095
            Index           =   0
            Left            =   -74880
            TabIndex        =   50
            Top             =   630
            Width           =   7575
            Begin lvButton.lvButtons_H cmdNextUpt 
               Height          =   375
               Left            =   5400
               TabIndex        =   51
               Top             =   3480
               Width           =   1935
               _ExtentX        =   3413
               _ExtentY        =   661
               Caption         =   "   Next    >>"
               CapAlign        =   2
               BackStyle       =   2
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
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
               Image           =   "BackupOnline.frx":493E6
               cBack           =   -2147483633
            End
            Begin VB.Label Label19 
               Caption         =   "Update Server and Branch Databases"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   2400
               TabIndex        =   53
               Top             =   1080
               Width           =   3615
            End
            Begin VB.Label Label18 
               Caption         =   "Would you like to run Backup?  Clik Next Button.."
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   2400
               TabIndex        =   52
               Top             =   1680
               Width           =   4575
            End
            Begin VB.Image Image8 
               Height          =   840
               Left            =   600
               Picture         =   "BackupOnline.frx":49980
               Stretch         =   -1  'True
               Top             =   600
               Width           =   840
            End
         End
         Begin VB.Frame fmUpdate 
            Caption         =   "Step-3"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   4095
            Index           =   3
            Left            =   -74880
            TabIndex        =   4
            Top             =   630
            Visible         =   0   'False
            Width           =   7575
            Begin lvButton.lvButtons_H cmdNextUpt3 
               Height          =   375
               Left            =   5280
               TabIndex        =   5
               Top             =   3480
               Width           =   1935
               _ExtentX        =   3413
               _ExtentY        =   661
               Caption         =   " Done    >>"
               CapAlign        =   2
               BackStyle       =   2
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
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
               Image           =   "BackupOnline.frx":4A382
               Enabled         =   0   'False
               cBack           =   -2147483633
            End
            Begin lvButton.lvButtons_H cmdexecute 
               Height          =   495
               Left            =   3240
               TabIndex        =   6
               Top             =   1440
               Width           =   3840
               _ExtentX        =   6773
               _ExtentY        =   873
               Caption         =   "Execute"
               CapAlign        =   2
               BackStyle       =   2
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Mode            =   0
               Value           =   0   'False
               ImgAlign        =   1
               Image           =   "BackupOnline.frx":92C1C
               cBack           =   -2147483633
            End
            Begin lvButton.lvButtons_H cmdBackUpt2 
               Height          =   375
               Left            =   3240
               TabIndex        =   7
               Top             =   3480
               Width           =   1935
               _ExtentX        =   3413
               _ExtentY        =   661
               Caption         =   "     <<   Back"
               CapAlign        =   2
               BackStyle       =   2
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
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
               Image           =   "BackupOnline.frx":931B6
               cBack           =   -2147483633
            End
            Begin VB.Label Label7 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   375
               Left            =   4365
               TabIndex        =   9
               Top             =   495
               Width           =   1575
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               Caption         =   "Update Databases from Branches"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Left            =   1920
               TabIndex        =   8
               Top             =   510
               Width           =   3045
            End
            Begin VB.Image Image5 
               Appearance      =   0  'Flat
               Height          =   720
               Left            =   960
               Picture         =   "BackupOnline.frx":93750
               Top             =   240
               Width           =   720
            End
         End
         Begin VB.Frame frRmBackup 
            Caption         =   "Step-2"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   4095
            Index           =   2
            Left            =   120
            TabIndex        =   35
            Top             =   570
            Visible         =   0   'False
            Width           =   7575
            Begin VB.ListBox lstRemoteFile 
               Height          =   2010
               Left            =   120
               TabIndex        =   39
               Top             =   960
               Width           =   1935
            End
            Begin VB.DirListBox localDir 
               Height          =   1665
               Left            =   3120
               TabIndex        =   38
               Top             =   1320
               Width           =   2175
            End
            Begin VB.DriveListBox DrvLocUpdt 
               Height          =   315
               Left            =   3120
               TabIndex        =   37
               Top             =   960
               Width           =   2175
            End
            Begin VB.FileListBox filLocalFile 
               Height          =   2040
               Left            =   5520
               TabIndex        =   36
               Top             =   960
               Width           =   1935
            End
            Begin lvButton.lvButtons_H cmdGetFile 
               Height          =   375
               Left            =   2280
               TabIndex        =   40
               Top             =   1800
               Width           =   600
               _ExtentX        =   1058
               _ExtentY        =   661
               CapAlign        =   2
               BackStyle       =   2
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Mode            =   0
               Value           =   0   'False
               ImgAlign        =   4
               Image           =   "BackupOnline.frx":DBFDA
               Enabled         =   0   'False
               cBack           =   -2147483633
            End
            Begin lvButton.lvButtons_H cmdUp 
               Height          =   375
               Left            =   2280
               TabIndex        =   41
               Top             =   960
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   661
               Caption         =   "Up"
               CapAlign        =   2
               BackStyle       =   2
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
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
               Image           =   "BackupOnline.frx":DC574
               cBack           =   -2147483633
            End
            Begin lvButton.lvButtons_H cmdRMBack 
               Height          =   375
               Left            =   3480
               TabIndex        =   42
               Top             =   3480
               Width           =   1935
               _ExtentX        =   3413
               _ExtentY        =   661
               Caption         =   "     <<   Back"
               CapAlign        =   2
               BackStyle       =   2
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
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
               Image           =   "BackupOnline.frx":DCB0E
               cBack           =   -2147483633
            End
            Begin lvButton.lvButtons_H cmdDone 
               Height          =   375
               Left            =   5520
               TabIndex        =   43
               Top             =   3480
               Width           =   1935
               _ExtentX        =   3413
               _ExtentY        =   661
               Caption         =   "Done"
               CapAlign        =   2
               BackStyle       =   2
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
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
               Image           =   "BackupOnline.frx":DD0A8
               Enabled         =   0   'False
               cBack           =   -2147483633
            End
            Begin MSComctlLib.ProgressBar PrgBar 
               Height          =   255
               Left            =   3120
               TabIndex        =   44
               Top             =   3120
               Visible         =   0   'False
               Width           =   3720
               _ExtentX        =   6562
               _ExtentY        =   450
               _Version        =   393216
               Appearance      =   1
            End
            Begin VB.Label Label8 
               Caption         =   "File Selected"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   49
               Top             =   3000
               Width           =   1935
            End
            Begin VB.Label lblpctg 
               AutoSize        =   -1  'True
               Caption         =   "100%"
               Height          =   195
               Left            =   6960
               TabIndex        =   48
               Top             =   3120
               Visible         =   0   'False
               Width           =   495
            End
            Begin VB.Image Image2 
               Appearance      =   0  'Flat
               Height          =   720
               Left            =   1080
               Picture         =   "BackupOnline.frx":125942
               Top             =   120
               Width           =   720
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Download Files from FTP server"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Left            =   1920
               TabIndex        =   47
               Top             =   480
               Width           =   2850
            End
            Begin VB.Label lblsize 
               Alignment       =   2  'Center
               Caption         =   "0 mb"
               Height          =   255
               Left            =   240
               TabIndex        =   46
               Top             =   3720
               Width           =   1755
            End
            Begin VB.Shape shpDownload 
               Height          =   675
               Left            =   120
               Top             =   3360
               Width           =   1905
            End
            Begin VB.Label lblfilename 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               BackStyle       =   0  'Transparent
               Caption         =   "......"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   120
               TabIndex        =   45
               Top             =   3480
               Width           =   1905
            End
         End
         Begin VB.Frame frRmBackup 
            Caption         =   "Step-1"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   4095
            Index           =   1
            Left            =   120
            TabIndex        =   20
            Top             =   570
            Width           =   7575
            Begin VB.ComboBox cmbBranch 
               Height          =   315
               Left            =   600
               Style           =   2  'Dropdown List
               TabIndex        =   21
               Top             =   2040
               Width           =   3375
            End
            Begin lvButton.lvButtons_H cmdConnect 
               Height          =   375
               Left            =   4920
               TabIndex        =   22
               Top             =   2040
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   661
               Caption         =   "Log-In"
               CapAlign        =   2
               BackStyle       =   2
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
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
               Image           =   "BackupOnline.frx":16E1CC
               Enabled         =   0   'False
               cBack           =   -2147483633
            End
            Begin lvButton.lvButtons_H cmdRMNext 
               Height          =   375
               Left            =   4920
               TabIndex        =   23
               Top             =   3480
               Width           =   1935
               _ExtentX        =   3413
               _ExtentY        =   661
               Caption         =   "   Next    >>"
               CapAlign        =   2
               BackStyle       =   2
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
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
               Image           =   "BackupOnline.frx":16E766
               Enabled         =   0   'False
               cBack           =   -2147483633
            End
            Begin lvButton.lvButtons_H cmdCancel 
               Height          =   375
               Left            =   2880
               TabIndex        =   54
               Top             =   3480
               Width           =   1935
               _ExtentX        =   3413
               _ExtentY        =   661
               Caption         =   "   Cancel"
               CapAlign        =   2
               BackStyle       =   2
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
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
               Image           =   "BackupOnline.frx":16ED00
               cBack           =   -2147483633
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "Select Branch"
               Height          =   195
               Left            =   600
               TabIndex        =   27
               Top             =   1680
               Width           =   1005
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               Caption         =   "Select Branch and Log-In to FTP server"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Left            =   2040
               TabIndex        =   26
               Top             =   720
               Width           =   3585
            End
            Begin VB.Image Image1 
               Appearance      =   0  'Flat
               Height          =   720
               Left            =   960
               Picture         =   "BackupOnline.frx":16F712
               Top             =   480
               Width           =   720
            End
            Begin VB.Label lblstatus 
               Alignment       =   2  'Center
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   600
               TabIndex        =   25
               Top             =   2640
               Width           =   6375
            End
            Begin VB.Label Label15 
               Caption         =   "Status"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   600
               TabIndex        =   24
               Top             =   2400
               Width           =   1575
            End
         End
         Begin VB.Frame fmUpdate 
            Caption         =   "Step-2"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   4095
            Index           =   2
            Left            =   -74880
            TabIndex        =   10
            Top             =   600
            Visible         =   0   'False
            Width           =   7575
            Begin VB.DirListBox localDir1 
               Height          =   1215
               Left            =   840
               TabIndex        =   13
               Top             =   1680
               Width           =   3135
            End
            Begin VB.DriveListBox DrvLocUpdt1 
               Height          =   315
               Left            =   840
               TabIndex        =   12
               Top             =   1320
               Width           =   3135
            End
            Begin VB.FileListBox filLocalFile1 
               Height          =   1650
               Left            =   4200
               TabIndex        =   11
               Top             =   1320
               Width           =   3015
            End
            Begin lvButton.lvButtons_H cmdNextUpt2 
               Height          =   375
               Left            =   5280
               TabIndex        =   14
               Top             =   3480
               Width           =   1935
               _ExtentX        =   3413
               _ExtentY        =   661
               Caption         =   "   Next    >>"
               CapAlign        =   2
               BackStyle       =   2
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
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
               Image           =   "BackupOnline.frx":1B7F9C
               Enabled         =   0   'False
               cBack           =   -2147483633
            End
            Begin lvButton.lvButtons_H cmdBackUpt1 
               Height          =   375
               Left            =   3240
               TabIndex        =   15
               Top             =   3480
               Width           =   1935
               _ExtentX        =   3413
               _ExtentY        =   661
               Caption         =   "     <<   Back"
               CapAlign        =   2
               BackStyle       =   2
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
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
               Image           =   "BackupOnline.frx":1B8536
               cBack           =   -2147483633
            End
            Begin VB.Label Label13 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   375
               Left            =   4365
               TabIndex        =   19
               Top             =   495
               Width           =   1575
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               Caption         =   "Select script from Local Drive"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Left            =   1920
               TabIndex        =   18
               Top             =   480
               Width           =   2670
            End
            Begin VB.Label Label14 
               Caption         =   "Local Drive"
               Height          =   255
               Left            =   840
               TabIndex        =   17
               Top             =   1080
               Width           =   1215
            End
            Begin VB.Label Label16 
               Caption         =   "File List"
               Height          =   255
               Left            =   4200
               TabIndex        =   16
               Top             =   1080
               Width           =   1215
            End
            Begin VB.Image Image4 
               Appearance      =   0  'Flat
               Height          =   720
               Left            =   960
               Picture         =   "BackupOnline.frx":1B8AD0
               Top             =   360
               Width           =   720
            End
         End
         Begin VB.Frame fmUpdate 
            Caption         =   "Step-1"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   4095
            Index           =   1
            Left            =   -74880
            TabIndex        =   28
            Top             =   600
            Visible         =   0   'False
            Width           =   7575
            Begin VB.ComboBox cmbUpdate 
               Height          =   315
               Left            =   1920
               Style           =   2  'Dropdown List
               TabIndex        =   29
               Top             =   2040
               Width           =   2895
            End
            Begin lvButton.lvButtons_H cmdNextUpt1 
               Height          =   375
               Left            =   5280
               TabIndex        =   30
               Top             =   3480
               Width           =   1935
               _ExtentX        =   3413
               _ExtentY        =   661
               Caption         =   "   Next    >>"
               CapAlign        =   2
               BackStyle       =   2
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
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
               Image           =   "BackupOnline.frx":20135A
               Enabled         =   0   'False
               cBack           =   -2147483633
            End
            Begin lvButton.lvButtons_H cmdcancel2 
               Height          =   375
               Left            =   3240
               TabIndex        =   59
               Top             =   3480
               Width           =   1935
               _ExtentX        =   3413
               _ExtentY        =   661
               Caption         =   "   Cancel"
               CapAlign        =   2
               BackStyle       =   2
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
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
               Image           =   "BackupOnline.frx":2018F4
               cBack           =   -2147483633
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Server Branch:"
               Height          =   195
               Left            =   1920
               TabIndex        =   33
               Top             =   1680
               Width           =   1080
            End
            Begin VB.Label lblbranch 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   375
               Left            =   4365
               TabIndex        =   32
               Top             =   495
               Width           =   1575
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               Caption         =   "Select Branch name from server"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Left            =   1920
               TabIndex        =   31
               Top             =   480
               Width           =   2910
            End
            Begin VB.Image Image3 
               Appearance      =   0  'Flat
               Height          =   720
               Left            =   960
               Picture         =   "BackupOnline.frx":202306
               Top             =   240
               Width           =   720
            End
         End
      End
      Begin VB.Image Image7 
         Height          =   6480
         Left            =   120
         Picture         =   "BackupOnline.frx":24AB90
         Stretch         =   -1  'True
         Top             =   240
         Width           =   2925
      End
      Begin VB.Label Label17 
         Caption         =   "Backup History"
         Height          =   255
         Left            =   3120
         TabIndex        =   34
         Top             =   5040
         Width           =   1935
      End
   End
   Begin InetCtlsObjects.Inet FTPOnline 
      Left            =   5880
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin lvButton.lvButtons_H cmdExit 
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      Caption         =   "E&xit"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
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
      Image           =   "BackupOnline.frx":25EA93
      cBack           =   -2147483633
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Online Backup"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   1260
   End
End
Attribute VB_Name = "BackupOnline"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RecievingSize, RecievingTm As Boolean
Dim Prgbr As Long
Dim rsBackup As New ADODB.Recordset

Private Sub ConnectFtp()
On Error GoTo errhand
    'Server Access type
    
    FTPOnline.AccessType = icUseDefault
    'Protocol to be used
    FTPOnline.Protocol = icFTP

    'Server session timeout
    'Remote host name
    ' FTPOnline.Execute , "CLOSE"
    FTPOnline.RemoteHost = "122.53.79.35"
    'Server port number, usually its 21
    FTPOnline.RemotePort = "21"
    'Server password
    FTPOnline.Password = "republicoflyr"
    'Server Username
    FTPOnline.UserName = "Administrator"
    
    'Add list of Backup from Branch
    'FPath$ = "Online_DB\motorulas.rar"
    
    'TPath$ = "C:\ACCTBACKUP\OnlineDB\motorulas.rar"
    
lstRemoteFile.Clear
FTPOnline.RequestTimeout = "60"

If InStr(cmbBranch.Text, "MOTO-R ULAS") > 0 Then
    FTPOnline.Execute , "DIR \MOTOR_ULAS"
ElseIf InStr(cmbBranch.Text, "YAMAHA") > 0 Then
    FTPOnline.Execute , "DIR \MOTORHUB"
ElseIf InStr(cmbBranch.Text, "DEPOT") > 0 Then
    FTPOnline.Execute , "DIR \MCPARTSDEPOT"
End If

Do While FTPOnline.StillExecuting
    DoEvents: DoEvents: DoEvents
Loop

errhand:
    If Err.Number <> 0 Then
       MsgBox (Err.Number & " " & Err.Description), vbExclamation
       cmdConnect.Enabled = True
    Else
        cmdConnect.Enabled = False
    End If

End Sub
Private Sub cboServerDb_Click()
    Dim qryTbl As String
    
    If cboServerDb.Text = "SQLMTRULAS" Then
        lblbranch.Caption = "Moto-r Ulas"
    End If
    If rsBackup.State = adStateOpen Then rsBackup.Close
    If chkDbUpdate.Value = 0 Then
        If cmdexecute.Enabled = False Then cmdexecute.Enabled = True
        mhflxtbl.Clear
    Else
        qryTbl = "SHOW TABLES FROM " & cboServerDb.Text
        rsBackup.Open qryTbl, LYRcon, adOpenStatic, adLockReadOnly
        Set mhflxtbl.Recordset = rsBackup
        mhflxtbl.ColWidth(0) = mhflxtbl.Width
    End If
End Sub
Private Sub chkAll_Click()
    If chkAll.Value = 1 Then
        mhflxtbl.Enabled = False
        cmdAdd.Enabled = False
        cmdremove.Enabled = False
    Else
        mhflxtbl.Enabled = True
        cmdAdd.Enabled = True
        cmdremove.Enabled = True
    End If
End Sub
Private Sub chkDbUpdate_Click()
    cmdexecute.Enabled = False
    If chkDbUpdate.Value = 1 Then
        
        chkDbUpdate.Caption = "Update Warehouse"
        mhflxtbl.Enabled = True
        mshflxSelTbl.Enabled = True
        cmdAdd.Enabled = True
        cmdremove.Enabled = True
        chkAll.Enabled = True
        
        DrvLocUpdt.Enabled = False
        localDir.Enabled = False
        filLocalFile.Enabled = False
        
        cboServerDb.Clear
        cboServerDb.AddItem "SQLMTRULAS"
       
    Else
        
        chkDbUpdate.Caption = "Update Server"
        'Call DisAbleCompn
        cboServerDb.Clear
        mhflxtbl.Clear
        mhflxtbl.Rows = 0
        cboServerDb.AddItem "SQLMTRULAS"
        
    End If
        
End Sub
Private Sub cmbBranch_Click()
    If cmbBranch.Text <> "" Then cmdConnect.Enabled = True
End Sub
'Private Sub cmdhistory_Click()
' If rsBackup.State = adStateOpen Then rsBackup.Close
'        rsBackup.Open "SELECT * From backuphistory WHERE backuphistory.updatedb='" & cboUpdateHistory.Text & "' Order By backuphistory.datebackup", LYRcon, adOpenStatic, adLockReadOnly
'        Set mshFlxHis.Recordset = rsBackup
'         mshFlxHis.FormatString = "<DATE    |<FILENAME        |<DATE CREATED  |^FILE SIZE            |<BRANCH  |<UPDATED"
    
'End Sub
Private Sub cmbUpdate_Click()
    If cmbUpdate.Text <> "" Then cmdNextUpt1.Enabled = True
End Sub
Private Sub cmdBackUpt1_Click()
    Call DBBackupStep(1)
End Sub
Private Sub cmdCancel_Click()
    Call RmBackupStep(0)
End Sub
Private Sub cmdcancel2_Click()
    Call DBBackupStep(0)
End Sub
Private Sub cmdNextRm_Click()
    Call RmBackupStep(1)
End Sub
Private Sub cmdNextUpt_Click()
    Call DBBackupStep(1)
End Sub

Private Sub cmdNextUpt1_Click()
    Call DBBackupStep(2)
End Sub
Private Sub cmdNextUpt2_Click()
    Dim isSQL As String
    
    If filLocalFile1.filename = "" Then MsgBox ("Select script  from File list"), vbInformation, "Script": Exit Sub
    isSQL = localDir1.Path & "\" & filLocalFile1.filename
    If InStr(UCase$(isSQL), "RAR") = 0 And InStr(UCase$(isSQL), "SQL") = 0 Then MsgBox ("Not valid SQL Backup"), vbExclamation, "Not Valid": Exit Sub
    
    Select Case cmbUpdate.ListIndex
            Case 0
                If InStr(UCase$(cmbUpdate.Text), "ULAS") > 0 Then
                    If InStr(UCase$(filLocalFile1.filename), "MTRULAS") = 0 Then
                        MsgBox ("Invalid Moto-R Ulas script"), vbExclamation, "Invalid"
                        Exit Sub
                    Else
                         Call DBBackupStep(3)
                    End If
                End If
            Case 1
                If InStr(UCase$(cmbUpdate.Text), "YAMAHA") > 0 Then
                    If InStr(UCase$(filLocalFile1.filename), "YAMAHA") = 0 Then
                        MsgBox ("Invalid Moto-R Yamaha script"), vbExclamation, "Invalid"
                        Exit Sub
                    Else
                         Call DBBackupStep(3)
                    End If
                End If
            Case 2
                If InStr(UCase$(cmbUpdate.Text), "DEPOT") > 0 Then
                    If InStr(UCase$(filLocalFile1.filename), "DEPOT") = 0 Then
                        MsgBox ("Invalid Moto-R Yamaha script"), vbExclamation, "Invalid"
                        Exit Sub
                    Else
                         Call DBBackupStep(3)
                    End If
                End If
          End Select
     End Sub

Private Sub cmdNextUpt3_Click()
            Call RmBackupStep(1)
        
End Sub
Private Sub cmdRMBack_Click()
    Call RmBackupStep(1)
End Sub
Private Sub cmdRMNext_Click()
    Call RmBackupStep(2)
    Call localDir_Change
End Sub

Private Sub cmdUp_Click()
    On Error Resume Next
    If FTPReady(True) Then
        If InStr(UCase(cmbBranch.Text), "ULAS") > 0 Then
            lstfilename = "mtrulas*.rar"
        ElseIf InStr(UCase(cmbBranch.Text), "YAMAHA") > 0 Then
            lstfilename = "yamaha*.rar"
        ElseIf InStr(UCase(cmbBranch.Text), "DEPOT") > 0 Then
            lstfilename = "mcpdepot*.rar"
        End If
        
        Do While FTPOnline.StillExecuting
            DoEvents: DoEvents: DoEvents
        Loop
        Do Until FTPReady(False)
            DoEvents: DoEvents: DoEvents: DoEvents
        Loop
        lstRemoteFile.Clear
        
        FTPOnline.Execute , "DIR " & lstfilename
    End If
End Sub

Private Sub DrvLocUpdt1_Change()
    On Error GoTo Errhandlr:
    localDir1.Path = DrvLocUpdt1.Drive

Errhandlr:
    If Not Err.Number > 0 Then Exit Sub
    MsgBox ("Error accessing local Drive"), vbCritical, "Error"
    Resume Next
End Sub

Private Sub filLocalFile1_Click()
    cmdNextUpt2.Enabled = True
End Sub

Private Sub localDir_Click()
    Call localDir_Change
End Sub

Private Sub localDir1_Change()
     filLocalFile1.Path = localDir1.Path
        If InStr(UCase$(cmbUpdate.Text), "ULAS") > 0 Then
            filLocalFile1.filename = "mtrulas*.rar"
        ElseIf InStr(UCase$(cmbUpdate.Text), "YAMAHA") > 0 Then
            filLocalFile1.filename = "yamaha*.rar"
        ElseIf InStr(UCase$(cmbUpdate.Text), "DEPOT") > 0 Then
            filLocalFile1.filename = "mcpdepot*.rar"
        End If
End Sub
Private Sub lstLoc_Click()
    If lstLoc.filename <> "" Then
        cmdNextUpt2.Enabled = True
    Else
        cmdNextUpt2.Enabled = False
    End If
End Sub

Private Sub cmdAdd_Click()
    If mhflxtbl.Rows = 0 Then Exit Sub
    mshflxSelTbl.AddItem mhflxtbl.TextMatrix(mhflxtbl.Row, 0) & vbTab & mhflxtbl.Row
    mshflxSelTbl.ColWidth(0) = mshflxSelTbl.Width
    mhflxtbl.RemoveItem mhflxtbl.Row
    If mshflxSelTbl.Rows > 0 Then
        If cmdexecute.Enabled = False Then cmdexecute.Enabled = True
    Else
         cmdexecute.Enabled = False
    End If
End Sub
Private Sub cmdConnect_Click()
   
    Call ConnectFtp
End Sub
Private Sub cmdGetFile_Click()
    If lstRemoteFile.Text = "" Then Exit Sub
    Call DownloadSQLFile
End Sub
Private Sub cmdDone_Click()
    'Dim AnsMsg As Integer
    'AnsMsg = MsgBox("Logoff FTP of " & cmbBranch & " ?", vbQuestion + vbOKCancel, "Log-Off")
    'If AnsMsg = vbOK Then
        'FTPOnline.Execute , "CLOSE"
        Call RmBackupStep(1)
        If FTPReady(False) Then
            lblstatus.Caption = "Disconnected"
            lstRemoteFile.Clear
            
        End If
    'End If
End Sub

'Private Sub cmdSendReqst_Click()
    'MsgBox ("Send Request to " & cmbBranch & " Server"), vbInformation
    'FTPOnline.Execute , "PUT " & App.Path & "\backup.bat" & " " & "\online_request\backup.bat"
'End Sub
'Private Sub cmdUp_Click()
'If the itc is ready then move up one directory and refresh the remote files list
'On Error Resume Next
'If FTPReady(True) Then
'    FTPOnline.Execute , "CDUP"
    
'    Do Until FTPReady(False)
'        DoEvents: DoEvents: DoEvents: DoEvents
'    Loop
'    lstRemoteFile.Clear
'    FTPOnline.Execute , "DIR"
'End If
'End Sub
Private Sub DrvLocUpdt_Change()
    On Error GoTo Errhandlr:
    localDir.Path = DrvLocUpdt.Drive

Errhandlr:
    If Not Err.Number > 0 Then Exit Sub
    MsgBox ("Error accessing local Drive"), vbCritical, "Error"
    Resume Next
End Sub
Private Sub Form_Load()
    Dim myDatabase As String
    
    cmbBranch.AddItem "MOTO-R ULAS"
    cmbBranch.AddItem "YAMAHA MOTORHUB"
    cmbBranch.AddItem "MC PARTS DEPOT CORP."
    
    cmbUpdate.AddItem "MOTO-R ULAS"
    cmbUpdate.AddItem "YAMAHA MOTORHUB"
    cmbUpdate.AddItem "MC PARTS DEPOT CORP."
    
    
    Call RmBackupStep(0)
    Call BackupHistory
    'cboUpdateHistory.AddItem "Main Server"
    'cboUpdateHistory.AddItem "Warehouse"
    'mshflxHis.FormatString = "<DATE    |<FILENAME       |<DATE CREATED  |^FILE SIZE                    |<BRANCH  |<UPDATED"
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If FTPOnline.StillExecuting = False Then
        'FTPOnline.Execute , "CLOSE"
    End If
End Sub
Private Sub FTPOnline_StateChanged(ByVal State As Integer)
    On Error GoTo Errhandlr
    Dim Data1, RemoteFiles
    Dim RemoteFileName As String
    'Index of FTP status constant
    Select Case State
        Case 1
            'icHostResolvingHost
            lblstatus.Caption = "Resolving Connection"
            
        Case 2
            'icHostResolved
            lblstatus.Caption = "Resolved Connection"
        Case 3
            'icConnecting
            lblstatus.Caption = "Connecting " & cmbBranch.Text & " Server"
        Case 4
            'icConnected
             lblstatus.Caption = "Connected"
        Case 5
            'icRequesting
            lblstatus.Caption = "Requesting Command"
        Case 6
            'icRequestSent
             lblstatus.Caption = "Send Command"
        Case 7
            'icReceivingResponse
             lblstatus.Caption = "Receiving Response"
        Case 8
            'icResponseReceived
            lblstatus.Caption = "Response has been received"
        Case 9
            'icDisconnecting
            lblstatus.Caption = "Disconnect to " & cmbBranch.Text & " Server"
        Case 10
            'icDisconnected
             lblstatus.Caption = "You are no longer Connected"
        Case 12  'request complete get the data
             'loop until you get all data
            Do While True
                Data1 = FTPOnline.GetChunk(4096, icString)
                If Len(Data1) = 0 Then Exit Do
                DoEvents
                RemoteFiles = RemoteFiles & Data1
            Loop
            Beep
            'If it is recieving size data tell the user the size and then exit the sub
            If RecievingSize Then
                'MsgBox "The size of file " & lstRemoteFile.Text & " is " & RemoteFiles & " bytes", vbInformation + vbOKOnly, "Size"
                
                lblsize.Caption = Format(Val(RemoteFiles) / 1024000, "###,###.#0") & " MB"
                Prgbr = Val(RemoteFiles)
                Exit Sub
            End If
            
            'Loop through, check for carriage returns to get each file name and add to listbox
            For i = 1 To Len(RemoteFiles)
                If Mid(RemoteFiles, i, 1) = Chr(13) Then
                    If Trim(RemoteFileName) <> "" Then
                        lstRemoteFile.AddItem RemoteFileName
                        RemoteFileName = ""
                    End If
                Else
                    If Mid(RemoteFiles, i, 1) <> Chr(10) Then
                        RemoteFileName = RemoteFileName & Mid(RemoteFiles, i, 1)
                    End If
                End If
            Next i
                 lblstatus.Caption = "Connected " & cmbBranch.Text & " Server"
                 Call ButtonForm(True)
Errhandlr:
    If Err.Number <> 0 Then
        MsgBox (Err.Number & " " & Err.Description), vbExclamation, "Error"
        cmdConnect.Enabled = True
        Exit Sub
    End If
End Select
End Sub
Private Sub cmdExecute_Click()
Dim BranchDB, isSQL As String
Dim myScrft As Integer, FileINfo As String
Dim i As Integer
        isSQL = localDir1.Path & "\" & filLocalFile1.filename
        
        mybats = "extract.bat"
        Open App.Path & "\" & mybats For Output As #1
        Print #1, "@echo off"
        Print #1, "echo Extract " & cmbUpdate.Text & " Databases. Please wait....."
        myScrft = InStr(UCase$(isSQL), "RAR")
        If myScrft > 0 Then
            Print #1, App.Path & "\rar e -plaborday " & isSQL & " " & localDir1.Path
        End If
           
        If InStr(UCase$(cmbUpdate.Text), "ULAS") > 0 Then
            BranchDB = "SQLMTRULAS"
        ElseIf InStr(UCase$(cmbUpdate.Text), "MOTORHUB") > 0 Then
            BranchDB = "YAMAHA"
        ElseIf InStr(UCase$(cmbUpdate.Text), "DEPOT") > 0 Then
            BranchDB = "SQLMCPDEPOT"
        End If
        
        Print #1, "echo Please wait...Updating server databases"
        If myScrft > 0 Then
            'if RAR
            Print #1, App.Path & "\" & "mysql -h192.168.2.27 -uSERVER -plaborday " & BranchDB & " < " & Left(isSQL, Len(isSQL) - 3) & "SQL"
        Else
            'if SQL
            Print #1, App.Path & "\" & "mysql -h192.168.2.27 -uSERVER -plaborday " & BranchDB & " < " & isSQL
        End If
        Print #1, "echo Data Successfully loaded"
        Print #1, "pause"
        Print #1, App.Path & "\" & "del " & mybats
        Close #1
        DoEvents
        Shell App.Path & "\" & mybats, vbNormalFocus
        
        FileINfo = Format(FileDateTime(isSQL), "hh:mm:ss ampm") & " " & Format(FileDateTime(isSQL), "mm/dd/yyyy")
        
        LYRcon.Execute "INSERT INTO backuphistory(datebackup, filename,datecreated,filesize,branch)" & _
            "VALUES('" & Format(Now, "yyyy-mm-dd") & "','" & filLocalFile1.filename & "','" & FileINfo & "',  '" & FileLen(isSQL) / 1024 & "MB" & "','" & BranchDB & "');"
        DoEvents
End Sub
Private Sub cmdExit_Click()
    Call cmdDone_Click
    Unload Me
End Sub
Private Sub localDir_Change()
    filLocalFile.Path = localDir.Path
    If InStr(UCase$(cmbBranch.Text), "ULAS") > 0 Then
        filLocalFile.filename = "mtrulas*.rar"
    ElseIf InStr(UCase$(cmbBranch.Text), "YAMAHA") > 0 Then
        filLocalFile.filename = "yamaha*.rar"
    ElseIf InStr(UCase$(cmbBranch.Text), "DEPOT") > 0 Then
        filLocalFile.filename = "mcpdepot*.rar"
    End If
End Sub

Private Sub localDir1_Click()
    Call localDir_Change
End Sub

Private Sub lstRemoteFile_Click()
    If Right(lstRemoteFile.Text, 1) = "/" Then: cmdGetFile.Enabled = False: Exit Sub
    cmdGetFile.Enabled = False
    Call GetFileSize
End Sub
Private Sub lstRemoteFile_DblClick()
Dim lstfilename As String
'If the itc is ready, check that the selected is a folder and change the directory
If FTPReady(True) Then
    If Right(lstRemoteFile.Text, 1) = "/" Then
        FTPOnline.Execute , "CD " & Chr(34) & lstRemoteFile.Text & Chr(34)
    Else
        Call imgReceiveFile_Click
        Exit Sub
    End If
    Do Until FTPReady(False)
        DoEvents: DoEvents: DoEvents: DoEvents
    Loop
    lstRemoteFile.Clear
   
    If InStr(UCase(cmbBranch.Text), "ULAS") > 0 Then
        lstfilename = "mtrulas*.rar"
    ElseIf InStr(UCase(cmbBranch.Text), "YAMAHA") > 0 Then
        lstfilename = "yamaha*.rar"
    ElseIf InStr(UCase(cmbBranch.Text), "DEPOT") > 0 Then
        lstfilename = "mcpdepot*.rar"
    End If
  
    FTPOnline.Execute , "DIR " & lstfilename
    lblstatus = "Connected"
End If
End Sub
Private Function FTPReady(ShowMessage As Boolean)
'Check the state of itc, if it is not executing return true
If FTPOnline.StillExecuting Then
    FTPReady = False
    If ShowMessage Then
        MsgBox "Please wait.  FTP is still executing", vbInformation + vbOKOnly, "Busy"
    End If
Else
    FTPReady = True
End If
End Function
Private Sub imgReceiveFile_Click()
'If the ITC is not still executing then receive the file
If FTPReady(True) = True Then
    'Check that they are not recieving a folder
    If Right(lstRemoteFile.Text, 1) = "/" Then
        MsgBox lstRemoteFile.Text & " is a folder and cannot be sent.", vbCritical + vbOKOnly, "Error"
        Exit Sub
    End If
    
    'Check that the file does not already exist on the computer, if it does exit sub
    For i = 0 To filLocalFile.ListCount
        If lstRemoteFile.Text = filLocalFile.List(i) Then
            MsgBox "File " & lstRemoteFile.Text & " already exists.", vbInformation + vbOKOnly, "Recieve"
            Exit Sub
        End If
    Next i
    
    FTPOnline.Execute , "GET " & Chr(34) & lstRemoteFile.Text & Chr(34) & " " & Chr(34) & filLocalFile.Path & "\" & lstRemoteFile.Text & Chr(34)
    Do Until FTPReady(False)
        DoEvents: DoEvents: DoEvents: DoEvents
    Loop
    'filLocalFile.Refresh
    'lblStatus = "Connected"
End If
End Sub
Sub GetFileSize()
    If FTPReady(True) Then
        RecievingSize = True
        FTPOnline.Execute , "SIZE " & Chr(34) & lstRemoteFile.Text & Chr(34)
        Do Until FTPReady(False)
                DoEvents: DoEvents: DoEvents: DoEvents
        Loop
        RecievingSize = False
        lblfilename.Caption = lstRemoteFile.Text
        cmdGetFile.Enabled = True
     End If
End Sub
Sub DownloadSQLFile()

    'If the ITC is not still executing then send the file
    If FTPReady(True) = True Then
        'Check that they are not recieving a folder
        If Right(lstRemoteFile.Text, 1) = "/" Then
            MsgBox lstRemoteFile.Text & " is a folder and cannot be sent.", vbCritical + vbOKOnly, "Error"
            Exit Sub
        End If
        
        'Check that the file does not already exist on the computer, if it does exit sub
        For i = 0 To filLocalFile.ListCount
            If lstRemoteFile.Text = filLocalFile.List(i) Then
                MsgBox "File " & lstRemoteFile.Text & " already exists.", vbInformation + vbOKOnly, "Recieve"
                Exit Sub
            End If
        Next i
        PrgBar.Visible = True
        lblpctg.Visible = True
        
        FTPOnline.Execute , "GET " & Chr(34) & lstRemoteFile.Text & Chr(34) & " " & Chr(34) & filLocalFile.Path & "\" & lstRemoteFile.Text & Chr(34)
        PrgBar.Max = Prgbr
        PrgBar.Value = 0
        
        Do Until FTPReady(False) And VerifyFileDownload = PrgBar.Max
            DoEvents: DoEvents: DoEvents: DoEvents
            PrgBar.Value = VerifyFileDownload
            PrcntRec = PrgBar.Value / PrgBar.Max
            lblpctg.Caption = Format$(PrcntRec, "##0.0%")
        Loop
        
        filLocalFile.Refresh
        'Call ReadFileInfo
        If FTPReady(False) Then
            MsgBox ("Finished downloading file"), vbInformation, "Finished"
            PrgBar.Visible = False
            lblpctg.Visible = False
            cmdDone.Enabled = True
        End If
    End If
End Sub
'Sub ReadFileInfo()
'    LocPath = localDir.Path & "\" & lblfilename
'    If Dir(LocPath) <> "" Then
        'lbldate.Caption = Format(FileDateTime(LocPath), "hh:mm:ss ampm") & " " & Format(FileDateTime(LocPath), "mm/dd/yyyy")
        'lblfilename.Caption = lblfname.Caption
'    End If
'End Sub
Function VerifyFileDownload() As Long
    LocPath = localDir.Path & "\" & lblfilename
    If Dir(LocPath) <> "" Then
        VerifyFileDownload = FileLen(LocPath)
    End If
End Function
'Sub LocalInfor()
'    locfile = localDir.Path & "\" & filLocalFile.filename
'    lbldate.Caption = Format(FileDateTime(locfile), "hh:mm:ss ampm") & " " & Format(FileDateTime(locfile), "mm/dd/yyyy")
'    lblfname.Caption = filLocalFile.filename
'    lblfsize.Caption = Format(FileLen(locfile) / 1024000, "###,###.#0") & " mb"
'End Sub
Public Function StringExistsInFile(ByVal TheString As String, ByVal TheFile As String) As Boolean
  Dim L As Long, S As String, FileNum As Integer
  FileNum = FreeFile
  Open TheFile For Binary Access Read Shared As #FileNum
  L = LOF(FileNum)
  S = Space$(L)
  Get #1, , S
  Close #FileNum
  If InStr(1, S, TheString) Then
    StringExistsInFile = True
  End If
End Function

Sub ButtonForm(EnDisTF As Boolean)
    
    'cmdSendReqst.Enabled = EnDisTF
    cmdRMNext.Enabled = EnDisTF
End Sub
Sub DBBackupStep(indx As Integer)
    fmUpdate(0).Visible = False
    fmUpdate(1).Visible = False
    fmUpdate(2).Visible = False
    fmUpdate(3).Visible = False
    fmUpdate(indx).Visible = True
    If InStr(UCase$(cmbUpdate.Text), "ULAS") > 0 Then
        filLocalFile1.filename = "mtrulas*.rar"
    ElseIf InStr(UCase$(cmbUpdate.Text), "YAMAHA") > 0 Then
        filLocalFile1.filename = "yamaha*.rar"
    ElseIf InStr(UCase$(cmbUpdate.Text), "DEPOT") > 0 Then
        filLocalFile1.filename = "mcpdepot*.rar"
    End If
     
End Sub
Sub RmBackupStep(indx As Integer)
    frRmBackup(0).Visible = False
    frRmBackup(1).Visible = False
    frRmBackup(2).Visible = False
    
    frRmBackup(indx).Visible = True
    
End Sub
Private Sub cmdBackUpt2_Click()
    Call DBBackupStep(2)
End Sub
Private Sub stabBk_Click(PreviousTab As Integer)
    If stabBk.Tab = 1 Then
        Call DBBackupStep(0)
    End If
End Sub
Sub BackupHistory()
Dim Qry As String
    If rsBackup.State = adStateOpen Then rsBackup.Close
    Qry = "backuphistory"
    rsBackup.Open Qry, LYRcon, adOpenKeyset, adLockReadOnly
    
    Set mshFlxHis.Recordset = rsBackup
    mshFlxHis.FormatString = "<DATE           |<FILENAME                 |<DATE CREATED            |^FILE SIZE     |<BRANCH                 "
  
End Sub

