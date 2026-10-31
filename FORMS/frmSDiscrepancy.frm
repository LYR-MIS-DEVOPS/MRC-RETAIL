VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmSDiscrepancy 
   Caption         =   "Processing of Discrepancy"
   ClientHeight    =   8745
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13785
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8745
   ScaleWidth      =   13785
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   1500
      Left            =   120
      TabIndex        =   18
      Tag             =   "XY"
      Top             =   720
      Width           =   13575
      Begin VB.CheckBox chkFinal 
         Caption         =   "FINAL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5640
         TabIndex        =   40
         Tag             =   "XY"
         Top             =   960
         Width           =   975
      End
      Begin MSMask.MaskEdBox txtDateRef 
         Height          =   360
         Left            =   11640
         TabIndex        =   12
         Tag             =   "XY"
         Top             =   600
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   635
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtDateRec 
         Height          =   360
         Left            =   11640
         TabIndex        =   11
         Tag             =   "XY"
         Top             =   960
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   635
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtSTFDate 
         Height          =   360
         Left            =   3840
         TabIndex        =   9
         Tag             =   "XY"
         Top             =   960
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   635
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtRRNo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   8640
         TabIndex        =   10
         Tag             =   "XY"
         Top             =   960
         Width           =   1815
      End
      Begin VB.TextBox txtRefNo 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   11640
         TabIndex        =   22
         Tag             =   "XY"
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox txtSTFNo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   960
         TabIndex        =   8
         Tag             =   "XY"
         Top             =   960
         Width           =   1695
      End
      Begin VB.ComboBox cmbBranch 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   960
         TabIndex        =   7
         Tag             =   "XY"
         Top             =   480
         Width           =   5295
      End
      Begin VB.Label lblRRNo 
         Alignment       =   1  'Right Justify
         Caption         =   "RR #:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   8040
         TabIndex        =   26
         Tag             =   "XY"
         Top             =   960
         Width           =   495
      End
      Begin VB.Label lblDateRec 
         Alignment       =   1  'Right Justify
         Caption         =   "DATE REC:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   10440
         TabIndex        =   25
         Tag             =   "XY"
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label lblDateRef 
         Alignment       =   1  'Right Justify
         Caption         =   "DATE REF:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   10440
         TabIndex        =   24
         Tag             =   "XY"
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label lblRefNo 
         Alignment       =   1  'Right Justify
         Caption         =   "REF. #:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   10440
         TabIndex        =   23
         Tag             =   "XY"
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lblSTFDate 
         Alignment       =   1  'Right Justify
         Caption         =   "STF DATE:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2640
         TabIndex        =   21
         Tag             =   "XY"
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label lblSTFNo 
         Caption         =   "STF #:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Tag             =   "XY"
         Top             =   960
         Width           =   975
      End
      Begin VB.Label lblBranch 
         Caption         =   "SOURCE:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Tag             =   "XY"
         Top             =   480
         Width           =   975
      End
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   3
      Left            =   4620
      TabIndex        =   3
      Top             =   0
      Width           =   1620
      _ExtentX        =   2858
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
      Image           =   "frmSDiscrepancy.frx":0000
      cBack           =   -2147483633
   End
   Begin VB.TextBox txtSDiscID 
      Height          =   285
      Left            =   13920
      TabIndex        =   37
      Top             =   360
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Frame Frame2 
      Height          =   4695
      Left            =   120
      TabIndex        =   27
      Tag             =   "XY"
      Top             =   2160
      Width           =   13575
      Begin MSFlexGridLib.MSFlexGrid flxSDiscrepancy 
         Height          =   4335
         Left            =   120
         TabIndex        =   34
         Tag             =   "XY"
         Top             =   240
         Width           =   13335
         _ExtentX        =   23521
         _ExtentY        =   7646
         _Version        =   393216
         Cols            =   7
         FormatString    =   "|<PARTNUMBER |<DESCRIPTION |^QTY |<REMARKS |<RECOMMENDATION"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame Frame3 
      Height          =   735
      Left            =   120
      TabIndex        =   28
      Tag             =   "XY"
      Top             =   6840
      Width           =   13575
      Begin VB.TextBox txtRemarks 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4440
         MaxLength       =   52
         TabIndex        =   15
         Tag             =   "XY"
         Top             =   240
         Width           =   8895
      End
      Begin VB.TextBox txtStatus 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1200
         TabIndex        =   14
         Tag             =   "XY"
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label lblRemarks 
         Alignment       =   1  'Right Justify
         Caption         =   "REMARKS:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3240
         TabIndex        =   30
         Tag             =   "XY"
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lblStatus 
         Alignment       =   2  'Center
         Caption         =   "STATUS:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   29
         Tag             =   "XY"
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.TextBox txtRecom 
      Height          =   285
      Left            =   13920
      MaxLength       =   50
      TabIndex        =   13
      Top             =   1680
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtHoldWhiteBackColor 
      Height          =   375
      Left            =   13920
      TabIndex        =   36
      Text            =   "HoldBackColor"
      Top             =   720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtHoldGrayBackColor 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   13920
      Locked          =   -1  'True
      TabIndex        =   35
      Top             =   1200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame4 
      Height          =   1095
      Left            =   120
      TabIndex        =   31
      Tag             =   "XY"
      Top             =   7560
      Width           =   13575
      Begin VB.TextBox txtApprvBy 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   8400
         MaxLength       =   50
         TabIndex        =   17
         Tag             =   "XY"
         Top             =   360
         Width           =   5055
      End
      Begin VB.TextBox txtPrepBy 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   240
         TabIndex        =   16
         Tag             =   "XY"
         Top             =   360
         Width           =   5055
      End
      Begin VB.Label lblApprvBy 
         Caption         =   "APPROVED BY"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   10320
         TabIndex        =   33
         Tag             =   "XY"
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label lblPrepBy 
         Caption         =   "PREPARED BY"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1920
         TabIndex        =   32
         Tag             =   "XY"
         Top             =   720
         Width           =   1695
      End
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   0
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
      Image           =   "frmSDiscrepancy.frx":059A
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   2
      Left            =   3120
      TabIndex        =   2
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
      Image           =   "frmSDiscrepancy.frx":0B34
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   4
      Left            =   6240
      TabIndex        =   4
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
      Image           =   "frmSDiscrepancy.frx":10CE
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   5
      Left            =   7740
      TabIndex        =   5
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
      Image           =   "frmSDiscrepancy.frx":1668
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   6
      Left            =   9240
      TabIndex        =   6
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
      Image           =   "frmSDiscrepancy.frx":1C02
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   1
      Left            =   1620
      TabIndex        =   1
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
      Image           =   "frmSDiscrepancy.frx":219C
      cBack           =   -2147483633
   End
   Begin VB.Label lblHeading 
      BackStyle       =   0  'Transparent
      Caption         =   "SETTLEMENT OF DISCREPANCY"
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
      Left            =   120
      TabIndex        =   38
      Top             =   450
      Width           =   3450
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   14280
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   14025
      m_MDI_Height    =   9315
   End
   Begin VB.Label lblHeading 
      BackStyle       =   0  'Transparent
      Caption         =   "SETTLEMENT OF DISCREPANCY"
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
      Left            =   150
      TabIndex        =   39
      Top             =   450
      Width           =   3450
   End
   Begin VB.Menu mnuStatus 
      Caption         =   "PopupMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuAdd 
         Caption         =   "&Add Row"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "&Delete Row"
      End
   End
End
Attribute VB_Name = "frmSDiscrepancy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Ans As String
Dim Qry As String
Dim arrBRANCH() As String
Dim i As Long, j As Long
Dim ActRow As Integer, ActCol As Integer
Dim tmpRs  As New ADODB.Recordset
Dim mainRs As New ADODB.Recordset
Dim StatSrch As Boolean
Dim strPartial As String, strTotal As String
Dim formAction As String
Dim PrtBln As Boolean, PrtSrch As Boolean, ProperExit As Boolean
Dim chkStock As String, ChkRec As String, branchCode As String

Public Sub SaveOnline()
    Dim rsSDisc As New ADODB.Recordset, myLastUpdate As String
    Dim timeOut As String, ErrMsg As String
 '----------- SETTLEMENT OF DISCREPANCY REQUEST SAVE ONLINE
    
    On Error GoTo LocalErr
    ErrMsg = "Saving failed while processing online"
    
    Call ConnectOnline
    
    myLastUpdate = Format(Now, "yyyy-mm-dd hh:mm:ss")
    Qry = "SELECT * From sdiscrepancy Where sdiscrepancy.REFNO = '" & txtRefNo.Text & "';"
    rsSDisc.Open Qry, OnlineCon, adOpenKeyset, adLockOptimistic
    If rsSDisc.RecordCount >= 1 Then
        rsSDisc.Close
        Exit Sub
    End If
    rsSDisc.Close
 
    OnlineCon.Execute "Insert Into sdiscrepancy(SDISCID , RefNo, DATEREF, Source, STFNO, STFDATE, RRNO, DATEREC, STATUS, REMARKS, " & _
        "PREPBY,APPRVBY, UPDATEDBY, LASTUPDATE, BCODE)" & _
        "VALUES('" & Val(txtSDiscID.Text) & "', '" & txtRefNo.Text & "', '" & Format(txtDateRef.Text, "YYYY-MM-DD") & "', " & _
        "'" & cmbBranch.Text & "', '" & txtSTFNo.Text & "', '" & Format(txtSTFDate.Text, "yyyy-mm-dd") & "', " & _
        "'" & txtRRno.Text & "', '" & Format(txtDateRec.Text, "yyyy-mm-dd") & "', '" & "NEW" & "', '" & txtRemarks.Text & "', " & _
        "" & UserID & ", '" & txtApprvBy.Text & "', " & UserID & ", '" & myLastUpdate & "', '" & sysBranchCode & "');"
    
    'Creating Batch Inserting
    OnlineCon.Execute "Drop Temporary Table IF  Exists sp_mrcretails.TMP_SDISCREPANCYDTLonline;"

    DoEvents
    OnlineCon.Execute " Create Temporary Table IF Not Exists sp_mrcretails.TMP_SDISCREPANCYDTLonline( " & _
        "   SDISCID             INT(11),        " & _
        "   QTY                 INT(11),        " & _
        "   REMARKS             Varchar(40),    " & _
        "   RECOMMENDATION      Varchar(100),   " & _
        "   SPAREPARTID         INT(11),        " & _
        "   SEQNO               INT(11)         " & _
        ");"

    Qry = "Insert Into sp_mrcretails.TMP_SDISCREPANCYDTLonline(SDISCID,QTY,REMARKS,RECOMMENDATION,SPAREPARTID,SEQNO) VALUES "
        
    For i = 1 To flxSDiscrepancy.Rows - 1
        If i > 1 Then Qry = Qry & ","
        Qry = Qry & "( '" & Val(txtSDiscID.Text) & "','" & flxSDiscrepancy.TextMatrix(i, 3) & "', '" & flxSDiscrepancy.TextMatrix(i, 4) & "', '" & flxSDiscrepancy.TextMatrix(i, 5) & "', '" & flxSDiscrepancy.TextMatrix(i, 6) & "', " & i & ")"
    Next
    Call OnlineCon.Execute(Qry)
    'End of Creating Batch Inserting
    
    OnlineCon.Execute "Insert Into sdiscrepancydtl( SDISCID, QTY, REMARKS, BRECOMMENDATION, SPAREPARTID, SEQNO)" & _
                "SELECT SDISCID,QTY,REMARKS,RECOMMENDATION,SPAREPARTID,SEQNO FROM sp_mrcretails.TMP_SDISCREPANCYDTLonline"

    OnlineCon.Close
    
    MsgBox "Settlement of Discrepancy Request successfully saved online!", vbInformation + vbOKOnly, "Save"
        
   Exit Sub
LocalErr:
    MsgBox Err.Description
End Sub

Public Sub CreateTempTable()
On Error GoTo er
    LYRcon.Execute "Drop Temporary Table IF  Exists sp_mrcretails.TMP_SDISCREPANCYDTL;"

    DoEvents
    LYRcon.Execute " Create Temporary Table IF Not Exists sp_mrcretails.TMP_SDISCREPANCYDTL( " & _
        "   QTY                 INT(11),        " & _
        "   REMARKS             Varchar(40),    " & _
        "   RECOMMENDATION      Varchar(100),   " & _
        "   SPAREPARTID         INT(11),        " & _
        "   SEQNO               INT(11)         " & _
        ");"

        Qry = "Insert Into sp_mrcretails.TMP_SDISCREPANCYDTL(QTY,REMARKS,RECOMMENDATION,SPAREPARTID,SEQNO) VALUES "
        
        For i = 1 To flxSDiscrepancy.Rows - 1
            If i > 1 Then Qry = Qry & ","
            Qry = Qry & "('" & flxSDiscrepancy.TextMatrix(i, 3) & "', '" & flxSDiscrepancy.TextMatrix(i, 4) & "', '" & flxSDiscrepancy.TextMatrix(i, 5) & "', '" & flxSDiscrepancy.TextMatrix(i, 6) & "', " & i & ")"
        Next
        Call LYRcon.Execute(Qry)
Exit Sub
er:
 Call modObjTools.ErrorMessage(Err.Description, "Error->CreateTempTable()")
End Sub

Function ValidateEntry() As Boolean
  ValidateEntry = False
    
    If Trim(cmbBranch.Text) = "" Then
        MsgBox "BRANCH IS REQUIRED!", vbExclamation, "Validation"
        cmbBranch.SetFocus
        Exit Function
    End If
    
    If Trim(txtSTFNo.Text) = "" Then
        MsgBox "STF # IS REQUIRED!", vbExclamation, "Validation"
        txtSTFNo.SetFocus
        Exit Function
    End If
        
    If Trim(txtSTFDate.Text) = "__/__/____" Then
        MsgBox "STF DATE IS REQUIRED!", vbExclamation, "Validation"
        txtSTFDate.SetFocus
        Exit Function
    End If
    
    If Trim(txtRRno.Text) = "" Then
        MsgBox "RECEIVING RECEIPT # IS REQUIRED!", vbExclamation, "Validation"
        txtRRno.SetFocus
        Exit Function
    End If
    
    If Trim(txtDateRec.Text) = "__/__/____" Then
        MsgBox "DATE RECEIVE IS REQUIRED!", vbExclamation, "Validation"
        txtDateRec.SetFocus
        Exit Function
    End If
    
    If Trim(txtPrepby.Text) = "" Then
        MsgBox "PREPARED BY IS REQUIRED!", vbExclamation, "Validation"
        txtPrepby.SetFocus
        Exit Function
    End If
    
    If Trim(txtApprvBy.Text) = "" Then
        MsgBox "APPROVED BY IS REQUIRED!", vbExclamation, "Validation"
        txtApprvBy.SetFocus
        Exit Function
    End If
    
    ValidateEntry = True

End Function

Private Sub cmbBranch_Change()
On Error GoTo er
    If StatSrch Then
        StatSrch = False
        Exit Sub
    End If

    With cmbBranch
        strPartial = .Text
        i = SendMessage(.hwnd, CB_FINDSTRING, -1, ByVal strPartial)  'If match found, append unmatched characters
        If i <> -1 Then
            strTotal = .List(i)
        
            j = Len(strTotal) - Len(strPartial)
           
            If j <> 0 Then
                StatSrch = True
                .SelText = Right$(strTotal, j)
                .SelStart = Len(strPartial)
                .SelLength = j
            End If
        End If
    End With
    
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->cboBranch-Change()")
End Sub

Private Sub cmbBranch_Click()
    branchCode = arrBRANCH(cmbBranch.ListIndex, 0)
End Sub

Private Sub cmbBranch_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then StatSrch = True
End Sub

Private Sub cmbBranch_KeyPress(KeyAscii As Integer)
Dim cbolistIndex As Integer

    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(cmbBranch, KeyAscii, ValidChars)
    If KeyAscii = 13 Then
        cbolistIndex = modObjTools.getListIndex(cmbBranch, cmbBranch.Text)
        
        If cbolistIndex < 0 Then
            cmbBranch.SetFocus
        Else
            cmbBranch.ListIndex = cbolistIndex
            Call cmbBranch_Click
            txtRRno.SetFocus
        End If
    
        
    End If
End Sub

Private Sub cmdTrans_Click(Index As Integer)
Dim rsSDisc As New ADODB.Recordset
On Error GoTo er

    If Index = 0 Then 'add
        Call clearform
        Call CboDisplay
        Call LockData(True)
        Call SetButtonVisibility(True)
        chkFinal.Enabled = True
        chkFinal.Value = 0
        formAction = "add"
    ElseIf Index = 1 Then 'edit
        If txtStatus.Text = "FINAL" Then
            MsgBox "This Transaction is already Final", vbExclamation + vbOKOnly, "Edit"
            Exit Sub
        End If
        Call SetButtonVisibility(True)
        Call LockData(True)
        cmbBranch.BackColor = txtHoldGrayBackColor.BackColor
        cmbBranch.Locked = True
        If txtStatus.Text = "PENDING" Then
            chkFinal.Enabled = True
            chkFinal.Value = 0
        End If
        formAction = "edit"
    ElseIf Index = 2 Then 'view summary
        FrmSummary.mySrc = "SETTLEMENT DISCREPANCY"
        FrmSummary.Show
        Unload Me
    ElseIf Index = 3 Then
        If txtRefNo.Text = "<AUTO>" Then
            MsgBox "No Records Found", vbExclamation + vbOKOnly, "PRINT"
            Exit Sub
        End If
        
        Qry = "call sp_mrcretails.sp_SDiscrepancy_getDetails('" & sysBranchCode & "'," & Val(txtSDiscID.Text) & ")"
        frmPrint.QryPrn = Qry
        
        frmPrint.ParamValue1 = SysBranch
        frmPrint.ParamValue2 = Format(cmbBranch.Text, "")
        frmPrint.ParamValue3 = Format(txtSTFNo.Text, "")
        frmPrint.ParamValue4 = Format(txtSTFDate.Text, "mm-dd-yyyy")
        frmPrint.ParamValue5 = Format(txtRefNo.Text, "")
        frmPrint.ParamValue6 = Format(txtDateRef.Text, "mm-dd-yyyy")
        frmPrint.ParamValue7 = Format(txtDateRec.Text, "mm-dd-yyyy")
        frmPrint.ParamValue8 = Format(txtRRno.Text, "")
        frmPrint.ParamValue9 = Format(txtStatus.Text, "")
        frmPrint.ParamValue10 = Format(txtRemarks.Text, "")
        frmPrint.ParamValue11 = Format(txtPrepby.Text, "")
        frmPrint.ParamValue12 = Format(txtApprvBy.Text, "")
        frmPrint.ParamValue13 = Format(SysAddress, "")
        frmPrint.MySource = "SETTLEMENT DISCREPANCY"
        frmPrint.Show
    ElseIf Index = 4 Then 'exit
        formAction = "EXIT"
        ProperExit = True
        Call clearform
        formAction = ""
        Unload Me
    ElseIf Index = 5 Then 'save
        If ValidateEntry() = False Then Exit Sub

        Ans = MsgBox("Do you want to save?", vbQuestion + vbYesNo, "Save")
        
        If Ans = vbNo Then Exit Sub
        
        If formAction = "add" Then
            Call CreateTempTable
            Call ConnectOnline
            If modConnection.TestConnection() = False Then Exit Sub
            DoEvents
            Qry = "call sp_mrcretails.sp_SDiscrepancy_Create('" & sysBranchCode & "', '" & cmbBranch.Text & "', '" & txtSTFNo.Text & "', '" & Format(txtSTFDate.Text, "YYYY-mm-dd") & "', '" & txtRRno.Text & "', '" & Format(txtDateRec.Text, "YYYY-mm-dd") & "', '" & Format(txtDateRef.Text, "YYYY-mm-dd") & "', '" & txtRemarks.Text & "', " & UserID & ", '" & txtApprvBy.Text & "', " & UserID & ", " & Val(chkFinal.Value) & ")"
            Set rsSDisc = LYRcon.Execute(Qry)
            If rsSDisc.RecordCount > 0 Then
                txtRefNo.Text = rsSDisc!RefNo
                txtStatus.Text = rsSDisc!STATUS
                txtRemarks.Text = IIf(IsNull(rsSDisc!Remarks), "", rsSDisc!Remarks)
                txtSDiscID.Text = rsSDisc!SDISCID
            End If
            If Val(chkFinal.Value) = 1 Then
                Call SaveOnline
            End If
        Else
            Call CreateTempTable
            Call ConnectOnline
            If modConnection.TestConnection() = False Then Exit Sub
            DoEvents
            Qry = "call sp_mrcretails.sp_SDiscrepancy_Update('" & sysBranchCode & "', '" & txtSTFNo.Text & "', '" & Format(txtSTFDate.Text, "YYYY-mm-dd") & "', '" & txtRRno.Text & "', '" & Format(txtDateRec.Text, "YYYY-mm-dd") & "', '" & txtRemarks.Text & "', '" & txtApprvBy.Text & "', " & UserID & ", " & txtSDiscID.Text & ", " & Val(chkFinal.Value) & ")"
            Set rsSDisc = LYRcon.Execute(Qry)
            If rsSDisc.RecordCount > 0 Then
                txtStatus.Text = rsSDisc!STATUS
            End If
            If Val(chkFinal.Value) = 1 Then
                Call SaveOnline
            End If
            
        End If
        Call SetButtonVisibility(False)
        Call LockData(False)
        chkFinal.Enabled = False
        formAction = ""
    
    ElseIf Index = 6 Then 'cancel
        Ans = MsgBox("Do you want to cancel?", vbQuestion + vbYesNo)
        If Ans = vbNo Then Exit Sub
        Call SetButtonVisibility(False)
        Call LockData(False)
        chkFinal.Enabled = False
        formAction = ""
    End If
    
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->cmdTrans-Click()")
End Sub

Private Sub flxSDiscrepancy_Click()
    ActRow = flxSDiscrepancy.Row
    ActCol = flxSDiscrepancy.Col
    
    'FlxParts Columns: 1-PartNumber, 2-Desc, 3-Qty, 4-Price, 5-Total, 6-ID, 7-Remarks
    If ActCol = 0 Or ActRow = 0 Then Exit Sub
    Select Case ActCol
'        Case 1  'PartNumber
'            If txtSTFNo.Text <> "" And txtRRno.Text <> "" Then
'                With txtPartNo
'                    If flxSDiscrepancy.TextMatrix(ActRow, 1) <> "" Then .Tag = flxSDiscrepancy.TextMatrix(ActRow, 1)
'                    .Top = Frame2.Top + flxSDiscrepancy.CellTop + flxSDiscrepancy.Top
'                    .Left = Frame2.Left + flxSDiscrepancy.CellLeft + flxSDiscrepancy.Left
'                    .Height = flxSDiscrepancy.CellHeight
'                    .Width = flxSDiscrepancy.CellWidth
'                    .Text = flxSDiscrepancy.Text
'                    .Visible = True
'                    .SelStart = 0
'                    .SelLength = Len(.Text)
'                    .ZOrder
'                    .SetFocus
'                End With
'            Else
'                Exit Sub
'            End If
'        Case 3 'Qty.
'            If flxSDiscrepancy.TextMatrix(flxSDiscrepancy.Row, 1) <> "" Then
'                With txtQty
'                    If flxSDiscrepancy.TextMatrix(ActRow, 3) <> "" Then .Tag = flxSDiscrepancy.TextMatrix(ActRow, 3)
'                    .Top = Frame2.Top + flxSDiscrepancy.CellTop + flxSDiscrepancy.Top
'                    .Left = Frame2.Left + flxSDiscrepancy.CellLeft + flxSDiscrepancy.Left
'                    .Height = flxSDiscrepancy.CellHeight
'                    .Width = flxSDiscrepancy.CellWidth
'                    .Text = flxSDiscrepancy.Text
'                    .Visible = True
'                    .SelStart = 0
'                    .SelLength = Len(.Text)
'                    .ZOrder
'                    .SetFocus
'                End With
'            Else
'                Exit Sub
'            End If
'        Case 4  'REMARKS
'            If flxSDiscrepancy.TextMatrix(flxSDiscrepancy.Row, 3) <> "" Then
'                With txtRem
'                    If flxSDiscrepancy.TextMatrix(ActRow, 4) <> "" Then .Tag = flxSDiscrepancy.TextMatrix(ActRow, 4)
'                    .Top = Frame2.Top + flxSDiscrepancy.CellTop + flxSDiscrepancy.Top
'                    .Left = Frame2.Left + flxSDiscrepancy.CellLeft + flxSDiscrepancy.Left
'                    .Height = flxSDiscrepancy.CellHeight
'                    .Width = flxSDiscrepancy.CellWidth
'                    .Text = flxSDiscrepancy.Text
'                    .Visible = True
'                    .SelStart = 0
'                    .SelLength = Len(.Text)
'                    .ZOrder
'                    .SetFocus
'                End With
'            Else
'                Exit Sub
'            End If
        Case 5  'RECOMMENDATION
            If formAction = "add" Or formAction = "edit" Then
                With txtRecom
                    If flxSDiscrepancy.TextMatrix(ActRow, 5) <> "" Then .Tag = flxSDiscrepancy.TextMatrix(ActRow, 5)
                    .Top = Frame2.Top + flxSDiscrepancy.CellTop + flxSDiscrepancy.Top
                    .Left = Frame2.Left + flxSDiscrepancy.CellLeft + flxSDiscrepancy.Left
                    .Height = flxSDiscrepancy.CellHeight
                    .Width = flxSDiscrepancy.CellWidth
                    .Text = flxSDiscrepancy.Text
                    .Visible = True
                    .SelStart = 0
                    .SelLength = Len(.Text)
                    .ZOrder
                    .SetFocus
                End With
            Else
                Exit Sub
            End If
       End Select
End Sub

Private Sub flxSDiscrepancy_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Call flxSDiscrepancy_Click
End Sub

Private Sub Form_Load()
    Call LockData(False)
    txtDateRef.Text = Format$(Now, "mm/dd/yyyy")
    txtPrepby.Text = UserName
    txtRefNo.Text = "<AUTO>"
    chkFinal.Enabled = False
    
    '===================BUTTON POSITION===================
    If cmdTrans(0).Visible = False Then
        'Align Save Button
        cmdTrans(5).Left = cmdTrans(0).Left
        cmdTrans(5).Top = cmdTrans(0).Top
        'Align Cancel Button
        cmdTrans(6).Left = cmdTrans(1).Left
        cmdTrans(6).Top = cmdTrans(1).Top
    End If
End Sub

Sub clearform()

    flxSDiscrepancy.Rows = 1
    flxSDiscrepancy.Rows = 2
    'Frame 1
    cmbBranch.Text = ""
    txtSTFNo.Text = ""
    txtSTFDate.Text = "__/__/____"
    txtRefNo.Text = "<AUTO>"
    txtDateRef.Text = Format(Now, "mm/dd/yyyy")
    txtDateRec.Text = "__/__/____"
    txtRRno.Text = ""
    
    'Frame 3
    txtStatus.Text = ""
    txtRemarks.Text = ""
    
    'Frame 4
    txtApprvBy.Text = ""
    
End Sub

Sub LockData(DataEntry As Boolean)
    DataEntry = Not DataEntry
    'Frame 1
    cmbBranch.Locked = DataEntry
    txtSTFNo.Locked = DataEntry
    'txtSTFDate.Locked = DataEntry
    txtRefNo.Locked = DataEntry
    'txtDateRef.Locked = DataEntry
    'txtDateRec.Locked = DataEntry
    txtRRno.Locked = DataEntry
    'Frame 3
    txtStatus.Locked = DataEntry
    txtRemarks.Locked = DataEntry
    'Frame 4
    txtPrepby.Locked = DataEntry
    txtApprvBy.Locked = DataEntry
    
    If cmbBranch.Locked Then
        'Frame 1
        cmbBranch.BackColor = txtHoldGrayBackColor.BackColor
        txtSTFNo.BackColor = txtHoldGrayBackColor.BackColor
        txtSTFDate.BackColor = txtHoldGrayBackColor.BackColor
        txtRefNo.BackColor = txtHoldGrayBackColor.BackColor
        txtDateRef.BackColor = txtHoldGrayBackColor.BackColor
        txtDateRec.BackColor = txtHoldGrayBackColor.BackColor
        txtRRno.BackColor = txtHoldGrayBackColor.BackColor
        'Frame 3
        txtStatus.BackColor = txtHoldGrayBackColor.BackColor
        txtRemarks.BackColor = txtHoldGrayBackColor.BackColor
        'Frame 4
        txtPrepby.BackColor = txtHoldGrayBackColor.BackColor
        txtApprvBy.BackColor = txtHoldGrayBackColor.BackColor
    Else
        'Frame 1
        cmbBranch.BackColor = txtHoldWhiteBackColor.BackColor
        txtSTFNo.Locked = True '.BackColor = txtHoldWhiteBackColor.BackColor
        txtSTFDate.BackColor = txtHoldGrayBackColor.BackColor
        txtRefNo.Locked = True
        'txtDateRef.BackColor = txtHoldWhiteBackColor.BackColor
        txtDateRec.BackColor = txtHoldGrayBackColor.BackColor
        txtRRno.BackColor = txtHoldWhiteBackColor.BackColor
        'Frame 3
        txtStatus.Locked = True
        txtRemarks.BackColor = txtHoldWhiteBackColor.BackColor
        'Frame 4
        txtPrepby.Locked = True
        txtApprvBy.BackColor = txtHoldWhiteBackColor.BackColor
    End If
End Sub

Private Sub SetButtonVisibility(DataEntry As Boolean)
'cmdTrans Legend
'0 ->ADD
'1 ->EDIT
'2 ->VIEW SUMM
'3 ->EXIT
'4 ->PRINT
'5 ->SAVE
'6 ->CANCEL

    'hide all button
    For i = 0 To 5
        cmdTrans(i).Visible = False
    Next i

    If DataEntry Then
        cmdTrans(5).Visible = True
        cmdTrans(6).Visible = True
    Else
        For i = 0 To 5
            cmdTrans(i).Visible = True
            cmdTrans(5).Visible = False
            cmdTrans(6).Visible = False
        Next i
    End If
End Sub

Public Sub flxChanger()
    With flxSDiscrepancy
        .Cols = 7
        .FormatString = " |<PARTNUMBER |<DESCRIPTION |^QTY |<REMARKS |<RECOMMENDATION"
        
        .ColWidth(0) = 200
        .ColWidth(1) = (.Width - .ColWidth(0)) * 0.15   'PARTS NUMBER
        .ColWidth(2) = (.Width - .ColWidth(0)) * 0.25  'DESCRIPTION
        .ColWidth(3) = (.Width - .ColWidth(0)) * 0.1 'QTY
        .ColWidth(4) = (.Width - .ColWidth(0)) * 0.15 'REMARKS
        .ColWidth(5) = (.Width - .ColWidth(0)) * 0.33 'RECOMMENDATION
        .ColWidth(6) = 0 '(.Width - .ColWidth(0)) * 0.06 'SPID
    End With
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo er

    If ProperExit Then
        If mainRs.State = adStateOpen Then mainRs.Close
        Call cmdTrans_Click(4)
        ProperExit = False
        MotoRep.toolBarButton.Visible = True
    Else
        MotoRep.toolBarButton.Visible = False
        Cancel = True
    End If
    
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->QueryUnload()")
End Sub

Private Sub Form_Resize()
Call RXResizer1_ResizeComplete
End Sub

Private Sub RXResizer1_ResizeComplete()
Call flxChanger
End Sub

Public Sub CboDisplay()
On Error GoTo er
    'Get Data from database
    DoEvents
    Call modMotoR.sp_get_Branch(mainRs)
    If mainRs.State = 1 Then
        Call modObjTools.addToListObjectWithArrayGivenField(cmbBranch, mainRs, arrBRANCH, "BRANCHCODE", "BRANCHNAME")
    End If
    If mainRs.State <> 0 Then mainRs.Close
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->CboDisplay()")
End Sub

Private Sub txtApprvBy_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtApprvBy, KeyAscii, ValidChars)
    If KeyAscii = 13 Then cmbBranch.SetFocus
End Sub

Private Sub txtRecom_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    'Call ObjKeyTrapping.Validation(txtSPECMODEL, KeyAscii, ValidChars)
    If Chr(KeyAscii) = "'" Then KeyAscii = 0
    If KeyAscii = 13 Then
        Call txtRecom_LostFocus
    End If
End Sub

Private Sub txtRecom_LostFocus()
    If txtRecom.Tag = "skiplostfocus" Then
        txtRecom.Tag = ""
        Exit Sub
    End If
    If txtRecom.Text = "" Then
        txtRecom.Visible = False
    End If
    
    flxSDiscrepancy.TextMatrix(ActRow, 5) = txtRecom.Text
    txtRecom.Visible = False
End Sub

Private Sub txtRemarks_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtRemarks, KeyAscii, ValidChars)
    If KeyAscii = 13 Then txtApprvBy.SetFocus
End Sub

Private Sub txtRRNo_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtRRno, KeyAscii, ValidChars)
    If KeyAscii = 13 Then
        Call txtRRNo_LostFocus
    End If
End Sub

Private Sub txtRRNo_LostFocus()
Dim Qry As String, tmpDateRec As String, tmpDateTrans As String, chkExists As String, ChkStatus As String
Dim rsRec As New ADODB.Recordset
On Error GoTo er
    tmpDateRec = ""
    chkStock = ""
    ChkRec = ""
    If txtRRno.Text <> "" Then
        chkExists = GetFieldValue("SELECT RRNO FROM SDISCREPANCY WHERE RRNO = '" & txtRRno.Text & "'")
        
        If chkExists = "" Then
            ChkRec = GetFieldValue("SELECT PACKINGNO FROM SPRECEIVE WHERE PACKINGNO = '" & txtRRno.Text & "'")
            If ChkRec <> "" Then
                txtSTFNo.Locked = False
                txtSTFDate.BackColor = txtHoldWhiteBackColor.BackColor
                txtSTFNo.BackColor = txtHoldWhiteBackColor.BackColor
                
                tmpDateRec = Format(GetFieldValue("Select spreceive.DATEREC FROM spreceive where spreceive.PACKINGNO = '" & txtRRno.Text & "'"), "mm/dd/YYYY")
                
                If tmpDateRec = "" Then
                    txtDateRec.Text = "__/__/____"
                Else
                    txtDateRec.Text = tmpDateRec
                End If
                
                Qry = "call sp_partsmatrix.sp_getReceiving('" & sysBranchCode & "', '" & txtRRno.Text & "')"
                With rsRec
                    .Open Qry, LYRcon, adOpenStatic, adLockOptimistic
                    If .RecordCount > 0 Then
                        .MoveFirst
                        flxSDiscrepancy.Rows = .RecordCount + 1
                        
                        For i = 1 To flxSDiscrepancy.Rows - 1
                            flxSDiscrepancy.TextMatrix(i, 1) = !PartNumber
                            flxSDiscrepancy.TextMatrix(i, 2) = !Description
                            flxSDiscrepancy.TextMatrix(i, 3) = !Qty
                            flxSDiscrepancy.TextMatrix(i, 4) = !Remarks
                            flxSDiscrepancy.TextMatrix(i, 6) = !SparepartID
                            .MoveNext
                        Next i
                    Else
                        MsgBox "No Record Found", vbInformation, "Error"
                    End If
                    .Close
                End With
                txtSTFNo.SetFocus
            Else
                MsgBox "No Discrepancy to process because '" & txtRRno.Text & "' does not exists", vbInformation, "Error"
                Exit Sub
            End If
        Else
            ChkStatus = GetFieldValue("SELECT STATUS FROM SDISCREPANCY WHERE RRNO ='" & txtRRno.Text & "'")
            MsgBox "'" & txtRRno.Text & "' is already process and Status is '" & ChkStatus & "'", vbInformation, "Error"
            Exit Sub
        End If
    End If
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->txtRRNO()")
End Sub

Public Sub ShowRecord(myStatus As String, mySDiscID As Long)
On Error GoTo er

    txtSDiscID.Text = mySDiscID 'Used for Edit Verification
    Screen.MousePointer = vbHourglass
    
    If myStatus = "PENDING" Then
        chkFinal.Value = 0
    Else
        chkFinal.Value = 1
    End If
    
    DoEvents
    Qry = "call sp_partsmatrix.sp_SDiscrepancy_getHeader('" & sysBranchCode & "', " & mySDiscID & ")"
    With mainRs
        .Open Qry, LYRcon, adOpenStatic, adLockOptimistic
        
        If .RecordCount > 0 Then
            .MoveFirst
            cmbBranch.Text = IIf(IsNull(!TRANSTO), "", !TRANSTO)
            txtSTFNo.Text = IIf(IsNull(!STFNO), "", !STFNO)
            txtSTFDate.Text = Format(!STFDATE, "MM/DD/YYYY")
            txtRRno.Text = IIf(IsNull(!RRNO), "", !RRNO)
            txtDateRec.Text = Format(!DATEREC, "MM/DD/YYYY")
            txtRefNo.Text = IIf(IsNull(!RefNo), "", !RefNo)
            txtDateRef.Text = Format(!DATEREF, "MM/DD/YYYY")
            txtStatus.Text = IIf(IsNull(!STATUS), "", !STATUS)
            txtRemarks.Text = IIf(IsNull(!Remarks), "", !Remarks)
            txtPrepby.Text = IIf(IsNull(!PREPBY), "", !PREPBY)
            txtApprvBy.Text = IIf(IsNull(!APPRVBY), "", !APPRVBY)
        End If
        .Close
    End With
    
    DoEvents
    Qry = "call sp_partsmatrix.sp_SDiscrepancy_getDetails('" & sysBranchCode & "', " & mySDiscID & ")"
    With mainRs
        .Open Qry, LYRcon, adOpenStatic, adLockOptimistic
        
        If .RecordCount > 0 Then
            .MoveFirst
            flxSDiscrepancy.Rows = .RecordCount + 1
            
            For i = 1 To flxSDiscrepancy.Rows - 1
                flxSDiscrepancy.TextMatrix(i, 1) = !PartNumber
                flxSDiscrepancy.TextMatrix(i, 2) = !Description
                flxSDiscrepancy.TextMatrix(i, 3) = !Qty
                flxSDiscrepancy.TextMatrix(i, 4) = !Remarks
                flxSDiscrepancy.TextMatrix(i, 5) = !BRECOMMENDATION
                flxSDiscrepancy.TextMatrix(i, 6) = !SparepartID
                .MoveNext
            Next i
        End If
    End With
    
    If mainRs.State <> 0 Then
        Screen.MousePointer = vbDefault
        mainRs.Close
    End If
    
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->ShowRecord()")
End Sub

Private Sub txtSTFDate_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtSTFDate, KeyAscii, ValidChars)
    If KeyAscii = 13 Then txtRRno.SetFocus
End Sub

Private Sub txtSTFNo_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    Call ObjKeyTrapping.Validation(txtSTFNo, KeyAscii, ValidChars)
    If KeyAscii = 13 Then
        Call txtSTFNo_LostFocus
    End If
End Sub

Private Sub txtSTFNo_LostFocus()
Dim Qry As String, tmpDateTrans As String, cntSDTypeLack As Integer, cntSDTypeDamage As Integer
Dim chkSupplier As String, tmpTransferid As String, chkExistRec As String, rsOut As New ADODB.Recordset
On Error GoTo er
    chkStock = ""
    cntSDTypeLack = 0
    cntSDTypeDamage = 0
    If txtSTFNo.Text <> "" Then
        For i = 1 To flxSDiscrepancy.Rows - 1
            If flxSDiscrepancy.TextMatrix(i, 4) = "LACKING" Then
                cntSDTypeLack = cntSDTypeLack + 1
            ElseIf flxSDiscrepancy.TextMatrix(i, 4) = "DAMAGE" Then
                cntSDTypeDamage = cntSDTypeDamage + 1
            End If
        Next i
      
        If cntSDTypeLack > 0 Then
            chkExistRec = GetFieldValue("Select PACKINGNO FROM SPRECEIVE where PACKINGNO = '" & txtSTFNo.Text & "'")
            If chkExistRec <> "" And txtRRno.Text = txtSTFNo.Text Then
                'connect to warehouse database
                Call ConnectMain
                tmpTransferid = Replace(txtSTFNo.Text, "-MWH", "")
                
                'query the transdate from warehouse database
                Qry = "SELECT TRANSDATE FROM STOCKTRANSFER WHERE TRANSFERID = '" & tmpTransferid & "'"
                Set rsOut = MainCon.Execute(Qry)
                
                'if record was not found
                If rsOut.RecordCount <= 0 Then
                    MsgBox "STFNO does not Exists in Warehouse. Please Enter STF DATE Manually.", vbInformation + vbOKOnly
                    txtSTFDate.SetFocus
                    rsOut.Close 'close recordset
                    Exit Sub
                End If
                
                'if record was found
                If rsOut.RecordCount > 0 Then
                    txtSTFDate.Text = Format(rsOut!TRANSDATE, "mm/dd/yyyy")
                End If
                
                rsOut.Close ' close recordset
            Else
                MsgBox "The Discrepancy Type is Lacking. Make it sure that the RRNO '" & txtRRno.Text & "' and STFNO '" & txtSTFNo.Text & "' must be the same Value.", vbInformation + vbOKOnly
                Exit Sub
            End If
            
        ElseIf cntSDTypeDamage > 0 Then
            chkStock = GetFieldValue("Select STOCKTRANSFER.TRANSNO FROM STOCKTRANSFER where STOCKTRANSFER.TRANSNO = '" & txtSTFNo.Text & "'")
            tmpDateTrans = Format(GetFieldValue("Select STOCKTRANSFER.TRANSDATE FROM STOCKTRANSFER where STOCKTRANSFER.TRANSNO = '" & txtSTFNo.Text & "'"), "mm/dd/YYYY")
        
            If tmpDateTrans = "" Then
                txtSTFDate.Text = "__/__/____"
            Else
                txtSTFDate.Text = tmpDateTrans
            End If
            If chkStock = "" Then
                MsgBox "The Discrepancy Type is Damage. Make it sure that the STFNO '" & txtSTFNo.Text & "' is exists in your STOCKTRANSFER Record", vbInformation + vbOKOnly
                Exit Sub
            End If
        ElseIf cntSDTypeLack > 0 And cntSDTypeDamage > 0 Then
            chkStock = GetFieldValue("Select STOCKTRANSFER.TRANSNO FROM STOCKTRANSFER where STOCKTRANSFER.TRANSNO = '" & txtSTFNo.Text & "'")
            tmpDateTrans = Format(GetFieldValue("Select STOCKTRANSFER.TRANSDATE FROM STOCKTRANSFER where STOCKTRANSFER.TRANSNO = '" & txtSTFNo.Text & "'"), "mm/dd/YYYY")
        
            If tmpDateTrans = "" Then
                txtSTFDate.Text = "__/__/____"
            Else
                txtSTFDate.Text = tmpDateTrans
            End If
            If chkStock = "" Then
                MsgBox "One of the Partnumber is Damage. Make it sure that the STFNO '" & txtSTFNo.Text & "' is exists in your STOCKTRANSFER Record", vbInformation + vbOKOnly
                Exit Sub
            End If
        End If
        
        flxSDiscrepancy.SetFocus
        flxSDiscrepancy.Row = 1
        flxSDiscrepancy.Col = 5
    End If 'end of If txtSTFNo.Text <> "" Then
Exit Sub
er:
    DoEvents
    Call modObjTools.ErrorMessage(Err.Description, "Error->txtSTFNo()")
End Sub
