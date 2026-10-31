VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form TLSprintPref 
   Caption         =   "Printing Preferences"
   ClientHeight    =   8910
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12090
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8910
   ScaleWidth      =   12090
   WindowState     =   2  'Maximized
   Begin VB.CheckBox ContractVisible 
      Caption         =   "Check1"
      Height          =   255
      Left            =   1080
      TabIndex        =   35
      Top             =   9000
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdA 
      Caption         =   "&A"
      Height          =   495
      Left            =   0
      TabIndex        =   34
      Top             =   9000
      Width           =   75
   End
   Begin VB.TextBox txtPrefID 
      Height          =   285
      Left            =   6480
      TabIndex        =   33
      Top             =   720
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtHeight 
      Height          =   315
      Left            =   10080
      TabIndex        =   2
      Tag             =   "XYF"
      Top             =   780
      Width           =   1935
   End
   Begin VB.ComboBox cboAlign 
      Height          =   315
      Left            =   9840
      TabIndex        =   5
      Top             =   5280
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.ComboBox cboBold 
      Height          =   315
      Left            =   8520
      TabIndex        =   4
      Top             =   5280
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtFlx 
      Height          =   285
      Left            =   3480
      TabIndex        =   3
      Top             =   5280
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtForm 
      Height          =   315
      Left            =   1680
      MaxLength       =   50
      TabIndex        =   1
      Tag             =   "XYF"
      Top             =   780
      Width           =   3615
   End
   Begin VB.CheckBox chkGrid 
      Caption         =   "FORM WITH GRID"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Tag             =   "XYF"
      Top             =   6600
      Width           =   2895
   End
   Begin MSFlexGridLib.MSFlexGrid flxPrn 
      Height          =   5295
      Left            =   240
      TabIndex        =   6
      Tag             =   "XYF"
      Top             =   1200
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   9340
      _Version        =   393216
      Cols            =   9
      FixedCols       =   0
      RowHeightMin    =   315
      BackColorBkg    =   -2147483633
      ScrollBars      =   2
      FormatString    =   $"TLSprintPref.frx":0000
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
   Begin VB.Frame Frame1 
      Height          =   1935
      Left            =   240
      TabIndex        =   0
      Tag             =   "XYF"
      Top             =   6840
      Visible         =   0   'False
      Width           =   11775
      Begin VB.TextBox txtSubLen 
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
         Left            =   4920
         TabIndex        =   14
         Tag             =   "XYF"
         Top             =   512
         Width           =   6495
      End
      Begin VB.TextBox txtSubFBold 
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
         Left            =   4920
         TabIndex        =   15
         Tag             =   "XYF"
         Top             =   844
         Width           =   6495
      End
      Begin VB.TextBox txtSubLeft 
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
         Left            =   4920
         TabIndex        =   13
         Tag             =   "XYF"
         Top             =   180
         Width           =   6495
      End
      Begin VB.TextBox txtSubAlign 
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
         Left            =   4920
         TabIndex        =   16
         Tag             =   "XYF"
         Top             =   1176
         Width           =   6495
      End
      Begin VB.TextBox txtInsertIndex 
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
         Left            =   1800
         TabIndex        =   8
         Tag             =   "XYF"
         Top             =   210
         Width           =   1335
      End
      Begin VB.TextBox txtLoopInterval 
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
         Left            =   1800
         TabIndex        =   12
         Tag             =   "XYF"
         Top             =   1520
         Width           =   1335
      End
      Begin VB.TextBox txtLoopFSize 
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
         Left            =   1800
         TabIndex        =   11
         Tag             =   "XYF"
         Top             =   1206
         Width           =   1335
      End
      Begin VB.TextBox txtLoopHeight 
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
         TabIndex        =   10
         Tag             =   "XYF"
         Top             =   864
         Width           =   1335
      End
      Begin VB.TextBox txtLoopStartTop 
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
         TabIndex        =   9
         Tag             =   "XYF"
         Top             =   522
         Width           =   1335
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Grid Row Interval"
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
         TabIndex        =   32
         Tag             =   "XYF"
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Grid Cell Height"
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
         TabIndex        =   28
         Tag             =   "XYF"
         Top             =   930
         Width           =   1455
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Grid Cell Alignment"
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
         Index           =   8
         Left            =   3240
         TabIndex        =   23
         Tag             =   "XYF"
         Top             =   1215
         Width           =   1575
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Grid Cell Left"
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
         Left            =   3240
         TabIndex        =   22
         Tag             =   "XYF"
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Grid Font Size"
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
         TabIndex        =   21
         Tag             =   "XYF"
         Top             =   1245
         Width           =   1455
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Grid Font Bold"
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
         Left            =   3240
         TabIndex        =   20
         Tag             =   "XYF"
         Top             =   930
         Width           =   1575
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Grid Cell Length"
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
         Left            =   3240
         TabIndex        =   19
         Tag             =   "XYF"
         Top             =   585
         Width           =   1575
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Grid Top Row"
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
         TabIndex        =   18
         Tag             =   "XYF"
         Top             =   585
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Insert Grid at Row"
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
         TabIndex        =   17
         Tag             =   "XYF"
         Top             =   240
         Width           =   1455
      End
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   36
      Top             =   0
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   661
      Caption         =   "Add &New"
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
      Image           =   "TLSprintPref.frx":0090
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   1
      Left            =   1440
      TabIndex        =   37
      Top             =   0
      Width           =   1440
      _ExtentX        =   2540
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
      Image           =   "TLSprintPref.frx":062A
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   3
      Left            =   4320
      TabIndex        =   38
      Top             =   0
      Width           =   1440
      _ExtentX        =   2540
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
      Image           =   "TLSprintPref.frx":0BC4
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   4
      Left            =   5760
      TabIndex        =   39
      Top             =   0
      Visible         =   0   'False
      Width           =   1440
      _ExtentX        =   2540
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
      Image           =   "TLSprintPref.frx":115E
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   5
      Left            =   7200
      TabIndex        =   40
      Top             =   0
      Visible         =   0   'False
      Width           =   1440
      _ExtentX        =   2540
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
      Image           =   "TLSprintPref.frx":16F8
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   2
      Left            =   2880
      TabIndex        =   41
      Top             =   0
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   661
      Caption         =   "&Find"
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
      Image           =   "TLSprintPref.frx":1C92
      cBack           =   -2147483633
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "FORM HEIGHT"
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
      Left            =   7680
      TabIndex        =   43
      Tag             =   "XYF"
      Top             =   810
      Width           =   2295
   End
   Begin VB.Label Label5 
      Caption         =   "NAME OF FORM"
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
      TabIndex        =   42
      Tag             =   "XYF"
      Top             =   840
      Width           =   1455
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   3840
      Top             =   8040
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   12330
      m_MDI_Height    =   9480
   End
   Begin VB.Label lblMax 
      Height          =   255
      Left            =   6840
      TabIndex        =   31
      Top             =   2400
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblReo 
      Height          =   255
      Left            =   5520
      TabIndex        =   30
      Top             =   2400
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label lblLoc 
      Height          =   255
      Left            =   3720
      TabIndex        =   29
      Top             =   2400
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Printing Preferences"
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
      Index           =   5
      Left            =   9000
      TabIndex        =   27
      Top             =   480
      Width           =   3015
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Printing Preferences"
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
      Index           =   4
      Left            =   9030
      TabIndex        =   26
      Top             =   510
      Width           =   3015
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "MOTO-R REPUBLIC CORP."
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
      Index           =   3
      Left            =   240
      TabIndex        =   25
      Top             =   480
      Width           =   3500
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "MOTO-R REPUBLIC CORP."
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
      Index           =   2
      Left            =   270
      TabIndex        =   24
      Top             =   510
      Width           =   3500
   End
   Begin VB.Menu mnuPref 
      Caption         =   "Printing Preferences"
      Visible         =   0   'False
      Begin VB.Menu mnuAdd 
         Caption         =   "Add Row"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete Row"
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInsert 
         Caption         =   "Insert Row"
      End
   End
End
Attribute VB_Name = "TLSprintPref"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsPref As New ADODB.Recordset

Dim ActRow As Integer, ActCol As Integer
Dim NewPref As Boolean

Private Sub ClearEntries()
    txtForm.Text = ""
    txtHeight.Text = ""
    flxPRN.Rows = 1
    flxPRN.Rows = 2
    txtInsertIndex.Text = ""
    txtLoopStartTop.Text = ""
    txtLoopHeight.Text = ""
    txtLoopFSize.Text = ""
    txtLoopInterval.Text = ""
    txtSubLeft.Text = ""
    txtSubLen.Text = ""
    txtSubFBold.Text = ""
    txtSubAlign.Text = ""
    txtPrefID.Text = ""
    chkGrid.Value = 0
End Sub

Private Sub tbSaveCancel(bln As Boolean)
    Dim i As Integer
    
    For i = 0 To 3
        cmdTrans(i).Visible = Not bln
    Next i
    cmdTrans(4).Visible = bln
    cmdTrans(5).Visible = bln
    cmdTrans(4).Left = 0
    cmdTrans(5).Left = cmdTrans(4).Width
End Sub

Private Sub LockData(bln As Boolean)
    txtForm.Locked = bln
    txtHeight.Locked = bln
    chkGrid.Enabled = Not bln
    flxPRN.Enabled = Not bln
    Frame1.Enabled = Not bln
End Sub

Private Sub cboAlign_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call cboAlign_LostFocus
        flxPRN.SetFocus
        flxPRN.Col = 0
    End If
End Sub

Private Sub cboAlign_LostFocus()
    If cboAlign.ListIndex = -1 Then
        MsgBox "Please select your answer from list."
        cboAlign.Text = ""
        cboAlign.Visible = False
        Exit Sub
    End If
    flxPRN.TextMatrix(ActRow, ActCol) = cboAlign.Text
    cboAlign.Visible = False
End Sub

Private Sub cboBold_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call cboBold_LostFocus
        flxPRN.SetFocus
        flxPRN.Col = 8
    End If
End Sub

Private Sub cboBold_LostFocus()
    If cboBold.ListIndex = -1 Then
        MsgBox "Please select your answer from list."
        cboBold.Text = ""
        cboBold.Visible = False
        Exit Sub
    End If
    flxPRN.TextMatrix(ActRow, ActCol) = cboBold.Text
    cboBold.Visible = False
End Sub

Public Sub ShowRecord(MyID As Integer)
    Dim Qry As String
    Dim X As Integer
    Dim MyDesc() As String, myTop() As String, myLeft() As String, myLen() As String
    Dim myHeight() As String, myFsize() As String, myFbold() As String, myAlign() As String
            
    'Qry = "SELECT PRINTPREF.* From PRINTPREF " & _
          "WHERE (((PRINTPREF.PREFID)=" & MyId & "));"
    'mySQL
    Qry = "SELECT PRINTPREF.* From PRINTPREF " & _
          "WHERE PRINTPREF.PREFID=" & MyID & ";"
    
    rsPref.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If rsPref.RecordCount > 0 Then
            rsPref.MoveFirst
            txtPrefID.Text = MyID
            txtForm.Text = rsPref!FormName
            txtHeight.Text = rsPref!FORMHEIGHT
            MyDesc = Split(rsPref!Description, ",")
            myTop = Split(rsPref!Top, ",")
            myLeft = Split(rsPref!Left, ",")
            myLen = Split(rsPref!Length, ",")
            myHeight = Split(rsPref!Height, ",")
            myFsize = Split(rsPref!FontSize, ",")
            myFbold = Split(rsPref!FontBold, ",")
            myAlign = Split(rsPref!Alignment, ",")
            flxPRN.Rows = 1
            For X = LBound(MyDesc) To UBound(MyDesc)
                flxPRN.AddItem vbTab & MyDesc(X) & vbTab & myTop(X) & vbTab & myLeft(X) & vbTab & myLen(X) & vbTab & myHeight(X) & vbTab & myFsize(X) & vbTab & IIf(myFbold(X) = 1, "True", "False") & vbTab & GetAlignRev(myAlign(X))
            Next X
            If rsPref!INSERTINDEX <> 0 Then
                chkGrid.Value = 1
                txtInsertIndex.Text = rsPref!INSERTINDEX
                txtLoopStartTop.Text = rsPref!LOOPSTARTTOP
                txtLoopHeight.Text = rsPref!LOOPHEIGHT
                txtLoopFSize.Text = rsPref!LOOPFSIZE
                txtLoopInterval.Text = rsPref!LOOPINTERVAL
                txtSubLeft.Text = rsPref!SUBLEFT
                txtSubLen.Text = rsPref!SUBLEN
                txtSubFBold.Text = rsPref!SUBFBOLD
                txtSubAlign.Text = rsPref!SUBALIGN
            End If
        Else
            flxPRN.Rows = 1
            flxPRN.Rows = 2
            txtInsertIndex.Text = ""
            txtLoopStartTop.Text = ""
            txtLoopHeight.Text = ""
            txtLoopFSize.Text = ""
            txtLoopInterval.Text = ""
            txtSubLeft.Text = ""
            txtSubLen.Text = ""
            txtSubFBold.Text = ""
            txtSubAlign.Text = ""
            txtPrefID.Text = ""
        End If
    rsPref.Close
End Sub



Private Sub chkGrid_Click()
    If chkGrid.Value = 1 Then
        Frame1.Visible = True
    Else
        Frame1.Visible = False
    End If
End Sub

Private Sub cmdA_Click()
    Call mnuAdd_Click
    flxPRN.SetFocus
    flxPRN.Col = 1
End Sub

Private Sub cmdTrans_Click(Index As Integer)
    Dim myMsg As String, Qry As String, MyAcc As String
    Dim Ans As Integer
    
    Select Case Index
        Case 0  '"new"
            'Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE (((USERACCESS.EMPID)=" & UserID & ") AND ((USERACCESS.MENUINDX)=95));"
            
            'mySQL
            Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE USERACCESS.EMPID=" & UserID & " AND USERACCESS.MENUINDX=95;"
            
            MyAcc = GetFieldValue(Qry)
            If MyAcc <> "FULL ACCESS" Then
                MsgBox "You don't have full access to this transaction." & vbCrLf & "Please contact your System Administrator", vbExclamation + vbOKOnly, "Access Denied"
                Exit Sub
            End If
        
            Call tbSaveCancel(True)
            Call LockData(False)
            Call ClearEntries
            NewPref = True
        
        Case 1  '"edit"
            'Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE (((USERACCESS.EMPID)=" & UserID & ") AND ((USERACCESS.MENUINDX)=95));"
            
            'mySQL
            Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE USERACCESS.EMPID=" & UserID & " AND USERACCESS.MENUINDX=95;"
            
            MyAcc = GetFieldValue(Qry)
            If MyAcc <> "FULL ACCESS" Then
                MsgBox "You don't have full access to this transaction." & vbCrLf & "Please contact your System Administrator", vbExclamation + vbOKOnly, "Access Denied"
                Exit Sub
            End If
            If txtPrefID.Text = "" Then
                MsgBox "No Printing Preference available!", vbExclamation + vbOKOnly, "Edit Denied"
                Exit Sub
            End If
            Call LockData(False)
            Call tbSaveCancel(True)
            NewPref = False
        
        Case 2  '"search"
            Set frmSearch.FrmSource = TLSprintPref
            frmSearch.QrySearch = "SELECT PRINTPREF.PREFID, PRINTPREF.FORMNAME From PRINTPREF ORDER BY PRINTPREF.FORMNAME;"
            frmSearch.Show vbModal
        
        Case 3  '"quit"
            'If ContractVisible.Value = 0 Then LYRmain.tbMain.Visible = True
            Unload Me
            'MotoRep.tbMain.Visible = True
            'MotoRep.imgBG.Visible = True
            MotoRep.toolBarButton.Visible = True
        Case 4  '"save"
            myMsg = ValidatePref
            If myMsg <> "" Then
                MsgBox myMsg, vbExclamation + vbOKOnly, "Validation Error"
                Exit Sub
            End If
            Ans = MsgBox("Do you want to save preferences?", vbQuestion + vbYesNo, "Save")
            If Ans = vbNo Then Exit Sub
            Call SavePrintPref
            Call tbSaveCancel(False)
            Call LockData(True)
            MsgBox "Print Preference for " & txtForm.Text & " successfully saved!", vbInformation + vbOKOnly
        
        Case 5  '"cancel"
            Ans = MsgBox("Do you want to Cancel?", vbQuestion + vbYesNo, "Cancel")
            If Ans = vbNo Then Exit Sub
            Call tbSaveCancel(False)
            Call LockData(True)
            Call ClearEntries

    End Select

End Sub

Private Sub flxPrn_Click()
    Dim Tmp
    Const CB_SHOWDROPDOWN = &H14F
    ActRow = flxPRN.Row
    ActCol = flxPRN.Col
    
    If ActCol = 0 Or ActRow = 0 Then Exit Sub
    Select Case ActCol
        Case 1 To 6    '
            With txtFlx
                .Top = flxPRN.CellTop + flxPRN.Top
                .Left = flxPRN.CellLeft + flxPRN.Left
                .Width = flxPRN.CellWidth
                .Height = flxPRN.CellHeight
                .Text = flxPRN.Text
                .Visible = True
                .SelStart = 0
                .SelLength = Len(.Text)
                .ZOrder
                .SetFocus
            End With
        Case 7  'Font Bold
            With cboBold
                .Top = flxPRN.CellTop + flxPRN.Top
                .Left = flxPRN.CellLeft + flxPRN.Left
                .Width = flxPRN.CellWidth
                .Text = flxPRN.Text
                '.Height = flxPrn.CellHeight
                .Visible = True
                .SelStart = 0
                .SelLength = Len(.Text)
                .ZOrder
                .SetFocus
                 Tmp = SendMessage(.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
            End With
        Case 8    'Alignment
            With cboAlign
                .Top = flxPRN.CellTop + flxPRN.Top
                .Left = flxPRN.CellLeft + flxPRN.Left
                .Width = flxPRN.CellWidth
                '.Height = flxPrn.CellHeight
                .Text = flxPRN.Text
                .Visible = True
                .SelStart = 0
                .SelLength = Len(.Text)
                .ZOrder
                .SetFocus
       '          Const CB_SHOWDROPDOWN = &H14F
                 Tmp = SendMessage(.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
            End With
    End Select

End Sub

Private Sub flxPrn_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call flxPrn_Click
End Sub

Private Sub flxPrn_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    If Button = vbRightButton Then
        Call PopupMenu(mnuPref)
    End If
End Sub

Private Sub Form_Load()
    Dim Qry As String
    Dim pPct As Single
        
    'Add values to combo for flxPrn
    cboAlign.AddItem "Left"
    cboAlign.AddItem "Center"
    cboAlign.AddItem "Right"
    cboBold.AddItem "True"
    cboBold.AddItem "False"
    cmdA.Top = 15000
    Call tbSaveCancel(False)
    Call LockData(True)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If rsPref.State = adStateOpen Then rsPref.Close
End Sub

Private Sub mnuAdd_Click()
    Dim i As Integer
    Dim myMsg As String
    
    'Validate last row
    myMsg = ""
    With flxPRN
        For i = 1 To 8
            If .TextMatrix(ActRow, i) = "" Then
                myMsg = myMsg & vbCrLf & "Incomplete entry on field " & .TextMatrix(0, i)
            End If
        Next i
        If myMsg <> "" Then
            MsgBox myMsg, vbExclamation + vbOKOnly, "Validation Error"
            Exit Sub
        End If
        .AddItem ""
    End With
End Sub

Private Sub mnuDelete_Click()
    Dim Ans As Integer
    
    Ans = MsgBox("Do you want to Delete Row?", vbQuestion + vbYesNo, "Delete")
    If Ans = vbNo Then Exit Sub
    With flxPRN
        If ActRow = 1 And .Rows = 2 Then
            .Rows = 1
            .Rows = 2
        Else
            .RemoveItem ActRow
        End If
    End With
End Sub

Private Sub mnuInsert_Click()
    flxPRN.AddItem "", ActRow
End Sub


Function GetAlign(aStr As String) As String
    Select Case aStr
        Case "Left"
            GetAlign = "1"
        Case "Center"
            GetAlign = "2"
        Case "Right"
            GetAlign = "3"
    End Select
End Function

Function GetAlignRev(aStr As String) As String
    Select Case aStr
        Case "1"
            GetAlignRev = "Left"
        Case "2"
            GetAlignRev = "Center"
        Case "3"
            GetAlignRev = "Right"
    End Select
End Function
Private Sub SavePrintPref()
    Dim myTop As String, myLeft As String, myLen As String, myHeight As String, MyDesc As String
    Dim Qry As String, myFsize As String, myFbold As String, myAlign As String
    Dim i As Integer
    
    'If rsPref.State = adStateOpen Then rsPref.Close
    'If NewPref Then
    '    rsPref.Open "PrintPref", LYRcon, adOpenKeyset, adLockOptimistic
    '    rsPref.AddNew
    'Else
    '    Qry = "SELECT PRINTPREF.* From PRINTPREF " & _
    '          "WHERE (((PRINTPREF.PREFID)=" & txtPrefID.Text & "));"
    '    rsPref.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    '    rsPref.MoveFirst
    'End If
    MyDesc = "": myTop = "": myLeft = "": myLen = "": myHeight = "": myFsize = "": myFbold = "": myAlign = ""
    'Assign data to variables
    With flxPRN
        For i = 1 To .Rows - 1
            MyDesc = IIf(MyDesc = "", .TextMatrix(i, 1), MyDesc & "," & .TextMatrix(i, 1))
            myTop = IIf(myTop = "", .TextMatrix(i, 2), myTop & "," & .TextMatrix(i, 2))
            myLeft = IIf(myLeft = "", .TextMatrix(i, 3), myLeft & "," & .TextMatrix(i, 3))
            myLen = IIf(myLen = "", .TextMatrix(i, 4), myLen & "," & .TextMatrix(i, 4))
            myHeight = IIf(myHeight = "", .TextMatrix(i, 5), myHeight & "," & .TextMatrix(i, 5))
            myFsize = IIf(myFsize = "", .TextMatrix(i, 6), myFsize & "," & .TextMatrix(i, 6))
            myFbold = IIf(myFbold = "", IIf(.TextMatrix(i, 7) = "True", "1", "0"), myFbold & "," & IIf(.TextMatrix(i, 7) = "True", "1", "0"))
            myAlign = IIf(myAlign = "", GetAlign(.TextMatrix(i, 8)), myAlign & "," & GetAlign(.TextMatrix(i, 8)))
        Next i
    End With
    'Assign variables to fields
    'If NewPref Then rsPref!FormName = txtForm.Text
    'rsPref!FORMHEIGHT = txtHeight.Text
    'rsPref!Description = MyDesc
    'rsPref!Top = myTop
    'rsPref!Left = myLeft
    'rsPref!Length = myLen
    'rsPref!Height = myHeight
    'rsPref!FontSize = myFsize
    'rsPref!FontBold = myFbold
    'rsPref!Alignment = myAlign
    'Save grid preference
    'If chkGrid.Value = 1 Then
    '    rsPref!INSERTINDEX = txtInsertIndex.Text
    '    rsPref!LOOPSTARTTOP = txtLoopStartTop.Text
    '    rsPref!LOOPHEIGHT = txtLoopHeight.Text
    '    rsPref!LOOPFSIZE = txtLoopFSize.Text
    '    rsPref!LOOPINTERVAL = txtLoopInterval.Text
    '    rsPref!SUBLEFT = txtSubLeft.Text
    '    rsPref!SUBLEN = txtSubLen.Text
    '    rsPref!SUBFBOLD = txtSubFBold.Text
    '    rsPref!SUBALIGN = txtSubAlign.Text
    'End If
    'rsPref.Update
    'txtPrefID.Text = rsPref!PREFID
    'rsPref.Close
    
    
    'mySQL
    If NewPref Then
        'insert statement
        LYRcon.Execute "INSERT INTO printpref(FORMNAME, FORMHEIGHT, DESCRIPTION, TOP, `LEFT`, `LENGTH`, HEIGHT, FONTSIZE, FONTBOLD, ALIGNMENT)" & _
            "VALUES('" & txtForm.Text & "', " & txtHeight.Text & ", '" & MyDesc & "', '" & myTop & "', '" & myLeft & "', '" & myLen & "', " & _
            "'" & myHeight & "', '" & myFsize & "', '" & myFbold & "', '" & myAlign & "');"
        If chkGrid.Value = 1 Then
            LYRcon.Execute "INSERT INTO printpref(INSERTINDEX, LOOPSTARTTOP, LOOPHEIGHT, LOOPFSIZE, LOOPINTERVAL, SUBLEFT, SUBLEN, SUBFBOLD, " & _
                "SUBALIGN)VALUES(" & txtInsertIndex.Text & ", " & txtLoopStartTop.Text & ", " & txtLoopHeight.Text & ", " & txtLoopFSize.Text & ", " & _
                "" & txtLoopInterval.Text & ", '" & txtSubLeft.Text & "', '" & txtSubLen.Text & "', '" & txtSubFBold.Text & "', '" & txtSubAlign.Text & "');"
        End If
    Else
        LYRcon.Execute "Update printpref SET FORMNAME = '" & txtForm.Text & "', FORMHEIGHT = " & txtHeight.Text & ", DESCRIPTION = '" & MyDesc & "', " & _
            "TOP = '" & myTop & "', `LEFT` = '" & myLeft & "', `LENGTH` = '" & myLen & "', HEIGHT = '" & myHeight & "', FONTSIZE = '" & myFsize & "', " & _
            "FONTBOLD = '" & myFbold & "', ALIGNMENT = '" & myAlign & "' Where printpref.PREFID = " & txtPrefID.Text & ";"
        If chkGrid.Value = 1 Then
            'update Grip preference
            LYRcon.Execute "Update printpref SET INSERTINDEX = " & txtInsertIndex.Text & ", LOOPSTARTTOP = " & txtLoopStartTop.Text & ", " & _
                "LOOPHEIGHT = " & txtLoopHeight.Text & ", LOOPFSIZE = " & txtLoopFSize.Text & ", LOOPINTERVAL = " & txtLoopInterval.Text & ", " & _
                "SUBLEFT = '" & txtSubLeft.Text & "', SUBLEN = '" & txtSubLen.Text & "', SUBFBOLD = '" & txtSubFBold.Text & "', " & _
                "SUBALIGN = '" & txtSubAlign.Text & "' Where printpref.PREFID = " & txtPrefID.Text & ";"
        Else
            LYRcon.Execute "Update printpref set insertindex = 0, loopstarttop = 0, LOOPHEIGHT = 0, LOOPFSIZE = 0, LOOPINTERVAL = 0, SUBLEFT = '', SUBLEN = '', SUBFBOLD = '', SUBALIGN = '' Where printpref.PREFID = " & txtPrefID.Text & ";"
        End If
    End If
        Qry = ""
        Qry = "SELECT printpref.PREFID, printpref.INSERTINDEX, printpref.LOOPSTARTTOP, printpref.LOOPHEIGHT, printpref.LOOPFSIZE, printpref.LOOPINTERVAL, printpref.SUBLEFT, " & _
          "printpref.SUBLEN, printpref.SUBFBOLD, printpref.SUBALIGN From printpref Where printpref.FORMNAME = '" & txtForm.Text & "';"
        If rsPref.State = adStateOpen Then rsPref.Close
        rsPref.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        rsPref.MoveFirst
        
        txtPrefID.Text = rsPref!pREFID
        rsPref.Close
End Sub


Function ValidatePref() As String
    Dim MyErr As String, Qry As String
    Dim i As Integer
    
    MyErr = ""
    If NewPref Then
        If txtForm.Text = "" Then MyErr = "Incomplete entry on field Form Name"
        'Qry = "SELECT PRINTPREF.* From PRINTPREF " & _
            "WHERE (((PRINTPREF.FORMNAME)='" & txtForm.Text & "'));"
        'mySQL
        Qry = "SELECT PRINTPREF.* From PRINTPREF " & _
            "WHERE PRINTPREF.FORMNAME='" & txtForm.Text & "';"
        
        If CtrlNumExist(Qry) Then MyErr = "Form Name already exist"
    End If
    If txtHeight.Text = "" Then MyErr = MyErr & vbCrLf & "Incomplete entry on field Form Height"
    For i = 1 To 8
        If flxPRN.TextMatrix(flxPRN.Rows - 1, i) = "" Then
            MyErr = MyErr & vbCrLf & "Incomplete entry on field " & flxPRN.TextMatrix(0, i)
        End If
    Next i
    If chkGrid.Value = 1 Then
        If txtInsertIndex.Text = "" Or txtLoopStartTop.Text = "" Or txtLoopHeight.Text = "" Or txtLoopFSize.Text = "" Or txtLoopInterval.Text = "" Or txtSubLeft.Text = "" Or txtSubLen.Text = "" Or txtSubFBold.Text = "" Or txtSubAlign.Text = "" Then
            MyErr = MyErr & vbCrLf & "Incomplete entry in some field of Grid section."
        End If
    End If
'    If InStrRev(txtSubLeft.Text, ",") <> InStrRev(txtSubLen.Text, ",") Or InStrRev(txtSubFBold.Text, ",") Or InStrRev(txtSubAlign.Text, ",") Then
'        MyErr = MyErr & vbCrLf & "No. of values Mismatch: Cell Left, Length, Font Bold, Alignment"
'    End If
    ValidatePref = MyErr
End Function

Private Sub RXResizer1_ResizeComplete()
    Dim i As Integer
    
    With flxPRN
        .ColWidth(0) = 400
        .ColWidth(1) = (.Width - 400) * 0.22
        For i = 2 To 8
            .ColWidth(i) = (.Width - 400) * 0.11
        Next i
        Label3(4).Left = (.Left + .Width) - Label3(4).Width
        Label3(5).Left = (.Left + .Width) - Label3(5).Width + 30
    End With
    
            
End Sub
Private Sub txtFlx_KeyPress(KeyAscii As Integer)
    Dim strValid As String
    strValid = "0123456789"
    If ActCol >= 2 And ActCol <= 6 Then
        If KeyAscii > 26 Then
            If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
        End If
    End If
    If KeyAscii = 13 Then
        flxPRN.SetFocus
        flxPRN.Col = ActCol + 1
    End If
End Sub

Private Sub txtFlx_LostFocus()
    flxPRN.TextMatrix(ActRow, ActCol) = txtFlx.Text
    txtFlx.Visible = False
End Sub


Private Sub txtForm_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub


Private Sub txtForm_LostFocus()
    Dim Qry As String
    
    If NewPref Then
        Qry = "SELECT PRINTPREF.* From PRINTPREF " & _
              "WHERE (((PRINTPREF.FORMNAME)='" & txtForm.Text & "'));"
        If CtrlNumExist(Qry) Then
            MsgBox txtForm.Text & " already exist!", vbExclamation + vbOKOnly, "Warning"
        End If
    End If
End Sub

Private Sub txtHeight_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        flxPRN.SetFocus
        flxPRN.Col = 1
    End If
End Sub

Private Sub txtInsertIndex_KeyPress(KeyAscii As Integer)
    Dim strValid As String
    
    strValid = "0123456789"
    If KeyAscii > 26 Then
        If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtLoopFSize_KeyPress(KeyAscii As Integer)
    Dim strValid As String
    
    strValid = "0123456789"
    If KeyAscii > 26 Then
        If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtLoopHeight_KeyPress(KeyAscii As Integer)
    Dim strValid As String
    
    strValid = "0123456789"
    If KeyAscii > 26 Then
        If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtLoopInterval_KeyPress(KeyAscii As Integer)
    Dim strValid As String
    
    strValid = "0123456789"
    If KeyAscii > 26 Then
        If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtLoopStartTop_KeyPress(KeyAscii As Integer)
    Dim strValid As String
    
    strValid = "0123456789"
    If KeyAscii > 26 Then
        If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtSubAlign_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtInsertIndex.SetFocus
    End If
End Sub

Private Sub txtSubFBold_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtSubLeft_KeyPress(KeyAscii As Integer)
    Dim strValid As String
    
    strValid = "0123456789,"
    If KeyAscii > 26 Then
        If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub


Private Sub txtSubLen_KeyPress(KeyAscii As Integer)
    Dim strValid As String
    
    strValid = "0123456789,"
    If KeyAscii > 26 Then
        If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub
