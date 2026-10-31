VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form SPInsXst 
   BorderStyle     =   0  'None
   Caption         =   "Insert Existing Spare Parts"
   ClientHeight    =   6735
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12315
   Icon            =   "SPInsXst.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6735
   ScaleWidth      =   12315
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComCtl2.UpDown UpDRef 
      Height          =   315
      Left            =   1800
      TabIndex        =   37
      Top             =   240
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   556
      _Version        =   393216
      Max             =   1
      Enabled         =   -1  'True
   End
   Begin MSComctlLib.ProgressBar Pbar 
      Height          =   255
      Left            =   240
      TabIndex        =   21
      Top             =   6120
      Visible         =   0   'False
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.TextBox txtOrder 
      Height          =   285
      Left            =   3960
      TabIndex        =   20
      Top             =   6360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin lvButton.lvButtons_H cmdReplenish 
      Height          =   375
      Left            =   9960
      TabIndex        =   19
      Top             =   240
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      Caption         =   "Stocks for Replenish"
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
      Image           =   "SPInsXst.frx":058A
      cBack           =   -2147483633
   End
   Begin VB.TextBox txtSuppCode 
      Height          =   285
      Left            =   -120
      TabIndex        =   18
      Top             =   6840
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.ComboBox cboFind 
      Height          =   315
      Left            =   2040
      TabIndex        =   0
      Top             =   240
      Width           =   4335
   End
   Begin lvButton.lvButtons_H cmdInsert 
      Height          =   375
      Left            =   7920
      TabIndex        =   7
      Top             =   6240
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "&Insert"
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
      Image           =   "SPInsXst.frx":0B24
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdOptHide 
      Height          =   375
      Left            =   3600
      TabIndex        =   5
      Top             =   600
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      Caption         =   "&Hide Option"
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
      ImgAlign        =   3
      Image           =   "SPInsXst.frx":10BE
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdOptShow 
      Height          =   375
      Left            =   480
      TabIndex        =   4
      Top             =   600
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      Caption         =   "&Show Option"
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
      ImgAlign        =   3
      Image           =   "SPInsXst.frx":1658
      cBack           =   -2147483633
   End
   Begin VB.Frame FrameOpt 
      Height          =   1095
      Left            =   240
      TabIndex        =   3
      Top             =   720
      Width           =   11895
      Begin MSComCtl2.UpDown OptrSQ 
         Height          =   285
         Left            =   10560
         TabIndex        =   36
         Top             =   255
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   503
         _Version        =   393216
         Value           =   1
         Max             =   5
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown OptrBO 
         Height          =   285
         Left            =   8280
         TabIndex        =   35
         Top             =   615
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   503
         _Version        =   393216
         Max             =   5
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtQH 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   8535
         TabIndex        =   32
         Text            =   "0"
         Top             =   255
         Width           =   480
      End
      Begin MSComCtl2.UpDown updBO 
         Height          =   285
         Left            =   9000
         TabIndex        =   28
         Top             =   615
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   503
         _Version        =   393216
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtBO"
         BuddyDispid     =   196614
         OrigLeft        =   9840
         OrigTop         =   600
         OrigRight       =   10095
         OrigBottom      =   855
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtBO 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   8520
         TabIndex        =   27
         Text            =   "0"
         Top             =   615
         Width           =   480
      End
      Begin MSComCtl2.UpDown updSQ 
         Height          =   285
         Left            =   11280
         TabIndex        =   26
         Top             =   255
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   503
         _Version        =   393216
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtSQ"
         BuddyDispid     =   196615
         OrigLeft        =   11280
         OrigTop         =   255
         OrigRight       =   11535
         OrigBottom      =   540
         Max             =   1000
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtSQ 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   10800
         TabIndex        =   25
         Text            =   "0"
         Top             =   255
         Width           =   480
      End
      Begin MSComCtl2.DTPicker SDfrom 
         Height          =   315
         Left            =   1800
         TabIndex        =   15
         Top             =   720
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         Format          =   58785793
         CurrentDate     =   39498
      End
      Begin VB.CheckBox chkAllSupp 
         Caption         =   "All Suppliers"
         Height          =   255
         Left            =   4920
         TabIndex        =   13
         Top             =   360
         Width           =   1335
      End
      Begin VB.ComboBox cboSupplier 
         Height          =   315
         Left            =   1800
         TabIndex        =   12
         Top             =   360
         Width           =   3015
      End
      Begin MSComCtl2.DTPicker SDto 
         Height          =   315
         Left            =   4920
         TabIndex        =   17
         Top             =   720
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         Format          =   58785793
         CurrentDate     =   39498
      End
      Begin MSComCtl2.UpDown OptrQH 
         Height          =   285
         Left            =   8280
         TabIndex        =   31
         Top             =   255
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   503
         _Version        =   393216
         Max             =   5
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown updQH 
         Height          =   285
         Left            =   9000
         TabIndex        =   33
         Top             =   255
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   503
         _Version        =   393216
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtQH"
         BuddyDispid     =   196613
         OrigLeft        =   9015
         OrigTop         =   255
         OrigRight       =   9270
         OrigBottom      =   540
         Max             =   1000
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.Label lblOprQH 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "="
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
         Height          =   285
         Left            =   7800
         TabIndex        =   34
         Top             =   255
         Width           =   480
      End
      Begin VB.Label lblOprSQ 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   ">"
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
         Height          =   285
         Left            =   10080
         TabIndex        =   30
         Top             =   255
         Width           =   480
      End
      Begin VB.Label lblOprBO 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "="
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
         Height          =   285
         Left            =   7800
         TabIndex        =   29
         Top             =   615
         Width           =   480
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "SQ:"
         Height          =   255
         Index           =   2
         Left            =   9600
         TabIndex        =   24
         ToolTipText     =   "SALES QTY."
         Top             =   270
         Width           =   375
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "BO:"
         Height          =   255
         Index           =   1
         Left            =   7320
         TabIndex        =   23
         ToolTipText     =   "BACK ORDER"
         Top             =   630
         Width           =   375
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "QH:"
         Height          =   255
         Index           =   0
         Left            =   7320
         TabIndex        =   22
         ToolTipText     =   "QTY-ON-HAND"
         Top             =   270
         Width           =   375
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Sales Date To:"
         Height          =   255
         Left            =   3480
         TabIndex        =   16
         Top             =   750
         Width           =   1335
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Sales Date From:"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   750
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Supplier:"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   390
         Width           =   1215
      End
   End
   Begin lvButton.lvButtons_H cmdFind 
      Height          =   495
      Left            =   6435
      TabIndex        =   2
      Top             =   135
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      CapAlign        =   2
      BackStyle       =   2
      Shape           =   4
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
      Image           =   "SPInsXst.frx":1BF2
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdOK 
      Height          =   375
      Left            =   9360
      TabIndex        =   8
      Top             =   6240
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "&OK"
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
      Image           =   "SPInsXst.frx":218C
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdCancel 
      Height          =   375
      Left            =   10800
      TabIndex        =   9
      Top             =   6240
      Width           =   1335
      _ExtentX        =   2355
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
      Image           =   "SPInsXst.frx":2726
      cBack           =   -2147483633
   End
   Begin MSFlexGridLib.MSFlexGrid flxSP 
      Height          =   4215
      Left            =   240
      TabIndex        =   10
      Top             =   1920
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   7435
      _Version        =   393216
      Cols            =   12
      FixedCols       =   0
      RowHeightMin    =   300
      BackColorBkg    =   16777215
      FocusRect       =   2
      AllowUserResizing=   1
      FormatString    =   $"SPInsXst.frx":2CC0
   End
   Begin VB.Label lblChk 
      Caption         =   "ü"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   11
      Top             =   6360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblRefSearch 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FIND PARTS DESC."
      Height          =   195
      Left            =   180
      TabIndex        =   1
      Top             =   315
      Width           =   1485
   End
   Begin VB.Menu mnuSP 
      Caption         =   "SP"
      Visible         =   0   'False
      Begin VB.Menu mnuAct 
         Caption         =   "set to ""ACTIVE"""
      End
      Begin VB.Menu mnuInAct 
         Caption         =   "set to ""IN-ACTIVE"""
      End
      Begin VB.Menu sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMark 
         Caption         =   "Mark"
      End
      Begin VB.Menu mnuUnmark 
         Caption         =   "Unmark"
      End
      Begin VB.Menu sep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMarkAll 
         Caption         =   "Mark All"
      End
      Begin VB.Menu mnuUnmarkAll 
         Caption         =   "UnMark All"
      End
   End
End
Attribute VB_Name = "SPInsXst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsSupplier As New ADODB.Recordset
Dim ShowOpt As Boolean, SuppSrch As Boolean
Dim ActRow As Integer, ActCol As Integer
Public MySource As String

Private Sub cboFind_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call cmdFind_Click
End Sub

Private Sub cboSupplier_Change()
    Dim i As Long, j As Long
    Dim strPartial As String, strTotal As String
    Dim X As Integer
    
    If SuppSrch Then
        SuppSrch = False
        Exit Sub
    End If
    With cboSupplier
        strPartial = .Text
        i = SendMessage(.hwnd, CB_FINDSTRING, -1, ByVal strPartial) 'If match found, append unmatched characters
        If i <> -1 Then
            strTotal = .List(i)
            rsSupplier.MoveFirst
            rsSupplier.Move i
            txtSuppCode.Text = IIf(IsNull(rsSupplier!SUPPLIERCODE), "", rsSupplier!SUPPLIERCODE)
            chkAllSupp.Value = 0
            j = Len(strTotal) - Len(strPartial)
            If j <> 0 Then
                SuppSrch = True
                .SelText = Right$(strTotal, j)
                .SelStart = Len(strPartial)
                .SelLength = j
            End If
        Else
            txtSuppCode.Text = ""
        End If
    End With

End Sub

Private Sub cboSupplier_Click()
    Call cboSupplier_Change
End Sub

Private Sub cboSupplier_GotFocus()
   Dim Tmp
    
   Const CB_SHOWDROPDOWN = &H14F
   Tmp = SendMessage(cboSupplier.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub

Private Sub cboSupplier_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then SuppSrch = True
End Sub

Private Sub cboSupplier_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then cboFind.SetFocus
End Sub

Private Sub chkAllSupp_Click()
    If chkAllSupp.Value = 1 Then 'All Supplier
        cboSupplier.Text = ""
        txtSuppCode.Text = ""
    Else
        If txtSuppCode.Text = "" Then
            cboSupplier.Text = SPordering.cboSupplier.Text
            txtSuppCode.Text = SPordering.txtScode.Text
        End If
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    Dim rsSP As New ADODB.Recordset
    Dim MySP() As String, Mdl() As String, Qry As String
    Dim myW As String, myW1 As String, myW2 As String, myW3 As String
    Dim MySel As String, i As Integer
    
    If cboFind.Text = "" Then Exit Sub
    MySP = Split(cboFind.Text, " ")
    myW1 = ""
    For i = LBound(MySP) To UBound(MySP)
        If myW1 = "" Then
            'myW1 = "(SPAREPARTS.DESCRIPTION) Like '%" & MySP(i) & "%'"
            
            If UpDRef.Value = 0 Then
                myW1 = "(SPAREPARTS.PARTNUMBER) Like '%" & MySP(i) & "%'"
            Else
                myW1 = "(SPAREPARTS.DESCRIPTION) Like '%" & MySP(i) & "%'"
            End If
        Else
            'myW1 = myW1 & " AND (SPAREPARTS.DESCRIPTION) Like '%" & MySP(i) & "%'"
            If UpDRef.Value = 0 Then
                myW1 = myW1 & " AND (SPAREPARTS.PARTNUMBER) Like '%" & MySP(i) & "%'"
            Else
                myW1 = myW1 & " AND (SPAREPARTS.DESCRIPTION) Like '%" & MySP(i) & "%'"
            End If
        End If
    Next i
    myW = myW1
    If txtSuppCode.Text <> "" Then
        myW = "(" & myW & ") AND ((SPAREPARTS.SUPPLIERCODE) = '" & txtSuppCode.Text & "')"
    End If
    
    MySel = "SELECT  SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPAREPARTS.QTYONHAND, SPAREPARTS.BACKORDER, SPAREPARTS.SALES, SPAREPARTS.COST, SPAREPARTS.SPAREPARTID, SPAREPARTS.APPLICMODEL, SPAREPARTS.DESCRIPTION2 AS MYSTAT From SPAREPARTS WHERE (" & myW & ")  ORDER BY SPAREPARTS.PARTNUMBER;"
    With rsSP
        .Open MySel, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            flxSP.Rows = 1
            pBar.Value = 0
            pBar.Visible = True
            pBar.Max = .RecordCount * 2
            Do While Not .EOF
                flxSP.AddItem vbTab & !PartNumber & vbTab & !Description & vbTab & !QtyOnHand & vbTab & IIf(IsNull(!BACKORDER), "", !BACKORDER) & vbTab & IIf(IsNull(!Sales), "", !Sales) & vbTab & vbTab & Format(!Cost, "#,##0.00") & vbTab & vbTab & !SPAREPARTID & vbTab & IIf(IsNull(!ApplicModel), "", !ApplicModel) & vbTab & IIf(IsNull(!MyStat), "", !MyStat)
                pBar.Value = pBar.Value + 1
                .MoveNext
                DoEvents
            Loop
            flxSP.Visible = True
            If ShowOpt Then
                flxSP.Top = 1920
            Else
                flxSP.Top = 1080
            End If
            cmdInsert.Top = flxSP.Top + flxSP.Height + 120
            cmdOK.Top = flxSP.Top + flxSP.Height + 120
            cmdCancel.Top = flxSP.Top + flxSP.Height + 120
            pBar.Top = flxSP.Top + flxSP.Height
            Me.Height = cmdOK.Top + cmdOK.Height + 580
        Else
            With flxSP
                .Clear
                .Rows = 2
                .FormatString = "  |<PART NUMBER|<DESCRIPTION|^QH |^BO|^SQ|^ST|>COST |>ORDER |  |<APPLICABLE MODEL"
                .ColWidth(0) = 400
                .ColWidth(1) = (.Width - 400) * 0.15
                .ColWidth(2) = (.Width - 400) * 0.4
                .ColWidth(3) = (.Width - 400) * 0.07
                .ColWidth(4) = (.Width - 400) * 0.07
                .ColWidth(5) = (.Width - 400) * 0.07
                .ColWidth(6) = (.Width - 400) * 0.07
                .ColWidth(7) = (.Width - 400) * 0.1
                .ColWidth(8) = (.Width - 400) * 0.1
                .ColWidth(9) = 0
                .ColWidth(10) = (.Width - 400) * 0.1
                .ColWidth(11) = 0
                .Visible = True
            End With
        End If
        .Close
    End With
    With flxSP
        .ColWidth(9) = 0
        .ColWidth(11) = 0
        For i = 1 To .Rows - 1
            pBar.Value = IIf(pBar.Value = pBar.Max, pBar.Max, pBar.Value + 1)
            .TextMatrix(i, 5) = GetSalesQty(Val(.TextMatrix(i, 9)))
            'Qry = "SELECT Sum(SPITR.OUT) AS SumOfOUT From SPITR " & _
                  "WHERE (((SPITR.TRANSDATE)>=#" & SDfrom.Value & "# And (SPITR.TRANSDATE)<=#" & SDto.Value & "#) AND ((SPITR.TRANSTYPE)='OUT') AND ((SPITR.REMARKS)='TRANSFER') AND ((SPITR.SPAREPARTID)=" & Val(.TextMatrix(i, 9)) & "));"
            'mySQL
            Qry = "SELECT Sum(SPITR.OUT) AS SumOfOUT From SPITR " & _
                  "WHERE SPITR.TRANSDATE>='" & Format(SDfrom.Value, "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(SDto.Value, "yyyy-mm-dd") & "' AND SPITR.TRANSTYPE='OUT' AND SPITR.REMARKS='TRANSFER' AND SPITR.SPAREPARTID=" & Val(.TextMatrix(i, 9)) & ";"
            
            .TextMatrix(i, 6) = GetFieldValue(Qry)
            If .TextMatrix(i, 11) = "INACTIVE" Then
                .Col = 1: .Row = i: .ColSel = 11: .RowSel = i: .FillStyle = flexFillRepeat: .CellForeColor = RGB(255, 0, 0)
            End If
        Next i
    End With
    pBar.Visible = False

    Call AddInCombo
    If flxSP.Rows >= 2 Then
        If flxSP.TextMatrix(1, 3) <> "" Then
            flxSP.SetFocus
        Else
            cboFind.SetFocus
        End If
    Else
        cboFind.SetFocus
    End If
End Sub

Function GetSalesQty(MyID As Long) As Long
    Dim rsQty As New ADODB.Recordset
    Dim Qry As String
    Dim MySales As Long, MyRet As Long
    
    'old Qry = "SELECT Sum(SPITR.OUT) AS SumOfOUT From SPITR " & _
          "WHERE (((SPITR.TRANSDATE)>=#" & SDfrom.Value & "# And (SPITR.TRANSDATE)<=#" & SDto.Value & "#) AND ((SPITR.TRANSTYPE)='OUT') AND ((SPITR.SPAREPARTID)=" & SDto.Value & "));"
    'Qry = "SELECT Sum(SPITR.OUT) AS SumOfOUT From SPITR " & _
          "WHERE (((SPITR.TRANSDATE)>=#" & SDfrom.Value & "# And (SPITR.TRANSDATE)<=#" & SDto.Value & "#) AND ((SPITR.TRANSTYPE)='OUT') AND ((SPITR.REMARKS)<>'TRANSFER') AND ((SPITR.SPAREPARTID)=" & MyID & "));"
    'mySQL
    Qry = "SELECT Sum(SPITR.OUT) AS SumOfOUT From SPITR " & _
          "WHERE SPITR.TRANSDATE>='" & Format(SDfrom.Value, "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(SDto.Value, "yyyy-mm-dd") & "' AND SPITR.TRANSTYPE='OUT' AND SPITR.REMARKS<>'TRANSFER' AND SPITR.SPAREPARTID=" & MyID & ";"
   
    With rsQty
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            MySales = IIf(IsNull(.Fields(0)), 0, .Fields(0))
        Else
            MySales = 0
        End If
        .Close
    End With
    'Qry = "SELECT Sum(SPITR.RECEIVE) AS SumOfRECEIVE From SPITR " & _
          "WHERE (((SPITR.SPAREPARTID)=" & MyID & ") AND ((SPITR.TRANSTYPE)='SALES RETURN') AND ((SPITR.TRANSDATE)>=#" & SDfrom.Value & "# And (SPITR.TRANSDATE)<=#" & SDto.Value & "#));"
    'mySQL
    Qry = "SELECT Sum(SPITR.RECEIVE) AS SumOfRECEIVE From SPITR " & _
          "WHERE SPITR.SPAREPARTID=" & MyID & " AND SPITR.TRANSTYPE='SALES RETURN' AND SPITR.TRANSDATE>='" & Format(SDfrom.Value, "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(SDto.Value, "yyyy-mm-dd") & "';"
  
    MyRet = Val(GetFieldValue(Qry))
    GetSalesQty = MySales - MyRet
End Function

Private Sub AddInCombo()
    Dim X As Integer, w As Boolean
    
    If cboFind.Text = "" Then Exit Sub
    w = False
    For X = 0 To cboFind.ListCount - 1
        If cboFind.Text = cboFind.List(X) Then
            If X = 0 Then
                w = True
                Exit For
            Else
                cboFind.RemoveItem X
                Exit For
            End If
        End If
    Next X
    If w = False Then cboFind.AddItem cboFind.Text, 0
End Sub

Private Sub cmdInsert_Click()
    Dim i As Integer
    
    
    If MySource = "SPORDERING" Then
        With SPordering.flxParts
            If .TextMatrix(.Rows - 1, 1) = "" Then  'Remove last row if empty
                If .Rows = 2 Then
                    .Rows = 1
                Else
                    .RemoveItem .Rows - 1
                End If
            End If
            For i = 1 To flxSP.Rows - 1
                If flxSP.TextMatrix(i, 0) <> "" Then
                    If Not SPinFlx(flxSP.TextMatrix(i, 1)) Then
                        '.AddItem .Rows & vbTab & flxSP.TextMatrix(i, 1) & vbTab & flxSP.TextMatrix(i, 2) & vbTab & flxSP.TextMatrix(i, 8) & vbTab & flxSP.TextMatrix(i, 7) & vbTab & vbTab & flxSP.TextMatrix(i, 9) & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & .TextMatrix(i, 13)
                        'Add applicable model
                        .AddItem .Rows & vbTab & flxSP.TextMatrix(i, 1) & vbTab & flxSP.TextMatrix(i, 2) & vbTab & flxSP.TextMatrix(i, 8) & vbTab & flxSP.TextMatrix(i, 7) & vbTab & vbTab & flxSP.TextMatrix(i, 9) & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & flxSP.TextMatrix(i, 10)
                        
                    End If
                End If
            Next i
            If .Rows = 1 Then .AddItem "1"
        End With
    End If
    If MySource = "RECEIVING" Then
        With SPincoming.flxOrd
            If .TextMatrix(.Rows - 1, 7) = "" And .TextMatrix(.Rows - 1, 2) = "" Then 'Remove last row if empty
                If .Rows = 2 Then
                    .Rows = 1
                Else
                     .RemoveItem .Rows - 1
                End If
            End If
            For i = 1 To flxSP.Rows - 1
                If flxSP.TextMatrix(i, 0) <> "" Then
                    If Not SPinFlx(flxSP.TextMatrix(i, 9)) Then
                        .AddItem vbTab & vbTab & vbTab & flxSP.TextMatrix(i, 2) & vbTab & flxSP.TextMatrix(i, 3) & vbTab & flxSP.TextMatrix(i, 5) & vbTab & vbTab & flxSP.TextMatrix(i, 9) & vbTab & flxSP.TextMatrix(i, 1)
                    End If
                End If
            Next i
            If .Rows = 1 Then .Rows = 2
        End With
    End If
End Sub

Function SPinFlx(myPN As String) As Boolean
    Dim i As Integer
    
    If MySource = "SPORDERING" Then
        With SPordering.flxParts
            For i = 1 To .Rows - 1
                If myPN = .TextMatrix(i, 1) Then
                    SPinFlx = True
                    Exit Function
                End If
            Next i
        End With
        SPinFlx = False
        Exit Function
    End If
    If MySource = "RECEIVING" Then
        With SPincoming.flxOrd
            For i = 1 To .Rows - 1
                If MyID = .TextMatrix(i, 7) Then
                    SPinFlx = True
                    Exit Function
                End If
            Next i
        End With
        SPinFlx = False
        Exit Function
    End If
End Function

Private Sub cmdOK_Click()
    Call cmdInsert_Click
    Call SPordering.ComputeTotal
    Unload Me
End Sub

Private Sub cmdOptHide_Click()
    FrameOpt.Visible = False
    cmdOptHide.Left = cmdOptShow.Left
    cmdOptHide.Visible = False
    cmdOptShow.Visible = True
    flxSP.Top = 1080
    cmdInsert.Top = flxSP.Top + flxSP.Height + 120
    cmdOK.Top = flxSP.Top + flxSP.Height + 120
    cmdCancel.Top = flxSP.Top + flxSP.Height + 120
    pBar.Top = flxSP.Top + flxSP.Height
    Me.Height = cmdOK.Top + cmdOK.Height + 580
    'txtModel.Text = ""
    ShowOpt = False
End Sub

Private Sub cmdOptShow_Click()
    FrameOpt.Visible = True
    cmdOptHide.Left = cmdOptShow.Left
    cmdOptHide.Visible = True
    cmdOptShow.Visible = False
    flxSP.Visible = True
    flxSP.Top = 1920
    cmdInsert.Top = flxSP.Top + flxSP.Height + 120
    cmdOK.Top = flxSP.Top + flxSP.Height + 120
    cmdCancel.Top = flxSP.Top + flxSP.Height + 120
    pBar.Top = flxSP.Top + flxSP.Height
    Me.Height = cmdOK.Top + cmdOK.Height + 580
    ShowOpt = True
End Sub

Private Sub cmdReplenish_Click()
    Dim rsRpnsh As New ADODB.Recordset
    Dim Qry As String
    
    cmdReplenish.Enabled = False
    cmdInsert.Enabled = False
    cmdOK.Enabled = False
    cmdCancel.Enabled = False
    If txtSuppCode.Text = "" Then
        'Qry = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPAREPARTS.QTYONHAND, SPAREPARTS.BACKORDER, SPAREPARTS.SALES, SPAREPARTS.COST, SPAREPARTS.SPAREPARTID, SPAREPARTS.APPLICMODEL, SPAREPARTS.DESCRIPTION2 AS MYSTAT " & _
              "From SPAREPARTS WHERE (((SPAREPARTS.QTYONHAND)" & lblOprQH.Caption & Val(txtQH.Text) & ") AND ((SPAREPARTS.BACKORDER)" & lblOprBO.Caption & Val(txtBO.Text) & ")) ORDER BY SPAREPARTS.PARTNUMBER;"
        'mySQL
        Qry = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPAREPARTS.QTYONHAND, SPAREPARTS.BACKORDER, SPAREPARTS.SALES, SPAREPARTS.COST, SPAREPARTS.SPAREPARTID, SPAREPARTS.APPLICMODEL, SPAREPARTS.DESCRIPTION2 AS MYSTAT " & _
              "From SPAREPARTS WHERE SPAREPARTS.QTYONHAND" & lblOprQH.Caption & Val(txtQH.Text) & " AND SPAREPARTS.BACKORDER" & lblOprBO.Caption & Val(txtBO.Text) & " ORDER BY SPAREPARTS.PARTNUMBER;"

    Else
        'Qry = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPAREPARTS.QTYONHAND, SPAREPARTS.BACKORDER, SPAREPARTS.SALES, SPAREPARTS.COST, SPAREPARTS.SPAREPARTID, SPAREPARTS.APPLICMODEL, SPAREPARTS.DESCRIPTION2 AS MYSTAT " & _
              "From SPAREPARTS WHERE (((SPAREPARTS.QTYONHAND)" & lblOprQH.Caption & Val(txtQH.Text) & ") AND ((SPAREPARTS.BACKORDER)" & lblOprBO.Caption & Val(txtBO.Text) & ") AND ((SPAREPARTS.SUPPLIERCODE)='" & txtSuppCode.Text & "')) ORDER BY SPAREPARTS.PARTNUMBER;"
        'mySQL
        Qry = "SELECT SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPAREPARTS.QTYONHAND, SPAREPARTS.BACKORDER, SPAREPARTS.SALES, SPAREPARTS.COST, SPAREPARTS.SPAREPARTID, SPAREPARTS.APPLICMODEL, SPAREPARTS.DESCRIPTION2 AS MYSTAT " & _
              "From SPAREPARTS WHERE SPAREPARTS.QTYONHAND" & lblOprQH.Caption & Val(txtQH.Text) & " AND SPAREPARTS.BACKORDER" & lblOprBO.Caption & Val(txtBO.Text) & " AND SPAREPARTS.SUPPLIERCODE='" & txtSuppCode.Text & "' ORDER BY SPAREPARTS.PARTNUMBER;"

       
    End If
    With rsRpnsh
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            flxSP.Rows = 1
            pBar.Value = 0
            pBar.Visible = True
            pBar.Max = .RecordCount * 2
            Do While Not .EOF
                      '1-PART NUMBER, 2-DESCRIPTION, 3- QH, 4- BO, 5-SQ, 6-COST, 7-ORDER, 8-SPID, 9, applicmodel, 10-MYSTAT
                      '1-PART NUMBER, 2-DESCRIPTION, 3- QH, 4- BO, 5-SQ, 6-ST, 7-COST, 8-ORDER, 9-SPID, 10- applicmodel, 11-MYSTAT
                flxSP.AddItem vbTab & !PartNumber & vbTab & !Description & vbTab & !QtyOnHand & vbTab & IIf(IsNull(!BACKORDER), "", !BACKORDER) & vbTab & IIf(IsNull(!Sales), "", !Sales) & vbTab & vbTab & Format(!Cost, "#,##0.00") & vbTab & vbTab & !SPAREPARTID & vbTab & IIf(IsNull(!ApplicModel), "", !ApplicModel) & vbTab & IIf(IsNull(!MyStat), "", !MyStat)
                pBar.Value = pBar.Value + 1
                .MoveNext
                DoEvents
            Loop
            flxSP.Visible = True
            If ShowOpt Then
                flxSP.Top = 1920
            Else
                flxSP.Top = 1080
            End If
            cmdInsert.Top = flxSP.Top + flxSP.Height + 120
            cmdOK.Top = flxSP.Top + flxSP.Height + 120
            cmdCancel.Top = flxSP.Top + flxSP.Height + 120
            pBar.Top = flxSP.Top + flxSP.Height
            Me.Height = cmdOK.Top + cmdOK.Height + 580
        Else
            With flxSP
                .Clear
                .Rows = 2
                      '1-PART NUMBER, 2-DESCRIPTION, 3- QH, 4- BO, 5-SQ, 6-ST, 7-COST, 8-ORDER, 9-SPID, 10- applicmodel, 11-MYSTAT
                .FormatString = "  |<PART NUMBER|<DESCRIPTION|^QH |^BO |^SQ |^ST |>COST |>ORDER |  |<APPLICABLE MODEL|"
                .ColWidth(0) = 400
                .ColWidth(1) = (.Width - 400) * 0.15
                .ColWidth(2) = (.Width - 400) * 0.4
                .ColWidth(3) = (.Width - 400) * 0.07
                .ColWidth(4) = (.Width - 400) * 0.07
                .ColWidth(5) = (.Width - 400) * 0.07
                .ColWidth(6) = (.Width - 400) * 0.07
                .ColWidth(7) = (.Width - 400) * 0.1
                .ColWidth(8) = (.Width - 400) * 0.1
                .ColWidth(9) = 0
                .ColWidth(10) = (.Width - 400) * 0.1
                .ColWidth(11) = 0
                .Visible = True
            End With
        End If
        .Close
    End With
    flxSP.ColWidth(9) = 0
    flxSP.ColWidth(11) = 0
    For i = flxSP.Rows - 1 To 1 Step -1
        flxSP.TextMatrix(i, 5) = GetSalesQty(Val(flxSP.TextMatrix(i, 9)))
        'Qry = "SELECT Sum(SPITR.OUT) AS SumOfOUT From SPITR " & _
              "WHERE (((SPITR.TRANSDATE)>=#" & SDfrom.Value & "# And (SPITR.TRANSDATE)<=#" & SDto.Value & "#) AND ((SPITR.TRANSTYPE)='OUT') AND ((SPITR.REMARKS)='TRANSFER') AND ((SPITR.SPAREPARTID)=" & Val(flxSP.TextMatrix(i, 9)) & "));"
        'mySQL
        Qry = "SELECT Sum(SPITR.OUT) AS SumOfOUT From SPITR " & _
              "WHERE SPITR.TRANSDATE>='" & Format(SDfrom.Value, "yyyy-mm-dd") & "' And SPITR.TRANSDATE<='" & Format(SDto.Value, "yyyy-mm-dd") & "' AND SPITR.TRANSTYPE='OUT' AND SPITR.REMARKS='TRANSFER' AND SPITR.SPAREPARTID=" & Val(flxSP.TextMatrix(i, 9)) & ";"
        
        flxSP.TextMatrix(i, 6) = GetFieldValue(Qry)
        pBar.Value = IIf(pBar.Value = pBar.Max, pBar.Max, pBar.Value + 1)
        DoEvents
        If flxSP.TextMatrix(i, 11) = "INACTIVE" Then
            flxSP.Col = 1: flxSP.Row = i: flxSP.ColSel = 11: flxSP.RowSel = i: flxSP.FillStyle = flexFillRepeat: flxSP.CellForeColor = RGB(255, 0, 0)
        End If
        Select Case OptrSQ.Value
            Case 0
                If Val(flxSP.TextMatrix(i, 5)) <> Val(txtSQ.Text) Then 'flxSP.TextMatrix(i, 9) = "x"
                    If flxSP.Rows = 2 Then
                        flxSP.Rows = 1
                        flxSP.Rows = 2
                    Else
                        flxSP.RemoveItem i
                    End If
                End If
            Case 1
                If Val(flxSP.TextMatrix(i, 5)) <= Val(txtSQ.Text) Then
                    If flxSP.Rows = 2 Then
                        flxSP.Rows = 1
                        flxSP.Rows = 2
                    Else
                        flxSP.RemoveItem i
                    End If
                End If
            Case 2
                If Val(flxSP.TextMatrix(i, 5)) < Val(txtSQ.Text) Then
                    If flxSP.Rows = 2 Then
                        flxSP.Rows = 1
                        flxSP.Rows = 2
                    Else
                        flxSP.RemoveItem i
                    End If
                End If
            Case 3
                If Val(flxSP.TextMatrix(i, 5)) >= Val(txtSQ.Text) Then
                    If flxSP.Rows = 2 Then
                        flxSP.Rows = 1
                        flxSP.Rows = 2
                    Else
                        flxSP.RemoveItem i
                    End If
                End If
            Case 4
                If Val(flxSP.TextMatrix(i, 5)) > Val(txtSQ.Text) Then
                    If flxSP.Rows = 2 Then
                        flxSP.Rows = 1
                        flxSP.Rows = 2
                    Else
                        flxSP.RemoveItem i
                    End If
                End If
            Case 5
                If Val(flxSP.TextMatrix(i, 5)) = Val(txtSQ.Text) Then
                    If flxSP.Rows = 2 Then
                        flxSP.Rows = 1
                        flxSP.Rows = 2
                    Else
                        flxSP.RemoveItem i
                    End If
                End If
        End Select
    Next i
    Call AddInCombo
    pBar.Visible = False
    If flxSP.Rows >= 2 Then
        If flxSP.TextMatrix(1, 3) <> "" Then
            flxSP.SetFocus
        Else
            cboFind.SetFocus
        End If
    Else
        cboFind.SetFocus
    End If
        
   cmdReplenish.Enabled = True
   cmdInsert.Enabled = True
   cmdOK.Enabled = True
   cmdCancel.Enabled = True
End Sub

Private Sub flxSP_Click()

    ActRow = flxSP.Row
    ActCol = flxSP.Col
    If ActRow = 0 Or ActCol <> 8 Then Exit Sub
    With txtOrder
        .Top = flxSP.CellTop + flxSP.Top
        .Left = flxSP.CellLeft + flxSP.Left
        .Width = flxSP.CellWidth
        .Text = Trim(flxSP.Text)
        .Visible = True
        .SelStart = 0
        .SelLength = Len(.Text)
        .ZOrder
        .SetFocus
    End With
    
End Sub

Private Sub flxSP_DblClick()
        If flxSP.TextMatrix(flxSP.Row, 0) = "" And Val(flxSP.TextMatrix(flxSP.Row, 1)) > 0 Then
            Call mnuMark_Click
        Else
            Call mnuUnmark_Click
        End If
End Sub

Private Sub flxSP_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
        If flxSP.TextMatrix(flxSP.Row, 0) = "" And Val(flxSP.TextMatrix(flxSP.Row, 1)) > 0 Then
            Call mnuMark_Click
        Else
            Call mnuUnmark_Click
        End If
    End If
End Sub


Private Sub flxSP_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call flxSP_Click
End Sub

Private Sub flxSP_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
    If Val(flxSP.TextMatrix(flxSP.Row, 1)) = 0 Then Exit Sub
    If Button = vbRightButton Then
        If flxSP.TextMatrix(flxSP.Row, 11) = "INACTIVE" Then
            mnuInAct.Visible = False
            mnuAct.Visible = True
        Else
            mnuInAct.Visible = True
            mnuAct.Visible = False
        End If
        Call PopupMenu(mnuSP)
    End If
End Sub

Private Sub flxSP_Scroll()
    txtOrder.Visible = False
End Sub

Private Sub Form_Load()
    Dim Qry As String
    
    With flxSP
        .ColWidth(0) = 400
        .ColWidth(1) = (.Width - 400) * 0.15
        .ColWidth(2) = (.Width - 400) * 0.4
        .ColWidth(3) = (.Width - 400) * 0.07
        .ColWidth(4) = (.Width - 400) * 0.07
        .ColWidth(5) = (.Width - 400) * 0.07
        .ColWidth(6) = (.Width - 400) * 0.07
        .ColWidth(7) = (.Width - 400) * 0.1
        .ColWidth(8) = (.Width - 400) * 0.1
        .ColWidth(9) = 0
        .ColWidth(10) = (.Width - 400) * 0.1
        .ColWidth(11) = 0
    End With
    
    Call cmdOptHide_Click
    
    With SPordering
        If .LockFlx = "True" Then
            cmdInsert.Enabled = False
            cmdOK.Enabled = False
        End If
        SDfrom.Value = .SalesDateFrom.Value
        SDto.Value = .SalesDateTo.Value
    End With
    
    Qry = "SELECT SUPPLIER.SUPPLIERID, SUPPLIER.SUPPLIERNAME, SUPPLIER.SUPPLIERCODE " & _
          "From SUPPLIER ORDER BY SUPPLIER.SUPPLIERNAME;"
    With rsSupplier
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            cboSupplier.Clear
            Do While Not .EOF
                cboSupplier.AddItem !SUPPLIERNAME
                .MoveNext
            Loop
        End If
    End With
    
    cboSupplier.Text = SPordering.cboSupplier.Text
    txtSuppCode.Text = SPordering.txtScode.Text
    'Set Search Description
    UpDRef.Value = 1
    'LYRmain.sBar.Panels(6).Text = "Find Spare Parts"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'LYRmain.sBar.Panels(6).Text = ""
    If rsSupplier.State = adStateOpen Then rsSupplier.Close
End Sub


Private Sub mnuAct_Click()
    Dim Ans As Integer
    
    Ans = MsgBox("Do you want ACTIVATE Partnumber " & flxSP.TextMatrix(flxSP.Row, 1) & "?", vbQuestion + vbYesNo, "In-Active")
    If Ans = vbNo Then Exit Sub
    
    flxSP.Col = 1:  flxSP.ColSel = 11: flxSP.RowSel = flxSP.Row: flxSP.FillStyle = flexFillRepeat: flxSP.CellForeColor = RGB(0, 0, 0)
    flxSP.TextMatrix(flxSP.Row, 11) = ""
    
    'LYRcon.Execute "UPDATE SPAREPARTS SET SPAREPARTS.DESCRIPTION2 = '' WHERE SPAREPARTS.SPAREPARTID=" & Val(flxSP.TextMatrix(flxSP.Row, 9)) & ";"
    'mySQL
    LYRcon.Execute "UPDATE SPAREPARTS SET SPAREPARTS.DESCRIPTION2 = '' WHERE SPAREPARTS.SPAREPARTID=" & Val(flxSP.TextMatrix(flxSP.Row, 9)) & ";"

End Sub

Private Sub mnuInAct_Click()
    Dim Ans As Integer
    
    Ans = MsgBox("Do you want Partnumber " & flxSP.TextMatrix(flxSP.Row, 1) & " to be IN-ACTIVE?", vbQuestion + vbYesNo, "In-Active")
    If Ans = vbNo Then Exit Sub
    
    flxSP.Col = 1:  flxSP.ColSel = 11: flxSP.RowSel = flxSP.Row: flxSP.FillStyle = flexFillRepeat: flxSP.CellForeColor = RGB(255, 0, 0)
    flxSP.TextMatrix(flxSP.Row, 11) = "INACTIVE"
    
    'LYRcon.Execute "UPDATE SPAREPARTS SET SPAREPARTS.DESCRIPTION2 = 'INACTIVE' WHERE (((SPAREPARTS.SPAREPARTID)=" & Val(flxSP.TextMatrix(flxSP.Row, 9)) & "));"
    'mySQL
    LYRcon.Execute "UPDATE SPAREPARTS SET SPAREPARTS.DESCRIPTION2 = 'INACTIVE' WHERE SPAREPARTS.SPAREPARTID=" & Val(flxSP.TextMatrix(flxSP.Row, 9)) & ";"
    
End Sub

Private Sub mnuMark_Click()
    With flxSP
        .Col = 0
        .CellFontName = "WINGDINGS"
        .CellFontSize = 12
        .Text = lblChk.Caption
        .CellAlignment = 4
    End With
End Sub

Private Sub mnuMarkAll_Click()
    Dim i As Integer
    
    With flxSP
        For i = 1 To .Rows - 1
            .Row = i
            .Col = 0
            .CellFontName = "WINGDINGS"
            .CellFontSize = 12
            .Text = lblChk.Caption
            .CellAlignment = 4
            .TextMatrix(i, 0) = lblChk.Caption
        Next i
    End With
    
End Sub

Private Sub mnuUnmark_Click()
    flxSP.TextMatrix(flxSP.Row, 0) = ""
End Sub

Private Sub mnuUnmarkAll_Click()
    Dim i As Integer
    
    For i = 1 To flxSP.Rows - 1
        flxSP.TextMatrix(i, 0) = ""
    Next i
End Sub


Private Sub OptrBO_Change()
    Select Case OptrBO.Value
        Case 0
            lblOprBO.Caption = "="
        Case 1
            lblOprBO.Caption = ">"
        Case 2
            lblOprBO.Caption = ">="
        Case 3
            lblOprBO.Caption = "<"
        Case 4
            lblOprBO.Caption = "<="
        Case 5
            lblOprBO.Caption = "<>"
    End Select

End Sub

Private Sub OptrQH_Change()
    Select Case OptrQH.Value
        Case 0
            lblOprQH.Caption = "="
        Case 1
            lblOprQH.Caption = ">"
        Case 2
            lblOprQH.Caption = ">="
        Case 3
            lblOprQH.Caption = "<"
        Case 4
            lblOprQH.Caption = "<="
        Case 5
            lblOprQH.Caption = "<>"
    End Select
End Sub

Private Sub OptrSQ_Change()
    Select Case OptrSQ.Value
        Case 0
            lblOprSQ.Caption = "="
        Case 1
            lblOprSQ.Caption = ">"
        Case 2
            lblOprSQ.Caption = ">="
        Case 3
            lblOprSQ.Caption = "<"
        Case 4
            lblOprSQ.Caption = "<="
        Case 5
            lblOprSQ.Caption = "<>"
    End Select
End Sub

Private Sub txtOrder_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        txtOrder.Text = flxSP.TextMatrix(ActRow, 8)
        flxSP.SetFocus
        flxSP.Row = ActRow
        flxSP.Col = 8
        txtOrder.Visible = False
    End If
End Sub

Private Sub txtOrder_KeyPress(KeyAscii As Integer)
    Dim strValid As String
    
    strValid = "1234567890"
    
    If KeyAscii > 26 Then
        If InStr(strValid, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
    If KeyAscii = 13 Then
        flxSP.SetFocus
        flxSP.Row = ActRow
        flxSP.Col = 8
    End If
End Sub

Private Sub txtOrder_LostFocus()
    
    With flxSP
        .TextMatrix(ActRow, 8) = txtOrder.Text
        txtOrder.Visible = False
        If Val(.TextMatrix(ActRow, 8)) > 0 Then
            .Row = ActRow
            .Col = 0
            .CellFontName = "WINGDINGS"
            .CellFontSize = 12
            .Text = lblChk.Caption
            .CellAlignment = 4
            .Col = 8
        Else
            .TextMatrix(ActRow, 8) = ""
            .TextMatrix(ActRow, 0) = ""
            .Col = 8
        End If
    End With
End Sub


Private Sub UpDRef_Change()
    If UpDRef.Value = 0 Then
        lblRefSearch.Caption = "FIND PARTNUMBER"
    Else
        lblRefSearch.Caption = "FIND PARTS DESC."
    End If
   
End Sub
