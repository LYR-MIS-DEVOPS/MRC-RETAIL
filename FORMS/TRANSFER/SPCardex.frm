VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form SPCardex 
   Caption         =   "CARDEX"
   ClientHeight    =   9405
   ClientLeft      =   3000
   ClientTop       =   1200
   ClientWidth     =   7185
   Icon            =   "SPCardex.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9405
   ScaleWidth      =   7185
   Begin MSComctlLib.ProgressBar prgBar 
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   9120
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.TextBox txtSPid 
      Height          =   285
      Left            =   5160
      TabIndex        =   16
      Top             =   9480
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   5160
      Top             =   240
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxParts 
      Height          =   1815
      Left            =   840
      TabIndex        =   13
      Top             =   6000
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   3201
      _Version        =   393216
      BackColor       =   16777215
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
      Height          =   1575
      Left            =   120
      TabIndex        =   9
      Tag             =   "XYF"
      Top             =   720
      Width           =   6975
      Begin VB.TextBox txtLocation 
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
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   3
         Tag             =   "XYF"
         Top             =   1080
         Width           =   4095
      End
      Begin VB.TextBox txtPrice 
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
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   2
         Tag             =   "XYF"
         Top             =   1080
         Width           =   2535
      End
      Begin VB.TextBox txtDesc 
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
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   1
         Tag             =   "XYF"
         Top             =   480
         Width           =   4095
      End
      Begin VB.TextBox txtPartNo 
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
         Left            =   120
         TabIndex        =   0
         Tag             =   "XYF"
         Top             =   480
         Width           =   2535
      End
      Begin VB.Label Label5 
         Caption         =   "LOCATION"
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
         TabIndex        =   15
         Tag             =   "XYF"
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "PRICE"
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
         TabIndex        =   12
         Tag             =   "XYF"
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "DESCRIPTION"
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
         TabIndex        =   11
         Tag             =   "XYF"
         Top             =   240
         Width           =   3255
      End
      Begin VB.Label Label1 
         Caption         =   "PART NUMBER"
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
         TabIndex        =   10
         Tag             =   "XYF"
         Top             =   240
         Width           =   2415
      End
   End
   Begin MSFlexGridLib.MSFlexGrid flxStock 
      Height          =   6735
      Left            =   120
      TabIndex        =   4
      Tag             =   "XYF"
      Top             =   2400
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   11880
      _Version        =   393216
      Cols            =   5
      FixedCols       =   0
      RowHeightMin    =   315
      BackColorBkg    =   16777215
      ScrollBars      =   2
      FormatString    =   "^    Date    |^   Reference     |^ In  |^ Out |^  Balance   "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin lvButton.lvButtons_H cmdExit 
      Height          =   375
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
      Caption         =   "E&xit Cardex"
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
      Image           =   "SPCardex.frx":058A
      cBack           =   -2147483633
   End
   Begin VB.Label lblPrceRec 
      Height          =   255
      Left            =   6600
      TabIndex        =   18
      Top             =   9120
      Width           =   495
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   3960
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   7425
      m_MDI_Height    =   9975
      OriginalBorder  =   10
      OriginalCaption =   31
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "CARDEX"
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
      Left            =   4080
      TabIndex        =   8
      Top             =   480
      Width           =   3015
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "CARDEX"
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
      Left            =   4110
      TabIndex        =   7
      Top             =   510
      Width           =   3015
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "MOTO-R REPUBLIC"
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
      Index           =   3
      Left            =   120
      TabIndex        =   6
      Top             =   480
      Width           =   2415
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "MOTO-R REPUBLIC"
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
      Index           =   2
      Left            =   150
      TabIndex        =   5
      Top             =   510
      Width           =   2415
   End
   Begin VB.Menu mnuStockCard 
      Caption         =   "StockCard"
      Visible         =   0   'False
   End
End
Attribute VB_Name = "SPCardex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsParts As New ADODB.Recordset
Dim PrtSrch As Boolean, PrtBln As Boolean, FromFlx As Boolean
Public PNadj As String
Private Sub LoadSpareParts()
    Dim rsSet As New ADODB.Recordset
    Dim Qry As String
    Dim pPct As Single
        
    Qry = "SELECT SPAREPARTS.SPAREPARTID, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPAREPARTS.STOREPRICE, SPAREPARTS.QTYONHAND, SPAREPARTS.LOCATION " & _
          "From SPAREPARTS ORDER BY SPAREPARTS.PARTNUMBER;"
    rsParts.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    Set flxParts.DataSource = rsParts
    With flxParts
        .Width = txtPartNo.Width
        .ColWidth(0) = 0
        .ColWidth(1) = 0
        .ColWidth(2) = txtPartNo.Width
        .ColWidth(3) = 0
        .ColWidth(4) = 0
        .ColWidth(5) = 0
        .ColWidth(6) = 0
    End With
End Sub

Private Sub cmdExit_Click()
    If PNadj <> "" Then
        Me.Hide
        PNadj = ""
    Else
        Unload Me
    End If
    'MotoRep.tbMain.Visible = True
End Sub


Private Sub flxParts_Click()
    FromFlx = True
End Sub

Private Sub flxParts_EnterCell()
    With flxParts
        FromFlx = True
        txtPartNo.Text = .TextMatrix(.Row, 2)
    End With
End Sub

Private Sub flxParts_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        flxParts.Visible = False
        flxStock.SetFocus
        flxStock.Row = flxStock.Rows - 1
        flxStock.Col = 4
        flxStock.RowSel = flxStock.Rows - 1
        flxStock.ColSel = 0
        Call DisplayCardex(flxParts.TextMatrix(flxParts.Row, 1))
    End If
End Sub

Private Sub flxParts_LostFocus()
    flxParts.Visible = False
End Sub

Private Sub Form_Activate()
    If PNadj <> "" Then
        If PNadj <> "INVADJ" Then
            txtPartNo.Text = PNadj
            flxParts.Visible = False
        Else
            txtPartNo.Text = ""
            txtDesc.Text = ""
            txtPrice.Text = ""
            txtLocation.Text = ""
            flxStock.Rows = 1
            flxStock.Rows = 2
        End If
    End If
End Sub

Private Sub Form_Load()
        
    'Call LoadSpareParts
    Timer1.Enabled = True
    
    'MotoRep.sBar.Panels(6).Text = "Spare Parts CARDEX"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If rsParts.State = adStateOpen Then rsParts.Close
End Sub

Private Sub RXResizer1_ResizeComplete()
    With flxStock
        .ColWidth(0) = .Width * 0.18
        .ColWidth(1) = .Width * 0.27
        .ColWidth(2) = .Width * 0.18
        .ColWidth(3) = .Width * 0.18
        .ColWidth(4) = .Width * 0.18
    End With
    flxParts.ColWidth(2) = txtPartNo.Width
    'flxParts.RowHeightMin = txtPartNo.Height
    flxParts.Font.Size = txtPartNo.FontSize
End Sub

Private Sub DisplayCardex(MyID As Long)
    Dim rsCardex As New ADODB.Recordset
    Dim Qry As String
    'Dim myIN As Integer, myOUT As Integer, PrvBal As Integer
    Dim myIN As Long, myOUT As Long, PrvBal As Long
    
    'Clear FlxStock
    flxStock.Rows = 1
    
    '***Access
    'Qry = "SELECT SPITR.TRANSDATE, SPITR.REFERENCE, SPITR.RECEIVE, SPITR.`OUT`, SPITR.BALANCE, SPITR.TRANSTYPE, SPITR.REMARKS From SPITR " & _
          "WHERE SPITR.TRANSDATE>=#7/31/2007# AND ((SPITR.TRANSTYPE)='RECEIVE' Or (SPITR.TRANSTYPE)='BEGBAL' Or (SPITR.TRANSTYPE)='SALES RETURN' Or (SPITR.TRANSTYPE)='INVENTORY ADJ.' Or (SPITR.TRANSTYPE)='OUT' Or (SPITR.TRANSTYPE)='PURCHASE RETURN') AND ((SPITR.REMARKS) Is Null Or (SPITR.REMARKS)<>'Temp') AND ((SPITR.SPAREPARTID)=" & MyId & ")) " & _
          "ORDER BY SPITR.TRANSDATE, SPITR.LASTUPDATE;"
    Qry = "SELECT " & _
            "SPITR.TRANSDATE," & _
            "SPITR.REFERENCE," & _
            "SPITR.RECEIVE," & _
            "SPITR.`OUT`, " & _
            "SPITR.BALANCE," & _
            "SPITR.TRANSTYPE," & _
            "SPITR.REMARKS " & _
          "From SPITR " & _
          "Where SPITR.TRANSDATE>='2007-07-31' AND " & _
            "(SPITR.TRANSTYPE='RECEIVE' Or " & _
            "SPITR.TRANSTYPE='BEGBAL' Or " & _
            "SPITR.TRANSTYPE='SALES RETURN' Or " & _
            "SPITR.TRANSTYPE='INVENTORY ADJ.' Or " & _
            "SPITR.TRANSTYPE='OUT' Or " & _
            "SPITR.TRANSTYPE='PURCHASE RETURN') AND " & _
            "(SPITR.REMARKS Is Null Or " & _
            "SPITR.REMARKS<>'Temp') AND " & _
            "(SPITR.SPAREPARTID=" & MyID & ") " & _
          "Order By SPITR.TRANSDATE, " & _
            "SPITR.LASTUPDATE"

    With rsCardex
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            PrvBal = 0
            .MoveFirst
             prgBar.Value = 0
             prgBar.Max = .RecordCount
             prgBar.Visible = True
             lblPrceRec.Visible = True
            Do While Not .EOF
                If !TRANSTYPE = "INVENTORY ADJ." Then
                    If !RECEIVE > 0 Then
                        myIN = !RECEIVE
                        myOUT = 0
                    ElseIf !OUT > 0 Then
                        myOUT = !OUT
                        myIN = 0
                    End If
                Else
                    myIN = IIf(!TRANSTYPE = "OUT" Or !TRANSTYPE = "PURCHASE RETURN", 0, !RECEIVE)
                    myOUT = IIf(!TRANSTYPE <> "OUT" And !TRANSTYPE <> "PURCHASE RETURN", 0, !OUT)
                End If
                PrvBal = PrvBal + (myIN - myOUT)
                flxStock.AddItem Format$(!TRANSDATE, "mm/dd/yyyy") & vbTab & !REFERENCE & vbTab & IIf(myIN = 0, "", myIN) & vbTab & IIf(myOUT = 0, "", myOUT) & vbTab & PrvBal
                .MoveNext
                
                prgBar.Value = prgBar.Value + 1
                PrcntRec = prgBar.Value / prgBar.Max
                DoEvents
                lblPrceRec.Caption = Format$(PrcntRec, "##0.0%")
                Loop
            flxStock.TopRow = flxStock.Rows - 1
            flxStock.RowSel = flxStock.Rows - 1
            flxStock.ColSel = 0
            flxStock.Row = flxStock.Rows - 1
            flxStock.Col = 4
        Else
            flxStock.Rows = 1
        End If
        .Close
    End With
    prgBar.Visible = False
    lblPrceRec.Visible = False
End Sub

Private Sub Timer1_Timer()
    Call LoadSpareParts
    Timer1.Enabled = False
End Sub

Private Sub txtPartNo_Change()
    Dim mRow As Integer
    
    With flxParts
        If PrtSrch Then
            PrtSrch = False
            Exit Sub
        End If
        .Top = Frame1.Top + txtPartNo.Height + txtPartNo.Top
        .Left = Frame1.Left + txtPartNo.Left
        .Width = txtPartNo.Width
        .Visible = True
        If txtPartNo.Text = "" Then .Visible = False: Exit Sub
        rsParts.Find "[PARTNUMBER] LIKE '" + txtPartNo.Text + "%'", , adSearchForward, 1
        If Not rsParts.EOF Then
            If Not PrtBln Then
                txtPartNo.Tag = Len(txtPartNo.Text)
                PrtBln = True
            End If
            If Not FromFlx Then     'top row only if user type @ txtpartno
                .TopRow = rsParts.AbsolutePosition - 1
                FromFlx = False
            End If
            .Row = rsParts.AbsolutePosition - 1
            .Col = 0: .ColSel = 3
            txtPartNo.Text = rsParts!PartNumber
            'Insert Additional retrieval here
            txtDesc.Text = flxParts.TextMatrix(flxParts.Row, 3)
            txtPrice.Text = Format$(flxParts.TextMatrix(flxParts.Row, 4), "###,##0.00")
            txtLocation.Text = flxParts.TextMatrix(flxParts.Row, 6)
            txtSPid.Text = flxParts.TextMatrix(flxParts.Row, 1)
            'Call DisplayCardex(flxParts.TextMatrix(flxParts.Row, 1))
            
            txtPartNo.SelStart = Val(txtPartNo.Tag)
            txtPartNo.SelLength = Len(txtPartNo.Text)
            PrtBln = False
        Else
            txtSPid.Text = ""
            txtDesc.Text = ""
            txtPrice.Text = ""
            flxStock.Rows = 1
            .Visible = False
        End If
    End With
    
End Sub

Private Sub txtPartNo_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            cmdExit.SetFocus
            flxParts.Visible = False
        Case vbKeyDelete
            PrtSrch = Truem
        Case vbKeyBack
            PrtSrch = True
        Case vbKeyDown
            flxParts.Visible = True
            flxParts.SetFocus
        Case Else
            FromFlx = False
    End Select

End Sub
Private Sub txtPartNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        flxParts.Visible = False
        flxStock.RowSel = flxStock.Rows - 1
        flxStock.Col = 4
        flxStock.SetFocus
        flxStock.ColSel = 0
        flxStock.Row = flxStock.Rows - 1
        Call DisplayCardex(flxParts.TextMatrix(flxParts.Row, 1))
    End If
End Sub
