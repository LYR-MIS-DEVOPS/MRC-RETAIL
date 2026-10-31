VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form SPRecSummary 
   Caption         =   "Spare Parts Receiving Summary"
   ClientHeight    =   8535
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11865
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8535
   ScaleWidth      =   11865
   WindowState     =   2  'Maximized
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxREC 
      Height          =   7095
      Left            =   120
      TabIndex        =   17
      Tag             =   "XYF"
      Top             =   1350
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   12515
      _Version        =   393216
      RowHeightMin    =   285
      SelectionMode   =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Tag             =   "XY"
      Top             =   720
      Width           =   11655
      Begin VB.TextBox txtFind 
         Height          =   285
         Left            =   9600
         TabIndex        =   6
         Tag             =   "XYF"
         Top             =   195
         Width           =   1935
      End
      Begin lvButton.lvButtons_H cmdRefresh 
         Height          =   375
         Left            =   5400
         TabIndex        =   5
         ToolTipText     =   "Execute"
         Top             =   150
         Width           =   375
         _ExtentX        =   661
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
         Image           =   "SPrecSummary.frx":0000
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdFind 
         Height          =   375
         Left            =   8280
         TabIndex        =   7
         Tag             =   "XY"
         Top             =   150
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         Caption         =   "Find Next"
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
         Image           =   "SPrecSummary.frx":059A
         cBack           =   -2147483633
      End
      Begin MSComCtl2.DTPicker dTo 
         Height          =   315
         Left            =   3720
         TabIndex        =   8
         Top             =   180
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         Format          =   45547521
         CurrentDate     =   39322
      End
      Begin MSComCtl2.DTPicker dFrom 
         Height          =   315
         Left            =   1440
         TabIndex        =   9
         Top             =   180
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         Format          =   45547521
         CurrentDate     =   39322
      End
      Begin lvButton.lvButtons_H cmdNext 
         Height          =   375
         Left            =   3240
         TabIndex        =   15
         ToolTipText     =   "Execute"
         Top             =   150
         Width           =   375
         _ExtentX        =   661
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
         Image           =   "SPrecSummary.frx":0B34
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdPrev 
         Height          =   375
         Left            =   5880
         TabIndex        =   16
         ToolTipText     =   "Execute"
         Top             =   150
         Visible         =   0   'False
         Width           =   375
         _ExtentX        =   661
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
         Image           =   "SPrecSummary.frx":10CE
         cBack           =   -2147483633
      End
      Begin VB.Label Label1 
         Caption         =   "Date Covered:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   210
         Width           =   1335
      End
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   661
      Caption         =   "&Show Packing list"
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
      Image           =   "SPrecSummary.frx":1668
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   1
      Left            =   1800
      TabIndex        =   11
      Top             =   0
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   661
      Caption         =   "&Refresh "
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
      Image           =   "SPrecSummary.frx":1C02
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   2
      Left            =   3600
      TabIndex        =   12
      Top             =   0
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   661
      Caption         =   "&Print Summary"
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
      Image           =   "SPrecSummary.frx":219C
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   3
      Left            =   5400
      TabIndex        =   13
      Top             =   0
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   661
      Caption         =   "E&xit Report"
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
      Image           =   "SPrecSummary.frx":2736
      cBack           =   -2147483633
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   360
      Top             =   8640
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   11985
      m_MDI_Height    =   9045
   End
   Begin VB.Label lblHeading 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Spare Parts Receiving Summary"
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
      Left            =   7200
      TabIndex        =   2
      Top             =   450
      Width           =   4500
   End
   Begin VB.Label lblHeading 
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
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   450
      Width           =   3500
   End
   Begin VB.Label lblHeading 
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
      Index           =   1
      Left            =   150
      TabIndex        =   1
      Top             =   475
      Width           =   3500
   End
   Begin VB.Label lblHeading 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Spare Parts Receiving Summary"
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
      Left            =   7215
      TabIndex        =   3
      Top             =   480
      Width           =   4500
   End
   Begin VB.Menu mnuSPinv 
      Caption         =   "SP Invoice"
      Visible         =   0   'False
      Begin VB.Menu mnuCancel 
         Caption         =   "Cancel Invoice"
      End
   End
End
Attribute VB_Name = "SPRecSummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsSPrec As New ADODB.Recordset
Dim ActRow As Integer, fRow As Integer, fCol As Integer
Dim Qry As String

Private Sub RefreshRecord(dCode As Integer)
    Dim i As Integer
    If dCode = 1 Then
        Qry = "SELECT SPRECEIVE.DATEREC, SPRECEIVE.PACKINGNO, SPRECEIVE.SUPPLIER, SPRECEIVE.SHIPPER, SPRECEIVE.WAYBILLNO, SPRECEIVE.PACKDATE, SPRECEIVE.REMARKS, SPRECEIVE.CHECKEDBY, SPRECEIVE.PRINTCOPY, SPRECEIVE.SPRRNO " & _
              "From SPRECEIVE Where (((SPRECEIVE.DATEREC) = #" & dFrom.Value & "#)) ORDER BY SPRECEIVE.DATEREC, SPRECEIVE.PACKINGNO DESC;"
    Else
        Qry = "SELECT SPRECEIVE.DATEREC, SPRECEIVE.PACKINGNO, SPRECEIVE.SUPPLIER, SPRECEIVE.SHIPPER, SPRECEIVE.WAYBILLNO, SPRECEIVE.PACKDATE, SPRECEIVE.REMARKS, SPRECEIVE.CHECKEDBY, SPRECEIVE.PRINTCOPY, SPRECEIVE.SPRRNO " & _
              "From SPRECEIVE Where (((SPRECEIVE.DATEREC) >= #" & dFrom.Value & "# And (SPRECEIVE.DATEREC) <= #" & dTo.Value & "#)) ORDER BY SPRECEIVE.DATEREC, SPRECEIVE.PACKINGNO DESC;"
    End If
    If rsSPrec.State = adStateOpen Then rsSPrec.Close
    rsSPrec.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    If rsSPrec.RecordCount = 0 Then
        flxREC.Clear
        flxREC.Rows = 2
        flxREC.Cols = 10
        Call RXResizer1_ResizeComplete
        Exit Sub
    End If
    Set flxREC.DataSource = rsSPrec
    flxREC.Visible = True
    Call RXResizer1_ResizeComplete
End Sub


Private Sub cmdFind_Click()
    Dim i As Integer, j As Integer
        
    If txtFind.Text = "" Then Exit Sub
    With flxREC
        For i = fRow To .Rows - 1
            For j = fCol To 7
                If InStr(.TextMatrix(i, j), txtFind.Text) > 0 Then
                    .Col = 1: .Row = i
                    .ColSel = 8:  .RowSel = i
                    .TopRow = i
                    fRow = IIf(j = 11, i + 1, i)
                    fCol = IIf(j = 11, 1, j + 1)
                    Exit Sub
                End If
            Next j
            fCol = 1
        Next i
    End With
    MsgBox "No match found!", vbInformation + vbOKOnly, "Find"
    fRow = 1
    fCol = 1
End Sub

Private Sub cmdNext_Click()
    cmdNext.Tag = "2"
    cmdNext.Visible = False
    cmdPrev.Left = cmdNext.Left
    cmdPrev.Top = cmdNext.Top
    cmdPrev.Visible = True
    dTo.Visible = True
    cmdRefresh.Visible = True
    dTo.Value = dFrom.Value
End Sub

Private Sub cmdPrev_Click()
    cmdNext.Tag = "1"
    cmdNext.Visible = True
    cmdPrev.Visible = False
    dTo.Visible = False
    cmdRefresh.Visible = False
    Call RefreshRecord(1)
End Sub

Private Sub cmdRefresh_Click()
    Call RefreshRecord(2)
End Sub

Private Sub cmdTrans_Click(Index As Integer)
    Dim dCov As String
    
    Select Case Index
        Case 0  'Show Packing No
            If flxREC.Rows = 2 Then ActRow = 1
            If rsSPrec.RecordCount > 0 And ActRow <> 0 Then
                SPincoming.Show
                Call SPincoming.ShowRecord(flxREC.TextMatrix(ActRow, 2))
                Unload Me
            Else
                SPincoming.Show
                Unload Me
            End If
        Case 1  'Refresh
            Call RefreshRecord(Val(cmdNext.Tag))
        Case 2  'Print
            If Val(cmdNext.Tag) = 2 Then
                dCov = "Date Covered: " & dFrom.Value & " to " & dTo.Value
            Else
                dCov = "Date Covered: " & dFrom.Value
            End If
            PRNmc.MySource = "SPRECEIVESUMM"
            PRNmc.QryPrn = Qry
            PRNmc.ParamValue1 = dCov
            PRNmc.Show
        Case 3  'Exit
            Unload Me
            MotoRep.tbMain.Visible = True
    End Select
        
End Sub

Private Sub dFrom_CloseUp()
    If cmdNext.Tag = "1" Then
        Call RefreshRecord(1)
    End If
End Sub

Private Sub dTo_CloseUp()
    If dTo.Value < dFrom.Value Then
        MsgBox "Please choose a date not less than " & dFrom.Value, vbExclamation + vbOKOnly, "Validation Error"
        dTo.Value = dFrom.Value
    End If
End Sub


Private Sub flxrec_EnterCell()
    ActRow = flxREC.Row
End Sub


Private Sub Form_Load()
    
    If Day(Date) <= 7 Then
        dFrom.Value = DateAdd("m", -1, DateSerial(Year(Date), Month(Date), 1))
    Else
        dFrom.Value = DateSerial(Year(Date), Month(Date), 1)
    End If
    dTo.Value = Date
    cmdNext.Tag = "2"
    'Call cmdNext_Click
    Call RefreshRecord(2)
    ActRow = 0
    MotoRep.sBar.Panels(6).Text = "Spare Parts Receiving Summary"
    MotoRep.tbMain.Visible = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    MotoRep.sBar.Panels(6).Text = ""
    If rsSPrec.State = adStateOpen Then rsSPrec.Close
End Sub



Private Sub RXResizer1_ResizeComplete()
    lblHeading(2).Left = Frame1.Left + (Frame1.Width - lblHeading(2).Width) - 30
    lblHeading(3).Left = Frame1.Left + (Frame1.Width - lblHeading(3).Width)
    
    With flxREC               '1          2            3         4        5             6           7       8        9
        .FormatString = "   |^DATE REC.|<PACKING NO.|<SUPPLIER|<SHIPPER|<WAY BILL NO.|^PACK DATE|<REMARKS|<CHECKER|<PRINT COPY"
        .ColWidth(0) = 300
        .ColWidth(1) = (.Width - 300) * 0.12
        .ColWidth(2) = (.Width - 300) * 0.12
        .ColWidth(3) = (.Width - 300) * 0.17
        .ColWidth(4) = (.Width - 300) * 0.17
        .ColWidth(5) = (.Width - 300) * 0.12
        .ColWidth(6) = (.Width - 300) * 0.12
        .ColWidth(7) = (.Width - 300) * 0.15
        .ColWidth(8) = (.Width - 300) * 0.15
        .ColWidth(9) = (.Width - 300) * 0.15
        .ColWidth(10) = 0
        If rsSPrec.RecordCount > 0 Then
            For i = 1 To .Rows - 1
                .TextMatrix(i, 1) = Format$(.TextMatrix(i, 1), "mm/dd/yyyy")
                .TextMatrix(i, 6) = Format$(.TextMatrix(i, 6), "mm/dd/yyyy")
                If Val(.TextMatrix(i, 9)) = 0 Then 'Not yet printed
                    .Col = 1: .Row = i: .ColSel = 10: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 250, 200)
                    .Col = 2
                End If
                If InStr(.TextMatrix(i, 7), "CANCELLED") > 0 Then 'Cancelled
                    .Col = 1: .Row = i: .ColSel = 10: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 100, 100)
                    .Col = 2
                End If
            Next i
        End If
    End With
    
    cmdFind.Height = cmdRefresh.Height
    cmdFind.Top = cmdRefresh.Top
    txtFind.Top = cmdFind.Top
    
End Sub

Private Sub txtFind_Change()
    fRow = 1
    fCol = 1
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then Call cmdFind_Click
End Sub

