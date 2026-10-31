VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form SPorderSum 
   Caption         =   "Spare Parts P.O. Summary"
   ClientHeight    =   8145
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11460
   LinkTopic       =   "Form1"
   ScaleHeight     =   8145
   ScaleWidth      =   11460
   WindowState     =   2  'Maximized
   Begin MSFlexGridLib.MSFlexGrid flxOrder 
      Height          =   7695
      Left            =   0
      TabIndex        =   8
      Tag             =   "XY"
      Top             =   480
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   13573
      _Version        =   393216
      Cols            =   9
      RowHeightMin    =   275
      BackColorBkg    =   -2147483633
      SelectionMode   =   1
      AllowUserResizing=   1
      FormatString    =   $"SPOrderSum.frx":0000
   End
   Begin CRVIEWERLibCtl.CRViewer vwrDmCm 
      Height          =   7695
      Left            =   0
      TabIndex        =   3
      Tag             =   "XY"
      Top             =   480
      Width           =   11415
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
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
   Begin VB.Timer tmrDropRev 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   4560
      Top             =   8280
   End
   Begin VB.Timer tmrFloatRev 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   4560
      Top             =   7800
   End
   Begin VB.Timer tmrDrop 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3960
      Top             =   8280
   End
   Begin VB.Timer tmrFloat 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3960
      Top             =   7800
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   2
      Left            =   3360
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
      Caption         =   "&View Order"
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
      Image           =   "SPOrderSum.frx":00B7
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
      Caption         =   "&Grid View"
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
      Image           =   "SPOrderSum.frx":0651
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   1
      Left            =   1680
      TabIndex        =   2
      Top             =   0
      Width           =   1680
      _ExtentX        =   2963
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
      Image           =   "SPOrderSum.frx":0BEB
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   3
      Left            =   5040
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
      Caption         =   "&Print Preview"
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
      Image           =   "SPOrderSum.frx":1185
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   4
      Left            =   6720
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   1680
      _ExtentX        =   2963
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
      Image           =   "SPOrderSum.frx":171F
      cBack           =   -2147483633
   End
   Begin VB.Label lblRow 
      Height          =   255
      Left            =   10320
      TabIndex        =   7
      Top             =   120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblP 
      Caption         =   "è"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   8280
      Width           =   255
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   480
      Top             =   7680
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   11580
      m_MDI_Height    =   8550
   End
   Begin VB.Menu nmuSales 
      Caption         =   "Sales Report"
      Visible         =   0   'False
   End
End
Attribute VB_Name = "SPorderSum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsRPT As New ADODB.Recordset

Dim BtnIndx As Integer, Qry As String
Dim ActRow As Integer, dFrm As Date, dTo As Date
Public ViewList As Boolean



Private Sub cmdTrans_Click(Index As Integer)
    Dim MyQry As String
    
    Select Case Index
        Case 0  'Grid View
            BtnIndx = 1
            tmrFloat.Enabled = True
            Call PrintPreView(False)
           ' Call cboMonth_Click
        Case 1  'Exit
            MotoRep.tbMain.Visible = True
            Unload Me
        Case 2  'SP P.O.
            With flxORDER
                If .TextMatrix(ActRow, 3) <> "" Then
                    If ActRow = 0 And .Rows > 2 Then ActRow = 1
                    SPordering.Show
                    Call SPordering.DisplaySPOrder(.TextMatrix(ActRow, 3))
                End If
            End With
            Unload Me
        Case 3  'Print Preview
'            BtnIndx = 4
'            tmrFloatRev.Enabled = True
'            Call PrintPreView(True)
'            Call DisplayReport
        Case 4  'Exit
            MotoRep.tbMain.Visible = True
            Unload Me
    End Select
End Sub

Private Sub flxOrder_Click()
    ActRow = flxORDER.Row
End Sub

Private Sub flxOrder_EnterCell()
    ActRow = flxORDER.Row
End Sub

Private Sub Form_Load()
    Dim rsOrder As New ADODB.Recordset
    Dim Qry As String
    Dim i As Integer

    Qry = "SELECT SPORDER.ORDERDATE, SPORDER.REFERENCE, SUPPLIER.SUPPLIERNAME, SPORDER.ORDERTYPE, SPORDER.CONFIRMNO, SPORDER.SUBMITTEDBY, SPORDER.APPROVEDBY " & _
          "FROM SPORDER INNER JOIN SUPPLIER ON SPORDER.SUPPLIERID = SUPPLIER.SUPPLIERID ORDER BY SPORDER.ORDERDATE DESC;"

    With rsOrder
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            flxORDER.Rows = 1
            Do While Not .EOF
                flxORDER.AddItem vbTab & vbTab & Format$(!ORDERDATE, "MM/DD/YYYY") & vbTab & !REFERENCE & vbTab & !SUPPLIERNAME & vbTab & !OrderType & vbTab & !ConfirmNo & vbTab & !SubmittedBy & vbTab & !ApprovedBy
                .MoveNext
            Loop
        End If
        .Close
    End With
    
    BtnIndx = 1
    tmrFloat.Enabled = True
    MotoRep.sBar.Panels(6).Text = "Spare Parts Order Summary"
End Sub

Private Sub DisplayReport()
End Sub

Private Sub PrintPreView(Ppvw As Boolean)
    If Ppvw Then
'        Frame1.Visible = False
        flxORDER.Visible = False
        vwrDmCm.Visible = True
    Else
        vwrDmCm.Visible = False
 '       Frame1.Visible = True
        flxORDER.Visible = True
    End If
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If rsRPT.State = adStateOpen Then rsRPT.Close
End Sub

Private Sub RXResizer1_ResizeComplete()
    Dim i As Integer
    With flxORDER
        .ColWidth(0) = 300
        .ColWidth(1) = 0
        .ColWidth(2) = (.Width - 600) * 0.12
        .ColWidth(3) = (.Width - 600) * 0.12
        .ColWidth(4) = (.Width - 600) * 0.2
        .ColWidth(5) = (.Width - 600) * 0.12
        .ColWidth(6) = (.Width - 600) * 0.12
        .ColWidth(7) = (.Width - 600) * 0.16
        .ColWidth(8) = (.Width - 600) * 0.16
    End With
End Sub

Private Sub tmrDrop_Timer()
    cmdTrans(BtnIndx).Top = cmdTrans(BtnIndx).Top + 50
    If cmdTrans(BtnIndx).Top >= 0 Then
        cmdTrans(BtnIndx).Top = 0
        BtnIndx = BtnIndx + 1
        If BtnIndx = 5 Then     'last index + 1
            tmrDrop.Enabled = False
        End If
    End If
End Sub

Private Sub tmrDropRev_Timer()
    cmdTrans(BtnIndx).Top = cmdTrans(BtnIndx).Top + 50
    If cmdTrans(BtnIndx).Top >= 0 Then
        cmdTrans(BtnIndx).Top = 0
        BtnIndx = BtnIndx + 1
        If BtnIndx = 2 Then
            tmrDropRev.Enabled = False
        End If
    End If
End Sub

Private Sub tmrFloat_Timer()
    Dim i As Integer
    
    cmdTrans(BtnIndx).Top = cmdTrans(BtnIndx).Top - 50
    If cmdTrans(BtnIndx).Top <= -375 Then
        cmdTrans(BtnIndx).Visible = False
        BtnIndx = BtnIndx - 1
        If BtnIndx = -1 Then
            tmrFloat.Enabled = False
            BtnIndx = 2 'Index of Button view grid
            For i = 2 To 4
                cmdTrans(i).Top = -375
                cmdTrans(i).Left = IIf(i = 2, 0, cmdTrans(2).Width * (i - 2))
                cmdTrans(i).Visible = True
            Next i
            tmrDrop.Enabled = True
        End If
    End If
End Sub

Private Sub tmrFloatRev_Timer()
    
    cmdTrans(BtnIndx).Top = cmdTrans(BtnIndx).Top - 50
    If cmdTrans(BtnIndx).Top <= -375 Then
        cmdTrans(BtnIndx).Visible = False
        BtnIndx = BtnIndx - 1
        If BtnIndx = 1 Then 'Index of Button Exit
            tmrFloatRev.Enabled = False
            BtnIndx = 0
            cmdTrans(0).Visible = True
            cmdTrans(1).Visible = True
            tmrDropRev.Enabled = True
        End If
    End If
End Sub

Private Sub updYear_Change()
    txtYear.Text = updYear.Value
    'Call cboMonth_Click
End Sub
