VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form BranchPOSumOld 
   Caption         =   "Branch Purchase Order"
   ClientHeight    =   8190
   ClientLeft      =   3930
   ClientTop       =   2610
   ClientWidth     =   11475
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   8190
   ScaleWidth      =   11475
   Begin VB.TextBox txtFind 
      Height          =   285
      Left            =   8595
      TabIndex        =   7
      Tag             =   "XYF"
      Top             =   3735
      Width           =   2610
   End
   Begin VB.Frame Frame1 
      Height          =   1065
      Left            =   90
      TabIndex        =   0
      Tag             =   "XY"
      Top             =   495
      Width           =   11280
      Begin VB.TextBox txtRefNo 
         Height          =   330
         Left            =   1350
         TabIndex        =   15
         Top             =   585
         Width           =   3030
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Show All"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4500
         TabIndex        =   14
         Top             =   585
         Width           =   1095
      End
      Begin VB.ComboBox cboBranch 
         Height          =   315
         Left            =   1350
         TabIndex        =   4
         Top             =   225
         Width           =   3030
      End
      Begin VB.Label Label3 
         Caption         =   "Reference # :"
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
         Left            =   90
         TabIndex        =   10
         Top             =   585
         Width           =   1605
      End
      Begin VB.Label Label2 
         Caption         =   "Branches     :"
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
         Left            =   90
         TabIndex        =   3
         Top             =   225
         Width           =   1245
      End
   End
   Begin lvButton.lvButtons_H cmdOk 
      Height          =   465
      Left            =   7695
      TabIndex        =   5
      Top             =   7650
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   820
      Caption         =   "&Load"
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
      Image           =   "BranchPOSumOld.frx":0000
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H lvButtons_H1 
      Height          =   465
      Left            =   9540
      TabIndex        =   6
      Top             =   7650
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   820
      Caption         =   "&Close"
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
      Image           =   "BranchPOSumOld.frx":059A
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdFind 
      Height          =   375
      Left            =   6930
      TabIndex        =   8
      Tag             =   "XY"
      Top             =   3690
      Width           =   1665
      _ExtentX        =   2937
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
      Image           =   "BranchPOSumOld.frx":0B34
      cBack           =   -2147483633
   End
   Begin VB.Frame Frame2 
      Height          =   6090
      Left            =   90
      TabIndex        =   9
      Top             =   1485
      Width           =   11265
      Begin VB.ComboBox cboAllocated 
         Height          =   315
         Left            =   90
         TabIndex        =   13
         Top             =   2250
         Width           =   3030
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid FlxServeOrder 
         Height          =   3225
         Left            =   90
         TabIndex        =   11
         Top             =   2700
         Width           =   11010
         _ExtentX        =   19420
         _ExtentY        =   5689
         _Version        =   393216
         RowHeightMin    =   285
         SelectionMode   =   1
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid FlxSum 
         Height          =   1665
         Left            =   90
         TabIndex        =   16
         Top             =   495
         Width           =   11025
         _ExtentX        =   19447
         _ExtentY        =   2937
         _Version        =   393216
         BackColor       =   12648447
         RowHeightMin    =   285
         BackColorBkg    =   -2147483638
         FocusRect       =   0
         GridLinesUnpopulated=   1
         SelectionMode   =   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label Label4 
         Caption         =   "Purchase Order"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   135
         TabIndex        =   12
         Top             =   225
         Width           =   1605
      End
   End
   Begin VB.Label lblHeading 
      BackStyle       =   0  'Transparent
      Caption         =   "BRANCH PURCHASE ORDER"
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
      TabIndex        =   1
      Top             =   225
      Width           =   2970
   End
   Begin VB.Label lblHeading 
      BackStyle       =   0  'Transparent
      Caption         =   "BRANCH PURCHASE ORDER"
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
      TabIndex        =   2
      Top             =   210
      Width           =   3060
   End
   Begin VB.Menu mnudummy 
      Caption         =   ""
      Enabled         =   0   'False
      Visible         =   0   'False
   End
End
Attribute VB_Name = "BranchPOSumOld"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rssummary As New ADODB.Recordset
Dim rsServe As New ADODB.Recordset
Dim ActRow As Integer, fRow As Integer, fCol As Integer, TCol As Integer, AccID As Integer, ColPrn As Integer
Dim Qry As String
Public mySrc As String
Dim vbBack As Boolean
Private m_bEditFromCode As Boolean
Dim bNone As Boolean
Private Sub cboBranch_Click()
    Call RefreshRecord
End Sub
Private Sub cmdOK_Click()
Dim dCov As String, pQRY As String, JOid As String
    With FlxSum
        If .Rows = 2 And FlxSum.TextMatrix(1, 1) = "" Then
            MsgBox "There is no record to view!", vbInformation + vbOKOnly, "Confirm"
            Exit Sub
        End If
        
        Call BranchPO.ShowRecord(.TextMatrix(ActRow, 8))
        BranchPO.cmdLoadPO.Enabled = False
        Unload Me
    End With
End Sub
Private Sub cmdTrans_Click(Index As Integer)
    Select Case Index
        Case 0  'Show Form
            
    End Select
End Sub
Public Sub RefreshRecord()
    Dim i As Integer
    
    Qry = "SELECT sporderbranch.REFERENCE, sporderbranch.ORDERTYPE, sporderbranch.ORDERDATE, sporderbranch.QTYORDER, sporderbranch.DATEDOWNLOAD, sporderbranch.BRANCHNAME, sporderbranch.BRANCHID, sporderbranch.SPORDERID From sporderbranch Where sporderbranch.BRANCHID= " & cboBranch.ItemData(cboBranch.ListIndex) & " ORDER BY sporderbranch.ORDERDATE"
    
    If rssummary.State = adStateOpen Then rssummary.Close
    rssummary.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    If rssummary.RecordCount = 0 Then
        FlxSum.Clear
        FlxSum.Rows = 2
        'Getting the Total Column
        FlxSum.Cols = TCol
        Call RXResizer1_ResizeComplete
        Exit Sub
    End If
    Set FlxSum.DataSource = rssummary
    ActRow = 1
    FlxSum.Visible = True
    Call RXResizer1_ResizeComplete
End Sub
Sub SPOrderTaking()
    Dim i As Integer
    
    With FlxServeOrder
        Qry = "SELECT spordrconfirm.DATECONFIRM, spordrconfirm.CONFIRMNO, spordrconfirm.ORDERTYPE, spordrconfirm.BACKORDER, spordrconfirm.QTYTAKING, spordrconfirm.DATEUPLOAD, spordrconfirm.DATEDOWNLOAD From spordrconfirm Where spordrconfirm.BRANCHID = " & cboBranch.ItemData(cboBranch.ListIndex) & " AND spordrconfirm.REFERENCE = '" & FlxSum.TextMatrix(FlxSum.Row, FlxSum.Col) & "' Order By spordrconfirm.SQSERVE"
        If rsServe.State = adStateOpen Then rsServe.Close
        rsServe.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If rsServe.RecordCount = 0 Then
            .Clear
            .Rows = 2
            Exit Sub
        End If
        Set .DataSource = rsServe
        .FormatString = "   |^DATE CONFIRM|^CONF. #|<ORDER TYPE|^BACKORDER|>QTY TAKING|^DATEUPLOAD|^DOWNLOAD"
           
        .ColWidth(0) = 400
        .ColWidth(1) = (.Width - 300) * 0.12
        .ColWidth(2) = (.Width - 300) * 0.12
        .ColWidth(3) = (.Width - 300) * 0.12
        .ColWidth(4) = (.Width - 300) * 0.13
        .ColWidth(5) = (.Width - 300) * 0.13
        .ColWidth(6) = (.Width - 300) * 0.13
        .ColWidth(7) = (.Width - 300) * 0.13
       
    End With
End Sub
Private Sub cmdFind_Click()
    Dim i As Integer, j As Integer
        
    If txtFind.Text = "" Then Exit Sub
    With FlxSum
        For i = fRow To .Rows - 1
            For j = fCol To TCol - 1
                If InStr(.TextMatrix(i, j), txtFind.Text) > 0 Then
                    .Col = 1: .Row = i
                    .ColSel = TCol - 1: .RowSel = i
                    .TopRow = i
                    fRow = IIf(j = TCol - 1, i + 1, i)
                    fCol = IIf(j = TCol - 1, 1, j + 1)
                    ActRow = i
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

Private Sub flxSum_EnterCell()
    ActRow = FlxSum.Row
    
End Sub

Private Sub flxSum_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    If FlxSum.TextMatrix(ActRow, 1) = "" Or mySrc = " " Then Exit Sub
    If Not IsDate(FlxSum.TextMatrix(ActRow, 1)) Then
        Exit Sub
    Else
        If DateValue(FlxSum.TextMatrix(ActRow, 1)) <> Date Then Exit Sub
    End If
    
End Sub
Private Sub Form_Load()
    
    Dim Qry As String
    
    If mySrc = "" Then
        'TCol = 14: ColPrn = 12: AccID = 43
        TCol = 14: ColPrn = 12: AccID = 43
        ActRow = 0
        MotoRep.sBar.Panels(6).Text = mySrc & " Summary"
        Me.Caption = mySrc & " Summary"
        MotoRep.tbMain.Visible = False
        ActRow = 1
    End If
    
    cboBranch.AddItem "MOTO-R ULAS"
    cboBranch.ItemData(cboBranch.ListCount - 1) = 81
    cboBranch.AddItem "YAMAHA MOTOR-HUB"
    cboBranch.ItemData(cboBranch.ListCount - 1) = 99
    cboBranch.AddItem "MC PARTS DEPOT CORP."
    cboBranch.ItemData(cboBranch.ListCount - 1) = 99
    cboBranch.Text = "MOTO-R ULAS"
    cboBranch.ListIndex = cboBranch.ListCount - 3
    
    cboAllocated.AddItem "ALL"
    cboAllocated.AddItem "ALLOCATED"
    cboAllocated.AddItem "BACKORDER"
    cboAllocated.ListIndex = 2
    Call RefreshRecord
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    MotoRep.sBar.Panels(6).Text = ""
    If rssummary.State = adStateOpen Then rssummary.Close
End Sub

Private Sub lvButtons_H1_Click()
    Unload Me
End Sub
Private Sub RXResizer1_ResizeComplete()
    Dim i As Integer, j As Integer
    With FlxSum
                               '1=Reference Number   2=Order Type 3=Order Date 4=Qty Order 5=Date Download 7=Branch Name 8=Branch ID
            .FormatString = "   |^REFERENCE NUMBER|^ORDER TYPE|<ORDER DATE|< QTY |^DATE DOWNLOAD|>BRANCH NAME|^BRANCH ID|^ORDER ID"
            .ColWidth(0) = 400
            .ColWidth(1) = (.Width - 300) * 0.2
            .ColWidth(2) = (.Width - 300) * 0.15
            .ColWidth(3) = (.Width - 300) * 0.25
            .ColWidth(4) = (.Width - 300) * 0.2
            .ColWidth(5) = (.Width - 300) * 0.19
            .ColWidth(6) = 0
            .ColWidth(7) = 0
            .ColWidth(8) = 0
            
    End With
End Sub
Private Sub txtFind_Change()
    fRow = 1
    fCol = 1
End Sub
Private Sub txtFind_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then Call cmdFind_Click
End Sub
Private Sub txtRefNo_Change()
With FlxSum
   'Before.........
   'If m_bEditFromCode Then
   '     m_bEditFromCode = False
   '    Exit Sub
   'End If
   '.Visible = True
   'New.......
   'If txtRefNo.Text = "" Or vbBack = True Then .Visible = False: vbBack = False: Exit Sub
   
   If txtRefNo.Text = "" Or vbBack = True Then: vbBack = False: Exit Sub
   
   'Before
   'If Text1 = "" Then .Visible = False: Exit Sub
   '----------------
    'rs.Find "[f_name] LIKE '" + Text1 + "%'", , adSearchForward, 1
    rssummary.Find "REFERENCE LIKE '" + txtRefNo + "%'", , adSearchForward, 1
    
     If Not rssummary.EOF Then
             If Not bNone Then
              txtRefNo.Tag = Len(txtRefNo.Text)
              bNone = True
             End If
             '.TopRow = rssummary.AbsolutePosition - 1
             .TopRow = rssummary.AbsolutePosition
             '.Row = rssummary.AbsolutePosition - 1
             'Select Column from col 1 to col 8
             .Col = 1
             .ColSel = 8
               
              txtRefNo.Text = rssummary!REFERENCE
              txtRefNo.SelStart = Val(txtRefNo.Tag)
              txtRefNo.SelLength = Len(txtRefNo.Text)
              FlxSum.Tag = txtRefNo.Text
              
              bNone = False
             
     Else
         'rs.Requery
         .Refresh
         '.Visible = False
        
     End If
End With
End Sub

Private Sub txtRefNo_Click()
    SendKeys "{home}+{end}"
    txtRefNo.SetFocus
    'FlxSum.Visible = False
End Sub

Private Sub txtRefNo_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyDown
        FlxSum.SetFocus
        Call SPOrderTaking
    Case vbKeyEscape
        txtRefNo = Empty
        txtRefNo.SetFocus
    Case vbKeyDelete
        m_bEditFromCode = True
    Case vbKeyBack
        '***New update search by: rvp
        'm_bEditFromCode = True
        If Val(txtRefNo.Tag) > 0 Then txtRefNo.SelStart = Val(txtRefNo.Tag) - 1
        If Len(txtRefNo.Text) > 0 Then txtRefNo.SelLength = Len(txtRefNo.Text) - 1
        If txtRefNo.SelStart = 0 Then
            txtRefNo = ""
            vbBack = True
        End If
    Case vbKeyRight
       
        
End Select
End Sub
Private Sub txtrefno_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        bNone = False
        'FlxSum.Visible = False
    End If
    If KeyAscii > 32 Then KeyAscii = UpperCaseText(KeyAscii)
End Sub
Function UpperCaseText(iKeyAscii As Integer) As Integer
Dim char As String
    char = Chr(iKeyAscii)
    UpperCaseText = Asc(UCase(char))
End Function

