VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form BranchPOSum 
   Caption         =   "Branch Purchase Order"
   ClientHeight    =   8130
   ClientLeft      =   2325
   ClientTop       =   2475
   ClientWidth     =   11865
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   8130
   ScaleWidth      =   11865
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxAllPO 
      Height          =   1695
      Left            =   2130
      TabIndex        =   35
      Top             =   8280
      Visible         =   0   'False
      Width           =   2355
      _ExtentX        =   4154
      _ExtentY        =   2990
      _Version        =   393216
      Cols            =   1
      FixedCols       =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   1
   End
   Begin VB.ComboBox cboAllocated 
      Height          =   315
      Left            =   9045
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1950
      Width           =   2730
   End
   Begin VB.CheckBox chkMark 
      Height          =   225
      Left            =   180
      TabIndex        =   4
      Tag             =   "XYF"
      Top             =   2310
      Width           =   225
   End
   Begin MSComctlLib.ProgressBar PgBar 
      Height          =   210
      Left            =   1620
      TabIndex        =   29
      Top             =   5790
      Width           =   4770
      _ExtentX        =   8414
      _ExtentY        =   370
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxsp 
      Height          =   3405
      Left            =   105
      TabIndex        =   5
      Top             =   2280
      Width           =   11685
      _ExtentX        =   20611
      _ExtentY        =   6006
      _Version        =   393216
      Cols            =   10
      FixedCols       =   0
      SelectionMode   =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   10
   End
   Begin VB.TextBox txtGTotBackOdr 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   9150
      TabIndex        =   24
      Top             =   5730
      Width           =   855
   End
   Begin VB.TextBox txtGTotalAllQ 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   8160
      TabIndex        =   23
      Top             =   5730
      Width           =   960
   End
   Begin VB.TextBox txtGTotalAmount 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   10680
      TabIndex        =   22
      Top             =   5730
      Width           =   1110
   End
   Begin VB.Frame Frame1 
      Height          =   1185
      Left            =   90
      TabIndex        =   0
      Tag             =   "XY"
      Top             =   750
      Width           =   11700
      Begin VB.CheckBox chkPO 
         BackColor       =   &H00FFC0C0&
         Caption         =   "All P.O"
         Height          =   345
         Left            =   3210
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   240
         Width           =   975
      End
      Begin VB.ComboBox cboRefNo 
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
         Left            =   1350
         TabIndex        =   2
         Tag             =   "XYF"
         Top             =   240
         Width           =   1830
      End
      Begin VB.ComboBox cboBranch 
         Height          =   315
         Left            =   5805
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   1980
      End
      Begin VB.TextBox txtPODate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
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
         Left            =   9210
         Locked          =   -1  'True
         TabIndex        =   7
         Tag             =   "XYF"
         Top             =   285
         Width           =   1965
      End
      Begin VB.TextBox txtOdrtype 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
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
         Left            =   9210
         Locked          =   -1  'True
         TabIndex        =   8
         Tag             =   "XYF"
         Top             =   645
         Width           =   1965
      End
      Begin VB.TextBox txtdateUpload 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
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
         Left            =   5820
         Locked          =   -1  'True
         TabIndex        =   6
         Tag             =   "XYF"
         Top             =   645
         Width           =   1965
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "REFERENCE#:"
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
         Left            =   285
         TabIndex        =   28
         Top             =   285
         Width           =   1065
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "BRANCHES:"
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
         Left            =   4800
         TabIndex        =   27
         Top             =   285
         Width           =   1065
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "ORDER TYPE"
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
         Left            =   7770
         TabIndex        =   11
         Tag             =   "XYF"
         Top             =   690
         Width           =   1185
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "P.O. DATE"
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
         Left            =   8130
         TabIndex        =   10
         Tag             =   "XYF"
         Top             =   330
         Width           =   825
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "DATE UPLOAD"
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
         Left            =   4560
         TabIndex        =   9
         Tag             =   "XYF"
         Top             =   690
         Width           =   1095
      End
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   661
      Caption         =   "Load Order"
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
      Image           =   "BranchPOSum.frx":0000
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   1
      Left            =   1800
      TabIndex        =   13
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
      Image           =   "BranchPOSum.frx":059A
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   2
      Left            =   3600
      TabIndex        =   14
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
      Image           =   "BranchPOSum.frx":0B34
      cBack           =   -2147483633
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxTrans 
      Height          =   1905
      Left            =   90
      TabIndex        =   32
      Top             =   6120
      Width           =   11685
      _ExtentX        =   20611
      _ExtentY        =   3360
      _Version        =   393216
      Cols            =   7
      FixedCols       =   0
      SelectionMode   =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   7
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxPartsno 
      Height          =   1665
      Left            =   5340
      TabIndex        =   34
      Top             =   8340
      Visible         =   0   'False
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   2937
      _Version        =   393216
      Rows            =   1
      Cols            =   1
      FixedRows       =   0
      FixedCols       =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   1
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
      Left            =   -30
      TabIndex        =   31
      Top             =   8220
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "ORDER STATUS:"
      Height          =   240
      Left            =   7380
      TabIndex        =   30
      Top             =   1995
      Width           =   1455
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL :"
      Height          =   195
      Left            =   7320
      TabIndex        =   21
      Top             =   5790
      Width           =   615
   End
   Begin VB.Label lblHeading 
      BackStyle       =   0  'Transparent
      Caption         =   "ORDERED SPARE PARTS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   4
      Left            =   105
      TabIndex        =   19
      Tag             =   "XYF"
      Top             =   2025
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
      Height          =   255
      Index           =   0
      Left            =   90
      TabIndex        =   16
      Tag             =   "TLF"
      Top             =   450
      Width           =   2895
   End
   Begin VB.Label lblHeading 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PURCHASE ORDER"
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
      Height          =   240
      Index           =   2
      Left            =   9975
      TabIndex        =   15
      Tag             =   "TRF"
      Top             =   480
      Width           =   1800
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
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   17
      Tag             =   "TLF"
      Top             =   480
      Width           =   2895
   End
   Begin VB.Label lblHeading 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PURCHARE ORDER"
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
      Height          =   240
      Index           =   3
      Left            =   10020
      TabIndex        =   18
      Tag             =   "TRF"
      Top             =   510
      Width           =   1800
   End
   Begin VB.Label lblHeading 
      BackStyle       =   0  'Transparent
      Caption         =   "ORDERED SPARE PARTS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   105
      TabIndex        =   20
      Tag             =   "XYF"
      Top             =   2055
      Width           =   2415
   End
   Begin VB.Label lblHeading 
      BackStyle       =   0  'Transparent
      Caption         =   "STOCK TRANSFER"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   6
      Left            =   105
      TabIndex        =   25
      Tag             =   "XYF"
      Top             =   5820
      Width           =   1605
   End
   Begin VB.Label lblHeading 
      BackStyle       =   0  'Transparent
      Caption         =   "STOCK TRANSFER"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   7
      Left            =   60
      TabIndex        =   26
      Tag             =   "XYF"
      Top             =   5805
      Width           =   1605
   End
   Begin VB.Menu mnudummy 
      Caption         =   ""
      Enabled         =   0   'False
      Visible         =   0   'False
   End
   Begin VB.Menu mnumarkOrder 
      Caption         =   "SPORDER"
      Visible         =   0   'False
      Begin VB.Menu mnuMark 
         Caption         =   "Mark Row"
      End
      Begin VB.Menu mnuUnmark 
         Caption         =   "Unmark"
      End
      Begin VB.Menu mnusep1 
         Caption         =   ""
      End
      Begin VB.Menu mnuMarkall 
         Caption         =   "Mark All"
      End
      Begin VB.Menu mnuUnmakrAll 
         Caption         =   "Unmark All"
      End
   End
End
Attribute VB_Name = "BranchPOSum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rssummary As New ADODB.Recordset
Dim rsServe As New ADODB.Recordset
Dim ActRow As Integer, fRow As Integer, fCol As Integer, TCol As Integer, AccID As Integer, ColPrn As Integer
Dim Qry As String
Public mySrc As String
Dim SuppSrch As Boolean
Private Sub cboAllocated_Click()
    Call FormatFlxSPEmpty
    Call chkPO_Click
    DoEvents
    Call OrderStatus
    flxSP.FixedRows = 1
End Sub
Private Sub cboAllocated_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call cboAllocated_Click
    End If
End Sub
Private Sub cmdOK_Click()
Dim dCov As String, pQRY As String, JOid As String
    With flxSum
        If .Rows = 2 And flxSum.TextMatrix(1, 1) = "" Then
            MsgBox "There is no record to view!", vbInformation + vbOKOnly, "Confirm"
            Exit Sub
        End If
        BranchPO.mtrBName = cboBranch.Text
        Call BranchPO.ShowRecord(cboBranch.ItemData(cboBranch.ListIndex))
        BranchPO.cmdLoadPO.Enabled = False
        Unload Me
    End With
End Sub

Private Sub cboBranch_Click()
    'Call ClearHPO
    Call cboBranch_KeyPress(13)
End Sub
Private Sub cboBranch_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call ClearHPO
        Call RefreshRecord
        Call cboRefNo_Click
        'Call OrderStatus
        flxSP.FixedRows = 1
        Call cboAllocated_Click
    End If
End Sub
Private Sub cboRefNo_Change()
 Dim i As Long, j As Long
    Dim strPartial As String, strTotal As String
    Dim X As Integer
    
    If SuppSrch Then
        SuppSrch = False
        Exit Sub
    End If
    With cboRefNo
        strPartial = .Text
        i = SendMessage(.hwnd, CB_FINDSTRING, -1, ByVal strPartial) 'If match found, append unmatched characters
        If i <> -1 Then
            strTotal = .List(i)
            rssummary.MoveFirst
            rssummary.Move i
            
             'txtPONumber.Text = rssummary!REFERENCE
             txtdateUpload.Text = IIf(IsNull(rssummary!DATEUPLOAD), "", rssummary!DATEUPLOAD)
             txtPODate.Text = rssummary!ORDERDATE
             txtOdrtype.Text = rssummary!ORDERTYPE
            
            j = Len(strTotal) - Len(strPartial)
            If j <> 0 Then
                SuppSrch = True
                .SelText = Right$(strTotal, j)
                .SelStart = Len(strPartial)
                .SelLength = j
            End If
        Else
             'txtPONumber.Text = ""
             txtdateUpload.Text = ""
             txtPODate.Text = ""
             txtOdrtype.Text = ""
        End If
    End With
End Sub
Private Sub cboRefNo_Click()
    Call cboRefNo_Change
    Call cboRefNo_KeyPress(13)
    Call OrderStatus
End Sub
Private Sub cboRefNo_GotFocus()
Dim Tmp
    
   Const CB_SHOWDROPDOWN = &H14F
   Tmp = SendMessage(cboRefNo.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&)
End Sub

Private Sub cboRefNo_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then SuppSrch = True
End Sub
Private Sub cboRefNo_KeyPress(KeyAscii As Integer)
Dim i As Integer
    If cboRefNo.Text = "" Then: Exit Sub
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        If cboRefNo.Text <> "" Then
            Call DisplayPO
            Call cboRefNo_Change
            chkMark.Value = 1
            Call StockTrans
            Call ReQueryPO
            'flxAllPO.FixedRows = 1
            For i = flxAllPO.Rows - 1 To 1 Step -1
                If flxAllPO.TextMatrix(i, 0) <> cboRefNo.Text Then
                    flxAllPO.RemoveItem i
                End If
                
            Next i
            
        Else
            flxSP.Clear
            Call FormatFlxSPEmpty
            flxSP.Rows = 2
            flxTrans.Clear
            Call FormatFlxTrans
            flxTrans.Row = 2
        End If
    End If
End Sub

Private Sub chkMark_Click()
    If chkMark.Value = 0 Then
        mnuUnmarkAll_Click
    Else
        Call mnuMarkAll_Click
    End If
End Sub

Private Sub chkPO_Click()
    If chkPO.Value = 1 Then
        cboRefNo.Enabled = False
        Call ReQueryPO
        Call DisplayPO
        Call StockTrans
        Call OrderStatus
        
    Else
        
        cboRefNo.Enabled = True
        Call cboRefNo_Change
        Call cboRefNo_KeyPress(13)
        Call OrderStatus
    
    End If
End Sub

Private Sub cmdTrans_Click(Index As Integer)
Dim Irs As New ADODB.Recordset
Dim Qry As String, ViewPOID As Integer, i As Long, rRROW As Long, myW1 As String, myW As String
    Select Case Index
        Case 0  'Show Form
          If BranchPO.isEdit Then
                Call ClearExistSP
                For i = 1 To flxSP.Rows - 1
                    If flxSP.TextMatrix(i, 0) <> "" Then
                        If BranchPO.flxSP.Rows = 2 And BranchPO.flxSP.TextMatrix(1, 1) = "" Then
                            BranchPO.flxSP.Rows = 1
                        End If
                        BranchPO.flxSP.Rows = (BranchPO.flxSP.Rows) + 1
                        rRROW = BranchPO.flxSP.Rows - 1
                        BranchPO.flxSP.TextMatrix(rRROW, 8) = flxSP.TextMatrix(i, 3)
                        BranchPO.flxSP.TextMatrix(rRROW, 9) = ""
                        BranchPO.flxSP.TextMatrix(rRROW, 13) = flxSP.TextMatrix(i, 8)
                        BranchPO.flxSP.TextMatrix(rRROW, 7) = flxSP.TextMatrix(i, 7)
                        BranchPO.flxSP.TextMatrix(rRROW, 12) = ""
    
                        If Irs.State = adStateOpen Then Irs.Close
                         Qry = "SELECT spbranchdtl.SPAREPARTID, spareparts.PARTNUMBER, spareparts.DESCRIPTION, spareparts.QTYONHAND, spareparts.BACKORDER, spbranchdtl.BQTYONHAND, " & _
                              "spareparts.cost From spbranchdtl INNER JOIN spareparts ON (spbranchdtl.SPAREPARTID = spareparts.SPAREPARTID) " & _
                              "Where spareparts.SPAREPARTID =  " & Val(flxSP.TextMatrix(i, 8)) & " AND spbranchdtl.REFERENCE ='" & cboRefNo.Text & "' AND spbranchdtl.BRANCHNAME = '" & cboBranch.Text & "';"
                        If Irs.State = adStateOpen Then Irs.Close
                        Irs.Open Qry, LYRcon, adOpenKeyset, adLockPessimistic
                        BranchPO.flxSP.TextMatrix(rRROW, 1) = Irs!PartNumber
                        BranchPO.flxSP.TextMatrix(rRROW, 2) = Irs!Description
                        BranchPO.flxSP.TextMatrix(rRROW, 3) = Irs!QtyOnHand
                        BranchPO.flxSP.TextMatrix(rRROW, 4) = Irs!BACKORDER
                        BranchPO.flxSP.TextMatrix(rRROW, 5) = Val(BranchPO.GetSalesQtyMain(flxSP.TextMatrix(i, 8)))
                        BranchPO.flxSP.TextMatrix(rRROW, 6) = Irs!BQTYONHAND
                        BranchPO.flxSP.TextMatrix(rRROW, 10) = Format(Irs!Cost, "#,###,###.#0")
                        BranchPO.flxSP.TextMatrix(rRROW, 11) = "0.00"
                    
                    End If
                Next i
            Else
                If flxSP.Rows = 2 And flxSP.TextMatrix(1, 1) = "" Then Exit Sub
                If txtGTotBackOdr = 0 Then
                    MsgBox ("All Spareparts of P.O " & cboRefNo & " Already Allocated"), vbExclamation, "Allocated"
                    Exit Sub
                End If
                'If CheckPendingPO = True Then
                '    MsgBox ("Previous allocation in this P.O not yet Final. Update Stock Tranfer"), vbExclamation, "Not Yet Final"
                '    Exit Sub
                'End If
                
                BranchPO.mtrBName = cboBranch.Text
                Call BranchPO.ShowRecord(cboBranch.ItemData(cboBranch.ListIndex))
                Unload Me
            End If
           
        Case 1
        
        Case 2
           
            Unload Me
    End Select
End Sub
Private Sub ClearExistSP()
Dim Added As Long, Add As Long
    For Added = 1 To BranchPO.flxSP.Rows - 1
        For Add = 1 To flxSP.Rows - 1
            If flxSP.TextMatrix(Add, 8) = BranchPO.flxSP.TextMatrix(Added, 13) And flxSP.TextMatrix(Add, 0) <> "" Then
                flxSP.TextMatrix(Add, 0) = ""
            End If
        Next Add
    Next Added
End Sub
Public Sub RefreshRecord()
    Dim i As Integer, Allcd As Integer
   
    If cboAllocated.Text = "BACKORDER" Then
        Allcd = 0
    End If
    If cboAllocated.Text = "ALLOCATED" Then
        Allcd = 1
    End If
    
    If cboAllocated.Text = "ALL" Then
        Qry = "SELECT sporderbranch.REFERENCE, sporderbranch.ORDERTYPE, sporderbranch.ORDERDATE, sporderbranch.QTYORDER, sporderbranch.DATEUPLOAD, sporderbranch.BRANCHNAME, sporderbranch.BRANCHID, sporderbranch.SPORDERID From sporderbranch Where sporderbranch.BRANCHID= " & cboBranch.ItemData(cboBranch.ListIndex) & " ORDER BY sporderbranch.ORDERDATE"
    Else
        Qry = "SELECT sporderbranch.REFERENCE, sporderbranch.ORDERTYPE, sporderbranch.ORDERDATE, sporderbranch.QTYORDER, sporderbranch.DATEUPLOAD, sporderbranch.BRANCHNAME, sporderbranch.BRANCHID, sporderbranch.SPORDERID From sporderbranch Where sporderbranch.BRANCHID= " & cboBranch.ItemData(cboBranch.ListIndex) & " AND sporderbranch.ALLOCATED = " & Allcd & " ORDER BY sporderbranch.ORDERDATE"
    End If
  
    With rssummary
        If .State = adStateOpen Then .Close
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount = 0 Then
            chkMark.Value = 0
            Call FormatFlxSPEmpty
            Exit Sub
        End If
        'Put Validation to all p.o base on ordertaking and order serve
        cboRefNo.Clear
        flxSP.Clear
        flxAllPO.Clear
        flxAllPO.Rows = 1
        Call FormatFlxSP
        .MoveFirst
        For i = 1 To .RecordCount
            cboRefNo.AddItem !REFERENCE
            flxAllPO.AddItem !REFERENCE
            .MoveNext
        Next i
        flxAllPO.FixedRows = 1
    End With
End Sub
Private Sub ReQueryPO()
    Dim rsRpo As New ADODB.Recordset
    'Dim rsCh As New ADODB.Recordset
    'Dim Qry As String, Qry1 As String
    
       
        
    With rsRpo
        If .State = adStateOpen Then .Close
        Qry = "SELECT sporderbranch.REFERENCE From sporderbranch Where sporderbranch.BRANCHNAME= '" & cboBranch.Text & "' ORDER BY sporderbranch.ORDERDATE"

        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount = 0 Then
            chkMark.Value = 0
            Call FormatFlxSPEmpty
            Exit Sub
        End If
        flxAllPO.Clear
        flxAllPO.Rows = 1
        Call FormatFlxSP
        .MoveFirst
        For i = 1 To .RecordCount
            'If rsCh.State = adStateOpen Then rsCh.Close
            'Qry1 = "SELECT spordrconfirm.REFERENCE From spordrconfirm Where spordrconfirm.REFERENCE = '" & !REFERENCE & "' AND spordrconfirm.BRANCHNAME= '" & cboBranch.Text & "' ORDER BY spordrconfirm.ORDERDATE"
            'rsCh.Open Qry1, LYRcon, adOpenKeyset, adLockOptimistic
            'If rsCh.RecordCount = 0 Then
                flxAllPO.AddItem !REFERENCE
            'End If
            .MoveNext
        Next i
        
        .Close
    End With
End Sub
Private Sub cmdFind_Click()
    Dim i As Integer, j As Integer
        
    If txtFind.Text = "" Then Exit Sub
    With flxSum
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


Private Sub flxSP_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
        If flxSP.TextMatrix(flxSP.Row, 0) = "" And Val(flxSP.TextMatrix(flxSP.Row, 7)) > 0 Then
            Call mnuMark_Click
        Else
            Call mnuUnmark_Click
        End If
    End If
End Sub


Private Sub Form_Load()
    
    With cboBranch
         .AddItem "MOTO-R ULAS"
         .ItemData(cboBranch.ListCount - 1) = 13
         .AddItem "YAMAHA MOTORHUB"
         .ItemData(cboBranch.ListCount - 1) = 14
         .AddItem "MC PARTS DEPOT CORP."
         .ItemData(cboBranch.ListCount - 1) = 15
         .ListIndex = 0
    End With
         flxAllPO.Clear
         PgBar.Visible = False
         cboAllocated.AddItem "ALL"
         cboAllocated.AddItem "BACKORDER"
         cboAllocated.AddItem "ALLOCATED"
         cboAllocated.ListIndex = 0
         
         Call RefreshRecord
         Call FormatFlxSPEmpty
         Call FormatFlxTrans
    If BranchPO.isEdit = True Then
        cboBranch.Text = BranchPO.txtbranch.Text
        cboBranch.Enabled = False
        cmdTrans(0).Caption = "Insert"
    Else
        cboBranch.Enabled = True
        cmdTrans(0).Caption = "Load Order"
        
    End If
End Sub
Private Sub FormatFlxSPEmpty()
    With flxSP
        '.Clear
        .Rows = 2
        .FormatString = " |<PARTNUMBER|<DESCRIPTION |^P.O NUMBER|^ORDER |^ALLTD.QTY|^B.O QTY|^CWOV|^ AMOUNT |^SPID"
      
        .ColWidth(0) = 300
        .ColWidth(1) = (.Width - 300) * 0.12
        .ColWidth(2) = (.Width - 300) * 0.28
        .ColWidth(3) = (.Width - 300) * 0.1
        .ColWidth(4) = (.Width - 300) * 0.096
        .ColWidth(5) = (.Width - 300) * 0.09
        .ColWidth(6) = (.Width - 300) * 0.065
        .ColWidth(7) = (.Width - 300) * 0.12
        .ColWidth(8) = (.Width - 300) * 0.12
        .ColWidth(9) = 0
    
    End With
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    MotoRep.sBar.Panels(6).Text = ""
    If rssummary.State = adStateOpen Then rssummary.Close
End Sub
Private Sub lvButtons_H1_Click()
    Unload Me
End Sub
Private Sub txtFind_Change()
    fRow = 1
    fCol = 1
End Sub
Private Sub txtFind_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr(KeyAscii)))
    If KeyAscii = 13 Then Call cmdFind_Click
End Sub
Function UpperCaseText(iKeyAscii As Integer) As Integer
Dim char As String
    char = Chr(iKeyAscii)
    UpperCaseText = Asc(UCase(char))
End Function
Private Sub DisplayPO()
 Dim rsFil As New ADODB.Recordset
 Dim NoOfSP As Long, NoOfRow As Long, nPO As Long, TotAllPO As Currency
 Dim myW As String, myW1 As String, fl As Long, TOTPOOrder As Long, MyPO As String
    
    myW1 = ""
    If chkPO.Value = 0 Then
        myW1 = "AND spbranchdtl.REFERENCE = '" & cboRefNo.Text & "'"
    Else
        For i = 1 To flxAllPO.Rows - 1
            If i = 1 Then
                myW1 = myW1 & " AND (spbranchdtl.REFERENCE = '" & flxAllPO.TextMatrix(i, 0) & "'"
            Else
                myW1 = myW1 & " OR spbranchdtl.REFERENCE = '" & flxAllPO.TextMatrix(i, 0) & "'"
            End If
        Next i
    End If
    myW = myW1
    
 
    'Qry = "SELECT DISTINCT spbranchdtl.SPAREPARTID From spbranchdtl " & _
    '      "LEFT OUTER JOIN spareparts ON (spbranchdtl.SPAREPARTID = spareparts.SPAREPARTID) " & _
    '      "Where  spbranchdtl.BRANCHNAME = '" & cboBranch.Text & "' " & myW & ") ORDER BY spareparts.PARTNUMBER;"
    
    'If rsFil.State = adStateOpen Then rsFil.Close
    '    rsFil.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    '    If rsFil.RecordCount > 0 Then
    '        flxPartsno.Clear
    '        Set flxPartsno.Recordset = rsFil
    '    End If
    
    'flxsp.Rows = 1
    'For fl = 0 To flxPartsno.Rows - 1
    '    Qry = "SELECT spareparts.PARTNUMBER, spareparts.DESCRIPTION, spbranchdtl.REFERENCE, '', " & _
    '          "' ',' ',spbranchdtl.PRICE, spbranchdtl.TOTAMOUNT, spbranchdtl.SPAREPARTID From spbranchdtl " & _
    '          "LEFT OUTER JOIN spareparts ON (spbranchdtl.SPAREPARTID = spareparts.SPAREPARTID) " & _
    '          "Where  spbranchdtl.BRANCHNAME = '" & cboBranch.Text & "' " & myW & ") AND spbranchdtl.SPAREPARTID = '" & flxPartsno.TextMatrix(fl, 0) & "' ORDER BY spareparts.PARTNUMBER;"
    '    With rsServe
    '        If .State = adStateOpen Then .Close
    '            .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    '            If .RecordCount = 0 Then
    '                .Close
    '                flxsp.Clear
    '                Call FormatFlxSPEmpty
    '                flxsp.Rows = 2
    '                Exit Sub
    '            Else
    '                Dim p%
    '                .MoveFirst
    '                myPO = ""
    '                For p = 1 To .RecordCount
    '                    myPO = myPO & IIf((myPO = ""), "", ",") & !REFERENCE
    '                    .MoveNext
    '                Next p
                    
    '                .MoveFirst
    '                Qry = "SELECT SUM(spbranchdtl.`ORDER`) AS TOTOrder From spbranchdtl " & _
    '                      "LEFT OUTER JOIN spareparts ON (spbranchdtl.SPAREPARTID = spareparts.SPAREPARTID) " & _
    '                      "Where  spbranchdtl.BRANCHNAME = '" & cboBranch.Text & "' " & myW & ") AND spbranchdtl.SPAREPARTID = '" & flxPartsno.TextMatrix(fl, 0) & "' ORDER BY spareparts.PARTNUMBER;"
                            
    '                      TOTPOOrder = Val(GetFieldValue(Qry))
    '            End If
    '            flxsp.AddItem vbTab & !PartNumber & vbTab & !Description & vbTab & myPO & vbTab & TOTPOOrder & vbTab & vbTab & vbTab & Format(!Price, "###,###,###.#0") & vbTab & Format(!TotAmount, "###,###,###.#0") & vbTab & !SparepartID
    '            .MoveNext
    '    End With
    '    flxsp.FixedRows = 1
    'Next fl
    
    '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    'Partnumber,Description, Order, Storeprice,TotalAmount, Allocated QTY, B.O Qty,SPAREPARTID
    
    If chkPO.Value = 1 And flxAllPO.Rows > 1 Then
        Qry = "SELECT spareparts.PARTNUMBER, spareparts.DESCRIPTION, spbranchdtl.REFERENCE,spbranchdtl.`ORDER`, '','', spbranchdtl.PRICE, spbranchdtl.TOTAMOUNT, spbranchdtl.SPAREPARTID From spbranchdtl " & _
          "INNER JOIN spareparts ON (spbranchdtl.SPAREPARTID = spareparts.SPAREPARTID) Where  spbranchdtl.BRANCHNAME = '" & cboBranch.Text & "' " & myW & ")" & " ORDER BY spareparts.PARTNUMBER;"
    ElseIf chkPO.Value = 0 Then
        Qry = "SELECT spareparts.PARTNUMBER, spareparts.DESCRIPTION, spbranchdtl.REFERENCE,spbranchdtl.`ORDER`, '','', spbranchdtl.PRICE, spbranchdtl.TOTAMOUNT, spbranchdtl.SPAREPARTID From spbranchdtl " & _
          "INNER JOIN spareparts ON (spbranchdtl.SPAREPARTID = spareparts.SPAREPARTID) Where  spbranchdtl.BRANCHNAME = '" & cboBranch.Text & "' " & myW & " ORDER BY spareparts.PARTNUMBER;"
    ElseIf chkPO.Value = 1 And flxAllPO.Rows = 1 Then
          Qry = "SELECT spareparts.PARTNUMBER, spareparts.DESCRIPTION, spbranchdtl.REFERENCE,spbranchdtl.`ORDER`, '','', spbranchdtl.PRICE, spbranchdtl.TOTAMOUNT, spbranchdtl.SPAREPARTID From spbranchdtl " & _
          "INNER JOIN spareparts ON (spbranchdtl.SPAREPARTID = spareparts.SPAREPARTID) Where  spbranchdtl.BRANCHNAME = '" & cboBranch.Text & "' " & myW & " ORDER BY spareparts.PARTNUMBER;"
    End If
    
    With rsServe
        If .State = adStateOpen Then .Close
            .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
            If .RecordCount = 0 Then
                .Close
                flxSP.Clear
                Call FormatFlxSPEmpty
                flxSP.Rows = 2
                Exit Sub
            End If
            flxSP.Rows = 1
            .MoveFirst
            For NoOfRow = 0 To .RecordCount - 1
                flxSP.AddItem vbTab & !PartNumber & vbTab & !Description & vbTab & !REFERENCE & vbTab & rsServe.Fields("ORDER").Value & vbTab & vbTab & vbTab & Format(!Price, "###,###,###.#0") & vbTab & Format(!TotAmount, "###,###,###.#0") & vbTab & !SPAREPARTID
                .MoveNext
            Next NoOfRow
         End With
    
    Call FormatFlxSP
    Call CalculatePO
    ActRow = 1
End Sub
Private Sub FormatFlxSP()
    Dim i As Integer
    With flxSP
        .FixedRows = 1
        .FormatString = " |<PARTNUMBER|<DESCRIPTION |^P.O NUMBER |^ORDER |^ALLTD.QTY|^B.O QTY|^CWOV|^ AMOUNT |^SPID"
        
        .ColWidth(0) = 300
        .ColWidth(1) = (.Width - 300) * 0.12
        .ColWidth(2) = (.Width - 300) * 0.28
        .ColWidth(3) = (.Width - 300) * 0.1
        .ColWidth(4) = (.Width - 300) * 0.096
        .ColWidth(5) = (.Width - 300) * 0.09
        .ColWidth(6) = (.Width - 300) * 0.065
        .ColWidth(7) = (.Width - 300) * 0.12
        .ColWidth(8) = (.Width - 300) * 0.12
        .ColWidth(9) = 0
        
        If flxSP.TextMatrix(1, 1) = "" Then Call FormatFlxSPEmpty: Exit Sub
        For i = 1 To flxSP.Rows - 1
            .TextMatrix(i, 8) = Format(Val(Replace(.TextMatrix(i, 8), ",", "")), "#,###,##0.#0")
            .TextMatrix(i, 7) = Format(Val(Replace(.TextMatrix(i, 7), ",", "")), "#,###,##0.#0")
        Next i
    End With
End Sub
Private Sub CalculatePO()
    Dim myW As String, myW1 As String, fxRw As Long
    Dim LastOrderID As Long, Qry As String, i As Long
    Dim GTotalAllQ As Long, GTotBackOdr As Long
    Dim TotAmntPO As Currency
    
    If flxSP.TextMatrix(1, 1) = "" Then
        txtGTotalAmount.Text = "0.00": txtGTotalAllQ.Text = 0: txtGTotBackOdr.Text = 0
        Exit Sub
    End If
   
    
    If chkPO.Value = 0 Then
        myW1 = "AND spbranchdtl.REFERENCE = '" & cboRefNo & "'"
    Else
        For i = 1 To flxAllPO.Rows - 1
            If i = 1 Then
                myW1 = myW1 & " AND spbranchdtl.REFERENCE = '" & flxAllPO.TextMatrix(i, 0) & "'"
            Else
                myW1 = myW1 & " OR spbranchdtl.REFERENCE = '" & flxAllPO.TextMatrix(i, 0) & "'"
            End If
        
        Next i
    End If
    myW = myW1

    For i = 1 To flxSP.Rows - 1
        Qry = ""
        Qry = "SELECT MAX(spordertaking.ORDERIDSQ) As LASTFINALSERVE From spordertaking " & _
            "Where spordertaking.BRANCHNAME = '" & cboBranch.Text & "' AND spordertaking.FINALQT = 1"
        LastOrderID = Val(GetFieldValue(Qry))
        Qry = ""
        Qry = "SELECT SUM(spbrchconfirmdtl.ORDERTAKING) AS CONORDER From spbrchconfirmdtl Where spbrchconfirmdtl.SPAREPARTID = " & flxSP.TextMatrix(i, 9) & " AND spbrchconfirmdtl.REFERENCE = '" & flxSP.TextMatrix(i, 3) & "' AND spbrchconfirmdtl.BRANCHID = " & cboBranch.ItemData(cboBranch.ListIndex) & ";"
        If rsServe.State = adStateOpen Then rsServe.Close
        rsServe.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If rsServe.RecordCount > 0 Then
            If rsServe!CONORDER > 0 Then
                flxSP.TextMatrix(i, 5) = Val(rsServe!CONORDER)
            Else
                flxSP.TextMatrix(i, 5) = 0
            End If
        Else
             flxSP.TextMatrix(i, 5) = 0
        End If
        rsServe.Close
        TotAmntPO = TotAmntPO + Val(Replace(flxSP.TextMatrix(i, 8), ",", ""))
        GTotalAllQ = GTotalAllQ + Val(Replace(flxSP.TextMatrix(i, 5), ",", ""))
        Qry = ""
        Qry = "SELECT spbrchconfirmdtl.SPAREPARTID, spbrchconfirmdtl.BBACKORDER, spbrchconfirmdtl.ORDERTAKING From spbrchconfirmdtl Where spbrchconfirmdtl.ORDERID = " & LastOrderID & " AND spbrchconfirmdtl.SPAREPARTID = " & flxSP.TextMatrix(i, 9) & " AND spbrchconfirmdtl.REFERENCE = '" & flxSP.TextMatrix(i, 3) & "';"
        If rsServe.State = adStateOpen Then rsServe.Close
        rsServe.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If rsServe.RecordCount > 0 Then
            If rsServe!BBACKORDER > 0 Then
                flxSP.TextMatrix(i, 6) = rsServe!BBACKORDER
            Else
                flxSP.TextMatrix(i, 6) = 0
            End If
        Else
            If LastOrderID <> 0 Then
                flxSP.TextMatrix(i, 6) = 0
            Else
                flxSP.TextMatrix(i, 6) = flxSP.TextMatrix(i, 4)
            End If
        End If
        If Val(flxSP.TextMatrix(i, 5)) = 0 Then
            flxSP.TextMatrix(i, 6) = flxSP.TextMatrix(i, 4)
        End If
        rsServe.Close
        GTotBackOdr = GTotBackOdr + Val(flxSP.TextMatrix(i, 6))
   Next i
   
        txtGTotBackOdr.Text = Format(GTotBackOdr, "#,###,##0.00")
        txtGTotalAllQ.Text = Format(GTotalAllQ, "#,###,##0.00")
        txtGTotalAmount.Text = Format(TotAmntPO, "#,###,##0.00")
End Sub
Private Sub StockTrans()
Dim rsTrans As New ADODB.Recordset
Dim Qty As String, i As Long, nPO As Long, AllPO As String, ref As Long

    'If chkPO.Value = 1 Then
    For ref = 0 To flxAllPO.Rows - 1
        AllPO = AllPO & IIf((AllPO = ""), "", ",") & flxAllPO.TextMatrix(ref, 0)
                   
    Next ref
    
    For nPO = 1 To flxAllPO.Rows - 1
        Qry = "SELECT stocktransfer.TRANSDATE, stocktransfer.TRANSFERID,' ', stocktransfer.TOTALCOST, stocktransfer.`STATUS`, stocktransfer.PONO, stocktransfer.REMARKS " & _
                    "From stocktransfer Where stocktransfer.PONO = '" & AllPO & "' AND stocktransfer.RECBRANCH = '" & cboBranch.Text & "';"
        With rsTrans
            .Open Qry, LYRcon, adOpenKeyset, adLockPessimistic
            If .RecordCount > 0 Then
                Set flxTrans.Recordset = rsTrans
                Call FormatFlxTrans
                
                For i = 1 To flxTrans.Rows - 1
                    flxTrans.TextMatrix(i, 0) = Format(flxTrans.TextMatrix(i, 0), "mm/dd/yyyy")
                    Qry = "SELECT SUM(stocktransdtl.Qty) As TOTALQTY From stocktransdtl Where stocktransdtl.TRANSFERID = " & flxTrans.TextMatrix(i, 1) & ";"
                    flxTrans.TextMatrix(i, 2) = Val(GetFieldValue(Qry))
                    flxTrans.TextMatrix(i, 3) = Format(flxTrans.TextMatrix(i, 3), "###,###,##0.00")
                Next i
            Else
                Call FormatFlxTrans
            End If
            .Close
        End With
    Next nPO
        
    'Else
    '    Qry = "SELECT stocktransfer.TRANSDATE, stocktransfer.TRANSFERID,' ', stocktransfer.TOTALCOST, stocktransfer.`STATUS`, stocktransfer.PONO, stocktransfer.REMARKS " & _
                    "From stocktransfer Where stocktransfer.PONO = '" & cboRefNo & "' AND stocktransfer.RECBRANCH = '" & cboBranch.Text & "';"
    '    With rsTrans
    '        .Open Qry, LYRcon, adOpenKeyset, adLockPessimistic
        
    '        If .RecordCount > 0 Then
    '            Set flxTrans.Recordset = rsTrans
    '            Call FormatFlxTrans
                
    '            For i = 1 To flxTrans.Rows - 1
    '                flxTrans.TextMatrix(i, 0) = Format(flxTrans.TextMatrix(i, 0), "mm/dd/yyyy")
    '                Qry = "SELECT SUM(stocktransdtl.Qty) As TOTALQTY From stocktransdtl Where stocktransdtl.TRANSFERID = " & flxTrans.TextMatrix(i, 1) & ";"
    '                flxTrans.TextMatrix(i, 2) = Val(GetFieldValue(Qry))
    '                flxTrans.TextMatrix(i, 3) = Format(flxTrans.TextMatrix(i, 3), "###,###,##0.00")
    '            Next i
                
    '        Else
    '            Call FormatFlxTrans
    '        End If
            
    '        .Close
    '    End With
    'End If
End Sub
Private Sub FormatFlxTrans()
    With flxTrans
        .FormatString = "<REF. DATE|<REF. NO |^QTY |^TOTAL PRICE |^ STATUS |^ P.O NUMBER |^REMARKS"
        
        .ColWidth(0) = .Width * 0.12
        .ColWidth(1) = .Width * 0.12
        .ColWidth(2) = .Width * 0.12
        .ColWidth(3) = .Width * 0.12
        .ColWidth(4) = .Width * 0.12
        .ColWidth(5) = .Width * 0.12
        .ColWidth(6) = .Width * 0.27
    End With
End Sub
Private Sub mnuMark_Click()
    With flxSP
        If Val(flxSP.TextMatrix(flxSP.Row, 8)) > 0 Then
            .Col = 0
            .CellFontName = "WINGDINGS"
            .CellFontSize = 12
            .Text = lblChk.Caption
            .CellAlignment = 4
        End If
    End With
End Sub

Private Sub mnuMarkAll_Click()
    Dim i As Integer
    
    With flxSP
        For i = 1 To .Rows - 1
            If Val(flxSP.TextMatrix(i, 8)) > 0 Then
                .Row = i
                .Col = 0
                .CellFontName = "WINGDINGS"
                .CellFontSize = 12
                .Text = lblChk.Caption
                .CellAlignment = 4
                .TextMatrix(i, 0) = lblChk.Caption
            End If
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
Private Sub OrderStatus()
    Dim nRow As Integer
        Call FormatFlxSP
        Call FormatFlxTrans
        
        'If flxsp.TextMatrix(1, 1) = "" Then Exit Sub
        
        If cboAllocated.Text = "BACKORDER" Then
            For nRow = flxSP.Rows - 1 To 1 Step -1
                If Val(flxSP.TextMatrix(nRow, 6)) = 0 Then
                    If flxSP.Rows = 2 Then
                        flxSP.Rows = 1
                        flxSP.Rows = 2
                        flxSP.FixedRows = 1
                        Call FormatFlxSPEmpty
                    Else
                        flxSP.RemoveItem (nRow)
                        'Call FormatFlxSPEmpty
                    End If
                End If
            Next nRow
        End If
        
        If cboAllocated.Text = "ALLOCATED" Then
            For nRow = flxSP.Rows - 1 To 1 Step -1
                If Val(flxSP.TextMatrix(nRow, 5)) = 0 Then
                    If flxSP.Rows = 2 Then
                        flxSP.Rows = 1
                        flxSP.Rows = 2
                        flxSP.FixedRows = 1
                    Else
                        flxSP.RemoveItem (nRow)
                    End If
                End If
            Next nRow
        End If
        
        Call CalculatePO
End Sub
Private Sub ClearHPO()
    cboRefNo.Clear
    chkMark.Value = 0
    flxSP.Clear
    flxTrans.Clear
    'txtPONumber.Text = ""
    txtdateUpload.Text = ""
    txtPODate.Text = ""
    txtOdrtype.Text = ""
    cboRefNo.Text = ""
    txtGTotalAmount.Text = ""
    txtGTotalAllQ.Text = ""
    txtGTotBackOdr.Text = ""
End Sub
Private Function CheckPendingPO() As Boolean
Dim rsPenPO As New ADODB.Recordset
Dim Qry As String, i As Long, fSQSERVE As Long

    Qry = "SELECT spordertaking.FINALQT From spordertaking Where spordertaking.BRANCHID = " & cboBranch.ItemData(cboBranch.ListIndex) & " AND spordertaking.FINALQT=0;"
    fSQSERVE = Val(GetFieldValue(Qry))
    Qry = "SELECT spordrconfirm.ORDERID From spordrconfirm Where spordrconfirm.SQSERVE = " & fSQSERVE & ";"
    
    With rsPenPO
        If .State = adStateOpen Then .Close
        .Open Qry, LYRcon, adOpenKeyset, adLockPessimistic
            If .RecordCount > 0 Then
                CheckPendingPO = True
                Exit Function
            Else
                CheckPendingPO = False
            End If
    End With
    
    'If chkPO.Value = 0 Then
        
    '        Qry = "SELECT spordrconfirm.FINALQT From spordrconfirm Where spordrconfirm.REFERENCE = '" & cboRefNo.Text & "' AND spordrconfirm.BRANCHID = " & cboBranch.ItemData(cboBranch.ListIndex) & " AND spordrconfirm.FINALQT=0;"
            
    '        With rsPenPO
    '            .Open Qry, LYRcon, adOpenKeyset, adLockPessimistic
    '                If .RecordCount > 0 Then
    '                    CheckPendingPO = True
    '                Else
    '                    CheckPendingPO = False
    '                End If
    '        End With
        
    'Else
        'For i = 1 To flxAllPO.Rows - 1
            'Qry = "SELECT spordrconfirm.FINALQT From spordrconfirm Where spordrconfirm.REFERENCE = '" & flxAllPO.TextMatrix(i, 0) & "' AND spordrconfirm.BRANCHID = " & cboBranch.ItemData(cboBranch.ListIndex) & " AND spordrconfirm.FINALQT=0;"
    '    Qry = "SELECT spordertaking.FINALQT From spordertaking Where spordertaking.BRANCHID = " & cboBranch.ItemData(cboBranch.ListIndex) & " AND spordertaking.FINALQT=0;"
        
    '    With rsPenPO
    '        If .State = adStateOpen Then .Close
    '        .Open Qry, LYRcon, adOpenKeyset, adLockPessimistic
    '            If .RecordCount > 0 Then
    '                CheckPendingPO = True
    '                Exit Function
    '            Else
    '                CheckPendingPO = False
    '            End If
    '    End With
        'Next i
    'End If
End Function

