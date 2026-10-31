VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form SPSalesRpt 
   Caption         =   "SpareParts Sales Report"
   ClientHeight    =   8565
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11865
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8565
   ScaleWidth      =   11865
   WindowState     =   2  'Maximized
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   3
      Left            =   3600
      TabIndex        =   10
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
      Image           =   "SPSalesRpt.frx":0000
      cBack           =   -2147483633
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxAcct 
      Height          =   7095
      Left            =   120
      TabIndex        =   13
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
         TabIndex        =   5
         Tag             =   "XYF"
         Top             =   195
         Width           =   1935
      End
      Begin lvButton.lvButtons_H cmdFind 
         Height          =   375
         Left            =   8280
         TabIndex        =   6
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
         Image           =   "SPSalesRpt.frx":059A
         cBack           =   -2147483633
      End
      Begin MSComCtl2.DTPicker dFrom 
         Height          =   315
         Left            =   1440
         TabIndex        =   7
         Top             =   180
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         Format          =   103481345
         CurrentDate     =   39322
      End
      Begin lvButton.lvButtons_H cmdNext 
         Height          =   375
         Left            =   3240
         TabIndex        =   12
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
         Image           =   "SPSalesRpt.frx":0B34
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdRefresh 
         Height          =   375
         Left            =   5400
         TabIndex        =   14
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
         Image           =   "SPSalesRpt.frx":10CE
         cBack           =   -2147483633
      End
      Begin MSComCtl2.DTPicker dTo 
         Height          =   315
         Left            =   3720
         TabIndex        =   15
         Top             =   180
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         Format          =   103481345
         CurrentDate     =   39322
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
         Image           =   "SPSalesRpt.frx":1668
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
         TabIndex        =   11
         Top             =   210
         Width           =   1335
      End
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   1
      Left            =   0
      TabIndex        =   8
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
      Image           =   "SPSalesRpt.frx":1C02
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   2
      Left            =   1800
      TabIndex        =   9
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
      Image           =   "SPSalesRpt.frx":219C
      cBack           =   -2147483633
   End
   Begin VB.Label lblHeading 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Sales Report"
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
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   360
      Top             =   8640
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   12105
      m_MDI_Height    =   9135
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
      Top             =   450
      Width           =   2615
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
      TabIndex        =   1
      Top             =   475
      Width           =   2615
   End
   Begin VB.Label lblHeading 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Sales Report"
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
      Begin VB.Menu mnuDetails 
         Caption         =   "Show Details"
      End
      Begin VB.Menu mnuCancel 
         Caption         =   "Cancel Invoice"
      End
   End
End
Attribute VB_Name = "SPSalesRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsOnAcct As New ADODB.Recordset
Dim ActRow As Integer, fRow As Integer, fCol As Integer

Private Sub RefreshRecord(dCode As Integer)
    Dim Qry As String, myValidation As String
    Dim i As Integer
    
    If dCode = 1 Then
        Qry = "Select Group_concat(inv.TRANSNO) as TRANSNO From spinvoice as inv " & _
              "Left Join (Select REFERENCE From spitr Where REMARKS IN ('CASH','CREDIT') and TRANSDATE = '" & Format(dFrom.Value, "yyyy-mm-dd") & "' Group by REFERENCE) as tr ON tr.REFERENCE = inv.TRANSNO " & _
              "Where inv.INVDATE = '" & Format(dFrom.Value, "yyyy-mm-dd") & "' and inv.REMARKS NOT IN ('CANCELLED') and tr.REFERENCE is NULL;"
    Else
        Qry = "Select Group_concat(inv.TRANSNO) as TRANSNO From spinvoice as inv " & _
              "Left Join (Select REFERENCE From spitr Where REMARKS IN ('CASH','CREDIT') and TRANSDATE Between '" & Format(dFrom.Value, "yyyy-mm-dd") & "' and '" & Format(dTo.Value, "yyyy-mm-dd") & "' Group by REFERENCE) as tr ON tr.REFERENCE = inv.TRANSNO " & _
              "Where inv.INVDATE Between '" & Format(dFrom.Value, "yyyy-mm-dd") & "' and '" & Format(dTo.Value, "yyyy-mm-dd") & "' and inv.REMARKS NOT IN ('CANCELLED') and tr.REFERENCE is NULL;"
    End If
    If rsOnAcct.State = adStateOpen Then rsOnAcct.Close
    rsOnAcct.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    If rsOnAcct.RecordCount > 0 Then
        rsOnAcct.MoveFirst
        Do While Not rsOnAcct.EOF
            If IIf(IsNull(rsOnAcct!TransNo), "", rsOnAcct!TransNo) <> "" Then
                myValidation = "This Invoice had no Spareparts Details Found: " & rsOnAcct!TransNo
            End If
            rsOnAcct.MoveNext
        Loop
        rsOnAcct.Close
    End If
    
    If myValidation <> "" Then
        MsgBox myValidation, vbInformation + vbOKOnly, "Validation"
        ActRow = 1
        flxAcct.Clear
        flxAcct.Rows = 2
        flxAcct.Visible = True
        Call RXResizer1_ResizeComplete
        Exit Sub
    End If
    If dCode = 1 Then
        'Qry = "SELECT SPINVOICE.INVDATE, SPINVOICE.SPINVNO, SPINVOICE.CHARGETO, SPINVOICE.ADDRESS, SPINVOICE.AMOUNT, SPINVOICE.REMARKS, SPINVOICE.INVTYPE " & _
              "From SPINVOICE Where (((SPINVOICE.INVDATE) = #" & dFrom.Value & "#)) ORDER BY SPINVOICE.INVTYPE, SPINVOICE.INVDATE DESC , Val([SPINVNO]);"
        'mySQL
        Qry = "SELECT SPINVOICE.INVDATE, SPINVOICE.SPINVNO, SPINVOICE.CHARGETO, SPINVOICE.ADDRESS, SPINVOICE.AMOUNT, SPINVOICE.REMARKS, SPINVOICE.INVTYPE " & _
              "From SPINVOICE Where SPINVOICE.INVDATE = '" & Format(dFrom.Value, "yyyy-mm-dd") & "' ORDER BY SPINVOICE.INVTYPE, SPINVOICE.INVDATE DESC , round(SPINVOICE.SPINVNO);"
        
    Else
        'Qry = "SELECT SPINVOICE.INVDATE, SPINVOICE.SPINVNO, SPINVOICE.CHARGETO, SPINVOICE.ADDRESS, SPINVOICE.AMOUNT, SPINVOICE.REMARKS, SPINVOICE.INVTYPE " & _
              "From SPINVOICE Where (((SPINVOICE.INVDATE) >= #" & dFrom.Value & "# And (SPINVOICE.INVDATE) <= #" & dTo.Value & "#)) ORDER BY SPINVOICE.INVTYPE, SPINVOICE.INVDATE, Val([SPINVNO]);"
        'mySQL
        Qry = "SELECT SPINVOICE.INVDATE, SPINVOICE.SPINVNO, SPINVOICE.CHARGETO, SPINVOICE.ADDRESS, SPINVOICE.AMOUNT, SPINVOICE.REMARKS, SPINVOICE.INVTYPE " & _
              "From SPINVOICE Where SPINVOICE.INVDATE >= '" & Format(dFrom.Value, "yyyy-mm-dd") & "' And SPINVOICE.INVDATE <= '" & Format(dTo.Value, "yyyy-mm-dd") & "' ORDER BY SPINVOICE.INVTYPE, SPINVOICE.INVDATE, round(SPINVOICE.SPINVNO);"
        
    End If
    If rsOnAcct.State = adStateOpen Then rsOnAcct.Close
    rsOnAcct.Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
    If rsOnAcct.RecordCount = 0 Then
        flxAcct.Clear
        flxAcct.Rows = 2
        Call RXResizer1_ResizeComplete
        Exit Sub
    End If
    Set flxAcct.DataSource = rsOnAcct
    ActRow = 1
    flxAcct.Visible = True
    'flxacct2.Visible = False
    Call RXResizer1_ResizeComplete
End Sub


Private Sub cmdFind_Click()
    Dim i As Integer, j As Integer
        
    If txtFind.Text = "" Then Exit Sub
    With flxAcct
        For i = fRow To .Rows - 1
            For j = fCol To 6
                If InStr(.TextMatrix(i, j), txtFind.Text) > 0 Then
                    .Col = 1: .Row = i
                    .ColSel = 6:  .RowSel = i
                    .TopRow = i
                    fRow = IIf(j = 6, i + 1, i)
                    fCol = IIf(j = 6, 1, j + 1)
                    ActRow = fRow
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
    Dim dCov As String, MyQry As String, Qry As String
    
    
    Select Case Index
        Case 1  'Refresh
            Call RefreshRecord(Val(cmdNext.Tag))
        Case 2  'Print
            If Val(cmdNext.Tag) = 1 And Date = dFrom.Value Then
                Qry = "SELECT * From payments Where payments.UPDATEDBY = " & UserID & " AND payments.PAYDATE = '" & Format(Date, "yyyy-mm-dd") & "';"
                If CtrlNumExist(Qry) Then
                    MsgBox "You are restricted to view this report!" & vbCrLf & "Please contact your OIC for further notice.", vbExclamation + vbOKOnly, "Warning"
                    Exit Sub
                End If
            ElseIf Val(cmdNext.Tag) = 2 And Date >= dFrom.Value And Date <= dTo.Value Then
                Qry = "SELECT * From payments Where payments.UPDATEDBY = " & UserID & " AND payments.PAYDATE = '" & Format(Date, "yyyy-mm-dd") & "';"
                If CtrlNumExist(Qry) Then
                    MsgBox "You are restricted to view this report!" & vbCrLf & "Please contact your OIC for further notice.", vbExclamation + vbOKOnly, "Warning"
                    Exit Sub
                End If
            End If
        
            If Val(cmdNext.Tag) = 1 Then
                'MyQry = "SELECT SPINVOICE.INVDATE, SPINVOICE.SPINVNO, SPINVOICE.CHARGETO, SPINVOICE.ADDRESS, SPINVOICE.AMOUNT, SPINVOICE.REMARKS, SPINVOICE.INVTYPE " & _
                      "From SPINVOICE Where (((SPINVOICE.INVDATE) = #" & dFrom.Value & "#)) ORDER BY SPINVOICE.INVTYPE, SPINVOICE.INVDATE DESC, SPINVOICE.SPINVNO;"
                'mySQL
                MyQry = "SELECT SPINVOICE.INVDATE, SPINVOICE.SPINVNO, SPINVOICE.CHARGETO, SPINVOICE.ADDRESS, SPINVOICE.AMOUNT, SPINVOICE.REMARKS, SPINVOICE.INVTYPE " & _
                      "From SPINVOICE Where SPINVOICE.INVDATE = '" & Format(dFrom.Value, "yyyy-mm-dd") & "' ORDER BY SPINVOICE.INVTYPE, SPINVOICE.INVDATE DESC, SPINVOICE.SPINVNO;"
                
            Else
                'MyQry = "SELECT SPINVOICE.INVDATE, SPINVOICE.SPINVNO, SPINVOICE.CHARGETO, SPINVOICE.ADDRESS, SPINVOICE.AMOUNT, SPINVOICE.REMARKS, SPINVOICE.INVTYPE " & _
                      "From SPINVOICE Where (((SPINVOICE.INVDATE) >= #" & dFrom.Value & "# And (SPINVOICE.INVDATE) <= #" & dTo.Value & "#)) ORDER BY SPINVOICE.INVTYPE, SPINVOICE.INVDATE,  SPINVOICE.SPINVNO;"
                'mySQL
                MyQry = "SELECT SPINVOICE.INVDATE, SPINVOICE.SPINVNO, SPINVOICE.CHARGETO, SPINVOICE.ADDRESS, SPINVOICE.AMOUNT, SPINVOICE.REMARKS, SPINVOICE.INVTYPE " & _
                      "From SPINVOICE Where SPINVOICE.INVDATE >= '" & Format(dFrom.Value, "yyyy-mm-dd") & "' And SPINVOICE.INVDATE <= '" & Format(dTo.Value, "yyyy-mm-dd") & "' ORDER BY SPINVOICE.INVTYPE, SPINVOICE.INVDATE,  SPINVOICE.SPINVNO;"
                
            End If
            dCov = "Date Covered: " & Format(dFrom.Value, "Mmm. dd, yyyy") & IIf(Val(cmdNext.Tag) <> 1, " to " & Format(dTo.Value, "Mmm. dd, yyyy"), "")
            PRNmc.QryPrn = MyQry
            PRNmc.ParamValue1 = dCov
            PRNmc.ParamValue2 = GetUserName
            PRNmc.MySource = "SPSALESRPT"
            PRNmc.Show
        Case 3  'Exit
            Unload Me
            'MotoRep.tbMain.Visible = True
            'MotoRep.imgBG.Visible = True
            MotoRep.toolBarButton.Visible = True
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
    Call RefreshRecord(Val(cmdNext.Tag))
End Sub

Private Sub flxAcct_Click()
    ActRow = flxAcct.Row
End Sub


Private Sub flxacct_EnterCell()
    ActRow = flxAcct.Row
End Sub

Private Sub flxAcct_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
    If Button = vbRightButton Then
        Call PopupMenu(mnuSPinv)
    End If
End Sub

Private Sub Form_Load()
    'Old Code
    'If Day(Date) <= 7 Then
    '    dFrom.Value = DateAdd("m", -1, DateSerial(Year(Date), Month(Date), 1))
    'Else
    '    dFrom.Value = DateSerial(Year(Date), Month(Date), 1)
    'End If
    
    dFrom.Value = Date
    cmdNext.Tag = "1"
    Call RefreshRecord(1)
    ActRow = 0
    MotoRep.sBar.Panels(6).Text = "Spare Parts Sales Report"
    'MotoRep.tbMain.Visible = False
    MotoRep.imgBG.Visible = False
    MotoRep.toolBarButton.Visible = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    MotoRep.sBar.Panels(6).Text = ""
    If rsOnAcct.State = adStateOpen Then rsOnAcct.Close
End Sub

Private Sub mnuCancel_Click()
    MsgBox "Not applicable to this Transaction. Please use Cancellation Module", vbInformation + vbOKOnly, "Validation"
'    Dim MyAcc As String, Qry As String
'
'    'Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE (((USERACCESS.EMPID)=" & UserID & ") AND ((USERACCESS.MENUINDX)=73));"
'
'    'mySQL
'    Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE USERACCESS.EMPID=" & UserID & " AND USERACCESS.MENUINDX=73;"
'
'    MyAcc = GetFieldValue(Qry)
'    If MyAcc <> "FULL ACCESS" Then
'        MsgBox "You don't have full access to this transaction." & vbCrLf & "Please contact your System Administrator", vbExclamation + vbOKOnly, "Access Denied"
'        Exit Sub
'    End If
'    InvDetails.MyInv = flxAcct.TextMatrix(ActRow, 2)
'    InvDetails.MyType = flxAcct.TextMatrix(ActRow, 7)
'    InvDetails.CancelInv = True
'    InvDetails.iRow = ActRow
'    InvDetails.Show vbModal
End Sub

Private Sub mnuDetails_Click()
    InvDetails.MyInv = flxAcct.TextMatrix(ActRow, 2)
    InvDetails.MyType = flxAcct.TextMatrix(ActRow, 7)
    InvDetails.CancelInv = False
    InvDetails.Show vbModal
End Sub



Private Sub RXResizer1_ResizeComplete()
    lblHeading(2).Left = Frame1.Left + (Frame1.Width - lblHeading(2).Width) - 30
    lblHeading(3).Left = Frame1.Left + (Frame1.Width - lblHeading(3).Width)
    
    With flxAcct
        .FormatString = "   |^ DATE|^INVOICE NO.|<CUSTOMER|<ADDRESS |>AMOUNT|<REMARKS"
        .ColWidth(0) = 300
        .ColWidth(1) = (.Width - 300) * 0.12
        .ColWidth(2) = (.Width - 300) * 0.12
        .ColWidth(3) = (.Width - 300) * 0.22
        .ColWidth(4) = (.Width - 300) * 0.22
        .ColWidth(5) = (.Width - 300) * 0.12
        .ColWidth(6) = (.Width - 300) * 0.18
        .ColWidth(7) = 0
        For i = 1 To .Rows - 1
            .TextMatrix(i, 1) = Format$(.TextMatrix(i, 1), "mm/dd/yyyy")
            .TextMatrix(i, 5) = Format$(.TextMatrix(i, 5), "###,##0.00")
        Next i
        .ZOrder
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

