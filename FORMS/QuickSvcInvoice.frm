VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{9636DF31-3E13-4EC2-AA5D-37D97950BC8F}#1.0#0"; "ResizeXtra.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form QuickSrvInvoice 
   Caption         =   "Cash Quick Service with Credit Invoice"
   ClientHeight    =   8550
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11865
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8550
   ScaleWidth      =   11865
   WindowState     =   2  'Maximized
   Begin MSFlexGridLib.MSFlexGrid FlxQS 
      Height          =   6975
      Left            =   120
      TabIndex        =   18
      Tag             =   "XY"
      Top             =   1440
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   12303
      _Version        =   393216
      Cols            =   10
      BackColorBkg    =   -2147483633
      SelectionMode   =   1
      AllowUserResizing=   1
      FormatString    =   $"QuickSvcInvoice.frx":0000
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
         Image           =   "QuickSvcInvoice.frx":0098
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
         Format          =   78315521
         CurrentDate     =   39322
      End
      Begin lvButton.lvButtons_H cmdNext 
         Height          =   375
         Left            =   3240
         TabIndex        =   13
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
         Image           =   "QuickSvcInvoice.frx":0632
         cBack           =   -2147483633
      End
      Begin lvButton.lvButtons_H cmdRefresh 
         Height          =   375
         Left            =   5400
         TabIndex        =   14
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
         Image           =   "QuickSvcInvoice.frx":0BCC
         cBack           =   -2147483633
      End
      Begin MSComCtl2.DTPicker dTo 
         Height          =   315
         Left            =   3720
         TabIndex        =   15
         Top             =   180
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         Format          =   101318657
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
         Image           =   "QuickSvcInvoice.frx":1166
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
         TabIndex        =   12
         Top             =   210
         Width           =   1335
      End
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   661
      Caption         =   "&Unpaid Only"
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
      Image           =   "QuickSvcInvoice.frx":1700
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   1
      Left            =   1800
      TabIndex        =   9
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
      Image           =   "QuickSvcInvoice.frx":1C9A
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   2
      Left            =   3600
      TabIndex        =   10
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
      Image           =   "QuickSvcInvoice.frx":2234
      cBack           =   -2147483633
   End
   Begin lvButton.lvButtons_H cmdTrans 
      Height          =   375
      Index           =   3
      Left            =   5400
      TabIndex        =   11
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
      Image           =   "QuickSvcInvoice.frx":27CE
      cBack           =   -2147483633
   End
   Begin MSFlexGridLib.MSFlexGrid flxCancel 
      Height          =   1695
      Left            =   1560
      TabIndex        =   17
      Top             =   8640
      Visible         =   0   'False
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   2990
      _Version        =   393216
      Rows            =   1
      Cols            =   7
      FixedCols       =   0
      FormatString    =   "Qty|<Unit|<Part Number|<Description |<Price | Amount|<SPID"
   End
   Begin ResizeXtra.RXResizer RXResizer1 
      Left            =   360
      Top             =   8640
      _ExtentX        =   847
      _ExtentY        =   847
      RegistryKey     =   "RXB49E"
      Registered      =   -1  'True
      m_MDI_Width     =   12105
      m_MDI_Height    =   9120
   End
   Begin VB.Label lblHeading 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Unpaid Quick Service Invoice"
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
      Caption         =   "Unpaid Quick Service Invoice"
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
   Begin VB.Menu mnuSalRet 
      Caption         =   "Sales Return"
      Visible         =   0   'False
      Begin VB.Menu mnuCancel 
         Caption         =   "Cancel Sales Return"
      End
   End
End
Attribute VB_Name = "QuickSrvInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ActRow As Integer, fRow As Integer, fCol As Integer
Dim Qry As String

Private Sub RefreshRecord(dCode As Integer)
    Dim rsQck As New ADODB.Recordset
    Dim i As Integer
    
    If dCode = 1 Then
        'Qry = "SELECT SRVCJOBORDER.CUSTOMER, SPINVOICE.INVDATE, SPINVOICE.SPINVNO, SPINVOICE.AMOUNT, REQUISITION.REQNO, SRVCJOBORDER.JOBORDERNO, REQUISITION.REQDATE, SPINVOICE.INVTYPE, SPINVOICE.ORNO " & _
              "FROM (REQUISITION INNER JOIN SRVCJOBORDER ON REQUISITION.JOBORDERID = SRVCJOBORDER.JOBORDERID) INNER JOIN SPINVOICE ON REQUISITION.INVOICENO = SPINVOICE.SPINVNO " & _
              "WHERE (((SPINVOICE.INVDATE)=#" & dFrom.Value & "#) AND ((SPINVOICE.INVTYPE)='CREDIT') AND ((SRVCJOBORDER.PAYTYPE)='CASH')) ORDER BY SPINVOICE.SPINVNO DESC , REQUISITION.REQNO;"
        'mySQL
        Qry = "SELECT SRVCJOBORDER.CUSTOMER, SPINVOICE.INVDATE, SPINVOICE.SPINVNO, SPINVOICE.AMOUNT, REQUISITION.REQNO, SRVCJOBORDER.JOBORDERNO, REQUISITION.REQDATE, SPINVOICE.INVTYPE, SPINVOICE.ORNO " & _
              "FROM REQUISITION INNER JOIN SRVCJOBORDER ON REQUISITION.JOBORDERID = SRVCJOBORDER.JOBORDERID INNER JOIN SPINVOICE ON REQUISITION.INVOICENO = SPINVOICE.SPINVNO " & _
              "WHERE SPINVOICE.INVDATE='" & Format(dFrom.Value, "yyyy-mm-dd") & "' AND SPINVOICE.INVTYPE='CREDIT' AND SRVCJOBORDER.PAYTYPE='CASH' ORDER BY SPINVOICE.SPINVNO DESC , REQUISITION.REQNO;"
        
    Else
        'Qry = "SELECT SRVCJOBORDER.CUSTOMER, SPINVOICE.INVDATE, SPINVOICE.SPINVNO, SPINVOICE.AMOUNT, REQUISITION.REQNO, SRVCJOBORDER.JOBORDERNO, REQUISITION.REQDATE, SPINVOICE.INVTYPE, SPINVOICE.ORNO " & _
              "FROM (REQUISITION INNER JOIN SRVCJOBORDER ON REQUISITION.JOBORDERID = SRVCJOBORDER.JOBORDERID) INNER JOIN SPINVOICE ON REQUISITION.INVOICENO = SPINVOICE.SPINVNO " & _
              "WHERE (((SPINVOICE.INVDATE)>=#" & dFrom.Value & "# And (SPINVOICE.INVDATE)<=#" & dTo.Value & "#) AND ((SPINVOICE.INVTYPE)='CREDIT') AND ((SRVCJOBORDER.PAYTYPE)='CASH')) ORDER BY SPINVOICE.SPINVNO DESC , REQUISITION.REQNO;"
        'mySQL
        Qry = "SELECT SRVCJOBORDER.CUSTOMER, SPINVOICE.INVDATE, SPINVOICE.SPINVNO, SPINVOICE.AMOUNT, REQUISITION.REQNO, SRVCJOBORDER.JOBORDERNO, REQUISITION.REQDATE, SPINVOICE.INVTYPE, SPINVOICE.ORNO " & _
              "FROM REQUISITION INNER JOIN SRVCJOBORDER ON REQUISITION.JOBORDERID = SRVCJOBORDER.JOBORDERID INNER JOIN SPINVOICE ON REQUISITION.INVOICENO = SPINVOICE.SPINVNO " & _
              "WHERE SPINVOICE.INVDATE>='" & Format(dFrom.Value, "yyyy-mm-dd") & "' And SPINVOICE.INVDATE<='" & Format(dTo.Value, "yyyy-mm-dd") & "' AND SPINVOICE.INVTYPE='CREDIT' AND SRVCJOBORDER.PAYTYPE='CASH' ORDER BY SPINVOICE.SPINVNO DESC , REQUISITION.REQNO;"
        
    End If
    If rsQck.State = adStateOpen Then rsQck.Close
    With rsQck
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        If .RecordCount > 0 Then
            .MoveFirst
            FlxQS.Rows = 1
            Do While Not .EOF
                '                  |<CUSTOMER NAME    |^INV. DATE  |^INVOICE NO.    |> INV. AMOUNT   |^REQUISITION#   |^ JOB ORDER# |^SALES RET.#   |^ RET. DATE |^   O.R. NO.    |^ BALANCE
                FlxQS.AddItem vbTab & !CUSTOMER & vbTab & Format(!InvDate, "mm/dd/yyyy") & vbTab & !SPINVNO & vbTab & Format(!Amount, "#,##0.00") & vbTab & !REQNO & vbTab & !JOBORDERNO & vbTab & vbTab & IIf(IsNull(!ORNO), "", !ORNO)
                .MoveNext
            Loop
        Else
            FlxQS.Rows = 1
            FlxQS.Rows = 2
        End If
        .Close
    End With
    If Not (FlxQS.Rows = 2 And FlxQS.TextMatrix(1, 1) = "") Then Call DisplayBalance
    ActRow = 1
End Sub

Private Sub DisplayBalance()
    Dim i As Integer, ORAmt As Currency, InvAmt As Currency, RetAmt As Currency, Bal As Currency
    Dim SRet As String, Qry As String
    
    With FlxQS
        For i = 1 To .Rows - 1
            If .TextMatrix(i, 8) <> "" Then 'Invoice with OR
                'Qry = "SELECT [SPAREPARTS]+[DOWNPAYMENT] AS ORAMT FROM PAYMENTS WHERE (((PAYMENTS.ORNO)='" & .TextMatrix(i, 8) & "'));"
                
                'mySQL
                Qry = "SELECT PAYMENTS.SPAREPARTS+PAYMENTS.DOWNPAYMENT AS ORAMT FROM PAYMENTS WHERE PAYMENTS.ORNO='" & .TextMatrix(i, 8) & "';"
                
                ORAmt = Val(GetFieldValue(Qry))
                InvAmt = Val(Replace(.TextMatrix(i, 4), ",", ""))
                Bal = InvAmt - ORAmt
                .TextMatrix(i, 9) = Format(Bal, "#,##0.00")
                'Qry = "SELECT SPITR.REFERENCE From SPITR " & _
                      "WHERE (((SPITR.TRANSTYPE)='SALES RETURN') AND ((SPITR.PACKINGLIST)='" & .TextMatrix(i, 3) & "'));"
                'mySQL
                Qry = "SELECT SPITR.REFERENCE From SPITR " & _
                      "WHERE SPITR.TRANSTYPE='SALES RETURN' AND SPITR.PACKINGLIST='" & .TextMatrix(i, 3) & "';"
                
                SRet = GetFieldValue(Qry)
                If SRet <> "" Then
                    .TextMatrix(i, 7) = SRet
                    'Qry = "SELECT Sum([RECEIVE]*[PRICE]) AS AMTRET From SPITR " & _
                          "WHERE (((SPITR.TRANSTYPE)='SALES RETURN') AND ((SPITR.PACKINGLIST)='" & .TextMatrix(i, 3) & "'));"
                    'mySQL
                    Qry = "SELECT Sum(SPITR.RECEIVE*SPITR.PRICE) AS AMTRET From SPITR " & _
                          "WHERE SPITR.TRANSTYPE='SALES RETURN' AND SPITR.PACKINGLIST='" & .TextMatrix(i, 3) & "';"
                    
                    RetAmt = Val(GetFieldValue(Qry))
                    .TextMatrix(i, 9) = Format(Bal - RetAmt, "#,##0.00")
                End If
                If Bal - RetAmt < 0 Then
                    Bal = GetAllocateDP(.TextMatrix(i, 8), .TextMatrix(i, 3))
                    .TextMatrix(i, 9) = Format((InvAmt - RetAmt) - Bal, "#,##0.00")
                    .TextMatrix(i, 8) = .TextMatrix(i, 8) & " [DP]"
                End If
                If Val(.TextMatrix(i, 9)) > 0 Then
                    .Col = 1: .Row = i: .ColSel = 9: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 250, 180)
                    .Col = 2
                End If
                RetAmt = 0
            Else
                'Qry = "SELECT SPITR.REFERENCE From SPITR " & _
                      "WHERE (((SPITR.TRANSTYPE)='SALES RETURN') AND ((SPITR.PACKINGLIST)='" & .TextMatrix(i, 3) & "'));"
                'mySQL
                Qry = "SELECT SPITR.REFERENCE From SPITR " & _
                      "WHERE SPITR.TRANSTYPE='SALES RETURN' AND SPITR.PACKINGLIST='" & .TextMatrix(i, 3) & "';"
                
                SRet = GetFieldValue(Qry)
                If SRet <> "" Then .TextMatrix(i, 7) = SRet
                .TextMatrix(i, 9) = .TextMatrix(i, 4)
                .Col = 1: .Row = i: .ColSel = 9: .RowSel = i: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 250, 180)
                .Col = 2
            End If
        Next i
   End With
End Sub

Function GetAllocateDP(MyOR As String, MyInv As String) As Currency
    Dim Qry As String, AmtAlloc As String, MyAA() As String, MyBB As String
    Dim j As Integer
    
    'Qry = "SELECT DPLEDGER.AMTALLOC From DPLEDGER WHERE (((DPLEDGER.ORNO)='" & MyOR & "'));"
    'mySQl
    Qry = "SELECT DPLEDGER.AMTALLOC From DPLEDGER WHERE DPLEDGER.ORNO='" & MyOR & "';"
    
    AmtAlloc = GetFieldValue(Qry)
    If AmtAlloc <> "" Then
        If InStr(AmtAlloc, ",") > 0 Then
            MyAA = Split(AmtAlloc, ",")
            For j = LBound(MyAA) To UBound(MyAA)
                If InStr(MyAA(j), MyInv) > 0 Then
                    MyBB = MyAA(j)
                    Exit For
                End If
            Next j
        Else
            MyBB = AmtAlloc
        End If
        GetAllocateDP = Val(Replace(MyBB, MyInv & "=", ""))
    Else
        GetAllocateDP = 0
    End If
    
End Function

Private Sub cmdFind_Click()
    Dim i As Integer, j As Integer
        
    If txtFind.Text = "" Then Exit Sub
    With FlxQS
        For i = fRow To .Rows - 1
            For j = fCol To 9
                If InStr(.TextMatrix(i, j), txtFind.Text) > 0 Then
                    .Col = 1: .Row = i
                    .ColSel = 9:  .RowSel = i
                    .TopRow = i
                    fRow = IIf(j = 9, i + 1, i)
                    fCol = IIf(j = 9, 1, j + 1)
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
    Dim i As Integer, MyBal As Currency
    
    Select Case Index
        Case 0  'Show Unpaid only
            With FlxQS
                For i = .Rows - 1 To 1 Step -1
                    MyBal = Val(Replace(.TextMatrix(i, 9), ",", ""))
                    If MyBal = 0 Then
                        If i = 1 And .Rows = 2 Then
                            .Rows = 1
                            .Rows = 2
                        Else
                            .RemoveItem i
                        End If
                    End If
                Next i
            End With
        Case 1  'Refresh
            Call RefreshRecord(Val(cmdNext.Tag))
        Case 2  'Print
            MsgBox "Not Yet Available", vbInformation, "Print"
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
End Sub


Private Sub flxQS_Click()
    ActRow = FlxQS.Row
End Sub

Private Sub flxQS_EnterCell()
    ActRow = FlxQS.Row
End Sub


Private Sub flxQS_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
    If ActRow = 0 Then Exit Sub
    'If DateValue(FlxQS.TextMatrix(ActRow, 1)) <> Date Then Exit Sub
    'If InStr(UCase$(FlxQS.TextMatrix(ActRow, 5)), "CANCELLED") > 0 Or Val(FlxQS.TextMatrix(ActRow, 6)) = 0 Then Exit Sub
    If Button = vbRightButton Then
       Call PopupMenu(mnuSalRet)
    End If
End Sub

Private Sub Form_Load()
    
    If Day(Date) <= 7 Then
        dFrom.Value = DateAdd("m", -1, DateSerial(Year(Date), Month(Date), 1))
    Else
        dFrom.Value = DateSerial(Year(Date), Month(Date), 1)
    End If
    dTo.Value = Date
    cmdNext.Tag = "2"
    Call RefreshRecord(2)
    ActRow = 0
    MotoRep.sBar.Panels(6).Text = "Cash Quick Service w/ Credit Invoice"
    'MotoRep.tbMain.Visible = False
    MotoRep.imgBG.Visible = False
    MotoRep.toolBarButton.Visible = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    MotoRep.sBar.Panels(6).Text = ""
'    If rsQck.State = adStateOpen Then rsQck.Close
End Sub

Private Sub mnuCancel_Click()
    Dim rsCDtls As New ADODB.Recordset
    Dim rsLedger As New ADODB.Recordset
    Dim Ans As Integer, MyReason As String, Qry As String, MyDetails As String, MyAcc As String
    Dim MyUnit As String, MyDesc As String, MyInv As String, MyProc As String
    Dim i As Integer, CustID As Integer, InvBal As Single, TtlAmtPd As Currency, MyCredit As Currency, MyDebit As Currency
    
    
    Qry = "SELECT USERACCESS.ACCESSTYPE From UserAccess WHERE (((USERACCESS.EMPID)=" & UserID & ") AND ((USERACCESS.MENUINDX)=29));"
    MyAcc = GetFieldValue(Qry)
    If MyAcc <> "FULL ACCESS" Then
        MsgBox "You don't have full access to this transaction." & vbCrLf & "Please contact your System Administrator", vbExclamation + vbOKOnly, "Access Denied"
        Exit Sub
    End If
    
    'Display Details in flxCancel
    'Qry = "SELECT SPITR.RECEIVE, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPITR.PRICE, [RECEIVE]*[PRICE] AS AMOUNT, SPITR.SPAREPARTID " & _
          "FROM SPITR INNER JOIN SPAREPARTS ON SPITR.SPAREPARTID = SPAREPARTS.SPAREPARTID " & _
          "WHERE (((SPITR.REFERENCE)='RET#" & FlxQS.TextMatrix(ActRow, 2) & "') AND ((SPITR.TRANSTYPE)='SALES RETURN'));"
    'mySQL
    Qry = "SELECT SPITR.RECEIVE, SPAREPARTS.PARTNUMBER, SPAREPARTS.DESCRIPTION, SPITR.PRICE, SPITR.RECEIVE]*SPITR.PRICE AS AMOUNT, SPITR.SPAREPARTID " & _
          "FROM SPITR INNER JOIN SPAREPARTS ON SPITR.SPAREPARTID = SPAREPARTS.SPAREPARTID " & _
          "WHERE SPITR.REFERENCE='RET#" & FlxQS.TextMatrix(ActRow, 2) & "' AND SPITR.TRANSTYPE='SALES RETURN';"
    
    With rsCDtls
        .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
        flxCancel.Rows = 1
        If .RecordCount > 0 Then
            .MoveFirst
            MyDetails = ""
            Do While Not .EOF
                '0-QTY,  1-PARTNO,  2-DESCRIPTION, 3-PRICE, 4-AMOUNT, 5-SPID
                flxCancel.AddItem !RECEIVE & vbTab & !PartNumber & vbTab & !Description & vbTab & !Price & vbTab & !Amount & vbTab & !SparepartID
                If InStr(UCase$(!Description), "4T OIL") > 0 Then
                    MyUnit = IIf(!RECEIVE > 1, "QRTS", "QRT")
                Else
                    MyUnit = IIf(!RECEIVE > 1, "PCS", "PC")
                End If
                If Len(!Description) < 30 Then
                    MyDesc = !Description & Space$(30 - Len(!Description))
                Else
                    MyDesc = Left$(!Description, 30)
                End If
                MyDetails = IIf(MyDetails <> "", MyDetails & vbCrLf, "") & !RECEIVE & vbTab & MyUnit & vbTab & !PartNumber & Space$(20 - Len(!PartNumber)) & vbTab & MyDesc & vbTab & Format(!Price, "#,##0.00") & vbTab & Format(!Amount, "#,##0.00")
                .MoveNext
            Loop
        End If
        .Close
    End With
    Ans = MsgBox("Do you want to Cancel Sales Return?" & vbCrLf & vbCrLf & _
        "Sales Return No.: " & FlxQS.TextMatrix(ActRow, 2) & vbTab & vbTab & vbTab & vbTab & "Sales Return Date : " & FlxQS.TextMatrix(ActRow, 1) & vbCrLf & vbCrLf & _
        "QTY." & vbTab & "UNIT" & vbTab & "PART NUMBER" & vbTab & vbTab & "DESCRIPTION" & vbTab & vbTab & vbTab & "PRICE" & vbTab & "AMOUNT" & vbCrLf & MyDetails, vbYesNo, "Cancel")
    If Ans = vbNo Then Exit Sub
    MyReason = InputBox("Please indicate reason for Cancellation of Sales Return:", "Reason for Cancel")
    If MyReason = "" Then
        MsgBox "Unable to indicate Reason..." & vbCrLf & "Cancellation of Sales Return failed!", vbExclamation + vbOKOnly, "Cancel Failed"
        Exit Sub
    End If
    'Retrace transaction made
    'Update Sales Return Remarks to CANCELLED
    LYRcon.Execute "UPDATE SRETURNSP SET SRETURNSP.REMARKS = 'CANCELLED', SRETURNSP.TOTALAMT = 0 WHERE (((SRETURNSP.RETNO)='" & FlxQS.TextMatrix(ActRow, 2) & "'));"
    
    With flxCancel
        For i = 1 To .Rows - 1
            'Cancel SPTIR
            'LYRcon.Execute "UPDATE SPITR SET SPITR.TRANSTYPE = 'CANCELLED', SPITR.RECEIVE = 0 " & _
                           "WHERE (((SPITR.REFERENCE)='RET#" & FlxQS.TextMatrix(ActRow, 2) & "') AND ((SPITR.SPAREPARTID)=" & .TextMatrix(i, 5) & "));"
            'mySQL
            LYRcon.Execute "UPDATE SPITR SET SPITR.TRANSTYPE = 'CANCELLED', SPITR.RECEIVE = 0 " & _
                           "WHERE SPITR.REFERENCE='RET#" & FlxQS.TextMatrix(ActRow, 2) & "' AND SPITR.SPAREPARTID=" & .TextMatrix(i, 5) & ";"
            
            'Update Qty-On Hand
            'LYRcon.Execute "UPDATE SPAREPARTS SET SPAREPARTS.QTYONHAND = [SPAREPARTS].[QTYONHAND]-" & Val(.TextMatrix(i, 0)) & _
                           " WHERE (((SPAREPARTS.SPAREPARTID)=" & .TextMatrix(i, 5) & "));"
            'mySQL
            LYRcon.Execute "UPDATE SPAREPARTS SET SPAREPARTS.QTYONHAND = SPAREPARTS.QTYONHAND-" & Val(.TextMatrix(i, 0)) & _
                           " WHERE SPAREPARTS.SPAREPARTID=" & .TextMatrix(i, 5) & ";"
            
        Next i
    End With
    
    'Qry = "SELECT SPINVOICE.SPINVNO From SPINVOICE WHERE (((SPINVOICE.SPINVNO)='" & FlxQS.TextMatrix(ActRow, 3) & "'));"
    'mySQL
    Qry = "SELECT SPINVOICE.SPINVNO From SPINVOICE WHERE SPINVOICE.SPINVNO='" & FlxQS.TextMatrix(ActRow, 3) & "';"
    
    If CtrlNumExist(Qry) Then   'Update Source Invoice field remarks to Null
        'LYRcon.Execute "UPDATE SPINVOICE SET SPINVOICE.REMARKS = Null WHERE (((SPINVOICE.SPINVNO)='" & FlxQS.TextMatrix(ActRow, 3) & "'));"
        'mySQL
        LYRcon.Execute "UPDATE SPINVOICE SET SPINVOICE.REMARKS = Null WHERE SPINVOICE.SPINVNO='" & FlxQS.TextMatrix(ActRow, 3) & "';"
        
    End If
    'Qry = "SELECT PAYMENTS.INVOICENO From PAYMENTS WHERE (((PAYMENTS.INVOICENO)='" & FlxQS.TextMatrix(ActRow, 3) & "'));"
    
    'mySQL
    Qry = "SELECT PAYMENTS.INVOICENO From PAYMENTS WHERE PAYMENTS.INVOICENO='" & FlxQS.TextMatrix(ActRow, 3) & "';"
    
    If CtrlNumExist(Qry) Then   'If Cash Update Payment's remarks to null
        'LYRcon.Execute "UPDATE PAYMENTS SET PAYMENTS.REMARKS = Null WHERE (((PAYMENTS.INVOICENO)='" & FlxQS.TextMatrix(ActRow, 3) & "'));"
        'mySQL
        LYRcon.Execute "UPDATE PAYMENTS SET PAYMENTS.REMARKS = Null WHERE PAYMENTS.INVOICENO='" & FlxQS.TextMatrix(ActRow, 3) & "';"
     
    End If
    'Update Ledger and On-Account Tables if Credit
    'CustID = GetFieldValue("SELECT SPINVOICE.CUSTOMERID, SPINVOICE.SPINVNO FROM SPINVOICE WHERE (((SPINVOICE.SPINVNO)='" & FlxQS.TextMatrix(ActRow, 3) & "'));")
    'mySQl
    CustID = GetFieldValue("SELECT SPINVOICE.CUSTOMERID, SPINVOICE.SPINVNO FROM SPINVOICE WHERE SPINVOICE.SPINVNO='" & FlxQS.TextMatrix(ActRow, 3) & "';")
    
    'Qry = "SELECT LEDGER.* From LEDGER WHERE (((LEDGER.REFNO)='" & FlxQS.TextMatrix(ActRow, 2) & "') AND ((LEDGER.CUSTOMERID)=" & CustID & "));"
    'mySQl
    Qry = "SELECT LEDGER.* From LEDGER WHERE LEDGER.REFNO='" & FlxQS.TextMatrix(ActRow, 2) & "' AND LEDGER.CUSTOMERID=" & CustID & ";"
   
    
    If CtrlNumExist(Qry) Then
        'Cancel Sales Return in Ledger
        'LYRcon.Execute "UPDATE LEDGER SET LEDGER.REFTYPE = 'CANCELLED', LEDGER.CUSTOMERID = 0 WHERE (((LEDGER.REFNO)='" & FlxQS.TextMatrix(ActRow, 2) & "'));"
        
        'mySQL
        LYRcon.Execute "UPDATE LEDGER SET LEDGER.REFTYPE = 'CANCELLED', LEDGER.CUSTOMERID = 0 WHERE LEDGER.REFNO='" & FlxQS.TextMatrix(ActRow, 2) & "';"
        
        'Update OS Bal in Ledger
        'InvBal = Val(GetFieldValue("SELECT SPINVOICE.AMOUNT From SPINVOICE WHERE (((SPINVOICE.SPINVNO)='" & FlxQS.TextMatrix(ActRow, 3) & "'));"))
        'mySQl
        InvBal = Val(GetFieldValue("SELECT SPINVOICE.AMOUNT From SPINVOICE WHERE SPINVOICE.SPINVNO='" & FlxQS.TextMatrix(ActRow, 3) & "';"))
        
        TtlAmtPd = 0
        'Qry = "SELECT LEDGER.* From LEDGER WHERE (((LEDGER.SPINVNO)='" & FlxQS.TextMatrix(ActRow, 3) & "') AND ((LEDGER.CUSTOMERID)=" & CustID & ")) ORDER BY LEDGER.SEQNO;"
        'mySQl
        Qry = "SELECT LEDGER.* From LEDGER WHERE LEDGER.SPINVNO='" & FlxQS.TextMatrix(ActRow, 3) & "' AND LEDGER.CUSTOMERID=" & CustID & " ORDER BY LEDGER.SEQNO;"
        
        With rsLedger   'Update Ledger
            .Open Qry, LYRcon, adOpenKeyset, adLockOptimistic
            If .RecordCount > 0 Then
                .MoveFirst
                Do While Not .EOF
                    MyCredit = IIf(IsNull(!Credit), 0, !Credit)
                    TtlAmtPd = TtlAmtPd + MyCredit
                    MyDebit = IIf(IsNull(!Debit), 0, !Debit)
                    If MyCredit > 0 Then
                        InvBal = InvBal - MyCredit
                    ElseIf MyDebit > 0 Then
                        InvBal = InvBal + MyDebit
                    End If
                    !OSbal = InvBal
                    .MoveNext
                Loop
            End If
            .Close
        End With
        'Update Onaccount table
        'LYRcon.Execute "UPDATE ONACCTERMS SET ONACCTERMS.TTLAMTPAID = " & TtlAmtPd & ", ONACCTERMS.BALANCE = " & InvBal & ", ONACCTERMS.STATUS = 'OPEN', ONACCTERMS.LASTUPDATE = #" & Now & "# " & _
                       "WHERE (((ONACCTERMS.SPINVNO)='" & FlxQS.TextMatrix(ActRow, 3) & "'));"
        'mySQL
        LYRcon.Execute "UPDATE ONACCTERMS SET ONACCTERMS.TTLAMTPAID = " & TtlAmtPd & ", ONACCTERMS.BALANCE = " & InvBal & ", ONACCTERMS.STATUS = 'OPEN', ONACCTERMS.LASTUPDATE = '" & Format(Now, "yyyy-mm-dd hh:mm:ss") & "' " & _
                       "WHERE ONACCTERMS.SPINVNO='" & FlxQS.TextMatrix(ActRow, 3) & "';"
        
        'Update Customer Outstanding Balance
        'LYRcon.Execute "UPDATE CUSTOMER SET CUSTOMER.OSBAL = [CUSTOMER].[OSBAL]+" & PureNum(FlxQS.TextMatrix(ActRow, 6)) & " WHERE (((CUSTOMER.CUSTOMERID)=" & CustID & "));"
        'mySQl
        LYRcon.Execute "UPDATE CUSTOMER SET CUSTOMER.OSBAL = CUSTOMER.OSBAL+" & Val(Replace(FlxQS.TextMatrix(ActRow, 6), ",", "")) & " WHERE CUSTOMER.CUSTOMERID=" & CustID & ";"
       
    End If
    'Update Audit Trans
    'LYRcon.Execute "INSERT INTO AUDITRANS ( TRANSDATE, TRANSTYPE, TRANSDOC, REFNO, REFDATE, REASON, EMPID ) " & _
            "VALUES (#" & Date & "#, 'CANCEL', 'SALES RETURN', '" & FlxQS.TextMatrix(ActRow, 2) & "', #" & FlxQS.TextMatrix(ActRow, 1) & "#, '" & MyReason & "', " & UserID & ");"
    'mySQl
    LYRcon.Execute "INSERT INTO AUDITRANS ( TRANSDATE, TRANSTYPE, TRANSDOC, REFNO, REFDATE, REASON, EMPID ) " & _
            "VALUES ('" & Format(Date, "yyyy-mm-dd") & "', 'CANCEL', 'SALES RETURN', '" & FlxQS.TextMatrix(ActRow, 2) & "', '" & Format(FlxQS.TextMatrix(ActRow, 1), "yyyy-mm-dd") & "', '" & MyReason & "', " & UserID & ");"
    
    'Update Grid
    With FlxQS
        .TextMatrix(ActRow, 7) = "CANCELLED"
        .TextMatrix(ActRow, 6) = "0.00"
        .Col = 1: .Row = ActRow: .ColSel = 8: .RowSel = ActRow: .FillStyle = flexFillRepeat: .CellBackColor = RGB(250, 250, 180)
        .Col = 2
    End With
    MsgBox "Sales Return Cancelled!", vbInformation + vbOKOnly, "Cancel"
                    
End Sub

Private Sub RXResizer1_ResizeComplete()
    Dim i As Integer
    
    lblHeading(2).Left = Frame1.Left + (Frame1.Width - lblHeading(2).Width) - 30
    lblHeading(3).Left = Frame1.Left + (Frame1.Width - lblHeading(3).Width)
    
    With FlxQS
        .ColWidth(0) = 300
        .ColWidth(1) = (.Width - 300) * 0.2
        For i = 2 To 9
            .ColWidth(i) = (.Width - 300) * 0.09
        Next i
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

